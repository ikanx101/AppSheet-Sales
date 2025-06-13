rm(list=ls())
gc()

library(dplyr)
library(tidyr)
library(readxl)

setwd("~/AppSheet-Sales/Versi 2025 baru/algoritma/jawa timur")

# ==============================================================================
# kita mengambil master data item
file = "Master Data.xlsx"
sht  = "Product"
df_master = 
  read_excel(file,sheet = sht) |> 
  janitor::clean_names() |> 
  select(item_group,category,brand) |> 
  rename("Nama Item" = item_group,
         Brand       = brand,
         Category    = category)
# ==============================================================================


# ==============================================================================
# sekarang kita ambil data utamanya
file = "01. Data NutriGO - Jan 2025.xlsx"
sht  = c("Order","Call","Merchandise")

# kita ambil data order, call, dan merchandise
df_order = 
  read_excel(file,sheet = sht[1]) %>% 
  janitor::clean_names() %>% 
  mutate(tanggal = as.Date(tanggal,"%Y-%m-%d"))
colnames(df_order) = read_excel(file,sheet = sht[1]) %>% colnames()

df_call = 
  read_excel(file,sheet = sht[2]) %>% 
  janitor::clean_names() %>% 
  mutate(check_in  = as.POSIXlt(check_in, format="%Y-%m-%d %H:%M:%S")) %>% 
  mutate(check_in  = format(check_in,"%H:%M:%S")) %>% 
  mutate(check_out = as.POSIXlt(check_out, format="%Y-%m-%d %H:%M:%S")) %>% 
  mutate(check_out = format(check_out,"%H:%M:%S")) %>% 
  mutate(durasi    = as.POSIXlt(durasi, format="%Y-%m-%d %H:%M:%S")) %>% 
  mutate(durasi    = format(durasi,"%H:%M:%S")) %>% 
  mutate(tanggal   = as.Date(tanggal,"%Y-%m-%d"))
colnames(df_call) = read_excel(file,sheet = sht[2]) %>% colnames()

df_av = 
  read_excel(file,sheet = sht[3]) %>% 
  janitor::clean_names() %>% 
  mutate(tanggal = as.Date(tanggal,"%Y-%m-%d"))
colnames(df_av) = read_excel(file,sheet = sht[3]) %>% colnames()
# ==============================================================================


# ==============================================================================
# sekarang kita akan ambil gabung call dan order
df_order$`Tipe Transaksi` = "Call"

df_call_order = 
  df_call %>% 
  select(`ID Call`,`Check In`,
         `Nama Cluster Firestart`,
         `Koordinat RO`,`Koordinat Call`,`Jarak (Meter)`,`Kesesuaian Titik`,
         `Check Out`,`Durasi`)

df_gabung = merge(df_order,df_call_order,by = "ID Call") %>% distinct()

df_gabung_1 =
  df_gabung %>% 
  mutate(`Peserta Display Wow Operator` = NA,
         `Peserta Loyalty Sachet`       = NA,
         `Project OTG`                  = NA,
         `Nama Distributor`             = NA,
         `Kecamatan Distributor`        = NA,
         `Project 1`                    = NA,
         `Project 2`                    = NA,
         `Status AV`                    = NA,
         `Berat(Gram)`                  = NA,
         `Firestart NS`                 = NA,
         `Firestart Hilo`               = NA
         ) %>% 
  mutate(`Status Sekolah` = ifelse(Sekolah == "Bukan Sekolah",FALSE,TRUE)) %>% 
  select(`ID Order`,`ID Call`,`Check In`,`Tanggal`,`Bulan`,`ID MDS`,`Nama MDS`,
         `PIC`,`Area MDS`,`Region MDS`,`Kode Customer`,`Nama Customer`,
         `Kecamatan`,`Kabupaten`,`Provinsi`,`Alamat`,`Detail Klasifikasi`,
         `Klasifikasi`,`Tipe Customer`,`Sekolah`,`Nama Cluster Firestart`,
         `Koordinat RO`,`Koordinat Call`,`Jarak (Meter)`,`Kesesuaian Titik`,
         `Peserta Display Wow Operator`,`Peserta Loyalty Sachet`,`Project OTG`,
         `Nama Distributor`,`Kecamatan Distributor`,`Project 1`,`Project 2`,
         `Check Out`,`Durasi`,`Brand`,`Category`,`Nama Item`,`Qty`,`Value`,
         `Status AV`,`Berat(Gram)`,`Tipe Transaksi`,`Firestart NS`,`Firestart Hilo`,
         `Status Sekolah`)
# ==============================================================================


# ==============================================================================
# sekarang kita akan gabung av dan merchandise
df_av_merch = 
  df_av %>% 
  select(`ID Merch`,`ID Call`,`Tanggal`,`Bulan`,`ID MDS`,`Nama MDS`,`PIC`,
         `Area MDS`,`Region MDS`,`Kode Customer`,`Nama Customer`,`Kecamatan`,
         `Kabupaten`,`Provinsi`,`Alamat`,`Detail Klasifikasi`,`Klasifikasi`,
         `Tipe Customer`,`Sekolah`) %>% 
  rename(`ID Order` = `ID Merch`) %>% 
  mutate(`Tipe Transaksi` = "AV") %>% 
  mutate(`Peserta Display Wow Operator` = NA,
         `Peserta Loyalty Sachet`       = NA,
         `Project OTG`                  = NA,
         `Nama Distributor`             = NA,
         `Kecamatan Distributor`        = NA,
         `Project 1`                    = NA,
         `Project 2`                    = NA,
         `Status AV`                    = NA,
         `Berat(Gram)`                  = NA,
         `Firestart NS`                 = NA,
         `Firestart Hilo`               = NA) %>% 
  mutate(`Status Sekolah` = ifelse(Sekolah == "Bukan Sekolah",FALSE,TRUE))

batas       = "Tier Lokalate & Wdank"
batas_hapus = 2:which(colnames(df_av) == batas)

df_av_produk = 
  df_av[-batas_hapus] %>% 
  rename(`ID Order` = `ID Merch`) %>% 
  melt(id.vars = "ID Order",na.rm = T) %>% 
  filter(value > 0) %>% 
  rename("Nama Item" = variable) %>% 
  rename(Qty = value) %>% 
  merge(df_master) %>% 
  mutate(Value = NA)

df_av_order = 
  df_call %>% 
  select(`ID Call`,`Check In`,
         `Nama Cluster Firestart`,
         `Koordinat RO`,`Koordinat Call`,`Jarak (Meter)`,`Kesesuaian Titik`,
         `Check Out`,`Durasi`)

df_gabung_order_1 = merge(df_av_merch,df_av_order,by = "ID Call") %>% distinct()
df_gabung_2       = merge(df_gabung_order_1,df_av_produk,by = "ID Order") %>% distinct()

df_gabung_2 = 
  df_gabung_2 %>% 
  select(`ID Order`,`ID Call`,`Check In`,`Tanggal`,`Bulan`,`ID MDS`,`Nama MDS`,
         `PIC`,`Area MDS`,`Region MDS`,`Kode Customer`,`Nama Customer`,
         `Kecamatan`,`Kabupaten`,`Provinsi`,`Alamat`,`Detail Klasifikasi`,
         `Klasifikasi`,`Tipe Customer`,`Sekolah`,`Nama Cluster Firestart`,
         `Koordinat RO`,`Koordinat Call`,`Jarak (Meter)`,`Kesesuaian Titik`,
         `Peserta Display Wow Operator`,`Peserta Loyalty Sachet`,`Project OTG`,
         `Nama Distributor`,`Kecamatan Distributor`,`Project 1`,`Project 2`,
         `Check Out`,`Durasi`,`Brand`,`Category`,`Nama Item`,`Qty`,`Value`,
         `Status AV`,`Berat(Gram)`,`Tipe Transaksi`,`Firestart NS`,`Firestart Hilo`,
         `Status Sekolah`)

df_final = rbind(df_gabung_1,df_gabung_2)










