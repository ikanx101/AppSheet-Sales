rm(list=ls())
gc()

library(readxl)
library(dplyr)
library(tidyr)
library(janitor)

setwd("~/AppSheet-Sales/Jevon")

# ==============================================================================
# tahap 1
# baca semua data frame dalam file
file = "X lap penjualan Nutrisari dan Hilo.xlsx"
df   = read_excel(file,col_names = F) %>% clean_names()
# ==============================================================================


# ==============================================================================
# tahap 2
# berikutnya kita akan pilah-pilah tabelnya
# caranya dengan mengecek ada di mana saja tabelnya

# kita cek dulu tabel header ada di mana saja
marker_1 = (1:nrow(df))[which(grepl("customer",df$x2,ignore.case = T))]
# karena yang akhir gak ada pasangan, kita tambahin manual
marker_1 = c(marker_1,nrow(df))

# Fungsi untuk generate angka antara dua nilai
generate_angka_antara <- function(start, end) {
  if (end - start <= 1) {
    return(NULL)
  }
  seq(from = start, to = end - 1, by = 1)
}

# Generate angka untuk setiap pasangan berurutan
hasil_generate <- list()

for (i in 1:(length(marker_1) - 1)) {
  start <- marker_1[i]
  end <- marker_1[i + 1]
  
  # Generate angka antara start dan end
  angka_antara <- generate_angka_antara(start, end)
  
  if (!is.null(angka_antara)) {
    hasil_generate[[paste0("marker[", i, "] - marker[", i + 1, "]")]] <- angka_antara
  }
}
# ==============================================================================


# ==============================================================================
# tahap 3
# kita pisahkan tabelnya
rumah_kita = list()
for(ikanx in 1:length(hasil_generate)){
  # ambil tabel
  temp = df[hasil_generate[[ikanx]],]
  # ambil nama customer
  nama_customer = temp$x2[1]
  # ambil tabel transaksi
  temp = temp %>% filter(!is.na(x4)) 
  # rapihin nama kolomnya
  colnames(temp) = temp[1,]
  # hapus kolom tak perlu
  temp = temp %>% clean_names() %>% filter(unit != "Unit")
  # tambahin nama costumer
  temp = temp %>% mutate(nama_customer)
  
  rumah_kita[[ikanx]] = temp
}

# kita gabungin untuk dirapikan kembali
gabung_tabel = data.table::rbindlist(rumah_kita) %>% as.data.frame()
# ==============================================================================


# ==============================================================================
# tahap 4
# kita benerin nama toko dan kode toko
# trus kita rapihin tabelnya
hasil_akhir = 
  gabung_tabel %>% 
  mutate(nama_customer = gsub("CUSTOMER : ","",nama_customer,fixed = T)) %>%
  mutate(nama_toko = stringr::str_squish(nama_customer)) %>% 
  separate(nama_customer,
           into = c("hapus","kode_toko"),
           sep = "\\(") %>% 
  mutate(kode_toko = gsub("\\)","",kode_toko)) %>% 
  select(-hapus) %>% 
  mutate(no_faktur = NA,
         bln_faktur = NA,
         tgl_faktur = NA,
         alamat_toko = NA,
         nama_salesman = NA,
         jenis_faktur = "PENJUALAN",
         hrg_satuan = NA,
         disc_total = NA) %>% 
  mutate(brand = case_when(
    grepl("NS|nutri|sari",product_name,ignore.case = T) ~ "NS",
    grepl("hilo|hi lo|hl",product_name,ignore.case = T) ~ "HiLo",
    grepl("ts|tropica",product_name,ignore.case = T) ~ "TS",
    grepl("lmen|l men|l-men",product_name,ignore.case = T) ~ "LMen"
  )) %>% 
  rename(kode_item = product_id,
         nama_item = product_name,
         qty_barang = qty_amount,
         satuan = unit,
         disc_satuan = discount,
         total_net = total_invoice) %>% 
  select(no_faktur,bln_faktur,tgl_faktur,kode_toko,nama_toko,alamat_toko,
         nama_salesman,kode_item,nama_item,qty_barang,satuan,
         hrg_satuan,disc_satuan,disc_total,total_net,brand,jenis_faktur)

openxlsx::write.xlsx(hasil_akhir,file = "converted.xlsx")
# ==============================================================================



