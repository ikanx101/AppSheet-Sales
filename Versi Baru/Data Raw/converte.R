# ==============================================================================
#
# CONVERTER APPSHEET
# Versi Neo 2023
# by: ikanx101.com
# last update: 7 Februari 2024
#
# latar belakang
# [Tuesday 8:34 AM] Chandra Bimantara
# Mas, saat ini Appsheet akan di develop lagi utk menjadi tools mds motoris nasional
# Jadi, ada beberapa penyesuaian mas pd struktur datanya, utk mengakomodir kebutuhan di area2 lain juga agar bisa serentak digunakan nasional
# nah, seperti biasa mas saya ada kebutuhan untuk membuat converternya, agar data dapat di convert dan di sambungkan ke tabi nantinya
# saya bisa minta bantuannya lagi kah mas untuk pembuatan converternya? kali ini satu data, akan di convert menjadi 2 data mas. ada data penjualan dan data AV item di operator2
#
# ==============================================================================


# ==============================================================================
# libraries yang terlibat
library(dplyr)
library(tidyr)
library(parallel)
library(readxl)

# bebersih dulu
rm(list=ls())

# set working directory
setwd("~/AppSheet-Sales/Versi Baru/Data Raw")

# menentukan berapa banyak cores yang terlibat
n_core = detectCores()

# si warna merah yang tak diperlukan AV
# warna_merah = readLines("warna merah.txt") %>% as.numeric()

# si warna merah yang tak diperlukan availability
# warna_merah_av = readLines("warna merah - av.txt") %>% as.numeric()

# function ntuk mengembalikan nama produk
benerin = function(tes){
  gsub("_"," ",tes) %>% toupper()
}

# function untuk benerin nama kolom
nama_judul = function(tes){
  benerin(tes) %>% stringr::str_to_title()
}

# ==============================================================================


# ==============================================================================
# baca file harga
file_harga = "Template Harga 1.xlsx"
df_harga   = 
  read_excel(file_harga) %>% 
  janitor::clean_names() %>% 
  rename(item_penjualan = nama_item) %>% 
  # ini revisi terbaru
  mutate(item_penjualan = janitor::make_clean_names(item_penjualan),
         item_penjualan = benerin(item_penjualan))

# baca file yang hendak dikonversi
file_conv  = "Call (35).xlsx"
df_raw     = 
  read_excel(file_conv) %>% 
  janitor::clean_names() 

# ambil nama kolom untuk omset
nama_kolom = colnames(df_raw)
# pertengahan pg3 dan pg5 harus kita hapus
awal  = which(nama_kolom == "pg3")
akhir = which(nama_kolom == "pg_5")
hapus = awal:akhir

# ini yang perlu diambil
df_raw = df_raw[-hapus]
# ==============================================================================


# ==============================================================================
# tahap 1
# kita pisahkan untuk df omset terlebih dahulu
selection = c("waktu","tanggal","bulan","nama_mds","id_mds","area_mds",
              "region_mds","pic","kode_customer","nama_customer",
              "no_hp_customer","kecamatan","kabupaten","provinsi",
              "detail_klasifikasi","klasifikasi","sekolah",
              "koordinat_ro","koordinat_call","jarak_meter",
              "kesesuaian_titik","peserta_display_wow_operator",
              "peserta_loyalty_sachet","project_1","project_2",
              "transaksi_penjualan","av_item","check_out","durasi")

# pemisahan pertama
df_omset_raw_1 = 
  df_raw %>% 
  select(id,contains(selection))
# colnames(df_omset_raw_1)

# pemisahan kedua
# ambil nama kolom untuk omset
nama_kolom = colnames(df_raw)
# pertengahan pg3 dan pg5 harus kita hapus
awal   = which(nama_kolom == "pg")
akhir  = which(nama_kolom == "pg_8")
simpan = c(1,awal:akhir)

df_omset_raw_2 = 
  df_raw[simpan] %>% 
  select(-contains("pg")) %>% 
  reshape2::melt(id.vars = "id") %>% 
  filter(!is.na(value)) %>% 
  rename(item_penjualan = variable,
         qty_penjualan  = value) %>% 
  rowwise() %>% 
  mutate(item_penjualan = benerin(item_penjualan)) %>% 
  ungroup() %>% 
  merge(df_harga,by = "item_penjualan") %>% 
  mutate(omzet = harga * qty_penjualan)

# kita gabung kembali ke format yang diinginkan
df_gabung = 
  merge(df_omset_raw_1,df_omset_raw_2,by = "id") %>% 
  relocate(brand,sub_brand,harga,.after = "item_penjualan") %>% 
  relocate(av_item,check_out,durasi,.after = "omzet") %>% 
  # ini yang kita hapus
  select(-check_out,-durasi)

# benerin nama kolom finalnya
colnames(df_gabung) = nama_judul(colnames(df_gabung))

openxlsx::write.xlsx(df_gabung,file = "Omzet_converted.xlsx")
# ==============================================================================


# ==============================================================================
# tahap 2
# baca file yang hendak dikonversi
df_raw     = 
  read_excel(file_conv) %>% 
  janitor::clean_names() 

# ambil nama kolom untuk av
nama_kolom = colnames(df_raw)
# pertengahan pg3 dan pg5 harus kita hapus
awal     = which(nama_kolom == "pg3")
akhir    = which(nama_kolom == "pg_5")
ambil_av = c(1,awal:akhir)

# kita lakukan pemecahan kembali
# pemecahan 1
# pemisahan pertama
df_av_raw_1 = 
  df_raw %>% 
  select(id,contains(selection))

# pemisahan kedua
df_av_raw_2 = 
  df_raw[ambil_av] %>% 
  reshape2::melt(id.vars = "id") %>% 
  mutate(value = as.numeric(value)) %>% 
  filter(value > 0) %>% 
  select(-value) %>% 
  rename(availability_item = variable) %>% 
  mutate(availability_item = benerin(availability_item))

# kita gabung kembali ke format yang diinginkan
df_gabung = 
  merge(df_av_raw_1,df_av_raw_2,by = "id") %>% 
  relocate(availability_item,.before = "peserta_display_wow_operator") %>% 
  relocate(check_out,durasi,.after = "project_2") %>% 
  # ini yang kita hapus
  select(-check_out,-durasi)

# benerin nama kolom finalnya
colnames(df_gabung) = nama_judul(colnames(df_gabung))

openxlsx::write.xlsx(df_gabung,file = "AV_converted.xlsx")
# ==============================================================================
