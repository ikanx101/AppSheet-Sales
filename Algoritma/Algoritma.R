# jangan lupa ganti working directory
#setwd("~/Documents/AppSheet-Sales/Algoritma")

# ========================
# start
rm(list=ls())
library(readxl)
library(dplyr)
library(tidyr)

# ========================
# jangan lupa ganti path file
#nama_file = "~/Documents/AppSheet-Sales/Damen/Convert Data Appsheet.xlsx"
nama_file = "/cloud/project/Damen/Convert Data Appsheet.xlsx"

# ========================
# extract nama sheets
shits = excel_sheets(nama_file)

# ========================
# extract database produk
dbase = read_excel(nama_file,
                   sheet = shits[3]) %>% 
  janitor::clean_names() %>% 
  mutate(item_standar = janitor::make_clean_names(item),
         brand = ifelse(brand == "TS","Tropicana Slim",brand),
         brand = ifelse(brand == "NS","NutriSari",brand))

# ========================
# extract data target utama
data = read_excel(nama_file,
                  sheet = shits[1]) %>% 
  janitor::clean_names()

# ========================
# ambil informasi yang diperlukan
header_data = colnames(data)
nama_item = dbase$item_standar
item_yg_dijual = header_data[header_data %in% nama_item]

# ========================
# ========================
# kita bagi-bagi datanya berdasarkan informasi yang ada
# pertama dari produk
data_1 = data[colnames(data) %in% c("id",item_yg_dijual)]
data = data[!colnames(data) %in% item_yg_dijual]

# kedua dari gimmick
data_2 = data %>% select(id,contains("gimmick"))
data = data %>% select(-contains("gimmick"))

# ketiganya sisanya
data_3 = data %>% select(-starts_with("pg"),
                         -starts_with("sec"),
                         -contains("omzet"),
                         -transaksi_penjualan,
                         -ns_tea_sweet_tea,
                         -ns_wdank_bajigur,
                         -ns_wdank_kopi_bajigur)

# ========================
# ========================
# proses pengerjaannya mungkin akan rumit. kenapa? 
# kalau kita lihat di sheet after, banyaknya baris akan tergantung dari banyaknya baris yang ada di gimmick dan produk yang terjual.
# maka dari itu, lebih baik semua data dbuat dalam bentuk list saja.
# nanti tinggal ditempel saja ke kanan.
data_1 = data_1 %>% split(.,.$id)
data_2 = data_2 %>% split(.,.$id)
data_3 = data_3 %>% split(.,.$id)

# sekarang kita akan kerjakan yang data_2
# kita rapikan gimmick
# rules: saat tidak ada gimmick, maka sisanya dbuat nol alias NA
