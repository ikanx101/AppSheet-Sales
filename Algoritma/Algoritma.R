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




# ===============================================
# ide dasarnya seperti ini:
# ini data profil toko
tes = data.frame(id = 1,
                 nama = "ikang"
                )
# ini data gimmick
tes_2 = data.frame(id = 1,x = 3:9,y = 3:9)
# ini data jualan
tes_3 = data.frame(id = 1,z = 1:2, f = 4:5)

# kita merge dulu masing2 data gimmick dengan data profilnya
dr_1 = merge(tes,tes_2)
dr_2 = merge(tes,tes_3)

# kita cek berapa banyak baris dari kedua data hasil mergenya
m1 = nrow(dr_1)
m2 = nrow(dr_2)

# dari banyak baris tersebut terlihat, siapa yang harus ditambahkan agar proses penempelan berikutnya sesuai
# saya menggunakan perintah cbind() untuk menempel kedua dataset
if(m1<m2){
  dr_1[(m2-m1):m2,] = NA
  dr_1= dr_1%>% fill(id,nama) %>% select(-id,-nama)
  final = cbind(dr_2,dr_1)
} else if(m1>m2){
  dr_2[(m1-m2):m1,] = NA
  dr_2 = dr_2 %>% fill(id,nama) %>% select(-id,-nama)
  final = cbind(dr_1,dr_2)
}

# hasil akhir
# PR nya selanjutnya tinggal mengubah urutan saja ya
final
