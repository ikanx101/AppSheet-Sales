# jangan lupa ganti working directory
setwd("D:/AppSheet-Sales/Algoritma")

# ========================
# start
rm(list=ls())
library(readxl)
library(dplyr)
library(tidyr)
library(reshape2)

# ========================
# jangan lupa ganti path file
#nama_file = "~/Documents/AppSheet-Sales/Damen/Convert Data Appsheet.xlsx"
nama_file = "D:/AppSheet-Sales/Damen/Convert Data Appsheet.xlsx"

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

# kita siapkan rumahnya dulu
ikanx = vector("list",length(data_1))

# sekarang kita mulai dari i = 1
i = 2

temp_1 = data_1[[i]] %>% as.data.frame()
temp_2 = data_2[[i]] %>% as.data.frame()
temp_3 = data_3[[i]] %>% as.data.frame()

# sekarang kita akan kerjakan yang data_2
# kita rapikan gimmick
# rules: saat tidak ada gimmick, maka sisanya dbuat nol alias NA
if(temp_2$pemberian_gimmick == "Ada"){
  temp_2 = 
    temp_2 %>% 
    melt(id.vars = c("id","pemberian_gimmick")) %>% 
    mutate(variable = ifelse(grepl("item",variable),
                             "item_gimmick",
                             "qty_gimmick")
    )
  temp_2_1 = 
    temp_2 %>% 
    filter(grepl("item",variable)) %>% 
    select(-variable) %>% 
    rename(item_gimmick = value)
  temp_2_2 = 
    temp_2 %>% 
    filter(grepl("qty",variable)) %>% 
    select(-variable) %>% 
    rename(qty_gimmick = value) %>% 
    select(qty_gimmick)
  temp_2 = cbind(temp_2_1,temp_2_2)
} else {
  temp_2 = data.frame(
    id = temp_2$id,
    pemberian_gimmick = temp_2$pemberian_gimmick,
    item_gimmick = NA,
    qty_gimmick = NA
  )
}

# sekarang kita akan kerjakan yang data_1
# item penjualan kita buat tabular
temp_1 = 
  temp_1 %>% 
  melt(id.vars = "id") %>% 
  filter(!is.na(value)) %>% 
  rename(item_standar = variable) %>% 
  merge(dbase) %>% 
  mutate(omzet = value*harga) %>% 
  select(-item_standar) %>% 
  rename(qty_penjualan = value,
         item_penjualan = item) %>% 
  relocate(id,item_penjualan,brand,qty_penjualan,harga,omzet)

# sekarang saatnya moment of truth
m1 = nrow(temp_1)
m2 = nrow(temp_2)
m3 = nrow(temp_3)

max_m = max(c(m1,m2,m3))

if(m1 < max_m){
  temp_1[(m1+1):max_m,] = NA
}
if(m2 < max_m){
  temp_2[(m2+1):max_m,] = NA
}
if(m3 < max_m){
  temp_3[(m3+1):max_m,] = temp_3[1,]
}

temp_1$id = NULL
temp_2$id = NULL
temp_3$id = NULL

final = cbind(temp_3,temp_2,temp_1)
ikanx[[i]] = final
