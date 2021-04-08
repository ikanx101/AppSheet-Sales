#setwd("~/Documents/AppSheet-Sales/Algoritma")
setwd("/cloud/project/Algoritma")

rm(list=ls())
library(readxl)
library(dplyr)
library(tidyr)

#nama_file = "~/Documents/AppSheet-Sales/Damen/Convert Data Appsheet.xlsx"
nama_file = "/cloud/project/Damen/Convert Data Appsheet.xlsx"

# extract nama sheets
shits = excel_sheets(nama_file)

# extract database produk
dbase = read_excel(nama_file,
                   sheet = shits[3]) %>% 
  janitor::clean_names() %>% 
  mutate(item_standar = janitor::make_clean_names(item),
         brand = ifelse(brand == "TS","Tropicana Slim",brand),
         brand = ifelse(brand == "NS","NutriSari",brand))

# extract data target utama
data = read_excel(nama_file,
                  sheet = shits[1]) %>% 
  janitor::clean_names()

# ambil informasi yang diperlukan
header_data = colnames(data)
nama_item = dbase$item_standar
item_yg_dijual = header_data[header_data %in% nama_item]

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
colnames(data)
