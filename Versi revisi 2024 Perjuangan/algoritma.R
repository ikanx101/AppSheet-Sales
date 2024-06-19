# ==============================================================================
# ini adalah algoritma untuk melakukan dan mengekstraksi data dari survey appsheet
# dibuat oleh: ikanx101.com
# versi: 19 juni 2024
# ==============================================================================


# ==============================================================================
# CHUNK 1

# kita mulai dari nol ya
rm(list = ls())
# set working directory
setwd("~/AppSheet-Sales/Versi revisi 2024 Perjuangan")

# libraries yang dibutuhkan
library(readxl)
library(dplyr)
library(tidyr)
library(reshape2)
library(parallel)

# kita tentukan berapa banyak cores yang terlibat
n_core = detectCores()

# kita mulai dengan memasukkan semua files yang terlibat
# dimulai dari memasukkan nama files excelnya terlebih dahulu
nama_file_utama = "Data Appsheet.xlsx"
nama_file_harga = "Template Harga 2.xlsx"
# ==============================================================================


# ==============================================================================
# CHUNK 2
# pada baris ini ke bawah, jangan diubah-ubah skripnya ya

# kita ekstrak data base harga
dbase =
  read_excel(nama_file_harga) |>
  janitor::clean_names() |>
  mutate(
    item_standar = janitor::make_clean_names(nama_item),
    brand = ifelse(brand == "TS", "TROPICANA SLIM", brand),
    brand = ifelse(brand == "NS", "NUTRISARI", brand)
  )

# kita simpan nama-nama item yang dijual
item_yang_dijual = dbase$item_standar

# kita akan mulai ekstrak data yang file utama
data = read_excel(nama_file_utama) %>% janitor::clean_names()

# kita simpan dulu nama variabel yang ada pada data utama
nama_var = colnames(data)

# kita akan simpan data pertama, yakni data yang bersifat umum
  # dimulai dari id
  id_1 = which(nama_var == "id")
  # diakhiri dengan pg3
  id_2 = which(nama_var == "pg3")
  # berikut adalah nama variabel yang dibutuhkan
  variabel_untuk_data_1 = nama_var[id_1:id_2]
  # berikut adalah data base pertama
  data_1 = data |> select(all_of(variabel_untuk_data_1))

# berikutnya kita akan ambil data untuk penjualannya
  # dimulai dari id lalu kita tambahin nama item yang dijual
  data_2 = 
    data |> 
    select(id,contains(item_yang_dijual)) |> 
    reshape2::melt(id.vars = "id") |> 
    # kita ganti namanya dulu agar bisa dimerge
    rename(item_standar = variable,
           qty_sold     = value) |> 
    # kita ubah dulu qty sold nya dulu ke numerik dan ganti nol
    mutate(qty_sold = as.numeric(qty_sold),
           qty_sold = ifelse(is.na(qty_sold),0,qty_sold)) |> 
    # kita merge dulu
    merge(dbase) |> 
    # lalu kita filter dulu dan ambil hanya variabel yang diperlukan
    select(-item_standar) |> 
    filter(qty_sold > 0) |> 
    # kita hitung omsetnya
    mutate(omset = qty_sold * harga) |> 
    merge(data_1) |> 
    arrange(id) |> 
    select(c(all_of(variabel_untuk_data_1),
             "nama_item","brand","sub_brand","harga",
             "qty_sold","omset")) |> 
    select(-contains("pg"))
  
  # kita simpan dulu ya hasilnya yang penjualan dulu
  openxlsx::write.xlsx(data_2,file = "penjualan_converted.xlsx")

# berikutnya adalah data ketiga, yakni availability dari produk-produk 
  # dimulai dari id
  id_1 = which(nama_var == "pg3")
  # diakhiri dengan pg3
  id_2 = which(nama_var == "pg_5")
  # berikut adalah nama variabel yang dibutuhkan
  variabel_untuk_data_3 = nama_var[id_1:id_2]
  # berikut adalah data base pertama
  data_3 = data |> select(id,all_of(variabel_untuk_data_3))
  
  # kita rapi-rapi dulu de ya
  data_3 =
    data_3 |> 
    reshape2::melt(id.vars = "id") |> 
    filter(value == T) |> 
    select(-value) |> 
    rename(availability = variable) |> 
    group_by(id) |> 
    mutate(av_item = length(availability)) |> 
    ungroup() |> 
    merge(data_1) |> 
    arrange(id) |> 
    select(c(all_of(variabel_untuk_data_1),
             "availability","av_item")) |> 
    select(-contains("pg")) |> 
    mutate(availability = gsub("\\_"," ",availability),
           availability = toupper(availability))
  
  # kita simpan dulu ya hasilnya yang penjualan dulu
  openxlsx::write.xlsx(data_3,file = "av_converted.xlsx")
  
  
  