# ==============================================================================
# ini adalah algoritma untuk melakukan dan mengekstraksi data dari survey appsheet
# dibuat oleh: ikanx101.com
# versi: 20 juni 2024
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

# fungsi untuk bikin judul proper
proper_new = function(x) {
  tess = stringi::stri_trans_general(x, id = "Title")
  gsub("\\_", " ", tess)
}
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
  
  # update pada tanggal 20 Juni
  # jangan lupa harus dimasukkan firestart dan juga kolom yang mengandung nama "project"
  tambah_var = nama_var[grepl("project|firestart|sahabat|wow|loyalty",nama_var)]
  # kita gabung lagi ke nama_var deh ya
  variabel_untuk_data_1 = c(variabel_untuk_data_1,tambah_var)

  # berikut adalah data base pertama
  data_1 = data |> select(all_of(variabel_untuk_data_1))
# ==============================================================================
  
  
  
# ==============================================================================
# CHUNK 3
  
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
    select(-contains("pg")) |> 
    rename(item_penjualan = nama_item) |> 
    mutate(tipe_transaksi = "Call")
  
  colnames(data_2) = proper_new(colnames(data_2))
  # kita simpan dulu ya hasilnya yang penjualan dulu
  openxlsx::write.xlsx(data_2,file = "penjualan_converted.xlsx")
# ==============================================================================
  
  
# ==============================================================================
# CHUNK 4
  
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
           availability = toupper(availability)) |> 
    mutate(tipe_transaksi = "AV") 
  
  colnames(data_3) = proper_new(colnames(data_3))
  # kita simpan dulu ya hasilnya yang penjualan dulu
  openxlsx::write.xlsx(data_3,file = "av_converted.xlsx")
# ==============================================================================

  
# ==============================================================================
# CHUNK 5
  
# kita akan gabung semua jadi satu data ke bawah
# pake apa? ya pakai bind_rows() aja donk 

# buat ngecek nama-nama kolom  
nama_kolom_2 = data_2 |> colnames()
nama_kolom_3 = data_3 |> colnames()
# melihat apakah ada perbedaan antara keduanya
setdiff(nama_kolom_2,nama_kolom_3)

# kita gabung
data_4 = bind_rows(data_2,data_3)

# kita simpan dulu ya hasilnya yang penjualan dulu
openxlsx::write.xlsx(data_4,file = "av_sales_gabung_converted.xlsx")
# ==============================================================================