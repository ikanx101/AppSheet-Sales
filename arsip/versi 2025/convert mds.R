rm(list=ls())
gc()

library(dplyr)
library(tidyr)
library(openxlsx)
library(janitor)
library(readxl)

file = "MDS MODERN (2).xlsx"
sht  = "Rekap"

df = read_excel(file,sheet = sht)

nama_kol    = colnames(df)
batas_akhir = which(nama_kol == "Klasifikasi Customer")

id_penting  = 1:batas_akhir
kol_penting = nama_kol[id_penting]

df_1 = df |> select(all_of(kol_penting))

kol_penting = nama_kol[2:batas_akhir]
df_2 = 
  df |> 
  select(-all_of(kol_penting)) |> 
  reshape2::melt(id.vars = "ID") |> 
  mutate(variable = as.character(variable)) |>
  filter(value != "N/A") |> 
  filter(value != "Tidak Jual") |> 
  rename("Products"      = variable,
         "Status Produk" = value) 

df_final = 
  df_1 |> 
  merge(df_2) |> 
  janitor::clean_names() |> 
  mutate(amount = 0) |> 
  mutate(brand = case_when(
    grepl("hi lo",products,ignore.case = T) ~ "Hi Lo",
    grepl("l-men",products,ignore.case = T) ~ "L-Men",
    grepl("ns",products,ignore.case = T) ~ "NutriSari",
    grepl("ts",products,ignore.case = T) ~ "Tropicana Slim",
    grepl("hi lo",products,ignore.case = T) ~ "Hi Lo"
  )) |> 
  select(tanggal,nama_mds_spg,region,klasifikasi_customer,customer_name,customer_code,
         products,amount,status_produk,brand,divisi)

colnames(df_final) = c("Submission Date","Pic","Area","Klasifikasi Customer","Customer Name",
                       "Customer Code","Products","Amount","Status Produk","Brand","DIV")

openxlsx::write.xlsx(df_final,file = "MDS Converted.xlsx",overwrite = T)

