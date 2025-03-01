rm(list=ls())
gc()

library(dplyr)
library(tidyr)
library(openxlsx)
library(janitor)
library(readxl)

file = "MDS MODERN (2).xlsx"
sht  = "MODERN"

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
  rename("Products"      = variable,
         "Status Produk" = value) 

df_final = 
  df_1 |> 
  merge(df_2) 

openxlsx::write.xlsx(df_final,file = "MDS Converted.xlsx",overwrite = T)

