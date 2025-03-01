rm(list=ls())
gc()

library(dplyr)
library(tidyr)
library(openxlsx)
library(janitor)
library(readxl)

# call
file = "Call.xlsx"
sht  = "Merchandise"
df   = read_excel(file,sheet = sht,col_types = "text")

nama_kol    = colnames(df)
batas_akhir = which(nama_kol == "Durasi")

id_penting  = 1:batas_akhir
kol_penting = nama_kol[id_penting]

df_1 = df |> select(all_of(kol_penting))

kol_penting = nama_kol[2:batas_akhir]
df_2 = 
  df |> 
  select(-all_of(kol_penting)) |> 
  reshape2::melt(id.vars = "ID Merch") |> 
  filter(!grepl("sku|tier",variable,ignore.case = T)) |> 
  mutate(variable = as.character(variable)) |> 
  rename("Nama Item" = variable,
         Status      = value) 

# master item
file = "Master Item.xlsx"
df_master = 
  read_excel(file) |> 
  janitor::clean_names() |> 
  select(item_group,category,brand) |> 
  rename("Nama Item" = item_group,
         Brand       = brand,
         Category    = category)

df_3 = df_2 |> merge(df_master,all.x = T)

df_final = 
  df_1 |> 
  merge(df_3) |> 
  relocate(Brand,.after = "Durasi") |> 
  relocate(Category,.after = "Brand") |> 
  relocate(`Nama Item`,.after = "Category")

openxlsx::write.xlsx(df_final,file = "Call Converted.xlsx",overwrite = T)




