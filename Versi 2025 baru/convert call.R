rm(list=ls())
gc()

library(dplyr)
library(tidyr)
library(openxlsx)
library(janitor)
library(readxl)

# call
file = "SAMPLE.xlsx"
sht  = "Merchandise"
df   = read_excel(file,sheet = sht,col_types = "text")

nama_kol    = colnames(df)
batas_akhir = which(nama_kol == "Durasi")

# kita save dulu variabel yang baru
df_baru = 
  df %>% 
  select(`ID Merch`,`Firestart NS`,`Firestart Hilo`,`Status Sekolah`) %>% 
  mutate(`Tipe Transaksi` = "AV") %>% 
  relocate(`Tipe Transaksi`,.after = `ID Merch`)

id_penting  = 1:batas_akhir
kol_penting = nama_kol[id_penting]

df_1 = df |> select(all_of(kol_penting))

kol_penting = nama_kol[2:batas_akhir]
df_2 = 
  df |> 
  select(-all_of(kol_penting)) |> 
  reshape2::melt(id.vars = "ID Merch") |> 
  filter(!grepl("sku|tier|fire|status",variable,ignore.case = T)) |> 
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
  relocate(`Nama Item`,.after = "Category") |> 
  # revisi 7 maret
  filter(!Status %in% c("Tidak Jual")) |> 
  filter(Status != "<NA>")

df_final = df_final |> merge(df_baru)

df_final$Tanggal = as.Date(as.numeric(df_final$Tanggal),
                           origin = "1900-01-01")
df_final$Tanggal = df_final$Tanggal - 2


# sekarang kita akan pecah ya
marker = nrow(df_final)

if(marker < 10^6){
  output = df_final
}

openxlsx::write.xlsx(output,file = "Call Converted.xlsx",overwrite = T)



