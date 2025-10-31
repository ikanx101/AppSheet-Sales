rm(list=ls())
gc()

library(dplyr)
library(tidyr)
library(openxlsx)
library(janitor)
library(readxl)
library(expss)

# call
file = "Call sample.xlsx"
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
  # ini revisi 31 Oktober 2025
  select(item_group,category,brand,item_group_code) |> 
  rename("Nama Item" = item_group,
         Brand       = brand,
         Category    = category,
         "Kode Item" = item_group_code)

df_3 = df_2 |> merge(df_master,all.x = T)

df_final = 
  df_1 |> 
  merge(df_3) |> 
  relocate(Brand,.after = "Durasi") |> 
  relocate(Category,.after = "Brand") |> 
  relocate(`Nama Item`,.after = "Category") |> 
  # ini revisi 31 Oktober 2025
  relocate("Kode Item",.after = "Nama Item") %>% 
  # revisi 7 maret
  filter(!Status %in% c("Tidak Jual")) |> 
  filter(Status != "<NA>")

df_final = df_final |> merge(df_baru)

df_final$Tanggal = as.Date(as.numeric(df_final$Tanggal),
                           origin = "1900-01-01")
df_final$Tanggal = df_final$Tanggal - 2


# ini revisi 31 Oktober 2025
# ini tambahannya
# kita ambil dari sheet call
sht  = "Call"
df   = read_excel("Call sample.xlsx",sheet = sht,col_types = "text")
df_call = df %>% select("ID Call","Check In","Check Out","Durasi")

# sekarang kita ganti isinya
df_finalista = 
  df_final %>% 
  select(-"Check In",-"Check Out",-"Durasi") %>% 
  merge(df_call,all.x = T) %>% 
  arrange("ID Call") %>% 
  relocate("Check In",.after = "ID Call") %>% 
  relocate("Durasi",.before = "Brand") %>% 
  relocate("Check Out",.before = "Durasi")

# kita tambahin lagi ya
df_finalista = 
  df_finalista %>% 
  mutate(Qty = NA,
         Value = NA,
         "Berat(Gram)" = NA,
         "Tipe Transaksi" = "AV",
         "Jenis MDS" = NA) %>% 
  relocate("Qty",.after = "Kode Item") %>% 
  relocate("Value",.after = "Qty") %>%
  rename("Status AV" = Status) %>% 
  relocate("Berat(Gram)",.after = "Status AV") %>% 
  mutate(Tanggal = as.Date(Tanggal,"%Y-%m-%d"))

# ini kita simpulkan cerita akhirnya
df_final = df_finalista

# sekarang kita akan pecah ya
marker      = nrow(df_final)
batas_pisah = 10^6

if(marker < batas_pisah){
  output = list(df_final)
}
if(marker > batas_pisah){
  output_1 = df_final[1:batas_pisah,]
  output_2 = df_final[(batas_pisah + 1):marker,]
  output   = list(output_1,output_2)
}

data_jatim = output
wb <- createWorkbook()

for(ikanx in 1:length(data_jatim)){
  sh = addWorksheet(wb, paste0("Sheet ",ikanx))
  xl_write(data_jatim[ikanx], wb, sh)
}

# Menyimpan workbook ke file
saveWorkbook(wb, file = "Hasil convert Merchandise.xlsx",overwrite = T)

