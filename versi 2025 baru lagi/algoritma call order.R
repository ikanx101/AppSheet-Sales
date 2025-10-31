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
sht  = "Order"
df   = read_excel(file,sheet = sht,col_types = "text")

# master item
file = "Master Item.xlsx"
df_master = 
  read_excel(file) |> 
  janitor::clean_names() |> 
  # ini revisi 31 Oktober 2025
  select(item_group,item_group_code) |> 
  rename("Nama Item" = item_group,
         "Kode Item" = item_group_code)

# kita gabung dulu
df_1 = merge(df,df_master,by = "Nama Item") %>% select(-"Check In",-"Check Out",-"Durasi")

# kita ambil dari sheet call
sht  = "Call"
df   = read_excel("Call sample.xlsx",sheet = sht,col_types = "text")
df_call = df %>% select("ID Call","Check In","Check Out","Durasi")

# sekarang kita gabung semua
df_final = 
  merge(df_1,df_call,by = "ID Call") %>% 
  select("ID Order","ID Call","Check In","Tanggal","Bulan","ID MDS","Nama MDS","PIC","Area MDS",
         "Region MDS","Kode Customer","Nama Customer","Kecamatan","Kabupaten","Provinsi","Alamat",
         "Detail Klasifikasi","Klasifikasi","Tipe Customer","Sekolah","Nama Cluster Firestart",
         "Koordinat RO","Koordinat Call","Jarak (Meter)","Kesesuaian Titik","Peserta Display Wow Operator",
         "Peserta Loyalty Sachet","Project OTG","Nama Distributor","Kecamatan Distributor",
         "Project 1","Project 2","Check Out","Durasi","Brand","Category","Nama Item","Kode Item","Qty",
         "Value","Firestart NS","Firestart Hilo","Status Sekolah") %>% 
  mutate("Status AV" = NA,
         "Berat(Gram)" = NA,
         "Tipe Transaksi" = "Call",
         "Jenis MDS" = NA) %>% 
  relocate("Status AV",.after = "Value") %>% 
  relocate("Berat(Gram)",.after = "Status AV") %>% 
  relocate("Tipe Transaksi",.after = "Berat(Gram)") %>% 
  relocate("Jenis MDS",.after = "Jenis MDS")


df_final 

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
saveWorkbook(wb, file = "Hasil convert Order.xlsx",overwrite = T)

