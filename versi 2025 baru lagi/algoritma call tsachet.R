rm(list=ls())
gc()

library(dplyr)
library(tidyr)
library(openxlsx)
library(janitor)
library(readxl)
library(expss)

# mau mengajukan perubahan tipe data terkait kolom Check in, Check out dan Durasi ,,formatnya dibuat time,,

# call
nama_file_call = "Call (30).xlsx"
sht  = "TSachet"
df   = read_excel(nama_file_call,sheet = sht,col_types = "text") %>% rename("ID Order" = IDdetail)

# benerin tanggal
# 25 november 2025
df = 
  df |> 
  mutate(Tanggal = as.numeric(Tanggal)) |> 
  rowwise() |> 
  mutate(Tanggal = as.Date(Tanggal, origin = "1899-12-30")) |> 
  ungroup()

# master item
nama_file_master = "Master Item (4).xlsx"
df_master = 
  read_excel(nama_file_master) |> 
  janitor::clean_names() |> 
  # ini revisi 31 Oktober 2025
  select(item_group,item_group_code) |> 
  rename("Nama Item" = item_group,
         "Kode Item" = item_group_code)

# kita gabung dulu
df_1 = merge(df,df_master,by = "Nama Item")

# sekarang kita gabung semua
df_final = 
  df_1 %>%  
  select("ID Order","ID Call","Check In","Tanggal","Bulan","ID MDS","Nama MDS","PIC",
         "Area MDS","Region MDS","Kode Customer","Nama Customer","Kecamatan","Kabupaten","Provinsi",
         "Alamat","Detail Klasifikasi","Klasifikasi","Tipe Customer","Sekolah","Nama Cluster Firestart",
         "Koordinat RO","Koordinat Call","Jarak (Meter)","Kesesuaian Titik","Peserta Display Wow Operator",
         "Peserta Loyalty Sachet","Project OTG","Nama Distributor","Kecamatan Distributor",
         "Project 1","Project 2","Check Out","Durasi","Brand","Category","Nama Item","Kode Item",
         "Qty","Value"
  ) %>% 
  mutate("Status AV" = NA,
         "Berat(Gram)" = NA,
         "Tipe Transaksi" = "Call",
         "Jenis MDS" = NA) %>% 
  relocate("Status AV",.after = "Value") %>% 
  relocate("Berat(Gram)",.after = "Status AV") %>% 
  relocate("Tipe Transaksi",.after = "Berat(Gram)") %>% 
  relocate("Jenis MDS",.after = "Jenis MDS") %>% 
  mutate(Tanggal = as.Date(Tanggal,"%Y-%m-%d"))

# ====================================================================
# ini adalah tambahan daripada request akhir taun

# fungsi untuk mengubah menjadi check in dan check out
ubahin_waktu = function(tes){
  # tes <- "0.35410879629629627"
  tes_numeric <- as.numeric(tes)
  
  # Konversi
  detik_total <- tes_numeric * 86400
  jam <- floor(detik_total / 3600)
  menit <- floor((detik_total %% 3600) / 60)
  detik <- round(detik_total %% 60)
  
  # Format AM/PM
  if (jam >= 12) {
    periode <- "PM"
    if (jam > 12) jam <- jam - 12
  } else {
    periode <- "AM"
    if (jam == 0) jam <- 12
  }
  
  hasil <- sprintf("%d:%02d:%02d %s", jam, menit, detik, 
                   periode)
  # hasil
  return(hasil)
}

df_final$`Check In` = sapply(df_final$`Check In`,ubahin_waktu)
df_final$`Check Out`= sapply(df_final$`Check Out`,ubahin_waktu)

# fungsi untuk mengubah menjadi durasi
ubahin_waktu = function(tes){
  # tes <- "0.35410879629629627"
  tes_numeric <- as.numeric(tes)
  
  # Konversi
  detik_total <- tes_numeric * 86400
  jam <- floor(detik_total / 3600)
  menit <- floor((detik_total %% 3600) / 60)
  detik <- round(detik_total %% 60)
  
  hasil <- sprintf("%d:%02d:%02d", jam, menit, detik)
  # hasil
  return(hasil)
}

df_final$Durasi = sapply(df_final$Durasi,ubahin_waktu)
# ====================================================================


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
# saveWorkbook(wb, file = "Hasil convert TS Sachet v2.xlsx",overwrite = T)