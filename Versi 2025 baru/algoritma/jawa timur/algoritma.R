rm(list=ls())
gc()

library(dplyr)
library(tidyr)
library(readxl)

setwd("~/AppSheet-Sales/Versi 2025 baru/algoritma/jawa timur")

# kita mengambil master data item
file = "Master Data.xlsx"
sht  = "Product"
df_master = 
  read_excel(file,sheet = sht) |> 
  janitor::clean_names() |> 
  select(item_group,category,brand) |> 
  rename("Nama Item" = item_group,
         Brand       = brand,
         Category    = category)

# sekarang kita ambil data utamanya
file = "01. Data NutriGO - Jan 2025.xlsx"
sht  = c("Order","Call","Merchandise")

# kita ambil data order, call, dan merchandise
df_order = 
  read_excel(file,sheet = sht[1]) %>% 
  janitor::clean_names() %>% 
  mutate(tanggal = as.Date(tanggal,"%Y-%m-%d"))

df_call = 
  read_excel(file,sheet = sht[2]) %>% 
  janitor::clean_names() %>% 
  mutate(check_in  = as.POSIXlt(check_in, format="%Y-%m-%d %H:%M:%S")) %>% 
  mutate(check_in  = format(check_in,"%H:%M:%S")) %>% 
  mutate(check_out = as.POSIXlt(check_out, format="%Y-%m-%d %H:%M:%S")) %>% 
  mutate(check_out = format(check_out,"%H:%M:%S")) %>% 
  mutate(durasi    = as.POSIXlt(durasi, format="%Y-%m-%d %H:%M:%S")) %>% 
  mutate(durasi    = format(durasi,"%H:%M:%S")) %>% 
  mutate(tanggal   = as.Date(tanggal,"%Y-%m-%d"))

df_av = 
  read_excel(file,sheet = sht[3]) %>% 
  janitor::clean_names()



