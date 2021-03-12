setwd("~/Documents/AppSheet-Sales/Algoritma")

library(readxl)
library(dplyr)
library(tidyr)

nama_file = "~/Documents/AppSheet-Sales/Damen/Convert Data Appsheet.xlsx"
shits = excel_sheets(nama_file)

dbase = read_excel(nama_file,
                   sheet = shits[3]) %>% 
  janitor::clean_names() %>% 
  mutate(item = janitor::make_clean_names(item))

before = read_excel(nama_file,
                    sheet = shits[1]) %>% 
  janitor::clean_names()

header_before = colnames(before)
item_yg_ada = dbase$item
marker_non_item = header_before[!header_before %in% item_yg_ada]
