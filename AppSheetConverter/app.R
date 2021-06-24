# =======================================
# ikanx101.com
# produly presents
# AppSheet Sales Converter v2.0
# =======================================

# alamatnya di: https://ikanx.shinyapps.io/AppSheetConverter


# =======================================
# panggil libraries yang dibutuhkan
rm(list=ls())
library(readxl)
library(dplyr)
library(tidyr)
library(reshape2)
library(shiny)
library(shinymanager)
# =======================================



# =======================================
# buat credential
credentials = data.frame(
    user = c("ikanx", "nutrifood"), # mandatory
    password = c("ikanx", "sales"), # mandatory
    admin = c(TRUE, TRUE),
    stringsAsFactors = FALSE
)
# =======================================



# =======================================
# user interface
title_pane = titlePanel("AppSheet Converter v1.0")
isi = fluidRow(column(width = 8,
                      h3("== Read Me! =="),
                      h4("Web apps converter ini digunakan untuk mengubah raw data AppSheet ke format yang dibutuhkan."),
                      h4("Pastikan bahwa file yang hendak di-convert harus berupa file Excel (format .xlsx) dengan ukuran < 5 Mb."),
                      h4("Ada dua files yang perlu dimasukkan ke dalam converter ini, yakni: raw data AppSheet dan database harga."),
                      h4("Pastikan bahwa raw data AppSheet yang dimasukkan sesuai dengan standar yang telah disepakati bersama. Silakan download template data untuk mengecek apakah data Anda sudah sesuai dengan format yang ada."),
                      tags$a(href="https://github.com/ikanx101/AppSheet-Sales/raw/main/Damen/Template.xlsx", "Template Raw Data AppSheet."),
                      br(),
                      tags$a(href="https://github.com/ikanx101/AppSheet-Sales/raw/main/Damen/Template%20Harga.xlsx", "Template database harga produk."),
                      h4("Jika terjadi kendala dan pertanyaan silakan hubungi saya di rizka.fadhli@nutrifood.co.id"),
                      h5("Dibuat menggunakan R Studio"),
                      h6("Last update: 24 Juni 2021 10:44 WIB")
                      ),
               column(width = 4,
                      h3("== Converter =="),
                      h4("Silakan upload files Anda:"),
                      fileInput('nama_file_utama', 'Pilih file: raw data AppSheet'),
                      fileInput('nama_file_harga', 'Pilih file: database harga'),
                      h4("Harap tunggu 30-40 detik untuk proses di server."),
                      h5(textOutput("currentTime")),
                      downloadButton("downloadData", "Download Converted Data")
                      )
               )

# membuat user interface menjadi hidup
ui = fluidPage(title_pane,isi)
ui = secure_app(ui)
# =======================================




# =======================================
# otak dari converter ini
server <- function(input, output, session) {
    
    
    
    # =======================================
    # check credential dulu gaes
    res_auth <- secure_server(
        check_credentials = check_credentials(credentials)
    )
    # =======================================
    
    
    # =======================================
    # menampilkan waktu server
    output$currentTime <- renderText({
        invalidateLater(1000, session)
        paste("Waktu server saat ini:", Sys.time() + (7*60*60))
    })
    # =======================================
    
    
    
    # =======================================
    # proses perhitungan: otak dari segala otak
    data_final = reactive({
        # extract data target utama
        data = 
            read_excel(input$nama_file_utama$datapath) %>% 
            janitor::clean_names()
        
        # extract database produk
        dbase = 
            read_excel(input$nama_file_harga$datapath,
                       col_types = c("text", "text", "numeric")) %>% 
            janitor::clean_names() %>% 
            mutate(item_standar = janitor::make_clean_names(item),
                   brand = ifelse(brand == "TS","Tropicana Slim",brand),
                   brand = ifelse(brand == "NS","NutriSari",brand))
        
        # ========================
        # ambil informasi yang diperlukan
        header_data = colnames(data)
        nama_item = dbase$item_standar
        item_yg_dijual = header_data[header_data %in% nama_item]
        
        # ========================
        # ========================
        # kita bagi-bagi datanya berdasarkan informasi yang ada
        # pertama dari produk
        data_1 = data[colnames(data) %in% c("id",item_yg_dijual)]
        data = data[!colnames(data) %in% item_yg_dijual]
        
        # kedua dari gimmick
        data_2 = data %>% select(id,contains("gimmick"))
        data = data %>% select(-contains("gimmick"))
        
        # ketiganya sisanya
        data_3 = data %>% select(-starts_with("pg"),
                                 -starts_with("sec"),
                                 -contains("omzet"),
                                 -transaksi_penjualan,
                                 -ns_tea_sweet_tea,
                                 -ns_wdank_bajigur,
                                 -ns_wdank_kopi_bajigur)
        
        # ========================
        # ========================
        # proses pengerjaannya mungkin akan rumit. kenapa? 
        # kalau kita lihat di sheet after, banyaknya baris akan tergantung dari banyaknya baris yang ada di gimmick dan produk yang terjual.
        # maka dari itu, lebih baik semua data dbuat dalam bentuk list saja.
        # nanti tinggal ditempel saja ke kanan.
        data_1 = data_1 %>% split(.,.$id)
        data_2 = data_2 %>% split(.,.$id)
        data_3 = data_3 %>% split(.,.$id)
        
        # kita siapkan rumahnya dulu
        ikanx = vector("list",length(data_1))
        
        # sekarang kita mulai looping dari i = 1 sampai selesai
        for(i in 1: length(data_1)){
            temp_1 = data_1[[i]] %>% as.data.frame()
            temp_2 = data_2[[i]] %>% as.data.frame()
            temp_3 = data_3[[i]] %>% as.data.frame()
            
            # sekarang kita akan kerjakan yang data_2
            # kita rapikan gimmick
            # rules: saat tidak ada gimmick, maka sisanya dbuat nol alias NA
            if(temp_2$pemberian_gimmick == "Ada"){
                temp_2 = 
                    temp_2 %>% 
                    melt(id.vars = c("id","pemberian_gimmick")) %>% 
                    mutate(variable = ifelse(grepl("item",variable),
                                             "item_gimmick",
                                             "qty_gimmick")
                    )
                temp_2_1 = 
                    temp_2 %>% 
                    filter(grepl("item",variable)) %>% 
                    select(-variable) %>% 
                    rename(item_gimmick = value)
                temp_2_2 = 
                    temp_2 %>% 
                    filter(grepl("qty",variable)) %>% 
                    select(-variable) %>% 
                    rename(qty_gimmick = value) %>% 
                    select(qty_gimmick)
                temp_2 = cbind(temp_2_1,temp_2_2)
            } else {
                temp_2 = data.frame(
                    id = temp_2$id,
                    pemberian_gimmick = temp_2$pemberian_gimmick,
                    item_gimmick = NA,
                    qty_gimmick = NA
                )
            }
            
            # sekarang kita akan kerjakan yang data_1
            # item penjualan kita buat tabular
            temp_1 = 
                temp_1 %>% 
                melt(id.vars = "id") %>% 
                filter(!is.na(value)) %>% 
                rename(item_standar = variable) %>% 
                merge(dbase) %>% 
                mutate(omzet = value*harga) %>% 
                select(-item_standar) %>% 
                rename(qty_penjualan = value,
                       item_penjualan = item) %>% 
                relocate(id,item_penjualan,brand,qty_penjualan,harga,omzet)
            
            # sekarang saatnya moment of truth
            m1 = nrow(temp_1)
            m2 = nrow(temp_2)
            m3 = nrow(temp_3)
            
            max_m = max(c(m1,m2,m3))
            
            if(m1 < max_m){
                temp_1[(m1+1):max_m,] = NA
            }
            if(m2 < max_m){
                temp_2[(m2+1):max_m,] = NA
            }
            if(m3 < max_m){
                temp_3[(m3+1):max_m,] = temp_3[1,]
            }
            
            temp_1$id = NULL
            temp_2$id = NULL
            temp_3$id = NULL
            
            final = cbind(temp_3,temp_2,temp_1)
            ikanx[[i]] = final
        }
        
        # ========================
        # saatnya kita gabung kembali
        printed_data = data.frame()
        for(i in 1:length(data_1)){
            temp = ikanx[[i]]
            printed_data = rbind(temp,printed_data)
        }
        
        # hasil finalnya
        return(printed_data)
        
    })
    
    
    
    # =======================================
    # fungsi untuk download data
    data = data_final
    output$downloadData <- downloadHandler(
        filename = function() {
            paste0("Converted AppSheet ", Sys.time(), ".xlsx")
        },
        content = function(file) {
            openxlsx::write.xlsx(data(), file)
        }
    )
    # =======================================
    
    
    
}
# =======================================



# =======================================
# Run the application 
shinyApp(ui = ui, server = server)
# =======================================