# converter appsheet milik sales operation


# ==============================================================================
# libraries
library(reshape2)
library(readxl)
library(dplyr)
library(tidyr)
library(janitor)
library(openxlsx)
library(shinydashboard)
library(shiny)
library(expss)
library(shinymanager)
# ==============================================================================

# ==============================================================================
# dimulai dari hati yang bersih
rm(list=ls())

# buat credential
credentials = data.frame(
  user             = c("sales", "a"), # mandatory
  password         = c("operation", "a"),  # mandatory
  admin            = c(TRUE, TRUE),
  stringsAsFactors = FALSE
)


# ==============================================================================
# USER INTERFACE PART

# header
header = dashboardHeader(title = "AppSheet Converter SO 2026",
                         titleWidth = 300)

# sidebar menu
sidebar = dashboardSidebar(width = 300,
                           sidebarMenu(
                             menuItem(tabName = 'readme',
                                      text = 'Read Me',icon = icon('check')),
                             menuItem(tabName = 'converter_1',
                                      text = 'Call Merchandise',icon = icon('dollar')),
                             menuItem(tabName = 'converter_2',
                                      text = 'Call Order',icon = icon('dollar')),
                             menuItem(tabName = 'converter_3',
                                      text = 'Call TSachet',icon = icon('dollar'))
                           )
)

# tab Read Me
readme = tabItem(tabName = 'readme',
                 fluidRow(
                   column(width = 12,
                          h1('Read Me'),
                          br(),
                          h4("Web apps converter ini digunakan untuk mengubah struktur data hasil export dari AppSheet menjadi struktur yang mudah dipivot di Ms. Excel."),
                          br(),
                          h5("Jika terjadi kendala atau pertanyaan, feel free to discuss ya: rizka.fadhli@nutrifood.co.id"),
                          br(),
                          br(),
                          h4(paste0("update 6 Januari 2026 15:12 WIB")),
                          h5("Copyright 2026"),
                          h5("Dibuat menggunakan R")
                   )
                 )
)

# tab converter 1
convert_1 = tabItem(tabName = 'converter_1',
                    fluidRow(
                      column(width = 12,
                             h1('Converter Call Merchandise'),
                             h4("Silakan upload dua files Anda:"),
                             fileInput('file_call_1', 'File Call (Merchandise)',
                                       accept = c('xlsx')
                             ),
                             br(),
                             fileInput('file_mtr_1', 'File Master Item',
                                       accept = c('xlsx')
                             ),
                             downloadButton("downloadData_1", "Download")
                      )
                    )
)


# tab converter 2
convert_2 = tabItem(tabName = 'converter_2',
                    fluidRow(
                      column(width = 12,
                             h1('Converter Call Order'),
                             h4("Silakan upload dua files Anda:"),
                             fileInput('file_call_2', 'File Call (Order)',
                                       accept = c('xlsx')
                             ),
                             br(),
                             fileInput('file_mtr_2', 'File Master Item',
                                       accept = c('xlsx')
                             ),
                             downloadButton("downloadData_2", "Download")
                      )
                    )
)

# tab converter 3
convert_3 = tabItem(tabName = 'converter_3',
                    fluidRow(
                      column(width = 12,
                             h1('Converter Call TSachet'),
                             h4("Silakan upload dua files Anda:"),
                             fileInput('file_call_3', 'File Call (TSachet)',
                                       accept = c('xlsx')
                             ),
                             br(),
                             fileInput('file_mtr_3', 'File Master Item',
                                       accept = c('xlsx')
                             ),
                             downloadButton("downloadData_3", "Download")
                      )
                    )
)
# body
body = dashboardBody(tabItems(readme,convert_1,convert_2,convert_3))

# ui all
ui = secure_app(dashboardPage(skin = "green",header,sidebar,body))
# ==============================================================================


# ==============================================================================
# SERVER PART
# Define server logic required to do a lot of things
server <- function(input,output,session){
  # credential untuk masalah login
  res_auth = secure_server(check_credentials = check_credentials(credentials))
  
  # konverter dimulai dari sini
  data_upload <- reactive({
    # tahap pertama adalah mengambil data yang diupload
    inFile <- input$target_upload
    if (is.null(inFile))
      return(NULL)
    
    # kita ambil nama kota dulu
    nama_kota_save = input$judul_area
    
    # baca data
    df <- read_excel(inFile$datapath) %>% janitor::clean_names() 
    
    # =================================
    # mulai paste dari sini
    
    
    # =================================
    # akhir paste di sini
    return(df_final)
    
  })
  
  data = data_upload
  
  output$downloadData <- downloadHandler(
    filename = function() {
      paste("MSA Jotform ", nama_kota_save(),Sys.time(), ".xlsx", sep=" ")
    },
    content = function(file) {
      openxlsx::write.xlsx(data(), file)
    })
  
  
}
# ==============================================================================

# ==============================================================================
# Run the application 
shinyApp(ui = ui, server = server)
# alhamdulillah selesai
# ==============================================================================

