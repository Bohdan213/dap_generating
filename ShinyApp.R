library(shiny)
library(openxlsx)

source("logic.R")

ui <- fluidPage(
  tabsetPanel(
    tabPanel("DAP to tool converter",
             sidebarLayout(
               sidebarPanel(
                 fileInput("dap", "Choose your DAP file", accept = ".xlsx"),
                 width = 3
               ),
               mainPanel(
                 downloadButton("download_tool", "Download Tool File")
               )
             )
    ),
    tabPanel("Tool to dap converter",
             sidebarLayout(
               sidebarPanel(
                 fileInput("tool", "Choose your kobo tool", accept = ".xlsx"),
                 width = 3
               ),
               mainPanel(
                 downloadButton("download_dap", "Download DAP File")
               )
             )
    ),
    tabPanel("Check changes in the dap",
             sidebarLayout(
               sidebarPanel(
                 fileInput("old_dap_tool", "Choose your kobo tool", accept = ".xlsx"),
                 fileInput("old_dap", "Choose your old DAP file", accept = ".xlsx"),
                 width = 3
               ),
               mainPanel(
                 downloadButton("download_changes", "Download changes File")
               )
             )
    )
  )
)

server <- function(input, output) {
  process_dap_tool <- function(dap) {
    if (is.null(file))
      return(NULL)
    
    tool <- create.tool(dap$datapath, "")
    
    return(tool)
  }
  process_tool_dap <- function(tool) {
    if (is.null(file))
      return(NULL)
    
    dap <- create.dap(tool$datapath, "dap_3.xlsx", "")
    
    return(dap)
  }
  
  process_tool_old_dap <- function(tool, old_dap) {
    if (is.null(file))
      return(NULL)
    
    dap <- create.changes.dap(tool$datapath, old_dap$datapath)
    
    return(dap)
  }
  
  processed_dap <- reactive({
    process_dap_tool(input$dap)
  })
  
  processed_tool <- reactive({
    process_tool_dap(input$tool)
  })
  
  processed_tool_old_dap <- reactive({
    process_tool_old_dap(input$old_dap_tool, input$old_dap)
  })
  
  output$download_tool <- downloadHandler(
    filename = function() {
      paste("tool.xlsx", sep = "")
    },
    content = function(file) {
      openxlsx::saveWorkbook(processed_dap(), file)
    }
  )
  
  output$download_dap <- downloadHandler(
    filename = function() {
      paste("dap.xlsx", sep = "")
    },
    content = function(file) {
      openxlsx::saveWorkbook(processed_tool(), file)
    }
  )
  
  output$download_changes <- downloadHandler(
    filename = function() {
      paste("dap_changes.xlsx", sep = "")
    },
    content = function(file) {
      openxlsx::saveWorkbook(processed_tool_old_dap(), file)
    }
  )
}

shinyApp(ui, server)
