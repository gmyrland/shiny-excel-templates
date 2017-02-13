library(shiny)
library(dplyr)
library(XLConnect)

## Defaults
group_choices <- c("Pizza", "Pop", "Chips")
default_custom_text <- "default"

## Template Information
data_path <- "www/"
templates <- data_frame(
    name = c("Type A", "Type B", "Type C"),
    path = paste0(data_path, c("TemplateA.xlsx", "TemplateB.xlsx", "TemplateC.xlsx")),
    img  = paste0(data_path, c("a.png", "b.png", "c.png"))
)

## Generate spreadsheet based on form inputs
generate_spreadsheet <- function(options, ...) {
    # User choices...
    template_path <- options$template_path
    image_path <- options$image_path
    custom_text <- options$custom_text
    choices <- options$choices
    
    # Sanity checks
    if (!file.exists(template_path)) {
        return(list(status=1, message=paste("Couldn't find the template:", template_path)))
    }
    if (!file.exists(image_path)) {
        return(list(status=1, message=paste("Couldn't find the image:", image_path)))
    }
    
    # Path to write
    path <- tempfile(pattern = "file", tmpdir = tempdir(), fileext = ".xlsx")
    
    # Copy chosen template
    file.copy(template_path, path)
    
    # Edit workbook based on selected options
    wb <- loadWorkbook(path)
    ws <- "Sheet1" # Edit as req'd
    
    # Formatting
    csHlight = createCellStyle(wb, name = "highlight")
    setFillPattern(csHlight, fill = XLC$FILL.SOLID_FOREGROUND)
    setFillForegroundColor(csHlight, color = XLC$COLOR.CORNFLOWER_BLUE)
    setCellStyle(wb, sheet=ws, row=10, col=seq(8,11), cellstyle = csHlight)
    csBorder = createCellStyle(wb, name = "border")
    setBorder(csBorder, side=c("top", "bottom", "left", "right"), type=XLC$BORDER.THIN, color = XLC$COLOR.BLACK)
    #setCellStyle(wb, sheet=ws, formula="Sheet1!$H$10", cellstyle = csBorder)
    writeWorksheet(wb, "Here's some more formatting created programmatically", sheet=ws, startRow=9, startCol=8, header=FALSE)
    setCellStyle(wb, sheet=ws, row=11, col=seq(8,11), cellstyle = csBorder)

    # Image
    createName(wb, name="img", formula="Sheet1!$A$4:$B$8")
    addImage(wb, filename = image_path, name="img", originalSize = FALSE)
    writeWorksheet(wb, "Custom image based on selected template", sheet=ws, startRow=9, startCol=1, header=FALSE)
    
    # Custom text
    writeWorksheet(wb, c("This was your custom text: ", custom_text), sheet=ws, startRow = 13, startCol = 1, header=FALSE)
    
    # Multiple value selection
    writeWorksheet(wb, data.frame(Choices=choices), sheet=ws, startRow = 16, startCol = 1, header=TRUE, rownames = TRUE)
    
    # Save result
    saveWorkbook(wb)
    
    return(list(status=0, message="Success", path=path))
}

## Open Spreadsheet
open_spreadsheet <- function(path) {
    system2("xdg-open", paste0('"', path, '"'))
}

## Shiny UI
ui <- fluidPage(
    tags$title("Custom Spreadsheet Builder"),
    tags$h1("Custom Spreadsheet Builder"),
    tags$hr(),
    tags$p(
        "A demo spreadsheet builder. This app demonstrates customization",
        "by choosing among different templates, and also by applying",
        "programmatic edits to the document at runtime."
    ),
    tags$hr(),
    tags$h4("Inputs:"),
    textInput("name", "Name:", value = "User"),
    selectInput("template", "Select template type:",
                choices = templates$name, 
                selected = templates$name[1],
                selectize = TRUE, width = NULL, size = NULL),
    checkboxGroupInput("choices", "Make some choices:", choices = group_choices),
    textInput("custom", "Enter custom text:", value = default_custom_text),
    tags$hr(),
    tags$h4("Generate!"),
    tags$p("Click the Button to make your spreadsheet"),
    actionButton("generate", "the Button"),
    tags$hr(),
    textOutput("results")
)

## Shiny Server
server <- function(input, output, session) {
    output$results <- renderText("Status: Nothing to report")
    observeEvent(input$generate, {
        res <- generate_spreadsheet(
            options = list(
                template_path = templates$path[templates$name == input$template],
                image_path = templates$img[templates$name == input$template],
                custom_text = input$custom,
                choices = input$choices
            )
        )
        output$results <- renderText(paste("Status:", res$message))
        if (res$status == 0) {
            open_spreadsheet(res$path)
        }
    })
    session$onSessionEnded(function() stopApp(returnValue=NULL))
}

## Run the App!
shinyApp(ui, server)

