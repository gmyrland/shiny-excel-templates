library(shiny)
library(dplyr)

## Defaults
group_choices <- c("Pizza", "Pop", "Chips")
default_custom_text <- "default"

## Template Information
templates <- data_frame(
    name = c("Type A", "Type B", "Type C"),
    path = c("TemplateA.xlsx", "TemplateB.xlsx", "TemplateC.xlsx"),
    img  = c("a.png", "b.png", "c.png")
)

## Generate spreadsheet based on form inputs
generate_spreadsheet <- function(options, ...) {
    
    return(list(status=0, message="Success"))
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
            options = list()
        )
        output$results <- renderText(paste("Status:", res$message))
    })
    session$onSessionEnded(function() stopApp(returnValue=NULL))
}

## Run the App!
shinyApp(ui, server)

