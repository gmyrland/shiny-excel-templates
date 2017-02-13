library(shiny)

## Shiny UI
ui <- fluidPage(
    tags$title("Title"),
)

## Shiny Server
server <- function(input, output, session) {
    session$onSessionEnded(function() stopApp(returnValue=NULL))
}

shinyApp(ui, server)

