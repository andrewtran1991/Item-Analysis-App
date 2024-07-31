# Prerequisite ----
# Load the libraries
library(shiny)
library(tidyverse)
library(ips.tools)
library(htmltools)
library(htmlwidgets)
library(DT)
library(gridExtra)
library(grid)
library(purrr)
library(tcltk)


IAapp_function <- function(exam = exam,
                           form = form,
                           window = NULL, #only needed for on-demand
                           exam_start = exam_start,
                           on_demand = TRUE,
                           scored_items = NULL, # For window exam, input value for Raw_Score and IRT_Score is 0:number of scored items
                           path,
                           filenames = NULL,
                           probuild = F)
{
  # Data Prep
  if(!is.null(filenames)){
    IA_result <- openxlsx::read.xlsx(paste0(path, filenames[1], ".xlsx"))

  }else{
    IA_result <- openxlsx::read.xlsx(paste0(path,"/", form, "_InitialItemAnalysis.xlsx"))
  }

  IA_result_original <- data.table::copy(IA_result)
  IA_result <- IA_result %>%
    filter(Taker_Status=="First-Timers") %>%
    select(Item, Comments, Bucket)

  if(on_demand == T) {
    if(!is.null(filenames)){
      IA_finalstats <- openxlsx::read.xlsx(paste0(path,filenames[2], ".xlsx"))
      IA_conv <- do.call("get.scoretable", list(form = form))


    }else{
      IA_finalstats <- openxlsx::read.xlsx(paste0(path, form, "_Final_Calibration_Results.xlsx"))

    }
  } else {
    IA_conv <- cbind.data.frame(Form = form, Raw_Score = scored_items, IRT_Score = scored_items)
  }


  IA_att <- do.call("get.itemattributes", list(exam,form)) #Check to ensure item attributes were uploaded
  IA_er <- do.call("get.examresults", list(exam = exam, form = form))
  IA_resp <- do.call("get.itemresponses", list(exam = exam, form = form, filter = T))
  IA_resp <- do.call("clean.resp", list(resp = IA_resp, er = IA_er, form = form, exam_start = exam_start))

  IA_er <- filter(IA_er, Taker_Status == "First-Timers")
  IA_resp <- filter(IA_resp, Taker_Status == "First-Timers")

  #IA_att <- rename(IA_att, c(LongID = Item, Item = Short_Id))
  #IA_resp <- rename(IA_resp, c(LongID = Item, Item = Short_Id))

  retired <- read.csv(paste0(path, "/",form, "_RetiredItems.csv"))$x

  # Use crosswalk for forms built using Pro
  if(probuild){
    cw <- do.call("get.crosswalk", list(exam = exam)) %>% mutate_if(is.numeric, as.character)
    IA_result <- left_join(IA_result %>% mutate_if(is.numeric, as.character), cw %>% mutate_if(is.numeric, as.character), by = c("Item" = "Pro_Id")) %>% select(Short_Id, Bucket, Comments) %>% rename(Item = Short_Id)

    if(on_demand){
      IA_finalstats <- left_join(IA_finalstats , cw, by = c("Item" = "Pro_Id")) %>% rename(c(ProID = Item, Item = Short_Id))
    }

  }

  ui <- createUI(IA_att, IA_result)
  server <- createServer(IA_att, IA_result, IA_result_original, IA_resp, IA_er, on_demand, IA_conv, retired)

  # Run the application ----
  shinyApp(ui = ui, server = server)
}

# Create function to output the server function
createServer <- function(IA_att, IA_result, IA_result_original, IA_resp, IA_er, on_demand, IA_conv, retired){
  serverfunc <- function(input, output, session) {
    # showModal(modalDialog(
    #   title = "Attention",
    #   "Please set a folder that you want to save the output files first. Then rerun the app.",
    #   easyClose = TRUE
    # ))

    ## Filtering items update ----
    # define IA_result variable outside of observeEvent section
    ia_result <- ""
    item_type <<- ""
    item_status <<- ""

    # Item Status observer
    observeEvent(input$item_status, {
      if (input$filter_type == "Item Status") {
        updateSelectInput(session, "n_breaks", choices = unique(sort(IA_att$Item[IA_att$Status == input$item_status])), selected =IA_att$Item[IA_att$Status == input$item_status])
        resetInputs()
      }
    }, ignoreNULL = FALSE, ignoreInit = TRUE)

    # Item Type observer
    observeEvent(input$item_type, {
      if (input$filter_type == "Item Type") {
        updateSelectInput(session, "n_breaks", choices = unique(sort(IA_att$Item[IA_att$Item_Type == input$item_type])), selected =IA_att$Item[IA_att$Item_Type == input$item_type])
        resetInputs()
      }
    }, ignoreNULL = FALSE, ignoreInit = TRUE)

    # Item Status from IA observer
    observeEvent(input$item_status_IA, {
      if (input$filter_type == "Item Status from IA") {
        updateSelectInput(session, "n_breaks", choices = unique(sort(IA_result$Item[IA_result$Bucket == input$item_status_IA])), selected = IA_result$Item[IA_result$Bucket == input$item_status_IA])
        resetInputs()
      }
    }, ignoreNULL = FALSE, ignoreInit = TRUE)

    # Filter Type observer
    observeEvent(input$filter_type, {
      resetInputs()
    }, ignoreNULL = FALSE, ignoreInit = TRUE)

    resetInputs <- function() {
      updateSelectInput(session, "item_decision", selected = "") # reset default item decision to blank
      updateTextInput(session, "comment", value = "") # reset default comment to blank
    }


    ## Selection panel setup ----
    # reset item decision and comment inputs to blank when a new item is selected
    observeEvent(input$n_breaks, {
      updateSelectInput(session, "item_decision", selected = "") # reset default item decision to blank
      updateTextInput(session, "comment", value = "") # reset default comment to blank
      # Get IA IA_er for selected item
      if (input$n_breaks %in% IA_result$Item) {
        ia_result <<- IA_result %>%
          filter(Item == input$n_breaks) %>%
          select(Bucket) %>%
          pull()

      } else {
        ia_result <<- ""
      }
      updateTextInput(session, "ia_result", value = ia_result)

      # Get item type for selected item
      if (input$n_breaks %in%unique(IA_att$Item)) {
        item_type <<- IA_att %>%
          distinct(Item, .keep_all = T) %>%
          filter(Item == input$n_breaks) %>%
          select(Item_Type) %>%
          pull()

      } else {
        item_type <<- ""
      }
      updateTextInput(session, "item_type", value = item_type)

      # Get Domain for selected item
      if (input$n_breaks %in%IA_att$Item) {
        domain <<- IA_att %>%
          distinct(Item, .keep_all = T) %>%
          filter(Item == input$n_breaks) %>%
          select(Domain) %>%
          pull()

      }

      # Get Max_Score for selected item
      if (input$n_breaks %in%IA_att$Item) {
        max_score <<- IA_att %>%
          distinct(Item, .keep_all = T) %>%
          filter(Item == input$n_breaks) %>%
          select(Max_Score) %>%
          pull()

      }


      # Get item status for selected item
      if (input$n_breaks %in% IA_att$Item) {
        item_status <<- IA_att %>%
          distinct(Item, .keep_all = T) %>%
          filter(Item == input$n_breaks) %>%
          select(Status) %>%
          pull()

      } else {
        item_status <<- ""
      }
      updateTextInput(session, "item_status", value = item_status)


    })

    ## Main reactive df ----
    # create reactive dataframe to store selected items and their item status
    item_data <- reactiveValues(data = data.frame(Item = character(), Item_status = character(),  Domain = character(), Max_Score = character(), Item_type = character(), IA_result = character(), Decision = character(), Comment = character()))

    # Define reactive expression for summary table
    summary_data <- reactive({
      Selected_item <- data.frame(IA_attribute = "Item_Count", Selected_items_N = as.integer(nrow(item_data$data)))
      Selected_item_type <- item_data$data %>% group_by(Item_type) %>% summarise(Selected_items_N = as.integer(n())) %>% ungroup() %>% rename(IA_attribute = Item_type)
      Selected_item_status <- item_data$data %>% group_by(Item_status) %>% summarise(Selected_items_N = as.integer(n())) %>% ungroup() %>% rename(IA_attribute = Item_status)
      Selected_domain <- item_data$data %>% group_by(Domain) %>% summarise(Selected_items_N = as.integer(n())) %>% ungroup() %>% rename(IA_attribute = Domain) %>% mutate(IA_attribute = paste0("Domain_", IA_attribute))
      Selected_max_score <- item_data$data %>% group_by(Max_Score) %>% summarise(Selected_items_N = as.integer(n())) %>% ungroup() %>% rename(IA_attribute = Max_Score) %>% mutate(IA_attribute = paste0("Max_Score_", IA_attribute))
      Selected_IA_status <- item_data$data %>% group_by(IA_result) %>% summarise(Selected_items_N = as.integer(n())) %>% ungroup() %>% rename(IA_attribute = IA_result) %>% mutate(IA_attribute = paste0("IA_Status_", IA_attribute))
      Selected_Decision_status <- item_data$data %>% group_by(Decision) %>% summarise(Selected_items_N = as.integer(n())) %>% ungroup() %>% rename(IA_attribute = Decision) %>% mutate(IA_attribute = paste0("Decision_", IA_attribute))

      rbind(Selected_item, Selected_item_type, Selected_domain, Selected_max_score, Selected_item_status, Selected_IA_status, Selected_Decision_status)
    })

    ### Add items to the df ----
    # add selected item and item status to reactive dataframe when add button is clicked
    observeEvent(input$add_button, {

      # Check if item already exists in the table
      if (input$n_breaks %in% item_data$data$Item) {

        # If item already has a decision, replace it with the new decision
        item_data$data$Decision[item_data$data$Item == input$n_breaks] <- input$item_decision
        # If item already has a comment, replace it with the new comment
        item_data$data$Comment[item_data$data$Item == input$n_breaks] <- input$comment

      } else {

        # Add new row to reactive dataframe if item has no previous decision
        data <- tibble(Item = input$n_breaks, Item_status = item_status, Domain = as.character(domain), Max_Score = as.character(max_score), Item_type = item_type, IA_result = ia_result, Decision = input$item_decision, Comment = input$comment)
        item_data$data <- bind_rows(item_data$data, data)

      }

    })

    ### Go to the next item ----
    observeEvent(input$next_button, {
      # Get the currently selected item
      currentItem <- input$n_breaks

      # Get the list of choices depending on the filter type
      if (input$filter_type == "Item Status") {
        items <- sort(IA_att$Item[IA_att$Status == input$item_status])
      } else if (input$filter_type == "Item Type") {
        items <- sort(IA_att$Item[IA_att$Item_Type == input$item_type])
      } else if (input$filter_type == "Item Status from IA") {
        items <- sort(IA_result$Item[IA_result$Bucket == input$item_status_IA])
      }

      # Find the next item in the list
      nextItemIndex <- match(currentItem, items) + 1
      if (nextItemIndex > length(items)) {
        nextItemIndex <- 1  # loop back to the first item if we've reached the end
      }

      # Update the select input
      updateSelectInput(session, "n_breaks", selected = items[nextItemIndex])
    })


    ### Decision and Summary tables ----


    # render item decision table
    output$item_decision_table <- renderDataTable({
      # filter item_data based on selected item status
      item_data_filtered <- item_data$data %>%
        select(Item, Item_status, Domain, Max_Score, Item_type, IA_result, Decision, Comment) %>%
        datatable(escape = FALSE, options = list(dom = 't', paging = FALSE, ordering = FALSE))

    })

    # render item decision text
    output$item_decision_text <- renderText({
      # filter item_data based on selected item status and item name
      item_data_filtered <- item_data$data %>%
        filter(ifelse(input$item_status == "", TRUE, Item_status == input$item_status),
               Item == input$n_breaks)

      # extract the decision for the selected item
      decision <- item_data_filtered$Decision

      # if the decision is empty, display a message indicating that no decision has been made
      if (is.null(decision) || length(decision) == 0 || decision == "") {
        return("No decision has been made for this item.")
      }

      # otherwise, display the decision
      return(paste("The decision for", input$n_breaks, "is:", decision))
    })

    # summary table
    output$summary_table <- renderDataTable({
      summary_data() %>%
        select(IA_attribute, Selected_items_N) %>%
        datatable(escape = FALSE, options = list(dom = 't', paging = FALSE, ordering = FALSE))
    })

    ### Display IA_att ----
    # show domain and item type of the current item
    observeEvent(input$n_breaks, {
      output$IA_att_text <- renderText({
        paste("Domain: ", as.character(domain), " . Item Type: ", item_type)
      })
    })

    ### Display IRTb option ----
    # show whether the exam has IRT_b or not
    output$IRTb_text <- renderText({
      if (on_demand == TRUE) {
        paste("This exam has IRT_b values.")
      } else {
        "This exam has no IRT_b value."
      }
    })



    ## Plots ----
    ### Item characteristic curve plot (IA) ----
    # Update this in your server code to account for polytomous items
    output$iaplot_output <- renderUI({
      if (input$item_type != "PACS" | max_score == "1") {
        plotOutput("iaplot", width = "600", height = "500")
      } else {
        HTML("<div style='width: 500px; height: 400px; display: flex; align-items: center; justify-content: center; border: 1px solid #ddd;'><h4>This is a polytomous item.</h4></div>")
      }
    })

    output$iaplot <- renderPlot({
      if (input$item_type != "PACS" | max_score == "1") {
        iaplots2(IA_resp, IA_er, IA_att, IA_conv,
                 item = as.character(input$n_breaks),
                 score = as.character(input$score_type),
                 grouped = input$group_by,
                 byform = input$by_form,
                 scored = input$scored)
      }
    })


    ### IRT plot ----
    irtplot_output <- reactive({
      # plot the item IA_response theory plot (for on-demand exams only)
      if (on_demand == TRUE) {
        irt.fitplot2(IA_resp, IA_att, IA_finalstats, IA_er, as.character(input$n_breaks), IA_conv, input$bin_adjust)
      } else {
        "This exam has no IRT_b plot."
      }
    })

    output$irt_plot <- renderPlot({
      if (is.character(irtplot_output())) {
        # if the output is text, display it as a text output
        output$text_irtplot <- renderText(irtplot_output())
        textOutput("text_irtplot")
      } else {
        # if the output is a plot, display it as a plot output
        irtplot_output()
      }
    })

    ## Downloads ----
    ### Download the item status table ----
    output$download_item_status <- downloadHandler(
      filename = paste(form, "Table_output_", Sys.Date(), ".csv", sep = ""),
      content = function(file) {
        write.csv(item_data$data, file, row.names = F)
      }
    )


    ### Download the graphs ----
    output$pdf_button <- downloadHandler(
      filename = paste(form, "Graph_output_", Sys.Date(), ".pdf", sep = ""),
      content = function(file) {
        # create PDF device
        pdf(file)

        # loop over all items in the decision table
        for (i in 1:nrow(item_data$data)) {
          item_type <-IA_att$Item_Type[IA_att$Item == item_data$data$Item[i]]
          if (item_type != "PACS" | max_score == "1") {
            graph1 <- iaplots2(IA_resp, IA_er, IA_att, IA_conv,
                               item = item_data$data$Item[i],
                               score = as.character(input$score_type),
                               grouped = input$group_by,
                               byform = input$by_form,
                               scored = input$scored)
          } else {
            graph1 <- ggplot() +
              theme_void() +
              theme(plot.margin = margin(5.5, 5.5, 5.5, 5.5, "points")) +
              annotation_custom(
                grob = textGrob(
                  "This is a polytomous item. There is no IA plot",
                  gp = gpar(fontsize = 14, col = "black")
                ),
                xmin = -Inf, xmax = Inf, ymin = -Inf, ymax = Inf
              ) +
              coord_cartesian(xlim = c(0, 1), ylim = c(0, 1))

          }

          # print plot to PDF
          print(graph1)

          if (on_demand == TRUE) {
            # plot the item IA_response theory plot (for on-demand exams only)
            graph2 <- irt.fitplot2(IA_resp, IA_att, IA_finalstats, IA_er, as.character(item_data$data$Item[i]), IA_conv, input$bin_adjust)

            # print plot to PDF
            print(graph2)
          }
        }

        dev.off()


      }
    )

    # Generate a new IA file and create a Content Review workbook
    #observeEvent(input$generate, {
      #setwd(tcltk::tk_choose.dir())

    output$generate <- downloadHandler(
      filename = function() {
        paste0(form, "_ItemAnalysis_Final", ".xlsx")
      },
      content = function(file) {

        result_IAapp <- item_data$data %>%
          dplyr::select(Item, Decision, Comment) %>%
          mutate(Item = as.character(Item))

        IA_result_original <- IA_result_original %>% mutate(Item = as.character(Item))

        result_updated <- IA_result_original %>%
          left_join(result_IAapp, by = "Item") %>%
          mutate(Bucket2 = ifelse(!is.na(Bucket) & !is.na(Decision), Decision, Bucket)) %>%
          select(-Decision, -Bucket) %>%
          rename(Bucket = Bucket2) %>%
          rename(Psychometrician_Comment = Comment)

        # Write the updated data to the specified file path
        openxlsx::write.xlsx(result_updated, file)

        # Generate IA workbook
        ia.workbook2(result_updated[which(!result_updated$Item %in% retired),], form = form)

        # Show a message when the process is finished
        showNotification("Workbook generated successfully!", type = "message")
      }
    )


  }
  return(serverfunc)
}

createUI <- function(IA_att, IA_result){
  ## Selection panel ----
  ui <- fluidPage(
    titlePanel("IA Content Review Tool"),

    sidebarLayout(
      sidebarPanel(
        selectInput("exam", label = "Exam",
                    choices = c("", unique(IA_att$Exam)), selected = ""),

        ### Item filter options ----
        radioButtons("filter_type", label = "Filter By:",
                     choices = c("Item Status", "Item Type", "Item Status from IA"), selected = "Item Status"),

        conditionalPanel(condition = "input.filter_type == 'Item Status'",
                         selectInput("item_status", label = "Item Status",
                                     choices = c(unique(IA_att$Status)), selected = "Operational")
        ),

        conditionalPanel(condition = "input.filter_type == 'Item Type'",
                         selectInput("item_type", label = "Item Type",
                                     choices = c(unique(IA_att$Item_Type)), selected = "SSMC")
        ),

        conditionalPanel(condition = "input.filter_type == 'Item Status from IA'",
                         selectInput("item_status_IA", label = "Item Status from IA",
                                     choices = c(unique(IA_result$Bucket)), selected = "Become/Remain Operational")
        ),


        selectInput("n_breaks", label = "Item:",
                    choices = sort(unique(IA_att$Item)), selected = unique(IA_att$Item)[1]),

        ### Parameter options ----
        # allow the user to select the score type
        radioButtons("score_type", label = "Score:",
                     choices = c("raw", "base"), selected = "raw"),

        # allow the user to select whether to group by form and score type
        checkboxInput("group_by", label = "Grouped", value = TRUE),
        checkboxInput("by_form", label = "By Form", value = FALSE),
        checkboxInput("scored", label = "Scored", value = FALSE),

        # IRT scale adjustment
        conditionalPanel(condition = "on_demand == TRUE",
                         sliderInput("bin_adjust", "Bin Width (IRT)", min = 0.1, max = 1, value = 0.5, step = 0.1)),

        ### Decision and Comments ----
        selectInput("item_decision", label = "Review Decision:",
                    choices = c("", "Become/Remain Operational", "Needs Content Review", "Rework/Retired"), selected = ""),

        textInput("comment", label = "Comments:", value = ""),

        ### Add items and Download ----
        actionButton("add_button", "Add to the list"),
        actionButton("next_button", "Next item"),
        downloadButton("download_item_status", "Download Item Review Decision Table"),
        downloadButton("pdf_button", "Export the graphs and Outputs"),
        #actionButton("generate", "Generate new IA and Content Review Workbook")
        downloadButton("generate", "Generate new IA and Content Review Workbook")

      ),
      ## Graphs and Table Display ----
      mainPanel(
        tabsetPanel(
          tabPanel("IA Plot",
                   uiOutput("iaplot_output", width = "500", height = "1500")),
          tabPanel("IRT Plot",
                   plotOutput("irt_plot", width = "500", height = "400"))
        ),
        textOutput("IA_att_text"),
        textOutput("IRTb_text"),
        verbatimTextOutput("item_decision_text"),
        DT::dataTableOutput("item_decision_table"),
        h3("Summary Table"),
        DT::dataTableOutput("summary_table")
      )
    )
  )

  return(ui)
}




# Plots for the distractor analysis
iaplots2 <- function(resp, result, att, conv, item, score = "base", grouped, byform = F, scored = F){

  scorescale <- ifelse(score == "base", "Base_Form_Scale_Score", "Raw_Score")

  #If scoring hasn't happened yet, create the raw score from the response table.
  if(scored == F){
    resptemp <- resp %>% group_by(Exam_Result_Id) %>% mutate(Raw_Score = sum(Item_Score, na.rm = T)) %>% as.data.frame() %>% select( Exam_Result_Id, Raw_Score) %>% unique()
    result <- result %>% select(-Raw_Score) %>% left_join(., resptemp, by = "Exam_Result_Id")

  }

  # Create score and response data frames in wide format (edited to remove records without responses for an item)
  itemscore <- resp %>% filter(Item == item) %>% select(Exam_Result_Id, Item_Score, Form,  Item) %>%  tidyr::spread(., Item, Item_Score) %>% left_join(., dplyr::select(result %>% filter(Taker_Status == "First-Timers"), Exam_Result_Id, Raw_Score), by = "Exam_Result_Id")
  itemresp <-  resp %>% filter(Item == item) %>% select(Exam_Result_Id, Candidate_Response, Form,  Item) %>% tidyr::spread(., Item, Candidate_Response) %>% left_join(., dplyr::select(result, Exam_Result_Id), by = "Exam_Result_Id")

  #Change exam result id to character
  itemscore$Exam_Result_Id <- as.character(itemscore$Exam_Result_Id)

  #Add base form scale score to itemscore
  if(scorescale == "Base_Form_Scale_Score"){
    itemscore <- itemscore %>% left_join(conv, by = c("Form", "Raw_Score")) # %>% select(Exam_Result_Id, Form, Raw_Score, sym(scorescale)) #, all_of(ci)

    #Round the base form scale score
    itemscore$Base_Form_Scale_Score <- ips.tools::round2(as.numeric(itemscore$Base_Form_Scale_Score), 0)
  }


  # Remove scores of 10 or less. hopefully the exclusion rules will help with this.
  itemscore <- filter(itemscore, Raw_Score > 10)
  # Make sure the response table matches
  itemresp <- itemresp[itemresp$Exam_Result_Id %in% itemscore$Exam_Result_Id, ]


  # Create data frame with keys
  key <- unique(cbind.data.frame(Item = as.character(resp$Item), key = as.character(resp$Key)))
  correct <- key$key[which(key$Item == item)]

  #top 10% of examinees (could probably roll this all into one) #Moved to create table function. Needed here for the plots
  # The previous version used all of the examinees who responded to all items. This can be updated potentially to include the whole population.
  top10 <- quantile(as.numeric(itemscore[,scorescale]), probs = .9)
  bottom10 <- quantile(as.numeric(itemscore[,scorescale]), probs = .1)
  med <- quantile(as.numeric(itemscore[,scorescale]), probs = .5)

  tempsmooth <- gksmooth(itemscore[!is.na(itemscore[,item]),], itemresp[!is.na(itemscore[,item]),], h = .5, item = item, form = NULL, scorescale = scorescale, correct = correct)

  # If grouped == T, the x-axis is grouped into intervals with roughly 2.5% in each. Depending on the sample size and the scale, there might be varied results.
  if(grouped == T){
    tempsmoothG <- inner_join(filter(itemscore[,c("Exam_Result_Id", scorescale, item)], !is.na(itemscore[,item])), cbind.data.frame(Exam_Result_Id = itemresp$Exam_Result_Id, response = itemresp[,item]), by = "Exam_Result_Id") %>%
      mutate(respA = as.numeric(response == ifelse(correct %in% as.character(1:4), "1", "A")), respB = as.numeric(response == ifelse(correct %in% as.character(1:4), "2", "B")), respC = as.numeric(response == ifelse(correct %in% as.character(1:4), "3", "C")), respD = as.numeric(response == ifelse(correct %in% as.character(1:4), "4", "D"))) %>%
      group_by( v1 = cut(!!sym(scorescale), breaks=c(unique(quantile(itemscore[,scorescale], probs = seq(0, 1, 0.025), type = 4)))))  %>%
      summarize(n = n(), meanA = mean(respA), meanB = mean(respB), meanC = mean(respC), meanD = mean(respD), mp = mean(!!sym(scorescale)))



    p <- ggplot(tempsmoothG, aes(x = mp, y = get(paste0("mean", ifelse(correct %in% as.character(1:4), c("A", "B", "C", "D")[as.numeric(correct)], correct))))) +
      geom_point()

  }else{
    p <- ggplot(tempsmooth, aes(x = !!sym(scorescale), y = get(paste0("mean", ifelse(correct %in% as.character(1:4), c("A", "B", "C", "D")[as.numeric(correct)], correct))))) +
      geom_point()
  }

  fig <- p + geom_line(data = tempsmooth, mapping = aes(x = !!sym(scorescale), y = gkerA, color = "blue") ) +
    geom_line(data = tempsmooth, mapping = aes(x = !!sym(scorescale), y = gkerB, color = "red" )) +
    geom_line(data = tempsmooth, mapping = aes(x = !!sym(scorescale), y = gkerC, color = "green") ) +
    geom_line(data = tempsmooth, mapping = aes(x = !!sym(scorescale), y = gkerD, color = "black") ) +
    ylim(0,1) +
    theme_bw() +
    ggtitle(paste0("Item - ", item)) +
    xlab(gsub("_", " ", scorescale)) +
    ylab("Response Proportion") +
    scale_color_identity(name = "Answer Options",
                         breaks = c("blue", "red", "green", "black"),
                         #labels = if(correct  %in% as.character(1:4)){1:4}else{c("A", "B", "C", "D")},
                         labels = sapply(if(correct  %in% as.character(1:4)){1:4}else{c("A", "B", "C", "D")}, function(x){if(x==correct){paste0(x, "*")}else{x}}, USE.NAMES = F),
                         guide = "legend")+
    geom_vline(xintercept = top10, linetype = 2, alpha = .5) +
    geom_vline(xintercept = bottom10, linetype = 2, alpha = .5) +
    geom_vline(xintercept = med, linetype = 2, alpha = .5) +
    theme(legend.position="bottom")

  # Create table with summary statistics ####

  #Final part #Can it be combined with below?
  tempdat <- left_join(itemscore[,c("Exam_Result_Id", scorescale)], itemresp[,c("Exam_Result_Id", item)], by = "Exam_Result_Id") %>% rename(item = )

  # Calculater distractor point biserials
  pbis <- inner_join(itemscore[,c("Exam_Result_Id", scorescale, item)], cbind.data.frame(Exam_Result_Id = itemresp$Exam_Result_Id, response = itemresp[,item]), by = "Exam_Result_Id") %>%
    #mutate(respA = as.numeric(response == "A"), respB = as.numeric(response == "B"), respC = as.numeric(response == "C"), respD = as.numeric(response == "D"))%>%
    mutate(respA = as.numeric(response == ifelse(correct %in% as.character(1:4), "1", "A")), respB = as.numeric(response == ifelse(correct %in% as.character(1:4), "2", "B")), respC = as.numeric(response == ifelse(correct %in% as.character(1:4), "3", "C")), respD = as.numeric(response == ifelse(correct %in% as.character(1:4), "4", "D"))) %>%
    select(!!(all_of(scorescale)), respA, respB, respC, respD) %>%
    drop_na %>%
    #pbis %>% filter(any(is.na(respA), is.na(respB), is.na(respC), is.na(respD)))

    cor %>%
    data.frame() %>%
    mutate(Option = if(correct %in% as.character(1:4)){c("x", 1:4)}else{toupper(c("x", letters[1:4]))}) %>%
    filter(!is.na(!!sym(scorescale))) %>%
    rename(Pbis = scorescale) %>%
    select(Pbis, Option) %>%
    filter(Pbis != 1) %>%
    #select(., -!!sym(c(scorescale, "respA", "respB", "respC", "respD", "Option")[which(apply(., 2, function(x)any(is.na(x))))  ]))  %>%
    mutate(across(.cols = Pbis, .fns = ~ips.tools::round2(.x, digits = 2)))

  # mutate(Option = substr(rownames(.), 5,5))


  sumtable <- tempdat %>% group_by(!!sym(item)) %>%
    summarize(n = n(), Prop = round(n/sum(nrow(tempdat)), 3), Mean = round(mean(!!sym(scorescale)), 2), Top10 = round(sum(!!sym(scorescale) >= top10)/(sum(tempdat[,scorescale] >= top10)), 2), Bottom10 = round(sum(!!sym(scorescale) < bottom10)/(sum(tempdat[,scorescale] < bottom10)), 2)) %>%
    data.frame() %>%
    rename(Option = paste0("X", item)) %>%
    left_join(pbis, by = "Option") %>%
    relocate(Pbis, .after = Mean)

  sumtable <- ggpubr::ggtexttable(sumtable, rows = NULL, theme = ggpubr::ttheme("light", base_size = 15) ) #%>% tab_add_title(., "a;kjh")
  #, tbody.style = tbody_style(size = 15)

  fig2 <- fig+ ggpubr::theme_classic2() + ggpubr::labs_pubr()
  #sumtable2 <- ggtexttable(sumtable, rows = NULL, theme = ttheme("light", base_size = 9))


  if(byform){

    fig2 <- fig2 + geom_line(data = purrr::map2_df(split(itemscore, itemscore$Form), split(itemresp, itemresp$Form), gksmooth, item = item, scorescale = scorescale, correct = correct),
                             mapping = aes(x = !!sym(scorescale),
                                           y = !!sym(paste0("gker", if(correct  %in% as.character(1:4)){c("A", "B", "C", "D")[as.numeric(correct)]}else{c("A", "B", "C", "D")[which(ifelse(correct %in% as.character(1:4), c("A", "B", "C", "D")[as.numeric(correct)], correct) == c("A", "B", "C", "D"))]})),
                                           linetype = Form,
                                           color = c("Blue", "Red", "Green", "Black")[which(ifelse(correct %in% as.character(1:4), c("A", "B", "C", "D")[as.numeric(correct)], correct) == c("A", "B", "C", "D"))]  )) +
      scale_linetype_manual(values=c("dotdash", "dotted"))

    tablebyform <- map2_df(split(itemscore, itemscore$Form), split(itemresp, itemresp$Form), createtable, item = item, correct = correct, scorescale = scorescale) %>%
      map(ggpubr::ggtexttable, theme = ggpubr::ttheme("light", base_size = 15),  rows = NULL) %>%
      map2(as.list(names(.)), function(x,y) x %>% ggpubr::tab_add_title(text = y))

    # Arrange tablebyform[[1]] and tablebyform[[2]] side by side
    tablebyform_side <- ggpubr::ggarrange(tablebyform[[1]], tablebyform[[2]], ncol = 2)

    # Arrange fig2 on top, sumtable in the middle and tablebyform_side at the bottom
    return(
      ggpubr::ggarrange(fig2, sumtable, tablebyform_side, nrow = 3, heights = c(2, 1, 1))
    )
  }else{
    ggpubr::ggarrange(fig2, sumtable, nrow =2, heights =  c(4, 2))
  }


}


##############################################

# Required for the distractor analysis
gksmooth <- function(itemscore, itemresp, h = .5, item, form = NULL, scorescale, correct ){

  #Subset only a particular form
  if(!is.null(form)){
    itemscore <- itemscore %>% filter(Form == form)
    itemresp <- itemresp %>% filter(Form == form)
  }

  distmean <- mean(itemscore[,scorescale])
  distsd <- sd(itemscore[,scorescale])


  dat <- inner_join(itemscore[,c("Exam_Result_Id", scorescale, item)], cbind.data.frame(Exam_Result_Id = itemresp[, c("Exam_Result_Id")] , response = itemresp[, item]), by = "Exam_Result_Id") %>%
    mutate(respA = as.numeric(response == ifelse(correct %in% as.character(1:4), "1", "A")), respB = as.numeric(response == ifelse(correct %in% as.character(1:4), "2", "B")), respC = as.numeric(response == ifelse(correct %in% as.character(1:4), "3", "C")), respD = as.numeric(response == ifelse(correct %in% as.character(1:4), "4", "D")), respOmit = as.numeric(is.na(response))) %>%
    group_by(!!sym(scorescale))  %>%
    summarize(n = n(), meanA = mean(respA), meanB = mean(respB), meanC = mean(respC), meanD = mean(respD), meanOmit = mean(respOmit), mp = mean(!!sym(scorescale))) %>%
    mutate_at(paste0("mean", c("A", "B", "C", "D")), ~replace(., is.na(.), 0)) %>%
    mutate(Form = paste0(unique(itemscore$Form), collapse = ","))


  dat$gkerA <- NA
  dat$gkerB <- NA
  dat$gkerC <- NA
  dat$gkerD <- NA

  for(i in 1:nrow(dat)){
    xk <- as.numeric(dat[i, scorescale])

    dat$gkerA[i] <- sum(exp(-1/(2*h)*((itemscore[,scorescale] - xk)/distsd)^2) * as.numeric(itemresp[, item]==ifelse(correct %in% as.character(1:4), "1", "A")), na.rm = T   )/sum(exp(-1/(2*h)*((itemscore[,scorescale] - xk)/distsd)^2), na.rm = T )
    dat$gkerB[i] <- sum(exp(-1/(2*h)*((itemscore[,scorescale] - xk)/distsd)^2) * as.numeric(itemresp[, item]==ifelse(correct %in% as.character(1:4), "2", "B")), na.rm = T   )/sum(exp(-1/(2*h)*((itemscore[,scorescale] - xk)/distsd)^2), na.rm = T )
    dat$gkerC[i] <- sum(exp(-1/(2*h)*((itemscore[,scorescale] - xk)/distsd)^2) * as.numeric(itemresp[, item]==ifelse(correct %in% as.character(1:4), "3", "C")), na.rm = T   )/sum(exp(-1/(2*h)*((itemscore[,scorescale] - xk)/distsd)^2), na.rm = T )
    dat$gkerD[i] <- sum(exp(-1/(2*h)*((itemscore[,scorescale] - xk)/distsd)^2) * as.numeric(itemresp[, item]==ifelse(correct %in% as.character(1:4), "4", "D")), na.rm = T   )/sum(exp(-1/(2*h)*((itemscore[,scorescale] - xk)/distsd)^2), na.rm = T )
  }
  return(dat)
}

##############################################

#need "scorescale" variable in this and other functions?
createtable <- function(itemscore, itemresp, item, correct, addFormName = F, scorescale){
  pbis <- inner_join(itemscore[,c("Exam_Result_Id", scorescale, item, "Form")], cbind.data.frame(Exam_Result_Id = itemresp$Exam_Result_Id, response = itemresp[,item]), by = "Exam_Result_Id") %>%
    #mutate(respA = as.numeric(response == "A"), respB = as.numeric(response == "B"), respC = as.numeric(response == "C"), respD = as.numeric(response == "D"))%>%
    mutate(respA = as.numeric(response == ifelse(correct %in% as.character(1:4), "1", "A")), respB = as.numeric(response == ifelse(correct %in% as.character(1:4), "2", "B")), respC = as.numeric(response == ifelse(correct %in% as.character(1:4), "3", "C")), respD = as.numeric(response == ifelse(correct %in% as.character(1:4), "4", "D"))) %>%
    select(!!sym(scorescale), respA, respB, respC, respD) %>%
    drop_na() %>%
    cor %>%
    data.frame() %>%
    mutate(Option = if(correct %in% as.character(1:4)){c("x", 1:4)}else{toupper(c("x", letters[1:4]))}) %>%
    filter(!is.na(!!sym(scorescale))) %>%
    rename(Pbis = scorescale) %>%
    select(Pbis, Option) %>%
    filter(Pbis != 1) %>%
    #select(., -!!sym(c(scorescale, "respA", "respB", "respC", "respD", "Option")[which(apply(., 2, function(x)any(is.na(x))))  ]))  %>%
    mutate(across(.cols = Pbis, .fns = ~round2(.x, digits = 2)))

  tempdat <- left_join(itemscore[,c("Exam_Result_Id", scorescale)], itemresp[,c("Exam_Result_Id", item)], by = "Exam_Result_Id") %>% rename(item = ) %>% drop_na()

  top10 <- quantile(as.numeric(itemscore[,scorescale]), probs = .9)
  bottom10 <- quantile(as.numeric(itemscore[,scorescale]), probs = .1)
  med <- quantile(as.numeric(itemscore[,scorescale]), probs = .5)

  sumtable <- tempdat %>% group_by(!!sym(item)) %>%
    summarize(n = n(), Prop = round(n/sum(nrow(tempdat)), 3), Mean = round(mean(!!sym(scorescale)), 2), Top10 = round(sum(!!sym(scorescale) >= top10)/(sum(tempdat[,scorescale] >= top10)), 2), Bottom10 = round(sum(!!sym(scorescale) < bottom10)/(sum(tempdat[,scorescale] < bottom10)), 2)) %>%
    data.frame() %>%
    rename(Option = paste0("X", item)) %>%
    left_join(pbis, by = "Option") %>%
    relocate(Pbis, .after = Mean)

  return(sumtable)
}


##############################################


# IRT function
irt.fitplot2 <- function(resp, att, irtstats, result, item, scoreTable, binwidth){


  #### Data Preparation ####

  #Specify the form
  #form <- unique(resp$Form)
  form <- att$Form[which(att$Item == item)]

  #Read in conversion table
  #scoreTable <- do.call("get.scoretable", list(form = form))
  # which form is the item on.

  #Subset only the responses for first-time test-takers.
  resp <- resp[if(length(form) == 1){resp$Form==form}else{TRUE} & resp$Taker_Status=="First-Timers",]

  #Rearrange response data
  itemscore <- tidyr::spread(resp[, c("Exam_Result_Id", "Item_Score",  "Item")], Item, Item_Score)

  #Assign each person a theta score.
  conversion <- scoreTable[if(length(form) >1){1:nrow(scoreTable)}else{scoreTable$Form==form}, c("Form", "Raw_Score", "IRT_Score")]
  #newdat <- cbind(result[, c("Exam_Result_Id", "Raw_Score")], Theta=sapply(result$Raw_Score,  assigntheta, conversion=conversion))
  newdat <- left_join(result[, c("Exam_Result_Id", "Raw_Score", "Form")], conversion, by = c("Form", "Raw_Score")) %>% rename(Theta = IRT_Score)

  #Create final data set with ID, raw score, theta, and examinee responses.
  finaldat <- merge(newdat, itemscore, by="Exam_Result_Id", all.y=T)

  #Select the percentile ranks for the vertical lines. These lines provide additional information about the distribution of test-takers.
  #Percentile rank is defined as the proportion of examinees below a certain score.
  pr <- c(.10, .50, .90)
  #ppoint <- conversion[which( scoreTable$Raw_Score %in% round(quantile(as.numeric(finaldat$Raw_Score), pr))),"IRT_Score"]
  ppoint <- quantile(as.numeric(finaldat$Theta), pr)


  #### Scored items ####

  #Settings for the plots (e.g., four plots per page, margin dimensions, etc.)
  #par(mfrow=c(2,2))
  par(mar=c(5,3,2,4))
  par(mgp=c(1.5,.5,0))
  par(oma=c(1,1,1,1))


  #finaldat$range <- cut(finaldat$Theta, seq(-1, 4, by = binwidth))
  finaldat$range <- cut(finaldat$Theta, seq(round2(min(finaldat$Theta), 0), ceiling(max(finaldat$Theta)), by = binwidth))

  if(att$Item_Type[which(att$Item == item)][1] %in% c("SSMC", "HotSpot")){
    #Aggregate the item scores by theta. For scored items, examinee theta scores are rounded to the nearest tenth, then grouped together.
    #y<-aggregate(finaldat[, item] , by=list(round2(finaldat$Theta,1)), mean, na.rm = T)
    y<-aggregate(finaldat[, c("Theta", item)] , by=list(finaldat$range), mean, na.rm = T)


    #The b parameter for the Rasch model so the model-based probability can be calculated.
    b<-irtstats$IRT_b[which(irtstats$Item==item)]

    #Generate the plot
    plot(y[,-1], las=1, pch=19, ylim=c(0,1), main=paste(form, "-",  "Item", item), ylab=expression(paste("P(", theta, ")")), xlab=expression(theta), cex.axis=.75)  #, xlim=c(min(y[,1])-.5)

    #Add lines for the model-based item characteristic curve
    lines(y[,"Theta"], probrasch(y[,"Theta"], b), col="red")

    #Calculate the number of examinees at each theta level. This is used for the confidence interval calculation.
    #n<-aggregate(as.numeric(finaldat[, item]), by=list(round2(finaldat$Theta,1)), length)
    n<-aggregate(as.numeric(finaldat[, item]), by=list(finaldat$range), length)

    #Create the lower bound and the upper bound for a 90% confidence interval.
    lb<-(y[, item]-1.645*sqrt(y[, item]*(1-y[, item])/n[,2]))
    ub<-(y[, item]+1.645*sqrt(y[, item]*(1-y[, item])/n[,2]))

    #Creates the error bands on the plot. In R's normal plot function, there is no way to add error bands.
    #Arrows with flat ends are drawn from the observed data point to the upper and lower bounds to look like normal error bands.
    #There are warnings when the frequency is one and there is no error band.
    suppressWarnings(arrows(y[,"Theta"], lb, y[,"Theta"], ub, length=0.05, angle=90, code=3))


    #Adds dashed vertical grey lines to the plot to denote the theta values associated with different percentile ranks.
    abline(v=ppoint, lty=2, lwd=.5, col="grey")

    #Create the item data to be displayed to the right of the plot. Pvalue and point biserial are calculated here. Other values are taken from Winsteps Output.
    istats<-rbind.data.frame(Pvalue = round2(mean(as.numeric(finaldat[!is.na(finaldat[, item]), item])), 2), round(cor(as.numeric(finaldat[!is.na(finaldat[, item]), "Raw_Score"]), as.numeric(finaldat[!is.na(finaldat[, item]), item])), 2), round(t(data.frame(irtstats[irtstats$Item==item ,c("IRT_b", "Outfit_MSQ", "Outfit_Z", "Infit_MSQ", "Infit_Z")])), 2))

    rownames(istats)[c(2, 4, 5, 6, 7)]<-c("PBis", "O_MSQ", "O_Z", "I_MSQ", "I_Z")
    mtext(text=stringr::str_c(rownames(istats), ": ", as.matrix(istats),  collapse ="\n" ),side=4,line=0,outer=F, cex=.60, las=2)

  }


  #Polytomous items
  #Plot for each category in a loop.

  if(att$Max_Score[which(att$Item == item)][1] > 1)
  {
    par(mfrow=c(3,2))
    # for(j in 1:length(scoredpoly))
    # {

    #Get the number of categories for the item
    ncat<-sum(!is.na(irtstats[which(irtstats$Item == item),paste0("IRT_d", 1:5)]))+1   #which(pretest$Item==pretest[i])

    #Aggregate the item scores by theta. For scored items, examinee theta scores are rounded to the nearest tenth, then grouped together.
    y <- aggregate(finaldat[!is.na(finaldat[, item]), item], by=list(finaldat$range), simplify=F, function(dat, ncat){matrix(table(factor(dat, levels=0:(ncat-1)))/length(dat))}, ncat=ncat)
    y <- cbind(aggregate(finaldat[, c("Theta", item)] , by=list(finaldat$range), mean, na.rm = T)[,-c(1,3)], matrix(unlist(y[,2]), ncol=ncat, byrow=T))



    #The b and d parameters are used to calculate the model-based probability for the partial credit model.
    b<-irtstats$IRT_b[which(irtstats$Item == item)]
    d<-as.matrix(irtstats[which(irtstats$Item == item), paste0("IRT_d",1:(ncat-1)),])

    #Calculate the probability of receiving each score using the partial credit model
    modelprob<-t(sapply(y[,1], probraschPoly, b=b, d=d))

    #Calculate the number of examinees at each theta level. This is used for the confidence interval calculation.
    n<-aggregate(as.numeric(finaldat[!is.na(finaldat[, item]), item]), by=list(finaldat$range), length)

    for(i in 1:ncat)
    {

      plot(y[,c(1,i+1)], las=1, pch=19, ylim=c(0,1), main=paste(form, "-",  "Item", item, "Cat.", i), ylab=expression(paste("P(", theta, ")")), xlab=expression(theta) )
      lines(y[,1], modelprob[,i],  col="red")

      lb<-(y[,i+1]-1.645*sqrt(y[,i+1]*(1-y[,i+1])/n[,2]))
      ub<-(y[,i+1]+1.645*sqrt(y[,i+1]*(1-y[,i+1])/n[,2]))

      suppressWarnings(arrows(y[,1], lb, y[,1], ub, length=0.05, angle=90, code=3))
      abline(v=ppoint, lty=2, lwd=.5, col="grey")

      #Creates the item data to be displayed to the right of the plot. Pvalue and point biserial are calculated here. Other values are taken from Winsteps Output.
      istats<-rbind.data.frame(Pvalue = round2(mean(as.numeric(finaldat[!is.na(finaldat[, item]), item])), 2), round(cor(as.numeric(finaldat[!is.na(finaldat[, item]), "Raw_Score"]), as.numeric(finaldat[!is.na(finaldat[, item]), item])), 2), round(t(data.frame(irtstats[irtstats$Item==item ,c("IRT_b", "Outfit_MSQ", "Outfit_Z", "Infit_MSQ", "Infit_Z")])), 2))

      # Add statistics to the last figure
      if(i == ncat){
        rownames(istats)[c(2, 4, 5, 6, 7)]<-c("PBis", "O_MSQ", "O_Z", "I_MSQ", "I_Z")
        mtext(text=stringr::str_c(rownames(istats), ": ", as.matrix(istats),  collapse ="\n" ),side=4,line=0,outer=F, cex=.60, las=2)
      }


    }

  }
}


# IRT function
irt.fitplot2_ggplot <- function(resp, att, irtstats, result, item, scoreTable, binwidth){


  #### Data Preparation ####

  #Specify the form
  form<-unique(resp$Form)

  #Read in conversion table
  #scoreTable <- do.call("get.scoretable", list(form = form))

  #Subset only the responses for first-time test-takers.
  resp<-resp[resp$Form==form & resp$Taker_Status=="First-Timers",]

  #Rearrange response data
  itemscore <- tidyr::spread(resp[, c("Exam_Result_Id", "Item_Score",  "Item")], Item, Item_Score)

  #Assign each person a theta score.
  conversion<-scoreTable[scoreTable$Form==form, c("Raw_Score", "IRT_Score")]
  newdat<-cbind(result[, c("Exam_Result_Id", "Raw_Score")], Theta=sapply(result$Raw_Score,  assigntheta, conversion=conversion))

  #Create final data set with ID, raw score, theta, and examinee responses.
  finaldat<-merge(newdat, itemscore, by="Exam_Result_Id", all.y=T)

  #Select the percentile ranks for the vertical lines. These lines provide additional information about the distribution of test-takers.
  #Percentile rank is defined as the proportion of examinees below a certain score.
  pr<-c(.10, .50, .90)
  ppoint<-conversion[which( scoreTable$Raw_Score %in% round(quantile(as.numeric(finaldat$Raw_Score), pr))),"IRT_Score"]

  #### Scored items ####

  #Get the scored items from the resp data. Does this make sense? Needs adjustment if more than form is required. irt attributes table works better.
  scored<-att[att$Max_Score == 1 & att$Scored == T & att$Form == form, "Item"]

  scoredpoly <- att[att$Scored ==T &  att$Max_Score > 1 & att$Form == form, "Item" ]


  #Settings for the plots (e.g., four plots per page, margin dimensions, etc.)
  #par(mfrow=c(2,2))
  par(mar=c(5,3,2,4))
  par(mgp=c(1.5,.5,0))
  par(oma=c(1,1,1,1))


  #finaldat$range <- cut(finaldat$Theta, seq(-1, 4, by = binwidth))
  finaldat$range <- cut(finaldat$Theta, seq(round2(min(finaldat$Theta), 0), ceiling(max(finaldat$Theta)), by = binwidth))


  #Aggregate the item scores by theta. For scored items, examinee theta scores are rounded to the nearest tenth, then grouped together.
  y<-aggregate(finaldat[, item] , by=list(round2(finaldat$Theta,1)), mean, na.rm = T)
  y<-aggregate(finaldat[, c("Theta", item)] , by=list(finaldat$range), mean, na.rm = T)


  #The b parameter for the Rasch model so the model-based probability can be calculated.
  b<-irtstats$IRT_b[which(irtstats$Item==item)]

  p <- ggplot(data = y, aes(x = Theta, y = y[, item])) +
    geom_point(pch = 19) +
    geom_line(aes(y = probrasch(Theta, b)), col = "red") +
    ylim(c(0, 1)) +
    labs(title = paste(form, "-",  "Item", item),
         y = expression(paste("P(", theta, ")")),
         x = expression(theta)) +
    theme_classic() +
    theme(axis.text = element_text(size = 8),
          axis.title = element_text(size = 10),
          plot.title = element_text(size = 12, face = "bold"))

  # Calculate the number of examinees at each theta level. This is used for the confidence interval calculation.
  n <- aggregate(as.numeric(finaldat[, item]), by = list(finaldat$range), length)

  # Create the lower bound and the upper bound for a 90% confidence interval.
  lb <- y[, item] - 1.645 * sqrt(y[, item] * (1 - y[, item]) / n[, 2])
  ub <- y[, item] + 1.645 * sqrt(y[, item] * (1 - y[, item]) / n[, 2])

  # Create error bars using geom_errorbar
  p <- p + geom_errorbar(aes(ymin = lb, ymax = ub),
                         width = 0.05,
                         length = unit(0.1, "inches"),
                         angle = 90,
                         position = position_identity())

  # Adds dashed vertical grey lines to the plot to denote the theta values associated with different percentile ranks.
  p <- p + geom_vline(xintercept = ppoint, linetype = "dashed", color = "grey")

  # Create the item data to be displayed to the right of the plot
  istats <- rbind.data.frame(Pvalue = round2(mean(as.numeric(finaldat[!is.na(finaldat[, item]), item])), 2),
                             round(cor(as.numeric(finaldat[!is.na(finaldat[, item]), "Raw_Score"]),
                                       as.numeric(finaldat[!is.na(finaldat[, item]), item])), 2),
                             round(t(data.frame(irtstats[irtstats$Item == item, c("IRT_b", "Outfit_MSQ", "Outfit_Z", "Infit_MSQ", "Infit_Z")])), 2))

  rownames(istats)[c(2, 4, 5, 6, 7)] <- c("PBis", "O_MSQ", "O_Z", "I_MSQ", "I_Z")
  p <- p + labs(caption = paste(rownames(istats), ": ", as.matrix(istats),  collapse = "\n"))


  return(p)

}

# Modify ia.workbook to enable users to select directory before exporting
ia.workbook2 <- function(result, form){

  result$Item_Status <- NULL
  # Create content review with drop down value
  wb <- createWorkbook()



  itemNames <- oapi.getItems(result[which(result$Bucket != "Become/Remain Operational"),]$Short_Id)
  result <- dplyr::left_join(result, itemNames[,c("Item", "Item_Name")])
  result <- result[order(result$Item_Name),
                   c("Forms", "Item", "Short_Id", "Pro_Id", "Item_Name",
                     setdiff(names(result), c("Forms", "Item", "Short_Id", "Pro_Id", "Item_Name")))]

  if("PACS-Subitem" %in% result$Item_Type){
    rest <- subset(result,!Item_Type %in% c("PACS","PACS-Subitem"))
    restPAC <-  subset(result,Item_Type %in% c("PACS","PACS-Subitem"))
    if(nrow(restPAC) > 0){
      restPAC[which(restPAC$Item_Type=="PACS-Subitem"),]$Bucket <- NA
    }
  } else{
    rest <- result
    restPAC <- data.frame()
  }

  ## Add the first tab:Require Content Review ##
  if(c("Needs Content Review") %in% rest$Bucket){
    addWorksheet(wb, "Require Content Review")
    content_df <- subset(rest,rest$Bucket=="Needs Content Review" & rest$Official_Stat==1,select = colnames(rest) %in%
                           c("Item","Surpass_Id","Short_Id","Pro_Id","Item_Name","Forms","Domain","Item_Type","Key","Total_Count","Mean_Time_Seconds","PValue","Corrected_PBis","A_Proportion","B_Proportion","C_Proportion","D_Proportion","A_PBis","B_PBis","C_PBis","D_PBis","Comments"))
    content_df[,"Is the content of this item correct and appropriate?"] <- NA
    # Create drop-down values dataframe
    content_df[1:3,"Drop-down Values"] = c("Yes", "No, the item needs to be reworked.","No, the item needs to be retired.")

    # Add content_df dataframe to the sheet "Require Content Review"
    writeData(wb, sheet = "Require Content Review", x = content_df, startCol = 1)

    # Add drop-downs to the decision column on the worksheet "Require Content Review"
    dataValidation(wb, "Require Content Review", type = "list", cols=match("Is the content of this item correct and appropriate?",names(content_df)),
                   rows = 2:(nrow(content_df)+1),value = "$X$2:$X$4")

    #Highlight Column T
    conditionalFormatting(wb, "Require Content Review", cols = grep("Is the content of this item correct and appropriate?",names(content_df)),
                          rows = 1:(nrow(content_df)+1), rule = ">=0", style = createStyle(bgFill = "yellow"))

    #Hide Column U
    setColWidths(wb, "Require Content Review", cols = ncol(content_df), hidden = TRUE)
    setColWidths(wb, "Require Content Review", cols = ncol(content_df)-1, widths = "auto")
  }

  ## Add the second tab:Rework or Retire ##
  if(c("Rework/Retired") %in% rest$Bucket){
    addWorksheet(wb, "Rework or Retire")
    retire_df<-subset(rest,rest$Bucket=="Rework/Retired" & rest$Official_Stat==1,select = colnames(rest) %in%
                        c("Item","Surpass_Id","Short_Id","Pro_Id","Item_Name","Forms","Domain","Item_Type","Key","Total_Count","Mean_Time_Seconds","PValue","Corrected_PBis","A_Proportion","B_Proportion","C_Proportion","D_Proportion","A_PBis","B_PBis","C_PBis","D_PBis","Comments"))
    retire_df[,"Decision (Retire or Rework)"] <- NA
    # Create drop-down values dataframe
    retire_df[1:2,"Drop-down Values"] = c("Retire", "Rework")

    # Add retire_df dataframe to the sheet "Rework or Retire"
    writeData(wb, sheet = "Rework or Retire", x = retire_df, startCol = 1)

    #Highlight Column T
    conditionalFormatting(wb, "Rework or Retire", cols = grep("Decision (Retire or Rework)", fixed = TRUE, names(retire_df)),
                          rows = 1:(nrow(retire_df)+1), rule = ">=0", style = createStyle(bgFill = "yellow"))

    #Hide Column U
    setColWidths(wb, "Rework or Retire", cols = ncol(retire_df), hidden = TRUE)
    setColWidths(wb, "Rework or Retire", cols = ncol(retire_df)-1, widths = "auto")

    # Add drop-downs to the decision column on the worksheet "Retire or Rework"
    dataValidation(wb, "Rework or Retire", type = "list", cols=match("Decision (Retire or Rework)",names(retire_df)),
                   rows = 2:(nrow(retire_df)+1),value = "$X$2:$X$3")
  }



  if(nrow(restPAC) > 0){
    restPAC <- subset(restPAC,restPAC$Official_Stat==1)
    #get PACS that need content review
    if(c("Needs Content Review") %in% restPAC$Bucket){
      addWorksheet(wb, "PACS - Require Review")

      crPAC <- subset(restPAC, Bucket == "Needs Content Review")
      crPAC2 <- data.frame()
      for(i in 1:nrow(crPAC)) {
        crPAC_temp <- subset(restPAC, grepl(crPAC[i,]$Item, restPAC$Item))
        crPAC2 <- rbind(crPAC_temp,crPAC2)
      }

      crPAC2<-subset(crPAC2,select = colnames(crPAC2) %in%
                       c("Item","Surpass_Id","Short_Id","Pro_Id","Forms","Domain","Item_Type","Key","Total_Count","Mean_Time_Seconds",
                         "PValue","Corrected_PBis","A_Proportion","B_Proportion","C_Proportion","D_Proportion",
                         "A_PBis","B_PBis","C_PBis","D_PBis","Bucket","Comments"))

      crPAC2[,"Decision"] <- NA

      # Create drop-down values dataframe
      crPAC2[1:3,"Drop-down Values"] <- c("Yes", "No, the item needs to be reworked.",
                                          "No, the item needs to be retired.")

      # Add all flagged PACS items& PACS-Subitem dataframe to the sheet "PACS Review"
      writeData(wb, sheet = "PACS - Require Review", x = crPAC2, startCol = 1)

      #Highlight Column T
      conditionalFormatting(wb, "PACS - Require Review", cols = grep("Decision",names(crPAC2)),
                            rows = 1:(nrow(crPAC2)+1), rule = ">=0", style = createStyle(bgFill = "yellow"))

      #Lengthen decision and bucket columns
      setColWidths(wb, "PACS - Require Review", cols = c(ncol(crPAC2)-1,ncol(crPAC2)-2), widths = "auto")

    }
    #get PACS that should be retired
    if(c("Rework/Retired") %in% restPAC$Bucket){
      addWorksheet(wb, "PACS - Rework or Retire")

      rrPAC <- subset(restPAC, Bucket == "Rework/Retired")
      rrPAC2 <- data.frame()
      for(i in 1:nrow(rrPAC)) {
        rrPAC_temp <- subset(restPAC, grepl(rrPAC[i,]$Item, restPAC$Item))
        rrPAC2 <- rbind(rrPAC_temp,rrPAC2)
      }

      rrPAC2<-subset(rrPAC2,select = colnames(rrPAC2) %in%
                       c("Item","Forms","Domain","Item_Type","Key","Total_Count","Mean_Time_Seconds",
                         "PValue","Corrected_PBis","A_Proportion","B_Proportion","C_Proportion","D_Proportion",
                         "A_PBis","B_PBis","C_PBis","D_PBis","Bucket","Comments"))

      rrPAC2[,"Decision"] <- NA

      # Create drop-down values dataframe
      rrPAC2[1:2,"Drop-down Values"] <- c("Retire", "Rework")

      # Add all flagged PACS items& PACS-Subitem dataframe to the sheet "PACS Review"
      writeData(wb, sheet = "PACS - Rework or Retire", x = rrPAC2, startCol = 1)

      #Highlight Column T
      conditionalFormatting(wb, "PACS - Rework or Retire", cols = grep("Decision",names(rrPAC2)),
                            rows = 1:(nrow(rrPAC2)+1), rule = ">=0", style = createStyle(bgFill = "yellow"))

      #Lengthen decision and bucket columns
      setColWidths(wb, "PACS - Rework or Retire", cols = c(ncol(rrPAC2)-1,ncol(rrPAC2)-2), widths = "auto")

    }
  } # end if PACS

  # Choose directory
  # choose.dir()

  # Save workbook
  saveWorkbook(wb, sprintf("%s_Content Review.xlsx",form), overwrite = TRUE)

}


