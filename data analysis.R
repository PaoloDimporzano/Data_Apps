# Function to check and install required packages
install_if_missing <- function(packages) {
  for (package in packages) {
    if (!requireNamespace(package, quietly = TRUE)) {
      message(paste("Installing package:", package))
      install.packages(package, dependencies = TRUE)
    }
    library(package, character.only = TRUE)
  }
}

# List of required packages (including writexl for Excel export)
required_packages <- c(
  "shiny",
  "readxl",
  "dplyr",
  "DT",
  "FactoMineR",
  "factoextra",
  "colourpicker",
  "ggplot2",
  "corrplot",
  "plotly",
  "grid",
  "gridExtra",
  "reshape2",
  "cluster",
  "heatmaply",
  "dendextend",
  "viridis",
  "RColorBrewer",
  "apcluster",
  "dbscan",
  "kernlab",
  "writexl"
)

# Check and install missing packages
install_if_missing(required_packages)

message("All required packages are installed and loaded successfully!")

# Function to remove outliers per group using the 1.5*IQR rule for each selected numeric variable
removeOutliers <- function(dataset, group_col, selected_vars) {
  dataset %>%
    group_by_at(group_col) %>%
    filter(if_all(all_of(selected_vars), function(x) {
      q1 <- quantile(x, 0.25, na.rm = TRUE)
      q3 <- quantile(x, 0.75, na.rm = TRUE)
      iqr <- q3 - q1
      (x >= (q1 - 1.5 * iqr)) & (x <= (q3 + 1.5 * iqr))
    })) %>%
    ungroup()
}

# UI Definition
ui <- navbarPage("Data Analysis App",
                 
                 # Overview Tab - New first page with summary and guide for each section.
                 tabPanel("Overview",
                          fluidPage(
                            h2("Welcome to the Data Analysis App"),
                            p("This application helps you perform several important statistical analyses, including:"),
                            tags$ul(
                              tags$li(
                                strong("Data Upload & New Dataset Creation:"), 
                                " Load your data in Excel or CSV format and create a new dataset by selecting key variables. The app also offers an outlier test to filter your data based on the 1.5*IQR rule. For further details, consider reading [Tukey's Exploratory Data Analysis](https://en.wikipedia.org/wiki/Exploratory_data_analysis).",
                                p("Outlier Test Equation: A data point is considered an outlier if it falls outside the range:"),
                                tags$p("Lower Bound = Q1 - 1.5 * IQR, Upper Bound = Q3 + 1.5 * IQR"),
                                tags$p("where IQR = Q3 - Q1")
                              ),
                              tags$li(
                                strong("PCA (Principal Component Analysis):"),
                                " Reduce dimensionality by selecting variables for PCA. The app computes eigenvalues, a covariance matrix, and principal components to help you understand the underlying correlation structure.",
                                p("The covariance matrix indicates how pairs of variables vary together, while the loadings plot displays the weight of each original variable on the principal components. This orthogonal transformation reveals the directions with the maximum variance. For further reading, refer to [Jolliffe's Principal Component Analysis](https://www.wiley.com/en-us/Principal+Component+Analysis+with+Applications+in+R-p-9780470050961).")
                              ),
                              tags$li(
                                strong("Violin Plot & Box Plot:"),
                                " Visualize the distribution of your data with violin plots that can be optionally overlaid with box plots and data points.",
                                p("Box Plot Explanation:"),
                                tags$p("• The box represents the interquartile range (IQR) containing the middle 50% of the data."),
                                tags$p("• The line within the box indicates the median value."),
                                tags$p("• The 'whiskers' extend to the smallest and largest values within 1.5 * IQR from the lower and upper quartiles respectively."),
                                tags$p("• Points outside the whiskers are considered outliers."),
                                p("Below is an illustrative image for the box plot components (Attribution: [Jhguch](https://en.wikipedia.org/wiki/Box_plot)) :"),
                                tags$img(
                                  src = "https://upload.wikimedia.org/wikipedia/commons/8/89/Boxplot_vs_PDF.png", 
                                  alt = "Box Plot Explanation Diagram",
                                  style = "width:600px; height:400px;"
                                )
                              ),
                              tags$li(
                                strong("K-means Clustering:"),
                                " Group similar observations using k-means clustering. In this method, the data are first scaled so that all variables contribute equally to the distance calculations. The app provides multiple distance metrics (Euclidean, Manhattan, Maximum, Minkowski) and suggests the optimal number of clusters using the elbow method.",
                                p("K-means works by partitioning the dataset into 'k' clusters, minimizing the within-cluster sum of squares. Scaling is crucial here to ensure that variables on different scales do not disproportionately influence the clustering outcome. For more background, see [MacQueen's work on k-means clustering](https://en.wikipedia.org/wiki/K-means_clustering).")
                              )
                            ),
                            h3("Guide to Using the App"),
                            h4("1. Data Upload and Dataset Setup"),
                            p("Upload your dataset via the 'New Dataset' tab. Select the appropriate ID and grouping columns, choose the variables you want to work with, and optionally apply the outlier test to exclude anomalous data points."),
                            h4("2. PCA Analysis"),
                            p("In the 'PCA' tab, select the numeric variables to include in the analysis and choose which principal components to plot on the X and Y axes. The PCA and loadings plots will display using the same IDs as in the new dataset to ensure consistency."),
                            h4("3. Violin Plot (with Box Plot)"),
                            p("Use the 'Violin Plot' tab to visualize the distribution of a selected numeric variable across groups. Adjust the transparency and overlay a box plot to emphasize medians and quartile ranges, along with individual data points."),
                            h4("4. K-means Clustering"),
                            p("In the 'K-means Clustering' tab, select the variables to use for clustering. The app allows you to choose from various distance metrics:"),
                            tags$ul(
                              tags$li(
                                strong("Euclidean:"),
                                " Ideal when variables are on the same scale and typically used by default."
                              ),
                              tags$li(
                                strong("Manhattan:"),
                                " Useful when differences are measured in absolute terms."
                              ),
                              tags$li(
                                strong("Maximum:"),
                                " Considers the maximum difference among dimensions."
                              ),
                              tags$li(
                                strong("Minkowski:"),
                                " A flexible option that can be adjusted with a power parameter for different distance sensitivities."
                              )
                            ),
                            p("The optimal number of clusters is suggested by the elbow method, which is displayed within the app."),
                            h4("5. Complete Dataset"),
                            p("The 'Complete Dataset' tab merges the newly created dataset with PCA dimensions and k-means clustering results. This combined view uses unique IDs to ensure every sample is consistently represented across all analyses.")
                          )
                          
                 ),
                 
                 tabPanel("New Dataset",
                          sidebarLayout(
                            sidebarPanel(
                              fileInput("file", "Choose Excel or CSV File", accept = c(".xlsx", ".xls", ".csv")),
                              uiOutput("sheet_ui"),
                              numericInput("headerRow", "Row number for header", value = 1, min = 1),
                              actionButton("loadData", "Load Data"),
                              tags$hr(),
                              h4("New Dataset Setup"),
                              uiOutput("id_col_ui"),
                              uiOutput("group_col_ui"),
                              uiOutput("select_vars_ui"),
                              uiOutput("rename_vars_ui"),
                              checkboxInput("remove_outliers", "Identify and Remove Outliers", value = FALSE),
                              actionButton("createDb", "Create New Dataset")
                            ),
                            mainPanel(
                              h4("Original Data Preview"),
                              DTOutput("orig_data"),
                              tags$hr(),
                              h4("New Dataset Preview"),
                              DTOutput("new_data"),
                              downloadButton("downloadData", "Download Dataset")
                            )
                          )
                 ),
                 
                 tabPanel("PCA",
                          sidebarLayout(
                            sidebarPanel(
                              uiOutput("pca_vars_ui"),
                              uiOutput("pca_axes_ui"),
                              uiOutput("pca_group_filter_ui"),
                              uiOutput("pca_group_colors_ui"),
                              checkboxInput("show_id_labels", "Show ID Labels", value = FALSE),
                              hr(),
                              h4("Axis Formatting Options"),
                              numericInput("pca_axis_title_size", "Axis Title Size:", value = 14, min = 8),
                              numericInput("pca_axis_text_size", "Axis Text Size:", value = 12, min = 8),
                              hr(),
                              h4("Download Options"),
                              selectInput("pca_download_format", "Select Format:",
                                          choices = c("PDF" = "pdf",
                                                      "PNG" = "png",
                                                      "JPEG" = "jpeg",
                                                      "TIFF" = "tiff")),
                              numericInput("pca_download_width", "Width (inches):", 10, min = 1),
                              numericInput("pca_download_height", "Height (inches):", 12, min = 1),
                              numericInput("pca_download_dpi", "DPI:", 300, min = 72),
                              downloadButton("download_pca_plot", "Download PCA Plot"),
                              downloadButton("download_contrib_plot", "Download Contributions Plot"),
                              downloadButton("download_corr_plot", "Download Correlation Plot")
                            ),
                            mainPanel(
                              tabsetPanel(
                                tabPanel("PCA Analysis", 
                                         plotOutput("pcaCombinedPlot", height = "900px")
                                ),
                                tabPanel("Contributions", plotOutput("pcaContributions")),
                                tabPanel("Variable Correlation", plotOutput("variableCorrelation", height = "600px")),
                                tabPanel("Sample-Variable Correlation", plotlyOutput("sampleCorrelation", height = "1600px")),
                                tabPanel("PC Matrix", 
                                         DTOutput("pc_matrix"),
                                         br(),
                                         downloadButton("download_pc_matrix", "Download PC Matrix as Excel")
                                )
                              )
                            )
                          )
                 ),
                 
                 tabPanel("Violin Plot",
                          sidebarLayout(
                            sidebarPanel(
                              uiOutput("violin_var_ui"),
                              uiOutput("violin_group_filter_ui"),
                              uiOutput("violin_group_colors_ui"),
                              sliderInput("violin_alpha", "Violin Plot Transparency:", 
                                          min = 0, max = 1, value = 0.8, step = 0.1),
                              checkboxInput("show_violin", "Show Violin Plot", value = TRUE),
                              checkboxInput("show_boxplot", "Show Box Plot", value = TRUE),
                              checkboxInput("show_points", "Show Data Points", value = TRUE),
                              hr(),
                              h4("Axis Formatting Options"),
                              numericInput("violin_axis_title_size", "Axis Title Size:", value = 14, min = 8),
                              numericInput("violin_axis_text_size", "Axis Text Size:", value = 12, min = 8),
                              hr(),
                              h4("Download Options"),
                              selectInput("violin_download_format", "Select Format:",
                                          choices = c("PDF" = "pdf",
                                                      "PNG" = "png",
                                                      "JPEG" = "jpeg",
                                                      "TIFF" = "tiff")),
                              numericInput("violin_download_width", "Width (inches):", 10, min = 1),
                              numericInput("violin_download_height", "Height (inches):", 8, min = 1),
                              numericInput("violin_download_dpi", "DPI:", 300, min = 72),
                              downloadButton("download_violin_plot", "Download Violin Plot")
                            ),
                            mainPanel(
                              plotOutput("violinPlot", height = "600px")
                            )
                          )
                 ),
                 
                 tabPanel("K-means Clustering",
                          sidebarLayout(
                            sidebarPanel(
                              uiOutput("kmeans_vars_ui"),
                              uiOutput("kmeans_group_select_ui"),
                              selectInput("distance_method", "Distance Method:",
                                          choices = c("euclidean", "manhattan", "maximum", "minkowski"),
                                          selected = "euclidean"),
                              conditionalPanel(
                                condition = "input.distance_method == 'minkowski'",
                                numericInput("minkowski_power", "Minkowski Power:", 
                                             value = 2, min = 1)
                              ),
                              sliderInput("max_clusters", "Maximum number of clusters to test:",
                                          min = 2, max = 15, value = 10),
                              actionButton("run_elbow", "Run Elbow Analysis", class = "btn-primary"),
                              hr(),
                              numericInput("selected_k", "Number of clusters (k):", 
                                           value = 3, min = 2),
                              actionButton("run_final_kmeans", "Run K-means Analysis", class = "btn-primary"),
                              hr(),
                              uiOutput("kmeans_group_filter_ui"),
                              hr(),
                              h4("Axis Formatting Options"),
                              numericInput("kmeans_axis_title_size", "Axis Title Size:", value = 14, min = 8),
                              numericInput("kmeans_axis_text_size", "Axis Text Size:", value = 12, min = 8),
                              hr(),
                              selectInput("kmeans_download_format", "Download plot format:",
                                          choices = c("pdf", "jpeg", "tiff"),
                                          selected = "pdf"),
                              numericInput("kmeans_download_width", "Plot width (inches):",
                                           value = 10, min = 1),
                              numericInput("kmeans_download_height", "Plot height (inches):",
                                           value = 8, min = 1),
                              numericInput("kmeans_download_dpi", "Plot resolution (DPI):",
                                           value = 300, min = 72),
                              downloadButton("download_elbow_plot", "Download Elbow Plot"),
                              downloadButton("download_kmeans_plot", "Download K-means Plot"),
                              br(), br(),
                              downloadButton("download_clustered_data", "Download Clustered Dataset")
                            ),
                            mainPanel(
                              tabsetPanel(
                                tabPanel("Elbow Analysis",
                                         br(),
                                         plotOutput("elbow_plot"),
                                         br(),
                                         verbatimTextOutput("optimal_clusters")
                                ),
                                tabPanel("K-means Analysis",
                                         br(),
                                         plotOutput("kmeans_plot"),
                                         hr(),
                                         h4("Cluster Summary Statistics"),
                                         DTOutput("cluster_summary"),
                                         hr(),
                                         h4("Sample Details"),
                                         DTOutput("sample_cluster_details")
                                )
                              )
                            )
                          )
                 ),
                 
                 tabPanel("Complete Dataset",
                          sidebarLayout(
                            sidebarPanel(
                              actionButton("generateCompleteData", "Generate Complete Dataset"),
                              br(), br(),
                              downloadButton("download_complete_data", "Download Complete Dataset as Excel")
                            ),
                            mainPanel(
                              DTOutput("complete_data")
                            )
                          )
                 )
)

# Server Definition
server <- function(input, output, session) {
  
  # Data Import Functions
  rawData <- eventReactive(input$loadData, {
    req(input$file, input$headerRow)
    filepath <- input$file$datapath
    if(grepl("\\.csv$", input$file$name, ignore.case = TRUE)) {
      read.csv(filepath, header = TRUE, check.names = FALSE)
    } else {
      sheets <- excel_sheets(filepath)
      updateSelectInput(session, "sheet", choices = sheets)
      read_excel(filepath, sheet = sheets[1], skip = input$headerRow - 1)
    }
  })
  
  output$sheet_ui <- renderUI({
    req(input$file)
    if(!grepl("\\.csv$", input$file$name, ignore.case = TRUE)) {
      sheets <- excel_sheets(input$file$datapath)
      selectInput("sheet", "Select Sheet", choices = sheets)
    }
  })
  
  dataImport <- reactive({
    req(rawData())
    if (!is.null(input$sheet)) {
      read_excel(input$file$datapath, sheet = input$sheet, skip = input$headerRow - 1)
    } else {
      rawData()
    }
  })
  
  output$orig_data <- renderDT({
    req(dataImport())
    datatable(dataImport(), options = list(scrollX = TRUE))
  })
  
  # UI Elements for Dataset Creation
  output$id_col_ui <- renderUI({
    req(dataImport())
    selectInput("id_col", "Select ID Column:",
                choices = names(dataImport()), selected = names(dataImport())[1])
  })
  
  output$group_col_ui <- renderUI({
    req(dataImport())
    selectInput("group_col", "Select Grouping Column (Labels):",
                choices = names(dataImport()), selected = names(dataImport())[1])
  })
  
  output$select_vars_ui <- renderUI({
    req(dataImport())
    selectizeInput("selected_vars", "Select Variables to include:",
                   choices = names(dataImport()), multiple = TRUE)
  })
  
  output$rename_vars_ui <- renderUI({
    req(input$selected_vars)
    lapply(input$selected_vars, function(var) {
      textInput(paste0("newname_", var), label = paste("New name for", var), value = var)
    })
  })
  
  # New Dataset Creation (before outlier removal)
  newDatabase <- eventReactive(input$createDb, {
    req(dataImport(), input$id_col, input$group_col, input$selected_vars)
    df <- dataImport()
    newdf <- df %>% select(all_of(c(input$id_col, input$group_col, input$selected_vars)))
    new_names <- sapply(input$selected_vars, function(var) {
      new_name <- input[[paste0("newname_", var)]]
      ifelse(is.null(new_name), var, new_name)
    })
    newnames_final <- c("ID", input$group_col, new_names)
    names(newdf) <- newnames_final
    newdf
  })
  
  # Reactive final new dataset after optional outlier removal
  finalNewData <- reactive({
    req(newDatabase())
    if (isTRUE(input$remove_outliers)) {
      removeOutliers(newDatabase(), input$group_col, input$selected_vars)
    } else {
      newDatabase()
    }
  })
  
  output$new_data <- renderDT({
    req(finalNewData())
    datatable(finalNewData(), options = list(scrollX = TRUE))
  })
  
  output$downloadData <- downloadHandler(
    filename = function() {
      paste("new_dataset-", Sys.Date(), ".csv", sep = "")
    },
    content = function(file) {
      write.csv(finalNewData(), file, row.names = FALSE)
    }
  )
  
  # PCA Section using finalNewData()
  output$pca_vars_ui <- renderUI({
    req(finalNewData())
    dat <- finalNewData()
    num_vars <- names(dat)[sapply(dat, is.numeric)]
    selectizeInput("pca_vars", "Select variables for PCA:", choices = num_vars, multiple = TRUE)
  })
  
  output$pca_group_filter_ui <- renderUI({
    req(finalNewData())
    groups <- sort(unique(finalNewData()[[input$group_col]]))
    checkboxGroupInput("pca_group_filter", "Select groups to include:", choices = groups, selected = groups)
  })
  
  output$pca_group_colors_ui <- renderUI({
    req(input$pca_group_filter)
    lapply(input$pca_group_filter, function(g) {
      colourInput(paste0("pca_color_", make.names(g)), paste("Color for", g), value = "#56B4E9")
    })
  })
  
  filtered_data <- reactive({
    req(finalNewData(), input$pca_group_filter)
    dat <- finalNewData()
    dat[dat[[input$group_col]] %in% input$pca_group_filter, ]
  })
  
  # Updated PCA reactive: Select complete cases and set row IDs accordingly.
  pcaResult <- reactive({
    req(filtered_data(), input$pca_vars)
    dat <- filtered_data()
    pca_data <- dat %>% select(all_of(input$pca_vars))
    complete_idx <- complete.cases(pca_data)
    valid_ids <- dat$ID[complete_idx]
    pca_data <- pca_data[complete_idx, , drop = FALSE]
    pca_out <- PCA(pca_data, scale.unit = TRUE, graph = FALSE)
    rownames(pca_out$ind$coord) <- valid_ids
    pca_out
  })
  
  output$pca_axes_ui <- renderUI({
    req(pcaResult())
    ncomp <- nrow(pcaResult()$eig)
    comp_choices <- paste0("PC", 1:ncomp)
    tagList(
      selectInput("pca_xaxis", "Select X-axis PC:", choices = comp_choices, selected = "PC1"),
      selectInput("pca_yaxis", "Select Y-axis PC:", choices = comp_choices, selected = "PC2")
    )
  })
  
  output$pcaCombinedPlot <- renderPlot({
    req(filtered_data(), pcaResult(), input$pca_xaxis, input$pca_yaxis)
    dat_filt <- filtered_data()
    group_factor <- factor(dat_filt[[input$group_col]], levels = input$pca_group_filter)
    color_vector <- vapply(input$pca_group_filter,
                           function(g) { input[[paste0("pca_color_", make.names(g))]] },
                           FUN.VALUE = "")
    xAxis <- as.numeric(gsub("PC", "", input$pca_xaxis))
    yAxis <- as.numeric(gsub("PC", "", input$pca_yaxis))
    pca_ids <- rownames(pcaResult()$ind$coord)
    
    plot_df <- data.frame(
      x = pcaResult()$ind$coord[, xAxis],
      y = pcaResult()$ind$coord[, yAxis],
      Group = group_factor[match(pca_ids, dat_filt$ID)],
      ID = pca_ids
    )
    p1 <- ggplot(plot_df, aes(x = x, y = y, color = Group)) +
      geom_point(size = 3, alpha = 0.86) +
      scale_color_manual(values = setNames(color_vector, input$pca_group_filter)) +
      labs(title = "PCA Individuals Plot",
           x = paste0(input$pca_xaxis, " (", round(pcaResult()$eig[xAxis, 2], 1), "%)"),
           y = paste0(input$pca_yaxis, " (", round(pcaResult()$eig[yAxis, 2], 1), "%)")) +
      theme_minimal() +
      theme(axis.title = element_text(size = input$pca_axis_title_size),
            axis.text = element_text(size = input$pca_axis_text_size))
    if(input$show_id_labels) {
      p1 <- p1 + geom_text(aes(label = ID), hjust = -0.2, vjust = 0.5)
    }
    loadings_coords <- as.data.frame(pcaResult()$var$coord)
    selected_loadings <- loadings_coords[, c(xAxis, yAxis)]
    colnames(selected_loadings) <- c("x", "y")
    selected_loadings$Variable <- rownames(selected_loadings)
    p2 <- ggplot(selected_loadings, aes(x = x, y = y, label = Variable)) +
      geom_segment(aes(x = 0, y = 0, xend = x, yend = y),
                   arrow = arrow(length = unit(0.2, "cm")), color = "grey50") +
      geom_text(color = "black", hjust = -0.2) +
      coord_fixed() +
      labs(title = "PCA Loadings Plot",
           x = input$pca_xaxis,
           y = input$pca_yaxis) +
      xlim(-2, 2) + ylim(-2, 2) +
      theme_minimal() +
      theme(axis.title = element_text(size = input$pca_axis_title_size),
            axis.text = element_text(size = input$pca_axis_text_size))
    grid.arrange(p1, p2, ncol = 1, heights = c(3, 2))
  })
  
  output$pcaContributions <- renderPlot({
    req(pcaResult(), input$pca_xaxis)
    xAxis <- as.numeric(gsub("PC", "", input$pca_xaxis))
    fviz_contrib(pcaResult(), choice = "var", axes = xAxis, top = 10) +
      theme_minimal() +
      theme(axis.title = element_text(size = input$pca_axis_title_size),
            axis.text = element_text(size = input$pca_axis_text_size))
  })
  
  output$variableCorrelation <- renderPlot({
    req(filtered_data(), input$pca_vars)
    dat <- filtered_data()
    cor_matrix <- cor(dat[, input$pca_vars], method = "pearson")
    corrplot(cor_matrix,
             method = "color",
             type = "upper",
             order = "hclust",
             addCoef.col = "black",
             tl.col = "black",
             tl.srt = 45,
             col = colorRampPalette(c("blue", "white", "red"))(100),
             diag = FALSE,
             title = "Variable Correlation Matrix",
             mar = c(0,0,1,0))
    # Note: corrplot has its own text options; one can adjust these parameters via tl.cex if needed.
  })
  
  output$sampleCorrelation <- renderPlotly({
    req(filtered_data(), input$pca_vars)
    dat <- filtered_data()
    scaled_data <- scale(dat[, input$pca_vars])
    colnames(scaled_data) <- input$pca_vars
    rownames(scaled_data) <- dat$ID
    sample_var_cor <- cor(t(scaled_data))
    plot_ly(
      x = input$pca_vars,
      y = dat$ID,
      z = sample_var_cor,
      type = "heatmap",
      colors = colorRamp(c("green", "white", "blue")),
      text = matrix(sprintf("%.2f", sample_var_cor), ncol = ncol(sample_var_cor))
    ) %>%
      layout(
        title = "Sample-Variable Correlation Matrix",
        xaxis = list(title = "Variables", 
                     titlefont = list(size = input$pca_axis_title_size),
                     tickfont = list(size = input$pca_axis_text_size),
                     tickangle = 45),
        yaxis = list(title = "Samples",
                     titlefont = list(size = input$pca_axis_title_size),
                     tickfont = list(size = input$pca_axis_text_size))
      )
  })
  
  output$download_pca_plot <- downloadHandler(
    filename = function() {
      paste("pca_plot.", input$pca_download_format, sep = "")
    },
    content = function(file) {
      tryCatch({
        if (input$pca_download_format == "pdf") {
          pdf(file, width = input$pca_download_width, height = input$pca_download_height)
        } else if (input$pca_download_format == "jpeg") {
          jpeg(file, width = input$pca_download_width * input$pca_download_dpi, 
               height = input$pca_download_height * input$pca_download_dpi,
               res = input$pca_download_dpi, quality = 100)
        } else if (input$pca_download_format == "tiff") {
          tiff(file, width = input$pca_download_width * input$pca_download_dpi, 
               height = input$pca_download_height * input$pca_download_dpi,
               res = input$pca_download_dpi, compression = "lzw")
        }
        dat_filt <- filtered_data()
        group_factor <- factor(dat_filt[[input$group_col]], levels = input$pca_group_filter)
        color_vector <- vapply(input$pca_group_filter,
                               function(g) { input[[paste0("pca_color_", make.names(g))]] },
                               FUN.VALUE = "")
        xAxis <- as.numeric(gsub("PC", "", input$pca_xaxis))
        yAxis <- as.numeric(gsub("PC", "", input$pca_yaxis))
        plot_df <- data.frame(
          x = pcaResult()$ind$coord[, xAxis],
          y = pcaResult()$ind$coord[, yAxis],
          Group = group_factor[match(rownames(pcaResult()$ind$coord), dat_filt$ID)],
          ID = rownames(pcaResult()$ind$coord)
        )
        p1 <- ggplot(plot_df, aes(x = x, y = y, color = Group)) +
          geom_point(size = 3, alpha = 0.86) +
          scale_color_manual(values = setNames(color_vector, input$pca_group_filter)) +
          labs(title = "PCA Individuals Plot",
               x = paste0(input$pca_xaxis, " (", round(pcaResult()$eig[xAxis, 2], 1), "%)"),
               y = paste0(input$pca_yaxis, " (", round(pcaResult()$eig[yAxis, 2], 1), "%)")) +
          theme_minimal() +
          theme(axis.title = element_text(size = input$pca_axis_title_size),
                axis.text = element_text(size = input$pca_axis_text_size))
        if(input$show_id_labels) {
          p1 <- p1 + geom_text(aes(label = ID), hjust = -0.2, vjust = 0.5)
        }
        loadings_coords <- as.data.frame(pcaResult()$var$coord)
        selected_loadings <- loadings_coords[, c(xAxis, yAxis)]
        colnames(selected_loadings) <- c("x", "y")
        selected_loadings$Variable <- rownames(selected_loadings)
        p2 <- ggplot(selected_loadings, aes(x = x, y = y, label = Variable)) +
          geom_segment(aes(x = 0, y = 0, xend = x, yend = y),
                       arrow = arrow(length = unit(0.2, "cm")), color = "grey50") +
          geom_text(color = "black", hjust = -0.2) +
          coord_fixed() +
          labs(title = "PCA Loadings Plot",
               x = input$pca_xaxis,
               y = input$pca_yaxis) +
          xlim(-1, 1) + ylim(-1, 1) +
          theme_minimal() +
          theme(axis.title = element_text(size = input$pca_axis_title_size),
                axis.text = element_text(size = input$pca_axis_text_size))
        grid.arrange(p1, p2, ncol = 1, heights = c(3, 2))
        dev.off()
      }, error = function(e) {
        if (dev.cur() > 1) dev.off()
        stop(e)
      })
    }
  )
  
  output$download_contrib_plot <- downloadHandler(
    filename = function() {
      paste("contribution_plot.", input$pca_download_format, sep = "")
    },
    content = function(file) {
      tryCatch({
        if (input$pca_download_format == "pdf") {
          pdf(file, width = input$pca_download_width, height = input$pca_download_height)
        } else if (input$pca_download_format == "jpeg") {
          jpeg(file, width = input$pca_download_width * input$pca_download_dpi, 
               height = input$pca_download_height * input$pca_download_dpi,
               res = input$pca_download_dpi, quality = 100)
        } else if (input$pca_download_format == "tiff") {
          tiff(file, width = input$pca_download_width * input$pca_download_dpi, 
               height = input$pca_download_height * input$pca_download_dpi,
               res = input$pca_download_dpi, compression = "lzw")
        }
        xAxis <- as.numeric(gsub("PC", "", input$pca_xaxis))
        print(fviz_contrib(pcaResult(), choice = "var", axes = xAxis, top = 10) +
                theme_minimal() +
                theme(axis.title = element_text(size = input$pca_axis_title_size),
                      axis.text = element_text(size = input$pca_axis_text_size)))
        dev.off()
      }, error = function(e) {
        if (dev.cur() > 1) dev.off()
        stop(e)
      })
    }
  )
  
  output$download_corr_plot <- downloadHandler(
    filename = function() {
      paste("correlation_plot.", input$pca_download_format, sep = "")
    },
    content = function(file) {
      tryCatch({
        if (input$pca_download_format == "pdf") {
          pdf(file, width = input$pca_download_width, height = input$pca_download_height)
        } else if (input$pca_download_format == "jpeg") {
          jpeg(file, width = input$pca_download_width * input$pca_download_dpi, 
               height = input$pca_download_height * input$pca_download_dpi,
               res = input$pca_download_dpi, quality = 100)
        } else if (input$pca_download_format == "tiff") {
          tiff(file, width = input$pca_download_width * input$pca_download_dpi, 
               height = input$pca_download_height * input$pca_download_dpi,
               res = input$pca_download_dpi, compression = "lzw")
        }
        dat <- filtered_data()
        cor_matrix <- cor(dat[, input$pca_vars], method = "pearson")
        corrplot(cor_matrix,
                 method = "color",
                 type = "upper",
                 order = "hclust",
                 addCoef.col = "black",
                 tl.col = "black",
                 tl.srt = 45,
                 col = colorRampPalette(c("blue", "white", "red"))(100),
                 diag = FALSE,
                 title = "Variable Correlation Matrix",
                 mar = c(0,0,1,0))
        dev.off()
      }, error = function(e) {
        if (dev.cur() > 1) dev.off()
        stop(e)
      })
    }
  )
  
  # New output for PC Matrix tab in PCA.
  output$pc_matrix <- renderDT({
    req(pcaResult())
    dat <- as.data.frame(pcaResult()$ind$coord)
    dat$ID <- rownames(dat)
    datatable(dat, options = list(scrollX = TRUE))
  })
  
  output$download_pc_matrix <- downloadHandler(
    filename = function() {
      paste("pc_matrix-", Sys.Date(), ".xlsx", sep = "")
    },
    content = function(file) {
      req(pcaResult())
      dat <- as.data.frame(pcaResult()$ind$coord)
      dat$ID <- rownames(dat)
      writexl::write_xlsx(dat, path = file)
    }
  )
  
  # Violin Plot Section
  output$violin_var_ui <- renderUI({
    req(finalNewData())
    dat <- finalNewData()
    num_vars <- names(dat)[sapply(dat, is.numeric)]
    selectInput("violin_var", "Select variable for Violin Plot:", choices = num_vars)
  })
  
  output$violin_group_filter_ui <- renderUI({
    req(finalNewData())
    groups <- sort(unique(finalNewData()[[input$group_col]]))
    checkboxGroupInput("violin_group_filter", "Select groups to display:", choices = groups, selected = groups)
  })
  
  output$violin_group_colors_ui <- renderUI({
    req(input$violin_group_filter)
    lapply(input$violin_group_filter, function(g) {
      colourInput(paste0("violin_color_", make.names(g)), paste("Color for", g), value = "#56B4E9")
    })
  })
  
  output$violinPlot <- renderPlot({
    req(finalNewData(), input$violin_var, input$violin_group_filter)
    dat <- finalNewData()
    dat_filtered <- dat[dat[[input$group_col]] %in% input$violin_group_filter, ]
    dat_filtered[[input$group_col]] <- factor(dat_filtered[[input$group_col]], levels = input$violin_group_filter)
    color_vector <- vapply(input$violin_group_filter,
                           function(g) { input[[paste0("violin_color_", make.names(g))]] },
                           FUN.VALUE = "")
    p <- ggplot(dat_filtered, aes(x = .data[[input$group_col]], y = .data[[input$violin_var]], fill = .data[[input$group_col]]))
    if(input$show_violin) {
      p <- p + geom_violin(alpha = input$violin_alpha, scale = "width")
    }
    p <- p + scale_fill_manual(values = setNames(color_vector, input$violin_group_filter)) +
      labs(x = input$group_col, y = input$violin_var, fill = input$group_col) +
      theme_minimal() +
      theme(axis.title = element_text(size = input$violin_axis_title_size),
            axis.text = element_text(size = input$violin_axis_text_size),
            legend.position = "none")
    if(input$show_boxplot) {
      p <- p + geom_boxplot(width = 0.2, fill = "white", alpha = 0.5, outlier.shape = NA)
    }
    if(input$show_points) {
      p <- p + geom_jitter(width = 0.1, alpha = 0.5)
    }
    p
  })
  
  output$download_violin_plot <- downloadHandler(
    filename = function() {
      paste("violin_plot.", input$violin_download_format, sep = "")
    },
    content = function(file) {
      tryCatch({
        if (input$violin_download_format == "pdf") {
          pdf(file, width = input$violin_download_width, height = input$violin_download_height)
        } else if (input$violin_download_format == "jpeg") {
          jpeg(file, width = input$violin_download_width * input$violin_download_dpi, 
               height = input$violin_download_height * input$violin_download_dpi,
               res = input$violin_download_dpi, quality = 100)
        } else if (input$violin_download_format == "tiff") {
          tiff(file, width = input$violin_download_width * input$violin_download_dpi, 
               height = input$violin_download_height * input$violin_download_dpi,
               res = input$violin_download_dpi, compression = "lzw")
        }
        dat <- finalNewData()
        dat_filtered <- dat[dat[[input$group_col]] %in% input$violin_group_filter, ]
        dat_filtered[[input$group_col]] <- factor(dat_filtered[[input$group_col]], levels = input$violin_group_filter)
        color_vector <- vapply(input$violin_group_filter,
                               function(g) { input[[paste0("violin_color_", make.names(g))]] },
                               FUN.VALUE = "")
        p <- ggplot(dat_filtered, aes(x = .data[[input$group_col]], y = .data[[input$violin_var]], fill = .data[[input$group_col]]))
        if(input$show_violin) {
          p <- p + geom_violin(alpha = input$violin_alpha, scale = "width")
        }
        p <- p + scale_fill_manual(values = setNames(color_vector, input$violin_group_filter)) +
          labs(x = input$group_col, y = input$violin_var, fill = input$group_col) +
          theme_minimal() +
          theme(axis.title = element_text(size = input$violin_axis_title_size),
                axis.text = element_text(size = input$violin_axis_text_size),
                legend.position = "none")
        if(input$show_boxplot) {
          p <- p + geom_boxplot(width = 0.2, fill = "white", alpha = 0.5, outlier.shape = NA)
        }
        if(input$show_points) {
          p <- p + geom_jitter(width = 0.1, alpha = 0.5)
        }
        print(p)
        dev.off()
      }, error = function(e) {
        if (dev.cur() > 1) dev.off()
        stop(e)
      })
    }
  )
  
  # K-means Analysis Section
  output$kmeans_vars_ui <- renderUI({
    req(finalNewData())
    dat <- finalNewData()
    num_vars <- names(dat)[sapply(dat, is.numeric)]
    selectizeInput("kmeans_vars", "Select variables for clustering:",
                   choices = num_vars, multiple = TRUE)
  })
  
  output$kmeans_group_select_ui <- renderUI({
    req(finalNewData())
    groups <- unique(finalNewData()[[input$group_col]])
    selectizeInput("kmeans_groups_selected", "Select groups for analysis:",
                   choices = groups, multiple = TRUE, selected = groups)
  })
  
  output$kmeans_group_filter_ui <- renderUI({
    req(kmeans_result())
    clusters <- sort(unique(kmeans_result()$cluster))
    checkboxGroupInput("kmeans_group_filter", "Select clusters to display:",
                       choices = clusters, selected = clusters)
  })
  
  kmeans_data <- reactive({
    req(finalNewData(), input$kmeans_vars, input$kmeans_groups_selected)
    filtered_data <- finalNewData()[finalNewData()[[input$group_col]] %in% input$kmeans_groups_selected, ]
    dat <- filtered_data[, input$kmeans_vars, drop = FALSE]
    scale(dat)
  })
  
  perform_clustering <- function(data, k, distance_method, power = 2) {
    if(distance_method == "euclidean") {
      res <- kmeans(data, centers = k, nstart = 25)
      return(res)
    } else {
      dist_matrix <- switch(distance_method,
                            "manhattan" = dist(data, method = "manhattan"),
                            "maximum" = dist(data, method = "maximum"),
                            "minkowski" = dist(data, method = "minkowski", p = power))
      pam_result <- pam(dist_matrix, k = k)
      centers <- as.data.frame(data)[pam_result$medoids, , drop = FALSE]
      list(
        cluster = pam_result$clustering,
        centers = centers,
        size = pam_result$clusinfo[, "size"],
        withinss = pam_result$objective
      )
    }
  }
  
  elbow_data <- eventReactive(input$run_elbow, {
    req(kmeans_data())
    power <- if(input$distance_method == "minkowski") input$minkowski_power else 2
    wss <- sapply(1:input$max_clusters, function(k) {
      if(k == 1) return(sum(kmeans_data()^2))
      clustering <- perform_clustering(kmeans_data(), k, input$distance_method, power)
      if(input$distance_method == "euclidean") {
        return(clustering$tot.withinss)
      } else {
        return(sum(clustering$withinss))
      }
    })
    data.frame(k = 1:input$max_clusters, wss = wss)
  })
  
  output$elbow_plot <- renderPlot({
    req(elbow_data())
    ggplot(elbow_data(), aes(x = k, y = wss)) +
      geom_line() +
      geom_point() +
      labs(x = "Number of Clusters (k)",
           y = "Total Within Sum of Squares",
           title = paste("Elbow Method for Optimal k Selection\nusing", input$distance_method, "distance")) +
      theme_minimal() +
      theme(axis.title = element_text(size = input$kmeans_axis_title_size),
            axis.text = element_text(size = input$kmeans_axis_text_size),
            plot.title = element_text(hjust = 0.5))
  })
  
  output$optimal_clusters <- renderText({
    req(elbow_data())
    wss_diff <- diff(elbow_data()$wss)
    wss_diff2 <- diff(wss_diff)
    optimal_k <- which.max(abs(wss_diff2)) + 1
    paste("Suggested optimal number of clusters (k):", optimal_k,
          "\nUsing", input$distance_method, "distance")
  })
  
  kmeans_result <- eventReactive(input$run_final_kmeans, {
    req(kmeans_data())
    power <- if(input$distance_method == "minkowski") input$minkowski_power else 2
    perform_clustering(kmeans_data(), input$selected_k, input$distance_method, power)
  })
  
  output$kmeans_plot <- renderPlot({
    req(kmeans_result(), input$kmeans_group_filter)
    if(length(input$kmeans_vars) < 2) return(NULL)
    data_plot <- as.data.frame(kmeans_data())
    names(data_plot) <- input$kmeans_vars
    data_plot$Cluster <- as.factor(kmeans_result()$cluster)
    data_plot$Selected <- data_plot$Cluster %in% input$kmeans_group_filter
    xvar <- paste0("`", input$kmeans_vars[1], "`")
    yvar <- paste0("`", input$kmeans_vars[2], "`")
    p <- ggplot(data_plot, aes_string(x = xvar, y = yvar, color = "Cluster")) +
      geom_point(data = subset(data_plot, !Selected), color = "grey80", alpha = 0.3) +
      geom_point(data = subset(data_plot, Selected), alpha = 0.6) +
      stat_ellipse(data = subset(data_plot, Selected), type = "norm", level = 0.95) +
      scale_color_brewer(palette = "Set1") +
      labs(title = paste("Clustering with", input$distance_method, "distance"),
           x = input$kmeans_vars[1],
           y = input$kmeans_vars[2]) +
      theme_minimal() +
      theme(axis.title = element_text(size = input$kmeans_axis_title_size),
            axis.text = element_text(size = input$kmeans_axis_text_size),
            legend.position = "right")
    centers_df <- as.data.frame(kmeans_result()$centers)
    names(centers_df) <- input$kmeans_vars
    p + geom_point(data = centers_df,
                   aes_string(x = paste0("`", input$kmeans_vars[1], "`"), 
                              y = paste0("`", input$kmeans_vars[2], "`")),
                   color = "black", size = 4, shape = 8)
  })
  
  output$cluster_summary <- renderDT({
    req(kmeans_result(), finalNewData())
    filtered_data_clust <- finalNewData()[finalNewData()[[input$group_col]] %in% input$kmeans_groups_selected, ]
    cluster_data <- data.frame(filtered_data_clust[, input$kmeans_vars, drop = FALSE],
                               Cluster = as.factor(kmeans_result()$cluster))
    summary_stats <- cluster_data %>%
      group_by(Cluster) %>%
      summarise(across(everything(), list(mean = ~mean(.), sd = ~sd(.)), .names = "{.col}_{.fn}"),
                Size = n())
    datatable(summary_stats, options = list(scrollX = TRUE),
              caption = paste("Cluster Summary Statistics using", input$distance_method, "distance"))
  })
  
  output$sample_cluster_details <- renderDT({
    req(kmeans_result(), finalNewData())
    filtered_data_clust <- finalNewData()[finalNewData()[[input$group_col]] %in% input$kmeans_groups_selected, ]
    sample_data <- data.frame(ID = filtered_data_clust$ID,
                              Group = filtered_data_clust[[input$group_col]],
                              Cluster = as.factor(kmeans_result()$cluster),
                              Selected = kmeans_result()$cluster %in% input$kmeans_group_filter)
    for(var in input$kmeans_vars) {
      sample_data[[var]] <- filtered_data_clust[[var]]
    }
    sample_data <- sample_data[order(sample_data$Cluster),]
    datatable(sample_data,
              options = list(scrollX = TRUE, pageLength = 15, dom = 'Bfrtip', buttons = c('copy', 'csv', 'excel')),
              rownames = FALSE) %>% 
      formatStyle('Selected', target = 'row', backgroundColor = styleEqual(c(TRUE, FALSE), c('#fff3cd', 'white')))
  })
  
  clusteredData <- reactive({
    req(kmeans_result(), finalNewData())
    filtered_data_clust <- finalNewData()[finalNewData()[[input$group_col]] %in% input$kmeans_groups_selected, ]
    filtered_data_clust$Cluster <- kmeans_result()$cluster
    filtered_data_clust
  })
  
  output$download_clustered_data <- downloadHandler(
    filename = function() {
      paste("clustered_dataset-", Sys.Date(), ".csv", sep = "")
    },
    content = function(file) {
      write.csv(clusteredData(), file, row.names = FALSE)
    }
  )
  
  output$download_elbow_plot <- downloadHandler(
    filename = function() {
      paste("elbow_plot", input$kmeans_download_format, sep = ".")
    },
    content = function(file) {
      ggsave(file,
             plot = {
               ggplot(elbow_data(), aes(x = k, y = wss)) +
                 geom_line() +
                 geom_point() +
                 labs(x = "Number of Clusters (k)",
                      y = "Total Within Sum of Squares",
                      title = paste("Elbow Method for Optimal k Selection\nusing", input$distance_method, "distance")) +
                 theme_minimal() +
                 theme(axis.title = element_text(size = input$kmeans_axis_title_size),
                       axis.text = element_text(size = input$kmeans_axis_text_size),
                       plot.title = element_text(hjust = 0.5))
             },
             width = input$kmeans_download_width,
             height = input$kmeans_download_height,
             dpi = input$kmeans_download_dpi)
    }
  )
  
  output$download_kmeans_plot <- downloadHandler(
    filename = function() {
      paste("kmeans_plot", input$kmeans_download_format, sep = ".")
    },
    content = function(file) {
      ggsave(file,
             plot = {
               data_plot <- as.data.frame(kmeans_data())
               names(data_plot) <- input$kmeans_vars
               data_plot$Cluster <- as.factor(kmeans_result()$cluster)
               data_plot$Selected <- data_plot$Cluster %in% input$kmeans_group_filter
               xvar <- paste0("`", input$kmeans_vars[1], "`")
               yvar <- paste0("`", input$kmeans_vars[2], "`")
               p <- ggplot(data_plot, aes_string(x = xvar, y = yvar, color = "Cluster")) +
                 geom_point(data = subset(data_plot, !Selected), color = "grey80", alpha = 0.3) +
                 geom_point(data = subset(data_plot, Selected), alpha = 0.6) +
                 stat_ellipse(data = subset(data_plot, Selected), type = "norm", level = 0.95) +
                 scale_color_brewer(palette = "Set1") +
                 labs(title = paste("Clustering with", input$distance_method, "distance"),
                      x = input$kmeans_vars[1],
                      y = input$kmeans_vars[2]) +
                 theme_minimal() +
                 theme(axis.title = element_text(size = input$kmeans_axis_title_size),
                       axis.text = element_text(size = input$kmeans_axis_text_size),
                       legend.position = "right")
               centers_df <- as.data.frame(kmeans_result()$centers)
               names(centers_df) <- input$kmeans_vars
               p + geom_point(data = centers_df,
                              aes_string(x = paste0("`", input$kmeans_vars[1], "`"), y = paste0("`", input$kmeans_vars[2], "`")),
                              color = "black", size = 4, shape = 8)
             },
             width = input$kmeans_download_width,
             height = input$kmeans_download_height,
             dpi = input$kmeans_download_dpi)
    }
  )
  
  # Complete Dataset Section: Merge new dataset with PCA dimensions and clustering.
  completeDataset <- reactive({
    validate(
      need(!is.null(finalNewData()), "generate missing dataset"),
      need(!is.null(pcaResult()), "generate missing dataset"),
      need(!is.null(clusteredData()), "generate missing dataset")
    )
    new_ds <- finalNewData()[, c("ID", input$group_col, input$selected_vars)]
    pc_mat <- as.data.frame(pcaResult()$ind$coord)
    pc_mat$ID <- rownames(pc_mat)
    cluster_ds <- clusteredData()[, c("ID", "Cluster")]
    complete_ds <- merge(new_ds, pc_mat, by = "ID", all = FALSE)
    complete_ds <- merge(complete_ds, cluster_ds, by = "ID", all = FALSE)
    complete_ds
  })
  
  observeEvent(input$generateCompleteData, {
    output$complete_data <- renderDT({
      req(completeDataset())
      datatable(completeDataset(), options = list(scrollX = TRUE))
    })
  })
  
  output$download_complete_data <- downloadHandler(
    filename = function() {
      paste("complete_dataset-", Sys.Date(), ".xlsx", sep = "")
    },
    content = function(file) {
      req(completeDataset())
      writexl::write_xlsx(completeDataset(), path = file)
    }
  )
  
}

# Run the application
shinyApp(ui = ui, server = server)
