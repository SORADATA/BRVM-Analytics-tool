# ===================== CHARGEMENT DES PACKAGES =====================
packages <- c(
  "shiny", "shinythemes", "shinycssloaders", "readxl", "dplyr", "stringr", 
  "rlang", "tibble", "tm", "tidytext", "wordcloud", "RColorBrewer", 
  "ggplot2", "tidyr", "widyr", "igraph", "ggraph", "scales", 
  "topicmodels", "slam", "LDAvis", "DT", "base64enc", "pdftools"
)

for (pkg in packages) {
  if (!require(pkg, character.only = TRUE)) {
    install.packages(pkg)
    library(pkg, character.only = TRUE)
  }
}

# ===================== FONCTIONS ET DONNÉES GLOBALES =====================
stopwords_fr <- tibble(word = stopwords("fr"))
mots_a_supprimer <- tibble(word = c(
  "impact", "d", "impacts", "k", "à", "a", "aucune", "absolument", "actuellement", 
  "ainsi", "alors", "apparemment", "approximativement", "apres", "assez", "assurement", 
  "au", "aucunement", "aucuns", "aujourd'hui", "auparavant", "aussi", "aussitot", 
  "autant", "autre", "autrefois", "autrement", "avant", "avec", "avoir", "beaucoup", 
  "bien", "bientot", "bon", "car", "carrement", "ce", "cela", "cependant", 
  "certainement", "certes", "ces", "ceux", "chaque", "ci", "comme", "comment", 
  "completement", "d'", "abord", "dans", "davantage", "de", "debut", "dedans", 
  "dehors", "deja", "demain", "depuis", "derechef", "des", "desormais", "deux", 
  "devrait", "diablement", "divinement", "doit", "donc", "dorenavant", "dos", 
  "droite", "drolement", "du", "elle", "elles", "en", "en verite", "encore", 
  "enfin", "ensuite", "entierement", "environ", "essai", "est", "et", "etaient", 
  "etat", "ete", "etions", "etre", "eu", "extremement", "fait", "faites", "fois", 
  "font", "force", "grandement", "guere", "habituellement", "haut", "hier", "hors", 
  "ici", "il", "ils", "insuffisamment", "jadis", "jamais", "je", "joliment", "la", 
  "le", "les", "leur", "leurs", "lol", "longtemps", "lors", "ma", "maintenant", 
  "mais", "mdr", "meme", "mes", "moins", "mon", "mot", "naguere", "ne", "ni", 
  "nommes", "non", "notre", "nous", "nouveaux", "nullement", "ou", "oui", "par", 
  "parce", "parfois", "parole", "pas", "pas mal", "passablement", "personne", 
  "personnes", "peu", "peut", "piece", "plupart", "plus", "plutot", "point", 
  "pour", "pourquoi", "precisement", "premierement", "presque", "prou", "puis", 
  "quand", "quasi", "quasiment", "que", "quel", "quelle", "quelles", "quelque", 
  "quelquefois", "quels", "qui", "quoi", "quotidiennement", "rien", "rudement", 
  "s'", "sa", "sans", "sans doute", "ses", "seulement", "si", "sien", "sitot", 
  "soit", "son", "sont", "soudain", "sous", "souvent", "soyez", "subitement", 
  "suffisamment", "sur", "t'", "ta", "tandis", "tant", "tantot", "tard", 
  "tellement", "tels", "terriblement", "tes", "ton", "tot", "totalement", 
  "toujours", "tous", "tout", "toutefois", "tres", "trop", "tu", "un", "une", 
  "valeur", "vers", "voie", "voient", "volontiers", "vont", "votre", "vous", 
  "vraiment", "vraisemblablement", "ras"
))
all_stopwords_base <- bind_rows(stopwords_fr, mots_a_supprimer) %>% distinct()

# ===================== FONCTION DE PARSING "BULLETPROOF" =====================
extract_finance_syscohada <- function(text_full) {
  lines <- unlist(strsplit(text_full, "\n"))
  
  parse_metric <- function(pattern) {
    idx <- grep(pattern, lines, ignore.case = TRUE)
    if(length(idx) > 0) {
      line <- lines[idx[1]]
      matches <- stringr::str_extract_all(line, "\\(?[0-9]{1,3}(?:[\\s\\.\\,][0-9]{3})+\\)?")[[1]]
      if(length(matches) > 0) {
        val_str <- matches[1]
        is_neg <- grepl("\\(", val_str) || grepl("-", val_str)
        val_num <- as.numeric(stringr::str_replace_all(val_str, "[^0-9]", ""))
        if(is_neg) val_num <- -val_num
        return(val_num)
      }
    }
    return(0)
  }
  
  res <- list(
    ca = parse_metric("Chiffre d.Affaires total"),
    rn = parse_metric("R.sultat net de l.exercice"), 
    stocks = parse_metric("Stocks nets"),
    creances = parse_metric("Cr.ances et emplois"),
    passif_hao = parse_metric("Passif circulant H\\.A\\.O"),
    passif_autres = parse_metric("Autres Dettes"),
    caf = parse_metric("Capacit. d.autofinancement"),
    cfo = parse_metric("Flux de tr.sorerie des activit.s"),
    treso_nette = parse_metric("Tr.sorerie Nette au 31 d.cembre") 
  )
  
  res$passif_circulant <- res$passif_hao + res$passif_autres
  res$bfr <- (res$stocks + res$creances) - res$passif_circulant
  res$dso <- ifelse(res$ca > 0, (res$creances / res$ca) * 360, 0)
  
  return(res)
}

myValueBox <- function(title, value, color = "#003366", icon_name = "chart-line") {
  div(style = paste0("background-color: ", color, "; color: white; padding: 20px; border-radius: 8px; text-align: center; margin-bottom: 20px; box-shadow: 0 4px 8px rgba(0,0,0,0.1);"),
      h2(value, style = "margin-top: 0; font-weight: bold; font-size: 28px;"),
      h5(icon(icon_name), " ", title, style = "margin-bottom: 0; font-size: 16px; opacity: 0.9;")
  )
}

# ===================== INTERFACE UTILISATEUR (UI) =====================
ui <- fluidPage(
  theme = shinytheme("lumen"),
  tags$head(
    tags$style(HTML("
      .header-container { display: flex; align-items: center; gap: 20px; margin-bottom: 20px; }
      .title-group h1 { color: #003366; margin-bottom: 0; font-weight: bold;}
      .title-group h2, h3, h4 { color: #003366; }
      .well { background-color: #f5faff; border: 1px solid #e6f2ff; }
      .nav-tabs > li > a { color: #0056b3; font-weight: bold; }
      .nav-tabs > li.active > a, .nav-tabs > li.active > a:hover, .nav-tabs > li.active > a:focus {
        color: #fff; background-color: #003366; border-color: #003366;
      }
      .btn-primary { background-color: #003366; border-color: #002244; }
      .shiny-notification { position:fixed; top: calc(50% - 50px); left: calc(50% - 150px); width: 300px; }
    "))
  ),
  
  div(class = "header-container",
      uiOutput("logo_ui"),
      div(class = "title-group",
          h1("BRVM Analytics Hub"),
          h4("Text Mining & Analyse Financière SYSCOHADA")
      )
  ),
  
  sidebarLayout(
    sidebarPanel(
      width = 3,
      h4("1. Source des données"),
      fileInput("upload_file", "Charger un fichier (.xlsx ou .pdf)", 
                accept = c(".xlsx", ".pdf"), buttonLabel = "Parcourir..."),
      uiOutput("sheet_selector_ui"),
      uiOutput("column_selector_ui"),
      hr(),
      h4("2. Options d'analyse (Texte)"),
      textAreaInput("custom_stopwords", "Mots à ignorer (séparés par une virgule)", 
                    placeholder = "ex: projet, client, entreprise"),
      actionButton("analyser", "Lancer l'analyse", icon = icon("play"), class = "btn-primary btn-lg"),
      hr(),
      h4("3. Visualisation avancée (Texte)"),
      textInput("mot_cible", "Mot à analyser (corrélations) :", placeholder = "Entrez un mot..."),
      actionButton("ldavis", "LDA Interactive", icon = icon("search-plus"))
    ),
    
    mainPanel(
      width = 9,
      tabsetPanel(
        id = "main_tabs",
        tabPanel("Analyse Financière", icon = icon("chart-line"), uiOutput("finance_ui")),
        tabPanel("Aperçu des données", icon = icon("eye"), h4("Données nettoyées prêtes pour l'analyse"), withSpinner(DTOutput("cleaned_data_table"))),
        tabPanel("Fréquences", icon = icon("table"), withSpinner(DTOutput("freq_table"))),
        tabPanel("Nuage de mots", icon = icon("cloud"),
                 tabsetPanel(
                   tabPanel("Global", sliderInput("nb_mots_global", "Nombre de mots :", min = 10, max = 200, value = 100), withSpinner(plotOutput("wordcloud_global", height = "600px")), downloadButton("download_wordcloud_global", "Télécharger")),
                   tabPanel("Spécifique", textAreaInput("specific_words", "Mots spécifiques à afficher (séparés par une virgule)"), sliderInput("nb_mots_specifique", "Nombre de mots :", min = 5, max = 100, value = 50), withSpinner(plotOutput("wordcloud_specifique", height = "600px")), downloadButton("download_wordcloud_specifique", "Télécharger"))
                 )),
        tabPanel("N-grammes", icon = icon("project-diagram"), selectInput("ngram_n", "Type de n-gramme", choices = c("Bigrammes (2)" = 2, "Trigrammes (3)" = 3), selected = 2), sliderInput("nb_ngrams", "Nombre de n-grammes à afficher :", min = 5, max = 50, value = 15), withSpinner(plotOutput("ngram_plot", height = "700px"))),
        tabPanel("Relations entre mots", icon = icon("network-wired"),
                 tabsetPanel(
                   tabPanel("Graphe de cooccurrences", sliderInput("cooc_freq", "Fréquence minimale de cooccurrence :", min = 1, max = 20, value = 2), withSpinner(plotOutput("cooc_graph", height = "700px"))),
                   tabPanel("Corrélations (mot-cible)", withSpinner(plotOutput("correlation_plot", height = "700px")))
                 )),
        tabPanel("Analyse de Sentiment", icon = icon("smile-beam"), withSpinner(plotOutput("sentiment_plot", height = "600px"))),
        tabPanel("Thèmes (LDA)", icon = icon("boxes"),
                 tabsetPanel(
                   tabPanel("Graphique des thèmes", sliderInput("lda_k", "Nombre de thèmes :", min = 2, max = 10, value = 4), sliderInput("lda_n_words", "Nombre de mots par thème :", min = 4, max = 20, value = 10), withSpinner(plotOutput("lda_plot", height = "700px"))),
                   tabPanel("Visualisation Interactive (LDAvis)", actionButton("run_ldavis", "Lancer la visualisation interactive", icon = icon("rocket")), withSpinner(visOutput("ldavis_output")))
                 ))
      )
    )
  )
)

# ===================== SERVEUR =====================
server <- function(input, output, session) {
  
  output$logo_ui <- renderUI({
    logo_path <- "www/logo.png"
    if (file.exists(logo_path)) {
      img_data <- base64enc::dataURI(file = logo_path, mime = "image/png")
      tags$img(src = img_data, height = "80px")
    } else { NULL }
  })
  
  file_ext <- reactive({
    req(input$upload_file)
    tools::file_ext(input$upload_file$name)
  })
  
  sheet_names <- reactive({
    req(input$upload_file)
    if(file_ext() == "xlsx") excel_sheets(input$upload_file$datapath) else NULL
  })
  
  output$sheet_selector_ui <- renderUI({
    if(file_ext() == "xlsx") {
      selectInput("selected_sheet", "Choisir la feuille :", choices = sheet_names())
    } else {
      helpText(icon("file-pdf"), " Mode PDF détecté : Analyse financière activée.", style="color:#0056b3; font-weight:bold;")
    }
  })
  
  raw_data <- reactive({
    req(input$upload_file)
    if(file_ext() == "xlsx") {
      req(input$selected_sheet)
      read_excel(input$upload_file$datapath, sheet = input$selected_sheet)
    } else {
      txt <- pdftools::pdf_text(input$upload_file$datapath) %>% paste(collapse = "\n")
      tibble(Texte_PDF = txt)
    }
  })
  
  output$column_selector_ui <- renderUI({
    req(raw_data())
    if(file_ext() == "xlsx") {
      selectInput("selected_column", "Choisir la colonne de texte :", choices = colnames(raw_data()))
    } else {
      selectInput("selected_column", "Colonne de texte :", choices = "Texte_PDF", selected = "Texte_PDF")
    }
  })
  
  data_fi <- reactive({
    req(input$upload_file, file_ext() == "pdf")
    text_brut <- pdftools::pdf_text(input$upload_file$datapath) %>% paste(collapse = "\n")
    extract_finance_syscohada(text_brut)
  })
  
  output$finance_ui <- renderUI({
    req(input$upload_file)
    if(file_ext() != "pdf") {
      return(div(style="text-align:center; padding: 50px;",
                 h3(icon("info-circle"), " Veuillez charger un rapport financier au format PDF pour afficher l'analyse SYSCOHADA.", style="color:#6c757d;")))
    }
    
    d <- data_fi()
    
    tagList(
      br(),
      fluidRow(
        column(4, myValueBox("Chiffre d'Affaires", paste0(round(d$ca/1e9, 2), " Mds"), "#0056b3", "shopping-cart")),
        column(4, myValueBox("Résultat Net", paste0(round(d$rn/1e9, 2), " Mds"), "#28a745", "piggy-bank")),
        column(4, myValueBox("Marge Nette", paste0(round(ifelse(d$ca>0, (d$rn/d$ca)*100, 0), 1), " %"), "#6f42c1", "percent"))
      ),
      fluidRow(
        column(4, myValueBox("BFR (Immobilisé)", paste0(round(d$bfr/1e9, 2), " Mds"), "#fd7e14", "warehouse")),
        column(4, myValueBox("CAF (Génération)", paste0(round(d$caf/1e9, 2), " Mds"), "#17a2b8", "cogs")),
        column(4, myValueBox("Trésorerie Nette", paste0(round(d$treso_nette/1e9, 2), " Mds"), ifelse(d$cfo < 0, "#dc3545", "#20c997"), "wallet"))
      ),
      fluidRow(
        column(6, div(class="well", h4("Structure du Besoin en Fonds de Roulement"), plotOutput("plot_bfr_fi"))),
        column(6, div(class="well", h4("Diagnostic Automatisé (Expertise FI)"), uiOutput("diagnostic_ui_fi")))
      )
    )
  })
  
  output$plot_bfr_fi <- renderPlot({
    req(data_fi())
    d <- data_fi()
    df_plot <- data.frame(
      Poste = c("Stocks (+)", "Créances (+)", "Passif Circ. (-)"),
      Valeur = c(d$stocks, d$creances, -d$passif_circulant),
      Type = c("Emploi", "Emploi", "Ressource")
    )
    if(sum(abs(df_plot$Valeur)) == 0) return(ggplot() + theme_void() + ggtitle("En attente de données..."))
    
    ggplot(df_plot, aes(x = Poste, y = Valeur / 1e9, fill = Type)) +
      geom_bar(stat = "identity") +
      scale_fill_manual(values = c("Emploi" = "#fd7e14", "Ressource" = "#17a2b8")) +
      theme_minimal(base_size = 14) +
      labs(y = "Milliards FCFA", x = "") + theme(legend.position = "none")
  })
  
  output$diagnostic_ui_fi <- renderUI({
    req(data_fi())
    d <- data_fi()
    diag_cfo <- ifelse(d$cfo < 0, 
                       "<span style='color:#dc3545;'><b>Alerte Cash :</b> Le flux d'exploitation est négatif. L'entreprise brûle du cash.</span>", 
                       "<span style='color:#28a745;'><b>Cash Flow Sain :</b> L'exploitation génère de la trésorerie positive.</span>")
    diag_dso <- paste0("<b>Délai Client (DSO) :</b> Les clients paient en moyenne en <b>", round(d$dso, 0), " jours</b>.")
    couverture <- ifelse(d$ca > 0, round((d$caf / d$ca)*100, 1), 0)
    
    HTML(paste("<ul><li style='margin-bottom:10px;'>", diag_cfo, "</li>",
               "<li style='margin-bottom:10px;'>", diag_dso, "</li>",
               "<li><b>Couverture CAF :</b> La Capacité d'Autofinancement représente ", couverture, "% du chiffre d'affaires.</li></ul>"))
  })
  
  analysis_data <- eventReactive(input$analyser, {
    req(raw_data(), input$selected_column)
    showNotification("Lancement de l'analyse...", type = "message", duration = 3)
    
    custom_words <- str_split(input$custom_stopwords, ",\\s*")[[1]]
    final_stopwords <- bind_rows(all_stopwords_base, tibble(word = custom_words[custom_words != ""])) %>% distinct()
    
    withProgress(message = 'Nettoyage et préparation des données', value = 0, {
      text_data <- raw_data() %>%
        mutate(doc_id = row_number()) %>%
        rename(text_raw = !!sym(input$selected_column)) %>%
        filter(!is.na(text_raw) & text_raw != "") %>%
        mutate(text_clean = text_raw %>%
                 str_to_lower() %>%
                 str_replace_all("[[:punct:]]", " ") %>%
                 str_replace_all("[[:digit:]]", " ") %>%
                 str_squish() %>%
                 str_remove_all("\\b\\w{1,2}\\b"))
      
      incProgress(0.5, detail = "Tokenisation et filtrage...")
      
      tokens_df <- text_data %>%
        select(doc_id, text_clean) %>%
        unnest_tokens(word, text_clean) %>%
        filter(!word %in% final_stopwords$word)
      
      words_freq <- tokens_df %>% count(word, sort = TRUE)
      incProgress(1, detail = "Terminé !")
    })
    return(list(cleaned_data = tokens_df, words_freq = words_freq))
  })
  
  output$cleaned_data_table <- renderDT({
    df <- req(analysis_data())$cleaned_data
    datatable(head(df, 1000), options = list(pageLength = 15, scrollX = TRUE), caption = "1000 premiers mots (tokens) après nettoyage.")
  })
  
  output$freq_table <- renderDT({
    df <- req(analysis_data())$words_freq
    datatable(df, options = list(pageLength = 20, autoWidth = TRUE), caption = "Fréquence de chaque mot dans le corpus.")
  })
  
  output$wordcloud_global <- renderPlot({
    df <- req(analysis_data())$words_freq
    req(nrow(df) > 0)
    set.seed(123)
    wordcloud(words = df$word, freq = df$n, max.words = input$nb_mots_global, random.order = FALSE, colors = brewer.pal(8, "Dark2"), scale = c(3.5, 0.5))
  })
}

# ===================== LANCEMENT DE L'APPLICATION =====================
shinyApp(ui = ui, server = server)