library(shiny)
library(leaflet)
library(sf)
library(dplyr)
library(readxl)
library(janitor)
library(base64enc)
library(DT)
library(writexl)
library(blastula)
library(webshot2)
library(htmlwidgets)

`%||%` <- function(x, y) {
  if (is.null(x) || length(x) == 0 || identical(x, "")) y else x
}

app_dir <- "."
data_dir <- file.path(app_dir, "data")
logo_path <- file.path(app_dir, "www", "logo_cmpc.png")

inventario_path <- file.path(data_dir, "dados_ultima_med_sao_gabriel.xlsx")
sig_path <- file.path(data_dir, "SG_PRODUTIVO.gpkg")
sig_rds_path <- file.path(data_dir, "SG_PRODUTIVO.rds")

email_remetente <- Sys.getenv("SHINY_EMAIL_FROM", unset = "")
email_senha <- Sys.getenv("SHINY_EMAIL_PASSWORD", unset = "")
email_destinatarios <- Sys.getenv(
  "SHINY_EMAIL_TO",
  unset = "pietro.duran@cmpc.com,matheus.roberto@cmpc.com"
)

destinatarios <- strsplit(email_destinatarios, "\\s*,\\s*")[[1]]
destinatarios <- destinatarios[nzchar(destinatarios)]

if (!file.exists(inventario_path)) {
  stop("Erro critico: o arquivo de inventario nao foi encontrado em 'data/'.")
}

if (!file.exists(sig_path)) {
  if (file.exists(sig_rds_path)) {
    sig_path <- sig_rds_path
  } else {
    stop("Erro critico: o arquivo SIG nao foi encontrado em 'data/'.")
  }
}

base_inventario <- read_excel(inventario_path, sheet = "Planilha3") %>%
  clean_names()

if (grepl("\\.gpkg$", sig_path, ignore.case = TRUE)) {
  base_sig <- st_read(sig_path, quiet = TRUE) %>%
    clean_names()
} else {
  base_sig <- readRDS(sig_path) %>%
    clean_names()
}

base_sig <- st_make_valid(base_sig)

suppressWarnings({
  base_sig <- base_sig %>%
    mutate(data_plant = as.Date(data_plant))

  base_inventario <- base_inventario %>%
    mutate(
      produtividade = as.numeric(produtividade),
      vlr_area_gis = as.numeric(vlr_area_gis),
      idade_inteira = as.numeric(idade_inteira),
      n = as.numeric(n),
      ht = as.numeric(ht),
      dap = as.numeric(dap),
      mg = as.numeric(mg),
      ab = as.numeric(ab)
    )
})

base_sig_filtro <- base_sig %>%
  dplyr::select(
    id_regiao,
    id_projeto,
    projeto,
    cd_talhao,
    caracteris,
    data_plant,
    espacament,
    geom
  ) %>%
  filter(
    caracteris == "Plantio Comercial",
    !is.na(data_plant),
    data_plant < as.Date("2025-01-01")
  )

if (is.na(sf::st_crs(base_sig_filtro))) {
  sf::st_crs(base_sig_filtro) <- 4326
}

base_sig_new_wgs84 <- st_transform(base_sig_filtro, crs = 4326)

base_inventario <- base_inventario %>%
  mutate(id = paste(id_projeto, cd_talhao, sep = "_"))

base_sig_new_wgs84 <- base_sig_new_wgs84 %>%
  mutate(id = paste(id_projeto, cd_talhao, sep = "_"))

base_info <- base_sig_new_wgs84 %>%
  left_join(base_inventario, by = "id") %>%
  filter(!is.na(produtividade)) %>%
  dplyr::select(
    id_projeto.x,
    cd_talhao.x,
    projeto,
    produtividade,
    ab,
    vlr_area_gis,
    idade_inteira,
    ht,
    dap,
    n,
    mg,
    geom
  ) %>%
  rename(
    id_projeto = id_projeto.x,
    cd_talhao = cd_talhao.x
  ) %>%
  mutate(
    pro_tlh = paste(id_projeto, cd_talhao, sep = "_"),
    vmi = round(produtividade / n, 2),
    vcsc = round(produtividade * vlr_area_gis, 2),
    ima = round(produtividade / idade_inteira, 2)
  )

make_pal <- function(values) {
  colorNumeric(palette = "RdYlGn", domain = values, reverse = FALSE)
}

formata_br <- function(x) {
  format(round(x, 2), nsmall = 2, big.mark = ".", decimal.mark = ",")
}

gerar_tabela_exportacao <- function(ids, dados) {
  dados %>%
    st_drop_geometry() %>%
    filter(pro_tlh %in% ids) %>%
    slice(match(ids, pro_tlh)) %>%
    mutate(ordem_sequencia = row_number()) %>%
    dplyr::select(
      `Ordem sequencia` = ordem_sequencia,
      `Fazenda/Talhao` = pro_tlh,
      `Area (ha)` = vlr_area_gis,
      `VCSC (m3)` = vcsc,
      `VMI (m3)` = vmi
    )
}

capturar_mapa <- function(dados_sel, arquivo_saida) {
  if (nrow(dados_sel) == 0) {
    return(FALSE)
  }

  pal_print <- make_pal(base_info$produtividade)
  mapa_print <- leaflet(dados_sel) %>%
    addTiles() %>%
    addProviderTiles(providers$Esri.WorldImagery) %>%
    addPolygons(
      color = "black",
      weight = 2,
      fillColor = ~pal_print(produtividade),
      fillOpacity = 0.7
    )

  bbox <- st_bbox(dados_sel)
  if (all(is.finite(bbox))) {
    mapa_print <- mapa_print %>%
      fitBounds(bbox[["xmin"]], bbox[["ymin"]], bbox[["xmax"]], bbox[["ymax"]])
  }

  html_temp <- tempfile(fileext = ".html")
  saveWidget(mapa_print, html_temp, selfcontained = FALSE)

  tryCatch({
    webshot2::webshot(
      url = html_temp,
      file = arquivo_saida,
      delay = 5,
      vwidth = 1000,
      vheight = 600
    )
    file.exists(arquivo_saida)
  }, error = function(e) {
    FALSE
  })
}

email_habilitado <- function() {
  nzchar(email_remetente) && nzchar(email_senha) && length(destinatarios) > 0
}

ui <- fluidPage(
  tags$head(
    tags$style(HTML("
      .title-container { display: flex; align-items: center; justify-content: space-between; padding: 10px 20px; border-bottom: 2px solid #000; background-color: #fcfcfc; }
      .main-title { color: black !important; font-weight: bold; margin: 0; font-size: 30px; }
      .title-logo { max-height: 100px; width: auto; }
      #tabela_sequencia { margin-top: 10px; font-size: 11px; }
      .status-email { margin-top: 10px; font-size: 12px; color: #666; }
    "))
  ),
  div(
    class = "title-container",
    h1("Inventario Florestal", class = "main-title"),
    uiOutput("logo_renderizado")
  ),
  fluidRow(
    column(
      width = 4,
      wellPanel(
        selectInput(
          "filtro_fazenda",
          "Selecionar Fazenda:",
          choices = c("Todas", sort(unique(base_info$projeto))),
          selected = "Todas"
        ),
        actionButton(
          "limpar_selecao",
          "Limpar Selecao",
          icon = icon("eraser"),
          style = "width: 100%; margin-bottom: 10px;"
        ),
        actionButton(
          "processar_envio",
          "Enviar Sequencia por E-mail",
          icon = icon("paper-plane"),
          style = "width: 100%; margin-bottom: 15px; background-color: #28a745; color: white; font-weight: bold;"
        ),
        div(
          class = "status-email",
          if (email_habilitado()) {
            "Envio de e-mail configurado."
          } else {
            "Envio de e-mail desabilitado ate configurar as variaveis de ambiente."
          }
        ),
        uiOutput("card_info"),
        hr(),
        h4(strong("Sequencia talhonar de colheita")),
        DTOutput("tabela_sequencia")
      )
    ),
    column(width = 8, leafletOutput("mapa", height = "800px"))
  )
)

server <- function(input, output, session) {
  output$logo_renderizado <- renderUI({
    if (file.exists(logo_path)) {
      tags$img(
        src = base64enc::dataURI(file = logo_path, mime = "image/png"),
        class = "title-logo"
      )
    }
  })

  ids_selecionados <- reactiveVal(character(0))

  output$mapa <- renderLeaflet({
    pal <- make_pal(base_info$produtividade)

    leaflet(base_info, options = leafletOptions(minZoom = 7, maxZoom = 18)) %>%
      addTiles() %>%
      addProviderTiles(providers$Esri.WorldImagery) %>%
      addLegend(
        pal = pal,
        values = ~produtividade,
        position = "bottomright",
        title = "VCSC (m3/ha)"
      ) %>%
      setView(lng = -54.0, lat = -30.3, zoom = 9)
  })

  observe({
    df <- base_info
    if (input$filtro_fazenda != "Todas") {
      df <- df %>% filter(projeto == input$filtro_fazenda)
    }

    proxy <- leafletProxy("mapa")
    proxy %>%
      clearShapes() %>%
      clearMarkers() %>%
      clearGroup("labels_talhao")

    if (nrow(df) > 0) {
      pal <- make_pal(base_info$produtividade)

      proxy %>%
        addPolygons(
          data = df,
          layerId = ~pro_tlh,
          fillColor = ~pal(produtividade),
          color = "white",
          weight = 1,
          fillOpacity = 0.7,
          label = ~paste("Talhao:", cd_talhao),
          highlightOptions = highlightOptions(
            weight = 3,
            color = "#FFFF00",
            fillOpacity = 0.9,
            bringToFront = TRUE
          )
        )

      if (input$filtro_fazenda != "Todas") {
        bbox <- st_bbox(df)
        proxy %>%
          flyToBounds(
            bbox[["xmin"]],
            bbox[["ymin"]],
            bbox[["xmax"]],
            bbox[["ymax"]]
          )

        proxy %>%
          addLabelOnlyMarkers(
            data = st_point_on_surface(df),
            label = ~as.character(cd_talhao),
            group = "labels_talhao",
            labelOptions = labelOptions(
              noHide = TRUE,
              direction = "bottom",
              textOnly = TRUE,
              style = list(
                "color" = "#d9534f",
                "font-weight" = "bold",
                "font-size" = "11px",
                "text-shadow" = "1px 1px 2px white"
              )
            )
          )
      }
    }
  })

  observeEvent(input$mapa_shape_click, {
    id <- input$mapa_shape_click$id
    if (is.null(id) || length(id) == 0) {
      return()
    }

    atual <- ids_selecionados()
    if (id %in% atual) {
      ids_selecionados(setdiff(atual, id))
    } else {
      ids_selecionados(c(atual, id))
    }
  })

  observeEvent(input$limpar_selecao, {
    ids_selecionados(character(0))
  })

  observe({
    proxy <- leafletProxy("mapa")
    proxy %>%
      clearGroup("destaque") %>%
      clearGroup("setas") %>%
      clearGroup("direcao")

    sel <- ids_selecionados()
    if (length(sel) == 0) {
      return()
    }

    dados_destaque <- base_info %>%
      filter(pro_tlh %in% sel) %>%
      slice(match(sel, pro_tlh))

    if (nrow(dados_destaque) == 0) {
      return()
    }

    proxy %>%
      addPolygons(
        data = dados_destaque,
        group = "destaque",
        fill = FALSE,
        color = "#222",
        weight = 4,
        opacity = 1
      )

    if (nrow(dados_destaque) > 1) {
      coords <- st_coordinates(st_point_on_surface(dados_destaque))

      proxy %>%
        addPolylines(
          lng = coords[, 1],
          lat = coords[, 2],
          group = "setas",
          color = "#FFFF00",
          weight = 5,
          opacity = 1,
          dashArray = "8, 12"
        )

      for (i in seq_len(nrow(coords))) {
        proxy %>%
          addLabelOnlyMarkers(
            lng = coords[i, 1],
            lat = coords[i, 2],
            group = "direcao",
            label = as.character(i),
            labelOptions = labelOptions(
              noHide = TRUE,
              direction = "top",
              textOnly = FALSE,
              style = list(
                "color" = "black",
                "font-weight" = "bold",
                "font-size" = "12px",
                "background-color" = "white",
                "border" = "2px solid #222",
                "border-radius" = "4px",
                "padding" = "2px 6px"
              )
            )
          )
      }
    }
  })

  observeEvent(input$processar_envio, {
    req(length(ids_selecionados()) > 0)

    if (!email_habilitado()) {
      showModal(
        modalDialog(
          title = "Envio indisponivel",
          "Configure SHINY_EMAIL_FROM, SHINY_EMAIL_PASSWORD e SHINY_EMAIL_TO antes de usar o envio por e-mail.",
          easyClose = TRUE
        )
      )
      return()
    }

    nome_fazenda_limpo <- if (input$filtro_fazenda == "Todas") {
      "Geral"
    } else {
      gsub("[^A-Za-z0-9_-]", "_", input$filtro_fazenda)
    }

    nome_dinamico_xlsx <- paste0("HF_", nome_fazenda_limpo, "_Sequencia_Talhonar.xlsx")
    nome_dinamico_png <- paste0("HF_", nome_fazenda_limpo, "_Mapa_Logistico.png")

    showNotification("Preparando arquivos...", type = "message", id = "envio", duration = NULL)

    tmp_excel <- tempfile(fileext = ".xlsx")
    tmp_png <- tempfile(fileext = ".png")

    df_export <- gerar_tabela_exportacao(ids_selecionados(), base_info)
    writexl::write_xlsx(df_export, path = tmp_excel)

    dados_sel <- base_info %>%
      filter(pro_tlh %in% ids_selecionados()) %>%
      slice(match(ids_selecionados(), pro_tlh))

    mapa_anexado <- capturar_mapa(dados_sel, tmp_png)

    tryCatch({
      corpo_email <- compose_email(
        body = md(
          paste0(
            "Ola,<br><br>Sequencia planejada para a fazenda **",
            input$filtro_fazenda,
            "**."
          )
        )
      )

      corpo_email <- corpo_email %>%
        add_attachment(file = tmp_excel, filename = nome_dinamico_xlsx)

      if (mapa_anexado) {
        corpo_email <- corpo_email %>%
          add_attachment(file = tmp_png, filename = nome_dinamico_png)
      }

      blastula::smtp_send(
        email = corpo_email,
        from = email_remetente,
        to = destinatarios,
        subject = paste("Sequencia -", input$filtro_fazenda),
        credentials = blastula::creds(
          user = email_remetente,
          provider = "gmail",
          host = "smtp.gmail.com",
          port = 587,
          use_ssl = TRUE
        ),
        password = email_senha
      )

      removeNotification("envio")
      showModal(modalDialog(title = "Sucesso", "E-mail enviado.", easyClose = TRUE))
    }, error = function(e) {
      removeNotification("envio")
      showModal(modalDialog(title = "Falha no envio", paste("Erro:", e$message)))
    })
  })

  output$tabela_sequencia <- renderDT({
    df_tab <- gerar_tabela_exportacao(ids_selecionados(), base_info)
    datatable(
      df_tab,
      options = list(dom = "t", pageLength = -1, ordering = FALSE),
      rownames = FALSE
    )
  })

  output$card_info <- renderUI({
    ids <- ids_selecionados()
    if (length(ids) == 0) {
      return(div(style = "color: #777; text-align: center;", "Nenhum talhao selecionado."))
    }

    resumo_dados <- base_info %>% filter(pro_tlh %in% ids)
    total_area <- sum(resumo_dados$vlr_area_gis, na.rm = TRUE)
    total_vcsc <- sum(resumo_dados$vcsc, na.rm = TRUE)

    div(
      style = "background-color:#ffffff; padding:15px; border-radius:8px; border: 1px solid #ddd;",
      p(strong("Talhoes selecionados: "), length(ids)),
      p(strong("Area total: "), formata_br(total_area), " ha"),
      hr(style = "margin: 10px 0;"),
      div(
        style = "text-align: center; background-color: #fcfcfc; padding: 10px; border: 1px solid #eee;",
        span("Volume Comercial (VCSC)"),
        span(
          style = "font-size: 1.5em; font-weight: bold; display: block;",
          formata_br(total_vcsc),
          " m3"
        )
      )
    )
  })
}

shinyApp(ui, server)
