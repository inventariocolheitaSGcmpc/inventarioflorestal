if (!requireNamespace("rsconnect", quietly = TRUE)) {
  stop("Instale o pacote 'rsconnect' antes de publicar.")
}

app_dir <- normalizePath(".", winslash = "/", mustWork = TRUE)

message("Publicando app em shinyapps.io a partir de: ", app_dir)

rsconnect::deployApp(
  appDir = app_dir,
  appFiles = c("app.R", "DESCRIPTION", "data", "www"),
  appName = "site-inventario-florestal",
  forceUpdate = TRUE
)
