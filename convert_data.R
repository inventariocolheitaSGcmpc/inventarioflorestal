library(sf)
library(dplyr)
library(readxl)
library(janitor)
library(jsonlite)

app_dir <- normalizePath(".", winslash = "/", mustWork = TRUE)
data_dir <- file.path(app_dir, "data")
public_data_dir <- file.path(app_dir, "site-data")

dir.create(public_data_dir, showWarnings = FALSE, recursive = TRUE)

inventario_path <- file.path(data_dir, "dados_ultima_med_sao_gabriel.xlsx")
sig_path <- file.path(data_dir, "SG_PRODUTIVO.gpkg")
sig_rds_path <- file.path(data_dir, "SG_PRODUTIVO.rds")

if (!file.exists(inventario_path)) {
  stop("Arquivo de inventario nao encontrado em data/.")
}

if (!file.exists(sig_path)) {
  if (file.exists(sig_rds_path)) {
    sig_path <- sig_rds_path
  } else {
    stop("Arquivo SIG nao encontrado em data/.")
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

if (is.na(st_crs(base_sig_filtro))) {
  st_crs(base_sig_filtro) <- 4326
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

st_write(
  base_info,
  dsn = file.path(public_data_dir, "base_info.geojson"),
  driver = "GeoJSON",
  delete_dsn = TRUE,
  quiet = TRUE
)

projects <- base_info %>%
  st_drop_geometry() %>%
  count(projeto, name = "talhoes") %>%
  arrange(projeto)

summary_data <- list(
  generated_at = format(Sys.time(), "%Y-%m-%d %H:%M:%S"),
  total_talhoes = nrow(base_info),
  total_fazendas = length(unique(base_info$projeto)),
  projects = projects
)

write_json(
  summary_data,
  path = file.path(public_data_dir, "summary.json"),
  auto_unbox = TRUE,
  pretty = TRUE,
  na = "null"
)
