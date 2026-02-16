#!/usr/bin/env Rscript

# ============================================================
# Script: Análisis descriptivo + confiabilidad + CFA
# Proyecto: Relevo Generacional
# ============================================================

suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(tidyr)
  library(forcats)
  library(stringr)
  library(ggplot2)
  library(scales)
  library(purrr)
  library(psych)
  library(lavaan)
  library(officer)
  library(flextable)
})

# ----------------------------
# 1) Parámetros de entrada/salida
# ----------------------------

# Ruta entregada por el usuario
input_file <- "C:/Users/johna/Desktop/Research/Relevo Generacional/data_cleaned_feb2026.xlsx"

# Carpeta de salida (no muestra gráficos en interfaz; solo guarda archivos)
output_dir <- "resultados_relevo_generacional"
plots_dir <- file.path(output_dir, "plots")

dir.create(output_dir, recursive = TRUE, showWarnings = FALSE)
dir.create(plots_dir, recursive = TRUE, showWarnings = FALSE)

# ----------------------------
# 2) Configuración del instrumento
# ----------------------------

demograficas <- c("sexo", "rural", "semestre", "carrera", "universidad")
items <- paste0("relevo_", 1:27)

dimensiones <- list(
  "D1_Oportunidades_laborales_salario" = paste0("relevo_", 1:3),
  "D2_Calidad_vida_servicios" = paste0("relevo_", 4:6),
  "D3_Futuro_crecimiento_personal" = paste0("relevo_", 7:9),
  "D4_Infraestructura_comunicaciones" = paste0("relevo_", 10:12),
  "D5_Bienestar_cultura" = paste0("relevo_", 13:15),
  "D6_Incentivos_economicos_finca" = paste0("relevo_", 16:18),
  "D7_Incentivos_no_economicos_finca" = paste0("relevo_", 19:21),
  "D8_Derechos_propiedad_control" = paste0("relevo_", 22:24),
  "D9_Apoyo_programas_institucionales" = paste0("relevo_", 25:27)
)

etiquetas_dimensiones <- c(
  D1_Oportunidades_laborales_salario = "D1: Oportunidades laborales y salario",
  D2_Calidad_vida_servicios = "D2: Calidad de vida y servicios básicos",
  D3_Futuro_crecimiento_personal = "D3: Percepción del futuro y crecimiento personal",
  D4_Infraestructura_comunicaciones = "D4: Infraestructura y comunicaciones",
  D5_Bienestar_cultura = "D5: Bienestar y cultura",
  D6_Incentivos_economicos_finca = "D6: Incentivos económicos en fincas familiares",
  D7_Incentivos_no_economicos_finca = "D7: Incentivos no económicos en la finca familiar",
  D8_Derechos_propiedad_control = "D8: Derechos de propiedad y control",
  D9_Apoyo_programas_institucionales = "D9: Apoyo de programas institucionales"
)

# ----------------------------
# 3) Funciones auxiliares
# ----------------------------

normalizar_rural <- function(x) {
  x_chr <- as.character(x)
  x_low <- tolower(trimws(x_chr))

  case_when(
    x_low %in% c("1", "si", "sí", "s", "true", "verdadero", "rural") ~ "Rural",
    x_low %in% c("0", "no", "n", "false", "falso", "urbano", "no rural") ~ "No rural",
    is.na(x) | x_low %in% c("", "na", "nan") ~ NA_character_,
    TRUE ~ str_to_title(x_chr)
  )
}

freq_table <- function(data, var, group_var = "grupo") {
  data %>%
    mutate(valor = .data[[var]]) %>%
    mutate(valor = ifelse(is.na(valor), "No responde", as.character(valor))) %>%
    count(.data[[group_var]], valor, name = "n") %>%
    group_by(.data[[group_var]]) %>%
    mutate(porcentaje = n / sum(n)) %>%
    ungroup() %>%
    rename(grupo = .data[[group_var]]) %>%
    mutate(variable = var) %>%
    select(variable, grupo, valor, n, porcentaje)
}

save_plot <- function(p, filename, width = 11, height = 7) {
  ggsave(
    filename = file.path(plots_dir, filename),
    plot = p,
    width = width,
    height = height,
    dpi = 320,
    bg = "white"
  )
}

fmt_n_pct <- function(n, pct) {
  paste0(n, " (", percent(pct, accuracy = 0.1), ")")
}

# ----------------------------
# 4) Carga y validaciones de datos
# ----------------------------

if (!file.exists(input_file)) {
  stop(paste0(
    "No se encontró el archivo de entrada en: ", input_file, "\n",
    "Actualiza la variable 'input_file' al path correcto."
  ))
}

raw <- read_excel(input_file)

faltantes <- setdiff(c(demograficas, items), names(raw))
if (length(faltantes) > 0) {
  stop(paste0("Faltan columnas requeridas en la base: ", paste(faltantes, collapse = ", ")))
}

df <- raw %>%
  mutate(
    rural = normalizar_rural(rural),
    grupo = case_when(
      is.na(rural) ~ "No especifica",
      TRUE ~ rural
    )
  )

# Asegurar codificación numérica en items para confiabilidad/CFA
for (it in items) {
  df[[it]] <- as.numeric(df[[it]])
}

# ----------------------------
# 5) Tabla 1: sociodemográficos (frecuencia absoluta y relativa)
# ----------------------------

tabla1 <- map_dfr(demograficas, ~freq_table(df, .x, "grupo"))

tabla1_total <- map_dfr(demograficas, function(v) {
  df %>%
    mutate(valor = .data[[v]]) %>%
    mutate(valor = ifelse(is.na(valor), "No responde", as.character(valor))) %>%
    count(valor, name = "n") %>%
    mutate(
      porcentaje = n / sum(n),
      variable = v,
      grupo = "Total"
    ) %>%
    select(variable, grupo, valor, n, porcentaje)
})

tabla1_largo <- bind_rows(tabla1, tabla1_total)

tabla1_final <- tabla1_largo %>%
  mutate(
    porcentaje = percent(porcentaje, accuracy = 0.1),
    valor = as.character(valor)
  ) %>%
  arrange(variable, factor(grupo, levels = c("Rural", "No rural", "No especifica", "Total")), desc(n))

write.csv(
  tabla1_final,
  file.path(output_dir, "tabla1_sociodemograficos.csv"),
  row.names = FALSE,
  fileEncoding = "UTF-8"
)

# Table 1 visible: columnas Rural / No rural / Total con n (p%)
tabla1_wide <- tabla1_largo %>%
  filter(grupo %in% c("Rural", "No rural", "Total")) %>%
  mutate(valor_np = fmt_n_pct(n, porcentaje)) %>%
  select(variable, categoria = valor, grupo, valor_np) %>%
  pivot_wider(names_from = grupo, values_from = valor_np) %>%
  mutate(variable = str_to_title(variable)) %>%
  select(Variable = variable, Categoría = categoria, Rural, `No rural`, Total)

write.csv(
  tabla1_wide,
  file.path(output_dir, "tabla1_sociodemograficos_visible.csv"),
  row.names = FALSE,
  fileEncoding = "UTF-8"
)

# Plot sociodemográficos (etiquetas pequeñas y sin sobreposición)
plot_demo_df <- tabla1_largo %>%
  mutate(
    variable = str_to_title(variable),
    grupo = factor(grupo, levels = c("Rural", "No rural", "No especifica", "Total"))
  )

p_demo <- ggplot(plot_demo_df, aes(x = fct_reorder(valor, n, .desc = TRUE), y = porcentaje, fill = grupo)) +
  geom_col(position = position_dodge2(width = 0.8, preserve = "single", padding = 0.2), width = 0.68, alpha = 0.95) +
  geom_text(
    aes(label = ifelse(porcentaje < 0.03, "", percent(porcentaje, accuracy = 0.1))),
    position = position_dodge2(width = 0.8, preserve = "single", padding = 0.2),
    vjust = -0.25,
    size = 2.2,
    check_overlap = TRUE
  ) +
  facet_wrap(~variable, scales = "free_x") +
  scale_y_continuous(labels = percent_format(accuracy = 1), expand = expansion(mult = c(0, 0.15))) +
  scale_fill_brewer(type = "qual", palette = "Set2") +
  labs(
    title = "Perfil sociodemográfico de la muestra",
    subtitle = "Frecuencia relativa por categoría: Rural, No rural y Total",
    x = "Categoría",
    y = "Porcentaje",
    fill = "Grupo"
  ) +
  theme_minimal(base_size = 12) +
  theme(
    legend.position = "top",
    axis.text.x = element_text(angle = 35, hjust = 1),
    panel.grid.minor = element_blank(),
    plot.title = element_text(face = "bold")
  )

save_plot(p_demo, "plot_sociodemograficos.png", width = 13, height = 8)

# ----------------------------
# 6) Descriptivos por dimensión (frecuencia de respuestas)
# ----------------------------

long_items <- map_dfr(names(dimensiones), function(dim_name) {
  df %>%
    select(grupo, all_of(dimensiones[[dim_name]])) %>%
    pivot_longer(
      cols = all_of(dimensiones[[dim_name]]),
      names_to = "item",
      values_to = "respuesta"
    ) %>%
    mutate(dimension = etiquetas_dimensiones[[dim_name]])
})

long_items <- long_items %>%
  mutate(
    respuesta = ifelse(is.na(respuesta), "No responde", as.character(respuesta))
  )

freq_dim_grupo <- long_items %>%
  count(dimension, grupo, respuesta, name = "n") %>%
  group_by(dimension, grupo) %>%
  mutate(porcentaje = n / sum(n)) %>%
  ungroup()

freq_dim_total <- long_items %>%
  count(dimension, respuesta, name = "n") %>%
  group_by(dimension) %>%
  mutate(
    porcentaje = n / sum(n),
    grupo = "Total"
  ) %>%
  ungroup()

freq_dim_final <- bind_rows(freq_dim_grupo, freq_dim_total) %>%
  mutate(
    grupo = factor(grupo, levels = c("Rural", "No rural", "No especifica", "Total")),
    porcentaje_fmt = percent(porcentaje, accuracy = 0.1)
  ) %>%
  arrange(dimension, grupo, respuesta)

write.csv(
  freq_dim_final %>% select(dimension, grupo, respuesta, n, porcentaje_fmt),
  file.path(output_dir, "descriptivos_dimensiones_frecuencias.csv"),
  row.names = FALSE,
  fileEncoding = "UTF-8"
)

# Tabla visible por dimensión: n (p%)
freq_dim_wide <- freq_dim_final %>%
  filter(grupo %in% c("Rural", "No rural", "Total")) %>%
  mutate(valor_np = fmt_n_pct(n, porcentaje)) %>%
  select(dimension, respuesta, grupo, valor_np) %>%
  pivot_wider(names_from = grupo, values_from = valor_np) %>%
  select(Dimensión = dimension, Respuesta = respuesta, Rural, `No rural`, Total)

write.csv(
  freq_dim_wide,
  file.path(output_dir, "descriptivos_dimensiones_visible.csv"),
  row.names = FALSE,
  fileEncoding = "UTF-8"
)

# Plot dimensiones (etiquetas pequeñas y sin sobreposición)
p_dim <- ggplot(freq_dim_final, aes(x = respuesta, y = porcentaje, fill = grupo)) +
  geom_col(position = position_dodge2(width = 0.85, preserve = "single", padding = 0.15), width = 0.7) +
  geom_text(
    aes(label = ifelse(porcentaje < 0.03, "", percent(porcentaje, accuracy = 0.1))),
    position = position_dodge2(width = 0.85, preserve = "single", padding = 0.15),
    vjust = -0.25,
    size = 2.1,
    check_overlap = TRUE
  ) +
  facet_wrap(~dimension, scales = "free_x") +
  scale_y_continuous(labels = percent_format(accuracy = 1), expand = expansion(mult = c(0, 0.18))) +
  scale_fill_brewer(type = "qual", palette = "Set2") +
  labs(
    title = "Distribución de respuestas por dimensión del instrumento",
    subtitle = "Comparación entre Rural, No rural y Total",
    x = "Respuesta",
    y = "Porcentaje",
    fill = "Grupo"
  ) +
  theme_minimal(base_size = 12) +
  theme(
    legend.position = "top",
    axis.text.x = element_text(angle = 35, hjust = 1),
    strip.text = element_text(face = "bold", size = 9),
    panel.grid.minor = element_blank(),
    plot.title = element_text(face = "bold")
  )

save_plot(p_dim, "plot_dimensiones_frecuencias.png", width = 15, height = 10)

# ----------------------------
# 7) Confiabilidad: alfa y omega (por dimensión y total)
# ----------------------------

reliability_by_dimension <- map_dfr(names(dimensiones), function(dim_name) {
  datos_dim <- df %>% select(all_of(dimensiones[[dim_name]])) %>% drop_na()

  if (nrow(datos_dim) < 10) {
    return(tibble(
      dimension = etiquetas_dimensiones[[dim_name]],
      n_valido = nrow(datos_dim),
      alpha = NA_real_,
      omega_total = NA_real_
    ))
  }

  a <- psych::alpha(datos_dim, warnings = FALSE, check.keys = FALSE)
  o <- suppressWarnings(psych::omega(datos_dim, nfactors = 1, plot = FALSE, warnings = FALSE))

  tibble(
    dimension = etiquetas_dimensiones[[dim_name]],
    n_valido = nrow(datos_dim),
    alpha = unname(a$total$raw_alpha),
    omega_total = unname(o$omega.tot)
  )
})

# Total: usar omega total con estructura multidimensional (9 factores)
datos_total <- df %>% select(all_of(items)) %>% drop_na()

alpha_total <- psych::alpha(datos_total, warnings = FALSE, check.keys = FALSE)
omega_total_obj <- suppressWarnings(psych::omega(datos_total, nfactors = 9, plot = FALSE, warnings = FALSE))

reliability_total <- tibble(
  dimension = "Instrumento total (27 ítems)",
  n_valido = nrow(datos_total),
  alpha = unname(alpha_total$total$raw_alpha),
  omega_total = unname(omega_total_obj$omega.tot),
  omega_hierarchical = unname(omega_total_obj$omega_h)
)

reliability_table <- bind_rows(reliability_by_dimension, reliability_total) %>%
  mutate(
    across(c(alpha, omega_total, omega_hierarchical), ~round(., 3))
  )

write.csv(
  reliability_table,
  file.path(output_dir, "confiabilidad_alpha_omega.csv"),
  row.names = FALSE,
  fileEncoding = "UTF-8"
)

# ----------------------------
# 8) CFA de 9 dimensiones
# ----------------------------

modelo_cfa <- '
D1 =~ relevo_1 + relevo_2 + relevo_3
D2 =~ relevo_4 + relevo_5 + relevo_6
D3 =~ relevo_7 + relevo_8 + relevo_9
D4 =~ relevo_10 + relevo_11 + relevo_12
D5 =~ relevo_13 + relevo_14 + relevo_15
D6 =~ relevo_16 + relevo_17 + relevo_18
D7 =~ relevo_19 + relevo_20 + relevo_21
D8 =~ relevo_22 + relevo_23 + relevo_24
D9 =~ relevo_25 + relevo_26 + relevo_27
'

fit_cfa <- lavaan::cfa(
  model = modelo_cfa,
  data = df,
  ordered = items,
  estimator = "WLSMV",
  std.lv = TRUE,
  missing = "pairwise"
)

indices_ajuste <- tibble(
  Indicador = c("CFI", "TLI", "RMSEA", "SRMR", "Chi-cuadrado", "gl", "p-valor"),
  Valor = c(
    fitMeasures(fit_cfa, "cfi"),
    fitMeasures(fit_cfa, "tli"),
    fitMeasures(fit_cfa, "rmsea"),
    fitMeasures(fit_cfa, "srmr"),
    fitMeasures(fit_cfa, "chisq"),
    fitMeasures(fit_cfa, "df"),
    fitMeasures(fit_cfa, "pvalue")
  )
) %>%
  mutate(Valor = round(Valor, 4))

cargas <- standardizedSolution(fit_cfa) %>%
  as_tibble() %>%
  filter(op == "=~") %>%
  transmute(
    Dimension = lhs,
    Item = rhs,
    Carga_estandarizada = round(est.std, 3),
    Error_estandar = round(se, 3),
    p_valor = round(pvalue, 4)
  )

write.csv(indices_ajuste, file.path(output_dir, "cfa_indices_ajuste.csv"), row.names = FALSE, fileEncoding = "UTF-8")
write.csv(cargas, file.path(output_dir, "cfa_cargas_factoriales.csv"), row.names = FALSE, fileEncoding = "UTF-8")

# ----------------------------
# 9) Reportes Word
# ----------------------------

# 9.1 Reporte descriptivo visible
doc_desc <- read_docx() %>%
  body_add_par("Reporte descriptivo - Relevo generacional", style = "heading 1") %>%
  body_add_par(format(Sys.time(), "%Y-%m-%d %H:%M"), style = "Normal") %>%
  body_add_par("", style = "Normal") %>%
  body_add_par("1) Table 1 sociodemográficos", style = "heading 2") %>%
  body_add_par("Formato de celda: frecuencia absoluta (frecuencia relativa).", style = "Normal")

ft_tabla1 <- flextable(tabla1_wide) %>% theme_booktabs() %>% autofit()

doc_desc <- doc_desc %>%
  body_add_flextable(ft_tabla1) %>%
  body_add_par("", style = "Normal") %>%
  body_add_par("2) Frecuencias por dimensión del instrumento", style = "heading 2") %>%
  body_add_par("Formato de celda: frecuencia absoluta (frecuencia relativa).", style = "Normal")

ft_dims <- flextable(freq_dim_wide) %>% theme_booktabs() %>% autofit()

doc_desc <- doc_desc %>%
  body_add_flextable(ft_dims)

print(doc_desc, target = file.path(output_dir, "reporte_descriptivos_relevo.docx"))

# 9.2 Reporte psicométrico
doc <- read_docx() %>%
  body_add_par("Reporte psicométrico - Instrumento de relevo generacional", style = "heading 1") %>%
  body_add_par(format(Sys.time(), "%Y-%m-%d %H:%M"), style = "Normal") %>%
  body_add_par("", style = "Normal") %>%
  body_add_par("1) Confiabilidad (alfa de Cronbach y omega)", style = "heading 2") %>%
  body_add_par(
    "Para cada dimensión se reporta alfa y omega total. Para el instrumento total se incluye omega total y omega jerárquico (apropiado en escalas multidimensionales).",
    style = "Normal"
  )

ft_rel <- flextable(reliability_table) %>%
  theme_booktabs() %>%
  autofit()

doc <- doc %>%
  body_add_flextable(ft_rel) %>%
  body_add_par("", style = "Normal") %>%
  body_add_par("2) Modelo de Análisis Factorial Confirmatorio (9 dimensiones)", style = "heading 2") %>%
  body_add_par("Indicadores de ajuste global del modelo:", style = "Normal")

ft_fit <- flextable(indices_ajuste) %>%
  theme_booktabs() %>%
  autofit()

doc <- doc %>%
  body_add_flextable(ft_fit) %>%
  body_add_par("", style = "Normal") %>%
  body_add_par("Cargas factoriales estandarizadas por ítem:", style = "Normal")

ft_load <- flextable(cargas) %>%
  theme_booktabs() %>%
  autofit()

doc <- doc %>%
  body_add_flextable(ft_load)

print(doc, target = file.path(output_dir, "reporte_psicometrico_relevo.docx"))

message("Análisis finalizado correctamente. Archivos generados en: ", normalizePath(output_dir))
