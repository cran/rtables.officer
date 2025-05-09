## ----include = FALSE----------------------------------------------------------
suggested_dependent_pkgs <- c("dplyr", "tern")
knitr::opts_chunk$set(
  collapse = TRUE,
  comment = "#>",
  eval = all(vapply(
    suggested_dependent_pkgs,
    requireNamespace,
    logical(1),
    quietly = TRUE
  ))
)

## ----echo=FALSE---------------------------------------------------------------
knitr::opts_chunk$set(comment = "#")

## -----------------------------------------------------------------------------
library(tern)
library(dplyr)
library(rtables.officer)

# Load example datasets
adsl <- formatters::ex_adae
adlb <- formatters::ex_adlb

# Convert character variables to factors and handle missing levels
adsl <- df_explicit_na(adsl)
adlb <- df_explicit_na(adlb)

# Create a temporary file for the output
tf <- tempfile(fileext = ".docx")

## -----------------------------------------------------------------------------
adlb_f <- adlb %>%
  dplyr::filter(
    PARAM %in% c("Alanine Aminotransferase Measurement", "C-Reactive Protein Measurement") &
      !(ACTARM == "B: Placebo" & AVISIT == "WEEK 1 DAY 8") &
      AVISIT != "SCREENING"
  )

## -----------------------------------------------------------------------------
afun <- function(x, .var, .spl_context, ...) {
  n_fun <- sum(!is.na(x), na.rm = TRUE)
  mean_sd_fun <- if (n_fun == 0) c(NA, NA) else c(mean(x, na.rm = TRUE), sd(x, na.rm = TRUE))
  median_fun <- if (n_fun == 0) NA else median(x, na.rm = TRUE)
  min_max_fun <- if (n_fun == 0) c(NA, NA) else c(min(x), max(x))

  is_chg <- .var == "CHG"
  is_baseline <- .spl_context$value[which(.spl_context$split == "AVISIT")] == "BASELINE"
  if (is_baseline && is_chg) n_fun <- mean_sd_fun <- median_fun <- min_max_fun <- NULL

  in_rows(
    "n" = n_fun,
    "Mean (SD)" = mean_sd_fun,
    "Median" = median_fun,
    "Min - Max" = min_max_fun,
    .formats = list("n" = "xx", "Mean (SD)" = "xx.xx (xx.xx)", "Median" = "xx.xx", "Min - Max" = "xx.xx - xx.xx"),
    .format_na_strs = list("n" = "NE", "Mean (SD)" = "NE (NE)", "Median" = "NE", "Min - Max" = "NE - NE")
  )
}

## -----------------------------------------------------------------------------
lyt <- basic_table() %>%
  split_cols_by("ACTARM", show_colcounts = TRUE, split_fun = keep_split_levels(levels(adlb_f$ACTARM)[c(1, 2)])) %>%
  split_rows_by("PARAM",
    split_fun = drop_split_levels, label_pos = "topleft",
    split_label = obj_label(adlb_f$PARAM), page_by = TRUE
  ) %>%
  split_rows_by("AVISIT",
    split_fun = drop_split_levels, label_pos = "topleft",
    split_label = obj_label(adlb_f$AVISIT)
  ) %>%
  split_cols_by_multivar(
    vars = c("AVAL", "CHG"),
    varlabels = c("Value at Visit", "Change from Baseline")
  ) %>%
  analyze_colvars(afun = afun)

## -----------------------------------------------------------------------------
result <- build_table(lyt, adlb_f)
result

## -----------------------------------------------------------------------------
main_title(result) <- "Alanine Aminotransferase Measurement"
subtitles(result) <- c("This is a subtitle.", "This is another subtitle.")
main_footer(result) <- "This is a demo table for illustration purpose."
prov_footer(result) <- "Program: demo_poc_docx.R\nDate: 2024-11-06\nVersion: 0.0.1\n"

## -----------------------------------------------------------------------------
flx_res <- tt_to_flextable(result)
export_as_docx(flx_res,
  file = tf,
  section_properties = section_properties_default(orientation = "landscape")
)
flx_res

## -----------------------------------------------------------------------------
cw <- propose_column_widths(result)
cw <- cw / sum(cw)
cw <- c(0.6, 0.1, 0.1, 0.1, 0.1)
spd <- section_properties_default(orientation = "landscape")
fin_cw <- cw * spd$page_size$width / 2 / sum(cw)

flex_tbl <- tt_to_flextable(result,
  total_page_width = spd$page_size$width / 2,
  counts_in_newline = TRUE,
  autofit_to_page = FALSE,
  bold_titles = TRUE,
  colwidths = cw
)

export_as_docx(flex_tbl, file = tf)
flex_tbl

## -----------------------------------------------------------------------------
flx_res <- tt_to_flextable(
  result,
  paginate = TRUE,
  titles_as_header = FALSE,
  lpp = 250,
  counts_in_newline = TRUE,
  bold_titles = TRUE,
  theme = theme_docx_default()
)
export_as_docx(flx_res, file = tf, add_page_break = TRUE)
flx_res[[1]]

## -----------------------------------------------------------------------------
tbl <- basic_table() %>%
  split_cols_by("ACTARM") %>%
  split_rows_by("PARAM", split_fun = drop_split_levels, section_div = "-") %>%
  split_rows_by("AVISIT", split_fun = drop_split_levels, section_div = " ") %>%
  split_cols_by_multivar(
    vars = c("AVAL", "CHG"), varlabels = c("Value at Visit", "Change from Baseline")
  ) %>%
  analyze_colvars(afun = afun) %>%
  build_table(adlb_f)
flx_res <- tt_to_flextable(tbl)
export_as_docx(flx_res, file = tf)
flx_res

