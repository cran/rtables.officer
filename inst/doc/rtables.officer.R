## ----include = FALSE----------------------------------------------------------
suggested_dependent_pkgs <- c("dplyr")
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

## ----message=FALSE------------------------------------------------------------
library(rtables.officer)
library(dplyr)

## ----data---------------------------------------------------------------------
n <- 400

set.seed(1)

df <- tibble(
  arm = factor(sample(c("Arm A", "Arm B"), n, replace = TRUE), levels = c("Arm A", "Arm B")),
  country = factor(sample(c("CAN", "USA"), n, replace = TRUE, prob = c(.55, .45)), levels = c("CAN", "USA")),
  gender = factor(sample(c("Female", "Male"), n, replace = TRUE), levels = c("Female", "Male")),
  handed = factor(sample(c("Left", "Right"), n, prob = c(.6, .4), replace = TRUE), levels = c("Left", "Right")),
  age = rchisq(n, 30) + 10
) %>% mutate(
  weight = 35 * rnorm(n, sd = .5) + ifelse(gender == "Female", 140, 180)
)

head(df)

## ----tbl, echo=FALSE----------------------------------------------------------
lyt <- basic_table(show_colcounts = TRUE) %>%
  split_cols_by("arm") %>%
  split_cols_by("gender") %>%
  split_rows_by("country") %>%
  summarize_row_groups() %>%
  split_rows_by("handed") %>%
  summarize_row_groups() %>%
  analyze("age", afun = mean, format = "xx.xx")

tbl <- build_table(lyt, df)
tbl

## ----tbl, echo=TRUE-----------------------------------------------------------
lyt <- basic_table(show_colcounts = TRUE) %>%
  split_cols_by("arm") %>%
  split_cols_by("gender") %>%
  split_rows_by("country") %>%
  summarize_row_groups() %>%
  split_rows_by("handed") %>%
  summarize_row_groups() %>%
  analyze("age", afun = mean, format = "xx.xx")

tbl <- build_table(lyt, df)
tbl

## ----eval=FALSE---------------------------------------------------------------
# output_dir <- "." # specify here where the file should be saved
# export_as_docx(tbl, file = file.path(output_dir, "example.docx"))

