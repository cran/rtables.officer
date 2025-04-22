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

## -----------------------------------------------------------------------------
library(rlistings)
library(dplyr)
library(rtables.officer)

lsting <- as_listing(
  df = head(formatters::ex_adae, n = 50),
  key_cols = c("USUBJID", "ARM"),
  disp_cols = c("AETOXGR", "AEDECOD", "AESEV"),
  main_title = "Listing of Adverse Events (First 50 Records)",
  main_footer = "Source: formatters::ex_adae example dataset",
  add_trailing_sep = "ARM" # for readability adds a space line between differen ARMs
)

## -----------------------------------------------------------------------------
# 1. Add Subtitles and Provenance Footer
subtitles(lsting) <- c(
  "Subset: Treatment-Emergent Events",
  "Protocol: XYZ-123"
)
prov_footer(lsting) <- c(
  paste("R Version:", R.version.string),
  paste("rlistings Version:", packageVersion("rlistings")),
  paste("Generated on:", Sys.time()), # Use current time
  "File: your_script_name.R" # Add script name if applicable
)

## -----------------------------------------------------------------------------
flx_res <- tt_to_flextable(lsting)

# Create a temporary file for the output
tf <- tempfile(fileext = ".docx")

export_as_docx(lsting,
  file = tf,
  section_properties = section_properties_default(orientation = "landscape")
)
flx_res

