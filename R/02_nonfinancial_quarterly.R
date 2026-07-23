# =========================================================
# OECD Non-Financial Accounts Relabeling Script
# Quarterly Version 
#
# Description:
# This script retrieves OECD quarterly non-financial accounts data, matches OECD
# transaction codes to the project glossary, builds final variable labels,
# and exports sector tables and QA sheets to Excel.
#
# Script scope:
#   - quarterly data only
#   - sector-level export for institutional sectors
#   - Total Economy aggregate output based on S1
#   - QA sheets for relabeling, transaction availability, and output consistency
#   - user-defined start quarter
#
# Required input workbook sheets:
#   - glossary
#   - sectors
#   - api_code_dict
#   - total_economy_labels
#
# Optional input workbook sheet:
#   - mapping
#
# Outputs:
#   - one sheet per sector
#   - QA sheets for validation and review
#
# Notes:
# - Data are retrieved from the selected start period onward.
# - The query does not impose a fixed end period, so the script reads
#   up to the latest published data available from the source.
# - In the quarterly OECD dataset used here, S1M represents households and NPISH.
# - If direct S13 rows are unavailable for Italy, selected GG series are derived transparently from S1 minus S11/S12/S1M, with dedicated QA sheets.
# - OECD accounting entry C is treated as the resources/received side
#   and D as the uses/spent side. Internally these are named REV and EXP.
# =========================================================


# -------------------------
# Packages
# -------------------------
install_if_missing <- function(pkgs) {
  to_install <- pkgs[!pkgs %in% rownames(installed.packages())]
  if (length(to_install) > 0) install.packages(to_install)
}

install_if_missing(c("openxlsx", "dplyr", "tidyr", "stringr", "tibble", "rsdmx", "writexl"))

library(openxlsx)
library(dplyr)
library(tidyr)
library(stringr)
library(tibble)
library(rsdmx)
library(writexl)


# -------------------------
# User inputs
# -------------------------

# Country shown in the output workbook
country_name <- "Italy"

# Country code used in the OECD query
country_code <- "ITA"

# Repository-relative input / output files
input_file <- file.path("data", "FINAL_nonfinancial_accounts_input.xlsx")
dir.create("outputs", showWarnings = FALSE, recursive = TRUE)
output_file <- file.path(
  "outputs",
  paste0(country_code, "_stage1_nonfinancial_accounts_quarterly_output.xlsx")
)

# Define the first quarter to request from the source.
# The query has no fixed end period, so all observations from
# `start_period` onward are retrieved automatically.
start_year <- 2010
start_quarter <- 1
start_period <- paste0(start_year, "-Q", start_quarter)

# Quarterly adjustment basis used throughout Stage 1.
# N = non-seasonally-adjusted series. This is used here because quarterly
# values are later compared/aggregated to annual national-accounts totals.
# IMPORTANT: never sum N and Y adjustment variants together.
adjustment_code <- "N"

# OECD API settings used in this script
dataset_id <- "OECD.SDD.NAD,DSD_NASEC1@DF_QSA,1.1"

# Sectors included in the quarterly query.
# In this dataset, S1M represents households and NPISH.
sector_codes <- c("S2", "S11", "S12", "S1M", "S13")

# Accounting entries:
# C = resources/received side
# D = uses/spent side
entry_codes <- c("D", "C")

# OECD transaction codes included in the quarterly query.
# The same core non-financial transaction universe is used across the
# annual and quarterly relabeling pipelines where the OECD source provides it.
transaction_codes <- c(
  "P1", "P2", "B2A3G", "B3G", "D21", "D31",
  "D1", "D2", "D5", "D3", "D61", "D62", "D7",
  "D41", "D42", "D43", "D44", "D45", "D8", "D9", "NP",
  "P3", "P5", "P6", "P7",
  "B12", "B8G", "B9"
)

# Default ordering used when a row is not explicitly controlled in mapping.
# This keeps automatically generated variables aligned with the shared
# transaction structure.
transaction_order_tbl <- tibble(
  transaction_code = transaction_codes,
  transaction_order = seq_along(transaction_codes) * 10L
)

# Output units label
units_label <- "EUR, current prices, millions"


# -------------------------
# Build OECD API URL
# -------------------------
build_oecd_url <- function(dataset_id,
                           country_code,
                           sector_codes,
                           entry_codes,
                           transaction_codes,
                           start_period) {
  key <- paste0(
    "Q..",
    country_code, ".",
    paste(sector_codes, collapse = "+"),
    "..",
    paste(entry_codes, collapse = "+"), ".",
    paste(transaction_codes, collapse = "+"),
    "......."
  )
  
  paste0(
    "https://sdmx.oecd.org/public/rest/data/",
    dataset_id, "/",
    key,
    "?startPeriod=", start_period
  )
}

url_api <- build_oecd_url(
  dataset_id = dataset_id,
  country_code = country_code,
  sector_codes = sector_codes,
  entry_codes = entry_codes,
  transaction_codes = transaction_codes,
  start_period = start_period
)


# -------------------------
# Total Economy aggregate selection
# -------------------------
# The S1 aggregate economy selection is retrieved as a separate output sheet
# to complement the institutional-sector accounts. It is not controlled by
# the sector mapping rules because it represents the whole economy rather
# than an additional institutional sector.


total_economy_transaction_codes <- c(
  "P3", "P31", "P32", "P5", "B1G", "D1", "D11", "D12",
  "B2A3G", "D21X31", "D2", "D3", "P52", "P53", "D21", "D31",
  "B101", "B1N", "B2G", "B3G", "B5G", "B6G", "B7G", "B8G",
  "B9", "B9FX9", "D211", "D212", "D214", "D29", "D39", "D4",
  "D41", "D41G", "D42", "D43", "D44", "D441", "D442", "D443",
  "D45", "D5", "D51", "D59", "D6", "D61", "D611", "D612",
  "D613", "D614", "D61N", "D61SC", "D62", "D63", "D631", "D632",
  "D7", "D71", "D72", "D74", "D75", "D76", "D8", "D9",
  "D91", "D92", "D99", "NP", "P1", "P11", "P12", "P13",
  "P2", "P51C", "TR1", "TR211", "TR212", "TR22", "TR241",
  "TR311", "TR312", "TU1", "TU211", "TU212", "TU22", "TU241",
  "TU311", "TU312", "D612_D614", "P51CB"
)

total_economy_start_period <- start_period

build_total_economy_url <- function(dataset_id,
                                    country_code,
                                    entry_codes,
                                    transaction_codes,
                                    start_period) {
  key <- paste0(
    "Q..",
    country_code,
    ".S1..",
    paste(entry_codes, collapse = "+"),
    ".",
    paste(transaction_codes, collapse = "+"),
    "......."
  )

  paste0(
    "https://sdmx.oecd.org/public/rest/data/",
    dataset_id,
    "/",
    key,
    "?startPeriod=", start_period,
    "&dimensionAtObservation=AllDimensions"
  )
}

url_total_economy <- build_total_economy_url(
  dataset_id = dataset_id,
  country_code = country_code,
  entry_codes = entry_codes,
  transaction_codes = total_economy_transaction_codes,
  start_period = total_economy_start_period
)

message("Using sector URL: ", url_api)
message("Using S1 aggregate economy URL: ", url_total_economy)
message("Output file: ", output_file)
message("Quarterly adjustment basis: ", adjustment_code,
        " (N = non-seasonally adjusted; Y = seasonally adjusted)")

# -------------------------
# Helpers
# -------------------------

# Read one sheet from the input workbook
# and standardize column names to lowercase
read_input_sheet <- function(file, sheet_name) {
  x <- openxlsx::read.xlsx(file, sheet = sheet_name)
  names(x) <- tolower(names(x))
  x
}

# Trim spaces and convert to character
normalize_text <- function(x) {
  stringr::str_squish(as.character(x))
}

# Stop if required columns are missing
required_cols <- function(df, cols, df_name) {
  missing <- setdiff(cols, names(df))
  if (length(missing) > 0) {
    stop(df_name, " is missing required columns: ", paste(missing, collapse = ", "))
  }
}

# Standardize specific OECD wording so joins stay consistent
normalize_oecd_label <- function(x) {
  x |>
    normalize_text() |>
    str_replace_all(
      fixed("Net lending/borrowing"),
      "Net lending (+) / net borrowing (-)"
    ) |>
    str_replace_all(
      fixed("Net lending / borrowing"),
      "Net lending (+) / net borrowing (-)"
    )
}

# Identify items that should not receive revenue / expenditure suffixes
is_no_flow_item <- function(oecd_label, notes = "") {
  text <- normalize_text(paste(oecd_label, notes))

  str_detect(
    text,
    regex(
      "balancing item|accounting item|production item|net lending|net borrowing|saving|external balance|^output$|gross output|operating surplus|mixed income",
      ignore_case = TRUE
    )
  )
}

# Build the final variable label
build_label <- function(tx_code, sector_letter, entry, no_flow_item) {
  base <- paste0(tx_code, sector_letter)
  
  case_when(
    no_flow_item ~ base,
    entry == "REV" ~ paste0(base, "R"),
    entry == "EXP" ~ paste0(base, "S"),
    TRUE ~ base
  )
}

# Build the variable description used in the exported table
build_desc <- function(oecd_label, entry, no_flow_item) {
  case_when(
    no_flow_item ~ oecd_label,
    entry == "REV" ~ paste0(oecd_label, ", received"),
    entry == "EXP" ~ paste0(oecd_label, ", spent"),
    TRUE ~ oecd_label
  )
}

# Standardize quarterly period formatting
parse_period <- function(x) {
  x <- as.character(x)
  x <- stringr::str_replace_all(x, "Q([1-4])-(\\d{4})", "\\2-Q\\1")
  x
}


# SDMX output column names can differ across OECD queries.
# The normal sector query often returns obsTime/obsValue, while the
# S1 aggregate economy query may return TIME_PERIOD/OBS_VALUE. This helper
# creates standard obsTime and obsValue columns so the rest of the
# script works for both cases.
find_sdmx_col <- function(df, candidates, df_name) {
  for (cand in candidates) {
    hit <- which(tolower(names(df)) == tolower(cand))
    if (length(hit) > 0) return(names(df)[hit[1]])
  }

  stop(
    df_name,
    " is missing an expected column. Tried: ",
    paste(candidates, collapse = ", "),
    ". Available columns are: ",
    paste(names(df), collapse = ", ")
  )
}

standardize_sdmx_obs_columns <- function(df, df_name = "SDMX data") {
  period_col <- find_sdmx_col(
    df,
    c("obsTime", "TIME_PERIOD", "time_period", "Time", "TIME"),
    df_name
  )

  value_col <- find_sdmx_col(
    df,
    c("obsValue", "OBS_VALUE", "obs_value", "Value", "VALUE"),
    df_name
  )

  df$obsTime <- df[[period_col]]
  df$obsValue <- df[[value_col]]
  df
}

# Read an OECD SDMX URL and return a data frame.
# This gives clearer feedback if the API server is unavailable,
# the connection times out, or the generated URL is not valid.
read_oecd_sdmx <- function(url, label) {
  tryCatch(
    {
      rsdmx::readSDMX(url) |> as.data.frame()
    },
    error = function(e) {
      stop(
        label, " OECD API call failed. ",
        "Please check the URL, internet connection, or OECD server availability. ",
        "Original error: ", conditionMessage(e),
        call. = FALSE
      )
    }
  )
}

# Write one QA sheet to the workbook
write_qa_sheet <- function(wb, sheet_name, df, title) {
  addWorksheet(wb, sheet_name)
  writeData(
    wb, sheet_name,
    data.frame(Message = title),
    startRow = 1, startCol = 1,
    colNames = FALSE
  )
  
  if (nrow(df) == 0) {
    writeData(
      wb, sheet_name,
      data.frame(Status = "No issues found."),
      startRow = 3, startCol = 1
    )
  } else {
    writeData(wb, sheet_name, df, startRow = 3, startCol = 1)
  }
  
  setColWidths(wb, sheet_name, cols = 1:max(1, ncol(df)), widths = "auto")
}

# Build one sector sheet in wide format for export
build_sheet_df <- function(mapped_long_api, sheet_name, country_name, units_label) {
  x <- mapped_long_api |>
    filter(sheet == sheet_name) |>
    arrange(period, ord, label)
  
  if (nrow(x) == 0) return(NULL)
  
  labels <- x |>
    distinct(ord, label, desc) |>
    arrange(ord, label)
  
  wide <- x |>
    select(period, label, value) |>
    pivot_wider(
      names_from = label,
      values_from = value
    ) |>
    arrange(period)
  
  keep_cols <- names(wide)[
    vapply(
      wide,
      function(col) !all(is.na(col) | col == ""),
      logical(1)
    )
  ]
  
  if (!"period" %in% keep_cols) {
    keep_cols <- c("period", keep_cols)
  }
  
  wide <- wide |>
    select(all_of(unique(keep_cols)))
  
  labels <- labels |>
    filter(label %in% names(wide)[names(wide) != "period"])
  
  sector_title <- x$sector_title[1]
  
  meta <- data.frame(
    X1 = c("Country name", "Institutional sector", "Units", "Date/time of data extraction"),
    X2 = c(country_name, sector_title, units_label, as.character(Sys.Date())),
    stringsAsFactors = FALSE
  )
  
  header_desc  <- c("", labels$desc)
  header_label <- c("Quarter", labels$label)
  
  desc_row  <- as.data.frame(as.list(header_desc), stringsAsFactors = FALSE)
  label_row <- as.data.frame(as.list(header_label), stringsAsFactors = FALSE)
  
  names(desc_row)  <- c("V1", paste0("V", seq_len(length(header_desc) - 1) + 1))
  names(label_row) <- names(desc_row)
  
  wide_chr <- wide |>
    mutate(across(everything(), as.character))
  
  names(wide_chr) <- names(desc_row)
  
  out <- bind_rows(desc_row, label_row, wide_chr)
  
  list(meta = meta, table = out)
}


# Build the Total Economy sheet in wide format for export.
# This sheet is generated from the S1 aggregate economy selection. Labels use
# the same transaction abbreviations as the sector sheets where available,
# but omit the sector letter because the absence of a sector marker identifies
# the aggregate Total Economy variable.
build_total_economy_sheet_df <- function(total_economy_mapped,
                                         country_name,
                                         units_label,
                                         url_total_economy) {
  x <- total_economy_mapped |>
    arrange(period, ord, label)

  if (nrow(x) == 0) return(NULL)

  labels <- x |>
    distinct(ord, label, desc) |>
    arrange(ord, label)

  wide <- x |>
    select(period, label, value) |>
    pivot_wider(
      names_from = label,
      values_from = value
    ) |>
    arrange(period)

  keep_cols <- names(wide)[
    vapply(
      wide,
      function(col) !all(is.na(col) | col == ""),
      logical(1)
    )
  ]

  if (!"period" %in% keep_cols) {
    keep_cols <- c("period", keep_cols)
  }

  wide <- wide |>
    select(all_of(unique(keep_cols)))

  labels <- labels |>
    filter(label %in% names(wide)[names(wide) != "period"])

  meta <- data.frame(
    X1 = c(
      "Country name",
      "Institutional sector",
      "Units",
      "OECD source URL",
      "Date/time of data extraction"
    ),
    X2 = c(
      country_name,
      "Total economy (S1)",
      units_label,
      url_total_economy,
      as.character(Sys.Date())
    ),
    stringsAsFactors = FALSE
  )

  header_desc  <- c("", labels$desc)
  header_label <- c("Quarter", labels$label)

  desc_row  <- as.data.frame(as.list(header_desc), stringsAsFactors = FALSE)
  label_row <- as.data.frame(as.list(header_label), stringsAsFactors = FALSE)

  names(desc_row)  <- c("V1", paste0("V", seq_len(length(header_desc) - 1) + 1))
  names(label_row) <- names(desc_row)

  wide_chr <- wide |>
    mutate(across(everything(), as.character))

  names(wide_chr) <- names(desc_row)

  out <- bind_rows(desc_row, label_row, wide_chr)

  list(meta = meta, table = out)
}

# -------------------------
# Read structured inputs
# -------------------------

glossary_raw <- read_input_sheet(input_file, "glossary")
required_cols(glossary_raw, c("oecd_label", "tx_code"), "glossary")

if (!"notes" %in% names(glossary_raw)) glossary_raw$notes <- ""
if (!"manual_comment" %in% names(glossary_raw)) glossary_raw$manual_comment <- ""

glossary_tbl <- glossary_raw |>
  mutate(
    oecd_label     = normalize_oecd_label(oecd_label),
    tx_code        = normalize_text(tx_code),
    notes          = if_else(is.na(notes), "", normalize_text(notes)),
    manual_comment = if_else(is.na(manual_comment), "", normalize_text(manual_comment))
  ) |>
  filter(oecd_label != "", tx_code != "") |>
  distinct(oecd_label, .keep_all = TRUE)

# One project tx_code must represent one OECD concept. Otherwise distinct
# source concepts can silently collapse to the same final variable label.
duplicate_project_tx_codes <- glossary_tbl |>
  distinct(oecd_label, tx_code) |>
  count(tx_code, name = "concept_count") |>
  filter(tx_code != "", concept_count > 1)

if (nrow(duplicate_project_tx_codes) > 0) {
  stop(
    "glossary contains project tx_code values assigned to more than one OECD concept: ",
    paste(duplicate_project_tx_codes$tx_code, collapse = ", "),
    ". Use distinct project labels before running the pipeline."
  )
}

sector_raw <- read_input_sheet(input_file, "sectors")
required_cols(sector_raw, c("sheet", "sector_letter", "sector_title"), "sectors")

sector_dict <- sector_raw |>
  mutate(
    sheet         = normalize_text(sheet),
    sector_letter = normalize_text(sector_letter),
    sector_title  = normalize_text(sector_title)
  ) |>
  filter(sheet != "") |>
  distinct(sheet, .keep_all = TRUE) |>
  select(sheet, sector_letter, sector_title)

api_code_raw <- read_input_sheet(input_file, "api_code_dict")
required_cols(api_code_raw, c("transaction_code", "oecd_label"), "api_code_dict")

if (!"status" %in% names(api_code_raw)) api_code_raw$status <- "CONFIRMED"

api_code_dict <- api_code_raw |>
  mutate(
    transaction_code = normalize_text(transaction_code),
    oecd_label       = normalize_oecd_label(oecd_label),
    status           = normalize_text(status)
  ) |>
  filter(transaction_code != "", oecd_label != "") |>
  filter(status == "" | is.na(status) | toupper(status) == "CONFIRMED") |>
  distinct(transaction_code, .keep_all = TRUE) |>
  select(transaction_code, oecd_label)


total_economy_labels_raw <- read_input_sheet(input_file, "total_economy_labels")
required_cols(
  total_economy_labels_raw,
  c("transaction_code", "oecd_label", "short_label", "no_flow_item"),
  "total_economy_labels"
)

if (!"notes" %in% names(total_economy_labels_raw)) {
  total_economy_labels_raw$notes <- ""
}

total_economy_labels_tbl <- total_economy_labels_raw |>
  mutate(
    transaction_code = normalize_text(transaction_code),
    oecd_label_total = normalize_oecd_label(oecd_label),
    short_label      = normalize_text(short_label),
    no_flow_item     = toupper(normalize_text(no_flow_item)) %in% c("TRUE", "T", "YES", "Y", "1"),
    notes_total      = if_else(is.na(notes), "", normalize_text(notes))
  ) |>
  filter(transaction_code != "", oecd_label_total != "", short_label != "") |>
  distinct(transaction_code, .keep_all = TRUE) |>
  select(transaction_code, oecd_label_total, short_label, no_flow_item, notes_total)

# Total Economy short labels must be unique across transaction codes.
# A duplicated short_label would collapse different S1 concepts together.
duplicate_total_short_labels <- total_economy_labels_tbl |>
  distinct(transaction_code, short_label) |>
  count(short_label, name = "code_count") |>
  filter(short_label != "", code_count > 1)

if (nrow(duplicate_total_short_labels) > 0) {
  stop(
    "total_economy_labels contains short_label values assigned to more than one transaction code: ",
    paste(duplicate_total_short_labels$short_label, collapse = ", "),
    ". Use distinct short labels before running the pipeline."
  )
}

sheet_names <- openxlsx::getSheetNames(input_file)
mapping_exists <- "mapping" %in% tolower(sheet_names)

if (mapping_exists) {
  mapping_raw <- read_input_sheet(input_file, "mapping")
  
  if (!"sheet" %in% names(mapping_raw)) mapping_raw$sheet <- ""
  if (!"entry" %in% names(mapping_raw)) mapping_raw$entry <- ""
  if (!"oecd_label" %in% names(mapping_raw)) mapping_raw$oecd_label <- ""
  if (!"custom_label" %in% names(mapping_raw)) mapping_raw$custom_label <- ""
  if (!"custom_desc" %in% names(mapping_raw)) mapping_raw$custom_desc <- ""
  if (!"include_final" %in% names(mapping_raw)) mapping_raw$include_final <- TRUE
  if (!"ord" %in% names(mapping_raw)) mapping_raw$ord <- NA
  if (!"review_note" %in% names(mapping_raw)) mapping_raw$review_note <- ""
  
  mapping_tbl <- mapping_raw |>
    mutate(
      sheet         = normalize_text(sheet),
      entry         = normalize_text(entry),
      oecd_label    = normalize_oecd_label(oecd_label),
      ord           = suppressWarnings(as.integer(ord)),
      include_final = toupper(normalize_text(include_final)) %in% c("TRUE", "T", "YES", "Y", "1"),
      custom_label  = if_else(is.na(custom_label), "", normalize_text(custom_label)),
      custom_desc   = if_else(is.na(custom_desc), "", normalize_text(custom_desc)),
      review_note   = if_else(is.na(review_note), "", normalize_text(review_note))
    ) |>
    filter(sheet != "", entry != "", oecd_label != "") |>
    distinct(sheet, entry, oecd_label, .keep_all = TRUE) |>
    select(sheet, entry, oecd_label, ord, include_final, custom_label, custom_desc, review_note)
} else {
  mapping_tbl <- tibble(
    sheet = character(),
    entry = character(),
    oecd_label = character(),
    ord = integer(),
    include_final = logical(),
    custom_label = character(),
    custom_desc = character(),
    review_note = character()
  )
}


# -------------------------
# Pull and clean OECD API data
# -------------------------
df_api <- read_oecd_sdmx(url_api, "Sector accounts")
df_api <- standardize_sdmx_obs_columns(df_api, "Sector API data")

api_pre_agg <- df_api |>
  mutate(
    entry = case_when(
      ACCOUNTING_ENTRY == "C" ~ "REV",
      ACCOUNTING_ENTRY == "D" ~ "EXP",
      TRUE ~ NA_character_
    ),
    sheet = case_when(
      SECTOR == "S11" ~ "NFC",
      SECTOR == "S12" ~ "FC",
      SECTOR == "S13" ~ "GG",
      SECTOR == "S1M" ~ "HH",
      SECTOR == "S2" ~ "ROW",
      TRUE ~ NA_character_
    ),
    period = parse_period(obsTime),
    value = as.numeric(obsValue),
    transaction_code = as.character(TRANSACTION)
  ) |>
  filter(
    !is.na(sheet),
    !is.na(entry),
    !is.na(period),
    ADJUSTMENT == adjustment_code
  )

# After selecting one adjustment basis, verify that no other OECD dimension
# still creates multiple observations for the same sector-entry-code-period.
qa_sector_source_duplicates <- api_pre_agg |>
  count(SECTOR, entry, transaction_code, period, name = "source_rows") |>
  filter(source_rows > 1) |>
  arrange(SECTOR, entry, transaction_code, period)

if (nrow(qa_sector_source_duplicates) > 0) {
  stop(
    "Sector API returned multiple source observations after ADJUSTMENT filtering. ",
    "Refusing to sum potentially different OECD variants. ",
    "Inspect qa_sector_source_duplicates in the script environment."
  )
}

api_long <- api_pre_agg |>
  group_by(sheet, entry, transaction_code, period) |>
  summarise(
    value = if (all(is.na(value))) NA_real_ else sum(value, na.rm = TRUE),
    .groups = "drop"
  )


# -------------------------
# Pull S1 aggregate economy API data
# -------------------------
df_total_api <- read_oecd_sdmx(url_total_economy, "S1 aggregate economy")
df_total_api <- standardize_sdmx_obs_columns(df_total_api, "S1 aggregate economy API data")

total_api_pre_agg <- df_total_api |>
  mutate(
    entry = case_when(
      ACCOUNTING_ENTRY == "C" ~ "REV",
      ACCOUNTING_ENTRY == "D" ~ "EXP",
      TRUE ~ NA_character_
    ),
    sheet = "TOTAL_ECONOMY",
    period = parse_period(obsTime),
    value = as.numeric(obsValue),
    transaction_code = as.character(TRANSACTION)
  ) |>
  filter(
    !is.na(entry),
    !is.na(period),
    ADJUSTMENT == adjustment_code
  )

qa_total_source_duplicates <- total_api_pre_agg |>
  count(entry, transaction_code, period, name = "source_rows") |>
  filter(source_rows > 1) |>
  arrange(entry, transaction_code, period)

if (nrow(qa_total_source_duplicates) > 0) {
  stop(
    "S1 Total Economy API returned multiple source observations after ADJUSTMENT filtering. ",
    "Refusing to sum potentially different OECD variants. ",
    "Inspect qa_total_source_duplicates in the script environment."
  )
}

total_api_long <- total_api_pre_agg |>
  group_by(sheet, entry, transaction_code, period) |>
  summarise(
    value = if (all(is.na(value))) NA_real_ else sum(value, na.rm = TRUE),
    .groups = "drop"
  )


# -------------------------
# Derive quarterly General Government (GG / S13) when direct S13 rows
# are unavailable in the selected OECD response
# -------------------------
# For the exact Italy DF_QSA response used by this pipeline, the selected
# non-seasonally-adjusted sector query may return no direct S13 observations.
# In that case, selected GG series are reconstructed transparently from the
# institutional-sector partition:
#
#   S1 = S11 + S12 + S13 + S1M
#
# hence:
#
#   S13 = S1 - S11 - S12 - S1M
#
# Only explicitly listed, accounting-justified series are derived below.
# P1/P2 are NOT invented when the required sector coverage is unavailable.
# For D2/D3, the non-overlapping S1 components D29/D39 are used so product
# taxes/subsidies are not double counted.
#
# Direct S13 observations, if present in a future OECD response, always take
# precedence; derived rows only fill missing S13 code-entry-period keys.

gg_periods <- sort(unique(total_api_long$period))

series_vec <- function(df,
                       transaction_code_value,
                       entry_value = NULL,
                       sheet_value = NULL,
                       periods = gg_periods) {
  x <- df |>
    filter(.data$transaction_code == transaction_code_value)

  if (!is.null(entry_value)) {
    x <- x |>
      filter(.data$entry == entry_value)
  }

  if (!is.null(sheet_value)) {
    x <- x |>
      filter(.data$sheet == sheet_value)
  }

  if (nrow(x) == 0) {
    return(rep(NA_real_, length(periods)))
  }

  # For no-flow/balancing items returned on both accounting sides,
  # choose one deterministic representation. REV is preferred, then EXP.
  x <- x |>
    mutate(
      entry_priority = case_when(
        .data$entry == "REV" ~ 1L,
        .data$entry == "EXP" ~ 2L,
        TRUE ~ 3L
      )
    ) |>
    arrange(.data$period, .data$entry_priority) |>
    group_by(.data$period) |>
    summarise(
      value = {
        non_missing <- .data$value[!is.na(.data$value)]
        if (length(non_missing) == 0) NA_real_ else non_missing[1]
      },
      .groups = "drop"
    )

  x$value[match(periods, x$period)]
}

s1v <- function(code, entry_value = NULL) {
  series_vec(
    total_api_long,
    transaction_code_value = code,
    entry_value = entry_value,
    periods = gg_periods
  )
}

secv <- function(sheet_name, code, entry_value = NULL) {
  series_vec(
    api_long,
    transaction_code_value = code,
    entry_value = entry_value,
    sheet_value = sheet_name,
    periods = gg_periods
  )
}

make_gg_row <- function(code, entry_value, values, formula_text) {
  tibble(
    sheet = "GG",
    entry = entry_value,
    transaction_code = code,
    period = gg_periods,
    value = as.numeric(values),
    derivation_formula = formula_text
  )
}

gg_derived_with_formula <- bind_rows(
  # Production-account balancing items available without inventing P1/P2.
  make_gg_row(
    "B2A3G", "REV",
    s1v("B2A3G") -
      secv("NFC", "B2A3G") -
      secv("FC",  "B2A3G") -
      secv("HH",  "B2A3G"),
    "S1 B2A3G - S11 - S12 - S1M"
  ),
  make_gg_row(
    "B3G", "REV",
    s1v("B3G") - secv("HH", "B3G"),
    "S1 B3G - S1M B3G; S11/S12 mixed income are structural zero"
  ),

  # Product taxes/subsidies and totals.
  make_gg_row(
    "D21", "REV",
    s1v("D21", "REV"),
    "S1 D21 REV; non-GG domestic receipts are structural zero"
  ),
  make_gg_row(
    "D31", "EXP",
    s1v("D31", "EXP"),
    "S1 D31 EXP; non-GG domestic expenditures are structural zero"
  ),
  make_gg_row(
    "D2", "REV",
    s1v("D2", "REV"),
    "S1 D2 REV; non-GG domestic receipts are structural zero"
  ),
  make_gg_row(
    "D2", "EXP",
    s1v("D29", "EXP") -
      secv("NFC", "D2", "EXP") -
      secv("FC",  "D2", "EXP") -
      secv("HH",  "D2", "EXP"),
    "S1 D29 EXP - S11 D2 EXP - S12 D2 EXP - S1M D2 EXP; GG D21 EXP structural zero"
  ),
  make_gg_row(
    "D3", "REV",
    s1v("D39", "REV") -
      secv("NFC", "D3", "REV") -
      secv("FC",  "D3", "REV") -
      secv("HH",  "D3", "REV"),
    "S1 D39 REV - S11 D3 REV - S12 D3 REV - S1M D3 REV; GG D31 REV structural zero"
  ),
  make_gg_row(
    "D3", "EXP",
    s1v("D3", "EXP"),
    "S1 D3 EXP; non-GG domestic expenditures are structural zero"
  ),

  # Compensation and current taxes.
  make_gg_row(
    "D1", "EXP",
    s1v("D1", "EXP") -
      secv("NFC", "D1", "EXP") -
      secv("FC",  "D1", "EXP") -
      secv("HH",  "D1", "EXP"),
    "S1 D1 EXP - S11 - S12 - S1M"
  ),
  make_gg_row(
    "D5", "REV",
    s1v("D5", "REV"),
    "S1 D5 REV; non-GG domestic receipts are structural zero"
  ),
  make_gg_row(
    "D5", "EXP",
    s1v("D5", "EXP") -
      secv("NFC", "D5", "EXP") -
      secv("FC",  "D5", "EXP") -
      secv("HH",  "D5", "EXP"),
    "S1 D5 EXP - S11 - S12 - S1M"
  ),

  # Social contributions and benefits.
  make_gg_row(
    "D61", "REV",
    s1v("D61", "REV") -
      secv("NFC", "D61", "REV") -
      secv("FC",  "D61", "REV") -
      secv("HH",  "D61", "REV"),
    "S1 D61 REV - S11 - S12 - S1M"
  ),
  make_gg_row(
    "D61", "EXP",
    s1v("D61", "EXP") -
      secv("HH", "D61", "EXP"),
    "S1 D61 EXP - S1M; S11/S12 expenditures are structural zero"
  ),
  make_gg_row(
    "D62", "REV",
    s1v("D62", "REV") -
      secv("HH", "D62", "REV"),
    "S1 D62 REV - S1M; S11/S12 receipts are structural zero"
  ),
  make_gg_row(
    "D62", "EXP",
    s1v("D62", "EXP") -
      secv("NFC", "D62", "EXP") -
      secv("FC",  "D62", "EXP") -
      secv("HH",  "D62", "EXP"),
    "S1 D62 EXP - S11 - S12 - S1M"
  ),

  # Other current transfers.
  make_gg_row(
    "D7", "REV",
    s1v("D7", "REV") -
      secv("NFC", "D7", "REV") -
      secv("FC",  "D7", "REV") -
      secv("HH",  "D7", "REV"),
    "S1 D7 REV - S11 - S12 - S1M"
  ),
  make_gg_row(
    "D7", "EXP",
    s1v("D7", "EXP") -
      secv("NFC", "D7", "EXP") -
      secv("FC",  "D7", "EXP") -
      secv("HH",  "D7", "EXP"),
    "S1 D7 EXP - S11 - S12 - S1M"
  ),

  # Property income.
  make_gg_row(
    "D41", "REV",
    s1v("D41", "REV") -
      secv("NFC", "D41", "REV") -
      secv("FC",  "D41", "REV") -
      secv("HH",  "D41", "REV"),
    "S1 D41 REV - S11 - S12 - S1M"
  ),
  make_gg_row(
    "D41", "EXP",
    s1v("D41", "EXP") -
      secv("NFC", "D41", "EXP") -
      secv("FC",  "D41", "EXP") -
      secv("HH",  "D41", "EXP"),
    "S1 D41 EXP - S11 - S12 - S1M"
  ),
  make_gg_row(
    "D42", "REV",
    s1v("D42", "REV") -
      secv("NFC", "D42", "REV") -
      secv("FC",  "D42", "REV") -
      secv("HH",  "D42", "REV"),
    "S1 D42 REV - S11 - S12 - S1M"
  ),
  make_gg_row(
    "D42", "EXP",
    s1v("D42", "EXP") -
      secv("NFC", "D42", "EXP") -
      secv("FC",  "D42", "EXP"),
    "S1 D42 EXP - S11 - S12; S1M expenditure structural zero"
  ),
  make_gg_row(
    "D43", "REV",
    s1v("D43", "REV") -
      secv("NFC", "D43", "REV") -
      secv("FC",  "D43", "REV") -
      secv("HH",  "D43", "REV"),
    "S1 D43 REV - S11 - S12 - S1M"
  ),
  make_gg_row(
    "D43", "EXP",
    s1v("D43", "EXP") -
      secv("NFC", "D43", "EXP") -
      secv("FC",  "D43", "EXP"),
    "S1 D43 EXP - S11 - S12; S1M expenditure structural zero"
  ),
  make_gg_row(
    "D44", "REV",
    s1v("D44", "REV") -
      secv("NFC", "D44", "REV") -
      secv("FC",  "D44", "REV") -
      secv("HH",  "D44", "REV"),
    "S1 D44 REV - S11 - S12 - S1M"
  ),
  make_gg_row(
    "D44", "EXP",
    s1v("D44", "EXP") -
      secv("NFC", "D44", "EXP") -
      secv("FC",  "D44", "EXP"),
    "S1 D44 EXP - S11 - S12; S1M expenditure structural zero"
  ),
  make_gg_row(
    "D45", "REV",
    s1v("D45", "REV") -
      secv("NFC", "D45", "REV") -
      secv("FC",  "D45", "REV") -
      secv("HH",  "D45", "REV"),
    "S1 D45 REV - S11 - S12 - S1M"
  ),
  make_gg_row(
    "D45", "EXP",
    s1v("D45", "EXP") -
      secv("NFC", "D45", "EXP") -
      secv("FC",  "D45", "EXP") -
      secv("HH",  "D45", "EXP"),
    "S1 D45 EXP - S11 - S12 - S1M"
  ),

  # Pension entitlement adjustment.
  make_gg_row(
    "D8", "REV",
    s1v("D8", "REV") -
      secv("HH", "D8", "REV"),
    "S1 D8 REV - S1M; S11/S12 receipts are structural zero"
  ),
  make_gg_row(
    "D8", "EXP",
    s1v("D8", "EXP") -
      secv("NFC", "D8", "EXP") -
      secv("FC",  "D8", "EXP") -
      secv("HH",  "D8", "EXP"),
    "S1 D8 EXP - S11 - S12 - S1M"
  ),

  # Capital account.
  make_gg_row(
    "D9", "REV",
    s1v("D9", "REV") -
      secv("NFC", "D9", "REV") -
      secv("FC",  "D9", "REV") -
      secv("HH",  "D9", "REV"),
    "S1 D9 REV - S11 - S12 - S1M"
  ),
  make_gg_row(
    "D9", "EXP",
    s1v("D9", "EXP") -
      secv("NFC", "D9", "EXP") -
      secv("FC",  "D9", "EXP") -
      secv("HH",  "D9", "EXP"),
    "S1 D9 EXP - S11 - S12 - S1M"
  ),
  make_gg_row(
    "NP", "EXP",
    s1v("NP") -
      secv("NFC", "NP", "EXP") -
      secv("FC",  "NP", "EXP") -
      secv("HH",  "NP", "EXP"),
    "S1 net NP - S11 NP EXP - S12 NP EXP - S1M NP EXP"
  ),
  make_gg_row(
    "P5", "EXP",
    s1v("P5", "EXP") -
      secv("NFC", "P5", "EXP") -
      secv("FC",  "P5", "EXP") -
      secv("HH",  "P5", "EXP"),
    "S1 P5 EXP - S11 - S12 - S1M"
  ),

  # Consumption, saving and net lending.
  make_gg_row(
    "P3", "EXP",
    s1v("P3", "EXP") -
      secv("HH", "P3", "EXP"),
    "S1 P3 EXP - S1M; S11/S12 final consumption structural zero"
  ),
  make_gg_row(
    "B8G", "REV",
    s1v("B8G") -
      secv("NFC", "B8G") -
      secv("FC",  "B8G") -
      secv("HH",  "B8G"),
    "S1 B8G - S11 - S12 - S1M"
  ),
  make_gg_row(
    "B9", "EXP",
    s1v("B9") -
      secv("NFC", "B9") -
      secv("FC",  "B9") -
      secv("HH",  "B9"),
    "S1 B9 - S11 - S12 - S1M"
  )
) |>
  filter(!is.na(.data$value))

gg_direct_keys <- api_long |>
  filter(.data$sheet == "GG") |>
  distinct(entry, transaction_code, period)

gg_derived_to_add <- gg_derived_with_formula |>
  anti_join(
    gg_direct_keys,
    by = c("entry", "transaction_code", "period")
  )

qa_gg_derivation <- gg_derived_to_add |>
  group_by(.data$transaction_code, .data$entry, .data$derivation_formula) |>
  summarise(
    rows_derived = n(),
    first_period = min(.data$period, na.rm = TRUE),
    last_period = max(.data$period, na.rm = TRUE),
    .groups = "drop"
  ) |>
  arrange(.data$transaction_code, .data$entry)

qa_gg_source_status <- tibble(
  direct_S13_rows_returned = nrow(gg_direct_keys),
  derived_S13_rows_added = nrow(gg_derived_to_add),
  adjustment_basis = adjustment_code,
  method = paste0(
    "Direct S13 rows take precedence. Missing GG rows are derived from ",
    "S1 = S11 + S12 + S13 + S1M using explicit formulas; D2/D3 use ",
    "S1 D29/D39 to avoid product-tax/subsidy overlap."
  )
)

api_long <- bind_rows(
  api_long,
  gg_derived_to_add |>
    select(sheet, entry, transaction_code, period, value)
)


# -------------------------
# Attach labels, glossary, and sector info
# -------------------------
api_named <- api_long |>
  left_join(api_code_dict, by = "transaction_code") |>
  mutate(
    oecd_label = normalize_oecd_label(oecd_label)
  )

api_enriched <- api_named |>
  left_join(glossary_tbl, by = "oecd_label") |>
  left_join(sector_dict, by = "sheet") |>
  mutate(
    tx_code_final   = if_else(is.na(tx_code) | tx_code == "", transaction_code, tx_code),
    no_flow_item    = is_no_flow_item(oecd_label, notes),
    generated_label = build_label(tx_code_final, sector_letter, entry, no_flow_item),
    generated_desc  = build_desc(oecd_label, entry, no_flow_item)
  )


# -------------------------
# Apply optional mapping overrides
# -------------------------
pre_final_long <- api_enriched |>
  left_join(mapping_tbl, by = c("sheet", "entry", "oecd_label")) |>
  left_join(transaction_order_tbl, by = "transaction_code") |>
  mutate(
    include_final = if_else(is.na(include_final), TRUE, include_final),
    label = if_else(!is.na(custom_label) & custom_label != "", custom_label, generated_label),
    desc  = if_else(!is.na(custom_desc)  & custom_desc  != "", custom_desc,  generated_desc)
  ) |>
  filter(include_final) |>
  group_by(sheet, entry, oecd_label) |>
  mutate(
    ord = if_else(
      is.na(ord),
      if_else(
        is.na(transaction_order),
        1000L + dense_rank(paste(entry, transaction_code, oecd_label)),
        as.integer(transaction_order)
      ),
      ord
    )
  ) |>
  ungroup() |>
  select(-transaction_order)

# -------------------------
# Resolve label collisions
# -------------------------
# Some balancing/no-flow items are returned under both REV and EXP.
# Because these items should not receive R/S suffixes, both rows can create
# the same final label. We keep one final series per sector-period-label.

mapped_long_api <- pre_final_long |>
  mutate(
    collapse_priority = case_when(
      # After filtering to one ADJUSTMENT basis, REV/EXP duplicates of
      # balancing/no-flow items are equivalent representations. Keep the
      # selection deterministic without using it to repair data problems.
      entry == "REV" ~ 1L,
      entry == "EXP" ~ 2L,
      TRUE ~ 3L
    )
  ) |>
  arrange(sheet, period, label, collapse_priority, ord, entry) |>
  group_by(sheet, period, label) |>
  summarise(
    sector_title = first(sector_title),
    ord = first(ord),
    desc = first(desc),
    value = {
      non_missing <- value[!is.na(value)]
      if (length(non_missing) == 0) NA_real_ else non_missing[1]
    },
    .groups = "drop"
  ) |>
  arrange(sheet, period, ord, label)


# -------------------------
# Build Total Economy table
# -------------------------
# The label dictionary for the S1 aggregate economy output is maintained in
# the input workbook so that naming choices remain outside the script logic.
# Any requested code not returned by the source is documented in QA.

total_order_tbl <- tibble(
  transaction_code = total_economy_transaction_codes,
  ord = seq_along(total_economy_transaction_codes) * 10L
)

total_economy_mapped <- total_api_long |>
  left_join(api_code_dict, by = "transaction_code") |>
  mutate(
    oecd_label_api = oecd_label
  ) |>
  select(-oecd_label) |>
  left_join(total_economy_labels_tbl, by = "transaction_code") |>
  left_join(total_order_tbl, by = "transaction_code") |>
  mutate(
    oecd_label = dplyr::coalesce(
      oecd_label_total,
      oecd_label_api,
      paste0("OECD transaction ", transaction_code)
    ),
    oecd_label = normalize_oecd_label(oecd_label),
    notes = if_else(is.na(notes_total), "", notes_total),
    tx_code_final = if_else(
      !is.na(short_label) & short_label != "",
      short_label,
      transaction_code
    ),
    no_flow_item = if_else(is.na(no_flow_item), FALSE, no_flow_item),
    ord = if_else(is.na(ord), 10000L + dense_rank(transaction_code), ord),
    entry_suffix = case_when(
      no_flow_item ~ "",
      entry == "REV" ~ "R",
      entry == "EXP" ~ "S",
      TRUE ~ ""
    ),
    label = paste0(tx_code_final, entry_suffix),
    desc = case_when(
      no_flow_item ~ oecd_label,
      entry == "REV" ~ paste0(oecd_label, ", received / resources (C)"),
      entry == "EXP" ~ paste0(oecd_label, ", spent / uses (D)"),
      TRUE ~ oecd_label
    )
  ) |>
  select(sheet, period, entry, transaction_code, oecd_label, ord, label, desc, value) |>
  arrange(sheet, period, ord, label, entry) |>
  group_by(sheet, period, label) |>
  summarise(
    entry = paste(sort(unique(entry)), collapse = ", "),
    transaction_code = first(transaction_code),
    oecd_label = first(oecd_label),
    ord = first(ord),
    desc = first(desc),
    value = {
      non_missing <- value[!is.na(value)]
      if (length(non_missing) == 0) NA_real_ else non_missing[1]
    },
    .groups = "drop"
  ) |>
  arrange(period, ord, label)


# -------------------------
# QA checks
# -------------------------
qa_missing_api_label <- api_named |>
  filter(is.na(oecd_label) | oecd_label == "") |>
  distinct(transaction_code) |>
  arrange(transaction_code)

qa_missing_glossary <- api_enriched |>
  filter(is.na(tx_code) | tx_code == "") |>
  distinct(transaction_code, oecd_label, entry, sheet) |>
  arrange(transaction_code, sheet, entry)

qa_duplicate_glossary <- glossary_raw |>
  mutate(oecd_label = normalize_oecd_label(oecd_label)) |>
  filter(oecd_label != "") |>
  count(oecd_label) |>
  filter(n > 1)

qa_duplicate_api_codes <- api_code_raw |>
  mutate(transaction_code = normalize_text(transaction_code)) |>
  count(transaction_code) |>
  filter(transaction_code != "", n > 1)

qa_duplicate_generated_labels <- mapped_long_api |>
  count(sheet, period, label) |>
  filter(n > 1) |>
  arrange(sheet, period, label)

qa_label_collisions_before_collapse <- pre_final_long |>
  count(sheet, period, label, name = "rows_before_collapse") |>
  filter(rows_before_collapse > 1) |>
  arrange(sheet, period, label)

qa_unused_mapping_rows <- mapping_tbl |>
  anti_join(
    api_enriched |> distinct(sheet, entry, oecd_label),
    by = c("sheet", "entry", "oecd_label")
  ) |>
  arrange(sheet, entry, ord, oecd_label)

qa_api_all_rows <- api_enriched |>
  distinct(sheet, entry, transaction_code, oecd_label, tx_code_final, generated_label, generated_desc) |>
  arrange(sheet, entry, transaction_code)

qa_collapsed_balance_items <- pre_final_long |>
  distinct(sheet, period, label, oecd_label, entry, transaction_code, tx_code_final) |>
  count(sheet, period, label) |>
  filter(n > 1) |>
  arrange(sheet, period, label)

# For every collapsed label, show whether candidate source values agree.
# This prevents a collision from being mistaken for harmless REV/EXP duplication.
qa_collapsed_value_check <- pre_final_long |>
  group_by(sheet, period, label) |>
  summarise(
    candidate_rows = n(),
    non_missing_values = sum(!is.na(value)),
    distinct_non_missing_values = n_distinct(value[!is.na(value)]),
    min_value = if (all(is.na(value))) NA_real_ else min(value, na.rm = TRUE),
    max_value = if (all(is.na(value))) NA_real_ else max(value, na.rm = TRUE),
    value_range = if (all(is.na(value))) NA_real_ else max(value, na.rm = TRUE) - min(value, na.rm = TRUE),
    entries = paste(sort(unique(entry)), collapse = ", "),
    transaction_codes = paste(sort(unique(transaction_code)), collapse = ", "),
    .groups = "drop"
  ) |>
  filter(candidate_rows > 1) |>
  arrange(desc(value_range), sheet, period, label)

qa_sector_transaction_expected <- tibble(
  sheet = "ROW",
  transaction_code = c("P6", "P7"),
  expected_entry = c("EXP", "REV"),
  oecd_label = c("Exports of goods and services", "Imports of goods and services"),
  expected_accounting_position = c(
    paste0(country_name, " exports are recorded as expenditure by the Rest of the World."),
    paste0(country_name, " imports are recorded as receipts by the Rest of the World.")
  )
)

qa_sector_transaction_returned <- api_long |>
  filter(sheet == "ROW", transaction_code %in% c("P6", "P7")) |>
  group_by(transaction_code, entry) |>
  summarise(
    rows_returned = n(),
    first_period = min(period, na.rm = TRUE),
    last_period = max(period, na.rm = TRUE),
    .groups = "drop"
  )

qa_sector_transaction_availability <- qa_sector_transaction_expected |>
  left_join(
    qa_sector_transaction_returned,
    by = c("transaction_code", "expected_entry" = "entry")
  ) |>
  mutate(
    rows_returned = if_else(is.na(rows_returned), 0L, as.integer(rows_returned)),
    present = rows_returned > 0,
    status = if_else(
      present,
      "OK: expected transaction returned for the expected accounting entry.",
      "MISSING: expected transaction was not returned for the expected accounting entry."
    )
  ) |>
  arrange(transaction_code)

qa_total_economy_codes <- tibble(
  transaction_code = total_economy_transaction_codes
) |>
  left_join(
    total_api_long |>
      group_by(transaction_code) |>
      summarise(
        rows_returned = n(),
        entries_returned = paste(sort(unique(entry)), collapse = ", "),
        first_period = min(period, na.rm = TRUE),
        last_period = max(period, na.rm = TRUE),
        .groups = "drop"
      ),
    by = "transaction_code"
  ) |>
  mutate(
    rows_returned = if_else(is.na(rows_returned), 0L, as.integer(rows_returned)),
    present = rows_returned > 0,
    status = if_else(
      present,
      "OK: code returned in the S1 aggregate economy API response.",
      "MISSING: code is included in the S1 aggregate economy transaction list but was not returned by the API."
    )
  ) |>
  arrange(match(transaction_code, total_economy_transaction_codes))

# -------------------------
# Build clean workbook
# -------------------------
# The output workbook is assembled as a named list of sheets and written with
# writexl, after converting all sheets to a consistent rectangular structure.

as_char_df <- function(df) {
  df <- as.data.frame(df, stringsAsFactors = FALSE)
  if (ncol(df) == 0) {
    df <- data.frame(V1 = character(), stringsAsFactors = FALSE)
  }
  df[] <- lapply(df, function(x) {
    x <- as.character(x)
    x[is.na(x)] <- ""
    x
  })
  names(df) <- paste0("V", seq_len(ncol(df)))
  df
}

pad_cols <- function(df, n) {
  df <- as_char_df(df)
  if (ncol(df) < n) {
    for (j in (ncol(df) + 1):n) {
      df[[paste0("V", j)]] <- ""
    }
  }
  df <- df[, paste0("V", seq_len(n)), drop = FALSE]
  df
}

stack_rows <- function(...) {
  pieces <- list(...)
  max_cols <- as.integer(max(vapply(pieces, function(x) max(1L, ncol(as.data.frame(x))), integer(1))))
  dplyr::bind_rows(lapply(pieces, pad_cols, n = max_cols))
}

make_sector_export <- function(obj) {
  blank <- data.frame(V1 = "", stringsAsFactors = FALSE)
  stack_rows(obj$meta, blank, obj$table)
}

make_qa_export <- function(df, title) {
  df <- as.data.frame(df, stringsAsFactors = FALSE)

  if (nrow(df) == 0 || ncol(df) == 0) {
    df <- data.frame(Status = "No issues found.", stringsAsFactors = FALSE)
  }

  title_row <- data.frame(V1 = title, stringsAsFactors = FALSE)
  blank_row <- data.frame(V1 = "", stringsAsFactors = FALSE)

  header_row <- as.data.frame(as.list(names(df)), stringsAsFactors = FALSE)
  names(header_row) <- paste0("V", seq_len(ncol(header_row)))

  data_rows <- df
  names(data_rows) <- paste0("V", seq_len(ncol(data_rows)))

  stack_rows(title_row, blank_row, header_row, data_rows)
}

sheet_list <- list()

for (sh in sector_dict$sheet) {
  obj <- build_sheet_df(mapped_long_api, sh, country_name, units_label)
  if (!is.null(obj)) {
    sheet_list[[sh]] <- make_sector_export(obj)
  }
}

total_obj <- build_total_economy_sheet_df(
  total_economy_mapped,
  country_name,
  units_label,
  url_total_economy
)

if (!is.null(total_obj)) {
  sheet_list[["TOTAL_ECONOMY"]] <- make_sector_export(total_obj)
}

sheet_list[["QA_missing_api_label"]] <- make_qa_export(
  qa_missing_api_label,
  "API transaction codes not matched to an OECD label in api_code_dict."
)
sheet_list[["QA_missing_glossary"]] <- make_qa_export(
  qa_missing_glossary,
  "API rows whose OECD label is not matched to a tx_code in glossary. Fallback label used transaction_code."
)
sheet_list[["QA_duplicate_glossary"]] <- make_qa_export(
  qa_duplicate_glossary,
  "Duplicate OECD labels in the glossary sheet."
)
sheet_list[["QA_duplicate_api_codes"]] <- make_qa_export(
  qa_duplicate_api_codes,
  "Duplicate transaction codes in api_code_dict."
)
sheet_list[["QA_duplicate_generated_labels"]] <- make_qa_export(
  qa_duplicate_generated_labels,
  "More than one final row still has the same label after collision resolution."
)
sheet_list[["QA_label_collisions_before"]] <- make_qa_export(
  qa_label_collisions_before_collapse,
  "Final-label collisions detected before collision resolution."
)
sheet_list[["QA_collapsed_balance_items"]] <- make_qa_export(
  qa_collapsed_balance_items,
  "Final-label collisions collapsed to one series per sheet-period-label."
)
sheet_list[["QA_collapsed_value_check"]] <- make_qa_export(
  qa_collapsed_value_check,
  "Candidate source values behind every collapsed final label; non-zero ranges require review."
)
sheet_list[["QA_sector_source_duplicates"]] <- make_qa_export(
  qa_sector_source_duplicates,
  "Duplicate raw sector observations before Stage-1 aggregation. The script stops if any are present."
)
sheet_list[["QA_total_source_duplicates"]] <- make_qa_export(
  qa_total_source_duplicates,
  "Duplicate raw S1 observations before Stage-1 aggregation. The script stops if any are present."
)
sheet_list[["QA_transaction_availability"]] <- make_qa_export(
  qa_sector_transaction_availability,
  "Validation of expected transaction availability in the sectoral output."
)
sheet_list[["QA_total_economy_codes"]] <- make_qa_export(
  qa_total_economy_codes,
  "Validation of transaction-code availability in the S1 aggregate economy query."
)
sheet_list[["QA_GG_source_status"]] <- make_qa_export(
  qa_gg_source_status,
  "Quarterly General Government source status and derivation policy."
)
sheet_list[["QA_GG_derivation"]] <- make_qa_export(
  qa_gg_derivation,
  "Quarterly GG rows derived transparently from S1 and observed non-GG sectors when direct S13 rows are unavailable."
)
sheet_list[["QA_unused_mapping_rows"]] <- make_qa_export(
  qa_unused_mapping_rows,
  "Rows in mapping that do not match any row actually returned by the API."
)
sheet_list[["API_rows_used"]] <- make_qa_export(
  qa_api_all_rows,
  "All API rows actually returned and relabeled."
)
sheet_list[["Glossary_used"]] <- make_qa_export(
  glossary_tbl,
  "Structured glossary used by the script."
)
sheet_list[["Mapping_used"]] <- make_qa_export(
  mapping_tbl,
  "Optional mapping overrides used by the script."
)

writexl::write_xlsx(sheet_list, path = output_file, col_names = FALSE)

# Workbook readability check.
written_sheets <- openxlsx::getSheetNames(output_file)
if (length(written_sheets) == 0) {
  stop("Workbook was written, but openxlsx still cannot read any worksheets.")
}

message("Clean workbook written to: ", output_file)
message("Sheets written: ", paste(written_sheets, collapse = ", "))
