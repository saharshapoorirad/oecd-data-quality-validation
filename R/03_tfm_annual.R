# =============================================================================
# TFM_master_annual_final_CORRECTED_FULL.R
#
# Annual non-financial Transaction-Flow Matrix (TFM) builder
# Italy — Stage 2 of the OECD non-financial accounts pipeline
#
# INPUT:
#   outputs/ITA_stage1_nonfinancial_accounts_annual_output.xlsx
#   Numbered copies such as (...)(1).xlsx are also detected automatically.
#
# OUTPUT:
#   outputs/ITA_stage2_TFM_annual_output.xlsx
#
# =============================================================================

options(stringsAsFactors = FALSE)

# -------------------------
# Packages
# -------------------------
install_if_missing <- function(pkgs) {
  to_install <- pkgs[!pkgs %in% rownames(installed.packages())]
  if (length(to_install) > 0) {
    install.packages(to_install)
  }
}

install_if_missing(c(
  "openxlsx", "dplyr", "tidyr", "tibble",
  "stringr", "purrr", "writexl"
))

suppressPackageStartupMessages({
  library(openxlsx)
  library(dplyr)
  library(tidyr)
  library(tibble)
  library(stringr)
  library(purrr)
  library(writexl)
})

# -------------------------
# User inputs
# -------------------------
country_name <- "Italy"
country_code <- "ITA"
input_dir <- "outputs"
output_dir <- "outputs"
dir.create(output_dir, showWarnings = FALSE, recursive = TRUE)

exact_input <- file.path(
  input_dir,
  paste0(
    country_code,
    "_stage1_nonfinancial_accounts_annual_output.xlsx"
  )
)

input_pattern <- paste0(
  "^",
  country_code,
  "_stage1_nonfinancial_accounts_annual_output(\\([0-9]+\\))?\\.xlsx$"
)

find_input_file <- function(exact_name, pattern, search_dir) {
  if (file.exists(exact_name)) {
    return(exact_name)
  }

  hits <- list.files(
    path = search_dir,
    pattern = pattern,
    full.names = TRUE
  )

  if (length(hits) == 0) {
    stop(
      "Stage-1 annual output not found in:\n",
      normalizePath(search_dir, winslash = "/", mustWork = FALSE),
      "\n\nExpected a file named like:\n  ",
      exact_name,
      "\n  ",
      sub("\\.xlsx$", "(1).xlsx", exact_name)
    )
  }

  info <- file.info(hits)
  hits[order(info$mtime, decreasing = TRUE)][1]
}

input_file <- find_input_file(exact_input, input_pattern, input_dir)
output_file <- file.path(
  output_dir,
  paste0(country_code, "_stage2_TFM_annual_output.xlsx")
)

message("Using input file: ", input_file)
message("Output file: ", output_file)
message("Standalone corrected annual Stage-2 TFM script.")
message("No KADJ/balancing adjustment is inserted.")

# -------------------------
# Core dictionaries
# -------------------------
sector_dictionary <- tribble(
  ~Sheet, ~Matrix_column, ~Official_sector, ~Meaning,
  "NFC", "S11_NFC",     "S11",     "Non-financial corporations",
  "FC",  "S12_FC",      "S12",     "Financial corporations",
  "GG",  "S13_GG",      "S13",     "General government",
  "HH",  "S14_S15_HH",  "S14+S15", "Households and NPISH combined",
  "ROW", "S2_ROW",      "S2",      "Rest of the world"
)

matrix_sector_cols <- sector_dictionary$Matrix_column

row_dictionary <- tribble(
  ~Row_order, ~Code,           ~Item,                                                               ~Row_type,      ~Zero_sum_test,
  10L,        "P1",            "Output",                                                             "Transaction", "Production_constructed",
  20L,        "P2",            "Intermediate consumption",                                          "Transaction", "Production_constructed",
  30L,        "D21",           "Taxes on products",                                                  "Transaction", "Production_constructed",
  40L,        "D31",           "Subsidies on products",                                              "Transaction", "Production_constructed",
  50L,        "B1GQ",          "Derived production balance: P1 + P2 + D21 + D31",                    "Derived",     "Production_constructed",
  55L,        "B3G",           "Mixed income, gross",                                                    "Balancing",   "Not_applicable",
  60L,        "P3",            "Final consumption expenditure",                                      "Transaction", "Production_constructed",
  70L,        "P5",            "Gross capital formation",                                            "Transaction", "Production_constructed",
  80L,        "P7",            "Imports of goods and services / ROW receipt",                        "Transaction", "Production_constructed",
  90L,        "P6",            "Exports of goods and services / ROW payment",                        "Transaction", "Production_constructed",
  100L,       "D1",            "Compensation of employees",                                          "Transaction", "Observed_sector_closure",
  110L,       "D29",           "Other taxes on production",                                          "Transaction", "Observed_sector_closure",
  120L,       "D5",            "Current taxes on income, wealth, etc.",                              "Transaction", "Observed_sector_closure",
  130L,       "D39",           "Other subsidies on production",                                      "Transaction", "Observed_sector_closure",
  140L,       "D61",           "Net social contributions",                                           "Transaction", "Observed_sector_closure",
  150L,       "D62",           "Social benefits other than social transfers in kind",                "Transaction", "Observed_sector_closure",
  160L,       "D7",            "Other current transfers",                                            "Transaction", "Observed_sector_closure",
  170L,       "D41",           "Interest",                                                            "Transaction", "Observed_sector_closure",
  180L,       "D42",           "Distributed income of corporations",                                 "Transaction", "Observed_sector_closure",
  190L,       "D43",           "Reinvested earnings on foreign direct investment",                   "Transaction", "Observed_sector_closure",
  200L,       "D44",           "Other investment income",                                            "Transaction", "Observed_sector_closure",
  210L,       "D45",           "Rent",                                                                "Transaction", "Observed_sector_closure",
  220L,       "D8",            "Adjustment for the change in pension entitlements",                  "Transaction", "Observed_sector_closure",
  230L,       "B8G",           "Saving, gross",                                                       "Balancing",   "Not_applicable",
  240L,       "D9",            "Capital transfers",                                                   "Transaction", "Observed_sector_closure",
  250L,       "NP",            "Acquisitions less disposals of non-produced non-financial assets",   "Transaction", "Observed_sector_closure",
  260L,       "XB",            "Current external balance",                                            "Balancing",   "Not_applicable",
  270L,       "NL_CALC",       "Net lending calculated from saving and capital account",             "Validation",  "Not_applicable",
  280L,       "B9",            "Official OECD net lending / net borrowing",                           "Benchmark",   "Not_applicable",
  290L,       "NL_ERROR",      "B9 minus calculated net lending",                                     "Validation",  "Not_applicable",
  300L,       "ROW_EXPECTED",  "ROW net lending implied by domestic official B9",                    "Validation",  "Not_applicable",
  310L,       "ROW_ERROR",     "ROW expected minus ROW calculated net lending",                       "Validation",  "Not_applicable"
)

production_rows <- c(
  "P1", "P2", "D21", "D31", "B1GQ",
  "P3", "P5", "P7", "P6"
)

# -------------------------
# Read helpers
# -------------------------
read_stage1_sheet <- function(file, sheet) {
  available_sheets <- openxlsx::getSheetNames(file)

  if (!sheet %in% available_sheets) {
    stop("Missing required sheet: ", sheet)
  }

  preview <- openxlsx::read.xlsx(
    file,
    sheet = sheet,
    startRow = 1,
    rows = 1:20,
    colNames = FALSE,
    skipEmptyRows = FALSE,
    skipEmptyCols = FALSE
  )

  year_names <- c("Year", "Period", "year")
  header_row <- NA_integer_

  for (i in seq_len(nrow(preview))) {
    vals <- stringr::str_squish(
      as.character(unlist(preview[i, ], use.names = FALSE))
    )

    if (any(vals %in% year_names, na.rm = TRUE)) {
      header_row <- i
      break
    }
  }

  if (is.na(header_row)) {
    stop(
      "Sheet ", sheet,
      " does not contain a Year/Period header in the first 20 rows."
    )
  }

  x <- openxlsx::read.xlsx(
    file,
    sheet = sheet,
    startRow = header_row,
    colNames = TRUE
  )

  names(x) <- stringr::str_squish(names(x))

  empty_col <- vapply(
    x,
    function(z) all(is.na(z) | trimws(as.character(z)) == ""),
    logical(1)
  )

  x <- x[, !empty_col, drop = FALSE]

  year_col <- intersect(year_names, names(x))

  if (length(year_col) == 0) {
    stop(
      "Sheet ", sheet,
      " header was detected at row ", header_row,
      " but no usable Year/Period column remained after reading."
    )
  }

  names(x)[names(x) == year_col[1]] <- "Year"

  x$Year <- suppressWarnings(
    as.integer(
      sub(
        "^([0-9]{4}).*$",
        "\\1",
        stringr::str_squish(as.character(x$Year))
      )
    )
  )

  x <- x[!is.na(x$Year), , drop = FALSE]

  x
}

year_order <- function(x) {
  suppressWarnings(as.integer(as.character(x)))
}

# Missing direct source series stay NA.
col_or_na <- function(df, colname) {
  if (colname %in% names(df)) {
    return(
      suppressWarnings(as.numeric(df[[colname]]))
    )
  }

  rep(NA_real_, nrow(df))
}

# Deliberate model/design zero only.
structural_zero <- function(df) {
  rep(0, nrow(df))
}

# Resources minus uses.
# If both source columns are absent => NA.
# If one side is present => absent side treated as zero.
net_rs <- function(df, received_col, spent_col) {
  has_r <- received_col %in% names(df)
  has_s <- spent_col %in% names(df)

  if (!has_r && !has_s) {
    return(rep(NA_real_, nrow(df)))
  }

  r <- if (has_r) {
    suppressWarnings(as.numeric(df[[received_col]]))
  } else {
    rep(0, nrow(df))
  }

  s <- if (has_s) {
    suppressWarnings(as.numeric(df[[spent_col]]))
  } else {
    rep(0, nrow(df))
  }

  r - s
}

# Sum available components; NA only when all are unavailable.
sum_available <- function(...) {
  xs <- list(...)
  m <- do.call(cbind, xs)

  out <- rowSums(m, na.rm = TRUE)
  out[rowSums(!is.na(m)) == 0] <- NA_real_

  out
}

# Sum only when every required component exists.
sum_required <- function(...) {
  xs <- list(...)
  m <- do.call(cbind, xs)

  out <- rowSums(m, na.rm = FALSE)
  out[rowSums(is.na(m)) > 0] <- NA_real_

  out
}

subtract_required <- function(total, component) {
  out <- total - component
  out[is.na(total) | is.na(component)] <- NA_real_
  out
}

# -------------------------
# Read Stage-1 sheets
# -------------------------
data_N  <- read_stage1_sheet(input_file, "NFC")
data_F  <- read_stage1_sheet(input_file, "FC")
data_G  <- read_stage1_sheet(input_file, "GG")
data_H  <- read_stage1_sheet(input_file, "HH")
data_R  <- read_stage1_sheet(input_file, "ROW")
data_TE <- read_stage1_sheet(input_file, "TOTAL_ECONOMY")

# -------------------------
# Input availability QA
# -------------------------
expected_inputs <- tribble(
  ~Sheet, ~TFM_code, ~Source_columns, ~Use_note,

  "NFC", "P1",  "OUTn",          "Direct",
  "NFC", "P2",  "ICnS",          "Direct spent side",
  "NFC", "D1",  "WBnS",          "Direct spent side",
  "NFC", "D29", "ITnS",          "Derived as D2 total minus structural-zero D21",
  "NFC", "D5",  "CTAXnS",          "Direct spent side",
  "NFC", "D39", "SUBnR",         "Derived as D3 total minus structural-zero D31",
  "NFC", "D61", "SCnR",          "Direct received side",
  "NFC", "D62", "SBnS",          "Direct spent side",
  "NFC", "D7",  "TRnR|TRnS",     "Received minus spent",
  "NFC", "D41", "INTnR|INTnS",   "Received minus spent",
  "NFC", "D42", "DIVnR|DIVnS",   "Received minus spent",
  "NFC", "D43", "REInR|REInS",   "Received minus spent",
  "NFC", "D44", "OInR|OInS",     "Received minus spent",
  "NFC", "D45", "RNTnR|RNTnS",   "Received minus spent",
  "NFC", "D8",  "PENAnR|PENAnS", "Received minus spent",
  "NFC", "D9",  "KTnR|KTnS",     "Received minus spent",
  "NFC", "NP",  "NPnR|NPnS",     "Received minus spent",
  "NFC", "P5",  "GCFnS",         "Direct spent side",
  "NFC", "B8G", "SAVn",          "Direct balancing item",
  "NFC", "B9",  "NETLn",         "Direct benchmark",

  "FC",  "P1",  "OUTf",          "Direct",
  "FC",  "P2",  "ICfS",          "Direct spent side",
  "FC",  "D1",  "WBfS",          "Direct spent side",
  "FC",  "D29", "ITfS",          "Derived as D2 total minus structural-zero D21",
  "FC",  "D5",  "CTAXfS",          "Direct spent side",
  "FC",  "D39", "SUBfR",         "Derived as D3 total minus structural-zero D31",
  "FC",  "D61", "SCfR",          "Direct received side",
  "FC",  "D62", "SBfS",          "Direct spent side",
  "FC",  "D7",  "TRfR|TRfS",     "Received minus spent",
  "FC",  "D41", "INTfR|INTfS",   "Received minus spent",
  "FC",  "D42", "DIVfR|DIVfS",   "Received minus spent",
  "FC",  "D43", "REIfR|REIfS",   "Received minus spent",
  "FC",  "D44", "OIfR|OIfS",     "Received minus spent",
  "FC",  "D45", "RNTfR|RNTfS",   "Received minus spent",
  "FC",  "D8",  "PENAfR|PENAfS", "Received minus spent",
  "FC",  "D9",  "KTfR|KTfS",     "Received minus spent",
  "FC",  "NP",  "NPfR|NPfS",     "Received minus spent",
  "FC",  "P5",  "GCFfS",         "Direct spent side",
  "FC",  "B8G", "SAVf",          "Direct balancing item",
  "FC",  "B9",  "NETLf",         "Direct benchmark",

  "GG",  "P1",  "OUTg",          "Direct",
  "GG",  "P2",  "ICgS",          "Direct spent side",
  "GG",  "D21", "TPgR|TPgS",     "Received minus spent",
  "GG",  "D31", "SPgR|SPgS",     "Received minus spent",
  "GG",  "D1",  "WBgR|WBgS",     "Received minus spent",
  "GG",  "D29", "ITgR|ITgS|TPgR|TPgS", "Derived as D2 total minus D21",
  "GG",  "D5",  "CTAXgR|CTAXgS",     "Received minus spent",
  "GG",  "D39", "SUBgR|SUBgS|SPgR|SPgS", "Derived as D3 total minus D31",
  "GG",  "D61", "SCgR|SCgS",     "Received minus spent",
  "GG",  "D62", "SBgR|SBgS",     "Received minus spent",
  "GG",  "D7",  "TRgR|TRgS",     "Received minus spent",
  "GG",  "D41", "INTgR|INTgS",   "Received minus spent",
  "GG",  "D42", "DIVgR|DIVgS",   "Received minus spent",
  "GG",  "D43", "REIgR|REIgS",   "Received minus spent",
  "GG",  "D44", "OIgR|OIgS",     "Received minus spent",
  "GG",  "D45", "RNTgR|RNTgS",   "Received minus spent",
  "GG",  "D8",  "PENAgR|PENAgS", "Received minus spent",
  "GG",  "D9",  "KTgR|KTgS",     "Received minus spent",
  "GG",  "NP",  "NPgR|NPgS",     "Received minus spent",
  "GG",  "P3",  "CHgS",          "Direct spent side",
  "GG",  "P5",  "GCFgS",         "Direct spent side",
  "GG",  "B8G", "SAVg",          "Direct balancing item",
  "GG",  "B9",  "NETLg",         "Direct benchmark",

  "HH",  "P1",  "IChS|WBhS|IThS|SUBhR|GOSh|MIh", "Derived from annual production/generation-of-income identity when direct OUTh is unavailable",
  "HH",  "P2",  "IChS",          "Direct spent side",
  "HH",  "D1",  "WBhR|WBhS",     "Received minus spent",
  "HH",  "D29", "IThR|IThS",     "Derived as D2 total minus structural-zero D21",
  "HH",  "D5",  "CTAXhR|CTAXhS",     "Received minus spent",
  "HH",  "D39", "SUBhR|SUBhS",   "Derived as D3 total minus structural-zero D31",
  "HH",  "D61", "SChR|SChS",     "Received minus spent",
  "HH",  "D62", "SBhR|SBhS",     "Received minus spent",
  "HH",  "D7",  "TRhR|TRhS",     "Received minus spent",
  "HH",  "D41", "INThR|INThS",   "Received minus spent",
  "HH",  "D42", "DIVhR|DIVhS",   "Received minus spent",
  "HH",  "D43", "REIhR|REIhS",   "Received minus spent",
  "HH",  "D44", "OIhR|OIhS",     "Received minus spent",
  "HH",  "D45", "RNThR|RNThS",   "Received minus spent",
  "HH",  "D8",  "PENAhR|PENAhS", "Received minus spent",
  "HH",  "D9",  "KThR|KThS",     "Received minus spent",
  "HH",  "NP",  "NPhR|NPhS",     "Received minus spent",
  "HH",  "P3",  "CHhS",          "Direct spent side",
  "HH",  "P5",  "GCFhS",         "Direct spent side",
  "HH",  "B8G", "SAVh",          "Direct balancing item",
  "HH",  "B9",  "NETLh",         "Direct benchmark",
  "HH", "B3G", "MIh",          "Direct mixed-income balancing item",

  "ROW", "D21", "TPrR|TPrS",     "Received minus spent",
  "ROW", "D31", "SPrR|SPrS",     "Received minus spent",
  "ROW", "D1",  "WBrR|WBrS",     "Received minus spent",
  "ROW", "D29", "ITrR|ITrS|TPrR|TPrS", "Derived as D2 total minus D21",
  "ROW", "D5",  "CTAXrR|CTAXrS",     "Received minus spent",
  "ROW", "D39", "SUBrR|SUBrS|SPrR|SPrS", "Derived as D3 total minus D31",
  "ROW", "D61", "SCrR|SCrS",     "Received minus spent",
  "ROW", "D62", "SBrR|SBrS",     "Received minus spent",
  "ROW", "D7",  "TRrR|TRrS",     "Received minus spent",
  "ROW", "D41", "INTrR|INTrS",   "Received minus spent",
  "ROW", "D42", "DIVrR|DIVrS",   "Received minus spent",
  "ROW", "D43", "REIrR|REIrS",   "Received minus spent",
  "ROW", "D44", "OIrR|OIrS",     "Received minus spent",
  "ROW", "D45", "RNTrR|RNTrS",   "Received minus spent; annual orientation closes directly",
  "ROW", "D8",  "PENArR|PENArS", "Received minus spent",
  "ROW", "D9",  "KTrR|KTrS",     "Received minus spent",
  "ROW", "NP",  "NPrR|NPrS",     "Received minus spent",
  "ROW", "P6",  "XGSrS",         "Direct ROW payment",
  "ROW", "P7",  "MGSrR",         "Direct ROW receipt",
  "ROW", "XB",  "XBr",           "Direct balancing item"
)

sheet_lookup <- list(
  NFC = data_N,
  FC = data_F,
  GG = data_G,
  HH = data_H,
  ROW = data_R
)

input_availability <- expected_inputs %>%
  rowwise() %>%
  mutate(
    Present_columns = {
      cols <- strsplit(Source_columns, "\\|")[[1]]
      paste(
        cols[cols %in% names(sheet_lookup[[Sheet]])],
        collapse = " | "
      )
    },

    Missing_columns = {
      cols <- strsplit(Source_columns, "\\|")[[1]]
      paste(
        cols[!cols %in% names(sheet_lookup[[Sheet]])],
        collapse = " | "
      )
    },

    Any_source_present = {
      cols <- strsplit(Source_columns, "\\|")[[1]]
      any(cols %in% names(sheet_lookup[[Sheet]]))
    },

    All_sources_present = {
      cols <- strsplit(Source_columns, "\\|")[[1]]
      all(cols %in% names(sheet_lookup[[Sheet]]))
    },

    Availability_status = case_when(
      All_sources_present ~
        "All listed source columns present",

      Any_source_present ~
        "Partial source pair/list; absent side treated as zero where net flow logic applies",

      TRUE ~
        "Unavailable in Stage-1 sheet; resulting direct TFM value is NA"
    )
  ) %>%
  ungroup() %>%
  select(
    Sheet, TFM_code, Source_columns, Use_note,
    Present_columns, Missing_columns,
    Any_source_present, All_sources_present,
    Availability_status
  )

# -------------------------
# Sector builders
# -------------------------
make_S11 <- function(d) {
  p1 <- col_or_na(d, "OUTn")
  p2 <- -col_or_na(d, "ICnS")

  d21 <- structural_zero(d)
  d31 <- structural_zero(d)

  d2_total <- -col_or_na(d, "ITnS")
  d3_total <- col_or_na(d, "SUBnR")

  d29 <- subtract_required(d2_total, d21)
  d39 <- subtract_required(d3_total, d31)

  p5 <- -col_or_na(d, "GCFnS")
  d9 <- net_rs(d, "KTnR", "KTnS")
  np <- net_rs(d, "NPnR", "NPnS")
  b8g <- col_or_na(d, "SAVn")
  b9 <- col_or_na(d, "NETLn")

  nl_calc <- sum_required(b8g, d9, p5, np)

  tibble(
    Year = d$Year,

    P1 = p1,
    P2 = p2,
    D21 = d21,
    D31 = d31,
    B1GQ = sum_required(p1, p2, d21, d31),
    B3G = structural_zero(d),

    P3 = structural_zero(d),
    P5 = p5,
    P7 = structural_zero(d),
    P6 = structural_zero(d),

    D1 = -col_or_na(d, "WBnS"),
    D29 = d29,
    D5 = -col_or_na(d, "CTAXnS"),
    D39 = d39,

    D61 = col_or_na(d, "SCnR"),
    D62 = -col_or_na(d, "SBnS"),
    D7 = net_rs(d, "TRnR", "TRnS"),

    D41 = net_rs(d, "INTnR", "INTnS"),
    D42 = net_rs(d, "DIVnR", "DIVnS"),
    D43 = net_rs(d, "REInR", "REInS"),
    D44 = net_rs(d, "OInR", "OInS"),
    D45 = net_rs(d, "RNTnR", "RNTnS"),

    D8 = net_rs(d, "PENAnR", "PENAnS"),

    B8G = b8g,
    D9 = d9,
    NP = np,
    XB = NA_real_,

    NL_CALC = nl_calc,
    B9 = b9,
    NL_ERROR = b9 - nl_calc,

    ROW_EXPECTED = NA_real_,
    ROW_ERROR = NA_real_
  )
}

make_S12 <- function(d) {
  p1 <- col_or_na(d, "OUTf")
  p2 <- -col_or_na(d, "ICfS")

  d21 <- structural_zero(d)
  d31 <- structural_zero(d)

  d2_total <- -col_or_na(d, "ITfS")
  d3_total <- col_or_na(d, "SUBfR")

  d29 <- subtract_required(d2_total, d21)
  d39 <- subtract_required(d3_total, d31)

  p5 <- -col_or_na(d, "GCFfS")
  d9 <- net_rs(d, "KTfR", "KTfS")
  np <- net_rs(d, "NPfR", "NPfS")
  b8g <- col_or_na(d, "SAVf")
  b9 <- col_or_na(d, "NETLf")

  nl_calc <- sum_required(b8g, d9, p5, np)

  tibble(
    Year = d$Year,

    P1 = p1,
    P2 = p2,
    D21 = d21,
    D31 = d31,
    B1GQ = sum_required(p1, p2, d21, d31),
    B3G = structural_zero(d),

    P3 = structural_zero(d),
    P5 = p5,
    P7 = structural_zero(d),
    P6 = structural_zero(d),

    D1 = -col_or_na(d, "WBfS"),
    D29 = d29,
    D5 = -col_or_na(d, "CTAXfS"),
    D39 = d39,

    D61 = col_or_na(d, "SCfR"),
    D62 = -col_or_na(d, "SBfS"),
    D7 = net_rs(d, "TRfR", "TRfS"),

    D41 = net_rs(d, "INTfR", "INTfS"),
    D42 = net_rs(d, "DIVfR", "DIVfS"),
    D43 = net_rs(d, "REIfR", "REIfS"),
    D44 = net_rs(d, "OIfR", "OIfS"),
    D45 = net_rs(d, "RNTfR", "RNTfS"),

    D8 = net_rs(d, "PENAfR", "PENAfS"),

    B8G = b8g,
    D9 = d9,
    NP = np,
    XB = NA_real_,

    NL_CALC = nl_calc,
    B9 = b9,
    NL_ERROR = b9 - nl_calc,

    ROW_EXPECTED = NA_real_,
    ROW_ERROR = NA_real_
  )
}

make_S13 <- function(d) {
  p1 <- col_or_na(d, "OUTg")
  p2 <- -col_or_na(d, "ICgS")

  d21 <- net_rs(d, "TPgR", "TPgS")
  d31 <- net_rs(d, "SPgR", "SPgS")

  d2_total <- net_rs(d, "ITgR", "ITgS")
  d3_total <- net_rs(d, "SUBgR", "SUBgS")

  d29 <- subtract_required(d2_total, d21)
  d39 <- subtract_required(d3_total, d31)

  p5 <- -col_or_na(d, "GCFgS")
  d9 <- net_rs(d, "KTgR", "KTgS")
  np <- net_rs(d, "NPgR", "NPgS")
  b8g <- col_or_na(d, "SAVg")
  b9 <- col_or_na(d, "NETLg")

  nl_calc <- sum_required(b8g, d9, p5, np)

  tibble(
    Year = d$Year,

    P1 = p1,
    P2 = p2,
    D21 = d21,
    D31 = d31,
    B1GQ = sum_required(p1, p2, d21, d31),
    B3G = structural_zero(d),

    P3 = -col_or_na(d, "CHgS"),
    P5 = p5,
    P7 = structural_zero(d),
    P6 = structural_zero(d),

    D1 = net_rs(d, "WBgR", "WBgS"),
    D29 = d29,
    D5 = net_rs(d, "CTAXgR", "CTAXgS"),
    D39 = d39,

    D61 = net_rs(d, "SCgR", "SCgS"),
    D62 = net_rs(d, "SBgR", "SBgS"),
    D7 = net_rs(d, "TRgR", "TRgS"),

    D41 = net_rs(d, "INTgR", "INTgS"),
    D42 = net_rs(d, "DIVgR", "DIVgS"),
    D43 = net_rs(d, "REIgR", "REIgS"),
    D44 = net_rs(d, "OIgR", "OIgS"),
    D45 = net_rs(d, "RNTgR", "RNTgS"),

    D8 = net_rs(d, "PENAgR", "PENAgS"),

    B8G = b8g,
    D9 = d9,
    NP = np,
    XB = NA_real_,

    NL_CALC = nl_calc,
    B9 = b9,
    NL_ERROR = b9 - nl_calc,

    ROW_EXPECTED = NA_real_,
    ROW_ERROR = NA_real_
  )
}

make_S14_S15 <- function(d) {
  p1_direct <- col_or_na(d, "OUTh")

  # Annual HH output can be reconstructed from the production /
  # generation-of-income identity when direct P1 is unavailable:
  # P1 = P2 + D1_paid + D2_paid - D3_received + B2G + B3G.
  p1_derived <- sum_required(
    col_or_na(d, "IChS"),
    col_or_na(d, "WBhS"),
    col_or_na(d, "IThS"),
    -col_or_na(d, "SUBhR"),
    col_or_na(d, "GOSh"),
    col_or_na(d, "MIh")
  )

  p1 <- ifelse(
    !is.na(p1_direct),
    p1_direct,
    p1_derived
  )

  p2 <- -col_or_na(d, "IChS")

  d21 <- structural_zero(d)
  d31 <- structural_zero(d)

  d2_total <- net_rs(d, "IThR", "IThS")
  d3_total <- net_rs(d, "SUBhR", "SUBhS")

  d29 <- subtract_required(d2_total, d21)
  d39 <- subtract_required(d3_total, d31)

  p5 <- -col_or_na(d, "GCFhS")
  d9 <- net_rs(d, "KThR", "KThS")
  np <- net_rs(d, "NPhR", "NPhS")
  b8g <- col_or_na(d, "SAVh")
  b9 <- col_or_na(d, "NETLh")

  # No balancing adjustment.
  nl_calc <- sum_required(b8g, d9, p5, np)

  tibble(
    Year = d$Year,

    P1 = p1,
    P2 = p2,
    D21 = d21,
    D31 = d31,
    B1GQ = sum_required(p1, p2, d21, d31),
    B3G = col_or_na(d, "MIh"),

    P3 = -col_or_na(d, "CHhS"),
    P5 = p5,
    P7 = structural_zero(d),
    P6 = structural_zero(d),

    D1 = net_rs(d, "WBhR", "WBhS"),
    D29 = d29,
    D5 = net_rs(d, "CTAXhR", "CTAXhS"),
    D39 = d39,

    D61 = net_rs(d, "SChR", "SChS"),
    D62 = net_rs(d, "SBhR", "SBhS"),
    D7 = net_rs(d, "TRhR", "TRhS"),

    D41 = net_rs(d, "INThR", "INThS"),
    D42 = net_rs(d, "DIVhR", "DIVhS"),
    D43 = net_rs(d, "REIhR", "REIhS"),
    D44 = net_rs(d, "OIhR", "OIhS"),
    D45 = net_rs(d, "RNThR", "RNThS"),

    D8 = net_rs(d, "PENAhR", "PENAhS"),

    B8G = b8g,
    D9 = d9,
    NP = np,
    XB = NA_real_,

    NL_CALC = nl_calc,
    B9 = b9,
    NL_ERROR = b9 - nl_calc,

    ROW_EXPECTED = NA_real_,
    ROW_ERROR = NA_real_
  )
}

make_S2 <- function(d, domestic_b9) {
  d21 <- net_rs(d, "TPrR", "TPrS")
  d31 <- net_rs(d, "SPrR", "SPrS")

  d2_total <- net_rs(d, "ITrR", "ITrS")
  d3_total <- net_rs(d, "SUBrR", "SUBrS")

  d29 <- subtract_required(d2_total, d21)
  d39 <- subtract_required(d3_total, d31)

  d9 <- net_rs(d, "KTrR", "KTrS")
  np <- net_rs(d, "NPrR", "NPrS")
  xb <- col_or_na(d, "XBr")

  nl_calc <- sum_required(xb, d9, np)

  expected <- domestic_b9 %>%
    transmute(
      Year,
      ROW_EXPECTED = -(S11_B9 + S12_B9 + S13_B9 + HH_B9)
    )

  out <- tibble(
    Year = d$Year,

    P1 = structural_zero(d),
    P2 = structural_zero(d),
    D21 = d21,
    D31 = d31,

    # Keep the row definition consistent for ROW as well:
    # P1 + P2 + D21 + D31, with structural-zero P1/P2.
    B1GQ = sum_required(d21, d31),
    B3G = structural_zero(d),

    P3 = structural_zero(d),
    P5 = structural_zero(d),

    P7 = col_or_na(d, "MGSrR"),
    P6 = -col_or_na(d, "XGSrS"),

    D1 = net_rs(d, "WBrR", "WBrS"),
    D29 = d29,
    D5 = net_rs(d, "CTAXrR", "CTAXrS"),
    D39 = d39,

    D61 = net_rs(d, "SCrR", "SCrS"),
    D62 = net_rs(d, "SBrR", "SBrS"),
    D7 = net_rs(d, "TRrR", "TRrS"),

    D41 = net_rs(d, "INTrR", "INTrS"),
    D42 = net_rs(d, "DIVrR", "DIVrS"),
    D43 = net_rs(d, "REIrR", "REIrS"),
    D44 = net_rs(d, "OIrR", "OIrS"),

    # Annual ROW D45 closes in the native received-minus-spent orientation.
    D45 = net_rs(d, "RNTrR", "RNTrS"),

    D8 = net_rs(d, "PENArR", "PENArS"),

    B8G = NA_real_,
    D9 = d9,
    NP = np,
    XB = xb,

    NL_CALC = nl_calc,
    B9 = NA_real_,
    NL_ERROR = NA_real_,

    ROW_EXPECTED = NA_real_,
    ROW_ERROR = NA_real_
  ) %>%
    left_join(
      expected,
      by = "Year",
      suffix = c("", "_joined")
    ) %>%
    mutate(
      ROW_EXPECTED = ROW_EXPECTED_joined,
      ROW_ERROR = ROW_EXPECTED - NL_CALC
    ) %>%
    select(-ROW_EXPECTED_joined)

  out
}

# -------------------------
# Build sector blocks
# -------------------------
s11_wide <- make_S11(data_N)
s12_wide <- make_S12(data_F)
s13_wide <- make_S13(data_G)
hh_wide  <- make_S14_S15(data_H)

domestic_b9 <- s11_wide %>%
  select(Year, S11_B9 = B9) %>%
  full_join(
    s12_wide %>% select(Year, S12_B9 = B9),
    by = "Year"
  ) %>%
  full_join(
    s13_wide %>% select(Year, S13_B9 = B9),
    by = "Year"
  ) %>%
  full_join(
    hh_wide %>% select(Year, HH_B9 = B9),
    by = "Year"
  )

s2_wide <- make_S2(data_R, domestic_b9)

wide_to_long <- function(df, value_col) {
  df %>%
    pivot_longer(
      cols = -Year,
      names_to = "Code",
      values_to = value_col
    )
}

s11_long <- wide_to_long(s11_wide, "S11_NFC")
s12_long <- wide_to_long(s12_wide, "S12_FC")
s13_long <- wide_to_long(s13_wide, "S13_GG")
hh_long  <- wide_to_long(hh_wide,  "S14_S15_HH")
s2_long  <- wide_to_long(s2_wide,  "S2_ROW")

# -------------------------
# Assemble matrix
# -------------------------
all_periods <- unique(c(
  s11_wide$Year,
  s12_wide$Year,
  s13_wide$Year,
  hh_wide$Year,
  s2_wide$Year
))

all_periods <- all_periods[
  !is.na(all_periods) & all_periods != ""
]

all_periods <- all_periods[
  order(year_order(all_periods), all_periods)
]

matrix_base <- row_dictionary %>%
  crossing(Year = all_periods) %>%
  left_join(s11_long, by = c("Year", "Code")) %>%
  left_join(s12_long, by = c("Year", "Code")) %>%
  left_join(s13_long, by = c("Year", "Code")) %>%
  left_join(hh_long,  by = c("Year", "Code")) %>%
  left_join(s2_long,  by = c("Year", "Code"))

matrix_complete <- matrix_base %>%
  rowwise() %>%
  mutate(
    Missing_sector_cells = sum(
      is.na(c_across(all_of(matrix_sector_cols)))
    ),

    Observed_sector_sum = {
      vals <- c_across(all_of(matrix_sector_cols))

      if (all(is.na(vals))) {
        NA_real_
      } else {
        sum(vals, na.rm = TRUE)
      }
    },

    # Never manufacture a production counterpart from incomplete coverage.
    PRODUCTION = case_when(
      Code %in% production_rows &
        Missing_sector_cells == 0 ~ -Observed_sector_sum,

      Code %in% production_rows ~ NA_real_,

      TRUE ~ 0
    ),

    System_row_residual = case_when(
      Zero_sum_test == "Not_applicable" ~ NA_real_,
      Missing_sector_cells > 0 ~ NA_real_,
      is.na(PRODUCTION) ~ NA_real_,
      TRUE ~ PRODUCTION + Observed_sector_sum
    ),

    Closure_status = case_when(
      Zero_sum_test == "Not_applicable" ~
        "N/A: balance, benchmark, or validation row",

      Zero_sum_test == "Production_constructed" &
        Missing_sector_cells > 0 ~
        "UNAVAILABLE: production counterpart not constructed because sector coverage is incomplete",

      Zero_sum_test == "Production_constructed" ~
        "CONSTRUCTED: counterpart from complete coverage; zero residual is mechanical",

      Missing_sector_cells > 0 ~
        "REVIEW: source coverage incomplete; residual not evaluated",

      abs(System_row_residual) <= 0.001 ~
        "OK",

      abs(System_row_residual) <= 1 ~
        "Small rounding difference",

      TRUE ~
        "REVIEW: row does not close"
    )
  ) %>%
  ungroup() %>%
  mutate(
    Period_order = year_order(Year)
  ) %>%
  arrange(
    Period_order,
    Row_order
  ) %>%
  select(
    Year,
    Row_order,
    Code,
    Item,
    Row_type,

    PRODUCTION,

    S11_NFC,
    S12_FC,
    S13_GG,
    S14_S15_HH,
    S2_ROW,

    Observed_sector_sum,
    System_row_residual,
    Missing_sector_cells,
    Closure_status,
    Zero_sum_test
  )

# -------------------------
# Net-lending validation
# -------------------------
nl_validation <- matrix_complete %>%
  filter(
    Code %in% c(
      "NL_CALC", "B9", "NL_ERROR",
      "ROW_EXPECTED", "ROW_ERROR"
    )
  ) %>%
  select(
    Year,
    Code,
    Item,
    S11_NFC,
    S12_FC,
    S13_GG,
    S14_S15_HH,
    S2_ROW
  ) %>%
  arrange(
    year_order(Year),
    match(
      Code,
      c(
        "NL_CALC", "B9", "NL_ERROR",
        "ROW_EXPECTED", "ROW_ERROR"
      )
    )
  )

row_closure <- matrix_complete %>%
  filter(
    Row_type == "Transaction" |
      Code == "B1GQ"
  ) %>%
  select(
    Year,
    Code,
    Item,
    Zero_sum_test,

    PRODUCTION,

    S11_NFC,
    S12_FC,
    S13_GG,
    S14_S15_HH,
    S2_ROW,

    Observed_sector_sum,
    System_row_residual,
    Missing_sector_cells,
    Closure_status
  )

# -------------------------
# TOTAL_ECONOMY (S1) validation
# -------------------------
# S1 is a benchmark only; it is NOT inserted as another sector column.

# Robust alias reader:
#   - accepts one or more aliases
#   - tries unsuffixed direct/balancing item first
#   - otherwise returns resources minus uses
#   - avoids the previous length > 1 failure
te_signed <- function(df, alias) {
  aliases <- unique(
    as.character(
      unlist(alias, use.names = FALSE)
    )
  )

  aliases <- stringr::str_squish(aliases)
  aliases <- aliases[
    !is.na(aliases) & aliases != ""
  ]

  if (length(aliases) == 0) {
    return(rep(NA_real_, nrow(df)))
  }

  direct_hit <- aliases[
    aliases %in% names(df)
  ]

  if (length(direct_hit) > 0) {
    return(
      suppressWarnings(
        as.numeric(df[[direct_hit[1]]])
      )
    )
  }

  for (a in aliases) {
    r_name <- paste0(a, "R")
    s_name <- paste0(a, "S")

    has_r <- r_name %in% names(df)
    has_s <- s_name %in% names(df)

    if (has_r || has_s) {
      r <- if (has_r) {
        suppressWarnings(as.numeric(df[[r_name]]))
      } else {
        rep(0, nrow(df))
      }

      s <- if (has_s) {
        suppressWarnings(as.numeric(df[[s_name]]))
      } else {
        rep(0, nrow(df))
      }

      return(r - s)
    }
  }

  rep(NA_real_, nrow(df))
}

te_code_map <- tribble(
  ~Code,  ~TE_alias, ~Comparison_basis,        ~S1_multiplier,

  "P1",   "OUT",     "Domestic sectors",        1,
  "P3",   "CH",      "Domestic sectors",        1,
  "P5",   "GCF",     "Domestic sectors",        1,

  "D21",  "TP",      "Domestic + PRODUCTION",   1,
  "D31",  "SP",      "Domestic + PRODUCTION",   1,

  "D29",  "D29",     "Domestic sectors",        1,
  "D39",  "D39",     "Domestic sectors",        1,

  "D1",   "WB",      "Domestic sectors",        1,
  "D5",   "CTAX",    "Domestic sectors",        1,
  "D61",  "SC",      "Domestic sectors",        1,
  "D62",  "SB",      "Domestic sectors",        1,
  "D7",   "TR",      "Domestic sectors",        1,

  "D41",  "INT",     "Domestic sectors",        1,
  "D42",  "DIV",     "Domestic sectors",        1,
  "D43",  "REI",     "Domestic sectors",        1,
  "D44",  "OI",      "Domestic sectors",        1,
  "D45",  "RNT",     "Domestic sectors",        1,

  "D8",   "PENA",    "Domestic sectors",        1,

  "B3G",  "MI",      "Domestic sectors",        1,
  "B8G",  "SAV",     "Domestic sectors",        1,
  "D9",   "KT",      "Domestic sectors",        1,

  # S1 NP is direct acquisition-less-disposal; reverse to TFM sign.
  "NP",   "NP",      "Domestic sectors",       -1,

  "B9",   "NETL",    "Domestic sectors",        1
)

# Explicit row-wise map: no pmap_dfr() ambiguity.
te_long <- map_dfr(
  seq_len(nrow(te_code_map)),
  function(i) {
    code_i <- as.character(
      te_code_map$Code[[i]]
    )

    alias_i <- as.character(
      te_code_map$TE_alias[[i]]
    )

    basis_i <- as.character(
      te_code_map$Comparison_basis[[i]]
    )

    multiplier_i <- as.numeric(
      te_code_map$S1_multiplier[[i]]
    )

    tibble(
      Year = data_TE$Year,
      Code = code_i,
      TE_alias = alias_i,
      Comparison_basis = basis_i,
      S1_multiplier = multiplier_i,

      Total_economy_S1 =
        multiplier_i *
        te_signed(data_TE, alias_i)
    )
  }
)

domestic_benchmark <- matrix_complete %>%
  rowwise() %>%
  mutate(
    Domestic_missing_cells = sum(
      is.na(
        c_across(
          c(
            S11_NFC,
            S12_FC,
            S13_GG,
            S14_S15_HH
          )
        )
      )
    ),

    Domestic_sector_sum = {
      vals <- c_across(
        c(
          S11_NFC,
          S12_FC,
          S13_GG,
          S14_S15_HH
        )
      )

      if (any(is.na(vals))) {
        NA_real_
      } else {
        sum(vals)
      }
    }
  ) %>%
  ungroup() %>%
  select(
    Year,
    Code,
    Item,
    PRODUCTION,
    Domestic_sector_sum,
    Domestic_missing_cells
  )

total_economy_validation <- te_long %>%
  left_join(
    domestic_benchmark,
    by = c("Year", "Code")
  ) %>%
  mutate(
    Comparison_value = case_when(
      Comparison_basis == "Domestic sectors" ~
        Domestic_sector_sum,

      Comparison_basis == "Domestic + PRODUCTION" &
        !is.na(Domestic_sector_sum) &
        !is.na(PRODUCTION) ~
        Domestic_sector_sum + PRODUCTION,

      TRUE ~ NA_real_
    ),

    Difference_comparison_minus_S1 =
      Comparison_value - Total_economy_S1,

    Validation_status = case_when(
      is.na(Total_economy_S1) ~
        "S1 benchmark unavailable for this alias/period",

      Domestic_missing_cells > 0 ~
        "REVIEW: domestic sector source coverage incomplete",

      is.na(Comparison_value) ~
        "REVIEW: comparison value unavailable",

      abs(Difference_comparison_minus_S1) <= 0.001 ~
        "OK",

      abs(Difference_comparison_minus_S1) <= 1 ~
        "Small rounding difference",

      abs(Difference_comparison_minus_S1) <= 5 ~
        "Small source/rounding difference",

      TRUE ~
        "REVIEW: comparison differs materially from S1 benchmark"
    )
  ) %>%
  arrange(
    year_order(Year),
    match(Code, te_code_map$Code)
  ) %>%
  select(
    Year,
    Code,
    Item,

    TE_alias,
    Comparison_basis,
    S1_multiplier,

    Domestic_sector_sum,
    PRODUCTION,
    Comparison_value,

    Total_economy_S1,
    Difference_comparison_minus_S1,

    Domestic_missing_cells,
    Validation_status
  )

# -------------------------
# Latest-quarter view and long format
# -------------------------
latest_year <- matrix_complete %>%
  distinct(Year) %>%
  mutate(
    Period_order = year_order(Year)
  ) %>%
  filter(
    !is.na(Period_order)
  ) %>%
  arrange(
    Period_order
  ) %>%
  slice_tail(n = 1) %>%
  pull(Year)

latest_year_matrix <- matrix_complete %>%
  filter(
    Year == latest_year
  ) %>%
  select(
    -Year,
    -Row_order
  )

matrix_long <- matrix_complete %>%
  select(
    Year,
    Row_order,
    Code,
    Item,
    Row_type,

    PRODUCTION,

    S11_NFC,
    S12_FC,
    S13_GG,
    S14_S15_HH,
    S2_ROW
  ) %>%
  pivot_longer(
    cols = c(
      PRODUCTION,
      S11_NFC,
      S12_FC,
      S13_GG,
      S14_S15_HH,
      S2_ROW
    ),
    names_to = "Matrix_column",
    values_to = "Value"
  ) %>%
  arrange(
    year_order(Year),
    Row_order,
    Matrix_column
  )

column_dictionary <- bind_rows(
  tibble(
    Matrix_column = "PRODUCTION",
    Official_sector =
      "Production account / constructed counterpart",
    Meaning = paste0(
      "Constructed for P1, P2, D21, D31, B1GQ, P3, P5, P7 and P6 ",
      "only when all five institutional-sector/ROW cells are available. ",
      "With incomplete coverage, PRODUCTION remains NA. ",
      "It is not an independently observed institutional sector."
    )
  ),

  sector_dictionary %>%
    transmute(
      Matrix_column,
      Official_sector,
      Meaning
    )
)

notes <- tribble(
  ~Item, ~Note,

  "Source workbook",
  input_file,

  "Output workbook",
  output_file,

  "Frequency",
  "Annual",

  "Year format",
  "Year is read from Year/Period labels and ordered chronologically with year_order().",

  "Stage-1 adjustment basis",
  "The final annual Stage-1 pipeline must use one OECD ADJUSTMENT basis consistently and must never sum N and Y variants together.",

  "HH annual saving",
  "No KADJ or balancing residual is inserted. NL_CALC is computed directly from Stage-1 saving and capital-account flows.",

  "Matrix columns",
  "PRODUCTION, S11_NFC, S12_FC, S13_GG, S14_S15_HH, S2_ROW",

  "Naming rule",
  "Official sector code first, readable acronym second; HH explicitly shows that S14 and S15 are combined.",

  "Total Economy",
  "S1 TOTAL_ECONOMY is used only as a benchmark in Total_economy_validation and is not added as another sector column, avoiding double counting.",

  "S1 comparison basis",
  "Most rows compare the domestic institutional-sector sum to S1. D21 and D31 compare domestic sectors plus PRODUCTION. NP reverses the direct S1 sign to match the TFM convention.",

  "Mixed income",
  "B3G mixed income, gross is carried as a household-sector balancing item. Non-HH sector cells are structural zeros; it is not merged with B2G operating surplus or B2A3G operating surplus plus mixed income.",
  "HH annual output",
  "When direct HH P1 is unavailable, annual HH output is derived as P2 + compensation paid + taxes on production and imports paid - subsidies received + gross operating surplus + gross mixed income. Direct P1 is preferred if it becomes available.",

  "Current taxes",
  "The TFM uses D5 current taxes on income, wealth, etc. (Stage-1 prefix CTAX) consistently across annual and quarterly data; D51 taxes on income remains available in Stage 1 as a narrower component.",

  "Non-overlapping tax rows",
  "D2 is not inserted alongside D21. The matrix uses D21 plus D29, where D29 = D2 total - D21.",

  "Non-overlapping subsidy rows",
  "D3 is not inserted alongside D31. The matrix uses D31 plus D39, where D39 = D3 total - D31.",

  "HH block",
  "Uses all available household Stage-1 flows, including wages, consumption, taxes, social flows, transfers, property income, capital transfers, NP, saving and B9.",

  "ROW block",
  "Uses detailed available ROW flows plus trade, current external balance, capital transfers and NP.",

  "ROW D45 orientation",
  "Annual ROW D45 uses the native received-minus-spent orientation. The annual source already closes against the domestic sector entries, so no sign reversal is applied.",

  "Missing direct source series",
  "Kept as NA. They are not silently converted to zero.",

  "Partial R/S pair",
  "If one side of a received/spent pair is present and the other side is absent, the absent side is treated as zero. This is flagged in Input_availability.",

  "Structural zero",
  "Used only where the matrix design deliberately specifies no sector entry, such as corporate final consumption or domestic-sector trade rows allocated through ROW.",

  "Production column",
  "Constructed for P1, P2, D21, D31, B1GQ, P3, P5, P7 and P6 only when all five sector/ROW cells are available. With incomplete coverage, PRODUCTION and the closure residual remain NA.",

  "Observed row closure",
  "For non-production transaction rows, System_row_residual is evaluated only when all five sector/ROW cells are available. Missing coverage is flagged before interpreting a residual.",

  "Domestic net lending",
  "NL_CALC = B8G + D9 + signed P5 + NP. NL_ERROR = B9 - NL_CALC. No KADJ or forced balancing term is used.",

  "ROW net lending",
  "NL_CALC = XB + D9 + NP. ROW_EXPECTED = negative sum of domestic official B9. ROW_ERROR compares the two.",

  "Scope limitation",
  "This is the annual non-financial TFM construction relative to the Stage-1 transaction universe and available OECD sector data. Financial accounts, stocks, revaluations and stock-flow links are separate later stages."
)

# -------------------------
# Export workbook
# -------------------------
writexl::write_xlsx(
  list(
    TFM_latest_year =
      latest_year_matrix,

    TFM_all_years =
      matrix_complete %>%
      select(-Row_order),

    TFM_long =
      matrix_long %>%
      select(-Row_order),

    Row_closure =
      row_closure,

    NL_validation =
      nl_validation,

    Total_economy_validation =
      total_economy_validation,

    Input_availability =
      input_availability,

    Column_dictionary =
      column_dictionary,

    Row_dictionary =
      row_dictionary,

    Notes =
      notes
  ),
  path = output_file
)

message(
  "Complete annual non-financial TFM workbook written: ",
  output_file
)

message(
  "Latest quarter included in TFM_latest_year: ",
  latest_year
)

message(
  "Review Input_availability, Row_closure, NL_validation and ",
  "Total_economy_validation before treating any residual as an ",
  "accounting inconsistency."
)
