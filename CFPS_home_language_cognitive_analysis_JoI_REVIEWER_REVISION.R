# =============================================================================
# CFPS HOME LANGUAGE AND COGNITIVE TEST PERFORMANCE
# Journal of Intelligence - REVIEWER-REVISION ANALYSIS PIPELINE
#
# Data used:
#   2010 adult + 2010 family economy + 2012 adult + 2014 adult + 2018 person
#
# Core additions relative to the previous script:
#   1) Cross-domain equivalence / TOST and one-sided non-inferiority tests
#   2) Common-support diagnostics + restricted-sample sensitivity analyses
#   3) Nonlinear baseline sensitivity using natural splines
#   4) 2010-2014-2018 balanced-panel trajectory models
#   5) Correct domain figure: contrasts versus Mandarin, not adjusted means
#   6) ALL tabular/data-like results are written into ONE Excel workbook
#   7) Full CFPS survey design: strata=subpopulation, PSU=psu, longitudinal weights
#   8) Province x language and ethnicity x language overlap diagnostics
#   9) Binary propensity/overlap-weight diagnostics for minority-language vs Mandarin
#  10) Weighted score distributions, floor/ceiling diagnostics, floor-exclusion sensitivity
#  11) Detailed language-specific sample-flow/retention audit
#  12) Expanded Table 1 and pooled-category ethnicity/province composition
#  13) Representativeness audit for the 2012 Mandarin-proficiency subsample
#  14) Multiple-imputation sensitivity for missing covariates
#
# 2022 selective retest has been intentionally removed from this version.
# Respondent-level data are NEVER exported.
# =============================================================================

# 0. USER SETTINGS ------------------------------------------------------------

message("请选择任意一个 CFPS analysis-ready Excel 文件（程序会自动识别同文件夹内其他文件）")
chosen_file <- file.choose()

if (tolower(tools::file_ext(chosen_file)) != "xlsx") {
  stop("选择的文件不是 .xlsx，请重新运行并选择任意一个 analysis-ready Excel 文件。")
}

DATA_DIR <- dirname(normalizePath(chosen_file, winslash = "/", mustWork = TRUE))
OUT_DIR  <- file.path(DATA_DIR, "CFPS_home_language_results_JoI_REVIEWER_REVISION")
FIG_DIR  <- file.path(OUT_DIR, "figures")

dir.create(OUT_DIR, showWarnings = FALSE, recursive = TRUE)
dir.create(FIG_DIR, showWarnings = FALSE, recursive = TRUE)

# Optional module switches
RUN_2012_MECHANISM      <- TRUE
RUN_BALANCED_TRAJECTORY <- TRUE
RUN_MATH_TRAJECTORY     <- TRUE
RUN_MI_SENSITIVITY      <- TRUE

# Equivalence / SESOI settings
# Primary SESOI: +/- 0.10 SD
# Sensitivity SESOI: +/- 0.15 SD
SESOI_BOUNDS <- c(0.10, 0.15)
ALPHA <- 0.05

# Common-support trimming rules.
# Primary: intersection of group-specific 1st-99th weighted percentiles.
# Strict sensitivity: intersection of group-specific 5th-95th weighted percentiles.
SUPPORT_PRIMARY <- c(0.01, 0.99)
SUPPORT_STRICT  <- c(0.05, 0.95)

# Reviewer-requested additional sensitivity settings
FLOOR_CUTOFFS <- c(0, 3, 5)     # exact floor; bottom ~10%; bottom ~15% of 0-34 scale
GEO_MIN_COUNTS <- c(1, 5, 10)   # minimum unweighted n in BOTH Mandarin and minority groups
PS_SUPPORT_PROBS <- c(0.01, 0.99)

set.seed(20260824)
options(survey.lonely.psu = "adjust")

cat("\n数据文件夹：", DATA_DIR, "\n", sep = "")
cat("结果文件夹：", OUT_DIR, "\n\n", sep = "")


# 1. PACKAGES -----------------------------------------------------------------

required_packages <- c(
  "readxl", "dplyr", "tidyr", "purrr", "stringr", "readr",
  "ggplot2", "survey", "broom", "emmeans", "lavaan", "writexl",
  "scales", "rlang"
)

if (RUN_MI_SENSITIVITY) {
  required_packages <- c(required_packages, "mice", "mitools")
}

missing_packages <- setdiff(required_packages, rownames(installed.packages()))
if (length(missing_packages) > 0) {
  message("需要安装以下R包：", paste(missing_packages, collapse = ", "))
  install.packages(missing_packages, dependencies = TRUE)
}

suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(tidyr)
  library(purrr)
  library(stringr)
  library(readr)
  library(ggplot2)
  library(survey)
  library(broom)
  library(emmeans)
  library(lavaan)
  library(writexl)
  library(scales)
  library(rlang)
})


# 2. GENERAL HELPERS -----------------------------------------------------------

locate_one <- function(pattern, label, required = TRUE) {
  hits <- list.files(
    DATA_DIR,
    pattern = pattern,
    full.names = TRUE,
    recursive = FALSE,
    ignore.case = TRUE
  )

  if (length(hits) == 0) {
    if (required) stop("Missing ", label, ". Expected: ", pattern)
    message("Optional file not found: ", label)
    return(NA_character_)
  }

  if (length(hits) > 1) {
    stop(
      "More than one file matches ", label, ":\n",
      paste(hits, collapse = "\n"),
      "\nPlease keep only one intended version in the selected folder."
    )
  }

  normalizePath(hits, winslash = "/", mustWork = TRUE)
}

read_selected <- function(path, wanted, wave_label) {
  dat <- readxl::read_excel(
    path,
    sheet = "数据",
    guess_max = 100000,
    na = c("", "NA", "NaN")
  )

  names(dat) <- tolower(names(dat))
  lookup <- setNames(names(dat), tolower(names(dat)))
  actual <- unname(lookup[tolower(wanted)])

  if (anyNA(actual)) {
    missing_vars <- wanted[is.na(actual)]
    stop(
      "Variables missing from ", wave_label, ": ",
      paste(missing_vars, collapse = ", ")
    )
  }

  dat %>%
    select(all_of(actual)) %>%
    rename_with(tolower)
}

make_dictionary <- function(path, wave_label) {
  x <- readxl::read_excel(path, sheet = "变量字典")
  if (ncol(x) < 2) stop("Variable dictionary is incomplete: ", path)

  tibble(
    wave = wave_label,
    variable = as.character(x[[1]]),
    label = as.character(x[[2]])
  )
}

num_clean <- function(x) {
  suppressWarnings(as.numeric(as.character(x)))
}

valid_range <- function(x, lower, upper) {
  y <- num_clean(x)
  ifelse(!is.na(y) & y >= lower & y <= upper, y, NA_real_)
}

valid_set <- function(x, allowed) {
  y <- num_clean(x)
  ifelse(y %in% allowed, y, NA_real_)
}

# Preserve IDs as character and avoid unnecessary floating-point conversion.
clean_id <- function(x) {
  y <- trimws(as.character(x))
  y[y %in% c("", "NA", "NaN", ".")] <- NA_character_
  y <- sub("\\.0+$", "", y)
  bad <- !is.na(y) & (!grepl("^[0-9]+$", y) | y == "0")
  y[bad] <- NA_character_
  y
}

assert_unique_key <- function(dat, key = "pid", wave_label = "data") {
  dup <- dat %>%
    filter(!is.na(.data[[key]])) %>%
    count(.data[[key]], name = "n") %>%
    filter(n > 1) %>%
    arrange(desc(n))

  if (nrow(dup) > 0) {
    stop(
      wave_label, " does not have a unique ", key,
      ". Largest duplicate count = ", max(dup$n),
      ". Please inspect duplicates before analysis."
    )
  }
  invisible(TRUE)
}

weighted_z <- function(x, w) {
  ok <- is.finite(x) & is.finite(w) & w > 0
  if (sum(ok) < 2) return(rep(NA_real_, length(x)))
  mu <- weighted.mean(x[ok], w[ok])
  sd_w <- sqrt(weighted.mean((x[ok] - mu)^2, w[ok]))
  if (!is.finite(sd_w) || sd_w <= 0) return(rep(NA_real_, length(x)))
  (x - mu) / sd_w
}

weighted_quantile <- function(x, w, probs) {
  ok <- is.finite(x) & is.finite(w) & w > 0
  x <- x[ok]
  w <- w[ok]
  if (length(x) == 0) return(rep(NA_real_, length(probs)))

  ord <- order(x)
  x <- x[ord]
  w <- w[ord]
  cw <- cumsum(w) / sum(w)

  vapply(probs, function(p) {
    idx <- which(cw >= p)[1]
    if (is.na(idx)) tail(x, 1) else x[idx]
  }, numeric(1))
}

make_svy_design <- function(dat, weight_var = "weight18_panel") {
  if (!all(c("psu_f", "strata_f", weight_var) %in% names(dat))) {
    stop("Survey design variables are missing: psu_f, strata_f, or ", weight_var)
  }
  svydesign(
    ids = ~psu_f,
    strata = ~strata_f,
    weights = as.formula(paste0("~", weight_var)),
    data = dat,
    nest = TRUE
  )
}

extract_language_terms <- function(fit, model_label) {
  broom::tidy(fit, conf.int = TRUE) %>%
    filter(grepl("^home_language10", term)) %>%
    mutate(model = model_label, .before = 1)
}

safe_regterm_table <- function(test_obj, label) {
  fval <- if (!is.null(test_obj$Ftest)) as.numeric(test_obj$Ftest)[1] else NA_real_
  df_num <- if (!is.null(test_obj$df)) as.numeric(test_obj$df)[1] else NA_real_
  df_den <- if (!is.null(test_obj$ddf)) as.numeric(test_obj$ddf)[1] else NA_real_

  pval <- NA_real_
  for (nm in c("p", "pvalue", "p.value")) {
    if (!is.null(test_obj[[nm]])) {
      pval <- as.numeric(test_obj[[nm]])[1]
      break
    }
  }

  tibble(
    test = label,
    F = fval,
    df_num = df_num,
    df_den = df_den,
    p_value = pval,
    printed_output = paste(capture.output(test_obj), collapse = " | ")
  )
}

safe_confint_df <- function(emm_obj) {
  as.data.frame(confint(emm_obj)) %>% as_tibble()
}

# Results are accumulated here and written once at the end.
RESULTS <- list()
STATUS  <- tibble(module = character(), status = character(), detail = character())

add_status <- function(module, status, detail = "") {
  STATUS <<- bind_rows(
    STATUS,
    tibble(module = module, status = status, detail = detail)
  )
}


# 3. LOCATE DATA FILES --------------------------------------------------------

path_2010_person <- locate_one(
  "^CFPS_2010_adult_analysis_ready(\\([0-9]+\\))?\\.xlsx$",
  "CFPS 2010 adult Excel file"
)
path_2010_family <- locate_one(
  "^CFPS_2010_family_economy_analysis_ready(\\([0-9]+\\))?\\.xlsx$",
  "CFPS 2010 family economy Excel file"
)
path_2012_person <- locate_one(
  "^CFPS_2012_adult_analysis_ready(\\([0-9]+\\))?\\.xlsx$",
  "CFPS 2012 adult Excel file",
  required = FALSE
)
path_2014_person <- locate_one(
  "^CFPS_2014_adult_analysis_ready(\\([0-9]+\\))?\\.xlsx$",
  "CFPS 2014 adult Excel file",
  required = FALSE
)
path_2018_person <- locate_one(
  "^CFPS_2018_person_analysis_ready(\\([0-9]+\\))?\\.xlsx$",
  "CFPS 2018 person Excel file"
)

required_file_audit <- tibble(
  dataset = c("2010 person", "2010 family", "2012 person", "2014 person", "2018 person"),
  path = c(path_2010_person, path_2010_family, path_2012_person, path_2014_person, path_2018_person),
  found = !is.na(c(path_2010_person, path_2010_family, path_2012_person, path_2014_person, path_2018_person))
)
RESULTS[["00_file_audit"]] <- required_file_audit


# 4. READ VARIABLES -----------------------------------------------------------

vars_2010_person <- c(
  "pid", "fid", "provcd", "psu", "rswt_nat", "rswt_res",
  "subsample", "subpopulation", "urban", "gender", "qa1age",
  "qa5code", "qd2", "qd101", "qd102", "qz205", "wordtest",
  "mathtest", "cfps2010eduy_best", "feduc", "meduc", "income"
)

vars_2010_family <- c(
  "fid", "familysize", "faminc", "indinc", "faminc_net", "indinc_net"
)

vars_2018_person <- c(
  "pid", "psu", "provcd18", "urban18", "age", "gender", "qz103",
  "wordtest18", "mathtest18", "wordtest18_sc2", "mathtest18_sc2",
  "cfps2018eduy", "cfps2018eduy_im", "rswt_natpn1018n", "rswt_natcs18n"
)

raw10 <- read_selected(path_2010_person, vars_2010_person, "CFPS 2010 person")
fam10 <- read_selected(path_2010_family, vars_2010_family, "CFPS 2010 family")
raw18 <- read_selected(path_2018_person, vars_2018_person, "CFPS 2018 person")


# 5. VARIABLE DECODING --------------------------------------------------------

person10 <- raw10 %>%
  transmute(
    pid = clean_id(pid),
    fid = clean_id(fid),
    province10 = valid_range(provcd, 1, 99),
    psu10 = clean_id(psu),
    weight10_national = if_else(num_clean(rswt_nat) > 0, num_clean(rswt_nat), NA_real_),
    weight10_resample = if_else(num_clean(rswt_res) > 0, num_clean(rswt_res), NA_real_),
    subsample10 = valid_set(subsample, c(0, 1)),
    # CFPS official guidance identifies subpopulation as the six sampling strata.
    subpopulation10 = valid_range(subpopulation, 1, 6),
    urban10 = factor(
      valid_set(urban, c(0, 1)),
      levels = c(0, 1),
      labels = c("Rural", "Urban")
    ),
    sex = factor(
      valid_set(gender, c(0, 1)),
      levels = c(0, 1),
      labels = c("Female", "Male")
    ),
    age10 = valid_range(qa1age, 16, 110),
    birth_year = 2010 - age10,
    ethnicity_code10 = valid_range(qa5code, 1, 56),
    ethnicity = factor(
      case_when(
        ethnicity_code10 == 1 ~ "Han",
        ethnicity_code10 > 1 ~ "Ethnic minority",
        TRUE ~ NA_character_
      ),
      levels = c("Han", "Ethnic minority")
    ),
    home_language10 = factor(
      case_when(
        valid_set(qd2, c(1, 3, 5, 77)) == 1 ~ "Mandarin",
        valid_set(qd2, c(1, 3, 5, 77)) == 3 ~ "Sinitic dialect",
        valid_set(qd2, c(1, 3, 5, 77)) == 5 ~ "Minority language",
        valid_set(qd2, c(1, 3, 5, 77)) == 77 ~ "Other",
        TRUE ~ NA_character_
      ),
      levels = c("Mandarin", "Sinitic dialect", "Minority language", "Other")
    ),
    mandarin_importance10 = valid_range(qd101, 1, 5),
    dialect_importance10 = valid_range(qd102, 1, 5),
    mandarin_fluency10 = valid_range(qz205, 1, 7),
    word10 = valid_range(wordtest, 0, 34),
    math10 = valid_range(mathtest, 0, 24),
    education_years10 = valid_range(cfps2010eduy_best, 0, 22),
    father_education_code = valid_set(feduc, 1:8),
    mother_education_code = valid_set(meduc, 1:8),
    personal_income10 = if_else(num_clean(income) >= 0, num_clean(income), NA_real_)
  )

family10 <- fam10 %>%
  transmute(
    fid = clean_id(fid),
    family_size10 = valid_range(familysize, 1, 30),
    family_income10 = if_else(num_clean(faminc) >= 0, num_clean(faminc), NA_real_),
    family_income_pc10 = if_else(num_clean(indinc) >= 0, num_clean(indinc), NA_real_),
    family_net_income10 = if_else(num_clean(faminc_net) >= 0, num_clean(faminc_net), NA_real_),
    family_net_income_pc10 = if_else(num_clean(indinc_net) >= 0, num_clean(indinc_net), NA_real_)
  )

person18 <- raw18 %>%
  transmute(
    pid = clean_id(pid),
    linked18 = 1L,
    psu18 = clean_id(psu),
    province18 = valid_range(provcd18, 1, 99),
    urban18 = factor(
      valid_set(urban18, c(0, 1)),
      levels = c(0, 1),
      labels = c("Rural", "Urban")
    ),
    age18 = valid_range(age, 10, 110),
    sex18 = factor(
      valid_set(gender, c(0, 1)),
      levels = c(0, 1),
      labels = c("Female", "Male")
    ),
    interview_language18 = factor(
      valid_set(qz103, c(1, 5)),
      levels = c(1, 5),
      labels = c("Mandarin", "Dialect")
    ),
    word18 = valid_range(wordtest18, 0, 34),
    math18 = valid_range(mathtest18, 0, 24),
    word18_wave_algorithm = valid_range(wordtest18_sc2, 0, 34),
    math18_wave_algorithm = valid_range(mathtest18_sc2, 0, 24),
    education_years18 = valid_range(cfps2018eduy, 0, 22),
    education_years18_imputed = valid_range(cfps2018eduy_im, 0, 22),
    weight18_panel = if_else(
      num_clean(rswt_natpn1018n) > 0,
      num_clean(rswt_natpn1018n),
      NA_real_
    ),
    weight18_cross_section = if_else(
      num_clean(rswt_natcs18n) > 0,
      num_clean(rswt_natcs18n),
      NA_real_
    )
  )

assert_unique_key(person10, "pid", "2010_person")
assert_unique_key(family10, "fid", "2010_family")
assert_unique_key(person18, "pid", "2018_person")


# 6. BUILD 2010-2018 PANEL ----------------------------------------------------

panel <- person10 %>%
  left_join(family10, by = "fid") %>%
  left_join(person18, by = "pid") %>%
  mutate(
    linked18 = if_else(is.na(linked18), 0L, linked18),
    retained18 = as.integer(!is.na(word18)),
    strata = subpopulation10,
    psu = coalesce(psu10, psu18),
    province = coalesce(province10, province18),
    age10_c = age10 - 45,
    birth_year_c = birth_year - 1965,
    education10_c = education_years10 - 9,
    parent_education_max = pmax(
      father_education_code,
      mother_education_code,
      na.rm = TRUE
    ),
    parent_education_max = if_else(
      is.infinite(parent_education_max),
      NA_real_,
      parent_education_max
    ),
    parent_education_missing = as.integer(is.na(parent_education_max)),
    parent_education_max_i = if_else(
      is.na(parent_education_max),
      median(parent_education_max, na.rm = TRUE),
      parent_education_max
    ),
    income_pc_source = coalesce(family_income_pc10, family_net_income_pc10),
    log_income_pc10 = if_else(
      !is.na(income_pc_source) & income_pc_source >= 0,
      log1p(income_pc_source),
      NA_real_
    ),
    income_missing = as.integer(is.na(log_income_pc10)),
    log_income_pc10_i = if_else(
      is.na(log_income_pc10),
      median(log_income_pc10, na.rm = TRUE),
      log_income_pc10
    )
  )

primary_language_levels <- c("Mandarin", "Sinitic dialect", "Minority language")

flow <- tibble(
  step = c(
    "2010 adult respondents",
    "Valid focal 2010 home-language category",
    "+ valid 2010 word score",
    "+ linked to 2018 person file",
    "+ valid 2018 word score",
    "+ positive 2018 longitudinal weight",
    "+ complete design/core covariates"
  ),
  n = c(
    nrow(panel),
    sum(panel$home_language10 %in% primary_language_levels, na.rm = TRUE),
    sum(panel$home_language10 %in% primary_language_levels & !is.na(panel$word10), na.rm = TRUE),
    sum(panel$home_language10 %in% primary_language_levels & !is.na(panel$word10) & panel$linked18 == 1L, na.rm = TRUE),
    sum(panel$home_language10 %in% primary_language_levels & !is.na(panel$word10) & panel$linked18 == 1L & !is.na(panel$word18), na.rm = TRUE),
    sum(panel$home_language10 %in% primary_language_levels & !is.na(panel$word10) & panel$linked18 == 1L & !is.na(panel$word18) & !is.na(panel$weight18_panel) & panel$weight18_panel > 0, na.rm = TRUE),
    sum(
      panel$home_language10 %in% primary_language_levels &
        !is.na(panel$word10) & panel$linked18 == 1L & !is.na(panel$word18) &
        !is.na(panel$weight18_panel) & panel$weight18_panel > 0 &
        !is.na(panel$psu) & !is.na(panel$strata) & !is.na(panel$age10) & !is.na(panel$sex) &
        !is.na(panel$urban10) & !is.na(panel$education_years10) &
        !is.na(panel$ethnicity) & !is.na(panel$ethnicity_code10) & !is.na(panel$province),
      na.rm = TRUE
    )
  )
) %>%
  mutate(
    pct_of_previous = 100 * n / lag(n),
    pct_of_previous = if_else(row_number() == 1, 100, pct_of_previous)
  )
RESULTS[["01_sample_flow"]] <- flow

analysis_primary <- panel %>%
  filter(
    home_language10 %in% primary_language_levels,
    !is.na(word10), !is.na(word18),
    !is.na(weight18_panel), weight18_panel > 0,
    !is.na(psu), !is.na(strata), !is.na(age10), !is.na(sex),
    !is.na(urban10), !is.na(education_years10),
    !is.na(ethnicity), !is.na(ethnicity_code10), !is.na(province)
  ) %>%
  mutate(
    home_language10 = factor(home_language10, levels = primary_language_levels),
    province_f = factor(province),
    ethnicity_code_f = factor(ethnicity_code10),
    psu_f = factor(psu),
    strata_f = factor(strata),
    male_ind = as.integer(sex == "Male"),
    urban_ind = as.integer(urban10 == "Urban"),
    ethnic_minority_ind = as.integer(ethnicity == "Ethnic minority"),
    word10_z = weighted_z(word10, weight18_panel),
    word18_z = weighted_z(word18, weight18_panel)
  )

RESULTS[["02_group_counts"]] <- analysis_primary %>%
  count(home_language10, name = "unweighted_n")


# 7. VARIABLE DICTIONARY ------------------------------------------------------

dictionary_files <- list(
  `2010 person` = path_2010_person,
  `2010 family economy` = path_2010_family,
  `2018 person` = path_2018_person
)
if (!is.na(path_2012_person)) dictionary_files[["2012 person"]] <- path_2012_person
if (!is.na(path_2014_person)) dictionary_files[["2014 person"]] <- path_2014_person

RESULTS[["03_variable_dictionary"]] <- imap_dfr(
  dictionary_files,
  ~ make_dictionary(.x, .y)
)


# 8. DESCRIPTIVE STATISTICS ---------------------------------------------------

design_primary <- make_svy_design(analysis_primary)

weighted_means <- bind_rows(
  as_tibble(svyby(~word10, ~home_language10, design_primary, svymean, na.rm = TRUE)) %>%
    mutate(outcome = "Word score, 2010"),
  as_tibble(svyby(~word18, ~home_language10, design_primary, svymean, na.rm = TRUE)) %>%
    mutate(outcome = "Word score, 2018"),
  as_tibble(svyby(~age10, ~home_language10, design_primary, svymean, na.rm = TRUE)) %>%
    mutate(outcome = "Age, 2010"),
  as_tibble(svyby(~education_years10, ~home_language10, design_primary, svymean, na.rm = TRUE)) %>%
    mutate(outcome = "Education years, 2010")
) %>%
  relocate(outcome)

RESULTS[["04_weighted_descriptives"]] <- weighted_means


# 9. PRIMARY BASELINE-ADJUSTED MODELS ----------------------------------------

formula_minimal <- word18_z ~
  home_language10 + word10_z + age10_c + I(age10_c^2) + sex

formula_social <- word18_z ~
  home_language10 + word10_z + age10_c + I(age10_c^2) + sex +
  urban10 + ethnicity + education10_c +
  parent_education_max_i + parent_education_missing +
  log_income_pc10_i + income_missing

formula_full <- update(formula_social, . ~ . + province_f)

fit_minimal <- svyglm(formula_minimal, design = design_primary)
fit_social  <- svyglm(formula_social,  design = design_primary)
fit_full    <- svyglm(formula_full,    design = design_primary)

main_models <- list(
  Model_1_baseline = fit_minimal,
  Model_2_social = fit_social,
  Model_3_province = fit_full
)

RESULTS[["05_primary_models"]] <- imap_dfr(
  main_models,
  ~ broom::tidy(.x, conf.int = TRUE) %>% mutate(model = .y, .before = 1)
)

RESULTS[["06_primary_language_terms"]] <- imap_dfr(
  main_models,
  ~ extract_language_terms(.x, .y)
)

# Quantify attenuation in absolute magnitude from Model 1 to Model 2/3.
primary_lang_terms <- RESULTS[["06_primary_language_terms"]] %>%
  select(model, term, estimate) %>%
  pivot_wider(names_from = model, values_from = estimate) %>%
  mutate(
    abs_attenuation_M1_to_M2 = abs(Model_1_baseline) - abs(Model_2_social),
    pct_attenuation_M1_to_M2 = 100 * (abs(Model_1_baseline) - abs(Model_2_social)) / abs(Model_1_baseline),
    abs_change_M2_to_M3 = abs(Model_3_province) - abs(Model_2_social)
  )
RESULTS[["07_model_attenuation"]] <- primary_lang_terms


# 10. CROSS-DOMAIN MODEL ------------------------------------------------------

analysis_domain <- panel %>%
  filter(
    home_language10 %in% primary_language_levels,
    !is.na(word10), !is.na(word18), !is.na(math10), !is.na(math18),
    !is.na(weight18_panel), weight18_panel > 0,
    !is.na(psu), !is.na(strata), !is.na(age10), !is.na(sex),
    !is.na(urban10), !is.na(education_years10),
    !is.na(ethnicity), !is.na(province)
  ) %>%
  mutate(
    home_language10 = factor(home_language10, levels = primary_language_levels),
    province_f = factor(province),
    psu_f = factor(psu),
    strata_f = factor(strata),
    pid_cluster = factor(pid),
    word10_z = weighted_z(word10, weight18_panel),
    word18_z = weighted_z(word18, weight18_panel),
    math10_z = weighted_z(math10, weight18_panel),
    math18_z = weighted_z(math18, weight18_panel)
  )

RESULTS[["08_domain_sample_counts"]] <- analysis_domain %>%
  count(home_language10, name = "unweighted_n")

domain_long <- bind_rows(
  analysis_domain %>%
    transmute(
      pid, pid_cluster, psu_f, strata_f, weight18_panel, home_language10,
      age10_c, sex, urban10, ethnicity, education10_c,
      parent_education_max_i, parent_education_missing,
      log_income_pc10_i, income_missing, province_f,
      domain = "Verbal",
      baseline_z = word10_z,
      outcome_z = word18_z
    ),
  analysis_domain %>%
    transmute(
      pid, pid_cluster, psu_f, strata_f, weight18_panel, home_language10,
      age10_c, sex, urban10, ethnicity, education10_c,
      parent_education_max_i, parent_education_missing,
      log_income_pc10_i, income_missing, province_f,
      domain = "Math",
      baseline_z = math10_z,
      outcome_z = math18_z
    )
) %>%
  mutate(domain = factor(domain, levels = c("Math", "Verbal")))

design_domain <- svydesign(
  ids = ~psu_f + pid_cluster,
  strata = ~strata_f,
  weights = ~weight18_panel,
  data = domain_long,
  nest = TRUE
)

fit_domain <- svyglm(
  outcome_z ~
    domain * home_language10 + domain * baseline_z +
    domain * (
      age10_c + I(age10_c^2) + sex + urban10 + ethnicity + education10_c +
      parent_education_max_i + parent_education_missing +
      log_income_pc10_i + income_missing + province_f
    ),
  design = design_domain
)

RESULTS[["09_domain_model"]] <- broom::tidy(fit_domain, conf.int = TRUE)

domain_wald <- regTermTest(
  fit_domain,
  ~domain:home_language10,
  method = "Wald"
)
RESULTS[["10_domain_wald"]] <- safe_regterm_table(
  domain_wald,
  "Joint domain x home-language interaction"
)

# IMPORTANT: These are true contrasts versus Mandarin within each domain.
# The previous script plotted adjusted means; this version corrects that.
emm_domain <- emmeans(
  fit_domain,
  ~ home_language10 | domain,
  weights = "proportional"
)

domain_contrasts <- contrast(
  emm_domain,
  method = "trt.vs.ctrl",
  ref = 1,
  adjust = "none"
)

domain_contrasts_df <- as.data.frame(
  summary(domain_contrasts, infer = c(TRUE, TRUE), level = 0.95)
) %>%
  as_tibble() %>%
  rename(language_contrast = contrast) %>%
  mutate(
    language_group = case_when(
      grepl("Sinitic dialect", language_contrast, fixed = TRUE) ~ "Sinitic dialect",
      grepl("Minority language", language_contrast, fixed = TRUE) ~ "Minority language",
      TRUE ~ language_contrast
    )
  )
RESULTS[["11_domain_contrasts"]] <- domain_contrasts_df

# Main cross-domain contrast figure (true differences relative to Mandarin)
p_domain <- ggplot(
  domain_contrasts_df,
  aes(
    x = language_group,
    y = estimate,
    ymin = lower.CL,
    ymax = upper.CL,
    color = domain
  )
) +
  geom_hline(yintercept = 0, linetype = 2, color = "grey65") +
  geom_pointrange(
    position = position_dodge(width = 0.38),
    linewidth = 0.55
  ) +
  scale_color_manual(values = c("Math" = "#E07A1F", "Verbal" = "#2457C5")) +
  labs(
    x = NULL,
    y = "Difference from Mandarin (weighted SD units)",
    color = "Domain",
    title = "Home-language contrasts by test domain",
    subtitle = "Baseline-adjusted contrasts versus Mandarin; 95% confidence intervals"
  ) +
  theme_minimal(base_size = 13) +
  theme(
    panel.grid.minor = element_blank(),
    legend.position = "top"
  )

print(p_domain)
ggsave(
  file.path(FIG_DIR, "Figure_2_domain_contrasts_vs_Mandarin.png"),
  p_domain,
  width = 8.4,
  height = 5.2,
  dpi = 450
)


# 11. EQUIVALENCE / TOST FOR CROSS-DOMAIN DIFFERENCES ------------------------

# With Math as the reference domain, the domainVerbal:home_language interaction
# is exactly: (Verbal language contrast) - (Math language contrast).
coefs_domain <- coef(fit_domain)
vc_domain <- vcov(fit_domain)
coef_names <- names(coefs_domain)
model_df <- suppressWarnings(as.numeric(df.residual(fit_domain)))
if (!is.finite(model_df) || model_df <= 0) model_df <- survey::degf(design_domain)

find_interaction_term <- function(group_label) {
  candidates <- c(
    paste0("domainVerbal:home_language10", group_label),
    paste0("home_language10", group_label, ":domainVerbal")
  )
  hit <- intersect(candidates, coef_names)
  if (length(hit) != 1) {
    stop(
      "Could not uniquely identify domain interaction term for: ", group_label,
      "\nCandidate terms: ", paste(candidates, collapse = " | ")
    )
  }
  hit
}

cross_domain_differences <- map_dfr(
  c("Sinitic dialect", "Minority language"),
  function(group_label) {
    term <- find_interaction_term(group_label)
    est <- unname(coefs_domain[[term]])
    se  <- sqrt(unname(vc_domain[term, term]))

    tibble(
      language_group = group_label,
      contrast_definition = "Verbal minus Math language contrast",
      estimate = est,
      SE = se,
      df = model_df,
      CI90_low = est - qt(0.95, model_df) * se,
      CI90_high = est + qt(0.95, model_df) * se,
      CI95_low = est - qt(0.975, model_df) * se,
      CI95_high = est + qt(0.975, model_df) * se,
      two_sided_p = 2 * pt(-abs(est / se), df = model_df)
    )
  }
)
RESULTS[["12_domain_differences"]] <- cross_domain_differences

# TOST + theory-focused one-sided non-inferiority.
# Negative values mean an additional verbal disadvantage relative to math.
# The one-sided test asks whether we can rule out an extra verbal penalty
# larger than the prespecified SESOI (e.g., worse than -0.10 SD).
tost_results <- tidyr::expand_grid(
  language_group = unique(cross_domain_differences$language_group),
  SESOI = SESOI_BOUNDS
) %>%
  left_join(
    cross_domain_differences %>% select(language_group, estimate, SE, df),
    by = "language_group"
  ) %>%
  rowwise() %>%
  mutate(
    t_lower = (estimate + SESOI) / SE,
    p_lower = 1 - pt(t_lower, df),
    t_upper = (estimate - SESOI) / SE,
    p_upper = pt(t_upper, df),
    TOST_p = max(p_lower, p_upper),
    equivalent_two_sided = p_lower < ALPHA & p_upper < ALPHA,
    noninferior_extra_verbal = p_lower < ALPHA,
    CI90_low = estimate - qt(0.95, df) * SE,
    CI90_high = estimate + qt(0.95, df) * SE,
    equivalence_interval = paste0("[-", sprintf("%.2f", SESOI), ", +", sprintf("%.2f", SESOI), "]")
  ) %>%
  ungroup()
RESULTS[["13_TOST_equivalence"]] <- tost_results


# 12. COMMON-SUPPORT DIAGNOSTICS ---------------------------------------------

calc_support <- function(dat, probs, label) {
  qtab <- dat %>%
    group_by(home_language10) %>%
    summarise(
      word_low = weighted_quantile(word10_z, weight18_panel, probs[1]),
      word_high = weighted_quantile(word10_z, weight18_panel, probs[2]),
      edu_low = weighted_quantile(education_years10, weight18_panel, probs[1]),
      edu_high = weighted_quantile(education_years10, weight18_panel, probs[2]),
      .groups = "drop"
    )

  bounds <- tibble(
    support_rule = label,
    lower_prob = probs[1],
    upper_prob = probs[2],
    word_lower = max(qtab$word_low, na.rm = TRUE),
    word_upper = min(qtab$word_high, na.rm = TRUE),
    edu_lower = max(qtab$edu_low, na.rm = TRUE),
    edu_upper = min(qtab$edu_high, na.rm = TRUE)
  )

  qtab <- qtab %>% mutate(support_rule = label, .before = 1)
  list(quantiles = qtab, bounds = bounds)
}

support_primary <- calc_support(analysis_primary, SUPPORT_PRIMARY, "1st-99th percentile intersection")
support_strict  <- calc_support(analysis_primary, SUPPORT_STRICT,  "5th-95th percentile intersection")

RESULTS[["14_support_quantiles"]] <- bind_rows(
  support_primary$quantiles,
  support_strict$quantiles
)
RESULTS[["15_support_bounds"]] <- bind_rows(
  support_primary$bounds,
  support_strict$bounds
)

apply_support <- function(dat, bounds) {
  dat %>%
    filter(
      word10_z >= bounds$word_lower,
      word10_z <= bounds$word_upper,
      education_years10 >= bounds$edu_lower,
      education_years10 <= bounds$edu_upper
    )
}

support_primary_data <- apply_support(analysis_primary, support_primary$bounds)
support_strict_data  <- apply_support(analysis_primary, support_strict$bounds)

support_counts <- bind_rows(
  analysis_primary %>% count(home_language10, name = "n") %>% mutate(sample = "Full primary sample"),
  support_primary_data %>% count(home_language10, name = "n") %>% mutate(sample = "1st-99th support"),
  support_strict_data %>% count(home_language10, name = "n") %>% mutate(sample = "5th-95th support")
) %>%
  relocate(sample)
RESULTS[["16_support_counts"]] <- support_counts

fit_support_model <- function(dat, label) {
  dat <- droplevels(dat)
  if (nrow(dat) < 100 || n_distinct(dat$home_language10) < 3) {
    return(tibble(
      model = label,
      term = "MODEL_NOT_FIT",
      estimate = NA_real_,
      std.error = NA_real_,
      statistic = NA_real_,
      p.value = NA_real_,
      conf.low = NA_real_,
      conf.high = NA_real_
    ))
  }

  des <- make_svy_design(dat)
  fit <- svyglm(formula_full, design = des)
  broom::tidy(fit, conf.int = TRUE) %>% mutate(model = label, .before = 1)
}

support_models <- bind_rows(
  fit_support_model(support_primary_data, "1st-99th common support"),
  fit_support_model(support_strict_data,  "5th-95th common support")
)
RESULTS[["17_support_models"]] <- support_models
RESULTS[["18_support_language_terms"]] <- support_models %>%
  filter(grepl("^home_language10", term) | term == "MODEL_NOT_FIT")

# Diagnostic plots: weighted distributions of the two key overlap variables.
p_support_word <- ggplot(
  analysis_primary,
  aes(x = word10, color = home_language10, weight = weight18_panel)
) +
  geom_density(linewidth = 0.9, adjust = 1.1, na.rm = TRUE) +
  labs(
    x = "2010 word score",
    y = "Weighted density",
    color = "2010 home language",
    title = "Baseline-score overlap across home-language groups",
    subtitle = "Weighted density; used to assess common support"
  ) +
  theme_minimal(base_size = 12) +
  theme(panel.grid.minor = element_blank(), legend.position = "top")

ggsave(
  file.path(FIG_DIR, "Figure_S1_common_support_word10.png"),
  p_support_word,
  width = 8.4,
  height = 5.0,
  dpi = 400
)

p_support_edu <- ggplot(
  analysis_primary,
  aes(x = education_years10, color = home_language10, weight = weight18_panel)
) +
  geom_density(linewidth = 0.9, adjust = 1.1, na.rm = TRUE) +
  labs(
    x = "Education years in 2010",
    y = "Weighted density",
    color = "2010 home language",
    title = "Education overlap across home-language groups",
    subtitle = "Weighted density; used to assess common support"
  ) +
  theme_minimal(base_size = 12) +
  theme(panel.grid.minor = element_blank(), legend.position = "top")

ggsave(
  file.path(FIG_DIR, "Figure_S2_common_support_education.png"),
  p_support_edu,
  width = 8.4,
  height = 5.0,
  dpi = 400
)


# 13. NONLINEAR BASELINE SENSITIVITY -----------------------------------------

# Natural spline for baseline word score; province and all social covariates
# are retained. df=4 is prespecified and avoids selecting curvature after
# looking at the result.
formula_spline_baseline <- update(
  formula_full,
  . ~ . - word10_z + splines::ns(word10_z, df = 4)
)

# More demanding sensitivity: allow nonlinear baseline and nonlinear education.
formula_spline_both <- update(
  formula_full,
  . ~ . - word10_z - education10_c +
    splines::ns(word10_z, df = 4) +
    splines::ns(education10_c, df = 4)
)

fit_spline_baseline <- svyglm(formula_spline_baseline, design = design_primary)
fit_spline_both     <- svyglm(formula_spline_both,     design = design_primary)

nonlinear_models <- bind_rows(
  broom::tidy(fit_spline_baseline, conf.int = TRUE) %>%
    mutate(model = "Natural spline baseline df=4", .before = 1),
  broom::tidy(fit_spline_both, conf.int = TRUE) %>%
    mutate(model = "Spline baseline + education df=4", .before = 1)
)
RESULTS[["19_nonlinear_models"]] <- nonlinear_models
RESULTS[["20_nonlinear_language_terms"]] <- nonlinear_models %>%
  filter(grepl("^home_language10", term))

# Combine common support + nonlinear baseline as the strongest extrapolation check.
fit_support_spline <- NULL
if (nrow(support_primary_data) >= 100 && n_distinct(support_primary_data$home_language10) == 3) {
  support_primary_data <- droplevels(support_primary_data)
  design_support_primary <- make_svy_design(support_primary_data)
  fit_support_spline <- svyglm(formula_spline_baseline, design = design_support_primary)
  RESULTS[["21_support_spline_model"]] <- broom::tidy(
    fit_support_spline,
    conf.int = TRUE
  ) %>% mutate(model = "1st-99th support + baseline spline", .before = 1)
  RESULTS[["22_support_spline_lang"]] <- RESULTS[["21_support_spline_model"]] %>%
    filter(grepl("^home_language10", term))
} else {
  RESULTS[["21_support_spline_model"]] <- tibble(
    note = "Model not fit because common-support sample was too small or lacked all three language groups."
  )
  RESULTS[["22_support_spline_lang"]] <- RESULTS[["21_support_spline_model"]]
}


# 14. MODERATION BY SCHOOLING AND BIRTH COHORT -------------------------------

fit_education_moderation <- svyglm(
  update(formula_full, . ~ . + home_language10:education10_c),
  design = design_primary
)

fit_cohort_moderation <- svyglm(
  update(formula_full, . ~ . + home_language10:birth_year_c),
  design = design_primary
)

education_wald <- regTermTest(
  fit_education_moderation,
  ~home_language10:education10_c,
  method = "Wald"
)
cohort_wald <- regTermTest(
  fit_cohort_moderation,
  ~home_language10:birth_year_c,
  method = "Wald"
)

RESULTS[["23_moderation_models"]] <- bind_rows(
  broom::tidy(fit_education_moderation, conf.int = TRUE) %>%
    mutate(model = "Education moderation", .before = 1),
  broom::tidy(fit_cohort_moderation, conf.int = TRUE) %>%
    mutate(model = "Birth-cohort moderation", .before = 1)
)
RESULTS[["24_moderation_wald"]] <- bind_rows(
  safe_regterm_table(education_wald, "Home language x education"),
  safe_regterm_table(cohort_wald, "Home language x birth cohort")
)

# Translate education interactions into a 6-year contrast for interpretability.
edu_terms <- broom::tidy(fit_education_moderation, conf.int = TRUE) %>%
  filter(grepl("home_language10.*:education10_c|education10_c:home_language10", term)) %>%
  mutate(
    change_per_6_years = 6 * estimate,
    CI_low_per_6_years = 6 * conf.low,
    CI_high_per_6_years = 6 * conf.high
  )
RESULTS[["25_education_6yr_scale"]] <- edu_terms


# 15. 2012 SELF-RATED MANDARIN / LANGUAGE-STABILITY MODULE -------------------

mechanism_status <- "not run"

if (RUN_2012_MECHANISM && !is.na(path_2012_person)) {
  tryCatch({
    raw12 <- read_selected(
      path_2012_person,
      c("pid", "qd1011", "qd1012", "qd201", "qz205"),
      "CFPS 2012 person"
    )

    duplicate_audit12 <- tibble(
      original_rows = nrow(raw12),
      exact_distinct_rows = nrow(distinct(raw12)),
      exact_duplicates_removed = nrow(raw12) - nrow(distinct(raw12))
    )
    RESULTS[["26_2012_duplicate_audit"]] <- duplicate_audit12

    raw12 <- raw12 %>% distinct()

    person12 <- raw12 %>%
      transmute(
        pid = clean_id(pid),
        mandarin_ability12 = valid_range(qd1011, 1, 5),
        dialect_ability12 = valid_range(qd1012, 1, 5),
        family_language12 = factor(
          case_when(
            valid_set(qd201, c(1, 2, 3, 4, 5, 6)) == 1 ~ "Mandarin",
            valid_set(qd201, c(1, 2, 3, 4, 5, 6)) == 2 ~ "Local dialect",
            valid_set(qd201, c(1, 2, 3, 4, 5, 6)) == 5 ~ "Non-local dialect",
            valid_set(qd201, c(1, 2, 3, 4, 5, 6)) == 6 ~ "Minority language",
            valid_set(qd201, c(1, 2, 3, 4, 5, 6)) %in% c(3, 4) ~ "Other",
            TRUE ~ NA_character_
          )
        ),
        interviewer_fluency12 = valid_range(qz205, 1, 7)
      )

    assert_unique_key(person12, "pid", "2012_person_after_exact_dedup")

    mechanism_data <- analysis_primary %>%
      left_join(person12, by = "pid") %>%
      filter(!is.na(mandarin_ability12)) %>%
      mutate(
        mandarin_ability12_z = weighted_z(mandarin_ability12, weight18_panel),
        dialect10 = as.integer(home_language10 == "Sinitic dialect"),
        minority_language10 = as.integer(home_language10 == "Minority language"),
        male = as.integer(sex == "Male"),
        urban = as.integer(urban10 == "Urban"),
        ethnic_minority = as.integer(ethnicity == "Ethnic minority"),
        psu_cluster12 = as.integer(factor(psu))
      )

    RESULTS[["27_mechanism_sample"]] <- mechanism_data %>%
      count(home_language10, name = "n")

    # Reviewer-requested representativeness audit for the 2012 proficiency subsample.
    mechanism_flag <- analysis_primary %>%
      left_join(person12 %>% select(pid, mandarin_ability12), by = "pid") %>%
      mutate(has_mandarin_ability12 = as.integer(!is.na(mandarin_ability12)))

    RESULTS[["27b_2012_inclusion_by_lang"]] <- mechanism_flag %>%
      group_by(home_language10) %>%
      summarise(
        primary_n = n(),
        proficiency_n = sum(has_mandarin_ability12),
        unweighted_inclusion_pct = 100 * mean(has_mandarin_ability12),
        weighted_inclusion_pct = 100 * weighted.mean(has_mandarin_ability12, weight18_panel),
        .groups = "drop"
      )

    sample_profile <- bind_rows(
      analysis_primary %>% mutate(sample = "Primary analytic sample"),
      mechanism_data %>% mutate(sample = "2012 Mandarin-proficiency subsample")
    ) %>%
      group_by(sample, home_language10) %>%
      summarise(
        n = n(),
        weight_sum = sum(weight18_panel),
        word10_mean = weighted.mean(word10, weight18_panel),
        education_mean = weighted.mean(education_years10, weight18_panel),
        age_mean = weighted.mean(age10, weight18_panel),
        male_pct = 100 * weighted.mean(as.integer(sex == "Male"), weight18_panel),
        urban_pct = 100 * weighted.mean(as.integer(urban10 == "Urban"), weight18_panel),
        ethnic_minority_pct = 100 * weighted.mean(as.integer(ethnicity == "Ethnic minority"), weight18_panel),
        .groups = "drop"
      ) %>%
      group_by(sample) %>%
      mutate(weighted_language_share_pct = 100 * weight_sum / sum(weight_sum)) %>%
      ungroup()
    RESULTS[["27c_2012_sample_profile"]] <- sample_profile

    province_matrix <- model.matrix(~province_f, data = mechanism_data)[, -1, drop = FALSE]
    colnames(province_matrix) <- make.names(colnames(province_matrix), unique = TRUE)
    mechanism_data <- bind_cols(mechanism_data, as.data.frame(province_matrix))

    control_names <- c(
      "word10_z", "age10_c", "male", "urban", "ethnic_minority",
      "education10_c", "parent_education_max_i",
      "parent_education_missing", "log_income_pc10_i", "income_missing",
      colnames(province_matrix)
    )
    controls <- paste(control_names, collapse = " + ")

    mechanism_model <- paste0(
      "mandarin_ability12_z ~ a_d*dialect10 + a_m*minority_language10 + ", controls, "\n",
      "word18_z ~ c_d*dialect10 + c_m*minority_language10 + b*mandarin_ability12_z + ", controls, "\n",
      "ind_dialect := a_d*b\n",
      "ind_minority := a_m*b\n",
      "total_dialect := c_d + ind_dialect\n",
      "total_minority := c_m + ind_minority"
    )

    fit_mechanism <- lavaan::sem(
      mechanism_model,
      data = mechanism_data,
      estimator = "MLR",
      missing = "fiml",
      sampling.weights = "weight18_panel",
      cluster = "psu_cluster12",
      fixed.x = FALSE
    )

    mechanism_estimates <- parameterEstimates(
      fit_mechanism,
      standardized = TRUE,
      ci = TRUE
    )
    RESULTS[["28_mechanism_path"]] <- mechanism_estimates

    # Short-term family-language stability sensitivity.
    language_trajectory_data <- mechanism_data %>%
      mutate(
        language12_binary = case_when(
          family_language12 == "Mandarin" ~ "Mandarin",
          family_language12 %in% c("Local dialect", "Non-local dialect", "Minority language") ~ "Non-Mandarin",
          TRUE ~ NA_character_
        ),
        language10_binary = if_else(home_language10 == "Mandarin", "Mandarin", "Non-Mandarin"),
        language_trajectory = factor(
          case_when(
            language10_binary == "Mandarin" & language12_binary == "Mandarin" ~ "Stable Mandarin",
            language10_binary == "Non-Mandarin" & language12_binary == "Mandarin" ~ "Shift to Mandarin",
            language10_binary == "Mandarin" & language12_binary == "Non-Mandarin" ~ "Shift from Mandarin",
            language10_binary == "Non-Mandarin" & language12_binary == "Non-Mandarin" ~ "Stable non-Mandarin",
            TRUE ~ NA_character_
          ),
          levels = c("Stable Mandarin", "Shift to Mandarin", "Shift from Mandarin", "Stable non-Mandarin")
        )
      ) %>%
      filter(!is.na(language_trajectory))

    design_trajectory12 <- svydesign(
      ids = ~psu_f,
      strata = ~strata_f,
      weights = ~weight18_panel,
      data = language_trajectory_data,
      nest = TRUE
    )

    fit_trajectory12 <- svyglm(
      update(formula_full, . ~ . - home_language10 + language_trajectory),
      design = design_trajectory12
    )

    RESULTS[["29_language_stability"]] <- broom::tidy(
      fit_trajectory12,
      conf.int = TRUE
    )

    mechanism_status <- paste0("completed; analytic n=", nrow(mechanism_data))
    add_status("2012 mechanism and language stability", "completed", mechanism_status)
  }, error = function(e) {
    mechanism_status <<- paste0("stopped safely: ", conditionMessage(e))
    add_status("2012 mechanism and language stability", "stopped safely", mechanism_status)
    RESULTS[["28_mechanism_path"]] <<- tibble(note = mechanism_status)
    RESULTS[["29_language_stability"]] <<- tibble(note = mechanism_status)
  })
} else {
  add_status("2012 mechanism and language stability", "not run", "2012 file unavailable or module disabled")
}


# 16. 2010-2014-2018 BALANCED-PANEL TRAJECTORY -------------------------------

if (RUN_BALANCED_TRAJECTORY && !is.na(path_2014_person)) {
  tryCatch({
    raw14 <- read_selected(
      path_2014_person,
      c("pid", "wordtest14", "mathtest14", "rswt_natpn1014"),
      "CFPS 2014 person"
    )

    person14 <- raw14 %>%
      transmute(
        pid = clean_id(pid),
        word14 = valid_range(wordtest14, 0, 34),
        math14 = valid_range(mathtest14, 0, 24),
        weight14_panel = if_else(num_clean(rswt_natpn1014) > 0, num_clean(rswt_natpn1014), NA_real_)
      )
    assert_unique_key(person14, "pid", "2014_person")

    panel14 <- panel %>% left_join(person14, by = "pid")

    balanced_word <- panel14 %>%
      filter(
        home_language10 %in% primary_language_levels,
        !is.na(word10), !is.na(word14), !is.na(word18),
        !is.na(weight18_panel), weight18_panel > 0,
        !is.na(psu), !is.na(strata), !is.na(age10), !is.na(sex),
        !is.na(urban10), !is.na(education_years10),
        !is.na(ethnicity), !is.na(province)
      ) %>%
      mutate(
        home_language10 = factor(home_language10, levels = primary_language_levels),
        province_f = factor(province),
        psu_f = factor(psu),
        strata_f = factor(strata),
        pid_cluster = factor(pid)
      )

    RESULTS[["30_balanced_word_counts"]] <- balanced_word %>%
      count(home_language10, name = "n")

    balanced_word_long <- balanced_word %>%
      select(
        pid, pid_cluster, psu_f, strata_f, weight18_panel, home_language10,
        age10_c, sex, urban10, ethnicity, education10_c,
        parent_education_max_i, parent_education_missing,
        log_income_pc10_i, income_missing, province_f,
        word10, word14, word18
      ) %>%
      pivot_longer(
        cols = c(word10, word14, word18),
        names_to = "wave",
        values_to = "word_score"
      ) %>%
      mutate(
        year = recode(wave, word10 = 2010L, word14 = 2014L, word18 = 2018L),
        year_f = factor(year, levels = c(2010, 2014, 2018))
      )

    design_balanced_word <- svydesign(
      ids = ~psu_f + pid_cluster,
      strata = ~strata_f,
      weights = ~weight18_panel,
      data = balanced_word_long,
      nest = TRUE
    )

    fit_balanced_word <- svyglm(
      word_score ~
        year_f * home_language10 +
        age10_c + I(age10_c^2) + sex + urban10 + ethnicity + education10_c +
        parent_education_max_i + parent_education_missing +
        log_income_pc10_i + income_missing + province_f,
      design = design_balanced_word
    )

    RESULTS[["31_balanced_word_model"]] <- broom::tidy(
      fit_balanced_word,
      conf.int = TRUE
    )

    balanced_word_wald <- regTermTest(
      fit_balanced_word,
      ~year_f:home_language10,
      method = "Wald"
    )
    RESULTS[["32_balanced_word_wald"]] <- safe_regterm_table(
      balanced_word_wald,
      "Year x home-language interaction, balanced word panel"
    )

    emm_balanced_word <- emmeans(
      fit_balanced_word,
      ~ year_f * home_language10,
      weights = "proportional"
    )
    balanced_word_emm <- safe_confint_df(emm_balanced_word)
    RESULTS[["33_balanced_word_emm"]] <- balanced_word_emm

    # Descriptive raw means in the SAME balanced sample.
    balanced_word_desc <- balanced_word_long %>%
      group_by(home_language10, year) %>%
      summarise(
        n = n(),
        weighted_mean = weighted.mean(word_score, weight18_panel),
        .groups = "drop"
      )
    RESULTS[["34_balanced_word_desc"]] <- balanced_word_desc

    p_balanced_word <- ggplot(
      balanced_word_emm,
      aes(
        x = as.numeric(as.character(year_f)),
        y = emmean,
        ymin = lower.CL,
        ymax = upper.CL,
        color = home_language10,
        group = home_language10
      )
    ) +
      geom_line(linewidth = 0.9) +
      geom_point(size = 2.5) +
      geom_errorbar(width = 0.25, linewidth = 0.55) +
      scale_x_continuous(breaks = c(2010, 2014, 2018)) +
      scale_color_manual(values = c(
        "Mandarin" = "#2457C5",
        "Sinitic dialect" = "#1F9D76",
        "Minority language" = "#E07A1F"
      )) +
      labs(
        x = NULL,
        y = "Adjusted comparable word score",
        color = "2010 home language",
        title = "Balanced-panel word-score trajectories",
        subtitle = "Same respondents observed in 2010, 2014, and 2018; 95% confidence intervals"
      ) +
      theme_minimal(base_size = 13) +
      theme(panel.grid.minor = element_blank(), legend.position = "top")

    print(p_balanced_word)
    ggsave(
      file.path(FIG_DIR, "Figure_1_balanced_word_trajectory.png"),
      p_balanced_word,
      width = 8.8,
      height = 5.4,
      dpi = 450
    )

    add_status(
      "2010-2014-2018 balanced word trajectory",
      "completed",
      paste0("balanced respondents n=", nrow(balanced_word))
    )

    # Optional parallel math trajectory, useful as supplementary evidence.
    if (RUN_MATH_TRAJECTORY) {
      balanced_math <- panel14 %>%
        filter(
          home_language10 %in% primary_language_levels,
          !is.na(math10), !is.na(math14), !is.na(math18),
          !is.na(weight18_panel), weight18_panel > 0,
          !is.na(psu), !is.na(strata), !is.na(age10), !is.na(sex),
          !is.na(urban10), !is.na(education_years10),
          !is.na(ethnicity), !is.na(province)
        ) %>%
        mutate(
          home_language10 = factor(home_language10, levels = primary_language_levels),
          province_f = factor(province),
          psu_f = factor(psu),
          strata_f = factor(strata),
          pid_cluster = factor(pid)
        )

      RESULTS[["35_balanced_math_counts"]] <- balanced_math %>%
        count(home_language10, name = "n")

      balanced_math_long <- balanced_math %>%
        select(
          pid, pid_cluster, psu_f, strata_f, weight18_panel, home_language10,
          age10_c, sex, urban10, ethnicity, education10_c,
          parent_education_max_i, parent_education_missing,
          log_income_pc10_i, income_missing, province_f,
          math10, math14, math18
        ) %>%
        pivot_longer(
          cols = c(math10, math14, math18),
          names_to = "wave",
          values_to = "math_score"
        ) %>%
        mutate(
          year = recode(wave, math10 = 2010L, math14 = 2014L, math18 = 2018L),
          year_f = factor(year, levels = c(2010, 2014, 2018))
        )

      design_balanced_math <- svydesign(
        ids = ~psu_f + pid_cluster,
        strata = ~strata_f,
        weights = ~weight18_panel,
        data = balanced_math_long,
        nest = TRUE
      )

      fit_balanced_math <- svyglm(
        math_score ~
          year_f * home_language10 +
          age10_c + I(age10_c^2) + sex + urban10 + ethnicity + education10_c +
          parent_education_max_i + parent_education_missing +
          log_income_pc10_i + income_missing + province_f,
        design = design_balanced_math
      )

      RESULTS[["36_balanced_math_model"]] <- broom::tidy(
        fit_balanced_math,
        conf.int = TRUE
      )

      balanced_math_wald <- regTermTest(
        fit_balanced_math,
        ~year_f:home_language10,
        method = "Wald"
      )
      RESULTS[["37_balanced_math_wald"]] <- safe_regterm_table(
        balanced_math_wald,
        "Year x home-language interaction, balanced math panel"
      )

      emm_balanced_math <- emmeans(
        fit_balanced_math,
        ~ year_f * home_language10,
        weights = "proportional"
      )
      balanced_math_emm <- safe_confint_df(emm_balanced_math)
      RESULTS[["38_balanced_math_emm"]] <- balanced_math_emm

      p_balanced_math <- ggplot(
        balanced_math_emm,
        aes(
          x = as.numeric(as.character(year_f)),
          y = emmean,
          ymin = lower.CL,
          ymax = upper.CL,
          color = home_language10,
          group = home_language10
        )
      ) +
        geom_line(linewidth = 0.9) +
        geom_point(size = 2.5) +
        geom_errorbar(width = 0.25, linewidth = 0.55) +
        scale_x_continuous(breaks = c(2010, 2014, 2018)) +
        scale_color_manual(values = c(
          "Mandarin" = "#2457C5",
          "Sinitic dialect" = "#1F9D76",
          "Minority language" = "#E07A1F"
        )) +
        labs(
          x = NULL,
          y = "Adjusted comparable math score",
          color = "2010 home language",
          title = "Balanced-panel math-score trajectories",
          subtitle = "Supplementary trajectory model; same respondents across three waves"
        ) +
        theme_minimal(base_size = 13) +
        theme(panel.grid.minor = element_blank(), legend.position = "top")

      ggsave(
        file.path(FIG_DIR, "Figure_S3_balanced_math_trajectory.png"),
        p_balanced_math,
        width = 8.8,
        height = 5.4,
        dpi = 400
      )

      add_status(
        "2010-2014-2018 balanced math trajectory",
        "completed",
        paste0("balanced respondents n=", nrow(balanced_math))
      )
    }

  }, error = function(e) {
    msg <- paste0("stopped safely: ", conditionMessage(e))
    add_status("2010-2014-2018 balanced trajectory", "stopped safely", msg)
    RESULTS[["31_balanced_word_model"]] <<- tibble(note = msg)
    RESULTS[["32_balanced_word_wald"]] <<- tibble(note = msg)
    RESULTS[["33_balanced_word_emm"]] <<- tibble(note = msg)
  })
} else {
  add_status("2010-2014-2018 balanced trajectory", "not run", "2014 file unavailable or module disabled")
}


# 17. OPTIONAL MULTIPLE-IMPUTATION SENSITIVITY -------------------------------

if (RUN_MI_SENSITIVITY) {
  suppressPackageStartupMessages({
    library(mice)
    library(mitools)
  })

  tryCatch({
    mi_source <- panel %>%
      filter(
        home_language10 %in% primary_language_levels,
        !is.na(word10), !is.na(word18),
        !is.na(weight18_panel), weight18_panel > 0,
        !is.na(psu), !is.na(strata), !is.na(age10), !is.na(sex), !is.na(province)
      ) %>%
      mutate(
        home_language10 = factor(home_language10, levels = primary_language_levels),
        province_f = factor(province),
        psu_f = factor(psu),
        strata_f = factor(strata),
        word10_z = weighted_z(word10, weight18_panel),
        word18_z = weighted_z(word18, weight18_panel)
      ) %>%
      select(
        word18_z, home_language10, word10_z, age10_c, sex, urban10,
        ethnicity, education10_c, parent_education_max,
        log_income_pc10, province_f, psu_f, strata_f, weight18_panel
      )

    method <- mice::make.method(mi_source)
    method[c("word18_z", "home_language10", "word10_z", "province_f", "psu_f", "strata_f", "weight18_panel")] <- ""

    predictor_matrix <- mice::make.predictorMatrix(mi_source)
    predictor_matrix[, c("psu_f", "strata_f")] <- 0
    predictor_matrix[c("psu_f", "strata_f", "weight18_panel"), ] <- 0

    imp <- mice::mice(
      mi_source,
      m = 20,
      maxit = 20,
      method = method,
      predictorMatrix = predictor_matrix,
      seed = 20260824,
      printFlag = FALSE
    )

    completed <- mice::complete(imp, action = "all")
    imp_list <- mitools::imputationList(completed)
    designs_mi <- svydesign(
      ids = ~psu_f,
      strata = ~strata_f,
      weights = ~weight18_panel,
      data = imp_list,
      nest = TRUE
    )

    fits_mi <- with(
      designs_mi,
      svyglm(
        word18_z ~
          home_language10 + word10_z + age10_c + I(age10_c^2) + sex +
          urban10 + ethnicity + education10_c + parent_education_max +
          log_income_pc10 + province_f
      )
    )

    pooled_mi <- mitools::MIcombine(fits_mi)
    mi_coef <- pooled_mi$coefficients
    mi_se <- sqrt(diag(pooled_mi$variance))
    mi_df <- pooled_mi$df
    if (length(mi_df) == 1) mi_df <- rep(mi_df, length(mi_coef))
    mi_t <- mi_coef / mi_se
    RESULTS[["39_MI_sensitivity"]] <- tibble(
      term = names(mi_coef),
      estimate = as.numeric(mi_coef),
      std.error = as.numeric(mi_se),
      df = as.numeric(mi_df),
      statistic = as.numeric(mi_t),
      p.value = 2 * pt(-abs(mi_t), df = mi_df),
      conf.low = mi_coef - qt(0.975, df = mi_df) * mi_se,
      conf.high = mi_coef + qt(0.975, df = mi_df) * mi_se
    )
    RESULTS[["39b_MI_language_terms"]] <- RESULTS[["39_MI_sensitivity"]] %>%
      filter(grepl("^home_language10", term))
    RESULTS[["39c_MI_summary_text"]] <- tibble(
      output = capture.output(summary(pooled_mi))
    )
    add_status("Multiple imputation sensitivity", "completed", "m=20; survey strata and PSU retained")
  }, error = function(e) {
    msg <- paste0("stopped safely: ", conditionMessage(e))
    RESULTS[["39_MI_sensitivity"]] <<- tibble(note = msg)
    add_status("Multiple imputation sensitivity", "stopped safely", msg)
  })
} else {
  add_status("Multiple imputation sensitivity", "not run", "RUN_MI_SENSITIVITY=FALSE")
}


# 18. ATTRITION AUDIT ---------------------------------------------------------

attrition_base <- panel %>%
  filter(
    home_language10 %in% primary_language_levels,
    !is.na(word10), !is.na(weight10_national), weight10_national > 0,
    !is.na(psu), !is.na(strata), !is.na(age10), !is.na(sex),
    !is.na(urban10), !is.na(education_years10),
    !is.na(ethnicity), !is.na(province)
  ) %>%
  mutate(
    home_language10 = factor(home_language10, levels = primary_language_levels),
    psu_f = factor(psu),
    strata_f = factor(strata),
    province_f = factor(province)
  )

design_attrition <- svydesign(
  ids = ~psu_f,
  strata = ~strata_f,
  weights = ~weight10_national,
  data = attrition_base,
  nest = TRUE
)

fit_attrition <- svyglm(
  retained18 ~
    home_language10 + word10 + age10_c + I(age10_c^2) + sex + urban10 +
    ethnicity + education10_c + parent_education_max_i +
    parent_education_missing + log_income_pc10_i + income_missing + province_f,
  design = design_attrition,
  family = quasibinomial()
)

RESULTS[["40_attrition_OR"]] <- broom::tidy(
  fit_attrition,
  conf.int = TRUE,
  exponentiate = TRUE
)



# 18B. REVIEWER-REQUESTED TRANSPARENCY AND OVERLAP ANALYSES -------------------

# These analyses directly address the reviewer requests concerning:
# - pooled minority-language composition;
# - province/ethnicity overlap and geographic identification;
# - multivariate overlap/propensity diagnostics;
# - floor/ceiling and score-distribution diagnostics;
# - language-specific sample flow;
# - expanded Table 1;
# - complex survey strata/PSU degrees of freedom.

RESULTS[["R00_reviewer_analysis_map"]] <- tibble(
  reviewer_issue = c(
    "Pooled minority-language heterogeneity",
    "Geographic and ethnic overlap",
    "Multivariate common support",
    "Floor/scale compression",
    "2012 proficiency-subsample representativeness",
    "Sample flow by language group",
    "Expanded Table 1",
    "Multiple imputation",
    "Complex survey specification"
  ),
  output_sheets = c(
    "R05-R07", "R08-R14", "R15-R21", "R22-R26",
    "27b-27c", "R01-R02", "R03-R04", "39-39c", "R27-R29"
  )
)

# ---- Labels for transparent descriptive tables ----
province_lookup <- c(
  `11`="Beijing", `12`="Tianjin", `13`="Hebei", `14`="Shanxi", `15`="Inner Mongolia",
  `21`="Liaoning", `22`="Jilin", `23`="Heilongjiang", `31`="Shanghai", `32`="Jiangsu",
  `33`="Zhejiang", `34`="Anhui", `35`="Fujian", `36`="Jiangxi", `37`="Shandong",
  `41`="Henan", `42`="Hubei", `43`="Hunan", `44`="Guangdong", `45`="Guangxi",
  `46`="Hainan", `50`="Chongqing", `51`="Sichuan", `52`="Guizhou", `53`="Yunnan",
  `54`="Tibet", `61`="Shaanxi", `62`="Gansu", `63`="Qinghai", `64`="Ningxia",
  `65`="Xinjiang"
)

ethnicity_lookup <- c(
  `1`="Han", `2`="Mongol", `3`="Hui", `4`="Tibetan", `5`="Uyghur", `6`="Miao",
  `7`="Yi", `8`="Zhuang", `9`="Bouyei", `10`="Korean", `11`="Manchu", `12`="Dong",
  `13`="Yao", `14`="Bai", `15`="Tujia", `16`="Hani", `17`="Kazakh", `18`="Dai",
  `19`="Li", `20`="Lisu", `21`="Va", `22`="She", `23`="Gaoshan", `24`="Lahu",
  `25`="Shui", `26`="Dongxiang", `27`="Naxi", `28`="Jingpo", `29`="Kyrgyz",
  `30`="Tu", `31`="Daur", `32`="Mulam", `33`="Qiang", `34`="Blang", `35`="Salar",
  `36`="Maonan", `37`="Gelao", `38`="Xibe", `39`="Achang", `40`="Pumi",
  `41`="Tajik", `42`="Nu", `43`="Uzbek", `44`="Russian", `45`="Ewenki",
  `46`="De'ang", `47`="Bonan", `48`="Yugur", `49`="Jing", `50`="Tatar",
  `51`="Dulong", `52`="Oroqen", `53`="Hezhe", `54`="Monba", `55`="Lhoba", `56`="Jino"
)

label_code <- function(x, lookup, prefix) {
  key <- as.character(x)
  lab <- unname(lookup[key])
  ifelse(is.na(lab), paste0(prefix, " ", key), lab)
}

analysis_primary <- analysis_primary %>%
  mutate(
    province_name = label_code(province, province_lookup, "Province code"),
    ethnicity_name = label_code(ethnicity_code10, ethnicity_lookup, "Ethnicity code")
  )

# ---- A. Detailed sample flow by home-language group ----
flow_base <- panel %>%
  filter(home_language10 %in% primary_language_levels) %>%
  mutate(
    home_language10 = factor(home_language10, levels = primary_language_levels),
    stage_focal_language = TRUE,
    stage_valid_2010_word = stage_focal_language & !is.na(word10),
    stage_linked_2018 = stage_valid_2010_word & linked18 == 1L,
    stage_valid_2018_word = stage_linked_2018 & !is.na(word18),
    stage_positive_panel_weight = stage_valid_2018_word & !is.na(weight18_panel) & weight18_panel > 0,
    stage_complete_core = stage_positive_panel_weight &
      !is.na(psu) & !is.na(strata) & !is.na(age10) & !is.na(sex) &
      !is.na(urban10) & !is.na(education_years10) & !is.na(ethnicity) &
      !is.na(ethnicity_code10) & !is.na(province)
  )

flow_stage_names <- c(
  stage_focal_language = "Valid focal 2010 home-language category",
  stage_valid_2010_word = "+ valid 2010 word score",
  stage_linked_2018 = "+ linked to 2018 person file",
  stage_valid_2018_word = "+ valid 2018 word score",
  stage_positive_panel_weight = "+ positive 2018 longitudinal weight",
  stage_complete_core = "+ complete design/core covariates"
)

review_flow_long <- purrr::imap_dfr(flow_stage_names, function(stage_label, stage_var) {
  flow_base %>%
    group_by(home_language10) %>%
    summarise(n = sum(.data[[stage_var]], na.rm = TRUE), .groups = "drop") %>%
    mutate(stage = stage_label, stage_var = stage_var, .before = 1)
}) %>%
  group_by(home_language10) %>%
  arrange(match(stage_var, names(flow_stage_names)), .by_group = TRUE) %>%
  mutate(
    denominator_focal = first(n),
    pct_of_focal = 100 * n / denominator_focal,
    pct_of_previous = 100 * n / lag(n),
    pct_of_previous = if_else(row_number() == 1, 100, pct_of_previous)
  ) %>%
  ungroup()
RESULTS[["R01_flow_by_language"]] <- review_flow_long

review_flow_total <- review_flow_long %>%
  group_by(stage, stage_var) %>%
  summarise(n = sum(n), .groups = "drop") %>%
  arrange(match(stage_var, names(flow_stage_names))) %>%
  mutate(
    pct_of_focal = 100 * n / first(n),
    pct_of_previous = 100 * n / lag(n),
    pct_of_previous = if_else(row_number() == 1, 100, pct_of_previous)
  )
RESULTS[["R02_flow_total"]] <- review_flow_total

# ---- B. Expanded Table 1 ----
weighted_mean_by_group <- function(var, label, multiplier = 1) {
  f <- as.formula(paste0("~", var))
  out <- as_tibble(svyby(f, ~home_language10, design_primary, svymean, na.rm = TRUE, vartype = "se"))
  value_col <- setdiff(names(out), c("home_language10", paste0("se.", var)))[1]
  se_col <- grep("^se", names(out), value = TRUE)[1]
  out %>%
    transmute(
      home_language10,
      characteristic = label,
      estimate = .data[[value_col]] * multiplier,
      SE = .data[[se_col]] * multiplier
    )
}

expanded_table1_long <- bind_rows(
  analysis_primary %>% count(home_language10, name = "unweighted_n") %>%
    transmute(home_language10, characteristic = "Unweighted N", estimate = as.numeric(unweighted_n), SE = NA_real_),
  weighted_mean_by_group("word10", "2010 word score"),
  weighted_mean_by_group("word18", "2018 word score"),
  weighted_mean_by_group("age10", "Age in 2010"),
  weighted_mean_by_group("education_years10", "Education years in 2010"),
  weighted_mean_by_group("male_ind", "Male, %", 100),
  weighted_mean_by_group("urban_ind", "Urban, %", 100),
  weighted_mean_by_group("ethnic_minority_ind", "Ethnic minority, %", 100)
)
RESULTS[["R03_expanded_Table1_long"]] <- expanded_table1_long
RESULTS[["R04_expanded_Table1_wide"]] <- expanded_table1_long %>%
  mutate(display = if_else(is.na(SE), sprintf("%.0f", estimate), sprintf("%.2f (%.2f)", estimate, SE))) %>%
  select(characteristic, home_language10, display) %>%
  pivot_wider(names_from = home_language10, values_from = display)

# ---- C. Pooled minority-language category: ethnicity and province composition ----
minority_comp <- analysis_primary %>%
  filter(home_language10 == "Minority language")

RESULTS[["R05_minority_ethnicity_comp"]] <- minority_comp %>%
  group_by(ethnicity_code10, ethnicity_name) %>%
  summarise(
    unweighted_n = n(),
    weight_sum = sum(weight18_panel, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  mutate(weighted_percent = 100 * weight_sum / sum(weight_sum)) %>%
  arrange(desc(weighted_percent), desc(unweighted_n))

RESULTS[["R06_minority_province_comp"]] <- minority_comp %>%
  group_by(province, province_name) %>%
  summarise(
    unweighted_n = n(),
    weight_sum = sum(weight18_panel, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  mutate(weighted_percent = 100 * weight_sum / sum(weight_sum)) %>%
  arrange(desc(weighted_percent), desc(unweighted_n))

RESULTS[["R07_language_measure_limit"]] <- tibble(
  item = c(
    "2010 home-language item",
    "Specific minority-language identity",
    "What can be disaggregated in the current analysis-ready file"
  ),
  detail = c(
    "qd2: Mandarin / Sinitic dialect / minority language / other",
    "Not available as a separate variable in the supplied 2010 analysis-ready file",
    "Ethnicity code and province can be described; language-specific leave-one-language-out analysis cannot be identified from qd2"
  )
)

# ---- D. Province x language and ethnicity x language overlap ----
province_lang <- analysis_primary %>%
  mutate(lang_key = recode(as.character(home_language10),
                           "Mandarin" = "Mandarin",
                           "Sinitic dialect" = "Sinitic",
                           "Minority language" = "Minority")) %>%
  group_by(province, province_name, lang_key) %>%
  summarise(n = n(), weight_sum = sum(weight18_panel), .groups = "drop") %>%
  pivot_wider(
    names_from = lang_key,
    values_from = c(n, weight_sum),
    values_fill = list(n = 0, weight_sum = 0)
  ) %>%
  mutate(
    overlap_MM_any = n_Mandarin > 0 & n_Minority > 0,
    overlap_MM_n5 = n_Mandarin >= 5 & n_Minority >= 5,
    overlap_MM_n10 = n_Mandarin >= 10 & n_Minority >= 10
  ) %>%
  arrange(desc(n_Minority), province)
RESULTS[["R08_province_language_overlap"]] <- province_lang

ethnicity_lang <- analysis_primary %>%
  mutate(
    ethnicity_name = label_code(ethnicity_code10, ethnicity_lookup, "Ethnicity code"),
    lang_key = recode(as.character(home_language10),
                      "Mandarin" = "Mandarin",
                      "Sinitic dialect" = "Sinitic",
                      "Minority language" = "Minority")
  ) %>%
  group_by(ethnicity_code10, ethnicity_name, lang_key) %>%
  summarise(n = n(), weight_sum = sum(weight18_panel), .groups = "drop") %>%
  pivot_wider(
    names_from = lang_key,
    values_from = c(n, weight_sum),
    values_fill = list(n = 0, weight_sum = 0)
  ) %>%
  mutate(
    overlap_MM_any = n_Mandarin > 0 & n_Minority > 0,
    overlap_MM_n5 = n_Mandarin >= 5 & n_Minority >= 5,
    overlap_MM_n10 = n_Mandarin >= 10 & n_Minority >= 10
  ) %>%
  arrange(desc(n_Minority), ethnicity_code10)
RESULTS[["R09_ethnicity_language_overlap"]] <- ethnicity_lang

# Share of minority-language respondents located in provinces that also contain Mandarin comparators.
minority_prov_weights <- minority_comp %>%
  group_by(province) %>%
  summarise(minority_w = sum(weight18_panel), minority_n = n(), .groups = "drop")

geo_overlap_summary <- purrr::map_dfr(GEO_MIN_COUNTS, function(k) {
  eligible <- province_lang %>% filter(n_Mandarin >= k, n_Minority >= k) %>% pull(province)
  tibble(
    minimum_n_in_each_group = k,
    eligible_provinces = length(eligible),
    minority_unweighted_n_in_eligible = sum(minority_comp$province %in% eligible),
    minority_unweighted_percent = 100 * mean(minority_comp$province %in% eligible),
    minority_weighted_percent = 100 * sum(minority_comp$weight18_panel[minority_comp$province %in% eligible]) /
      sum(minority_comp$weight18_panel)
  )
})
RESULTS[["R10_geo_overlap_summary"]] <- geo_overlap_summary

# Geographic concentration: weighted share of minority-language group in the top 1/3/5 provinces.
minority_prov_rank <- RESULTS[["R06_minority_province_comp"]] %>%
  arrange(desc(weighted_percent)) %>%
  mutate(rank = row_number(), cumulative_weighted_percent = cumsum(weighted_percent))
RESULTS[["R11_minority_province_rank"]] <- minority_prov_rank

# Full-model coefficient within progressively stricter province-overlap sets.
fit_full_subset <- function(dat, label) {
  dat <- droplevels(dat)
  if (nrow(dat) < 100 || n_distinct(dat$home_language10) < 2) {
    return(tibble(model = label, n = nrow(dat), provinces = n_distinct(dat$province),
                  term = NA_character_, estimate = NA_real_, std.error = NA_real_,
                  conf.low = NA_real_, conf.high = NA_real_, p.value = NA_real_))
  }
  des <- make_svy_design(dat)
  fit <- svyglm(formula_full, design = des)
  broom::tidy(fit, conf.int = TRUE) %>%
    filter(term == "home_language10Minority language") %>%
    mutate(model = label, n = nrow(dat), provinces = n_distinct(dat$province), .before = 1)
}

RESULTS[["R12_geo_restricted_models"]] <- purrr::map_dfr(GEO_MIN_COUNTS, function(k) {
  eligible <- province_lang %>% filter(n_Mandarin >= k, n_Minority >= k) %>% pull(province)
  fit_full_subset(
    analysis_primary %>% filter(province %in% eligible),
    paste0("Provinces with >=", k, " Mandarin and >=", k, " minority-language respondents")
  )
})

# Leave-one-province-out sensitivity: does any single province drive the pooled coefficient?
full_minority_est <- unname(coef(fit_full)["home_language10Minority language"])
provinces_with_minority <- sort(unique(minority_comp$province))
RESULTS[["R13_leave_one_province_out"]] <- purrr::map_dfr(provinces_with_minority, function(p) {
  dat <- analysis_primary %>% filter(province != p) %>% droplevels()
  out <- tryCatch({
    fit <- svyglm(formula_full, design = make_svy_design(dat))
    tt <- broom::tidy(fit, conf.int = TRUE) %>%
      filter(term == "home_language10Minority language")
    if (nrow(tt) == 0) stop("minority-language coefficient not estimable")
    tt %>% transmute(
      omitted_province = p,
      omitted_province_name = label_code(p, province_lookup, "Province code"),
      omitted_minority_n = sum(analysis_primary$province == p & analysis_primary$home_language10 == "Minority language"),
      estimate = estimate,
      std.error = std.error,
      conf.low = conf.low,
      conf.high = conf.high,
      p.value = p.value,
      change_from_full = estimate - full_minority_est
    )
  }, error = function(e) {
    tibble(
      omitted_province = p,
      omitted_province_name = label_code(p, province_lookup, "Province code"),
      omitted_minority_n = sum(analysis_primary$province == p & analysis_primary$home_language10 == "Minority language"),
      estimate = NA_real_, std.error = NA_real_, conf.low = NA_real_, conf.high = NA_real_, p.value = NA_real_,
      change_from_full = NA_real_
    )
  })
  out
}) %>% arrange(desc(abs(change_from_full)))

# Stronger ethnicity adjustment: exact ethnicity-code fixed effects instead of Han/minority binary.
RESULTS[["R14_exact_ethnicity_FE"]] <- tryCatch({
  formula_exact_ethnicity <- update(formula_full, . ~ . - ethnicity + ethnicity_code_f)
  fit_exact_ethnicity <- svyglm(formula_exact_ethnicity, design = design_primary)
  broom::tidy(fit_exact_ethnicity, conf.int = TRUE) %>%
    filter(grepl("^home_language10", term)) %>%
    mutate(model = "Exact ethnicity-code FE + province FE", .before = 1)
}, error = function(e) {
  tibble(note = paste0("Exact ethnicity fixed-effect sensitivity could not be estimated: ", conditionMessage(e)))
})

# ---- E. Multivariate propensity/overlap diagnostics: minority language vs Mandarin ----
ps_data <- analysis_primary %>%
  filter(home_language10 %in% c("Mandarin", "Minority language")) %>%
  droplevels() %>%
  mutate(minority_ind = as.integer(home_language10 == "Minority language"))

ps_design <- make_svy_design(ps_data)
ps_formula <- minority_ind ~
  word10_z + age10_c + I(age10_c^2) + sex + urban10 + ethnicity + education10_c +
  parent_education_max_i + parent_education_missing + log_income_pc10_i + income_missing + province_f

fit_ps <- svyglm(ps_formula, design = ps_design, family = quasibinomial())
ps_data$propensity_minority <- pmin(pmax(as.numeric(predict(fit_ps, type = "response")), 1e-6), 1 - 1e-6)

RESULTS[["R15_propensity_model"]] <- broom::tidy(fit_ps, conf.int = TRUE)

ps_quantiles <- ps_data %>%
  group_by(home_language10) %>%
  group_modify(~{
    q <- weighted_quantile(.x$propensity_minority, .x$weight18_panel,
                           c(.01, .05, .25, .50, .75, .95, .99))
    tibble(probability = c(.01, .05, .25, .50, .75, .95, .99), propensity = q)
  }) %>%
  ungroup()
RESULTS[["R16_propensity_quantiles"]] <- ps_quantiles

ps_bounds_by_group <- ps_data %>%
  group_by(home_language10) %>%
  summarise(
    p01 = weighted_quantile(propensity_minority, weight18_panel, PS_SUPPORT_PROBS[1]),
    p99 = weighted_quantile(propensity_minority, weight18_panel, PS_SUPPORT_PROBS[2]),
    .groups = "drop"
  )
ps_lower <- max(ps_bounds_by_group$p01, na.rm = TRUE)
ps_upper <- min(ps_bounds_by_group$p99, na.rm = TRUE)

RESULTS[["R17_propensity_support"]] <- ps_data %>%
  group_by(home_language10) %>%
  summarise(
    n = n(),
    weighted_pct_ps_lt_005 = 100 * weighted.mean(propensity_minority < .05, weight18_panel),
    weighted_pct_ps_gt_095 = 100 * weighted.mean(propensity_minority > .95, weight18_panel),
    weighted_pct_inside_1_99_intersection = 100 * weighted.mean(
      propensity_minority >= ps_lower & propensity_minority <= ps_upper,
      weight18_panel
    ),
    .groups = "drop"
  ) %>%
  mutate(common_support_lower = ps_lower, common_support_upper = ps_upper)

# Standardized mean differences before and after overlap weighting.
ps_data <- ps_data %>%
  mutate(
    overlap_weight = weight18_panel * if_else(minority_ind == 1, 1 - propensity_minority, propensity_minority)
  )

weighted_mean_var <- function(x, w) {
  ok <- is.finite(x) & is.finite(w) & w > 0
  if (sum(ok) < 2) return(c(mean = NA_real_, var = NA_real_))
  mu <- sum(x[ok] * w[ok]) / sum(w[ok])
  vv <- sum(w[ok] * (x[ok] - mu)^2) / sum(w[ok])
  c(mean = mu, var = vv)
}

smd_matrix <- model.matrix(
  ~ word10_z + age10_c + I(age10_c^2) + sex + urban10 + ethnicity + education10_c +
    parent_education_max_i + parent_education_missing + log_income_pc10_i + income_missing + province_f,
  data = ps_data
)[, -1, drop = FALSE]

calc_smd_table <- function(w, label) {
  purrr::map_dfr(seq_len(ncol(smd_matrix)), function(j) {
    x <- smd_matrix[, j]
    st <- weighted_mean_var(x[ps_data$minority_ind == 1], w[ps_data$minority_ind == 1])
    sc <- weighted_mean_var(x[ps_data$minority_ind == 0], w[ps_data$minority_ind == 0])
    den <- sqrt((st["var"] + sc["var"]) / 2)
    smd <- ifelse(is.finite(den) && den > 0, (st["mean"] - sc["mean"]) / den, NA_real_)
    tibble(weighting = label, covariate = colnames(smd_matrix)[j], SMD = smd, abs_SMD = abs(smd))
  })
}

smd_before <- calc_smd_table(ps_data$weight18_panel, "Original longitudinal weights")
smd_after  <- calc_smd_table(ps_data$overlap_weight, "Overlap weights")
RESULTS[["R18_propensity_SMD"]] <- bind_rows(smd_before, smd_after) %>% arrange(weighting, desc(abs_SMD))
RESULTS[["R19_propensity_balance_summary"]] <- bind_rows(smd_before, smd_after) %>%
  group_by(weighting) %>%
  summarise(
    max_abs_SMD = max(abs_SMD, na.rm = TRUE),
    mean_abs_SMD = mean(abs_SMD, na.rm = TRUE),
    n_abs_SMD_gt_010 = sum(abs_SMD > .10, na.rm = TRUE),
    n_abs_SMD_gt_020 = sum(abs_SMD > .20, na.rm = TRUE),
    .groups = "drop"
  )

# Propensity common-support restriction sensitivity.
ps_support_data <- ps_data %>%
  filter(propensity_minority >= ps_lower, propensity_minority <= ps_upper) %>%
  droplevels()
RESULTS[["R20_propensity_support_model"]] <- tryCatch({
  fit <- svyglm(formula_full, design = make_svy_design(ps_support_data))
  broom::tidy(fit, conf.int = TRUE) %>%
    filter(term == "home_language10Minority language") %>%
    mutate(
      model = "Model 3 within 1st-99th percentile propensity-support intersection",
      n = nrow(ps_support_data),
      minority_n = sum(ps_support_data$home_language10 == "Minority language"),
      .before = 1
    )
}, error = function(e) tibble(note = paste0("PS-support model failed: ", conditionMessage(e))))

# Overlap-weighted outcome contrast; this targets the covariate-overlap population and is a sensitivity estimand.
design_ow <- svydesign(
  ids = ~psu_f,
  strata = ~strata_f,
  weights = ~overlap_weight,
  data = ps_data,
  nest = TRUE
)
fit_ow <- svyglm(
  word18_z ~ minority_ind + word10_z + age10_c + I(age10_c^2) + sex + urban10 + ethnicity +
    education10_c + parent_education_max_i + parent_education_missing +
    log_income_pc10_i + income_missing + province_f,
  design = design_ow
)
RESULTS[["R21_overlap_weighted_outcome"]] <- broom::tidy(fit_ow, conf.int = TRUE) %>%
  filter(term == "minority_ind") %>%
  mutate(model = "Overlap-weighted adjusted minority-language vs Mandarin contrast", .before = 1)

# Propensity density figure.
p_ps <- ggplot(ps_data, aes(x = propensity_minority, color = home_language10, weight = weight18_panel)) +
  geom_density(linewidth = .9, adjust = 1.1, na.rm = TRUE) +
  geom_vline(xintercept = c(ps_lower, ps_upper), linetype = 2, color = "grey50") +
  labs(
    x = "Estimated probability of minority-language home use",
    y = "Weighted density",
    color = "2010 home language",
    title = "Multivariate propensity-score overlap",
    subtitle = "Minority-language versus Mandarin homes; dashed lines mark the 1st-99th percentile overlap intersection"
  ) +
  theme_minimal(base_size = 12) +
  theme(panel.grid.minor = element_blank(), legend.position = "top")
ggsave(file.path(FIG_DIR, "Figure_R1_propensity_overlap.png"), p_ps, width = 8.6, height = 5.2, dpi = 450)

# ---- F. Weighted score distributions, floor/ceiling diagnostics, and floor sensitivity ----
score_distribution <- function(dat, score_var, score_max, wave_label) {
  dat %>%
    filter(!is.na(.data[[score_var]])) %>%
    group_by(home_language10, score = .data[[score_var]]) %>%
    summarise(unweighted_n = n(), weight_sum = sum(weight18_panel), .groups = "drop") %>%
    group_by(home_language10) %>%
    complete(score = 0:score_max, fill = list(unweighted_n = 0, weight_sum = 0)) %>%
    mutate(weighted_percent = 100 * weight_sum / sum(weight_sum), wave = wave_label, .before = 1) %>%
    ungroup()
}
RESULTS[["R22_word_score_distribution"]] <- bind_rows(
  score_distribution(analysis_primary, "word10", 34, "2010"),
  score_distribution(analysis_primary, "word18", 34, "2018")
)
RESULTS[["R23_math_score_distribution"]] <- bind_rows(
  score_distribution(analysis_domain, "math10", 24, "2010"),
  score_distribution(analysis_domain, "math18", 24, "2018")
)

floor_ceiling_summary <- function(dat, score_var, score_max, wave_label) {
  dat %>%
    group_by(home_language10) %>%
    summarise(
      n = sum(!is.na(.data[[score_var]])),
      weighted_mean = weighted.mean(.data[[score_var]], weight18_panel, na.rm = TRUE),
      weighted_p01 = weighted_quantile(.data[[score_var]], weight18_panel, .01),
      weighted_p05 = weighted_quantile(.data[[score_var]], weight18_panel, .05),
      weighted_p25 = weighted_quantile(.data[[score_var]], weight18_panel, .25),
      weighted_median = weighted_quantile(.data[[score_var]], weight18_panel, .50),
      weighted_p75 = weighted_quantile(.data[[score_var]], weight18_panel, .75),
      exact_floor_pct = 100 * weighted.mean(.data[[score_var]] == 0, weight18_panel, na.rm = TRUE),
      bottom_10pct_scale_pct = 100 * weighted.mean(.data[[score_var]] <= floor(.10 * score_max), weight18_panel, na.rm = TRUE),
      bottom_15pct_scale_pct = 100 * weighted.mean(.data[[score_var]] <= floor(.15 * score_max), weight18_panel, na.rm = TRUE),
      exact_ceiling_pct = 100 * weighted.mean(.data[[score_var]] == score_max, weight18_panel, na.rm = TRUE),
      top_10pct_scale_pct = 100 * weighted.mean(.data[[score_var]] >= ceiling(.90 * score_max), weight18_panel, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(wave = wave_label, outcome = score_var, .before = 1)
}
RESULTS[["R24_floor_ceiling_summary"]] <- bind_rows(
  floor_ceiling_summary(analysis_primary, "word10", 34, "2010"),
  floor_ceiling_summary(analysis_primary, "word18", 34, "2018"),
  floor_ceiling_summary(analysis_domain, "math10", 24, "2010"),
  floor_ceiling_summary(analysis_domain, "math18", 24, "2018")
)

fit_floor_sensitivity <- function(cutoff, require_outcome_above = FALSE) {
  dat <- analysis_primary %>%
    filter(word10 > cutoff)
  if (require_outcome_above) dat <- dat %>% filter(word18 > cutoff)
  dat <- droplevels(dat)
  label <- if (require_outcome_above) {
    paste0("Exclude word scores <=", cutoff, " at both 2010 and 2018")
  } else {
    paste0("Exclude baseline word scores <=", cutoff)
  }
  tryCatch({
    fit <- svyglm(formula_full, design = make_svy_design(dat))
    broom::tidy(fit, conf.int = TRUE) %>%
      filter(grepl("^home_language10", term)) %>%
      mutate(
        model = label,
        n = nrow(dat),
        minority_n = sum(dat$home_language10 == "Minority language"),
        .before = 1
      )
  }, error = function(e) tibble(model = label, note = conditionMessage(e)))
}
RESULTS[["R25_floor_exclusion_sensitivity"]] <- bind_rows(
  purrr::map_dfr(FLOOR_CUTOFFS, ~fit_floor_sensitivity(.x, FALSE)),
  purrr::map_dfr(FLOOR_CUTOFFS, ~fit_floor_sensitivity(.x, TRUE))
)

# The supplied analysis-ready files contain total scores, not item-level response strings or stopping positions.
RESULTS[["R26_item_stop_availability"]] <- tibble(
  requested_diagnostic = c("Education-dependent entry rule", "Three-error stopping-point distribution", "Item-level reliability / DIF / invariance"),
  status_in_supplied_analysis_ready_files = c(
    "Education years are available; administration rule must be documented from CFPS instrument materials",
    "Not recoverable from the supplied total-score variables alone",
    "Not recoverable because item-level responses are not included in the supplied analysis-ready files"
  )
)

# ---- G. Survey-design specification and denominator degrees of freedom ----
strata_psu <- analysis_primary %>% distinct(strata_f, psu_f) %>% count(strata_f, name = "n_psu")
RESULTS[["R27_strata_PSU_counts"]] <- strata_psu

review_design_audit <- tibble(
  item = c(
    "Sampling strata variable",
    "Number of represented strata",
    "PSU variable",
    "Number of represented PSUs",
    "Survey design degrees of freedom",
    "Singleton-stratum handling",
    "Model 1 residual df",
    "Model 2 residual df",
    "Model 3 residual df",
    "Cross-domain model residual df"
  ),
  value = c(
    "2010 subpopulation (six CFPS sampling frames)",
    as.character(n_distinct(analysis_primary$strata_f)),
    "2010 PSU (coalesced with 2018 PSU only if needed)",
    as.character(n_distinct(analysis_primary$psu_f)),
    as.character(survey::degf(design_primary)),
    "survey.lonely.psu = 'adjust'",
    as.character(df.residual(fit_minimal)),
    as.character(df.residual(fit_social)),
    as.character(df.residual(fit_full)),
    as.character(df.residual(fit_domain))
  )
)
RESULTS[["R28_survey_design_DF_audit"]] <- review_design_audit

# ---- H. Three-wave / repeated-measures design audit ----
trajectory_design_audit <- tibble(
  model = c("Cross-domain repeated records", "Balanced word trajectory", "Balanced math trajectory"),
  weight = c("2018 national longitudinal weight", "2018 national longitudinal weight", "2018 national longitudinal weight"),
  first_stage_strata = c("2010 subpopulation", "2010 subpopulation", "2010 subpopulation"),
  first_stage_cluster = c("PSU", "PSU", "PSU"),
  second_stage_cluster = c("respondent", "respondent", "respondent"),
  interpretation = c(
    "Repeated domain records nested within respondent and PSU",
    "Repeated wave records nested within respondent and PSU; observed-score trajectory",
    "Repeated wave records nested within respondent and PSU; observed-score trajectory"
  )
)
RESULTS[["R29_repeated_design_audit"]] <- trajectory_design_audit


# 19. FINAL AUDIT -------------------------------------------------------------

audit_summary <- tibble(
  item = c(
    "2010 person rows",
    "2018 person rows",
    "Primary analytic rows",
    "Primary analytic PSUs",
    "Primary analytic strata",
    "Survey design df",
    "Domain-complete analytic rows",
    "Primary word outcome minimum",
    "Primary word outcome maximum",
    "2022 module included",
    "Respondent-level data exported"
  ),
  value = c(
    as.character(nrow(person10)),
    as.character(nrow(person18)),
    as.character(nrow(analysis_primary)),
    as.character(n_distinct(analysis_primary$psu)),
    as.character(n_distinct(analysis_primary$strata_f)),
    as.character(survey::degf(design_primary)),
    as.character(nrow(analysis_domain)),
    as.character(min(analysis_primary$word18, na.rm = TRUE)),
    as.character(max(analysis_primary$word18, na.rm = TRUE)),
    "No",
    "No"
  )
)
RESULTS[["41_final_audit"]] <- audit_summary

RESULTS[["42_module_status"]] <- STATUS

session_lines <- capture.output(sessionInfo())
RESULTS[["43_session_info"]] <- tibble(output = session_lines)

# Compact manuscript-oriented summary so the key results can be reviewed
# without searching dozens of model sheets.
key_primary <- RESULTS[["06_primary_language_terms"]] %>%
  transmute(
    analysis = paste0("Primary word model: ", model),
    group_or_contrast = term,
    estimate = estimate,
    SE = std.error,
    CI_low = conf.low,
    CI_high = conf.high,
    p_value = p.value,
    note = "Outcome SD units"
  )

key_domain <- RESULTS[["11_domain_contrasts"]] %>%
  transmute(
    analysis = paste0("Cross-domain contrast: ", domain),
    group_or_contrast = language_contrast,
    estimate = estimate,
    SE = SE,
    CI_low = lower.CL,
    CI_high = upper.CL,
    p_value = p.value,
    note = "Difference versus Mandarin; domain-specific SD units"
  )

key_tost <- RESULTS[["13_TOST_equivalence"]] %>%
  transmute(
    analysis = paste0("TOST SESOI +/-", sprintf("%.2f", SESOI)),
    group_or_contrast = language_group,
    estimate = estimate,
    SE = SE,
    CI_low = CI90_low,
    CI_high = CI90_high,
    p_value = TOST_p,
    note = paste0(
      "90% CI; equivalent=", equivalent_two_sided,
      "; rules out extra verbal penalty > SESOI=", noninferior_extra_verbal
    )
  )

key_support <- RESULTS[["18_support_language_terms"]] %>%
  filter(term != "MODEL_NOT_FIT") %>%
  transmute(
    analysis = paste0("Common support: ", model),
    group_or_contrast = term,
    estimate = estimate,
    SE = std.error,
    CI_low = conf.low,
    CI_high = conf.high,
    p_value = p.value,
    note = "Same outcome scaling as full primary sample"
  )

key_spline <- RESULTS[["20_nonlinear_language_terms"]] %>%
  transmute(
    analysis = paste0("Nonlinear sensitivity: ", model),
    group_or_contrast = term,
    estimate = estimate,
    SE = std.error,
    CI_low = conf.low,
    CI_high = conf.high,
    p_value = p.value,
    note = "Natural-spline sensitivity"
  )

RESULTS[["00_key_results"]] <- bind_rows(
  key_primary, key_domain, key_tost, key_support, key_spline
)


# 20. README SHEET ------------------------------------------------------------

sheet_descriptions <- tibble(
  sheet_prefix = c(
    "01", "04", "05-07", "09-13", "14-18", "19-22", "23-25",
    "26-29", "30-38", "40", "41-43"
  ),
  content = c(
    "Sample flow",
    "Weighted descriptives",
    "Primary 2010-2018 word-score models and attenuation",
    "Cross-domain model, true contrasts versus Mandarin, Wald test, TOST/equivalence",
    "Common-support quantiles, overlap bounds, restricted-sample models",
    "Nonlinear baseline / education spline sensitivity",
    "Education and birth-cohort moderation",
    "2012 self-rated Mandarin and short-term language-stability analyses",
    "2010-2014-2018 balanced-panel word and math trajectory analyses",
    "2010-2018 attrition model",
    "Audit, module status, session information"
  )
)

readme <- bind_rows(
  tibble(
    item = c(
      "Purpose",
      "Primary SESOI",
      "Sensitivity SESOI",
      "Common support primary",
      "Common support strict",
      "Primary trajectory weight",
      "2022 analysis",
      "Tabular output",
      "Figures"
    ),
    detail = c(
      "Journal of Intelligence enhanced analysis for home language and cognitive test score interpretation",
      "+/- 0.10 SD",
      "+/- 0.15 SD",
      "Intersection of group-specific weighted 1st-99th percentiles for baseline word score and education",
      "Intersection of group-specific weighted 5th-95th percentiles",
      "2018 national panel weight applied to the balanced 2010-2014-2018 respondent sample",
      "Removed intentionally",
      "All data-like/tabular results are stored in one multi-sheet Excel workbook",
      "Only figures are saved separately as PNG files"
    )
  ),
  tibble(
    item = paste0("Sheet map ", sheet_descriptions$sheet_prefix),
    detail = sheet_descriptions$content
  )
)
RESULTS <- append(list(`README` = readme), RESULTS)


# 21. WRITE ONE COMBINED EXCEL WORKBOOK --------------------------------------

# writexl requires unique Excel-safe names <=31 characters.
clean_sheet_name <- function(x) {
  x <- gsub("[:\\\\/?*\\[\\]]", "_", x)
  substr(x, 1, 31)
}

sheet_names <- vapply(names(RESULTS), clean_sheet_name, character(1))

# Guarantee uniqueness after truncation.
if (anyDuplicated(sheet_names)) {
  for (i in seq_along(sheet_names)) {
    if (sum(sheet_names[seq_len(i)] == sheet_names[i]) > 1) {
      suffix <- paste0("_", sum(sheet_names[seq_len(i)] == sheet_names[i]))
      sheet_names[i] <- paste0(substr(sheet_names[i], 1, 31 - nchar(suffix)), suffix)
    }
  }
}

names(RESULTS) <- sheet_names

# Ensure every element is a plain data.frame/tibble.
RESULTS <- lapply(RESULTS, function(x) {
  if (is.null(x)) return(data.frame(note = "No result generated"))
  as.data.frame(x)
})

combined_path <- file.path(OUT_DIR, "CFPS_all_results_REVIEWER_REVISION.xlsx")
writexl::write_xlsx(
  RESULTS,
  path = combined_path,
  col_names = TRUE,
  format_headers = TRUE
)

# Session info is also kept as a plain text file for reproducibility.
writeLines(session_lines, file.path(OUT_DIR, "sessionInfo.txt"), useBytes = TRUE)

cat("\n============================================================\n")
cat("分析完成。\n")
cat("所有表格型/数据型结果已合并到：\n")
cat("  ", combined_path, "\n", sep = "")
cat("图形保存在：\n")
cat("  ", FIG_DIR, "\n", sep = "")
cat("Excel工作表数量：", length(RESULTS), "\n")
cat("2022模块：已完全移除\n")
cat("审稿补充模块：province/ethnicity overlap, propensity overlap, floor sensitivity, sample flow, expanded Table 1, MI：已加入\n")
cat("============================================================\n")

message("Completed successfully.")
