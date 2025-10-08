# Clear R environment
rm(list = ls(all = TRUE))

# Load libraries
library(readxl)
library(dplyr)
library(effsize)
library(kableExtra)
library(knitr)
library(bootmatch)

# Compare standardized test measures across groups
compare_matched_groups <- function(input_df, group_col, metrics) {
  # Create a results dataframe
  desc <- matrix(NA, nrow = length(unique(metrics)), ncol = 1)
  desc <- as.data.frame(desc)
  colnames(desc) <- c("Score")
  desc$Score <- c(unique(metrics))

  # Calculate descriptive statistics and compare diagnostic groups
  for (i in seq_along(metrics)) {
    # Subset to one metric and subset groups
    test_scores <- dplyr::select(input_df, group_col, metrics[i])
    test_scores <- na.omit(test_scores)
    asd_scores <- test_scores %>% filter(.data[[group_col]] == "ASD")
    td_scores <- test_scores %>% filter(.data[[group_col]] == "TD")

    # Calculate descriptive statistics for ASD group
    desc$ASD_n[i] <- nrow(asd_scores)
    desc$ASD_mean[i] <- round(mean(asd_scores[[metrics[i]]]), 2)
    desc$ASD_sd[i] <- round(sd(asd_scores[[metrics[i]]]), 2)
    desc$ASD_min[i] <- round(min(asd_scores[[metrics[i]]]), 2)
    desc$ASD_max[i] <- round(max(asd_scores[[metrics[i]]]), 2)

    # Calculate descriptive statistics for TD group
    desc$TD_n[i] <- nrow(td_scores)
    desc$TD_mean[i] <- round(mean(td_scores[[metrics[i]]]), 2)
    desc$TD_sd[i] <- round(sd(td_scores[[metrics[i]]]), 2)
    desc$TD_min[i] <- round(min(td_scores[[metrics[i]]]), 2)
    desc$TD_max[i] <- round(max(td_scores[[metrics[i]]]), 2)

    # Calculate effect size (Cohen's d)
    d <- effsize::cohen.d(
      asd_scores[[metrics[i]]],
      td_scores[[metrics[i]]],
      paired = FALSE
    )
    desc$d[i] <- abs(round(d$estimate, digits = 2))

    # Calculate variance ratio
    desc$v[i] <- round(
      var(asd_scores[[metrics[i]]]) / var(td_scores[[metrics[i]]]),
      digits = 2
    )

    # Test for normality
    test1 <- try(
      shapiro.test(
        as.numeric(asd_scores[[metrics[i]]])
      ),
      silent = TRUE
    )
    test2 <- try(
      shapiro.test(
        as.numeric(td_scores[[metrics[i]]])
      ),
      silent = TRUE
    )

    # If either distribution is non-normal, use a wilcox test, else use a t-test
    if (test1$p.value < 0.05 || test2$p.value < 0.05) {
      test <- try(
        wilcox.test(
          asd_scores[[2]],
          td_scores[[2]],
          paired = FALSE
        ),
        silent = TRUE
      )
      desc$p[i] <- ifelse(is(test, "try-error"), NA,
        round(as.numeric(test$p.value), 3)
      )
    } else {
      test <- try(
        t.test(
          asd_scores[[2]],
          td_scores[[2]],
          paired = FALSE
        ),
        silent = TRUE
      )
      desc$p[i] <- ifelse(is(test, "try-error"), NA,
        round(as.numeric(test$p.value), 3)
      )
    }
  }

  # Improve table legibility
  desc <- na.omit(desc)
  desc$p <- ifelse(desc$p < 0.001, "< 0.001", desc$p)
  colnames(desc)[1:14] <- c(
    "Measure", "N", "Mean", "SD", "Min", "Max",
    "N", "Mean", "SD", "Min", "Max",
    "Cohen's d", "Var Ratio", "p-value"
  )

  # Render table
  kable(desc, format = "latex", booktabs = TRUE, row.names = FALSE) %>%
    add_header_above(c(
      " " = 1,
      "ASD Group" = 5,
      "NT Group" = 5,
      "Group Differences" = 3
    ))
}

# Define the raw GitHub URL of the Excel file
excel_url <- "https://github.com/tracyreuter/bootstrap-match-groups/raw/refs/heads/main/subjectlog.xlsx" # nolint

# Create a temporary file to download the Excel file
temp_file <- tempfile(fileext = ".xlsx")

# Download to the temporary file
download.file(excel_url, destfile = temp_file, mode = "wb")

# Load data
all_subjects <- read_excel(temp_file)

# Select metrics and rename group column
all_subjects <- dplyr::select(
  all_subjects, subject, group, sex, age,
  PLS.AC.Raw, PLS.AC.AE, PLS.EC.Raw, PLS.EC.AE,
  Mul.VR.Raw, Mul.VR.AE, Mul.Ratio.IQ
) %>%
  dplyr::rename(diagnostic_group = group)

# Convert all metrics to numeric
all_subjects <- all_subjects %>% mutate_at(c(4:11), as.numeric)

# Match groups
foo <- bootmatch::boot_match_univariate(
  data = all_subjects,
  y = diagnostic_group,
  x = PLS.AC.Raw,
  id = subject,
  caliper = 5,
  boot = 100
)

# Subset to matched subjects
matched_subjects <- all_subjects %>%
  filter(subject %in% unique(foo$Matching_MatchID))

# Call function
compare_matched_groups(
  input_df = matched_subjects,
  group_col = "diagnostic_group",
  metrics = colnames(matched_subjects[4:11])
)