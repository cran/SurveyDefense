#' Fraud Detection Analysis Tool 1
#'
#' This function analyzes survey data based on up to 5 Fraud Detection Questions and generates results in Word and HTML formats.
#'
#' @param output_dir Path specifying where the Word and HTML files will be saved.
#' @param data The data frame containing all the survey data.
#' @param FraudList A character vector of up to 5 Fraud Detection Questions.
#' @param correct_answers A numeric vector representing correct answers for each question. Default is \code{c(0, 0, 0, 0, 0)}.
#' @param ... Survey questions to be analyzed.
#' @return A flextable object with the fraud detection analysis results. The results include summary statistics and metrics comparing responses from reliable and fraudulent participants.
#' @importFrom flextable flextable
#' @importFrom flextable set_header_labels
#' @importFrom flextable colformat_double
#' @importFrom flextable merge_at
#' @importFrom flextable align
#' @importFrom flextable theme_vanilla
#' @importFrom flextable bold
#' @importFrom flextable save_as_docx
#' @importFrom flextable save_as_html
#' @importFrom utils write.csv
#' @import dplyr
#' @examples
#' if (requireNamespace("flextable", quietly = TRUE) && requireNamespace("officer", quietly = TRUE)) {
#'   library(flextable)
#'   library(officer)
#'
#'   # Example data for fraud detection analysis
#'   Q1 <- c(4, 5, 3, 2, 5, 2)
#'   Q2 <- c(3, 4, 2, 5, 4, 3)
#'   Q3 <- c(5, 4, 3, 5, 4, 5)
#'   Q4 <- c(1, 2, 3, 4, 5, 2)
#'   Q5 <- c(5, 2, 2, 1, 4, 1)
#'   Q6 <- c(5, 2, 3, 5, 1, 2)
#'   Q7 <- c(5, 2, 4, 5, 3, 4)
#'
#'   Fraud1 <- c(0, 1, 0, 0, 0, 0)
#'   Fraud2 <- c(0, 0, 0, 0, 0, 0)
#'   Fraud3 <- c(0, 1, 0, 0, 0, 0)
#'   Fraud4 <- c(0, 0, 1, 0, 0, 1)
#'   Fraud5 <- c(0, 0, 0, 1, 1, 1)
#'
#'   Test_Data_Fraud <- data.frame(Q1, Q2, Q3, Q4, Q5, Q6, Q7, Fraud1, Fraud2, Fraud3, Fraud4, Fraud5)
#'
#'   temp_dir <- tempdir()
#'
#'   FraudDetec1(
#'     output_dir = temp_dir,
#'     data = Test_Data_Fraud,
#'     FraudList = c("Fraud1", "Fraud2", "Fraud3", "Fraud4", "Fraud5"),
#'     correct_answers = c(0, 0, 0, 0, 0),
#'     Q1, Q2, Q3, Q4, Q5, Q6, Q7
#'   )
#' }

#' @export
FraudDetec1 <- function(output_dir, data, FraudList, correct_answers = c(0, 0, 0, 0, 0), ...) {
  # Evaluate question_vars names in the data context
  question_vars <- as.character(substitute(list(...))[-1])

  # Combine FraudList and question_vars for NA removal
  all_vars <- c(FraudList, question_vars)

  # Remove rows with NAs in any of the relevant columns
  data_cleaned <- data %>%
    filter(if_all(all_of(all_vars), ~ !is.na(.)))

  # Extract the cleaned columns from the data
  FraudList_cleaned <- lapply(FraudList, function(x) data_cleaned[[x]])
  questions_cleaned <- lapply(question_vars, function(x) data_cleaned[[x]])

  # Define a function to convert values to percentages
  to_percentage <- function(values) {
    return(paste0(round(values * 100, 2), "%"))
  }

  # Check if at least one question was passed
  if (length(questions_cleaned) == 0) {
    stop("At least one question must be passed.")
  }

  # Check if FraudList and correct_answers contain the same number of elements
  if (length(FraudList) != length(correct_answers)) {
    stop("FraudList and correct_answers must have the same length.")
  }

  # Calculate whether each fraud variable was answered incorrectly
  fraud_mismatch <- mapply(function(fraud, correct) {
    return(fraud != correct)
  }, FraudList_cleaned, correct_answers)

  # Calculate the FraudCheck variable: 1 if at least 1 fraud variable was answered incorrectly, else 0
  FraudCheck <- rowSums(fraud_mismatch) >= 1
  FraudCheck <- as.integer(FraudCheck)  # Convert TRUE/FALSE to 1/0

  # Calculate the share of cases where at least 1 fraud variable was  answered correctly
  share_of_reliable_participations <- mean(FraudCheck == 0)

  # Calculate the Full Sample Mean for each question
  full_sample_mean <- sapply(questions_cleaned, mean, na.rm = TRUE)

  # Calculate the Cleaned Sample Mean for each question
  cleaned_sample_indices <- which(FraudCheck == 0)
  cleaned_sample_mean <- sapply(questions_cleaned, function(q) mean(q[cleaned_sample_indices], na.rm = TRUE))

  # Calculate the Fraud Sample Mean for each question
  fraud_sample_indices <- which(FraudCheck == 1)
  fraud_sample_mean <- sapply(questions_cleaned, function(q) mean(q[fraud_sample_indices], na.rm = TRUE))

  # Calculate the difference and relative value _fraud_cleaned
  difference_fraud_cleaned <- fraud_sample_mean - cleaned_sample_mean
  relative_difference_fraud_cleaned <- cleaned_sample_mean / fraud_sample_mean - 1

  # Calculate the mean of the difference and relative value across all questions _fraud_cleaned
  mean_difference_fraud_cleaned <- mean(abs(difference_fraud_cleaned))
  relative_mean_difference_fraud_cleaned <- mean(abs(relative_difference_fraud_cleaned))

  # Calculate the difference and relative value _full_cleaned
  difference_full_cleaned <- full_sample_mean - cleaned_sample_mean
  relative_difference_full_cleaned <- cleaned_sample_mean / full_sample_mean - 1

  # Calculate the mean of the difference and relative value across all questions _full_cleaned
  mean_difference_full_cleaned <- mean(abs(difference_full_cleaned))
  relative_mean_difference_full_cleaned <- mean(abs(relative_difference_full_cleaned))

  # Calculate the number of participants for each group
  response_count_full <- length(unique(unlist(lapply(questions_cleaned, function(q) which(!is.na(q))))))
  response_count_fraud <- sum(FraudCheck == 1)
  response_count_cleaned <- sum(FraudCheck == 0)

  # Create an empty row for separation
  empty_row1 <- data.frame(
    Explanation = "Metrics per Survey Question",
    Question = NA,
    Full_Sample_Mean = NA,
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Create a table with the calculated values
  results_table <- data.frame(
    Explanation = NA,
    Question = question_vars,  # Use the variable names as question labels
    Full_Sample_Mean = round(full_sample_mean, 2),
    cleaned_sample_mean = cleaned_sample_mean,
    fraud_sample_mean = fraud_sample_mean,
    difference_fraud_cleaned = difference_fraud_cleaned,
    difference_full_cleaned = difference_full_cleaned,
    relative_difference_fraud_cleaned = to_percentage(relative_difference_fraud_cleaned),
    relative_difference_full_cleaned = to_percentage(relative_difference_full_cleaned)
  )

  # Add the number of responses
  response_row <- data.frame(
    Explanation = "Number of Participants",
    Question = NA,
    Full_Sample_Mean = response_count_full,
    cleaned_sample_mean = response_count_cleaned,
    fraud_sample_mean = response_count_fraud,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Add Share of Reliable Participants (SRP)
  share_row <- data.frame(
    Explanation = "Share of Reliable Participants (SRP)",
    Question = NA,
    Full_Sample_Mean = to_percentage(share_of_reliable_participations),
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Add the Fraud Detection Threshold row
  threshold_row <- data.frame(
    Explanation = "Fraud Detection Threshold",
    Question = "At least 1 Fraud Detection Questions answered wrong.",
    Full_Sample_Mean = NA,
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Create an empty row for separation
  empty_row2 <- data.frame(
    Explanation = "Analysis Summary",
    Question = NA,
    Full_Sample_Mean = NA,
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Create a row for "Aggregated Metrics"
  aggregated_metrics_row <- data.frame(
    Explanation = "Aggregated Metrics",
    Question = NA,
    Full_Sample_Mean = NA,
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Add the relative average Difference between Fraud and Cleaned Sample (RelFC)
  mean_row1 <- data.frame(
    Explanation = "\u00d8 Relative Difference between Fraud and Cleaned Sample",
    Question = NA,
    Full_Sample_Mean = to_percentage(relative_mean_difference_fraud_cleaned),
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Add the relative average Difference between Full and Cleaned Sample (RelFC)
  mean_row2 <- data.frame(
    Explanation = "\u00d8 Relative Difference between Full and Cleaned Sample",
    Question = NA,
    Full_Sample_Mean = to_percentage(relative_mean_difference_full_cleaned),
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Combine all results into a table
  results_table <- rbind(
    empty_row2, response_row, share_row, threshold_row, empty_row1, results_table, aggregated_metrics_row, mean_row1, mean_row2
  )

  # Format the table with flextable
  ft <- flextable(results_table)

  # Set the table headers
  ft <- set_header_labels(
    ft,
    Question = "",
    Full_Sample_Mean = "Full Sample Mean",
    fraud_sample_mean = "Fraud Sample Mean",
    cleaned_sample_mean = "Cleaned Sample Mean",
    difference_fraud_cleaned = "Difference between Fraud and Cleaned Sample",
    difference_full_cleaned = "Difference between Full and Cleaned Sample",
    relative_difference_fraud_cleaned = "Relative Difference between Fraud and Cleaned Sample",
    relative_difference_full_cleaned = "Relative Difference between Full and Cleaned Sample"
  )

  # Format all numerical values to two decimal places
  ft <- colformat_double(ft, j = c("Full_Sample_Mean", "fraud_sample_mean", "cleaned_sample_mean", "difference_fraud_cleaned", "difference_full_cleaned"), digits = 2)

  # Merge and left-align only for the "Analysis Summary" row and make it bold
  ft <- merge_at(ft, i = which(results_table$Explanation == "Analysis Summary"), j = 1:ncol(results_table))
  ft <- align(ft, i = which(results_table$Explanation == "Analysis Summary"), align = "left", part = "body")
  ft <- bold(ft, i = which(results_table$Explanation == "Analysis Summary"), bold = TRUE, part = "body")

  # Merge and left-align only for the "Aggregated Metrics" row and make it bold
  ft <- merge_at(ft, i = which(results_table$Explanation == "Aggregated Metrics"), j = 1:ncol(results_table))
  ft <- align(ft, i = which(results_table$Explanation == "Aggregated Metrics"), align = "left", part = "body")
  ft <- bold(ft, i = which(results_table$Explanation == "Aggregated Metrics"), bold = TRUE, part = "body")

  # Merge and left-align only for the "Metrics per Survey Question" row and make it bold
  ft <- merge_at(ft, i = which(results_table$Explanation == "Metrics per Survey Question"), j = 1:ncol(results_table))
  ft <- align(ft, i = which(results_table$Explanation == "Metrics per Survey Question"), align = "left", part = "body")
  ft <- bold(ft, i = which(results_table$Explanation == "Metrics per Survey Question"), bold = TRUE, part = "body")

  # Merge and left-align only for the "Number of Participants" row
  ft <- merge_at(ft, i = which(results_table$Explanation == "Number of Participants"), j = 1:2)
  ft <- align(ft, i = which(results_table$Explanation == "Number of Participants"), align = "left", part = "body")

  # Merge and left-align only for the "Share of Reliable Participants (SRP)" row
  ft <- merge_at(ft, i = which(results_table$Explanation == "Share of Reliable Participants (SRP)"), j = 1:2)
  ft <- align(ft, i = which(results_table$Explanation == "Share of Reliable Participants (SRP)"), align = "left", part = "body")

  # Merge and left-align only for the "Fraud Detection Threshold" row, spanning the text across the remaining columns
  ft <- merge_at(ft, i = which(results_table$Explanation == "Fraud Detection Threshold"), j = 2:ncol(results_table))
  ft <- align(ft, i = which(results_table$Explanation == "Fraud Detection Threshold"), align = "left", part = "body")

  # Choose a theme for the table appearance
  ft <- theme_vanilla(ft)

  # Save the table as a Word file
  doc_file_path <- file.path(output_dir, "FraudDetec1_results.docx")
  save_as_docx(ft, path = doc_file_path)

  # Save the table as an HTML file
  html_file_path <- file.path(output_dir, "FraudDetec1_results.html")
  save_as_html(ft, path = html_file_path)

  # Display message to user
  message("Note: Output has been saved as Microsoft Word (.docx) and HTML (.html) files under ", output_dir)
  message("Note: The cleaned data set has been saved as .CSV and as RData files under ", output_dir)
  message("Note: Fraud Detection Specification: At least 1 Fraud Detection Questions answered wrong")
  message("Note: Missing values for all variables used in the analysis were cleaned prior to conducting the analysis.")

  # Save the cleaned sample as a CSV file
  cleaned_sample <- data_cleaned[cleaned_sample_indices, !(names(data_cleaned) %in% c("Fraud1", "Fraud2", "Fraud3", "Fraud4", "Fraud5"))]
  csv_file_path <- file.path(output_dir, "FraudDetec1_cleaned_sample.csv")
  write.csv(cleaned_sample, csv_file_path, row.names = FALSE)

  # Save the cleaned sample as an RData file
  rdata_file_path <- file.path(output_dir, "FraudDetec1_cleaned_sample.RData")
  save(cleaned_sample, file = rdata_file_path)

  # Return the flextable
  return(ft)
}



#' Fraud Detection Analysis Tool 2
#'
#' This function analyzes survey data using up to 5 Fraud Detection Questions and generates a report in Word and HTML formats.
#'
#' @param output_dir Path specifying where the Word and HTML files will be saved.
#' @param data The data frame containing all the survey data.
#' @param FraudList A character vector of up to 5 Fraud Detection Questions.
#' @param correct_answers A numeric vector representing correct answers for each question. Default is \code{c(0, 0, 0, 0, 0)}.
#' @param ... Survey questions to be analyzed.
#' @return A flextable object with the fraud detection analysis results, including summary statistics for the overall sample and identified fraudulent responses.
#' @importFrom flextable flextable
#' @importFrom flextable set_header_labels
#' @importFrom flextable colformat_double
#' @importFrom flextable merge_at
#' @importFrom flextable align
#' @importFrom flextable theme_vanilla
#' @importFrom flextable save_as_docx
#' @importFrom flextable save_as_html
#' @importFrom utils write.csv
#' @import dplyr
#' @examples
#' if (requireNamespace("flextable", quietly = TRUE) && requireNamespace("officer", quietly = TRUE)) {
#'   library(flextable)
#'   library(officer)
#'
#'   # Example data for fraud detection analysis
#'   Q1 <- c(4, 5, 3, 2, 5, 2)
#'   Q2 <- c(3, 4, 2, 5, 4, 3)
#'   Q3 <- c(5, 4, 3, 5, 4, 5)
#'   Q4 <- c(1, 2, 3, 4, 5, 2)
#'   Q5 <- c(5, 2, 2, 1, 4, 1)
#'   Q6 <- c(5, 2, 3, 5, 1, 2)
#'   Q7 <- c(5, 2, 4, 5, 3, 4)
#'
#'   Fraud1 <- c(0, 1, 0, 0, 0, 0)
#'   Fraud2 <- c(0, 0, 0, 0, 0, 0)
#'   Fraud3 <- c(0, 1, 0, 0, 0, 0)
#'   Fraud4 <- c(0, 0, 1, 0, 0, 1)
#'   Fraud5 <- c(0, 0, 0, 1, 1, 1)
#'
#'   Test_Data_Fraud <- data.frame(Q1, Q2, Q3, Q4, Q5, Q6, Q7, Fraud1, Fraud2, Fraud3, Fraud4, Fraud5)
#'
#'   temp_dir <- tempdir()
#'
#'   FraudDetec2(
#'     output_dir = temp_dir,
#'     data = Test_Data_Fraud,
#'     FraudList = c("Fraud1", "Fraud2", "Fraud3", "Fraud4", "Fraud5"),
#'     correct_answers = c(0, 0, 0, 0, 0),
#'     Q1, Q2, Q3, Q4, Q5, Q6, Q7
#'   )
#' }

#' @export
FraudDetec2 <- function(output_dir, data, FraudList, correct_answers = c(0, 0, 0, 0, 0), ...) {
  # Evaluate question_vars names in the data context
  question_vars <- as.character(substitute(list(...))[-1])

  # Combine FraudList and question_vars for NA removal
  all_vars <- c(FraudList, question_vars)

  # Remove rows with NAs in any of the relevant columns
  data_cleaned <- data %>%
    filter(if_all(all_of(all_vars), ~ !is.na(.)))

  # Extract the cleaned columns from the data
  FraudList_cleaned <- lapply(FraudList, function(x) data_cleaned[[x]])
  questions_cleaned <- lapply(question_vars, function(x) data_cleaned[[x]])

  # Define a function to convert values to percentages
  to_percentage <- function(values) {
    return(paste0(round(values * 100, 2), "%"))
  }

  # Check if at least one question was passed
  if (length(questions_cleaned) == 0) {
    stop("At least one question must be passed.")
  }

  # Check if FraudList and correct_answers contain the same number of elements
  if (length(FraudList) != length(correct_answers)) {
    stop("FraudList and correct_answers must have the same length.")
  }

  # Calculate whether each fraud variable was answered incorrectly
  fraud_mismatch <- mapply(function(fraud, correct) {
    return(fraud != correct)
  }, FraudList_cleaned, correct_answers)

  # Calculate the FraudCheck variable: 1 if at least 2 fraud variable was answered incorrectly, else 0
  FraudCheck <- rowSums(fraud_mismatch) >= 2
  FraudCheck <- as.integer(FraudCheck)  # Convert TRUE/FALSE to 1/0

  # Calculate the share of cases where at least 1 fraud variable was not answered incorrectly
  share_of_reliable_participations <- mean(FraudCheck == 0)

  # Calculate the Full Sample Mean for each question
  full_sample_mean <- sapply(questions_cleaned, mean, na.rm = TRUE)

  # Calculate the Cleaned Sample Mean for each question
  cleaned_sample_indices <- which(FraudCheck == 0)
  cleaned_sample_mean <- sapply(questions_cleaned, function(q) mean(q[cleaned_sample_indices], na.rm = TRUE))

  # Calculate the Fraud Sample Mean for each question
  fraud_sample_indices <- which(FraudCheck == 1)
  fraud_sample_mean <- sapply(questions_cleaned, function(q) mean(q[fraud_sample_indices], na.rm = TRUE))

  # Calculate the difference and relative value _fraud_cleaned
  difference_fraud_cleaned <- fraud_sample_mean - cleaned_sample_mean
  relative_difference_fraud_cleaned <- cleaned_sample_mean / fraud_sample_mean - 1

  # Calculate the mean of the difference and relative value across all questions _fraud_cleaned
  mean_difference_fraud_cleaned <- mean(abs(difference_fraud_cleaned))
  relative_mean_difference_fraud_cleaned <- mean(abs(relative_difference_fraud_cleaned))

  # Calculate the difference and relative value _full_cleaned
  difference_full_cleaned <- full_sample_mean - cleaned_sample_mean
  relative_difference_full_cleaned <- cleaned_sample_mean / full_sample_mean - 1

  # Calculate the mean of the difference and relative value across all questions _full_cleaned
  mean_difference_full_cleaned <- mean(abs(difference_full_cleaned))
  relative_mean_difference_full_cleaned <- mean(abs(relative_difference_full_cleaned))

  # Calculate the number of participants for each group
  response_count_full <- length(unique(unlist(lapply(questions_cleaned, function(q) which(!is.na(q))))))
  response_count_fraud <- sum(FraudCheck == 1)
  response_count_cleaned <- sum(FraudCheck == 0)

  # Create an empty row for separation
  empty_row1 <- data.frame(
    Explanation = "Metrics per Survey Question",
    Question = NA,
    Full_Sample_Mean = NA,
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Create a table with the calculated values
  results_table <- data.frame(
    Explanation = NA,
    Question = question_vars,  # Use the variable names as question labels
    Full_Sample_Mean = round(full_sample_mean, 2),
    cleaned_sample_mean = cleaned_sample_mean,
    fraud_sample_mean = fraud_sample_mean,
    difference_fraud_cleaned = difference_fraud_cleaned,
    difference_full_cleaned = difference_full_cleaned,
    relative_difference_fraud_cleaned = to_percentage(relative_difference_fraud_cleaned),
    relative_difference_full_cleaned = to_percentage(relative_difference_full_cleaned)
  )

  # Add the number of responses
  response_row <- data.frame(
    Explanation = "Number of Participants",
    Question = NA,
    Full_Sample_Mean = response_count_full,
    cleaned_sample_mean = response_count_cleaned,
    fraud_sample_mean = response_count_fraud,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Add Share of Reliable Participants (SRP)
  share_row <- data.frame(
    Explanation = "Share of Reliable Participants (SRP)",
    Question = NA,
    Full_Sample_Mean = to_percentage(share_of_reliable_participations),
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Add the Fraud Detection Threshold row
  threshold_row <- data.frame(
    Explanation = "Fraud Detection Threshold",
    Question = "At least 2 Fraud Detection Questions answered wrong.",
    Full_Sample_Mean = NA,
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Create an empty row for separation
  empty_row2 <- data.frame(
    Explanation = "Analysis Summary",
    Question = NA,
    Full_Sample_Mean = NA,
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Create a row for "Aggregated Metrics"
  aggregated_metrics_row <- data.frame(
    Explanation = "Aggregated Metrics",
    Question = NA,
    Full_Sample_Mean = NA,
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Add the relative average Difference between Fraud and Cleaned Sample (RelFC)
  mean_row1 <- data.frame(
    Explanation = "\u00d8 Relative Difference between Fraud and Cleaned Sample",
    Question = NA,
    Full_Sample_Mean = to_percentage(relative_mean_difference_fraud_cleaned),
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Add the relative average Difference between Full and Cleaned Sample (RelFC)
  mean_row2 <- data.frame(
    Explanation = "\u00d8 Relative Difference between Full and Cleaned Sample",
    Question = NA,
    Full_Sample_Mean = to_percentage(relative_mean_difference_full_cleaned),
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Combine all results into a table
  results_table <- rbind(
    empty_row2, response_row, share_row, threshold_row, empty_row1, results_table, aggregated_metrics_row, mean_row1, mean_row2
  )

  # Format the table with flextable
  ft <- flextable(results_table)

  # Set the table headers
  ft <- set_header_labels(
    ft,
    Question = "",
    Full_Sample_Mean = "Full Sample Mean",
    fraud_sample_mean = "Fraud Sample Mean",
    cleaned_sample_mean = "Cleaned Sample Mean",
    difference_fraud_cleaned = "Difference between Fraud and Cleaned Sample",
    difference_full_cleaned = "Difference between Full and Cleaned Sample",
    relative_difference_fraud_cleaned = "Relative Difference between Fraud and Cleaned Sample",
    relative_difference_full_cleaned = "Relative Difference between Full and Cleaned Sample"
  )

  # Format all numerical values to two decimal places
  ft <- colformat_double(ft, j = c("Full_Sample_Mean", "fraud_sample_mean", "cleaned_sample_mean", "difference_fraud_cleaned", "difference_full_cleaned"), digits = 2)

  # Merge and left-align only for the "Analysis Summary" row and make it bold
  ft <- merge_at(ft, i = which(results_table$Explanation == "Analysis Summary"), j = 1:ncol(results_table))
  ft <- align(ft, i = which(results_table$Explanation == "Analysis Summary"), align = "left", part = "body")
  ft <- bold(ft, i = which(results_table$Explanation == "Analysis Summary"), bold = TRUE, part = "body")

  # Merge and left-align only for the "Aggregated Metrics" row and make it bold
  ft <- merge_at(ft, i = which(results_table$Explanation == "Aggregated Metrics"), j = 1:ncol(results_table))
  ft <- align(ft, i = which(results_table$Explanation == "Aggregated Metrics"), align = "left", part = "body")
  ft <- bold(ft, i = which(results_table$Explanation == "Aggregated Metrics"), bold = TRUE, part = "body")

  # Merge and left-align only for the "Metrics per Survey Question" row and make it bold
  ft <- merge_at(ft, i = which(results_table$Explanation == "Metrics per Survey Question"), j = 1:ncol(results_table))
  ft <- align(ft, i = which(results_table$Explanation == "Metrics per Survey Question"), align = "left", part = "body")
  ft <- bold(ft, i = which(results_table$Explanation == "Metrics per Survey Question"), bold = TRUE, part = "body")

  # Merge and left-align only for the "Number of Participants" row
  ft <- merge_at(ft, i = which(results_table$Explanation == "Number of Participants"), j = 1:2)
  ft <- align(ft, i = which(results_table$Explanation == "Number of Participants"), align = "left", part = "body")

  # Merge and left-align only for the "Share of Reliable Participants (SRP)" row
  ft <- merge_at(ft, i = which(results_table$Explanation == "Share of Reliable Participants (SRP)"), j = 1:2)
  ft <- align(ft, i = which(results_table$Explanation == "Share of Reliable Participants (SRP)"), align = "left", part = "body")

  # Merge and left-align only for the "Fraud Detection Threshold" row, spanning the text across the remaining columns
  ft <- merge_at(ft, i = which(results_table$Explanation == "Fraud Detection Threshold"), j = 2:ncol(results_table))
  ft <- align(ft, i = which(results_table$Explanation == "Fraud Detection Threshold"), align = "left", part = "body")

  # Choose a theme for the table appearance
  ft <- theme_vanilla(ft)

  # Save the table as a Word file
  doc_file_path <- file.path(output_dir, "FraudDetec2_results.docx")
  save_as_docx(ft, path = doc_file_path)

  # Save the table as an HTML file
  html_file_path <- file.path(output_dir, "FraudDetec2_results.html")
  save_as_html(ft, path = html_file_path)

  # Display message to user
  message("Note: Output has been saved as Microsoft Word (.docx) and HTML (.html) files under ", output_dir)
  message("Note: The cleaned data set has been saved as .CSV and as RData files under ", output_dir)
  message("Note: Fraud Detection Specification: At least 2 Fraud Detection Questions answered wrong")
  message("Note: Missing values for all variables used in the analysis were cleaned prior to conducting the analysis.")

  # Save the cleaned sample as a CSV file
  cleaned_sample <- data_cleaned[cleaned_sample_indices, !(names(data_cleaned) %in% c("Fraud1", "Fraud2", "Fraud3", "Fraud4", "Fraud5"))]
  csv_file_path <- file.path(output_dir, "FraudDetec2_cleaned_sample.csv")
  write.csv(cleaned_sample, csv_file_path, row.names = FALSE)

  # Save the cleaned sample as an RData file
  rdata_file_path <- file.path(output_dir, "FraudDetec2_cleaned_sample.RData")
  save(cleaned_sample, file = rdata_file_path)

  # Return the flextable
  return(ft)
}


#' Fraud Detection Analysis Tool 3
#'
#' @param output_dir Path specifying where the Word and HTML files will be saved.
#' @param data The data frame containing all the survey data.
#' @param FraudList A character vector of up to 5 Fraud Detection Questions.
#' @param correct_answers A numeric vector representing correct answers for each question. Default is \code{c(0, 0, 0, 0, 0)}.
#' @param ... Survey questions to be analyzed.
#' @return A flextable object with the results.
#' @importFrom flextable flextable
#' @importFrom flextable set_header_labels
#' @importFrom flextable colformat_double
#' @importFrom flextable merge_at
#' @importFrom flextable align
#' @importFrom flextable theme_vanilla
#' @importFrom flextable bold
#' @importFrom flextable save_as_docx
#' @importFrom flextable save_as_html
#' @importFrom utils write.csv
#' @import dplyr
#' @examples
#' if (requireNamespace("flextable", quietly = TRUE) && requireNamespace("officer", quietly = TRUE)) {
#'   library(flextable)
#'   library(officer)
#'
#'   # Example data for fraud detection analysis
#'   Q1 <- c(4, 5, 3, 2, 5, 2)
#'   Q2 <- c(3, 4, 2, 5, 4, 3)
#'   Q3 <- c(5, 4, 3, 5, 4, 5)
#'   Q4 <- c(1, 2, 3, 4, 5, 2)
#'   Q5 <- c(5, 2, 2, 1, 4, 1)
#'   Q6 <- c(5, 2, 3, 5, 1, 2)
#'   Q7 <- c(5, 2, 4, 5, 3, 4)
#'
#'   Fraud1 <- c(0, 1, 0, 0, 0, 0)
#'   Fraud2 <- c(0, 0, 0, 0, 0, 0)
#'   Fraud3 <- c(0, 1, 0, 0, 0, 0)
#'   Fraud4 <- c(0, 0, 1, 0, 0, 1)
#'   Fraud5 <- c(0, 0, 0, 1, 1, 1)
#'
#'   Test_Data_Fraud <- data.frame(Q1, Q2, Q3, Q4, Q5, Q6, Q7, Fraud1, Fraud2, Fraud3, Fraud4, Fraud5)
#'
#'   temp_dir <- tempdir()
#'
#'   FraudDetec3(
#'     output_dir = temp_dir,
#'     data = Test_Data_Fraud,
#'     FraudList = c("Fraud1", "Fraud2", "Fraud3", "Fraud4", "Fraud5"),
#'     correct_answers = c(0, 0, 0, 0, 0),
#'     Q1, Q2, Q3, Q4, Q5, Q6, Q7
#'   )
#' }
#' @export
FraudDetec3 <- function(output_dir, data, FraudList, correct_answers = c(0, 0, 0, 0, 0), ...) {
  # Evaluate question_vars names in the data context
  question_vars <- as.character(substitute(list(...))[-1])

  # Combine FraudList and question_vars for NA removal
  all_vars <- c(FraudList, question_vars)

  # Remove rows with NAs in any of the relevant columns
  data_cleaned <- data %>%
    filter(if_all(all_of(all_vars), ~ !is.na(.)))

  # Extract the cleaned columns from the data
  FraudList_cleaned <- lapply(FraudList, function(x) data_cleaned[[x]])
  questions_cleaned <- lapply(question_vars, function(x) data_cleaned[[x]])

  # Define a function to convert values to percentages
  to_percentage <- function(values) {
    return(paste0(round(values * 100, 2), "%"))
  }

  # Check if at least one question was passed
  if (length(questions_cleaned) == 0) {
    stop("At least one question must be passed.")
  }

  # Check if FraudList and correct_answers contain the same number of elements
  if (length(FraudList) != length(correct_answers)) {
    stop("FraudList and correct_answers must have the same length.")
  }

  # Calculate whether each fraud variable was answered incorrectly
  fraud_mismatch <- mapply(function(fraud, correct) {
    return(fraud != correct)
  }, FraudList_cleaned, correct_answers)

  # Calculate the FraudCheck variable: 1 if at least 3 fraud variable was answered incorrectly, else 0
  FraudCheck <- rowSums(fraud_mismatch) >= 3
  FraudCheck <- as.integer(FraudCheck)  # Convert TRUE/FALSE to 1/0

  # Calculate the share of cases where at least 3 fraud variable was  answered correctly
  share_of_reliable_participations <- mean(FraudCheck == 0)

  # Calculate the Full Sample Mean for each question
  full_sample_mean <- sapply(questions_cleaned, mean, na.rm = TRUE)

  # Calculate the Cleaned Sample Mean for each question
  cleaned_sample_indices <- which(FraudCheck == 0)
  cleaned_sample_mean <- sapply(questions_cleaned, function(q) mean(q[cleaned_sample_indices], na.rm = TRUE))

  # Calculate the Fraud Sample Mean for each question
  fraud_sample_indices <- which(FraudCheck == 1)
  fraud_sample_mean <- sapply(questions_cleaned, function(q) mean(q[fraud_sample_indices], na.rm = TRUE))

  # Calculate the difference and relative value _fraud_cleaned
  difference_fraud_cleaned <- fraud_sample_mean - cleaned_sample_mean
  relative_difference_fraud_cleaned <- cleaned_sample_mean / fraud_sample_mean - 1

  # Calculate the mean of the difference and relative value across all questions _fraud_cleaned
  mean_difference_fraud_cleaned <- mean(abs(difference_fraud_cleaned))
  relative_mean_difference_fraud_cleaned <- mean(abs(relative_difference_fraud_cleaned))

  # Calculate the difference and relative value _full_cleaned
  difference_full_cleaned <- full_sample_mean - cleaned_sample_mean
  relative_difference_full_cleaned <- cleaned_sample_mean / full_sample_mean - 1

  # Calculate the mean of the difference and relative value across all questions _full_cleaned
  mean_difference_full_cleaned <- mean(abs(difference_full_cleaned))
  relative_mean_difference_full_cleaned <- mean(abs(relative_difference_full_cleaned))

  # Calculate the number of participants for each group
  response_count_full <- length(unique(unlist(lapply(questions_cleaned, function(q) which(!is.na(q))))))
  response_count_fraud <- sum(FraudCheck == 1)
  response_count_cleaned <- sum(FraudCheck == 0)

  # Create an empty row for separation
  empty_row1 <- data.frame(
    Explanation = "Metrics per Survey Question",
    Question = NA,
    Full_Sample_Mean = NA,
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Create a table with the calculated values
  results_table <- data.frame(
    Explanation = NA,
    Question = question_vars,  # Use the variable names as question labels
    Full_Sample_Mean = round(full_sample_mean, 2),
    cleaned_sample_mean = cleaned_sample_mean,
    fraud_sample_mean = fraud_sample_mean,
    difference_fraud_cleaned = difference_fraud_cleaned,
    difference_full_cleaned = difference_full_cleaned,
    relative_difference_fraud_cleaned = to_percentage(relative_difference_fraud_cleaned),
    relative_difference_full_cleaned = to_percentage(relative_difference_full_cleaned)
  )

  # Add the number of responses
  response_row <- data.frame(
    Explanation = "Number of Participants",
    Question = NA,
    Full_Sample_Mean = response_count_full,
    cleaned_sample_mean = response_count_cleaned,
    fraud_sample_mean = response_count_fraud,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Add Share of Reliable Participants (SRP)
  share_row <- data.frame(
    Explanation = "Share of Reliable Participants (SRP)",
    Question = NA,
    Full_Sample_Mean = to_percentage(share_of_reliable_participations),
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Add the Fraud Detection Threshold row
  threshold_row <- data.frame(
    Explanation = "Fraud Detection Threshold",
    Question = "At least 3 Fraud Detection Questions answered wrong.",
    Full_Sample_Mean = NA,
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Create an empty row for separation
  empty_row2 <- data.frame(
    Explanation = "Analysis Summary",
    Question = NA,
    Full_Sample_Mean = NA,
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Create a row for "Aggregated Metrics"
  aggregated_metrics_row <- data.frame(
    Explanation = "Aggregated Metrics",
    Question = NA,
    Full_Sample_Mean = NA,
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Add the relative average Difference between Fraud and Cleaned Sample (RelFC)
  mean_row1 <- data.frame(
    Explanation = "\u00d8 Relative Difference between Fraud and Cleaned Sample",
    Question = NA,
    Full_Sample_Mean = to_percentage(relative_mean_difference_fraud_cleaned),
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Add the relative average Difference between Full and Cleaned Sample (RelFC)
  mean_row2 <- data.frame(
    Explanation = "\u00d8 Relative Difference between Full and Cleaned Sample",
    Question = NA,
    Full_Sample_Mean = to_percentage(relative_mean_difference_full_cleaned),
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Combine all results into a table
  results_table <- rbind(
    empty_row2, response_row, share_row, threshold_row, empty_row1, results_table, aggregated_metrics_row, mean_row1, mean_row2
  )

  # Format the table with flextable
  ft <- flextable(results_table)

  # Set the table headers
  ft <- set_header_labels(
    ft,
    Question = "",
    Full_Sample_Mean = "Full Sample Mean",
    fraud_sample_mean = "Fraud Sample Mean",
    cleaned_sample_mean = "Cleaned Sample Mean",
    difference_fraud_cleaned = "Difference between Fraud and Cleaned Sample",
    difference_full_cleaned = "Difference between Full and Cleaned Sample",
    relative_difference_fraud_cleaned = "Relative Difference between Fraud and Cleaned Sample",
    relative_difference_full_cleaned = "Relative Difference between Full and Cleaned Sample"
  )

  # Format all numerical values to two decimal places
  ft <- colformat_double(ft, j = c("Full_Sample_Mean", "fraud_sample_mean", "cleaned_sample_mean", "difference_fraud_cleaned", "difference_full_cleaned"), digits = 2)

  # Merge and left-align only for the "Analysis Summary" row and make it bold
  ft <- merge_at(ft, i = which(results_table$Explanation == "Analysis Summary"), j = 1:ncol(results_table))
  ft <- align(ft, i = which(results_table$Explanation == "Analysis Summary"), align = "left", part = "body")
  ft <- bold(ft, i = which(results_table$Explanation == "Analysis Summary"), bold = TRUE, part = "body")

  # Merge and left-align only for the "Aggregated Metrics" row and make it bold
  ft <- merge_at(ft, i = which(results_table$Explanation == "Aggregated Metrics"), j = 1:ncol(results_table))
  ft <- align(ft, i = which(results_table$Explanation == "Aggregated Metrics"), align = "left", part = "body")
  ft <- bold(ft, i = which(results_table$Explanation == "Aggregated Metrics"), bold = TRUE, part = "body")

  # Merge and left-align only for the "Metrics per Survey Question" row and make it bold
  ft <- merge_at(ft, i = which(results_table$Explanation == "Metrics per Survey Question"), j = 1:ncol(results_table))
  ft <- align(ft, i = which(results_table$Explanation == "Metrics per Survey Question"), align = "left", part = "body")
  ft <- bold(ft, i = which(results_table$Explanation == "Metrics per Survey Question"), bold = TRUE, part = "body")

  # Merge and left-align only for the "Number of Participants" row
  ft <- merge_at(ft, i = which(results_table$Explanation == "Number of Participants"), j = 1:2)
  ft <- align(ft, i = which(results_table$Explanation == "Number of Participants"), align = "left", part = "body")

  # Merge and left-align only for the "Share of Reliable Participants (SRP)" row
  ft <- merge_at(ft, i = which(results_table$Explanation == "Share of Reliable Participants (SRP)"), j = 1:2)
  ft <- align(ft, i = which(results_table$Explanation == "Share of Reliable Participants (SRP)"), align = "left", part = "body")

  # Merge and left-align only for the "Fraud Detection Threshold" row, spanning the text across the remaining columns
  ft <- merge_at(ft, i = which(results_table$Explanation == "Fraud Detection Threshold"), j = 2:ncol(results_table))
  ft <- align(ft, i = which(results_table$Explanation == "Fraud Detection Threshold"), align = "left", part = "body")

  # Choose a theme for the table appearance
  ft <- theme_vanilla(ft)

  # Save the table as a Word file
  doc_file_path <- file.path(output_dir, "FraudDetec3_results.docx")
  save_as_docx(ft, path = doc_file_path)

  # Save the table as an HTML file
  html_file_path <- file.path(output_dir, "FraudDetec3_results.html")
  save_as_html(ft, path = html_file_path)

  # Display message to user
  message("Note: Output has been saved as Microsoft Word (.docx) and HTML (.html) files under ", output_dir)
  message("Note: The cleaned data set has been saved as .CSV and as RData files under ", output_dir)
  message("Note: Fraud Detection Specification: At least 3 Fraud Detection Questions answered wrong")
  message("Note: Missing values for all variables used in the analysis were cleaned prior to conducting the analysis.")

  # Save the cleaned sample as a CSV file
  cleaned_sample <- data_cleaned[cleaned_sample_indices, !(names(data_cleaned) %in% c("Fraud1", "Fraud2", "Fraud3", "Fraud4", "Fraud5"))]
  csv_file_path <- file.path(output_dir, "FraudDetec3_cleaned_sample.csv")
  write.csv(cleaned_sample, csv_file_path, row.names = FALSE)

  # Save the cleaned sample as an RData file
  rdata_file_path <- file.path(output_dir, "FraudDetec3_cleaned_sample.RData")
  save(cleaned_sample, file = rdata_file_path)

  # Return the flextable
  return(ft)
}


#' Fraud Detection Analysis Tool 4
#'
#' @param output_dir Path specifying where the Word and HTML files will be saved.
#' @param data The data frame containing all the survey data.
#' @param FraudList A character vector of up to 5 Fraud Detection Questions.
#' @param correct_answers A numeric vector representing correct answers for each question. Default is \code{c(0, 0, 0, 0, 0)}.
#' @param ... Survey questions to be analyzed.
#' @return A flextable object with the results.
#' @importFrom flextable flextable
#' @importFrom flextable set_header_labels
#' @importFrom flextable colformat_double
#' @importFrom flextable merge_at
#' @importFrom flextable align
#' @importFrom flextable theme_vanilla
#' @importFrom flextable bold
#' @importFrom flextable save_as_docx
#' @importFrom flextable save_as_html
#' @importFrom utils write.csv
#' @import dplyr
#' @examples
#' if (requireNamespace("flextable", quietly = TRUE) && requireNamespace("officer", quietly = TRUE)) {
#'   library(flextable)
#'   library(officer)
#'
#'   # Example data for fraud detection analysis
#'   Q1 <- c(4, 5, 3, 2, 5, 2)
#'   Q2 <- c(3, 4, 2, 5, 4, 3)
#'   Q3 <- c(5, 4, 3, 5, 4, 5)
#'   Q4 <- c(1, 2, 3, 4, 5, 2)
#'   Q5 <- c(5, 2, 2, 1, 4, 1)
#'   Q6 <- c(5, 2, 3, 5, 1, 2)
#'   Q7 <- c(5, 2, 4, 5, 3, 4)
#'
#'   Fraud1 <- c(0, 1, 0, 0, 0, 0)
#'   Fraud2 <- c(0, 0, 0, 0, 0, 0)
#'   Fraud3 <- c(0, 1, 0, 0, 0, 0)
#'   Fraud4 <- c(0, 0, 1, 0, 0, 1)
#'   Fraud5 <- c(0, 0, 0, 1, 1, 1)
#'
#'   Test_Data_Fraud <- data.frame(Q1, Q2, Q3, Q4, Q5, Q6, Q7, Fraud1, Fraud2, Fraud3, Fraud4, Fraud5)
#'
#'   temp_dir <- tempdir()
#'
#'   FraudDetec4(
#'     output_dir = temp_dir,
#'     data = Test_Data_Fraud,
#'     FraudList = c("Fraud1", "Fraud2", "Fraud3", "Fraud4", "Fraud5"),
#'     correct_answers = c(0, 0, 0, 0, 0),
#'     Q1, Q2, Q3, Q4, Q5, Q6, Q7
#'   )
#' }
#' @export
FraudDetec4 <- function(output_dir, data, FraudList, correct_answers = c(0, 0, 0, 0, 0), ...) {
  # Evaluate question_vars names in the data context
  question_vars <- as.character(substitute(list(...))[-1])

  # Combine FraudList and question_vars for NA removal
  all_vars <- c(FraudList, question_vars)

  # Remove rows with NAs in any of the relevant columns
  data_cleaned <- data %>%
    filter(if_all(all_of(all_vars), ~ !is.na(.)))

  # Extract the cleaned columns from the data
  FraudList_cleaned <- lapply(FraudList, function(x) data_cleaned[[x]])
  questions_cleaned <- lapply(question_vars, function(x) data_cleaned[[x]])

  # Define a function to convert values to percentages
  to_percentage <- function(values) {
    return(paste0(round(values * 100, 2), "%"))
  }

  # Check if at least one question was passed
  if (length(questions_cleaned) == 0) {
    stop("At least one question must be passed.")
  }

  # Check if FraudList and correct_answers contain the same number of elements
  if (length(FraudList) != length(correct_answers)) {
    stop("FraudList and correct_answers must have the same length.")
  }

  # Calculate whether each fraud variable was answered incorrectly
  fraud_mismatch <- mapply(function(fraud, correct) {
    return(fraud != correct)
  }, FraudList_cleaned, correct_answers)

  # Calculate the FraudCheck variable: 1 if at least 4 fraud variable was answered incorrectly, else 0
  FraudCheck <- rowSums(fraud_mismatch) >= 4
  FraudCheck <- as.integer(FraudCheck)  # Convert TRUE/FALSE to 1/0

  # Calculate the share of cases where at least 4 fraud variable was  answered correctly
  share_of_reliable_participations <- mean(FraudCheck == 0)

  # Calculate the Full Sample Mean for each question
  full_sample_mean <- sapply(questions_cleaned, mean, na.rm = TRUE)

  # Calculate the Cleaned Sample Mean for each question
  cleaned_sample_indices <- which(FraudCheck == 0)
  cleaned_sample_mean <- sapply(questions_cleaned, function(q) mean(q[cleaned_sample_indices], na.rm = TRUE))

  # Calculate the Fraud Sample Mean for each question
  fraud_sample_indices <- which(FraudCheck == 1)
  fraud_sample_mean <- sapply(questions_cleaned, function(q) mean(q[fraud_sample_indices], na.rm = TRUE))

  # Calculate the difference and relative value _fraud_cleaned
  difference_fraud_cleaned <- fraud_sample_mean - cleaned_sample_mean
  relative_difference_fraud_cleaned <- cleaned_sample_mean / fraud_sample_mean - 1

  # Calculate the mean of the difference and relative value across all questions _fraud_cleaned
  mean_difference_fraud_cleaned <- mean(abs(difference_fraud_cleaned))
  relative_mean_difference_fraud_cleaned <- mean(abs(relative_difference_fraud_cleaned))

  # Calculate the difference and relative value _full_cleaned
  difference_full_cleaned <- full_sample_mean - cleaned_sample_mean
  relative_difference_full_cleaned <- cleaned_sample_mean / full_sample_mean - 1

  # Calculate the mean of the difference and relative value across all questions _full_cleaned
  mean_difference_full_cleaned <- mean(abs(difference_full_cleaned))
  relative_mean_difference_full_cleaned <- mean(abs(relative_difference_full_cleaned))

  # Calculate the number of participants for each group
  response_count_full <- length(unique(unlist(lapply(questions_cleaned, function(q) which(!is.na(q))))))
  response_count_fraud <- sum(FraudCheck == 1)
  response_count_cleaned <- sum(FraudCheck == 0)

  # Create an empty row for separation
  empty_row1 <- data.frame(
    Explanation = "Metrics per Survey Question",
    Question = NA,
    Full_Sample_Mean = NA,
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Create a table with the calculated values
  results_table <- data.frame(
    Explanation = NA,
    Question = question_vars,  # Use the variable names as question labels
    Full_Sample_Mean = round(full_sample_mean, 2),
    cleaned_sample_mean = cleaned_sample_mean,
    fraud_sample_mean = fraud_sample_mean,
    difference_fraud_cleaned = difference_fraud_cleaned,
    difference_full_cleaned = difference_full_cleaned,
    relative_difference_fraud_cleaned = to_percentage(relative_difference_fraud_cleaned),
    relative_difference_full_cleaned = to_percentage(relative_difference_full_cleaned)
  )

  # Add the number of responses
  response_row <- data.frame(
    Explanation = "Number of Participants",
    Question = NA,
    Full_Sample_Mean = response_count_full,
    cleaned_sample_mean = response_count_cleaned,
    fraud_sample_mean = response_count_fraud,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Add Share of Reliable Participants (SRP)
  share_row <- data.frame(
    Explanation = "Share of Reliable Participants (SRP)",
    Question = NA,
    Full_Sample_Mean = to_percentage(share_of_reliable_participations),
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Add the Fraud Detection Threshold row
  threshold_row <- data.frame(
    Explanation = "Fraud Detection Threshold",
    Question = "At least 4 Fraud Detection Questions answered wrong.",
    Full_Sample_Mean = NA,
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Create an empty row for separation
  empty_row2 <- data.frame(
    Explanation = "Analysis Summary",
    Question = NA,
    Full_Sample_Mean = NA,
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Create a row for "Aggregated Metrics"
  aggregated_metrics_row <- data.frame(
    Explanation = "Aggregated Metrics",
    Question = NA,
    Full_Sample_Mean = NA,
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Add the relative average Difference between Fraud and Cleaned Sample (RelFC)
  mean_row1 <- data.frame(
    Explanation = "\u00d8 Relative Difference between Fraud and Cleaned Sample",
    Question = NA,
    Full_Sample_Mean = to_percentage(relative_mean_difference_fraud_cleaned),
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Add the relative average Difference between Full and Cleaned Sample (RelFC)
  mean_row2 <- data.frame(
    Explanation = "\u00d8 Relative Difference between Full and Cleaned Sample",
    Question = NA,
    Full_Sample_Mean = to_percentage(relative_mean_difference_full_cleaned),
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Combine all results into a table
  results_table <- rbind(
    empty_row2, response_row, share_row, threshold_row, empty_row1, results_table, aggregated_metrics_row, mean_row1, mean_row2
  )

  # Format the table with flextable
  ft <- flextable(results_table)

  # Set the table headers
  ft <- set_header_labels(
    ft,
    Question = "",
    Full_Sample_Mean = "Full Sample Mean",
    fraud_sample_mean = "Fraud Sample Mean",
    cleaned_sample_mean = "Cleaned Sample Mean",
    difference_fraud_cleaned = "Difference between Fraud and Cleaned Sample",
    difference_full_cleaned = "Difference between Full and Cleaned Sample",
    relative_difference_fraud_cleaned = "Relative Difference between Fraud and Cleaned Sample",
    relative_difference_full_cleaned = "Relative Difference between Full and Cleaned Sample"
  )

  # Format all numerical values to two decimal places
  ft <- colformat_double(ft, j = c("Full_Sample_Mean", "fraud_sample_mean", "cleaned_sample_mean", "difference_fraud_cleaned", "difference_full_cleaned"), digits = 2)

  # Merge and left-align only for the "Analysis Summary" row and make it bold
  ft <- merge_at(ft, i = which(results_table$Explanation == "Analysis Summary"), j = 1:ncol(results_table))
  ft <- align(ft, i = which(results_table$Explanation == "Analysis Summary"), align = "left", part = "body")
  ft <- bold(ft, i = which(results_table$Explanation == "Analysis Summary"), bold = TRUE, part = "body")

  # Merge and left-align only for the "Aggregated Metrics" row and make it bold
  ft <- merge_at(ft, i = which(results_table$Explanation == "Aggregated Metrics"), j = 1:ncol(results_table))
  ft <- align(ft, i = which(results_table$Explanation == "Aggregated Metrics"), align = "left", part = "body")
  ft <- bold(ft, i = which(results_table$Explanation == "Aggregated Metrics"), bold = TRUE, part = "body")

  # Merge and left-align only for the "Metrics per Survey Question" row and make it bold
  ft <- merge_at(ft, i = which(results_table$Explanation == "Metrics per Survey Question"), j = 1:ncol(results_table))
  ft <- align(ft, i = which(results_table$Explanation == "Metrics per Survey Question"), align = "left", part = "body")
  ft <- bold(ft, i = which(results_table$Explanation == "Metrics per Survey Question"), bold = TRUE, part = "body")

  # Merge and left-align only for the "Number of Participants" row
  ft <- merge_at(ft, i = which(results_table$Explanation == "Number of Participants"), j = 1:2)
  ft <- align(ft, i = which(results_table$Explanation == "Number of Participants"), align = "left", part = "body")

  # Merge and left-align only for the "Share of Reliable Participants (SRP)" row
  ft <- merge_at(ft, i = which(results_table$Explanation == "Share of Reliable Participants (SRP)"), j = 1:2)
  ft <- align(ft, i = which(results_table$Explanation == "Share of Reliable Participants (SRP)"), align = "left", part = "body")

  # Merge and left-align only for the "Fraud Detection Threshold" row, spanning the text across the remaining columns
  ft <- merge_at(ft, i = which(results_table$Explanation == "Fraud Detection Threshold"), j = 2:ncol(results_table))
  ft <- align(ft, i = which(results_table$Explanation == "Fraud Detection Threshold"), align = "left", part = "body")

  # Choose a theme for the table appearance
  ft <- theme_vanilla(ft)

  # Save the table as a Word file
  doc_file_path <- file.path(output_dir, "FraudDetec4_results.docx")
  save_as_docx(ft, path = doc_file_path)

  # Save the table as an HTML file
  html_file_path <- file.path(output_dir, "FraudDetec4_results.html")
  save_as_html(ft, path = html_file_path)

  # Display message to user
  message("Note: Output has been saved as Microsoft Word (.docx) and HTML (.html) files under ", output_dir)
  message("Note: The cleaned data set has been saved as .CSV and as RData files under ", output_dir)
  message("Note: Fraud Detection Specification: At least 4 Fraud Detection Questions answered wrong")
  message("Note: Missing values for all variables used in the analysis were cleaned prior to conducting the analysis.")

  # Save the cleaned sample as a CSV file
  cleaned_sample <- data_cleaned[cleaned_sample_indices, !(names(data_cleaned) %in% c("Fraud1", "Fraud2", "Fraud3", "Fraud4", "Fraud5"))]
  csv_file_path <- file.path(output_dir, "FraudDetec4_cleaned_sample.csv")
  write.csv(cleaned_sample, csv_file_path, row.names = FALSE)

  # Save the cleaned sample as an RData file
  rdata_file_path <- file.path(output_dir, "FraudDetec4_cleaned_sample.RData")
  save(cleaned_sample, file = rdata_file_path)

  # Return the flextable
  return(ft)
}


#' Fraud Detection Analysis Tool 5
#'
#' @param output_dir Path specifying where the Word and HTML files will be saved.
#' @param data The data frame containing all the survey data.
#' @param FraudList A character vector of up to 5 Fraud Detection Questions.
#' @param correct_answers A numeric vector representing correct answers for each question. Default is \code{c(0, 0, 0, 0, 0)}.
#' @param ... Survey questions to be analyzed.
#' @return A flextable object with the results.
#' @importFrom flextable flextable
#' @importFrom flextable set_header_labels
#' @importFrom flextable colformat_double
#' @importFrom flextable merge_at
#' @importFrom flextable align
#' @importFrom flextable theme_vanilla
#' @importFrom flextable bold
#' @importFrom flextable save_as_docx
#' @importFrom flextable save_as_html
#' @importFrom utils write.csv
#' @import dplyr
#' @examples
#' if (requireNamespace("flextable", quietly = TRUE) && requireNamespace("officer", quietly = TRUE)) {
#'   library(flextable)
#'   library(officer)
#'
#'   # Example data for fraud detection analysis
#'   Q1 <- c(4, 5, 3, 2, 5, 2)
#'   Q2 <- c(3, 4, 2, 5, 4, 3)
#'   Q3 <- c(5, 4, 3, 5, 4, 5)
#'   Q4 <- c(1, 2, 3, 4, 5, 2)
#'   Q5 <- c(5, 2, 2, 1, 4, 1)
#'   Q6 <- c(5, 2, 3, 5, 1, 2)
#'   Q7 <- c(5, 2, 4, 5, 3, 4)
#'
#'   Fraud1 <- c(0, 1, 0, 0, 0, 0)
#'   Fraud2 <- c(0, 0, 0, 0, 0, 0)
#'   Fraud3 <- c(0, 1, 0, 0, 0, 0)
#'   Fraud4 <- c(0, 0, 1, 0, 0, 1)
#'   Fraud5 <- c(0, 0, 0, 1, 1, 1)
#'
#'   Test_Data_Fraud <- data.frame(Q1, Q2, Q3, Q4, Q5, Q6, Q7, Fraud1, Fraud2, Fraud3, Fraud4, Fraud5)
#'
#'   temp_dir <- tempdir()
#'
#'   FraudDetec5(
#'     output_dir = temp_dir,
#'     data = Test_Data_Fraud,
#'     FraudList = c("Fraud1", "Fraud2", "Fraud3", "Fraud4", "Fraud5"),
#'     correct_answers = c(0, 0, 0, 0, 0),
#'     Q1, Q2, Q3, Q4, Q5, Q6, Q7
#'   )
#' }
#' @export
FraudDetec5 <- function(output_dir, data, FraudList, correct_answers = c(0, 0, 0, 0, 0), ...) {
  # Evaluate question_vars names in the data context
  question_vars <- as.character(substitute(list(...))[-1])

  # Combine FraudList and question_vars for NA removal
  all_vars <- c(FraudList, question_vars)

  # Remove rows with NAs in any of the relevant columns
  data_cleaned <- data %>%
    filter(if_all(all_of(all_vars), ~ !is.na(.)))

  # Extract the cleaned columns from the data
  FraudList_cleaned <- lapply(FraudList, function(x) data_cleaned[[x]])
  questions_cleaned <- lapply(question_vars, function(x) data_cleaned[[x]])

  # Define a function to convert values to percentages
  to_percentage <- function(values) {
    return(paste0(round(values * 100, 2), "%"))
  }

  # Check if at least one question was passed
  if (length(questions_cleaned) == 0) {
    stop("At least one question must be passed.")
  }

  # Check if FraudList and correct_answers contain the same number of elements
  if (length(FraudList) != length(correct_answers)) {
    stop("FraudList and correct_answers must have the same length.")
  }

  # Calculate whether each fraud variable was answered incorrectly
  fraud_mismatch <- mapply(function(fraud, correct) {
    return(fraud != correct)
  }, FraudList_cleaned, correct_answers)

  # Calculate the FraudCheck variable: 1 if 5 fraud variable was answered incorrectly, else 0
  FraudCheck <- rowSums(fraud_mismatch) >= 5
  FraudCheck <- as.integer(FraudCheck)  # Convert TRUE/FALSE to 1/0

  # Calculate the share of cases where 5 fraud variable was answered correctly
  share_of_reliable_participations <- mean(FraudCheck == 0)

  # Calculate the Full Sample Mean for each question
  full_sample_mean <- sapply(questions_cleaned, mean, na.rm = TRUE)

  # Calculate the Cleaned Sample Mean for each question
  cleaned_sample_indices <- which(FraudCheck == 0)
  cleaned_sample_mean <- sapply(questions_cleaned, function(q) mean(q[cleaned_sample_indices], na.rm = TRUE))

  # Calculate the Fraud Sample Mean for each question
  fraud_sample_indices <- which(FraudCheck == 1)
  fraud_sample_mean <- sapply(questions_cleaned, function(q) mean(q[fraud_sample_indices], na.rm = TRUE))

  # Calculate the difference and relative value _fraud_cleaned
  difference_fraud_cleaned <- fraud_sample_mean - cleaned_sample_mean
  relative_difference_fraud_cleaned <- cleaned_sample_mean / fraud_sample_mean - 1

  # Calculate the mean of the difference and relative value across all questions _fraud_cleaned
  mean_difference_fraud_cleaned <- mean(abs(difference_fraud_cleaned))
  relative_mean_difference_fraud_cleaned <- mean(abs(relative_difference_fraud_cleaned))

  # Calculate the difference and relative value _full_cleaned
  difference_full_cleaned <- full_sample_mean - cleaned_sample_mean
  relative_difference_full_cleaned <- cleaned_sample_mean / full_sample_mean - 1

  # Calculate the mean of the difference and relative value across all questions _full_cleaned
  mean_difference_full_cleaned <- mean(abs(difference_full_cleaned))
  relative_mean_difference_full_cleaned <- mean(abs(relative_difference_full_cleaned))

  # Calculate the number of participants for each group
  response_count_full <- length(unique(unlist(lapply(questions_cleaned, function(q) which(!is.na(q))))))
  response_count_fraud <- sum(FraudCheck == 1)
  response_count_cleaned <- sum(FraudCheck == 0)

  # Create an empty row for separation
  empty_row1 <- data.frame(
    Explanation = "Metrics per Survey Question",
    Question = NA,
    Full_Sample_Mean = NA,
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Create a table with the calculated values
  results_table <- data.frame(
    Explanation = NA,
    Question = question_vars,  # Use the variable names as question labels
    Full_Sample_Mean = round(full_sample_mean, 2),
    cleaned_sample_mean = cleaned_sample_mean,
    fraud_sample_mean = fraud_sample_mean,
    difference_fraud_cleaned = difference_fraud_cleaned,
    difference_full_cleaned = difference_full_cleaned,
    relative_difference_fraud_cleaned = to_percentage(relative_difference_fraud_cleaned),
    relative_difference_full_cleaned = to_percentage(relative_difference_full_cleaned)
  )

  # Add the number of responses
  response_row <- data.frame(
    Explanation = "Number of Participants",
    Question = NA,
    Full_Sample_Mean = response_count_full,
    cleaned_sample_mean = response_count_cleaned,
    fraud_sample_mean = response_count_fraud,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Add Share of Reliable Participants (SRP)
  share_row <- data.frame(
    Explanation = "Share of Reliable Participants (SRP)",
    Question = NA,
    Full_Sample_Mean = to_percentage(share_of_reliable_participations),
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Add the Fraud Detection Threshold row
  threshold_row <- data.frame(
    Explanation = "Fraud Detection Threshold",
    Question = "5 Fraud Detection Questions answered wrong.",
    Full_Sample_Mean = NA,
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Create an empty row for separation
  empty_row2 <- data.frame(
    Explanation = "Analysis Summary",
    Question = NA,
    Full_Sample_Mean = NA,
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Create a row for "Aggregated Metrics"
  aggregated_metrics_row <- data.frame(
    Explanation = "Aggregated Metrics",
    Question = NA,
    Full_Sample_Mean = NA,
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Add the relative average Difference between Fraud and Cleaned Sample (RelFC)
  mean_row1 <- data.frame(
    Explanation = "\u00d8 Relative Difference between Fraud and Cleaned Sample",
    Question = NA,
    Full_Sample_Mean = to_percentage(relative_mean_difference_fraud_cleaned),
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Add the relative average Difference between Full and Cleaned Sample (RelFC)
  mean_row2 <- data.frame(
    Explanation = "\u00d8 Relative Difference between Full and Cleaned Sample",
    Question = NA,
    Full_Sample_Mean = to_percentage(relative_mean_difference_full_cleaned),
    cleaned_sample_mean = NA,
    fraud_sample_mean = NA,
    difference_fraud_cleaned = NA,
    difference_full_cleaned = NA,
    relative_difference_fraud_cleaned = NA,
    relative_difference_full_cleaned = NA
  )

  # Combine all results into a table
  results_table <- rbind(
    empty_row2, response_row, share_row, threshold_row, empty_row1, results_table, aggregated_metrics_row, mean_row1, mean_row2
  )

  # Format the table with flextable
  ft <- flextable(results_table)

  # Set the table headers
  ft <- set_header_labels(
    ft,
    Question = "",
    Full_Sample_Mean = "Full Sample Mean",
    fraud_sample_mean = "Fraud Sample Mean",
    cleaned_sample_mean = "Cleaned Sample Mean",
    difference_fraud_cleaned = "Difference between Fraud and Cleaned Sample",
    difference_full_cleaned = "Difference between Full and Cleaned Sample",
    relative_difference_fraud_cleaned = "Relative Difference between Fraud and Cleaned Sample",
    relative_difference_full_cleaned = "Relative Difference between Full and Cleaned Sample"
  )

  # Format all numerical values to two decimal places
  ft <- colformat_double(ft, j = c("Full_Sample_Mean", "fraud_sample_mean", "cleaned_sample_mean", "difference_fraud_cleaned", "difference_full_cleaned"), digits = 2)

  # Merge and left-align only for the "Analysis Summary" row and make it bold
  ft <- merge_at(ft, i = which(results_table$Explanation == "Analysis Summary"), j = 1:ncol(results_table))
  ft <- align(ft, i = which(results_table$Explanation == "Analysis Summary"), align = "left", part = "body")
  ft <- bold(ft, i = which(results_table$Explanation == "Analysis Summary"), bold = TRUE, part = "body")

  # Merge and left-align only for the "Aggregated Metrics" row and make it bold
  ft <- merge_at(ft, i = which(results_table$Explanation == "Aggregated Metrics"), j = 1:ncol(results_table))
  ft <- align(ft, i = which(results_table$Explanation == "Aggregated Metrics"), align = "left", part = "body")
  ft <- bold(ft, i = which(results_table$Explanation == "Aggregated Metrics"), bold = TRUE, part = "body")

  # Merge and left-align only for the "Metrics per Survey Question" row and make it bold
  ft <- merge_at(ft, i = which(results_table$Explanation == "Metrics per Survey Question"), j = 1:ncol(results_table))
  ft <- align(ft, i = which(results_table$Explanation == "Metrics per Survey Question"), align = "left", part = "body")
  ft <- bold(ft, i = which(results_table$Explanation == "Metrics per Survey Question"), bold = TRUE, part = "body")

  # Merge and left-align only for the "Number of Participants" row
  ft <- merge_at(ft, i = which(results_table$Explanation == "Number of Participants"), j = 1:2)
  ft <- align(ft, i = which(results_table$Explanation == "Number of Participants"), align = "left", part = "body")

  # Merge and left-align only for the "Share of Reliable Participants (SRP)" row
  ft <- merge_at(ft, i = which(results_table$Explanation == "Share of Reliable Participants (SRP)"), j = 1:2)
  ft <- align(ft, i = which(results_table$Explanation == "Share of Reliable Participants (SRP)"), align = "left", part = "body")

  # Merge and left-align only for the "Fraud Detection Threshold" row, spanning the text across the remaining columns
  ft <- merge_at(ft, i = which(results_table$Explanation == "Fraud Detection Threshold"), j = 2:ncol(results_table))
  ft <- align(ft, i = which(results_table$Explanation == "Fraud Detection Threshold"), align = "left", part = "body")

  # Choose a theme for the table appearance
  ft <- theme_vanilla(ft)

  # Save the table as a Word file
  doc_file_path <- file.path(output_dir, "FraudDetec5_results.docx")
  save_as_docx(ft, path = doc_file_path)

  # Save the table as an HTML file
  html_file_path <- file.path(output_dir, "FraudDetec5_results.html")
  save_as_html(ft, path = html_file_path)

  # Display message to user
  message("Note: Output has been saved as Microsoft Word (.docx) and HTML (.html) files under ", output_dir)
  message("Note: The cleaned data set has been saved as .CSV and as RData files under ", output_dir)
  message("Note: Fraud Detection Specification: 5 Fraud Detection Questions answered wrong")
  message("Note: Missing values for all variables used in the analysis were cleaned prior to conducting the analysis.")

  # Save the cleaned sample as a CSV file
  cleaned_sample <- data_cleaned[cleaned_sample_indices, !(names(data_cleaned) %in% c("Fraud1", "Fraud2", "Fraud3", "Fraud4", "Fraud5"))]
  csv_file_path <- file.path(output_dir, "FraudDetec5_cleaned_sample.csv")
  write.csv(cleaned_sample, csv_file_path, row.names = FALSE)

  # Save the cleaned sample as an RData file
  rdata_file_path <- file.path(output_dir, "FraudDetec5_cleaned_sample.RData")
  save(cleaned_sample, file = rdata_file_path)

  # Return the flextable
  return(ft)
}


