check.other.option <- function(tool.survey, question.name) {
  if (!question.name %in% tool.survey$name) {
    stop(paste0("Question ", question.name, " not found in tool survey!"))
  }
  question.name.other <- paste0(question.name, "_other")
  return (question.name.other %in% tool.survey$name)
}


load.responses <- function(choice_list_name, tool.choices) {
  choices <- tool.choices %>%
    dplyr::filter(list_name == choice_list_name)
  responses_eng <- paste0(choices$`label::English`, collapse = "\n")
  responses_rus <- paste0(choices$`label::Russian`, collapse = "\n")
  responses_ukr <- paste0(choices$`label::Ukrainian`, collapse = "\n")
  if (any(is.na(choices$`label::English`))) {
    responses_eng <- NA
  }
  if (any(is.na(choices$`label::Russian`))) {
    responses_rus <- NA
  }
  if (any(is.na(choices$`label::Ukrainian`))) {
    responses_ukr <- NA
  }
  return (list(
    responses_eng = responses_eng,
    responses_rus = responses_rus,
    responses_ukr = responses_ukr
  ))
}


scrap.formulas <- function(dap, colname, question.name) {
  formula <- dap[dap$`Indicator / Variable (name)` == question.name, colname][1]
  return (ifelse(is.na(formula), NA, formula))
}


check.row.identical <- function(row1, row2) {
  if (ncol(row1) != ncol(row2)) {
    return(FALSE)
  }
  if (!all(colnames(row1) == colnames(row2))) {
    return(FALSE)
  }
  for (colname in colnames(row1)) {
    if ((is.na(row1[1, colname]) & !is.na(row2[1, colname])) | (!is.na(row1[1, colname]) & is.na(row2[1, colname]))) {
      return(FALSE)
    }
    if (is.na(row1[1, colname]) & is.na(row2[1, colname])) {
      next
    }
    if (!(row1[1, colname] == row2[1, colname])) {
      # print(row1[1, colname])
      # print(row2[1, colname])
      return(FALSE)
    }
  }
  return(TRUE)
}


cast.strings <- function(data) {
  
  data <- data %>%
    dplyr::mutate(across(where(is.character), ~ gsub("\r\n", "\n", .x))) %>%
    dplyr::mutate(across(where(is.character), ~ gsub("\r", "", .x)))
  
  return(data)
}


dap.preparation <- function(data) {
  # remove rows with all NA
  data <- data[!apply(is.na(data), 1, all), ]
  # remove \n in the end of the strings
  for (colname in colnames(data)) {
    data[, colname] <- gsub("\n$", "", data[, colname])
  }
  # remove \r in the end of the strings
  for (colname in colnames(data)) {
    data[, colname] <- gsub("\r$", "", data[, colname])
  }
  # remove \n in the beginning of the strings
  for (colname in colnames(data)) {
    data[, colname] <- gsub("^\n", "", data[, colname])
  }
  # remove \r in the beginning of the strings
  for (colname in colnames(data)) {
    data[, colname] <- gsub("^\r", "", data[, colname])
  }
  # cast \r\n to the \r\n to the \n
  data <- cast.strings(data)
  # create data.frame from data
  data <- as.data.frame(data)
  
  return(data)
}
