library(openxlsx)
library(readxl)
library(utilityR)


source("www/utils.R")

create.tool <- function(dap.file, survey.file) {

  dap <- as.data.frame(readxl::read_xlsx(dap.file, sheet = "DAP__R_"))

  dap <- dap[!is.na(dap$`change question`),]
  dap <- dap.preparation(dap)
  saved.questions <- dap[dap$`change question` != "deletion",]

  ################################### create survey sheet

  survey <- data.frame(
    number_indicator = rep(NA, 5),
    type = c("start", "end", "today", "deviceid", "audit"),
    sector = rep(NA, 5),
    name = c("start", "end", "today", "deviceid", "audit"),
    xml = c("start", "end", "today", "deviceid", "audit"),
    `label::English` = rep(NA, 5),
    `label::Ukrainian` = rep(NA, 5),
    `label::Russian` = rep(NA, 5),
    `hint::English` = rep(NA, 5),
    `hint::Ukrainian` = rep(NA, 5),
    `hint::Russian` = rep(NA, 5),
    required = rep(NA, 5),
    appearance = rep(NA, 5),
    relevant = rep(NA, 5),
    constraint = rep(NA, 5),
    `constraint_message::English` = rep(NA, 5),
    `constraint_message::Ukrainian` = rep(NA, 5),
    `constraint_message::Russian` = rep(NA, 5),
    default = rep(NA, 5),
    choice_filter = rep(NA, 5),
    calculation = rep(NA, 5),
    parameters = c(NA, NA, NA, NA, "track-changes=true"),
    check.names = FALSE
  )

  for (i in 1:nrow(saved.questions)) {
    question <- saved.questions[i,]
    question_number = dplyr::if_else(is.na(question$`new number`), question$`old number`, question$`new number`)
    if (grepl("group", question$`Groups`)) {
      type <-  question$`Groups`
      name <- question$`Indicator / Variable (name)`
    } else {
      if (!is.na(question_number)) {
        name <- paste(question_number, question$`Indicator / Variable (name)`, sep = "_")
      } else {
        name <- question$`Indicator / Variable (name)`
      }

      if (grepl("select_", question$`Question Type`)) {
        type <- paste(question$`Question Type`, question$`Indicator / Variable (name)`)
      } else {
        type <- question$`Question Type`
      }
    }
    survey <- dplyr::bind_rows(survey, data.frame(
      number_indicator = question_number,
      type = type,
      sector = question$`Indicator group / sector`,
      name = name,
      xml = question$`Indicator / Variable (name)`,
      `label::English` = question$`Questionnaire Question`,
      `label::Ukrainian` = question$`Questionnaire Question UKR`,
      `label::Russian` = question$`Questionnaire Question RUS`,
      `hint::English` = question$`Hint`,
      `hint::Ukrainian` = question$`Hint UKR`,
      `hint::Russian` = question$`Hint RUS`,
      relevant = question$`Relevance`,
      relevant_do_text = question$`Relevance_do_text`,
      constraint = question$`Constraint`,
      constraint_do_text = question$`Constraint_do_text`,
      `constraint_message::English` = question$`Сonstraint message English`,
      `constraint_message::Ukrainian` = question$`Сonstraint message UKR`,
      `constraint_message::Russian` = question$`Сonstraint message RUS`,
      calculation = question$`Calculation`,
      required = ifelse(type != "note" & type != "calculate" & type != "" & !is.na(type), "true", "false"),
      appearance = NA,
      check.names = FALSE
    ))
  }

  ################################### create choices sheet

  choices <- data.frame(
    list_name = as.character(),
    name = as.character(),
    `label::English` = as.character(),
    `label::Ukrainian` = as.character(),
    `label::Russian` = as.character(),
    check.names = FALSE
  )

  for (i in 1:nrow(saved.questions)) {
    question <- saved.questions[i,]
    if (grepl("^select_", question$`Question Type`)) {
      question.name <- question$`Indicator / Variable (name)`
      choices.eng <- unlist(stringr::str_split(question$`Questionnaire Responses`, "\n"))
      choices.ukr <- unlist(stringr::str_split(question$`Questionnaire Responses UKR`, "\n"))
      choices.rus <- unlist(stringr::str_split(question$`Questionnaire Responses RUS`, "\n"))
      # strip \r
      choices.eng <- gsub("\r", "", choices.eng)
      choices.ukr <- gsub("\r", "", choices.ukr)
      choices.rus <- gsub("\r", "", choices.rus)
      for (j in 1:length(choices.eng)) {
        if (is.na(choices.eng[j])) {
          next
        }
        choice <- sub("\\([^)]*\\)", "", choices.eng[j])
        choice <- gsub("\\s{2,}", " ", choice)
        choice <- gsub(' ', '_', tolower(trimws(choice)))
        if (choice != "") {
          choices <- rbind(choices, data.frame(list_name = question.name,
                                               name = choice,
                                               `label::English` = choices.eng[j],
                                               `label::Ukrainian` = choices.ukr[j],
                                               `label::Russian` = choices.rus[j],
                                               check.names = FALSE))
        }
      }
    }
  }

  ################################### create settings sheet

  date <- Sys.Date()
  formatted_date <- format(date, "%Y-%m-%d")

  settings <- data.frame(
    id_string = NA,
    form_title = NA,
    version = c(formatted_date),
    default_language = c("Ukrainian"),
    allow_choice_duplicates = c("yes"),
    check.names = FALSE
  )

  ################################### combine into a doc
  # wb <- createWorkbook()
  wb <- loadWorkbook(dap.file)
  existing_sheets <- openxlsx::getSheetNames(dap.file)
  for (sheet in existing_sheets) {
    if (sheet != "READ_ME") {
      openxlsx::removeWorksheet(wb, sheet)
    }
  }
  readme.worksheet <- read.xlsx(wb, sheet = "READ_ME", colNames = FALSE)

  settings$form_title[1] <- readme.worksheet[1,2]
  settings$id_string[1] <- readme.worksheet[2,2]

  addWorksheet(wb, "survey")
  writeData(wb, "survey", survey)
  freezePane(wb, sheet = "survey", firstRow = T)
  addFilter(wb, sheet = "survey", row = 1, cols = 1:ncol(survey))

  addWorksheet(wb, "choices")
  writeData(wb, "choices", choices)

  addWorksheet(wb, "settings")
  writeData(wb, "settings", settings)

  headerStyle <- createStyle(
    fontSize = 13, halign = "center",
    fgFill = "#4F81BD", borderColour = "#4F81BD",

  )

  style <- createStyle(
    fontSize = 12, halign = "left", valign = "center",
    borderColour = "#000000"
  )


  addStyle(wb, sheet = "survey", style = headerStyle, rows = 1, cols = 1:ncol(survey))
  addStyle(wb, sheet = "survey", style = style, rows = 1:nrow(survey) + 1, cols = 1:ncol(survey), gridExpand = T)

  addStyle(wb, sheet = "choices", style = headerStyle, rows = 1, cols = 1:ncol(choices))
  addStyle(wb, sheet = "choices", style = style, rows = 1:nrow(choices) + 1, cols = 1:ncol(choices), gridExpand = T)

  setColWidths(wb, sheet = "survey", cols = 1:ncol(survey), widths = rep(45, ncol(survey)))
  setColWidths(wb, sheet = "choices", cols = 1:ncol(choices), widths = rep(30, ncol(choices)))


  # saveWorkbook(wb, paste0("resources/", survey.file), overwrite = TRUE)

  cat(paste0("Created ", survey.file, " in resources folder\n"))
  cat("Edit settings sheet for usage")
  return(wb)
}


create.dap <- function(survey.file, old_dap.file, new.dap) {

  tool.survey <- utilityR::load.tool.survey(survey.file, keep_cols = TRUE)
  tool.survey <- tool.survey[!tool.survey$type %in% c("start", "end", "today", "deviceid", "audit"),]
  tool.choices <- openxlsx::read.xlsx(survey.file, sheet = "choices")
  tool.survey <- cast.strings(tool.survey)
  readme.page <- openxlsx::read.xlsx(survey.file, sheet = "READ_ME")
  new.dap <- data.frame(check.names = FALSE)
  for (i in 1:nrow(tool.survey)) {

    if (!is.na(tool.survey$name[i])) {
      choice_list_name <- utilityR::get.choice.list.from.name(tool.survey$name[i], tool.survey)
      resposes <- load.responses(choice_list_name, tool.choices)
    }
    else {
      resposes <- list(
        responses_eng = NA,
        responses_rus = NA,
        responses_ukr =  NA
      )
    }

    is_group <- grepl("_group", tool.survey$`type`[i])
    new.dap <- dplyr::bind_rows(new.dap, data.frame(
      `Groups` = ifelse(is_group, tool.survey$`type`[i], NA),
      `change question` = "no",
      `old number` = NA,
      `new number` = tool.survey$number_indicator[i],
      `Question Type` = ifelse(is_group, NA, stringr::str_split(tool.survey$type[i], " ")[[1]][1]),
      `Indicator / Variable (name)` = tool.survey$xml[i],
      `old Questionnaire Question` = NA,
      `Questionnaire Question` = tool.survey$`label::English`[i],
      `Questionnaire Question RUS` = tool.survey$`label::Russian`[i],
      `Questionnaire Question UKR` = tool.survey$`label::Ukrainian`[i],
      `old Questionnaire Responses` = NA,
      `Questionnaire Responses` = resposes$responses_eng,
      `Questionnaire Responses RUS` = resposes$responses_rus,
      `Questionnaire Responses UKR` = resposes$responses_ukr,
      `Hint` = tool.survey$`hint::English`[i],
      `Hint RUS` = tool.survey$`hint::Russian`[i],
      `Hint UKR` = tool.survey$`hint::Ukrainian`[i],
      `Relevance` = tool.survey$`relevant`[i],
      `Relevance_do_text` = ifelse("relevant_do_text" %in% colnames(tool.survey), tool.survey$`relevant_do_text`[i], NA),
      `Constraint` = tool.survey$`constraint`[i],
      `Constraint_do_text` = ifelse("constraint_do_text" %in% colnames(tool.survey), tool.survey$`constraint_do_text`[i], NA),
      `Сonstraint message English` = tool.survey$`constraint_message::English`[i],
      `Сonstraint message UKR` = tool.survey$`constraint_message::Ukrainian`[i],
      `Сonstraint message RUS` = tool.survey$`constraint_message::Russian`[i],
      `Calculation` = tool.survey$`calculation`[i],
      `Data collection method` = NA,
      `Indicator group / sector` = tool.survey$sector[i],
      `Other (specify) Question` = NA,
      check.names = FALSE
    ))
  }

  wb <- openxlsx::loadWorkbook(paste0("./resources/", old_dap.file))
  old_dap.sheet <- openxlsx::read.xlsx(paste0("./resources/", old_dap.file), sheet = "DAP__R_")
  empty.data <- data.frame(matrix(NA, nrow = nrow(old_dap.sheet), ncol = ncol(old_dap.sheet)))
  openxlsx::writeData(wb, "DAP__R_", empty.data, startRow = 1, startCol = 1)

  existing_sheets <- openxlsx::getSheetNames(paste0("./resources/", old_dap.file))
  for (sheet in existing_sheets) {
    if (sheet != "DAP__R_" & sheet != "type" & sheet != "change" & sheet != "READ_ME") {
      openxlsx::removeWorksheet(wb, sheet)
    }
  }
  openxlsx::writeData(wb, "DAP__R_", new.dap, startRow = 1, startCol = 1)
  openxlsx::writeData(wb, "READ_ME", readme.page, startRow = 1, startCol = 1)

  other_formulas <- c()
  for (rowid in 2:999) {
    formula <- paste0('=IF(OR(B', rowid, '="new",B', rowid, '="yes"),IF(ISNUMBER(SEARCH("Other",L', rowid, ')),"add an additional text question below",""),IF(AND(B', rowid, '="deletion",ISNUMBER(SEARCH("Other",L', rowid, '))),"deletion an additional text question below",""))')
    other_formulas <- c(other_formulas, formula)
  }
  constraint_formulas <- c()
  last_rowid <- 2
  for (rowid in 1:999) {
    if (is.na(tool.survey$`constraint`[rowid])) {

      formula <- paste0('=IF(AND(OR(B', rowid + 1, '="new",B', rowid + 1, '="yes",B', rowid + 1, '="no"),E', rowid + 1, '="select_multiple"),IF(ISNUMBER(SEARCH("None",L', rowid + 1, ')),"not(selected(., ', "'none'", ') and (count-selected(.)>1))",""),"")')
      constraint_formulas <- c(constraint_formulas, formula)
    } else {
      print(length(constraint_formulas))
      print(last_rowid)
      print(rowid)
      if (length(constraint_formulas) > 0) {
        writeFormula(wb, sheet = "DAP__R_", x = constraint_formulas, startCol = which(colnames(new.dap) == "Constraint"), startRow = last_rowid)
      }
      constraint_formulas <- c()
      last_rowid <- rowid + 2
    }
  }
  writeFormula(wb, sheet = "DAP__R_", x = constraint_formulas, startCol = which(colnames(new.dap) == "Constraint"), startRow = last_rowid)
  writeFormula(wb, sheet = "DAP__R_", x = other_formulas, startCol = which(colnames(new.dap) == "Other (specify) Question"), startRow = 2)

  return(wb)
}


create.validation.dap <- function(survey.file, old_dap.file, new.dap) {

  tool.survey <- utilityR::load.tool.survey(survey.file, keep_cols = TRUE)
  tool.survey <- tool.survey[!tool.survey$type %in% c("start", "end", "today", "deviceid", "audit"),]
  tool.choices <- openxlsx::read.xlsx(survey.file, sheet = "choices")
  tool.survey <- cast.strings(tool.survey)
  readme.page <- openxlsx::read.xlsx(survey.file, sheet = "READ_ME")
  new.dap <- data.frame(check.names = FALSE)
  for (i in 1:nrow(tool.survey)) {

    if (!is.na(tool.survey$name[i])) {
      choice_list_name <- utilityR::get.choice.list.from.name(tool.survey$name[i], tool.survey)
      resposes <- load.responses(choice_list_name, tool.choices)
    }
    else {
      resposes <- list(
        responses_eng = NA,
        responses_rus = NA,
        responses_ukr =  NA
      )
    }

    is_group <- grepl("_group", tool.survey$`type`[i])
    new.dap <- dplyr::bind_rows(new.dap, data.frame(
      `new number` = tool.survey$number_indicator[i],
      `Question Type` = ifelse(is_group, NA, stringr::str_split(tool.survey$type[i], " ")[[1]][1]),
      `Indicator / Variable (name)` = tool.survey$xml[i],
      `Questionnaire Question` = tool.survey$`label::English`[i],
      `Questionnaire Question RUS` = tool.survey$`label::Russian`[i],
      `Questionnaire Question UKR` = tool.survey$`label::Ukrainian`[i],
      `Questionnaire Responses` = resposes$responses_eng,
      `Questionnaire Responses RUS` = resposes$responses_rus,
      `Questionnaire Responses UKR` = resposes$responses_ukr,
      `Hint` = tool.survey$`hint::English`[i],
      `Hint RUS` = tool.survey$`hint::Russian`[i],
      `Hint UKR` = tool.survey$`hint::Ukrainian`[i],
      `Relevance` = tool.survey$`relevant`[i],
      `Constraint` = tool.survey$`constraint`[i],
      `Required` = tool.survey$`required`[i],
      `Data collection method` = NA,
      `Indicator group / sector` = tool.survey$sector[i],
      check.names = FALSE
    ))
  }

  wb <- openxlsx::loadWorkbook(paste0("./resources/", old_dap.file))
  old_dap.sheet <- openxlsx::read.xlsx(paste0("./resources/", old_dap.file), sheet = "DAP__R_")
  empty.data <- data.frame(matrix(NA, nrow = nrow(old_dap.sheet), ncol = ncol(old_dap.sheet)))
  openxlsx::writeData(wb, "DAP__R_", empty.data, startRow = 1, startCol = 1)

  existing_sheets <- openxlsx::getSheetNames(paste0("./resources/", old_dap.file))
  for (sheet in existing_sheets) {
    if (sheet != "DAP__R_" & sheet != "type" & sheet != "change" & sheet != "READ_ME") {
      openxlsx::removeWorksheet(wb, sheet)
    }
  }
  openxlsx::writeData(wb, "DAP__R_", new.dap, startRow = 1, startCol = 1)
  openxlsx::writeData(wb, "READ_ME", readme.page, startRow = 1, startCol = 1)

  other_formulas <- c()
  for (rowid in 2:999) {
    formula <- paste0('=IF(OR(B', rowid, '="new",B', rowid, '="yes"),IF(ISNUMBER(SEARCH("Other",K', rowid, ')),"add an additional text question below",""),IF(AND(B', rowid, '="deletion",ISNUMBER(SEARCH("Other",K', rowid, '))),"deletion an additional text question below",""))')
    other_formulas <- c(other_formulas, formula)
  }
  constraint_formulas <- c()
  last_rowid <- 2
  for (rowid in 1:999) {
    if (is.na(new.dap$`Constraint`[rowid])) {
      formula <- paste0('=IF(AND(OR(B', rowid + 1, '="new",B', rowid + 1, '="yes",B', rowid + 1, '="no"),E', rowid + 1, '="select_multiple"),IF(ISNUMBER(SEARCH("None",K', rowid + 1, ')),"not(selected(., ', "'none'", ') and (count-selected(.)>1))",""),"")')
      constraint_formulas <- c(constraint_formulas, formula)
    } else {
      if (length(constraint_formulas) > 0) {
        writeFormula(wb, sheet = "DAP__R_", x = constraint_formulas, startCol = which(colnames(new.dap) == "Constraint"), startRow = last_rowid)
      }
      constraint_formulas <- c()
      last_rowid <- rowid + 1
    }
  }

  writeFormula(wb, sheet = "DAP__R_", x = other_formulas, startCol = which(colnames(new.dap) == "Other (specify) Question"), startRow = 2)

  return(wb)
}


create.changes.dap <- function(survey.file, old_dap.file) {

  tool.survey <- utilityR::load.tool.survey(survey.file, keep_cols = TRUE)
  tool.survey <- tool.survey[!tool.survey$type %in% c("start", "end", "today", "deviceid", "audit"),]
  tool.choices <- as.data.frame(readxl::read_xlsx(survey.file, sheet = "choices"))
  tool.survey <- cast.strings(tool.survey)
  tool.choices <- cast.strings(tool.choices)
  old.dap <- as.data.frame(readxl::read_xlsx(old_dap.file, sheet = "DAP__R_"))
  # TODO: a checker for columns consistency
  old.dap <- old.dap[!is.na(old.dap$`Indicator / Variable (name)`),]
  old.dap <- dap.preparation(old.dap)
  changes.dap <- data.frame(
    `Groups` = as.character(),
    `number` = as.character(),
    `Indicator group / sector` = as.character(),
    `Question Type` = as.character(),
    `Indicator / Variable (name)` = as.character(),
    `Questionnaire Question` = as.character(),
    `Questionnaire Question RUS` = as.character(),
    `Questionnaire Question UKR` = as.character(),
    `Questionnaire Responses` = as.character(),
    `Questionnaire Responses RUS` = as.character(),
    `Questionnaire Responses UKR` = as.character(),
    `Hint` = as.character(),
    `Hint RUS` = as.character(),
    `Hint UKR` = as.character(),
    `Relevance` = as.character(),
    `Relevance_do_text` = as.character(),
    `Constraint` = as.character(),
    `Constraint_do_text` = as.character(),
    `Сonstraint message English` = as.character(),
    `Сonstraint message UKR` = as.character(),
    `Сonstraint message RUS` = as.character(),
    `Calculation` = as.character(),
    check.names = FALSE)

  for  (i in 1:nrow(tool.survey)) {
    resposes <- load.responses(utilityR::get.choice.list.from.name(tool.survey$name[i], tool.survey), tool.choices)
    is_group <- grepl("_group", tool.survey$`type`[i])
    row <- data.frame(
      `Groups` = ifelse(is_group, tool.survey$`type`[i], NA),
      `number` = tool.survey$number_indicator[i],
      `Indicator group / sector` = tool.survey$sector[i],
      `Question Type` = ifelse(is_group, NA, stringr::str_split(tool.survey$type[i], " ")[[1]][1]),
      `Indicator / Variable (name)` = tool.survey$xml[i],
      `Questionnaire Question` = tool.survey$`label::English`[i],
      `Questionnaire Question RUS` = tool.survey$`label::Russian`[i],
      `Questionnaire Question UKR` = tool.survey$`label::Ukrainian`[i],
      `Questionnaire Responses` = resposes$responses_eng,
      `Questionnaire Responses RUS` = resposes$responses_rus,
      `Questionnaire Responses UKR` = resposes$responses_ukr,
      `Hint` = tool.survey$`hint::English`[i],
      `Hint RUS` = tool.survey$`hint::Russian`[i],
      `Hint UKR` = tool.survey$`hint::Ukrainian`[i],
      `Relevance` = tool.survey$relevant[i],
      `Relevance_do_text` = tool.survey$relevant_do_text[i],
      `Constraint` = tool.survey$constraint[i],
      `Constraint_do_text` = tool.survey$constraint_do_text[i],
      `Сonstraint message English` = tool.survey$`constraint_message::English`[i],
      `Сonstraint message UKR` = tool.survey$`constraint_message::Ukrainian`[i],
      `Сonstraint message RUS` = tool.survey$`constraint_message::Russian`[i],
      `Calculation` = tool.survey$calculation[i],
      check.names = FALSE
    )
    changes <- TRUE
    for (j in 1:nrow(old.dap)) {
      old.dap.row <- old.dap[j, ]
      number = ifelse(is.na(old.dap.row$`new number`), old.dap.row$`old number`, old.dap.row$`new number`)
      old.dap.row <- old.dap.row %>%
        dplyr::select(-c("old number", "new number", "old Questionnaire Responses", "Data collection method", "Indicator group / sector", "Other (specify) Question", "change question")) %>%
        dplyr::mutate(number = number) %>%
        dplyr::select(`Groups`, `number`, everything()) %>%
        dplyr::mutate_all(list(~ ifelse(. == "", NA, .)))

      if (is.na(old.dap.row[, "Question Type"])) {
        old.dap.row$`Question Type` <- NA
      } else {
        old.dap.row$`Question Type` <- old.dap.row$`Question Type`
      }
      if (check.row.identical(row[1, ], old.dap.row[1, ])) {
        changes <- FALSE
        break
      }
    }
    if (changes) {
      changes.dap <- dplyr::bind_rows(changes.dap, row)
    }
  }
  changes.dap <- changes.dap %>%
    dplyr::select(-c("Groups"))

  wb <- openxlsx::createWorkbook("changes")

  openxlsx::addWorksheet(wb, "suggestions")

  openxlsx::writeData(wb, "suggestions", changes.dap)

  headerStyle <- createStyle(
    fontSize = 14, fontColour = "black", halign = "center",
    fgFill = "#4F81BD", borderColour = "black", border = "left,right,top,bottom"
  )
  openxlsx::addStyle(wb, sheet = 1, style = headerStyle, rows = 1, cols = 1:ncol(changes.dap))

  openxlsx::setColWidths(wb, sheet = "suggestions", cols = 1:ncol(changes.dap), widths = rep(45, ncol(changes.dap)))
  openxlsx::setRowHeights(wb, sheet = "suggestions", rows = 1, heights = 45)
  openxlsx::setRowHeights(wb, sheet = "suggestions", rows = 2:nrow(changes.dap) + 1, heights = 30)

  return(wb)
}

# create.tool("resources/dap.xlsx", "tool.xlsx")
#
# create.dap("resources/msna.xlsx", "dap_3.xlsx", "resources/generated_dap.xlsx")
#
# create.changes.dap("tool.xlsx", "dap_3.xlsx")

