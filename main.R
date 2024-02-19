library(openxlsx)
library(readxl)
library(utilityR)


source("utils.R")

create.survey <- function(dap.file, survey.file) {
  
  dap <- as.data.frame(readxl::read_xlsx(paste0("./resources/", dap.file), sheet = "DAP__R_"))
  
  dap <- dap[!is.na(dap$`change question`),]
  dap <- dap.preparation(dap)
  saved.questions <- dap[dap$`change question` != "deletion",]
  
  ################################### create survey sheet
  
  survey <- data.frame(
    number_indicator = rep(NA, 5),
    type = c("start", "end", "today", "deviceid", "audit"),
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
    choice_filter = rep(NA, 5),
    relevant = rep(NA, 5),
    constraint = rep(NA, 5),
    `constraint_message::English` = rep(NA, 5),
    `constraint_message::Ukrainian` = rep(NA, 5),
    `constraint_message::Russian` = rep(NA, 5),
    default = rep(NA, 5),
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
      name <- paste(question_number, question$`Indicator / Variable (name)`, sep = "_")
      if (grepl("select_", question$`Question Type`)) {
        type <- paste(question$`Question Type`, question$`Indicator / Variable (name)`)
      } else {
        type <- question$`Question Type`
      }
    }
    survey <- dplyr::bind_rows(survey, data.frame(
      number_indicator = question_number,
      type = type,
      name = name,
      xml = question$`Indicator / Variable (name)`,
      `label::English` = question$`Questionnaire Question`,
      `label::Ukrainian` = question$`Questionnaire Question UKR`,
      `label::Russian` = question$`Questionnaire Question RUS`,
      `hint::English` = question$`Hint`,
      `hint::Ukrainian` = question$`Hint UKR`,
      `hint::Russian` = question$`Hint RUS`,
      relevant = question$`Relevance`,
      constraint = question$`Constraint`,
      calculation = NA,
      required = NA,
      appearance = NA,
      choice_filter = NA,
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
  
  settings <- data.frame(
    id_string = as.character(),
    form_title = as.character(),
    version = as.character(),
    default_language = as.character(),
    allow_choice_duplicates = as.character(),
    check.names = FALSE
  )
  
  ################################### combine into a doc
  wb <- createWorkbook()
  addWorksheet(wb, "survey")
  writeData(wb, "survey", survey)
  freezePane(wb, sheet = "survey", firstRow = T)
  addFilter(wb, sheet = "survey", row = 1, cols = 1:ncol(survey))
  
  addWorksheet(wb, "choices")
  writeData(wb, "choices", choices)
  
  addWorksheet(wb, "settings")
  writeData(wb, "settings", settings)
  
  headerStyle <- createStyle(
    fontSize = 12, halign = "center", 
    fgFill = "#4F81BD", borderColour = "#4F81BD",
    
  )
  
  style <- createStyle(
    fontSize = 12, halign = "left", valign = "center",
    borderColour = "#000000"
  )
  
  addStyle(wb, sheet = "survey", style = headerStyle, rows = 1, cols = 1:ncol(survey))
  addStyle(wb, sheet = "survey", style = style, rows = 2:nrow(survey), cols = 1:ncol(survey), gridExpand = T)
  
  addStyle(wb, sheet = "choices", style = headerStyle, rows = 1, cols = 1:ncol(choices))
  addStyle(wb, sheet = "choices", style = style, rows = 2:nrow(choices), cols = 1:ncol(choices), gridExpand = T)
  
  setColWidths(wb, sheet = "survey", cols = 1:ncol(survey), widths = rep(45, ncol(survey)))
  setColWidths(wb, sheet = "choices", cols = 1:ncol(choices), widths = rep(30, ncol(choices)))
  
  
  saveWorkbook(wb, paste0("resources/", survey.file), overwrite = TRUE)
  
}


create.dap <- function(survey.file, old_dap.file, new_dap.file) {
  
  tool.survey <- utilityR::load.tool.survey(paste0("./resources/", survey.file), keep_cols = TRUE)
  tool.survey <- tool.survey[!tool.survey$type %in% c("start", "end", "today", "deviceid", "audit"),]
  tool.choices <- openxlsx::read.xlsx(paste0("./resources/", survey.file), sheet = "choices")
  tool.survey <- cast.strings(tool.survey)
  # old.dap <- xlsx::read.xlsx(paste0("./resources/", old_dap.file), sheetName = "DAP__R_", keepFormulas=TRUE)
  # old.dap <- dap.preparation(old.dap)
  new.dap <- data.frame(check.names = FALSE)
  
  for (i in 1:nrow(tool.survey)) {
    resposes <- load.responses(utilityR::get.choice.list.from.name(tool.survey$name[i], tool.survey), tool.choices)
    is_group <- grepl("_group", tool.survey$`type`[i])
    new.dap <- dplyr::bind_rows(new.dap, data.frame(
      `Groups` = ifelse(is_group, tool.survey$`type`[i], NA),
      `change question` = "no",
      `old number` = tool.survey$number_indicator[i],
      `new number` = NA,
      `Question Type` = ifelse(is_group, NA, tool.survey$type[i]),
      `Indicator / Variable (name)` = tool.survey$xml[i],
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
      `Relevance` = tool.survey$relevant[i],
      `Constraint` = tool.survey$constraint[i],
      `Data collection method` = NA,
      `Indicator group / sector` = NA,
      `Other (specify) Question` = NA,
      check.names = FALSE
    ))
  }
  
  wb <- openxlsx::loadWorkbook(paste0("./resources/", old_dap.file))
  
  existing_sheets <- openxlsx::getSheetNames(paste0("./resources/", old_dap.file))
  for (sheet in existing_sheets) {
    if (sheet != "DAP__R_" & sheet != "type" & sheet != "change") {
      openxlsx::removeWorksheet(wb, sheet)
    }
  }
  
  openxlsx::writeData(wb, "DAP__R_", new.dap, startRow = 1, startCol = 1)
  
  # print(nrow(new.dap))
  # dataValidation(wb, sheet="DAP__R_", rows = 2:(nrow(new.dap) + 1),
  #                col = 2,
  #                type = 'list', value = "'change'!$B$2:$B$5")
  # 
  # dataValidation(wb, sheet="DAP__R_", rows = 2:(nrow(new.dap) + 1),
  #                col = 5,
  #                type = 'list', value = "'change'!$A$2:$A$6")
  
  formula <- '=IF(OR(B2="new";B2="yes");IF(ISNUMBER(SEARCH("Other";K2));"add an additional text question below";"");IF(AND(B2="deletion";ISNUMBER(SEARCH("Other";K2)));"deletion an additional text question below";""))'
  writeFormula(wb, sheet = "DAP__R_", x = formula, startCol = ncol(new.dap), startRow = 2)
  
  openxlsx::saveWorkbook(wb, paste0("resources/", new_dap.file), overwrite = TRUE)
}


create.changes.dap <- function(survey.file, old_dap.file) {

  tool.survey <- utilityR::load.tool.survey(paste0("./resources/", survey.file), keep_cols = TRUE)
  tool.survey <- tool.survey[!tool.survey$type %in% c("start", "end", "today", "deviceid", "audit"),]
  tool.choices <- as.data.frame(readxl::read_xlsx(paste0("./resources/", survey.file), sheet = "choices"))
  tool.survey <- cast.strings(tool.survey)
  tool.choices <- cast.strings(tool.choices)
  old.dap <- as.data.frame(readxl::read_xlsx(paste0("./resources/", old_dap.file), sheet = "DAP__R_"))
  # TODO: a checker for columns consistency
  old.dap <- old.dap[!is.na(old.dap$`Indicator / Variable (name)`),]
  old.dap <- dap.preparation(old.dap)
  changes.dap <- data.frame(check.names = FALSE)
  
  for  (i in 1:nrow(tool.survey)) {
    resposes <- load.responses(utilityR::get.choice.list.from.name(tool.survey$name[i], tool.survey), tool.choices)
    is_group <- grepl("_group", tool.survey$`type`[i])
    row <- data.frame(
      `Groups` = ifelse(is_group, tool.survey$`type`[i], NA),
      `number` = tool.survey$number_indicator[i],
      `Question Type` = ifelse(is_group, NA, tool.survey$type[i]),
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
      `Constraint` = tool.survey$constraint[i],
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
        if (grepl("select_", old.dap.row$`Question Type`)) {
          old.dap.row$`Question Type` <- paste(old.dap.row$`Question Type`, old.dap.row$`Indicator / Variable (name)`)
        }
      }
      if (check.row.identical(row[1, ], old.dap.row[1, ])) {
        changes <- FALSE
        break
      }
    }
    if (changes) {
      print("changes")
      print(row$`Indicator / Variable (name)`)
      changes.dap <- dplyr::bind_rows(changes.dap, row)
    } else {
      print("0 changes")
      print(row$`Indicator / Variable (name)`)
    }
  }
  
  changes.dap <- changes.dap %>%
    dplyr::select(-c("Groups"))
  
  # changes.dap <- dap.preparation(changes.dap)
  
  wb <- openxlsx::createWorkbook("changes")
  # create sheet suggestions
  openxlsx::addWorksheet(wb, "suggestions")
  
  openxlsx::writeData(wb, "suggestions", changes.dap)
  
  headerStyle <- createStyle(
    fontSize = 14, fontColour = "black", halign = "center",
    fgFill = "#4F81BD", borderColour = "black", border = "left,right,top,bottom"
  )
  openxlsx::addStyle(wb, sheet = 1, style = headerStyle, rows = 1, cols = 1:ncol(changes.dap))
  # openxlsx::addStyle(wb, sheet = 1, style = style, rows = 2:nrow(changes.dap), cols = 1:ncol(changes.dap), gridExpand = TRUE)
  
  openxlsx::setColWidths(wb, sheet = "suggestions", cols = 1:ncol(changes.dap), widths = rep(45, ncol(changes.dap)))
  openxlsx::setRowHeights(wb, sheet = "suggestions", rows = 1, heights = 45)
  openxlsx::setRowHeights(wb, sheet = "suggestions", rows = 2:nrow(changes.dap), heights = 30)
  
  openxlsx::saveWorkbook(wb, paste0("resources/", "changes.xlsx"), overwrite = TRUE)
}

create.survey("dap_3.xlsx", "test_tool.xlsx")

create.dap("test_tool.xlsx", "dap_3.xlsx", "new_dap.xlsx")

create.changes.dap("test_tool.xlsx", "dap_3.xlsx")
