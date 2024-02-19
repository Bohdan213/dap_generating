

create.survey <- function(dap.file, survey.file) {
  dap <- read_xlsx(paste0("./resources/", dap.file), sheet = "DAP__R_")
  
  dap <- dap[!is.na(dap$`change question`),]
  
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

create.survey("dap_3.xlsx", "test_tool.xlsx")