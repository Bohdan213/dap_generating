source("utils.R")

create.dap <- function(survey.file, old_dap.file, new_dap.file) {
  
  
  tool.survey <- utilityR::load.tool.survey(paste0("./resources/", survey.file), keep_cols = TRUE)
  tool.survey <- tool.survey[!tool.survey$type %in% c("start", "end", "today", "deviceid", "audit"),]
  tool.choices <- openxlsx::read.xlsx(paste0("./resources/", survey.file), sheet = "choices")
  
  old.dap <- xlsx::read.xlsx(paste0("./resources/", old_dap.file), sheetName = "DAP__R_", keepFormulas=TRUE)
  
  new.dap <- data.frame(check.names = FALSE)
  
  
  for (i in 1:nrow(tool.survey)) {
    
    resposes <- load.responses(utilityR::get.choice.list.from.name(tool.survey$name[i], tool.survey), tool.choices)
    is_group <- grepl("_group", tool.survey$`type`[i])
    new.dap <- dplyr::bind_rows(new.dap, data.frame(
      `Groups` = ifelse(is_group, tool.survey$`type`[i], NA),
      `change question` = "no",
      `old number` = tool.survey$number_indicator[i],
      `new number` = NA,
      `Question Type` = tool.survey$type[i],
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
      `Relevance` = ifelse(is.na(tool.survey$relevant[i]), old.dap$Relevance[old.dap$`Indicator / Variable (name)` == tool.survey$xml[i]], tool.survey$relevant[i]),
      `Constraint` = tool.survey$constraint[i],
      `Data collection method` = NA,
      `Indicator group / sector` = NA,
      `Other (specify) Question` = NA,
      check.names = FALSE
    ))
  }
  
  wb <- openxlsx::loadWorkbook(paste0("./resources/", old_dap.file))
  deleteData(wb, "DAP__R_", cols=1:ncol(old.dap), rows=1:nrow(old.dap), gridExpand = TRUE)
  
  existing_sheets <- openxlsx::getSheetNames(paste0("./resources/", old_dap.file))
  for (sheet in existing_sheets) {
    if (sheet != "DAP__R_") {
      openxlsx::removeWorksheet(wb, sheet)
    }
  }
  openxlsx::writeData(wb, "DAP__R_", new.dap, startRow = 1, startCol = 1)
  
  openxlsx::saveWorkbook(wb, paste0("resources/", new_dap.file), overwrite = TRUE)
}

create.dap("test_tool.xlsx", "dap_3.xlsx", "new_dap.xlsx")

