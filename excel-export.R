#### Settings ####

#### For input
excelInFile <- "excel-in/calculations.xlsx"

# Words that whould be interpreted as functions, not variables
functionWords <- c("IF")
# Mathematical operators used to split formulas (in order to find variables)
opRegex <- "\\*|\\-|\\+|/|\\^|=|<|>|%%|\\(| |\\)|\n|,"

#### For output
scriptCommentPrefix <- "# "
cellAsCommentAfterLine <- TRUE
scriptOutFile <- "script-out/script.R"

if (!dir.exists("script-out"))
  dir.create("script-out")

#### Packages ####
if (!require("XLConnect")) {
  install.packages("XLConnect")
  require("XLConnect")
}

#### Functions ####
getSheet <- function(cell) {
  strsplit(cell, "!")[[1]][[1]]
}

getRowCol <- function(cell) {
  if (!grepl("!", cell))
    stop("Sheet name must be contained in cell identifier.")
  rowCol <- strsplit(cell, "!")[[1]][[2]]
  rowCol <- strsplit(rowCol, "$", fixed = TRUE)[[1]]
  rowCol <- rowCol[!rowCol %in% ""]
  if (length(rowCol) != 2)
    stop("Error when splitting row and col.")
  row <- as.numeric(rowCol[2])
  if (is.na(row))
    stop("Error when converting row to numeric.")
  col <- rowCol[1]
  mapping <- c(LETTERS, paste0(LETTERS[1], LETTERS), paste0(LETTERS[2], LETTERS), paste0(LETTERS[3], LETTERS))
  mapping <- data.frame(char = mapping, num = 1:length(mapping), stringsAsFactors = FALSE)
  col1 <- mapping[mapping[, "char"] == col, "num"]
  if (length(col1) == 0)
    stop("Could not find number to column: ", col)
  return(list(row = row, col = col1))
}

dropSheetFromCell <- function(cell) {
  if (!grepl("!", cell))
    stop("Sheet name must be contained in cell identifier.")
  strsplit(cell, "!")[[1]][[2]]
}

getVars <- function(form) {
  if (is.numeric(form))
    return(NULL)
  vars <- strsplit(form, opRegex)[[1]]
  vars <- unique(trimws(vars[!vars %in% unique(c("", functionWords))]))
  vars <- vars[!grepl("^[0-9]+$", vars)]
  return(vars)
}

#### Load sheets ####
calc <- XLConnect::loadWorkbook(excelInFile)

#### Get calculations ####
calcList <- list()
sheetList <- list()

allVarNames <- XLConnect::getDefinedNames(calc)
dupl <- allVarNames[duplicated(allVarNames)]
if (length(dupl) > 0)
  stop("There are duplicated variable names: ", paste0(dupl, collapse = ", "))

for (varName in allVarNames) {
  cell <- XLConnect::getReferenceFormula(calc, varName)
  sheet <- getSheet(cell)
  if (!sheet %in% names(sheetList)) {
    sheetList[[sheet]] <-
      XLConnect::readWorksheet(calc, sheet = sheet, startRow = 1, startCol = 1, autofitRow = FALSE, autofitCol = FALSE, header = FALSE)
  }
  rowCol <- getRowCol(cell)


  form <- tryCatch({
    XLConnect::getCellFormula(calc, sheet = sheet, row = rowCol$row, col = rowCol$col)
  }, error = function(x) {
    sheetList[[sheet]][rowCol$row, rowCol$col]
  })
  form <- unname(form)

  formVars <- getVars(form)

  calcList[[length(calcList) + 1]] <- list(
    var = varName,
    sheet = sheet,
    cell = cell,
    form = form,
    formVars = formVars
  )
}

names(calcList) <- allVarNames

#### Find out correct ordering ####
calcList2 <- calcList

dropStage <- 1
lengthList <- length(calcList2)
while(length(calcList2) > 0) {
  varsToDrop <- unlist(sapply(calcList2, function(x) {
    if (length(x[["formVars"]]) == 0)
      return(x[["var"]])
    NULL
  }))

  for (var in varsToDrop) {
    calcList[[var]][["dropStage"]] <- dropStage
  }

  calcList2 <- calcList2[!names(calcList2) %in% varsToDrop]

  if (length(calcList2) == 1)
    break()

  if (length(calcList2) == length(lengthList)) { # length(calcList2) != 1
    print(calcList2)
    stop("The length of the list has not changed. Dead lock?")
  }
  lengthList <- length(calcList2)

  calcList2 <- lapply(calcList2, function(x) {
    x[["formVars"]] <- x[["formVars"]][!x[["formVars"]] %in% varsToDrop]
    x
  })

  dropStage <- dropStage + 1
}

# When only 1 var is left in the end, it can happen that no drop stage has been applied.
# Therefore, apply now.
calcList <- lapply(calcList, function(x) {
  if (is.null(x[["dropStage"]]))
    x[["dropStage"]] <- dropStage + 1
  return(x)
})


#### Apply ordering ####

calcDf <- do.call("rbind", lapply(calcList, function(x) {
  x <- x[!names(x) %in% "formVars"]
  data.frame(x)
}))

calcDf <- calcDf[order(calcDf[, "dropStage"], calcDf[, "sheet"]), ]

script <- character()

lastSheet <- "wefplxwerplwef"
for (i in 1:nrow(calcDf)) {
  if (calcDf[i, "sheet"] != lastSheet) {
    script[length(script) + 1] <- ""
    script[length(script) + 1] <- paste0(scriptCommentPrefix, calcDf[i, "sheet"])
    lastSheet <- calcDf[i, "sheet"]
  }
  script[length(script) + 1] <- paste0(calcDf[i, "var"], " = ", calcDf[i, "form"])
  if (cellAsCommentAfterLine) {
    script[length(script)] <- paste0(script[length(script)], " ", scriptCommentPrefix, dropSheetFromCell(calcDf[i, "cell"]))
  }
}

message("*** Script start ***")
invisible(lapply(script, function(line) cat(line, "\n", sep = "")))
message("*** Script end ***")
writeLines(script, scriptOutFile, useBytes = TRUE)


