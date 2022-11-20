#### Settings ####

excelInFile <- "excel-in/calculations.xlsx"
scriptOutFile <- "script-out/script.R"

# Comment character for the script that will be created
scriptCommentPrefix <- "# "
# Add cell in which the variable was defined as comment to script?
addCellAsComment <- TRUE
# Add comments from excel to script?
addCommentAsComment <- TRUE
# Add content of cell left to variable as comment to script?
addLeftCellValueAsComment <- TRUE
# Add content of cell right to variable as comment to script?
addRightCellValueAsComment <- TRUE
# If there are no real variable names in the excel sheet, then you can choose to get the variable names from the left cell.
getVarnamesFromLeftCell <- TRUE
# If no name in the left cell was found, how should the the prefix of the artificial variable name given?
varPrefix <- "VAR_"

# Mathematical operators used to split formulas (in order to find variables)
opRegex <- "\\*|\\-|\\+|/|\\^|=|<|>|%%|\\(| |\\)|\n|,"
# Words that whould be interpreted as functions, not variables are filtered out before variable detection.
preVarDetectFuncTransformer <- function(formula) {
  formula <- gsub("[a-zA-Z0-9_]+\\(", "", formula)
  return(formula)
}
# Funciton to change functions from excel to other language
functionTransformer <- function(formula) {
  formula <- gsub("IF(", "ifelse(", formula, fixed = TRUE)
  return(formula)
}



if (!dir.exists("script-out"))
  dir.create("script-out")

#### Packages ####
if (!"XLConnect" %in% .packages(all.available = TRUE))
  install.packages("XLConnect")

#### Functions ####
# Prerequisite constants for functions
colCharNumbMapping <- c(LETTERS, paste0(LETTERS[1], LETTERS), paste0(LETTERS[2], LETTERS), paste0(LETTERS[3], LETTERS), paste0(LETTERS[4], LETTERS), paste0(LETTERS[5], LETTERS))
colCharNumbMapping <- data.frame(char = colCharNumbMapping, num = 1:length(colCharNumbMapping), stringsAsFactors = FALSE)
rownames(colCharNumbMapping) <- colCharNumbMapping[, "char"]
colNumbCharMapping <- colCharNumbMapping
rownames(colNumbCharMapping) <- colNumbCharMapping[, "num"]

# getSheet("sheet1!$A$1")
getSheet <- function(cell) {
  strsplit(cell, "!")[[1]][[1]]
}

# getVars("a + 2 + b")
getVars <- function(form, keepNumbers = FALSE) {
  if (is.numeric(form))
    return(NULL)
  form <- preVarDetectFuncTransformer(form)
  vars <- strsplit(form, opRegex)[[1]]
  if (!keepNumbers)
    vars <- vars[!grepl("^[0-9]+$", vars)]
  return(vars)
}

# getRowCol("sheet1!$A$1")
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
  col1 <- colCharNumbMapping[col, "num"]
  if (length(col1) == 0 || is.na(col1))
    stop("Could not find number to column: ", col)
  return(list(row = row, col = col1))
}

# cellToChar(1, 1, TRUE)
cellToChar <- function(row, col, dollar = TRUE) {
  if (is.na(colNumbCharMapping[col, "char"]))
    stop("Cannot convert to character-like cell identifiert.")
  col <-colNumbCharMapping[col, "char"]
  if (dollar) {
    col <- paste0("$", col)
    row <- paste0("$", row)
  }
  return(paste0(col, row))
}

# fullyQualifiedCell(c("A1", 1), "sheet1")
fullyQualifiedCell <- function(cell, sheet) {
  for (i in 1:length(cell)) {
    if (grepl("^[0-9]$", cell[i]))
      next
    if (!grepl("$", cell[i], fixed = TRUE)) {
      col <- gsub("[0-9]+", "", cell[i])
      row <- gsub("[A-Z]+", "", cell[i])
      cell[i] <- paste0("$", col, "$", row)
    }
    if (!grepl("!", cell[i]))
      cell[i] <- paste0(sheet, "!", cell[i])
  }
  return(cell)
}

# fullyQualifiedFormula("A1+1", "sheet1")
fullyQualifiedFormula <- function(form, sheet) {
  if (!grepl("[A-Z]+[0-9]+", form))
    return(form)
  formVars <- getVars(form, keepNumbers = TRUE)
  if (length(formVars) > 0)
    for (i in 1:length(formVars)) {
      if (grepl("^[A-Z]+[0-9]+$", formVars[i])) {
        form <- gsub(formVars[i], fullyQualifiedCell(formVars[i], sheet = sheet), form, fixed = TRUE)
      }
    }
  return(form)
}

# sanitizeCellName("sheet1!$A$1", dropSheet = TRUE)
sanitizeCellName <- function(cell, dropSheet = FALSE) {
  if (dropSheet && grepl("!", cell)) {
    cell <- strsplit(cell, "!")[[1]][[2]]
  }
  cell <- gsub("$", "", cell, fixed = TRUE)
  return(cell)
}

# Extract all comments form an Excel Workbook
extractAllComments <- function(file) {
  if (!"openxlsx" %in% .packages(all.available = TRUE))
    install.packages("openxlsx")
  sheetNames <- openxlsx::sheets(openxlsx::loadWorkbook(excelInFile))

  exdir <- tempdir()
  utils::unzip(file, exdir = exdir)
  files <- list.files(paste0(exdir, "/xl"), pattern = "^comment", full.names = TRUE)
  #fileNo <- gsub("^comments|\\.xml$", "", basename(files))
  if (length(sheetNames) != length(files))
    stop("Each sheet must contain at least one comment, otherwise the extraction of comments does not work.")

  comments <- lapply(files, function(x) {
    txt <- paste0(suppressWarnings(readLines(x, encoding = "UTF-8")), collapse = "\n")
    xml <- xml2::read_xml(txt, encoding = "UTF-8")
    xml <- xml2::as_list(xml)[["comments"]][["commentList"]]
    out <- lapply(xml, function(y) {
      data.frame(cell = attr(y, "ref"),
                 value = y[["text"]][["r"]][["t"]][[1]])
    })
    do.call("rbind", out)
  })

  suppressWarnings(file.remove(files))
  suppressWarnings(unlink(exdir))

  names(comments) <- sheetNames
  for (sheet in names(comments)) {
    comments[[sheet]][, "sheet_cell"] <- paste0(sheet, "!", comments[[sheet]][, "cell"])
  }
  comments <- do.call("rbind", comments)
  rownames(comments) <- comments[, "sheet_cell"]
  return(comments)
}

#' Make all entires in vector the same length.
#' @keywords internal
#' @author Daniel Hoop
#' @param add The sign that will be added.
#' @param where Character indicating if `add` should be added at the beginning or at the end of the string.
#' @examples
#' x <- data.frame(a=seq(0.1, 1, 0.1) , b=c(1:10))
#' equal.length(equal.n.decimals(x))
equal.length <- function(x, add=0, where=c("beginning","end"), minlength=0, margin=2) {
  if(!is.null(dim(x))) {
    if(is.matrix(x)) return(apply(x,margin,function(x)equal.length(x=x, add=add, where=where, minlength=minlength, margin=margin)))
    if(is.data.frame(x)) return( as.data.frame( lapply(x,function(x)equal.length(x=x, add=add, where=where, minlength=minlength, margin=margin)) ,stringsAsFactors=FALSE) )
  }
  if(is.list(x)) return(lapply(x,function(x)equal.length(x=x, add=add, where=where, margin=margin)))

  where <- match.arg(where)
  nchar.x <- nchar(x)
  nchar.x[is.na(nchar.x)] <- 0
  n.add <- max(minlength, nchar.x)  - nchar.x
  x.new <- character()
  if(where=="beginning") {
    for(i in sort(unique(n.add))) x.new[n.add==i] <- paste0(paste0(rep(add,i),collapse=""), x[n.add==i])
  } else if(where=="end") {
    for(i in sort(unique(n.add))) x.new[n.add==i] <- paste0(x[n.add==i], paste0(rep(add,i),collapse="") )
  }
  return(x.new)
}

#### Load sheets ####
calc <- XLConnect::loadWorkbook(excelInFile)

if (addCommentAsComment)
  comments <- extractAllComments(excelInFile)

#### Preparations to get calculations ####
calcList <- list()
sheetList <- list()
varCellsList1 <- list()
varCellsList2 <- list()
varCounter <- 1

varCells <- XLConnect::getDefinedNames(calc)
dupl <- varCells[duplicated(varCells)]
if (length(dupl) > 0)
  stop("There are duplicated variable names: ", paste0(dupl, collapse = ", "))

#### Get calculations for defined variables ####
for (varName in varCells) {
  cell <- unname(XLConnect::getReferenceFormula(calc, varName))
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
  form <- functionTransformer(form)
  form <- fullyQualifiedFormula(form, sheet = sheet)

  varCellsList1[[length(varCellsList1) + 1]] <- list(
    form = form,
    sheet = sheet,
    name = varName,
    cell = cell)
  varCellsList1[[length(varCellsList1)]] <- as.data.frame(varCellsList1[[length(varCellsList1)]], stringsAsFactors = FALSE)
}
if (length(varCellsList1) > 0) {
  varCellsList1 <- do.call("rbind", varCellsList1)
  for (i in 1:nrow(varCellsList1)) {
    varCellsList1[, "form"] <- gsub(varCellsList1[i, "cell"], varCellsList1[i, "name"], varCellsList1[, "form"], fixed = TRUE)
  }
}
rm(varCells)


#### Get calculations for variables with name to the left of the cell ####

if (getVarnamesFromLeftCell) {
  for (sheet in XLConnect::getSheets(calc)) {
    if (!sheet %in% names(sheetList)) {
      sheetList[[sheet]] <-
        XLConnect::readWorksheet(calc, sheet = sheet, startRow = 1, startCol = 1, autofitRow = FALSE, autofitCol = FALSE, header = FALSE)
    }
    calcDat <- sheetList[[sheet]]
    # Loop through all cells of sheet
    for (row in 1:nrow(calcDat)) {
      for (col in 1:ncol(calcDat)) {
        if (is.na(calcDat[row, col]))
          next
        form <- try(XLConnect::getCellFormula(calc, sheet = sheet, row = row, col = col), silent = TRUE)
        if (class(form) %in% "try-error") {
          form <- sheetList[[sheet]][row, col]
          # Only numeric values are taken, else all characters would be imported as variables as well.
          if (!is.numeric(form))
            next
        }
        form <- functionTransformer(form)
        form <- fullyQualifiedFormula(form, sheet = sheet)
        # If there is a name to the left, add this name
        if (col > 1 && !is.na(sheetList[[sheet]][row, col - 1])) {
          name1 <- sheetList[[sheet]][row, col - 1]
          name <- gsub("[^a-zA-Z0-9]", "_", name1)
          # If the name in the cell to the left is totally useless, create a new one.
          if (grepl("^_+$", name) && name1 != name) {
            name <- paste0(varPrefix, varCounter)
            varCounter <- varCounter + 1
          }
        } else {
          # If there is no name, make up a new name.
          name <- paste0(varPrefix, varCounter)
          varCounter <- varCounter + 1
        }
        varCellsList2[[length(varCellsList2) + 1]] <- list(
          form = form,
          sheet = sheet,
          name = name,
          cell = paste0(sheet, "!", cellToChar(row = row, col = col, dollar = TRUE)))
        varCellsList2[[length(varCellsList2)]] <- as.data.frame(varCellsList2[[length(varCellsList2)]], stringsAsFactors = FALSE)
        #names(varCellsList2)[length(varCellsList2)] <- name
      }
    }
  }
  varCellsList2 <- do.call("rbind", varCellsList2)

  # If some variables were found...
  if (length(varCellsList1) > 0) {
    for (i in 1:nrow(varCellsList2)) {
      # If the variable which name was guessed was acually defined as real variable name in the workbook, then take that one.
      if (!grepl("^[0-9]+$", varCellsList2[i, "form"]) &&
          varCellsList2[i, "form"] %in% varCellsList1[, "form"]) {
        varCellsList2[i, "name"] <- varCellsList1[, "name"][which(varCellsList1[, "form"] == varCellsList2[i, "form"])[1]]
      }
    }
    # Loop through all cell names in varCellsList2 and update cell references in formulas of varCellsList1 and varCellsList2
    for (i in 1:nrow(varCellsList2)) {
      varCellsList1[, "form"] <- gsub(varCellsList2[i, "cell"], varCellsList2[i, "name"], varCellsList1[, "form"], fixed = TRUE)
      varCellsList2[, "form"] <- gsub(varCellsList2[i, "cell"], varCellsList2[i, "name"], varCellsList2[, "form"], fixed = TRUE)
    }
    # Loop through all cell names in varCellsList1 and update cell references in formulas of varCellsList1 and varCellsList2
    if (nrow(varCellsList1) > 0) {
      for (i in 1:nrow(varCellsList1)) {
        varCellsList1[, "form"] <- gsub(varCellsList1[i, "cell"], varCellsList1[i, "name"], varCellsList1[, "form"], fixed = TRUE)
        varCellsList2[, "form"] <- gsub(varCellsList1[i, "cell"], varCellsList1[i, "name"], varCellsList2[, "form"], fixed = TRUE)
      }
      # Drop all variables of which the name was guessed, if that variable was defined as real variable in the workbook
      varCellsList2 <- varCellsList2[!varCellsList2[, "name"] %in% varCellsList1[, "name"],]
    }
  }
}

#### Combine variables of which the names were defined and guessed ####
varCellsList <- rbind(varCellsList1, varCellsList2)
if (nrow(varCellsList) == 0)
  stop("No variables found.")

# Throw error if duplicated variable names were created.
dupl <- duplicated(varCellsList[, "name"])
if (sum(dupl) > 0) {
  dupl <- duplicated(varCellsList[, "name"]) | duplicated(varCellsList[, "name"], fromLast = TRUE)
  print(varCellsList[dupl, "name"])
  stop("There are duplicated variable names")
}

# Add all variables from calculation as additional list place into intermediate result.
rownames(varCellsList) <- varCellsList[, "name"]

for (varName in rownames(varCellsList)) {
  rowCol <- getRowCol(varCellsList[varName, "cell"])
  calcList[[length(calcList) + 1]] <- list(
    var = varName,
    sheet = varCellsList[varName, "sheet"],
    cell = varCellsList[varName, "cell"],
    row = rowCol$row,
    col = rowCol$col,
    form = varCellsList[varName, "form"],
    formVars = getVars(varCellsList[varName, "form"])
  )
}

names(calcList) <- rownames(varCellsList)

#### Find out correct ordering ####
calcList2 <- calcList

dropStage <- 1
nRuns <- 0
lengthList <- length(calcList2)
while(length(calcList2) > 0) {
  nRuns <- nRuns + 1
  if (nRuns >= 100000)
    stop("Infinite loop. Detecting the ordering of the variables was not possible.")

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

oldList <- NULL
while(!assertthat::are_equal(oldList, calcList)) {
  oldList <- calcList
  calcList <- lapply(calcList, function(x) {
    x[["formVars"]] <- sort(unique(c(x[["formVars"]], unlist(lapply(calcList[x[["formVars"]]], function(y) y[["formVars"]])))))
    x
  })
}

# Get sheet dependency ordering. This will be necessary to have the right sheet block ordering in the end.
sheetOrdering <- do.call("rbind", lapply(calcList, function(x) data.frame(x[c("sheet", "dropStage")])))
sheetOrdering <- tapply(sheetOrdering[, "dropStage"], sheetOrdering[, "sheet"], max)
#sheetOrdering[] <- equal.length(x = sheetOrdering, add = "0", where = "beginning")

calcList <- lapply(calcList, function(x) {
  x[["dependsOnSheet"]] <- sort(unique(c(x[["sheet"]], unlist(sapply(calcList[x[["formVars"]]], function(y) y[["sheet"]])))))
  return(x)
})
calcList <- lapply(calcList, function(x) {
  x[["sheetOrder"]] <- unname(max(sheetOrdering[x[["dependsOnSheet"]]]))
  return(x)
})


#### Apply ordering ####

calcDf <- do.call("rbind", lapply(calcList, function(x) {
  x[["dependsOnSheet"]] <- paste0(x[["dependsOnSheet"]], collapse = ", ")
  x <- x[!names(x) %in% "formVars"]
  data.frame(x)
}))

calcDf <- calcDf[order(calcDf[, "sheetOrder"], nchar(calcDf[, "dependsOnSheet"]), calcDf[, "dropStage"]), ]
script <- character()

lastSheet <- "wefplxwerplwef"
for (i in 1:nrow(calcDf)) {
  if (calcDf[i, "dependsOnSheet"] != lastSheet) {
    script[length(script) + 1] <- ""
    script[length(script) + 1] <- paste0(scriptCommentPrefix, calcDf[i, "dependsOnSheet"])
    lastSheet <- calcDf[i, "dependsOnSheet"]
  }
  script[length(script) + 1] <- paste0(calcDf[i, "var"], " = ", functionTransformer(calcDf[i, "form"]))
  if (sum(addCellAsComment, addCommentAsComment, addLeftCellValueAsComment, addRightCellValueAsComment) > 0) {
    script[length(script)] <- paste0(script[length(script)], " ", scriptCommentPrefix)
    sanitizedCellname <- sanitizeCellName(calcDf[i, "cell"])
  }
  if (addCellAsComment)
    script[length(script)] <- paste0(script[length(script)], sanitizedCellname, " | ")
  if (addCommentAsComment) {
    addVal <- if (sanitizedCellname %in% rownames(comments))
      comments[sanitizedCellname, "value"] else
        ""
    script[length(script)] <- paste0(script[length(script)], addVal, " | ")
  }
  if (addLeftCellValueAsComment) {
    addVal <- if (calcDf[i, "col"] > 1)
      sheetList[[calcDf[i, "sheet"]]][calcDf[i, "row"], calcDf[i, "col"] - 1] else
        ""
    if (is.na(addVal))
      addVal <- ""
    script[length(script)] <- paste0(script[length(script)], addVal, " | ")
  }
  # Will be NULL if i is out of range.
  if (addRightCellValueAsComment)
    script[length(script)] <- paste0(script[length(script)], sheetList[[calcDf[i, "sheet"]]][calcDf[i, "row"], calcDf[i, "col"] + 1], " | ")

  script[length(script)] <- gsub("|[| ]* *$", "", script[length(script)], fixed = FALSE)
  script[length(script)] <- trimws(script[length(script)])
}

{
  message("*** Script start ***")
  invisible(lapply(script, function(line) cat(line, "\n", sep = "")))
  message("\n*** Script end ***")
}
writeLines(script, scriptOutFile, useBytes = TRUE)


