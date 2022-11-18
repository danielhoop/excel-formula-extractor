#### Settings ####

#### For input
excelInFile <- "excel-in/calculations.xlsx"

# Words that whould be interpreted as functions, not variables
functionWords <- c("IF")
# Mathematical operators used to split formulas (in order to find variables)
opRegex <- "\\*|\\-|\\+|/|\\^|=|<|>|%%|\\(| |\\)|\n|,"
# Add comments from excel to script?
addCommentAsComment <- TRUE
# Add content of cell left to variable as comment to script?
addLeftCellValueAsComment <- TRUE
# Add content of cell right to variable as comment to script?
addRightCellValueAsComment <- TRUE

#### For output
scriptCommentPrefix <- "# "
cellAsCommentAfterLine <- TRUE
scriptOutFile <- "script-out/script.R"

if (!dir.exists("script-out"))
  dir.create("script-out")

#### Packages ####
if (!"XLConnect" %in% .packages(all.available = TRUE))
  install.packages("XLConnect")

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

sanitizeCellName <- function(cell, scriptNamePrefix = NULL, dropSheet = FALSE) {
  if (dropSheet && grepl("!", cell)) {
    cell <- strsplit(cell, "!")[[1]][[2]]
  }
  cell <- gsub("$", "", cell, fixed = TRUE)
  if (!is.null(scriptNamePrefix))
    cell <- gsub(scriptNamePrefix, "", cell, fixed = FALSE)
  return(cell)
}

getVars <- function(form) {
  if (is.numeric(form))
    return(NULL)
  vars <- strsplit(form, opRegex)[[1]]
  vars <- unique(trimws(vars[!vars %in% unique(c("", functionWords))]))
  vars <- vars[!grepl("^[0-9]+$", vars)]
  return(vars)
}

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
    row = rowCol$row,
    col = rowCol$col,
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
    script[length(script) + 1] <- paste0(scriptCommentPrefix, calcDf[i, "dependsOnSheet"]) # sanitizeCellName(calcDf[i, "sheet"], scriptNamePrefix = paste0(scriptNamePrefix, "[0-9]* ")))
    lastSheet <- calcDf[i, "dependsOnSheet"]
  }
  script[length(script) + 1] <- paste0(calcDf[i, "var"], " = ", calcDf[i, "form"])
  if (sum(cellAsCommentAfterLine, addCommentAsComment, addLeftCellValueAsComment, addRightCellValueAsComment) > 0) {
    script[length(script)] <- paste0(script[length(script)], " ", scriptCommentPrefix)
    sanitizedCellname <- sanitizeCellName(calcDf[i, "cell"])
  }
  if (cellAsCommentAfterLine)
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
    script[length(script)] <- paste0(script[length(script)], addVal, " | ")
  }
  if (addRightCellValueAsComment)
    script[length(script)] <- paste0(script[length(script)], sheetList[[calcDf[i, "sheet"]]][calcDf[i, "row"], calcDf[i, "col"] + 1], " | ")
}

{
  message("*** Script start ***")
  invisible(lapply(script, function(line) cat(line, "\n", sep = "")))
  message("\n*** Script end ***")
}
writeLines(script, scriptOutFile, useBytes = TRUE)


