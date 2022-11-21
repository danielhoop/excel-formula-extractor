#### Settings ####

excelInFile <- "excel-in/calculations.xlsx"
scriptOutFile <- "script-out/script.R"
lookupTableOutDir <- "script-out"
logOutFile <- "script-out/log.txt"

# Skip sheet with this pattern
skipSheetPattern <- "^lookup_"
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
# If there are no real variable names in the excel sheet, or they are not defined for all cells,
# then you can choose to get the variable names from the left cell.
getVarnamesFromLeftCell <- TRUE
# If no name in the left cell was found, how should the the prefix of the artificial variable name given?
varPrefix <- "VAR_"

# Mathematical operators used to split formulas (in order to find variables)
opRegex <- "\\*|\\-|\\+|/|\\^|=|<|>|%%|\\(| |\\)|\\[|\\]|\n|,"

# Words that whould be interpreted as functions, not variables and therefore are filtered out before variable detection.
preVarDetectFuncTransformer <- function(formula) {
  formula <- gsub("[a-zA-Z0-9_]+\\(", "", formula) # Function calls
  formula <- gsub("[a-zA-Z0-9_]+\\[", "", formula) # Accessing matrix/data.frame
  formula <- gsub("^TRUE$|^FALSE$", "", formula) # Logical values -> Not variables.
  return(formula)
}

# Funciton to change functions from excel to other language
# formula <- "HLOOKUP(3,$I$6:$L$10,2,FALSE) + HLOOKUP(3,$I$6:$L$10,2,FALSE)"
functionTransformer <- function(formula, sheet, wb = NULL) {
  # Add simple transformations here, such as IF -> ifelse, MAX -> max, etc.
  formula <- gsub("IF(", "ifelse(", formula, fixed = TRUE)
  formula <- gsub("MAX(", "max(", formula, fixed = TRUE)
  formula <- gsub("MIN(", "min(", formula, fixed = TRUE)

  # VLOOKUP and HLOOKUP transformation.
  if (grepl("[VH]LOOKUP+\\(", formula)) {
    for (i in (nchar(formula)-nchar("HLOOKUP")):1) {
      if (substr(formula, i, i + 7) %in% c("HLOOKUP(", "VLOOKUP(")) {
        # Split up the arguments in call to the lookup function.
        HV <- substr(formula, i, i)
        form1 <- substring(formula, i + 8, nchar(formula))
        formSplit <- strsplit(form1, ",")[[1]]
        searchForOrig <- formSplit[1]
        rangeOrig <- formSplit[2]
        takeWhich <- as.numeric(formSplit[3])
        exactMatch <- strsplit(formSplit[4], "\\)")[[1]][1]

        # When a column lookup is needed and the search key is numeric, it has to be converted by
        # adding a "X" in the front. Because R will add this automatically to numeric column names.
        searchFor <- searchForOrig
        if (HV == "H" && grepl("^[0-9]+$", searchForOrig))
          searchFor <- paste0("X", searchForOrig)
        range <- fullyQualifiedCell(rangeOrig, sheet = sheet)

        # If this cell range has never been used before, then save it and add
        # it to pre script block where the lookup table will be read in.
        fileOrObj <- paste0(HV, "_", gsub("[^a-zA-Z0-9_]", "_", gsub("$", "", range, fixed = TRUE)))
        .keyValueStore$getOrSet("lookup_tables", list())
        if (!fileOrObj %in% names(.keyValueStore$get("lookup_tables"))) {
          # Add to lookup ranges. Variables that are contained in these, are kicked later.
          .keyValueStore$append("lookup_ranges", expandRange(range))
          # Lookup table preparations and saving.
          tabList <- .keyValueStore$get("lookup_tables")
          filePath <- paste0(lookupTableOutDir, "/", fileOrObj, ".csv")
          tabList[[fileOrObj]] <- list(
            range = range,
            file = filePath
          )
          .keyValueStore$set("lookup_tables", tabList)
          saveRegionToFile(wb = wb, range = range, file = filePath)

          # If the script block is empty, add a function to make values numeric, if possible.
          # It is called 'tryNum'.
          scriptBlock <- .keyValueStore$getOrSet("pre_script_block", "")
          if (scriptBlock == "") {
            addFunc <- deparse(function(x) {
              if (is.numeric(x))
                return(x)
              out <- suppressWarnings(as.numeric(x))
              if (is.na(out))
                return(x)
              return(out)
            })
            scriptBlock <- paste0(
              "#### pre script block ####\n\n",
              "tryNum <- ", paste0(addFunc, collapse = "\n"),
              "\n")
          }

          if (HV == "H") {
            # When horizontal lookup, read in colnames.
            scriptBlock <- paste0(scriptBlock, '\n', fileOrObj, ' <- read.table("', filePath, '", sep=";", header=TRUE, stringsAsFactors=FALSE, quote = "\\\"", comment.char="", na.strings=c(""))')
          } else {
            # When vertical lookup, read in rownames.
            scriptBlock <- paste0(scriptBlock, '\n', fileOrObj, ' <- read.table("', filePath, '", sep=";", header=FALSE, row.names = 1, stringsAsFactors=FALSE, quote = "\\\"", comment.char="", na.strings=c(""))')
          }

          .keyValueStore$set("pre_script_block", scriptBlock)

        }

        # In the formula, replace the lookup with a read out from a matrix.¨
        if (isNum(searchFor) || searchForOrig != searchFor)
          searchFor <- paste0('"', searchFor, '"')
        lookupSearch <- paste0(HV, "LOOKUP(", searchForOrig, ",", rangeOrig, ",", takeWhich, ",", exactMatch, ")")
        lookupReplace <- paste0("tryNum(", fileOrObj, if (HV == "H") paste0('[', takeWhich-1, ',', searchFor, ']') else paste0('[', searchFor, ', ', takeWhich-1, ']'), ")")
        formulaOrig <- formula
        formula <- gsub(lookupSearch, lookupReplace, formula, fixed = TRUE)
        #if (searchFor == "r_") browser()
        if (formula == formulaOrig)
          warning("Conversion of VLOOKUP/HLOOKUP (SVERWEIS/WVERWEIS) has probably not worked for: ", lookupSearch)
        # In case double quotes have been produced, reduce them to single ones.
        formula <- gsub('"+', '"', formula)
      }
    }
  }

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

.keyValueStoreEnvir <- new.env()
.keyValueStore <- list(
  initialize = function() {
  }
  ,contains = function(key) {
    return (key %in% base::ls(envir = .keyValueStoreEnvir))
  }
  ,setAndReturn = function(key, value) {
    base::assign(key, value, envir = .keyValueStoreEnvir)
    return (value)
  }
  ,get = function(key) {
    return (base::get(key, envir = .keyValueStoreEnvir))
  }
  # The only function of .KeyValueStore that should be called is "getOrSet". All other functions are helper functions.
  ,getOrSet = function(key, value) {
    if (.keyValueStore$contains(key))
      return (.keyValueStore$get(key))
    return (.keyValueStore$setAndReturn(key, value))
  }
  ,remove = function(key) {
    if (.keyValueStore$contains(key))
      base::rm(list = key, envir = .keyValueStoreEnvir)
  }
  ,append = function(key, value) {
    if (!.keyValueStore$contains(key))
      return (.keyValueStore$set(key, value))
    invisible(.keyValueStore$set(key, c(.keyValueStore$get(key), value)))
  }
)

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
  vars <- vars[!vars %in% ""]
  vars <- vars[!grepl('\"', vars)]
  if (!keepNumbers)
    vars <- vars[!grepl("^[0-9]+$", vars)]
  return(vars)
}

# getRowCol("sheet1!$A$1")
getRowCol <- function(cell) {
  if (grepl("!", cell)) {
    rowCol <- strsplit(cell, "!")[[1]][[2]]
  } else {
    rowCol <- cell
  }
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
cellToChar <- function(sheet = NULL, row, col, dollar = TRUE) {
  if (is.na(colNumbCharMapping[col, "char"]))
    stop("Cannot convert to character-like cell identifiert.")
  col <-colNumbCharMapping[col, "char"]
  if (dollar) {
    col <- paste0("$", col)
    row <- paste0("$", row)
  }
  if (!is.null(sheet))
    sheet <- paste0(sheet, "!")
  return(paste0(sheet, col, row))
}

# fullyQualifiedCell(c("A1", 1), "sheet1")
# fullyQualifiedCell(c("$I$6:$L$10"), "sheet1")
fullyQualifiedCell <- function(cell, sheet) {
  for (i in 1:length(cell)) {
    if (grepl("^[0-9]$", cell[i]))
      next
    if (grepl(":", cell[i])) {
      cell1 <- strsplit(cell[i], ":")[[1]]
      cell1 <- fullyQualifiedCell(cell1, sheet = NULL)
      cell[i] <- paste0(cell1[1], ":", cell1[2])
    }
    if (!grepl("$", cell[i], fixed = TRUE)) {
      col <- gsub("[0-9]+", "", cell[i])
      row <- gsub("[A-Z]+", "", cell[i])
      cell[i] <- paste0("$", col, "$", row)
    }
    if (!is.null(sheet) && !grepl("!", cell[i]))
      cell[i] <- paste0(sheet, "!", cell[i])
  }
  return(cell)
}

# isCellNotVariable(cell = c("asdf12", "sheet!$A$1", "BB99"))
isCellNotVariable <- function(cell) {
  if (length(cell) == 0)
    return(cell)
  return(grepl("!|\\$", cell) | grepl("^[A-Z]+[0-9]+$", cell))
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

# expandRange("calculations1!$I$6:$L$10")
expandRange <- function(range) {
  sheet <- ""

  if (grepl("!", range)) {
    range <- strsplit(range, "!")[[1]]
    sheet <- range[1]
    rest <- range[2]
  } else {
    rest <- range
  }
  rest <- strsplit(rest, ":")[[1]]

  if (length(rest) == 1)
    return(fullyQualifiedCell(rest, sheet))
  col <- gsub("\\$|[0-9]*", "", rest)
  row <- as.numeric(gsub("\\$|[A-Z]*", "", rest))
  col <- try(colCharNumbMapping[which(colCharNumbMapping[, "char"] == col[1]):which(colCharNumbMapping[, "char"] == col[2]) , "char"])
  if (class(col) == "try-error")
    stop("Expansion of column range did not work. Probably the `colCharNumbMapping` has to be expanded.")
  row <- row[1]:row[2]
  allRowCol <- expand.grid(col, row)
  out <- paste0(if (sheet != "") paste0(sheet, "!"), "$", allRowCol[, 1], "$", allRowCol[, 2])
  return(out)
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

# region = "calculations1!$I$6:$L$10"; file = "script-out/H_calculations1-I6-L10.csv"
saveRegionToFile <- function(wb, range, file) {
  if (!grepl("!", range))
    stop("Region name is not fully qualified.")
  range <- strsplit(range, "!")[[1]]
  sheet <- range[1]
  range <- range[2]
  range <- strsplit(range, ":")[[1]]
  if (length(range) != 2)
    stop("Not a range given in `region`.")
  rowColStart <- getRowCol(range[1])
  rowColStopp <- getRowCol(range[2])

  out <- XLConnect::readWorksheet(wb, sheet = sheet, startRow = rowColStart$row, startCol = rowColStart$col,
                                  endRow = rowColStopp$row, endCol = rowColStopp$col,
                                  autofitRow = FALSE, autofitCol = FALSE, header = FALSE)

  write.table(out, file=file, sep = ";", eol = "\n", quote=FALSE, col.names=FALSE, row.names=FALSE) # KEINE Colnames oder Rownames
}

# Is kind of numeric or can be coerced to.
isNum <- function(x) {
  !is.na(suppressWarnings(as.numeric(x)))
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


#### Functions for circular references
#' @export
#' @title Find circular references
#' @description Finds circular references in an expression.
#' @author Daniel Hoop
#' @param expr The expression that should be evaluated. Must be quoted like \code{expr = quote({ a <- 1; b <- 2 })}
#' @param assignOp The assignment operator which assigns the values of the right hand side (RHS) to the left hand side (LHS). A character vector of any length, e.g. c("<-", "=")
#' @param filterPatternFunction A function to filter the parts of a line. Consider a line \code{b <- a * foo(b/a)}.\cr
#' The function \code{foo} might take the value of \code{b} from somewhere else, thus you don't want this to be identified as a circular reference.\cr
#' In this case, specify \code{filterPatternFunction} as follows: \code{filterPatternFunction = function(x) return(x[!startsWith(x, 'foo(')]))}
#' @return \code{NULL}, if no circular references are found, otherwise the lines of the expression that contain circular references.
#' @examples expr <- quote({ a <- b; b <- c; c <- a })
#' findCircularReferences(expr)
#' # [1] "a <- b" "b <- c" "c <- a"
#'
#' findCircularReferences(quote({ a <- if (a == 1) a + 1 else b }))
#' # [1] "a <- if (a == 1) a + 1 else b"
#'
#' findCircularReferences(quote({ a <- if (a == 1) 2 else a + b }))
#' # [1] "a <- if (a == 1) 2 else a + b"
#'
#' expr <- quote({
#'   a <- 1
#'   z <- 9
#'   b <- 2
#'   if (a == 1) y <- z else if (b == 1) a <- 1 else x <- z
#'   d <- 3
#'   z <- y
#' })
#' findCircularReferences(expr)
#' [1] "z <- 9" "y <- z" "z <- y"
#'
#' findCircularReferences(quote({ a <- b; c <- d }))
#' # NULL
#'
#' findCircularReferences(quote({ a <- if (a == 1) 2 else b }))
#' # NULL
findCircularReferences <- function (expr, assignOp=c("<-", "="), filterPatternFunction=function(x) return(x[!grepl("^(lookup|if) *\\(", x)])) {

  if (file.exists(expr)) {
    expr <- readLines(expr, encoding = "UTF-8")
    expr <- eval(parse(text = paste0("quote({", paste0(expr, collapse = "\n"), "\n})")))
  }
  expr <- .prepareExpression(expr, "expr")
  if (is.null(expr)) # This happens when an empty expression is given.
    return (NULL)

  # Split multiline if else statements like "if (a == b) c <- 1 else b <- 2"
  assignOpRegex <- paste0(escapeStr(assignOp), collapse = "|")
  expr <- unlist(apply(matrix(expr), 1, function(expr) {
    if (!(grepl("^if( |\\()", expr) && grepl(assignOpRegex, expr)))
      return (expr)
    expr <- .splitExceptInBracket(expr, splitIfElseOnly = TRUE)
    expr <- gsub("^if *\\(.*", "", expr)
    return(expr[expr != ""])
  }))

  # Convert multiple assignments per line to one assignment per line each.
  notAroundOperator <- if ("=" %in% assignOp) c("<", ">", "!", "=") else NULL
  expr0 <- unlist(mapply(
    expr = as.list(expr),
    splitted = lapply(as.list(expr), .splitExceptInBracket, operators = assignOp, ifIsAnOperator = FALSE, notAroundOperator = notAroundOperator),
    function (expr, splitted) {
      splitted <- trimws(splitted)
      if (length(splitted) > 2)
        return (rev(paste(c(splitted[-length(splitted)]), c(splitted[-1]), sep = paste0(" ", assignOp[1], " "))))
      return (expr)
    }))

  # Split again by assignment operator.
  expr1 <- lapply(expr0, .splitExceptInBracket, operators = assignOp, ifIsAnOperator = FALSE, notAroundOperator = notAroundOperator)

  # The next step splits each entry in expr1 accoding to the assignment operator. Then, right hand side will be split by operators
  # An entry in a list could look like this, where the original line in expression was z <- b + a
  #  list(list("z"), list(c("b","a")))
  errorMsg <- tryCatch({
    expr1 <- lapply(expr1, function (x) {
      if (length(x) == 1)
        return (NULL)
      if (length(x) > 2) {
        stop (paste0("Each row in the expression must contain maximally one assignment operator '", paste0(assignOp, collapse = "/"), "'. E.g.: a ", assignOp[1], " b.\n",
                     "An erroneouss row looks something like: ", paste0(x, collapse=paste0(" ", assignOp[1]," "))))
      }
      # Split outside of brackets
      res <- .splitExceptInBracket(as.list(x))
      if (length(res[[1]]) > 1)
        stop (paste0("Assigning to a LHS expression `a + b <- c` is not alloed.\n",
                     "An erroneouss row looks something like: ", paste0(x, collapse=paste0(" ", assignOp[1]," "))))
      if (!is.null(filterPatternFunction))
        res <- lapply(res, filterPatternFunction)
      # Now split completely all vars
      res[[2]] <- .splitAll(res[[2]], ignoreLhsAssignment = TRUE)
      if (!is.null(filterPatternFunction))
        res <- lapply(res, filterPatternFunction)
      res <- lapply(res, function (x) x[x!=""] )
      return (res)
    })
    "" # return
  }, error = function (e) return (e$message))
  if (errorMsg != "")
    return (errorMsg)

  expr0 <- expr0[!sapply(expr1, is.null)]
  expr1 <- expr1[!sapply(expr1, is.null)]

  # For each entry in left hand side (lhs), check if it appears in any right hand side (rhs).
  # If not, then delete from list, and carry on, as long as there are changes.
  # If there are no changes anymore, but the list is not empty, then there are circulars. If the list is empty in the end, everything is fine.
  while (TRUE) {
    lengthOld <- sum(!sapply(expr1, is.null))
    allx2 <- unlist(lapply(expr1, function(x)x[[2]]))
    expr1 <- lapply(expr1, function (x) {
      if (length(x) == 0)
        return (NULL)
      if (!x[[1]] %in% allx2)
        return (NULL)
      return (x)
    })
    lengthNew <- sum(!sapply(expr1, is.null))
    if (lengthNew == 0 | lengthNew == lengthOld)
      break
  }
  if (lengthNew > 0)
    return (expr0[!sapply(expr1, is.null)])
  return (NULL)
}

if (FALSE) {
  exprTest <- expression({
    a <- lookup(z)
    z <- b + a # z
    a <- b
    c <- b
    d <- c; b <- d
  })
  cat(paste0(findCircularReferences(exprTest), collapse="\n"))
}

#' @keywords internal
#' @author Daniel Hoop
.prepareExpression <- function (expr, argName) {
  if (is.expression(expr))
    stop (paste0("The argument '", argName, "' must not be an expression but quoted. Use something along the lines of:\n", argName, " = quote({\n  a <- 1\n  b <- 2\n})"))
  expr0 <- as.character(expr)
  expr0 <- if (length(expr0) > 0 && expr0[1] == "{") expr0[-1] else expr0
  expr0 <- if (length(expr0) > 0 && expr0[length(expr0)] == "}") expr0[-length(expr0)] else expr0
  # The next step is necessary, because inside if statements if(a == b) { a <- 1; b <- 2 }, "\n" characters are included, but these lines will not be in seperate vector places.
  expr0 <- unlist(strsplit(expr0, "\n"))
  expr0 <- expr0[expr0 != ""]
  return (expr0)
}

#' Escapes characters that have special meaning in regular expressions with backslashes.
#' @keywords internal
#' @author Daniel Hoop
#' @param x Character vector containing the strings that should be escaped.
#' @examples
#' escapeStr(c("asdf.asdf", ".asdf", "\\.asdf", "asdf\\.asdf") )
escapeStr <- function(x) {
  if (!is.character(x))
    stop("x must be a character vector.")
  regChars <- c(".","|","(",")","[","]","{","}","^","$","*","+","-","?")
  res <- apply(matrix(x), 1, function(y) {
    if (nchar(y)==1 && y%in%regChars)
      return (paste0("\\", y))
    for (i in nchar(y):2) {
      if ( substr(y, i-1, i-1) != "\\" && substr(y, i, i) %in% regChars )
        y <- paste0(substr(y, 1, i-1), "\\", substr(y, i, nchar(y)))
    }
    if (substring(y, 1, 1) %in% regChars)
      y <- paste0("\\", y)
    return (y)
  })
  return (res)
}

#' Splits an expression where operators occur, but not inside brackets.
#' @keywords internal
#' @author Daniel Hoop
#' @examples
#' .splitExceptInBracket("(a+b+c)+1++1+")
#' [1] "(a+b+c)" "1"       ""        "1"
#' .splitExceptInBracket("a+b", splitIfElseOnly=TRUE)
#' [1] "a+b"
#' .splitExceptInBracket("1 <- if (a == b) c else d", splitIfElseOnly=TRUE)
#' [1] "1 <- if (a == b)" "c"                "d"
#' .splitExceptInBracket(c("a = b+b", "a <-b+b"), operators = c("=", "<-"))
#' [[1]]
#' [1] "a"   "b+b"
#' [[2]]
#' [1] "a"   "b+b"
#' .splitExceptInBracket("a <= b", operators = c("<-", "="), notAroundOperator = c("<", ">", "!", "="))
.splitExceptInBracket <- function (txt,
                                   operators = NULL,
                                   notAroundOperator = NULL,
                                   ifIsAnOperator = TRUE,
                                   splitIfElseOnly = FALSE) {

  if (length(txt)>1)
    return (lapply(as.list(txt), .splitExceptInBracket,
                   operators = operators,
                   notAroundOperator = notAroundOperator,
                   ifIsAnOperator = ifIsAnOperator,
                   splitIfElseOnly = splitIfElseOnly))

  if (length(txt) == 0 || txt == "")
    return (txt)
  if (splitIfElseOnly) {
    op <- "else"
  } else {
    if (is.null(operators)) {
      op <- c("*", "-", "+", "/", "^", "=", "!", "<", ">", "%%", "} else {", "}else{", "} else ", " else ")
      opRegex <- "\\*|\\-|\\+|/|\\^|=|<|>|%%|(\n| |\\})+else(\\{| |\n)+"
    } else {
      op <- operators
      opRegex <- paste0(escapeStr(operators), collapse = "|")
    }
    if (is.null(notAroundOperator) && !grepl("\\(|\\)|\\[|\\]", txt))
      return (trimws(strsplit(txt, opRegex)[[1]]))
  }

  bBal <- 0 # Bracket balance
  parts <- character() # Splitted result
  lastSplit <- 1 # Index of last splitted operator
  hasIfOccured <- hasIfBracketOpened <- hasIfBracketClosed <- FALSE
  rest <- txt
  for(i in 1:nchar(txt)) {
    st <- substr(txt, i, i)
    restMin1 <- rest
    rest <- substring(txt, i)
    if (st == "(" || st == "[") bBal <- bBal + 1
    if (st == ")" || st == "]") bBal <- bBal - 1
    if (hasIfOccured && bBal == 1)
      hasIfBracketOpened <- TRUE
    if (bBal == 0) {
      hasIfOccured <- ifIsAnOperator && (hasIfOccured || grepl("if( |\n)*\\(", rest))
      if (hasIfBracketOpened)
        hasIfBracketClosed <- TRUE
      if (hasIfBracketClosed) {
        hasIfOccured <- hasIfBracketOpened <- hasIfBracketClosed <- FALSE
        parts <- c(parts, substr(txt, lastSplit, i))
        lastSplit <- i+1
      } else {
        # Normal case without if
        if (!is.null(notAroundOperator)) {
          opFound <- which(sapply(op, function (op) {
            foundTmp <- startsWith(rest, op)
            return (foundTmp
                    && !any(sapply(notAroundOperator,
                                   function(notOp) {
                                     startsWith(restMin1, notOp) ||
                                       startsWith(substring(rest, 2), notOp)
                                   })))
          }))
        } else {
          opFound <- which(sapply(op, function (op) startsWith(rest, op)))
        }
        if (length(opFound) > 0) {
          parts <- c(parts, substr(txt, lastSplit, i-1))
          lastSplit <- i + nchar(op[opFound]) # i+1
        }
      }
    }
  }
  if (lastSplit != i+1)
    parts <- c(parts, substr(txt, lastSplit, i))
  return (trimws(parts))
}

#' @keywords internal
#' @author Daniel Hoop
.splitAll <- function (txt, killFunctionCalls = TRUE, ignoreLhsAssignment = FALSE) {
  # Helper function for function 'findCircularReferences'
  if (length(txt) > 1)
    return (unlist(lapply(txt, .splitAll, ignoreLhsAssignment = ignoreLhsAssignment)))
  if (length(txt) == 0 || txt == "")
    return (txt)
  # Kill function calls
  if (killFunctionCalls)
    txt <- gsub("[a-zA-Z.][a-zA-Z0-9_]? *\\(", "", txt)
  if (ignoreLhsAssignment) {
    txt <- unlist(strsplit(txt, "\\*|\\+|/|\\^|%%|\\(|\\)|\\[|\\]|>|<=|==|=>|!=")) # Split everything except assignments
    txt <- unlist(strsplit(txt, "(?<!<)\\-", perl = TRUE)) # Negative lookbehind: "-" that is not preceeded by "<"
    txt <- unlist(strsplit(txt, "<(?!\\-)", perl = TRUE)) # Negative lookahead: "<" that is not followed by "-"
    # If there is an assignment in a verctor place, then take the right hand side.
    txt <- unlist(lapply(strsplit(txt, "<\\-|="), function (x) {
      if (length(x) == 2)
        return (x[2])
      return (x)
    }))
    return (trimws(txt))
  }
  opRegex <- "\\*|\\-|\\+|/|\\^|=|<|>|%%"
  return (trimws(unlist(strsplit(txt, opRegex))))
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
invisible(.keyValueStore$setAndReturn("lookup_ranges", NULL))

varCells <- XLConnect::getDefinedNames(calc)
dupl <- varCells[duplicated(varCells)]
if (length(dupl) > 0)
  stop("There are duplicated variable names: ", paste0(dupl, collapse = ", "))

#### Get calculations for defined variables ####
for (varName in varCells) {
  cell <- unname(XLConnect::getReferenceFormula(calc, varName))
  sheet <- getSheet(cell)
  if (!is.null(skipSheetPattern) && grepl(skipSheetPattern, sheet)) {
    .keyValueStore$append("log", paste0('\nSheet "', sheet, '" excluded by `skipSheetPattern`: "', skipSheetPattern, '".'))
    next
  }
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
  form <- functionTransformer(form, sheet = sheet, wb = calc)
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
    if (!is.null(skipSheetPattern) && grepl(skipSheetPattern, sheet)) {
      .keyValueStore$append("log", paste0('\nSheet "', sheet, '" excluded by `skipSheetPattern`: "', skipSheetPattern, '".'))
      next
    }
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
          form <- suppressWarnings(as.numeric(form))
          if (is.na(form))
            next
        }
        form <- functionTransformer(form, sheet = sheet, wb = calc)
        form <- fullyQualifiedFormula(form, sheet = sheet)

        # If it is a number and to the left or right, there is also a number, then it is probably a table and should not be saved as variable.
        if (grepl("^[0-9]+$", form)) {
          leftIsNumber <- col > 1 && grepl("^[0-9]+$", sheetList[[sheet]][row, col - 1])
          leftIsEmpty <- col > 1 && is.na(sheetList[[sheet]][row, col - 1])
          rightIsNumber <- local({
            if (length(sheetList[[sheet]][row, col + 1]) == 0 || is.na(sheetList[[sheet]][row, col + 1]))
              return(FALSE)
            isTrue <- try(grepl("^[0-9]+$", sheetList[[sheet]][row, col + 1]), silent = TRUE)
            if (class(isTrue) == "try-error" || length(isTrue) == 0)
              return(FALSE)
            return(isTrue)
          })
          if (leftIsNumber) { # || (leftIsEmpty && rightIsNumber)) {
            cell <- cellToChar(sheet = sheet, row = row, col = col, dollar = FALSE)
            .keyValueStore$append("log", paste0(cell, ": Discarded. Cell to left is number -> Table, not variable.")) # `leftIsNumber || (leftIsEmpty && rightIsNumber)`"))
            next
          }
        }

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
          cell = cellToChar(sheet = sheet, row = row, col = col, dollar = TRUE))
        varCellsList2[[length(varCellsList2)]] <- as.data.frame(varCellsList2[[length(varCellsList2)]], stringsAsFactors = FALSE)
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
      kickVars <- (varCellsList2[, "name"] %in% varCellsList1[, "name"] |
                     varCellsList2[, "cell"] %in% varCellsList1[, "cell"])
      if (sum(kickVars) > 0) {
        .keyValueStore$append(
          "log",
          paste0("\nThe following variables were discarded because their name was guessed, but there already existed",
                 " a variable with that defined name or range:\n",
                 paste0(paste0(varCellsList2[kickVars, "name"], " (", gsub("\\$", "", varCellsList2[kickVars, "cell"]), ")"),
                        collapse = ", ")))
        varCellsList2 <- varCellsList2[!kickVars,, drop = FALSE]
      }
    }
  }
}

#### Combine variables of which the names were defined and guessed ####
varCellsList <- rbind(varCellsList1, varCellsList2)

kickCells <- intersect(varCellsList[, "cell"], .keyValueStore$get("lookup_ranges"))
if (length(kickCells) > 0) {
  .keyValueStore$append("log", paste0("\nThe following cells were not considered a variable because they were part of lookup tables:\n",
                                      paste0(gsub("\\$", "", kickCells), collapse = ", "), "\n"))
  varCellsList <- varCellsList[!varCellsList[, "cell"] %in% kickCells, ]
}

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
  calcList[[length(calcList)]][["formVarsPasted"]] <- paste0(calcList[[length(calcList)]][["formVars"]], collapse = ",")
}

names(calcList) <- rownames(varCellsList)

notVariableConverted <- lapply(calcList, function(x) if(any(isCellNotVariable(x[["formVars"]]))) return(x[!names(x) %in% "formVars"]) else return(NULL))
notVariableConverted <- notVariableConverted[!sapply(notVariableConverted, is.null)]
if (length(notVariableConverted) > 0) {
  print(do.call("rbind", notVariableConverted))
  stop("The transformation of some cells into variables did not work. See printed output above. Problematic cells are found in columns 'form' and 'formVarsPasted'.")
}

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

  if (length(calcList2) == lengthList || nRuns >= 100000) { # length(calcList2) != 1
    print(do.call("rbind", lapply(calcList2, function(x) {
      x[["formVars"]] <- paste0(x[["formVars"]], collapse = ", ")
      x
    })))
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
script <- paste0(
  .keyValueStore$getOrSet("pre_script_block", ""),
  "\n\n", "#### script ####")

lastSheet <- "wefplxwerplwef"
for (i in 1:nrow(calcDf)) {
  if (calcDf[i, "dependsOnSheet"] != lastSheet) {
    script[length(script) + 1] <- ""
    script[length(script) + 1] <- paste0(scriptCommentPrefix, calcDf[i, "dependsOnSheet"])
    lastSheet <- calcDf[i, "dependsOnSheet"]
  }
  script[length(script) + 1] <- paste0(calcDf[i, "var"], " = ", calcDf[i, "form"])
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
  if (addRightCellValueAsComment) {
    addVal <- sheetList[[calcDf[i, "sheet"]]][calcDf[i, "row"], calcDf[i, "col"] + 1]
    if (length(addVal) == 0 || is.na(addVal))
      addVal <- ""
    script[length(script)] <- paste0(script[length(script)], addVal, " | ")
  }

  # Clean up unneccesary commentary "|" signs
  script[length(script)] <- gsub("|[| ]* *$", "", script[length(script)], fixed = FALSE)
  script[length(script)] <- trimws(script[length(script)])
}

{
  message("*** Script start ***")
  invisible(lapply(script, function(line) cat(line, "\n", sep = "")))
  message("\n*** Script end ***")
}
writeLines(script, scriptOutFile, useBytes = TRUE)

if (!is.null(logOutFile) && .keyValueStore$contains("log"))
  writeLines(.keyValueStore$get("log"), logOutFile, useBytes = TRUE)

circulars <- findCircularReferences(scriptOutFile)
if (length(circulars) > 1) {
  print(circulars)
  stop("Circular references found in the script.")
} else {
  message("*Success*\nIf you alter the script in the future, assure that it does not contain circular references using the function 'findCircularReferences(filename)'.")
}