# Excel formula extractor

## What it will do for you
Did you ever inherit a project from a colleague where all calculations were done in Excel?  
However, for reproducibility and version control reasons, you would prefer to use a programming language such as [R](https://www.r-project.org/)...  
This project can help you!

* **The script will extract all formulas defined in an Excel sheet into an R script**.
* It will use the defined variable names in the Excel sheet (if available).
  * If no names were defined, it will try to guess the names by accessing the cell left to a certain formula.
  * If no variable name can be guessed, it will create generic variable names.
* It will figure out the ordering in which the calculations have to be executed.
  * Still, it will try to keep variables defined in one sheet together.
* It will transform calls to VLOOKUP/HLOOKUP (in German: "SVERWEIS", "WVERWEIS") into calls applicable in R.
  * All data contained in LOOKUP tables will be written to separate csv files.
  * In the beginning of the automatically created script, all the referenced lookup tables will be read in.
* It can be adapted to your needs by adapting the function `functionTransformer` to handle other functions than VLOOKUP/HLOOKUP. E.g., you can adapt it to transform `IF` to `ifelse` function calls, etc.

## How To...
* Configure the script `excel-export.R` in the block `#### Settings ####`. Then, run it via the `source` command. If you use the standard configuration...    * the file `script-out/script.R` will be created containing the formulas from `excel-in/calculations.xlsx`.
  * a log of messages will be written to `script-out/log.txt`.
  * After each script line (variable definition and calculation), there will be a script comment containing the cell where the calculation was found, the content of the cells to the left and right, as well as the content of a comment (if there was any for this cell).
* If you alter the automatically created script (from Excel) in the future, assure that it does not contain circular references using the function `findCircularReferences(filename)`.

The output for the given Excel sheet will be as follows:
```R
#### pre script block ####

tryNum <- function (x) 
{
    if (is.numeric(x)) 
        return(x)
    out <- suppressWarnings(as.numeric(x))
    if (is.na(out)) 
        return(x)
    return(out)
}

V_calculations1_I6_L10 <- read.table("script-out/V_calculations1_I6_L10.csv", sep=";", header=FALSE, row.names = 1, stringsAsFactors=FALSE, quote = "\"", comment.char="", na.strings=c(""))
H_calculations1_I6_L10 <- read.table("script-out/H_calculations1_I6_L10.csv", sep=";", header=TRUE, stringsAsFactors=FALSE, quote = "\"", comment.char="", na.strings=c(""))
V_calculations1_I13_L17 <- read.table("script-out/V_calculations1_I13_L17.csv", sep=";", header=FALSE, row.names = 1, stringsAsFactors=FALSE, quote = "\"", comment.char="", na.strings=c(""))

#### Script ####

# b
w = 4 # b!C5 | b.w | w | is defined as w

# calculations2
e = 1 # calculations2!C2 |  | e | is defined as e
f = 2 # calculations2!C3 |  | f | is defined as f

# calculations1
b = 1 # calculations1!C3 |  | b | is defined as b
d = 4 # calculations1!C5 |  | d | is defined as d
j = 10 # calculations1!C11 |  |  | is defined as j
r_ = "h" # calculations1!C21 |  | r_ | variable for VLOOKUP
VAR_2 = 10 # calculations1!C8 |  |  | should become VAR_
VAR_5 = 10 # calculations1!C11 |  |  | is defined as j
u = tryNum(V_calculations1_I6_L10["a", 1]) # calculations1!C17 |  | u | VLOOKUP / SVERWEIS
u2 = tryNum(V_calculations1_I6_L10["a", 1]) # calculations1!C18 |  | u2 | VLOOKUP / SVERWEIS (duplicate)
t = tryNum(H_calculations1_I6_L10[1,"X3"]) # calculations1!C19 |  | t | HLOOKUP / WVERWEIS
c_ = d+1 # calculations1!C4 |  | c_ | is defined as c_
i = VAR_2+1 # calculations1!C10 |  |  | is defined as i
s = tryNum(V_calculations1_I13_L17[r_, 1]) # calculations1!C20 |  | s | VLOOKUP / SVERWEIS - using variable
a = ifelse(b<10,b+c_,0) # calculations1!C2 | calculations1.a | a | is defined as a
VAR_6 = i+1 # calculations1!C13 |  |  | should become VAR_
VAR_8 = i+1 # calculations1!C14 |  |  | should become VAR_
VAR_9 = i+1 # calculations1!C15 |  |  | should become VAR_

# calculations1, calculations2
g = e+f+a # calculations2!C4 |  | g | is defined as g
h = e+f+a # calculations2!C5 |  | h | is defined as h
z_ = h+1 # calculations2!C6 | calculations2.z_ | z_ | is defined as z_
x = z_+1 # calculations1!C6 | calculations1.x | x | is defined as x

# a, b, calculations1, calculations2
y = x+z_+w # a!C5 | a.y | y | final value
```
