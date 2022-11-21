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
* Configure the script `excel-export.R` in the block `#### Settings ####`. Then, run it via the `source` command.
  * If you use the standard configuration, the file `script-out/script.R` will be created containing the formulas from `excel-in/calculations.xlsx`. A log of messages will be written to `script-out/log.txt`
* If you alter the automatically created script (from Excel) in the future, assure that it does not contain circular references using the function `findCircularReferences(filename)`.
