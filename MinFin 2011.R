
## Enable Trust access to the VBA project object model
library(rvest)
library(stringr)
library(readxl)
library(dplyr)
library(reshape2)
library(lubridate)
library(ggplot2)
library(scales)

getOption("max.print")
options(max.print=1000000)

## Working on the sorted data as per structure: 
## Feb 15, 2010; Aug 31, 2011; current


## Enable Trust access to the VBA project object model
shell(shQuote(normalizePath("C:/Users/user.userov/Documents/Price/2011/VBS1.vbs")), "cscript", flag = "//nologo")

getwd()
setwd("C:/Users/user.userov/Documents/Price/2011")

file.list <- list.files(pattern='*.xls') ## get all the excel files in the directory
df.list <- lapply(file.list, read_excel) ## read them all

dat = Reduce(function(x, y) merge(x, y, by = "Code",
                                  all.x = TRUE, all.y = TRUE), df.list)
dt = dat[ , order(colnames(dat))]
Code = dt$Code

dt = dt[, -ncol(dt)]  ## Code column is at the end. Delete it 
dt = cbind(Code, dt)  ## and move it forward

ndt = dt              ## just to be safe
tst = ndt$Code        ## some Codes are not useful for web scraping
ntst = sub(" ", "", tst)    ## hence gotta modify 2 " " using "sub" twice which
ntst2 = sub(" 0", ".", ntst)## matches only the first occurrence of pattern
ndt$Code = ntst2      ## substitute the initial Code with the new one

                      ## now we have duplicated rows and deal with it as below:
agg.ndt = aggregate(x = ndt[, !(names(ndt) %in% "Code")],
                    by = list(Code = ndt$Code), min, na.rm = TRUE)

agg.ndt[ , "URL"] <- NA   ## create new NA URL columns
agg.ndt$URL = apply(agg.ndt, 1, function(x) paste0("http://old.kase.kz/en/gsecs/show/", toString(x[1])))

name = as.character(agg.ndt$Code)
URL = agg.ndt$URL

instruments <- data.frame(name, URL, stringsAsFactors = FALSE)

wanted <- c("NSIN" = "ISIN",
            "Nominal value in issue's currency" = "NominalValue",
            "Number of bonds outstanding" = "BondsQuantity",
            "Issue volume, KZT" = "IssueVolume",
            "Date of circulation start" = "StartDate",
            "Principal repayment date" = "EndDate")

getValues <- function (name, url) {
  df <- url %>%
    read_html() %>%
    html_nodes("table.top") %>%
    html_table()
  df = as.data.frame(df)
  names(df) <- c("full_name", "value")
  
  
  ## filter and remap wanted columns
  result <- df[df$full_name %in% names(wanted),]
  result$column_name <- sapply(result$full_name, function(x) {wanted[[x]]})
  
  ## add the identifier to every row
  result$name <- name
  return (result[,c("name", "column_name", "value")])
}

## invoke function for each name/URL pair - returns list of data frames
columns <- apply(instruments[,c("name", "URL")], 1, function(x) {getValues(x[["name"]], x[["URL"]])})

## bind using dplyr:bind_rows to make a tall data frame
tall <- bind_rows(columns)

## make wide using dcast from reshape2
wide <- dcast(tall, name ~ column_name, id.vars = "value")
colnames(wide)[1]= "Code"

## ***********************************************************
## ***********************************************************

nwd = merge(agg.ndt, wide, by = "Code")
nwd

write.csv(nwd, "nwd2011.csv")
