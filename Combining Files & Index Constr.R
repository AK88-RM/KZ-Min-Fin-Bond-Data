
## Enable Trust access to the VBA project object model
library(stringr)
library(readxl)
library(dplyr)
library(reshape2)
library(lubridate)
library(ggplot2)
library(scales)

getOption("max.print")
options(max.print = 1000000)

getwd()
setwd("C:/Users/user.userov/Documents/Price")

## Gotta put together all the files
## make sure that the files are in historical order (2010, 2011, curr)

file.list <- list.files(pattern='*.csv') ## get all the excel files in the directory
df.list <- lapply(file.list, read.csv)   ## read them all
new.df.list = lapply(df.list,
                     function(y) {y["X"] <- NULL; y}) ## delete numbering



df = Reduce(function(x, y) merge(x, y,
                                 by = c("Code", "URL", "BondsQuantity",
                                        "EndDate", "ISIN", "IssueVolume",
                                        "NominalValue", "StartDate"),
                                 all.x = TRUE, all.y = TRUE), new.df.list)
dim(df)
df[df == Inf] = NA

df$StartDate = mdy(df$StartDate)
df$EndDate = mdy(df$EndDate)
write.csv(df, "df.csv")

## Clean Price data
Clean = cbind(df[1:8], df[, seq(10, ncol(df), by = 3)])  ## extract only clean price
Clean$BondsQuantity = as.numeric(gsub(",", "", Clean$BondsQuantity))
Clean$NominalValue = as.numeric(gsub(",", "", Clean$NominalValue))
Clean$IssueVolume = as.numeric(gsub(",", "", Clean$IssueVolume))
colnames(Clean)[9:ncol(Clean)] = gsub("X", "",
                                      colnames(Clean)[9:ncol(Clean)])
colnames(Clean)[9:ncol(Clean)] = gsub("__1", "",
                                      colnames(Clean)[9:ncol(Clean)])
names(Clean)[9:ncol(Clean)] = format(as.Date(as.numeric(names(Clean)[9:ncol(Clean)]),
                                             origin="1899-12-30"))

sum(is.na(Clean$IssueVolume)) ## check if there are NAs in cols for weighting
                              ## if there are, run the code below
                              ## MinFin should be fine, NTK are problematic
# for (i in 1:nrow(Clean)) {
#   if (is.na(Clean[i, 3]) == TRUE) {
#     Clean[i, 3] = Clean[i, 5]/Clean[i, 4]
#   }
# }

## Dirty Price data
Dirty = cbind(df[1:8], df[, seq(11, ncol(df), by = 3)])  ## extract only clean price
Dirty$BondsQuantity = as.numeric(gsub(",", "", Dirty$BondsQuantity))
Dirty$NominalValue = as.numeric(gsub(",", "", Dirty$NominalValue))
Dirty$IssueVolume = as.numeric(gsub(",", "", Dirty$IssueVolume))
colnames(Dirty)[9:ncol(Dirty)] = gsub("X", "",
                                      colnames(Dirty)[9:ncol(Dirty)])
colnames(Dirty)[9:ncol(Dirty)] = gsub("__2", "",
                                      colnames(Dirty)[9:ncol(Dirty)])
names(Dirty)[9:ncol(Dirty)] = format(as.Date(as.numeric(names(Dirty)[9:ncol(Dirty)]),
                                             origin="1899-12-30"))

sum(is.na(Dirty$IssueVolume)) ## check if there are NAs in cols for weighting
## if there are, run the code below
## MinFin should be fine, NTK are problematic
# for (i in 1:nrow(Clean)) {
#   if (is.na(Clean[i, 3]) == TRUE) {
#     Clean[i, 3] = Clean[i, 5]/Clean[i, 4]
#   }
# }

## Create an empty vector to store Index values
temprow <- matrix(c(rep.int(NA, ncol(Dirty))), nrow=1, ncol = ncol(Dirty))
PI <- data.frame(temprow)
colnames(PI) = colnames(Dirty)
PI[, 9] = 100

## Get ready Accrued Interest data
AI = Dirty[, colnames(Dirty)[c(9:ncol(Dirty))]] -
  Clean[, colnames(Clean)[c(9:ncol(Clean))]]
AI[is.na(AI)] = 0
AI[1:12, 4] = 0
dim(AI)

## TEST area start
j = 10
tDr = Dirty  ## set all Dirty data to tDr
tCl = Clean  ## set all Clean data to tCl
tAI = AI     ## set all AI data to tAI
tDr = tDr[, colnames(tDr)[c(1:8, j - 1, j)]]  ## get 2 Dirty cols at a given point
tCl = tCl[, colnames(tCl)[c(j)]]              ## get the corresponding Clean col
tAI = tAI[,colnames(tAI)[c(j - 9, j - 8)]]    ## get 2 AI cols at a given point
tDr = cbind(tDr, tCl, tAI)                    ## bind them all to tDr

tDr = subset(tDr, tDr$StartDate <= as.Date(colnames(tDr)[10])) ## Issue date should be less than Portfolio date (today)
tDr = subset(tDr, tDr$EndDate < (as.Date(colnames(tDr)[10]) + years(3))) ## Maturity should be max 3 years
tDr = subset(tDr, !is.na(tDr[, 10]))

rownames(tDr) = seq(length = nrow(tDr))

cost = 0.05

tDr[, 9] = ifelse(is.na(tDr[, 9]) &
                    ((as.numeric(tDr$EndDate - tDr$StartDate)/365) > 3),
                  tDr[, 10] * (1 + cost), tDr[, 9])
tDr[, 9] = ifelse(is.na(tDr[, 9]) &
                    ((as.numeric(tDr$EndDate - tDr$StartDate)/365) <= 3),
                  100, tDr[, 9])
tDr[, 9] = ifelse(grepl("MOM", tDr[, 1]) | grepl("MUM", tDr[, 1]),
                  ifelse(tDr[, 12] > tDr[, 13], tDr[, 11], tDr[, 9]),
                  tDr[, 9])

TR = sum(tDr[, 10] * tDr[,7]/100 * 100 * tDr[,3])/
     sum(tDr[, 9] * tDr[,7]/100 * 100 * tDr[,3])
PI[, j] = PI[, j - 1] * TR
PI[,j]
## Test area end


## Define slippage cost
cost = 0.05

## Loop thru 
for (j in 10:ncol(Dirty)) {
  tDr = Dirty  ## set all Dirty data to tDr
  tCl = Clean  ## set all Clean data to tCl
  tAI = AI     ## set all AI data to tAI
  tDr = tDr[, colnames(tDr)[c(1:8, j - 1, j)]]  ## get 2 Dirty cols at a given point
  tCl = tCl[, colnames(tCl)[c(j)]]              ## get the corresponding Clean col
  tAI = tAI[,colnames(tAI)[c(j - 9, j - 8)]]    ## get 2 AI cols at a given point
  tDr = cbind(tDr, tCl, tAI)                    ## bind them all to tDr
  
  tDr = subset(tDr, tDr$StartDate <= as.Date(colnames(tDr)[10])) ## Issue date should be less than Portfolio date (today)
  tDr = subset(tDr, tDr$EndDate < (as.Date(colnames(tDr)[10]) - days(1) + years(3))) ## Maturity should be max 3 years
  tDr = subset(tDr, !is.na(tDr[, 10]))
  rownames(tDr) = seq(length = nrow(tDr))       ## reset numbering of the rows 
  
  tDr[, 9] = ifelse(is.na(tDr[, 9]) &
                      ((as.numeric(tDr$EndDate - tDr$StartDate)/365) > 3),
                    tDr[, 10] * (1 + cost), tDr[, 9])
  tDr[, 9] = ifelse(is.na(tDr[, 9]) &
                      ((as.numeric(tDr$EndDate - tDr$StartDate)/365) <= 3),
                    100, tDr[, 9])
  tDr[, 9] = ifelse(grepl("MOM", tDr[, 1]) | grepl("MUM", tDr[, 1]),
                    ifelse(tDr[, 12] > tDr[, 13], tDr[, 11], tDr[, 9]),
                    tDr[, 9])
  
  TR = sum(tDr[, 10] * tDr[,7]/100 * 100 * tDr[,3])/
       sum(tDr[, 9] * tDr[,7]/100 * 100 * tDr[,3])
  PI[, j] = PI[, j - 1] * TR
}

PI[1, 1] = "TRIndex"
PI

write.csv(PI, "PI.csv")


tot = rbind(Dirty, PI)
tot = tot[,colnames(tot)[c(1, 9:ncol(tot))]]
tail(tot)

BDmelt <- melt(tot, id.vars = "Code")
BDmelt$value <- as.numeric(BDmelt$value)
BDmelt$variable <- as.Date(BDmelt$variable)

re = ggplot() +
  geom_line(data = BDmelt, aes(x = variable, y = value, group = Code), color = "lightblue") +
  
  geom_line(data = subset(BDmelt, Code == "TRIndex"), aes(x = variable, y = value, group = Code), color = "black", size = 1)+
  
  theme(axis.text.x = element_text(angle = 90, hjust = 1), axis.text = element_text(size=7)) +
  theme(legend.position = "none") +
  xlab("Date") +
  ylab("Price") +
  ggtitle("KZ Gov Total Return Bond Index") +
  scale_x_date(date_breaks = "1 month", date_minor_breaks = "1 week", labels = date_format("%b-%y"))
re
