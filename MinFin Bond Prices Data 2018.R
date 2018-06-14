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

page <- read_html("http://old.kase.kz/ru/marketvaluation")

y2018 <- page %>%
  html_nodes("a") %>%       # find all links
  html_attr("href") %>%     # get the url
  str_subset("\\.zip") %>%  # find those that end in zip
  str_subset("/ru/2018/val")

linkdata = c(y2018)
linkdata = as.data.frame(linkdata)
linkdata[,"URL"] <- NA
linkdata$URL = apply(linkdata, 1,
                     function(x) paste0("http://old.kase.kz", toString(x[1])))
dim(linkdata)

getwd()
setwd("C:/Users/a.khussanov/Documents/Price/data")

for (i in 1 : nrow(linkdata)){
  URL = linkdata[i, 2]
  dir = basename(URL)
  download.file(URL, dir)
  unzip(dir)
  TXT <- list.files(pattern = "*.TXT")
  zip <- list.files(pattern = "*.zip")
  bad <- "Price08_11_16(1)_add_eurnot.xls"
  file.remove(TXT, zip, bad)
  Sys.sleep(sample(10, 1) * 0.5)
}	
