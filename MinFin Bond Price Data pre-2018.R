
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

page <- html("http://www.kase.kz/ru/marketvaluation")

y2010 <- page %>%
  html_nodes("a") %>%       # find all links
  html_attr("href") %>%     # get the url
  str_subset("\\.zip") %>%  # find those that end in zip
  str_subset("/ru/2010/val")

y2011 <- page %>%
  html_nodes("a") %>%       # find all links
  html_attr("href") %>%     # get the url
  str_subset("\\.zip") %>%  # find those that end in zip
  str_subset("/ru/2011/val")

y2012 <- page %>%
  html_nodes("a") %>%       # find all links
  html_attr("href") %>%     # get the url
  str_subset("\\.zip") %>%  # find those that end in zip
  str_subset("/ru/2012/val")

y2013 <- page %>%
  html_nodes("a") %>%       # find all links
  html_attr("href") %>%     # get the url
  str_subset("\\.zip") %>%  # find those that end in zip
  str_subset("/ru/2013/val")

y2014 <- page %>%
  html_nodes("a") %>%       # find all links
  html_attr("href") %>%     # get the url
  str_subset("\\.zip") %>%  # find those that end in zip
  str_subset("/ru/2014/val")

y2015 <- page %>%
  html_nodes("a") %>%       # find all links
  html_attr("href") %>%     # get the url
  str_subset("\\.zip") %>%  # find those that end in zip
  str_subset("/ru/2015/val")

y2016 <- page %>%
  html_nodes("a") %>%       # find all links
  html_attr("href") %>%     # get the url
  str_subset("\\.zip") %>%  # find those that end in zip
  str_subset("/ru/2016/val")

y2017 <- page %>%
  html_nodes("a") %>%       # find all links
  html_attr("href") %>%     # get the url
  str_subset("\\.zip") %>%  # find those that end in zip
  str_subset("/ru/2017/val")

linkdata = c(y2010, y2011, y2012, y2013,
  y2014, y2015, y2016, y2017)
linkdata = as.data.frame(linkdata)
linkdata[,"URL"] <- NA
linkdata$URL = apply(linkdata, 1,
                     function(x) paste0("http://www.kase.kz", toString(x[1])))
dim(linkdata)               
linkdata = linkdata[-322, ] # get rid off the file val161108161113_(1).zip
dim(linkdata)
linkdata