# 
# Written BY; Gus Segura
# Blueskymetrics.com  
# May 3, 2017
# See Git Repository for Updates

# load library
library("readxl")
library("openxlsx")
library("forecast")
library("TTR")
library("zoo")

# results list()
results <- list()

# read sample data - make sure to setwd() to your this local repository directory
champagne2 <- read_excel("champagne-sales.xlsx")


# create time series and seasonal difference series
champagne2timeseries <- ts(champagne2$`ChampagneSales`, frequency=12, start=c(2001,1))
champagne2timeseries

# load seasonal difference from excel data
champagne2tsseadiff <- ts(champagne2$`Seasonal Difference`, frequency=12, start=c(2001,1))
champagne2tsseadiff

# load 1st difference from excel data
champagne2ts1stdiff <- ts(champagne2$`1stDifference`, frequency=12, start=c(2001,1))
champagne2ts1stdiff


# use diff after you isolate the seasonal portion vs using excel and algorithim
firstdiff <- diff(champagne2tsseadiff)
firstdiff

# try some plots
plot.ts(champagne2timeseries)
champ2decomp <- decompose(champagne2timeseries)
plot(champ2decomp)

# try non-ts plots -  looking for stationary
plot(na.omit(champagne2tsseadiff))
plot(na.omit(champagne2ts1stdiff))

# try acf - pacf - LJung to visually access data
acf(na.omit(champagne2ts1stdiff), lag.max = 20)
pacf(na.omit(champagne2ts1stdiff), lag.max = 20)
Box.test(na.omit(champagne2ts1stdiff), lag = 20, type ="Ljung-Box")

# plot time series
plot(champagne2timeseries, ylab="", xlab="year")

# arima - auto arima will pick our terms : use auto to find parameters
auto.arima(champagne2timeseries)

# auto.arima(champagne2timeseries, stepwise=FALSE, approximation=FALSE)
# ARIMA(1,0,0)(1,1,0)[12] with drift
fit.champagne <- arima( champagne2timeseries, order=c(1,0,0), seasonal = c(1,1,0))
results[["champagne"]]  <- forecast.Arima(fit.champagne)
plot(results[["champagne"]], ylab="Forecast of Champagne Sales", xlab="Year")

# results
# convert ts into data frame for export to excel
df.champagne <- data.frame(F1=as.yearmon(time(champagne2timeseries), "%b%y"), 
                          Point.Forecast=as.matrix(champagne2timeseries))

# forecast - combine multiple result sets if needed.
results[["champagne"]]

# combine results
results_together <- do.call(rbind,lapply(names(results),function(x){
  transform(as.data.frame(results[[x]]), Name = x)
}))

# Note: Use this snippet to re-produce the source data from my Tableau Visual
wb <- createWorkbook()
addWorksheet(wb, "Forecasts")
writeData(wb, "Forecasts", df.champagne, rowNames = FALSE)
writeData(wb, "Forecasts", results_together, rowNames = TRUE)
saveWorkbook(wb, "Forcasts.xlsx", overwrite = TRUE)

# Note: Use this snippet to create two sheets in one book - Use Union of source data in Tableau
# to combine both sheets into one data set.  You may want to look at the data in different sets.
wb <- createWorkbook()
addWorksheet(wb, "HistoryData")
addWorksheet(wb, "Forecasts")
writeData(wb, "HistoryData", df.champagne, rowNames = FALSE)
writeData(wb, "Forecasts", results_together, rowNames = TRUE)
saveWorkbook(wb, "Forcasts2.xlsx", overwrite = TRUE)

