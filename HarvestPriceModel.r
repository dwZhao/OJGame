library("TTR")
library("forecast")

setwd("Documents/PrincetonSenior/ORF411/OJgame")
ALLprices = read.csv("HistoricalHarvestPrices.csv", header = TRUE, sep=",")
FLAprices = ALLprices$FLA
CALprices = ALLprices$CAL
TEXprices = ALLprices$TEX
ARZprices = ALLprices$ARZ
BRAprices = ALLprices$BRA
SPAprices = ALLprices$SPA

FLAtimeseries = ts(FLAprices, start = c(2003,9), frequency = 12)
CALtimeseries = ts(CALprices, start = c(2003,9), frequency = 12)
TEXtimeseries = ts(TEXprices, start = c(2003,9), frequency = 12)
ARZtimeseries = ts(ARZprices, start = c(2003,9), frequency = 12)
BRAtimeseries = ts(BRAprices, start = c(2003,9), frequency = 12)
SPAtimeseries = ts(SPAprices, start = c(2003,9), frequency = 12)

#components
FLAcomponents = decompose(FLAtimeseries)
plot(FLAcomponents)

#ARIMA: going with this one!!
timeseries = FLAtimeseries #use 2,1,0
timeseries = CALtimeseries #use 3,1,0
timeseries = TEXtimeseries #use 4,1,0
timeseries = ARZtimeseries #use 4,1,0
timeseries = BRAtimeseries #use 3,1,0
timeseries = SPAtimeseries #use 3,1,0

fit = Arima(timeseries, order=c(3,1,0), seasonal=list(order=c(3,1,0),period = 12))
plot(timeseries, col='blue')
lines(fitted(fit), col='red')
plot(forecast(fit))

#STLF
FLAmodel = stlf(FLAtimeseries, etsmodel="AAN", damped = FALSE)
plot(FLAmodel)

#find price multiplier cutoffs, for the next year
lower=forecast(fit)$lower
upper=forecast(fit)$upper
# winsorize negative prices
lower[lower < 0] = 0
upper[upper < 0] = 0 #should never happen but in case...
nextLow80 = mean(lower[1:12,1]) #average over twelves
nextHigh80 = mean(upper[1:12,1])

