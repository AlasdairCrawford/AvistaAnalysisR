#Load Libraries
library(gdata)
library(dplyr)
library(lubridate)
library(openxlsx)
Sys.setenv(R_ZIPCMD= "C:/Rtools/bin/zip")
#Load csv file
fname <- file.choose()
rawdata <- read.csv(fname,sep=",",header=TRUE)
#rawdata <- rawdata[seq(1,NROW(rawdata),by=10),]
#Filter out rows without data
#Only take the columns we are interested in raw
avistadata <- select(rawdata,DATETIME_STAMP,TES_STATE_OF_CHARGE_CAPY_PERCENT)
colnames(avistadata) <- c("Time","SOC")
avistadata$Temp <- rowMeans(rawdata[,c("TES_BATTERY1_STORAGE_TEMP_MAX_DEGC","TES_BATTERY2_STORAGE_TEMP_MIN_DEGC")])
avistadata$Power<-(as.numeric(as.character(rawdata$TES_BATTERY1_KW))+as.numeric(as.character(rawdata$TES_BATTERY2_KW)))*1000
#Convert from text into numerics
avistadata$Power <- as.numeric(as.character(avistadata$Power))
avistadata$SOC <- as.numeric(as.character(avistadata$SOC))/100
avistadata$Time <- as.POSIXct(as.character(avistadata$Time),format="%Y-%m-%d %T")
#Add an hours column for convenience
#avistadata$Hours <- seq(from=0,length.out=length(avistadata$Time),by=(0.00277777777777778))
avistadata$Hours <- as.numeric(avistadata$Time)/60/60
avistadata$Hours <- avistadata$Hours - avistadata$Hours[1]
seq <- rle(avistadata$SOC)$lengths
SOCmid <- round(cumsum(seq)-seq/2)
y <- avistadata$SOC[SOCmid]
x <- avistadata$Hours[SOCmid]
avistadata$SOC <- approx(x,y,xout=avistadata$Hours)$y
#avistadata$SOC <- spline(x,y,xout=avistadata$Hours)$y
#avistadata$SOC <- loess(avistadata$SOC ~ avistadata$Hours,span=0.1)$fitted
#Get rid of rows without any data
avistadata <- avistadata[!is.na(avistadata$Power),]
#Choose conditions to count as taper power
taperhigh <- 0.975
taperlow <- 0.05
dispower <- max(avistadata$Power)
chgpower <- min(avistadata$Power)
avistadata$State <- avistadata$Power %>% sapply(FUN = function(x) (abs(x) > taperlow*dispower+0)*sign(x))
#avistadata$StChange <- c(0,abs(diff(avistadata$State)))
avistadata$StChange <- c(0,1*(diff(avistadata$State)!=0))
avistadata$Index <- cumsum(avistadata$StChange)
avistadata$PowerSum <- cumsum(avistadata$Power)*10/60/60/1000
#disvoltage <- avistadata[avistadata$State == 1 & avistadata$Power > dispower*taperhigh,]
#chgvoltage <- avistadata[avistadata$State == -1 & avistadata$Power < chgpower*taperhigh,]
disvoltage <- avistadata[avistadata$State == 1,]
chgvoltage <- avistadata[avistadata$State == -1,]
cycles <- avistadata %>% group_by(Index)
peaks <- summarize(cycles,Peak = max(abs(Power)),minSOC = min(SOC), maxSOC = max(SOC), DOD=maxSOC-minSOC)
avistadata$Peak <- 0
avistadata$DOD <- 0
for(i in 1:nrow(avistadata)){
  avistadata$Peak[i] <- peaks$Peak[avistadata$Index[i]==peaks$Index]*avistadata$State[i]
  avistadata$DOD[i] <- peaks$DOD[avistadata$Index[i]==peaks$Index]
}
avistadata$PeakFrac <- avistadata$Power/avistadata$Peak
disvoltage <- avistadata[avistadata$State == 1 & avistadata$PeakFrac > taperhigh & avistadata$DOD>0.15,]
chgvoltage <- avistadata[avistadata$State == -1 & avistadata$PeakFrac > taperhigh & avistadata$DOD>0.15,]
summary <- cycles %>% summarize(Start=min(Time), St=head(State,n=1), Energy = sum(Power)*10/60/60/1000,Peak = max(abs(Power)), TaperStart=max(SOC[Power<dispower*taperhigh]),SOCMin=min(SOC),SOCMax=max(SOC),DOD=SOCMax-SOCMin, Duration = max(Hours)-min(Hours),AvgTemp=mean(Temp))
summary <- filter(summary, DOD>0 & St != 0)
for(i in 1:nrow(disvoltage)){
  disvoltage$Elapsed[i] <- disvoltage$Hours[i]-head(disvoltage[disvoltage$Index == disvoltage$Index[i],]$Hours,n=1)
  disvoltage$TimeRemaining[i] <- -disvoltage$Hours[i]+tail(disvoltage[disvoltage$Index == disvoltage$Index[i],]$Hours,n=1)
  disvoltage$Energy[i] <- disvoltage$PowerSum[i]-head(disvoltage[disvoltage$Index == disvoltage$Index[i],]$PowerSum,n=1)
  disvoltage$EnergyRemaining[i] <- -disvoltage$PowerSum[i]+tail(disvoltage[disvoltage$Index == disvoltage$Index[i],]$PowerSum,n=1)
}

for(i in 1:nrow(chgvoltage)){
  chgvoltage$Elapsed[i] <- chgvoltage$Hours[i]-head(chgvoltage[chgvoltage$Index == chgvoltage$Index[i],]$Hours,n=1)
  chgvoltage$TimeRemaining[i] <- -chgvoltage$Hours[i]+tail(chgvoltage[chgvoltage$Index == chgvoltage$Index[i],]$Hours,n=1)
  chgvoltage$Energy[i] <- chgvoltage$PowerSum[i]-head(chgvoltage[chgvoltage$Index == chgvoltage$Index[i],]$PowerSum,n=1)
  chgvoltage$EnergyRemaining[i] <- -chgvoltage$PowerSum[i]+tail(chgvoltage[chgvoltage$Index == chgvoltage$Index[i],]$PowerSum,n=1)
}
chgvoltage <- chgvoltage[!is.na(chgvoltage$Index),]
indexes <- unique(chgvoltage$Index)
parameters<-NULL
for(i in 1:length(indexes)){
  curve <- chgvoltage[chgvoltage$Index==indexes[i],c("SOC","Elapsed","Power","Temp","Time","DOD")]
  y<-curve$Elapsed
  x<-curve$SOC
  smoothed <- lm(y~poly(x,3,raw=TRUE))
  curve$dSOCdt <- 1/(3*smoothed$coefficients[[4]]*curve$SOC^2+2*smoothed$coefficients[[3]]*curve$SOC+smoothed$coefficients[[2]])
  y<-curve$dSOCdt
  a<-0.1
  b<-0.01
  c<--0.5
  model <- nls(log(curve$dSOCdt) ~ log(a)+c*log(curve$SOC-b),data=curve[,c("SOC","dSOCdt")],upper=c(1000,min(x)-.001,0),lower=c(0,0,-1),start=c(a=a,b=b,c=c),algorithm="port",control = nls.control(warnOnly = TRUE))
  a<- model$m$getAllPars()["a"]
  b<- model$m$getAllPars()["b"]
  c<- model$m$getAllPars()["c"]
  parameters <- rbind(parameters,list(head(curve$Time,1),indexes[i],mean(curve$Power),mean(curve$Temp),a,b,c))
  }
# Make a data table for points where power is tapering
distaper <- avistadata[avistadata$PeakFrac<taperhigh & avistadata$Power>0 & avistadata$SOC < 0.80,]
chgtaper <- avistadata[avistadata$PeakFrac<taperhigh & avistadata$Power<0 & avistadata$SOC > 0.80,]
#Output the power taper to csv file
distaper <- distaper %>% group_by(Index)
chgtaper <- chgtaper %>% group_by(Index)
tapertimes <- summarize(distaper,Start=min(Hours))$Start
taperenergies <- summarize(distaper,Energy=min(PowerSum))$Energy
distaper$Elapsed <- 0
distaper$Energy <- 0
for(i in 1:nrow(distaper)){
  distaper$Elapsed[i] <- distaper$Hours[i]-max(tapertimes[tapertimes<=distaper$Hours[i]])
  distaper$Energy[i] <- distaper$PowerSum[i]-max(taperenergies[taperenergies<=distaper$PowerSum[i]])
}
glimpse(distaper)
tapersummary <- distaper %>% summarize(Start=min(Time), St=head(State,n=1), Energy = sum(Power)*10/60/60/1000, SOCMin=min(SOC),SOCMax=max(SOC[SOC<80]),DOD=SOCMax-SOCMin, Duration = max(Hours)-min(Hours),datapoints=length(Power),Slope=(if(datapoints>1){coef(lsfit(SOC,Power))[[2]]} else{0}),Intercept=(if(datapoints>1){coef(lsfit(SOC,Power))[[1]]} else{Intercept=0}))
filename <- paste(gsub(".csv", "", fname),"Summary.xlsx")
#plot(avistadata$Time,avistadata$Power) 
wb <- createWorkbook()
sheet <- addWorksheet(wb, sheetName = "Cycle Summary")
writeDataTable(wb, sheet, summary)
sheet <- addWorksheet(wb, sheetName = "Taper Summary")
writeDataTable(wb, sheet, tapersummary)
sheet <- addWorksheet(wb, sheetName = "Dis Tapers")
writeDataTable(wb, sheet, distaper)
sheet <- addWorksheet(wb, sheetName = "Chg Tapers")
writeDataTable(wb, sheet, chgtaper)
sheet <- addWorksheet(wb, sheetName = "Discharge Periods")
writeDataTable(wb, sheet, disvoltage)
sheet <- addWorksheet(wb, sheetName = "Charge Periods")
writeDataTable(wb, sheet, chgvoltage)
sheet <- addWorksheet(wb, sheetName = "RawData")
writeDataTable(wb, sheet, avistadata)
saveWorkbook(wb, filename, overwrite = TRUE)
 