#Load Libraries
library(gdata)
library(dplyr)
library(lubridate)
library(openxlsx)
Sys.setenv(R_ZIPCMD= "C:/Rtools/bin/zip")
#Load csv file
fname <- file.choose()
rawdata <- read.csv(fname,sep=",",header=TRUE)
#Filter out rows without data
rawdata <- filter(rawdata,TES.CONN.ACTIVE_PWR.W !="No Data")
#Only take the columns we are interested in raw
if("TES.DC_INVERTER.V" %in% colnames(rawdata) ) {
avistadata <- select(rawdata,Time,TES.STATE_OF_CHARGE_CAPY.PERCENT,TES.CONN.ACTIVE_PWR.W,TES.DC_INVERTER.V,TES.DC_INVERTER.IN_PWR.W)
colnames(avistadata) <- c("Time","SOC","Power","Voltage","InverterPower")
avistadata$Voltage <- as.numeric(as.character(avistadata$Voltage))
} else {
  avistadata <- select(rawdata,Time,TES.STATE_OF_CHARGE_CAPY.PERCENT,TES.CONN.ACTIVE_PWR.W)
  colnames(avistadata) <- c("Time","SOC","Power")
}
#avistadata$Temp <- 0.25*(as.character(rawdata$TES_BATTERY1.STORAGE.TEMP.MAX.DEGC)+as.character(rawdata$TES_BATTERY1.STORAGE.TEMP.MIN.DEGC)+as.character(rawdata$TES_BATTERY2.STORAGE.TEMP.MAX.DEGC)+as.character(rawdata$TES_BATTERY2.STORAGE.TEMP.MIN.DEGC))
avistadata$Temp <- rowMeans(rawdata[,c("TES_BATTERY1.STORAGE.TEMP.MAX.DEGC","TES_BATTERY2.STORAGE.TEMP.MIN.DEGC")])
#avistadata$Temp <- 0
#Name the columns for convenience
avistadata$ConPower <- avistadata$Power
avistadata$Power<-(as.numeric(as.character(rawdata$TES.BATTERY1.KW))+as.numeric(as.character(rawdata$TES.BATTERY2.KW)))*1000
#Convert from text into numerics
avistadata$Power <- as.numeric(as.character(avistadata$Power))
avistadata$SOC <- as.numeric(as.character(avistadata$SOC))
avistadata$Time <- as.POSIXct(as.character(avistadata$Time),format="%m/%d/%Y %H:%M")
#Add an hours column for convenience
avistadata$Hours <- seq(from=0,length.out=length(avistadata$Time),by=(1/360))
#Get rid of rows without any data
avistadata <- avistadata[!is.na(avistadata$Power),]
#Look at the data

#Find the max discharge and charge power
dispower <- max(avistadata$Power, na.rm=TRUE)
chgpower <- min(avistadata$Power, na.rm=TRUE)
#Choose conditions to count as taper power
taperhigh <- 0.95
taperlow <- 0.05
dispower
chgpower
avistadata$State <- avistadata$Power %>% sapply(FUN = function(x) (abs(x) > taperlow*dispower+0)*sign(x))
avistadata$StChange <- c(0,abs(diff(avistadata$State)))
avistadata$Index <- cumsum(avistadata$StChange)
avistadata$OCV <- 1.4176*avistadata$SOC+743.91
avistadata$InvEff <- (avistadata$ConPower/as.numeric(avistadata$InverterPower))^(avistadata$State)
if("Voltage" %in% colnames(avistadata)) {
  avistadata$DCEff <- (avistadata$Voltage/avistadata$OCV)^(avistadata$State)
  avistadata$TotEff <- avistadata$DCEff*avistadata$InvEff
  }
avistadata$PowerSum <- cumsum(avistadata$Power)*10/60/60/1000
disvoltage <- avistadata[avistadata$State == 1 & avistadata$Power > dispower*taperhigh,]
chgvoltage <- avistadata[avistadata$State == -1 & avistadata$Power < chgpower*taperhigh,]

cycles <- avistadata %>% group_by(Index)
summary <- cycles %>% summarize(Start=min(Time), St=head(State,n=1), Energy = sum(Power)*10/60/60/1000, ConstEnergy = sum(Power[abs(Power)>(dispower*taperhigh)])*10/60/60/1000, TaperEnergy=Energy-ConstEnergy,MaxPower = max(abs(Power))/1000, TaperStart=max(SOC[Power<dispower*taperhigh]),SOCMin=min(SOC),SOCMax=max(SOC),DOD=SOCMax-SOCMin, Duration = max(Hours)-min(Hours),AvgTemp=mean(Temp))
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
#Make a data table for points where power is tapering
distaper <- avistadata[avistadata$Power<(dispower*taperhigh) & avistadata$Power>(dispower*taperlow) & avistadata$SOC < 80,]
#Output the power taper to csv file
distaper <- distaper %>% group_by(Index)
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
sheet <- addWorksheet(wb, sheetName = "Tapers")
writeDataTable(wb, sheet, distaper)
sheet <- addWorksheet(wb, sheetName = "Discharge Periods")
writeDataTable(wb, sheet, disvoltage)
sheet <- addWorksheet(wb, sheetName = "Charge Periods")
writeDataTable(wb, sheet, chgvoltage)
saveWorkbook(wb, filename, overwrite = TRUE)
 