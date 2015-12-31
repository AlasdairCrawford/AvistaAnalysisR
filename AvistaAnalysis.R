#Load Libraries
library(gdata)
library(dplyr)
library(lubridate)
library(r2excel)
options(java.parameters = "-Xmx8000m")
#Load csv file
fname <- file.choose()
rawdata <- read.csv(fname,sep=",",header=TRUE)
#Filter out rows without data
rawdata <- filter(rawdata,TES.CONN.ACTIVE_PWR.W !="No Data")
#Only take the columns we are interested in raw
if("TES.DC_INVERTER.V" %in% colnames(rawdata) ) {
avistadata <- select(rawdata,Time,TES.STATE_OF_CHARGE_CAPY.PERCENT,TES.CONN.ACTIVE_PWR.W,TES.DC_INVERTER.V)
colnames(avistadata) <- c("Time","SOC","Power","Voltage")
avistadata$Voltage <- as.numeric(as.character(avistadata$Voltage))
} else {
  avistadata <- select(rawdata,Time,TES.STATE_OF_CHARGE_CAPY.PERCENT,TES.CONN.ACTIVE_PWR.W)
  colnames(avistadata) <- c("Time","SOC","Power")
}
#Name the columns for convenience
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
if("Voltage" %in% colnames(avistadata)) {avistadata$DCEff <- (avistadata$Voltage/avistadata$OCV)^(avistadata$State)}
avistadata$PowerSum <- cumsum(avistadata$Power)*10/60/60/1000
disvoltage <- avistadata[avistadata$State == 1 & avistadata$DCEff < 1,]
chgvoltage <- avistadata[avistadata$State == -1 & avistadata$DCEff < 1,]
cycles <- avistadata %>% group_by(Index)
summary <- cycles %>% summarize(Start=min(Time), St=head(State,n=1), Energy = sum(Power)*10/60/60/1000, ConstEnergy = sum(Power[abs(Power)>(dispower*taperhigh)])*10/60/60/1000, TaperEnergy=Energy-ConstEnergy,MaxPower = max(abs(Power))/1000, TaperStart=max(SOC[Power<dispower*taperhigh]),SOCMin=min(SOC),SOCMax=max(SOC),DOD=SOCMax-SOCMin, Duration = max(Hours)-min(Hours))
summary <- filter(summary, DOD>0 & St != 0)
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
wb <- createWorkbook(type="xlsx")
sheet <- createSheet(wb, sheetName = "Cycle Summary")
xlsx.addTable(wb, sheet, data=summary)
sheet <- createSheet(wb, sheetName = "Taper Summary")
xlsx.addTable(wb, sheet, data=tapersummary)
sheet <- createSheet(wb, sheetName = "Tapers")
xlsx.addTable(wb, sheet, data=distaper)
sheet <- createSheet(wb, sheetName = "Dis Eff")
xlsx.addTable(wb, sheet, data=disvoltage)
sheet <- createSheet(wb, sheetName = "Chg Eff")
xlsx.addTable(wb, sheet, data=chgvoltage)
saveWorkbook(wb, filename)
#write.csv(distaper,file="Taper.csv",row.names=FALSE)
#write.csv(summary,file="Summary.csv",row.names=FALSE)
#write.csv(tapersummary,file="TaperSummary.csv",row.names=FALSE)
 