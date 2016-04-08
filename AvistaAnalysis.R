#Load Libraries
library(gdata)
library(dplyr)
library(lubridate)
library(openxlsx)
Sys.setenv(R_ZIPCMD= "C:/Rtools/bin/zip")
#Load csv file
fname <- file.choose()
rawdata <- read.csv(fname,sep=",",header=TRUE)
wb <- createWorkbook()
#rawdata <- rawdata[seq(1,NROW(rawdata),by=10),]
#Filter out rows without data
#Only take the columns we are interested in raw
avistadata <- select(rawdata,DATETIME_STAMP,TES_STATE_OF_CHARGE_CAPY_PERCENT)
colnames(avistadata) <- c("Time","SOC")
avistadata$Temp <- rowMeans(rawdata[,c("TES_BATTERY1_STORAGE_TEMP_MAX_DEGC","TES_BATTERY1_STORAGE_TEMP_MIN_DEGC","TES_BATTERY2_STORAGE_TEMP_MAX_DEGC","TES_BATTERY2_STORAGE_TEMP_MIN_DEGC")])
avistadata$Power<-(as.numeric(as.character(rawdata$TES_BATTERY1_KW))+as.numeric(as.character(rawdata$TES_BATTERY2_KW)))*1000
#Convert from text into numerics
avistadata$Power <- as.numeric(as.character(avistadata$Power))
avistadata$SOC <- as.numeric(as.character(avistadata$SOC))/100
avistadata$Time <- as.POSIXct(as.character(avistadata$Time),format="%Y-%m-%d %T")
if(is.na(head(avistadata$Time,1))){
  avistadata$Time <- as.POSIXct(as.character(rawdata$DATETIME_STAMP),format="%m/%d/%Y %H:%M")
}
avistadata$Time <- seq(from=head(avistadata$Time,1),by=10,length.out=nrow(avistadata))
#Add an hours column for convenience
#avistadata$Hours <- seq(from=0,length.out=length(avistadata$Time),by=(0.00277777777777778))

avistadata$Hours <- as.numeric(avistadata$Time)/60/60
avistadata$Hours <- avistadata$Hours - avistadata$Hours[1]

#seq <- rle(avistadata$SOC)$lengths
#SOCmid <- round(cumsum(seq)-seq/2)
#y <- avistadata$SOC[SOCmid]
#x <- avistadata$Hours[SOCmid]
#avistadata$SOC <- approx(x,y,xout=avistadata$Hours)$y

#avistadata$SOC <- spline(x,y,xout=avistadata$Hours)$y
#avistadata$SOC <- loess(avistadata$SOC ~ avistadata$Hours,span=0.1)$fitted

#Get rid of rows without any data
avistadata <- avistadata[!is.na(avistadata$Power),]
#Choose conditions to count as taper power
taperhigh <- 0.95
taperlow <- 0.05
dispower <- max(avistadata$Power)
chgpower <- min(avistadata$Power)
#avistadata$State <- avistadata$Power %>% sapply(FUN = function(x) (abs(x) > taperlow*dispower+0)*sign(x))
avistadata$State <- avistadata$Power %>% sapply(FUN = function(x) (abs(x) > 75000+0)*sign(x))
avistadata$StChange <- c(0,1*(diff(avistadata$State)!=0))
avistadata$Index <- cumsum(avistadata$StChange)
avistadata$PowerSum <- cumsum(avistadata$Power)*10/60/60/1000
#disvoltage <- avistadata[avistadata$State == 1 & avistadata$Power > dispower*taperhigh,]
#chgvoltage <- avistadata[avistadata$State == -1 & avistadata$Power < chgpower*taperhigh,]
disvoltage <- avistadata[avistadata$State == 1,]
chgvoltage <- avistadata[avistadata$State == -1,]
avistadata <- group_by(avistadata,Index)
avistadata <- mutate(avistadata,Peak=max(abs(Power))*State,PeakFrac=Power/Peak)
avistadata <- filter(avistadata, State != 0, PeakFrac > taperhigh)
avistadata <- mutate(avistadata,DOD=max(SOC)-min(SOC))
disvoltage <- filter(avistadata,State == 1, DOD > 0.30)
chgvoltage <- filter(avistadata,State == -1, DOD > 0.30)
summary <- avistadata %>% summarize(Start=min(Time), St=head(State,n=1), Energy = sum(Power)*10/60/60/1000,Peak = max(abs(Power)), TaperStart=max(SOC[Power<dispower*taperhigh]),SOCMin=min(SOC),SOCMax=max(SOC),DOD=SOCMax-SOCMin, Duration = max(Hours)-min(Hours),AvgTemp=mean(Temp))
summary <- filter(summary, DOD>0 & St != 0)
disvoltage <- group_by(disvoltage,Index)
disvoltage <- mutate(disvoltage, Elapsed = Hours - min(Hours),TimeRemaining=max(Hours)-Hours,Energy=PowerSum-head(PowerSum),EnergyRemaining=tail(PowerSum)-PowerSum)
chgvoltage <- group_by(chgvoltage,Index)
chgvoltage <- mutate(chgvoltage, Elapsed = Hours - min(Hours),TimeRemaining=max(Hours)-Hours,Energy=PowerSum-head(PowerSum),EnergyRemaining=tail(PowerSum)-PowerSum)

chgsummary <- chgvoltage %>% summarize(Start=min(Time), St=head(State,n=1), Energy = sum(Power)*10/60/60/1000,Peak = max(abs(Power))/1000, TaperStart=max(SOC[Power<dispower*taperhigh]),SOCMin=min(SOC),SOCMax=max(SOC),DOD=SOCMax-SOCMin, Duration = max(Hours)-min(Hours),AvgTemp=mean(Temp))
chgsummary$Peak[chgsummary$Peak==316 | chgsummary$Peak==265 | chgsummary$Peak==314 | chgsummary$Peak==315]=chgsummary$Peak[chgsummary$Peak==316 | chgsummary$Peak==265 | chgsummary$Peak==314 | chgsummary$Peak==315]*2
fitted_models <- chgvoltage %>% group_by(Index) %>% do(model = smooth.spline(.$Elapsed ~ .$SOC))
predictions <- lapply(fitted_models$model,predict,deriv=1)
predictions <- lapply(predictions,as.data.frame)
predictions <- lapply(predictions,mutate,SOC=x,dSOCdt = 1/y)
predictions <- lapply(predictions,select,SOC,dSOCdt)
predictions <- lapply(predictions,filter,dSOCdt>0,dSOCdt<1)
nonlinearmodels <- lapply(predictions, nls,formula= log(SOC) ~ (1/c)*log(dSOCdt/a),upper=list(100,-.01),lower=c(0,-2),start=list(a=0.1,c=-0.5),algorithm="port",control = nls.control(warnOnly = TRUE))
parameters <- lapply(nonlinearmodels,coef)
parameters <- as.data.frame(t(as.data.frame(parameters)))
Indexes <- unique(chgvoltage$Index)
curves <- fitted_models$model
chgsummary <- cbind(chgsummary,parameters)
chgwb <- createWorkbook()
sheet <- addWorksheet(chgwb, sheetName = "Charge Summary")
nonlinearpredicted <- NULL
generalmodel <- NULL
writeDataTable(chgwb, sheet, chgsummary)
#dev.off()
for(i in 1:length(Indexes)){
  plot(fitted_models$model[[i]]$x,fitted_models$model[[i]]$yin,xlab="SOC",ylab="Elapsed Time",xlim=c(0,1),ylim=c(0,10),main=chgsummary$Start[i])
  lines(curves[[i]],col=2,lwd=2)
  sheet <- addWorksheet(chgwb,sheetName=chgsummary$Index[i])
  writeDataTable(chgwb, sheet, chgsummary[i,])
  insertPlot(chgwb, sheet, 2, width = 5, height = 4, fileType = "png", units = "in",xy=c(1,3))
  title <- paste(round(chgsummary$Peak[i],-1), " kW ", round(chgsummary$AvgTemp[i]), " C")
  plot(predictions[[i]],xlim=c(0,1),ylim=c(0,1),col=1,main=title)
  nonlinearpredicted$x <- seq(chgsummary$SOCMin[i],chgsummary$SOCMax[i],by=0.01)
  nonlinearpredicted$y <- parameters[i,1]*(nonlinearpredicted$x)^(parameters[i,2])
  lines(nonlinearpredicted,col=2,lwd=2)
  a<-((7.35E-05)+(2.06E-08)*chgsummary$Peak[i]+(5.54E-07)*(chgsummary$AvgTemp[i]))*chgsummary$Peak[i]
  c<-(-1.352544791)+(0.000389547)*chgsummary$Peak[i]+(0.013571239)*(chgsummary$AvgTemp[i])
  generalmodel$x <- nonlinearpredicted$x
  generalmodel$y <- a*(nonlinearpredicted$x)^c
  lines(generalmodel,col=4,lwd=2)
  #lines(chgvoltage$SOC[chgvoltage$Index==Indexes[i]],chgvoltage$PeakFrac[chgvoltage$Index==Indexes[i]],col=3)
  insertPlot(chgwb, sheet, 2, width = 5, height = 4, fileType = "png", units = "in",xy=c(9,3))
}
filename <- paste(gsub(".csv", "", fname)," Charge Summary.xlsx")
saveWorkbook(chgwb, filename, overwrite = TRUE)

dissummary <- disvoltage %>% summarize(Start=min(Time), St=head(State,n=1), Energy = sum(Power)*10/60/60/1000,Peak = max(abs(Power))/1000, TaperStart=max(SOC[Power<dispower*taperhigh]),SOCMin=min(SOC),SOCMax=max(SOC),DOD=SOCMax-SOCMin, Duration = max(Hours)-min(Hours),AvgTemp=mean(Temp))
dissummary$Peak[dissummary$Peak==246 | dissummary$Peak==247]=dissummary$Peak[dissummary$Peak==246 | dissummary$Peak==247]*2
dissummary$Peak[dissummary$Peak==316]=dissummary$Peak[dissummary$Peak==316]*2
fitted_models <- disvoltage %>% group_by(Index) %>% do(model = smooth.spline(.$Elapsed ~ .$SOC))
predictions <- lapply(fitted_models$model,predict,deriv=1)
predictions <- lapply(predictions,as.data.frame)
predictions <- lapply(predictions,mutate,SOC=x,dSOCdt = abs(1/y))
predictions <- lapply(predictions,select,SOC,dSOCdt)
predictions <- lapply(predictions,filter,dSOCdt>0,dSOCdt<1)
#nonlinearmodels <- lapply(predictions, nls,formula= log(SOC) ~ log(dSOCdt^(1/c)*a^(-1/c)+b),upper=list(100,1,-.01),lower=c(0,0,-1),start=list(a=0.01,b=0,c=-0.5),algorithm="port",control = nls.control(warnOnly = TRUE))
nonlinearmodels <- lapply(predictions, nls,formula= log(dSOCdt) ~ c*log(abs(SOC-b)+.0001)+log(a),upper=list(100,1,-.01),lower=c(0,0,-1),start=list(a=0.1,b=0.15,c=-0.3),algorithm="port",control = nls.control(warnOnly = TRUE))
parameters <- lapply(nonlinearmodels,coef)
parameters <- as.data.frame(t(as.data.frame(parameters)))
Indexes <- unique(disvoltage$Index)
curves <- fitted_models$model
dissummary <- cbind(dissummary,parameters)
diswb <- createWorkbook()
sheet <- addWorksheet(diswb, sheetName = "Disharge Summary")
writeDataTable(diswb, sheet, dissummary)
#dev.off()
for(i in 1:length(Indexes)){
  plot(fitted_models$model[[i]]$x,fitted_models$model[[i]]$yin,xlab="SOC",ylab="Elapsed Time",xlim=c(0,1),ylim=c(0,10),main=dissummary$Start[i])
  lines(curves[[i]],col=2,lwd=2)
  sheet <- addWorksheet(diswb,sheetName=dissummary$Index[i])
  writeDataTable(diswb, sheet, dissummary[i,])
  insertPlot(diswb, sheet, 2, width = 5, height = 4, fileType = "png", units = "in",xy=c(1,3))
  title <- paste(round(dissummary$Peak[i],-1), " kW ", round(dissummary$AvgTemp[i]), " C")
  plotname <- paste("displot_",dissummary$Index[i],".png")
  mypath <- file.path("C:","Users","craw038","My Documents", "DischargeAnimation",plotname)
  #png(mypath)
  plot(predictions[[i]],xlim=c(0,1),ylim=c(0,0.5),col=1,main=title,xlab="SOC",ylab=expression('dSOC/dt (h'^-1*')'))
  nonlinearpredicted$x <- seq(dissummary$SOCMin[i],dissummary$SOCMax[i],by=0.01)
  nonlinearpredicted$y <- parameters[i,1]*(nonlinearpredicted$x-parameters[i,2])^(parameters[i,3])
  lines(nonlinearpredicted,col=2,lwd=2)
  a<-((1.169E-04)+(5.868E-08)*dissummary$Peak[i]+(-2.608E-08)*(dissummary$AvgTemp[i]))*dissummary$Peak[i]
  b<-((8.10E-04)+(1.26E-07)*dissummary$Peak[i]+(-9.85E-06)*(dissummary$AvgTemp[i]))*dissummary$Peak[i]
  c<-(-6.527E-01)+(6.018E-04)*dissummary$Peak[i]+(-1.048E-03)*(dissummary$AvgTemp[i])
  generalmodel$x <- nonlinearpredicted$x
  generalmodel$y <- a*(nonlinearpredicted$x-b)^c
  lines(generalmodel,col=4,lwd=2)
  #lines(disvoltage$SOC[disvoltage$Index==Indexes[i]],disvoltage$PeakFrac[disvoltage$Index==Indexes[i]],col=3)
  insertPlot(diswb, sheet, 2, width = 5, height = 4, fileType = "png", units = "in",xy=c(9,3))
  #dev.off()
}
filename <- paste(gsub(".csv", "", fname)," Discharge Summary.xlsx")
saveWorkbook(diswb, filename, overwrite = TRUE)

# Make a data table for points where power is tapering
distaper <- avistadata[avistadata$PeakFrac<taperhigh & avistadata$Power>0 & avistadata$SOC < 0.80,]
chgtaper <- avistadata[avistadata$PeakFrac<taperhigh & avistadata$Power<0 & avistadata$SOC > 0.80,]
#Output the power taper to csv file
distaper <- distaper %>% group_by(Index)
chgtaper <- chgtaper %>% group_by(Index)
tapertimes <- summarize(distaper,Start=min(Hours))$Start
taperenergies <- summarize(distaper,Energy=min(PowerSum))$Energy
#distaper$Elapsed <- 0

#distaper$Energy <- 0
#for(i in 1:nrow(distaper)){
  #distaper$Elapsed[i] <- distaper$Hours[i]-max(tapertimes[tapertimes<=distaper$Hours[i]])
  #distaper$Energy[i] <- distaper$PowerSum[i]-max(taperenergies[taperenergies<=distaper$PowerSum[i]])
#}
#tapersummary <- distaper %>% summarize(Start=min(Time), St=head(State,n=1), Energy = sum(Power)*10/60/60/1000, SOCMin=min(SOC),SOCMax=max(SOC[SOC<80]),DOD=SOCMax-SOCMin, Duration = max(Hours)-min(Hours),datapoints=length(Power),Slope=(if(datapoints>1){coef(lsfit(SOC,Power))[[2]]} else{0}),Intercept=(if(datapoints>1){coef(lsfit(SOC,Power))[[1]]} else{Intercept=0}))
filename <- paste(gsub(".csv", "", fname),"Summary.xlsx")
#plot(avistadata$Time,avistadata$Power) 

sheet <- addWorksheet(wb, sheetName = "Cycle Summary")
writeDataTable(wb, sheet, summary)
#sheet <- addWorksheet(wb, sheetName = "Taper Summary")
#writeDataTable(wb, sheet, tapersummary)
#sheet <- addWorksheet(wb, sheetName = "Dis Tapers")
#writeDataTable(wb, sheet, distaper)
#sheet <- addWorksheet(wb, sheetName = "Chg Tapers")
#writeDataTable(wb, sheet, chgtaper)
#sheet <- addWorksheet(wb, sheetName = "Discharge Periods")
#writeDataTable(wb, sheet, disvoltage)
#sheet <- addWorksheet(wb, sheetName = "Charge Periods")
#writeDataTable(wb, sheet, chgvoltage)
#sheet <- addWorksheet(wb, sheetName = "Fitting Parameters")
#writeDataTable(wb, sheet, parameters)
#sheet <- addWorksheet(wb, sheetName = "RawData")
#writeDataTable(wb, sheet, avistadata)
saveWorkbook(wb, filename, overwrite = TRUE)
 