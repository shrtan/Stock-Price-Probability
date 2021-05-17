library(readxl)
X_GSPC <- read_excel("C:/Users/Shreya/Downloads/^GSPC.xls")
View(X_GSPC)

orig <- X_GSPC

orig$day <- weekdays(as.Date(orig$Date))
orig$week = strftime(orig$Date, format = "%V")
orig$year <- substring(orig$Date,1,4)

#temp <- data.frame(Year='1990', Week='01', Open=100, Close=100)
#temp <- data.frame(Year=orig$year[[1]], Week=orig$week[[1]], Open=orig$Open[[1]], Close=100)
#Clean <- rbind(Clean, temp)

#creating empty year, week, open close df
Clean<-data.frame(Year=character(0),Week=character(0),Open=numeric(0), Close=numeric(0), Difference=numeric(0), Returns=numeric(0), Flag=character(0))

#adding year, week, open, close, difference, returns, flag
prevweek=0
for (i in seq(1, nrow(orig), by=1)) {
  presweek = orig$week[i]
  tempweek = orig$week[i]
  if(prevweek != tempweek)
  {
    tempopen = orig$Open[i]
    tempyear = orig$year[i]
    prevweek = tempweek
  }
  if(orig$week[i+1] != tempweek || i == nrow(orig))
  {
    tempclose = orig$Close[i]
    tempsub <- tempclose - tempopen
    tempreturns <- tempsub/tempopen
    tempflag = 'a'
    if(tempreturns >= 0) {
      tempflag = 'Positive'
    }
    else {
      tempflag = 'Negative'
    }
    temp <- data.frame(Year=tempyear, Week=tempweek, Open=tempopen, Close=tempclose, Difference=tempsub, Returns=tempreturns, Flag=tempflag)
    Clean <- rbind(Clean, temp)
  }
}

#initializing consecutive positive and negative column
poscolumn <- data.frame(Consecutive_Positive=numeric(0))
negcolumn <- data.frame(Consecutive_Negative=numeric(0))

countpos = 0
countneg = 0
poscolumn <- rbind(poscolumn, data.frame(Consecutive_Positive=0))
negcolumn <- rbind(negcolumn, data.frame(Consecutive_Negative=0))
if(Clean$Flag[1]=='Positive') {
  countpos = 1
} else {
  countneg = 1
}

#counting consecutive negative and positive flags
for(i in seq(2, nrow(Clean), by=1)) {
  poscolumn <- rbind(poscolumn, data.frame(Consecutive_Positive=countpos))
  negcolumn <- rbind(negcolumn, data.frame(Consecutive_Negative=countneg))
  
  if(Clean$Flag[i]=='Positive') {
    countpos = countpos+1
    countneg=0
  } else {
    countneg = countneg+1
    countpos=0
  }
}

#adding consective positive and negative columns to Clean
Clean <- cbind(Clean, poscolumn)
Clean <- cbind(Clean, negcolumn)

#total count of weeks with 3 prev positive flags
totalprob3 = 0
prob4 = 0

for(i in seq(3, nrow(Clean), by=1)) {
  if(Clean$Consecutive_Positive[i]==3) {
    totalprob3 = totalprob3+1
    #checking if present week is postive
    if(Clean$Flag[i]=='Positive') {
      prob4 = prob4+1
    }
  }
}

#total count of weeks with 2 prev positive flags
totalprob2 = 0
prob3 = 0

for(i in seq(3, nrow(Clean), by=1)) {
  if(Clean$Consecutive_Positive[i]==2) {
    totalprob2 = totalprob2+1
    #checking if present week is postive
    if(Clean$Flag[i]=='Positive') {
      prob3 = prob3+1
    }
  }
}

prob = prob3/totalprob2 

#no of years
years = strtoi(Clean$Year[nrow(Clean)]) - strtoi(Clean$Year[1]) + 1

library(xlsx)
write.xlsx(Clean, "C:\\Users\\Shreya\\Downloads\\SPData.xlsx")


