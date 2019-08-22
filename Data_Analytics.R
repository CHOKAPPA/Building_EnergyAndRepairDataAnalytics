# Initialize --------------------------------------------------------------
graphics.off()
rm(list=ls(pattern="^PP[.]"));gc()

# Load Libraries ----------------------------------------------------------
library(readxl)
library(openxlsx)
library(plyr)
library(dplyr)
library(data.table)
library(reshape2)
library(mclust)
library(arules)
library(arulesViz)
library(stringr)
library(magrittr)
library(dbplyr)
library(lubridate)

# Declare: Constants ------------------------------------------------------



# Load: Dataset -----------------------------------------------------------
PP.Dataset_FilePath   <- "C:/Users/E5251120/Desktop/Training_Material/Dataset.xlsx"
PP.Dataset_TableName  <- c("weather_history","work_order_history","work_order_grouping")
PP.Dataset_FileInfo   <- lapply(PP.Dataset_TableName, function(x) 
  as.data.frame(read_excel(path = PP.Dataset_FilePath,sheet=x,col_names = TRUE,col_types = NULL,skip = 0),row.names = NULL,guess_max = 5000))
names(PP.Dataset_FileInfo) <- c("PP.Info_Weather","PP.Info_WorkOrder","PP.Info_WorkOrder_Grouping")
list2env(PP.Dataset_FileInfo,.GlobalEnv)

# Load & PreProcess - Energy Data ----------------------------------
PP.Info_Energy_InWF <- as.data.frame(read_excel(path = PP.Dataset_FilePath,sheet= "energy_history",
                                                col_names = TRUE,col_types = c("text","numeric","numeric","numeric","numeric","numeric"),
                                                trim_ws = TRUE,skip = 0))
PP.Info_Energy_InWF$timestamp  <- as.POSIXct(str_replace(str_sub(PP.Info_Energy_InWF$timestamp,1,19),"T"," "), tz = "Australia/Melbourne", format = "%Y-%m-%d %H:%M:%OS")
PP.Info_Energy             <- melt(PP.Info_Energy_InWF, id.vars = c("timestamp"))
colnames(PP.Info_Energy)   <- c("Timestamp","Building_ID","kWH")
PP.Info_Energy             <- PP.Info_Energy[(!is.na(PP.Info_Energy$kWH)),]
PP.Info_Energy$Building_Ix <- as.integer(str_sub(PP.Info_Energy$Building_ID,10,10))
PP.Info_Energy             <- PP.Info_Energy[,c( "Timestamp","Building_Ix","Building_ID","kWH")]

PP.Info_Energy_Boxplot       <- boxplot(kWH ~ Building_Ix,data = PP.Info_Energy,plot = TRUE)
PP.Info_Energy_Boxplot_Stats <- as.data.frame(t(PP.Info_Energy_Boxplot$stats))
rownames(PP.Info_Energy_Boxplot_Stats) <- PP.Info_Energy_Boxplot$names
colnames(PP.Info_Energy_Boxplot_Stats) <- c("LT","Q1","Q2","Q3","UT")
PP.Info_Energy_Boxplot_Stats$Building_Ix  <- as.integer(rownames(PP.Info_Energy_Boxplot_Stats))

PP.Info_Energy <- PP.Info_Energy %>% left_join(PP.Info_Energy_Boxplot_Stats, by = c('Building_Ix'))
PP.Info_Energy$kWH_Valid <- ifelse((PP.Info_Energy$kWH >= PP.Info_Energy$LT) & (PP.Info_Energy$kWH <= PP.Info_Energy$UT),1,0)
PP.Info_Energy_Boxplot       <- boxplot(kWH ~ Building_Ix,data = PP.Info_Energy[(PP.Info_Energy$kWH_Valid == 1),] ,plot = TRUE,
                                        main = "Boxplot - Avg. Hourly Energy (kWh) Consumption By Building",xlab = "Building #", ylab = "Energy (kWh)")

# Load & PreProcess - Weather Data ----------------------------------
PP.Info_Weather$timestamp  <- as.POSIXct(str_replace(str_sub(PP.Info_Weather$timestamp,1,19),"T"," "), tz = "Australia/Melbourne", format = "%Y-%m-%d %H:%M:%OS")
colnames(PP.Info_Weather)  <- c("Timestamp","Temperature_DegC","RelativeHumdity_Prcnt","DewPoint_DegC")

# PreProcess - Work Order / Failure Data ----------------------------------
PP.Info_Failures <- PP.Info_WorkOrder %>% 
  mutate(Failure_DateTime = as.POSIXct(as.character(`Failure DateTime`), tz = "Australia/Melbourne", format = "%Y-%m-%d %H:%M:%OS")) %>% 
  arrange(location,Failure_DateTime) %>% 
  group_by(location) %>% 
  mutate(Hrs_SinceLastRepair = as.numeric(round(difftime(Failure_DateTime,lag(Failure_DateTime,order_by = Failure_DateTime),units = "hours"),1))) %>% 
  mutate(Hrs_SinceLastRepair = ifelse(Hrs_SinceLastRepair > 0,Hrs_SinceLastRepair,NA)) %>% 
  ungroup()

PP.Hrs_SinceLastFailure_Boxplot       <- boxplot(PP.Info_Failures$Hrs_SinceLastRepair ~ PP.Info_Failures$location,plot = FALSE)
PP.Hrs_SinceLastFailure_Boxplot_Stats <- as.data.frame(t(PP.Hrs_SinceLastFailure_Boxplot$stats))
rownames(PP.Hrs_SinceLastFailure_Boxplot_Stats) <- PP.Hrs_SinceLastFailure_Boxplot$names
colnames(PP.Hrs_SinceLastFailure_Boxplot_Stats) <- c("LT","Q1","Q2","Q3","UT")
PP.Hrs_SinceLastFailure_Boxplot_Stats$location  <- rownames(PP.Hrs_SinceLastFailure_Boxplot_Stats)

PP.Info_Failures                            <- PP.Info_Failures %>% left_join(PP.Hrs_SinceLastFailure_Boxplot_Stats, by = c("location"))
PP.Info_Failures$Hrs_SinceLastRepair_Valid  <- ifelse((PP.Info_Failures$Hrs_SinceLastRepair > PP.Info_Failures$LT) & (PP.Info_Failures$Hrs_SinceLastRepair < PP.Info_Failures$UT),1,0)
PP.Info_Failures$Building_Ix                <- as.numeric(str_sub(PP.Info_Failures$location,-1,-1))

PP.Info_Hrs_SinceLastFailure       <- as.data.frame(PP.Info_Failures %>% 
  filter((Hrs_SinceLastRepair_Valid == 1) & (Hrs_SinceLastRepair > 1)) %>% 
  select(Building_Ix,Hrs_SinceLastRepair) %>% 
  arrange(Building_Ix,Hrs_SinceLastRepair))

PP.Info_Hrs_SinceLastFailure_Model <- densityMclust(PP.Info_Hrs_SinceLastFailure$Hrs_SinceLastRepair)
summary(PP.Info_Hrs_SinceLastFailure_Model)
PP.Info_Hrs_SinceLastFailure$Classification <- PP.Info_Hrs_SinceLastFailure_Model$classification

# Build - Failure Group and Idenify Snapshot Frame - Before & After Failure Group --------
PP.Info_Failures <- PP.Info_Failures %>% 
  mutate(Group_Condition = ifelse((is.na(Hrs_SinceLastRepair) == TRUE) | (Hrs_SinceLastRepair <= (3*24)),0,1)) %>% 
  group_by(Building_Ix) %>% 
  mutate(Group_Sq = (Building_Ix * 10000) + cumsum(Group_Condition)) %>%
  ungroup()

PP.Info_Failures <- PP.Info_Failures %>% 
  arrange(Building_Ix,Group_Sq) %>% 
  group_by(Building_Ix,Group_Sq) %>%
  mutate(FailureGroup_DateTime_BEG  = min(Failure_DateTime,na.rm = TRUE),
         FailureGroup_DateTime_END  = max(Failure_DateTime,na.rm = TRUE)) %>% 
  mutate(BeforeFailure_DateTime_BEG = FailureGroup_DateTime_BEG - hours(36),
         BeforeFailure_DateTime_END = FailureGroup_DateTime_BEG - hours(4),
         AfterFailure_DateTime_BEG  = FailureGroup_DateTime_END + hours(4),
         AFterFailure_DateTime_END  = FailureGroup_DateTime_END + hours(36)) %>% 
  ungroup()
         
PP.Info_FailuresGroup <- as.data.frame(PP.Info_Failures %>% 
  arrange(Building_Ix,Group_Sq) %>%
  select(Building_Ix,Group_Sq,
         FailureGroup_DateTime_BEG,
         FailureGroup_DateTime_END,
         BeforeFailure_DateTime_BEG,
         BeforeFailure_DateTime_END,
         AfterFailure_DateTime_BEG,
         AFterFailure_DateTime_END) %>%
  distinct())

# Extract - Snapshot Before & After Failure Group For Energy & Weather ------------
PP.Info_FailuresGroup_ByGrSq  <- split(PP.Info_FailuresGroup,PP.Info_FailuresGroup$Group_Sq)
PP.Info_FailuresGroup_Weather <- do.call(rbind.fill,lapply(PP.Info_FailuresGroup_ByGrSq, function(x) {
  # rm(list=ls(pattern="^Fn[.]"))
  # x <- PP.Info_FailuresGroup_ByGrSq[[1]]
  
  Fn.Info_Weather_BeforeFailure       <- as.data.frame(PP.Info_Weather %>% filter((PP.Info_Weather$Timestamp >= x$BeforeFailure_DateTime_BEG) & (PP.Info_Weather$Timestamp <= x$BeforeFailure_DateTime_END)))
  Fn.Info_Weather_BeforeFailure$Type  <- if (nrow(Fn.Info_Weather_BeforeFailure) > 0) {"Before Failure"}
  
  Fn.Info_Weather_AfterFailure        <- as.data.frame(PP.Info_Weather %>% filter((PP.Info_Weather$Timestamp >= x$AfterFailure_DateTime_BEG)  & (PP.Info_Weather$Timestamp <= x$AFterFailure_DateTime_END)))
  Fn.Info_Weather_AfterFailure$Type   <- if (nrow(Fn.Info_Weather_AfterFailure) > 0) {"After Failure"}
  
  Fn.VARsOI_Weather <- c("Timestamp","Temperature_DegC","RelativeHumdity_Prcnt","DewPoint_DegC","Type")
  Fn.VARsOI_Group  <- c( "Building_Ix","Group_Sq")
  
  Fn.Info_Weather <- rbind.fill(Fn.Info_Weather_BeforeFailure,Fn.Info_Weather_AfterFailure)
  Fn.Info_Weather <- if (nrow(Fn.Info_Weather) > 0) {cbind(x[,Fn.VARsOI_Group],Fn.Info_Weather[,Fn.VARsOI_Weather])} else {NULL}
  return(Fn.Info_Weather)
}))
PP.Info_FailuresGroup_Energy  <- do.call(rbind.fill,lapply(PP.Info_FailuresGroup_ByGrSq, function(x) {
  # rm(list=ls(pattern="^Fn[.]"))
  # x <- PP.Info_FailuresGroup_ByGrSq[[83]]
  
  Fn.Info_Energy_BeforeFailure       <- as.data.frame(PP.Info_Energy %>% filter(
    (as.integer(PP.Info_Energy$Building_Ix) == as.integer(x$Building_Ix))
    & (PP.Info_Energy$Timestamp >= x$BeforeFailure_DateTime_BEG)
    & (PP.Info_Energy$Timestamp <= x$BeforeFailure_DateTime_END)))
  Fn.Info_Energy_BeforeFailure$Type <- if (nrow(Fn.Info_Energy_BeforeFailure) > 0) {"Before Failure"}
  
  Fn.Info_Energy_AfterFailure        <- as.data.frame(PP.Info_Energy %>% filter(
    (as.integer(PP.Info_Energy$Building_Ix) == as.integer(x$Building_Ix))
    & (PP.Info_Energy$Timestamp >= x$AfterFailure_DateTime_BEG)
    & (PP.Info_Energy$Timestamp <= x$AFterFailure_DateTime_END)))
  Fn.Info_Energy_AfterFailure$Type   <- if (nrow(Fn.Info_Energy_AfterFailure) > 0) {"After Failure"}
  
  Fn.VARsOI_Energy <- c("Timestamp","kWH","kWH_Valid","Type")
  Fn.VARsOI_Group  <- c( "Building_Ix","Group_Sq")
  
  Fn.Info_Energy <- rbind.fill(Fn.Info_Energy_BeforeFailure,Fn.Info_Energy_AfterFailure)
  Fn.Info_Energy <- if (nrow(Fn.Info_Energy) > 0) {cbind(x[,Fn.VARsOI_Group],Fn.Info_Energy[,Fn.VARsOI_Energy])} else {NULL}  
}))

# Mine - Parts that are repaired together ----------------------------------
PP.Info_TRNs  <- as.data.frame(PP.Info_Failures %>% 
                                 select(TRNs_ID = Group_Sq,TRNs_ITMs = `Failure Area`) %>% 
                                 distinct() %>% 
                                 filter(TRNs_ITMs %in% c("Unknown","Other") == FALSE) %>%
                                 arrange(TRNs_ID,TRNs_ITMs))
                                 

FUNC.Transactions <- as(split(PP.Info_TRNs$TRNs_ITMs,PP.Info_TRNs$TRNs_ID),"transactions")
summary(FUNC.Transactions)

#*** PLOT: Item Frequency Bar Chart
FUNC.ItmFrq  <- sort(round(itemFrequency(FUNC.Transactions,type="relative"),4))
FUNC.ItmFrq  <- FUNC.ItmFrq[FUNC.ItmFrq > 0]
itemFrequencyPlot(FUNC.Transactions, topN = 20,type="relative",names="TRUE",xlab="Replaced Parts",ylab="Percentage Items",ylim=c(0,1))

#*** DECLARE: "arules" Parameters
FUNC.suppMIN	<- 0.1
FUNC.confMIN	<- 0.6
FUNC.minLEN		<- 2
FUNC.maxLEN		<- 5
FUNC.aRules	<- apriori(FUNC.Transactions, parameter=list(supp=FUNC.suppMIN,conf=FUNC.confMIN, minlen=FUNC.minLEN,maxlen=FUNC.maxLEN),appearance = NULL,control = list(verbose = FALSE))

#*** VISUALIZE: RULES
inspect(FUNC.aRules)
plot(FUNC.aRules, method="grouped")
plot(FUNC.aRules, method="graph")

# WRITE: Output into Excel File -------------------------------------------
wb <-createWorkbook(creator = ifelse(.Platform$OS.type == "windows",Sys.getenv("USERNAME"), Sys.getenv("USER")),
                    title = "BUENO Case Study",
                    subject = "Building Energy Managament & Maintenace",
                    category = NULL)
addWorksheet(wb, "Info Energy")
addWorksheet(wb, "Info Energy  By FailureGroup")
addWorksheet(wb, "Info Weather")
addWorksheet(wb, "Info Weather By FailureGroup")
addWorksheet(wb, "Info Failures")
addWorksheet(wb, "Info Failure Groups")

writeData(wb, sheet = "Info Energy" ,PP.Info_Energy)
writeData(wb, sheet = "Info Energy  By FailureGroup",PP.Info_FailuresGroup_Energy)
writeData(wb, sheet = "Info Weather",PP.Info_Weather )
writeData(wb, sheet = "Info Weather By FailureGroup", PP.Info_FailuresGroup_Weather)
writeData(wb, sheet = "Info Failures",PP.Info_Failures)
writeData(wb, sheet = "Info Failure Groups",PP.Info_FailuresGroup)

saveWorkbook(wb,file.path(dirname(PP.Dataset_FilePath),"Dataset_ForAnalytics.xlsx",fsep = .Platform$file.sep), overwrite = TRUE)

a <- PP.Info_FailuresGroup_Energy %>% 
  group_by(Group_Sq) %>% 
  summarise(COV = (var(kWH,na.rm = TRUE)^0.5) / mean(kWH,na.rm = TRUE))
              

head(a)
