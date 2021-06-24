####################### Load Library and SQL Pull #########
library(odbc)
library(reshape2)
library(dplyr)
library(openxlsx)
library(writexl)
library(lubridate)
library(tidyr)


OldRoster<-"DOC MH Roster Philadelphia 2021-Jun24 Unified.xlsx"
NewFile<-"Copy of DOC MH Roster Philadelphia County 6.2021.xlsx"


con <- dbConnect(odbc(), Driver = "SQL Server", Server = "SQL-HQ", 
                 Database = "SCCJU", Trusted_Connection = "True")

CCSRoster <- dbGetQuery(con,"
                        
                        select 
                        [ClientID#] as 'ClientID'
                        ,[SS#] as 'SSN'
                        ,[LastName] 
                        ,[FirstName]
                        ,[DOB]
                        ,[PP#] as 'PID'
                        ,[SID]
                        ,DOC.Inmate_No as 'Inmate.Number'
                        
                        
                        from tblclient as c
                        left join tblDOCentry as DOC on c.ClientID#=DOC.clientID
                        
                        where deleted =0
                        and c.LastName not like '%test%'
                        
                        ")


############# SOURCE FILEs #################
NewPath<-"C:/Users/cmccrea/Box/BHJD/DOC/Rosters/NeedsMerged/"
NewFile<-read.xlsx(paste0(NewPath,NewFile))

OldPath<-"C:/Users/cmccrea/Box/BHJD/DOC/Rosters/" 
OldRoster<-read.xlsx(paste0(OldPath,OldRoster))

rm(OldPath,NewPath)

NewDOCRoster <- subset(NewFile, select= c("inmate_number","Lst_Nm","Frst_Nm","DOB","location_permanent","CurrLoc_Cd","offense","Parole_Violator","Parole_Status_Code","Min_date","Max_Dt"))
colnames(NewDOCRoster)<-c("Inmate.Number","LastName","FirstName","DOB","Location.Permanent","Current.Loc.CD","Offence","Parole.Violator","Parole.Status.Code","Min.Date","Max.Date")

OldDOCRoster <- subset(OldRoster, select= c("Inmate.Number","LastName","FirstName","DOB","Location.Permanent","Current.Loc.CD","Offence","Parole.Violator","Parole.Status.Code","Min.Date","Max.Date"))
colnames(OldDOCRoster)<-c("Inmate.Number","LastName","FirstName","DOB","Location.Permanent","Current.Loc.CD","Offence","Parole.Violator","Parole.Status.Code","Min.Date","Max.Date")
rm(NewFile,OldRoster)

NewDOCRoster[ ,c('DOB', 'Min.Date', 'Max.Date')] <- lapply(NewDOCRoster[ ,c('DOB', 'Min.Date', 'Max.Date')],as.Date, origin='1899-12-30')
OldDOCRoster[ ,c('DOB', 'Min.Date', 'Max.Date')] <- lapply(OldDOCRoster[ ,c('DOB', 'Min.Date', 'Max.Date')],as.Date, origin='1899-12-30')

############################ Mergeing #########################
SameClientsDOC <- merge.data.frame(OldDOCRoster,NewDOCRoster)
SameClientsDOC<- subset(SameClientsDOC, select=c("Inmate.Number","LastName","FirstName","DOB","Location.Permanent","Current.Loc.CD","Offence","Parole.Violator","Parole.Status.Code","Min.Date","Max.Date"))

NClientsDOC<-subset(NewDOCRoster,!(NewDOCRoster$Inmate.Number%in%OldDOCRoster$Inmate.Number))

RClientsDOC<- subset(OldDOCRoster,!(OldDOCRoster$Inmate.Number%in%NewDOCRoster$Inmate.Number))
RClientsDOC<-RClientsDOC[order(RClientsDOC$Max.Date),]

CCSRoster[,5]<- as.Date(CCSRoster[,5])
CCSRoster <-subset(SameClientsDOC,!(SameClientsDOC$Inmate.Number%in%CCSRoster$Inmate.Number))

TRoster<-merge.data.frame(NClientsDOC,SameClientsDOC,all = TRUE)
TRoster<-TRoster[order(TRoster$Max.Date),]
########### FORMATTING AND SUCH

TRoster$DOB<-format(as.Date(TRoster$DOB),format="%m/%d/%Y")
TRoster$Min.Date<-format(as.Date(TRoster$Min.Date),format="%m/%d/%Y")
TRoster$Max.Date<-format(as.Date(TRoster$Max.Date),format="%m/%d/%Y")

NClientsDOC$DOB<-format(as.Date(NClientsDOC$DOB),format="%m/%d/%Y")
NClientsDOC$Min.Date<-format(as.Date(NClientsDOC$Min.Date),format="%m/%d/%Y")
NClientsDOC$Max.Date<-format(as.Date(NClientsDOC$Max.Date),format="%m/%d/%Y")

RClientsDOC$DOB<-format(as.Date(RClientsDOC$DOB),format="%m/%d/%Y")
RClientsDOC$Min.Date<-format(as.Date(RClientsDOC$Min.Date),format="%m/%d/%Y")
RClientsDOC$Max.Date<-format(as.Date(RClientsDOC$Max.Date),format="%m/%d/%Y")

CCSRoster$DOB<-format(as.Date(CCSRoster$DOB),format="%m/%d/%Y")
CCSRoster$Min.Date<-format(as.Date(CCSRoster$Min.Date),format="%m/%d/%Y")
CCSRoster$Max.Date<-format(as.Date(CCSRoster$Max.Date),format="%m/%d/%Y")

#Add workbook and worksheets
wb<- createWorkbook("DOCrosters")
addWorksheet(wb,"Roster",gridLines = FALSE)
addWorksheet(wb,"New",gridLines = FALSE)
addWorksheet(wb,"Removed",gridLines = FALSE)
addWorksheet(wb,"Not in CCS",gridLines = FALSE)
#Write data to worksheets
writeData(wb, sheet = 1, TRoster, rowNames = FALSE)
writeData(wb, sheet = 2, NClientsDOC, rowNames = FALSE)
writeData(wb, sheet = 3, RClientsDOC, rowNames = FALSE)
writeData(wb, sheet = 4, CCSRoster, rowNames = FALSE)

#Create formatting styles
allformat<-createStyle(halign = 'left',
                       borderColour = getOption("openxlsx.borderColour", "black"),
                       borderStyle = getOption("openxlsx.borderStyle", "thin"),
                       border = 'TopBottomLeftRight'
)
centerformat<-createStyle(halign = 'center')
headerformat<-createStyle(textDecoration ='Bold',fgFill ='#88c2a0')
#Apply formatting styles
addStyle(wb, sheet = 1, allformat, rows = 1:400, cols = 1:11, gridExpand = TRUE)
addStyle(wb, sheet = 2, allformat, rows = 1:400, cols = 1:11, gridExpand = TRUE)
addStyle(wb, sheet = 3, allformat, rows = 1:400, cols = 1:11, gridExpand = TRUE)
addStyle(wb, sheet = 4, allformat, rows = 1:400, cols = 1:11, gridExpand = TRUE)


addStyle(wb, sheet = 1, headerformat, rows = 1, cols = c(1:11), gridExpand = TRUE, stack = TRUE)
addStyle(wb, sheet = 2, headerformat, rows = 1, cols = c(1:11), gridExpand = TRUE, stack = TRUE)
addStyle(wb, sheet = 3, headerformat, rows = 1, cols = c(1:11), gridExpand = TRUE, stack = TRUE)
addStyle(wb, sheet = 4, headerformat, rows = 1, cols = c(1:11), gridExpand = TRUE, stack = TRUE)

addStyle(wb, sheet = 1, centerformat, rows = 1:5000, cols = c(4:6,8:11), gridExpand = TRUE, stack = TRUE)
addStyle(wb, sheet = 2, centerformat, rows = 1:5000, cols = c(4:6,8:11), gridExpand = TRUE, stack = TRUE)
addStyle(wb, sheet = 3, centerformat, rows = 1:5000, cols = c(4:6,8:11), gridExpand = TRUE, stack = TRUE)
addStyle(wb, sheet = 4, centerformat, rows = 1:5000, cols = c(4:6,8:11), gridExpand = TRUE, stack = TRUE)

setColWidths(wb,sheet =1, col=1:11,widths = 'auto')
setColWidths(wb,sheet =2, col=1:11,widths = 'auto')
setColWidths(wb,sheet =3, col=1:11,widths = 'auto')
setColWidths(wb,sheet =4, col=1:11,widths = 'auto')

FileName<-paste0("DOC MH Roster Philadelphia Unified",format(Sys.Date(),"%Y-%b"),".xlsx")

Path<-paste0("C:/Users/cmccrea/Box/BHJD/DOC/Rosters/",FileName)
saveWorkbook(wb, Path, overwrite = TRUE)
