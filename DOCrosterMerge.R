####################### Load Library and SQL Pull #########
library(odbc)
library(reshape2)
library(dplyr)
library(openxlsx)
library(writexl)
library(lubridate)
library(tidyr)


OldRoster<-"DOC MH Roster Philadelphia 2021-Apr21 Unified.xlsx"
NewFile<-"Copy of DOC MH Roster Philadelphia County 4.1.2021.xlsx"


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


sheets <- list("Roster" = TRoster, "New" = NClientsDOC,"Removed" = RClientsDOC,"Not in CCS" = CCSRoster) #assume sheet1 and sheet2 are data frames
filepath<-("C:/Users/cmccrea/Box/BHJD/DOC/Rosters/")
filename<-paste0(filepath,"DOC MH Roster Philadelphia ",format(Sys.Date(),"%Y-%b%d")," Unified.xlsx")

write.xlsx(sheets,filename,colWidths="auto")
