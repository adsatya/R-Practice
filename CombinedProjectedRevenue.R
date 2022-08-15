#### Setup ####
library(pacman)
p_load(beepr, colorspace, extrafont, gizmo, knitr, lubridate,
       magrittr, officer, openxlsx, purrrlyr, RODBC, scales,
       tidyselect, tidyverse,
       readxl)
##library(magi)
#Variables
variable = list(
    template_path = file.path('data', 'November_Earnings_Summary-2021-11-08.xlsx'),
    revenue_path = file.path('data', 'proj rev combined 2021-10-25 pivot.xlsx')
)

source_files("src")
 
actearnings = read_excel(path = variable$template_path,
                     sheet = 'FMW_data',
                     range = cell_rows(c(4, NA)),
                     guess_max = 10000)

projectrevenue = read_excel(path = variable$revenue_path,
                           sheet = 'Proj rev combined 2021-10-25',
                           range = cell_rows(c(1, NA)),
                           guess_max = 10000)
#names(projectrevenue)
projectrevenue$Month = as.Date(projectrevenue$Month)
projectrevenue <- filter(projectrevenue, projectrevenue$Month == as.Date("2021-10-01"))


projorg <- actearnings[,1]
projnum <- actearnings[,2]
projnetrevenue <- actearnings[,15]
select(actearnings,"Client","Org")

dfProject <- data.frame(projorg, projnum, projnetrevenue)

library(tidyr)
dfProject = tidyr::separate(data = dfProject, col = Project, into = c("projectNumber", "ProjectName"), 10)
dfProject = tidyr::separate(data = dfProject, col = Org, into = c("projorg", "X"), 13)
dfProject$X = NULL
## dfProject = dfProject %>% separate(Org, c('projorg', 'x'))


library(RODBC)

connect <- odbcConnect("SQLConn")
SQL <- sqlQuery(connect, "SELECT VP.PracticeArea as [Practice Area], VO.Region, vP.Project as projectNumber, VP.Potentiality as [Real or Awaiting] FROM vma.project VP LEFT JOIN vma.org VO ON VP.Org  = VO.Org ORDER BY VP.Project")
dfmerger = merge(dfProject, SQL, by="projectNumber")
dfmerger <- unique(dfmerger)
write.csv(actearnings, "green_report_data.csv")

#names(dfmerger)
#grp <- c("projectNumber" ,"projorg", "ProjectName" ,"Practice Area", "Region" ,"Real or Awaiting")
dfmergerGroup <-  dfmerger %>% 
   group_by(projectNumber) %>%
    summarize(Current.Net.Revenue= sum(Current.Net.Revenue))

dfmergeCombine <- merge(x =dfmergerGroup , y = dfmerger, by = "projectNumber", all.x = TRUE, all.y = TRUE)
dfmergeCombine["Current.Net.Revenue.x"] <- NULL
names(dfmergeCombine)[4] <- "Current.Net.Revenue"

dfmerger <- distinct(dfmergeCombine)

dfProject <- data.frame(projorg, projnum, projnetrevenue)

colnames(dfmerger)[1] <- "Project"

colnames(projectrevenue)
projectrevenue <- projectrevenue[,c("Project","ProjectedRevUSD")]


CombinedProjectedRevenue <- merge(x =dfmerger , y = projectrevenue, by = "Project", all = TRUE)


colnames(CombinedProjectedRevenue) <- c("Project Number","Organization","Project Name","Actual Net Revenue"
                                        ,"Practice Area","Region","Real or Awaiting","Projected Revenue")

CombinedProjectedRevenue <- CombinedProjectedRevenue[,c("Practice Area","Region","Organization","Real or Awaiting"
                                                        ,"Project Number","Project Name","Projected Revenue","Actual Net Revenue")]
CombinedProjectedRevenue["Accounting Adjustment"] = 0.00
CombinedProjectedRevenue["Total Net Revenue (H+I)"] = CombinedProjectedRevenue["Actual Net Revenue"] + CombinedProjectedRevenue["Accounting Adjustment"]
CombinedProjectedRevenue["Variance (G-J)"] = CombinedProjectedRevenue["Projected Revenue"] - CombinedProjectedRevenue["Actual Net Revenue"]
CombinedProjectedRevenue["Comment"] = ""


library(dbplyr)

colnames(CombinedProjectedRevenue)[5] <- "projectNumber"
colnames(CombinedProjectedRevenue)[8] <- "ActualNetRevenue"


dfActNetRevGroup <-  CombinedProjectedRevenue %>% 
    group_by(projectNumber) %>%
    summarize(ActualNetRevenue= sum(ActualNetRevenue))

dfActNetRevCombine <- merge(x =dfActNetRevGroup , y = CombinedProjectedRevenue, by = "projectNumber", all.x = TRUE, all.y = TRUE)
dfActNetRevCombine$ActualNetRevenue.y <- NULL
names(dfActNetRevCombine)[2] <- "Actual Net Revenue"
names(dfActNetRevCombine)[1] <- "Project Number"



finaldf <- distinct(dfActNetRevCombine)
names(finaldf)
str(finaldf)
finaldf <-  finaldf[,c("Practice Area","Region","Organization","Real or Awaiting"
                       ,"Project Number","Project Name","Projected Revenue", "Actual Net Revenue"
                       ,"Accounting Adjustment", "Total Net Revenue (H+I)", "Variance (G-J)","Comment")]
str(finaldf)
finaldf$`Projected Revenue` <- as.numeric(format(round(finaldf$`Projected Revenue`, 2), nsmall = 2))
finaldf$`Actual Net Revenue` <- as.numeric(format(round(finaldf$`Actual Net Revenue`, 2), nsmall = 2))
finaldf$`Total Net Revenue (H+I)` <- as.numeric(format(round(finaldf$`Total Net Revenue (H+I)`, 2), nsmall = 2))
finaldf$`Variance (G-J)` <- as.numeric(format(round(finaldf$`Variance (G-J)`, 2), nsmall = 2))

finaldf$`Actual Net Revenue`[is.na(finaldf$`Actual Net Revenue`)] <- 0
finaldf$`Projected Revenue`[is.na(finaldf$`Projected Revenue`)] <- 0
finaldf$`Accounting Adjustment`[is.na(finaldf$`Accounting Adjustment` )] <- 0
finaldf$`Variance (G-J)`[is.na(finaldf$`Variance (G-J)`)] <- 0

finaldf$`Project Name` [is.na(finaldf$`Project Name` )] <- ""
#finaldf$Region [is.na(finaldf$Region )] <- ""
#finaldf$`Practice Area` [is.na(finaldf$`Practice Area` )] <- ""
finaldf$Organization [is.na(finaldf$Organization )] <- ""


write.csv(finaldf,"C:\\Users\\s.adusumilli\\Documents\\CombinedProjectedRevenue\\data\\FinalData1.csv", row.names = FALSE)




