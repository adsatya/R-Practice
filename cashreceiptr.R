## Libraries

library(dplyr)
library(xlsx)
library(tidyverse)
library(gizmo)
library(purrrlyr)
library(RODBCext)
std_library()


## Queries

visiondata <- file.path("query") %>%
  read_queries() %>%
  run_queries("Vision")

## Function

`%notin%` = function(x,y) !(x %in% y)

##Adjustment Out

adj_out <- visiondata$ledgerarquery %>%
  mutate(Amount = Amount * -1) %>%
  select(PostSeq , Amount , WBS1 , WBS2) %>%
  filter(Amount <= 0) %>%
  group_by(PostSeq ) %>%
  summarise(Amount = sum(Amount))


## Adjustmented Applied

adj_apl <- visiondata$ledgerarquery %>%
  mutate(Amount = Amount * -1) %>%
  select(WBS1 , WBS2 , PostSeq , Amount) %>%
  filter(Amount >= 0) %>%
  group_by( PostSeq) %>%
  summarize(Amount = sum(Amount))


## Final Adjustment Table

Adj_table <- adj_out %>%
  left_join(adj_apl, by = c("PostSeq")) %>%
  mutate(Adjustment = round(Amount.x + Amount.y)) %>%
  rename(PostSeq1 = PostSeq)


## Cash Receipt Data

cashreportnew <- visiondata$ledgerarquery %>%
  mutate(Amount = Amount * - 1) %>%
  left_join(visiondata$projectsquery %>%
              rename(PhaseName = Name) %>%
              mutate(WBS2 = as.integer(WBS2)),
            by = c("WBS1", "WBS2")) %>%
  left_join(visiondata$projectnames,
            by = c("WBS1" = "Project")) %>%
  left_join(visiondata$clientquery %>%
              rename(BillingClient = Name),
            by = c("BillingClientID" = "ClientID")) %>%
  left_join(visiondata$clientquery %>%
              rename(PrimaryClient = Name),
            by = c("ClientID")) %>%
  left_join(Adj_table, by = c("PostSeq" = "PostSeq1")) %>%
  filter(Adjustment != 0 | is.na(Adjustment) , Org != "PEA:FMW:00000" ,
         str_sub(WBS1,1,1) %notin% c("E","T")) %>%
  select( Org , WBS1 , ProjectName , WBS2 , PhaseName , Account , TransDate , Desc2 , Invoice , Amount,
          BillingClient , PrimaryClient, PostSeq ) %>%
  rename(ProjectNumber = WBS1 , PhaseNumber = WBS2 , DateEntered = TransDate ,
         Description = Desc2) %>%
  mutate(Invoice = ifelse( Account == 190.00, "Retainage",
                           ifelse( Account == 270.00, "Retainer" , Invoice)))


### 

projection <- cashreportnew %>% 
  group_by( Org, DateEntered ) %>% 
  summarise ( Amount = sum(Amount))

## Extract Report


write.xlsx(as.data.frame(cashreportnew),
           file.path("report",str_c("cashreceipt2", today(), ".xlsx")),
           sheetName = "receipt_data", row.names = FALSE)

