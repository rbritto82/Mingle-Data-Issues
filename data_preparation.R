library(openxlsx)
library(tidyverse)
library(lubridate)
library(stringr)


#Filepath, use pp snipet to help with pasting the file paths-----------------------------------------------------------------------------------------------
pathBUC <- "C:\\Users\\eriobro\\Documents\\Dashboard\\Data\\Mingle\\BUC\\BUC_Cards.csv"
pathBSUCRMCA <- "C:\\Users\\eriobro\\Documents\\Dashboard\\Data\\Mingle\\RMCA\\RMCA_BSUC_Cards.csv"
pathBSUCCHA <- "C:\\Users\\eriobro\\Documents\\Dashboard\\Data\\Mingle\\CHA\\CHA_BSUC_Cards.csv"
pathBSUCCPM <- "C:\\Users\\eriobro\\Documents\\Dashboard\\Data\\Mingle\\CPM\\CPM_BSUC_Cards.csv"
pathBSUCCIL <- "C:\\Users\\eriobro\\Documents\\Dashboard\\Data\\Mingle\\CIL\\CIL_BSUC_Cards.csv"
pathUSRMCA <- "C:\\Users\\eriobro\\Documents\\Dashboard\\Data\\Mingle\\RMCA\\RMCA_US_Cards.csv"
pathUSCHA <- "C:\\Users\\eriobro\\Documents\\Dashboard\\Data\\Mingle\\CHA\\CHA_US_Cards.csv"
pathUSCPM <- "C:\\Users\\eriobro\\Documents\\Dashboard\\Data\\Mingle\\CPM\\CPM_US_Cards.csv"
pathUSCIL <- "C:\\Users\\eriobro\\Documents\\Dashboard\\Data\\Mingle\\CIL\\CIL_US_Cards.csv"

#Reading file and cleaning data--------------------------------------------------------------------------------------------------------------------------------
rawBUC <- as_tibble(read.csv2(pathBUC, stringsAsFactors = FALSE))
rawBSUCRMCA <- as_tibble(read.csv2(pathBSUCRMCA, stringsAsFactors = FALSE))
rawBSUCCHA <- as_tibble(read.csv2(pathBSUCCHA, stringsAsFactors = FALSE))
rawBSUCCPM <- as_tibble(read.csv2(pathBSUCCPM, stringsAsFactors = FALSE))
rawBSUCCIL <- as_tibble(read.csv2(pathBSUCCIL, stringsAsFactors = FALSE))
rawUSRMCA <- as_tibble(read.csv2(pathUSRMCA, stringsAsFactors = FALSE))
rawUSCHA <- as_tibble(read.csv2(pathUSCHA, stringsAsFactors = FALSE))
rawUSCPM <- as_tibble(read.csv2(pathUSCPM, stringsAsFactors = FALSE))
rawUSCIL <- as_tibble(read.csv2(pathUSCIL, stringsAsFactors = FALSE))

#combining the data from all tpgs in only two data frames
rawBSUC <- rbind(rawBSUCRMCA, rawBSUCCHA, rawBSUCCPM, rawBSUCCIL)
rawUS <- rbind(rawUSRMCA, rawUSCHA, rawUSCPM, rawUSCIL)

#removing original dataframes
rm(rawBSUCRMCA, rawBSUCCHA, rawBSUCCPM, rawBSUCCIL, rawUSRMCA, rawUSCHA, rawUSCPM, rawUSCIL)
rm(pathBUC, pathBSUCRMCA, pathBSUCCHA, pathBSUCCPM, pathBSUCCIL, pathUSRMCA, pathUSCHA, pathUSCPM, pathUSCIL)

#dropping unused first column
rawBUC$query <- NULL
rawBSUC$query <- NULL
rawUS$query <- NULL

#fixing the BSUC name for all use cases
rawUS$business_sub_use_case <-  sub("^#.*?\\s","",rawUS$business_sub_use_case)

#removing empty BUCs
rawBUC <- rawBUC %>% filter(buc_id != "")
rawBSUC <- rawBSUC %>% filter(buc_id != "")

#Adding number of teams per BSUC-------------------------------------------------------------------------------------------------------------------------------
j <- nrow(rawBSUC)
numTeamsTPG <- numeric(j)

for(i in 1:j){ #identifying related use cases and associated teams
  indexes <- grep(rawBSUC$name[i], rawUS$business_sub_use_case)
  localUS <- rawUS$assigned_team[indexes]
  localUS <- localUS[!duplicated(localUS)]
  numTeamsTPG[i] <- length(localUS)
}

numTeamsTPG[numTeamsTPG == 0] <- 1 #for BSUCS with no US
rawBSUC <- rawBSUC %>% mutate(number_TPG_teams = numTeamsTPG) #adding number of teams per BSUC

#Adding number of involved TPGs----------------------------------------------------------------------------------------------------------------------------------
numTPGs <- numeric(j)
numTotalTeams <- numeric(j)

for(i in 1:j){ #identifying related BUCs and associated OCS TPGs
  indexes <- grep(rawBSUC$buc_id[i], rawBSUC$buc_id)
  localProduct <- rawBSUC$product[indexes]
  localProduct <- localProduct[!duplicated(localProduct)]
  numTPGs[i] <- length(localProduct)
  numTotalTeams[i] <- sum(rawBSUC$number_TPG_teams[indexes])
}

rawBSUC <- rawBSUC %>% mutate(number_TPGs = numTPGs, number_Total_Teams = numTotalTeams) #adding number of OCS TPGs collaborating, including the BSUC responsible

#Adding size for each BSUC. Convention is S <= 300h < M <= 600h < L <= 900 < XL------------------------------------------------------------------------------------
numericSize <- c(300, 600, 900)
categoricalSize <- c("S", "M", "L", "XL", "Undefined") #zero and empty are set to Undefined
BSUCSize <- character(j)

rawBSUC$estimated_effort[is.na(rawBSUC$estimated_effort)] <- 0 #setting empty rows to zero
rawBSUC$estimated_effort[rawBSUC$product == "RMCA"] <- rawBSUC$estimated_effort[rawBSUC$product == "RMCA"] * 300 #transforming RMCA effort from sprints to hours

BSUCSize[rawBSUC$estimated_effort <= numericSize[1]] <- categoricalSize[1] #S
BSUCSize[rawBSUC$estimated_effort > numericSize[1] & rawBSUC$estimated_effort <= numericSize[2]] <- categoricalSize[2] #M
BSUCSize[rawBSUC$estimated_effort > numericSize[2] & rawBSUC$estimated_effort <= numericSize[3]] <- categoricalSize[3] #L
BSUCSize[rawBSUC$estimated_effort > numericSize[3]] <- categoricalSize[4] #XL
BSUCSize[rawBSUC$estimated_effort == 0] <- categoricalSize[5] #Undefined

BSUCSize <- factor(BSUCSize, levels = categoricalSize) #transforming in factor

rawBSUC <- rawBSUC %>% mutate(size = BSUCSize) #adding size to BSUCS

#Adding lead-time---------------------------------------------------------------------------------------------------------------
#Just implementation lead-time accounted for
#not sure if these are the best dates to consider. However, the others look inconsistent 

leadTime <- ymd(rawBSUC$moved_to_context_verified_on) - ymd(rawBSUC$moved_to_ongoing_on) 
rawBSUC <- rawBSUC %>% mutate(lead_time = leadTime)

#The following BSUCS have issues with dates
# [1] "HV85240 SDDIT CHA D1720: CHA Internal Server Error for reading Refill Profiles through COBA" (No product for this one either)

# [2] "CPM_FT25-1: Threshold Triggered Promotion for Accumulated Service Usage, Activation of Bundled Promotion Product - CPM impact"
# [3] "CPM_FT25-4: Threshold Triggered Promotion Generates Promotion Notification"                                                   
# [4] "CPM_4B-10 One-time Charge for activation at first traffic event _Testdata"                                                    
# [5] "CPM_FT25-1:2 Product Status Notification Event enhancement with Unique Id"
# rawBSUC <- rawBSUC[rawBSUC$product != "",]

#Not possible to collect full lead-time yet, since many BSUCS don't have an associate BUC
# studyOnGoing <- character(j) 
# studyDone <- character(j)
# specStarted <- character(j)
# specDone <- character(j)
# 
# for(i in 1:j){ #identifying related BUCs and associated OCS TPGs
#   pattern <- paste("^", rawBSUC$buc_id[i], "$", sep = "") #pattern to grep the dates from BUC data
#   index <- grep(pattern, rawBUC$buc_id) #index of the associated BUC
#   print(i)
#   print(index)
#   studyOnGoing[i] <- rawBUC$moved_to_study_ongoing_on[index]
#   studyDone[i] <- rawBUC$moved_to_study_done_on[index]
#   specStarted[i] <- rawBUC$moved_to_specification_started_on[index]
#   specDone[i] <- rawBUC$moved_to_specification_done_on[index]
# }



#Creating workbook-----------------------------------------------------------------------------------------------------------
workbookName <- paste("C:\\Users\\eriobro\\Documents\\Dashboard\\Data\\Prepared\\BSUC_", Sys.Date(), ".xlsx", sep = "")
wb <- createWorkbook(workbookName)
addWorksheet(wb, "Prepared Data")


#Writing and saving the prepared data
writeData(wb, "Prepared Data", rawBSUC, withFilter = TRUE)
saveWorkbook(wb, file = workbookName, overwrite = TRUE)
