require("devtools")
install.packages("readxl")
library("readxl")
library(dplyr)
library(tidyr)
library(data.table)
# this was helpful
#http://stackoverflow.com/questions/12945687/how-to-read-all-worksheets-in-an-excel-workbook-into-an-r-list-with-data-frame-e 

rigFileName <- "USRigs.xlsx"
testSheet <- read_excel(rigFileName)
str(testSheet)
class(testSheet)
head(testSheet)

sheets <- readxl::excel_sheets(rigFileName)
sheets

#Function to read all the sheets and return them in a list of data tables
read_excel_allsheets <- function(filename){
  sheets <- readxl::excel_sheets(filename)
  x <- lapply(sheets, function(x) readxl::read_excel(filename, sheet = x))
  names(x) <- sheets
  x
}

AllRigs <- read_excel_allsheets(rigFileName)
str(AllRigs)

#Convert all to data.tables
AllRigs_DT <- lapply(AllRigs, function(x){
  as.data.table(x)
 
})
AllRigs_DT

#Add Cbind all the columns (!? willthis work?)
reduced_DT <- Reduce(cbind, AllRigs_DT)
str(reduced_DT)
# so cbind won't work - it doesn't match on the rig names


#First set all the rigs to hhave the location as their key
lapply(AllRigs_DT, function(x){
  setkey(x, Location)
})

#First remove the NA  and blank "" columns - they are probably screwing things up
AllRigs_DT <- lapply(AllRigs_DT, function(x){
  x[,!names(x) %in% c(NA, ""), with = FALSE]
})

#Remove the columns with "Avg" in them
Clean <- lapply(AllRigs_DT, function(x){
  x[,!grepl("Avg", names(x)), with = FALSE]
})

#Merge the data.tables to be the 
Clean[[1]][Clean[[2]], all = TRUE]
myMerge <- function(x,y){
  merge(x, y, all = TRUE)
}
myMerge2 <- function(x,y){
  dirty <- y[x, all = TRUE]
  cleaned <- dirty[Location != "NA",] 
  return(cleaned)
}

#Test
  Reduce(myMerge2, Clean[1:15])
final <- Reduce(myMerge2, Clean)


#Change all the names to be character dates
Long <- gather(final, Date, Count, -Location)

#Convert Date from Factor to a Date
Long <- Long %>%
  mutate(Date = as.Date(as.numeric(as.character(Date)), origin = "1899-12-30"))
Long
write.table(unique(Long$Location), row.names = FALSE, file = "clipboard", sep = "\t")

#Save the State Lookup Table
#Lookup <- read.table("clipboard", sep = "\t", header = TRUE)
#Lookup
#saveRDS(Lookup, file = "Lookup.rds")
Lookup <- readRDS("Lookup.rds")
Lookup <- as.data.table(Lookup)
setkey(Lookup, Location)
Long <- as.data.table(Long)
setkey(Long, Location)
Long <- Lookup[Long,]
Long <- Long[Country != "None"]

class(Long)
Long

