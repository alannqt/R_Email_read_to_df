#clear all memory
rm(list=ls())

#set the libraries needed
library(RDCOMClient)
library(stringr)
library(XML)

#if you do not have RDCOMClient installed, uncomment below to install devtools and from github
#devtools::install_github("alannqt/RDCOMClient")

#set working directory
setwd("C:/Users/alann/Documents/GitHub/R_Email_read_to_df/")
attachment_path <- "C:/Users/alann/Documents/GitHub/R_Email_read_to_df/Attachment"

#initialise the app
outApp <- COMCreate("Outlook.Application")
outlooknamespace <- outApp$GetNameSpace("MAPI")

#set the foldername to use. The folder name will follow the folder name in your outlook email
sourcefolder <- "Dev_1"
outputfolder <- "Dev_2"

#++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# This portion is important as the function will try to point to the right address in your outlook mail to refer to that folder of interest, if you do not know how it works. 
# change nothing here.
#set email poointer function
setEmail <- function(foldersrc){
  assigned <- FALSE
  i <- 1
  while (assigned == FALSE){
    readsourcecheck <- tryCatch({readsource<- outlooknamespace$Folders(i)$Folders(foldersrc)},error = function(e){return(0)})
    if (is.numeric(readsourcecheck)){
      i <- i+1
      next
    }
    assigned <- TRUE
    return(outlooknamespace$Folders(i)$Folders(foldersrc))
  }
}

#assigned the right pointer to email
emails <- setEmail(sourcefolder)$Items
processedfolder <- setEmail(outputfolder)

#sample of reading the content of email 1 (Put an email with attachment to see the output)
#to read content of all email in a folder use a while loop, the example shown is using a single email
# while (emails()$Count()>0){ }
emailsubject <- emails(1)$Subject()
emailbody <- emails(1)$Body() 
emailattachment <- emails(1)$Attachments()

if (emailattachment$Count() > 0){
  for (j in c(1:emailattachment$Count())){
    emails(1)$Attachments(j)$SaveAsFile(paste(attachment_path,emailattachment$Item(j)[['DisplayName']],sep ="/"))
  }
}

#URL pattern
url_pattern <- "http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+"
x <- str_match_all(emailbody, url_pattern)

#Read tables
#declare empty df
df <- data.frame(Name = character(),Code = character(),Price = character(), Volume = character())

#locate the start and stop of the table
tabledata <- str_locate_all(emailbody,"\r\n\r\n")
data <- substr(emailbody,tabledata[[1]][1],tabledata[[1]][2])
data <- gsub("<(.*?)>", "", data, perl=TRUE)
data_test <- unlist(str_extract_all(data,"\\(?[A-Za-z0-9),.]+\\)?"))

data_test <- strsplit(unlist(data_test)," ")

y <- [as.character(data_test[5]),as.character(data_test[6]),as.character(data_test[7]),as.character(data_test[8])]
y <- data_test[5:8]
df <-rbind(df,y)

write.csv(x,"output.csv")
write.csv(df,"df_output.csv")

#moving email from source folder to Test folder
emails(1)$Move(processedfolder)
