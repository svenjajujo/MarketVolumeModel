#Install packages if necessary
#install.packages('tidyverse')
#install.packages('dplyr')
#install.packages('insight')
#install.packages("xlsx")

#Change below with your project directory, keep only one uncommented setwd() with your correct project directory
#setwd("/home/developer/projects/MarketVolumeModel/MarketVolumeModel")
setwd("/Users/sjungjohann/OneDrive - Global Alliance for Improved Nutrition/Data for analysis/R syntax/")

#Might need to install javax64 and redefine JAVA_HOME if error
# Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jre-1.8')

#Import libraries (Install packages if not installed already)
library(tidyverse)
library(xlsx)
library(readxl)
library(dplyr)
library(insight)
library(arsenal)

#The following command extends the number of lines of printing your results in the console
options(scipen = 999)

#Clear console
cat("\014")  

#Suggestion : set Rstudio console line limit to 10000

#legend for color code
print_color(('---------------------------'),"violet")
print_color(('Legend for the color code'),"br_violet")
print_color(('---------------------------'),"violet")

#For new line 
cat('\n')
cat('\n')

#Informative message instruction, different color for different error and info messages etc.
print_color(c("Start work on a step"),"green")
cat('\n')
print_color(c("Completed work on a step"),"bold")
cat('\n')
print_color(c('Information to the user, no action required') ,"cyan")
cat('\n')
print_color(c('Error, required user correction in the source data and the R code must be rerun') ,"red")
cat('\n')
print_color(c('Need user input in the R console to continue, no need to rerun the R code'),"yellow")
cat('\n')
print_color(c('Data was missing, assumption was applied to deduce it'),"blue")
cat('\n')
cat('\n')
print_color(('-------------------------------------------------------------------------------'),"violet")
cat('\n')
cat('\n')

print_color(c("Starting work"),"green")
cat("\n")
cat("\n")


#It will prompt user to select if wants to use the R console or a text output file ('console.txt')
#Note:- if typed 'text' then file 'console.txt' will be stored in your current project directoy with all result executions and messages
print_color(c('To have the output in a text file type "text", otherwise just press the "Return" key ') ,"yellow")
Console_or_text = readline()
if (Console_or_text == "text"){
  sink('console.txt')
}

# Import excel data from excel sheet 'Données_Source.xlsx' and store it in dedicated dataframes(DF) and variables in R environment
# for loop to go through all worksheet of the 'Données_Source.xlsx' excel
for (i in excel_sheets("Données_Source.xlsx")) {
  Temporary_Object <- capture.output(cat("Worksheet_",i, sep = ""))
  #Create object with worksheet
  #Create variable for all worksheet in the R environment and store worksheet name in that variable name
  #Suppose worksheet name is 'ConsumptionFood' then will create variable with name 'Worksheet_ConsumptionFood' in R environment and store name 'ConsumptionFood' in to it 
  assign(Temporary_Object,i)
  #Create DF for each worksheet with worksheet name in R environment and store data in to it
  #list of worksheet with first line as header
  worksheet_no1stline <- c("Listes" , "CountryPop")
  if(!(i %in% worksheet_no1stline)){
    assign(i, read_xlsx("Données_Source.xlsx",
                        sheet = i,
                        skip = 1))
    TemporaryDF <- get(i)
    TemporaryDF <-mutate(TemporaryDF, 'Assumption applied' = NA)
    assign(i,TemporaryDF) 
    rm(TemporaryDF)
  } else {
    assign(i, read_xlsx("Données_Source.xlsx",
                        sheet = i))
  }
  
}
cat("\n")
cat("\n")




#For loop to go through each Worksheet of the excel 'Données_Source.xlsx' and check for mandatory columns and data
print_color(('---------------------------'),"violet")
print_color(('CHECKING FOR MANDATORY DATA'),"br_violet")
print_color(('---------------------------'),"violet")
cat('\n')
cat('\n')
cat('\n')

Donnée_obligatoire_manquante <- 0  #Initialized flag with zero if found missing column and data during checking then flag will update to 1 and at the end of file flag will be cross validated
for (i in excel_sheets("Données_Source.xlsx")) {
  # Code to be executed if the pattern "Supply" is found in the string 'i'
  # It will only look for the exact string "Supply" (case-sensitive) due to fixed = TRUE
  if(grepl("Supply",i,fixed = TRUE)){
    
    print_color(c("Checking for missing mandatory Data in",i),"green")
    cat("\n")
    TemporaryDF <- get(i)
    #in Supply : mandatory = Country, Domestic Supply Category, Food Staple, Reference for analysis
    #Check for mandatory above defined columns existence in each worksheet containing 'Supply' in their name
    
    if (!("Country" %in% colnames(TemporaryDF))){
      print_color(c('Mandatory Column "Country" was not found, please check that the column name as not been changed> Correct it') ,"red")
      Donnée_obligatoire_manquante <- 1
      cat("\n")
    }
    if (!("Domestic Supply Category" %in% colnames(TemporaryDF))){
      print_color(c('Mandatory Column "Domestic Supply Category" was not found, please check that the column name as not been changed> Correct it') ,"red")
      Donnée_obligatoire_manquante <- 1
      cat("\n")
    }
    if (!("Food Staple" %in% colnames(TemporaryDF))){
      print_color(c('Mandatory Column "Food Staple" was not found, please check that the column name as not been changed> Correct it') ,"red")
      Donnée_obligatoire_manquante <- 1
      cat("\n")
    }
    if (!("Reference for analysis" %in% colnames(TemporaryDF))){
      print_color(c('Mandatory Column "Reference for analysis" was not found, please check that the column name as not been changed> Correct it') ,"red")
      Donnée_obligatoire_manquante <- 1
      cat("\n")
    }

    #Check for mandatory Data
    #in Supply : mandatory data columns = Country, Domestic Supply Category, Food Staple, Reference for analysis
    # Check for 'NA' value in above each column 
    if (anyNA(TemporaryDF$Country)) {
      print_color(c('Mandatory Data missing in "Country" column of', i ,'> Correct it') ,"red")
      Donnée_obligatoire_manquante <- 1
      cat("\n")
    }
    if (anyNA(TemporaryDF$"Domestic Supply Category")) {
      print_color(c('Mandatory Data missing in "Domestic Supply Category" column of', i ,'> Correct it') ,"red")
      Donnée_obligatoire_manquante <- 1
      cat("\n")
    }
    if (anyNA(TemporaryDF$'Food Staple')) {
      print_color(c('Mandatory Data missing in "Food Staple" column of', i ,'> Correct it') ,"red")
      Donnée_obligatoire_manquante <- 1
      cat("\n")
    }
    if (anyNA(TemporaryDF$'Reference for analysis')) {
      print_color(c('Mandatory Data missing in "Reference for analysis" column of', i ,'> Automaticaly replaced by 1') ,"cyan")
      TemporaryDF$'Reference for analysis'[is.na(TemporaryDF$'Reference for analysis')] = 1
      cat("\n")
      
    }
    print_color(c("Finished checking for missing mandatory Data in",i),"bold")
    cat("\n")
    cat("\n")
    #Storing temporary DF in original DF
    assign(i,TemporaryDF)
    rm(TemporaryDF) #remove temporary DF
  } else if(grepl("Consumption",i,fixed = TRUE)){
    # Code to be executed if the pattern "Consumption" is found in the string 'i'
    # It will only look for the exact string "Consumption" (case-sensitive) due to fixed = TRUE
    print_color(c("Checking for missing mandatory Data in",i),"green")
    cat("\n")
    TemporaryDF <- get(i)
    
    #in Consumption : mandatory = Country, Consumption/Use category, Food Staple, Reference for analysis
    #check for mandatory above columns existence
    
    if (!("Country" %in% colnames(TemporaryDF))){
      print_color(c('Mandatory Column "Country" was not found, please check that the column name as not been changed> Correct it') ,"red")
      Donnée_obligatoire_manquante <- 1
      cat("\n")
    }
    if (!("Consumption/Use category" %in% colnames(TemporaryDF))){
      print_color(c('Mandatory Column "Consumption/Use category" was not found, please check that the column name as not been changed> Correct it') ,"red")
      Donnée_obligatoire_manquante <- 1
      cat("\n")
    }
    if (!("Food Staple" %in% colnames(TemporaryDF))){
      print_color(c('Mandatory Column "Food Staple" was not found, please check that the column name as not been changed> Correct it') ,"red")
      Donnée_obligatoire_manquante <- 1
      cat("\n")
    }
    if (!("Reference for analysis" %in% colnames(TemporaryDF))){
      print_color(c('Mandatory Column "Reference for analysis" was not found, please check that the column name as not been changed> Correct it') ,"red")
      Donnée_obligatoire_manquante <- 1
      cat("\n")
    }
    
    #Check for mandatory column Data
    #in Consumption : mandatory data column = Country, Consumption/Use category, Food Staple, Reference for analysis
    # Check for 'NA' value in above each column
    if (anyNA(TemporaryDF$Country)) {
      print_color(c('Mandatory Data missing in "Country" column of', i ,'> Correct it') ,"red")
      Donnée_obligatoire_manquante <- 1
      cat("\n")
    }
    if (anyNA(TemporaryDF$"Consumption/Use category")) {
      print_color(c('Mandatory Data missing in "Consumption/Use category" column of', i ,'> Correct it') ,"red")
      Donnée_obligatoire_manquante <- 1
      cat("\n")
    }
    if (anyNA(TemporaryDF$'Food Staple')) {
      print_color(c('Mandatory Data missing in "Food Staple" column of', i ,'> Correct it') ,"red")
      Donnée_obligatoire_manquante <- 1
      cat("\n")
    }
    if (anyNA(TemporaryDF$'Reference for analysis')) {
      print_color(c('Mandatory Data missing in "Reference for analysis" column of', i ,'> Automaticaly replaced by 1') ,"cyan")
      TemporaryDF$'Reference for analysis'[is.na(TemporaryDF$'Reference for analysis')] = 1
      cat("\n")
      
    }
    print_color(c("Finished checking for missing mandatory Data in",i),"bold")
    cat("\n")
    cat("\n")
    #Storing temporary DF in original DF
    assign(i,TemporaryDF)
    rm(TemporaryDF) #remove temporary DF
  } 
}

cat("\n")
cat("\n")

#for loop to go through each Worksheet and check for data limited by a list (data should be from predefined list for particular columns, for validation using Lists DF)
print_color(('---------------------------'),"violet")
print_color(('CHECKING FOR LIST LIMITED DATA'),"br_violet")
print_color(('---------------------------'),"violet")
cat('\n')
cat('\n')
cat('\n')
Donnée_liste_manquante <- 0  #Initialized flag with zero if found missing limited column and limited column data with predefined during checking then flag will update to 1 and at the end of file flag will be cross validated
for (i in excel_sheets("Données_Source.xlsx")) {
  TemporaryDF <- get(i)
  if(grepl("Supply",i,fixed = TRUE)){
    
    print_color(c("Checking for list limited Data in",i),"green")
    cat("\n")
    
    #in Supply : list limited Data = Country, Location (Region),Domestic Supply Category, Food Staple, Food Type, Packaged/Bulk, Reference for analysis
    Limited_Data_List = c('Country', 'Location (Region)','Domestic Supply Category', 'Food Staple', 'Food Type', 'Packaged/Bulk', 'Reference for analysis')
    
    for (j in Limited_Data_List){
      print_color(c('Checking for',j),"cyan")
      cat("\n")
      #check for list limited columns existence
      if (!(j %in% colnames(TemporaryDF))){
        print_color(c('List limited Column',paste0('"',j,'"'),'was not found, please check that the column name as not been changed> Correct it') ,"red")
        Donnée_liste_manquante <- 1
        cat("\n")
      }
      #check for list limited data means data should be from predefined data that is in Lists worksheet
      Counter = 3
      Temporary_List = TemporaryDF[[j]]
      Temporary_Listes_List = Listes[[j]]
      for (k in Temporary_List){
        #Check with already having 'Listes' data
        if((!toupper(k) %in% toupper(Temporary_Listes_List)) & (!is.na(k))){
          
            print_color(capture.output(cat(j,paste0('"',k,'"') ,'In',i,'at line',Counter,'does not belong to the list, correct it or update the list')),"red")
            Donnée_liste_manquante <- 1
            cat("\n")
        }
        Counter = Counter + 1
      }
    }
    print_color(c("Finished checking for list limited Data in",i),"bold")
    cat("\n")
    cat("\n")
  }
  if(grepl("Consumption",i,fixed = TRUE)){
    
    print_color(c("Checking for list limited Data in",i),"green")
    cat("\n")
    
    #in Consumption : list limited Data = Country, Location (Region),Consumption/Use category, Food Staple, Food Type, Packaged/Bulk, Reference for analysis
    Limited_Data_List = c('Country', 'Location (Region)','Consumption/Use category', 'Food Staple', 'Food Type', 'Packaged/Bulk', 'Reference for analysis')
    for (j in Limited_Data_List){
      print_color(c('Checking for',j),"cyan")
      cat("\n")
      #check for list limited columns existence
      if (!(j %in% colnames(TemporaryDF))){
        print_color(c('List limited Column',paste0('"',j,'"'),'was not found, please check that the column name as not been changed> Correct it') ,"red")
        Donnée_liste_manquante <- 1
        cat("\n")
      }
      #check for list limited data
      Counter = 3
      Temporary_List = TemporaryDF[[j]]
      Temporary_Listes_List = Listes[[j]]
      for (k in Temporary_List){
        #print(typeof(k))
        
        if (!is.double(k)){
          if((!toupper(k) %in% toupper(Temporary_Listes_List)) & (!is.na(k))){
            print_color(capture.output(cat(j,paste0('"',k,'"') ,'In',i,'at line',Counter,'does not belong to the list, correct it or update the list')),"red")
            Donnée_liste_manquante <- 1
            cat("\n")
          }
          Counter = Counter + 1
        } else {
          
        }
        
        
      }
      
    }
    print_color(c("Finished checking for list limited Data in",i),"bold")
    cat("\n")
    cat("\n")
  }
  rm(TemporaryDF)
  
 
}
cat("\n")
cat("\n")

#for loop to go through each Worksheet and check and validate for volume data
#check if volume data exist and if not, we can deduce it from consumption per capita
print_color(('---------------------------'),"violet")
print_color(('CHECKING FOR VOLUME DATA'),"br_violet")
print_color(('---------------------------'),"violet")
cat('\n')
cat('\n')
cat('\n')
Volume_manquant <- 0
for (i in excel_sheets("Données_Source.xlsx")) {
  TemporaryDF <- get(i)
  
    print_color(c("Checking for Volume Data in",i),"green")
    cat("\n")
    Counter = 3
    #Loop to check if volume data exist and if not, we can deduce it from consumption per capita
    if ('Volume Value (unit/y)' %in% colnames(TemporaryDF)){
      for (j in TemporaryDF$'Volume Value (unit/y)') {
        Found_Country = FALSE
        per_capita_available <- 0
        monetary_available <- 0
        choice <-0
        #check if the volume exist that unit and Unit quantity column are filled correctly
        if (!is.na(j)){
          if (is.na(TemporaryDF$Unit[Counter-2])){
            TemporaryDF$Unit[Counter-2] <- "MT"
            print_color(c('Data missing in "Unit" column of', i ,'> Automaticaly replaced by MT') ,"cyan")
            cat("\n")
          }
          if (is.na(TemporaryDF$'Unit quantity'[Counter-2])){
            TemporaryDF$'Unit quantity'[Counter-2] <- 1000
            print_color(c('Data missing in "Unit Quantity" column of', i ,'> Automaticaly replaced by 1000') ,"cyan")
            cat("\n")
          }
        }
        #check if volume not exist then calculate by using either consumption per capita (Consumption per Capita (kg/y)) or monetary value (Monetary Value of 1kt)
        if (is.na(j)){
          #check if either consumption per capita or monetary value data are available
          if (!is.na(TemporaryDF$'Consumption per Capita (kg/y)'[Counter-2])){
            per_capita_available <- 1
          }
          if (!is.na(TemporaryDF$'Monetary Value of 1kt'[Counter-2])){
            monetary_available <- 1
          }
          #check if both consumption per capita and monetary value are present then give choice to select any of one 
          if (per_capita_available == 1 && monetary_available == 1){
            print_color(c('line',Counter,'both consumption per capita and Monetary value are available, which one do you want to use to calculate the volume? Type "1" for consumption per capita and "2" for monetary value '),"yellow")
            choice <- readline()
            while (choice != 2 && choice != 1) {
              print_color(c('you must choose either 1 or 2 no other value is accepted'),"yellow")
              choice <- readline()
            }
          }
          #check with consumption per capita and calculate using this
          if (choice == 1 || (per_capita_available == 1 && monetary_available == 0)){
            Counter_k = 1
            #get population of the matching country from CountryPop DF and update 'Found_Country' flag to TRUE
            for (k in CountryPop$Country) {
              if (toupper(TemporaryDF$'Country'[Counter-2]) == toupper(k)) {
                Population = as.integer(CountryPop$Population[Counter_k]) 
                Found_Country = TRUE
                break
              }
              Counter_k = Counter_k+1
            }
            #if country found then go for the volume calculation
            if (Found_Country) {
              #the calculated volume is the population x consumption per capita in kg/y divided by 1 million (to obtain kt/y)
              New_Volume = Population*TemporaryDF$'Consumption per Capita (kg/y)'[Counter-2]/1000000
              TemporaryDF$'Volume Value (unit/y)'[Counter-2] = New_Volume
              
              print_color(c(paste('Line',Counter, 'Volume DATA not available, consumption per capita available, the new volume data is', New_Volume,'kt/year')),"blue")
              TemporaryDF$'Assumption applied'[Counter-2] <- "Volume deduced from consumption per capita"
              TemporaryDF$'Unit quantity'[Counter-2] <- 1000
              TemporaryDF$'Unit'[Counter-2] <- "MT"
              cat("\n")
            } else {
              print_color(c(paste('Line',Counter,'Unknown country',paste0(TemporaryDF$Country[Counter-2],','),'population is not known, volume cannot be calculated')),'red')
              cat("\n")
            }
            
            Found_Country = FALSE
          }
          #check with monetary value and calculate using this
          if (choice == 2 || (per_capita_available == 0 && monetary_available == 1)){
            New_Volume = TemporaryDF$'Monetary Value'[Counter-2]/TemporaryDF$'Monetary Value of 1kt'[Counter-2]
            TemporaryDF$'Volume Value (unit/y)'[Counter-2] = New_Volume
            print_color(c(paste('Line',Counter, 'Volume DATA not available, Monetary value available, the new volume data is', New_Volume,'kt/year')),"blue")
            TemporaryDF$'Assumption applied'[Counter-2] <- "Volume deduced from monetary value"
            TemporaryDF$'Unit quantity'[Counter-2] <- 1000
            TemporaryDF$'Unit'[Counter-2] <- "MT"
            cat("\n")
          }
           #  else {
           #  print_color(c('Line',Counter,'Volume DATA not available and no other DATA is available to deduce it'),"red")
           #  Volume_manquant <- 1
           #  cat("\n")
           # }
        }
        
        Counter = Counter +1
      } 
      print_color(c("Finished checking for Volume Data in",i),"bold")
    } else {
      print_color(c('No volume data in that worksheet, passing') ,"cyan")
    }
    #Storing temporary DF in original DF
    assign(i,TemporaryDF)
    cat("\n")
    cat("\n")
    rm(TemporaryDF) #remove temporary DF
}
cat("\n")
cat("\n")

#for loop to go through each Worksheet and perform cross validation for each above checking
cross_validation <- 1
if(Donnée_obligatoire_manquante == 1) {
  print_color(c('Mandatory Data still missing > Correct it, cross validation cannot be done') ,"red")
  cross_validation <- 0
  cat('\n')
  
  cat('\n')
}


if(Donnée_liste_manquante == 1) {
  print_color(c('List Limited Data incorrect > Correct it, cross validation cannot be done') ,"red")
  cross_validation <- 0
  cat('\n')
  cat('\n')
}

# if(Volume_manquant == 1) {
#   print_color(c('Volume data is missing > Correct it, cross validation cannot be done') ,"red")
#   cross_validation <- 0
#   cat('\n')
#   cat('\n')
# }

if (cross_validation == 1) {
  print_color(('---------------------------'),"violet")
  print_color(('Running cross validation'),"br_violet")
  print_color(('---------------------------'),"violet")
  cat('\n')
  cat('\n')
  cat('\n')
  #Check for total for each food staple
  #create DF to store totals per food staple from SupplyFood
  SupplyFood_FoodStaple_Totals <- data.frame(food_staple=c(unique(SupplyFood$`Food Staple`)),Total_Value=NA)
  Counter <- 3
  print_color(c("Checking for Total volume in SupplyFood"),"green")
  cat("\n")
  for (i in SupplyFood$`Food Staple`){
    if (toupper(SupplyFood$`Domestic Supply Category`[Counter-2])=="TOTAL"){
      Counter_unique <- 1
      
      for (j in SupplyFood_FoodStaple_Totals$food_staple){
        if (j == SupplyFood$`Food Staple`[Counter-2]){
          SupplyFood_FoodStaple_Totals$Total_Value[Counter_unique] <- SupplyFood$"Volume Value (unit/y)"[Counter-2]
          print_color(c(paste('Total volume for',j,'found :', SupplyFood$"Volume Value (unit/y)"[Counter-2],"kt/y")),"cyan")
          cat("\n")
        }
        Counter_unique = Counter_unique+1
      }
    }
    Counter = Counter +1
    
  }
  #warn the user if we did not found a total for a Food Staple
  Counter_i <-1
  for (i in SupplyFood_FoodStaple_Totals$food_staple ){
    if(is.na(SupplyFood_FoodStaple_Totals$Total_Value[Counter_i])){
      print_color(c('Missing Total value for',i,'> Total volume was calculated with the data provided, the total volume for xxxx is xx kt/y') ,"blue")
    }
    Counter_i = Counter_i+1
  }
  cat('\n')
  cat('\n')
  
  print_color(c("Checking for Total volume in ConsumptionFood"),"green")
  cat("\n")
  #create DF to store totals per food staple from ConsumptionFood and populate it with totals
  ConsumptionFood_FoodStaple_Totals <- data.frame(food_staple=c(unique(ConsumptionFood$`Food Staple`)),Total_Value=0)
  temporary_df <- aggregate(ConsumptionFood$`Volume Value (unit/y)`, by=list(ConsumptionFood$`Food Staple`), FUN = sum)
  Counter <- 1
  for (i in ConsumptionFood_FoodStaple_Totals$food_staple ){
    Counter_j <-1
    for (j in temporary_df$Group.1){
      if( i == j){
        ConsumptionFood_FoodStaple_Totals$Total_Value[Counter] <- temporary_df$x[Counter_j]
        print_color(c(paste('Total volume for',j,'found :', ConsumptionFood_FoodStaple_Totals$Total_Value[Counter],"kt/y")),"cyan")
        cat('\n')
      }
      Counter_j = Counter_j+1
    }
    Counter = Counter+1
  }
  rm(temporary_df)
  
  cat("\n")
  print_color(c("Cross validation of totals between SupplyFood and ConsumptionFood"),"green")
  cat("\n")
  #compare FoodStaple betweeen supply and consumption and detect missing ones
  #Check for Foodstaple in SupplyFood to be present in ConsumptionFood
  Counter_i <- 1
  for (i in SupplyFood_FoodStaple_Totals$food_staple){
    Found_it <- 0
    Counter_j <- 1
    for (j in ConsumptionFood_FoodStaple_Totals$food_staple){
      if (toupper(i) ==toupper(j) ){
        Found_it <- 1
        if (SupplyFood_FoodStaple_Totals$Total_Value[Counter_i]== ConsumptionFood_FoodStaple_Totals$Total_Value[Counter_j] ){
          print_color(c('Total supply and consumption for',i, 'are equal'),"cyan")
          cat('\n')
        } else {
          print_color(c('The total supply and Consumption for',i, 'are not equal. Total supply =',SupplyFood_FoodStaple_Totals$Total_Value[Counter_i],' and Total consumption=',ConsumptionFood_FoodStaple_Totals$Total_Value[Counter_j],"> Correct it"),"red")
          cat('\n')
        }
      }
      Counter_j = Counter_j + 1
    }
    
    if (Found_it == 0){
      print_color(c(paste0('"',i,'"'),'present in SupplyFood is not present in ConsumptionFood > assumption to apply') ,"blue")
      cat('\n')
    }
    Counter_i = Counter_i + 1
  }
  
  #Check for Foodstaple in ConsumptionFood that are not present in SupplyFood
  for (i in ConsumptionFood_FoodStaple_Totals$food_staple){
    Found_it <- 0
    for (j in SupplyFood_FoodStaple_Totals$food_staple){
      if (toupper(i) ==toupper(j) ){
        Found_it <- 1
      }
    }
    if (Found_it == 0){
      print_color(c(paste0('"',i,'"'),'present in ConsumptionFood is not present in SupplyFood > assumption to apply') ,"blue")
      cat('\n')
    }
  }
  cat('\n')
}



#export to Excel
#check if the 'Zambia_MM.xlsx' excel file already there then delete it
if (file.exists("Zambia_MM.xlsx")){
  file.remove("Zambia_MM.xlsx")
  print_color(c('Zambia_MM.xlsx file already exist, deleting it'),"cyan")
  cat('\n')
}
#check if the 'Zambia_MM_all_in_one.xlsx' excel file already there then delete it
if (file.exists("Zambia_MM_all_in_one.xlsx")){
  file.remove("Zambia_MM_all_in_one.xlsx")
  print_color(c('Zambia_MM_all_in_one.xlsx file already exist, deleting it'),"cyan")
  cat('\n')
}


#concatenate all DF in one single DF

Final_DF <- data.frame()
for (i in excel_sheets("Données_Source.xlsx")) {
  if(grepl("Supply",i,fixed = TRUE) || grepl("Consumption",i,fixed = TRUE)){
    df <- get(i)
    df <- mutate(df, 'Source worksheet' = i)
    Final_DF <- bind_rows(Final_DF, df)
  }

}

#export final DF to excel
Final_DF[is.na(Final_DF)] <- ""
write.xlsx(Final_DF,"Zambia_MM_all_in_one.xlsx")
print_color(c('Zambia_MM_all_in_one.xlsx file created, check the folder'),"cyan")
cat('\n')

#export to excel & remove NA from the individual DF
sheet_list <- excel_sheets("Données_Source.xlsx")
for (i in 1:length(sheet_list)){
  df <- get(sheet_list[[i]])
  for (j in 1:length(df)){
    df[j] <- lapply(df[j], as.character)
  }
  df[is.na(df)] <- ""
  file_name <- "Zambia_MM.xlsx"
  write.xlsx(df,file_name, sheetName = sheet_list[[i]], append = TRUE)
}
print_color(c('Zambia_MM.xlsx file created, check the folder'),"cyan")
cat('\n')
#copying log to text output file
if (Console_or_text == "text"){
  sink()
  print_color(c('Work completed, see "console.txt" for the log details'),"bold")
  cat('\n')
}else {
  print_color(c('Work completed'),"bold")
  cat('\n')
  }




