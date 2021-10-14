#' Read in the data dictionary
#'
#' This function reads in a data dictionary that was set up with the
#' PMH DataDictionary.xlsm spreadsheet to a data frame.
#'
#' It assumes that the columns names  have not been altered and are:
#' c('VariableName', 'Description (optional)', 'Type', 'Minimum', 'Maximum', 'Levels')
#' To override specify colnames as an argument, ensuring to place variables in the above order.
#
#' As of the time of writing, the origin date in Excel is 30 December 1899. To override this specify origin="yyy-mm-dd"
#'
#' To read in only part of the excel sheet specify the desrired range (ie range="A1:F6")
#'
#' @param excelFile Character, Path and filename of the data file
#' @param dictionarySheet Character, Name of the dictionary sheet within the file, defaults to 'DataDictionary'
#' @param colnames Optional, Column names of the DataDictionary, defaults to those used in the Excel template c('VariableName', 'Description (optional)', 'Type', 'Minimum', 'Maximum', 'Levels')
#' @param range Optional, Range of Excel sheet to restrict import to (ie. range="A1:F6")
#' @param origin Optional, the date origin of Excel dates, defaults to 30 December 1899
#' @export
readDataDict <- function(excelFile,dictionarySheet ='DataDictionary',range,colnames,origin){
  if (missing(range)) range = NULL
  if (!file.exists(excelFile)){
    stop('The specified file can not be found. Check that the file exists in the specified directory.')
  }
  dict <- try(readxl::read_excel(excelFile,sheet=dictionarySheet,col_names = T,range=range,col_types = 'text'),silent = T)
  if (class(dict)[1]=='try-error') stop(paste('File access failure. \n Check that sheet',dictionarySheet,'exists in file:',excelFile,
                                              '\n\nNOTE: It may be necessary to close Excel for this function to work.'))
  if (missing(colnames)) colnames <-  c('VariableName', 'Description (optional)', 'Type', 'Minimum', 'Maximum', 'Levels')
  if (missing(origin)) origin = "1899-12-30"
  # Ensures variable ordering is correct
  dict <- dict[,colnames]
  # Rename all columns to the defaults (except Description)
  names(dict) <- c('VariableName', 'Description', 'Type', 'Minimum', 'Maximum', 'Levels')
  # remove empty rows
  dict <- dict[rowSums(is.na(dict))<length(colnames),]

  # Every variable must have a VariableName
  if (sum(is.na(dict[['VariableName']]))!=0) {
    stop (paste('Every variable must have a variable name. \nVariables:', paste(which(is.na(dict[['VariableName']])),collapse=', '),
                'are missing names. \nUse the range argument to read in only part of the Excel sheet'))
  }

  # Every variable must have a Type that is an allowed value
  types <- c('calculated','category','character','codes','date','integer','numeric')
  if (!(all(dict[['Type']] %in% types) & sum(is.na(dict[["Type"]]))==0)) {
    stop (paste('Every variable must have a valid type declared. ',
                ifelse(length(setdiff(dict[['Type']],types))==0,'',
                       paste('\nInvalid Types:',paste(setdiff(dict[['Type']],types),collapse = ','))),
                ifelse(sum(is.na(dict[["Type"]]))==0,'',
                       paste('\nMissing Types:',paste(which(is.na(dict[['Type']])),collapse=', '),
                             '\nUse the range argument to read in only part of the Excel sheet'))
    ))
  }


  # Check that dates, integers and numeric variables have min and max values set
  minCheck = dict[['VariableName']][is.na(dict[['Minimum']]) &dict[['Type']] %in% c('integer','numeric','date')]
  maxCheck = dict[['VariableName']][is.na(dict[['Maximum']]) &dict[['Type']] %in% c('integer','numeric','date')]
  if (!(length(minCheck)==0 & length(maxCheck)==0)) stop(paste('All integer, numeric and date variables must have ranges set.',
                                                               '\n\nCheck these variables:',ifelse(length(minCheck)==0,'',minCheck),ifelse(length(maxCheck)==0,'',maxCheck)))

  # Check that categories and codes can be read by importCodes
  # read the codes for variables that are categorical or numeric codes
  codeCheck = numeric(0)
  for (v in which(dict[['Type']] %in% c('category','codes'))){
    tab = try(importCodes(dict[['Levels']][v]))
    if (class(tab)[1]=='try-error') codeCheck<-c(codeCheck,v)
  }
  if (!length(codeCheck)==0) stop(paste('The codes for the following variables could not be read:',dict[["VariableName"]][codeCheck],
                                        '\n\nPlease ensure that codes are in the form code=label and that codes and categories are comma-separated.'))

  # For any dates, convert the numbers to a character string of the date
  for (v in 1:nrow(dict)){
    if (dict[v,'Type']=='date'){
      if (testForNumeric(dict[v,'Minimum'])) dict[v,'Minimum'] <- as.character(as.Date(as.numeric(dict[v,'Minimum']), origin = origin))
      if (testForNumeric(dict[v,'Maximum'])) dict[v,'Maximum'] <- as.character(as.Date(as.numeric(dict[v,'Maximum']), origin = origin))
    }
  }

  # Create a log file if none exists
  if (Sys.getenv("EXCEL_LOG")=="") Sys.setenv(EXCEL_LOG=paste0(gsub("[.].*",'',excelFile),format(Sys.Date(),'%d%b%y'),'.log'))

  WriteToLog(paste('Data Dictionary imported from: ', excelFile, 'from sheet: ',dictionarySheet))

  # return the dictionary
  return(dict)
}

#' Import Excel Data from a DataDictionary file
#'
#' This function reads in a data dictionary and data entry table and converts
#' code and category variables to factors as outlined in the dictionary. This
#' code is to be used in conjection with the DataDictionary.xlsm template
#' template file according to the specifications in the dataDictionary
#'
#' Prior to reading in the data, the dataDictionary file must be imported using
#' readDataDict.
#'
#' Warning: If SetErrorsMissing = TRUE then a subsequent call to checkData will not return any errors, because the errors have been set to missing.
#'
#' NOTE: This function will only read in those columns present in the DataDictionary
#' @param excelFile path and filename of the data file
#' @param dictionarySheet the name of the sheet containing the data dictionary, defaults to 'DataDictionary'
#' @param dataSheet name of the data entry sheet within the file, defaults to 'DataEntry'
#' @param saveWarnings Boolean, if TRUE and there are any warnings then the function will return a list with the data frame and the import warnings
#' @param setErrorsMissing Boolean, if TRUE all values out of range will be set to NA
#' @param range Optional, Range of Excel sheet to restrict import to (ie. range="A1:F6")
#' @param origin Optional, the date origin of Excel dates, defaults to 30 December 1899
#' @param timeUnit character specifying the unit of time that should be used when created survival type variables
#' @return a list containing two data frames: the data dictionary and the data table
#' @export
importExcelData <- function(excelFile,dictionarySheet='DataDictionary',dataSheet='DataEntry',saveWarnings=FALSE,setErrorsMissing=TRUE,range,origin,timeUnit='month'){
  if (missing(excelFile) ) stop('The excel file containing the data dictionary and data entry table are required')
  if (missing(range)) range = NULL
  if (missing(origin)) origin = "1899-12-30"

  # Create a log file if none exists
  if (Sys.getenv("EXCEL_LOG")=="") Sys.setenv(EXCEL_LOG=paste0(gsub("[.].*",'',excelFile),format(Sys.Date(),'%d%b%y'),'.log'))
  WriteToLog(msg =  'Log File Created',timestamp = T,append=F)

  dictTable <- readDataDict(excelFile,dictionarySheet =dictionarySheet)

  dataTable <- readExcelData(excelFile,dataDictionary =  dictTable,dataSheet=dataSheet,saveWarnings=saveWarnings,setErrorsMissing=setErrorsMissing,range,origin)

  factorData <- addFactorVariables(dictTable,dataTable,keepOriginal = FALSE)

  fullData <- createCalculated(factorData,dictTable,timeUnit='month')


  cat('File import complete. Details of variables created are in the logfile: ',Sys.getenv("EXCEL_LOG"))


  # RUN CHECKS TO ENSURE THE Calculated Variables are Correct
  return(list(dictionary=dictTable,data=fullData))
}

#' Read Excel Data
#'
#' This function reads in an excel data table created by the DataDictionary.xlsm
#' template file according to the specifications in the dataDictionary
#'
#' Prior to reading in the data, the dataDictionary file must be imported using
#' readDataDict.
#'
#' Warning: If SetErrorsMissing = TRUE then a subsequent call to checkData will not return any errors, because the errors have been set to missing.
#'
#' NOTE: This function will only read in those columns present in the DataDictionary
#' @param dataDictionary a data frame returned by readDataDict
#' @param excelFile path and filename of the data file
#' @param dataSheet name of the data entry sheet within the file, defaults to 'DataEntry'
#' @param saveWarnings Boolean, if TRUE and there are any warnings then the function will return a list with the data frame and the import warnings
#' @param setErrorsMissing Boolean, if TRUE all values out of range will be set to NA
#' @param range Optional, Range of Excel sheet to restrict import to (ie. range="A1:F6")
#' @param origin Optional, the date origin of Excel dates, defaults to 30 December 1899
#' @return a data frame containing the imported data
#' @export
readExcelData <- function(excelFile,dataDictionary,dataSheet='DataEntry',saveWarnings=FALSE,setErrorsMissing=FALSE,range,origin){
  if (missing(excelFile) | missing(dataDictionary)) stop(paste('Both the excel data file and the data dictionary are required arguments.\n',
                                                               'Use exceldata::readDataDict to read in the data dictionary before importing data.\n',
                                                               'Or run exceldata::importExcelData to import the dictionary and data files and create factor variables.'))
  if (all(names(dataDictionary) != c('VariableName', 'Description', 'Type', 'Minimum', 'Maximum', 'Levels'))) {
    stop('The specified dictionary does not have the expected columns. \nTry running readDataDict again.')
  }
  if (missing(range)) range = NULL
  if (missing(origin)) origin = "1899-12-30"
  col_types = sapply(dataDictionary[["Type"]],function(x){
    if (x %in% c('integer','numeric','codes')) x <-'numeric'
    if (x %in% c('category','character','calculated')) x <-'text'
    return(x)},simplify = T)
  varLookup = data.frame(VariableName = dataDictionary[['VariableName']],
                         col_type = col_types)
  # Read in the Excel data
  dat <- try(readxl::read_excel(excelFile,sheet=dataSheet,col_names = T,range=range),silent = T)
  if (class(dat)[1]=='try-error') stop(paste('File access failure. \n Check that sheet',dataSheet,'exists in file:\n',excelFile,
                                             '\n\nNOTE: It may be necessary to close Excel for this \nfunction to work.'))

  # check that the variables match the data dictionary
  if (!all(dataDictionary[['VariableName']] %in% names(dat))){
    warning(paste('The following variable are missing from the datafile:\n',
                  setdiff(dataDictionary[['VariableName']], names(dat))))
  }

  # Assign the col_type of any variables missing in the dataDictionary to 'guess'
  dataFileColumns = data.frame(VariableName = names(dat),position=1:length(names(dat)))
  dataFileColumns = merge(dataFileColumns,varLookup,all.x = T)
  dataFileColumns = dataFileColumns[order(dataFileColumns[["position"]]),]
  import_types = dataFileColumns[["col_type"]]
  import_types[is.na(import_types)] <- 'guess'
  entry_warning <- NULL
  withCallingHandlers(dat <- try(readxl::read_excel(excelFile,sheet=dataSheet,col_names = T,range=range,col_types = import_types),silent=T),
                      warning = function(w) { entry_warning <<-append(entry_warning,w)})

  if (class(dat)[1]=='try-error') stop(paste('File import failure.'))

  # remove all columns not in the data dictionary
  dat <- dat[,which(import_types!='guess')]

  # assign integer types to integer fields
  for (v in dataDictionary$VariableName[dataDictionary$Type=='integer']) dat[[v]] <- as.integer(dat[[v]])

  # assign date types to date fields
  for (v in dataDictionary$VariableName[dataDictionary$Type=='date']) dat[[v]] <- as.Date(dat[[v]])

  # Set all out of range entries to missing if specified
  if (setErrorsMissing){

    # Range checks
    varsToCheck <- intersect(dataDictionary[['VariableName']][dataDictionary[['Type']] %in% c('numeric','codes','category','integer','date')],names(dat))

    for (v in varsToCheck){
      # Numeric Data
      if (dataDictionary[['Type']][dataDictionary[['VariableName']]==v] %in% c('integer','numeric')){
        minVal = dataDictionary[['Minimum']][dataDictionary[['VariableName']]==v]
        if (minVal %in% dataDictionary[["VariableName"]]) minVal = dat[[minVal]] else minVal = as.numeric(minVal)
        maxVal = dataDictionary[['Maximum']][dataDictionary[['VariableName']]==v]
        if (maxVal %in% dataDictionary[["VariableName"]]) maxVal = dat[[maxVal]] else maxVal = as.numeric(maxVal)

        check = minVal <= dat[[v]] & dat[[v]] <= maxVal
      }

      # Dates
      if (dataDictionary[['Type']][dataDictionary[['VariableName']]==v] =='date'){
        minVal = dataDictionary[['Minimum']][dataDictionary[['VariableName']]==v]
        if (minVal %in% dataDictionary[["VariableName"]]) minVal = as.Date(dat[[minVal]]) else if (minVal=='today') minVal=Sys.Date() else  minVal = as.Date(minVal)
        maxVal = dataDictionary[['Maximum']][dataDictionary[['VariableName']]==v]
        if (maxVal %in% dataDictionary[["VariableName"]]) maxVal = as.Date(dat[[maxVal]]) else if (maxVal=='today') maxVal=Sys.Date() else maxVal = as.Date(maxVal)

        check = as.numeric(minVal) <= as.numeric(as.Date(dat[[v]])) & as.numeric(as.Date(dat[[v]])) <= as.numeric(maxVal)
      }

      # Factors
      if (dataDictionary[['Type']][dataDictionary[['VariableName']]==v] %in% c('category','codes')){
        allowedCodes = importCodes(dataDictionary[['Levels']][dataDictionary[['VariableName']]==v])[['code']]
        if (dataDictionary[['Type']][dataDictionary[['VariableName']]==v] =='codes') allowedCodes = as.numeric(allowedCodes)
        check = dat[[v]]  %in% allowedCodes
      }
      dat[[v]][!check] <- NA
    }

  }


  if (!is.null(entry_warning) & saveWarnings){
    entry_warning <- entry_warning[names(entry_warning)=='message']
    entry_warning <- as.vector(do.call(rbind,entry_warning))
    return(list(data=dat,warnings=entry_warning))
  } else {
    if (!is.null(entry_warning)){
      cat('To store entry warnings, re-run function with saveWarnings=TRUE.\n A list with data,warnings will be returned.\n')
    }

    # Create a log file if none exists
    if (Sys.getenv("EXCEL_LOG")=="") Sys.setenv(EXCEL_LOG=paste0(gsub("[.].*",'',excelFile),format(Sys.Date(),'%d%b%y'),'.log'))

    # write to the logfile
    WriteToLog(paste('Data read in from Excel file :',excelFile,' from sheet: ', dataSheet))

    return(dat)
  }}

#' Check the entered data against the data dictionary
#'
#' This function compares the data in the data entry table against the
#' specifications in the DataDictionary.
#'
#' Prior to reading in the data, the dataDictionary must be imported using
#' readDataDict and the dataTable must be imported using readExcelData.
#'
#' The function will check all variables in the dataDictionary.
#' If variables are missing from the dataDictionary an error will occur.
#' If variables are missing from the data table a warning will be shown.
#'
#'
#' @param dataDictionary a data frame returned by readDataDict
#' @param dataTable a data frame returned by readExcelData
#' @param id a string indicating the ID variable, to display errors by ID instead of row number
#' @return A list with three data frames: one with all errors, one with errors by row
#' (or ID if supplied) and one with errors by variable. Also returns a check for duplicate rows.
#' @export
#'
checkData <-function(dataDictionary,dataTable,id){
  # excelFile = '/Users/lisaavery/OneDrive/HB_survey/dataFile.xlsx'
  # dataDictionary = exceldata::readDataDict(excelFile)
  # dataTable = exceldata::readExcelData(excelFile,dataDictionary)

  if (!('data.frame' %in% class(dataDictionary))) stop('dataDictionary must be a data dictionary imported using readDataDict')
  if (!('data.frame' %in% class(dataTable))) stop('dataTable must be a data entry table imported using readExcelData')
  if (!missing(id)) if (!id %in% names(dataTable)) stop(paste(id,'not found in the dataTable. Specify a valid ID variable.'))
  if (all(names(dataDictionary) != c('VariableName', 'Description', 'Type', 'Minimum', 'Maximum', 'Levels'))) {
    stop('The specified dictionary does not have the expected columns. \nTry running readDataDict again.')
  }

  if (!all(names(dataTable) %in% dataDictionary[['VariableName']])){
    stop(paste('Variables missing from the dataDictionary:\n',
               setdiff(names(dataTable),dataDictionary[['VariableName']])))
  }

  if (!all(dataDictionary[['VariableName']] %in% names(dataTable) )){
    warning(paste('Variables missing from the data table:\n',
                  paste(setdiff(dataDictionary[['VariableName']],names(dataTable)),collapse = ',')))
  }

  # check for duplicate rows
  if (any(duplicated(dataTable),na.rm = T)){
    dupl = ifelse(missing(id),(1:nrow(dataTable))[duplicated(dataTable)],dataTable[[id]][duplicated(dataTable)])
    dupl= paste0('The following rows (IDs) are duplicated: ',paste(dupl,collapse = ","))
  } else {
    dupl = 'No duplicated rows'
  }

  # Range checks
  varsToCheck <- intersect(dataDictionary[['VariableName']][dataDictionary[['Type']] %in% c('numeric','codes','category','integer','date')],names(dataTable))
  df_checks <- NULL
  for (v in varsToCheck){
    if (dataDictionary[['Type']][dataDictionary[['VariableName']]==v] %in% c('integer','numeric')){
      minVal = dataDictionary[['Minimum']][dataDictionary[['VariableName']]==v]
      if (minVal %in% dataDictionary[["VariableName"]]) minVal = dataTable[[minVal]] else minVal = as.numeric(minVal)
      maxVal = dataDictionary[['Maximum']][dataDictionary[['VariableName']]==v]
      if (maxVal %in% dataDictionary[["VariableName"]]) maxVal = dataTable[[maxVal]] else maxVal = as.numeric(maxVal)

      check = minVal <= dataTable[[v]] & dataTable[[v]] <= maxVal
    }

    if (dataDictionary[['Type']][dataDictionary[['VariableName']]==v] =='date'){
      minVal = dataDictionary[['Minimum']][dataDictionary[['VariableName']]==v]
      if (minVal %in% dataDictionary[["VariableName"]]) minVal = as.Date(dataTable[[minVal]]) else if (minVal=='today') minVal=Sys.Date() else  minVal = as.Date(minVal)
      maxVal = dataDictionary[['Maximum']][dataDictionary[['VariableName']]==v]
      if (maxVal %in% dataDictionary[["VariableName"]]) maxVal = as.Date(dataTable[[maxVal]]) else if (maxVal=='today') maxVal=Sys.Date() else maxVal = as.Date(maxVal)

      check = as.numeric(minVal) <= as.numeric(as.Date(dataTable[[v]])) & as.numeric(as.Date(dataTable[[v]])) <= as.numeric(maxVal)
    }

    if (dataDictionary[['Type']][dataDictionary[['VariableName']]==v] %in% c('category','codes')){
      allowedCodes = importCodes(dataDictionary[['Levels']][dataDictionary[['VariableName']]==v])[['code']]
      if (dataDictionary[['Type']][dataDictionary[['VariableName']]==v] =='codes') allowedCodes = as.numeric(allowedCodes)
      check = dataTable[[v]]  %in% allowedCodes
    }

    if (any(check==FALSE,na.rm=T)) df_checks[[v]] = check
  }
  if (is.null(df_checks)){
    message('No errors in data.')
    return(NULL)
  } else {
    entry_errors = as.data.frame(!do.call(cbind, df_checks))
    rowsToKeep = rowSums(entry_errors,na.rm=T)>0
    # keep only rows with errors
    if (missing(id)){
      id = 'originalRowID'
      rowIDs = 1:nrow(entry_errors)
    } else{
      rowIDs = dataTable[[id]]
    }
    entry_errors<- cbind(rowIDs,entry_errors)
    names(entry_errors)[1] <- id
    entry_errors <- entry_errors[rowsToKeep,]
    row_errors = data.frame(originalRowID = entry_errors[[id]],
                            Errors = sapply(1:nrow(entry_errors),function(i){
                              paste(names(entry_errors)[-1][as.logical(as.vector(entry_errors[i,-1]))],collapse = ",")
                            }))
    names(row_errors)[1] <- id
    var_errors = data.frame(Variable = names(entry_errors)[-1],
                            Row_Errors = sapply(names(entry_errors)[-1],function(v){
                              paste(entry_errors[[id]][as.logical(as.vector(entry_errors[[v]]))],collapse = ",")
                            }))
    rownames(var_errors) <- NULL
    colnames(var_errors)[2] <- ifelse(missing(id),'Row_Errors','IDs_With_Errors')
    return(list(errors_by_row=row_errors,errors_by_variable=var_errors,duplicated_entries=dupl,error_dataframe = entry_errors))
  }
}


#' Create factor variables from data dictionary
#'
#' This function will replace the code and category variables
#' with factors based on the factor levels provided in the data
#' dictionary. The original variables are retained with the suffix
#' '_orig'
#'
#' @param dataDictionary a data frame returned by readDataDict
#' @param dataTable a data frame returned by readExcelData
#' @param keepOriginal Boolean indicating if the original character variables should be kept
#' @return a data frame with the updated factor variables
#' @export
addFactorVariables <-function(dataDictionary,dataTable,keepOriginal=FALSE){
  if (!('data.frame' %in% class(dataDictionary))) stop('dataDictionary must be a data dictionary imported using readDataDict')
  if (!('data.frame' %in% class(dataTable))) stop('dataTable must be a data entry table imported using readExcelData')
  if (all(names(dataDictionary) != c('VariableName', 'Description', 'Type', 'Minimum', 'Maximum', 'Levels'))) {
    stop('The specified dictionary does not have the expected columns. \nTry running readDataDict again.')
  }

  if (!all(names(dataTable) %in% dataDictionary[['VariableName']])){
    stop(paste('Variables missing from the dataDictionary:\n',
               setdiff(names(dataTable),dataDictionary[['VariableName']])))
  }

  if (!all(dataDictionary[['VariableName']] %in% names(dataTable) )){
    warning(paste('Variables missing from the data table:\n',
                  paste(setdiff(dataDictionary[['VariableName']],names(dataTable)),collapse = ',')))
  }
  varsWithCodes <- intersect(dataDictionary[['VariableName']][dataDictionary[['Type']] %in% c('codes','category')],names(dataTable))
  for (v in varsWithCodes){
    factorLevels = try(importCodes(dataDictionary[["Levels"]][dataDictionary[['VariableName']]==v]),silent = T)
    if (!class(factorLevels)[1]=='try-error'){
      if (keepOriginal) dataTable[[paste0(v,'_orig')]] <- dataTable[[v]]
      if (dataDictionary[['Type']][dataDictionary[['VariableName']]==v]=='codes'){
        factorVar <- factor(dataTable[[v]],levels=factorLevels$code,labels=factorLevels$label)
      } else {
        factorVar <- factor(dataTable[[v]],levels=factorLevels$code)
      }
      dataTable[[v]] <-factorVar
    } else warning(paste('Factor could not be created for',v))
  }

  WriteToLog(paste('Factor structures added for: ',paste(varsWithCodes,collapse=', ')))

  return(dataTable)
}




testForNumeric <- function(str){
  testResults = sapply(str,function(x){
    if (is.na(suppressWarnings(as.character(as.numeric(x)))==x)) {
      rtn=FALSE } else {
        rtn=suppressWarnings(as.character(as.numeric(x)))==x
      }
    return(rtn)
  })
  unname(testResults)
}

#' Return a data frame of codes
#'
#' Accepts a string input in the form "code1=label1,code2=label2,.." and
#' returns a data frame with a column of codes and a column of labels
#'
#' @param labelStr in the format code1=label1,code2=label2
#' @param delim delimeter separating codes in labelStr, defaults to ','
#' @export
importCodes<-function(labelStr,delim=','){
  x=strsplit(labelStr,split=delim)[[1]]
  # check for multiple '=', probably indicates an error
  check = sapply(x,function(y){ifelse(nchar(gsub('=','',y))+1 < nchar(y),1,0)})
  if (sum(check)>0) stop(paste('Check these codes:\n',x[check==1],'\nMultiple "=" signs not allowed. Did you forget a delimiter?'))
  codeLst=strsplit(x,"=")
  tbl <- NULL
  for (i in seq_along(codeLst)) tbl<-rbind(tbl,cbind(code=codeLst[[i]][1],label=codeLst[[i]][2]))
  tbl <- as.data.frame(tbl)
  if (isTRUE(all.equal(as.character(suppressWarnings(as.numeric(tbl[[1]]))),tbl[[1]]))) tbl[[1]] <- as.numeric(tbl[[1]])
  names(tbl)=c('code','label')
  tbl[['code']] <- sapply(tbl[['code']],function(x) trimws(x))
  tbl[['label']] <- sapply(tbl[['label']],function(x) trimws(x))
  return(tbl)
}

#' Create calculated variables
#'
#' This function will create survival and recoded variables according to the
#' rules in the DataDictionary.xlsm file. See the Example sheet for an example.
#'
#' @param data is data returned by the importExcelData or readExcelData functions
#' @param dictTable is the data dictionary returned by importExcelData or readDataDict functions
#' @param timeUnit a string containing the desired unit of time for survival variables
#' @export
createCalculated<-function(data,dictTable,timeUnit='month'){
  calcVars <- dictTable$VariableName[dictTable$Type=='calculated']

  # Calculation instructions should start with 'start' for survival type variable or a variable name for recoding
  for (v in calcVars){

    # first make sure nothing has been entered for the calculated variable
    if (sum(is.na(data[[v]])) !=nrow(data)) {
      suffix <- 1
      repeat{
        newVarName <- paste0(v,'_',suffix)
        if (newVarName %in% names(data)) suffix=suffix+1 else break
      }
      warning(paste('There is already data entered for ',v,'.\n The calculated variable has been stored in ',newVarName))
    } else {
      newVarName <- v
    }

    instructions <- trimws(unlist(strsplit(dictTable$Levels[dictTable$VariableName==v],',')))

    # Determine what type of calculation to do
    if (substr(dictTable$Levels[dictTable$VariableName==v],1,5)=='start') {
      if (length(instructions)!=3) {
        warning(paste(newVarName,' not created. For a survival variable, a start date, event date and last date followd must be specified. Check dictTableionary.'))
        survVars <- sapply(instructions,function(x) trimws(unlist(strsplit(x,'='))[2]))
        if (!all(survVars %in% names(data))) {
          warning(paste(paste(survVars[!survVars %in% names(data)],collapse=','),' not in the data. Variable names are case sensitive.\n',newVarName,' not created.'))
        }} else{
          data <- createSurvVar(data,newVarName,survVars,timeUnit)
        }
    } else if (substr(dictTable$Levels[dictTable$VariableName==v],1,7)=='combine') {
      varsToCombine <- unlist(strsplit(trimws(gsub('combine|to','',instructions[1])),' '))
      varsToCombine <- varsToCombine[varsToCombine!=""]
      if (!all(varsToCombine %in% names(data))) {
        warning(paste(paste(varsToCombine[!varsToCombine %in% names(data)],collapse=','),' not in the data. Variable names are case sensitive.\n',newVarName,' not created.'))
      } else if (length(varsToCombine)!=2) {
        warning(paste0(newVarName,' not created, check instructions. Must be in the format: combine var1 to var2, response=value'))
      }  else if (substr(instructions[2],1,8)!='response'){
        warning(paste0(newVarName,' not created, check instructions. Must be in the format: combine var1 to var2, response=value'))
      } else{
        responseVal <- trimws(gsub('response|=','',instructions[2]))
        data <- createCombinedVar(data,dictTable,newVarName,varsToCombine,responseVal)
      }
    } else {

      if (!instructions[1] %in% names(data)) {
        warning(paste(instructions[1],' not in the data. Variable names are case sensitive.\n',newVarName,' not created.'))
      } else {
        if (dictTable$Type[dictTable$VariableName==instructions[1]] %in% c('integer','numeric')) {
          data <- createCategorisedVar(data,newVarName,instructions)

        } else{
          data <- createRecodedVar(data,dictTable,newVarName,instructions)
        }
      }
    }

  }
  return(data)
}

#' Create survival variables (survival duration + status)
#'
#' This function will create survival variables from an existing start variable
#' date of event variable and last date followed variable
#'
#' @param data is data returned by the importExcelData or readExcelData functions
#' @param newVarName the name of the new survival variable. The status variable will be suffixed with '_status'
#' @param survVars are, in order the start date, event date and date of last followup
#' @param timeUnit string, the unit of time to calculate survival variables for (day week month year)
createSurvVar <- function(data,newVarName,survVars,timeUnit='month'){

  # ensure that the status variable doesn't already exist, if it does create a new one with a suffix
  newStatusVar <- paste0(newVarName,'_status')
  if (newStatusVar %in% names(data)){
    suffix <- 1
    repeat{
      newStatusVar <- paste0(newStatusVar,suffix)
      if (newStatusVar %in% names(data)) suffix=suffix+1 else break
    }}

  # Determine if the event occured
  status <- ifelse(is.na(data[[survVars[2]]]),0L,1L)

  # Determine the end date for calculating the interval NOTE: requires dplyr::if_else
  end_date <- dplyr::if_else(is.na(data[[survVars[2]]]),data[[survVars[3]]],data[[survVars[2]]])

  # Determine time to event
  TTE <- lubridate::time_length(lubridate::interval(data[[survVars[1]]],end_date),timeUnit)
  attributes(TTE)$timeUnit <- timeUnit

  # Add the new variables to the data
  data[[newVarName]] <- TTE
  data[[newStatusVar]] <- status

  # Log
  WriteToLog(paste('New survival variables created: ',paste(c(newVarName,newStatusVar),collapse=', '),' from ', paste(survVars,collapse=', '), '- time unit is ',paste0(timeUnit,'s')))

  return(data)
}


#' Create survival variables (survival duration + status)
#'
#' This function will create survival variables from an existing start variable
#' date of event variable and last date followed variable
#'
#' The instructions are contained in the Levels column of the data dictionary and should be in the format:
#' original_varname,newCode1=oldcode1,oldCode2,...,newCode2=oldCode3,..
#' For Example:
#' instructions="T_Stage,T0=T0,T1up=T1,T2,T3,T4" will recode T1-T4 as T1up and retain T0
#'
#' @param data is data returned by the importExcelData or readExcelData functions
#' @param dictTable is the data dictionary returned by importExcelData or readDataDict functions
#' @param newVarName the name of the new variable.
#' @param instructions are from the data dictionary
createRecodedVar <- function(data,dictTable,newVarName,instructions){

  originalVar <- instructions[1]

  codeReps <-diff(  c(grep("=",instructions),length(instructions)+1))
  newCodes <- sapply(instructions[grep("=",instructions)],function(x) trimws(unlist(strsplit(x,"="))[1]))
  oldCodes=gsub('.*=','',instructions[-1])
  labels = unlist(mapply(rep,x=newCodes,times=codeReps))


  # If the variable is a codes variable, then we need the original entered levels
  if ('factor' %in% class(data[[originalVar]])) {
    factorLevels = try(importCodes(dictTable[["Levels"]][dictTable[['VariableName']]==originalVar]),silent = T)

    if (class(factorLevels)[1]=='try-error'){
      warning(paste('Data codes for',originalVar,'could not be extracted,', newVarName,'not created.'))
    } else {

      oldCodes <- data.frame(oldCodes=oldCodes)
      codeLookup <- merge(oldCodes,factorLevels,by.x = "oldCodes", by.y = "code" )
      recoded = droplevels(factor(data[[originalVar]],levels = codeLookup[,2],labels=labels))
    }
  } else{
    recoded = droplevels(factor(data[[originalVar]],levels = oldCodes,labels=labels))

  }
  # Add the variable to the data
  data[[newVarName]] <- recoded

  WriteToLog(paste('New recoded variables created: ',newVarName,' from ',instructions[1]))
  tbl_check <- capture.output(table(data[[newVarName]],data[[originalVar]]))
  WriteToLog(tbl_check)

  return(data)
}

#' Create a combined variable from several dichotomous variables
#'
#' This function will create a single variable from a set of dichotomous
#' variables, usually checkbox items from a survey. The combined variable may be
#' if there are a small number of popular response patterns. Currently this
#' function only works with dichotomous variables.
#'
#' The instructions are contained in the Levels column of the data dictionary and should be in the format:
#' original_varname,newCode1=oldcode1,oldCode2,...,newCode2=oldCode3,..
#' For Example:
#' instructions="T_Stage,T0=T0,T1up=T1,T2,T3,T4" will recode T1-T4 as T1up and retain T0
#'
#' @param data is data returned by the importExcelData or readExcelData functions
#' @param dictTable is the data dictionary returned by importExcelData or readDataDict functions
#' @param newVarName the name of the new variable.
#' @param varsToCombine a vector of the first and last variables to combine into the new variable. Note that the variables to be conbined mut be contiguous in the data.
#' @param responseVal the value of the variables to be combined, usually this will be 1 for 0,1 variables or Yes for Yes No or Checked for Checked Unchecked
createCombinedVar <- function(data,dictTable,newVarName,varsToCombine,responseVal){

  startVar = which(names(data)==varsToCombine[1])
  endVar = which(names(data)==varsToCombine[2])
  var_dat <- data[,startVar:endVar]

  # determine if data has already been turned into factors, or if it is numeric.
  # If all columns are not the same time warning user and don't create the variable
  types <- unlist(lapply(var_dat,function(x) class(x)[1]))

  if (length(unique(types))!=1) {
    warning(paste0('All variables to be combined must be the same type ',newVarName,' not created.') )
  } else if (types[1]=='factor') {
    factorLevels = try(importCodes(dictTable[["Levels"]][dictTable[['VariableName']]==varsToCombine[1]]),silent = T)
    responseIn <- unlist(lapply(factorLevels,function(x) responseVal %in%  x))
    if (sum(responseIn)==0) {
      warning(paste0('The response value ',responseVal,' is not a response to the variable to be combined. ',newVarName,' not created.') )
    } else {
      # check if the response is in var_dat or if the values and codes need to be swapped
      if (!any(unlist(lapply(var_dat, function(x) responseVal %in% x)))){
        need <- names(responseIn)[responseIn]
        if (need=='code'){
          new_dat <-NULL
          for (j in 1:ncol(var_dat)){
            x <- var_dat[[j]]
            for (i in 1:nrow(factorLevels)) {
              x <- gsub(factorLevels$label[i],factorLevels$code[i],x)
            }
            new_dat <-cbind(new_dat,x)
          }
          var_dat=as.data.frame(new_dat)
        } else {
          new_dat <-NULL
          for (j in 1:ncol(var_dat)){
            x <- var_dat[[j]]
            for (i in 1:nrow(factorLevels)) {
              x <- gsub(factorLevels$code[i],factorLevels$label[i],x)
            }
            new_dat <-cbind(new_dat,x)
          }
          var_dat=as.data.frame(new_dat)
        }
      }
    }

    check <- unlist(lapply(var_dat,function(x) responseVal %in%  x))
    if (!any(check)){
      warning(paste0('The response value ',responseVal,' was not found in the variables to combine. ',newVarName,' not created.') )
    } else{
      m <- as.matrix(var_dat)
      m[m!=responseVal] <- NA
      m[m==responseVal] <- 1L
      m[is.na(m)] <- 0L
      # row combine
      combinedVar <- apply(m,1,function(x) paste(x,collapse=''))

      # Add the variable to the data
      data[[newVarName]] <- combinedVar

      # give it an attribute so it can be plotted
      attributes(data[[newVarName]])$graphType = 'pareto'
      WriteToLog(paste('New combined variables created: ',newVarName,' from ',paste(varsToCombine,collapse=' to ')))
    }
  }
  return(data)
}

#' Create categories from continuous data
#'
#' This function will create categories based on the ranges provided in the instructions
#'
#' @param data is data returned by the importExcelData or readExcelData functions
#' @param newVarName the name of the new  variable. Must be empty in data
#' @param instructions category names and bounds
#'
createCategorisedVar <- function(data,newVarName,instructions){

  originalVar <- instructions[1]

  contData <- data[[originalVar]]

  categories <- trimws(gsub('=.*','',instructions[-1]))
  rules <- trimws(sub('.*=','',instructions[-1])[grep("=",instructions[-1])])

  # There should be one more category than rule
  if (length(categories)!=(length(rules)+1)) {
    warning(paste(newVarName,' not created, check the Levels column in the data dictionary, rule not correctly formatted.'))
    WriteToLog(paste('Variable creation failed: ',newVarName,' from ',instructions[1]))

  } else{
    recodingRule <-   'dplyr::case_when('
    for (i in seq_along(rules)){
      recodingRule <- paste0(recodingRule,originalVar,rules[i],'~"',categories[i],'",')
    }
    recodingRule <- paste0(recodingRule,'TRUE~"',categories[i+1],'")')

    eval(parse(text=paste("recoded <- with(data,",recodingRule,')')))
    recoded <- factor(recoded,levels = categories)

    # Add the variable to the data
    data[[newVarName]] <- recoded

    WriteToLog(paste('New recoded variables created: ',newVarName,' from ',instructions[1],'\n',recodingRule))
    tbl_check <- capture.output(aggregate(data[[originalVar]],by=  list(data[[newVarName]]),FUN=function(x) {c(min=min(x,na.rm=T),max=max(x,na.rm=T))}))
    WriteToLog(tbl_check)

  }
  return(data)
}

WriteToLog <- function(msg,append=T,timestamp=F){

  # Ensure the logfile exists, create if it doesn't
  logfile = ifelse(Sys.getenv("EXCEL_LOG")=="",paste0('logfile',format(Sys.Date(),'%d%b%y'),'.log'),Sys.getenv("EXCEL_LOG"))
  if (!file.exists(logfile)) file.create(logfile)

  if (timestamp) timePrint <- Sys.time() else timePrint <-""

  # Write the message to the log file
  cat(file = Sys.getenv("EXCEL_LOG"),
      paste('\n\n',timePrint,msg,sep=' '),
      append=append)

}

#' Return a list of univariate ggplots for each non-character variable
#'
#' This function should be run as the final step after the data have been
#' imported, checked and the factor variables created.
#' @param data is a data frame containing the variables to be plotted
#' @param dictTable optional, the data dictionary returned by importExcelData or readDataDict functions to provide plot titles
#' @param IDvar an optional string indicating the name of an identifying variable to highlight outliers
#' @param vars is an optional character vector of the names of variables to plot
#' @param showOutliers boolean, should outliers be labelled
#' @param nOut integer only used if showOutliers=TRUE the most extreme nOut values are highlighted, takes precedent over qOut
#' @param qOut proportion between 0 and 1 only used if showOutliers=TRUE the most extreme qOut values are highlighted
#' @import ggplot2
#' @importFrom scales dateformat
#' @export
plotVariables<-function(data,vars,dictTable,IDvar,nAsBar=6,showOutliers=FALSE,nOut=NULL,qOut=.05){
  varTypes = sapply(data,function(x) class(x)[1])
  if (missing(vars)){
    char_vars = names(data)[varTypes=='character']
    paretoVars <- NULL
    for (v in char_vars) if (!is.null(attributes(data[[v]])$graphType)) if (attributes(data[[v]])$graphType=='pareto') paretoVars <- c(paretoVars,v)
    vars <- c(names(data)[!varTypes=='character'],paretoVars)
  }
  N = nrow(data)
  plots <- NULL
  # Check that supplied variables are not character - if characters are found issue warning
  for (v in vars) {
    if ('character' %in% class(data[[v]])) {
      if (is.null(attributes(data[[v]])$graphType))  warning(paste(v,'is a character variable and will not be plotted. Convert to factor for plotting'))
    }
    # For each variable choose an appropriate plot based on sample size and variable type and save to the list

    if ('character' %in% class(data[[v]])){
      patterns <- sort(table(data[[v]]),decreasing = T)
      npat <- min(10,length(patterns))
      data[[v]] <- factor(data[[v]],levels = names(patterns)[1:10])
      p <- ggplot2::ggplot(data=subset(data,!is.na(data[[v]])),aes(x=.data[[v]])) +
        geom_bar() +
        theme_bw() +
        annotate('text',x=Inf,y=Inf,label=paste(npat,'most common patterns'),vjust=1,hjust=1,colour='red')
    } else if ('factor' %in% class(data[[v]]) | 'integer' %in% class(data[[v]]) ){
      p <- ggplot2::ggplot(data=data,aes(x=.data[[v]])) +
        geom_bar() +
        theme_bw()
    } else {
      # make dots smaller for larger files
      dotsize = ifelse(N<30,1,ifelse(N<200,.5,.25))

      p <- ggplot2::ggplot(data=data,aes(x=.data[[v]])) +
        geom_dotplot(dotsize=dotsize) +
        theme_bw() +
        theme(axis.title.y = element_blank(),axis.ticks.y=element_blank(),axis.text.y = element_blank())

      if (showOutliers){
        if (is.null(nOut)) {
          nOut = floor(qOut*nrow(data))
        } else{
          qOut = nOut/nrow(data)
        }
        highlights <- quantile(data[[v]],probs = c(qOut/2,1-qOut/2))
        data$highlight <- ifelse(data[[v]]<highlights[1]|data[[v]]>highlights[2],TRUE,FALSE)
        if (missing(IDvar)) {
          data$outLabel <- 1:nrow(data)
        } else  if (!IDvar %in% names(data)) {
          data$outLabel <- 1:nrow(data)
        } else {
          data$outLabel <- data[[IDvar]]
        }

        p <- p + geom_dotplot(data=subset(data,data$highlight),
                              dotsize = dotsize,
                              fill='red')
        p <- p + ggrepel::geom_text_repel(data=subset(data,data$highlight),aes(x=.data[[v]],y=0,label=.data[['outLabel']]),position = position_nudge(y=.2))
      }

    }

    # Show date axis a month-year for date variables
    if ("Date" %in% class(data[[v]])) p <- p +  scale_x_date(labels = scales::date_format("%m-%Y"))

    # Add a time unit to the x-lab if there is one
    if (!is.null(attributes(data[[v]])$timeUnit)) p <- p + xlab(paste0(v," (",attributes(data[[v]])$timeUnit,'s)'))

    # If the dictTable is provided at the description as a title
    if (!missing(dictTable)) p <- p + ggtitle(dictTable$Description[dictTable$VariableName==v])

    plots[[v]] <- p
  }

  return(plots)
}
