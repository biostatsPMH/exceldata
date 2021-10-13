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

  # Set all out of range entries to missing if specified
  if (setErrorsMissing){
    # Range checks
    varsToCheck <- intersect(dataDictionary[['VariableName']][dataDictionary[['Type']] %in% c('numeric','codes','category','integer','date')],names(dat))
    for (v in varsToCheck){
      if (dataDictionary[['Type']][dataDictionary[['VariableName']]==v] %in% c('integer','numeric')){
        if (dataDictionary[['Type']][dataDictionary[['VariableName']]==v] =='integer') data[[v]] <- as.integer(data[[v]])

        minVal = dataDictionary[['Minimum']][dataDictionary[['VariableName']]==v]
        if (minVal %in% dataDictionary[["VariableName"]]) minVal = dat[[minVal]] else minVal = as.numeric(minVal)
        maxVal = dataDictionary[['Maximum']][dataDictionary[['VariableName']]==v]
        if (maxVal %in% dataDictionary[["VariableName"]]) maxVal = dat[[maxVal]] else maxVal = as.numeric(maxVal)

        check = minVal <= dat[[v]] & dat[[v]] <= maxVal
      }

      if (dataDictionary[['Type']][dataDictionary[['VariableName']]==v] =='date'){
        minVal = dataDictionary[['Minimum']][dataDictionary[['VariableName']]==v]
        if (minVal %in% dataDictionary[["VariableName"]]) minVal = as.Date(dat[[minVal]]) else if (minVal=='today') minVal=Sys.Date() else  minVal = as.Date(minVal)
        maxVal = dataDictionary[['Maximum']][dataDictionary[['VariableName']]==v]
        if (maxVal %in% dataDictionary[["VariableName"]]) maxVal = as.Date(dat[[maxVal]]) else if (maxVal=='today') maxVal=Sys.Date() else maxVal = as.Date(maxVal)

        check = as.numeric(minVal) <= as.numeric(as.Date(dat[[v]])) & as.numeric(as.Date(dat[[v]])) <= as.numeric(maxVal)
      }

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
#' @param codeName column name for codes, defaults to code
#' @param lblName column name for labels, defaults to label
#' @export
importCodes<-function(labelStr,delim=',',codeName='code',lblName='label'){
  x=strsplit(labelStr,split=delim)[[1]]
  # check for multiple '=', probably indicates an error
  check = sapply(x,function(y){ifelse(nchar(gsub('=','',y))+1 < nchar(y),1,0)})
  if (sum(check)>0) stop(paste('Check these codes:\n',x[check==1],'\nMultiple "=" signs not allowed. Did you forget a delimiter?'))
  codeLst=strsplit(x,"=")
  tbl <- NULL
  for (i in seq_along(codeLst)) tbl<-rbind(tbl,cbind(code=codeLst[[i]][1],label=codeLst[[i]][2]))
  tbl <- as.data.frame(tbl)
  if (isTRUE(all.equal(as.character(suppressWarnings(as.numeric(tbl[[1]]))),tbl[[1]]))) tbl[[1]] <- as.numeric(tbl[[1]])
  names(tbl)=c(codeName,lblName)
  return(tbl)
}

#' Create calculated variables
#'
#' This function will create survival and recoded variables according to the
#' rules in the DataDictionary.xlsm file. See the Example sheet for an example.
#'
#' @param data is data returned by the importExcelData or readExcelData functions
#' @param dict is the data dictionary returned by importExcelData or readDataDict functions
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

    instructions <- unlist(strsplit(dictTable$Levels[dictTable$VariableName==v],','))

    # Determine what type of calculation to do
    if (substr(dictTable$Levels[dictTable$VariableName==v],1,5)=='start') {
      if (length(instructions)!=3) {
        warning(paste(newVarName,' not created. For a survival variable, a start date, event date and last date followd must be specified. Check dictTableionary.'))
        survVars <- sapply(instructions,function(x) unlist(strsplit(x,'='))[2])
        if (!all(survVars %in% names(data))) {
          warning(paste(paste(survVars[!survVars %in% names(data)],collapse=','),' not in the data. Variable names are case sensitive.\n',newVarName,' not created.'))
        }} else{
          data <- createSurvVar(data,newVarName,survVars,timeUnit)
        }
    } else {
      if (!instructions[1] %in% names(data)) {
        warning(paste(instructions[1],' not in the data. Variable names are case sensitive.\n',newVarName,' not created.'))
      } else {
        if (dictTable$Type[dictTable$VariableName==instructions[1]] %in% c('integer','numeric')) {
          data <- createCategorisedVar(data,newVarName,instructions)

        } else{
          data <- createRecodedVar(data,newVarName,instructions)
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
  status <- ifelse(is.na(data[[survVars[2]]]),0,1)

  # Determine the end date for calculating the interval NOTE: requires dplyr::if_else
  end_date <- dplyr::if_else(is.na(data[[survVars[2]]]),data[[survVars[3]]],data[[survVars[2]]])

  # Determine time to event
  TTE <- lubridate::time_length(lubridate::interval(data[[survVars[1]]],end_date),timeUnit)

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
#' @param newVarName the name of the new variable.
#' @param instructions are from the data dictionary
createRecodedVar <- function(data,newVarName,instructions){

  originalVar <- instructions[1]

  codeReps <-diff(  c(grep("=",instructions),length(instructions)+1))
  newCodes <- sapply(instructions[grep("=",instructions)],function(x) unlist(strsplit(x,"="))[1])
  labels = unlist(mapply(rep,x=newCodes,times=codeReps))

  oldCodes <- gsub('.*=','',instructions[-1])

  recoded = droplevels(factor(data[[originalVar]],levels = oldCodes,labels=labels))

  # Add the variable to the data
  data[[newVarName]] <- recoded

  WriteToLog(paste('New recoded variables created: ',newVarName,' from ',instructions[1]))
  tbl_check <- capture.output(table(data[[newVarName]],data[[originalVar]]))
   WriteToLog(tbl_check)

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

  categories <- gsub('=.*','',instructions[-1])
  rules <- sub('.*=','',instructions[-1])[grep("=",instructions[-1])]

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
#' @param IDvar an optional string indicating the name of an identifying variable to highlight outliers
#' @param vars is an optional character vector of the names of variables to plot
#' @param nAsBar the number of unique values to display as a bar plot for non-factor data, default is 6
#' @import ggplot2
#' @export
plotVariables<-function(data,IDvar,vars,nAsBar=6){
  ggplot2::theme_set(ggplot2::theme_bw())
  varTypes = sapply(data,function(x) class(x)[1])
  if (missing(vars)){
    vars = names(data)[varTypes!='character']
  }
  N = nrow(data)
  plots <- NULL
  # Check that supplied variables are not character - if characters are found issue warning
  for (v in vars) {
    if ('character' %in% class(data[[v]])) warning(paste(v,'is a character variable and will not be plotted. Convert to factor for plotting'))

    # For each variable choose an appropriate plot based on sample size and variable type and save to the list


    if ('factor' %in% class(data[[v]]) ){ #| length(unique(data[[v]]))<=nAsBar
      p <- ggplot2::ggplot(data=data,aes(x=.data[[v]])) +
        geom_bar()
    } else {
      if (N<50){
        p <- ggplot2::ggplot(data=data,aes(x=.data[[v]])) +
          geom_dotplot() +
          theme(axis.title.y = element_blank(),axis.ticks.y=element_blank(),axis.text.y = element_blank())
      } else{
        p <- ggplot2::ggplot(data=data,aes(x=.data[[v]])) +
          geom_dotplot(size=.5) +
          theme(axis.title.y = element_blank(),axis.ticks.y=element_blank(),axis.text.y = element_blank())

      }
    }
    plots[[v]] <- p
  }

  return(plots)
}
