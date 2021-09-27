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

  # return the dictionary
  return(dict)
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
#' @export
readExcelData <- function(excelFile,dataDictionary,dataSheet='DataEntry',saveWarnings=FALSE,setErrorsMissing=FALSE,range,origin){
  if (missing(excelFile) | missing(dataDictionary)) stop(paste('Both the excel data file and the data dictionary are required arguments.\n',
                                     'Use exceldata::readDataDict to read in the data dictionary before importing data.'))
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
#' @return a data frame with the updated variables
#'
addFactorVariables <-function(dataDictionary,dataTable){
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
      dataTable[[paste0(v,'_orig')]] <- dataTable[[v]]
      if (dataDictionary[['Type']][dataDictionary[['VariableName']]==v]=='codes'){
        factorVar <- factor(dataTable[[v]],levels=factorLevels$code,labels=factorLevels$label)
      } else {
        factorVar <- factor(dataTable[[v]],levels=factorLevels$code)
      }
      dataTable[[v]] <-factorVar
    } else warning(paste('Factor could not be created for',v))
  }
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



