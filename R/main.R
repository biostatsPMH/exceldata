# TODO:
# A function for data import and checking based on the data dictionary
# Allow variable names to be changed
# import based on data types
# perform checks on max and min values
# perform logit checks:
# review prelim check rmd file
# check_that_x_before_y
# check for duplicates on all (+most) variables
# check that values are between max and min
# check that values are only of the specified allowed values
# option to report missing values by rowID, subjectID separately for each column
# create factor variables from codes and categories
# # read the codes for variables that are categorical or numeric codes
# for (v in which(dict[['Type']] %in% c('category','codes'))){
#   dict[['Lookup']][v] = importCodes(dict[['Levels']][v])
# }

# maybe output a data check report in the style of Hmisc summary?
#-------------------------------

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
#' @param excelFile path and filename of the data file
#' @param dictionarySheet name of the dictionary sheet within the file, defaults to 'DataDictionary'
#'
readDataDict <- function(excelFile,dictionarySheet ='DataDictionary'){
  if (missing(range)) range = NULL
  dict <- try(readxl::read_excel(excelFile,sheet=dictionarySheet,col_names = T,range=range),silent = T)
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

  # Every variable must have a Type and that it is an allowed value
  types <- c('calculated','category','character','codes','date','integer','numeric')
  if (sum(is.na(dict[['Type']]))!=0) {
    stop (paste('Every variable must have a type declared. \nVariables:', paste(which(is.na(dict[['Type']])),collapse=', '),
                'are missing types. \nUse the range argument to read in only part of the Excel sheet'))
  }


  # Check that dates, integers and numeric variables have min and max values set
  minCheck = dict[['VariableName']][is.na(dict[['Minimum']]) &dict[['Type']] %in% c('integer','numeric','date')]
  maxCheck = dict[['VariableName']][is.na(dict[['Maximum']]) &dict[['Type']] %in% c('integer','numeric','date')]
  # minCheck = sum(is.na(dict[['Minimum']][dict[['Type']] %in% c('integer','numeric','date')]))
  #   maxCheck = sum(is.na(dict[['Maximum']][dict[['Type']] %in% c('integer','numeric','date')]))
  #  if (minCheck!=0 | maxcheck!=0) stop('All integer, numeric and date variables must have ranges set.')
  if (!(is.null(minCheck) & is.null(maxCheck))) stop(paste('All integer, numeric and date variables must have ranges set.',
                                                           '\n\nCheck these variables:',ifelse(is.null(minCheck),'',minCheck),ifelse(is.null(maxCheck),'',maxCheck)))

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


