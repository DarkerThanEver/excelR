library(openxlsx)
library(R6)

# base object

RsheetClass = "Rsheet"

Rsheet <- R6Class("Rsheet",
                  private = list(
                    dataSheet = NA
                  ),
                  public = list(
                    sheetName = "",
                    initialize = function(sheetData = NA, nameSheet = "", ...){
                      self$data <- sheetData
                      self$sheetName <- nameSheet
                      invisible(self)
                    },
                    addToWorkbook = function(workbook, ...){
                      if (!(self$sheetName %in% workbook$sheet_names)) {
                        addWorksheet(workbook,self$sheetName)
                      }
                      invisible(self)
                    },
                    removeData = function(...){
                      # placeholder, base object does not do anything
                      invisible(self)
                    }
                  ),
                  active = list(
                    data = function(value){
                      if (missing(value)){
                        return(private$dataSheet)
                      } else {
                        private$dataSheet <- value
                      }
                    }
                  )
)

# ----

RsheetTableClass = "RsheetTable"

RsheetTable <- R6Class("RsheetTable",
                       inherit = Rsheet,
                       public = list(
                         sheetName = "",
                         rowNames = FALSE,
                         colNames = TRUE,
                         columns = NA,
                         widths = NA,
                         sheet.startRow = 1,
                         sheet.startCol = 1,
                         initialize = function(sheetData = NA,
                                               name = "", rNames = FALSE, cNames = TRUE,
                                               columns = NA, widths = NA,
                                               startRow = 1, startCol = 1, ...){
                           self$data <- sheetData
                           self$sheetName <- name
                           self$colNames <- cNames    # set column names before putting table into this object
                           self$rowNames <- rNames
                           self$columns <- columns    # numeric indices
                           self$widths <- widths
                           self$sheet.startRow <- startRow
                           self$sheet.startCol <- startCol
                           invisible(self)
                         },
                         addToWorkbook = function(workbook, ...){
                           super$addToWorkbook(workbook, ...)
                           writeDataTable(wb = workbook, sheet = self$sheetName, self$data,
                                          startCol = self$sheet.startCol, startRow = self$sheet.startRow,
                                          colNames = self$colNames, rowNames = self$rowNames, ...)
                           if (!identical(self$columns, NA)) {
                             setColWidths(wb = workbook ,sheet = self$sheetName,
                                          cols = self$columns, widths = self$widths,...)
                           }
                         }
                       )
)

# -----

# note : to get figure to file
# png(filename = "temp__plot.png",width = 1200, height = 800, units = "px", pointsize = 12, bg = "white")
# code to create figure
# dev.off()

RsheetImageClass = "RsheetImage"

RsheetImage <- R6Class("RsheetImage",
                       inherit = Rsheet,
                       public = list(
                         removeTheFile = FALSE,
                         picture.width = 12,
                         picture.height = 8,
                         picture.dpi = 600,
                         picture.units = "in",
                         sheet.startRow = 1,
                         sheet.startCol = 1,
                         initialize = function(figureFileName = "", nameSheet = "",
                                               removeFile = FALSE,
                                               width = 12, height = 8, dpi = 600, units = "in",
                                               startRow = 1, startCol = 1, ...){
                           super$initialize(sheetData = figureFileName, nameSheet = nameSheet, ...)
                           self$removeTheFile = removeFile
                           self$picture.width = width
                           self$picture.height = height
                           self$picture.dpi = dpi
                           self$sheet.startRow <- startRow
                           self$sheet.startCol <- startCol
                           invisible(self)
                         },
                         addToWorkbook = function(workbook, ...){
                           super$addToWorkbook(workbook, ...)
                           insertImage(wb = workbook, sheet = self$sheetName, file = self$data,
                                       width = self$picture.width, height = self$picture.height,
                                       startRow = self$sheet.startRow, startCol = self$sheet.startCol,
                                       units = self$picture.units, dpi = self$picture.dpi)
                           invisible(self)
                         },
                         removeData = function(...){  # note: only remove file when asked (after save workbook)
                           if (self$removeTheFile) {
                             nothingVariable <- file.remove(self$data)
                           }
                           invisible(self)
                         }
                       )
)

# -----

# to put more than one table/image in a single sheet of an excel file
RsheetMultiClass = "RsheetMulti"

RsheetMulti <- R6Class("RsheetMulti",
                       inherit = Rsheet,
                       public = list(
                         columns = NA,
                         widths = NA,
                         initialize = function(nameSheet = "", columns = NA, widths = NA, ...){
                           super$initialize(sheetData = list(), nameSheet = nameSheet, ...)
                           self$setColumnWidths(columns = columns, widths = widths)
                           invisible(self)
                         },
                         setColumnWidths = function(columns = NA, widths = NA){
                           # only for whole sheet, not defined per table
                           self$columns = columns
                           self$widths = widths
                           invisible(self)
                         },
                         addTable = function(tableData, rNames = FALSE, cNames = TRUE,
                                             startRow = 1, startCol = 1, ...){
                           item = list(type = "table", rowNames = rNames, colNames = cNames,
                                       sheet.startRow = startRow, sheet.startCol = startCol,
                                       data = list(tableData))
                           private$dataSheet[[self$numberOfElements+1]] <- item
                           invisible(self)
                         },
                         addImage = function(figureFileName = "", removeFile = FALSE,
                                             width = 12, height = 8, dpi = 600, units = "in",
                                             startRow = 1, startCol = 1, ...){
                           item = list(type = "image",
                                       figureFileName = figureFileName,
                                       removeFile = removeFile,
                                       picture.width = width, picture.height = height,
                                       picture.dpi = dpi, picture.units = units,
                                       sheet.startRow = startRow, sheet.startCol = startCol)
                           private$dataSheet[[self$numberOfElements+1]] <- item
                           invisible(self)
                         },
                         addToWorkbook = function(workbook, ...){
                           super$addToWorkbook(workbook, ...)
                           for (counter in 1:self$numberOfElements){
                             if (private$dataSheet[[counter]]$type == "image"){
                               insertImage(wb = workbook, sheet = self$sheetName,
                                           file = private$dataSheet[[counter]]$figureFileName,
                                           width = private$dataSheet[[counter]]$picture.width,
                                           height = private$dataSheet[[counter]]$picture.height,
                                           startRow = private$dataSheet[[counter]]$sheet.startRow,
                                           startCol = private$dataSheet[[counter]]$sheet.startCol,
                                           units = private$dataSheet[[counter]]$picture.units,
                                           dpi = private$dataSheet[[counter]]$picture.dpi, ...)
                             } else {
                               if (private$dataSheet[[counter]]$type == "table"){
                                 writeDataTable(wb = workbook, sheet = self$sheetName,
                                                private$dataSheet[[counter]]$data[[1]],
                                                startCol = private$dataSheet[[counter]]$sheet.startCol,
                                                startRow = private$dataSheet[[counter]]$sheet.startRow,
                                                colNames = private$dataSheet[[counter]]$colNames,
                                                rowNames = private$dataSheet[[counter]]$rowNames, ...)
                               }
                             }
                           }
                           if (!identical(self$columns, NA)) {   # only for whole sheet, not for individual tables
                             setColWidths(wb = workbook ,sheet = self$sheetName,
                                          cols = self$columns, widths = self$widths,...)
                             
                           }
                           invisible(self)
                         },
                         removeData = function(...){  # note: only remove image files when asked (after save workbook)
                           for (counter in 1:self$numberOfElements){
                             if (private$dataSheet[[counter]]$type == "image"){                     # only images
                               if (private$dataSheet[[counter]]$removeFile){
                                 nothingVariable <- file.remove(private$dataSheet[[counter]]$figureFileName)  # only when removeFile == TRUE for that image
                               }
                             }
                           }
                           invisible(self)
                         }
                      ),
                      active = list(
                        numberOfElements = function(value){
                          if (missing(value)){
                            return(length(private$dataSheet))
                          } else {
                            # do nothing
                          }
                        }
                      )
)

# -----

RxcelClass = "Rxcel"

Rxcel <- R6Class("Rxcel",
                 private = list(
                   nameOfFile = as.character(NA),
                   extension = ".xlsx",
                   xcelSheets = NA
                 ),
                 public = list(
                   initialize = function(nameFile = as.character(NA)){
                     self$fileName <- nameFile
                     invisible(self)
                   },
                   addSheet = function(sheetData = NA, ...){
                     if (RsheetClass %in% class(sheetData)){
                       if (self$length == 0){
                         self$excelSheets <- list(sheetData)
                       } else {
                         self$excelSheets <- append(self$excelSheets,sheetData)
                       }
                       recentSheet <- self$length # sheet that was just added
                       if (private$xcelSheets[[recentSheet]]$sheetName == ""){
                         private$xcelSheets[[recentSheet]]$sheetName <- paste("sheet",toString(recentSheet),sep = "_")
                       }
                     }
                     invisible(self)
                   },
                   createExcel = function(...){
                     tempExcel <- createWorkbook(self$fileName)
                     for (counter in 1:(self$length)){
                       private$xcelSheets[[counter]]$addToWorkbook(workbook = tempExcel, ...)
                     }
                     return(tempExcel)
                   },
                   writeExcel = function(overwrite = FALSE, ...){
                     tempExcel <- self$createExcel(...)
                     saveWorkbook(wb = tempExcel,
                                  file = paste(self$fileName, self$fileExtension, sep = ""),
                                  overwrite = overwrite)
                     # note: only remove file when asked (after save workbook)
                     for (counter in 1:(self$length)){
                       private$xcelSheets[[counter]]$removeData()
                     }
                   }
                 ),
                 active = list(
                   fileName = function(value){
                     if (missing(value)){
                       return(private$nameOfFile)
                     } else {
                       private$nameOfFile <- value
                     }
                   },
                   fileExtension = function(value){
                     if (missing(value)){
                       return(private$extension)
                     } else {
                       private$extension <- value
                     }
                   },
                   length = function(value){
                     if (missing(value)){
                       if (identical(private$xcelSheets,NA)){
                         return(0)
                       } else {
                         return(length(private$xcelSheets))
                       }
                     } else {
                       # do nothing, cannot set
                     }
                   },
                   excelSheets = function(value){
                     if (missing(value)){
                       return(private$xcelSheets)
                     } else {
                       private$xcelSheets <- value
                     }
                   }
                 )
)
