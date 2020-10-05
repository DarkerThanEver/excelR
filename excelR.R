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
                         initialize = function(sheetData = NA,
                                               name = "", rNames = FALSE, cNames = TRUE,
                                               columns = NA, widths = NA){
                           self$data <- sheetData
                           self$sheetName <- name
                           self$colNames <- cNames
                           self$rowNames <- rNames
                           self$columns <- columns
                           self$widths <- widths
                           invisible(self)
                         },
                         addToWorkbook = function(workbook, ...){
                           super$addToWorkbook(workbook, ...)
                           writeDataTable(wb = workbook, sheet = self$sheetName, self$data,
                                          colNames = self$colNames, rowNames = self$rowNames, ...)
                           if (!identical(self$columns, NA)) {
                             setColWidths(wb = workbook ,sheet = self$sheetName,
                                          cols = self$columns, widths = self$widths,...)
                           }
                         }
                       )
)

# ----

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
                        private$xcelSheets[[counter]]$addToWorkbook(tempExcel, ...)
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
