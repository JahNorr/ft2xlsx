
XLSX_XML_Mgr <-
  R6::R6Class(
    classname = "XLSX_XML_Mgr",

    private = list(
      my_dir = NULL,
      xl_dir = NULL,
      workbook_file = NULL,
      styles_file = NULL,
      worksheets_dir = NULL,
      theme_dir = NULL,

      styles_xml_doc = NULL
    ),

    public = list(

      initialize = function(file) {

        dir <- tempdir()

        test_dir <- paste0(dir,"\\xlsx_to_xml")
        if(!dir.exists(test_dir)) dir.create(test_dir)

        file.copy(from = file, to = test_dir)

        file_xlsx <- paste0(test_dir,"\\",basename(file))
        file_zip <- gsub("xlsx$", "zip", file_xlsx)

        file.rename(from = file_xlsx, file_zip)

        unzip(file_zip, overwrite = TRUE, exdir = test_dir)
        file.remove(file_zip)

        private$my_dir <- test_dir

        private$xl_dir <- paste0(test_dir, "\\xl")
        private$workbook_file <- paste0(private$xl_dir, "\\workbook.xml")
        private$styles_file <- paste0(private$xl_dir, "\\styles.xml")
        private$worksheets_dir <-  paste0(private$xl_dir, "\\worksheets")
        private$theme_dir <- paste0(private$xl_dir, "\\theme")

        private$styles_xml_doc <- xml2::read_xml(x = private$styles_file)

      }


    ),

    active  = list(

      styles = function(value) {

        if(!missing(value)) {
          stop("styles is readonly")
        } else {
          return(private$styles_xml_doc)
        }
      }
    )
  )
