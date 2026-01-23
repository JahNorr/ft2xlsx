

#' @export
XL_Page <-
  R6::R6Class(
    classname = "XL_Page",
    inherit = XL_Mgr,

    private = list(

      filename_pvt = NULL,
      sheetname_pvt = NULL

    ),

    public = list(
      wb = NULL,
      ws = NULL,
      items = list(),

      left_margin = 0.25,
      right_margin = 0.25,
      top_margin = 0.75,
      bottom_margin = 0.75,

      initialize = function(filename = NULL, sheetname = NULL) {


        private$sheetname_pvt <- sheetname

        super$new(filename)

      },

      add_item = function(item) {

        if(!inherits(item, "XL_PageItem")) {
          message("item must be class <XL_PageItem")
        }

        nitems <- length(self$items)

        self$items[[nitems + 1]] <- item

      },

      add_sheet = function(filename = NULL, sheetname = NULL, replace = TRUE) {

        if(is.null(filename)) filename <- private$filename_pvt
        if(is.null(sheetname)) sheetname <- private$sheetname_pvt

        if(file.exists(filename)) {
          wb <- wb_load(filename)
        } else {
          wb <-  wb_workbook()
        }

        if(replace) {

          super$replace_sheet(wb, sheetname)

          wb$set_page_setup(
            sheet = sheetname,
            fitToWidth = 1,
            left = self$left_margin,
            right = self$right_margin,
            top = self$top_margin,
            bottom = self$bottom_margin,
            fitToHeight = FALSE
          )

          max_col <- 16
          max_row <- 44

          # Define the Print Area using the reserved name '_xlnm.Print_Area'
          wb$add_named_region(
            sheet = sheetname,
            name = "_xlnm.Print_Area",
            dims = rowcol_to_dims(row = 1:max_row, col = 1:max_col),
            local_sheet= TRUE  # 0 corresponds to the first sheet
          )

        } else {
          openxlsx2::wb$dims(sheet = sheetname)
          max_col <- my_dims$max_col
          max_row <- my_dims$max_row

        }

        center_style <- create_cell_style(horizontal = "center")

        # --------------  build some data  -----------------

        purrr::walk(self$items, \(item) {

          dims <- item$dims

          if(is.null(dims)) dims <- rowcol_to_dims(row = item$row, col = item$col)


          wb$add_data(sheet = sheetname, x = item$text, dims = dims)

          wb$add_font(dims = dims, name = item$style$name, color = wb_colour(item$style$color),
                      size = item$style$size,
                      bold = item$style$bold, underline = item$style$underline,
                      vert_align = item$style$vert_align
          )

          if(item$merge_across) {
            rc <- dims_to_rowcol(x = dims, as_integer = T)
            rc$col <- 1:max_col
            mrg_dims <- rowcol_to_dims(rc$row, rc$col)

            wb$merge_cells(sheet = sheetname, dims = mrg_dims)

            wb$add_cell_style(
              sheet = sheetname,
              dims = mrg_dims,
              horizontal = "center"
            )
          }

        })

        # ------------------- do the merge if needed  -----------------------------


        wb$save(filename, overwrite = TRUE)

      },

      ws_dims = function(wb) {

        sd <- wb$worksheets[[1]]$sheet_data

        val_dims <- sd$cc  %>%
          filter(is != "") %>%
          mutate(t = gsub(".*<t>(.*?)<.*","\\1",is)) %>%
          mutate(tok = !grepl("<is>", t, fixed = T)) %>%
          filter(tok) %>%
          pull(r)  %>% purrr::map(~as.data.frame(dims_to_rowcol(.x, as_integer = T))) %>% bind_rows() %>%
          summarise(mincol = min(col), minrow = min(row), maxcol = max(col), maxrow = max(row)) %>%
          as.list()

        val_dims
      }

    ),

    active = list(

    )
  )

#' @export
XL_PageItem <-  R6::R6Class(
  classname = "XL_PageItem",

  private = list(
  ),

  public = list(
    dims = NULL,
    text = "",
    style = NULL,
    row = NULL,
    col = NULL,
    merge_across = FALSE,

    initialize = function(dims = NULL,
                          text = "",
                          style = NULL,
                          row = NULL,
                          col = NULL,
                          merge_across = FALSE) {

      #args <- match.call()[-1]
      arg_names <- names(formals())
      args <- mget(arg_names, envir = sys.frame(sys.nframe()))

      purrr::imap(args, \(val, nm) {

        self[[nm]] <- val
      })


    }
  ),

  active = list(

  )
)

# b = b, color = wb_color(color), i = i, name = name, charset = "1",
# sz = as.character(as.integer(sz)), u = u,
# vert_align = vert_align, scheme = "none"

#' @export
XL_Styling <-  R6::R6Class(
  classname = "XL_Styling",

  private = list(

    bold_pvt = "",
    italic_pvt = "",
    outline_pvt = "",
    strike_pvt = "",
    name_pvt = "Aptos Narrow",
    size_pvt = "11",
    color_pvt = "",
    underline_pvt = "",
    vert_align_pvt = "",
    shadow_pvt = "",
    charset_pvt = "",
    condense_pvt = "",
    extend_pvt = "",
    family_pvt = "",
    scheme_pvt = "",
    horizontal_pvt = NULL

  ),

  public = list(
    initialize = function(name = NULL,
                          size = NULL,
                          color = NULL,
                          bold = NULL,
                          italic = NULL,
                          outline = NULL,
                          strike = NULL,
                          underline = NULL,
                          vert_align = NULL,
                          shadow = NULL,
                          charset = NULL,
                          condense = NULL,
                          extend = NULL,
                          family = NULL,
                          scheme = NULL,
                          horizontal = NULL) {


      args <- match.call()[-1]


      purrr::imap(args, \(val, nm) {

        nm <- paste0(nm ,"_pvt")
        private[[nm]] <- val
      })

    }

  ),

  active = list(

    bold = function(value) {

      item <- paste0("bold","_pvt")

      if(missing(value)) return(private[[item]])

      private[[item]] <- value
    },

    italic = function(value) {

      item <- paste0("italic","_pvt")

      if(missing(value)) return(private[[item]])

      private[[item]] <- value
    },

    size = function(value) {

      item <- paste0("size","_pvt")
      if(missing(value)) return(private[[item]])
      private[[item]] <- value
    },

    name = function(value) {
      item <- paste0("name","_pvt")
      if(missing(value)) return(private[[item]])
      private[[item]] <- value
    },

    color = function(value) {
      item <- paste0("color","_pvt")
      if(missing(value)) return(private[[item]])
      private[[item]] <- value
    },

    underline = function(value) {
      item <- paste0("underline","_pvt")
      if(missing(value)) return(private[[item]])

      if(!(inherits(value, "chjaracter") && value %in% c("single" , "double", ""))) {

        message("value must be one of single, double, or empty string")
      }

      private[[item]] <- value
    },

    vert_align = function(value) {
      item <- paste0("vert_align","_pvt")
      if(missing(value)) return(private[[item]])
      private[[item]] <- value
    },

    horizontal = function(value) {
      item <- paste0("horizontal","_pvt")
      if(missing(value)) return(private[[item]])
      private[[item]] <- value
    },

    outline = function(value) {
      item <- paste0("outline","_pvt")
      if(missing(value)) return(private[[item]])
      private[[item]] <- value
    },

    strike = function(value) {
      item <- paste0("strike","_pvt")
      if(missing(value)) return(private[[item]])
      private[[item]] <- value
    },

    shadow = function(value) {
      item <- paste0("strike","_pvt")
      if(missing(value)) return(private[[item]])
      private[[item]] <- value
    },

    charset = function(value) {
      item <- paste0("strike","_pvt")
      if(missing(value)) return(private[[item]])
      private[[item]] <- value
    },

    condense =function(value) {
      item <- paste0("strike","_pvt")
      if(missing(value)) return(private[[item]])
      private[[item]] <- value
    },

    extend = function(value) {
      item <- paste0("strike","_pvt")
      if(missing(value)) return(private[[item]])
      private[[item]] <- value
    },

    family = function(value) {
      item <- paste0("strike","_pvt")
      if(missing(value)) return(private[[item]])
      private[[item]] <- value
    },

    scheme = function(value) {
      item <- paste0("strike","_pvt")
      if(missing(value)) return(private[[item]])
      private[[item]] <- value
    }



  )
)

