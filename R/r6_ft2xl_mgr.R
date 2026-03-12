library(R6)


#------------------------------------------------------------------------------
# XL_Mgr
#
# Manager class for Excel workbooks.
#
# Handles:
#   - opening a workbook
#   - moving sheets
#   - renamingsheets
#   - saving a workbook
#------------------------------------------------------------------------------

#' @export
XL_Mgr <- R6Class(
  classname = "XL_Mgr",

  private = list(

    # Path to target Excel file
    filename_pvt = NULL,

    # Default sheet name (currently not heavily used)
    sheetname_pvt = NULL,

    # Default sheet name (currently not heavily used)
    wb_pvt = NULL,

    col2hex = function(x) {
      grDevices::rgb(t(grDevices::col2rgb(x)), maxColorValue = 255)
    }
  ),

  public = list(

    initialize = function(filename, append = FALSE) {

      if(append) {
        if (!file.exists(filename)) {
          warning("You must specify a file name for an existing file to append")
          return(FALSE)
        }

        private$wb_pvt <- wb_load(filename)
      } else {
        private$wb_pvt <- wb_workbook()
      }
      # Load workbook and metadata
      private$filename_pvt <- filename

    },

    save = function(overwrite = TRUE) {

      private$wb_pvt$save(private$filename_pvt, overwrite = overwrite)
    },

    #--------------------------------------------------------------------------
    # replace_sheet
    #
    # Creates a  worksheet, removing the old sheet if there is one
    #--------------------------------------------------------------------------
    replace_sheet = function(wb, sheetname) {

      if(missing(sheetname)) sheetname <- private$sheetname_pvt
      if (sheetname %in% wb$sheet_names) {
        wb$remove_worksheet(sheetname)
      }
      wb$add_worksheet(sheetname)
    },

    #--------------------------------------------------------------------------
    # rename_sheet
    #
    # Renames a  worksheet
    #--------------------------------------------------------------------------
    rename_sheet = function(...) {

      wb <- private$wb_pvt
      quos <- enquos(...)
      q <- quos[1]

      to <- names(q)

      from <- q[[1]] %>% as_label() %>% gsub("\"","", .)
      index <- which(wb$sheet_names == from)

      wb$set_sheet_names(old = index, new = to)
      wb$save(private$filename_pvt, overwrite = TRUE)



    },

    #--------------------------------------------------------------------------
    # move_sheet
    #
    # Reorders worksheets in an existing Excel file.
    # Supports moving by sheet name or index.
    #--------------------------------------------------------------------------
    move_sheet = function(from, to, .after, .before) {

      # Validate required arguments
      if (missing(from)) {
        warning("You must specify <from> as an integer index or by name")
        return(FALSE)
      }

      if (!file.exists(private$filename_pvt)) {
        warning("You must specify a file name for an existing file")
        return(FALSE)
      }

      # Load workbook and metadata
      wb <- wb_load(private$filename_pvt)
      sheets <- wb$sheet_names
      nsheets <- length(wb$worksheets)

      # Resolve source sheet index
      if (is.character(from)) {
        ifrom <- which(sheets == from)
      } else if (is.numeric(from)) {
        ifrom <- from

        if (!between(ifrom, 1, nsheets)) {
          warning(paste0("<from> must be between 1 and ", nsheets))
          return(FALSE)
        }
      } else {
        warning("<from> must be a valid sheet index or name")
        return(FALSE)
      }

      # Resolve destination index
      if (!missing(to)) {

        if (is.character(to)) {
          ito <- which(sheets == to)
        } else if (is.numeric(to)) {
          ito <- to

          if (!between(ito, 1, nsheets)) {
            warning(paste0("<to> must be between 1 and ", nsheets))
            return(FALSE)
          }
        } else {
          warning("<to> must be a valid sheet index or name")
          return(FALSE)
        }

      } else if (!missing(.before)) {

        if (is.character(.before)) {
          .before <- which(sheets == .before)

          if(length(.before) == 0)  {
            warning("<.before> must be a valid sheet index or name")
            return(FALSE)
          }

        } else if (is.numeric(.before)) {
          ito <- .before
        }

        ito <- .before - (ifrom < .before) * 1

        if (!between(ito, 1, nsheets)) {
          warning(paste0("<.before> must be between 1 and ", nsheets))
          return(FALSE)
        }

        # end if handle .before
        # ==========================================

      } else if (!missing(.after)) {

        if (is.character(.after)) {
          .after <- which(sheets == .after)

          if(length(.after) == 0)  {
            warning("<.after> must be a valid sheet index or name")
            return(FALSE)
          }

        } else if (is.numeric(.after)) {
          ito <- .after
        }

        ito <- .after + (ifrom > .after) * 1

        if (!between(ito, 1, nsheets)) {
          warning(paste0("<.after> must be between 1 and ", nsheets))
          return(FALSE)
        }


        # end if handle .after
        # ==========================================

      }

      # Build new sheet order
      isheets <- setdiff(seq_len(nsheets), ifrom)
      bef <- head(isheets, ito - 1)
      aft <- tail(isheets, nsheets - ito)

      order <- c(bef, ifrom, aft)

      # Apply and save
      wb$set_order(order)
      self$save()
    }
  ),

  active = list(


    #--------------------------------------------------------------------------
    # wb
    #
    # Getter / setter for Excel workbook (from openxlsx2::)
    #--------------------------------------------------------------------------
    wb = function(value) {

      if (missing(value)) return(private$wb_pvt)

      if(!inherits(wb, "wbWorkbook")) {
        message("<wb> must be a 'wbWorkbook' object ")
        return(NULL)
      }

      private$wb_pvt <- value

    },

    #--------------------------------------------------------------------------
    # filename
    #
    # Getter / setter for Excel filename.
    # Setting a new filename automatically resets append behavior.
    #--------------------------------------------------------------------------
    filename = function(value) {

      if (missing(value)) return(private$filename_pvt)

      private$filename_pvt <- value

    },

    #--------------------------------------------------------------------------
    # sheetname
    #
    # Getter / setter for Excel sheetname
    #--------------------------------------------------------------------------

    sheetnames = function(value) {

      if(!missing(value))  {

        message("this property is read-only")
      }

      wb <- private$wb_pvt
      wb$sheet_names


    }

  )

)

#------------------------------------------------------------------------------
# FT2XL_Mgr
#
# Manager class for exporting flextable objects to Excel workbooks.
# Handles:
#   - tracking output filename
#   - append vs overwrite behavior
#   - adding sheets incrementally
#   - reordering sheets in an existing workbook
#------------------------------------------------------------------------------

#' @export
FT2XL_Mgr <- R6Class(
  classname = "FT2XL_Mgr",
  inherit = XL_Mgr,

  private = list(

    # Optional upstream flextable manager (used by subclasses)
    ft_mgr = NULL,

    # Starting position for table insertion
    start_row = 1,
    start_col = 1,

    # margins
    top_margin_pvt = 1,
    left_margin_pvt = 0.75,
    right_margin_pvt = 0.75,
    bottom_margin_pvt = 0.75,

    # Whether styling should be applied (reserved for future use)
    style = TRUE,

    # Internal append flag:
    #   FALSE → overwrite file
    #   TRUE  → append new sheets
    append_pvt = FALSE,

    # Verbose logging flag
    verbose_pvt = TRUE
  ),

  public = list(

    #--------------------------------------------------------------------------
    # Constructor
    #--------------------------------------------------------------------------
    initialize = function(filename = NULL, verbose = FALSE) {

      # Filename can be provided at initialization or later via active binding
      super$filename  <-  filename
      private$verbose_pvt <- verbose

    },

    #--------------------------------------------------------------------------
    # Add a single worksheet to the Excel file
    #
    # This is a wrapper around ft2xlsx::ft_to_xlsx2()
    #--------------------------------------------------------------------------
    add_sheet = function(ft = NULL, sheetname = NULL, index = NULL,
                         append = NULL, tab_color = NULL) {

      # Fallback to default sheet name if not provided
      if (is.null(sheetname)) sheetname <- super$sheetname

      # If append is not explicitly provided, use internal state
      if (is.null(append)) append <- private$append_pvt

      if(!is.null(tab_color)) {
        if(!grepl("^[0-9A-Fa-f]{6}$", tab_color)) {
          if(!grepl("^#[0-9A-Fa-f]{6}$", tab_color)) {

            tab_color <- private$col2hex(tab_color)

          }
          tab_color <- gsub("^.","", tab_color)
        }

      }

      # Write flextable to Excel
      ft_to_xlsx2(
        ft            = ft,
        file          = super$filename,
        sheet_name    = sheetname,
        index         = index,
        left_margin   = private$left_margin_pvt,
        right_margin  = private$right_margin_pvt,
        top_margin    = private$top_margin_pvt,
        bottom_margin = private$bottom_margin_pvt,
        start_row     = private$start_row,
        start_col     = private$start_col,
        append        = append,
        tab_color     = tab_color,
        verbose = private$verbose_pvt
      )

      # Once a sheet has been written, all subsequent writes should append
      private$append_pvt <- TRUE
    }
          ),

          active = list(

            #--------------------------------------------------------------------------
            # filename
            #
            # Getter / setter for Excel filename.
            # Setting a new filename automatically resets append behavior.
            #--------------------------------------------------------------------------
            filename = function(value) {

              if (missing(value)) return(super$filename)

              super$filename <- value

              # New file → do not append until first write
              private$append_pvt <- FALSE
            },


            #--------------------------------------------------------------------------
            # append
            #
            # Controls whether sheets are appended to an existing file.
            #--------------------------------------------------------------------------
            append = function(value) {

              if (missing(value)) return(private$append_pvt)

              if (!inherits(value, "logical")) {
                warning("append must be logical")
                return(NULL)
              }

              private$append_pvt <- value
            },

            #--------------------------------------------------------------------------
            # left_margin
            #
            # sets a printing margin
            #--------------------------------------------------------------------------
            left_margin = function(value) {

              if (missing(value)) return(private$left_margin_pvt)

              if (!inherits(value, "numeric")) {
                warning("append must be numeric")
                return(NULL)
              }

              private$left_margin_pvt <- value


            },

            #--------------------------------------------------------------------------
            # right_margin
            #
            # sets a printing margin
            #--------------------------------------------------------------------------
            right_margin = function(value) {

              if (missing(value)) return(private$right_margin_pvt)

              if (!inherits(value, "numeric")) {
                warning("append must be numeric")
                return(NULL)
              }

              private$right_margin_pvt <- value


            },

            #--------------------------------------------------------------------------
            # top_margin
            #
            # sets a printing margin
            #--------------------------------------------------------------------------
            top_margin = function(value) {

              if (missing(value)) return(private$top_margin_pvt)

              if (!inherits(value, "numeric")) {
                warning("append must be numeric")
                return(NULL)
              }

              private$top_margin_pvt <- value


            },

            #--------------------------------------------------------------------------
            # bottom_margin
            #
            # sets a printing margin
            #--------------------------------------------------------------------------
            bottom_margin = function(value) {

              if (missing(value)) return(private$bottom_margin_pvt)

              if (!inherits(value, "numeric")) {
                warning("append must be numeric")
                return(NULL)
              }

              private$bottom_margin_pvt <- value


            }
          )
        )

        #------------------------------------------------------------------------------
        # FT2XL_Stats_Mgr
        #
        # Extension of FT2XL_Mgr specialized for statistical outputs.
        # Integrates with an FT_StatsMgr to generate flextables dynamically
        # from "columns of interest" (COIs).
        #------------------------------------------------------------------------------

        #' @export
        FT2XL_Stats_Mgr <- R6Class(
          classname = "FT2XL_Stats_Mgr",
          inherit   = FT2XL_Mgr,

          private = list(
            # Inherited ft_mgr is expected to be an FT_StatsMgr

          ),

          public = list(

            initialize = function(ft_stats_mgr = NULL, ...) {

              private$ft_mgr <- ft_stats_mgr

              super$initialize(...)


            },

            #--------------------------------------------------------------------------
            # add_sheet
            #
            # Adds a sheet using either:
            #   - an explicitly supplied flextable, or
            #   - a COI resolved via the FT_StatsMgr
            #--------------------------------------------------------------------------
            add_sheet = function(ft = NULL, coi = NULL, sheetname = NULL,
                                 index = NULL, append = NULL, tab_color = NULL) {

              # Default sheet name logic
              if (is.null(sheetname)) sheetname <- private$sheetname
              if (is.null(sheetname) && !is.null(coi)) sheetname <- coi

              #if (is.null(sheetname)) browser()

              if (is.null(append)) append <- private$append_pvt

              # Lazily generate flextable if not provided
              if (is.null(ft)) {
                ft <- private$ft_mgr$ft(coi = coi)
              }

              # Delegate to parent class
              super$add_sheet(
                ft        = ft,
                sheetname = sheetname,
                index     = index,
                append    = append,
                tab_color = tab_color
              )
            },

            #--------------------------------------------------------------------------
            # add_sheets
            #
            # Batch version of add_sheet.
            # Generates flextables from COIs if not explicitly provided.
            #--------------------------------------------------------------------------
            add_sheets = function(fts = NULL, cois = NULL,
                                  sheetnames = NULL, append = NULL, tab_color = NULL) {


              # Default sheet names map to COIs
              if (is.null(sheetnames)) sheetnames <- cois

              # Generate flextables if missing
              if (is.null(fts)) {
                fts <- purrr::map(cois, \(coi) {
                  private$ft_mgr$ft(coi = coi)
                })
              }

              nfts <- length(fts)

              # Append logic:
              #   first sheet → current append state
              #   remaining   → TRUE
              if (is.null(append)) {
                append <- private$append_pvt
              }

              append <- c(append, rep(TRUE, nfts - 1))
              # Sequential write
              purrr::walk(seq_len(nfts), \(ift) {
                cat(cois[ift], "\n")

                super$add_sheet(
                  ft        = fts[[ift]],
                  sheetname = sheetnames[ift],
                  append    = append[ift],
                  tab_color = tab_color
                )
              })
            },

            #--------------------------------------------------------------------------
            # add_section
            #
            # High-level helper for exporting structured report sections.
            # Relies heavily on external layout metadata and global context.
            #--------------------------------------------------------------------------
            add_section = function(type = "Core",
                                   section = NULL,
                                   sect_num = NULL,
                                   filename = NULL) {

              # At least one section selector must be provided
              if (is.null(sect_num) && is.null(section)) return(FALSE)

              # Load layout metadata
              lo_mgr <- Layout_Mgr$new(type = "data")

              df_lo <- lo_mgr$layout() %>%
                filter(sect_type == {{ type }}) %>%
                select(sect_num, section) %>%
                distinct() %>%
                arrange(sect_num)

              # Filter by section number or name
              if (!is.null(sect_num)) {
                df_lo <- df_lo %>% filter(sect_num == {{ sect_num }})
              } else if (!is.null(section)) {
                df_lo <- df_lo %>% filter(grepl({{ section }}, section))
              }

              # Iterate through sections and export each
              purrr::pwalk(df_sections, \(sect_num, section) {

                filename <- paste0(
                  "annual_", year, "_core_",
                  sect_num, "_", section, ".xlsx"
                ) %>%
                  gsub("[ ]", "_", .) %>%
                  gsub("/", "_", ., fixed = TRUE) %>%
                  paste0(dir, .)

                cat(filename, "\n")

                xl_mgr$filename <- filename

                cois <- df_lo_core %>%
                  filter(sect_num == {{ sect_num }}) %>%
                  pull(col_name)

                purrr::walk(cois, \(coi) {
                  xl_mgr$add_sheet(coi = coi)
                })
              })
            }
          ),

          active = list(

            #--------------------------------------------------------------------------
            # append (redeclared for clarity)
            #--------------------------------------------------------------------------
            append = function(value) {

              if (missing(value)) return(private$append_pvt)

              if (!inherits(value, "logical")) {
                warning("append must be logical")
                return(NULL)
              }

              private$append_pvt <- value
            },

            verbose = function(value) {

              if (missing(value)) return(private$verbose_pvt)

              if (!inherits(value, "logical")) {
                warning("verbose must be logical")
                return(NULL)
              }

              private$verbose_pvt <- value
            },

            #--------------------------------------------------------------------------
            # ft_stats_mgr
            #
            # Getter / setter for the FT_StatsMgr dependency.
            #--------------------------------------------------------------------------
            ft_stats_mgr = function(mgr) {

              if (missing(mgr)) return(private$ft_mgr)

              if (!inherits(mgr, "FT_StatsMgr")) {
                warning("mgr must be class FT_StatsMgr")
                return(NULL)
              }

              private$ft_mgr <- mgr
            }
          )
        )
