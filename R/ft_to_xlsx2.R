
library(openxlsx2)
library(dplyr)
library(flextable)

#' Convert Flextable Object to .xlsx File
#'
#' @param ft - flextable object
#' @param file - output file name
#' @param sheet_name  - sheet name
#' @param start_row - starting row on sheet
#' @param start_col - starting column on sheet
#' @param style - apply styling
#' @param append - append this to an already existing file
#' @param verbose - add some messages (via cat) to show progress
#'
#' @returns - NULL
#'
#' @import dplyr
#' @import openxlsx
#' @import flextable
#' @import officer
#'
#' @export
#'
#' @examples
#' library(flextable)
#' library(officer)
#'
#' ft <- mtcars %>%
#'   head(10) %>%
#'   select(1:7) %>%
#'   tibble::rownames_to_column(var = "Make") %>%
#'   flextable() %>%
#'   add_header_lines(c("ft_to_xlsx Example", "mtcars")) %>%
#'   align(align = "center", part = "header") %>%
#'   hline(i = 1, border = fp_border(width = 0), part = "header") %>%
#'   vline(j = 3:4, border = fp_border_default()) %>%
#'   hline(i = 6:7, border = fp_border(width = 1), part = "body") %>%
#'   bg(j = 4, bg = "#bbddcc") %>%
#'   bold(i = 4, bold = TRUE) %>%
#'   italic( i = 7, italic = TRUE) %>%
#'   color(i = 4, color = "#334499") %>%
#'   width(j = "Make", width = 1.5) %>%
#'   padding(padding.bottom = 1, padding.top = 2)
#'
#' ft %>% ft_to_xlsx2()
#'
ft_to_xlsx2 <- function(ft, file = "ft_to_xlsx.xlsx",
                        sheet_name = "ft1",
                        start_row = 1,
                        start_col = 1,
                        style = TRUE,
                        append = FALSE,
                        verbose = FALSE) {


  verbose = TRUE

  if(verbose) cat("-----------------------------------\n","
                   ... in ft_to_xlsx2\n")

  # Load or create workbook
  wb <- if (append && file.exists(file)) wb_load(file) else wb_workbook()

  # ===========================================
  # Add sheet

  wb$add_worksheet(sheet_name)

  ws <- sheet_name

  # get row offsets to body from header size

  start_header_row <- start_row
  start_body_row   <- start_header_row + ft$header$content$nrow
  start_footer_row <- start_body_row + ft$body$content$nrow

  ncols <- ft$body$content$ncol

  # Write header and body data

  df_header <- ft$header$dataset %>% select(all_of(ft$header$col_keys))
  df_body   <- ft$body$dataset   %>% select(all_of(ft$body$col_keys))

  if(verbose) cat(" ... adding header data\n")

  # --------------------------------------------------------------------------
  wb$add_data(ws, df_header, startRow = start_header_row,
              startCol = start_col, colNames = FALSE)

  if(verbose) cat(" ... adding body data\n")

  wb$add_data(ws, df_body,   startRow = start_body_row,
              startCol = start_col, colNames = FALSE)


  if(verbose) cat(" ... doing xlheader\n")
  wb <- wb %>% xl_header(ws = ws, data = ft$header$content$data,
                         start_row = start_header_row, start_col = start_col)


  # =========== determine if footer contains any information ===========
  #                  if there are footnotes the footer will
  #                  have empty data and the footnotes
  #                        are stored elsewhere

  ftr_len <- length(ft$footer$dataset %>% unlist())
  empty_len <- ft$footer$dataset %>% {length(which(.==""))}

  if(ftr_len != empty_len) {

    df_footer   <- ft$footer$dataset   %>% select(all_of(ft$footer$col_keys))

    if(verbose) cat(" ... adding footer data\n")
    wb$add_data(ws, df_footer,   startRow = start_footer_row,
                startCol = start_col, colNames = FALSE)

  } else {

    if(verbose) cat(" ... doing footnote\n")
    wb <- xl_footnotes(wb, ws, ft$footer$content$data, start_footer_row)
  }

  # Apply merges
  if(verbose) cat(" ... doing merging\n")

  wb <- xl_merge_from_spans2(wb, ws, ft$header$spans,
                             start_row = start_header_row, start_col = start_col)

  wb <- xl_merge_from_spans2(wb, ws, ft$body$spans,
                             start_row = start_body_row,   start_col = start_col)

  # Apply styling

  if(style) {

    if(verbose) cat(" ... doing styling\n")

    # ====================  test xl_styling3 =================================

    xl_styling3(wb, ws, part = "header",part_data = ft$header,
                      start_row = start_header_row, start_col = start_col)

    # ========================================================================

    if(verbose) cat(" ... doing header styling\n")

    wb <- xl_styling2(wb, ws, part = "header",part_data = ft$header,
                      start_row = start_header_row, start_col = start_col)

    if(verbose) cat(" ... doing body styling\n")
    wb <- xl_styling2(wb, ws, part = "body", part_data = ft$body,
                      start_row = start_body_row,   start_col = start_col)

    if(verbose) cat(" ... doing footer styling\n")
    wb <- xl_styling2(wb, ws, part = "footer", part_data = ft$body,
                      start_row = start_body_row,   start_col = start_col)
  }
  #
  # Column widths

  if(verbose) cat(" ... setting column widths\n")
  wb$set_col_widths(ws, cols = 1:ncols, widths = ft$header$colwidths * 12)

  # TODO: Footnotes (use ft$xl_footnotes or get_xl_footnotes() as before)

  # Save workbook

  if(verbose) cat(" ... saving\n")

  wb$save(file, overwrite = TRUE)
}


xl_header <- function(wb, ws, data, start_row = 1, start_col = 1) {

  nrows <- nrow(data)
  ncols <- ncol(data)

  irow <- start_row
  row_offset = start_row - 1
  col_offset = start_col - 1

  # for each row and col of data

  purrr::map(1:nrows, function(irow) {
    purrr::map(1:ncols, function( icol) {

      # ---- get the dims

      dims <- rowcol_to_dims(irow + row_offset, icol + col_offset)


      purrr::map(data[irow, icol], function(ftnt_lines) {

        x <- fmt_txt("")

        purrr::pmap(ftnt_lines, function(txt, font.size, italic, bold,
                                         underlined, color, shading.color, font.family,
                                         hansi.family, eastasia.family, cs.family,
                                         vertical.align, width, height,  url,
                                         eq_data, word_field_data,
                                         img_data, .chunk_index) {


          x <<- x + fmt_txt(txt,
                            bold = ifelse(is.na(bold), FALSE, bold),
                            #color = ifelse(is.na(color), wb_color("black"), wb_color(color)),
                            size = ifelse(is.na(font.size), 11, font.size),
                            vert_align = ifelse(is.na(vertical.align), "baseline",
                                                vertical.align))

        })
        wb <<- wb_add_data(wb = wb, sheet = ws, x = x, dims = dims)
      })
    })
  })

  wb
}

