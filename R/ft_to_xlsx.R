
library(openxlsx)
library(dplyr)


#' Convert Flextable Object to .xlsx File
#'
#' @param ft - flextable object
#' @param file - output file name
#' @param sheet_name  - sheet name
#' @param start_row - starting row on sheet
#' @param start_col - starting column on sheet
#' @param append - append this to an already existing file
#'
#' @returns - NULL
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
#' ft %>% ft_to_xlsx()
#'

ft_to_xlsx <- function(ft, file = "ft_to_xlsx.xlsx",
                       sheet_name ="ft1",
                       start_row = 1,
                       start_col = 1,
                       append = FALSE) {



  sheet_col_offset <- start_col - 1
  sheet_row_offset <- start_row - 1

  if(append) {
    wb <- openxlsx::loadWorkbook(file = file)
  } else {
    wb <- openxlsx::createWorkbook()
  }

  # ===========================================
  #   add sheet

  ws = wb %>% addWorksheet(sheetName = sheet_name)

  # get row offsets to body from header size

  start_header_row <-  start_row
  start_body_row <-  start_header_row + ft$header$content$nrow
  start_footer_row <-  start_body_row + ft$body$content$nrow

  ncols <- ft$body$content$ncol

  # add datasets

  df_header <- ft$header$dataset %>% select(all_of(ft$header$col_keys))
  df_body <- ft$body$dataset %>% select(all_of(ft$body$col_keys))

  wb %>% writeData(ws, df_header,
                   startCol = start_col,
                   startRow = start_header_row,
                   colNames = FALSE)

  wb %>% writeData(ws, df_body,
                   startCol = start_col,
                   startRow = start_body_row,
                   colNames = FALSE)

  # ===========================================
  #
  # *************** HEADER ********************

  # format the merge spans


  wb %>% xl_merge_from_spans(ws, ft$header$spans,
                             start_row = start_header_row,
                             start_col = start_col)

  wb %>% xl_styling(ws,
                    part = ft$header,
                    start_row = start_header_row,
                    start_col = start_col)

  # ===========================================
  #
  # *************** BODY ********************

  # format the merge spans

  # row_spans <- ft$body$spans$rows
  # col_spans <- ft$body$spans$columns

  wb %>% xl_merge_from_spans(ws, ft$body$spans,
                             start_row = start_body_row,
                             start_col = start_col)

  wb %>% xl_styling(ws,
                    part = ft$body,
                    start_row = start_body_row,
                    start_col = start_col)


  wb %>% setColWidths(ws, cols = 1:ncols, widths = ft$header$colwidths * 12) # "auto")

  # ===========================================
  #
  # *************** FOOTER ********************


  ftr_content <- ft$footer$content$data
  nlines <- ftr_content %>% dim() %>% {.[1]}

  if(nlines > 0) {

    sapply(1:nlines, function(iline) {

      line_info <- ftr_content[[iline]]

      nparts <- line_info %>% nrow()

      sapply(1:nparts, function(ipart) {

        #browser()


      })

    })

  }

  # ================================================
  #
  #       save what we have

  wb %>% saveWorkbook(file = file, overwrite = TRUE)

}

xl_styling <- function(wb,ws, part, start_row = 1, start_col = 1) {

  nrows <- part$content$nrow
  ncols <- part$content$ncol

  row_offset <- start_row - 1
  col_offset <- start_col - 1

  cell_styles = part$styles$cells
  par_styles = part$styles$pars
  text_styles = part$styles$text

  sapply(1:nrows, function(r) {
    sapply(1:ncols, function(c) {

      fill_val <- cell_styles$background.color$data[r,c]
      if(fill_val == "transparent") fill_val <- NULL

      border_data <- border_data(styles = cell_styles, row = r, col = c)
      border_widths <- border_data$width
      border_styles <- border_data$style %>%
        gsub("solid","thin",.) %>%
        {.[border_widths == 0] <- "none";.} %>%
        {.[border_widths > 0.7] <- "medium";.} %>%
        {.[border_widths > 1] <- "thick";.} %>%
        {.[border_data$color == "transparent"] <- "none";.}

      border_colors <- border_data$color %>%
        gsub("transparent","white",.)

      valign <- cell_styles$vertical.align$data[r,c]
      halign <- par_styles$text.align$data[r,c]
      color <- text_styles$color$data[r,c]
      font_family <- text_styles$font.family$data[r,c] %>% unname()
      font_size <- text_styles$font.size$data[r,c] %>% unname()

      decorations <- ft_text_decoration(text_styles = text_styles,r,c)

      if(decorations$bold) bold <- "bold" else bold <- ""
      if(decorations$italic) italic <- "italic" else italic <- ""
      if(decorations$underline) underline <- "underline" else underline <- ""

      text_dec <- c(bold, italic, underline)

      style <- createStyle(fgFill = fill_val,
                           fontName = font_family,
                           fontSize = font_size,
                           fontColour = color,
                           border = borders(),
                           borderStyle = border_styles,
                           borderColour = border_colors,
                           valign = valign, halign = halign,
                           textDecoration = text_dec
                           #bold = bold, italic = italic, underline = underline
      )

      style_col <- c + col_offset
      style_row <- r + row_offset
      wb %>% addStyle(ws, style = style, rows = style_row, cols = style_col)

    })
  })

  wb

}

ft_text_decoration <- function(text_styles, row, col) {


  bold <- text_styles$bold$data[row,col] %>% unname()
  italic <- text_styles$italic$data[row,col] %>% unname()
  underline <- text_styles$underlined$data[row,col] %>% unname()

  list(bold = bold, italic = italic, underline = underline)
}

borders <- function() {c("top", "bottom", "left", "right")}

border_element <- function(styles, type,  row, col) {

  bords <- borders()

  x <- sapply(bords, function(bord) {
    el <- paste0("border.",type,".", bord)

    styles[[el]]$data[row,col]

  })
  names(x) <- bords
  x
}

border_data <- function(styles, row, col) {

  types <- styles %>% names() %>% {.[grepl("^border[.]",.)]} %>%
    gsub("border.","",.) %>% gsub("(.*)[.].*","\\1",.) %>% unique()

  x <- lapply(types, function(type) {

    border_element(styles, type , row , col )

  })
  names(x) <- types
  x
}

padding_data <- function(styles, row, col) {

  bords <- borders()

  x <- sapply(bords, function(bord) {
    el <- paste0("padding.", bord)

    styles[[el]]$data[row,col]

  })

  names(x) <- bords
  x
}

margin_data <- function(styles, row, col) {

  bords <- borders()

  x <- sapply(bords, function(bord) {
    el <- paste0("margin.", bord)

    styles[[el]]$data[row,col]

  })

  names(x) <- bords
  x
}

ft_style_elements <- function(ft, type = "cells", part = "body") {

  ft[[part]]$styles[[type]] %>% names()

}

#' @keywords internal
xl_merge_from_spans <- function(wb, ws, spans,
                                start_row = 1,
                                start_col = 1) {



  row_spans <- spans$rows
  col_spans <- spans$columns

  wb %>% xl_merge_from_rowspans(ws, row_spans,
                                start_row = start_row,
                                start_col = start_col)

  wb %>% xl_merge_from_colspans(ws, col_spans,
                                start_row = start_row,
                                start_col = start_col)
  wb
}


xl_merge_from_rowspans <- function(wb,ws,spans, start_row = 1, start_col = 1) {

  nrows <- nrow(spans)
  ncols <- ncol(spans)

  row_offset <- start_row - 1
  col_offset <- start_col - 1

  sapply(1:nrows, function(r) {
    sapply(1:ncols, function(c) {
      span_val <- spans[r,c]
      if(span_val > 1) {
        spn_cols <- (c:(c + span_val - 1)) + col_offset
        spn_row <- r + row_offset
        wb %>% mergeCells(ws, cols = spn_cols, rows = spn_row)
      }

    })
  })

  wb
}


xl_merge_from_colspans <- function(wb,ws,spans, start_row = 1, start_col = 1) {

  nrows <- nrow(spans)
  ncols <- ncol(spans)

  row_offset <- start_row - 1
  col_offset <- start_col - 1

  sapply(1:ncols, function(c) {
    sapply(1:nrows, function(r) {
      span_val <- spans[r,c]
      if(span_val > 1) {

        # if(row_offset > 0) browser()
        spn_col <- c + col_offset
        spn_rows <- (r:(r + span_val - 1)) + row_offset
        wb %>% mergeCells(ws, cols = spn_col, rows = spn_rows)
      }

    })
  })

  wb
}
