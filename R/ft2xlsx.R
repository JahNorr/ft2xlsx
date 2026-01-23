
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
                        sheet_name = "ft1", index = NULL,
                        start_row = 1,
                        start_col = 1,
                        left_margin = 0.75,
                        right_margin = 0.75,
                        top_margin = 1,
                        bottom_margin = 0.75,
                        style = TRUE,
                        append = FALSE,
                        verbose = FALSE) {



  if(verbose) cat("-----------------------------------\n","
                   ... in ft_to_xlsx2\n")

  # Load or create workbook
  wb <- if (append && file.exists(file)) wb_load(file) else wb_workbook()

  # ===========================================
  # Add sheet

  if (sheet_name %in% wb$sheet_names) wb$remove_worksheet(sheet_name)

  wb$add_worksheet(sheet_name)

  ws <- sheet_name

  wb$set_page_setup(
    sheet = sheet_name,
    left = left_margin,
    right = right_margin,
    top = top_margin,
    bottom = bottom_margin
  )

  # get row offsets to body from header size

  start_header_row <- start_row
  start_body_row   <- start_header_row + ft$header$content$nrow
  start_footer_row <- start_body_row + ft$body$content$nrow

  ncols <- ft$body$content$ncol

  # Write header and body data

  df_header <- ft$header$dataset %>% select(all_of(ft$header$col_keys))
  df_body   <- ft$body$dataset   %>% select(all_of(ft$body$col_keys))

  # --------------------------------------------------------------------------

  if(verbose) cat(" ... adding header data\n")

  wb$add_data(ws, df_header, startRow = start_header_row,
              startCol = start_col, colNames = FALSE)

  if(verbose) cat(" ... adding body data\n")

  wb$add_data(ws, df_body,   startRow = start_body_row,
              startCol = start_col, colNames = FALSE)


  if(verbose) cat(" ... doing xlheader\n")

  stats_row <- ft$header$content$nrow

  xl_hdr_superscript(wb = wb, ws = ws, data = ft$header$content$data,
                     styles = ft$header$styles$text,
                     start_row = start_header_row, start_col = start_col)


  resps_row <- ft$header$content$nrow - 1

  # wrap_style

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


  wb <- xl_merge_from_spans2(wb, ws, ft$header$spans, start_row = start_header_row,
                             start_col = start_col)

  wb <- xl_merge_from_spans2(wb, ws, ft$body$spans,   start_row = start_body_row,
                             start_col = start_col)
  #
  # Apply styling

  if(style) {

    if(verbose) cat(" ... doing styling\n")

    add_styling(wb, ft, start_header_row, start_body_row , start_footer_row )

  }



  #wb_get_cell_style(wb, sheet = sheet_name, dims = "A5")  #
  # Column widths

  if(verbose) cat(" ... setting column widths\n")
  wb$set_col_widths(ws, cols = 1:ncols, widths = ft$header$colwidths * 12)

  # TODO: Footnotes (use ft$xl_footnotes or get_xl_footnotes() as before)

  if(!is.null(index) && index < nsheets) {
    nsheets <- length(wb$worksheets)
    isheets <- 1:(nsheets-1)
    order <- c(head(isheets, index-1), nsheets,  tail(isheets, nsheets - index))

    # wb$worksheets <- wb$worksheets[order]
    # wb$sheet_names <- wb$sheet_names[order]
    wb$set_order(order)
  }



  # Save workbook

  if(verbose) cat(" ... saving\n")

  wb$save(file, overwrite = TRUE)
}

inspect_style <- function(wb, style_name = NULL, xf_id = NULL) {

  # browser()
  if (!is.null(style_name)) {
    xf_id <- which(wb$styles_mgr$xf$name == style_name) - 1
  }
  if (is.null(xf_id)) stop("Must provide either style_name or xf_id")

  # get the XML string of the xf
  xf_xml <- wb$styles_mgr$xf$typ[xf_id + 1]  # +1 for R 1-based indexing

  # xf_xml contains something like:
  # "<xf applyAlignment=\"1\" applyBorder=\"1\" applyFill=\"1\" applyFont=\"1\" borderId=\"1\" fillId=\"2\" fontId=\"1\"/>"

  # parse with xml2
  library(xml2)
  xf_node <- read_xml(xf_xml)

  # extract attributes
  attrs <- xml_attrs(xf_node)

  # get the IDs
  font_id   <- as.integer(attrs["fontId"])
  fill_id   <- as.integer(attrs["fillId"])
  border_id <- as.integer(attrs["borderId"])

  # return a list
  list(
    style_name = style_name %||% paste0("xf-", xf_id),
    xf_flags   = attrs[c("applyFont","applyFill","applyBorder","applyAlignment")],
    font       = wb$styles_mgr$font[font_id + 1, ],
    fill       = wb$styles_mgr$fill[fill_id + 1, ],
    border     = wb$styles_mgr$border[border_id + 1, ]
  )
}

add_styling <- function(wb, ft, start_header_row, start_body_row , start_footer_row ) {

  cell_styles <- ft_cell_style_info(ft)
  font_styles <- ft_text_style_info(ft)
  pars_styles <- ft_pars_style_info(ft)

  pars_styles$data <- pars_styles$data %>% select(-line_spacing, -shading.color)

  add_my_border_styles(wb, cell_styles$border_styles)
  add_my_font_styles(wb, font_styles$font_styles)
  add_my_fill_styles(wb, cell_styles$fill_styles)

  my_style_data <- cell_styles$data %>%
    left_join(pars_styles$data, by = join_by(part, irow, jcol))%>%
    left_join(font_styles$data, by = join_by(part, irow, jcol)) %>%
    mutate(border_id = paste0("border_", border_id)) %>%
    mutate(font_id = paste0("font_", font_id)) %>%
    mutate(fill_id = paste0("fill_", fill_id)) %>%
    mutate(padding_id = paste0("padding_", padding_id)) %>%
    as.data.frame()

  my_styles <- my_style_data %>%
    select(-c(part, irow, jcol)) %>%
    distinct() %>%
    mutate(style_id = paste0("style_",row_number())) %>%
    as.data.frame()

  add_my_styles_to_xlsx(wb, my_styles)

  parts  <-  c("header", "body", "footer")
  starts <- purrr::map_int(parts, ~get(paste0("start_", .x, "_row")))

  df_starts <- data.frame(part = parts, start = starts)

  df_spans <- purrr::map(parts, ~ft[[.x]]$spans$rows %>% as.data.frame()) %>%
    bind_rows() %>%
    mutate(irow = row_number()) %>%
    relocate(irow)

  colnames(df_spans) <- c("irow", paste0("col_",1:(ncol(df_spans) - 1)))

  df_spans <- df_spans %>%
    tidyr::pivot_longer(cols = starts_with("col")) %>%
    mutate(jcol = as.integer(gsub("col_","",name))) %>%
    mutate(keep = value > 0) %>%
    select(-name, -value)%>%
    relocate(irow, jcol)

  data <- my_style_data %>%
    left_join(my_styles,
              by = join_by(vertical.align, text.direction, fill_color, hrule, border_id,
                           fill_id, margins_id, horizontal, padding_id, font_id)) %>%
    select(part, irow, jcol, style_id, font_id, border_id) %>%
    mutate(irow = as.integer(irow), jcol = as.integer(jcol)) %>%
    as.data.frame() %>%
    left_join(df_starts, by = join_by(part)) %>%
    mutate(irow = irow - 1 + start) %>%
    select(-part, -start) %>%
    left_join(df_spans, by = join_by(irow, jcol))

  purrr::pwalk(data, \(irow, jcol, style_id, font_id, border_id, keep) {

    if(keep) {
      xf_id <- wb$styles_mgr$get_xf_id(style_id)
      #if(xf_id == "1") xf_id <-  "2"
      dims <- rowcol_to_dims(irow, jcol)
      #if(jcol == 1) cat("\n")
      #cat(gsub("[:].*","",dims), ": ", xf_id, "  ", sep = "")
      wb$set_cell_style(
        dims = dims,
        style = xf_id
      )
    }
  })

  return(invisible(NULL))

}

add_my_styles_to_xlsx <- function(wb, styles) {

  purrr::pwalk(
    styles,
    \(vertical.align, text.direction, fill_color, hrule, border_id,
      fill_id, margins_id, horizontal, padding_id, font_id, style_id) {

      x <-  create_cell_style(border_id = wb$styles_mgr$get_border_id(border_id),
                              fill_id = wb$styles_mgr$get_fill_id(fill_id),
                              font_id = wb$styles_mgr$get_font_id(font_id),
                              horizontal = horizontal)


      wb$styles_mgr$add(x, paste0(style_id))
    })
}


add_my_fill_styles <- function(wb, styles) {

  purrr::pwalk(
    styles,
    \(fill_id, fg_color, pattern_type) {

      x <- create_fill(pattern_type = pattern_type, fg_color = wb_color_me(fg_color))

      wb$styles_mgr$add(x, paste0("fill_", fill_id))
    })

}

add_my_font_styles <- function(wb, styles) {

  purrr::pmap(
    styles,
    \(b,color,font_id,i,name,sz,u,vert_align) {

      x <- create_font(b = b, color = wb_color(color), i = i, name = name, charset = "1",
                       sz = as.character(as.integer(sz)), u = u,
                       vert_align = vert_align, scheme = "none")


      wb$styles_mgr$add(x, paste0("font_", font_id))

    })

}

add_my_border_styles <- function(wb, styles) {

  purrr::pwalk(
    styles,
    \(border_id, bottom_color, top_color, left_color, right_color,
      bottom, top, left, right) {

      x <- create_border(bottom_color = wb_color_me(bottom_color),
                         top_color = wb_color_me(top_color),
                         left_color = wb_color_me(left_color),
                         right_color = wb_color_me(right_color),
                         bottom = bottom, top = top, left = left, right = right)


      wb$styles_mgr$add(x, paste0("border_", border_id))
    })

}

wb_color_me <- function(x) {

  if (x == "transparent" || is.null(x)) {
    x <- wb_color(hex = "00FFFFFF",format = "ARGB")
  }
  wb_color(x)
}

xl_hdr_superscript <- function(wb, ws, data, styles, start_row = 1, start_col = 1) {

  nrows <- nrow(data)
  ncols <- ncol(data)

  irow <- start_row
  row_offset = start_row - 1
  col_offset = start_col - 1


  hdr_data <- list(
    color = styles$color$data,
    font.size = styles$font.size$data,
    bold = styles$bold$data,
    italic = styles$italic$data,
    underlined = styles$underlined$data,
    font.family = styles$font.family$data,
    vertical.align = styles$vertical.align$data,
    shading.color = styles$shading.color$data
  )

  # for each row and col of data

  purrr::map(1:nrows, function(irow) {
    purrr::map(1:ncols, function( icol) {

      # ---- get the dims

      dims <- rowcol_to_dims(irow + row_offset, icol + col_offset)


      #browser()
      purrr::map(data[irow, icol], function(ftnt_lines) {

        x <- NULL #fmt_txt("")
        #cat(irow, icol, class(ftnt_lines), nrow(ftnt_lines),"\n")
        if(nrow(ftnt_lines) > 1) {

          purrr::pmap(ftnt_lines, function(txt, font.size, italic, bold,
                                           underlined, color, shading.color, font.family,
                                           hansi.family, eastasia.family, cs.family,
                                           vertical.align, width, height,  url,
                                           eq_data, word_field_data,
                                           img_data, .chunk_index) {



            #if(irow == 4) browser()


            if(is.na(color)) {
              color <- wb_color(hdr_data$color[irow, icol] %>% unname())
            } else {
              color <- wb_color(color)
            }


            x <<- x + fmt_txt(txt,
                              font = ifelse(is.na(font.family),
                                            hdr_data$font.family[irow, icol], font.family),
                              bold = ifelse(is.na(bold),
                                            hdr_data$bold[irow, icol] ,
                                            bold),
                              color = color,

                              size = ifelse(is.na(font.size),
                                            hdr_data$font.size[irow, icol], font.size),
                              vert_align = ifelse(is.na(vertical.align), "baseline",
                                                  vertical.align))


          })


          wb <<- wb_add_data(wb = wb, sheet = ws, x = x, dims = dims)
        }
      })
    })
  })

  wb
}

xl_footnotes <- function(wb, ws, data, start_row) {

  nnotes <- nrow(data)

  irow <- start_row
  purrr::map(data, function(ftnt_lines) {

    dims <- rowcol_to_dims(irow, 1)
    x <- fmt_txt("")

    purrr::pmap(ftnt_lines, function(txt, font.size, italic, bold,
                                     underlined, color, shading.color, font.family,
                                     hansi.family, eastasia.family, cs.family, vertical.align,
                                     width, height,  url, eq_data, word_field_data,
                                     img_data, .chunk_index) {


      x <<- x + fmt_txt(txt,
                        bold = ifelse(is.na(bold), FALSE, bold),
                        #color = ifelse(is.na(color), wb_color("black"), wb_color(color)),
                        size = ifelse(is.na(font.size), 11, font.size),
                        vert_align = ifelse(is.na(vertical.align), "baseline", vertical.align))

    })
    wb <<- wb_add_data(wb = wb, sheet = ws, x = x, dims = dims)
    irow <<- irow + 1
  })

  wb
}


xl_merge_from_spans2 <- function(wb, ws, spans, start_row = 1, start_col = 1) {
  wb <- xl_merge_from_rowspans(wb, ws, spans$rows, start_row, start_col)
  wb <- xl_merge_from_colspans(wb, ws, spans$columns, start_row, start_col)
  wb
}

xl_merge_from_rowspans <- function(wb, ws, spans, start_row = 1, start_col = 1) {
  row_offset <- start_row - 1
  col_offset <- start_col - 1
  nrows <- nrow(spans)
  ncols <- ncol(spans)

  for (r in seq_len(nrows)) {
    for (c in seq_len(ncols)) {

      span_val <- spans[r, c]
      if (span_val > 1) {
        spn_cols <- (c:(c + span_val - 1)) + col_offset
        spn_row  <- r + row_offset
        dims <- rowcol_to_dims(spn_row, spn_cols)
        dims2 <- rowcol_to_dims(spn_row, tail(spn_cols, -1))


        wb$add_data(x = matrix(NA, nrow = 1, ncol = span_val - 1),
                    start_col = spn_cols[2],
                    start_row = spn_row,
                    col_names = FALSE,
                    row_names = FALSE
        )

        wb$merge_cells(ws, dims = dims)
      }
    }
  }
  wb
}


xl_merge_from_colspans <- function(wb, ws, spans, start_row = 1, start_col = 1) {
  row_offset <- start_row - 1
  col_offset <- start_col - 1
  nrows <- nrow(spans)
  ncols <- ncol(spans)

  for (c in seq_len(ncols)) {
    for (r in seq_len(nrows)) {
      span_val <- spans[r, c]
      if (span_val > 1) {
        spn_cols <- c + col_offset
        spn_rows <- (r:(r + span_val - 1)) + row_offset
        dims <- rowcol_to_dims(spn_rows, spn_cols)
        dims2 <- rowcol_to_dims(tail(spn_rows, -1), spn_cols)
        wb$add_data(x = NA, dims = dims2)

        wb$merge_cells(ws, dims = dims)
      }
    }
  }
  wb
}


border_data <- function(styles, row, col) {

  types <- names(styles) %>% {.[grepl("^border[.]", .)]} %>%
    gsub("border\\.", "", .) %>% gsub("(.*)[.].*", "\\1", .) %>% unique()

  x <- lapply(types, function(type) {
    bords <- c("top", "bottom", "left", "right")
    vals <- sapply(bords, function(bord) {
      el <- paste0("border.", type, ".", bord)
      styles[[el]]$data[row, col]
    })
    names(vals) <- bords
    vals
  })
  names(x) <- types

  x

}

ft_text_decoration <- function(text_styles, row, col) {
  list(
    bold = text_styles$bold$data[row, col] %>% unname(),
    italic = text_styles$italic$data[row, col] %>% unname(),
    underline = text_styles$underlined$data[row, col] %>% unname()
  )
}


