


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
