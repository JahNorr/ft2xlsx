

style_fill <- function(wb, cell_styles, start_row , start_col) {

  browser()
  # dims <- rowcol_to_dims(r + row_offset, c + col_offset)
  #
  fill_vals <- cell_styles$background.color$data #[r, c]

  purrr::walk(fill_vals,)
  if (fill_val == "transparent" || is.null(fill_val)) {
    fill_val <- wb_color(hex = "00FFFFFF",format = "ARGB")
    fill_pattern = "none"
  } else {
    fill_val <- wb_color(fill_val)
    fill_pattern = "solid"
  }


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

