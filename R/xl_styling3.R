

xl_styling3 <- function(wb, ws, part, part_data, start_row = 1, start_col = 1) {

  nrows <- part_data$content$nrow
  ncols <- part_data$content$ncol
  row_offset <- start_row - 1
  col_offset <- start_col - 1

  cell_styles <- part_data$styles$cells
  par_styles  <- part_data$styles$pars
  text_styles <- part_data$styles$text

  browser()
  style_fill(wb, cell_styles, start_row , start_col)
  # do cell by cell

  for (r in seq_len(nrows)) {
    for (c in seq_len(ncols)) {

      dims <- rowcol_to_dims(r + row_offset, c + col_offset)

      # handle fill

      fill_val <- cell_styles$background.color$data[r, c]

      if (fill_val == "transparent" || is.null(fill_val)) {
        fill_val <- wb_color(hex = "00FFFFFF",format = "ARGB")
        fill_pattern = "none"
      } else {
        fill_val <- wb_color(fill_val)
        fill_pattern = "solid"
      }

      # handle border

      border_data <- border_data(cell_styles, r, c)
      border_styles <- border_data$style %>%
        gsub("solid", "thin", .) %>%
        {.[border_data$width == 0] <- "none"; .} %>%
        {.[border_data$width > 0.7 & . == "thin"] <- "medium"; .} %>%
        {.[border_data$width > 1 & . == "thin"]   <- "thick"; .} %>%
        {.[border_data$color == "transparent"] <- "none"; .}

      border_colors <- border_data$color %>% gsub("transparent", "white", .)

      # handle alignment

      valign <- cell_styles$vertical.align$data[r, c] %>% as.character()
      halign <- par_styles$text.align$data[r, c] %>% as.character()

      # handle color/font

      color <- text_styles$color$data[r, c]
      font_family <- unname(text_styles$font.family$data[r, c])
      font_size   <- unname(text_styles$font.size$data[r, c])

      # handle text decorations (bold/italic, etc)

      dec <- ft_text_decoration(text_styles, r, c)

      text_dec <- c(if(dec$bold) "bold" else NULL,
                    if(dec$italic) "italic" else NULL,
                    if(dec$underline) "underline" else NULL)

      dec$underline <- ifelse(dec$underline, "single", "")

      #cat("[",r,",",c,"]  ","h:[",halign, "]  v:[",valign, "]\n", sep = "")

      # make the style
      wb <- wb %>%
        wb_add_fill(dims = dims, color = fill_val, pattern = fill_pattern) %>%

        wb_add_font(dims = dims,
                    name = font_family,
                    size = font_size,
                    color = wb_color(color),
                    bold = dec$bold,
                    italic =  dec$italic,
                    underline = dec$underline) %>%

        wb_add_border(dims = dims,
                      bottom_color = wb_color(border_colors[2]),
                      left_color =  wb_color(border_colors[3]),
                      right_color =  wb_color(border_colors[4]),
                      top_color =  wb_color(border_colors[1]),
                      bottom_border = unname(border_styles["bottom"]),
                      left_border = unname(border_styles["left"]),
                      right_border = unname(border_styles["right"]),
                      top_border = unname(border_styles["top"])) %>%


        wb_add_cell_style(dims = dims,
                          apply_alignment = TRUE,
                          horizontal = halign,
                          vertical = unname(as.character(valign)))
    }
  }

  wb
}
