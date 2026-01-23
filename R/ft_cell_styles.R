

part_styles <- function(ft, part, type = c("cells", "text", "pars"))  {

  type <- match.arg(type[1], c("cells", "text", "pars"))

  type_vars <- names(ft[[part]]$styles[[type]]) %>% {.[!. %in% c("width", "height")]}

  nrows <- ft[[part]]$content$nrow
  ncols <- length(ft$col_keys)

  df_sty <- purrr::map(1:nrows, \(irow) {

    purrr::map(1:ncols, \(jcol) {

      lst <- type_vars %>% purrr::map_chr( \(var) {

        x  <-  ft[[part]]$styles[[type]][[var]]$data[irow,jcol]
        x
      })

      names(lst) <- type_vars
      lst["part"] <- part
      lst["irow"] <- irow
      lst["jcol"] <- jcol
      tibble::tibble_row(!!!as.list(lst))

    })

  }) %>% bind_rows() %>%
    relocate(c(part, irow, jcol))

}

# ===========   borders  =========================

calc_borders <- function(df_styles) {

  sides <- c("top", "right", "bottom", "left")

  df_borders <- df_styles %>% select(starts_with("border")) %>%
    rename_with(.fn = ~gsub("border.style.", "", .x)) %>%
    rename_with(.fn = ~gsub("border[.](.*)[.](.*)", "\\2_\\1", .x))

  purrr::walk(sides, \(side) {
    colors <- df_borders[[paste0(side, "_color")]]
    widths <- df_borders[[paste0(side, "_width")]]

    df_borders[side] <<- ifelse(colors == "transparent" | widths == 0.0, "none",
                                df_borders[[side]])

    styles <- df_borders[[side]]

    df_borders[paste0(side, "_color")] <<- ifelse(widths == 0.0, "transparent",
                                                  df_borders[[paste0(side, "_color")]])
    df_borders[side] <<- ifelse( df_borders[[side]] == "solid" & widths < 0.2, "hairline",
                                 df_borders[[side]])
    df_borders[side] <<- ifelse( df_borders[[side]] == "solid" & widths < 0.6, "thin",
                                 df_borders[[side]])
    df_borders[side] <<- ifelse( df_borders[[side]] == "solid" & widths < 1.2, "medium",
                                 df_borders[[side]])
    df_borders[side] <<- ifelse( df_borders[[side]] == "solid" & widths >= 1.2, "thick",
                                 df_borders[[side]])

  })

  df_borders

}

get_border_styles <- function(df_borders) {

  df_borders %>%
    select(-ends_with("width")) %>% distinct() %>%
    mutate(border_id = row_number()) %>%
    relocate(border_id) %>%
    as.data.frame()
}

replace_border_styles <- function(df_styles, df_borders, df_border_styles) {

  df_borders <- df_borders %>%
    left_join(df_border_styles, by = join_by(bottom_color, top_color, left_color,
                                             right_color, bottom, top, left, right)) %>%
    select(border_id)

  df_styles %>% select(-starts_with("border")) %>% bind_cols(df_borders)
}

# ===========   fill  =========================

calc_fills <- function(df_styles) {

  df_styles %>% select(background.color) %>%
    rename(fg_color = background.color) %>%
    mutate(pattern_type = if_else(fg_color == "transparent", "none", "solid"))


}

get_fill_styles <- function(df_fills) {

  df_fills %>%
    distinct() %>%
    mutate(fill_id = row_number()) %>%
    relocate(fill_id) %>%
    as.data.frame()
}

replace_fill_styles <- function(df_styles, df_fills, df_fill_styles) {

  df_fills <- df_fills %>%
    left_join(df_fill_styles, by = join_by(fg_color, pattern_type)) %>%
    select(fill_id)

  df_styles %>% select(-starts_with("fill")) %>% bind_cols(df_fills)
}

# ===========   margins  =========================

calc_margins <- function(df_styles) {
  df_styles %>% select(starts_with("margin"))  %>%
    rename_with(.fn = ~gsub("margin[.](.*)", "\\1", .x))
}

get_margins_styles <- function(df_margins) {

  df_margins %>% distinct() %>%
    mutate(margins_id = row_number()) %>%
    relocate(margins_id) %>%
    as.data.frame()
}

replace_margins_styles <- function(df_styles, df_margins, df_margin_styles) {

  df_margins <- df_margins %>%
    left_join(df_margin_styles, by = join_by(bottom, top, left, right)) %>%
    select(margins_id)

  df_styles %>% select(-starts_with("margin")) %>% bind_cols(df_margins)
}
###################################################################################

#      script code for cells

ft_cell_style_info <- function(ft) {

  df_cell_styles <- part_styles(ft, "header", type = "cells") %>%
    bind_rows(part_styles(ft, "body", type = "cells")) %>%
    bind_rows(part_styles(ft, "footer", type = "cells"))

  # ==================  borders  ===========================================



  df_borders <- calc_borders(df_cell_styles)

  df_border_styles <- df_borders %>% get_border_styles()

  df_cell_styles <- df_cell_styles %>% replace_border_styles(df_borders, df_border_styles)


  #                    end borders
  # ================================================================

  df_fills <- calc_fills(df_cell_styles)

  df_fill_styles <- df_fills %>% get_fill_styles()

  df_cell_styles <- df_cell_styles %>% replace_fill_styles(df_fills, df_fill_styles)


  # ==================  margins  ===========================================



  df_margins <- calc_margins(df_cell_styles)
  df_margins_styles <-df_margins %>% get_margins_styles()

  df_cell_styles <- df_cell_styles %>% replace_margins_styles(df_margins, df_margins_styles)

  #                    end margins
  # ================================================================

  df_cell_styles <- df_cell_styles %>% rename(fill_color = background.color)

  list(
    data = df_cell_styles,
    border_styles = df_border_styles,
    fill_styles = df_fill_styles,
    margins_styles = df_margins_styles
    )
}

