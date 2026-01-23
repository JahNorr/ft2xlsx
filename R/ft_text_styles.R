

ft_text_style_info <- function(ft) {


  df_text_styles <- part_styles(ft, "header", type = "text") %>%
    bind_rows(part_styles(ft, "body", type = "text")) %>%
    bind_rows(part_styles(ft, "footer", type = "text")) %>%
    rename(sz = font.size, b = bold, i = italic, u = underlined,
           name = font.family, vert_align = vertical.align) %>%
    select(-contains("."))%>%
    mutate(b = as.character(as.integer(as.logical(b)))) %>%
    mutate(i = as.character(as.integer(as.logical(i)))) %>%
    mutate(u = as.logical(u)) %>%
    mutate(u = if_else(u, "single", "none"))

  df_font_styles <- df_text_styles %>% select(-c(part, irow, jcol)) %>%
    distinct() %>%
    mutate(font_id = row_number())

  df_text_styles <- df_text_styles %>%
    left_join(df_font_styles, by = join_by(color, sz, b, i, u, name, vert_align)) %>%
    select(part, irow, jcol, font_id)

  list(
    data = df_text_styles,
    font_styles = df_font_styles
  )
}


ft_pars_style_info <- function(ft) {


  df_pars_styles <- part_styles(ft, "header", type = "pars") %>%
    bind_rows(part_styles(ft, "body", type = "pars")) %>%
    bind_rows(part_styles(ft, "footer", type = "pars")) %>% as.data.frame() %>%
    select(-c("tabs", starts_with("keep"),starts_with("border"))) %>%
    rename(horizontal = text.align)

  df_styles <- df_pars_styles %>% select(-c(part, irow, jcol)) %>%
    distinct() %>%
    mutate(pars_id = row_number())

  # ==================  padding  ===========================================



  df_padding <- calc_padding(df_pars_styles)
  df_padding_styles <-df_padding %>% get_padding_styles()

  df_pars_styles <- df_pars_styles %>% replace_padding_styles(df_padding, df_padding_styles)

  #                    end padding
  # ================================================================


  list(
    data = df_pars_styles,
    padding_styles = df_padding_styles
  )
}

# ===========   padding  =========================

calc_padding <- function(df_styles) {
  df_styles %>% select(starts_with("padding"))  %>%
    rename_with(.fn = ~gsub("padding[.](.*)", "\\1", .x))
}

get_padding_styles <- function(df_padding) {

  df_padding %>% distinct() %>%
    mutate(padding_id = row_number()) %>%
    relocate(padding_id) %>%
    as.data.frame()
}

replace_padding_styles <- function(df_styles, df_padding, df_padding_styles) {

  df_padding <- df_padding %>%
    left_join(df_padding_styles, by = join_by(bottom, top, left, right)) %>%
    select(padding_id)

  df_styles %>% select(-starts_with("padding")) %>% bind_cols(df_padding)
}
