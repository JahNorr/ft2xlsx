
XLSX_XML_Mgr <-
  R6::R6Class(
    classname = "XLSX_XML_Mgr",

    private = list(
      my_dir = NULL,
      xl_dir = NULL,
      workbook_file = NULL,
      styles_file = NULL,
      worksheets_dir = NULL,
      theme_dir = NULL,

      styles_xml_doc = NULL,
      ns_styles = NULL,

      current_sheet = 1,
      worksheet_file = NULL,
      sheet_xml_doc = NULL,
      ns_sheet = NULL,

      # ==============  private functions  =================

      style_nodes = function(value) {

        return(xml2::xml_find_all(private$styles_xml_doc,
                                  paste0("//d1:", value),
                                  ns = private$ns_styles)
        )

      },

      load_worksheet = function(isheet) {

        private$current_sheet <- isheet
        private$worksheet_file <- paste0(private$worksheets_dir, "\\sheet", isheet,".xml")
        private$sheet_xml_doc <- xml2::read_xml(x = private$worksheet_file)

        private$ns_sheet <- xml2::xml_ns(private$sheet_xml_doc)
      }
    ),

    public = list(

      initialize = function(file) {

        dir <- tempdir()

        test_dir <- paste0(dir,"\\xlsx_to_xml")
        if(!dir.exists(test_dir)) dir.create(test_dir)

        file.copy(from = file, to = test_dir)

        file_xlsx <- paste0(test_dir,"\\",basename(file))
        file_zip <- gsub("xlsx$", "zip", file_xlsx)

        file.rename(from = file_xlsx, file_zip)

        unzip(file_zip, overwrite = TRUE, exdir = test_dir)
        file.remove(file_zip)

        private$my_dir <- test_dir

        private$xl_dir <- paste0(test_dir, "\\xl")
        private$workbook_file <- paste0(private$xl_dir, "\\workbook.xml")
        private$styles_file <- paste0(private$xl_dir, "\\styles.xml")
        private$worksheets_dir <-  paste0(private$xl_dir, "\\worksheets")
        private$theme_dir <- paste0(private$xl_dir, "\\theme")

        private$styles_xml_doc <- xml2::read_xml(x = private$styles_file)

        private$ns_styles <- xml2::xml_ns(private$styles_xml_doc)

        private$load_worksheet(1)

      },

      rc_totals = function() {

        nodes = xml2::xml_find_all(private$sheet_xml_doc,
                                   "//d1:dimension",
                                   ns = private$ns_sheet)

        rc  <-  xml2::xml_attr(nodes,attr = "ref")

        rc <- openxlsx2::dims_to_rowcol(rc)

        cols <- rc$col %>% length()
        rows <- rc$row %>% length()

        list(
          rows = rows,
          cols = cols)
      },

      rc_style = function(irow, icol) {

        nodes = xml2::xml_find_all(private$sheet_xml_doc,
                                   "//d1:sheetData/d1:row",
                                   ns = private$ns_sheet)
        node <- nodes[[irow]]


        col_nodes <- xml2::xml_find_all(node,
                                        "d1:c",
                                        ns = private$ns_sheet)

        ncols <- length(col_nodes)

        if(icol > ncols) return(NA)

        col_node <- col_nodes[[icol]]

        style_id <- xml2::xml_attr(col_node, "s")

        if(is.na(style_id)){
          return( list(xf = NULL ,
                       font = NULL,
                       fill = NULL,
                       border = NULL

          )
          )
        }
        xf <- self$xf_style(style_id)

        fill <- self$fill_style(xf$fill_id)%>% as.data.frame()
        font <- self$font_style(xf$font_id)%>% as.data.frame()
        border <- self$border_style(xf$border_id)%>% as.data.frame()

        xf <- xf %>% as.data.frame() %>%
          mutate(style_id = {{style_id}}) %>%
          relocate(style_id)

        list(xf = xf ,

             font = font,
             fill = fill,
             border = border

        )

      },

      xf_styles = function() {

        nodes = self$cellXfs

        purrr::map(nodes, function(xf) {
          attrs <- xml2::xml_attrs(xf)
          as.data.frame(as.list(attrs), stringsAsFactors = FALSE)
        }) %>% bind_rows %>%
          mutate(xfStyleId = as.character(row_number() - 1)) %>%
          relocate(xfStyleId)

      },

      xf_style = function(istyle) {

        if(is.character(istyle)) istyle <- as.integer(istyle) + 1

        nodes = self$cellXfs

        node <- nodes[[istyle]]

        align_node <- xml2::xml_find_all(node, "d1:alignment", ns = private$ns_styles)

        list(
          font_id = xml2::xml_attr(node, "fontId"),
          border_id = xml2::xml_attr(node, "borderId"),
          fill_id = xml2::xml_attr(node, "fillId"),
          horz_align = xml2::xml_attr(align_node, "horizontal")
        )

      },

      font_style = function(istyle) {

        if(is.character(istyle)) istyle <- as.integer(istyle) + 1

        nodes = self$fonts

        node <- nodes[[istyle]]

        list(
          bold = (self$subnode_attr(node, "b", "val") == "1"),
          color = self$subnode_attr(node, "color", "rgb"),
          family = self$subnode_attr(node, "family", "val"),
          italic = (self$subnode_attr(node, "i", "val") == "1"),
          name = self$subnode_attr(node, "name", "val"),
          size = self$subnode_attr(node, "sz", "val"),
          underline = self$subnode_attr(node, "u", "val"),
          vert_align = self$subnode_attr(node, "vertAlign")
        )

      },

      fill_style = function(istyle) {

        if(is.character(istyle)) istyle <- as.integer(istyle) + 1

        nodes = self$fills

        node <- nodes[[istyle]]

        ptype = self$subnode_attr(node, "patternFill ", "patternType")

        if(ptype == "solid") {
          rgb <- self$subnode_attr(node, c("patternFill", "fgColor"), "rgb")
        } else {
          rgb <- NA
        }

        list(
          patternType = ptype,
          rgb = rgb
        )

      },

      border_style = function(istyle) {

        if(is.character(istyle)) istyle <- as.integer(istyle) + 1

        nodes = self$borders

        node <- nodes[[istyle]]

        style_left <- self$subnode_attr(node, "left", "style")
        color_left <- self$subnode_attr(node, c("left", "color"), "rgb")

        style_right <- self$subnode_attr(node, "right", "style")
        color_right <- self$subnode_attr(node, c("right", "color"), "rgb")

        style_top <- self$subnode_attr(node, "top", "style")
        color_top <- self$subnode_attr(node, c("top", "color"), "rgb")

        style_bottom <- self$subnode_attr(node, "bottom", "style")
        color_bottom <- self$subnode_attr(node, c("bottom", "color"), "rgb")

        list(
          style_left = style_left,
          color_left = color_left,
          style_right = style_right,
          color_right = color_right,
          style_top = style_top,
          color_top = color_top,
          style_bottom = style_bottom,
          color_bottom = color_bottom
        )

      },

      subnode_attr = function(node, nm, attr = "val") {

        xpath <- paste0("d1:",paste0(nm, collapse = "/d1:"))

        val <- xml2::xml_find_all(node, xpath, ns = private$ns_styles) %>%
          xml2::xml_attr(attr)

        if(length(val) == 0) val <- NA

        val
      },

      rc_styles = function(irows, icols,
                           type = c("style", "font", "fill", "border"), subtype = NULL) {

        type = match.arg(type, c("style","font", "fill", "border"))

        rc <- self$rc_totals()
        if(missing(irows)) irows <- 1:rc$rows
        if(missing(icols)) icols <- 1:rc$cols

        id_col <- paste0(type, "_id")


        df <- purrr::map(irows, \(irow){
          x_col <- purrr::map(icols, \(icol){

            tryCatch({
              x <- self$rc_style(irow, icol)

              if(any(is.na(x[[1]]))) return(NULL)
              id_val <- x$xf[[id_col]]

              if(type == "style") x_type <- x[["xf"]] else x_type <- x[[type]]

              if(is.null(x_type)) return(NULL)

              x_type %>%
                mutate(row = irow, col = icol)%>%
                mutate(id = {{id_val}}) %>%
                relocate(row , col, id ) %>%
                rename({{id_col}} := id)
            },
            error = function(e) {
              browser()
            })
          })

          x_col

        }) %>%
          bind_rows()

        df
      },

      style_matrix = function(type = "style", subtype = NULL) {

        rc <- self$rc_totals()

        id <- paste0(type, "_id")
        if(!is.null(subtype)) id <- subtype

        x <- self$rc_styles(1:rc$rows, 1:rc$cols, type = type, subtype = subtype) %>%
          select(row, col, matches(id)) %>%
          rename(id = 3) %>%
          tidyr::pivot_wider(names_from = col, values_from = id,
                             names_prefix = "[") %>%
          rename_with(.fn = ~gsub("([[].*)", "\\1]", .), .cols = everything()) %>%
          as.data.frame()


        cnames <- suppressWarnings(
          colnames(mat) %>%
            gsub("[][]", "",.) %>%
            as.integer() %>%
            replace(is.na(.),0) %>%
            sort() %>%
            gsub("([0-9]*)","[\\1]", .) %>%
            gsub("[0]","row",., fixed = T)
        )

        x <- x %>% select(all_of(cnames)) %>%
          replace(is.na(.),"")

        cat("=============================================================\n",
            type, ifelse(is.null(subtype),"", paste0(": ", subtype)),"\n\n")
        print(x)
        structure(
          x,
          class = c("style_matrix", "data.frame"),
          style_type = type
        )
        return(invisible(NULL))
      }


    ),

    active  = list(

      styles = function(value) {

        if(!missing(value)) {
          stop("styles is readonly")
        } else {
          return(private$styles_xml_doc)
        }
      },

      fonts = function(value) {

        private$style_nodes("font")
      },

      borders = function(value) {

        private$style_nodes("border")
      },

      fills = function(value) {

        private$style_nodes("fill")
      },

      cellStyles = function(value) {

        private$style_nodes("cellStyle")
      },

      cellXfs = function(value) {

        private$style_nodes("cellXfs/d1:xf")
      },

      cellStyleXfs = function(value) {

        private$style_nodes("cellStyleXfs/d1:xf")
      }

    )
  )
