


xl_merge_from_spans2 <- function(wb, ws, spans, start_row = 1, start_col = 1) {
  wb <- xl_merge_from_rowspans2(wb, ws, spans$rows, start_row, start_col)
  wb <- xl_merge_from_colspans2(wb, ws, spans$columns, start_row, start_col)
  wb
}

xl_merge_from_rowspans2 <- function(wb, ws, spans, start_row = 1, start_col = 1) {
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

        wb$merge_cells(ws, dims = dims)
      }
    }
  }
  wb
}


xl_merge_from_colspans2 <- function(wb, ws, spans, start_row = 1, start_col = 1) {
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

        wb$merge_cells(ws, dims = dims)
      }
    }
  }
  wb
}

