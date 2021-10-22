### Range Write 2


# Generic: range_write2 ---------------------------------------------------


#' (Over)write new data into a range (v2)
#'
#' This is a generic function that wraps around `googlesheets4::range_write()`.
#' This variation is pipe-friendly because `data` argument has been swapped to the first argument.
#' If the `data` is a named list of data.frame, each data.frame will be written to the corresponding
#' sheets that its names were matched to the names of list of data.frame, but if any names of `data` was not found
#' in sheet names, new sheets will have been created before the data.frame is written.
#' If the `data` is a data.frame, it will simply pass to `googlesheets4::range_write()`, and If `sheet` can't be found
#' in existing sheets, new `sheet` will be created, then `data` is written.
#'
#' @param data A data frame or a named list of data frames
#' @param ss Something that identifies a Google Sheet
#' @param sheet (If `data` is a data frame) Sheet to write into, in the sense of "worksheet" or "tab". You can identify a sheet by name, with a string, or by position, with a number. Ignored if the sheet is specified via range. If neither argument specifies the sheet, defaults to the first visible sheet.
#' @param range (If `data` is a data frame) Where to write. This range argument has important similarities and differences to range elsewhere (
#' @param col_names Logical, indicates whether to send the column names of data.
#' @param reformat 	Logical, indicates whether to reformat the affected cells. This default to `FALSE` meaning leaves prior formatting as is.
#'
#' @return The (list of) input ss, as an instance of sheets_id
#' @export
#'
#' @examples NULL
range_write2 <- function(data,
                         ss,
                         sheet = NULL,
                         range = NULL,
                         col_names = TRUE,
                         reformat = FALSE # Leave existing format as is
) {

  UseMethod("range_write2")

}

# List Method: range_write2 ---------------------------------------------------

#' @export
range_write2.list <- function(data,
                              ss,
                              col_names = TRUE,
                              reformat = FALSE, # Leave existing format as is
                              ...
) {

  is_element_df <- purrr::every(data, is.data.frame)
  is_named <- !is.null(names(data))
  # Check if every element of named list is a data.frame
  if( !(is_element_df && is_named) ) stop("`data` must be a named list of data.frame",
                                          call. = FALSE)

  # Check whether names of list_df in existing sheet names
  data_nm <- names(data)
  sheet_nm <- googlesheets4::sheet_names(ss)
  new_nm <- data_nm[is.na(match(data_nm, sheet_nm))] # Names from `data` that not in GS

  if(length(new_nm) > 0){
    ## Add new sheets
    googlesheets4::sheet_add(ss, sheet = new_nm)
  }

  purrr::imap(
    data,
    ~googlesheets4::range_write(
      ss = ss,
      data = .x, sheet = .y,
      range = NULL, # No A1 notation for now
      col_names = col_names,
      reformat = reformat
    )
  )



}

# DF Method: range_write2 ---------------------------------------------------

#' @export
range_write2.data.frame <- function(data,
                                    ss,
                                    sheet = NULL,
                                    range = NULL,
                                    col_names = TRUE,
                                    reformat = FALSE # Leave existing format as is
) {

  # Check if sheet names is already exist or not
  if(!is.null(sheet)){
    sheet_nm <- googlesheets4::sheet_names(ss)
    ## If it not already exist, create New sheet
    if( !all(sheet %in% sheet_nm)){
      googlesheets4::sheet_add(ss, sheet = sheet)
    }
  }

  googlesheets4::range_write(
    ss = ss,
    data = data,
    sheet = sheet,
    range = range,
    col_names = col_names,
    reformat = reformat
  )
}
