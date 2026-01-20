#' Read data from a GLATOS project workbook
#'
#' Read data from a GLATOS project workbook (xlsm or xlsx file).
#'
#' @param wb_file A character string with path and name of workbook in standard
#'   GLATOS format (.xlsm or .xlsx). If only file name is given, then the file
#'   must be located in the working directory. See the GLATOSWeb Data
#'   Portal <https://glatos.glos.us> for file format definitions.
#'
#' @param wb_version An optional character string with the workbook version
#'   number. If NULL (default value) then version will be determined by
#'   evaluating workbook structure. Currently, the only allowed values are
#'   `NULL`, `"1.3"`, and `"1.4"`. See Details. Any other values will trigger
#'   an error.
#'
#' @param read_all If TRUE, then all columns and sheets (e.g., user-created
#'   "project-specific" columns or sheets) in the workbook will be imported. If
#'   FALSE (default value) then only columns and sheets in the standard GLATOS
#'   workbook will be imported (project-specific columns will be ignored).
#'
#' @param simplify If TRUE (default value), then the returned object is a
#'   `glatos_workbook` object. If FALSE, then the returned object is a list with
#'   an element for each sheet in `wb_file`. See Details.
#'
#' @details When `simplify = TRUE`, data in workbook sheets 'Deployment',
#'   'Recovery', and 'Location' are merged on columns 'GLATOS_PROJECT',
#'   'GLATOS_ARRAY', 'STATION_NO', 'CONSECUTIVE_DEPLOY_NO', AND 'INS_SERIAL_NO'
#'   to produce the output data frame `receivers`. Data in workbook sheets
#'   'Project' and 'Tagging' are passed through to new data frames named
#'   'project' and 'animals', respectively, and data from workbook sheet
#'   'Proposed' is not included in result. When `simplify = FALSE`, data in all
#'   sheets in the standard workbook are passed through to new data frames with
#'   like names (e.g., 'location', 'deployment', 'recovery').
#'
#' @details If `read_all = TRUE` then each sheet not included in the standard
#'   workbook (e.g., not named 'Project', 'Location', 'Deployment', 'Recovery',
#'   or 'Tagging') will be included as an element of the returned list; and in
#'   standard workbook sheets, any non-standard columns (i.e, 'project-specific
#'   fields') will be included in the result. Names of non-standard columns
#'   may be changed (e.g., for uniqueness), with warnings.
#'
#' @details Data are read from the input file using
#'   [read_excel][readxl::read_excel] in the 'readxl' package. If `read_all =
#'   TRUE` then the type of data in each user-defined column (and sheet) will be
#'   'guessed' by [read_excel][readxl::read_excel]. Therefore, if `read_all =
#'   TRUE` then the structure of those columns should be carefully reviewed in
#'   the result. See [read_excel][readxl::read_excel] for details.
#'
#' @details Column `animal_id` is considered a required column by many other
#'   functions in this package, so it will be created if any records are `NULL`.
#'   When created, it will be constructed from `tag_code_space` and
#'   `tag_id_code`, separated by '-'.
#'
#' @details Timezone attribute of all timestamp columns (class `POSIXct`) in
#'   output will be "UTC" and all 'glatos-specific' timestamp and timezone
#'   columns will be omitted from result.
#'
#' @details As of glatos 0.9.0, if a sheet contains two columns with the same
#'   name, then the sheet is not loaded and an error is returned. In earlier
#'   versions, a suffix was added to all but the first in each set of
#'   duplicate column names.
#'
#' @note ***On warnings and errors about date and timestamp formats.*** Date and
#'   time columns are sometimes stored as text in Excel. When those records are
#'   loaded by this function, there are two possible outcomes. \cr \cr 1. If the
#'   records are formatted according to the GLATOS Data Dictionary specification
#'   (e.g., "YYYY-MM-DD" for dates and "YYYY-MM-DD HH:MM" for timestamps; see
#'   <https:\\glatos.glos.us>) those records should be properly loaded into R,
#'   but the user is encouraged to verify that they were loaded correctly, so a
#'   warning points the user to those records in the workbook. Users may want to
#'   format as custom date in the workbook to avoid warnings in the future. \cr
#'   \cr 2. If the format of a date-as-text column is not consistent with GLATOS
#'   specification, then no data will be loaded and an error will alert the user
#'   to this condition. Similarly, if a date or date-time column is stored
#'   as a number in Excel, then no data will be loaded and an error will alert
#'   the user to this condition.\cr \cr
#'   ***On cells with locked formatting in Excel:*** Occasionally the
#'   format of a cell in Excel will be locked. In those cases, it is sometimes
#'   possible to force date formatting in Excel by (1) highlighting the columns
#'   that need reformatting, (2) select 'Text-to-columns' in the 'Data' menu,
#'   (3) select 'Delimited' and 'next', (4) uncheck all delimiters and 'next',
#'   (5) choose 'Date: YMD' in the 'Column data format' box, and (6) 'Finish'.
#'
#' @return If `simplify = TRUE`, a list of class `glatos_workbook` with three
#'   elements (described below) containing data from the standard GLATOS
#'   Workbook sheets. If `read_all = TRUE`, then additional elements will be
#'   added with names corresponding to non-standard sheet names.
#' \describe{
#'   \item{metadata}{A list with data about the project and workbook.}
#'   \item{animals}{A data frame of class `glatos_animals` with data about
#'   tagged animals.}
#'   \item{receivers}{A data frame of class `glatos_receivers` with data
#'   about telemetry receivers.}
#'
#'   If `simplify = FALSE`, a list (class = `list`) with one element for each
#'   sheet in `wb_file`.
#' }
#'
#' @author C. Holbrook \email{cholbrook@usgs.gov}
#'
#' @seealso [read_excel][readxl::read_excel]
#'
#' @examples
#'
#' # Example 1: Version 1.3 (xlsm)
#'
#' # get path to example GLATOS Data Workbook
#' wb_file <- system.file("extdata",
#'   "walleye_workbook.xlsm",
#'   package = "glatos"
#' )
#'
#' # note that code above is needed to find the example file
#' # for real glatos data, use something like below
#' # wb_file <- "c:/path_to_file/HECWL_GLATOS_20150321.xlsm"
#'
#' wb <- read_glatos_workbook(wb_file)
#'
#' wbr <- read_glatos_workbook(wb_file, simplify = FALSE)
#'
#'
#' # Example 2: Version 1.4 (xlsx)
#'
#' wb2_file <- system.file("extdata",
#'   "walleye_workbook.xlsx",
#'   package = "glatos"
#' )
#'
#' wb2 <- read_glatos_workbook(wb2_file)
#'
#' wbr2 <- read_glatos_workbook(wb2_file, simplify = FALSE)
#'
#' @import readxl
#'
#' @export
read_glatos_workbook <- function(
  wb_file,
  read_all = FALSE,
  wb_version = NULL,
  simplify = TRUE
) {
  # See version-specific file specifications
  # internal glatos_workbook_spec in R/sysdata.r

  # to avoid R CMD CHECK note "undefined global functions or variables"
  extra_sheets <- NULL

  # temporary remapping to internal data (for use externally)
  glatos_workbook_schema <- glatos:::glatos_workbook_schema


  # Get and check version if specified
  wb_version <- identify_workbook_version(
    wb_file,
    wb_version
  )


  #-Workbook v1.3, v1.4--------------------------------------------------------------
  if (wb_version %in% c("1.3", "1.4")) {
    # Get sheet names in external file
    wb_sheets <- readxl::excel_sheets(wb_file)


    # Subset sheet names to read
    sheets_to_read <-
      if (read_all) {
        wb_sheets
      } else {
        wb_sheets[tolower(wb_sheets) %in% names(glatos_workbook_schema$v1.3)]
      }


    # Preallocate glatos_workbook object
    wb <- setNames(
      vector("list", length(sheets_to_read)),
      tolower(sheets_to_read)
    )


    # Read project data (non-standard structure)
    wb$project <- read_workbook_project(wb_file)


    # Read all sheets except project
    sheets_to_read <- grep("project", sheets_to_read,
      ignore.case = TRUE,
      value = TRUE,
      invert = TRUE
    )

    for (i in 1:length(sheets_to_read)) {
      schema_i <- glatos_workbook_schema[["v1.3"]][[tolower(sheets_to_read[i])]]

      # Specify first row to read (with headers)
      xl_first_row <- switch(wb_version,
        "1.3" = 2,
        "1.4" = 1
      )

      # Specify last column to read (based on 'read_all' arg)
      xl_last_col <- if (read_all) NA_integer_ else nrow(schema_i)


      # Read sheet data
      sheet_i <-
        tryCatch(
          readxl::read_excel(
            wb_file,
            sheet = sheets_to_read[i],
            range = cellranger::cell_limits(
              ul = c(xl_first_row, 1),
              lr = c(NA, xl_last_col)
            ),
            col_types = "list",
            na = c("", "NA"),
            n_max = 0,
            guess_max = 1048576,
            .name_repair = "minimal"
          ),
          error = function(e) e
        )

      # Add context (sheet name) to errors
      if (inherits(sheet_i, "error")) {
        stop("In workbook '", basename(wb_file),
          "' sheet '", sheets_to_read[i], "':",
          "\n ", sheet_i$message,
          call. = FALSE
        )
      }

      # Preallocate error messages for ith sheet

      err_i <- NULL


      # Check column names

      col_names_i <- names(sheet_i)


      # Check for duplicate column names

      dupl_cols <- unique(col_names_i[duplicated(tolower(col_names_i))])

      if (length(dupl_cols) > 0) {
        err_i <-
          c(
            err_i,
            paste0(
              "In sheet '", sheets_to_read[i], "':",
              "\n Duplicate column names are not allowed: \n  '",
              paste0(dupl_cols, collapse = "'\n  '"), "'\n"
            )
          )
      }


      # Check for required columns

      missing_cols <- setdiff(schema_i$name, tolower(col_names_i))

      if (length(missing_cols) > 0) {
        err_i <-
          c(
            err_i,
            paste0(
              "In sheet '", sheets_to_read[i], "':",
              "\n Required columns not found: \n  '",
              paste0(missing_cols, collapse = "'\n  '"), "'\n"
            )
          )
      }

      # Preallocate new object for parsed/cast values
      # keep as tibble here to preserve structure
      sheet_i2 <- sheet_i[]


      if (nrow(sheet_i) > 0) {
        # Add attribute for warnings and errors
        # warnings are warning_cast_to_check
        # errors are error_input_class_skipped and error_cast_failed

        attr(sheet_i2, "warning_cast_to_check") <- list()
        attr(sheet_i2, "error_input_class_skipped") <- list()
        attr(sheet_i2, "error_cast_failed") <- list()


        # Coerce by expected column type

        # character
        char_cols <- col_names_i[tolower(col_names_i) %in%
          with(schema_i, name[type == "character"])]

        for (j in char_cols) {
          sheet_i2[[j]] <- cast(sheet_i[[j]],
            new_class = "character"
          )
        }


        # numeric
        num_cols <- col_names_i[tolower(col_names_i) %in%
          with(schema_i, name[type == "numeric"])]

        for (j in num_cols) {
          sheet_i2[[j]] <- cast(sheet_i[[j]],
            new_class = "numeric"
          )
        }

        # POSIXct

        # Only support POSIXct or character string that parses correctly
        # Do not accept numeric input.

        posix_cols <- col_names_i[tolower(col_names_i) %in%
          with(schema_i, name[type == "POSIXct"])]

        for (j in posix_cols) {
          # cast existing POSIXct or character to character
          sheet_i2[[j]] <- cast(sheet_i[[j]],
            new_class = "character",
            old_class = c("character", "POSIXct")
          )


          # cast character to POSIXct, enforce timezone, but return UTC

          args_ij <- schema_i$args[schema_i$name == tolower(j)]

          # strip spaces (for formatting consistency)
          args_ij <- gsub(" ", "", args_ij)

          if (grepl("tz=REFCOL", args_ij)) {
            tz_col <- gsub("tz=REFCOL\\(|\\)", "", args_ij)

            tz_ij <- paste0("US/", sheet_i[[grep(tz_col,
              names(sheet_i),
              ignore.case = TRUE
            )]])
          } else {
            tz_ij <- gsub("tz=|tz=\"|\"", "", args_ij)
          }

          sheet_i2[[j]] <- cast(sheet_i2[[j]],
            new_class = "POSIXct",
            old_class = c(
              "character",
              "POSIXct"
            ),
            tz = tz_ij
          )

          attr(sheet_i2[[j]], "tzone") <- "UTC"
        } # end j


        # Date

        # Only support POSIXct or character string that parses correctly
        # Do not accept numeric input.

        date_cols <- col_names_i[tolower(col_names_i) %in%
          with(schema_i, name[type == "Date"])]

        for (j in date_cols) {
          # cast existing POSIXct or character to character
          sheet_i2[[j]] <- cast(sheet_i[[j]],
            new_class = "Date",
            old_class = c("character", "POSIXct")
          )
        } # end j
      } # end if


      # Handle 'extra' columns (not in schema)
      # if multiple classes present in a column, cast to "highest-level" class

      extra_cols <- setdiff(tolower(col_names_i), schema_i$name)

      if (read_all) {
        if (nrow(sheet_i) > 0) {
          supported_classes <- c(
            "POSIXct",
            "Date",
            "numeric",
            "character",
            "logical"
          )

          for (j in posix_cols) {
            types_ij <- unique(unlist(lapply(sheet_i[[j]], class)))

            # expect 'highest-level' observed class
            type_exp <- intersect(supported_classes, types_ij)[1]

            # cast to type_exp
            # but if type_exp is POSIXct, cast to character

            sheet_i2[[j]] <- cast(sheet_i[[j]],
              new_class = ifelse(type_exp == "POSIXct",
                "character",
                type_exp
              )
            )
          }
        } # end j
      } else {
        std_names_i <- names(sheet_i2)[tolower(names(sheet_i2)) %in%
          schema_i$name]

        sheet_i2 <- sheet_i2[, std_names_i]
      }


      # Append to wb
      wb[[tolower(sheets_to_read[i])]] <- as.data.frame(sheet_i2)


      if (simplify) {
        # Change names of all standard columns to lowercase

        std_names_index <-
          match(
            schema_i$name,
            tolower(names(wb[[tolower(sheets_to_read[i])]]))
          )

        names(wb[[tolower(sheets_to_read[i])]])[std_names_index] <- schema_i$name
      }
    } # end i


    # Collect warnings across sheets
    warn <- c(
      attr(wb[["locations"]], "warn"),
      attr(wb[["proposed"]], "warn"),
      attr(wb[["deployment"]], "warn"),
      attr(wb[["recovery"]], "warn"),
      attr(wb[["tagging"]], "warn")
    )

    if (length(warn) > 0) warning(warn)


    # Collect errors across sheets
    err <- c(
      err_i,
      attr(wb[["locations"]], "err"),
      attr(wb[["proposed"]], "err"),
      attr(wb[["deployment"]], "err"),
      attr(wb[["recovery"]], "err"),
      attr(wb[["tagging"]], "err")
    )

    if (length(err) > 0) {
      stop(err,
        call. = FALSE
      )
    }


    # PICK UP HERE
    #
    #   if (!read_all) {
    #     # subset only columns in schema (by name)
    #     # - use match so that first column with each name is selected if > 1
    #     #tmp <- tmp[, match(schema_i$name, tolower(prj_cols_i))]
    #
    #     cols_to_read_i <- 1:length(std_cols_i)
    #
    #   } else {
    #
    #     # identify project-specific fields
    #     extra_cols_i <- col_names_i[(length(schema_i$name) + 1):length(col_names_i)]
    #
    #     # identify new columns to add
    #     if (length(extra_cols_i) > 0) {
    #
    #       # count column names to identify and rename any conflicting
    #       col_counts <- table(tolower(col_names_i))
    #       conflict_cols <- col_counts[col_counts > 1]
    #
    #       if (length(conflict_cols) > 0) {
    #         # rename conflict cols
    #         for (k in 1:length(conflict_cols)) {
    #           name_k <- names(conflict_cols)[k]
    #           extra_names_k <- c(
    #             name_k,
    #             paste0(name_k, "_x", 1:(conflict_cols[k] - 1))
    #           )
    #           names(tmp)[tolower(colnames(tmp)) == name_k] <- extra_names_k
    #
    #           warning(paste0(
    #             "Non-standard (project-specific) columns ",
    #             "were found with names matching standard \n  column names ",
    #             "in sheet '", sheets_to_read[i], "'.\n\n  The following ",
    #             "column names were assigned to avoid conflicts:",
    #             "\n    ", paste0(extra_names_k, collapse = ", "), "."
    #           ))
    #         }
    #       } # end if
    #     }
    #   } # end if else
    #
    #   # make column names lowercase
    #   names(tmp) <- tolower(names(tmp))
    #
    #   # set classes; by column name since conflicts resolved above
    #
    #   # character
    #   char_cols <- with(schema_i, name[type == "character"])
    #   for (j in char_cols) tmp[, j] <- as.character(data.frame(tmp)[, j])
    #
    #   # numeric
    #   num_cols <- with(schema_i, name[type == "numeric"])
    #   for (j in num_cols) tmp[, j] <- as.numeric(data.frame(tmp)[, j])
    #
    #   # POSIXct
    #   posixct_cols <- with(schema_i, name[type == "POSIXct"])
    #   for (j in posixct_cols) {
    #     schema_row <- match(j, schema_i$name)
    #
    #     # Get time zone
    #     # function to construct time zone string from reference column tmp
    #     REFCOL <- function(x) {
    #       col_x <- gsub(")$", "", strsplit(x, "REFCOL\\(")[[1]][2])
    #       x2 <- tmp[, col_x]
    #       utc_rows <- tolower(x2) %in% c("utc", "gmt")
    #       x2[utc_rows] <- "UTC"
    #       x2[!utc_rows] <- with(tmp, paste0("US/", x2[!utc_rows]))
    #       return(x2)
    #     }
    #
    #     # get timezone for this column
    #     tz_cmd <- gsub("^tz = |^tz=|\"", "", schema_i$arg[schema_row])
    #
    #     if (grepl("REFCOL", tz_cmd)) {
    #       tzone_j <- REFCOL(tz_cmd)
    #       tz_cmd <- unique(tzone_j)
    #     } else {
    #       tzone_j <- tz_cmd
    #     }
    #
    #     if (nrow(tmp) > 0) {
    #       # Handle mixture of timestamps as date and char
    #
    #       # identify timestamps that can be numeric; assume others character
    #       posix_na <- is.na(tmp[, j]) # identify missing first
    #
    #       if (inherits(tmp[, j], "POSIXct")) {
    #         posix_as_num <- tmp[, j]
    #       } else {
    #         posix_as_num <- suppressWarnings(as.numeric(tmp[, j]))
    #
    #         # convert excel number to posix
    #         posix_as_num <- as.POSIXct("1899-12-30", tz = "GMT") +
    #           (posix_as_num * 86400)
    #       }
    #
    #       posix_as_char <- !posix_na & is.na(posix_as_num)
    #
    #       if (any(posix_as_char)) {
    #         bad_pc_rows <- which(posix_as_char) + 2
    #         bad_pc_rows <- ifelse(length(bad_pc_rows) < 10,
    #           paste0(bad_pc_rows, collapse = ", "),
    #           paste0(paste(bad_pc_rows[1:10], collapse = ", "),
    #             "... +", length(bad_pc_rows) - 10, " more.",
    #             collapse = " "
    #           )
    #         )
    #         warning(paste0(
    #           "Some records (see below) ",
    #           "in '", sheets_to_read[i], "` were not recognized as Excel ",
    #           "datetime objects.\n  These should have imported correctly if ",
    #           "formatted as 'YYYY-MM-DD HH:MM',\n  but see 'Note' in ",
    #           "help(\"read_glatos_workbook\") to avoid this warning.\n\n ",
    #           "Column: '", j, "'\n   Rows:  ", bad_pc_rows, "\n "
    #         ))
    #       }
    #
    #       # handle multiple time zones
    #       for (k in 1:length(tz_cmd)) {
    #         rows_k <- tzone_j %in% tz_cmd[k] # get rows with kth tz
    #         # round to nearest minute and force to correct timezone
    #         posix_as_num[rows_k] <- as.POSIXct(
    #           format(
    #             round(
    #               posix_as_num[rows_k],
    #               "mins"
    #             ),
    #             "%Y-%m-%d %H:%M",
    #           ),
    #           tz = tz_cmd[k]
    #         )
    #
    #         # do same for posix_as_char and insert into posix_as_num
    #         if (any(posix_as_char[rows_k])) {
    #           posix_as_num[posix_as_char & rows_k] <-
    #             as.POSIXct(tmp[posix_as_char & rows_k, j],
    #               tz = tz_cmd[k]
    #             )
    #         }
    #       } # end k
    #
    #       tmp[, j] <- posix_as_num
    #     } else {
    #       tmp[, j] <- as.POSIXct(NA, tz = "UTC")[0]
    #     }
    #
    #     attr(tmp[, j], "tzone") <- "UTC"
    #   } # end j
    #
    #   # Date
    #   date_cols <- with(schema_i, name[type == "Date"])
    #   for (j in date_cols) {
    #     schema_row <- match(j, schema_i$name)
    #
    #
    #     # identify date that can be numeric; assume others character
    #     date_na <- is.na(tmp[, j]) # identify missing
    #
    #     if (inherits(tmp[, j], "POSIXct")) {
    #       date_as_num <- suppressWarnings(as.Date(tmp[, j]))
    #     } else {
    #       # convert excel number to R date
    #       date_as_num <- suppressWarnings(as.numeric(tmp[, j]))
    #       date_as_num <- as.Date("1899-12-30") + date_as_num
    #     }
    #
    #     date_as_char <- !date_na & is.na(date_as_num)
    #
    #     # do same for date_as_char and insert into ate_as_num
    #     if (any(date_as_char)) {
    #       bad_dc_rows <- which(date_as_char) + 2
    #       bad_dc_rows <- ifelse(length(bad_dc_rows) < 10,
    #         paste0(bad_dc_rows, collapse = ", "),
    #         paste0(paste(bad_dc_rows[1:10], collapse = ", "),
    #           "... +", length(bad_dc_rows) - 10, " more.",
    #           collapse = " "
    #         )
    #       )
    #       date_as_num[date_as_char] <- tryCatch(as.Date(tmp[date_as_char, j]),
    #         error = function(e) {
    #           if (e$message == "character string is not in a standard unambiguous format") {
    #             stop(paste0(
    #               "At least one of the records identified below in '",
    #               sheets_to_read[i], "`\n  could not be coerced to Date because ",
    #               "the format was invalid.\n  Dates stored as ",
    #               "text in GLATOS Workbooks must be formatted \n  ",
    #               "'YYYY-MM-DD'. See 'Note' in ",
    #               "help(\"read_glatos_workbook\") \n  about formatting ",
    #               "dates and times ",
    #               "in GLATOS Workbooks.\n\n ",
    #               "Column: '", j, "'\n   Row:  ", bad_dc_rows, "\n"
    #             ))
    #           } else {
    #             return(e)
    #           }
    #         }
    #       )
    #     }
    #
    #     # warn user if no error
    #     if (any(date_as_char)) {
    #       warning(paste0(
    #         "Some records (see below) in '",
    #         sheets_to_read[i], "` were not recognized as Excel date objects.\n",
    #         "  These should have imported correctly if formatted as ",
    #         "'YYYY-MM-DD',\n  but see 'Note' in ",
    #         "help(\"read_glatos_workbook\") to avoid this warning.\n\n ",
    #         "Column: '", j, "'\n    Row:  ", bad_dc_rows, "\n"
    #       ))
    #     }
    #
    #
    #     tmp[, j] <- date_as_num
    #   } # end j
    #
    # } # end i


    if (simplify == FALSE) {
      return(wb)
    }


    # Merge to glatos_workbook list object

    ins_key <- list(
      by.x = c(
        "glatos_project", "glatos_array", "station_no",
        "consecutive_deploy_no", "ins_serial_no"
      ),
      by.y = c(
        "glatos_project", "glatos_array", "station_no",
        "consecutive_deploy_no", "ins_serial_number"
      )
    )

    wb2 <- with(wb, list(
      metadata = project,
      animals = tagging,
      receivers = merge(
        deployment,
        recovery[
          ,
          unique(c(
            ins_key$by.y,
            setdiff(names(recovery), names(deployment))
          ))
        ],
        by.x = c(
          "glatos_project", "glatos_array", "station_no",
          "consecutive_deploy_no", "ins_serial_no"
        ),
        by.y = c(
          "glatos_project", "glatos_array", "station_no",
          "consecutive_deploy_no", "ins_serial_number"
        ),
        all.x = TRUE, all.y = TRUE
      )
    ))

    # add location descriptions
    wb2$receivers <- with(wb2, merge(receivers, wb$locations,
      by = "glatos_array"
    ))

    # Drop unwanted columns from receivers

    # coalesce deploy_date_time and glatos_deploy_date_time
    attr(wb2$receivers$glatos_deploy_date_time, "tzone") <- "UTC"
    ddt_na <- is.na(wb2$receivers$deploy_date_time)
    wb2$receivers$deploy_date_time[ddt_na] <-
      wb2$receivers$glatos_deploy_date_time[ddt_na]

    # coalesce recover_date_time and glatos_recover_date_time
    attr(wb2$receivers$glatos_recover_date_time, "tzone") <- "UTC"
    rdt_na <- is.na(wb2$receivers$recover_date_time)
    wb2$receivers$recover_date_time[rdt_na] <-
      wb2$receivers$glatos_recover_date_time[rdt_na]

    drop_cols_rec <- c(
      "glatos_deploy_date_time", "glatos_timezone",
      "glatos_recover_date_time"
    )
    wb2$receivers <- wb2$receivers[, -match(
      drop_cols_rec,
      names(wb2$receivers)
    )]

    # sort rows by deploy_date_time
    wb2$receivers <- wb2$receivers[with(
      wb2$receivers,
      order(deploy_date_time, glatos_array, station_no)
    ), ]
    row.names(wb2$receivers) <- NULL

    # Drop unwanted columns from animals

    # coalesce release_date_time and utc_release_date_time
    attr(wb2$animals$glatos_release_date_time, "tzone") <- "UTC"
    ardt_na <- is.na(wb2$animals$utc_release_date_time)
    wb2$animals$utc_release_date_time[ardt_na] <-
      wb2$animals$glatos_release_date_time[ardt_na]

    drop_cols_anim <- c("glatos_release_date_time", "glatos_timezone")
    wb2$animals <- wb2$animals[, -match(
      drop_cols_anim,
      names(wb2$animals)
    )]

    # sort animals
    # sort rows by deploy_date_time
    wb2$animals <- wb2$animals[with(
      wb2$animals,
      order(utc_release_date_time, animal_id)
    ), ]
    row.names(wb2$animals) <- NULL

    # create animal_id if missing
    anid_na <- is.na(wb2$animals$animal_id)
    wb2$animals$animal_id[anid_na] <- with(
      wb2$animals[anid_na, ],
      paste0(tag_code_space, "-", tag_id_code)
    )

    # Append new sheets if required
    if (read_all) {
      for (i in 1:length(extra_sheets)) {
        wb2[extra_sheets[i]] <- wb[extra_sheets[i]]
      }
    }
  }

  #-end v1.3, v1.4----------------------------------------------------------------

  # assign classes
  wb2$animals <- glatos::as_glatos_animals(wb2$animals)
  wb2$receivers <- glatos::as_glatos_receivers(wb2$receivers)

  # temporary mapping internal function for external use
  glatos_workbook <- glatos:::glatos_workbook
  wb2 <- glatos_workbook(wb2)

  return(wb2)
}


#' Cast a list of scalars to a new class
#'
#' Cast a list of scalars, with potentially mixed classes, to a new class.
#'
#' @param x A list of scalars.
#'
#' @param new_class A text string with name of new class.
#'
#' @param old_class A character vector with names of classes in `x` that will
#'  be cast to `new_class`. Any record with a different class will result in NA
#'  (with error). Default value is `c("logical", "numeric", "character", "Date",
#'  "POSIXct")`.
#'
#' @param defer_exceptions If TRUE (default value) then errors and warnings will
#'  be returned as attributes with prefix "error_" or "warning_".
#'
#' @param ... Other arguments passed to the casting function (e.g., `tz =
#'   "US/Eastern"` when `new_class` is `POSIXct`.
#'
#' @details
#'  Written specifically for `readxl::read_excel` with `col_types = "list"` to
#'  evaluate class of each record/row independently and then present user with
#'  a single report of all errors (instead of sequential one. at. a. time).
#'
#' @returns A vector of length same as `x` and class as `new_class`.
#'
#' @examples
#'
#' x <- list(TRUE, "A", NA, 3.1415, Sys.time(), Sys.Date(), "1997-05-13 12:43:21")
#'
#' sapply(x, class)
#'
#' cast(x, "character")
#'
#' cast(x, "numeric")
#'
#' cast(x, "Date")
#'
#' cast(x, "POSIXct")
#'
#' cast(x, "POSIXct", tz = "US/Pacific")
#'
#' # separate tz for each element
#' cast(x, "POSIXct", tz = c("US/Eastern", rep("US/Pacific", 5)))
#'
#' # Only cast from if class is character
#' cast(x, "POSIXct", old_class = "character")
#'
#' # Only cast from if class is character or POSIXct
#' cast(x, "character", old_class = c("character", "POSIXct"))
#'
#' \dontrun{
#'
#' # Bad (unsupported) new_class
#' cast(x, "foo")
#'
#' # Bad (unsupported) old_class
#' cast(x, "character", old_class = c("character", "foo"))
#' }
#'
#' @export
cast <- function(x,
                 new_class,
                 old_class = c(
                   "logical",
                   "character",
                   "numeric",
                   "Date",
                   "POSIXct"
                 ),
                 defer_exceptions = TRUE,
                 ...) {
  args_in <- list(...)

  x_class <- sapply(x, function(x) class(x)[1])
  x_na <- sapply(x, is.na)

  # Check if new_class is valid
  # (only default values of old_class are supported)
  invalid_new_class <- setdiff(
    new_class,
    eval(formals(cast)$old_class)
  )

  if (length(invalid_new_class) > 0) {
    stop(
      "Invalid 'new_class': \"",
      paste0(invalid_new_class,
        collapse = ", "
      ), "\""
    )
  }

  # Check if old_class is valid
  # (only default values of old_class are supported)
  invalid_old_class <- setdiff(
    old_class,
    eval(formals(cast)$old_class)
  )

  if (length(invalid_old_class) > 0) {
    stop(
      "Invalid 'old_class': \"",
      paste0(invalid_old_class,
        collapse = ", "
      ), "\""
    )
  }

  # Select casting function; based on new_class
  cast_with <- switch(new_class,
    logical = as.logical,
    numeric = as.numeric,
    character = as.character,
    Date = as.Date,
    POSIXct = as.POSIXct
  )


  # Cast; return NA on error
  x2 <- suppressWarnings(
    lapply(seq_along(x), function(k) {
      if (inherits(x[[k]], old_class)) {
        args_in_k <- args_in
        if (length(args_in_k) > 0) {
          for (m in 1:length(args_in_k)) args_in_k[[m]] <- args_in[[m]][k]
        }
        tryCatch(
          do.call(cast_with, c(
            list(x = x[[k]]),
            args_in_k
          )),
          error = function(e) cast_with(NA)
        )
      } else {
        cast_with(NA)
      }
    })
  )

  # combine
  x2 <- do.call(c, x2)


  # Assign code indicating status of the cast
  cast_status <- data.table::fcase(
    # success; record not cast because old and new class are same
    # or input was NA
    # result is same as input
    (x_class == new_class & x_class %in% old_class) | x_na,
    1, # old and new class same
    # apparent successful cast (not NA before and after)
    !x_na & !is.na(x2),
    2, # apparent successful cast (non-NA result)
    # record not cast; unsupported class (not in old_class)
    # result is NA
    !(x_class %in% old_class),
    3, # unsupported class
    # failed cast (NA after cast; but not before)
    # result in NA
    !x_na & is.na(x2),
    4, # failed cast
    default = NA
  )


  # Handle exceptions (warnings and errors)

  # preallocate vector for each specific error and warning by type

  error_input_class_skipped <- NULL
  error_cast_failed <- NULL
  warning_cast_to_check <- NULL

  # function to print a message that shows only first n records... & last
  print_first_n <- function(x, n = 3) {
    if (length(x) <= n) {
      paste(x, collapse = ", ")
    } else {
      paste0(
        paste(x[1]:min(length(x), n), collapse = ", "),
        ", ...",
        tail(x, 1),
        " (+ ", length(x) - n - 1, " others)"
      )
    }
  }


  # Invalid input class (cast not supported)
  if (any(cast_status == 3)) {
    eiic_index <- which(cast_status == 3)

    plural <- "s"[length(eiic_index) > 1]

    error_input_class_skipped <- c(
      error_input_class_skipped,
      paste0(
        "row", plural, " ",
        print_first_n(eiic_index, 5)
      )
    )
  }

  # Cast failed (bad input value?)
  if (any(cast_status == 4)) {
    ecf_index <- which(cast_status == 4)

    plural <- "s"[length(ecf_index) > 1]

    error_cast_failed <- c(
      error_cast_failed,
      paste0(
        "row", plural, " ",
        print_first_n(ecf_index, 5)
      )
    )
  }


  # Cast succeeded but needs to be checked

  if (any(cast_status == 2)) {
    wctc_index <- which(cast_status == 2)

    plural <- "s"[length(wctc_index) > 1]

    warning_cast_to_check <- c(
      warning_cast_to_check,
      paste0(
        "row", plural, " ",
        print_first_n(wctc_index, 5)
      )
    )
  }

  if (!defer_exceptions) {
    if (length(warning_cast_to_check) > 0) {
      warning(
        "\nCast was successful but should be verified: ",
        paste0(warning_cast_to_check, collapse = "\n ")
      )
    }

    if (length(error_input_class_skipped) > 0 |
      length(error_cast_failed) > 0) {
      stop(c(
        paste0(
          "\n\nInvalid input class: ",
          paste(error_input_class_skipped, collapse = "\n")
        ),
        paste0(
          "\n\nCast failed: ",
          paste(error_cast_failed, collapse = "\n")
        )
      ))
    }
  }

  return(structure(x2,
    cast_status = cast_status
  ))
}


#' Identify and check GLATOS workbook file version
#'
#' Identify and check version of a GLATOS workbook file (xlsm or xlsx) based on
#' its structure.
#'
#' @inheritParams read_glatos_workbook
#'
#' @returns A character string with version number ("1.3", "1.4").
#'
#' @examples
#'
#' # Example 1: Version 1.3 (xlsm)
#'
#' # get path to example GLATOS Data Workbook
#' wb_file <- system.file("extdata",
#'   "walleye_workbook.xlsm",
#'   package = "glatos"
#' )
#'
#' identify_workbook_version(wb_file)
#'
#'
#' # Example 2: Version 1.4 (xlsx)
#'
#' wb2_file <- system.file("extdata",
#'   "walleye_workbook.xlsx",
#'   package = "glatos"
#' )
#'
#' identify_workbook_version(wb2_file)
#'
#' @export
identify_workbook_version <- function(wb_file,
                                      wb_version = NULL) {
  # Get sheet names
  sheets <- tolower(readxl::excel_sheets(wb_file))


  # v1.3 and v1.4 have same sheet names
  if (all(names(glatos_workbook_schema$v1.3) %in% sheets)) {
    # Identify version based on structure of Project sheet
    # Get project data
    prj <- tryCatch(
      readxl::read_excel(wb_file,
        sheet = "Project",
        col_names = FALSE,
        .name_repair = "minimal"
      ),
      error = function(e) {
        if (e$message ==
          "Expecting a single string value: [type=character; extent=0].") {
          stop(
            "There was a problem reading from input file specified. It may ",
            "be protected.\n  Try again after opening, saving, and closing the ",
            "file."
          )
        } else {
          stop(e)
        }
      }
    )

    if (all(prj[[1]][1:3] == c(
      "GLATOS Project Code:",
      "Principal Investigator (PI):",
      "PI E-Mail Address:"
    ))) {
      if (ncol(prj) > 2 & nrow(prj) > 3) {
        if (prj[[3]][5] == "DataForm Version 1.4") {
          wbv <- "1.4"
        }
      } else {
        wbv <- "1.3"
      }
    } else {
      wbv <- NULL
    }
  } else {
    wbv <- NULL
  }

  if (is.null(wbv)) {
    stop(
      "Workbook version could not be identified. Double check ",
      "that you are using a standard GLATOS Workbook file. The ",
      "names of sheets must match standard file.",
      call. = FALSE
    )
  }

  # Check against input, if given

  if (is.null(wb_version) == FALSE) {
    if (wb_version != wbv) {
      stop("Workbook version looks like ",
        wbv, ", which is different than input ",
        "'wb_version' (", wb_version, ").",
        call. = FALSE
      )
    }
  }

  return(wbv)
}


#' Read Project sheet from GLATOS workbook file
#'
#' Read Project sheet from GLATOS workbook file
#'
#' @inheritParams read_glatos_workbook
#'
#' @returns A list with six elements:
#' 1. __project_code:__ GLATOS Project Code.
#' 2. __principle_investigator:__ Name of Principle Investigator.
#' 3. __pi_email:__ Email address of Principle Investigator.
#' 4. __source_file:__ Name of input `wb_file`.
#' 5. __wb_version:__ Version of GLATOS workbook file.
#' 6. __created:__ Timestamp when source file metadata was changed
#'  (from `file.info(wb_file)$ctime`).
#'
#' @examples
#'
#' # Example 1: Version 1.3 (xlsm)
#'
#' # get path to example GLATOS Data Workbook
#' wb_file <- system.file("extdata",
#'   "walleye_workbook.xlsm",
#'   package = "glatos"
#' )
#'
#' read_workbook_project(wb_file)
#'
#'
#' # Example 2: Version 1.4 (xlsx)
#'
#' wb2_file <- system.file("extdata",
#'   "walleye_workbook.xlsx",
#'   package = "glatos"
#' )
#'
#' read_workbook_project(wb2_file)
#'
#' @export
read_workbook_project <- function(wb_file) {
  # Get workbook version
  wb_version <- identify_workbook_version(wb_file)

  # Get project data
  prj_raw <- tryCatch(
    readxl::read_excel(wb_file,
      sheet = "Project",
      col_names = FALSE,
      .name_repair = "minimal"
    ),
    error = function(e) {
      if (e$message ==
        "Expecting a single string value: [type=character; extent=0].") {
        stop(
          "There was a problem reading from input file specified. It may ",
          "be protected.\n  Try again after opening, saving, and closing the ",
          "file."
        )
      } else {
        stop(e)
      }
    }
  )

  # Restructure as list
  prj <- list(
    project_code = prj_raw[[2]][1],
    principle_investigator = prj_raw[[2]][2],
    pi_email = prj_raw[[2]][3],
    source_file = basename(wb_file),
    wb_version = wb_version,
    created = file.info(wb_file)$ctime
  )

  return(prj)
}
