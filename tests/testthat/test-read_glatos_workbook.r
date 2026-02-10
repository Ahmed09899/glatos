
# Check workbook version 1.3

wb13_file <- system.file("extdata", "walleye_workbook.xlsm", package = "glatos")
wb13 <- read_glatos_workbook(wb13_file)

test_that("metadata element gives expected result; v1.3", {

  expect_type(wb13[["metadata"]], "list")

  # Check if expected and actual results are the same
  expect_equal(wb13[["metadata"]][1:5], walleye_workbook[["metadata"]][1:5])
})

test_that("receivers element gives expected result; v1.3", {

  expect_s3_class(wb13[["receivers"]], "glatos_receivers")
  expect_s3_class(wb13[["receivers"]], "data.frame")

  # Check if expected and actual results are the same
  expect_equal(wb13[["receivers"]], walleye_workbook[["receivers"]])
})

test_that("animals gives expected result; v1.3", {

  expect_s3_class(wb13[["animals"]], "glatos_animals")
  expect_s3_class(wb13[["animals"]], "data.frame")

  # Check if expected and actual results are the same
  expect_equal(wb13[["animals"]], walleye_workbook[["animals"]])
})


# optional arg. read_all = TRUE
wb13a <- read_glatos_workbook(wb13_file, read_all = TRUE)

test_that("receivers element gives expected result; v1.3 (read_all = TRUE)", {
  
  expect_s3_class(wb13a[["receivers"]], "glatos_receivers")
  expect_s3_class(wb13a[["receivers"]], "data.frame")
  
  # Check if project-specific columns are present
  expect_true("receiver" %in% names(wb13a[["receivers"]]))
})

test_that("animals gives expected result; v1.3 (read_all = TRUE)", {
  
  expect_s3_class(wb13a[["animals"]], "glatos_animals")
  expect_s3_class(wb13a[["animals"]], "data.frame")
  
  # Check if project-specific columns are present
  expect_true(all(c("release_year", "release_month", "release_doy") %in% 
                names(wb13a[["animals"]])))
})


# Check workbook version 1.4

wb14_file <- system.file("extdata", "walleye_workbook.xlsx", package = "glatos")
wb14 <- read_glatos_workbook(wb14_file)

test_that("metadata element gives expected result; v1.4", {

  expect_type(wb14[["metadata"]], "list")
  
  # Check if expected and actual results are the same
  # skip elements 4-6 (source_file, wb_version, created) because
  # the data object is format v.1.3.
  expect_equal(
    wb14[["metadata"]][c(1:3)],
    walleye_workbook[["metadata"]][c(1:3)]
  )
})

test_that("receivers element gives expected result; v1.4", {

  expect_s3_class(wb14[["receivers"]], "glatos_receivers")
  expect_s3_class(wb14[["receivers"]], "data.frame")
  
  # Check if expected and actual results are the same
  expect_equal(wb14[["receivers"]], walleye_workbook[["receivers"]])
})

test_that("animals gives expected result; v1.4", {

    expect_s3_class(wb14[["animals"]], "glatos_animals")
  expect_s3_class(wb14[["animals"]], "data.frame")
  
  # Check if expected and actual results are the same
  expect_equal(wb14[["animals"]], walleye_workbook[["animals"]])
})


# optional arg. read_all = TRUE
wb14a <- read_glatos_workbook(wb14_file, read_all = TRUE)

test_that("receivers element gives expected result; v1.4 (read_all = TRUE)", {
  
  expect_s3_class(wb14a[["receivers"]], "glatos_receivers")
  expect_s3_class(wb14a[["receivers"]], "data.frame")
  
  # Check if project-specific columns are present
  expect_true("receiver" %in% names(wb14a[["receivers"]]))
})

test_that("animals gives expected result; v1.4 (read_all = TRUE)", {
  
  expect_s3_class(wb14a[["animals"]], "glatos_animals")
  expect_s3_class(wb14a[["animals"]], "data.frame")
  
  # Check if project-specific columns are present
  expect_true(all(c("release_year", "release_month", "release_doy") %in% 
                    names(wb14a[["animals"]])))
})