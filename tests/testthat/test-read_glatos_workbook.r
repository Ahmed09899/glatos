test_that("metadata element gives expected result; v1.3", {
  wb_file <- system.file("extdata", "walleye_workbook.xlsm", package = "glatos")
  wb <- read_glatos_workbook(wb_file)

  expect_type(wb[["metadata"]], "list")

  # Check if expected and actual results are the same
  expect_equal(wb[["metadata"]][1:5], walleye_workbook[["metadata"]][1:5])
})

test_that("metadata element gives expected result; v1.4", {
  wb_file <- system.file("extdata", "walleye_workbook.xlsx", package = "glatos")
  wb <- read_glatos_workbook(wb_file)
  
  expect_type(wb[["metadata"]], "list")
  
  # Check if expected and actual results are the same
  # skip elements 4-6 (source_file, wb_version, created) because 
  # the data object is format v.1.3.
  expect_equal(wb[["metadata"]][c(1:3)], 
               walleye_workbook[["metadata"]][c(1:3)])
})

test_that("receivers element gives expected result; v1.3", {
  wb_file <- system.file("extdata", "walleye_workbook.xlsm", package = "glatos")
  wb <- read_glatos_workbook(wb_file)

  expect_s3_class(wb[["receivers"]], "glatos_receivers")
  expect_s3_class(wb[["receivers"]], "data.frame")

  # Check if expected and actual results are the same
  expect_equal(wb[["receivers"]], walleye_workbook[["receivers"]])
})

test_that("receivers element gives expected result; v1.4", {
  wb_file <- system.file("extdata", "walleye_workbook.xlsx", package = "glatos")
  wb <- read_glatos_workbook(wb_file)
  
  expect_s3_class(wb[["receivers"]], "glatos_receivers")
  expect_s3_class(wb[["receivers"]], "data.frame")
  
  # Check if expected and actual results are the same
  expect_equal(wb[["receivers"]], walleye_workbook[["receivers"]])
})

test_that("animals gives expected result; v1.3", {
  wb_file <- system.file("extdata", "walleye_workbook.xlsm", package = "glatos")
  wb <- read_glatos_workbook(wb_file)

  expect_s3_class(wb[["animals"]], "glatos_animals")
  expect_s3_class(wb[["animals"]], "data.frame")

  # Check if expected and actual results are the same
  expect_equal(wb[["animals"]], walleye_workbook[["animals"]])
})

test_that("animals gives expected result; v1.4", {
  wb_file <- system.file("extdata", "walleye_workbook.xlsm", package = "glatos")
  wb <- read_glatos_workbook(wb_file)
  
  expect_s3_class(wb[["animals"]], "glatos_animals")
  expect_s3_class(wb[["animals"]], "data.frame")
  
  # Check if expected and actual results are the same
  expect_equal(wb[["animals"]], walleye_workbook[["animals"]])
})
