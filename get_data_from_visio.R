# Install required package if you haven't already:
# install.packages("xml2")

########## basic strategy ######################################################
# Since a Visio (.vsdx) file is essentially a ZIP archive containing XML files,
# you can use R’s built‐in functions along with packages like xml2 to extract and parse the contents.
# Here’s an example workflow that shows how to start by processing one file from your visio_files folder:
#
# Unzip the file
# Use R’s unzip() function to extract the file into a temporary directory.
#
# Locate and Parse the XML Page
# Visio files usually store pages under the visio/pages/ folder. You can list those XML files and then parse one of them with xml2.
#
# Extract Data
# Once the XML is loaded, you can use xml2 functions (like xml_find_all() and xml_text()) to extract text from <Shape> elements and,
# if needed, handle the <Connects> elements.

# All clear!
rm(list = ls())

# Load necessary package
library(xml2)
library(openxlsx)

# Define the extraction function
extract_visio_data <- function(vsdx_file) {
  # Create a temporary directory for unzipping
  tmp_dir <- tempfile("vsdx_")
  dir.create(tmp_dir)

  # Unzip the file into the temporary directory
  unzip(vsdx_file, exdir = tmp_dir)

  # Locate the pages directory
  pages_dir <- file.path(tmp_dir, "visio", "pages")
  if (!dir.exists(pages_dir)) {
    stop("No 'visio/pages' folder found in the vsdx file: ", vsdx_file)
  }

  # List XML files in the pages folder (excluding pages.xml)
  page_files <- list.files(pages_dir, pattern = "\\.xml$", full.names = TRUE)
  page_files <- page_files[!grepl("pages\\.xml$", page_files)]
  if (length(page_files) == 0) {
    stop("No page XML files found in file: ", vsdx_file)
  }

  # For demonstration, pick the first page file
  page_file <- page_files[1]
  cat("Processing page file:", page_file, "\n")

  # Read and parse the XML page
  page_xml <- read_xml(page_file)

  # Handle namespaces (the default namespace is assigned a prefix, often "d1")
  ns <- xml_ns(page_xml)

  # Extract all <Shape> elements using a namespace-aware XPath
  shapes <- xml_find_all(page_xml, ".//d1:Shape", ns)
  cat("Found", length(shapes), "Shape element(s) in", vsdx_file, "\n")

  # Loop over each shape and extract attributes and text into a list of data frames
  shape_data <- lapply(shapes, function(shape) {
    # Extract attributes from the <Shape> element
    id <- xml_attr(shape, "ID")
    name <- xml_attr(shape, "Name")
    nameu <- xml_attr(shape, "NameU")

    # Extract the text from the first <Text> child element (if available)
    text_node <- xml_find_first(shape, ".//d1:Text", ns)
    text_content <- if (!is.na(text_node)) xml_text(text_node) else NA_character_

    # Return as a one-row data frame
    data.frame(
      ID = id,
      Name = name,
      NameU = nameu,
      Text = text_content,
      stringsAsFactors = FALSE
    )
  })

  # Combine all rows into a single data frame
  shapes_df <- do.call(rbind, shape_data)

  # Clean up temporary files
  unlink(tmp_dir, recursive = TRUE)
  cat("Temporary files cleaned up for", vsdx_file, "\n\n")

  # Return the resulting data frame
  return(shapes_df)
}

# Get a list of all .vsdx files in the visio_files folder
vsdx_files <- list.files("visio_files", pattern = "\\.vsdx$", full.names = TRUE)

# Initialize an empty list to store data frames
visio_data_list <- list()

# Loop through each file, run the extraction function, and store the result in the list
for (file in vsdx_files) {
  # Get the base name (without extension) to use as the list name
  base_name <- tools::file_path_sans_ext(basename(file))
  cat("Extracting data for:", base_name, "\n")

  # Extract data from the current file
  df <- extract_visio_data(file)

  # Store the data frame in the list using the base name
  visio_data_list[[base_name]] <- df
}

rm(df)

# Optionally, print the names of the data frames stored in the list
cat("Data frames stored in the list for the following files:\n")
print(names(visio_data_list))

# Now you can access each data frame by name from visio_data_list, for example:
# View(visio_data_list[["vs5"]])

# Create a new workbook
wb <- createWorkbook()

# Loop over each element in the visio_data_list
for(sheet_name in names(visio_data_list)) {
  # Add a worksheet with the sheet name (dataframe name)
  addWorksheet(wb, sheet_name)

  # Write the data frame into the worksheet
  writeData(wb, sheet = sheet_name, visio_data_list[[sheet_name]])
}

# Save the workbook to an Excel file
saveWorkbook(wb, "output/visio_data.xlsx", overwrite = TRUE)

