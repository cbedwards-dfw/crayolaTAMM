## helper function: convert vector to numeric
convert_cols_to_num <- function(x) {
  x[x == "" | tolower(x) == "na" | tolower(x) == "n/a"] <- NA_real_
  as.numeric(x)
}


#helper function: is a simple way of telling which columns should be processed Marked, unmarked, total, or stock
#Now using strings instead of column numbers
get_col_index <- function(stock) {
  switch(stock,
         "Still" = c("Marked Impacts", "Unmarked Impacts"),
         # "Nooksack" = 3,
         "Nooksack" = "Unmarked Impacts",
         # "Skagit S/F" = 4,
         "Skagit S/F" = "Cumulative Impacts",
         # "Skykomish" = 4,
         "Skykomish" = "Cumulative Impacts",
         # "White" = 4,
         "White" = "Cumulative Impacts",
         # "Skagit Sp" = 4,
         "Skagit Sp" = "Cumulative Impacts",
         # "Dung" = 2,
         "Dung" = "Stock",
         # "Nisq" = 2,
         "Nisq" = "Stock",
         # "Skok"=4,
         "Skok"="Cumulative Impacts",
         # "Lake" = 3,
         "Lake" = "Unmarked Impacts",
         # "Green" = 3,
         "Green" = "Unmarked Impacts",
         # "Puyallup"=3,
         "Puyallup"="Unmarked Impacts",
         "Mid"="Cumulative Impacts",
         # 2
         "Stock"
  )
}

# Apply conditional formatting
apply_diff_formatting <- function(rows, wb, sheet_name, red_style, green_style) {
  for (r in rows) {
    openxlsx::conditionalFormatting(wb, sheet = sheet_name, cols = 2:4, rows = r,
                          rule = "<0", style = red_style)
    openxlsx::conditionalFormatting(wb, sheet = sheet_name, cols = 2:4, rows = r,
                          rule = ">0", style = green_style)
  }
}
