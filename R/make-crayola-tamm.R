#' Make workbook summary of Chinook TAMM
#'
#' @param filepath Filepath of TAMM, including `.xlsx` suffix
#' @param output_path Filepath for resulting file. Optional; if not provided, will save in same folder as `filepath` with
#' name of "\{filepath\}_Impacts_crayola.xlsx"
#' @param plot_width Controls width of barplots in stock sheets. Numeric, defaults to 10
#' @param plot_height Controls height of barplots in stock sheets. Numeric, defaults to 8
#' @param plot_units Character string defining units of `plot_width` and `plot_height`. Defaults to "in", can also be "cm" or "px".
#' @param verbose Should progress messages and saved filepath be printed to screen?
#'
#' @return Invisibly returns filepath of the summary file.
#' @export
#'
#' @examples
#' \dontrun{
#' make_crayola_tamm("Chin2225.xlsx")
#' }
#'
#'
make_crayola_tamm <- function(filepath,
                              output_path = NULL,
                              plot_width = 10,
                              plot_height = 8,
                              plot_units = "in",
                              verbose = TRUE) {
  # Input validation
  if (!(is.character(filepath) && length(filepath) == 1)) {
    cli::cli_abort("{.arg filepath} must be a single character string.")
  }

  if (!file.exists(filepath)) {
    cli::cli_abort("File {.path {filepath}} does not exist.")
  }

  if (!(is.null(output_path) || (is.character(output_path) && length(output_path) == 1))) {
    cli::cli_abort("{.arg output_path} must be NULL or a single character string.")
  }

  if (!(is.numeric(plot_width) && length(plot_width) == 1 && plot_width > 0)) {
    cli::cli_abort("{.arg plot_width} must be a positive number.")
  }

  if (!(is.numeric(plot_height) && length(plot_height) == 1 && plot_height > 0)) {
    cli::cli_abort("{.arg plot_height} must be a positive number.")
  }

  if (!(is.character(plot_units) && length(plot_units) == 1 && plot_units %in% c("in", "cm", "px"))) {
    cli::cli_abort("{.arg plot_units} must be one of 'in', 'cm', or 'px'.")
  }

  ## if output_path is not provided, just put it in same folder as filepath, with "Impacts_crayola" suffix
  if (is.null(output_path)) {
    output_path <- paste0(tools::file_path_sans_ext(filepath), "_Impacts_crayola.xlsx")
  }

  ## reading in sections
  TAMM <- TAMMsupport::read_limiting_stock(filepath)
  Thresholdsnew <- TAMMsupport::read_overview_complete(filepath)
  twoa_sheets <- TAMMsupport::read_2a_sheets(filepath)
  JDF_new <- TAMMsupport::read_jdf(filepath)
  JDF_unmarked <- TAMMsupport::read_jdf_unmarked(filepath)

  temp_plot_paths <- NULL

  ## cleanup names to be consistent, make sure mid-canal has appropriate er_ceiling label
  Thresholdsnew <- Thresholdsnew |>
    dplyr::mutate(er_ceiling = dplyr::if_else(.data$primary_stock == "Mid-Hood Canal",
                                              "ESC Change",
                                              .data$er_ceiling
    )) |>
    ## some finangling to get stock names aligned into `stock` column
    dplyr::mutate(stock = dplyr::if_else(.data$stock == "Total",
                                         .data$primary_stock,
                                         .data$stock
    )) |>
    ## some spot replacements
    dplyr::mutate(stock = dplyr::if_else(.data$stock == "Skagit" & .data$season == "spring/early",
                                         "Skagit Sp",
                                         .data$stock
    )) |>
    dplyr::mutate(stock = dplyr::if_else(.data$stock == "Skagit" & .data$season == "summer/fall",
                                         "Skagit S/F",
                                         .data$stock
    )) |>
    dplyr::mutate(stock = dplyr::if_else(.data$primary_stock == "Stillaguamish" & .data$stock == "Unmarked ER",
                                         "Stillaguamish Unmarked",
                                         .data$stock
    )) |>
    dplyr::mutate(stock = dplyr::if_else(.data$primary_stock == "Stillaguamish" & .data$stock == "Marked ER",
                                         "Stillaguamish Marked",
                                         .data$stock
    )) |>
    dplyr::mutate(stock = dplyr::if_else(.data$stock == "Lake WA (Cedar R.)",
                                         "Lake WA",
                                         .data$stock
    ))

  ## Parse stock thresholds for later subsetting
  stock_thresholds <- Thresholdsnew |>
    dplyr::filter(!is.na(.data$er_ceiling)) |>
    dplyr::select(
      "stock_name" = "stock",
      "er_ceiling",
      "er_type"
    )

  wb <- openxlsx::createWorkbook()


  ## dataframe organizing our various labels
  ## `$tamm_2ac_label` is for the tamm column name as produced from read_2ac_sheets()
  ## The `2a_cu&m_n` sheet uses a different naming scheme, represented in the *_2 column
  stock_data <- dplyr::tribble(
    ~stocks, ~stocks_long, ~label, ~tamm_2ac_label, ~tamm_2ac_label_2, ~jdf_label,
    "Nooksack", "Nooksack Earlies", "Nooksack", "nooksack early", "nooksack early", NA,
    "Skagit Sp", "Skagit Springs", "Skagit Sp", "skagit spr nat", "skagit spr natural", NA,
    "White", "White Spring Fing", "White", "white r. spring", "white r. spring", NA,
    "Dung", "Dungeness Spring", "Dungeness", "elwha/dungeness s/f n&h", "dung/elwha natural", "Dungeness (N)",
    "Skagit S/F", "Skagit SF","Skagit SF", "skagit s/f nat", "skagit sf natural",  NA,##is fourth item correct??? column is either "n" or "hat" depending on the sheet
    "Still", "Stillaguamish", "Stillaguamish", "stillag. sum/fall", "stillag. sum/fall", NA,
    "Lake", "Lake Washingon", "Lake WA", "lake wash.", "lake wash.", NA,
    "Green", "Green", "Green", "green river", "green river", NA,
    "Puyallup", "Puyallup", "Puyallup", "puyallup river", "puyallup river", NA,
    "Nisq", "Nisqually", "Nisqually", "nisq. river", "nisq. river", NA,
    "Hoko", "Hoko", "Strait-Hoko", "hoko s/f n&h", "hoko natural", "Hoko (N)", ## this is labeled as nat and hatch in all sheets but 2a_cu_m_n. Is this good?
    "Elwha", "Elwha", "Elwha", "elwha/dungeness s/f n&h", "dung/elwha natural", "Elwha (N+H)",
    "Mid", "Mid-HC","Mid-Hood Canal",  "mid-hdcnl (12b) natural", "mid-hdcnl (12b) natural", NA,
    "Skok", "Skokomish", "Skokomish", "skok. r natural", "skok. r natural", NA,
    "Skykomish", "Skykomish", "Snohomish", "skykomish s/f nat", "skykomish natural", NA
  )
  stock_data$color <- grDevices::rainbow(nrow(stock_data))

  component_list <- list()
  for (i in 1:nrow(stock_data)) { # START Primary Loop
    stock <- stock_data$stocks[[i]]
    stock_long <- stock_data$stocks_long[[i]]
    if(verbose){cli::cli_alert("Processing {stock_long}")}

    Impacts <- TAMM |>
      dplyr::select("stock_type", "FisheryID", "Fishery", dplyr::starts_with(stock_long)) |>
      dplyr::select("stock_type", "FisheryID", "Fishery", dplyr::ends_with("aeq")) |>
      dplyr::filter(.data$stock_type %in% c("UM_H", "UM_N", "AD_H", "AD_N")) |>
      tidyr::pivot_longer(cols = dplyr::ends_with("aeq")) |>
      tidyr::pivot_wider(names_from = .data$stock_type, values_from = .data$value) |>
      dplyr::mutate(
        UM_Total = .data$UM_H + .data$UM_N,
        M_Total = .data$AD_H + .data$AD_N
      ) |>
      dplyr::select(-.data$AD_H, -.data$AD_N, -.data$UM_H, -.data$UM_N) |>
      dplyr::mutate(name = gsub(paste0(stock_long, "_"), "", .data$name)) |>
      tidyr::pivot_wider(names_from = .data$name,
                         values_from = c("UM_Total", "M_Total")) |>
      dplyr::select(
        FisheryName = .data$Fishery,
        .data$FisheryID,
        MarkedT2 = .data$M_Total_t2_aeq,
        MarkedT3 = .data$M_Total_t3_aeq,
        MarkedT4 = .data$M_Total_t4_aeq,
        UnmarkedT2 = .data$UM_Total_t2_aeq,
        UnmarkedT3 = .data$UM_Total_t3_aeq,
        UnmarkedT4 = .data$UM_Total_t4_aeq
      ) |>
      ## Cut out "Fishery" of escapement
      dplyr::filter(.data$FisheryName != "Escapement")

    ## Special rules for Treaty 3/4 troll (id 17), since timestep 4 has a different "region" designation.
    ## Save info for fishery 17 separately, then NA the timestep4 entries for fishery 17 in `Impacts`
    Impacts_17 <- Impacts |>
      dplyr::filter(
        .data$FisheryID == 17
      )
    Impacts <- Impacts |>
      dplyr::mutate(
        UnmarkedT4 = ifelse(.data$FisheryID == 17, NA, .data$UnmarkedT4),
        MarkedT4 = ifelse(.data$FisheryID == 17, NA, .data$MarkedT4)
      )


    ## Make pretty version of Impacts for printing.
    ## It looks like there is no further use of `Impacts`, so we could arguably simplify by just continuing to use "Impacts" in place of "Impacts_Print
    ## But it doesn't hurt to leave it as is
    Impacts_Print <- Impacts
    Impacts_Print <- Impacts_Print |> # produce artificial classes to catagorize impacts later
      dplyr::mutate(Class = dplyr::case_when(
        grepl("^SEAK", .data$FisheryName) ~ "SEAK", # Starts with "SEAK"
        .data$FisheryID %in% c("4", "5", "6", "7", "8", "9", "10", "11", "12", "13-15") ~ "Canada", ## id canada fisheries
        .data$FisheryID %in% as.character(16:35) &
          grepl("^Tr", .data$FisheryName, ignore.case = TRUE) ~ "Tribal Ocean", # Starts with "Tr"
        .data$FisheryID %in% as.character(16:35) &
          grepl("^NT|Buoy|Ar 1|Ar 3:4 Spt|Ar 2 Sport|NoWACstNet", .data$FisheryName, ignore.case = TRUE) ~ "WDFW Ocean",
        .data$FisheryID %in% as.character(16:35) &
          grepl("Cal|OR|KMZ", .data$FisheryName) ~ "SOF Ocean",
        .data$FisheryID %in% as.character(36:71) &
          grepl("^NT|Sport|Spt|Sprt|NT10:11Net|NT 7BCDNet", .data$FisheryName) ~ "WDFW P.Sound",
        .data$FisheryID %in% as.character(36:71) &
          grepl("Tr", .data$FisheryName) ~ "Tribal P.Sound",
        .data$FisheryID %in% as.character(72) &
          grepl("Sport", .data$FisheryName) ~ "WDFW River",
        .data$FisheryID %in% as.character(73) &
          grepl("Net", .data$FisheryName) ~ "Tribal River",
        .data$FisheryID %in% as.character(17) &
          grepl("^Tr", .data$FisheryName, ignore.case = TRUE) ~ "Tribal OCEAN[T2:3] & Sound[T4]",
        TRUE ~ "Other"
      ))

    ## update object "Impacts_17" to include class
    Impacts_17$Class <- "Tribal P.Sound" # this is because fishery 17 is in the ocean for timestep 2 and 3 and in puget sound for timestep 4
    Impacts_17 <- Impacts_17 |>
      dplyr::mutate(
        UnmarkedT2 = NA,
        MarkedT2 = NA,
        UnmarkedT3 = NA,
        MarkedT3 = NA,
        UnmarkedT4 = .data$UnmarkedT4,
        MarkedT4 = .data$MarkedT4
      )
    Impacts_Print <- rbind(Impacts_Print, Impacts_17)

    ## To check classifications:
    ## Add columns of total, total marked, total unmarked.
    ## `pick()` is a way to create a subsetted dataframe by selecting multiple columns within a
    ##     data masking function like mutate() or summarize()
    Impacts_Print <- Impacts_Print |>
      dplyr::mutate(
        Total = rowSums(
          dplyr::pick(
            .data$MarkedT2, .data$MarkedT3, .data$MarkedT4,
            .data$UnmarkedT2, .data$UnmarkedT3, .data$UnmarkedT4
          ),
          na.rm = T
        ),
        TotalMarked = rowSums(dplyr::pick(.data$MarkedT2, .data$MarkedT3, .data$MarkedT4),
                              na.rm = T
        ),
        TotalUnmarked = rowSums(dplyr::pick(.data$UnmarkedT2, .data$UnmarkedT3, .data$UnmarkedT4),
                                na.rm = T
        )
      ) |>
      dplyr::arrange(dplyr::desc(.data$Total)) |>
      ## Round to 2 digits for our counts of fish
      dplyr::mutate(dplyr::across(
        c(.data$MarkedT2, .data$MarkedT3, .data$MarkedT4,
          .data$UnmarkedT2, .data$UnmarkedT3, .data$UnmarkedT4,
          .data$Total, .data$TotalMarked, .data$TotalUnmarked),
        \(x) round(x, digits = 2)
      ))


    ## Want to have a semi-long-form, with a row for each timestep. Going to use pivot longer then pivot wider.
    long_df <- Impacts_Print |> # turn into a long df where each fishery has 3 timesteps (2,3,4)
      tidyr::pivot_longer(
        cols = dplyr::starts_with("Marked") | dplyr::starts_with("Unmarked"),
        names_to = c("Type", "TimeStep"),
        names_pattern = "(Marked|Unmarked)(T[2-4])",
        values_to = "Count"
      ) |>
      dplyr::filter(!is.na(.data$Count)) |>
      tidyr::pivot_wider(
        names_from = .data$Type,
        values_from = .data$Count
      ) |>
      ## make sure marked and unmarked have 0s for NAs
      dplyr::mutate(
        Total = dplyr::coalesce(.data$Marked, 0) + dplyr::coalesce(.data$Unmarked, 0),
        TimeStepLabel = paste(.data$FisheryName, .data$TimeStep)
      ) |>
      dplyr::select("FisheryName", "FisheryID", "TimeStep", "TimeStepLabel", "Class", "Marked", "Unmarked", "Total")


    # remove fisheries that have no impact on stock
    Impacts_Print <- Impacts_Print |>
      dplyr::filter(.data$Total != 0) |>
      dplyr::select(
        "FisheryName", "FisheryID", "Class", "MarkedT2",
        "MarkedT3", "MarkedT4", "UnmarkedT2", "UnmarkedT3", "UnmarkedT4", "Total",
        "TotalMarked", "TotalUnmarked"
      )

    sheet_name <- stock_data$label[i]
    sheet_name_save <- sheet_name
    openxlsx::addWorksheet(wb, sheet_name, tabColour = stock_data$color[i]) # create a sheet, give it a rainbow color because you can
    ## It looks like these should already be numerics -- can proably cut some of this.
    Impacts_Print <- as.data.frame(Impacts_Print) |>
      dplyr::mutate(dplyr::across(
        c("MarkedT2", "MarkedT3", "MarkedT4", "UnmarkedT2", "UnmarkedT3",
          "UnmarkedT4"),
        as.numeric
      ))
    openxlsx::writeDataTable(wb, sheet = sheet_name, x = Impacts_Print)

    # Conditional formatting for columns 4 to 6 to highlight highest impacts Marked
    for (col in grep("^Marked", names(Impacts_Print))) {
      openxlsx::conditionalFormatting(
        wb,
        sheet = sheet_name,
        cols = col,
        ## reminder: row 1 of the sheet is the column names, so add 1 to rows
        rows = 1:(nrow(Impacts_Print)) + 1,
        type = "colorScale",
        style = c("#f2e085", "#e0a31b") # light gold to goldenrod
      )
    }

    # Conditional formatting for columns 7 to 9 to highlight highest impacts Unmarked
    for (col in grep("^UnMarked", names(Impacts_Print))) {
      openxlsx::conditionalFormatting(
        wb,
        sheet = sheet_name,
        cols = col,
        ## reminder: row 1 of the sheet is the column names, so add 1 to rows
        rows = 1:(nrow(Impacts_Print)) + 1,
        type = "colorScale",
        style = c("#aed6f1", "#2e86c1") # light blue to steel blue
      )
    }

    # #grabbing stock threshold/ER
    Stock_Thresh <- stock_thresholds |>
      dplyr::filter(grepl(.env$stock, .data$stock_name))


    # apply the function to all the neceissary and preloaded tabs
    if(stock_data$tamm_2ac_label[i] == "skagit s/f nat"){
      ## special case for skagit, which has a different column name in the TAMM
      ##   for marked fish.
      Mrkd_filtered <- twoa_sheets$aeq |>
        dplyr::filter(.data$sheet == "2A_Cmrkd") |>
        dplyr::filter(.data$stock == "skagit s/f h") |>
        dplyr::select("fishery_joint_label", "value")
    } else {
      Mrkd_filtered <- twoa_sheets$aeq |>
        dplyr::filter(.data$sheet == "2A_Cmrkd") |>
        dplyr::filter(.data$stock == stock_data$tamm_2ac_label[i]) |>
        dplyr::select("fishery_joint_label", "value")
    }

    Unmrkd_filtered <- twoa_sheets$aeq |>
      dplyr::filter(.data$sheet == "2A_CUnmrkd") |>
      dplyr::filter(.data$stock == stock_data$tamm_2ac_label[i]) |>
      dplyr::select("fishery_joint_label", "value")

    Cumu_filtered <- twoa_sheets$aeq |>
      dplyr::filter(.data$sheet == "2A_CU&M") |>
      dplyr::filter(.data$stock == stock_data$tamm_2ac_label[i]) |>
      dplyr::select("fishery_joint_label", "value")

    Cumu_natural_filtered <- twoa_sheets$aeq |>
      dplyr::filter(.data$sheet == "2A_CU&M_N") |>
      dplyr::filter(.data$stock == stock_data$tamm_2ac_label[i]) |>
      dplyr::select("fishery_joint_label", "value")

    ## Remember, this sheet has a different labeling scheme, so diff column for stock filtering
    CumSP_filtered <- twoa_sheets$aeq |>
      dplyr::filter(.data$sheet == "2A_CU&M_N") |>
      dplyr::filter(.data$stock == stock_data$tamm_2ac_label_2[i]) |>
      dplyr::select("fishery_joint_label", "value")

    if (stock %in% c("Elwha", "Hoko")) {
      ## Remember, this sheet has a different labeling scheme, so using $jdf_label for stock filter
      Filtered <- JDF_new$aeq |>
        dplyr::filter(.data$stock == stock_data$jdf_label[i]) |>
        dplyr::select("fishery_joint_label", "value")
      colnames(Filtered) <- c("Fishery", "Stock")
    }

    if (stock == "Dung"){
      Filtered <- JDF_unmarked$aeq |>
        dplyr::filter(.data$stock == stock_data$jdf_label[i]) |>
        dplyr::select("fishery_joint_label", "value")
      colnames(Filtered) <- c("Fishery", "Stock")
    }

    if (stock %in% c("Nisq")) {
      Filtered <- CumSP_filtered
      colnames(Filtered) <- c("Fishery", "Stock")
    }

    if (stock %in% c("Nooksack")) {
      Filtered <- cbind(Mrkd_filtered, Cumu_natural_filtered$value, Cumu_natural_filtered$value)
      colnames(Filtered) <- c("Fishery", "Marked Impacts", "Unmarked Impacts", "Cumulative Impacts")
    }


    if (!stock %in% c("Dung", "Nisq", "Elwha", "Hoko", "Nooksack")) {
      Filtered <- cbind(Mrkd_filtered, Unmrkd_filtered$value, Cumu_filtered$value)
      colnames(Filtered) <- c("Fishery", "Marked Impacts", "Unmarked Impacts", "Cumulative Impacts")
    }


    Filtered <- Filtered |> # providing categorical classes for subsetting
      dplyr::mutate(Class = dplyr::case_when(
        grepl("^Alaska", .data$Fishery) ~ "Alaska",
        grepl("^Canada", .data$Fishery) ~ "Canada",
        grepl("S. Of", .data$Fishery) ~ "S. of Falcon",
        grepl("Ocean Troll: Trty", .data$Fishery) ~ "Tribal",
        grepl("Net: Trty|Snd Trty Troll|Out-of-Region net: Trty|Pgt Snd Troll Trty|Strait Net Trty|San Juans Net Trty", .data$Fishery) ~ "Tribal",
        grepl("Sport", .data$Fishery) ~ "WDFW",
        grepl(" NTrty|Ntrty", .data$Fishery) ~ "WDFW",
        grepl("Test", .data$Fishery) ~ "Test",
        TRUE ~ NA
      ))


    if (stock %in% c("Dung", "Nisq", "Elwha", "Hoko")) {
      Filtered <- Filtered |>
        dplyr::mutate(Stock = convert_cols_to_num(.data$Stock))

      df_summary <- Filtered |>
        dplyr::filter(!is.na(.data$Class)) |> # Exclude NA values in Class
        ## for some reason we are combining S of F with WDFW
        ## easiest way: relabel S. of Falcon as WDFW for this summarize
        ## UPDATE: this is because we're going to just be working with "Treaty" and "Nontreaty" in the SUS
        dplyr::mutate(Class = dplyr::if_else(.data$Class == "S. of Falcon",
                                             "WDFW",
                                             .data$Class
        )) |>
        dplyr::group_by(.data$Class) |>
        dplyr::summarise(Stock = sum(.data$Stock, na.rm = TRUE))
    } else {
      Filtered <- Filtered |>
        dplyr::mutate(dplyr::across(
          c("Marked Impacts", "Unmarked Impacts", "Cumulative Impacts"),
          convert_cols_to_num
        ))

      if (stock %in% c("Lake", "Green", "Puyallup")) { # removing terminal catch from WDFW and
        Filtered <- Filtered |>
          dplyr::filter(!grepl("Freshwater|Terminal", .data$Fishery))
      }
      df_summary <- Filtered |>
        dplyr::filter(!is.na(.data$Class)) |> # Exclude NA values in Class
        ## for some reason we are combining S of F with WDFW
        ## easiest way: relabel S. of Falcon as WDFW for this summarize
        ## UPDATE: this is because we're going to just be working with "Treaty" and "Nontreaty" in the SUS
        dplyr::mutate(Class = dplyr::if_else(.data$Class == "S. of Falcon",
                                             "WDFW",
                                             .data$Class
        )) |>
        dplyr::group_by(.data$Class) |>
        dplyr::summarise(dplyr::across(
          c("Marked Impacts", "Unmarked Impacts", "Cumulative Impacts"),
          \(x){sum(x, na.rm = TRUE)}
        )
        )
    }





    Filtered_ER <- Filtered |>
      dplyr::select(-.data$Class)


    combined_df <- as.data.frame(rbind( # grab starting information
      as.matrix(Filtered_ER |>
                  dplyr::filter(grepl("Abundance", .data$Fishery, ignore.case = TRUE))),
      as.matrix(df_summary |>
                  dplyr::filter(grepl("WDFW", .data$Class, ignore.case = TRUE))),
      as.matrix(df_summary |>
                  dplyr::filter(grepl("Tribal", .data$Class, ignore.case = TRUE)))
    )) |>
      dplyr::mutate(dplyr::across(
        -.data$Fishery,
        convert_cols_to_num
      ))

    col_count <- ncol(combined_df)

    col_idx <- get_col_index(stock)
    abund <- as.numeric(combined_df[1, col_idx])
    if(stock == "Mid"){
      thresh = NA
    } else {
      thresh = as.numeric(Stock_Thresh$er_ceiling)
    }
    ## special case for Stilly: two abundances, two thresholds. Need to make sure they align.
    ## col_idx is "Marked Impacts", "Unmarked Impacts", but if we switch the order, our abunds switch.
    ## For robustness, just going to do this the hard way: reorder based on matching the "Unmarked" vs "Marked" designation
    if (stock == "Still") {
      names(thresh) <- gsub(".* ", "", Stock_Thresh$stock_name)
      thresh <- thresh[gsub(" .*", "", col_idx)]
    }

    if (stock %in% c("Nisq", "Skagit Sp", "Skykomish", "Skok")) { # these 4 are "TOTAL ER" therefore we have to take into account northern impacts
      Northern_Impacts <- Filtered |>
        dplyr::filter(.data$Fishery %in% c("Alaska", "Canada"))
      if (stock %in% c("Skagit Sp", "Skok")){
        impact_col <- "Cumulative Impacts"
      } else {
        impact_col <- "Stock"
      }
      Northern_Impacts <- sum(Northern_Impacts[[impact_col]])
      threshold_val <- (abund * thresh - Northern_Impacts) # what is the total abundance * ER rate cap, divided by 2 for tribal and state
    } else {
      threshold_val <- (abund * thresh)
    }

    ################## changes to make people happy

    ## making an SUS total catch row
    sus_catch_row <- combined_df |>
      dplyr::filter(.data$Fishery != "TOTAL ABUNDANCE") |>
      dplyr::select(col_idx) |> ##col_idx is an environmental var, do NOT use .data$
      dplyr::summarize(dplyr::across(
        col_idx,
        \(x){sum(as.numeric(x), na.rm = T)}
      )) |>
      dplyr::mutate(Fishery = "Total SUS Catch")

    ## update labeling
    combined_df <- combined_df |>
      dplyr::mutate(Fishery = dplyr::case_match(.data$Fishery,
                                                "TOTAL ABUNDANCE" ~ "Abundance",
                                                "WDFW" ~ "NTRTY SUS CATCH",
                                                "Tribal" ~ "TRTY SUS CATCH",
                                                .default = .data$Fishery
      ))

    if(stock != "Mid"){

      ## making row for catch threshold
      threshold_row <- as.data.frame(t(threshold_val))
      names(threshold_row) <- col_idx
      threshold_row$Fishery <- "SUS Catch Threshold"

      sus_vals_only <- sus_catch_row |>
        dplyr::select(-.data$Fishery)
      thresh_vals_only <- threshold_row |>
        dplyr::select(-.data$Fishery)
      ## make sure alignment is right
      thresh_vals_only <- thresh_vals_only[, names(sus_vals_only)]
      sus_diff_row <- thresh_vals_only - sus_vals_only
      sus_diff_row <- sus_diff_row |>
        dplyr::mutate(Fishery = "Difference")

      thresh_er_row <- as.data.frame(t(thresh))
      names(thresh_er_row) <- col_idx
      thresh_er_row <- thresh_er_row |>
        dplyr::mutate(Fishery = "ER Ceiling")

      combined_df <- dplyr::bind_rows(
        combined_df,
        sus_catch_row,
        threshold_row,
        sus_diff_row,
        thresh_er_row
      )
      if (Stock_Thresh$er_type[1] == "Total") {
        new_row <- Filtered_ER |>
          dplyr::filter(.data$Fishery == "Total Exploitation")
      } else if (Stock_Thresh$er_type[1] == "PT-SUS") {
        new_row <- Thresholdsnew |>
          dplyr::select(stock_name = "stock", "pt_sus_er") |>
          dplyr::filter(grepl(.env$stock, .data$stock_name)) |>
          dplyr::mutate(Fishery = "PT-SUS") |>
          dplyr::select(-.data$stock_name)
        names(new_row)[which(names(new_row) == "pt_sus_er")] <- col_idx
      } else {
        new_row <- Filtered_ER |>
          dplyr::filter(.data$Fishery == "Exploitation in Southern U.S. Fisheries")
      }
      # add new row of ER based on calculation or for PT-SUS the TAMM ER_ESC tab
      # new_row <- as.matrix(Filtered_ER |> dplyr::filter(grepl(er_filter, Fishery, ignore.case = TRUE)))

      ## If stock type is PT-SUS,
      ertype <- Stock_Thresh$er_type[1]
      ## relabel "UM SUS" to "SUS"
      ertype <- ifelse(ertype == "UM SUS",
                       "SUS",
                       ertype
      )
      new_row$Fishery <- paste0("ER", ertype)
      combined_df <- dplyr::bind_rows(combined_df, new_row)
      # Rename final labels

      combined_df$Fishery <- c(
        "Abundance", "NTRTY SUS Catch", "TRTY SUS Catch",
        "Total SUS Catch",
        "SUS Catch Threshold", "Difference", "ER Ceiling", paste("ER ", ertype, sep = "")
      )[1:nrow(combined_df)]
    }
    if (stock %in% c("Skykomish")) { # skykomish requires two ER values and allocations being calculated for total and sus
      Northern_Impacts <- Filtered |>
        dplyr::filter(.data$Fishery %in% c("Alaska", "Canada"))
      Northern_Impacts <- sum(Northern_Impacts$`Cumulative Impacts`)
      threshold_val <- (abund * .2 - Northern_Impacts)
      new_threshold <- data.frame(t(threshold_val))
      names(new_threshold) <- col_idx
      new_threshold$Fishery <- "Catch Threshold"

      ## SUS difference row
      thresh_vals_only <- new_threshold |>
        dplyr::select(-.data$Fishery)
      sus_vals_only <- combined_df |>
        dplyr::filter(.data$Fishery == "Total SUS Catch") |>
        dplyr::select(names(thresh_vals_only))
      # make sure alignment is right
      thresh_vals_only <- thresh_vals_only[, names(sus_vals_only)]
      sus_diff_row <- thresh_vals_only - sus_vals_only
      sus_diff_row <- sus_diff_row |>
        dplyr::mutate(Fishery = "SUS Difference")

      ## total er ceiling row
      er_ceiling_new <- dplyr::tibble(Fishery = "Total ER Ceiling", `Cumulative Impacts` = 0.2)

      ## Total ER row
      er_total_new <- Filtered_ER |>
        dplyr::filter(.data$Fishery == "Total Exploitation") |>
        dplyr::mutate(Fishery = "Total ER")

      combined_df <- dplyr::bind_rows(
        combined_df,
        new_threshold,
        sus_diff_row,
        er_ceiling_new,
        er_total_new
      )
    }


    Impacts_Calc <- Impacts_Print |> # impact calc is just a way to calculate and bin fisheries impacts based on category.Turning limiting stock tabb into percents
      dplyr::filter(!is.na(.data$Class)) |> # Exclude NA values in Class
      dplyr::group_by(.data$Class) |>
      dplyr::summarise(dplyr::across(
        c("TotalMarked", "TotalUnmarked"),
        \(x){sum(x, na.rm = TRUE)}
      )) |>
      dplyr::arrange(.data$Class) |>
      dplyr::mutate(
        `TotalMarked%` = round(.data$TotalMarked / sum(.data$TotalMarked), 3),
        `TotalUnmarked%` = round(.data$TotalUnmarked / sum(.data$TotalUnmarked), 3),
        `Total%` = round((.data$TotalUnmarked + .data$TotalMarked) /
                           (sum(.data$TotalUnmarked) + sum(.data$TotalMarked)), 3)
      )

    # tier 1 fishery is a broad overview of what entity gets what share
    Class <- c("SEAK", "Canada", "SOF Ocean", "Tribal Ocean", "Tribal P.Sound", "Tribal River", "WDFW Ocean", "WDFW P.Sound", "WDFW River", "OTHER")
    color_scale <- grDevices::colorRampPalette(c("red", "black"))(length(Class))
    tier1color <- data.frame(
      Class = Class,
      colors = color_scale,
      stringsAsFactors = FALSE
    )
    tier1fisheries <- dplyr::left_join(tier1color, Impacts_Calc, by = "Class") # added by Yi
    tier1fisheries <- tier1fisheries |>
      dplyr::filter(!is.na("Total%"))

    ## DistributionPlot now creates the file as well, returns the file name of the temporary file
    plot_path <- distribution_plot(
      tier1fisheries = tier1fisheries,
      long_df = long_df,
      sheet_name = sheet_name,
      width = plot_width,
      height = plot_height,
      units = plot_units
    )
    ## save teh temporary file name for later deletion
    temp_plot_paths <- c(temp_plot_paths, plot_path)

    # Determine number of columns in combined_df to find where to place the image
    col_offset <- ncol(Impacts_Print) + 2 # +2 gives one column of space
    openxlsx::insertImage(wb,
                          sheet = sheet_name, file = plot_path,
                          startCol = col_offset, startRow = 1, width = plot_width, height = plot_height, units = plot_units
    )

    row_offset <- nrow(Impacts_Print) + 3 # +2 gives one column of space

    openxlsx::writeData(wb, sheet = sheet_name, x = Impacts_Calc,
                        startRow = row_offset,
                        startCol = 3, borders = "rows",
                        headerStyle = openxlsx::createStyle(textDecoration = "bold"))

    ## add percent formatting for the TotalMarked% through Total%
    ## Annoyingly, need two separate styles, one for the "main" columns and one for the "right" column
    ##   in order to keep the borders as we want them
    pct <- openxlsx::createStyle(
      numFmt = "0%",
      border = c("top", "bottom")
    )
    pct_right <- openxlsx::createStyle(
      numFmt = "0%",
      border = c("top", "bottom", "right")
    )
    openxlsx::addStyle(wb,
                       sheet = sheet_name,
                       style = pct,
                       cols = 6:7,
                       rows = (row_offset) + (1:nrow(Impacts_Calc)),
                       gridExpand = TRUE
    )
    openxlsx::addStyle(wb,
                       sheet = sheet_name,
                       style = pct_right,
                       cols = 8,
                       rows = (row_offset) + (1:nrow(Impacts_Calc)),
                       gridExpand = TRUE
    )

    openxlsx::saveWorkbook(wb, output_path, overwrite = TRUE)

    ## Storing the components lists for later adding. No need to make hidden sheets
    component_list[[stock]] <- list(combined_df = combined_df, Impacts_Calc = Impacts_Calc)
  } # END Primary LOOP


  openxlsx::addWorksheet(wb, "Crayon Box", tabColour = "tan3")

  current_row <- 1 # Keep track of where to paste next

  for (i in seq_along(component_list)) {
    stock_name <- stock_data$label[i]
    openxlsx::writeData(wb,
                        sheet = "Crayon Box", x = data.frame(stock_name),
                        startRow = current_row, startCol = 1, colNames = FALSE
    )
    current_row <- current_row + 1
    n_rows <- max(c(
      nrow(component_list[[1]]$combined_df),
      nrow(component_list[[1]]$Impacts_Calc)
    ))
    openxlsx::writeData(wb,
                        sheet = "Crayon Box", x = component_list[[i]]$combined_df,
                        startRow = current_row, startCol = 1, colNames = TRUE
    )
    openxlsx::writeData(wb,
                        sheet = "Crayon Box", x = component_list[[i]]$Impacts_Calc,
                        startRow = current_row, startCol = 6, colNames = TRUE
    )
    current_row <- current_row + n_rows + 4 # add gap after each block. Remember two of these rows are for the headers
  }

  ## Final bit ------------------------------------------------------------------

  # Load the combined workbook
  sheet_name <- "Crayon Box"
  #
  # Read the data in the overview sheet
  overview_df <- openxlsx::readWorkbook(wb,
                                        sheet = sheet_name, colNames = FALSE,
                                        skipEmptyCols = FALSE,
                                        skipEmptyRows = FALSE
  )
  ## Updated to not turn first row to headers. Makes row numbers align expectation.
  fishery_col <- 1

  block_header_rows <- which(overview_df[, 1] == "Fishery") ## rows for the headers
  cols_for_headers <- which(!is.na(overview_df[2, ])) ## columns that have headers in them.
  class_col <- grep("Class", overview_df[2, ]) ## col with "class" etc.




  ## More formatting ======================
  # REMINDER!! MUST PROVIDE ROWS/COLS AS NUMBERS, NOT LOGICAL VEC! OTHERWISE, CRASHES EXCEL
  ## add comma handling for non-percent numbers

  block_rows = which(!is.na(overview_df[ , class_col]) & overview_df[ , class_col] != "Class")
  # class blocks
  numeric_cols <- which(overview_df[2, ] %in% c("TotalMarked", "TotalUnmarked"))
  openxlsx::addStyle(wb,
                     sheet = sheet_name,
                     style = openxlsx::createStyle(numFmt = "#,##0.00"),
                     cols = numeric_cols,
                     rows = block_rows,
                     gridExpand = TRUE
  )

  ## abundance blocks
  er_rows <- c(grep("^ER", overview_df[, 1]), grep("^Total ER", overview_df[, 1]))
  ## want all rows that are not ER, not the header blocks, and not the fishery name (which is one above header block)
  unwanted_rows <- c(er_rows, block_header_rows, (block_header_rows - 1))
  numeric_rows <- setdiff(1:nrow(overview_df), unwanted_rows)
  numeric_cols <- grep(" Impacts$", overview_df[2, ])

  openxlsx::addStyle(wb,
                     sheet = sheet_name,
                     style = openxlsx::createStyle(numFmt = "#,##0.00"),
                     cols = 2:4,
                     rows = numeric_rows,
                     gridExpand = TRUE
  )



  # ## add percentage handling for ER
  er_cols <- grep(" Impacts$", overview_df[2, ]) # Marked Impacts through Cumulative Impacts

  openxlsx::addStyle(wb,
                     sheet = sheet_name,
                     style = openxlsx::createStyle(numFmt = "0.00%"),
                     cols = er_cols,
                     rows = er_rows,
                     gridExpand = TRUE
  )



  ## Add percentage handling for class sections
  percent_cols <- grep("%$", overview_df[2, ])
  block_rows <- which(!is.na(overview_df[, class_col]) & overview_df[, class_col] != "Class")

  openxlsx::addStyle(wb,
                     sheet = sheet_name,
                     style = openxlsx::createStyle(numFmt = "0.00%"),
                     cols = percent_cols,
                     rows = block_rows,
                     gridExpand = TRUE
  )


  ## add header formatting

  openxlsx::addStyle(wb,
                     sheet = sheet_name,
                     style = openxlsx::createStyle(textDecoration = "bold",
                                                   fontSize = 12,
                                                   halign = "center"),
                     cols = cols_for_headers,
                     rows = block_header_rows,
                     gridExpand = TRUE
  )

  ## add fishery label handling
  openxlsx::addStyle(wb,
                     sheet = sheet_name,
                     style = openxlsx::createStyle(
                       textDecoration = "bold",
                       fontSize = 14,
                       halign = "center",
                       fgFill = "#90D5FF"
                     ),
                     cols = 1:max(cols_for_headers),
                     rows = (block_header_rows - 1),
                     gridExpand = TRUE
  )

  for (fishery_row_cur in (block_header_rows - 1)) {
    openxlsx::mergeCells(wb, sheet_name,
                         cols = 1:max(cols_for_headers),
                         rows = fishery_row_cur
    )
  }

  #
  # ## adding colors ========================
  # Get actual row numbers in Excel
  diff_row <- which(overview_df[[fishery_col]] %in% c("Difference", "SUS Difference"))

  er_row <- which(overview_df[[fishery_col]] %in% c("ER", "ER Total", "ER SUS", "ER PT-SUS", "Total ER"))

  # Create styles
  red_style <- openxlsx::createStyle(bgFill = "tomato")
  green_style <- openxlsx::createStyle(bgFill = "palegreen1")



  # Apply to each group
  apply_diff_formatting(diff_row, wb = wb, sheet_name = sheet_name,
                        red_style = red_style, green_style = green_style)

  for (r in er_row) {
    for (col in 2:4) {
      cell <- paste0(LETTERS[col], r)
      cell_above <- paste0(LETTERS[col], r - 1)
      openxlsx::conditionalFormatting(wb,
                                      sheet = sheet_name, cols = col, rows = r,
                                      # rule = paste0("AND(NOT(ISBLANK(", cell_above, ")), ROUND(", cell, ", 3) > ROUND(", cell_above, ", 3))"),
                                      rule = glue::glue("AND(NOT(ISBLANK({cell_above})), ROUND({cell}, 3) > ROUND({cell_above}, 3))"),
                                      style = red_style
      )
      openxlsx::conditionalFormatting(wb,
                                      sheet = sheet_name, cols = col, rows = r,
                                      # rule = paste0("AND(NOT(ISBLANK(", cell_above, ")), ROUND(", cell, ", 3) <= ROUND(", cell_above, ", 3))"),
                                      rule = glue::glue("AND(NOT(ISBLANK({cell_above})), ROUND({cell}, 3) <= ROUND({cell_above}, 3))"),
                                      style = green_style
      )
    }
  }

  ## make sure the columns are wide enough to read the headers
  openxlsx::setColWidths(wb,
                         sheet = sheet_name,
                         cols = cols_for_headers, widths = 18
  )
  openxlsx::setColWidths(wb,
                         sheet = sheet_name,
                         cols = c(7, 8, 9, 11),
                         widths = 14.5
  )

  sheet_names <- names(wb)
  # Desired order: move "Crayon Box" to the first position
  new_order <- c(
    which(sheet_names == "Crayon Box"),
    which(sheet_names != "Crayon Box")
  )
  # Apply the new order using numeric indices
  openxlsx::worksheetOrder(wb) <- new_order
  # print(output_path)
  openxlsx::saveWorkbook(wb, output_path, overwrite = TRUE)
  if(verbose){cli::cli_alert_success("Summary saved as {output_path}")}
  return(invisible(output_path))
}
