distribution_plot <- function(tier1fisheries, long_df, sheet_name, width = 12, height = 8, units = "in", res = 300) { #create 3 % bar graphs for summary, wdfw and tribal that shows the distribution of impacts for that categorry
  plot_path <- tempfile(fileext = ".png")
  grDevices::png(plot_path, width = width, height = height, units = units, res = res)

  graphics::par(mfrow = c(1,3),
                 mar = c(5, 8, 2, 1))# Increase left margin for names
  graphics::barplot(
    height = as.numeric(tier1fisheries$`Total%`),
    names.arg = tier1fisheries$Class,
    horiz = TRUE,
    col = tier1fisheries$colors,
    border = NA,
    las = 1,  # Horizontal y-axis labels
    xlab = "Total %",  cex.names = .8, main = "Summary"
  )

  long_df <- long_df |>
    ## update to give Tr 3:4 Trl T4 the right classification
    dplyr::mutate(Class = dplyr::if_else(.data$TimeStepLabel == "Tr 3:4 Trl T4",
                           "Tribal P.Sound",
                           .data$Class)) |>
    dplyr::mutate(fishery_name_class = glue::glue("{TimeStepLabel} ({.data$Class})")) |>
    dplyr::mutate(fishery_name_class = gsub("WDFW ", "", .data$fishery_name_class)) |>
    dplyr::mutate(fishery_name_class = gsub("Tribal ", "", .data$fishery_name_class))

  tier2fisheries <- long_df |>
    dplyr::filter(stringr::str_detect(.data$Class, "WDFW")) |>
    dplyr::mutate(total_percent = .data$Total / sum(.data$Total)) |>
    dplyr::arrange(.data$FisheryID)

  tier2fisheries$colors <- grDevices::colorRampPalette(c("forestgreen", "goldenrod1"))(nrow(tier2fisheries))
  tier2fisheries <- tier2fisheries |>
    dplyr::filter(.data$Total != 0)
  graphics::barplot(
    height = as.numeric(tier2fisheries$total_percent),
    names.arg = tier2fisheries$fishery_name_class,
    horiz = TRUE,
    col = tier2fisheries$colors,
    border = NA, cex.axis = 1, cex.names = .8,
    las = 1,  # Horizontal y-axis labels
    xlab = "% of WDFW Total",  main = "WDFW"
  )

  tier3fisheries <- long_df |>
    dplyr::filter(stringr::str_detect(.data$Class, "Tribal")) |>
    dplyr::mutate(total_percent = .data$Total / sum(.data$Total)) |>
    dplyr::arrange(.data$FisheryID)
  tier3fisheries$colors <- grDevices::colorRampPalette(c("darkred", "darkturquoise"))(nrow(tier3fisheries))
  tier3fisheries<- tier3fisheries|>
    dplyr::filter(.data$Total != 0)

  graphics::barplot(
    height = as.numeric(tier3fisheries$total_percent),
    names.arg = tier3fisheries$fishery_name_class,
    horiz = TRUE,
    col = tier3fisheries$colors,
    border = NA, cex.axis = 1, cex.names = .8,
    las = 1,  # Horizontal y-axis labels
    xlab = "% of Tribal Total",  main = "Tribal"
  )
  grDevices::dev.off()
  return(plot_path)
}
