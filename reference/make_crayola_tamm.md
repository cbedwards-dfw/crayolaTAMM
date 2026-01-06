# Make workbook summary of Chinook TAMM

Make workbook summary of Chinook TAMM

## Usage

``` r
make_crayola_tamm(
  filepath,
  output_path = NULL,
  plot_width = 10,
  plot_height = 8,
  plot_units = "in",
  verbose = TRUE
)
```

## Arguments

- filepath:

  Filepath of TAMM, including `.xlsx` suffix

- output_path:

  Filepath for resulting file. Optional; if not provided, will save in
  same folder as `filepath` with name of
  "{filepath}\_Impacts_crayola.xlsx"

- plot_width:

  Controls width of barplots in stock sheets. Numeric, defaults to 10

- plot_height:

  Controls height of barplots in stock sheets. Numeric, defaults to 8

- plot_units:

  Character string defining units of `plot_width` and `plot_height`.
  Defaults to "in", can also be "cm" or "px".

- verbose:

  Should progress messages and saved filepath be printed to screen?

## Value

Invisibly returns filepath of the summary file.
