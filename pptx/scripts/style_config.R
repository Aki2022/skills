# Style configuration loader for R
# Reads style.yaml and provides helper functions for ggplot2 styling

library(yaml)

#' Load style configuration from style.yaml
#'
#' @param path Path to style.yaml file
#' @return List with style configuration
load_style <- function(path = NULL) {
  if (is.null(path)) {
    # Find style.yaml relative to this script
    script_dir <- dirname(sys.frame(1)$ofile)
    if (is.null(script_dir) || script_dir == "") {
      script_dir <- "."
    }
    path <- file.path(dirname(script_dir), "style.yaml")
    if (!file.exists(path)) {
      path <- "style.yaml"
    }
  }

  if (!file.exists(path)) {
    stop(paste("style.yaml not found at:", path))
  }

  yaml::read_yaml(path)
}

#' Get series colors from style configuration
#'
#' @param style Style configuration list
#' @param n Number of colors needed
#' @return Vector of hex color strings
get_series_colors <- function(style, n = 3) {
  colors <- c()
  series <- style$colors$series

  for (i in 1:min(n, length(series))) {
    s <- series[[i]]
    if (s$type == "rgb") {
      colors <- c(colors, s$value)
    } else if (s$type == "theme") {
      # Convert theme colors to approximate RGB
      # These are approximations - actual theme colors depend on PowerPoint theme
      theme_colors <- list(
        "bg1" = "#FFFFFF",
        "tx1" = "#000000",
        "accent1" = "#4472C4",
        "accent2" = "#ED7D31"
      )
      base_color <- theme_colors[[s$value]]
      if (is.null(base_color)) base_color <- "#808080"

      # Apply brightness adjustment
      brightness <- s$brightness
      if (!is.null(brightness)) {
        base_color <- adjust_brightness(base_color, brightness)
      }
      colors <- c(colors, base_color)
    }
  }

  # If we need more colors than defined, recycle
  if (length(colors) < n) {
    colors <- rep_len(colors, n)
  }

  colors
}

#' Adjust hex color brightness
#'
#' @param hex_color Hex color string (e.g., "#FFFFFF")
#' @param brightness Brightness adjustment (-1 to 1)
#' @return Adjusted hex color string
adjust_brightness <- function(hex_color, brightness) {
  # Parse hex
  r <- strtoi(substr(hex_color, 2, 3), 16)
  g <- strtoi(substr(hex_color, 4, 5), 16)
  b <- strtoi(substr(hex_color, 6, 7), 16)

  if (brightness < 0) {
    # Darken
    factor <- 1 + brightness
    r <- round(r * factor)
    g <- round(g * factor)
    b <- round(b * factor)
  } else {
    # Lighten
    r <- round(r + (255 - r) * brightness)
    g <- round(g + (255 - g) * brightness)
    b <- round(b + (255 - b) * brightness)
  }

  # Clamp values
  r <- max(0, min(255, r))
  g <- max(0, min(255, g))
  b <- max(0, min(255, b))

  sprintf("#%02X%02X%02X", r, g, b)
}

#' Get primary color
#'
#' @param style Style configuration list
#' @return Hex color string
get_primary_color <- function(style) {
  style$colors$primary
}

#' Create ggplot2 theme based on style.yaml
#'
#' @param style Style configuration list
#' @return ggplot2 theme object
theme_style <- function(style) {
  # Get font sizes
  cat_font_size <- style$category_axis$font$size_pt
  legend_font_size <- style$legend$font$size_pt

  # Get legend position
  legend_pos <- style$legend$position

  # Get axis visibility
  value_axis_visible <- style$value_axis$visible

  theme_minimal() +
    theme(
      # Legend
      legend.position = legend_pos,
      legend.text = element_text(size = legend_font_size),

      # Axis text
      axis.text = element_text(size = cat_font_size),
      axis.text.x = element_text(color = "#595959"),  # tx1 with brightness 0.35
      axis.text.y = if (value_axis_visible) element_text(color = "#595959") else element_blank(),

      # Axis lines
      axis.line.x = element_line(color = "#D9D9D9", linewidth = 0.5),  # tx1 with brightness 0.85

      # Grid
      panel.grid.major = element_blank(),
      panel.grid.minor = element_blank(),

      # Value axis ticks
      axis.ticks.y = if (value_axis_visible) element_line() else element_blank()
    )
}

#' Create scale_fill_manual with style colors
#'
#' @param style Style configuration list
#' @param ... Additional arguments passed to scale_fill_manual
#' @return ggplot2 scale object
scale_fill_style <- function(style, ...) {
  colors <- get_series_colors(style, 10)
  scale_fill_manual(values = colors, ...)
}

#' Create scale_color_manual with style colors
#'
#' @param style Style configuration list
#' @param ... Additional arguments passed to scale_color_manual
#' @return ggplot2 scale object
scale_color_style <- function(style, ...) {
  colors <- get_series_colors(style, 10)
  scale_color_manual(values = colors, ...)
}
