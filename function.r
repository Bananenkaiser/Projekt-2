dir <- getwd()
setwd(dir)
library(ggplot2)
library(readxl)
library(dplyr)
library(tsapp)
library(lmtest)
library(forecast)
library(tseries)
library(MASS)

plot_list <- function(df_list) {
  df <- ts(df_list$Abverkauf, frequency = 1)
  dfv <- as.vector(df)
  dfv <- as.numeric(dfv)
  plot(dfv, type = "l", col = "blue", xlab = "Zeit", ylab = "Abverkauf")
  title(main = df_list$Artikel_Nr[1], col.main = "black", font.main = 4, cex.main = 1.2)
}

arima_test <- function(zeitreihe) {
  
  # Check stationarity and get differencing order
  stationary <- calc_stationarity(zeitreihe)
  d_order <- stationary$ts_diff
  
  # Find significant lags
  acf_result <- acf(zeitreihe, plot = FALSE)
  pacf_result <- pacf(zeitreihe, plot = FALSE)
  significant_acf_lags <- which(abs(acf_result$acf) > 2 / sqrt(length(zeitreihe)))
  significant_pacf_lags <- which(abs(pacf_result$acf) > 2 / sqrt(length(zeitreihe)))
  # Define ARIMA order based on significant lags
  p_order <- length(significant_pacf_lags)
  q_order <- length(significant_acf_lags)
  # Adjust significant lags to be within the order limits
  adjusted_significant_pacf_lags <- significant_pacf_lags[significant_pacf_lags <= p_order]
  adjusted_significant_acf_lags <- significant_acf_lags[significant_acf_lags <= q_order]
  # Create fixed values
  fixed_values <- create_fixed_values(adjusted_significant_pacf_lags, p_order, adjusted_significant_acf_lags, q_order)

  if (length(fixed_values) != p_order + q_order) {
    fixed_values <- c(fixed_values, NA)
  }
  if (d_order == 0) {
    fixed_values <- c(fixed_values, NA)
  }
  # Fitting the model
  if (p_order == 0 || q_order == 0 || p_order > 20 || q_order > 20) {
    # No significant lags found, consider a simple modell
    order <- c(2, 0, 0)
    modell <- arima(zeitreihe, order = order)
  } else {
    # Fit ARIMA model with identified orders
    order <- c(p_order, d_order, q_order)
    modell <- try(arima(zeitreihe, order = order, fixed = fixed_values), silent = TRUE)
    if (inherits(modell, "try-error")) {
      order <- c(2, 0, 0)
      modell <- arima(zeitreihe, order = order)
    }
  }
  return(list(modell = modell, size = length(modell$coef), ar = p_order, i = d_order, ma = q_order, fixed = fixed_values))
}

calc_stationarity <- function(zeitreihe) {
  if (length(zeitreihe) <= 1) {
    stop("Time series is too short for stationarity testing.")
  }
  ts_diff <- 0
  original_zeitreihe <- zeitreihe
  for (ts_diff in 0:4) {
    if (ts_diff > 0) {
      zeitreihe <- diff(zeitreihe)
    }
    adf_result <- adf.test(zeitreihe, alternative = "stationary")
    if (adf_result$p.value < 0.05) {
      break
    }
  }
  if (ts_diff == 5) {
    warning("Maximum differencing reached. The time series might still be non-stationary.")
  }
  differenzierte_zeitreihe <- if (ts_diff > 0) diff(original_zeitreihe, differences = ts_diff) else original_zeitreihe
  return(list(ts_diff = ts_diff, differenzierte_zeitreihe = differenzierte_zeitreihe))
}

create_fixed_values <- function(significant_ar, p_order, significant_ma, q_order){
  # Initialisiere fixed_values für AR und MA Komponenten
  fixed_values_ar <- rep(0, p_order)
  fixed_values_ma <- rep(0, q_order)

  # Setze NA für signifikante Lags
  if (!is.null(significant_ar)) {
    fixed_values_ar[significant_ar] <- NA
  }
  if (!is.null(significant_ma)) {
    fixed_values_ma[significant_ma] <- NA
  }
  # Kombiniere AR und MA fixed_values
  return(c(fixed_values_ar, fixed_values_ma))
}

clear_data <- function(df_sorted) {
  colnames(df_sorted) <- gsub(" ", "_", colnames(df_sorted))
  df_sorted[order(df_sorted$Jahr, df_sorted$KW), ]

  # Überprüfe auf doppelte Werte in der Kombination Jahr und KW
  duplicates <- df_sorted[duplicated(df_sorted[c("Jahr", "KW")]) | duplicated(df_sorted[c("Jahr", "KW")], fromLast = TRUE), ]
  # Überprüfe, ob bei einem doppelten Paar der Wert bei Abverkauf gleich null ist
  duplicates_with_zero_abverkauf <- duplicates[duplicates$Abverkauf == 0, ]
  if (nrow(duplicates_with_zero_abverkauf) > 0) {
    # Lösche die Zeilen mit doppelten Werten und Abverkauf gleich null
    df_sorted <- df_sorted[!(duplicated(df_sorted[c("Jahr", "KW")]) & df_sorted$Abverkauf == 0), ]
  }

  artikel_nr <- df_sorted$Artikel_Nr[1]
  start_jahr <- 1995
  start_kw <- 44
  end_jahr <- 1997
  end_kw <- 22
  expected_kw <- start_kw
  expected_jahr <- start_jahr
  while (expected_jahr < end_jahr || (expected_jahr == end_jahr && expected_kw <= end_kw)) {
    if (!any(df_sorted$Jahr == expected_jahr & df_sorted$KW == expected_kw)) {
      # Fülle die Lücke mit 0 für "Abverkauf"
      new_row <- data.frame(
        Artikel_Nr = artikel_nr,
        Jahr = expected_jahr,
        KW = expected_kw,
        Abverkauf = 0
      )
      # Füge die neue Zeile an der richtigen Position ein
      df_sorted <- bind_rows(df_sorted, new_row)
    }
    # Aktualisiere die erwarteten Werte
    expected_kw <- (expected_kw %% 52) + 1  # Modulo 52, da es 52 Kalenderwochen im Jahr gibt
    expected_jahr <- ifelse(expected_kw == 1, expected_jahr + 1, expected_jahr)
  }
  df_sorted <- df_sorted[order(df_sorted$Jahr, df_sorted$KW), ]
  # Überprüfe auf doppelte Werte in der Kombination Jahr und KW
  duplicates <- df_sorted[duplicated(df_sorted[c("Jahr", "KW")]) | duplicated(df_sorted[c("Jahr", "KW")], fromLast = TRUE), ]
  # Überprüfe, ob bei einem doppelten Paar der Wert bei Abverkauf gleich null ist
  duplicates_with_zero_abverkauf <- duplicates[duplicates$Abverkauf == 0, ]
  if (nrow(duplicates_with_zero_abverkauf) > 0) {
    df_sorted <- df_sorted[!(duplicated(df_sorted[c("Jahr", "KW")]) & df_sorted$Abverkauf == 0), ]
  }
  return(df_sorted)
}

kostenberechnung <- function(prediction_values,actual_values,unteres_ki,oberes_ki){
  x <- length(actual_values)
  for (k in 1:x){
    kosten_pro_unterdeckung <- 2
    kosten_pro_ueberdeckung <- 1
    unterdeckung <- 0
    ueberdeckung <- 0
    if (actual_values[k] > prediction_values[k]) {
      unterdeckung <- (actual_values[k] - prediction_values[k])
    } else {
      ueberdeckung <- (prediction_values[k] - actual_values[k])
    }
    if(prediction_values[k] > unteres_ki[k]) {
      unterdeckung <- unterdeckung + (prediction_values[k] - unteres_ki[k])
    }
    if (actual_values[k] < oberes_ki[k]) {
      ueberdeckung <- ueberdeckung + (oberes_ki[k] - actual_values[k])
    }
    kosten[k] <- (unterdeckung * kosten_pro_unterdeckung) + (ueberdeckung * kosten_pro_ueberdeckung)
  }
  ges_kosten <- sum(kosten)
  return(ges_kosten)
}

aic_calc <- function(modell) {
  npar <- length(modell$coef) + 1
  aic <- -2 * modell$loglik + 2 * npar
  return(aic)
}

bic_calc <- function(modell) {
  npar <- length(modell$coef) + 1
  nobs <- length(modell$residuals)
  bic <- -2 * modell$loglik + npar * log(nobs)
  return(bic)
}

hq_calc <- function(modell) {
  npar <- length(modell$coef)
  nobs <- length(modell$residuals)
  hq <- -2 * modell$loglik + 2 * npar * (log(nobs) / log(2))
  return(hq)
}

info_krit <- function(modell) {
  aic <- aic_calc(modell)
  bic <- bic_calc(modell)
  hq <- hq_calc(modell)
  return(list(aic = aic, bic = bic, hq = hq))
}

poisson_test <- function(mittelwert) {
  # Initialisiere ki-Werte
  unteres_ki <- NA
  oberes_ki <- NA
  # Finde das untere ki
  for (i in 0:(2 * mittelwert)) {
    if (ppois(i, lambda = mittelwert) >= 0.05) {
      unteres_ki <- i
      break
    }
  }
  # Finde das obere ki
  for (l in unteres_ki:(4 * mittelwert)) {
    if (ppois(l, lambda = mittelwert) >= 0.95) {
      oberes_ki <- l
      break
    }
  }
  return(list(unteres_ki = unteres_ki, oberes_ki = oberes_ki))
}