
{
  rm(list = ls(all = TRUE))
  graphics.off()
  dir <- getwd()
  setwd(dir)

  source("function.r")
  library(writexl)
  library(ggplot2)
  library(readxl)
  library(dplyr)
  library(tsapp)
  library(lmtest)
  library(forecast)
  library(tseries)
  library(MASS)

  df <- read_excel("Referat II __ Student II.xls", skip = 2)
  df_liste <- split(df, df$"Artikel Nr")
  df_results <- data.frame(df_nr = numeric(),
                           Best_Modell_Öko = character(),
                           Öko_Wert = numeric(),
                           untereski = numeric(),
                           oberes_ki = numeric(),
                           Best_Modell_Stat = character(),
                           stat_Wert = numeric())
  df_sort <- lapply(df_liste, clear_data)
  #########Automatisierung für gesamte Excel_Datei##################
  x <- length(df_liste)
  for (i in 1:x){
  #for (i in 1:1){
    #Ich möchte die for-schleife ersetzen
    df_sorted <- df_sort[[i]]
    #plot_list(df_sorted)
    modell_results <- data.frame(pred_modell = character(),
                                 MAE = numeric(),
                                 RMSE = numeric(),
                                 AIC = numeric(),
                                 BIC = numeric(),
                                 HQ = numeric(),
                                 #Varianz_Abverkauf = numeric(),
                                 #Varianz_Vorhersage = numeric(),
                                 Gesamt_Kosten = numeric(),
                                 stringsAsFactors = FALSE)
    n <- 20
    kosten <- NULL
    while (n > 0) {
      df_trimmed <- head(df_sorted, n = nrow(df_sorted) - n)
    # plot_list(df_trimmed)
      ts_df_trimmed <- ts(df_trimmed$Abverkauf, frequency = 1)
      ts_dfv_trimmed <- as.vector(ts_df_trimmed)
      unteres_ki <- 0
      oberes_ki <- 0
      # Auswahl des Modells und Vorhersage
      mean_sales <- round(mean(ts_dfv_trimmed))
      #cat("Mittelwert: ", mean_sales, "\n")
      if (mean_sales > 10) {
        modell <- arima_test(ts_dfv_trimmed)
        pred_modell_string  <- c("ARIMA ", modell$ar, ".", modell$i, ".", modell$ma)
        pred_modell <- paste0(pred_modell_string, collapse = "")
        prediction <- predict(modell$modell, n.ahead = 4)
        
        prediction_values <- as.vector(round(prediction$pred))
        prediction_se <- prediction$se
        
        # Berechnen der potenziellen Über- und Unterdeckung
        oberes_ki <- abs(round(prediction_values + 1.645 * prediction_se))
        unteres_ki <- abs(round(prediction_values - 1.645 * prediction_se))
        info_calc <- info_krit(modell$modell)
        aic <- info_calc$aic
        bic <- info_calc$bic
        hq <- info_calc$hq
      } else {
              modell <- poisson_test(mean_sales)
              pred_modell <- "Poisson"

              unteres_ki <- modell$unteres_ki
              oberes_ki <- modell$oberes_ki
              oberes_ki <- c(oberes_ki, oberes_ki, oberes_ki, oberes_ki)
              unteres_ki <- c(unteres_ki, unteres_ki, unteres_ki, unteres_ki)
              prediction_values <- as.vector(round(c(mean_sales, mean_sales, mean_sales, mean_sales)))}
      
      oberes_ki <- as.vector(c(oberes_ki))
      unteres_ki <- as.vector(c(unteres_ki))

      # Berechnung von RMSE und Varianz der Vorhersagen
      actual_values <- as.vector(ts_dfv_trimmed[(length(ts_dfv_trimmed) - 3):length(ts_dfv_trimmed)])
      mae <- mean(abs(prediction_values - actual_values))
      rmse <- sqrt(mean((prediction_values - actual_values)^2))
      #varianz_vorhersagen <- var(prediction_values)
      
      ges_kosten <- kostenberechnung(prediction_values,actual_values,unteres_ki,oberes_ki)
      
      # Speichern der Ergebnisse
      modell_results <- rbind(modell_results, data.frame(Pred_modell = pred_modell,
                                                         MAE = mae,
                                                         RMSE = rmse,
                                                         AIC = ifelse(pred_modell == "Poisson", NA, round(aic,2)),
                                                         BIC = ifelse(pred_modell == "Poisson", NA, round(bic,2)),
                                                         HQ = ifelse(pred_modell == "Poisson", NA, round(hq,2)),
                                                         #Varianz_Abverkauf = varianz_abverkauf,
                                                         #Varianz_Vorhersage = varianz_vorhersagen,
                                                         Gesamtkosten = ges_kosten))
      n <- n - 1
    }
    #print(i)
    # Sortieren der Ergebnisse
    modell_results_aic <- modell_results[order(modell_results$AIC), ]
    modell_results_bic <- modell_results[order(modell_results$BIC), ]
    modell_results_hq <- modell_results[order(modell_results$HQ), ]
    
    modell_results_rmse <- modell_results[order(modell_results$RMSE), ]
    modell_results_mae <- modell_results[order(modell_results$MAE), ]
    
    modell_results_gkosten <- modell_results[order(modell_results$Gesamtkosten), ]
    
    best_modell_mae <- modell_results_mae$Pred_modell[1]
    best_modell_bic <- modell_results_bic$Pred_modell[1]

    mae <- modell_results_mae$MAE[1]
    bic <- modell_results_bic$BIC[1]
    
    df_results <- rbind(df_results, data.frame(df_nr = nrow(df_results) + 1,
                                               Artikel_Nr = df_sorted$Artikel_Nr[1],
                                               Öko_Wert = modell_results_gkosten$Gesamtkosten[1],
                                               Best_Modell_Öko = modell_results_gkosten$Pred_modell[1],
                                               MAE = mae,
                                               Best_Modell_MAE = best_modell_mae,
                                               BIC = bic,
                                               Best_Modell_BIC = best_modell_bic,
                                               unteres_ki = unteres_ki[1],
                                               Mittelwert = mean_sales,
                                               oberes_ki = oberes_ki[1]))
  }
  #write_xlsx(df_results, "df_results.xlsx")  
  View(df_results)
}