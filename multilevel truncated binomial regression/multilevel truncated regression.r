# Load required libraries
library(glmmTMB)
library(readxl)
library(dplyr)
library(broom.mixed)
library(performance)
library(DHARMa)
library(ggplot2)
library(officer)
library(flextable)
library(MASS)
library(car)

# Function to allow user to select input file
select_file <- function(title = "Select Excel Data File") {
  file_path <- file.choose()
  return(file_path)
}

# Allow user to select input file
cat("Please select the input Excel file...\n")
file_path <- select_file()

if (is.null(file_path) || file_path == "") {
  cat("No file selected. Exiting.\n")
  stop("No file selected")
}

# Import the data
cat(paste("Loading data from:", file_path, "\n"))
df <- read_excel(file_path, sheet = "Sheet")

# Define continuous and categorical variables
continuous_vars <- c(
  "import_from_slu_log", "age",
  "distance_miles_log", "state_percapita_income_log",
  "state_unemployment", "immigrant_density"
)

categorical_model_vars <- c(
  "sex_enc", "marital_status_enc", "employment_status_enc",
  "purpose_simple", "accomd_type_enc", "season_enc",
  "us_state_enc", "tourist_type_enc"
)

# Ensure all continuous variables are numeric
for (var in continuous_vars) {
  if (var %in% names(df)) {
    df[[var]] <- as.numeric(df[[var]])
  }
}

# Convert categorical variables to factors
for (var in categorical_model_vars) {
  if (var %in% names(df)) {
    df[[var]] <- as.factor(df[[var]])
  }
}

# Create the formula for the model
formula_vars <- c("los_capped", continuous_vars, categorical_model_vars)
existing_continuous <- continuous_vars[continuous_vars %in% names(df)]
cat_vars_in_df <- categorical_model_vars %in% names(df)
existing_categorical <- categorical_model_vars[cat_vars_in_df]
formula_parts <- c(existing_continuous, existing_categorical)
formula_string <- paste("los_capped ~", paste(formula_parts, collapse = " + "))

# Add random effect for us_state_enc
if ("us_state_enc" %in% names(df)) {
  fixed_effects <- formula_parts[formula_parts != "us_state_enc"]
  formula_string <- paste(
    "los_capped ~",
    paste(fixed_effects, collapse = " + "),
    "+ (1|us_state_enc)"
  )
} else {
  cat("Warning: us_state_enc not found in data. Running model without random effects.\n")
}

cat(paste("Model formula:", formula_string, "\n"))

# Clean data - remove rows with missing values
formula_vars_existing <- c("los_capped", existing_continuous, existing_categorical)
df_clean_nb <- df[, formula_vars_existing, drop = FALSE]
df_clean_nb <- df_clean_nb[complete.cases(df_clean_nb), ]

cat(paste("Number of rows after dropping missing values:", nrow(df_clean_nb), "\n"))

# Fit the multilevel truncated negative binomial model
cat("\nFitting multilevel truncated negative binomial regression model...\n")
trunc_nb_model <- glmmTMB(
  formula = as.formula(formula_string),
  data = df_clean_nb,
  family = truncated_nbinom2,
  REML = TRUE
)

# Print model summary
cat("\nMultilevel Truncated Negative Binomial Model Results:\n")
print(summary(trunc_nb_model))

# Extract fixed effects
cat("\nFixed Effects Coefficients:\n")
fixed_effects_summary <- summary(trunc_nb_model)$coefficients$cond
print(fixed_effects_summary)

# Calculate Incident Rate Ratios (IRR)
cat("\nIncident Rate Ratios (IRR) for Fixed Effects:\n")
irr_results <- data.frame(
  Variable = rownames(fixed_effects_summary),
  IRR = exp(fixed_effects_summary[, "Estimate"]),
  Lower_CI = exp(fixed_effects_summary[, "Estimate"] - 1.96 * fixed_effects_summary[, "Std. Error"]),
  Upper_CI = exp(fixed_effects_summary[, "Estimate"] + 1.96 * fixed_effects_summary[, "Std. Error"]),
  P_value = fixed_effects_summary[, "Pr(>|z|)"]
)
print(irr_results)

# Extract random effects
has_random_effects <- FALSE
if ("us_state_enc" %in% names(df)) {
  cat("\nRandom Effects Variance Components:\n")
  random_effects <- VarCorr(trunc_nb_model)
  print(random_effects)
  
  if (!is.null(random_effects$us_state_enc) && length(random_effects$us_state_enc) > 0) {
    state_var <- as.numeric(random_effects$us_state_enc[1])
    if (!is.na(state_var) && state_var >= 0) {
      has_random_effects <- TRUE
      dispersion <- sigma(trunc_nb_model) ^ 2
      icc <- state_var / (state_var + dispersion + (pi ^ 2) / 3)
      cat(paste("Intraclass Correlation Coefficient (ICC):", round(icc, 4), "\n"))
    } else {
      cat("Warning: Invalid or zero variance for us_state_enc random effect.\n")
    }
  } else {
    cat("Warning: No random effect variance estimated for us_state_enc.\n")
  }
}

# Model diagnostics
cat("\nModel Diagnostics:\n")
simulation_output <- simulateResiduals(fittedModel = trunc_nb_model, plot = FALSE)

# Save diagnostic plots
temp_dir <- tempdir()
residuals_plot_file <- file.path(temp_dir, "residuals_plot.png")
png(residuals_plot_file, width = 800, height = 600)
par(mfrow = c(2, 2))
plot(simulation_output, main = "DHARMa Residuals")
dev.off()

residuals_by_state_file <- file.path(temp_dir, "residuals_by_state.png")
png(residuals_by_state_file, width = 800, height = 600)
plotResiduals(simulation_output, form = df_clean_nb$us_state_enc, main = "Residuals by State")
dev.off()

# Model fit statistics
cat("\nModel Fit Statistics:\n")
cat(paste("AIC:", round(AIC(trunc_nb_model), 2), "\n"))
cat(paste("BIC:", round(BIC(trunc_nb_model), 2), "\n"))
cat(paste("Log-likelihood:", round(logLik(trunc_nb_model), 2), "\n"))

# Check for overdispersion
check_overdispersion(trunc_nb_model)

# Fit comparison model without truncation
cat("\nFitting comparison model without truncation...\n")
comparison_formula <- gsub("\\+ \\(1\\|us_state_enc\\)", "", formula_string)
comparison_formula <- paste(comparison_formula, "+ (1|us_state_enc)")
nb_model_comparison <- glmmTMB(
  formula = as.formula(comparison_formula),
  data = df_clean_nb,
  family = nbinom2,
  REML = TRUE
)

# Model comparison
cat("\nModel Comparison (Truncated vs Non-truncated):\n")
anova_results <- anova(trunc_nb_model, nb_model_comparison)
print(anova_results)

# Compute VIF
# Compute VIF
cat("\nComputing Variance Inflation Factor (VIF):\n")
# Fit a glm.nb model for VIF calculation (no random effects)
fixed_formula <- paste("los_capped ~", paste(formula_parts[formula_parts != "us_state_enc"], collapse = " + "))
vif_model <- tryCatch(
  glm.nb(as.formula(fixed_formula), data = df_clean_nb),
  error = function(e) {
    cat("Error fitting glm.nb model for VIF calculation:", e$message, "\n")
    return(NULL)
  }
)

if (!is.null(vif_model)) {
  vif_values <- tryCatch(
    vif(vif_model),
    error = function(e) {
      cat("Error computing VIF:", e$message, "\n")
      return(NULL)
    }
  )
  
  if (!is.null(vif_values)) {
    # Handle VIF output for categorical variables
    vif_matrix <- as.matrix(vif_values)
    predictor_names <- colnames(model.matrix(vif_model))[-1]  # Exclude intercept
    if (length(predictor_names) == nrow(vif_matrix)) {
      vif_results <- data.frame(
        Variable = predictor_names,
        VIF = vif_matrix[, 1]
      )
      cat("Note: VIF > 5 may indicate multicollinearity.\n")
      print(vif_results)
    } else {
      cat("Warning: Mismatch between VIF values and predictor names.\n")
      vif_results <- data.frame(
        Variable = paste("Predictor", seq_along(vif_matrix[, 1])),
        VIF = vif_matrix[, 1]
      )
      print(vif_results)
    }
  } else {
    cat("Warning: VIF calculation failed.\n")
    vif_results <- data.frame(Variable = "None", VIF = NA)
  }
} else {
  cat("Warning: glm.nb model fitting failed, skipping VIF calculation.\n")
  vif_results <- data.frame(Variable = "None", VIF = NA)
}
cat("Note: VIF > 5 may indicate multicollinearity.\n")
print(vif_results)

# Extract predictions and plot
df_clean_nb$predicted_truncated <- predict(trunc_nb_model, type = "response")
df_clean_nb$residuals_truncated <- residuals(trunc_nb_model, type = "pearson")

pred_vs_obs_plot <- ggplot(df_clean_nb, aes(x = predicted_truncated, y = los_capped)) +
  geom_point(alpha = 0.5) +
  geom_abline(intercept = 0, slope = 1, color = "red") +
  labs(title = "Predicted vs Observed Length of Stay", x = "Predicted LOS", y = "Observed LOS") +
  theme_minimal()

# Save predicted vs observed plot
pred_vs_obs_file <- file.path(temp_dir, "pred_vs_obs.png")
ggsave(pred_vs_obs_file, plot = pred_vs_obs_plot, width = 8, height = 6, dpi = 300)

# Save results
results_list <- list(
  model_summary = summary(trunc_nb_model),
  irr_results = irr_results,
  random_effects = if (has_random_effects) VarCorr(trunc_nb_model) else NULL,
  icc = if (has_random_effects) icc else NULL,
  fit_statistics = data.frame(
    AIC = AIC(trunc_nb_model),
    BIC = BIC(trunc_nb_model),
    LogLik = as.numeric(logLik(trunc_nb_model))
  ),
  vif_results = vif_results,
  anova_results = anova_results
)

# Ask user where to save results
cat("\nWould you like to save the results? (y/n): ")
save_choice <- readline()

if (tolower(save_choice) == "y") {
  output_file <- file.choose(new = TRUE)
  if (!is.null(output_file) && output_file != "") {
    # Save RDS file
    saveRDS(results_list, file = paste0(tools::file_path_sans_ext(output_file), ".rds"))

    # Save IRR as CSV
    write.csv(irr_results, file = paste0(tools::file_path_sans_ext(output_file), "_IRR.csv"), row.names = FALSE)

    # Create Word document
    doc <- read_docx()

    # Add title
    doc <- doc %>% body_add_par("Multilevel Truncated Negative Binomial Model Results", style = "heading 1")

    # Add Fixed Effects table
    doc <- doc %>% body_add_par("Fixed Effects Coefficients", style = "heading 2")
    fixed_effects_df <- as.data.frame(fixed_effects_summary)
    fixed_effects_df <- round(fixed_effects_df, 4)
    fixed_effects_df$Variable <- rownames(fixed_effects_summary)
    fixed_effects_ft <- flextable(fixed_effects_df) %>%
      set_header_labels(
        Variable = "Variable",
        Estimate = "Estimate",
        `Std. Error` = "Std. Error",
        `z value` = "z value",
        `Pr(>|z|)` = "P-value"
      ) %>%
      align(align = "right", part = "body", j = 2:5) %>%
      autofit()
    doc <- doc %>% body_add_flextable(fixed_effects_ft)

    # Add IRR table
    doc <- doc %>% body_add_par("Incident Rate Ratios (IRR) for Fixed Effects", style = "heading 2")
    irr_results_formatted <- irr_results
    irr_results_formatted$IRR <- formatC(irr_results$IRR, digits = 4, format = "f")
    irr_results_formatted$Lower_CI <- formatC(irr_results$Lower_CI, digits = 4, format = "f")
    irr_results_formatted$Upper_CI <- formatC(irr_results$Upper_CI, digits = 4, format = "f")
    irr_results_formatted$P_value <- formatC(irr_results$P_value, digits = 4, format = "f")
    irr_ft <- flextable(irr_results_formatted) %>%
      set_header_labels(
        Variable = "Variable",
        IRR = "IRR",
        Lower_CI = "Lower CI",
        Upper_CI = "Upper CI",
        P_value = "P-value"
      ) %>%
      align(align = "left", part = "body", j = 1) %>%
      align(align = "right", part = "body", j = 2:5) %>%
      colformat_double(j = 2:5, digits = 4) %>%
      width(j = 1, width = 2.5) %>%
      width(j = 2:5, width = 1) %>%
      fontsize(size = 10, part = "all") %>%
      set_caption("Incident Rate Ratios for Fixed Effects") %>%
      autofit()
    doc <- doc %>% body_add_flextable(irr_ft)

    # Add Random Effects
    if (has_random_effects) {
      doc <- doc %>% body_add_par("Random Effects Variance Components", style = "heading 2")
      random_effects_df <- data.frame(
        Group = "us_state_enc",
        Variance = state_var,
        StdDev = sqrt(state_var)
      )
      random_effects_ft <- flextable(random_effects_df) %>%
        autofit()
      doc <- doc %>% body_add_flextable(random_effects_ft)
      doc <- doc %>% body_add_par(paste("Intraclass Correlation Coefficient (ICC):", round(icc, 4)), style = "Normal")
    } else {
      doc <- doc %>% body_add_par("Random Effects Variance Components", style = "heading 2") %>%
        body_add_par("No random effect variance estimated for us_state_enc.", style = "Normal")
    }

    # Add VIF table
      # Add VIF table
    doc <- doc %>% body_add_par("Variance Inflation Factor (VIF)", style = "heading 2")
    if (!is.null(vif_results) && !all(is.na(vif_results$VIF))) {
      vif_results$VIF <- formatC(vif_results$VIF, digits = 2, format = "f")
      vif_ft <- flextable(vif_results) %>%
        set_header_labels(Variable = "Variable", VIF = "VIF") %>%
        align(align = "right", part = "body", j = 2) %>%
        set_caption("VIF Values (Note: VIF > 5 may indicate multicollinearity)") %>%
        autofit()
      doc <- doc %>% body_add_flextable(vif_ft)
    } else {
      doc <- doc %>% body_add_par("VIF calculation failed or returned no results.", style = "Normal")
    }

    # Add ANOVA results
    doc <- doc %>% body_add_par("Model Comparison (Truncated vs Non-truncated)", style = "heading 2")
    anova_df <- as.data.frame(anova_results)
    anova_df <- round(anova_df, 4)
    anova_ft <- flextable(anova_df) %>%
      set_caption("ANOVA Model Comparison Results") %>%
      autofit()
    doc <- doc %>% body_add_flextable(anova_ft)

    # Add Model Fit Statistics
    doc <- doc %>% body_add_par("Model Fit Statistics", style = "heading 2")
    fit_stats_df <- data.frame(
      Statistic = c("AIC", "BIC", "Log-likelihood"),
      Value = round(c(AIC(trunc_nb_model), BIC(trunc_nb_model), as.numeric(logLik(trunc_nb_model))), 2)
    )
    fit_stats_ft <- flextable(fit_stats_df) %>%
      autofit()
    doc <- doc %>% body_add_flextable(fit_stats_ft)

    # Add Predicted vs Observed Plot
    doc <- doc %>% body_add_par("Predicted vs Observed Length of Stay", style = "heading 2")
    doc <- doc %>% body_add_img(src = pred_vs_obs_file, width = 6, height = 4.5)

    # Add Diagnostic Plots
    doc <- doc %>% body_add_par("Diagnostic Plots", style = "heading 2")
    doc <- doc %>% body_add_par("DHARMa Residuals", style = "heading 3")
    doc <- doc %>% body_add_img(src = residuals_plot_file, width = 6, height = 4.5)
    doc <- doc %>% body_add_par("Residuals by State", style = "heading 3")
    doc <- doc %>% body_add_img(src = residuals_by_state_file, width = 6, height = 4.5)

    # Save Word document
    word_output <- paste0(tools::file_path_sans_ext(output_file), "_results.docx")
    print(doc, target = word_output)

    cat(paste("Results saved to:", output_file, "\n"))
    cat(paste("Word document saved to:", word_output, "\n"))
  }
}

cat("\nAnalysis complete.\n")