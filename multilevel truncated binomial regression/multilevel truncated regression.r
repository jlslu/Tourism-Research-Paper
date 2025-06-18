# Load required libraries
library(glmmTMB)
library(readxl)
library(dplyr)
library(broom.mixed)
library(performance)
library(DHARMa)
library(ggplot2)

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
17
# Define continuous and categorical variables following your naming convention
continuous_vars <- c(
  "immigrant_population_log", "import_from_slu_log", "age",
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
# Build formula parts
formula_vars <- c("los_capped", continuous_vars, categorical_model_vars)

# Filter to only include variables that exist in the dataframe
existing_continuous <- continuous_vars[continuous_vars %in% names(df)]
cat_vars_in_df <- categorical_model_vars %in% names(df)
existing_categorical <- categorical_model_vars[cat_vars_in_df]

# Create formula string
formula_parts <- c(existing_continuous, existing_categorical)
formula_string <- paste("los_capped ~", paste(formula_parts, collapse = " + "))

# Add random effect for us_state_enc (multilevel component)
if ("us_state_enc" %in% names(df)) {
  # Remove us_state_enc from fixed effects and add as random effect
  fixed_effects <- formula_parts[formula_parts != "us_state_enc"]
  formula_string <- paste(
    "los_capped ~",
    paste(fixed_effects, collapse = " + "),
    "+ (1|us_state_enc)"
  )
} else {
  cat(
    "Warning: us_state_enc not found in data. ",
    "Running model without random effects.\n"
  )
}

cat(paste("Model formula:", formula_string, "\n"))

# Clean data - remove rows with missing values
formula_vars_existing <- c(
  "los_capped", existing_continuous, existing_categorical
)
df_clean_nb <- df[, formula_vars_existing, drop = FALSE]
df_clean_nb <- df_clean_nb[complete.cases(df_clean_nb), ]

cat(
  paste(
    "Number of rows after dropping missing values:",
    nrow(df_clean_nb), "\n"
  )
)

# Check for truncation point - assuming truncation at 0 (left truncation)
# For right truncation, you would specify the upper bound
cat("\nFitting multilevel truncated negative binomial regression model...\n")

# Fit the multilevel truncated negative binomial model using glmmTMB
# The truncated family in glmmTMB handles zero-truncated distributions
trunc_nb_model <- glmmTMB(
  formula = as.formula(formula_string),
  data = df_clean_nb,
  family = truncated_nbinom2,  # Zero-truncated negative binomial
  REML = TRUE  # Use REML for random effects estimation
)

# Print model summary
cat("\nMultilevel Truncated Negative Binomial Model Results:\n")
print(summary(trunc_nb_model))

# Extract and display key results
cat("\nFixed Effects Coefficients:\n")
fixed_effects_summary <- summary(trunc_nb_model)$coefficients$cond
print(fixed_effects_summary)

# Calculate Incident Rate Ratios (IRR) for fixed effects
cat("\nIncident Rate Ratios (IRR) for Fixed Effects:\n")
irr_results <- data.frame(
  Variable = rownames(fixed_effects_summary),
  IRR = exp(fixed_effects_summary[, "Estimate"]),
  Lower_CI = exp(
    fixed_effects_summary[, "Estimate"] -
      1.96 * fixed_effects_summary[, "Std. Error"]
  ),
  Upper_CI = exp(
    fixed_effects_summary[, "Estimate"] +
      1.96 * fixed_effects_summary[, "Std. Error"]
  ),
  P_value = fixed_effects_summary[, "Pr(>|z|)"]
)
print(irr_results)

# Extract random effects variance components
if ("us_state_enc" %in% names(df)) {
  cat("\nRandom Effects Variance Components:\n")
  random_effects <- VarCorr(trunc_nb_model)
  print(random_effects)

  #Calculate Intraclass Correlation Coefficient (ICC)
  state_var <- as.numeric(random_effects$us_state_enc[1])
  # For negative binomial, the residual variance is not directly available
  # We'll use the dispersion parameter
  dispersion <- sigma(trunc_nb_model) ^ 2
  icc <- state_var / (state_var + dispersion + (pi ^ 2) / 3)
  # Approximation for count models
  cat(paste(
    "Intraclass Correlation Coefficient (ICC):",
    round(icc, 4), "\n"
  ))

  #DHARMa residuals for GLMMs
  simulation_output <- simulateResiduals(
    fittedModel = trunc_nb_model,
    plot = FALSE
  )
}

# Model diagnostics
cat("\nModel Diagnostics:\n")

# DHARMa residuals for GLMMs
simulation_output <- simulateResiduals(
  fittedModel = trunc_nb_model, plot = FALSE
)

# Plot residuals
par(mfrow = c(2, 2))
plot(simulationOutput, main = "DHARMa Residuals")

# Additional diagnostic plots
plotResiduals(simulationOutput, form = df_clean_nb$us_state_enc, 
              main = "Residuals by State")

# Model fit statistics
cat("\nModel Fit Statistics:\n")
cat(paste("AIC:", round(AIC(trunc_nb_model), 2), "\n"))
cat(paste("BIC:", round(BIC(trunc_nb_model), 2), "\n"))
cat(paste("Log-likelihood:", round(logLik(trunc_nb_model), 2), "\n"))

# Check for overdispersion
check_overdispersion(trunc_nb_model)

# Compare with non-truncated model (if desired)
cat("\nFitting comparison model without truncation...\n")
comparison_formula <- gsub("\\+ \\(1\\|us_state_enc\\)", "", formula_string)
comparison_formula <- paste(comparison_formula, "+ (1|us_state_enc)")

nb_model_comparison <- glmmTMB(
  formula = as.formula(comparison_formula),
  data = df_clean_nb,
  family = nbinom2,  # Standard negative binomial
  REML = TRUE
)

# Model comparison
cat("\nModel Comparison (Truncated vs Non-truncated):\n")
anova(trunc_nb_model, nb_model_comparison)

# Extract predictions
df_clean_nb$predicted_truncated <- predict(trunc_nb_model, type = "response")
df_clean_nb$residuals_truncated <- residuals(trunc_nb_model, type = "pearson")

# Plot predictions vs observed
ggplot(df_clean_nb, aes(x = predicted_truncated, y = los_capped)) +
  geom_point(alpha = 0.5) +
  geom_abline(intercept = 0, slope = 1, color = "red") +
  labs(title = "Predicted vs Observed Length of Stay",
       x = "Predicted LOS", y = "Observed LOS") +
  theme_minimal()

# Save results to file
results_list <- list(
  model_summary = summary(trunc_nb_model),
  irr_results = irr_results,
  random_effects = if ("us_state_enc" %in% names(df)) {
    VarCorr(trunc_nb_model)
  } else {
    NULL
  },
  icc = if ("us_state_enc" %in% names(df)) icc else NULL,
  fit_statistics = data.frame(
    AIC = AIC(trunc_nb_model),
    BIC = BIC(trunc_nb_model),
    LogLik = as.numeric(logLik(trunc_nb_model))
  )
)

# Ask user where to save results
cat("\nWould you like to save the results? (y/n): ")
save_choice <- readline()

if (tolower(save_choice) == "y") {
  output_file <- file.choose(new = TRUE)
  if (!is.null(output_file) && output_file != "") {
    saveRDS(
      results_list,
      file = paste0(
        tools::file_path_sans_ext(output_file),
        ".rds"
      )
    )

    # Also save as CSV for easy viewing
    write.csv(
      irr_results,
      file = paste0(
        tools::file_path_sans_ext(output_file),
        "_IRR.csv"
      ),
      row.names = FALSE
    )

    cat(paste("Results saved to:", output_file, "\n"))
  }
}

cat("\nAnalysis complete.\n")