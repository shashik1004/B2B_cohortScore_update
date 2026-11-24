######################################################################
# Recode categorical columns as factors and numeric columns as numeric
recode_columns <- function(df, categorical_columns, numeric_columns) {
  
  # Keep only columns that actually exist in df
  categorical_columns <- intersect(categorical_columns, names(df))
  numeric_columns <- intersect(numeric_columns, names(df))
  
  # Recode categorical columns as factors
  df[categorical_columns] <- lapply(df[categorical_columns], as.factor)
  
  # Recode numeric columns as numeric
  df[numeric_columns] <- lapply(df[numeric_columns], function(x) {
    as.numeric(as.character(x))
  })
  
  return(df)
}
######################################################################
# Drop rows where the target is NaN
drop_na_target <- function(df, target) {
  
  # Check target exists
  stopifnot(target %in% names(df))
  
  total_rows <- nrow(df)
  
  # Logical mask for non-missing target
  valid_mask <- !is.na(df[[target]])
  
  dropped <- sum(!valid_mask)
  dropped_pct <- if (total_rows > 0) {
    round(100 * dropped / total_rows, 2)
  } else {
    0
  }
  
  # Always print message
  message(
    sprintf(
      "[drop_na_target] Dropping %d rows (%.2f%%) with NA in target '%s'.",
      dropped, dropped_pct, target
    )
  )
  
  # Subset and reset row names
  df_clean <- df[valid_mask, , drop = FALSE]
  rownames(df_clean) <- NULL
  
  return(df_clean)
}
######################################################################
# Create missing indicators for numeric variables because CTREE drops rows with missing values in each split
create_missing_indicators <- function(df, numeric_cols) {
  new_cols <- c()
  for (var in numeric_cols) {
    new_col <- paste0(var, "_missing")
    df[[new_col]] <- as.factor(is.na(df[[var]]))
    new_cols <- c(new_cols, new_col)
  }
  return(list(
    data = df,
    missing_cols = new_cols
  ))
}
######################################################################
# Roll-up at a group level from loan level and recode all features
rollup_features <- function(
    df,
    rollup_var,
    categorical_features,
    numeric_features,
    target_var,
    min_loans_threshold = 10
) {
  
  library(dplyr)
  library(tidyr)
  
  # -----------------------------
  # Clean categorical upfront
  # -----------------------------
  df[categorical_features] <- lapply(
    df[categorical_features],
    function(x) ifelse(is.na(x), "Missing", as.character(x))
  )
  
  # -----------------------------
  # Base: exposure
  # -----------------------------
  base_df <- df %>%
    group_by(across(all_of(rollup_var))) %>%
    summarise(
      n_loans = n(),
      .groups = "drop"
    ) %>%
    mutate(
      small_client_flag = n_loans < 5
    ) %>%
    filter(n_loans >= min_loans_threshold)
  
  # -----------------------------
  # Target
  # -----------------------------
  target_df <- df %>%
    group_by(across(all_of(rollup_var))) %>%
    summarise(
      !!target_var := mean(.data[[target_var]], na.rm = TRUE),
      .groups = "drop"
    )
  
  # -----------------------------
  # Numeric features
  # -----------------------------
  numeric_features_clean <- setdiff(numeric_features, target_var)
  
  numeric_df <- df %>%
    group_by(across(all_of(rollup_var))) %>%
    summarise(
      across(
        all_of(numeric_features_clean),
        list(
          mean = ~mean(.x, na.rm = TRUE),
          median = ~median(.x, na.rm = TRUE),
          p75 = ~quantile(.x, 0.75, na.rm = TRUE),
          std_raw = ~sd(.x, na.rm = TRUE),
          missing_pct = ~mean(is.na(.x))
        ),
        .names = "{.col}_{.fn}"
      ),
      .groups = "drop"
    ) %>%
    left_join(base_df, by = rollup_var)
  
  # -----------------------------
  # Fix std / cv + missing
  # -----------------------------
  for (var in numeric_features_clean) {
    
    mean_col   <- paste0(var, "_mean")
    median_col <- paste0(var, "_median")
    p75_col    <- paste0(var, "_p75")
    std_col    <- paste0(var, "_std_raw")
    miss_col   <- paste0(var, "_missing_pct")
    
    # std
    numeric_df[[paste0(var, "_std")]] <- ifelse(
      numeric_df$n_loans > 1 & is.finite(numeric_df[[std_col]]),
      numeric_df[[std_col]],
      0
    )
    
    # cv
    numeric_df[[paste0(var, "_cv")]] <- ifelse(
      numeric_df$n_loans > 1 &
        is.finite(numeric_df[[std_col]]) &
        numeric_df[[mean_col]] != 0,
      numeric_df[[std_col]] / (numeric_df[[mean_col]] + 1e-6),
      0
    )
    
    # all-missing flag
    all_missing <- numeric_df[[miss_col]] == 1
    flag_col <- paste0(var, "_all_missing_flag")
    
    numeric_df[[flag_col]] <- all_missing
    
    # replace stats
    numeric_df[[mean_col]][all_missing]   <- 0
    numeric_df[[median_col]][all_missing] <- 0
    numeric_df[[p75_col]][all_missing]    <- 0
    
    # clean NaN / NA
    numeric_df[[mean_col]][!is.finite(numeric_df[[mean_col]])]     <- 0
    numeric_df[[median_col]][!is.finite(numeric_df[[median_col]])] <- 0
    numeric_df[[p75_col]][!is.finite(numeric_df[[p75_col]])]       <- 0
  }
  
  numeric_df <- numeric_df %>%
    select(-ends_with("_std_raw"), -n_loans, -small_client_flag)
  
  # -----------------------------
  # Categorical proportions
  # -----------------------------
  cat_list <- lapply(categorical_features, function(var) {
    
    df %>%
      group_by(across(all_of(c(rollup_var, var)))) %>%
      summarise(n = n(), .groups = "drop") %>%
      group_by(across(all_of(rollup_var))) %>%
      mutate(pct = n / sum(n)) %>%
      select(-n) %>%
      pivot_wider(
        names_from = all_of(var),
        values_from = pct,
        values_fill = 0,
        names_prefix = paste0(var, "_")
      )
    
  })
  
  cat_df <- Reduce(function(x, y) full_join(x, y, by = rollup_var), cat_list)
  
  # -----------------------------
  # Dominant category
  # -----------------------------
  dominant_list <- lapply(categorical_features, function(var) {
    
    df %>%
      group_by(across(all_of(c(rollup_var, var)))) %>%
      summarise(n = n(), .groups = "drop") %>%
      slice_max(n, n = 1, with_ties = FALSE) %>%
      select(-n) %>%
      rename(!!paste0(var, "_dominant") := all_of(var))
    
  })
  
  dominant_df <- Reduce(function(x, y) full_join(x, y, by = rollup_var), dominant_list)
  
  # FIX: force Missing AFTER aggregation
  dominant_cols <- grep("_dominant$", names(dominant_df), value = TRUE)
  dominant_df[dominant_cols] <- lapply(
    dominant_df[dominant_cols],
    function(x) ifelse(is.na(x), "Missing", x)
  )
  
  # -----------------------------
  # Final dataset
  # -----------------------------
  final_df <- base_df %>%
    left_join(target_df, by = rollup_var) %>%
    left_join(numeric_df, by = rollup_var) %>%
    left_join(cat_df, by = rollup_var) %>%
    left_join(dominant_df, by = rollup_var)
  
  # FINAL SAFETY CLEAN (CRITICAL)
  char_cols <- names(final_df)[sapply(final_df, is.character)]
  final_df[char_cols] <- lapply(
    final_df[char_cols],
    function(x) ifelse(is.na(x), "Missing", x)
  )
  
  # Convert flags to character
  flag_cols <- grep("flag$", names(final_df), value = TRUE)
  final_df[flag_cols] <- lapply(final_df[flag_cols], as.character)
  colnames(final_df) <- make.names(colnames(final_df))
  
  # -----------------------------
  # Feature lists
  # -----------------------------
  numeric_features_out <- names(final_df)[sapply(final_df, is.numeric)]
  character_features_out <- names(final_df)[sapply(final_df, function(x) is.character(x) | is.factor(x))]
  
  numeric_features_out <- setdiff(numeric_features_out, c(rollup_var, target_var))
  character_features_out <- setdiff(character_features_out, c(rollup_var, target_var))
  
  return(list(
    data = final_df,
    numeric_features = numeric_features_out,
    categorical_features = character_features_out
  ))
}
######################################################################
# Produce data quality report
feature_summary <- function(df, predictor_list) {
  
  # Combine all features
  features <- predictor_list
  features <- intersect(features, names(df))
  
  summary_list <- lapply(features, function(var) {
    
    x <- df[[var]]
    
    # Missing %
    missing_pct_num <- 100 * sum(is.na(x)) / length(x)
    missing_pct <- sprintf("%.2f%%", missing_pct_num)
    
    # Distinct count (excluding NA)
    distinct_n <- length(unique(x[!is.na(x)]))
    
    # Initialize
    min_val <- NA
    max_val <- NA
    mean_val <- NA
    median_val <- NA
    
    # Numeric stats
    if (is.numeric(x)) {
      min_val <- sprintf("%.2f", min(x, na.rm = TRUE))
      max_val <- sprintf("%.2f", max(x, na.rm = TRUE))
      mean_val <- sprintf("%.2f", mean(x, na.rm = TRUE))
      median_val <- sprintf("%.2f", median(x, na.rm = TRUE))
      
      # Handle all-NA case
      if (all(is.na(x))) {
        min_val <- max_val <- mean_val <- median_val <- NA
      }
    }
    
    data.frame(
      feature = var,
      missing_pct = missing_pct,
      missing_pct_num = missing_pct_num,  # helper column
      distinct_values = distinct_n,
      min = min_val,
      max = max_val,
      mean = mean_val, 
      median = median_val, 
      stringsAsFactors = FALSE
    )
  })
  
  summary_df <- do.call(rbind, summary_list)
  
  # Order by missing %
  summary_df <- summary_df[order(-summary_df$missing_pct_num), ]
  
  # Drop helper column
  summary_df$missing_pct_num <- NULL
  
  return(summary_df)
}
######################################################################
# Impute numeric feratures
impute_numeric <- function(df, numeric_cols) {
  medians <- list()
  for (var in numeric_cols) {
    if (var %in% names(df) && is.numeric(df[[var]])) {
      med <- median(df[[var]], na.rm = TRUE)
      # Store median
      medians[[var]] <- med
      # Replace NA
      df[[var]][is.na(df[[var]])] <- med
    }
  }
  return(list(
    data = df,
    medians = medians
  ))
}
######################################################################
# Remove outliers using IQR
cap_iqr_df <- function(df, features, iqr_coef = 1.5, min_cap = -Inf, max_cap = Inf) {
  
  df_out <- df
  dq_list <- list()
  
  for (var in features) {
    
    x <- df[[var]]
    
    if (!is.numeric(x)) next
    
    # Compute IQR bounds
    q1 <- quantile(x, 0.01, na.rm = TRUE)
    q3 <- quantile(x, 0.99, na.rm = TRUE)
    iqr <- q3 - q1
    
    lower_iqr <- q1 - iqr_coef * iqr
    upper_iqr <- q3 + iqr_coef * iqr
    
    # Apply global caps
    lower <- max(lower_iqr, min_cap)
    upper <- min(upper_iqr, max_cap)
    
    # Apply capping
    x_capped <- pmax(x, lower)
    x_capped <- pmin(x_capped, upper)
    
    # Data quality metrics
    replaced_min <- sum(x < lower, na.rm = TRUE)
    replaced_max <- sum(x > upper, na.rm = TRUE)
    total_n <- sum(!is.na(x))
    
    dq_list[[var]] <- data.frame(
      feature = var,
      pct_replaced_min = sprintf("%.2f%%", 100 * replaced_min / total_n),
      pct_replaced_max = sprintf("%.2f%%", 100 * replaced_max / total_n),
      min_applied = round(lower, 4),
      max_applied = round(upper, 4),
      stringsAsFactors = FALSE
    )
    
    # Update dataframe
    df_out[[var]] <- x_capped
  }
  
  dq_table <- do.call(rbind, dq_list)
  
  return(list(
    data = df_out,
    dq_table = dq_table
  ))
}
######################################################################
# Apply thresholds to test data
apply_caps <- function(df, caps_table) {
  df_out <- df
  for (i in 1:nrow(caps_table)) {
    var <- caps_table$feature[i]
    lower <- caps_table$min_applied[i]
    upper <- caps_table$max_applied[i]
    if (var %in% names(df_out)) {
      df_out[[var]] <- pmax(df_out[[var]], lower)
      df_out[[var]] <- pmin(df_out[[var]], upper)
    }
  }
  return(df_out)
}
######################################################################
# Develop the Decision Tree
decision_tree_fun <- function(
    df,
    target,
    predictor_list,
    tree_config,
    client_id_col = NULL,
    return_model = TRUE
) {
  
  # -----------------------------
  # Validate inputs
  # -----------------------------
  stopifnot(target %in% names(df))
  stopifnot(all(c("categorical", "numeric") %in% names(predictor_list)))
  stopifnot(tree_config$type %in% c("rpart", "ctree"))
  
  categorical_cols <- intersect(predictor_list$categorical, names(df))
  numeric_cols     <- intersect(predictor_list$numeric, names(df))
  predictors       <- c(categorical_cols, numeric_cols)
  
  # -----------------------------
  # Drop rows where target is NA
  # -----------------------------
  valid_mask <- !is.na(df[[target]])
  dropped <- sum(!valid_mask)
  dropped_pct <- round(100 * dropped / nrow(df), 2)
  
  message(
    sprintf(
      "[decision_tree_fun] Dropping %d rows (%.2f%%) with NA in target '%s'.",
      dropped, dropped_pct, target
    )
  )
  
  df_clean <- df[valid_mask, , drop = FALSE]
  rownames(df_clean) <- NULL
  
  # -----------------------------
  # CLEAN COLUMN NAMES (CRITICAL FIX)
  # -----------------------------
  colnames(df_clean) <- make.names(colnames(df_clean))
  
  # also update variable names used later
  target <- make.names(target)
  categorical_cols <- make.names(categorical_cols)
  numeric_cols     <- make.names(numeric_cols)
  predictors       <- c(categorical_cols, numeric_cols)

  # -----------------------------
  # Recode columns
  # -----------------------------
  df_clean[categorical_cols] <- lapply(df_clean[categorical_cols], as.factor)
  df_clean[numeric_cols] <- lapply(df_clean[numeric_cols], function(x) {
    as.numeric(as.character(x))
  })
  
  # -----------------------------
  # Optional: client-level weights
  # -----------------------------
  sample_weight <- NULL
  if (!is.null(client_id_col)) {
    stopifnot(client_id_col %in% names(df_clean))
    client_counts <- table(df_clean[[client_id_col]])
#    sample_weight <- 1 / as.numeric(client_counts[df_clean[[client_id_col]]])
    sample_weight <- 1 / sqrt(as.numeric(client_counts[df_clean[[client_id_col]]]))
  }
  
  # -----------------------------
  # Build formula
  # -----------------------------
  formula_str <- paste(target, "~", paste(predictors, collapse = " + "))
  model_formula <- as.formula(formula_str)
  
  # -----------------------------
  # Train model
  # -----------------------------
  
  # =====================================================
  # rpart IMPLEMENTATION
  # =====================================================
  if (tree_config$type == "rpart") {
    
    params <- if (!is.null(tree_config$params)) tree_config$params else list()
    
    model <- do.call(
      rpart,
      c(
        list(
          formula = model_formula,
          data    = df_clean,
          method  = "anova",
          weights = sample_weight
        ),
        params
      )
    )
    
    predictions <- predict(model, df_clean)
    
    if (!is.null(model$variable.importance)) {
      feature_importance <- data.frame(
        feature = names(model$variable.importance),
        importance = as.numeric(model$variable.importance),
        row.names = NULL
      )
    } else {
      feature_importance <- data.frame(
        feature = character(),
        importance = numeric(),
        row.names = NULL
      )
      message("[decision_tree_fun] rpart did not compute variable importance (no effective splits).")
    }
    
  } 
  
  # =====================================================
  # ctree IMPLEMENTATION
  # =====================================================
  else if (tree_config$type == "ctree") {
    
    params <- if (!is.null(tree_config$params)) tree_config$params else list()
    
    model <- do.call(
      ctree,
      c(
        list(
          formula = model_formula,
          data    = df_clean,
          weights = sample_weight
        ),
        params
      )
    )
    
    predictions <- predict(model)
    
    vi <- varimp(model)
    
    if (!is.null(vi) && length(vi) > 0) {
      feature_importance <- data.frame(
        feature = names(vi),
        importance = as.numeric(vi),
        row.names = NULL
      )
    } else {
      feature_importance <- data.frame(
        feature = character(),
        importance = numeric(),
        row.names = NULL
      )
      message("[decision_tree_fun] ctree did not compute variable importance (no effective splits).")
    }
    
  }
  
  # -----------------------------
  # Assemble results
  # -----------------------------
  feature_importance <- feature_importance[
    order(feature_importance$importance, decreasing = TRUE),
  ]
  
  results <- list(
    predictions = predictions,
    feature_importance = feature_importance,
    cat_mappings = lapply(df_clean[categorical_cols], levels)
  )
  
  if (return_model) {
    results$model <- model
  }
  
  print(feature_importance)
  
  return(results)
}
######################################################################
# Vusualize the Decision Tree
visualize_decision_tree <- function(
    model,
    data_input,   
    target,
    client_id_col = NULL,
    model_type = c("rpart", "ctree"),
    maxdepth = NULL,
    show_node_stats = TRUE,
    use_graphviz = FALSE, 
    rpart_tweak = 1.2,
    rpart_cex = 1.5, 
    ctree_fontsize = 12
) {
  
  model_type <- match.arg(model_type)
  
  # -----------------------------
  # rpart visualization
  # -----------------------------
  if (model_type == "rpart") {
    
    if (!requireNamespace("rpart.plot", quietly = TRUE)) {
      stop("Package 'rpart.plot' is required.")
    }
    
    if (!use_graphviz) {
      
      rpart.plot::rpart.plot(
        model,
        type = 2,
        extra = if (show_node_stats) 101 else 0,
        fallen.leaves = TRUE,
        digits = 3,
        roundint = FALSE,
        tweak = rpart_tweak,
        cex = rpart_cex
      )
      
    } else {
      
      if (!requireNamespace("DiagrammeR", quietly = TRUE)) {
        stop("Package 'DiagrammeR' is required.")
      }
      
      dot <- rpart.plot::prp(
        model,
        type = 2,
        extra = if (show_node_stats) 101 else 0,
        fallen.leaves = TRUE,
        digits = 3,
        nn = FALSE,
        branch = 1,
        faclen = 0,
        compress = TRUE,
        box.col = "lightblue",
        split.col = "black",
        cex = rpart_cex
      )
      
      DiagrammeR::grViz(dot)
    }
  }
  
  # -----------------------------
  # ctree visualization
  # -----------------------------
  if (model_type == "ctree") {
    
    if (!requireNamespace("partykit", quietly = TRUE)) {
      stop("Package 'partykit' is required.")
    }
    
    nodes_all <- partykit::nodeids(model, terminal = FALSE)
	# use actual data
	data <- data_input
	node_assign <- predict(model, newdata = data, type = "node")
    # node_assign <- predict(model, type = "node")
    # data <- partykit::data_party(model)
    total_n <- nrow(data)
	# Ensure correct column name
	target <- make.names(target)
	if (!(target %in% names(data))) {
	  stop("Target column not found in data_input")
	}
	y_all <- as.numeric(data[[target]])
    
    # Correct node aggregation
    node_stats <- data.frame(
      node = nodes_all,
      n = sapply(nodes_all, function(id)
        sum(node_assign %in% partykit::nodeids(partykit::node_party(model), from = id))
      )
    )
    
    node_stats$avg <- sapply(nodes_all, function(id) {
      obs <- node_assign %in% partykit::nodeids(partykit::node_party(model), from = id)
	if (sum(obs) == 0) {NA} else {mean(y_all[obs], na.rm = TRUE)}
	})
	
    node_stats$median <- sapply(nodes_all, function(id) {
      obs <- node_assign %in% partykit::nodeids(partykit::node_party(model), from = id)
	if (sum(obs) == 0) {NA} else {median(y_all[obs], na.rm = TRUE)}
    })

	if (!is.null(client_id_col) && client_id_col %in% names(data)) {
	  node_stats$max_client_pct <- sapply(nodes_all, function(id) {
		obs <- node_assign %in% partykit::nodeids(partykit::node_party(model), from = id)
		node_clients <- data[[client_id_col]][obs]
		# Remove NA clients (important)
		node_clients <- node_clients[!is.na(node_clients)]
		# If no valid clients → NA
		if (length(node_clients) == 0) return(NA)
		client_counts <- table(node_clients)
		max(client_counts) / sum(client_counts)
	  })
	} else {
	  warning("client_id_col not found in data. max_client_pct set to NA.")
	  node_stats$max_client_pct <- NA
	}
    
    node_stats$pct <- node_stats$n / total_n
    
    node_stats_list <- split(node_stats, node_stats$node)
    
    if (!use_graphviz) {
      
      plot(
        model,
        gp = grid::gpar(fontsize = ctree_fontsize),
        
        terminal_panel = function(node) {
          id <- partykit::id_node(node)
          stats <- node_stats_list[[as.character(id)]]
          
          grid::grid.text(
            paste0(
              "Node ", id,
              "\n n = ", sprintf("%.1f%%", 100 * stats$pct),
              " (", format(stats$n, big.mark = ","), ")",
              "\n avg = ", ifelse(is.na(stats$avg), "NA", round(stats$avg, 3)), 
              "\n med = ", ifelse(is.na(stats$avg), "NA", round(stats$median, 3)), 
              "\n max client = ", ifelse(
                is.na(stats$max_client_pct),
                "NA",
                sprintf("%.1f%%", 100 * stats$max_client_pct))
            ),
            gp = grid::gpar(fontsize = ctree_fontsize)
          )
        },
        
        inner_panel = function(node) {
          split <- partykit::split_node(node)
          if (is.null(split)) return()
          
          id <- partykit::id_node(node)
          stats <- node_stats_list[[as.character(id)]]
          
          data <- partykit::data_party(model)
          varname <- names(data)[split$varid]
          
          # Clean split text (no buggy reconstruction)
          split_txt <- if (!is.null(split$breaks)) {
            paste0(varname, " <= ", round(split$breaks, 3))
          } else {
            varname
          }
          
          grid::grid.text(
            paste0(
              "Node ", id,
              "\n", split_txt,
              "\n n = ", sprintf("%.1f%%", 100 * stats$pct),
              " (", format(stats$n, big.mark = ","), ")",
              "\n avg = ", ifelse(is.na(stats$avg), "NA", round(stats$avg, 3)), 
              "\n med = ", ifelse(is.na(stats$avg), "NA", round(stats$median, 3)), 
              "\n max client = ", ifelse(
                is.na(stats$max_client_pct),
                "NA",
                sprintf("%.1f%%", 100 * stats$max_client_pct)
              )
            ),
            gp = grid::gpar(fontsize = ctree_fontsize)
          )
        },
        
        legend = NULL
      )
      
      grid::grid.remove("legend", warn = FALSE)
      
    } else {
      
      if (!requireNamespace("DiagrammeR", quietly = TRUE)) {
        stop("Package 'DiagrammeR' is required.")
      }
      
      DiagrammeR::grViz(
        partykit:::.plot.party(
          model,
          terminal_panel = partykit::node_terminal(
            digits = 3,
            nobs = show_node_stats,
            gp = grid::gpar(fontsize = ctree_fontsize)
          ),
          inner_panel = partykit::node_inner(
            digits = 3,
            pval = FALSE,
            gp = grid::gpar(fontsize = ctree_fontsize)
          ),
          legend = NULL,
          return = TRUE
        )
      )
    }
  }
  
  print(model)
  print(node_stats)
}
######################################################################
# Validate/Test the Decision Tree - Applied only for CTREE
validate_tree <- function(
    model,
    train_data, 
    test_data, 
    target, 
    abs_diff_thres, 
    relative_diff_thres, 
    test_pct_thres, 
    test_count_thres
) {
  # -----------------------------
  # Align column names with model 
  # -----------------------------
  colnames(train_data) <- make.names(colnames(train_data))
  colnames(test_data)  <- make.names(colnames(test_data))
  target <- make.names(target)

  # Step 1 — Apply tree to BOTH datasets
  train_nodes <- predict(model, newdata = train_data, type = "node")
  test_nodes  <- predict(model, newdata = test_data,  type = "node")
  
  # Step 2 — Compare node-level performance
  # get ALL nodes (not just leaves)
  nodes_all <- partykit::nodeids(model, terminal = FALSE)
  
  # helper function to get observations in each node
#  get_obs <- function(node_id, node_assign, model) {
#    node_assign %in% partykit::nodeids(model, from = node_id)
#  }
  get_obs <- function(node_id, node_assign, model) {
    node_assign %in% partykit::nodeids(partykit::node_party(model), from = node_id)
  }
  

  # Training
  train_summary <- data.frame(
    node = nodes_all,
    train_avg = sapply(nodes_all, function(id) {
      obs <- get_obs(id, train_nodes, model)
      mean(train_data[[target]][obs], na.rm = TRUE)
    }),
    train_median = sapply(nodes_all, function(id) {
      obs <- get_obs(id, train_nodes, model)
      median(train_data[[target]][obs], na.rm = TRUE)
    }),
    train_n = sapply(nodes_all, function(id) {
      sum(get_obs(id, train_nodes, model))
    })
  )
#  colnames(train_summary)[2] <- "train_avg"
  # Test
  test_summary <- data.frame(
    node = nodes_all,
    test_avg = sapply(nodes_all, function(id) {
      obs <- get_obs(id, test_nodes, model)
      mean(test_data[[target]][obs], na.rm = TRUE)
    }),
    test_median = sapply(nodes_all, function(id) {
      obs <- get_obs(id, test_nodes, model)
      median(test_data[[target]][obs], na.rm = TRUE)
    }),
    test_n = sapply(nodes_all, function(id) {
      sum(get_obs(id, test_nodes, model))
    })
  )
#  colnames(test_summary)[2] <- "test_avg"
  
  # Step 3 — Merge
  comparison <- merge(train_summary, test_summary, by = "node", all = TRUE)

  # Skew metrics
  comparison$train_skew <- with(comparison,
                                ifelse(train_median == 0, NA,
                                       abs(train_avg - train_median) / abs(train_median))
  )
  
  comparison$test_skew <- with(comparison,
                               ifelse(test_median == 0, NA,
                                      abs(test_avg - test_median) / abs(test_median))
  )
  # Format as %
  comparison$train_skew_pct <- sprintf("%.1f%%", 100 * comparison$train_skew)
  comparison$test_skew_pct  <- sprintf("%.1f%%", 100 * comparison$test_skew)

  comparison$train_median <- round(comparison$train_median, 2)
  comparison$test_median  <- round(comparison$test_median, 2)

  # Add percentages
  comparison$train_pct <- sprintf("%.1f%%", 100 * comparison$train_n / nrow(train_data))
  comparison$test_pct  <- sprintf("%.1f%%", 100 * comparison$test_n  / nrow(test_data))

  # Convert % to numeric
  comparison$test_pct_num  <- as.numeric(sub("%", "", comparison$test_pct))
  comparison$train_pct_num <- as.numeric(sub("%", "", comparison$train_pct))
  # Absolute Differences
  comparison$abs_diff <- abs(comparison$test_avg - comparison$train_avg)
  # Relative difference
  comparison$relative_diff <- with(comparison,
                                   ifelse(train_avg == 0, NA,
                                          abs(test_avg - train_avg) / abs(train_avg))
  )
  
  comparison$relative_diff_pct <- sprintf(
    "%.1f%%",
    100 * comparison$relative_diff
  )

  # Flag unstable nodes
  comparison$unstable_reason <- apply(comparison, 1, function(row) {
    reasons <- c()
    if (!is.na(row["abs_diff"]) && as.numeric(row["abs_diff"]) > abs_diff_thres) {
      reasons <- c(reasons, "abs_diff")
    }
    if (!is.na(row["relative_diff"]) && as.numeric(row["relative_diff"]) > relative_diff_thres) {
      reasons <- c(reasons, "relative_diff")
    }
    if (!is.na(row["test_pct_num"]) && as.numeric(row["test_pct_num"]) < (test_pct_thres*100)) {
      reasons <- c(reasons, "low_test_pct")
    }
    if (!is.na(row["test_n"]) && as.numeric(row["test_n"]) < test_count_thres) {
      reasons <- c(reasons, "low_test_n")
    }
    if (!is.na(row["train_skew"]) &&  (as.numeric(row["train_skew"])>0.5)) {
      reasons <- c(reasons, "train_skew")
    }
    if (!is.na(row["test_skew"]) &&  (as.numeric(row["test_skew"])>0.5)) {
      reasons <- c(reasons, "train_skew")
    }
    if (length(reasons) == 0) {
      return("stable")
    } else {
      return(paste(reasons, collapse = "; "))
    }
  })
  comparison$unstable <- comparison$unstable_reason != "stable"

  # Formatting
  comparison$train_avg <- round(comparison$train_avg, 2)
  comparison$test_avg  <- round(comparison$test_avg, 2)
  comparison$abs_diff  <- round(comparison$abs_diff, 2)
  
  # Cleanup
  comparison <- comparison[, !names(comparison) %in% 
                             c("test_pct_num", "train_pct_num", "relative_diff")]

  comparison <- comparison[, c(
    "node",
    "unstable",
    "unstable_reason",
    "train_avg",
    "train_median",
    "train_skew_pct",
    "test_avg",
    "test_median",
    "test_skew_pct",
    "abs_diff",
    "relative_diff_pct",
    "train_n",
    "test_n",
    "train_pct",
    "test_pct"
  )]
    
  # Produce an Executive ready table
  # Step 1 — Extract rules from tree
  data_full <- partykit::data_party(model)
  get_all_rules <- function(node, data_full, path = list()) {
    
    node_id <- partykit::id_node(node)
    rules <- list()
    
    # Convert path list to string for output
    if (length(path) == 0) {
      rules[[as.character(node_id)]] <- ""
    } else {
      rules[[as.character(node_id)]] <- paste(
        sapply(names(path), function(v) {
          
          val <- path[[v]]
          
          if (is.list(val)) {
            # numeric range
            lower <- val$lower
            upper <- val$upper
            
            if (is.finite(lower) & is.finite(upper)) {
              paste0(lower, " < ", v, " <= ", upper)
            } else if (is.finite(lower)) {
              paste0(v, " > ", lower)
            } else if (is.finite(upper)) {
              paste0(v, " <= ", upper)
            } else {
              v
            }
            
          } else {
            # categorical
            paste0(v, " in ('", paste(val, collapse = "', '"), "')")
          }
          
        }),
        collapse = " & "
      )
      
    }
    
    if (!partykit::is.terminal(node)) {
      
      split <- partykit::split_node(node)
      kids  <- partykit::kids_node(node)
      varname <- names(data_full)[split$varid]
      
      for (i in seq_along(kids)) {
        
        new_path <- path
        
        if (!is.null(split$breaks)) {
          
          if (!(varname %in% names(new_path))) {
            new_path[[varname]] <- list(lower = -Inf, upper = Inf)
          }
          
          if (i == 1) {
            # <= split
            new_path[[varname]]$upper <- min(new_path[[varname]]$upper, split$breaks)
          } else {
            # > split
            new_path[[varname]]$lower <- max(new_path[[varname]]$lower, split$breaks)
          }
        }
        else if (!is.null(split$index)) {
          
          levels_var <- levels(data_full[[varname]])
          grp <- unique(levels_var[split$index == i])
          
          node_data <- data_full[get_obs(node_id, train_nodes, model), varname]
          actual_values <- unique(node_data)
          
          grp <- intersect(grp, actual_values)
          
          if (length(grp) > 0) {
            
            if (varname %in% names(new_path)) {
              # INTERSECTION happens here
              new_path[[varname]] <- intersect(new_path[[varname]], grp)
            } else {
              new_path[[varname]] <- grp
            }
          }
        }
        
        rules <- append(
          rules,
          get_all_rules(kids[[i]], data_full, new_path)
        )
      }
    }
    
    return(rules)
  }

  rules_list <- get_all_rules(partykit::node_party(model), data_full)

  rules_df <- data.frame(
    node = as.integer(names(rules_list)),
    rule = unlist(rules_list),
    stringsAsFactors = FALSE
  )
  
  rules_df$rule[rules_df$rule == ""] <- "All Customers"

  # Step 2 — merge with your validation output
  exec_table <- merge(comparison, rules_df, by = "node", all.x = TRUE)
  # Step 3 — keep only stable nodes
#  exec_table <- exec_table[!exec_table$unstable, ]
  # Step 4 — clean + rename for business
  exec_table$Segment <- paste0("Segment ", exec_table$node)
  exec_table <- exec_table[, c(
    "Segment",
    "rule",
    "train_pct",
    "test_pct",
    "train_avg",
    "train_median",
    "test_avg",
    "test_median",
    "relative_diff_pct"
  )]
  
  colnames(exec_table) <- c(
    "Segment",
    "Business Rule",
    "Train Population %",
    "Test Population %",
    "KPI Train Avg",
    "KPI Train Med",
    "KPI Test Avg",
    "KPI Test Med",
    "Stability"
  )

  # Produce a side by side comparison of train and test averages
  # Step 1 — reshape (only change needed before plot)
  library(tidyr)
  library(dplyr)
  plot_data <- exec_table %>%
    pivot_longer(
      cols = c("KPI Train Avg", "KPI Test Avg"),
      names_to = "Type",
      values_to = "KPI"
    ) %>%
    mutate(
      pop_pct = ifelse(Type == "KPI Train", `Train Population %`, `Test Population %`)
    )
  # Step 2 — produce the plot with side by side bars
  library(ggplot2)
  print(
    ggplot(plot_data, aes(x = Segment, y = KPI, fill = Type)) +
      geom_col(position = "dodge") +
      geom_text(
        aes(label = paste0(round(KPI, 2), "\n(", pop_pct, ")")),
        position = position_dodge(width = 0.9),
        vjust = -0.5
      ) +
      theme_minimal() +
      labs(
        title = "Segment Performance",
        y = "KPI",
        x = "",
        fill = ""
      )
  )
  
  return(list(
    comparison = comparison,
    exec_table = exec_table
  ))

}
######################################################################
######################################################################
######################################################################