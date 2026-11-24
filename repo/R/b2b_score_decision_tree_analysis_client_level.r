install.packages(
     "DiagrammeR",
     lib  = "C:/Program Files/R/R-4.3.1/library",
     type = "binary"
 )

install.packages("ggparty")

######################################################################
# Load the functions
source("C:/Users/adamakis/OneDrive - Onity Group Inc/GitHub repos/B2B-Pricing-Model/B2B-Pricing-Model/repo/R/functions.R")
######################################################################
# Pull the data from SQL server
library(DBI)
library(odbc)

# Create SQL Server connection (Windows authentication)
con <- dbConnect(
  odbc(),
  Driver   = "ODBC Driver 17 for SQL Server",
  Server   = "VMC2VMDBRADC01P",
  Database = "Analytics_Rpt",
  Trusted_Connection = "Yes"
)

# SQL query
base_sql <- "
SELECT *
FROM Analytics_Rpt.RPT.B2BPricing_ClientMetrics_DT
"

# Read data into R
df <- dbGetQuery(con, base_sql)

# Close connection
dbDisconnect(con)
######################################################################
# View the table
dim(df)
View(head(df, 10))

library(dplyr)
glimpse(df)
######################################################################
# Feature engineering
# Step 1 — make sure date is Date
df$FundDt <- as.Date(df$FundDt)
# Step 2 — generate month starts
month_starts <- seq.Date(
  from = as.Date("2025-05-01"),
  to   = as.Date("2026-04-01"),
  by   = "month"
)
# Step 3 — convert to month ends
month_ends <- as.Date(format(month_starts + 32, "%Y-%m-01")) - 1
# Step 3 — loop and create features
col_names <- c()
for (i in seq_along(month_ends)) {
  d <- month_ends[i]   # keeps Date class
  col_name <- paste0("before_", format(d, "%Y_%m"))
  col_names <- c(col_names, col_name)   # store name
  df[[col_name]] <- ifelse(df$FundDt <= d, 1, 0)
}
cat(paste0('"', col_names, '"', collapse = ",\n"))
######################################################################
# Create Regional features
library(dplyr)

df <- df %>%
  mutate(
    Region_1 = case_when(
      Region %in% c("East") |
        PropState %in% c("NY", "SC", "NH", "MN") |
        DivisionName %in% c("Middle Atlantic") |
        StateRegion %in% c("Northeast") |
        RegionName %in% c("Northeast") ~ "High Regions",
      TRUE ~ "Low Regions"
    ),
    
    Region_2 = case_when(
      Region %in% c("East") |
        PropState %in% c("NY", "SC", "NH", "MN") |
        DivisionName %in% c("Middle Atlantic") |
        StateRegion %in% c("Northeast") ~ "High Regions",
      TRUE ~ "Low Regions"
    ),
    
    Region_3 = case_when(
      Region %in% c("East") |
        PropState %in% c("NY", "SC", "NH", "MN") |
        DivisionName %in% c("Middle Atlantic") ~ "High Regions",
      TRUE ~ "Low Regions"
    ),
    
    Region_4 = case_when(
      Region %in% c("East") |
        PropState %in% c("NY", "SC", "NH", "MN") ~ "High Regions",
      TRUE ~ "Low Regions"
    ),
    
    Region_5 = case_when(
      Region %in% c("East") ~ "High Regions",
      TRUE ~ "Low Regions"
    )
  )
######################################################################
# Create the predictors list
exclude_cols <- c("asofdt", "FundDt", "loanid", "InvestorLoanNum", "Client", target_column)
predictor_list <- setdiff(names(df), exclude_cols)
predictor_list
######################################################################
# Create a list of numeric and character columns
categorical_columns <- c(
#  , "Channel"
#  , "SubChannel"
#  , "AmortizationTerm_Band"
#  , "AmortizationType"
#  , "ARMIndex_Band"
#  , "LoanInvestor"
#  , "MSROwner"
#  , "LoanPurposeGroup"
#  , "LoanTypeGroup"
#  , "SelfEmployedFlag"
#  , "buydown_flag"
#  , "OccupancyGroup"
#  , "ProductCategory"
#  , "PropertyTypeGroup"
#   "StateRegion"
#   "Region"
#  , "PropState"
#   , "RegionName"
    "Region_1"
   , "Region_2"
   , "Region_3"
   , "Region_4"
   , "Region_5"
#   "DivisionName"
#  , "ServicingType"
#  , "before_2025_05"
#  , "before_2025_06"
#  , "before_2025_07"
#  , "before_2025_08"
#  , "before_2025_09"
#  , "before_2025_10"
#  , "before_2025_11"
#  , "before_2025_12"
#  , "before_2026_01"
#  , "before_2026_02"
#  , "before_2026_03"
#  , "before_2026_04"
)
numeric_columns <- c(
#  "BaseLoanAmt"
#  , "TotalLoanAmt"
#  , "CLTV"
#  , "DTI"
#  , "FICO"
#  , "LockDuration"
#  , "LTV"
)
target_column = c("B2B_score")
######################################################################
# Recode categorical columns as factors and numeric columns as numeric
df <- recode_columns(
  df = df,
  categorical_columns = categorical_columns,
  numeric_columns = numeric_columns
)

glimpse(df)
######################################################################
# Produce data quality report
summary_table <- feature_summary(df, c(categorical_columns, numeric_columns))
View(summary_table)
######################################################################
# Drop rows where the target is NaN
df <- drop_na_target(
  df = df,
  target = target_column
)
######################################################################
# Treat NA explicitly as a level so that they will appear in the rules
library(forcats)

df[categorical_columns] <- lapply(df[categorical_columns], function(x) {
  forcats::fct_na_value_to_level(x, level = "Missing")
})
View(feature_summary(df, c(categorical_columns, numeric_columns)))
######################################################################
# Create missing indicators for numeric variables because CTREE drops rows with missing values in each split
result <- create_missing_indicators(df, numeric_columns)
df <- result$data
missing_cols <- result$missing_cols
glimpse(df)
######################################################################
# Impute numeric feratures
#impute_numeric_result <- impute_numeric(df, numeric_columns)

#df <- impute_numeric_result$data
#medians_used <- impute_numeric_result$medians
#View(feature_summary(df, c(categorical_columns, numeric_columns)))
######################################################################
# Split the data
####
# Random split
#set.seed(123)
#train_idx <- sample(1:nrow(df), 0.8 * nrow(df))
#train_data <- df[train_idx, ]
#test_data  <- df[-train_idx, ]

# Date split
train_data <- df[df$FundDt < as.Date("2026-02-28"), ]
test_data  <- df[df$FundDt >= as.Date("2026-02-28"), ]

# No split
#train_data <- df
#test_data  <- df
######################################################################
# Drop features that have length one
categorical_all <- c(categorical_columns, missing_cols)
categorical_filtered <- categorical_all[
  sapply(train_data[categorical_all], function(x) {
    is.factor(x) && length(unique(x)) > 1
  })
]

numeric_all <- c(numeric_columns)
numeric_filtered <- numeric_all[
  sapply(train_data[numeric_all], function(x) {
    is.numeric(x) && length(unique(x)) > 1
  })
]
######################################################################
# Roll-up at a group level from loan level and recode all features for the train data
result_roll_up_train <- rollup_features(
  df = train_data,
  rollup_var = "Client",
  categorical_features = categorical_filtered,
  numeric_features = numeric_filtered, 
  target_var = target_column,
  min_loans_threshold = 10
)
train_df <- result_roll_up_train$data
num_feats <- result_roll_up_train$numeric_features
num_feats
cat_feats <- result_roll_up_train$categorical_features
cat_feats
######################################################################
# Roll-up at a group level from loan level and recode all features for the test data
result_roll_up_test <- rollup_features(
  df = test_data,
  rollup_var = "Client",
  categorical_features = categorical_filtered,
  numeric_features = numeric_filtered, 
  target_var = target_column,
  min_loans_threshold = 10
)
test_df <- result_roll_up_test$data
######################################################################
# Recode categorical columns as factors and numeric columns as numeric
train_df <- recode_columns(
  df = train_df,
  categorical_columns = cat_feats,
  numeric_columns = num_feats
)
test_df <- recode_columns(
  df = test_df,
  categorical_columns = cat_feats,
  numeric_columns = num_feats
)
glimpse(train_df)
######################################################################
# Produce data quality report
summary_table_roll_up <- feature_summary(train_df, c(cat_feats, num_feats))
View(summary_table_roll_up)
######################################################################
## Remove outliers from the target variable using IQR and business rules
result_iqr <- cap_iqr_df(
  df = train_df,
  features = target_column,
  iqr_coef = 0.5,
  min_cap = 0.1,
  max_cap = 5
)
dq_table <- result_iqr$dq_table
dq_table
train_df <- result_iqr$data

# Apply thresholds to test data
test_df <- apply_caps(test_df, dq_table)
######################################################################
# Develop the Decision Tree
library(rpart)
library(partykit)   # if you use ctree

# Set seed for reproducibility
set.seed(123)

params = list(
#  type = "rpart", params = list(control = rpart.control(maxdepth = 4,minbucket = 1000,cp = 0))
  type = "ctree", params = list(control = ctree_control(maxdepth = 10, minbucket = 4, alpha = 0.05, 
                                                        multiway = FALSE))
)

decision_tree_results = decision_tree_fun(
  df = train_df,
  target = target_column,
  predictor_list = list(
    categorical = cat_feats,
    numeric = num_feats
  ),
  tree_config = params,
#  client_id_col = "Client",
  client_id_col = NULL,
  return_model = TRUE
)
######################################################################
# Vusualize the Decision Tree
visualize_decision_tree(
  model = decision_tree_results$model,
  data_input = train_df,   
  target = target_column,
  client_id_col = "Client",
  model_type = params$type,
  maxdepth = NULL,
  show_node_stats = TRUE, 
  use_graphviz = FALSE, 
  rpart_tweak = 1,
  rpart_cex = 0.7, 
  ctree_fontsize = 10
)
######################################################################
# Validate/Test the Decision Tree - Applied only for CTREE
validate_tree_results = validate_tree(
  model = decision_tree_results$model,
  train_data = train_df, 
  test_data = test_df, 
  target = target_column, 
  abs_diff_thres = 0.25, 
  relative_diff_thres = 0.25, 
  test_pct_thres = 0.01, 
  test_count_thres = 100
)
validate_tree_results = validate_tree(
  model = decision_tree_results$model,
  train_data = train_df, 
  test_data = test_df, 
  target = "RecapScore", 
  abs_diff_thres = 0.25, 
  relative_diff_thres = 0.25, 
  test_pct_thres = 0.01, 
  test_count_thres = 100
)
validate_tree_results = validate_tree(
  model = decision_tree_results$model,
  train_data = train_df, 
  test_data = test_df, 
  target = "CPRScore", 
  abs_diff_thres = 0.25, 
  relative_diff_thres = 0.25, 
  test_pct_thres = 0.01, 
  test_count_thres = 100
)

View(validate_tree_results$comparison)
View(validate_tree_results$exec_table)
######################################################################
# Validation
get_mean_with_filter <- function(df, condition, feature) {
  # Apply filter
  filtered_df <- df[eval(substitute(condition), df, parent.frame()), ]
  # Compute mean
  mean_val <- mean(filtered_df[[feature]], na.rm = TRUE)
  median_val <- median(filtered_df[[feature]], na.rm = TRUE)
  print(c(mean_val, median_val))
  return(filtered_df)
}
filtered_df = get_mean_with_filter(
  df = train_data,
  condition = SubChannel %in% ('GNMA PIIT') & ProductCategory %in% c('FHA', 'USDA'),
  feature = target_column
#  feature = "RecapScore"
#  feature = "CPRScore"
)
View(filtered_df)
openxlsx::write.xlsx(filtered_df, "C:/Users/adamakis/OneDrive - Onity Group Inc/Desktop/Projects/B2B Pricing Model/output/DT_outputs/filtered_df.xlsx")
######################################################################
train_df %>%
  dplyr::filter(
#    SubChannel %in% c("GNMA PIIT") & ProductCategory %in% c("FHA", "USDA")
    PropState_NY > 0.091
  ) %>%
  View()
######################################################################





## EXPERIMENTAL ##

######################################################################
# Option A — ggalluvial
library(ggplot2)
library(ggalluvial)
library(dplyr)
library(tidyr)
# Specify the dataset to use
dataset = test_data

dataset$node <- predict(model, newdata = dataset, type = "node")
flow_data <- dataset %>%
  group_by(SubChannel, Region, node) %>%
  summarise(
    n = n(),
    KPI = mean(B2B_score, na.rm = TRUE),
    .groups = "drop"
  )
flow_data$node <- paste0("Segment ", flow_data$node)
flow_data <- flow_data %>%
  mutate(
    SubChannel = paste0("SC: ", SubChannel),
    Region = paste0("R: ", Region),
    node = paste0("Node: ", node)
  )
ggplot(flow_data,
       aes(axis1 = SubChannel, axis2 = Region, axis3 = node, y = n)) +
  
  geom_alluvium(aes(fill = KPI), alpha = 0.8) +
  
  geom_stratum() +
  
  geom_text(stat = "stratum",
            aes(label = gsub("SC: |R: |Node: ", "", after_stat(stratum)))) +
  
  scale_x_discrete(limits = c("SubChannel", "Region", "Segment")) +
  
  scale_fill_gradient2(
    low = "blue",
    mid = "white",
    high = "red",
    midpoint = mean(flow_data$KPI, na.rm = TRUE)
  ) +
  
  theme_minimal() +
  labs(fill = "KPI")
######################################################################






# Option B — networkD3 (more “executive” look)
library(networkD3)


# 3. Clean ggparty tree (best native tree option)
library(ggparty)

ggparty(decision_tree_results$model) +
  geom_edge() +
  geom_edge_label() +
  geom_node_splitvar() +
  geom_node_label() +
  theme_minimal()



validate_tree_results$exec_table[1:10,]









