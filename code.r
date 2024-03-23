
install.packages("purrr")
install.packages("readxl")
install.packages("dplyr")
install.packages("ggplot2")
install.packages("ggthemes")


library(readxl)
library(dplyr)
library(purrr)

file_names <- c("C:\\Users\\tonis\\Downloads\\DOUTORADO\\Sheet-25-1Q.xlsx", "C:\\Users\\tonis\\Downloads\\DOUTORADO\\Sheet-25-2Q.xlsx", "C:\\Users\\tonis\\Downloads\\DOUTORADO\\Sheet-25-3Q.xlsx", "C:\\Users\\tonis\\Downloads\\DOUTORADO\\Sheet-25-4Q.xlsx")


# Function to read data from a single sheet of an Excel file
read_stock_data <- function(file, sheet) {
  read_excel(file, sheet = sheet) %>%
    rename(Date = 1, Time = 3, ClosingPrice = 7, Volume = 8, Returns = 10) %>%
    mutate(Stock = sheet, FileName = basename(file))
}

# Initialize an empty data frame to store the combined data
combined_data <- data.frame()

# Loop through each file and each sheet to read data
for (file in file_names) {
  # Get sheet names for the current file
  sheet_names <- excel_sheets(file)
  
  for (sheet in sheet_names) {
    stock_data <- tryCatch({
      read_stock_data(file, sheet)
    }, error = function(e) {
      NULL # Ignore sheets that do not exist or cause an error
    })
    if (!is.null(stock_data)) {
      combined_data <- rbind(combined_data, stock_data)
    }
  }
}

# View the first few rows of the combined data
head(combined_data)

# Convert ClosingPrice to numeric (if it's not already)
combined_data$ClosingPrice <- as.numeric(combined_data$ClosingPrice)



# Calculate absolute returns
combined_data <- combined_data %>%
  group_by(Stock, Date) %>%
  arrange(Date, Time) %>%
  mutate(
    AbsReturn = c(NA, abs(diff(ClosingPrice))),
    AbsReturnLag = lag(AbsReturn, 1)
  ) %>%
  ungroup()

# Calculate 5-minute variance for each stock on each day
combined_data <- combined_data %>%
  group_by(Stock, Date) %>%
  mutate(
    FiveMinInterval = ceiling(Time / 5),
    FiveMinVar = var(Returns, na.rm = TRUE)
  ) %>%
  ungroup()

combined_data <- combined_data %>%
  group_by(Stock, Date) %>%
  mutate(
    Zeta = exp(-AbsReturn * sqrt(390) / (10 * sqrt(FiveMinVar)))
  ) %>%
  ungroup()






# Calculate hat(p)^S_d for each day
ps_d_values <- combined_data %>%
  group_by(Date) %>%
  summarize(hat_psd = mean(Zeta, na.rm = TRUE))

# Checking the first few rows of the calculated values
head(ps_d_values)

library(ggplot2)

# Set a seed for reproducibility
set.seed(123)

# Sample approximately two-thirds of the rows
sampled_data <- merged_data_log %>% 
  sample_frac(size = 0.66)

# Check the number of rows in the sampled data
nrow(sampled_data)


# Adding log of hat_psd to ps_d_values
ps_d_values$log_hat_psd <- log(ps_d_values$hat_psd)

# Checking the first few rows
head(ps_d_values)

# Merge with daily average volume
merged_data_log <- merge(ps_d_values, daily_avg_volume, by = "Date")

# Checking the merged data
head(merged_data_log)

correlation_value_log <- cor(merged_data_log$log_hat_psd, merged_data_log$AvgVolume, use = "complete.obs")
print(correlation_value_log)


ggplot(merged_data_log, aes(x = AvgVolume, y = log_hat_psd)) +
  geom_point() +
  geom_smooth(method = "lm", color = "blue") +
  labs(
    title = "Correlation between Average Daily Volume and Log of Systematic Staleness",
    x = "Average Daily Volume",
    y = "Log of Systematic Staleness"
  ) +
  theme_minimal() +
  theme(
    plot.title = element_text(hjust = 0.5),
    axis.title = element_text(size = 12, face = "bold")
  )



# Calculate daily volatility (standard deviation of returns)
daily_volatility <- combined_data %>%
  group_by(Date, Stock) %>%
  summarize(Volatility = sd(Returns, na.rm = TRUE)) %>%
  ungroup()

# Checking the first few rows
head(daily_volatility)


# Merge with ps_d_values that includes log_hat_psd
merged_data_volatility <- merge(ps_d_values, daily_volatility, by = "Date")

# Checking the merged data
head(merged_data_volatility)


correlation_volatility <- cor(merged_data_volatility$log_hat_psd, merged_data_volatility$Volatility, use = "complete.obs")
print(correlation_volatility)


library(ggthemes) 

ggplot(merged_data_volatility, aes(x = Volatility, y = log_hat_psd)) +
  geom_point(alpha = 0.7, color = "grey") +  # Basic color for points
  geom_smooth(method = "lm", formula = y ~ x, color = "red", size = 1) +  # Smooth line in red
  labs(
    title = "Correlation between Daily Volatility and Log of Systematic Staleness",
    x = "Daily Volatility",
    y = "Log of Systematic Staleness"
  ) +
  theme_minimal() +
  theme(
    plot.title = element_text(hjust = 0.5, size = 14, face = "bold"),
    axis.title = element_text(size = 12, face = "bold"),
    legend.position = "none"  # Hiding the legend
  ) +
  theme_tufte()  # Tufte theme for a clean and elegant look







