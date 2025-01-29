#Code compilation of our teams analysis on the impact of Covid-19 on the workforce.

# packages
library(tidyverse)
library(readxl)
library(moments)
library(ggplot2)
library(caret)
library(randomForest)

# Data set we chose was the Impact of COVID-19 on Working Professionals data set. The data can be found at:
# https://www.kaggle.com/datasets/gcreatives/impact-of-covid-19-on-working-professionals/data

# get kaggle data, explore, calculate summary statistics
# get data download from 
data <- read_csv("C:/Users/Suresh/Desktop/MS in DS/Class Project/synthetic_covid_impact_on_work.csv")

# initial exploration
str(data)
summary(data)

# cleaning 
# less than 0 meetings/day isn't possible
data$Meetings_Per_Day[data$Meetings_Per_Day < 0] <- 0
# Affected_by_Covid is same value for everyone. Remove
data$Affected_by_Covid <- NULL

# function to run some descriptive statistics
descr_stat_func <- function(col_data, col_name){ 
  col_mean <- mean(col_data) 
  col_median <- median(col_data) 
  col_range <- range(col_data) 
  stdev <- sd(col_data)
  skew = skewness(col_data)
  kurt = kurtosis(col_data)
  return(cat(col_name, 
             "\nMean: ", col_mean,
             "\nMedian: ", col_median, 
             "\nRange: ", col_range,   
             "\nStandard Deviation: ", stdev,
             "\nSkewness: ", skew,
             "\nKurtosis: ", kurt)) 
} 

descr_stat_func(data$Hours_Worked_Per_Day, "Ave Hours Worked per Day")
descr_stat_func(data$Meetings_Per_Day, "Ave Number of Meetings per Day")

# checking for normality with distribution plot and qqplot for continuous variables
# orange bars
orange <- rgb(230, 140, 50, maxColorValue = 255)

ggplot(data, aes(x=Hours_Worked_Per_Day)) + 
  geom_histogram(binwidth=0.4, color="black", fill=orange) +
  labs(x="Average Hours Worked per Day", y="Count") +
  theme(plot.title=element_text(hjust=0.5), panel.background=element_rect(fill="white"), panel.grid.major=element_line(color="gray")) +
  ggtitle("Working Hours Histogram")

ggplot(data, aes(x=Hours_Worked_Per_Day)) +
  geom_histogram(binwidth=0.4, color="black", fill=orange) +
  labs(x="Average Meetings per Day", y="Count") +
  theme(plot.title=element_text(hjust=0.5), panel.background=element_rect(fill="white"), panel.grid.major=element_line(color="gray")) +
  ggtitle("Meetings Histogram") 

# dark red reference line
dark_red <- rgb(213, 94, 0, maxColorValue = 255)
qqnorm(data$Hours_Worked_Per_Day, cex=0.5, main="QQ Plot of Average Hours Worked per Day")
qqline(data$Hours_Worked_Per_Day, col=dark_red)

qqnorm(data$Meetings_Per_Day, cex=0.5, main="QQ Plot of Average Meetings per Day")
qqline(data$Meetings_Per_Day, col=dark_red)

# exploring hours worked by sector
sector_hours <- data %>%  
  group_by(Sector) %>% 
  summarize(Average_Hours_Worked = mean(Hours_Worked_Per_Day))

ggplot(sector_hours, aes(x = reorder(Sector, -Average_Hours_Worked), y = Average_Hours_Worked)) + 
  geom_bar(stat = "identity", fill = "light blue", color = "black", linewidth=0.5) + 
  labs( title = "Average Work Hours by Sector After COVID", x = "Sector", y = "Average Hours Worked Per Day" ) + 
  coord_cartesian(ylim = c(7.9, 8.1)) +  
  theme_minimal() + 
  theme( axis.text.x = element_text(angle = 45, hjust = 1), plot.title = element_text(hjust = 0.5) ) 

# run anova
anova_result <- aov(Hours_Worked_Per_Day ~ Sector, data=data)
summary(anova_result)


# feature engineering
# binary data for features that could be continuous or have more than two 
# possible observances -> productivity change 1, 0 doesn't give directionality
# will use published psychology research and present features to impute 
# productivity change direction (3 classes -1, 0, 1) and a continuous variable
# salary

# for those whose imputed productivity was decreased/unchanged and salary change = 1
# or salary change = 0 simulate using 2019 data here:

# Retail: https://www.bls.gov/oes/special.requests/oes_research_2019_sec_42-44-45.xlsx
# Healthcare: https://www.bls.gov/oes/special.requests/oes_research_2019_sec_62.xlsx
# Education: https://www.bls.gov/oes/special.requests/oes_research_2019_sec_61.xlsx
# IT: https://www.bls.gov/oes/special.requests/oes_research_2019_sec_51-54.xlsx

# for those whose imputed productivity was increased and salary change = 1
# function to get salary data using 2020 data here:
# Retail: https://www.bls.gov/oes/special.requests/oes_research_2020_sec_42-44-45.xlsx
# Healthcare: https://www.bls.gov/oes/special.requests/oes_research_2020_sec_62.xlsx
# Education: https://www.bls.gov/oes/special.requests/oes_research_2020_sec_61.xlsx
# IT: https://www.bls.gov/oes/special.requests/oes_research_2020_sec_51-54.xlsx

# add feature, convert ave hours/day to hours/week. Assuming closer to 40 hours is more
# productive based on research at stamford.
# assuming 5 working days insert column for hours/week
data$Hours_Worked_Per_Week <- data$Hours_Worked_Per_Day * 5

# function to get bls data
get_bls_salary_data <- function(file_path_and_name){
  df <- read_excel(file_path_and_name)
  return(df)
}

# function to clean BLS data
# between 2019 and 2020 data columns had been altered in some instances from upper to lower case
# Also, BLS data has thousands of jobs throughout multiple sectors. This function cleans all the 
# information out of the data that isn't used in the simulation. Just a general category
# matching what is conventionally know to be each sector was used as described in the report
# the function also removes everything except the 50 states, the 25th percentile, 50th percentile (median),
# and the 75th percentile
clean_salary_data <- function(df, title, sector){
  
  # make all col_names consistent
  colnames(df) <- tolower(colnames(df))
  colnames(df) <- gsub(" ", "_", colnames(df))
  
  # get the naics title  and occupations for the sector
  return_df <- df[df$naics_title == title, ]
  return_df <- return_df[return_df$occ_title == "All Occupations", ]
  
  # remove district of columnbia and territories
  return_df <- return_df[!return_df$area_title %in% c("Guam", "Puerto Rico", "District of Columbia", "Virgin Islands"),]
  return_df <- return_df[return_df$i_group == sector,]
  
  # filter dataframe to only have state, average salary, and IRQ values
  return_df <- return_df[, c("area_title", "a_pct25", "a_median", "a_pct75")]
  
  return_df$a_median[return_df$a_median == "*"] <- NA
  return_df$a_pct25[return_df$a_pct25 == "*"] <- NA
  return_df$a_pct75[return_df$a_pct75 == "*"] <- NA
  
  # coerce mean to numeric
  return_df$a_median <- as.numeric(return_df$a_median)
  return_df$a_pct25 <- as.numeric(return_df$a_pct25)
  return_df$a_pct75 <- as.numeric(return_df$a_pct75)
  
  # impute the NAs
  return_df$a_median[is.na(return_df$a_median)] <- mean(return_df$a_median, na.rm = TRUE)
  return_df$a_pct25[is.na(return_df$a_pct25)] <- mean(return_df$a_pct25, na.rm = TRUE)
  return_df$a_pct75[is.na(return_df$a_pct75)] <- mean(return_df$a_pct75, na.rm = TRUE)
  
  # estimate meanlog, standard deviation, and meansd
  return_df$standard_deviation <- (return_df$a_pct75 - return_df$a_pct25) / 1.35
  return_df$meanlog <- log(return_df$a_median)
  
  # log transform percentiles, subtract for spread
  # the numbers in the denominator come from the z-score for the 25th and 75th percentiles 
  # and then use that z-score multiply by two for interval
  # to calculate sdtev of interval log75%-25% to calculate parameters for estimate
  # of sdlog parameter for 
  return_df$sdlog <- (log(return_df$a_pct75)-log(return_df$a_pct25))/(2 * 0.6745)
  
  # delete quartile columns
  return_df$a_pct25 <- NULL
  return_df$a_pct75 <- NULL
  
  # round to nearest ten and 1/10th
  return_df$a_median <- round(return_df$a_median, -1)
  return_df$standard_deviation <- round(return_df$standard_deviation, -1)
  
  # rename to state
  colnames(return_df)[colnames(return_df) == "area_title"] <- "region"
  
  # convert names in state column to lowercase to matchup with map
  return_df$region <- tolower(return_df$region)
  return(return_df)
}

# function to simulate salaries using lognormal distribution
simulate_salaries <- function(df, state, n) {
  
  meanlog_salary <- df[df$region == state,]$meanlog
  stdevlog_salary <- df[df$region == state,]$sdlog
  
  # generate salaries based on the bls data
  salaries <- rlnorm(n, meanlog = meanlog_salary, sdlog = stdevlog_salary)
  return(salaries)
}

# path for data from 2019
retail_excel_file19 <- "C:/Users/Suresh/Desktop/MS in DS/Class Project/oes_research_2019_sec_42-44-45.xlsx"               
healthcare_file19 <- "C:/Users/Suresh/Desktop/MS in DS/Class Project/oes_research_2019_sec_62.xlsx"
education_file19 <- "C:/Users/Suresh/Desktop/MS in DS/Class Project/oes_research_2019_sec_61.xlsx"
IT_file19 <- "C:/Users/Suresh/Desktop/MS in DS/Class Project/oes_research_2019_sec_51-54.xlsx"
# path for data from 2020
retail_excel_file20 <- "C:/Users/Myers/Desktop/MS in DS/Class Project/oes_research_2020_sec_42-44-45.xlsx"               
healthcare_file20 <- "C:/Users/Myers/Desktop/MS in DS/Class Project/oes_research_2020_sec_62.xlsx"
education_file20 <- "C:/Users/Myers/Desktop/MS in DS/Class Project/oes_research_2020_sec_61.xlsx"
IT_file20 <- "C:/Users/Myers/Desktop/MS in DS/Class Project/oes_research_2020_sec_51-54.xlsx"

# create a new df for each sector by year BLS data
# 2019
retail_data19 <- get_bls_salary_data(retail_excel_file19)
ret_df19 <- clean_salary_data(retail_data19, "General Merchandise Stores", "3-digit")

healthcare_data19 <- get_bls_salary_data(healthcare_file19)
health_df19 <- clean_salary_data(healthcare_data19, "Health Care and Social Assistance", "sector")

education_data19 <- get_bls_salary_data(education_file19)
ed_df19 <- clean_salary_data(education_data19, "Elementary and Secondary Schools", "4-digit")

IT_data19 <- get_bls_salary_data(IT_file19)
IT_df19 <- clean_salary_data(IT_data19, "Information", "sector")

# 2020
retail_data20 <- get_bls_salary_data(retail_excel_file20)
ret_df20 <- clean_salary_data(retail_data20, "General Merchandise Stores", "3-digit")

healthcare_data20 <- get_bls_salary_data(healthcare_file20)
health_df20 <- clean_salary_data(healthcare_data20, "Health Care and Social Assistance", "sector")

education_data20 <- get_bls_salary_data(education_file20)
ed_df20 <- clean_salary_data(education_data20, "Elementary and Secondary Schools", "4-digit")

IT_data20 <- get_bls_salary_data(IT_file20)
IT_df20 <- clean_salary_data(IT_data20, "Information", "sector")

# impute productivity change direction
# Imputed a productivity change direction based on multiple factors.
# study showing that ~40 hours a week was the optimal hours for productivity
# Yerkes-Dodson Law that in a three category ranking system of stress level
# classification, low and high stress has a tendency to be less productive than
# medium. The purpose of the imputation is to simulate salaries.
# increases in productivity will have simulated salary data from 2020 and 
# decreases or no change in productivity resulted in simulating salary data from 2019
data$Productivity_Change_Direction <- NA
data$Productivity_Change_Direction[data$Hours_Worked_Per_Week > 40 & data$Increased_Work_Hours == 1 & data$Productivity_Change == 1] <- -1
data$Productivity_Change_Direction[data$Hours_Worked_Per_Week <= 40 & data$Increased_Work_Hours == 1 & data$Productivity_Change == 1] <- 1
data$Productivity_Change_Direction[data$Productivity_Change == 1 & data$Stress_Level == "Medium" & is.na(data$Productivity_Change_Direction)] <- 1
data$Productivity_Change_Direction[data$Productivity_Change == 1 & data$Stress_Level %in% c("Low", "High") & is.na(data$Productivity_Change_Direction)] <- -1
data$Productivity_Change_Direction[data$Productivity_Change == 0] <- 0

# values for subsetting the simulation by prod change direction and salary changes
retail_of_sal <- data[data$Sector=="Retail",]
num_ret_increased <- sum(data$Sector=="Retail" & data$Productivity_Change_Direction == 1 & data$Salary_Changes == 1)

health_of_sal <- data[data$Sector=="Healthcare",]
num_hel_increased <- sum(data$Sector=="Healthcare" & data$Productivity_Change_Direction == 1 & data$Salary_Changes == 1)

ed_of_sal <- data[data$Sector=="Education",]
num_ed_increased <- sum(data$Sector=="Education" & data$Productivity_Change_Direction == 1 & data$Salary_Changes == 1)

IT_of_sal <- data[data$Sector=="IT",]
num_IT_increased <- sum(data$Sector=="IT" & data$Productivity_Change_Direction == 1 & data$Salary_Changes == 1)

# chose a state to sample from - we chose new york. You could chose any of the 50 and it will sample from the
# BLS lognormal salary distributions
state <- "new york"

data$Salary <- NA
data[data$Sector == "Retail" & data$Productivity_Change_Direction == 1 & data$Salary_Changes == 1,]$Salary <- simulate_salaries(ret_df20, state, num_ret_increased)
num_ret_decreased <- sum(is.na(data$Salary) & data$Sector == "Retail")
data[is.na(data$Salary) & data$Sector == "Retail", ]$Salary <- simulate_salaries(ret_df19, state, num_ret_decreased)

data[data$Sector == "Education" & data$Productivity_Change_Direction == 1 & data$Salary_Changes == 1,]$Salary <- simulate_salaries(ed_df20, state, num_ed_increased)
num_ed_decreased <- sum(is.na(data$Salary) & data$Sector == "Education")
data[is.na(data$Salary) & data$Sector == "Education", ]$Salary <- simulate_salaries(ed_df19, state, num_ed_decreased)

data[data$Sector == "Healthcare" & data$Productivity_Change_Direction == 1 & data$Salary_Changes == 1,]$Salary <- simulate_salaries(health_df20, state, num_hel_increased)
num_hel_decreased <- sum(is.na(data$Salary) & data$Sector == "Healthcare")
data[is.na(data$Salary) & data$Sector == "Healthcare", ]$Salary <- simulate_salaries(health_df19, state, num_hel_decreased)

data[data$Sector == "IT" & data$Productivity_Change_Direction == 1 & data$Salary_Changes == 1,]$Salary <- simulate_salaries(IT_df20, state, num_IT_increased)
num_IT_decreased <- sum(is.na(data$Salary) & data$Sector == "IT")
data[is.na(data$Salary) & data$Sector == "IT", ]$Salary <- simulate_salaries(ret_df19, state, num_IT_decreased)

data$Salary <- round(data$Salary, 2)

# make distribution to check for log normal distribution simulation 
hist(data$Salary, breaks = 75, main = "Salary Histogram", col = "lightblue", xlab="Salary", ylab="Count")


# exploration of binary data for answering business questions

# function to plot bar chart for increased working hours by stress level
# could use this to plot any of the binary variables ---
create_plots <- function(split_df, x, y) {
  
  # bar colors
  orange <- rgb(230, 140, 50, maxColorValue = 255)
  blue <- rgb(50, 70, 160, maxColorValue = 255)
  
  # loop through each sector
  for (sector_name in names(split_df)) {
    # get data from the sector
    df <- split_df[[sector_name]]  
    # get string for the title
    plot_title <- paste("Stress Levels by Increased Working Hours in the", sector_name, "Sector")
    hours_plot <- ggplot(df, aes(x = factor(Stress_Level), fill = factor(Increased_Work_Hours))) +
      # place no and yes bars next to each other on the x axis
      geom_bar(position = "dodge", stat = "count", color="black") +  
      labs(x = "Stress Level", y = "Count", fill = "Increased Working Hours") +
      scale_fill_manual(values = c("0" = blue, "1" = orange),  
                        labels = c("No", "Yes")) + 
      ggtitle(plot_title) +
      theme_minimal() +
      theme(axis.text.x = element_text(hjust = 0.5))
    
    print(hours_plot)
    
  }
}

# function to create bar chart for productivity change direction by stress level
create_three_bar_plots <- function(split_df, x, y) {
  
  # bar colors
  orange <- rgb(230, 140, 50, maxColorValue = 255)  
  blue <- rgb(50, 70, 160, maxColorValue = 255)    
  green <- rgb(50, 160, 50, maxColorValue = 255)
  
  # loop through each sector
  for (sector_name in names(split_df)) {
    # get data from the secotr
    df <- split_df[[sector_name]]  
    # get string for title
    plot_title <- paste("Stress Levels by Productivity Change in the", sector_name, "Sector")
    prod_plot <- ggplot(df, aes(x = factor(Stress_Level), fill = factor(Productivity_Change_Direction))) +
      # place productivity change direction next to each other on the x axis
      geom_bar(position="dodge", stat="count", color="black") +  
      labs(x="Stress Level", y="Count", fill="Change in Productivity") +
      scale_fill_manual(values=c("-1"=green, "0"=blue, "1"=orange),  
                        labels=c("Decrease", "No Change", "Increase")) + 
      ggtitle(plot_title) +
      theme_minimal() +
      theme(axis.text.x = element_text(hjust = 0.5))
    # print the plot
    print(prod_plot)
  }
}

# chagne stress level to factor - label 1,2,3 for plot
data$Stress_Level <- factor(data$Stress_Level, levels = c("Low", "Medium", "High"), labels = c(1, 2, 3))

# group the data by sector 
split_data <- split(data, data$Sector)
create_plots(split_df = split_data, "Stress_Level", "Increased_Work_Hours")
create_three_bar_plots(split_df=split_data, "Stress_Level", "Productivity_Change_Direction")


# svm model to predict stress level
get a new 
data1 <- data
data1 <- data1[,-14]

# change Stress_level to factor
data1$Stress_Level<-as.factor(data1$Stress_Level)

# train test split
trainList<-createDataPartition(y=data1$Stress_Level,p=.6,list=FALSE)
trainSet<-data1[trainList,]
testSet<-data1[-trainList,]

## SVM model
# train the model
svmOutputStress<-train(Stress_Level~.,data=trainSet,method="svmRadial",preProc=c("center","scale"))

# use model to make predictions on test set
predicted<-predict(svmOutputStress,newdata=testSet)

# confusion matrix
confusion<-table(predicted,testSet$Stress_Level)

# random forest
covid <- data

# get percentages of each stress level in the data set to use as class weights
# in the random forest algorithm
stress_level_count <- table(covid$Stress_Level)
stress_level_percentage <- (stress_level_count / sum(stress_level_count))
class_weights <- stress_level_percentage

# set the seed for reproducibility
set.seed(123)

# split the data set
split_covid <- createDataPartition(covid$Stress_Level, p = 0.7, list = FALSE)
train_covid <- covid[split_covid, ]
test_covid <- covid[-split_covid, ]

# get a list of the binary columns
binary_columns <- c("Health_Issue", "Job_Security", "Childcare_Responsibilities", 
                    "Commuting_Changes", "Technology_Adaptation", "Salary_Changes", 
                    "Team_Collaboration_Challenges", "Increased_Work_Hours", "Work_From_Home", "Productivity_Change")

# change the binary columns to factor for learning
for (col in binary_columns) {
  train_covid[[col]] <- factor(train_covid[[col]], levels = c(0, 1))
  test_covid[[col]] <- factor(test_covid[[col]], levels = c(0, 1))
}

# train model for feature importance
rf_model <- randomForest(
  Stress_Level ~ .,
  data = covid_selected_features,
  importance = TRUE,
  ntree = 1000,
  classwt = class_weights
)

importance(rf_model)
varImpPlot(rf_model)

# select features for learning
selected_features <- c("Stress_Level","Hours_Worked_Per_Day", "Meetings_Per_Day", "Technology_Adaptation", 
                       "Increased_Work_Hours", "Job_Security", "Sector")

covid_selected_features <- train_covid[, selected_features]
class_weights <- stress_level_percentage

# train model on those features
rf_model <- randomForest(
  Stress_Level ~ .,
  data = covid_selected_features,
  importance = TRUE,
  ntree = 1000,
  classwt = class_weights
)

# make predictions 
rf_predictions <- predict(rf_model, newdata = covid_selected_features)

conf_matrix <- table(Predicted = rf_predictions, Actual = covid_selected_features$Stress_Level)
print(conf_matrix)

confusion_matrix <- confusionMatrix(rf_predictions, covid_selected_features$Stress_Level)
print(confusion_matrix)
