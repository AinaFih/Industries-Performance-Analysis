##If running this code, please don't forget to change the working directory in the code to your own.

library(tidyverse)
library(readxl)
library(stringr)
library(tools) 
library(lme4)
library(Hmisc) ### If applying changes in the following code, note that this library overrides some functions in the tidyverse library. 
### Therefore, it is recommended to call the functions explicitly within the library, e.g., dplyr::summarize.


### This code will not be executed because it is unnecessary for the data files provided with this analysis.
### However, if you are interested in running this code, you must first extract it from the function, and then run it independently 
### using the original files available on the specified website.

combination_code = function(){
  ## These lines of code were used to combine the data files from the website. 
  ## Initially, the files were exported to CSV format and then regionally appended together.
  
  ## Note: The CSV files for the US region were renamed in the file explorer to include "US" in their filenames.
  ## This was necessary because the original file names from the website did not include the region name, 
  ## which is required for subsequent processing.
  
  clean_and_convert_to_csv = function(input_folder, output_folder) {
    excel_files = list.files(input_folder, pattern = "\\.xls$", full.names = TRUE)
    
    for (file in excel_files) {
      file_name = file_path_sans_ext(basename(file))
      
      # Reading the Excel file, skipping the first 8 rows as they're not part of the data but just information
      data = read_excel(file, skip = 8)
      
      output_file = paste0(output_folder, file_name, ".csv")
      write.csv(data, output_file, row.names = FALSE)
    }
  }
  
  #Path to the directory folders, don't forget the "/" at the end
  input_folder = "C:/Path/"
  output_folder = "C:/Path/"
  
  clean_and_convert_to_csv(input_folder, output_folder)
  
  ##appending the file together
  append_files = function(input_folder1, output_file1) {
    csv_files = list.files(input_folder1, pattern = "\\.csv$", full.names = TRUE)
    data_list = list()
    
    for (file in csv_files) {
      file_name = basename(file)
      
      #Extract the two-digit year from the file name and convert it to the full year
      two_digit_year = as.numeric(str_extract(file_name, "\\d{2}(?=\\.csv$)"))
      year = 2000 + two_digit_year
      
      data = read_csv(file, col_types = cols(.default = "c")) # Reading all columns as characters, as there were issues when this was not done before
      data$year = year
      
      region = str_extract(file_name, "(?<=margin)[a-zA-Z]+(?=\\d+)")
      data$region = region
      
      # Append the data to the list
      data_list[[length(data_list) + 1]] = data
    }
    
    #Creating a unified "Industry" column by coalescing the columns that have the same information in it but with different names
    #The misspelled Induistry Name only exists in some of the region files
    
    ##coalescing columns
    final_data = bind_rows(data_list)
    final_data = if ("Induistry Name" %in% colnames(final_data)){
      final_data %>%
        mutate(industry = coalesce(Industry,
                                   `Industry Name`,
                                   `Induistry Name`))%>%
        select(!Industry & !`Industry Name` & !`Induistry Name`)%>%
        # Convert remaining columns to appropriate types
        mutate(across(everything(), type.convert, as.is = TRUE))
    } else{
      final_data%>%
        mutate(industry = coalesce(Industry,
                                   `Industry Name`))%>%
        select(!Industry & !`Industry Name`)%>%
        # Convert remaining columns to appropriate types
        mutate(across(everything(), type.convert, as.is = TRUE))
    }
    
    # Write the combined data to a CSV file
    write.csv(final_data, output_file1, row.names = FALSE)
  }
  
  #Path to the directory folders. Don't forget to include the file where the results will be exported to at the end
  ##Each region had a different path (e.g., /emerg for emerging countries, /Europe for Europe, etc.), so the code was run for each of them.
  input_folder1 = "C:/Path/"
  output_file1 = "C:/Path/combined_data.csv"
  
  append_files(input_folder1, output_file1)
}

### The files were previously combined regionally and transformed using the code described above. They were also renamed in the file folder to include the region name in the file names.  
### Subsequently, the files were grouped into a single folder to streamline the workflow. As a result, the files loaded from this point are not the original documents but the transformed datasets.  
### The original datasets are available online via the link provided in the report file. Running them through the code described above will produce the transformed files being loaded here.  


### Importing and Cleaning the Datasets
##Don't forget to change this path to your directory folders
setwd("C:/Path/")

##Importing the datasets
##The datasets from the above code have been moved to one folder and renamed according to the region
data_files = list.files(pattern = "*.csv$", full.names = TRUE)
for (file in data_files) {
  file_name = str_remove(basename(file), "\\.csv$")
  assign(file_name, read_csv(file))
}


##One column is totally empty in the US dataset, due to some typo mistake in the original file that haven't been removed during the combination process by region
combined_data_US = combined_data_US %>%
  select(!`Industry  Name`)

##Checking if there is any value equal to 0 in each region dataset
sum(combined_data_global ==0)
sum(combined_data_emerg ==0)
sum(combined_data_japan ==0)
sum(combined_data_US ==0)
sum(combined_data_europe ==0)

##Combining the datasets in one dataset
margin = bind_rows(combined_data_emerg, combined_data_europe, combined_data_global, combined_data_japan, combined_data_US)

##checking the types of each column to see if they're correct
sapply(margin, class)

## Renaming the columns
margin = margin %>%
  rename(
    number_of_firms = 'Number of firms',
    net_margin = 'Net Margin',
    UOM_pretax = 'Pre-tax Unadjusted Operating Margin',
    UOM_aftertax = 'After-tax Unadjusted Operating Margin',
    Lease_AM_pretax = 'Pre-tax Lease adjusted Margin',
    Lease_AM_aftertax = 'After-tax Lease Adjusted Margin',
    Lease_RD_AM_pretax = 'Pre-tax Lease & R&D adj Margin',
    Lease_RD_AM_aftertax = 'After-tax Lease & R&D adj Margin',
    EBITDA_Sales = 'EBITDA/Sales',
    EBITDA_SGA_Sales = 'EBITDASG&A/Sales',
    EBITDA_RD_Sales = 'EBITDAR&D/Sales',
    gross_margin = 'Gross Margin',
    Pretax_Prestock_OM = 'Pre-tax, Pre-stock compensation Operating Margin',
    COGS_Sales = 'COGS/Sales',
    RD_Sales = 'R&D/Sales',
    SGA_Sales = 'SG&A/ Sales' ,
    Stock_based_compensation_Sales = 'Stock-Based Compensation/Sales',
    Lease_expense_Sales = 'Lease Expense/Sales'
  )

## Dropping the rows with missing values
margin = margin %>%
    na.omit()



###Descriptive Analysis

##Global industry profitability

#checking the name of all the regions
unique(margin$region)

##Graphic function of the average profitability (top ten only), the parameter is the region and profitability ratio we want to see
profit_graph = function(Region, profit_KPI){
  average_KPI = margin %>%
    filter(region == Region) %>%
    group_by(industry) %>%
    dplyr::summarize(
      avg = mean(!!sym(profit_KPI))
    )
  avg_KPI = average_KPI %>%
    arrange(desc(avg))
  
##Transforming the type of avg_KPI$Industry into factor to keep its order untouched for the plots
  avg_KPI$industry = factor(avg_KPI$industry, levels = avg_KPI$industry)

  avg_KPI[1:10,] %>%
    ggplot(aes(x = industry, y = avg, fill = avg)) + 
    geom_bar(stat = "identity") +
    scale_fill_gradient(low = "lightgreen", high = "darkgreen") +
    labs(title = paste("Top 10 Most Profitable Industries by", toTitleCase(profit_KPI), ",", toTitleCase(Region)),
      x = "Industries",
      y = "Profit")+
    theme_minimal() + 
    theme(panel.grid = element_blank(),
          axis.text.x = element_text(angle = 45, hjust = 1))
}

profit_graph("Global", "net_margin")

##Getting the top 10 performing industries in each region
top_performing_industries= margin %>%
  group_by(region, industry) %>%
  dplyr::summarize(Average_net_margin = mean(net_margin)) %>%
  slice_max(order_by = Average_net_margin, n = 10) %>% 
  select(region, industry, Average_net_margin)

##Putting the top 10 performing industries by region
top_performing_industries_1= top_performing_industries %>%
  select(region, industry) %>%
  mutate(rank = row_number()) %>%
  pivot_wider(
    names_from = region,
    values_from = industry,
    names_sort = TRUE
  )
write_csv(top_performing_industries_1, "C:/Users/manit/Documents/M1 BDEEM/Data analysis with R/Project/xls_files/Margin/Regional Top performing industries.csv")


##Counting the most recurrent industries in the top ten of each region
industry_frequency = top_performing_industries %>%
  filter(region != "Global") %>%
  ungroup() %>%
  group_by(industry) %>%
  count(industry) %>%
  arrange(desc(n))

industry_frequency$industry = factor(industry_frequency$industry, levels = industry_frequency$industry)
industry_frequency %>%
  ggplot(aes(x = industry, y = n, fill = n)) + 
  geom_bar(stat = "identity") +
  scale_fill_gradient(low = "lightblue", high = "darkblue" ) +
  labs(title = paste("Most Frequent Top-Performing Industries Across Regions"),
       x = "Industries",
       y = "Frequency")+
  theme_minimal() + 
  theme(panel.grid = element_blank(),
        axis.text.x = element_text(angle = 75, hjust = 1))
print(industry_frequency, n=24)

###Global industry risk
risk_graph = function(Region, profit_KPI){
  average_risk_KPI = margin %>%
    filter(region == Region) %>%
    group_by(industry) %>%
    dplyr::summarize(
      var_risk = var(!!sym(profit_KPI))
    )
  avg_risk_KPI = average_risk_KPI %>%
    arrange(desc(var_risk))
  
  ##Keeping the industry order untouched for the plots
  avg_risk_KPI$industry = factor(avg_risk_KPI$industry, levels = avg_risk_KPI$industry)
  avg_risk_KPI[1:10,] %>%
    ggplot(aes(x = industry, y = var_risk, fill = var_risk)) + 
    geom_bar(stat = "identity") +
    scale_fill_gradient(low = "salmon", high = "firebrick" ) +
    labs(title = paste("Top 10 Riskiest Industries by", toTitleCase(profit_KPI), "Variability,", toTitleCase(Region)),
         x = "Industries",
         y = "Variance")+
    theme_minimal() + 
    theme(panel.grid = element_blank(),
          axis.text.x = element_text(angle = 35, hjust = 1))
}
risk_graph("Global", "Lease_expense_Sales")

## Variance for all the regions
box_plot_graph = function(KPI, Industry){
  margin %>%
    filter(industry == Industry)%>%
    ggplot(aes(x = region, y = !!sym(KPI), color = region)) +
    geom_boxplot() +
    labs(title = paste(KPI, "variability by Region for the", Industry, "Industry"),
         x = "",
         y = KPI) + 
    theme_minimal()
}

box_plot_graph("Lease_expense_Sales", "Telecom (Wireless)")

###Market Size analysis
growth_function = function(Region, KPI){
  average_growth = margin %>%
    filter(region == Region) %>%
    arrange(industry, year) %>%
    group_by(industry) %>%
    mutate(growth = (!!sym(KPI) - Lag(!!sym(KPI)))/ Lag(!!sym(KPI))) %>%
    dplyr::summarize(avg = mean(growth, na.rm = TRUE))
  avg_growth = average_growth %>%
    arrange(desc(avg))
  
  ##Transforming the type of avg_KPI$Industry into factor to keep its order untouched for the plots
  avg_growth$industry = factor(avg_growth$industry, levels = avg_growth$industry)
  
  graph_title = if(KPI == "number_of_firms"){
    paste("Top 10 Most Growing Industries,", toTitleCase(Region))
  } else if(KPI == "RD_Sales"){
    paste("Top 10 Most Innovative Industries,", toTitleCase(Region))
  } else{
    paste("Top 10 Most Growing Industries by", KPI, ",", toTitleCase(Region))
  }
  avg_growth[1:10,] %>%
    ggplot(aes(x = industry, y = avg, fill = avg)) + 
    geom_bar(stat = "identity") +
    coord_flip() +
    scale_fill_gradient(low = "lightgreen", high = "darkgreen") +
    labs(title = graph_title,
         x = "Industries",
         y = "Average Growth",
         fill = "")+
    theme_minimal() + 
    theme(panel.grid = element_blank(),
          axis.text.x = element_text(angle = 45, hjust = 1))
}

growth_function("Europe", "RD_Sales")
growth_function("Global", "number_of_firms")

##Largest industries
largest_indus = margin %>%
  filter(year == 2023 &  !str_detect(industry, "Total")) %>%
  group_by(region)%>%
  slice_max(order_by = number_of_firms, n = 3) %>% 
  select(region, industry, number_of_firms)
largest_indus
largest_indus %>%
  ggplot(aes(x = region, y = number_of_firms, fill = industry)) +
  geom_bar(stat = "identity") +
  labs(title = "Largest industries per region",
            x = "Region",
            y = "Industry size")+
  theme_minimal() + 
  theme(panel.grid = element_blank(),
        legend.position = "bottom")

###Econometric modelling
##Putting number_of_firms into ln form as the scale is too big compared to the others
margin1 = margin %>%
  filter(region != "Global") %>%
  mutate(ln_nof = log(number_of_firms))

mixed_model = lmer(net_margin ~
                     Lease_expense_Sales +
                     RD_Sales + ln_nof  +
                     (1 | industry) +
                     (1 | region),
                   data = margin1
                   )
summary(mixed_model)

model_results = summary(mixed_model)

## A 2nd model was made to compare and get the best one
mixed_model2 =lmer(net_margin ~ Lease_expense_Sales + RD_Sales + ln_nof +
                         (Lease_expense_Sales | industry)+
                         (1 | region), data = margin1)
summary(mixed_model2)

##Normality and Heteroskedasticity tests:
qqnorm(residuals(mixed_model))
qqline(residuals(mixed_model))

residuals = residuals(mixed_model)
fitted_values = fitted(mixed_model)
ggplot(data.frame(fitted = fitted_values, residuals = residuals), aes(x = fitted, y = sqrt(abs(residuals)))) +
  geom_point() +
  geom_smooth(se = FALSE, color = "red") +
  labs(title = "Scale-Location Plot", x = "Fitted Values", y = "Square Root of |Residuals|") +
  theme_minimal()
qqnorm(residuals(mixed_model2))
qqline(residuals(mixed_model2))

residuals = residuals(mixed_model2)
fitted_values = fitted(mixed_model2)
ggplot(data.frame(fitted = fitted_values, residuals = residuals), aes(x = fitted, y = sqrt(abs(residuals)))) +
  geom_point() +
  geom_smooth(se = FALSE, color = "red") +
  labs(title = "Scale-Location Plot", x = "Fitted Values", y = "Square Root of |Residuals|") +
  theme_minimal()

AIC(mixed_model, mixed_model2)