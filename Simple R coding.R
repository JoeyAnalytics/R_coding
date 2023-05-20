library(tidyverse)
library(readxl)
library(ggplot2)
library(lubridate)
library(stringr)
library(ggpubr)


### Loading Excel Files
# Directory
# file_path <- "/Users/mitchleeson/Documents/Education/Masters/Smith/860 Acuisition and Management of Data/Assignments/Assignment 1/MMA860_Assignment1_Data_vf.xlsx"
file_path <- "/Users/yufanwang/Desktop/MMA 860/Data/MMA860_Assignment1_Data_vf.xlsx"
product_sales <- read_excel(file_path, "Product_Sales" )
# Question 1
p_1 <- read_excel(file_path, sheet = "Island Airport Weather", skip=16, .name_repair = "universal")
p_1$index <- 1:nrow(p_1)

# Question 3
product_sales <- read_excel(file_path, sheet = "Product_Sales" )

# Question 4
collinear_25 <- read_excel(file_path, sheet = "Collinearity", n_max=25)
collinear_100 <- read_excel(file_path, sheet = "Collinearity")





### Question 1
# Add in constants
TurbineSweptArea <- 13273.23
MaxPowerCoefficient <- 0.35
TurbineNominalPower <- 4 * 1000000 #(in watts, converted from given Megawatts)
CutInWindspeed <- 4
CutOutWindspeed <- 32


# Summary statistics, testing
# summary(p_1)
# head(p_1,5)
# str(p_1)
# view(p_1)

# Question 1. a) Air Density
p_1$AirDensity <- (p_1$Stn.Press..kPa. * 1000) / (287.05 * (p_1$Temp...C. + 273.15))
head(p_1 %>% select(AirDensity),5)


# Question 1. b) Wind speed in meters per second
p_1$WindSpdMetersPerSecond <- p_1$Wind.Spd..km.h. * 0.277778
head(p_1$WindSpdMetersPerSecond, 5)
head(p_1 %>% select(WindSpdMetersPerSecond), n=5)
#testing that values with 0 wind speed are still showing
#view(filter(p_1,p_1$Wind.Spd..km.h. == 0))


# Question 1. c) Power production at each turbine
PowerWatts <- p_1$AirDensity * TurbineSweptArea * ((0.5 * (p_1$WindSpdMetersPerSecond)^3)) * MaxPowerCoefficient

p_1$PowerWatts <- ifelse(p_1$WindSpdMetersPerSecond > CutOutWindspeed,0,
                         ifelse(p_1$WindSpdMetersPerSecond < CutInWindspeed,0,
                                ifelse(PowerWatts > TurbineNominalPower,TurbineNominalPower,
                                       PowerWatts)))
head(p_1 %>% select(PowerWatts), n=5)
# Summary stats
# view(p_1) #opens new window
# p_1[71,]

# Extra add column for power but with no limits
p_1$PowerWattsNoLimit <- ifelse(p_1$WindSpdMetersPerSecond > CutOutWindspeed,0,
                                ifelse(p_1$WindSpdMetersPerSecond < CutInWindspeed,0,
                                       PowerWatts))


# Question 1. D) What is the total amount of electricity produced for the entire windfarm in January in Megawatts? 
aggregate(p_1$PowerWatts/1000000, by=list(p_1=p_1$Month), FUN=sum)


# Question 1. E) visualization showing power produced per day at the windfarm in January. 
PowerPerDay <- aggregate(p_1$PowerWatts/1000000, 
                         by=list(Day=p_1$Day), 
                         FUN=sum)

ggplot(PowerPerDay, aes(y=x, x=Day)) + 
  geom_line() +
  geom_smooth(method = "lm", se=F, color="Steel Blue") +
  labs(title = "Power Produced per Day in January",
       y = "Total Power in Megawatts",
       x = "Day of Month (in January)")





### Question 2 - Making visualizations to tell a story
# Creating table to get hypothetical power production if max turbine power were to double
ExtraProduction <- p_1 %>% 
  select(index, Day, PowerWattsNoLimit) %>% 
  arrange(desc(PowerWattsNoLimit)) %>%
  filter(PowerWattsNoLimit >= 4000000 & PowerWattsNoLimit <=8000000)
ExtraProduction
print(sum(ExtraProduction$PowerWattsNoLimit))

# Plot 1 - "Potential Power Production by Weather Type in January"
# Convert Time column to numeric
p_1$UnformattedTime <- as.numeric(format(p_1$Time, "%H"))
p_1$DayFormatted <- paste("2022", p_1$Month, p_1$Day, p_1$UnformattedTime, sep="-") %>% ymd_h() %>% as_datetime()

ggplot(p_1, aes(y=PowerWattsNoLimit/1000000, x=DayFormatted, colour = Weather, fill = Weather)) + 
  geom_col() +
  theme(axis.text.x = element_text(face = "bold", color = "black"),
        axis.title.x = element_blank()) +
  labs(title = "Potential Power Production by Weather Type in January",
       y = "Power Production")

# Plot 2 - "power by power"
p_1PlottedForPowerByPower <- p_1
p_1PlottedForPowerByPower <- p_1PlottedForPowerByPower[order(p_1PlottedForPowerByPower$PowerWattsNoLimit), ]
p_1PlottedForPowerByPower$index <- factor(p_1PlottedForPowerByPower$index, levels = p_1PlottedForPowerByPower$index)

ggplot(p_1PlottedForPowerByPower, aes(y=PowerWattsNoLimit/1000000, x=index, colour = Weather, fill = Weather)) + 
  geom_col() +
  theme(axis.text.x = element_blank(),
        axis.title.x = element_blank(),
        legend.position = c(0.25, 0.65)) +
  labs(title = "Power Distribution Curve",
       y = "Power Production")

#TESTING OTHER PLOTS - Plot 3
ggplot(p_1, aes(y=`Rel.Hum....`, x=`Temp...C.`, color=Weather, size=PowerWatts)) +
  geom_point(aes(fill=Weather)) +
  geom_point(shape=1, colour="black") +
  labs(title = "PowerProduction is a Function of Temp and Humidity",
       y = "Relative Humidity",
       x = "Temperature")





### Question 3 - Tidying datasets and basic correlations
# Question 3. a) Convert Price from string to number 

product_sales <- read_excel(file_path, "Product_Sales" )
product_sales$Price <- as.numeric(gsub("\\$","",product_sales$Price))

str(product_sales)

# Add leading zeros
product_sales$Product_ID <- str_pad(product_sales$Product_ID, 3, pad  ="0")
head(product_sales,10)
str(product_sales)


# Question 3. b) Tidy the dataset
#separate 2016 and 2017 data to different observations
gather_product_sales <- gather(product_sales, Year, Sales,Sales_2016, Sales_2017)
head(gather_product_sales,10)
gather_product_sales


# Question 3. c) Price vs Sales graph
# price_graph <- ggplot(gather_product_sales, aes(y = Price, x = Sales, shape = Year)) +
#   geom_point(aes( size = 2, color = as.factor(Year))) +
#   scale_shape_manual(values = c(18, 20)) +
#   ggtitle("Sales vs. price in year 2016 and 2017")
# price_graph

price_graph <- ggplot(gather_product_sales, aes(x=Sales, y=Price, color=Year, shape=Year)) +
  geom_point( size = 3) + 
  geom_smooth(formula = y ~ x, method=lm, se=FALSE, fullrange=TRUE) +
  labs(title = "Sales vs. price in year 2016 and 2017")
price_graph




# Question 3. d) Price correlations
#Price correlation calculation
price_retailers_correlation <- cor(gather_product_sales$Price, gather_product_sales$Num_Retailers)
price_retailers_correlation

#Price and Number of Retailers scatter plot
price_retailers_graph <- ggplot(gather_product_sales, aes( y = Price, x = Num_Retailers)) + geom_point(aes()) +
  geom_smooth(method="lm") + #best fit line
  scale_shape_identity() +
  labs(title = "Price vs Number of Retailers", 
       y = "Price", 
       x = "Number of Retailers")
price_retailers_graph


# Question 3. e) tell a story with the data
# Import vs. Non
Import_group_by <- gather_product_sales %>% group_by(Import, Year) %>%
  summarise(total_sales = sum(Sales))

Import_group_by$Import <- as.factor(Import_group_by$Import)
Import_group_by

Import_group_by_plot <- ggplot(Import_group_by) +
  geom_bar(aes(x = Import, y = total_sales), stat ="identity", width = 0.5, fill = "blue")+
  geom_text(aes(x = Import, y= total_sales, label = total_sales),position = position_dodge(width = 0.9), vjust = 0)+
  geom_line(aes(x = Import, y= total_sales, group = 1), color = "black", stat = "identity")+
  facet_grid(. ~ Year) + 
  labs(title = "Total Sales of Imported vs. Non-Imported Products", y = "Sales $", x = "Import")
Import_group_by_plot

reg_height_25 <- lm(Y ~ Experience + Height, collinear_25)
summary(reg_height_25)

reg_height_100 <- lm(Y ~ Experience + Height, collinear_100)
summary(reg_height_100)

# ii) Regression testing on Y and Experience + Weight
reg_weight_25 <- lm(Y ~ Experience + Weight, collinear_25)
summary(reg_weight_25)

reg_weight_100 <- lm(Y ~ Experience + Weight, collinear_100)
summary(reg_weight_100)

# iii) Regression testing on Y and Experience + Height + Weight
reg_height_weight_25 <- lm(Y ~ Experience + Height + Weight, collinear_25)
summary(reg_height_weight_25)

reg_height_weight_100 <- lm(Y ~ Experience + Height + Weight, collinear_100)
summary(reg_height_weight_100)

