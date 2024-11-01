# to3-inventory-analysis

## Instructions
Detailed usage instructions can be found in the User Guide

## Background
The Inventory Analysis Tool is designed to examine supply chain patterns over time and flag priority facilties and products for action. Two metrics, inventory turnover and the Coefficient of Variance for the consumption are calculated and used together to categories facilities/products into risk categories.

### Method
Both metrics are calculated on a rolling basis over a given period of months (typically 12 months) for a given facility/product combination

#### Inventory Turn Rate
Inventory Turn Rate = sum(consumption) / avg(stock on hand)
Unit - times stock turned over per period (typically 12 months)

#### COV Consumption
CoV Consumption = std(consumption) / avg(consumption)
Interpretation - the lower the COV the less variability there is. < .7 is low, > 1.5 is high

### Tool overview

The tool has been designed to work on any country's data that meets certain data requirements. The user inputs dataset-specific information which map to required arguments for the tool

#### Data requirements
Excel or csv data containing product, facility, consumption, stock on hand, date 

####
Input - excel or csv
User interface - python gui using pysimplegui
Analysis - python script
Ouput - txt tables
Visualization - excel dashboard, csv output connected via Power Query

## To run
pip install requirements.txt

python inventory_analysis.gui
