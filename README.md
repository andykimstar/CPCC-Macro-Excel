# CPCC-Macro-Excel-File

## Description
The development of this product is completely led by myself from design, development, testing and management (CI/CD) as part of the Part-Time work at the Canadian Phycological Culture Centre (CPCC) Biology Lab at University of Wateroo. 
It was initiated by myself with suggestion and presention to the Professor in effort to help the tracking of the Sales & Customers of the Lab. 

Canadian Phycological Culture Centre (CPCC) Link: [https://uwaterloo.ca/canadian-phycological-culture-centre/](https://uwaterloo.ca/canadian-phycological-culture-centre/about)

The Macro Excel (.xlsx) is a variation Excel File with a layer of coding embedded in the file. 
The goal of this Macro-Excel file is to eliminate redunant manual labour, while saving time though the automation process. 

## Purpose
This code is to automate the data processing and collection to evaluate the performance of the CPCC Biology Lab in University of Waterloo. 
It eliminates the manual work of collection as the code intends to clean data, construct variables, and run statistical analysis

This CPCC Excel computes & monitor the Orders into a CPCC Operation breakdown. Once the CPCC Orders has been entered correctly, the file is capable of tracking. 

## Design
![image](https://github.com/andykimstar/CPCC-Excel-Automation-Tool/assets/113536228/d2acc607-bdfe-4422-ac5c-c38f5962a732)



## List of Content

### Orders
* Entry of each request of Order

* Notes
    - Button in Top Left Corner will scroll to bottom of Table
    - Merge Requests with several Parts. BUT No Rows can be merged in-between Column A-I (avoid dup)

### Usage
* Tracks 12-Month periods usage/deatils of the Requests
  - Data From: Orders

* Notes
    - Must be 12-month period date range
    - Only the 12-Month period Table is automated
    - Only the ( Yearly Table + Chart ) must be edited manually 

![Usage](https://github.com/user-attachments/assets/3842d6ae-e6d2-4f8b-a7ca-8112317afcef)


### Media Requests
* Tracks 12-Month periods in ( Number & Volumes ) of each Media Requests
  - Data From: Orders

* Notes
    - Must be 12-month period date range
    - All tables are fully automated
    - Two Tables must have same list of Media
      
![MediaRequest](https://github.com/user-attachments/assets/0bb4c1b9-7984-4f71-8299-89e8cd14586c)

 
### Source Requests
* Tracks 12-Month periods in ( Country & Affiliation ) of each Requests
  - Data From: Orders

* Notes
    - Must be 12-month period date range
    - All tables and pie chart are fully automated
    - More countries can be added/removed in the table (it's not a fixed list of countries)
 
![SourceRequest](https://github.com/user-attachments/assets/8ad132dd-9b1e-4b5a-a52d-00e482cb7451)

    
### Strains Ordered
* Tracks any-entered periods in ( Total Count & Latest Order Date ) of each Strains Requests
  - Data From: Orders

* Notes
    - Enter any period date range
    - All tables are fully automated
    - Button in Top Left Corner will scroll to bottom of Table
    - More strains can be added/removed in the table (it's not a fixed list of strains)
    - If the Date Range include Dec/31/22 then Total Count of Requests from 1986 to 2022 will be included
 
    
### Service Revenue Breakdown
* Tracks 12-Month periods in Total Revenue of each CPCC Service Requests
  - Data From: Orders

* Notes
    - Must be 12-month period date range
    - All tables and pie charts are fully automated
    - Revenue is broken down into Ordered & Invoiced
    - More services can be added/removed in the table (it's not a fixed list of service)
 
### Users List
* Tracks any-entered periods oragnizes all Primary User of each Requests: (Primary User Details, Count of Request from Primary User, All associated Additional User for each Primary User)
  - Data From: Orders
    
* Notes
    - Enter any period date range
    - All tables are fully automated
    - All additional users associated to each primary user are listed
 
### Instituion List
* Tracks any-entered periods organizes all Institution of each Requests: (Primary User Details, Count of Request from Institution, All associated Primary/Additional User for each Institution )
  - Data From: User List
    
* Notes
    - Enter any period date range
    - All tables are fully automated
    - Table is divided by Affiliation
    - All users associated to each institution are listed

