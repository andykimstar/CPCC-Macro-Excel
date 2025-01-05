## Description
In collaboration with  [Canadian Phycological Culture Centre](https://uwaterloo.ca/canadian-phycological-culture-centre/about) (CPCC) Biology Lab at University of Wateroo, a Macro Excel file was created to automate the data processing and collection to evaluate the performance of the CPCC Biology Lab.
The project was initiated & led from the product from design, development, testing and management.

Contact: andykimstar@gmail.com

## Purpose
The Macro-Excel file carries the goal of eliminating the redunant manual labour, while saving time though the automation process. 
Hence the project automates the data processing and collection to evaluate the performance of the CPCC Biology Lab in University of Waterloo. 
It intends to organize the data, construct variables, and run statistical analysis.

This CPCC Excel computes & monitor the Orders into a CPCC Operation breakdown. Once the CPCC Orders has been entered correctly, the file is capable of tracking. The file & code here [github](https://github.com/andykimstar/CPCC-Macro-Excel)

I have recorded the overview of the content which can be find here: [Go to Demo](https://youtu.be/qNkoCGgKEuw).


## Design
![CPCC-Design](https://github.com/user-attachments/assets/2d1319d7-8fd6-4aba-9712-16ec8978b1a2)



## List of Content

### Watch out
* Allow Trust notification - Click
* Colors of the tab for the separation between automation/non-automation & merge/non-merge  (Green w/ Black Text = Automation & non-merge   +   Dark Green w/ White Text = Automation & Merge   +    White = Non-Automation)
* '"-" for column F if the DATE FROM is after December 2022 in the case of Zeros
* '"-" for column E if the DATE FROM in before December 2022 in the case of Zeros

### Orders
* Entry of each request of Order

* *Notes*
    - Button in Top Left Corner will scroll to bottom of Table
    - Watch out for merging rows for columns in Green Tab
    - Please read the notes attached at the top of each column

### Usage
* Tracks 12-Month periods usage/deatils of the Requests
  - data coming-from: **_Orders Tab_**

* *Notes*
    - Must be 12-month period date range
    - Only the 12-Month period Table is automated
    - Only the ( Yearly Table + Chart ) must be edited manually 

![Usage](https://github.com/user-attachments/assets/fb2e8faa-39e5-490d-8f63-d66957a65a0d)



### Media Requests
* Tracks 12-Month periods in ( Number & Volumes ) of each Media Requests
  - data coming-from: **_Orders Tab_**

* *Notes*
    - Must be 12-month period date range
    - All tables are fully automated
    - Two Tables must have same list of Media
      
![MediaRequest](https://github.com/user-attachments/assets/72d1240c-ad1b-4b26-b0fe-b4b3c8dd2dad)


 
### Source Requests
* Tracks 12-Month periods in ( Country & Affiliation ) of each Requests
  - data coming-from: **_Orders Tab_**

* *Notes*
    - Must be 12-month period date range
    - All tables and pie chart are fully automated
    - More countries can be added/removed in the table (it's not a fixed list of countries)
 
![SourceRequest](https://github.com/user-attachments/assets/901f1124-89c7-4d71-9ca2-6281006fa9f1)


    
### Strains Ordered
* Tracks any-entered periods in ( Total Count & Latest Order Date ) of each Strains Requests
  - data coming-from: **_Orders Tab_**

* *Notes*
    - Enter any period date range
    - All tables are fully automated
    - Button in Top Left Corner will scroll to bottom of Table
    - More strains can be added/removed in the table (it's not a fixed list of strains)
    - If the Date Range include Dec/31/22 then Total Count of Requests from 1986 to 2022 will be included
    - Select the Column then CTRL+Shift+L to Filter the Columns

![StrainsOrdered](https://github.com/user-attachments/assets/221deb48-9279-4d46-94e4-0fc046a202ea)



### Service Revenue Breakdown
* Tracks 12-Month periods in Total Revenue of each CPCC Service Requests
  - data coming-from: **_Orders Tab_**

* *Notes*
    - Must be 12-month period date range
    - All tables and pie charts are fully automated
    - Revenue is broken down into Ordered & Invoiced
    - More services can be added/removed in the table (it's not a fixed list of service)

![ServiceRevenue](https://github.com/user-attachments/assets/f0c53022-d301-42b6-8ba7-7ff427f8089c)



### Users List
* Tracks any-entered periods oragnizes all Primary User of each Requests: (Primary User Details, Count of Request from Primary User, All associated Additional User for each Primary User)
  - data coming-from: **_Orders Tab_**
    
* *Notes*
    - Enter any period date range
    - All tables are fully automated
    - All additional users associated to each primary user are listed
    - Select the Column then CTRL+Shift+L to Filter the Columns
 
![UserList](https://github.com/user-attachments/assets/def245ea-4db0-4fb8-ab21-681343314463)


 
### Instituion List
* Tracks any-entered periods organizes all Institution of each Requests: (Primary User Details, Count of Request from Institution, All associated Primary/Additional User for each Institution )
  - data coming-from: **_User List_**
    
* *Notes*
    - Enter any period date range
    - All tables are fully automated
    - Table is divided by Affiliation
    - All users associated to each institution are listed
    - Select the Zone (col/row) then CTRL+Shift+L to Filter the Columns
 
![InstitutionList](https://github.com/user-attachments/assets/0b055fc3-a10f-488d-91a4-4a231abb63f1)



## How to..

### How to save a copy of a file

1. Make a copy of the Excel file and re-name
2. Open the copied Excel File
3. Go to "FILE" Tab => "INFO" => "Protect Workbook" and select "Always Open Read-Only"
4. Save and CLOSE the Copied File
5. Re-Open

![MacroSetting](https://github.com/user-attachments/assets/4c2841e8-c8d4-4ebb-a523-5f2256afbff5)


### Override Excel Security to unblock and allow trusted file

1. Click on the Properties of the Excel File
2. Scroll Down to "Unblock"

![Security](https://github.com/user-attachments/assets/272cb18c-bbd9-4be1-9959-ee3676bc2b7f)


