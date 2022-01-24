# Railway decision support system
* Written by: [Poon Athit S. ](https://www.linkedin.com/in/athit-srimachand/)
* Technologies: VBA, User forms, Excel-Macros

## 1. Introduction
Since the COVID-19 pandemic-related travel restrictions were lifted, public transportation needs have returned to normal levels. Numerous surges in demand for travel, particularly during peak hours such as concerts and festivals, have caused train operating compananies some concerns about their route arrangement. This sparked my interest to develop this project which introduces the Railway Decision Support System (RDSS) to assist route managers in determining the best strategy for sending trains through each railway segment. It can provide details of railway segment planning by determining the maximum number of trains that can be sent from one origin to one destination using Ford-Fulkerson algorithm, represented in its user-friendly functions within an Excel. 

## 2. Functions
<img src="https://github.com/PoonAthitS/railway-decision-support-system/blob/main/IMAGES/Decision%20Support%20System%20Tab.png?raw=true" width="500">
All of the RDSS's functions are purposefully concentrated on the "Decision Support System" tab, which contains five commands which includes to the following userforms.

### 2.1 Set up sheets
The first 2 commands are 1. Setup “Stations” sheet and 2. Setup “Rail Segments” sheet which they are designed to activate by creating these 2 sheets if they have not been created. The “Stations” sheet aims to provide the table of the user-entered station names while the “Rail segments”, on the other hand, contains a table about each input railway segment which includes Station [from], Station [To] and capacity fields

### 2.2 Enter railway data
<img src="https://github.com/PoonAthitS/railway-decision-support-system/blob/main/IMAGES/Setup%20Form.png?raw=true" width="400">
The third command is “3. Enter railway data”. The command automatically generates the setup form if both aforementioned sheets have been created. Its purpose is broken down into 2 parts. Firstly, a new station can be entered by the user in the upper section of the form to populate the specified list in the “Stations” sheet. Meanwhile, the second section focuses on creating a new railway segment in the “Rail segments” sheet by using the names of two of the stations entered in the first section. This input also requires the capacity of the segment or the maximum number of trains that could pass through it. Additionally, the system may notify the user by displaying a warning pop-up regarding railway network connectivity following a close button press. This is because RDDS runs the algorithm to determine whether or not the input network is correctly connected in this step.

### 2.3 Run the algorithm
<img src="https://github.com/PoonAthitS/railway-decision-support-system/blob/main/IMAGES/Algorithm%20Setup.png?raw=true" width="450">
After populating all stations and railway segments in the preceding functions, the user may run the algorithm by pressing the “4. Run the algorithm” command. The provided origin and destination stations are processed with all station and railway segment data in the Ford-Fulkerson algorithm to determine the maximum number possible of trains that can be sent from the origin to the destination.

## 3. Outcomes and examples
<img src="https://github.com/PoonAthitS/railway-decision-support-system/blob/main/IMAGES/Trial%20with%20GWR.png?raw=true" width="600">

* The outcomes are displayed in the "Optimal flows" sheet, with the maximum number of running trains displayed on the righthand side in red text and the entered origin and destination in blue. Each row contains information about the railway segment to which the algorithm decides to send the train, including its Station [From], Station [To], Number of running trains, Initial Capacity, Capacity left (remaining capacity after sending the train), and percentage capacity used. Moreover, if the railway section’s capacity is * fully utilised, the percentage cell will be automatically highlighted in green.

* According to the photo, this is the example when we input the origin = Bristol Parkway station, destination = Eastleigh station. The result shows that the maximum running train sending from Bristol Parkway is 21 trains.

## 4. About the programming

### 4.1 Files
The system is built in an Excel with VBA codes and user forms embeded: [Railway_decision_support_system_on_maximum_flow.xlsm](https://github.com/PoonAthitS/railway-decision-support-system/blob/main/Railway_decision_support_system_on_maximum_flow.xlsm)

### 4.2 Data
We use the mock-up simplified list of Great Western Railway major segments and stations as provided within the XLSM file

### To learn more about Poon Athit S., visit his [LinkedIn profile](https://www.linkedin.com/in/athit-srimachand/)
All rights reserved 2022. All codes are developed and owned by Poon Athit S. or the metioned team member(s). If you use this code, please visit his LinkedIn and give him a skill endorsement in Data analytics and the aforementioned coding technologies.
