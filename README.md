# Sensortower.py
Scraping sensortower

Scrapping sensor tower site to find requested by client info about app. For scrapper i used Selenium, because it is interactive sice also we need to log in.

Challenge in the table. 
 I need to find app revenue and app download in table. Looks easy, but table organised like mapping particular cell to another. 
if so i cant find app by number and get like excel sheet "B42".
 I am going to find tag that wraping group of cells for this app by matching attribute "data-entity-id" to app id 
 
 Than write all data to xlsx sheet with openpyxl
 
