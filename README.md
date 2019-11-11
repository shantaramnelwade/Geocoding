# shanta
import pandas as pd #import pandas for accessing excel and making dataframe 
import geopy as gy # importing for geocoding
import openpyxl as xl #writing in new excel
locc=gy.geocoders.ArcGIS() #intialising the geocoder
wbs=xl.load_workbook('abc.xlsx') #reading excel sheet from directory
data=wbs.get_active_sheet() # getting active sheet

for i in range(1,len(data['A'])): #start the range from one because excel sheet start from first row not zero
    st_addr=data['P{}'.format(i)].value # read street adress from column P 
    strst_addr=str(st_addr) # converting to the string
    ld_addr=data['N{}'.format(i)].value #read landmark adress from excel 
    strld_adrr=str(ld_adrr) # convering to string
    ct_adrr=data['O{}'.format(i)].value # read city names
    strct_adrr=str(ct_adrr) # converting into string

    adress=strld_adrr+','+strst_addr+','+strct_adrr # constructing adress from concating the adress strings
    locadress=locc.geocode(adress,timeout=150)# timeout represents the running time out of geocoding
    latt=locadress.latitude # getting lattitude 
    print(latt)
    data.cell(row=i,column=24,value=latt) # writing into the excel sheet, specify the column number to write in excel
    long=locadress.longitude # getting longitude
    print(long)
    data.cell(row=i,column=25,value=long)# writing in excel sheet, specify the column number to write in excel
 # exit()  
wbs.save('abc.xlsx')# saved the file in current directory
    
