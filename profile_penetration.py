#Import packages
import pandas as pd
import numpy as np

#Set the output row, 365 for a year, 4380 for 10 years
row_output=int(input("Enter the row number (day) for output"))

#Set number of year
tahun=range(2021,2033)

#Set output columns
output_cols=['Year','Month','Day', 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24]

def output(row_output):

    sheet=pd.ExcelFile("Profil Wind.xlsx").sheet_names[1:]

    for sheets in sheet:
        #Make new output dataframe
        output=pd.DataFrame(index=range(0,row_output), columns=output_cols)

        #Read and modifiy input columns
        df=pd.read_excel("Profil Wind.xlsx", sheet_name=sheets)
        columns_=df.iloc[0, 4:21]
        df=df.iloc[1:, 4:21]
        df.columns=list(columns_)

        #Creating year data frame
        df_tahun=pd.DataFrame(tahun, columns=['Tahun'])
        df_tahun['Tahun']=pd.to_datetime(df_tahun['Tahun'], format='%Y')
        
        #Splitting to year, months, day
        date_year=pd.date_range(start=df_tahun['Tahun'][0], periods=row_output+3, freq='D')
        df_date=pd.DataFrame(date_year, columns=['Date']).astype(str)
        df_date[['Year','Month','Day']]=df_date['Date'].str.split('-', expand=True)
        df_date=df_date[['Year','Month','Day']]

        #Masking leap year
        mask_leap_year=((df_date['Year']=='2024') | (df_date['Year']=='2028') | (df_date['Year']=='2032')) & (df_date['Month']=='02') & (df_date['Day']=='29')
        df_date=df_date[~mask_leap_year].reset_index(drop=True)

        #Modify output data frame (year, months, day)
        output['Year']=df_date.iloc[:,0]
        output['Month']=df_date.iloc[:,1]
        output['Day']=df_date.iloc[:,2]

        #Resize input
        input=pd.concat([df.iloc[:,4]]*12, ignore_index=True).values
        input.resize(row_output,24)
        output.iloc[0:row_output+1, 3:]=input

        #Save output into csv
        output.to_csv(f'wind_profile_csv\{sheets}.csv', index=False)