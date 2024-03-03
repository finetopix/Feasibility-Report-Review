#import openpyxl
import pandas as pd
#import xlwings as xw
import sqlite3
import os

filename_nokia_hld = input('Please input Nokia HLD file name(HLD_reporting_2023-03-24.xlsx):')
if filename_nokia_hld == '':
    filename_nokia_hld = 'HLD_reporting_2023-05-05.xlsx'

filename_ejv_hld = input('Please input EJV HLD file name(JVD-008 - eJV 5G RAN HLD Report v2.xlsx):')
if filename_ejv_hld == '':
    filename_ejv_hld = 'JVD-008 - eJV 5G RAN HLD Report v2 17-04-2023.xlsx'

filename_rfmsl = input('Please input RFMSL file name(RF MSL.xlsx):')
if filename_rfmsl == '':
    filename_rfmsl = 'RF MSL.xlsx'
filename_rfnsa = input('Please input RFNSA file name(RFNSA_20230509.xlsx):')
if filename_rfnsa == '':
    filename_rfnsa = 'RFNSA_20230509.xlsx'
filename_nr35 = input('Please input RFMSL file name(NR35_Ideal_Spectrum_Allocation_C-Band_reshuffle.xlsx):')
if filename_nr35 == '':
    filename_nr35 = 'NR35_Ideal_Spectrum_Allocation_C-Band_reshuffle.xlsx'
        
if not os.path.isfile(filename_nokia_hld):
    print('Missing Nokia HLD file:',filename_nokia_hld)
    input('Press any key to exit...')
    exit()
        
if not os.path.isfile(filename_rfmsl):
    print('Missing RF MSL file:',filename_rfmsl)
    input('Press any key to exit...')
    exit()
        
if not os.path.isfile(filename_rfnsa):
    print('Missing RFNSA file:',filename_rfnsa)
    input('Press any key to exit...')
    exit()
        
#filename_ejv_hld = 'JVD-008 - eJV 5G RAN HLD Report v2.xlsx'
#filename_rfmsl = 'RF MSL.xlsx'
#filename_nokia_hld = 'HLD_reporting_2023-03-24.xlsx'
#filename_rfnsa = 'RFNSA.xlsx'
#filename_nr35 = 'NR35_Ideal_Spectrum_Allocation_C-Band_reshuffle.xlsx'


#read HLD site List into pandas dataframe
print('Reading NOKIA HLD data...')
df_nokia_hld = pd.read_excel(filename_nokia_hld,sheet_name='Site List',header=1)
#update data type
df_nokia_hld[['Vodafone_Site_ID']] = df_nokia_hld[['Vodafone_Site_ID']].fillna(value=0)
df_nokia_hld[['Vodafone_Site_ID']] = df_nokia_hld[['Vodafone_Site_ID']].astype(int)

#Filter columns
df_nokia_hld = df_nokia_hld[['JV Site Id',
                     'Site Name',
                     'Lat (GDA94)',
                     'Long (GDA94)',
                     'Vodafone_Site_ID',
                     'Program',
                     'Phase',
                     'Single Band RRU (LB)_LB Coverage Triggerred_Design Referrence',
                     'P.Antenna  Type 1 - V',
                     'P.Antenna Model 1 -V',
                     'Centreline 1',
                     'Azimuth 1 - V',
                     'M Tilt - V-P1',
                     'E Tilt LB - V-P1',
                     'E Tilt HB - V-P1',
                     'AAU1.Antenna Type - V',
                     'AAU2.Antenna Type - V',
                     'M Tilt AAU 1- V',
                     'E Tilt AAU1 - V',
                     'P.Antenna  Type 2 - V',
                     'P.Antenna Model 2 -V',
                     'Centreline 2',
                     'Azimuth 2 - V',
                     'P.Antenna  Type 3 - V',
                     'P.Antenna Model 3 -V',
                     'Centreline 3',
                     'Azimuth 3 - V',
                     'P.Antenna  Type 4 - V',
                     'P.Antenna Model 4 -V',
                     'Centreline 4',
                     'Azimuth 4 - V',
                     'P.Antenna  Type 5 - V',
                     'P.Antenna Model 5 -V',
                     'Centreline 5',
                     'Azimuth 5 - V',
                     'U900','U2100','M900','L700','L850','L900','L1800','L2100','L2600','N700','N3600',
                     'SAED','Preferred Design','Single Band RRU for U9 Capacity_Not for Design'
                    ]]


print('Reading EJV HLD data...')
df_ejv_hld = pd.read_excel(filename_ejv_hld)
#Filter some columns
df_ejv_hld = df_ejv_hld[['JV Site','TPG Site ID','Optus ACMA PreCheck Commentary','Optus ACMA PreCheck Result',
                 'Azimuth-1 -O','P.Antenna 1 Type -O',
                 'Azimuth-2 -O','P.Antenna 2 Type -O',
                 'Azimuth-3 -O','P.Antenna 3 Type -O',
                 'Azimuth-4 -O','P.Antenna 4 Type -O',
                 'Azimuth-5 -O','P.Antenna 5 Type -O',
                 'E-Tilt AAU1 -O','E-Tilt HB-P1 -O','E-Tilt LB-P1 -O','M-Tilt AAU1 -O','M-Tilt-P1 -O',
                 'Additional Requirements -O',
                 'L1800 -O','L2100 -O','L2300 -O','L2600 -O','L700 -O','L900 -O','NR2300 -O','NR3500 -O','U2100 -O','U900 -O','Additional Requirements -V'
    ]]
#fill NA value
df_ejv_hld[['TPG Site ID']] = df_ejv_hld[['TPG Site ID']].fillna(value=0)
df_ejv_hld[['TPG Site ID']] = df_ejv_hld[['TPG Site ID']].astype(int)

print('Reading RF MSL data...')
df_msl = pd.read_excel(filename_rfmsl,header=3)
#Filter some columns
df_msl = df_msl[['SiteID','Site','JV_Site_ID','NSA_Site_ID','ACMA_SiteID','LocationID',
                     'State','FacilityType','StructureOwner','Address','VHA_LRP_AREA','NRS_N700',
                     'NRS_L850','NRS_U900','NRS_M900','NRS_L1800','NRS_U2100','NRS_L2100','NRS_L2600',
                     'NRS_N3600','L700 BW','L850 BW','900 BW','L1800 BW','3GIS U2100 BW','VF U2100 BW',
                     'L2100_BW_APPARATUS','L2600 BW','SUA','3600 BW'
]]
#fill NA value
df_msl[['SiteID','NSA_Site_ID','ACMA_SiteID','LocationID']] = df_msl[['SiteID','NSA_Site_ID','ACMA_SiteID','LocationID']].fillna(value=0)
df_msl[['SiteID','NSA_Site_ID','ACMA_SiteID','LocationID']] = df_msl[['SiteID','NSA_Site_ID','ACMA_SiteID','LocationID']].astype(int)

#read RFNSA site into pandas dataframe
print('Reading RFNSA data')
df_rfnsa = pd.read_excel(filename_rfnsa,header=0)
df_rfnsa = df_rfnsa[['Site ID',
                     'Carrier Site Code',
                     'Carrier Site Name',
                     'Structure',
                     'Structure Latitude',
                     'Structure Longitude',
                     'ACMA Site IDs'
                     ]]
df_rfnsa.rename(columns={'Site ID': 'NSA', 'Carrier Site Code': 'Add ID'}, inplace=True)
print('Reading NR35 data...')
df_nr35 = pd.read_excel(filename_nr35,header=0)



#write to sqlite
print('Writing sqlite datebase...')
conn = sqlite3.connect('fr_database3.db')
#conn.commit()
df_nokia_hld.to_sql('NOK_HLD', conn, if_exists='replace', index = False)
df_ejv_hld.to_sql('EJV_HLD', conn, if_exists='replace', index = False)
df_msl.to_sql('RFMSL', conn, if_exists='replace', index = False)
df_rfnsa.to_sql('RFNSA', conn, if_exists='replace', index = False)
df_nr35.to_sql('NR35', conn, if_exists='replace', index = False)


print('Sqlite Database %s Updated' % 'fr_database3.db')
conn.commit()
conn.close()

