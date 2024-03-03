import openpyxl
#import pandas as pd
import xlwings as xw
import os
import sqlite3
import time

'''
v1.0 read NOKIA_HLD,EJV_HLD,RF_MSL separately;
siteid as filter for multiple sites.
#query sqlite datebase to get siteinfo
v2.0 read info from sqlite database
v2.1 add reviewer initial to file name.

'''

def list2sheet(wb,sheet,siteinfo):
    sht=wb.sheets(sheet)
    first_cell_row = 2 #start from row 2
    if len(siteinfo) == 0:
        print('    no site info found in the table',sheet)
        #exit()
    if len(siteinfo) > 0:
        row = first_cell_row
        for i in range(len(siteinfo)):
            row = row + i
            y = str(row)
            startcell = 'A' + y
            sht.range(startcell).value = siteinfo[i]

if __name__ == '__main__':
    databasename = input('Please input the sqlite database file name(fr_database3.db):')
    if databasename == '':
        databasename = 'fr_database3.db'

    file_sitelist = input('Please input site list(sitelist.text):')
    if file_sitelist == '':
        file_sitelist = 'sitelist.txt'

    review_template = input('Please input the template file(FREBP3.xlsm):')
    if review_template == '':
        review_template = 'FREBP3.xlsm'
                 
    if not os.path.isfile(databasename):
        print('Missing database file:',databasename)
        input('Press any key to exit...')
        exit()
        
    if not os.path.isfile(file_sitelist):
        print('Missing site id file:',file_sitelist)
        input('Press any key to exit...')
        exit()
        
    if not os.path.isfile(review_template):
        print('Missing review template file:',review_template)
        input('Press any key to exit...')
        exit()

    siteid_list = []
    file_sitelist = os.getcwd() + '\\' + file_sitelist
    with open(file_sitelist,mode='r') as f:
        lines = f.readlines()
        for line in lines:
            line = line.rstrip('\n')
            siteid_list.append(line) #TPG site id list
    
    conn = sqlite3.connect(databasename)
    c = conn.cursor()

    print('Generating review files...')
        
    for siteid in siteid_list:
        print('    ' + str(siteid))
        app = xw.App(visible=False,add_book=False)
        app.display_alerts=False
        app.interactive=False
        try:
            wb = app.books.open(review_template) #open excel file for FRreview site,using new template.
        except Exception as inst:
            pass

        #nokia hld information
        sql1 ='''
        SELECT * FROM NOK_HLD
        WHERE Vodafone_Site_ID = ''' + str(siteid)
        c.execute(sql1)
        siteinfo_nokia_hld = c.fetchall()
        
        if len(siteinfo_nokia_hld) == 0:
            print('    no site info in NOKIA HLD for site:',siteid)
            #continue #if site is not in Nokia HLD, GO TO next site id
            pass

        #ejv hld information:
        sql1 = '''
        SELECT * FROM EJV_HLD
        WHERE "TPG Site ID" = ''' + str(siteid)
        c.execute(sql1)
        siteinfo_ejv_hld = c.fetchall()
            
        #RFMSL information
        c.execute('''
        SELECT * FROM RFMSL
        WHERE SiteID = ''' + str(siteid)
        )
        siteinfo_msl = c.fetchall()
        if len(siteinfo_msl) > 0:
            site_rfnsa_id = siteinfo_msl[0][3]
            sitename = siteinfo_msl[0][1]
            c.execute('''SELECT * FROM RFNSA
        WHERE "Add ID" = ''' + str(siteid)
        )
            siteinfo_rfnsa = c.fetchall()
            list2sheet(wb,'RFNSA',siteinfo_rfnsa) # writing rfnsa information to sheet
        
        #NR35 reshuffle information
        c.execute('''SELECT * FROM NR35
        WHERE Add_ID = ''' + str(siteid)
        )
        siteinfo_nr35 = c.fetchall()
        

        if len(siteinfo_nokia_hld) > 0:
            sitename = siteinfo_nokia_hld[0][1]
        
        
        list2sheet(wb,'NOK_HLD',siteinfo_nokia_hld) # writing nokia hld information to sheet
        list2sheet(wb,'EJV_HLD',siteinfo_ejv_hld) # writing ejv hld information to sheet
        list2sheet(wb,'RFMSL',siteinfo_msl) # writing RFMSL information to sheet
        list2sheet(wb,'NR35',siteinfo_nr35) # writing NR3500 reshuffle information to sheet
        
        
        wb.sheets('Main').range('A1').value = siteid
        if wb.sheets('Main').range('B35').value is None:
            filename = wb.sheets('RFMSL').range('B2').value + '.xlsm'
            
        else:
            filename = wb.sheets('Main').range('B35').value + '.xlsm'
            
        print('    ' + filename + ' generated!')
        if os.path.isfile(filename):
            os.remove(filename)
            time.sleep(0.5)
        try:
            wb.save(filename)
        except Exception as inst:
            print('    Saving ',filename)
            print(type(inst))
            #print(inst.args)
            #print(inst)
            pass
        wb.close
        try:
            app.quit()
        except Exception as inst:
            print('quit excel app')
            print(type(inst))
            #print(inst.args)
            #print(inst)
            pass
        
    print('All files saved!')
    input("Press Enter to exit...")

