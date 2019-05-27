import xlsxwriter
import os
from datetime import datetime
import pandas as pd
import json

def header(json_data):
    json_data = json_data[0]
    today = datetime.today().strftime('%d-%B-%y')
    filename = 'filename_header_{0}'.format(datetime.now())
    workbook   = xlsxwriter.Workbook('{0}.xlsx'.format(filename))
    worksheet1 = workbook.add_worksheet()
    
   
    worksheet1.write(0,0,'Job No. :{0}'.format(json_data['jobno']), workbook.add_format({'bold': True,'border': 2}))
    
    worksheet1.write(0,1,'Customer :{0}'.format(json_data['customer']),workbook.add_format({'bold': True,'border': 2, 'align':'center'}))
    
    worksheet1.write(0,2,'Date :{0}'.format(today),workbook.add_format({'bold': True,'border': 2, 'align':'right'}))
    
    worksheet1.write(1,0,'Subject :150kvar Capacitor Bank',workbook.add_format({'bold': True,'border': 2}))

    worksheet1.write(1,1,None,workbook.add_format({'bold': True,'border': 2, 'align':'center'}))

    worksheet1.write(1,2,'Qty :1 Unit',workbook.add_format({'bold': True,'border': 2, 'align':'right'}))

    worksheet1.set_column(2, 2, width=30)
    worksheet1.set_column(1, 1, width=70)
    worksheet1.set_column(0, 0, 30)
    workbook.close()
    worksheet1.set_row(1,1,workbook.add_format({'border': 2}))
    file_path = '/home/thuya/Downloads/kopo/tutorial/{0}.xlsx'.format(filename)
    return file_path

def custom_xlsx_export(header_obj,body_obj):
    obj = body_obj 
    today = datetime.today().strftime('%d-%B-%y')
    filename = '/home/thuya/Desktop/filename_body_{0}'.format(datetime.now())
    workbook   = xlsxwriter.Workbook('{0}.xlsx'.format(filename))
    worksheet1 = workbook.add_worksheet()
    row_id = 0
    for sheet_index in obj:
        json_data = sheet_index
        num_column = 0
        bold = workbook.add_format({'bold': True,'border': 2})
        number_col_bold = workbook.add_format({'bold': True,'border': 2,'align':'center'})
        cell_number_format = workbook.add_format({'bold':True,'border': 2,'num_format': '#,##0'})
        """ Title start"""
        worksheet1.write(row_id,1,'Multi Power Engineering Co.,Ltd.', workbook.add_format({'bold': True,'align':'center'}))
        worksheet1.write(row_id + 1,1,'Quotation Report', workbook.add_format({'bold': True,'align':'center'}))  
        row_id = row_id + 2
        """ Title end """
        if header_obj:
            header = header_obj[0]
            """header level start"""
            worksheet1.write(row_id + 1,0,'Job No. :{0}'.format(header['jobno']), workbook.add_format({'bold': True,'border': 2}))   
            worksheet1.write(row_id + 1,1,'Customer :{0}'.format(header['customer']),workbook.add_format({'bold': True,'border': 2, 'align':'center'}))        
            worksheet1.write(row_id + 1,2,'Date :{0}'.format(today),workbook.add_format({'bold': True,'border': 2, 'align':'right'}))        
            row_id = row_id + 1
            worksheet1.write(row_id + 1,0,'Subject :150kvar Capacitor Bank',workbook.add_format({'bold': True,'border': 2}))
            worksheet1.write(row_id + 1,1,None,workbook.add_format({'bold': True,'border': 2, 'align':'center'}))
            worksheet1.write(row_id + 1,2,'Qty :1 Unit',workbook.add_format({'bold': True,'border': 2, 'align':'right'}))
            row_id = row_id + 2
            print("row_id", row_id)
            """header level end"""
            
        """sheet column name level start"""
        worksheet1.write(row_id + 1,0,'No.', number_col_bold)
        worksheet1.write(row_id + 1,1,'Description', workbook.add_format({'bold': True,'border': 2, 'align':'center'}))
        worksheet1.write(row_id + 1,2,'Amount (Kyat)', workbook.add_format({'bold': True,'border': 2,'align':'right'}))
        row_id = row_id + 1
        """ sheet column name level end """
        if "enclosure" in json_data and json_data['enclosure']['totalwithp'] > 0:
            row_data = json_data['enclosure']
            worksheet1.write((row_id + 1),0, '{0}{1}'.format(num_column + 1,'.'), number_col_bold)
            worksheet1.write((row_id + 1),1, row_data['label'], bold)
            worksheet1.write((row_id + 1),2, row_data['totalwithp'], cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1       

            """ Panel Enclosure Description """
            PED = 'Size : {0}mmH x {1}mmW x {2}mmD\nType : {3}\nMaterial : {4}\nPaint : {5}\nColour : {6}\nStandard : {7}'.format(
                row_data['height'],row_data['width'],row_data['depth']['name'],row_data['type'],
                row_data['material'],row_data['paint'],row_data['colour'],row_data['standard']
                )
            worksheet1.write('A{0}'.format(row_id+2),None, workbook.add_format({'border': 2}))
            worksheet1.write('B{0}'.format(row_id+2),PED, workbook.add_format({'border': 2}))
            worksheet1.write('C{0}'.format(row_id+2),None, workbook.add_format({'border': 2}))
            row_id = row_id + 1

        #level 4
        if "busbarfab" in json_data and json_data['busbarfab']['totalwithp'] > 0:
            row_data = json_data['busbarfab']
            worksheet1.write((row_id + 1),0, '{0}{1}'.format(num_column + 1,'.'), number_col_bold)
            worksheet1.write((row_id + 1),1, row_data['label'], bold)
            worksheet1.write((row_id + 1),2, row_data['totalwithp'], cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1
        #level 5
        if "busbar" in json_data and json_data['busbar']['totalwithp'] > 0:
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = json_data['busbar']
            worksheet1.write((row_id + 1),0, '{0}{1}'.format(num_column + 1,'A.'),number_col_bold)
            worksheet1.write((row_id + 1),1, row_data['label'], cell_format)
            worksheet1.write((row_id + 1),2, row_data['totalwithp'], cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1

            """ Bus bar: Heat Shrinking Tube, Accessory Description """
            if "items" in row_data:
                cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                BBD = ''
                for item in row_data["items"]:
                    UOM = ' {0}s'.format(item['uom']) if item['uom'] == 'No' and int(round(float(item['qty']))) > 1 else ' {0}'.format(item['uom'])
                    BBD = BBD  + ' -  ' + item['name'] + ' --- ' + item['qty'] + UOM + ' x ' + '{:1,}'.format(int(round(float(item['price'])))) +'\n'
                worksheet1.write('A{0}'.format(row_id+2),None, workbook.add_format({'border': 2}))
                worksheet1.write('B{0}'.format(row_id+2),BBD[:-1],cell_format)
                worksheet1.write('C{0}'.format(row_id+2),None, workbook.add_format({'border': 2}))
                row_id = row_id + 1
        if "cableandlug" in json_data and json_data['cableandlug']['totalwithp'] > 0:
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = json_data['cableandlug']
            worksheet1.write((row_id + 1),0, '{0}{1}'.format(num_column,'B.'), number_col_bold)
            worksheet1.write((row_id + 1),1, row_data['label'], cell_format)
            worksheet1.write((row_id + 1),2, row_data['totalwithp'], cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1
        elif "cableandlug" in json_data and json_data['cableandlug']['totalwithp'] == 0:
            num_column = num_column + 1 

        if "insulator" in json_data and json_data['insulator']['totalwithp'] > 0:
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = json_data['insulator']
            worksheet1.write((row_id + 1),0, '{0}{1}'.format(num_column,'.'), number_col_bold)
            worksheet1.write((row_id + 1),1, row_data['label'], cell_format)
            worksheet1.write((row_id + 1),2, row_data['totalwithp'], cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1    

            """ Insulator Description """
            if "items" in row_data:
                cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                ITD = ''
                for item in row_data["items"]:
                    UOM = ' {0}s'.format(item['uom']) if item['uom'] == 'No' and int(round(float(item['qty']))) > 1 else ' {0}'.format(item['uom'])
                    ITD = ITD  + ' -  ' + item['name'] + ' --- ' + str(item['qty']) + UOM + ' x ' + '{:1,}'.format(int(round(float(item['price'])))) +'\n'
                worksheet1.write('A{0}'.format(row_id+2),None, workbook.add_format({'border': 2}))
                worksheet1.write('B{0}'.format(row_id+2),ITD[:-1],cell_format)
                worksheet1.write('C{0}'.format(row_id+2),None, workbook.add_format({'border': 2}))
                row_id = row_id + 1
        
        if "metering" in json_data and json_data['metering']['totalwithp'] > 0:
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = json_data['metering']
            worksheet1.write((row_id + 1),0, '{0}{1}'.format(num_column,'.'), number_col_bold)
            worksheet1.write((row_id + 1),1, row_data['label'], cell_format)
            worksheet1.write((row_id + 1),2, row_data['totalwithp'], cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1    

            """ Metering Description """
            if "items" in row_data:
                cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                MTD = ''
                for item in row_data["items"]:
                    UOM = ' {0}s'.format(item['uom']) if item['uom'] == 'No' and int(round(float(item['qty']))) > 1 else ' {0}'.format(item['uom'])
                    MTD = MTD  + ' -  ' + item['name'] + ' --- ' + str(item['qty']) + UOM + ' x ' + str(item['price']) +'\n'
                worksheet1.write('A{0}'.format(row_id+2),None, workbook.add_format({'border': 2}))
                worksheet1.write('B{0}'.format(row_id+2),MTD[:-1],cell_format)
                worksheet1.write('C{0}'.format(row_id+2),None, workbook.add_format({'border': 2}))
                row_id = row_id + 1

        if "protection" in json_data and json_data['protection']['totalwithp'] > 0:
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = json_data['protection']
            worksheet1.write((row_id + 1),0, '{0}{1}'.format(num_column,'.'), number_col_bold)
            worksheet1.write((row_id + 1),1, row_data['label'], cell_format)
            worksheet1.write((row_id + 1),2, row_data['totalwithp'], cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1    

            """ Protection & Controls Description """
            if "items" in row_data and len(row_data["items"]) > 0:
                cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                PCD = ''
                for item in row_data["items"]:
                    PCD = PCD  + ' -  ' + item['name'] + ' --- ' + str(item['qty']) + ' Set x ' + str(item['price']) +'\n'
                worksheet1.write('A{0}'.format(row_id+2),None,cell_format)
                worksheet1.write('B{0}'.format(row_id+2),PCD[:-1],cell_format)
                worksheet1.write('C{0}'.format(row_id+2),None,cell_format)
                row_id = row_id + 1
                
        if "pricelist" in json_data and json_data['pricelist']['totalwithp'] > 0:
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = json_data['pricelist']
            if len(row_data["items"]) > 0:
                worksheet1.write((row_id + 1),0, '{0}{1}'.format(num_column,'.'), number_col_bold)
                worksheet1.write((row_id + 1),1, row_data['label'], cell_format)
                # worksheet1.write((row_id + 1),2, row_data['totalwithp'], cell_number_format)
                if row_data['customersupply'] == True:
                    worksheet1.write((row_id + 1),2, "Customer Supply", workbook.add_format({'bold': True,'border': 2,'align':'right'}))
                else:
                    worksheet1.write((row_id + 1),2, row_data["totalwithp"], workbook.add_format({'bold': True,'border': 2,'align':'right'}))    
                num_column = num_column + 1
                row_id = row_id + 1    

                """ Material List Description """
                if "items" in row_data:
                    cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                    MLD = ''
                    for item in row_data["items"]:
                        UOM = ' {0}s'.format('No') if int(round(float(item['qty']))) > 1 else ' {0}'.format('No')
                        MLD = MLD  + ' -  ' + item['brandname'] + ' - ' + item['description'] + ' --- ' + str(item['qty']) + UOM + ' x ' + str(item['price']) +'\n'
                    worksheet1.write('A{0}'.format(row_id+2),None, workbook.add_format({'border': 2}))
                    worksheet1.write('B{0}'.format(row_id+2),MLD[:-1],cell_format)
                    worksheet1.write('C{0}'.format(row_id+2),None, workbook.add_format({'border': 2}))
                    row_id = row_id + 1
            
        if "extrapricelist" in json_data and json_data['extrapricelist']['totalwithp'] > 0:
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = json_data['extrapricelist']
            if len(row_data["items"]) > 0:
                worksheet1.write((row_id + 1),0, '{0}{1}'.format(num_column,'.'), number_col_bold)
                worksheet1.write((row_id + 1),1, row_data['label'], cell_format)
                if row_data['customersupply'] == True:
                    worksheet1.write((row_id + 1),2, "Customer Supply", workbook.add_format({'bold': True,'border': 2,'align':'right'}))
                else:
                    worksheet1.write((row_id + 1),2, row_data["totalwithp"], workbook.add_format({'bold': True,'border': 2,'align':'right'}))
                num_column = num_column + 1
                row_id = row_id + 1    

                """Extra Material List Description """
                if "items" in row_data:
                    cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                    EMLD = ''
                    for item in row_data["items"]:
                        UOM = ' {0}s'.format('No') if int(round(float(item['qty']))) > 1 else ' {0}'.format('No')
                        EMLD = EMLD  + ' -  ' + item['brandname'] + '  -- ' + item['description'] + ' --- ' + str(item['qty']) + UOM +'\n'
                    worksheet1.write('A{0}'.format(row_id+2),None, workbook.add_format({'border': 2}))
                    worksheet1.write('B{0}'.format(row_id+2),EMLD[:-1],cell_format)
                    worksheet1.write('C{0}'.format(row_id+2),None, workbook.add_format({'border': 2}))
                    row_id = row_id + 1

        if "extraonepricelist" in json_data and json_data['extraonepricelist']['totalwithp'] > 0:
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = json_data['extraonepricelist']
            if len(row_data["items"]) > 0:
                worksheet1.write((row_id + 1),0, '{0}{1}'.format(num_column,'.'), number_col_bold)
                worksheet1.write((row_id + 1),1, row_data['label'], cell_format)
                if row_data['customersupply'] == True:
                    worksheet1.write((row_id + 1),2, "Customer Supply", workbook.add_format({'bold': True,'border': 2,'align':'right'}))
                else:    
                    worksheet1.write((row_id + 1),2, row_data["totalwithp"], workbook.add_format({'bold': True,'border': 2,'align':'right'}))
                num_column = num_column + 1
                row_id = row_id + 1    

                """Extra Material List Description """
                if "items" in row_data:
                    cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                    EMLD = ''
                    for item in row_data["items"]:
                        UOM = ' {0}s'.format('No') if int(round(float(item['qty']))) > 1 else ' {0}'.format('No')
                        EMLD = EMLD  + ' -  ' + item['brandname'] + '  -- ' + item['description'] + ' --- ' + str(item['qty']) + UOM +'\n'
                    worksheet1.write('A{0}'.format(row_id+2),None, workbook.add_format({'border': 2}))
                    worksheet1.write('B{0}'.format(row_id+2),EMLD[:-1],cell_format)
                    worksheet1.write('C{0}'.format(row_id+2),None, workbook.add_format({'border': 2}))
                    row_id = row_id + 1        
            
        if "extratwopricelist" in json_data and json_data['extratwopricelist']['totalwithp'] > 0: 
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = json_data['extratwopricelist']
            if len(row_data["items"]) > 0:
                worksheet1.write((row_id + 1),0, '{0}{1}'.format(num_column,'.'), number_col_bold)
                worksheet1.write((row_id + 1),1, row_data['label'], cell_format)
                if row_data['customersupply'] == True:
                    worksheet1.write((row_id + 1),2, "Customer Supply", workbook.add_format({'bold': True,'border': 2,'align':'right'}))
                else:
                    worksheet1.write((row_id + 1),2, row_data["totalwithp"], workbook.add_format({'bold': True,'border': 2,'align':'right'}))
                num_column = num_column + 1
                row_id = row_id + 1    

                """Extra Material List Description """
                if "items" in row_data:
                    cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                    EMLD = ''
                    for item in row_data["items"]:
                        UOM = ' {0}s'.format('No') if int(round(float(item['qty']))) > 1 else ' {0}'.format('No')
                        EMLD = EMLD  + ' -  ' + item['brandname'] + '  -- ' + item['description'] + ' --- ' + str(item['qty']) + UOM +'\n'
                    worksheet1.write('A{0}'.format(row_id+2),None, workbook.add_format({'border': 2}))
                    worksheet1.write('B{0}'.format(row_id+2),EMLD[:-1],cell_format)
                    worksheet1.write('C{0}'.format(row_id+2),None, workbook.add_format({'border': 2}))
                    row_id = row_id + 1  
                    
        



        cell_format = workbook.add_format({'bold':True,'border': 2})
        worksheet1.write((row_id + 1),0,None, workbook.add_format({'border': 2}))
        worksheet1.write((row_id + 1),1,"Total in (Kyat)", cell_format)
        total_in_kyat = 0
        for item in json_data:
            if isinstance(json_data[item], dict):
                #print("ITEM",type(json_data[item]))
                if 'customersupply' in  json_data[item]:
                    if json_data[item]['customersupply'] == False:
                        total_in_kyat += int(float(json_data[item]['totalwithp']))
                else:
                    total_in_kyat += int(float(json_data[item]['totalwithp']))
        worksheet1.write((row_id + 1),2, total_in_kyat, cell_number_format)

        row_id = row_id + 5
    #worksheet1.set_column(2, 2, 30, None, {'level': 1})
    worksheet1.set_column(2, 2, width=30)
    worksheet1.set_column(1, 1, width=70)
    worksheet1.set_column(0, 0, 10)
    workbook.close()
    file_path = '{0}.xlsx'.format(filename)
    return file_path
    