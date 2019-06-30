import xlsxwriter
import os
from datetime import datetime
import json
import inflect
import math
def rownumber_to_columnstring(n):
    return (n + 1)
def columnstring_to_rownumber(n):
    return (n -1)



def replace_all(text, dic):
    for i, j in dic.iteritems():
        text = text.replace(i, j)
    return text

def custom_xlsx_export(header_obj,body_obj):
    obj = body_obj 
    today = datetime.today().strftime('%d-%B-%y')
    filename = '/home/thuya/OUTSOURCE/test/filename_body_{0}'.format(datetime.now())
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

def quotation_xlsx_export(header_obj,unit_obj,body_obj):
    header = header_obj[0]
    unit = unit_obj[0]
    today = datetime.today().strftime('%d-%B-%y')
    filename = '/home/thuya/OUTSOURCE/test/quotation_{0}'.format(datetime.now())
    workbook   = xlsxwriter.Workbook('{0}.xlsx'.format(filename))
    worksheet = workbook.add_worksheet()

    bold = workbook.add_format({'bold': True, 'underline':True})
    merge_format = workbook.add_format({'align': 'left'})
    worksheet.merge_range('A1:E1',header['qdate'], merge_format)
    worksheet.merge_range('A2:E2',"To.", merge_format)
    worksheet.merge_range('A3:E3',header['companyname'], workbook.add_format({'border':1,'align': 'left'}))
    worksheet.merge_range('A4:E4',"Att       : {0}".format(header['customer']), merge_format)
    worksheet.merge_range('A5:E5',"Ph        : {0}".format(header['phoneno']), merge_format)
    worksheet.merge_range('A6:E6',"Email    : {0}".format(header['email']), merge_format)
    worksheet.merge_range('A7:E7',"CC        : {0}".format(header['ccemail']), merge_format)
    worksheet.merge_range('A8:E8',"Q No     : {0}".format(header['jobno']), merge_format)
    # worksheet.merge_range('A9:E9',"Subject : {0}".format(unit['subject']), merge_format)
    worksheet.write_rich_string(
        'A9:E9',
        'Subject : ',
        bold,unit['subject']
    )
    d = {"<h3>Dear Sir,</h3>":"","<p>":"","</p>":""}
    last = replace_all(unit['header'], d)
    worksheet.merge_range('A11:E11',"Dear Sir,", workbook.add_format({'bold': True}))
    worksheet.merge_range('A12:E12',last, merge_format)
    table_header_cell_format = workbook.add_format({'align': 'center','border':1})
    worksheet.write(12,0,"No.", table_header_cell_format)
    worksheet.write(12,1,"Description", table_header_cell_format)
    worksheet.write(12,2,"Qty.", table_header_cell_format)
    worksheet.write(12,3,"Price", table_header_cell_format)
    worksheet.write(12,4,"Amount", table_header_cell_format)

    table_data_cell_format = workbook.add_format({'align': 'center','border':1})
    row_id = 13
    table_col_num = 1
    total_units = 0
    total_amount = 0
    for obj in body_obj:
        print("obj['subject']", obj['subject'])
        u = "Units" if obj['unit'] > 1 else "Unit"
        amount = obj['unit'] * obj['price']
        worksheet.write(row_id,0,table_col_num, table_data_cell_format)
        worksheet.write(row_id,1,obj['subject'], workbook.add_format({'align': 'left','border':1}))
        worksheet.write(row_id,2,"{0}{1}".format(obj['unit'], u), table_data_cell_format)
        worksheet.write(row_id,3,obj['price'], workbook.add_format({'align': 'right','border':1, 'num_format': '#,##0'}))
        worksheet.write(row_id,4, amount, workbook.add_format({'align': 'right','border':1, 'num_format': '#,##0'}))
        total_units += obj['unit']
        total_amount += amount
        row_id += 1
        table_col_num += 1

    #Summary Row
    u = "Units" if total_units > 1 else "Unit"
    worksheet.write(row_id,0, None, table_data_cell_format)
    worksheet.write(row_id,1,"Total in (Kyat)", table_data_cell_format)
    worksheet.write(row_id,2, "{0}{1}".format(total_units, u), table_data_cell_format)
    worksheet.write(row_id,3, None, table_data_cell_format)
    worksheet.write(row_id,4, total_amount, workbook.add_format({'align': 'right','border':1, 'num_format': '#,##0'}))
    row_id += 1

    #Extra Header
    col_row_num = rownumber_to_columnstring(row_id)
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num),'Price Validity       :    {0}    '.format(unit['validity']), merge_format)
    col_row_num += 1
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num),'Payment Term     :    {0}'.format(unit['paymentterm']), merge_format)
    col_row_num += 1
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num),'Delivery date      :    {0}'.format(unit['delivery']), merge_format)
    col_row_num += 1
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num),'Drawing              :    {0}'.format(unit['drawing']), merge_format)
    col_row_num += 1

    """footer text start"""
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num),None, merge_format)
    col_row_num += 1
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num),None, merge_format)
    col_row_num += 1
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num),None, merge_format)
    col_row_num += 1
    d = {"<h3>":"","</h3>":"\n","<p>":"","</p>":""}
    footer_text = replace_all(unit['footer'], d)
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num), footer_text, merge_format)
    col_row_num += 1

    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num),None, merge_format)
    col_row_num += 1
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num),None, merge_format)
    col_row_num += 1
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num),"With Regard,", merge_format)
    col_row_num += 1
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num),None, merge_format)
    col_row_num += 1
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num),None, merge_format)
    col_row_num += 1
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num),None, merge_format)
    col_row_num += 1

    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num), "-------------------------------", merge_format)
    col_row_num += 1
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num), "Yan Tun (09-400686491/09-451162999)", merge_format)
    col_row_num += 1
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num), "Sale & Marketing Manager", merge_format)
    col_row_num += 1
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num), "MPEC", merge_format)
    col_row_num += 1
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num), "ZPW", merge_format)
    col_row_num += 1

    """footer text end"""

    """Start Data Table"""
    panels = unit['panels']
    panels.sort(key=lambda item: item.get("name"))
    print("gg", panels)
    row_id = columnstring_to_rownumber(col_row_num)
    for panel in panels:
        table_body_data = [x for x in body_obj if x["id"] == panel["id"] != None][0]
        worksheet.write(row_id,0,panel['name'], workbook.add_format({'bold': True,'align':'center'}))
        row_id += 1
        """sheet column name level start"""
        worksheet.write(row_id,0,'No.', workbook.add_format({'bold': True,'border': 1, 'align':'center'}))
        worksheet.write(row_id,1,'Description', workbook.add_format({'bold': True,'border': 1, 'align':'center'}))
        worksheet.write(row_id,2,'Qty.', workbook.add_format({'bold': True,'border': 1,'align':'center'}))
        worksheet.write(row_id,3,'Price', workbook.add_format({'bold': True,'border': 1,'align':'center'}))
        worksheet.write(row_id,4,'Amount', workbook.add_format({'bold': True,'border': 1,'align':'center'}))
        row_id += 1

        num_column = 0
        number_col_bold = workbook.add_format({'bold': True,'border': 1,'align':'center'})
        cell_number_format = workbook.add_format({'bold':True,'border': 1,'num_format': '#,##0'})
        """Enclosure Start"""
        if "enclosure" in table_body_data:
            unit = 1
            row_data = table_body_data['enclosure']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column + 1,'.'), number_col_bold)

            """ Panel Enclosure Description """
            PED = 'Panel Enclosure:\nSize : {0}mmH x {1}mmW x {2}mmD\nType : {3}\nMaterial : {4}\nPaint : {5}\nColour : {6}\nStandard : {7}'.format(
                row_data['height'],row_data['width'],row_data['depth']['name'],row_data['type'],
                row_data['material'],row_data['paint'],row_data['colour'],row_data['standard']
                )
            worksheet.write(row_id,1, PED,workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "Unit"),  workbook.add_format({'border': 1,'align':'center'}))
            worksheet.write(row_id,3, row_data['totalwithp'],cell_number_format)
            worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1
        """Enclosure End""" 

        if "busbarfab" in table_body_data:
            unit = 1
            row_data = table_body_data['busbarfab']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column + 1,'.'), number_col_bold)
            worksheet.write(row_id,1, row_data['label'], workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "No"), workbook.add_format({'border': 1,'align':'center'}))
            worksheet.write(row_id,3, row_data['totalwithp'], cell_number_format)
            worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1

        """BusBar Start"""
        if "busbar" in table_body_data:
            unit = 1
            row_data = table_body_data['busbar']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column + 1,'A.'), number_col_bold)

            if "items" in row_data:
                cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                BBD = '{0}\n'.format(row_data['label'])
                for item in row_data["items"]:
                    UOM = ' {0}s'.format(item['uom']) if item['uom'] == 'No' and int(round(float(item['qty']))) > 1 else ' {0}'.format(item['uom'])
                    BBD = BBD  + ' -  ' + item['name'] + ' --- ' + item['qty'] + UOM + ' x ' + '{:1,}'.format(int(round(float(item['price'])))) +'\n'
          
            worksheet.write(row_id,1, BBD[:-1],workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "Lot"),  workbook.add_format({'border': 1,'align':'center'}))
            worksheet.write(row_id,3, row_data['totalwithp'],cell_number_format)
            worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1

        if "cableandlug" in table_body_data and  table_body_data['cableandlug']['totalwithp'] > 0:
            unit = 1
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = table_body_data['cableandlug']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column,'B.'), number_col_bold)
            worksheet.write(row_id,1, row_data['label'], workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "Lot"),  workbook.add_format({'border': 1,'align':'center'}))
            worksheet.write(row_id,3, row_data['totalwithp'], cell_number_format)
            worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1    
        
        if "insulator" in table_body_data and table_body_data['insulator']['totalwithp'] > 0:
            unit = 1
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = table_body_data['insulator']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column + 1,'.'), number_col_bold)
            """ Insulator Description """
            if "items" in row_data:
                cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                ITD = '{0}\n'.format(row_data['label'])
                for item in row_data["items"]:
                    UOM = ' {0}s'.format(item['uom']) if item['uom'] == 'No' and int(round(float(item['qty']))) > 1 else ' {0}'.format(item['uom'])
                    ITD = ITD  + ' -  ' + item['name'] + ' --- ' + str(item['qty']) + UOM + ' x ' + '{:1,}'.format(int(round(float(item['price'])))) +'\n'
            worksheet.write(row_id,1, ITD[:-1],workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "Lot"),  workbook.add_format({'border': 1,'align':'center'}))
            worksheet.write(row_id,3, row_data['totalwithp'],cell_number_format)
            worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1

        if "metering" in table_body_data and table_body_data['metering']['totalwithp'] > 0:
            unit = 1
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = table_body_data['metering']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column + 1,'.'), number_col_bold)

            """ Metering Description """
            if "items" in row_data:
                cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                MTD = '{0}\n'.format(row_data['label'])
                for item in row_data["items"]:
                    UOM = ' {0}s'.format(item['uom']) if item['uom'] == 'No' and int(round(float(item['qty']))) > 1 else ' {0}'.format(item['uom'])
                    MTD = MTD  + ' -  ' + item['name'] + ' --- ' + str(item['qty']) + UOM + ' x ' + str(item['price']) +'\n'
            worksheet.write(row_id,1, MTD[:-1],workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "Lot"),  workbook.add_format({'border': 1,'align':'center'}))
            worksheet.write(row_id,3, row_data['totalwithp'],cell_number_format)
            worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1   

        if "protection" in table_body_data and table_body_data['protection']['totalwithp'] > 0:
            unit = 1
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = table_body_data['protection']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column + 1,'.'), number_col_bold)

            """ Protection & Controls Description """
            if "items" in row_data and len(row_data["items"]) > 0:
                cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                PCD = '{0}\n'.format(row_data['label'])
                for item in row_data["items"]:
                    PCD = PCD  + ' -  ' + item['name'] + ' --- ' + str(item['qty']) + ' Set x ' + str(item['price']) +'\n'
            worksheet.write(row_id,1, PCD[:-1],workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "Lot"),  workbook.add_format({'border': 1,'align':'center'}))
            worksheet.write(row_id,3, row_data['totalwithp'],cell_number_format)
            worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1 

        if "pricelist" in table_body_data and table_body_data['pricelist']['totalwithp'] > 0:
            unit = 1
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = table_body_data['pricelist']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column + 1,'.'), number_col_bold)

            """ Material List Description """
            if "items" in row_data:
                cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                MLD = '{0}\n'.format(row_data['label'])
                for item in row_data["items"]:
                    UOM = ' {0}s'.format('No') if int(round(float(item['qty']))) > 1 else ' {0}'.format('No')
                    MLD = MLD  + ' -  ' + item['brandname'] + ' - ' + item['description'] + ' --- ' + str(item['qty']) + UOM + ' x ' + str(item['price']) +'\n'
            worksheet.write(row_id,1, MLD[:-1],workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "Lot"),  workbook.add_format({'border': 1,'align':'center'}))
            if row_data['customersupply'] == True:
                worksheet.write(row_id,3, None,cell_number_format)
                worksheet.write(row_id,4, "Customer Supply", workbook.add_format({'valign':'middle','align':'right','border': 1,'num_format': '#,##0'}))
            else:
                worksheet.write(row_id,3, row_data['totalwithp'],cell_number_format)
                worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1

        if "extrapricelist" in table_body_data and table_body_data['extrapricelist']['totalwithp'] > 0:
            unit = 1
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = table_body_data['extrapricelist']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column + 1,'.'), number_col_bold)

            """Extra Material List Description """
            if "items" in row_data:
                cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                EMLD = ''
                for item in row_data["items"]:
                    UOM = ' {0}s'.format('No') if int(round(float(item['qty']))) > 1 else ' {0}'.format('No')
                    EMLD = EMLD  + ' -  ' + item['brandname'] + '  -- ' + item['description'] + ' --- ' + str(item['qty']) + UOM +'\n'
            worksheet.write(row_id,1, EMLD[:-1],workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "Lot"),  workbook.add_format({'border': 1,'align':'center'}))
            if row_data['customersupply'] == True:
                worksheet.write(row_id,3, None,cell_number_format)
                worksheet.write(row_id,4, "Customer Supply", workbook.add_format({'valign':'middle','align':'right','border': 1,'num_format': '#,##0'}))
            else:
                worksheet.write(row_id,3, row_data['totalwithp'],cell_number_format)
                worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1

        if "extraonepricelist" in table_body_data and table_body_data['extraonepricelist']['totalwithp'] > 0:
            unit = 1
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = table_body_data['extraonepricelist']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column + 1,'.'), number_col_bold)

            """Extra Material List Description """
            if "items" in row_data:
                cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                EMLD = ''
                for item in row_data["items"]:
                    UOM = ' {0}s'.format('No') if int(round(float(item['qty']))) > 1 else ' {0}'.format('No')
                    EMLD = EMLD  + ' -  ' + item['brandname'] + '  -- ' + item['description'] + ' --- ' + str(item['qty']) + UOM +'\n'
            worksheet.write(row_id,1, EMLD[:-1],workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "Lot"),  workbook.add_format({'border': 1,'align':'center'}))
            if row_data['customersupply'] == True:
                worksheet.write(row_id,3, None,cell_number_format)
                worksheet.write(row_id,4, "Customer Supply", workbook.add_format({'valign':'middle','align':'right','border': 1,'num_format': '#,##0'}))
            else:
                worksheet.write(row_id,3, row_data['totalwithp'],cell_number_format)
                worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1                 
        
        if "extratwopricelist" in table_body_data and table_body_data['extratwopricelist']['totalwithp'] > 0:
            unit = 1
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = table_body_data['extratwopricelist']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column + 1,'.'), number_col_bold)

            """Extra Material List Description """
            if "items" in row_data:
                cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                EMLD = ''
                for item in row_data["items"]:
                    UOM = ' {0}s'.format('No') if int(round(float(item['qty']))) > 1 else ' {0}'.format('No')
                    EMLD = EMLD  + ' -  ' + item['brandname'] + '  -- ' + item['description'] + ' --- ' + str(item['qty']) + UOM +'\n'
            worksheet.write(row_id,1, EMLD[:-1],workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "Lot"),  workbook.add_format({'border': 1,'align':'center'}))
            if row_data['customersupply'] == True:
                worksheet.write(row_id,3, None,cell_number_format)
                worksheet.write(row_id,4, "Customer Supply", workbook.add_format({'valign':'middle','align':'right','border': 1,'num_format': '#,##0'}))
            else:
                worksheet.write(row_id,3, row_data['totalwithp'],cell_number_format)
                worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1


        cell_format = workbook.add_format({'bold':True,'border': 1})
        worksheet.write(row_id,0,None, workbook.add_format({'border': 1}))
        worksheet.write(row_id,1,"Total in (Kyat)", cell_format)
        worksheet.write(row_id,2,None, cell_format)
        worksheet.write(row_id,3,None, cell_format)
        total_in_kyat = 0
        for item in table_body_data:
            if isinstance(table_body_data[item], dict):
                #print("ITEM",type(json_data[item]))
                if 'customersupply' in  table_body_data[item]:
                    if table_body_data[item]['customersupply'] == False:
                        total_in_kyat += int(float(table_body_data[item]['totalwithp']))
                else:
                    total_in_kyat += int(float(table_body_data[item]['totalwithp']))
        worksheet.write(row_id,4, total_in_kyat, cell_number_format)

        row_id += 2
        # worksheet.write(row_id,0,table_col_num, table_data_cell_format)
        



    """END Data Table"""




    worksheet.set_column(0, 0, 10)
    worksheet.set_column(1, 1, width=70)
    worksheet.set_column(2, 2, width=20)
    worksheet.set_column(3, 3, width=20)
    worksheet.set_column(4, 4, width=20)
    worksheet.set_column(5, 5, width=20)
    workbook.close()
    file_path = '{0}.xlsx'.format(filename)
    return file_path

def invoice_xlsx_export(header_obj,body_obj,currency_obj,payment_obj,signature_obj,invoice_list_obj):
    p = inflect.engine()
    header = header_obj[0]
    payment = payment_obj[0]
    invoice_list = invoice_list_obj[0]
    #unit = unit_obj[0]
    today = datetime.today().strftime('%d-%B-%y')
    filename = '/home/thuya/OUTSOURCE/test/invoice_{0}'.format(datetime.now())
    workbook   = xlsxwriter.Workbook('{0}.xlsx'.format(filename))    
    worksheet = workbook.add_worksheet()
    format = workbook.add_format()
    format.set_font_name('Times New Roman')
    row_id = 20
    merge_format = workbook.add_format({'align': 'left'})
    bold = workbook.add_format({'bold': True})
    col_row_num = rownumber_to_columnstring(row_id)
    worksheet.merge_range('A1:E1',"INVOICE", workbook.add_format({'align': 'center'}))
    worksheet.merge_range('A3:B5', None)
    worksheet.write_rich_string(
        'A3:B5',
        'To.\n',
        bold,header['companyname']
    )
    worksheet.write('C3',"Date", workbook.add_format({'valign': 'left', 'border':1}))
    invoice_date = datetime.strptime(invoice_list['invoice_date'], '%Y-%M-%d').strftime('%b %d,%Y')
    worksheet.write('D3', invoice_date, workbook.add_format({'valign': 'left', 'border':1}))
    worksheet.write('C4',"Your Ref: No", workbook.add_format({'valign': 'left', 'border':1}))
    worksheet.write('D4',"Order by {0}".format(header['customer']), workbook.add_format({'valign': 'left', 'border':1}))
    worksheet.write('C5',"Our Ref: No", workbook.add_format({'valign': 'left', 'border':1}))
    worksheet.write('D5',invoice_list['invNo'], workbook.add_format({'valign': 'left', 'border':1, 'bold': True }))

    worksheet.write('A6',"Att:", workbook.add_format({'valign': 'left', 'border':1}))
    worksheet.write('B6', header['companyname'], workbook.add_format({'valign': 'left', 'border':1, 'bold': True}))
    worksheet.write('C6',"Subject", workbook.add_format({'valign': 'left', 'border':1}))
    worksheet.write('D6', payment['subject'], workbook.add_format({'valign': 'left', 'border':1}))


    d = {"<h3>Dear Sir,</h3>":"","<p>":"","</p>":""}
    head_text = replace_all(invoice_list['header'], d)
    worksheet.merge_range('A8:E8',"Dear Sir,", workbook.add_format({'bold': True}))
    head_start = head_text.split("[#payment#]")[0]
    head_mid = "{0} Payment".format(invoice_list['paymenttype'].capitalize())
    head_end = head_text.split("[#payment#]")[1]
    worksheet.merge_range('A9:E9',None)
    worksheet.write_rich_string(
        'A9:E9',
        head_start,
        bold,head_mid,
        head_end
    )

    table_header_cell_format = workbook.add_format({'align': 'center','border':1})
    worksheet.write(10,0,"No.", table_header_cell_format)
    worksheet.write(10,1,"Description", table_header_cell_format)
    worksheet.write(10,2,"Qty.", table_header_cell_format)
    worksheet.write(10,3,"Price", table_header_cell_format)
    worksheet.write(10,4,"Amount", table_header_cell_format)

    table_data_cell_format = workbook.add_format({'align': 'center','border':1})
    row_id = 11
    table_col_num = 1
    total_units = 0
    total_amount = 0
    for obj in body_obj:
        u = "Units" if obj['unit'] > 1 else "Unit"
        amount = obj['unit'] * obj['price']
        worksheet.write(row_id,0,'{0}.'.format(table_col_num), table_data_cell_format)
        worksheet.write(row_id,1,obj['subject'], workbook.add_format({'align': 'left','border':1}))
        worksheet.write(row_id,2,"{0}{1}".format(obj['unit'], u), table_data_cell_format)
        worksheet.write(row_id,3,obj['price'], workbook.add_format({'align': 'right','border':1, 'num_format': '#,##0'}))
        worksheet.write(row_id,4, amount, workbook.add_format({'align': 'right','border':1, 'num_format': '#,##0'}))
        total_units += obj['unit']
        total_amount += amount
        row_id += 1
        table_col_num += 1

    #Summary Row
    u = "Units" if total_units > 1 else "Unit"
    worksheet.write(row_id,0, None, table_data_cell_format)
    worksheet.write(row_id,1,"Total in (Kyat)", workbook.add_format({'align': 'left','border':1}))
    if "discount" in payment and payment['discount'] > 0:
        worksheet.write(row_id,2, None, table_data_cell_format)
    elif "manual_discount" in payment and payment['manual_discount'] > 0:
        worksheet.write(row_id,2, None, table_data_cell_format)
    elif "tax" in payment and payment['tax'] > 0:
        worksheet.write(row_id,2, None, table_data_cell_format)
    else:
        worksheet.write(row_id,2, "{0}{1}".format(total_units, u), table_data_cell_format)
    worksheet.write(row_id,3, None, table_data_cell_format)
    worksheet.write(row_id,4, total_amount, workbook.add_format({'align': 'right','border':1, 'num_format': '#,##0'}))
    row_id += 1

    """ Discount """
    if "discount" in payment and payment['discount'] > 0:
        # worksheet.write_rich_string(
        # 'A9:E9',
        # 'Subject : ',
        # bold,unit['subject']
        # )
        discount = int(math.ceil((total_amount * payment['discount'] )/100.0))
        worksheet.write(row_id,0,None, table_data_cell_format)
        worksheet.write(row_id,1,"Discount ({0}%)".format(payment['discount']), workbook.add_format({'align': 'left','border':1}))
        worksheet.write(row_id,2,None, table_data_cell_format)
        worksheet.write(row_id,3,None, table_data_cell_format)
        worksheet.write(row_id,4,"(-){0}".format("{:,}".format(discount)), workbook.add_format({'align': 'right','border':1}))
        total_amount = total_amount - discount
        row_id += 1

        worksheet.write(row_id,0, None, table_data_cell_format)
        worksheet.write(row_id,1,"Total in (Kyat)", workbook.add_format({'align': 'left','border':1}))
        worksheet.write(row_id,2, None, table_data_cell_format)
        worksheet.write(row_id,3, None, table_data_cell_format)
        worksheet.write(row_id,4, total_amount, workbook.add_format({'align': 'right','border':1, 'num_format': '#,##0'}))
        row_id += 1
    """ Commercial Tax """
    if "tax" in payment and payment['tax'] > 0:
        commercial_tax = int(math.ceil((total_amount * payment['tax'] )/100.0))
        worksheet.write(row_id,0,None, table_data_cell_format)
        worksheet.write(row_id,1,"Commercial Tax ({0}%)".format(payment['tax']), workbook.add_format({'align': 'left','border':1}))
        worksheet.write(row_id,2,None, table_data_cell_format)
        worksheet.write(row_id,3,None, table_data_cell_format)
        worksheet.write(row_id,4,"(+){0}".format("{:,}".format(commercial_tax)), workbook.add_format({'align': 'right','border':1}))
        total_amount = total_amount + commercial_tax
        row_id += 1

        worksheet.write(row_id,0, None, table_data_cell_format)
        worksheet.write(row_id,1,"Total in (Kyat)", workbook.add_format({'align': 'left','border':1}))
        worksheet.write(row_id,2, None, table_data_cell_format)
        worksheet.write(row_id,3, None, table_data_cell_format)
        worksheet.write(row_id,4, total_amount, workbook.add_format({'align': 'right','border':1, 'num_format': '#,##0'}))
        row_id += 1

    """ Special Discount"""
    if "manual_discount" in payment and payment['manual_discount'] > 0:
        manual_discount = payment['manual_discount']
        worksheet.write(row_id,0,None, table_data_cell_format)
        worksheet.write(row_id,1,"Special Discount", workbook.add_format({'align': 'left','border':1}))
        worksheet.write(row_id,2,None, table_data_cell_format)
        worksheet.write(row_id,3,None, table_data_cell_format)
        worksheet.write(row_id,4,"(-){0}".format("{:,}".format(manual_discount)), workbook.add_format({'align': 'right','border':1, 'num_format': '#,##0'}))
        total_amount = total_amount - manual_discount
        row_id += 1

        worksheet.write(row_id,0, None, table_data_cell_format)
        worksheet.write(row_id,1,"Total in (Kyat)", workbook.add_format({'align': 'left','border':1}))
        worksheet.write(row_id,2, "{0}{1}".format(total_units, u), table_data_cell_format)
        worksheet.write(row_id,3, None, table_data_cell_format)
        worksheet.write(row_id,4, total_amount, workbook.add_format({'align': 'right','border':1, 'num_format': '#,##0'}))
        row_id += 1

    # #Summary Row
    # u = "Units" if total_units > 1 else "Unit"
    # worksheet.write(row_id,0, None, table_data_cell_format)
    # worksheet.write(row_id,1,"Total in (Kyat)", workbook.add_format({'align': 'left','border':1}))
    # worksheet.write(row_id,2, "{0}{1}".format(total_units, u), table_data_cell_format)
    # worksheet.write(row_id,3, None, table_data_cell_format)
    # worksheet.write(row_id,4, total_amount, workbook.add_format({'align': 'right','border':1, 'num_format': '#,##0'}))
    # row_id += 1

    """first Payment for Total in (Kyat) (34%)"""
    worksheet.write(row_id,0, None, table_data_cell_format)
    worksheet.write(row_id,1,"{0} Payment for Total in (Kyat) ({1}%)".format(invoice_list['paymenttype'].capitalize(),str(invoice_list['payments'])), workbook.add_format({'align': 'left','border':1}))
    worksheet.write(row_id,2, None, table_data_cell_format)
    worksheet.write(row_id,3, None, table_data_cell_format)
    first_payment = int(math.ceil((total_amount*invoice_list['payments'])/100.0))
    worksheet.write(row_id,4, first_payment, workbook.add_format({'align': 'right','border':1, 'num_format': '#,##0'}))
    row_id += 1

    col_row_num = rownumber_to_columnstring(row_id+1)
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num),"(Kyat: {0})".format(p.number_to_words(first_payment)), workbook.add_format({'align': 'left','bold':"True"}))
    col_row_num += 1
    d = {"<p>":"","</p>":""}
    footer_text = replace_all(payment['footer'], d)
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num),footer_text, workbook.add_format({'align': 'left'}))
    col_row_num += 1
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num),"With Regard,", workbook.add_format({'align': 'left'}))
    col_row_num += 4
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num),"-------------------------------", workbook.add_format({'align': 'left'}))
    col_row_num += 1

    signatureid = payment['signatureid']
    signature_id_data = [x for x in signature_obj if x["id"] == signatureid != None][0]
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num),"{0} ({1})".format(signature_id_data['name'], signature_id_data['phoneno']), workbook.add_format({'align': 'left'}))
    col_row_num += 1
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num),signature_id_data['designation'], workbook.add_format({'align': 'left'}))
    col_row_num += 1
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num),signature_id_data['companyname'], workbook.add_format({'align': 'left'}))
    col_row_num += 1
    worksheet.merge_range('A{}:E{}'.format(col_row_num, col_row_num),header['createdby'], workbook.add_format({'align': 'left'}))
    col_row_num += 1

    row_id = columnstring_to_rownumber(col_row_num)
    loop_count = 1
    for panel in body_obj:
        #table_body_data = [x for x in body_obj if x["id"] == panel["id"] != None][0]
        table_body_data = panel
        row_id += 1
        col_row_num = rownumber_to_columnstring(row_id+1)
        panel_title_format = workbook.add_format({'bold': True,'border': 1, 'align':'left','underline':True})
        worksheet.merge_range('A{}:E{}'.format(col_row_num,col_row_num),None)
        worksheet.write_rich_string(
            'A{}:E{}'.format(col_row_num,col_row_num),
            bold,'{0}.'.format(str(loop_count)),
            panel_title_format,table_body_data['subject']
        )
        row_id = columnstring_to_rownumber(col_row_num)
        loop_count += 1

        row_id += 1
        """sheet column name level start"""
        worksheet.write(row_id,0,'No.', workbook.add_format({'bold': True,'border': 1, 'align':'center'}))
        worksheet.write(row_id,1,'Description', workbook.add_format({'bold': True,'border': 1, 'align':'center'}))
        worksheet.write(row_id,2,'Qty.', workbook.add_format({'bold': True,'border': 1,'align':'center'}))
        worksheet.write(row_id,3,'Price', workbook.add_format({'bold': True,'border': 1,'align':'center'}))
        worksheet.write(row_id,4,'Amount', workbook.add_format({'bold': True,'border': 1,'align':'center'}))
        row_id += 1

        num_column = 0
        number_col_bold = workbook.add_format({'border': 1})
        number_col_bold.set_align('center')
        number_col_bold.set_align('vcenter')
        cell_number_format = workbook.add_format({'border': 1,'num_format': '#,##0'})
        cell_number_format.set_align('center')
        cell_number_format.set_align('vcenter')
        """Enclosure Start"""
        if "enclosure" in table_body_data and table_body_data['enclosure']['totalwithp'] > 0:
            unit = 1
            row_data = table_body_data['enclosure']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column + 1,'.'), number_col_bold)

            """ Panel Enclosure Description """
            PED = 'Size : {0}mmH x {1}mmW x {2}mmD\nType : {3}\nMaterial : {4}\nPaint : {5}\nColour : {6}\nStandard : {7}'.format(
                row_data['height'],row_data['width'],row_data['depth']['name'],row_data['type'],
                row_data['material'],row_data['paint'],row_data['colour'],row_data['standard']
                )
            bold = workbook.add_format({'bold': True})
            col_row_num = rownumber_to_columnstring(row_id)
            worksheet.write_rich_string(
                'B{}'.format(col_row_num,col_row_num),
                ' ',
                bold,"Panel Enclosure :\n",
                PED
            )
            row_id = columnstring_to_rownumber(col_row_num)                
            #worksheet.write(row_id,1, PED,workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "Unit"),  number_col_bold)
            worksheet.write(row_id,3, row_data['totalwithp'],cell_number_format)
            worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1
        """Enclosure End""" 

        if "busbarfab" in table_body_data and table_body_data['busbarfab']['totalwithp'] > 0:
            unit = 1
            row_data = table_body_data['busbarfab']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column + 1,'.'), number_col_bold)
            worksheet.write(row_id,1, row_data['label'], workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "No"), workbook.add_format({'border': 1,'align':'center'}))
            worksheet.write(row_id,3, row_data['totalwithp'], cell_number_format)
            worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1

        """BusBar Start"""
        if "busbar" in table_body_data and int(float(table_body_data['busbar']['totalwithp'])) > 0:
            unit = 1
            row_data = table_body_data['busbar']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column + 1,'A.'), number_col_bold)

            if "items" in row_data:
                cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                BBD = '{0}\n'.format(row_data['label'])
                for item in row_data["items"]:
                    UOM = ' {0}s'.format(item['uom']) if item['uom'] == 'No' and int(round(float(item['qty']))) > 1 else ' {0}'.format(item['uom'])
                    BBD = BBD  + ' -  ' + item['name'] + ' --- ' + item['qty'] + UOM + ' x ' + '{:1,}'.format(int(round(float(item['price'])))) +'\n'
            
            bold = workbook.add_format({'bold': True})
            col_row_num = rownumber_to_columnstring(row_id)
            worksheet.write_rich_string(
                'B{}'.format(col_row_num,col_row_num),
                ' ',
                bold,"Panel Enclosure :\n",
                PED
            )
            row_id = columnstring_to_rownumber(col_row_num)              
            #worksheet.write(row_id,1, BBD[:-1],workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "Lot"),  workbook.add_format({'border': 1,'align':'center'}))
            worksheet.write(row_id,3, row_data['totalwithp'],cell_number_format)
            worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1

        if "cableandlug" in table_body_data and  int(float(table_body_data['cableandlug']['totalwithp'])) > 0:
            unit = 1
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = table_body_data['cableandlug']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column,'B.'), number_col_bold)
            worksheet.write(row_id,1, row_data['label'], workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "Lot"),  workbook.add_format({'border': 1,'align':'center'}))
            worksheet.write(row_id,3, row_data['totalwithp'], cell_number_format)
            worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1    
        
        if "insulator" in table_body_data and int(float(table_body_data['insulator']['totalwithp'])) > 0:
            unit = 1
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = table_body_data['insulator']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column + 1,'.'), number_col_bold)
            """ Insulator Description """
            if "items" in row_data:
                cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                ITD = '{0}\n'.format(row_data['label'])
                for item in row_data["items"]:
                    UOM = ' {0}s'.format(item['uom']) if item['uom'] == 'No' and int(round(float(item['qty']))) > 1 else ' {0}'.format(item['uom'])
                    ITD = ITD  + ' -  ' + item['name'] + ' --- ' + str(item['qty']) + UOM + ' x ' + '{:1,}'.format(int(round(float(item['price'])))) +'\n'
            worksheet.write(row_id,1, ITD[:-1],workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "Lot"),  workbook.add_format({'border': 1,'align':'center'}))
            worksheet.write(row_id,3, row_data['totalwithp'],cell_number_format)
            worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1

        if "metering" in table_body_data and int(float(table_body_data['metering']['totalwithp'])) > 0:
            unit = 1
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = table_body_data['metering']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column + 1,'.'), number_col_bold)

            """ Metering Description """
            if "items" in row_data:
                cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                MTD = '{0}\n'.format(row_data['label'])
                for item in row_data["items"]:
                    UOM = ' {0}s'.format(item['uom']) if item['uom'] == 'No' and int(round(float(item['qty']))) > 1 else ' {0}'.format(item['uom'])
                    MTD = MTD  + ' -  ' + item['name'] + ' --- ' + str(item['qty']) + UOM + ' x ' + str(item['price']) +'\n'
            worksheet.write(row_id,1, MTD[:-1],workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "Lot"),  workbook.add_format({'border': 1,'align':'center'}))
            worksheet.write(row_id,3, row_data['totalwithp'],cell_number_format)
            worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1   

        if "protection" in table_body_data and int(float(table_body_data['protection']['totalwithp'])) > 0:
            unit = 1
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = table_body_data['protection']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column + 1,'.'), number_col_bold)

            """ Protection & Controls Description """
            if "items" in row_data and len(row_data["items"]) > 0:
                cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                PCD = '{0}\n'.format(row_data['label'])
                for item in row_data["items"]:
                    PCD = PCD  + ' -  ' + item['name'] + ' --- ' + str(item['qty']) + ' Set x ' + str(item['price']) +'\n'
            worksheet.write(row_id,1, PCD[:-1],workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "Lot"),  workbook.add_format({'border': 1,'align':'center'}))
            worksheet.write(row_id,3, row_data['totalwithp'],cell_number_format)
            worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1 

        if "pricelist" in table_body_data and int(float(table_body_data['pricelist']['totalwithp'])) > 0:
            print("pricelist")
            unit = 1
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = table_body_data['pricelist']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column + 1,'.'), number_col_bold)

            """ Material List Description """
            if "items" in row_data:
                cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                MLD = '{0}\n'.format(row_data['label'])
                for item in row_data["items"]:
                    UOM = ' {0}s'.format('No') if int(round(float(item['qty']))) > 1 else ' {0}'.format('No')
                    MLD = MLD  + ' -  ' + item['brandname'] + ' - ' + item['description'] + ' --- ' + str(item['qty']) + UOM + ' x ' + str(item['price']) +'\n'
            worksheet.write(row_id,1, MLD[:-1],workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "Lot"),  workbook.add_format({'border': 1,'align':'center'}))
            if row_data['customersupply'] == True:
                worksheet.write(row_id,3, None,cell_number_format)
                worksheet.write(row_id,4, "Customer Supply", workbook.add_format({'border': 1,'align':'right'}))
            else:
                worksheet.write(row_id,3, row_data['totalwithp'],cell_number_format)
                worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1

        if "extrapricelist" in table_body_data and int(float(table_body_data['extrapricelist']['totalwithp'])) > 0:
            unit = 1
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = table_body_data['extrapricelist']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column + 1,'.'), number_col_bold)

            """Extra Material List Description """
            if "items" in row_data:
                cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                EMLD = ''
                for item in row_data["items"]:
                    UOM = ' {0}s'.format('No') if int(round(float(item['qty']))) > 1 else ' {0}'.format('No')
                    EMLD = EMLD  + ' -  ' + item['brandname'] + '  -- ' + item['description'] + ' --- ' + str(item['qty']) + UOM +'\n'
            worksheet.write(row_id,1, EMLD[:-1],workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "Lot"),  workbook.add_format({'border': 1,'align':'center'}))
            if row_data['customersupply'] == True:
                worksheet.write(row_id,3, None,cell_number_format)
                worksheet.write(row_id,4, "Customer Supply", workbook.add_format({'border': 1,'align':'right'}))
            else:
                worksheet.write(row_id,3, row_data['totalwithp'],cell_number_format)
                worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1

        if "extraonepricelist" in table_body_data and int(float(table_body_data['extraonepricelist']['totalwithp'])) > 0:
            print("extraonepricelist")
            unit = 1
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = table_body_data['extraonepricelist']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column + 1,'.'), number_col_bold)

            """Extra Material List Description """
            if "items" in row_data:
                cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                EMLD = ''
                for item in row_data["items"]:
                    UOM = ' {0}s'.format('No') if int(round(float(item['qty']))) > 1 else ' {0}'.format('No')
                    EMLD = EMLD  + ' -  ' + item['brandname'] + '  -- ' + item['description'] + ' --- ' + str(item['qty']) + UOM +'\n'
            worksheet.write(row_id,1, EMLD[:-1],workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "Lot"),  workbook.add_format({'border': 1,'align':'center'}))
            if row_data['customersupply'] == True:
                worksheet.write(row_id,3, None,cell_number_format)
                worksheet.write(row_id,4, "Customer Supply", workbook.add_format({'border': 1,'align':'right'}))
            else:
                worksheet.write(row_id,3, row_data['totalwithp'],cell_number_format)
                worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1                 
        
        if "extratwopricelist" in table_body_data and int(float(table_body_data['extratwopricelist']['totalwithp'])) > 0:
            unit = 1
            cell_format = workbook.add_format({'bold':True,'border': 2})
            row_data = table_body_data['extratwopricelist']
            worksheet.write(row_id,0, '{0}{1}'.format(num_column + 1,'.'), number_col_bold)

            """Extra Material List Description """
            if "items" in row_data:
                cell_format = workbook.add_format({'text_wrap':1,'valign':'top', 'border': 2})
                EMLD = ''
                for item in row_data["items"]:
                    UOM = ' {0}s'.format('No') if int(round(float(item['qty']))) > 1 else ' {0}'.format('No')
                    EMLD = EMLD  + ' -  ' + item['brandname'] + '  -- ' + item['description'] + ' --- ' + str(item['qty']) + UOM +'\n'
            worksheet.write(row_id,1, EMLD[:-1],workbook.add_format({'border': 1}))
            worksheet.write(row_id,2, "{0}{1}".format(unit, "Lot"),  workbook.add_format({'border': 1,'align':'center'}))
            if row_data['customersupply'] == True:
                worksheet.write(row_id,3, None,cell_number_format)
                worksheet.write(row_id,4, "Customer Supply", workbook.add_format({'border': 1,'align':'right'}))
            else:
                worksheet.write(row_id,3, row_data['totalwithp'],cell_number_format)
                worksheet.write(row_id,4, (unit * row_data['totalwithp']), cell_number_format)
            num_column = num_column + 1
            row_id = row_id + 1


        cell_format = workbook.add_format({'bold':True,'border': 1})
        worksheet.write(row_id,0,None, workbook.add_format({'border': 1}))
        worksheet.write(row_id,1,"Total in (Kyat)", cell_format)
        worksheet.write(row_id,2,None, workbook.add_format({'border': 1}))
        worksheet.write(row_id,3,None, workbook.add_format({'border': 1}))
        total_in_kyat = 0
        for item in table_body_data:
            if isinstance(table_body_data[item], dict):
                #print("ITEM",type(json_data[item]))
                if 'customersupply' in  table_body_data[item]:
                    if table_body_data[item]['customersupply'] == False:
                        total_in_kyat += int(float(table_body_data[item]['totalwithp']))
                else:
                    total_in_kyat += int(float(table_body_data[item]['totalwithp']))
        worksheet.write(row_id,4, total_in_kyat, cell_number_format)

        row_id += 2
        # worksheet.write(row_id,0,table_col_num, table_data_cell_format)
        



    """END Data Table"""




    worksheet.set_column(0, 0, 10)
    worksheet.set_column(1, 1, width=70)
    worksheet.set_column(2, 2, width=20)
    worksheet.set_column(3, 3, width=20)
    worksheet.set_column(4, 4, width=20)
    worksheet.set_column(5, 5, width=20)
    workbook.close()
    file_path = '{0}.xlsx'.format(filename)
    return file_path