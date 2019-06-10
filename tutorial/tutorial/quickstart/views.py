import os
from django.http import HttpResponse, Http404
from .xlsx_template import custom_xlsx_export,invoice_xlsx_export,quotation_xlsx_export
import json
from rest_framework.views import APIView

class CustomView(APIView):
    def get(self, request, format=None):
        false = False
        true = True
        with open('/home/thuya/Downloads/kopo/quotation-header.json', 'r') as header_file:
            header_data=header_file.read()
        header_obj = json.loads(header_data)  
        with open('/home/thuya/Downloads/kopo/body2.json', 'r') as body_file:
            body_data=body_file.read()
        body_obj = json.loads(body_data)    
        file_path = custom_xlsx_export(header_obj,body_obj)
        print('file_path',file_path)
        with open(file_path, 'r') as fh:
            response = HttpResponse(fh.read(), content_type="application/vnd.ms-excel")
            response['Content-Disposition'] = 'inline; filename=' + os.path.basename(file_path)
            return response
        raise Http404        

class InvoiceView(APIView):
    def get(self, request, format=None):
        false = False
        true = True
        """HEADER"""
        with open('/home/thuya/OUTSOURCE/mpec/invoice/header.json', 'r') as header_file:
            header_data=header_file.read()

        """BODY"""
        header_obj = json.loads(header_data)  
        with open('/home/thuya/OUTSOURCE/mpec/invoice/body.json', 'r') as body_file:
            body_data=body_file.read()
        body_obj = json.loads(body_data)  

        """CURRENCY"""
        with open('/home/thuya/OUTSOURCE/mpec/invoice/currency.json', 'r') as body_file:
            currency_data=body_file.read()
        currency_obj = json.loads(currency_data)


        """PAYMENT"""
        with open('/home/thuya/OUTSOURCE/mpec/invoice/payment.json', 'r') as body_file:
            payment_data=body_file.read()
        payment_obj = json.loads(payment_data)


        """SIGNATURE"""
        with open('/home/thuya/OUTSOURCE/mpec/invoice/signature.json', 'r') as body_file:
            signature_data=body_file.read()
        signature_obj = json.loads(signature_data)  

        file_path = invoice_xlsx_export(header_obj,body_obj,currency_obj,payment_obj,signature_obj)
        print('file_path',file_path)
        with open(file_path, 'r') as fh:
            response = HttpResponse(fh.read(), content_type="application/vnd.ms-excel")
            response['Content-Disposition'] = 'inline; filename=' + os.path.basename(file_path)
            return response
        raise Http404   

class QuotationView(APIView):
    def get(self, request, format=None):
        false = False
        true = True
        with open('/home/thuya/OUTSOURCE/mpec/quotation/header.json', 'r') as header_file:
            header_data=header_file.read()
        header_obj = json.loads(header_data)
        with open('/home/thuya/OUTSOURCE/mpec/quotation/unit.json', 'r') as body_file:
            body_data=body_file.read()
        unit_obj = json.loads(body_data)   
        with open('/home/thuya/OUTSOURCE/mpec/quotation/body.json', 'r') as body_file:
            body_data=body_file.read()
        body_obj = json.loads(body_data) 
        file_path = quotation_xlsx_export(header_obj,unit_obj,body_obj)
        print('file_path',file_path)
        with open(file_path, 'r') as fh:
            response = HttpResponse(fh.read(), content_type="application/vnd.ms-excel")
            response['Content-Disposition'] = 'inline; filename=' + os.path.basename(file_path)
            return response
        raise Http404               
