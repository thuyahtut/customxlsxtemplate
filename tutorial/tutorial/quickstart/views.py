from django.contrib.auth.models import User, Group
from rest_framework import viewsets, generics
from rest_framework.permissions import IsAdminUser
from tutorial.quickstart.serializers import UserSerializer, GroupSerializer
from rest_framework.response import Response
from pandas.io.json import json_normalize
from datetime import datetime
import xlsxwriter
import os
from django.conf import settings
from django.http import HttpResponse, Http404
from .xlsx_template import header, custom_xlsx_export
import json
class UserViewSet(viewsets.ModelViewSet):
    """
    API endpoint that allows users to be viewed or edited.
    """
    queryset = User.objects.all().order_by('-date_joined')
    serializer_class = UserSerializer


class GroupViewSet(viewsets.ModelViewSet):
    """
    API endpoint that allows groups to be viewed or edited.
    """
    queryset = Group.objects.all()
    serializer_class = GroupSerializer

from rest_framework.views import APIView, Response
 
class CustomView1(APIView):
    def get(self, request, format=None):
        json_data = [
            {
                "id":448,
                "jobno":"IES-1203-R0",
                "designation":"-",
                "customer":"U Soe Kyaw Thu Hlaing",
                "email":"soekyawthuhlaing@iemmyanmar.com",
                "phoneno":"09-254024617",
                "address":"-",
                "companyname":"I.E.M Company Limited",
                "created":"2019-03-30",
                "qdate":"2019-03-30",
                "createdby":"ZPW",
                "updatedby":"admin",
                "ccemail":"winhlaingoo@iemmyanmar.com, hanaung@iemmyanmar.com, hanlynn@iemmyanmar.com"
            }
        ]
        file_path = header(json_data)
        print('file_path',file_path)
        with open(file_path, 'r') as fh:
            response = HttpResponse(fh.read(), content_type="application/vnd.ms-excel")
            response['Content-Disposition'] = 'inline; filename=' + os.path.basename(file_path)
            return response
        raise Http404

class CustomView2(APIView):
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

    def post(self, request, format=None):
        header = [
            {
                "id":448,
                "jobno":"IES-1203-R0",
                "designation":"-",
                "customer":"U Soe Kyaw Thu Hlaing",
                "email":"soekyawthuhlaing@iemmyanmar.com",
                "phoneno":"09-254024617",
                "address":"-",
                "companyname":"I.E.M Company Limited",
                "created":"2019-03-30",
                "qdate":"2019-03-30",
                "createdby":"ZPW",
                "updatedby":"admin",
                "ccemail":"winhlaingoo@iemmyanmar.com, hanaung@iemmyanmar.com, hanlynn@iemmyanmar.com"
            }
        ]
        json_data = header[0]
        pd_df = json_normalize(json_data)
        today = datetime.today().strftime('%d-%B-%y')
        workbook   = xlsxwriter.Workbook('filename.xlsx')
        worksheet1 = workbook.add_worksheet()
        worksheet1.write(0,0,'Job No. :{0}'.format(json_data['jobno']))
        # worksheet1.set_column(0,0,20)
        worksheet1.write(0,1,'Customer :{0}'.format(json_data['customer']))
        worksheet1.write(0,2,'Date :{0}'.format(today))
        worksheet1.write(1,0,'Subject :150kvar Capacitor Bank')
        worksheet1.write(1,2,'Qty :1 Unit')
        workbook.close()
        file_path = '/home/thuya/Downloads/kopo/tutorial/filename.xlsx'
        with open('/home/thuya/Downloads/kopo/tutorial/filename.xlsx', 'r') as fh:
            response = HttpResponse(fh.read(), content_type="application/vnd.ms-excel")
            response['Content-Disposition'] = 'inline; filename=' + os.path.basename(file_path)
            return response
        raise Http404      