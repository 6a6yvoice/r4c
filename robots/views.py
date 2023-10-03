from django.shortcuts import render
from .models import Robot
# import pandas as pd
from django.http import HttpResponse
from openpyxl import Workbook


# Create your views here.
# def export_data_to_excel(request):
#    objs = Robot.objects.all()
#    pd.DataFrame(objs).to_excel('output.xlsx')
#    return JsonResponse({
#        'status': 200
#    })

