from django.shortcuts import render

# Create your views here.
from django.http import HttpResponseForbidden, FileResponse
from django.views.decorators.csrf import csrf_exempt
from django.contrib.auth.decorators import login_required
from robots.utils import get_robot_list, create_robot, get_report_xlsx


@csrf_exempt
def robot(request):
    if request.method == "GET":
        return get_robot_list()
    if request.method == "POST":
        return create_robot(request)


@csrf_exempt
@login_required
def report(request):
    if request.method == "GET" and request.user.username == "director":
        report_xslx = get_report_xlsx()
        return FileResponse(report_xslx)
    else:
        return HttpResponseForbidden()