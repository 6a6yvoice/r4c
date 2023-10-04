import os
from django.conf import settings
from django.utils.timezone import make_aware, datetime, timedelta
from django.core.exceptions import ValidationError
from django.forms.models import model_to_dict
from xlsxwriter import Workbook
from robots.models import Robot
from django.http import JsonResponse
from django.core import serializers
import json


def get_robot_list():
    robots = Robot.objects.all()
    data = serializers.serialize("json", robots)
    return JsonResponse(json.loads(data), safe=False)


def create_robot(request):
    try:
        request_data = json.loads(request.body.decode("utf-8"))
        robot_data, status = save_robot_if_clean_data(request_data)
        return JsonResponse(data=robot_data, status=status)
    except json.JSONDecodeError:
        return JsonResponse(data={"error": "Invalid JSON format"}, status=400)


def save_robot_if_clean_data(data: dict):
    try:
        data["serial"] = f"{data['model']}-{data['version']}"
        data["created"] = make_aware(datetime.now())
        new_robot = Robot(**data)
        new_robot.clean_fields()
        new_robot.save()
        return model_to_dict(new_robot), 201
    except ValidationError as exc:
        errors = {k: v[0].messages[0] for k, v in exc.error_dict.items()}
        return {"error": errors}, 400


def get_report_xlsx():
    reports_storage = os.path.join(settings.MEDIA_ROOT, "reports")
    if not os.path.exists(reports_storage):
        os.makedirs(reports_storage)
    clean_folder(path_to_folder=reports_storage)

    datetime_stamp = datetime.now().strftime("(%d_%m_%Y_%H-%M-%S)")
    path_to_report = os.path.join(reports_storage, f"report{datetime_stamp}.xlsx")

    reporting_period_from = datetime.now() - timedelta(days=7)
    robots = Robot.objects.filter(created__gt=reporting_period_from)
    workbook = Workbook(path_to_report)

    for model in robots.values_list("model", flat=True).distinct():

        worksheet = workbook.add_worksheet(name=model)
        cell_format = workbook.add_format()
        cell_format.set_border()

        row = 0
        col = 0

        worksheet.write(row, col, "Модель", cell_format)
        worksheet.write(row, col + 1, "Версия", cell_format)
        worksheet.write(row, col + 2, "Количество за неделю", cell_format)
        worksheet.autofit()
        row += 1

        data = []

        for item in robots.filter(model=model):
            quantity = robots.filter(model=item.model, version=item.version).count()
            row_content = (item.model, item.version, quantity)
            data.append(row_content)

        for model, version, quantity in sorted(set(data)):
            worksheet.write(row, col, model, cell_format)
            worksheet.write(row, col + 1, version, cell_format)
            worksheet.write(row, col + 2, quantity, cell_format)
            row += 1

    workbook.close()
    report_file = open(path_to_report, "rb")
    return report_file


def clean_folder(path_to_folder):
    for report in os.listdir(path_to_folder):
        os.remove(os.path.join(path_to_folder, report))