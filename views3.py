from django.shortcuts import render
from django.http import HttpResponse
from django.http import FileResponse
from os.path import abspath
from django.shortcuts import render
from rest_framework.response import Response
from rest_framework import generics
import openpyxl
from .models import Spare_parts as sp

# Create your views here.
def home_page(request):
    "home page"
    print(request.build_absolute_uri())
    if request.method == "POST":
        return render(request, 'home.html', {"new_item_text": request.POST["item_text"]})
    return render(request, "home.html")

def get_file(request):
    response = FileResponse(open(abspath("SW_Repin_mart_2023.xlsx"), "rb"))
    return response

def load_excel_file_exe(request):
    with open (abspath("Load_Excel_File.exe"), "rb")as my_application:
        response = HttpResponse(my_application.read(), headers = {
            "Content-Type": "application/vnd.microsoft.portable-executable",
            "Content-Disposition": "attachment; filename = Load_Excel_File.exe"})
        return response

def load_excel_file(request):
    doc=openpyxl.load_workbook("second_example.xlsx")
    sheets = doc.get_sheet_names()
    sheet = doc[sheets[0]]
    n = 1i
    items = sp.objects.all()
    for i in items:
        if i["count"]<i["min"]:
            sheet["A"+str(n+10)]=str(n)
            sheet["B"+str(n+10)]=i.get("title")
            sheet["D"+str(n+10)]="Запасные части,  предназначается для службы главного механика, инициатором закупки является служба главного механика.\n"+i.get("title")
            sheet["E"+str(n+10)]=i.get("brand")
            sheet["F"+str(n+10)]="No/Нет"
            sheet["H"+str(n+10)]=i.get("unit")
            sheet["I"+str(n+10)]=i.get("count")
            sheet["K"+str(n+10)]="M&U (SW)"
            sheet["L"+str(n+10)]="Spare Parts and Service / Запасные части и сервис"
            sheet["M"+str(n+10)]=str(i.get("MABP"))+i.get("currency")
            sheet["N"+str(n+10)]="""Инициатор закупки Володи С.В.
Решение о закупки Володи С.В.
Согласование закупки Онуфриев С.Ю. 
Для быстрого ремонта оборудования в цехе сварки, в случае отказа в работе основного устройства, приобретается в рамках списка ключевых запасных частей.
На складе 0, на линии 2

Procurement initiator Volodya S.V.
The decision to purchase Volodya S.V.
Procurement approval Onufriev S.Yu.
For quick repair of equipment in the welding shop, in case of failure of the main device, it is purchased as part of the use of spare parts.
In warehouse 0, on line 2"""
            sheet["M"+str(n+10)]=str(i.get("MABP")*(i.get("min")-i.get("count")))+i.get("currency")
    doc.save("newfile.xlsx")
    print("End");
    with open(abspath("newfile.xlsx"), "rb") as file:
        my_data = file.read()
        response = HttpResponse(my_data, headers = {
            "Content-Type": "application/vnd.ms-excel",
            "Content-Disposition": "attachment; filename = 19_04_2023.xlsx"})
        return response
















    
    
def load_apk_file(request):
    with open(abspath("game.apk"), "rb") as file:
        data = file.read()
        response = HttpResponse(data, headers = {
            "Content-Type": "application/vnd.android.package-archive",
            "Content-Disposition": "attachment; filename = game.apk"})
        return response
    
def load_image(request):
    image_id = request.GET["image"]
    print(image_id)
    with open(abspath("lists/images/"+image_id+".png"), "rb") as image:
        data = image.read()
        response = HttpResponse(data, headers = {
            "Content-Type": "image/png",
            "Content-Disposition": "attachment; filename = "+image_id+".png"})
        return response
    

def show_image(request):
    image_id = request.GET["image"]
    response = FileResponse(open(abspath("lists/images/"+image_id+".png"), "rb"))
    return response
    
class Items_list(generics.ListAPIView):
    def get(self, request):
        return Response(sp.get_all())
    


class One_item(generics.GenericAPIView):
    def get(self, request):
        id = request.GET["id"]
        item = sp.objects.filter(id = id)[0]
        return Response(item.get_full_info())
    


def change_items_count(request):
    if request.method=="GET":
        item_id = request.GET["id"]
        count_difference = request.GET["difference"]
        print(item_id, count_difference)
        print(sp.get_all())
        return HttpResponse("Yes!")

    