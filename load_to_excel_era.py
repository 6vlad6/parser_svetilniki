# -*- coding: cp1251 -*-

import openpyxl
import json


workbook = openpyxl.load_workbook("result/eraworld/products_eraworld.xlsx")
sheet = workbook.active

with open("result/eraworld/data.json", 'r', encoding="utf-8") as file:
    products_str = file.read()
    products = json.loads(products_str)["products"]

    keys_places = {
        '��������': 21,
        '�������� �������': 23,
        '��������� �� ����������� �������': 25,
        '���� ������ ��� ���� ��������': 27,
        '������ ��������': 29,
        '�������� ����������, �': 31,
        '���� ������, ���': 33,
        '���������� �����������': 35,
        '����� ������ �� ��������� ������������� �����': 37,
        '�����': 39,
        '���������� �������': 41,
        '������ ������������': 43,
        '����������� ��������, PF': 45,
        '��������': 47,
        '������� ��������': 49,
        '������������': 51,
        '�������� �������� ���������� � ����, �': 53,
        '������� ������ �� ����������� ���������� �����, IP': 55,
        '��� �����������': 57,
        '����� �������������������': 59,
        '�������� ����������� ������, ��': 61,
        '������ �������������': 63,
        '������': 65,
        '��� ������������': 67,
        '����������� ��������� ��������� ������, %': 69,
        '������� ����������': 71,
        '������� ����, ��': 73,
        '������������� ����������� �������� � ���������': 75,
        '���� ������ �������': 77,
        '�������� �����������, �': 79,
        '���������� �����������, ��': 81,
        '������� ���������� � ������ �������': 83,
        '����� ������������������': 85,
        '�������� ������� ����������': 87,
        '��� ������': 89,
        '��� ������ ���� �����': 91,
        '�������� ������������': 93,
        '����': 95,
        '������� ��������� ��� � ��': 97,
        '�������� �����': 99,
        '����������, �': 101,
        '������������� ����������': 103,
        '����������� ������������': 105,
        '����������, ��/��': 107,
        '���� �������': 109
    }

    i = 2
    keys = []
    for product in products:
        first_separate_key_index = 21
        first_separate_value_index = 22

        separate_keys = []
        for key in product.keys():

            if key == "��� ������":
                sheet[f'A{i}'] = str(product[key])

            elif key == "��������":
                sheet[f'B{i}'] = str(product[key])

            elif key == "��������":
                sheet[f'D{i}'] = str(product[key]).strip()

            elif key == "���������":
                sheet[f'P{i}'] = str(product[key])

            elif key == "������������ ��������, ��":
                sheet[f'F{i}'] = str(product[key])

            elif key == "�������� �����, ��":
                sheet[f'H{i}'] = str(product[key])

            elif "��������" in str(key):
                try:
                    try:
                        gabs = str(product[key]).lower().split("x")

                        sheet[f'K{i}'] = str(gabs[0])
                        sheet[f'I{i}'] = str(gabs[1])
                        sheet[f'J{i}'] = str(gabs[2])
                    except:
                        gabs = str(product[key]).lower().split("�")

                        sheet[f'K{i}'] = str(gabs[0])
                        sheet[f'I{i}'] = str(gabs[1])
                        sheet[f'J{i}'] = str(gabs[2])
                except:
                    pass

            elif "�����" in str(key):
                sheet[f'L{i}'] = float(str(product[key]).replace(",", "."))


            elif "����".lower() in key.lower() or key == "�������":
                pass

            else:
                separate_keys.append(key)

        for key in separate_keys:
            title = key
            if key == "����������� ����, ���":
                title = '��������'

            cell_name = sheet.cell(row=i, column=keys_places[title], value=key)
            cell_value = sheet.cell(row=i, column=keys_places[title]+1, value=product[key])

        i += 1

    workbook.save("result/eraworld/products_eraworld.xlsx")
