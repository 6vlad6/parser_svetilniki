# -*- coding: cp1251 -*-

import openpyxl
import json


workbook = openpyxl.load_workbook("result/gcsvt/products_gcsvt.xlsx")
sheet = workbook.active

with open("result/gcsvt/data.json", 'r', encoding="utf-8") as file:
    products_str = file.read()
    products = json.loads(products_str)["products"]
    keys_places = {
        '���': 21,
        '������� ����, ��': 23,
        '������ �������������': 25,
        '��������� ��������� ������': 27,
        '������': 29,
        '��������': 31,
        '������ ������ �����������': 33,
        '�������� ����������� ����������': 35,
        '�����': 37,
        '����� ����� �� �����': 39,
        '�������� � ���� �������': 41,
        '��������� ��������� ������, %': 43,
        '�������': 45,
        'IP': 47,
        '���� ������ �����': 49,
        '����� ��������': 51,
        '��� �����': 53,
        '��� ��������� ���������': 55,
        '����������� ��������': 57,
        '�������� �������': 59,
        '���': 61,
        '����� ������ �� ��������� ��. �����': 63,
        '�������������, ��/��': 65,
        '���� �������': 67,
        '���� ������ �����������': 69,
        '������ ������������� � �������� �����������': 71,
        '����������': 73,
        '�������������� �����': 75,
        '��������� ������': 77,
        '������� ����������� ���. ����� ': 79,
        '�������� �������������� ��������� ����, ��': 81,
        '��������� ������/������������': 83,
        '������� ��� ����������': 85,
        '������������� ����������': 87,
        '���������� �������': 89,
        '������������������, �3/���': 91,
        '����������� �������': 93,
        '������� ��������� �������': 95,
        '���������': 97,
        '������� ����������� ���. �����': 99,
        '������������': 101,
    }
    i = 2
    codes = []
    for product in products:
        try:
            product_code = str(product["��� ������"])
            if product_code in codes:
                continue
            else:
                codes.append(product_code)
                sheet[f'A{i}'] = product_code
        except:
            pass


        separate_keys = []
        for key in product.keys():

            if key == "��������":
                sheet[f'B{i}'] = str(product[key])

            elif key == "��������":
                sheet[f'D{i}'] = str(product[key])

            elif key == "���������":
                sheet[f'P{i}'] = str(product[key])

            elif key == "��������, ��":
                sheet[f'F{i}'] = float(product[key])

            elif key == "�������� �����, ��" or key == "�����, ��":
                sheet[f'H{i}'] = float(product[key])

            elif "�������� �/�/�" in str(key):
                try:
                    try:
                        gabs = str(product[key]).split("x")

                        sheet[f'I{i}'] = float(gabs[1])
                        sheet[f'J{i}'] = float(gabs[2])
                        sheet[f'K{i}'] = float(gabs[0])
                    except:
                        gabs = str(product[key]).split("�")

                        sheet[f'I{i}'] = float(gabs[1])
                        sheet[f'J{i}'] = float(gabs[2])
                        sheet[f'K{i}'] = float(gabs[0])
                except:
                    pass

            elif "�����" in str(key):
                sheet[f'L{i}'] = str(product[key])

            elif "����".lower() in key.lower() or key == "�������":
                pass

            else:
                separate_keys.append(key)
        try:
            separate_keys["������������"] = separate_keys.pop("������������")
        except:
            pass

        if "������������" in separate_keys:
            options = product["������������"]['������������']
            for option in options:
                if option['�������� �����'] in separate_keys:
                    separate_keys.remove(option['�������� �����'])

        for key in separate_keys:

            if key == "������������":
                options = product[key]['������������']

                for option in options:
                    option_name = option["�������� �����"]

                    if str(option["�������� �����"]).strip() == '�������� ����������� � ������ �������������':
                        option_name = '������ ������������� � �������� �����������'

                    elif str(option["�������� �����"]).strip() == '������������':
                        option_name = '��������� ������/������������'

                    if option_name == "��������� ������":
                        cell_name = sheet.cell(row=i, column=103, value=option["�������� �����"])
                        cell_value = sheet.cell(row=i, column=104,
                                                value="; ".join(option["�����"]))

                    elif option_name == "���������� ��������":
                        cell_name = sheet.cell(row=i, column=105, value=option["�������� �����"])
                        cell_value = sheet.cell(row=i, column=106,
                                                value="; ".join(option["�����"]))

                    elif option_name == "�����":
                        cell_name = sheet.cell(row=i, column=107, value=option["�������� �����"])
                        cell_value = sheet.cell(row=i, column=108,
                                                value="; ".join(option["�����"]))

                    elif option_name == "������� ������":
                        cell_name = sheet.cell(row=i, column=109, value=option["�������� �����"])
                        cell_value = sheet.cell(row=i, column=110,
                                                value="; ".join(option["�����"]))

                    else:
                        cell_name = sheet.cell(row=i, column=keys_places[option_name], value=option["�������� �����"])
                        cell_value = sheet.cell(row=i, column=keys_places[option_name]+1,
                                                value="; ".join(option["�����"]))

            else:
                if key == "��� ������":
                    continue
                else:
                    cell_name = sheet.cell(row=i, column=keys_places[key], value=key)
                    cell_value = sheet.cell(row=i, column=keys_places[key]+1, value=product[key])

        i += 1

    workbook.save("result/gcsvt/products_gcsvt.xlsx")

print(len(codes))
print(len(set(codes)))