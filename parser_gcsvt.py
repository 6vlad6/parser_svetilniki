from selenium.webdriver.common.by import By
import time

import os
import requests as r

import json

from datetime import datetime

from create_driver import create_driver


urls = [
    'https://gcsvt.ru/svetodiodnyie-svetilniki/bakteritsidnye-retsirkulyatory/',
    'https://gcsvt.ru/svetodiodnyie-svetilniki/naruzhnoe-osveshchenie/',
    'https://gcsvt.ru/svetodiodnyie-svetilniki/vnutrennee-osveshchenie/',
    'https://gcsvt.ru/svetodiodnyie-svetilniki/spetsialnoe-osveshchenie/',
]

data = {}
data["products"] = []

for url in urls:
    mistakes = 0
    print("Категория началась, ", datetime.now().strftime("%H:%M:%S"))
    try:
        driver = create_driver(url+"?SHOWALL_1=1")

        products = driver.find_elements(By.CLASS_NAME, "product-item-title")

        for product in products:

            context = {}  # словарь для данных

            driver = create_driver(product.find_element(By.TAG_NAME, "a").get_attribute("href"))
            time.sleep(1)

            try:
                btn_show_all_props = driver.find_element(By.ID, "base_props_value_show")
                btn_show_all_props.click()
                time.sleep(3)
            except:
                pass

            features_div = driver.find_element(By.CLASS_NAME, "product-item-detail-tab-content")

            # заголовки характеристик
            feature_titles = features_div.find_elements(By.CLASS_NAME, "product-item-detail-properties-name")
            features = []
            for title in feature_titles:
                features.append(title.text)

            # значения характеристик
            feature_values = features_div.find_elements(By.CLASS_NAME, "product-item-detail-properties-value")
            values = []
            for value in feature_values:
                values.append(value.text)

            # связывание заголовка и значения
            for i in range(len(features)):
                context.update({f"{features[i].replace(':', '')}": values[i]})
            context["Название"] = driver.find_element(By.ID, "pagetitle").get_attribute("innerHTML")
            context["Описание"] = driver.find_elements(By.CLASS_NAME, "col")[1].text.replace("&quot;", "'").replace("<br />", "")
            context["Категория"] = url.split("svetodiodnyie-svetilniki/")[1][:-1]

            stroke_ancor_items = driver.find_elements(By.CLASS_NAME, "stroke_ancor_item")
            if len(stroke_ancor_items) >= 5:
                configs = {
                    "Конфигурации": []
                }
                config_elems = driver.find_elements(By.CLASS_NAME, "config_body_group")

                for config_elem in config_elems:

                    config = {
                        "Название опции": "",
                        "Опции": []
                    }

                    config["Название опции"] = config_elem.find_element(By.CLASS_NAME, "config_body_group_name").get_attribute("innerHTML")
                    config_options = []
                    config_labels = config_elem.find_elements(By.TAG_NAME, "label")

                    for config_label in config_labels:
                        text = config_label.find_element(By.TAG_NAME, "span").get_attribute("innerHTML")
                        config_options.append(text)

                    config["Опции"] = config_options
                    configs["Конфигурации"].append(config)
                context["Конфигуратор"] = configs

            replacements = {"Артикул": "Код товара"}
            for i in context:
                if i in replacements:
                    context[i] = driver.find_elements(By.CLASS_NAME,
                        "product-item-detail-properties-value")[2].get_attribute("innerHTML").split("\t")[0]
                    context[replacements[i]] = context.pop(i)
                    break

            data["products"].append(context)

            if not context['Код товара'] in os.listdir("result/gcsvt/"):
                os.makedirs(f"result/gcsvt/{context['Код товара']}", exist_ok=True)

                images = driver.find_elements(By.CLASS_NAME, "product-item-detail-slider-image")
                i = 1
                for image in images:

                    response = r.get(image.find_element(By.TAG_NAME, "img").get_attribute("src"))
                    file = open(f"result/gcsvt/{context['Код товара']}/{i}.png", 'bw')
                    for chunk in response.iter_content(4096):
                        file.write(chunk)

                    i += 1
    except:
        mistakes += 1
    print("Категория обработана, ", datetime.now().strftime("%H:%M:%S"), mistakes)
with open('result/gcsvt/data.json', 'a', encoding="utf-8") as file:
    json.dump(data, file, ensure_ascii=False)
