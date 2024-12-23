import ctypes

import mapsyst
import maptype
import mapapi
import rscapi
import seekapi
import logapi
import maperr
import mapselec
import doforeach
import tkinter
import sitapi
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog

def SemText(hmap: maptype.HMAP,hobj: maptype.HOBJ) -> str:
    result = 0

    if hmap == 0:
        print("Ошибка: Нет карты.")
        sys.exit()  # Прекращаем выполнение скрипта
    
    hobj = mapapi.mapCreateObject(hmap)
    if hobj == 0:
        print("Ошибка: Не удалось создать объект hobj.")
        sys.exit()  # Прекращаем выполнение скрипта
    
    # Инициализация списка выбранных объектов
    selected_objects = []
    
    # Запросить число семантических характеристик у объекта
    semamount = mapapi.mapSemanticAmount(hobj)
    
    # Запросить последовательный номер кода семантической харак (модифицировано)
    semNumb = mapapi.mapSemanticNumber(hobj, 801)
    
    # Запросить уникальный номер объекта
    objKey = mapapi.mapObjectKey(hobj)

    # сохранение XML документа как файл
    ftypes=[("XML files", "*.xml"), ('All files', '*.*')]
    initnameXML = str('тест объекта') + "_" + str (objKey) + str('.xml')
    filenameXML = filedialog.asksaveasfilename(filetypes=ftypes, title='Select file for points and semantics',defaultextension='.xml',initialfile=initnameXML)

    if len(filenameXML) == 0:
        result = 1  
    else:
        tfileXML = open(filenameXML, 'w', encoding='utf-8')
    if tfileXML is None:
        result = 1  
   
    # создание XML документа как строки с дополнительными атрибутами
    xml_doc = '<?xml version="1.0" encoding="UTF-8"?>\n'
    xml_doc += '<ArrayOfERSIntegrationXML xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">\n'

    flag = maptype.WO_FIRST
    while (seekapi.mapTotalSeekObject(hmap, hobj, flag) != 0):
        selected_objects.append(hobj)
    
        # Создаем новый XML элемент для объекта
        xml_doc += '  <ERSIntegrationXML>\n'
    
        # добавление координат объекта в XML-документ
        polycount = mapapi.mapPolyCount(hobj)  # Получаем количество полигонов
        pointstrings = []  # Список для хранения строк с координатами полигонов

        for i in range(0, polycount):  # Проходим по каждому полигону
            pointcount = mapapi.mapPointCount(hobj, i)  # Получаем количество точек в текущем полигоне
            pointstring = ""  # Инициализируем строку для хранения координат текущего полигона
            
            for j in range(1, pointcount + 1):  # Проходим по каждой точке в полигоне
                point = maptype.DOUBLEPOINT(0, 0)  # Создаем объект для хранения координат
                mapapi.mapGetPlanePoint(hobj, ctypes.byref(point), j, i)  # Получаем координаты точки
                # Добавляем координаты в строку pointstring с фиксированным количеством знаков после запятой
                pointstring += f"{point.X:.3f} {point.Y:.3f}, "
            
            # Удаляем последнюю запятую и пробел из pointstring
            if pointstring:  # Проверяем, что pointstring не пустой
                pointstring = pointstring.rstrip(", ")  # Удаляем лишние символы
            
            pointstrings.append(pointstring)  # Добавляем строку с координатами в список

        # Формируем строку POLYGON
        polygon_string = "POLYGON ((" + " ".join(pointstrings) + "))"  # Объединяем все строки в одну

        mapping = {
                "Вид пространственных данных": "subtype_ref",
                "Гриф секретности": "secretclass_ref",
                "Название, текст подписи": "name",
                "МСК": "referencesystem_ref",
                "Дата создания": "objectcreatedat",
                "Правообладатель": "rightholder_ref",
                "Производитель (автор)": "creator_ref",
                "Условия доступа": "accesscondition_ref",
                "Номенклатура": "nomenclature",
                "Масштаб": "scale",
                "Место хранения": "location_ref",
                "Формат хранения": "storageformat_ref",
                "Дата состояния местности": "areastatedate",
                polygon_string : "wgscoordinates"
        }
    
        for key in mapping:
            xml_doc += f'    <{mapping[key]}></{mapping[key]}>\n'
        
        # извлечение семантических характеристик
        semamount = mapapi.mapSemanticAmount(hobj)
        for i in range(semamount):
            semNumb = i + 1
            SemFullName = mapsyst.WTEXT(64)
            mapapi.mapSemanticFullName(hobj, semNumb, SemFullName, SemFullName.size())
            SemFullNametext = SemFullName.string()

            SemanticValue = mapsyst.WTEXT(64)
            mapapi.mapSemanticValuePro(hobj, semNumb, SemanticValue, SemanticValue.size(), 0)
            SemanticValuetext = SemanticValue.string()

            # сопоставление SemFullNametext с корректным значением в mapping 
            if SemFullNametext in mapping:
                if SemFullNametext == "МСК":
                    if SemanticValuetext == "-12":
                        SemanticValuetext = "12"
                xml_doc = xml_doc.replace(f'<{mapping[SemFullNametext]}></{mapping[SemFullNametext]}>\n', f'<{mapping[SemFullNametext]}>' + SemanticValuetext + f'</{mapping[SemFullNametext]}>\n')
            else:
                print(f"Warning: {SemFullNametext} not found in mapping dictionary")
            # добавление polygon_string в mapping
            xml_doc = xml_doc.replace(f'<{mapping[polygon_string]}></{mapping[polygon_string]}>\n', f'<{mapping[polygon_string]}>' + polygon_string + f'</{mapping[polygon_string]}>\n')  
        
        xml_doc += '  </ERSIntegrationXML>\n'       
        flag = maptype.WO_NEXT
    
    # закрытие элемента ArrayOfERSIntegrationXML
    xml_doc += '</ArrayOfERSIntegrationXML>\n'

    # запись XML документа в файл
    tfileXML.write(xml_doc)

    # закрываем файл
    tfileXML.close()
    
    # Создание окна Tkinter
    root = tk.Tk()
    root.withdraw()  # Скрытие основного окна
    
    # Вызов функции SemText
    #result = SemText(hmap, hobj)  

    # Показ сообщения "готово", если функция SemText завершилась успешно
    if result == 0:
        messagebox.showinfo("Готово",initnameXML +  " успешно сохранен")
    else:
        messagebox.showerror("Ошибка", "Произошла ошибка при выполнении SemText")

    # Закрытие окна Tkinter
    root.quit()
        
    if hobj != 0:
       mapapi.mapFreeObject(hobj) 
       
    return 0