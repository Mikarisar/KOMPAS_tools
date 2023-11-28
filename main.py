import pythoncom
from win32com.client import Dispatch, gencache


# Подключение к API7 программы Kompas 3D
def get_kompas_api7():
    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    api = module.IKompasAPIObject(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    return module, api, const


# Технические переменные
FileTypeSupported = False

# Подключение к нтерфейсу программы Kompas 3D
module7, api7, const7 = get_kompas_api7()   # Подключаемся к API7
app7 = api7.Application                     # Получаем основной интерфейс
app7.Visible = True                         # Показываем окно пользователю (если скрыто)
app7.HideMessage = const7.ksHideMessageNo   # Отвечаем НЕТ на любые вопросы программы
print("Версия КОМПАС:", app7.ApplicationName(FullName=True))  # Печатаем название программы

application = module7.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module7.IApplication.CLSID, pythoncom.IID_IDispatch))
Documents = application.Documents

print("Документов открыто:", application.Documents.Count)

if application.ActiveDocument is None:
    print("Документ не выбран!")
else:
    #  Получим активный документ
    kompas_document = application.ActiveDocument
    kompas_document_2d = module7.IKompasDocument2D(kompas_document)

    if kompas_document.Name == '':
        print("Активный документ не сохранён на диск!")
    else:
        print("Активный документ:", kompas_document.Name)
        print("Папка документа:", kompas_document.Path)

    # Узнаём тип документа
    if kompas_document.DocumentType == 1:  # Чертёж
        print("Тип документа: Чертёж")
        FileTypeSupported = True

        #  Количество листов
        print("Количество листов:", kompas_document.LayoutSheets.Count)

        # Количество видов
        print("Количество видов:", kompas_document_2d.ViewsAndLayersManager.Views.Count)

        # получаем активный вид
        active_view = kompas_document_2d.ViewsAndLayersManager.Views.ActiveView
        print("Активный вид:", active_view.Name)
        print("Масштаб вида:", active_view.Scale)

    elif kompas_document.DocumentType == 2:  # Фрагмент
        print("Тип документа: Фрагмент")

    elif kompas_document.DocumentType == 3:  # Спецификация
        print("Тип документа: Спецификация")

    elif kompas_document.DocumentType == 4:  # Деталь
        print("Тип документа: Деталь")

    elif kompas_document.DocumentType == 5:  # Сборка
        print("Тип документа: Сборка")

    elif kompas_document.DocumentType == 6:  # Текстовый документ
        print("Тип документа: Текстовый документ")

    else:
        print("Неизвестный тип документа:", kompas_document.DocumentType)

    if not FileTypeSupported:
        print("Этот тип документа не поддерживается!")

True

# input("\nДля завершения нажмите Enter")
