import pythoncom
from win32com.client import Dispatch, gencache


class Kompas(object):
    """
    Класс для взаимодействия с КОМПАС 3D
    """

    def __init__(self):
        # Подключение к API7 программы Kompas 3D

        # Подключаемся к API7
        self.module7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        self.api7 = self.module7.IKompasAPIObject(
            Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(self.module7.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch))
        self.const7 = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants

        # Подключение к нтерфейсу программы Kompas 3D
        print("Подключение к КОМПАС...")
        self.app7 = self.api7.Application                     # Получаем основной интерфейс
        self.app7.Visible = True                              # Показываем окно пользователю (если скрыто)
        self.app7.HideMessage = self.const7.ksHideMessageNo   # Отвечаем НЕТ на любые вопросы программы

        self.application = self.module7.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(self.module7.IApplication.CLSID, pythoncom.IID_IDispatch))
        self.Documents = self.application.Documents

    def info_general(self):
        """
        Вывод информации о программе
        """
        print("######### ИНФОРМАЦИЯ О ПРОГРАММЕ #########")
        print("Версия КОМПАС:", self.app7.ApplicationName(FullName=True))  # Печатаем название программы
        print("Документов открыто:", self.application.Documents.Count)
        print("##########################################")

    def info_active(self):
        """
        Вывод информации об активном документе
        """

        print("#### ИНФОРМАЦИЯ ОБ АКТИВНОМ ДОКУМЕНТЕ ####")
        FileTypeSupported = False

        if self.application.ActiveDocument is None:
            print("Документ не выбран!")
        else:
            #  Получим активный документ
            kompas_document = self.application.ActiveDocument
            kompas_document_2d = self.module7.IKompasDocument2D(kompas_document)

            if kompas_document.Name == '':
                print("Активный документ не сохранён на диск!")
            else:
                print("Активный документ:", kompas_document.Name)
                print("Папка документа:", kompas_document.Path)

            # Узнаём тип документа
            if kompas_document.DocumentType == 1:  # Чертёж
                FileTypeSupported = True

                # получаем активный вид
                active_view = kompas_document_2d.ViewsAndLayersManager.Views.ActiveView

                print("Тип документа: Чертёж")
                #  Количество листов
                print("Количество листов:", kompas_document.LayoutSheets.Count)

                # Количество видов
                print("Количество видов:", kompas_document_2d.ViewsAndLayersManager.Views.Count)

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

        print("##########################################")

    def new_drawing(self):
        """
        Создание нового чертежа
        """
        print("Создание нового чертежа...")
        self.kompas_document = self.Documents.AddWithDefaultSettings(self.const7.ksDocumentDrawing, True)
        self.kompas_document_2d = self.module7.IKompasDocument2D(self.kompas_document)
        #self.iDocument2D = self.api7.ActiveDocument2D()
        print("Новый чертёж создан.")
