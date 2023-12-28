import pythoncom
from win32com.client import Dispatch, gencache


class Kompas(object):
    """
    Класс для взаимодействия с КОМПАС 3D
    """

    def __init__(self):
        # Подключаем константы API Компас
        self.constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
        self.constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

        # Подключаемся к API5
        self.module5 = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
        self.object5 = self.module5.KompasObject(
            Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(self.module5.KompasObject.CLSID, pythoncom.IID_IDispatch))

        # Подключаемся к API7
        self.module7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        self.api7 = self.module7.IKompasAPIObject(
            Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(self.module7.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch))

        # Подключение к нтерфейсу программы Kompas 3D
        print("Подключение к КОМПАС...")
        self.app7 = self.api7.Application                     # Получаем основной интерфейс
        self.app7.Visible = True                              # Показываем окно пользователю (если скрыто)
        self.app7.HideMessage = self.constants.ksHideMessageNo   # Отвечаем НЕТ на любые вопросы программы

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

    def get_active_docs(self):
        """
        Возвращает объекты активного докумета
        :return: kompas_document, kompas_document_2d, iDocument2D
        """
        kompas_document = self.application.ActiveDocument
        kompas_document_2d = self.module7.IKompasDocument2D(kompas_document)
        idocument_2d = self.object5.ActiveDocument2D()
        return kompas_document, kompas_document_2d, idocument_2d

    def info_active(self):
        """
        Вывод информации об активном документе
        """

        print("#### ИНФОРМАЦИЯ ОБ АКТИВНОМ ДОКУМЕНТЕ ####")
        file_type_supported = False

        if self.application.ActiveDocument is None:
            print("Документ не выбран!")
        else:
            #  Получим активный документ
            kompas_document, kompas_document_2d, _ = self.get_active_docs()
            if kompas_document.Name == '':
                print("Активный документ не сохранён на диск!")
            else:
                print("Активный документ:", kompas_document.Name)
                print("Папка документа:", kompas_document.Path)

            # Узнаём тип документа
            if kompas_document.DocumentType == 1:  # Чертёж
                file_type_supported = True

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

            if not file_type_supported:
                print("Этот тип документа не поддерживается!")

        print("##########################################")

    def new_drawing(self):
        """
        Создание нового чертежа
        """
        print("Создание нового чертежа...")
        kompas_document = self.Documents.AddWithDefaultSettings(self.constants.ksDocumentDrawing, True)
        # kompas_document_2d = self.module7.IKompasDocument2D(kompas_document)
        print("Новый чертёж создан.")

    def draw_rectangle(self, x: int, y: int, height: int, width: int, style=1, ang=0):
        """
        Создание прямоугольника в активном документе
        :param x: координата x начальной точки
        :param y: координата y начальной точки
        :param height: высота
        :param width: ширина
        :param style: стиль линии
        :param ang: угол наклона в градусах
        :return: id прямоугольника
        """
        _, _, idoc2d = self.get_active_docs()

        i_rec_param = self.module5.ksRectangleParam(self.object5.GetParamStruct(self.constants.ko_RectangleParam))
        i_rec_param.Init()

        i_rec_param.x = x
        i_rec_param.y = y
        i_rec_param.ang = ang
        i_rec_param.height = height
        i_rec_param.width = width
        i_rec_param.style = style

        obj = idoc2d.ksRectangle(i_rec_param)

        print("Создан прямоугольник в точке (", x, ", ", y, ") размером ШxВ ", width, "x", height, ", вращ.: ", ang, ", стиль: ", style, sep="")

        return obj

    def draw_circle(self, x: int, y: int, radius: int, style=1):
        """
        Создание окружности в активном документе
        :param x: координата x начальной точки
        :param y: координата y начальной точки
        :param radius: радиус окружности
        :param style: стиль линии
        :return: id окружности
        """
        _, _, idoc2d = self.get_active_docs()
        obj = idoc2d.ksCircle(x, y, radius, style)

        print("Создана окружность в точке (", x, ", ", y, ") с радиусом ", radius, ", стиль: ", style, sep="")

        return obj

    def draw_line(self, x1: int, y1: int, x2: int, y2: int, style=1):
        """
        Создание отрезка в активном документе
        :param x1: координата x первой точки
        :param y1: координата y первой точки
        :param x2: координата x второй точки
        :param y2: координата y второй точки
        :param style: стиль линии
        :return: id отрезка
        """
        _, _, idoc2d = self.get_active_docs()
        obj = idoc2d.ksLineSeg(x1, y1, x2, y2, style)

        print("Создан отрезок с точками (", x1, ", ", y1, ") и (", x2, ", ", y2,"), стиль: ", style, sep="")

        return obj
