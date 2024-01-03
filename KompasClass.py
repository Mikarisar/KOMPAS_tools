import pythoncom
from win32com.client import Dispatch, gencache
from datetime import date


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

    # МЕТОДЫ ВЫВОДА ИНФОРМАЦИИ (info_)
    def info_general(self):
        """
        Вывод информации о программе
        """
        print("######### ИНФОРМАЦИЯ О ПРОГРАММЕ #########")
        print(f"Версия КОМПАС: {self.app7.ApplicationName(FullName=True)}")  # Печатаем название программы
        print(f"Документов открыто: {self.application.Documents.Count}")
        print("##########################################")

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
                print(f"Активный документ: {kompas_document.Name}")
                print(f"Папка документа: {kompas_document.Path}")

            # Узнаём тип документа
            if kompas_document.DocumentType == 1:  # Чертёж
                file_type_supported = True

                # получаем активный вид
                active_view = kompas_document_2d.ViewsAndLayersManager.Views.ActiveView

                print("Тип документа: Чертёж")
                #  Количество листов
                print(f"Количество листов: {kompas_document.LayoutSheets.Count}")

                # Количество видов
                print(f"Количество видов: {kompas_document_2d.ViewsAndLayersManager.Views.Count}")

                print(f"Активный вид:{active_view.Name}")
                print(f"Масштаб вида:{active_view.Scale}")

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

    # МЕТОДЫ ПОЛУЧЕНИЯ ОБЪЕКТОВ И ПАРАМЕТРОВ (get_)
    def get_active_docs(self):
        """
        Возвращает объекты активного докумета
        :return: kompas_document, kompas_document_2d, iDocument2D
        """
        kompas_document = self.application.ActiveDocument
        kompas_document_2d = self.module7.IKompasDocument2D(kompas_document)
        idocument_2d = self.object5.ActiveDocument2D()
        return kompas_document, kompas_document_2d, idocument_2d

    # МЕТОДЫ СОЗДАНИЯ ФАЙЛОВ (newfile_)
    def newfile_drawing(self):
        """
        Создание нового чертежа
        """
        print("\nСоздание нового чертежа...")
        kompas_document = self.Documents.AddWithDefaultSettings(self.constants.ksDocumentDrawing, True)
        print("Новый чертёж создан.\n")

    def newfile_fragment(self):
        """
        Создание нового фрагмента
        """
        print("\nСоздание нового фрагмента...")
        kompas_document = self.Documents.AddWithDefaultSettings(self.constants.ksDocumentFragment, True)
        print("Новый фрагмент создан.\n")

    # МЕТОДЫ СОЗДАНИЯ ОБЪЕКТОВ И РАБОЧИХ ПРОСТРАНСТВ (new_)
    def new_view(self, x: float, y: float, name: str, scale: float, angle=0, color=0xFF0000, state=3):
        """
        Создание нового вида в активном документе
        :param x: координата x начальной точки
        :param y: координата y начальной точки
        :param name: название вида
        :param scale: масштаб вида
        :param angle: угол наклона в градусах
        :param color: цвет (HEX BGR 000000)
        :param state: состояние???
        :return: id вида
        """
        _, _, idoc2d = self.get_active_docs()

        i_view_param = self.module5.ksViewParam(self.object5.GetParamStruct(self.constants.ko_ViewParam))
        i_view_param.Init()
        i_view_param.angle = angle
        i_view_param.color = color
        i_view_param.name = name
        i_view_param.scale_ = scale
        i_view_param.state = state
        i_view_param.x = x
        i_view_param.y = y

        obj = idoc2d.ksCreateSheetView(i_view_param, 0)

        print(f'\nСоздан вид "{name}" в точке ({x:.2f}, {y:.2f}), масштаб: {scale:.2f}, наклон: {angle:.2f}, цвет BGR: #{color:06X}, состояние: {state}\n')

        return obj

    # МЕТОДЫ СОЗДАНИЯ ГЕОМЕТРИИ (draw_)
    def draw_rectangle(self, x: float, y: float, height: float, width: float, style=1, ang=0):
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

        print(f"Создан прямоугольник в точке ({x:.2f}, {y:.2f}) размером ШxВ {width:.2f}x{height:.2f}, вращ.: {ang:.2f}, стиль: {style}")

        return obj

    def draw_circle(self, x: float, y: float, radius: float, style=1):
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

        print(f"Создана окружность в точке ({x:.2f}, {y:.2f}) с радиусом {radius:.2f}, стиль: {style}")

        return obj

    def draw_line(self, x1: float, y1: float, x2: float, y2: float, style=1):
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

        print(f"Создан отрезок с точками ({x1:.2f}, {y1:.2f}) и ({x2:.2f}, {y2:.2f}), стиль: {style}")

        return obj

    def draw_point(self, x: float, y: float, style=1):
        """
        Создание точки в активном документе
        :param x: координата x
        :param y: координата y
        :param style: стиль
        :return: id точки
        """
        _, _, idoc2d = self.get_active_docs()
        obj = idoc2d.ksPoint(x, y, style)

        print(f"Создана точка ({x:.2f}, {y:.2f}), стиль {style}")

        return obj

    # МЕТОДЫ ЗАДАНИЯ СВОЙСТВ ОБЪЕКТОВ И ПАРАМЕТРОВ (set_)
    def _set_frame_field(self, col_num: int, text='', color=0x000000, font_name="GOST type A", font_height=3.5, style=32768):
        """
        Устанавливает текст в заданной ячейке рамки основной надписи документа
        :param col_num: номер ячейки
        :param text: текст
        :param color: цвет
        :param font_name: шрифт
        :param font_height: высота текста
        :param style: стиль
        """
        _, _, idoc2d = self.get_active_docs()

        i_stamp = idoc2d.GetStamp()

        i_stamp.ksOpenStamp()
        i_stamp.ksColumnNumber(col_num)

        i_textline_param = self.module5.ksTextLineParam(self.object5.GetParamStruct(self.constants.ko_TextLineParam))
        i_textline_param.Init()
        i_textline_param.style = style
        i_textitem_array = self.object5.GetDynamicArray(4)
        i_textitem_param = self.module5.ksTextItemParam(self.object5.GetParamStruct(self.constants.ko_TextItemParam))
        i_textitem_param.Init()
        i_textitem_param.iSNumb = 0
        i_textitem_param.s = text
        i_textitem_param.type = 0
        i_textitem_font = self.module5.ksTextItemFont(i_textitem_param.GetItemFont())
        i_textitem_font.Init()
        i_textitem_font.bitVector = 4096
        i_textitem_font.color = color
        i_textitem_font.fontName = font_name
        i_textitem_font.height = font_height
        i_textitem_font.ksu = 1
        i_textitem_array.ksAddArrayItem(-1, i_textitem_param)
        i_textline_param.SetTextItemArr(i_textitem_array)

        i_stamp.ksTextLine(i_textline_param)
        # i_stamp.ksColumnNumber(col_num)

        i_stamp.ksCloseStamp()

    def set_developer_name(self, name: str):
        """
        Устанавливает имя разработчика в основной надписи
        :param name: имя
        """
        self._set_frame_field(col_num=110, text=name)
        print(f"Установлено имя разработчика: {name}")

    def set_inspector_name(self, name: str):
        """
        Устанавливает имя проверяющего в основной надписи
        :param name: имя
        """
        self._set_frame_field(col_num=111, text=name)
        print(f"Установлено имя проверяющего: {name}")

    def set_tech_control_name(self, name: str):
        """
        Устанавливает имя в строке "Тех. контроль" в основной надписи
        :param name: имя
        """
        self._set_frame_field(col_num=112, text=name)
        print(f"Установлено имя отв. за тех. контроль: {name}")

    def set_empty_field_name(self, name: str):
        """
        Устанавливает имя в пустой строке основной надписи
        :param name: имя
        """
        self._set_frame_field(col_num=113, text=name)
        print(f"Установлено имя в пустой строке: {name}")

    def set_reg_control_name(self, name: str):
        """
        Устанавливает имя в строке "Норм. контроль" в основной надписи
        :param name: имя
        """
        self._set_frame_field(col_num=115, text=name)
        print(f"Установлено имя отв. за норм. контроль: {name}")

    def set_approver_name(self, name: str):
        """
        Устанавливает имя в строке "Утвердил" в основной надписи
        :param name: имя
        """
        self._set_frame_field(col_num=114, text=name)
        print(f"Установлено имя утверждающего: {name}")

    def set_drawing_name(self, name: str):
        """
        Устанавливает наименование в основной надписи
        :param name: наименование
        """
        self._set_frame_field(col_num=1, text=name, font_height=10, style=32769)
        print(f"Установлено наименование: {name}")

    def set_drawing_designation(self, designation: str):
        """
        Устанавливает обозначение в основной надписи
        :param designation: обозначение
        """
        self._set_frame_field(col_num=2, text=designation, font_height=7, style=32770)
        print(f"Установлено обозначение: {designation}")

    def set_material_name(self, name: str):
        """
        Устанавливает название материала в основной надписи
        :param name: название материала
        """
        self._set_frame_field(col_num=3, text=name, font_height=7, style=32771)
        print(f"Установлено название материала: {name}")

    def set_company_name(self, name: str):
        """
        Устанавливает название предприятия в основной надписи
        :param name: название предприятия
        """
        self._set_frame_field(col_num=9, text=name, font_height=7, style=32771)
        print(f"Установлено название предприятия: {name}")

    def set_mass_text(self, mass: float):
        """
        Устанавливает значение массы в основной надписи
        *Учитывается 2 знака после запятой*
        :param mass: масса в кг
        """
        self._set_frame_field(col_num=5, text=str(f"{mass:.2f}"), font_height=5, style=32772)
        print(f"Установлено название предприятия: {mass:.2f}")

    def set_scale_text(self, scale: str):
        """
        Устанавливает значение массы в основной надписи
        :param scale: масштаб (пример: 1:10)
        """
        self._set_frame_field(col_num=5, text=scale, font_height=5, style=32772)
        print(f"Установлено название предприятия: {scale}")

    def set_developer_date(self, date=""):
        """
        Устанавливает дату в строке "Разработал" в основной надписи
        Устанавливает текущую дату при отсутствии аргументов
        :param date: дата в формате ДД.ММ.ГГГГ
        """
        if (date == ""):
            today = date.today()
            date = f"{today.day:02d}.{today.month:02d}.{today.year:04d}"
        self._set_frame_field(col_num=110, text=date)
