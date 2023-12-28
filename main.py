import pythoncom
from win32com.client import Dispatch, gencache
from KompasClass import *

kompas = Kompas()  # Запуск или подключение к Компас
kompas.info_general()  # Вывод информации о программе
kompas.info_active()  # Вывод информации об активном документе
kompas.new_drawing()  # Создание нового чертежа
kompas.info_active()
