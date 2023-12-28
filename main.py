from KompasClass import *

kompas = Kompas()  # Запуск или подключение к Компас
kompas.info_general()  # Вывод информации о программе
kompas.new_drawing()  # Создание нового чертежа
rectangle = kompas.draw_rectangle(10, 20, 100, 200, 2, 45)
circle = kompas.draw_circle(100, 100, 40, 3)
line = kompas.draw_line(-50, -60, 60, 50, 4)

True
