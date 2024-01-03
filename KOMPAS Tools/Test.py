from KompasClass import *

kompas = Kompas()  # Запуск или подключение к Компас
kompas.info_general()  # Вывод информации о программе

# kompas.newfile_drawing()  # Создание нового чертежа
# kompas.newfile_fragment()  # Создание нового фрагмента

view1_id = kompas.new_view(25, 50, "Тестовый вид", 1 / 15, state=0, color=0x00FF00)  # Создание нового вида

# Построение геометрии
rectangle_id = kompas.draw_rectangle(0, 0, 100, 200, 5, 0)
circle_id = kompas.draw_circle(200, 100, 40, 3)
line_id = kompas.draw_line(0, 0, 200, 100, 4)
point_id = kompas.draw_point(500, 500, 2)

kompas.set_developer_name("Инженер")
kompas.set_developer_date()
kompas.set_inspector_name("Инспектор")
kompas.set_inspector_date("10.01.2024")
kompas.set_tech_control_name("Технический")
kompas.set_tech_control_date()
kompas.set_reg_control_name("Нормативный")
kompas.set_reg_control_date("12.01.2024")
kompas.set_empty_field_name("Пустой")
kompas.set_empty_field_date("00.00.0000")
kompas.set_approver_name("Главный")
kompas.set_approver_date("14.01.2024")
kompas.set_mass_val(200)
kompas.set_scale_text("1:20")
kompas.set_drawing_name("Наименование")
kompas.set_drawing_designation("000.000.000АБВГД")
kompas.set_material_name("Космический сплав")
kompas.set_company_name("Mikarisar Co. Ltd.")

True
