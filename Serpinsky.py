"""
Построение треугольника Серпинского на фрагменте КОМПАС
"""
from KompasClass import *
import numpy as np

# ИЗНАЧАЛЬНЫЕ ПАРАМЕТРЫ
SCALE = 1000  # масштаб треугольника
ITERATIONS = 10000  # количество итераций

# ПЕРЕМЕННЫЕ
# вершины треугольника
A = [0, 0]
B = [1/2, np.sqrt(3)/2]
C = [1, 0]
ABC = [A, B, C]

# случайная точка внутри треугольника (первая)
fx = np.random.randint(1, 100)/100
if fx <= 0.5:
    fy = np.random.randint(1, fx*np.tan(np.pi/3)*100)/100
else:
    fy = np.random.randint(1, (1-fx)*np.tan(np.pi/3)*100)/100
p = [[fx, fy]]

# остальные точки
for _ in range(ITERATIONS):
    U = ABC[np.random.randint(0, len(ABC))]
    G = p[np.random.randint(0, len(p))]
    dx = np.absolute(U[0]-G[0])/2
    dy = np.absolute(U[1]-G[1])/2
    if U[0] > G[0]:
        nx = G[0]+dx
    else:
        nx = G[0]-dx
    if U[1] > G[1]:
        ny = G[1]+dy
    else:
        ny = G[1]-dy
    p.append([nx, ny])

# Инициализация КОМПАС
kompas = Kompas()  # Запуск или подключение к Компас
kompas.info_general()  # Вывод информации о программе

kompas.new_fragment()  # Создание нового фрагмента

# Черчение точек
for i in range(len(p)):
    kompas.draw_point(p[i][0]*SCALE, p[i][1]*SCALE, 0)
