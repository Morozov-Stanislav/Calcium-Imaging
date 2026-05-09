import glob
import numpy as np
import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QDesktopWidget, QFileDialog, QProgressBar, QApplication, QTreeWidgetItem
from PyQt5.Qt import Qt
from matplotlib import pyplot as plt, colors, colormaps, patches as patches
from matplotlib.widgets import CheckButtons, Slider, Button
import copy
import os
import colorsys
from numpy import AxisError
from scipy.ndimage import gaussian_filter, shift as sh, rotate
import multiprocessing


# Здесь стеки для каждого дня и пробы выравниваются и обсчитываются без цикла, последовательно вызывая
# метод compute_ori() по нажатию кнопки на окне с выравниванием картинок.
# Добавлена функция skip().
#
# Предположительно исправлена ошибка св. с сортировкой списка ориентаций - в старых версиях, возможно,
# после первой итерации ориентации интерпретировались неправильно, так как соотносились с уже отсортированным
# списком (нет, в старых версиях всё было нормально, просто через левое колено).
# Внимания также требует конвертация переменной res в глобальную self.res, так как исторически у них разные размерности.
# Для ускорения чтения файлов и поворота картинок применён модуль multiprocessing
# 25.07 - Объединены фрагменты с выравниванием и основное тело программы (визуализации). Работает
# 30.07 - добавлен расчёт ориентационной селективности (кнопка OI). Общепринятая формула даёт неинформативные рез-ты:
# Из-за наличия отриц. значений для ортогональных ориентаций индекс получается >100% для ROI, где ничего нет.
# Вместо этого сделал расчёт OI как величины векторной суммы ответов на все ориентации (!!! - после пересчёта ответов
# в формат 180 градусов). Значения очень наглядные, но количественно ничего не отражают - их надо бы к чему-то
# отнести (к макс. ответу в данной пробе?, или проводить этот расчёт только в групповом анализе?
# - а как же тогда статистика, коей нужны индивид. значения?)
# 04.08 - multiprocessing для ф-ции compute_squares(), переключение отображения настройки с 360 на 180 градусов.
# 05.08 - отображение pref ori на круговом графике; суммарный вектор (угол pref ori) вычисляется на этапе рассчёта,
# в функции compute_squares(), и сохраняется в self.data; переделал функцию сохранения более удобно
# 07.08 - во вращении стеков параметр order=0, что убирает интерполяцию и искажение. Значения за границей = 0.
# 03.09 - вроде бы наконец правильный расчёт углов pref_ori и pref_dir (с исп. np.arctan2), сохранение их, а также
# индексов ориентационной и дирекциональной селективности (OI и DI) в экселе на листе "Preferred ori".
# Расчёт суммарных векторов по формуле из статьи  https://doi.org/10.3389/fncir.2014.00092
# (предварительный перевод всех ответов в формат 180 град. исп-ся лишь для отображения на графике, но не для рассчёта);
# также перераспределил более логично методы в тексте кода.
# 04.09 - добавлена проверка на одинаковое количество файлов (кадров) для каждой ориентации в пределах одной пробы.
# 05.09 - добавлено сохранение ответов на все ориентации (переменная matrix) в виде стека tiff; добавлено переключение
# с OI на DI на картинке.
# 06.09 - добавлена возможность выравнивания стеков не только слайдерами, но и стрелками на клавиатуре, плюс колёсиком;
# Поменяны местами координаты X и Y при выравнивании (align, compute_squares), Y-нулевая ось, X-первая (как и должно);
# Activity trace теперь можно переключать между разными пробами во вкладке Multisession; Исправлен расчёт OI в average.
# 07.09 - в окне align проекции теперь не перевёрнутые (поменял границы оси местами); размеры экрана - глоб. переменные;
# оптимизирован импорт: убран импорт всего из PyQt5 и импорт math (все нужные ф-ции math есть в np); исправлена
# проблема с цветовой шкалой в OI.
# 8.09 - исправил границы отображаемых цветом значений в OI; исправил ошибку при сохранении усред. и неусред. сессий
# в один файл Excel.
# 9.09 - исправлено обновление и отображение предпочитаемой ориентации/направления (ф-ция show()); исправлено сохранение
# pref_dir в радианах, а не в градусах при сохранении всех ROI сразу.
# 11.09 - исправил ошибку: в параметры всех проб записывались параметры последней обработанной пробы (с пом. deepcopy);
# добавил сохранение значений сдвига каждой пробы отн-но первой в сохраняемые параметры.
# 14.09 - добавил возможность автоматического выравнивания стеков на основании значений сдвига и поворота из файла txt.
# Для этого переменная self.image1_init заменена на self.stack12ori. При наличии координат ф-ция align() пропускается и
# сразу запускается get_shift(). Файл с координатами автоматически сохранятеся после выравнивания всех проб.
# 14.09-2 - вместо отдельных ячеек словаря для pref_ori, pref_dir, OI, DI сделал одну переменную 'maps', содержащую
# все эти таблицы, плюс индекс целевой селективности TI.
# 15.09 - сохраняются tif-ы гиперстеками по дням. Сохраняется переменная matrix без отрицательных значений (matrix_corr)
# в отдельный лист Excel. Исправил мелкие ошибки, в.ч. сохранение параметров при
# child_index = 0 и 2; проверку cursess < len(checked). Иногда не замыкается кривая настройки!
# 17.09 - добавил везде, где что-то копируется в или из переменной self.data функции указания типа - чтобы создавалась
# настоящая копия, а не ссылка на тот же объект, т.к. из-за этого вероятно данные искажаются с каждой итерацией???
# 29-30.09 - Начал писать функции statistic и plot_statistic - считают тест Манна-уитни между выбранными днями,
# выводят графики изменения всех индексов по дням и таблицы с p-value (НЕ проверял!). Потом должны будут также
# считать % ROI, где есть значимые изменения, от кол-ва ROI, ответ в которых превосходит порог по амплитуде (mask_resp).
# 02.10 - Доделал предыдущее, сделал % ROI, сделал независимый выбор ROI для основных картинок и картинок статистики;
# также выводятся кривые настройки, усреднённые по пробам для каждого дня. Проверял - работает.
# 3.10 - при обработке больших данных - Memory Error. Автоматическое выравнивание по файлу с align_values заменил на
# автоматический поворот и сдвиг на нужные значения в ф. align. Далее нужно только нажать 'Apply'/'Enter'.
# 4.10 Исправил ошибку в расчёте теста М-У и др. ошибки. НО процент изменяющихся ROI считает неправильно!
# 5.10 Исправил ошибку в рассчёте процента изменяющихся ROI (учитывались нули в нижнем треугольнике матрицы) -> np.ones
# А также добавил строку, чтобы не учитывались области с NaN (т.е. ROI, которые выпадали в тот или иной день), даже
# если между какими-то днями вних было обнаружено значимое отличие. Надо ещё перепроверить.
# 6.10 - Изменил вид графиков, сделал автоматическое распределение сетки подграфиков ориентационной настройки.
# 7.10 - В графиках ф-ции statistics() и plot_grand_aver() деления осей линейных и полярных графиков сделал такими же,
# как на основных графиках (нагляднее).
# 20.10 - исправил ошибку отображения графика средней ампл. ответа на предпочитаемую ориентацию (добавлен ранее).
# 25.10 - Исправил проблему: не обновлялся диапазон слайдера, регулирующего порог амплитуды ответа для отображения
# при переключении пробы. Плюс добавил коррекцию порога маскирования по амплитуде, если он выходит за новые границы
# амплитуды (до этого проверка стояла в ф-ции change_sess, а надо было в separate).
# Проблема: коррекция порога маскирования из-за изменения диапазона амплитуды будет вызывать display() в тот момент,
# когда выбраны все пробы -> задача нарисовать кучу рафиков -> зависание. Нужно сделать отдельный слайдер для grand_av.
# 26.02.25 - исправил ориентацию полярных графиков - теперь 0 градусов сверху (что соотвествует реальным ориентациям
# решёток в Psychopy - 0-180 вертикаль), также поменял направление полярных графиков на по часовой стрелке.
# 01.04.25 - исправил отображение нескольких усреднённых ориентационных настроек (было смещение заливки STD от его
# границ) - просто надо было .canvas.draw поставить не в цикле для каждой пробы, а один раз в конце. И ещё поднял
# заголовок на графике ориентационной настройки.

# MULTIPROCESSING FUNCTIONS ---------------------------------------------------------------- MULTIPROCESSING FUNCTIONS
# 1. Function for reading tiffs of cur session and building a stack in Multiprocessing loop
def read_tiffs(subfolders, ori, filtr):
    # (f"    Reading ori {ori}")
    tiffs = glob.glob(f'{subfolders}\\*.tif')
    tiffs.remove(f'{subfolders}\\ChanB_Preview.tif')  # it's not an informative frame
    stack = []  # Stack of one ori: [frames, x, y]
    for tif in tiffs:  # Loop over frames
        img = plt.imread(tif)
        if filtr != 0:
            img = gaussian_filter(img, sigma=filtr)
        imgarray = np.array(img)
        stack.append(imgarray)
    stack_arr = np.array(stack)
    return stack_arr, ori


# 2. Function for implementing desired shift on a stack in Multiprocessing loop
def shift_rot(res, ori, deg, x, y):  # res: [frames y, x]
    # print(f"    Shifting {ori} ori", end='\r')
    res = rotate(res, angle=deg, reshape=False, axes=(1, 2), mode='constant', cval=0, order=0)  # Rotating initial stack
    res = sh(res, (0, y, x), mode='constant', cval=0, order=0)  # Shifting initial image [frames, y, x]
    return res, ori


# 3. Function for averaging deltaF/F signal over a ROIs, & computing response amplitude for each ROI
def comp_sq(stack_arr, ori, param):
    # param = [0 - h, 1 - w, 2 - trial_len, 3 - filt_time, 4 - wind_size, 5 - st_lat_fr, 6 - windbrdr]
    result = np.zeros((param[0], param[1], param[2]))
    matrix = np.zeros((param[0], param[1]))
    for i in range(param[0]):  # Что тут с X-ами и Y-ами??? Между res и self.res.
        for j in range(param[1]):
            trace = np.zeros(param[2])  # var for trace of current roi
            patch = stack_arr[:,
                    param[4] * i:param[4] * (i + 1),
                    param[4] * j:param[4] * (j + 1)
                    ]  # Fragment of the stack, i.e. curr ROI
            for t in range(param[2]):  # For each frame
                trace[t] = np.mean(patch[t, :, :])  # Averaging values in ROI
            if param[3] > 0:  # filtering row signal
                trace = gaussian_filter(trace, param[3])
            pre_aver = np.mean(trace[0:param[5] - 1])  # Prestimulus level (basal level,F)
            if pre_aver != 0:
                trace_delta = (trace - pre_aver) / pre_aver  # Converting into deltaF/F
            else:
                # trace_delta = np.zeros(trace.shape)  # Otherwise, zero division -> NaN
                trace_delta = np.full(trace.shape, np.nan)
            result[i, j, :] = trace_delta
            val = np.mean(trace_delta[param[5]:param[6]])  # It's already normalized by prestim
            # since intensity was converted into deltaF/F
            matrix[i, j] = val
    return result, matrix, ori


# AUXILIARY FUNCTIONS ----------------------------------------------------------------------------- AUXILIARY FUNCTIONS
def find_devisor(val, curr):  # find the divisor of image shape ('val') closest to current one (size of square, 'curr')
    print("find_devisor")
    divisors = []
    for i in range(1, int(np.sqrt(val)) + 1):
        if val % i == 0:
            divisors.append(i)
            if i != val // i:  # integer division
                divisors.append(val // i)  # Find all divisors
    new = min(divisors, key=lambda x: abs(x - curr))  # Select the nearest divisor
    return new


def get_date(path):
    import datetime
    # Getting the date when files in given list directories were created.
    # Listing files for 1st ori, then check date for 1st tif file
    file = [os.path.join(path[0], f) for f in os.listdir(path[0]) if f.endswith(".tif")]
    c_time = os.path.getmtime(file[0])
    c_date = datetime.datetime.fromtimestamp(c_time).date()
    return c_date


class Ui_MainWindow(object):
    def __init__(self):
        self.dispOI = 'matrix_OI'
        self.cursess = 0
        self.pbar = None
        self.check = None
        self.ori_disp = None
        self.activity = False
        self.proj_lay = None
        self.ax3 = None
        self.fig3 = None
        self.img = None
        self.ax2 = None
        self.made = False
        self.gen_img_made = False
        self.axim = None
        self.fig2 = None
        self.flat = None
        self.colormap = None
        self.coord = None
        self.ax = None
        self.fig1 = None
        self.res = None  # [y, x, 12 ori, frames]
        self.folder = ''
        self.folder_type = None
        self.initial_ori = [6.0e+01, 9.0e+01, 1.5e+02, 3.0e+01, 1.2e+02, 0.0e+00, 2.1e+02, 2.7e+02, 1.8e+02, 3.0e+02,
                            2.4e+02, 3.3e+02]
        self.ori_list = copy.deepcopy(self.initial_ori)
        self.ori_list.sort()
        self.cur_day = 0
        self.cur_sess = 0
        self.subfolders = None
        self.count = 0  # For displaying progress of computation

        self.rads = np.arange(0, (2 * np.pi), 0.01)  # For drawing circle at 0.
        self.radius0 = np.full(shape=self.rads.shape, fill_value=0, dtype=int)
        self.rad = np.zeros(len(self.ori_list) + 1)  # Degrees to radians:
        self.rad[:-1] = np.radians(self.ori_list)
        self.rad[-1] = self.rad[0]  # Copying last value to close the curve

        self.wind_size = 8  # size of square ROI in px
        self.filter = 2  # Equal to sigma of gaussian filter XY; if 0, not filter
        self.filt_time = 3  # Equal to sigma of gaussian filter T; if 0, not filter

        self.frrate = 12.719  # Hz
        self.duration = 2  # sec (duration of stimulation)
        self.stim_lat = 2  # sec (latency of stimulation onset, i.e. prestimulus interval)
        self.st_lat_fr = round(self.frrate * self.stim_lat)  # frame of stimulation onset
        self.windbrdr = round(self.frrate * (
                self.stim_lat + self.duration))  # End of stimulation
        self.trial_len = 85  # frames (will be recomputed further)
        self.clrs = ['#ffffff', '#000', '#f00', '#f80', '#ffe500', '#bf0', '#2f0', '#0f7', '#00ffd0', '#00e1ff',
                     '#0084ff', '#0400ff', '#ae00ff', '#ff00f7']
        self.curr_roi = (0, 0)  # Current square roi of the matrix
        self.cur_roi_stat = (0, 0)  # Current square roi of the matrix for statistics figures
        self.curr_ori = 1  # Current orientation (i.e. slice of the stack)
        self.mask_resp = 0  # Minimal value of deltaF/F to display in maps
        self.ampl_thrs = 0  # Minimal value of deltaF/F to display in grand_aver
        self.p_values = None

        self.name = ''
        self.par = {'day': '',
                    'name': self.name,
                    'frrate': self.frrate,
                    'Latency_of_stimulation_onset(s)': self.stim_lat,
                    'Poststimulus_interval_start(frames)': self.st_lat_fr,
                    'Stimulation_duration(s)': self.duration,
                    'Poststimulus_interval_end(frames)': self.windbrdr,
                    'Size_of_square_ROI(px)': self.wind_size,
                    'filterXY(px)': self.filter,
                    'filterT(frames)': self.filt_time,
                    'Shift_X': 0,
                    'Shift_Y': 0,
                    'Shift_Rot': 0
                    }

        self.aver = False
        self.disp180 = 'matrix'  # 'matrix' - plot 360 degree, 'matrix180' - plot 180 degree

        # Variables for alignment:
        self.ref_stack = None  # Reference MAX-stack projection for alignment
        self.fig_ali = None
        self.ax_ali = None
        self.reference = None
        self.image1 = None
        self.im = None
        self.shift = {'x': 0, 'y': 0, 'deg': 0}
        self.shifted = False
        self.img_ref = None
        self.img_res = None
        self.h_low = None
        self.h_high = None
        self.w_low = None
        self.w_high = None

        # Dimensions: [day session Ori X Y Frame]
        self.data = {}  # Collecting all the data

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(798, 265)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        font = QtGui.QFont()
        font.setPointSize(10)

        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setGeometry(QtCore.QRect(0, 0, 798, 265))
        self.tabWidget.setObjectName("tabWidget")
        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")
        self.tabWidget.addTab(self.tab, "")
        self.tab_2 = QtWidgets.QWidget()
        self.tab_2.setObjectName("tab_2")
        self.tabWidget.addTab(self.tab_2, "")
        self.tab_3 = QtWidgets.QWidget()
        self.tab_3.setObjectName("tab_3")
        self.tabWidget.addTab(self.tab_3, "")
        self.tabWidget.setFont(font)

        self.select_btn = QtWidgets.QPushButton(self.tab)
        self.select_btn.setGeometry(QtCore.QRect(40, 20, 141, 31))
        self.select_btn.setFont(font)
        self.select_btn.setObjectName("select_btn")
        self.select_btn.clicked.connect(lambda: self.search())

        self.go = QtWidgets.QPushButton(self.tab)
        self.go.setGeometry(QtCore.QRect(40, 60, 141, 31))
        self.go.setFont(font)
        self.go.setObjectName("go")
        self.go.clicked.connect(lambda: self.calculate())

        self.disp_act = QtWidgets.QPushButton(self.tab)
        self.disp_act.setGeometry(QtCore.QRect(40, 100, 141, 31))
        self.disp_act.setFont(font)
        self.disp_act.setObjectName("disp_act")
        self.disp_act.clicked.connect(lambda: self.disp_activ())

        self.show_btn = QtWidgets.QPushButton(self.tab)
        self.show_btn.setGeometry(QtCore.QRect(40, 140, 141, 31))
        self.show_btn.setFont(font)
        self.show_btn.setObjectName("show_btn")
        self.show_btn.clicked.connect(lambda: self.show())

        self.MU_btn = QtWidgets.QPushButton(self.tab)
        self.MU_btn.setGeometry(QtCore.QRect(190, 100, 50, 31))
        self.MU_btn.setFont(font)
        self.MU_btn.setObjectName("MU_btn")
        self.MU_btn.clicked.connect(lambda: self.statistic())

        self.OI_btn = QtWidgets.QPushButton(self.tab)
        self.OI_btn.setGeometry(QtCore.QRect(190, 140, 50, 31))
        self.OI_btn.setFont(font)
        self.OI_btn.setObjectName("OI_btn")
        self.OI_btn.clicked.connect(lambda: self.ori_index())

        self.save_label = QtWidgets.QLabel(self.tab)
        self.save_label.setGeometry(QtCore.QRect(60, 180, 60, 30))
        self.save_label.setText('Save:')
        self.save_label.setFont(font)

        self.save = QtWidgets.QPushButton(self.tab)
        self.save.setGeometry(QtCore.QRect(120, 180, 60, 31))
        self.save.setFont(font)
        self.save.setObjectName("save")
        self.save.clicked.connect(lambda: self.savef())

        self.save_tif = QtWidgets.QPushButton(self.tab)
        self.save_tif.setGeometry(QtCore.QRect(190, 180, 50, 31))
        self.save_tif.setFont(font)
        self.save_tif.setObjectName("save_tif")
        self.save_tif.clicked.connect(lambda: self.save_stacks())

        self.pbar = QProgressBar(self.tab)
        self.pbar.setGeometry(280, 64, 390, 25)
        self.pbar.setAlignment(Qt.AlignCenter)

        self.sl_label = QtWidgets.QLabel(self.tab)
        self.sl_label.setGeometry(QtCore.QRect(200, 62, 65, 25))
        self.sl_label.setText('Progress')
        self.sl_label.setFont(font)

        self.file_path = QtWidgets.QTextEdit(self.tab)
        self.file_path.setGeometry(QtCore.QRect(200, 20, 530, 31))
        self.file_path.setObjectName("file_path")
        self.file_path.setFont(font)

        # TAB_2 ------------------------------------------------------------------------------------------------TAB_2
        self.fps = QtWidgets.QTextEdit(self.tab_2)
        self.fps.setGeometry(QtCore.QRect(135, 20, 60, 30))
        self.fps.setObjectName("fps")
        font.setPointSize(9)
        self.fps.setFont(font)
        self.fps.setText(str(self.frrate))
        self.fps.textChanged.connect(lambda: self.get_parameters())

        self.fps_label = QtWidgets.QLabel(self.tab_2)
        self.fps_label.setGeometry(QtCore.QRect(20, 20, 30, 30))
        self.fps_label.setText('fps')
        self.fps_label.setFont(font)

        self.st_lat = QtWidgets.QTextEdit(self.tab_2)
        self.st_lat.setGeometry(QtCore.QRect(135, 55, 60, 30))
        self.st_lat.setObjectName("st_lat")
        self.st_lat.setFont(font)
        self.st_lat.setText("2")
        self.st_lat.textChanged.connect(lambda: self.get_parameters())

        self.st_lat_label = QtWidgets.QLabel(self.tab_2)
        self.st_lat_label.setGeometry(QtCore.QRect(20, 55, 95, 30))
        self.st_lat_label.setText('Stim onset, s')
        self.st_lat_label.setFont(font)

        self.dur = QtWidgets.QTextEdit(self.tab_2)
        self.dur.setGeometry(QtCore.QRect(135, 90, 60, 30))
        self.dur.setObjectName("dur")
        self.dur.setFont(font)
        self.dur.setText("2")
        self.dur.textChanged.connect(lambda: self.get_parameters())

        self.dur_label = QtWidgets.QLabel(self.tab_2)
        self.dur_label.setGeometry(QtCore.QRect(20, 90, 105, 30))
        self.dur_label.setText('Stim duration, s')
        self.dur_label.setFont(font)

        self.sq = QtWidgets.QTextEdit(self.tab_2)
        self.sq.setGeometry(QtCore.QRect(135, 125, 60, 30))
        self.sq.setObjectName("sq")
        self.sq.setFont(font)
        self.sq.setText("8")
        self.sq.textChanged.connect(lambda: self.get_parameters())

        self.sq_label = QtWidgets.QLabel(self.tab_2)
        self.sq_label.setGeometry(QtCore.QRect(20, 125, 105, 30))
        self.sq_label.setText('Square size, px')
        self.sq_label.setFont(font)

        self.filtXY = QtWidgets.QTextEdit(self.tab_2)
        self.filtXY.setGeometry(QtCore.QRect(385, 20, 60, 30))
        self.filtXY.setObjectName("filtXY")
        self.filtXY.setFont(font)
        self.filtXY.setText("2")
        self.filtXY.textChanged.connect(lambda: self.get_parameters())

        self.filtXY_label = QtWidgets.QLabel(self.tab_2)
        self.filtXY_label.setGeometry(QtCore.QRect(270, 20, 105, 30))
        self.filtXY_label.setText('Filter XY, px')
        self.filtXY_label.setFont(font)

        self.filtT = QtWidgets.QTextEdit(self.tab_2)
        self.filtT.setGeometry(QtCore.QRect(385, 55, 60, 30))
        self.filtT.setObjectName("filtT")
        self.filtT.setFont(font)
        self.filtT.setText("2")
        self.filtT.textChanged.connect(lambda: self.get_parameters())

        self.filtT_label = QtWidgets.QLabel(self.tab_2)
        self.filtT_label.setGeometry(QtCore.QRect(270, 55, 105, 30))
        self.filtT_label.setText('Filter T, fr')
        self.filtT_label.setFont(font)

        self.targ_ori = QtWidgets.QTextEdit(self.tab_2)
        self.targ_ori.setGeometry(QtCore.QRect(385, 90, 60, 30))
        self.targ_ori.setObjectName("targ_ori")
        self.targ_ori.setFont(font)
        self.targ_ori.setText("0")

        self.targ_ori_label = QtWidgets.QLabel(self.tab_2)
        self.targ_ori_label.setGeometry(QtCore.QRect(270, 90, 105, 30))
        self.targ_ori_label.setText('Target ori, deg')
        self.targ_ori_label.setFont(font)

        self.save_res = QtWidgets.QCheckBox(self.tab_2)
        self.save_res.setGeometry(QtCore.QRect(675, 20, 20, 20))
        self.save_res_label = QtWidgets.QLabel(self.tab_2)
        self.save_res_label.setGeometry(QtCore.QRect(510, 20, 130, 30))
        self.save_res_label.setText('Save traces')
        self.save_res_label.setFont(font)

        self.save_matrix = QtWidgets.QCheckBox(self.tab_2)
        self.save_matrix.setGeometry(QtCore.QRect(675, 55, 20, 20))
        self.save_matrix.setChecked(True)
        self.save_matrix_label = QtWidgets.QLabel(self.tab_2)
        self.save_matrix_label.setGeometry(QtCore.QRect(510, 55, 130, 30))
        self.save_matrix_label.setText('Save ori preference')
        self.save_matrix_label.setFont(font)

        self.save_map = QtWidgets.QCheckBox(self.tab_2)
        self.save_map.setGeometry(QtCore.QRect(675, 90, 20, 20))
        self.save_map.setChecked(True)
        self.save_map_label = QtWidgets.QLabel(self.tab_2)
        self.save_map_label.setGeometry(QtCore.QRect(510, 90, 130, 30))
        self.save_map_label.setText('Save pref. ori')
        self.save_map_label.setFont(font)

        self.save_param = QtWidgets.QCheckBox(self.tab_2)
        self.save_param.setGeometry(QtCore.QRect(675, 125, 20, 20))
        self.save_param.setChecked(True)
        self.save_param_label = QtWidgets.QLabel(self.tab_2)
        self.save_param_label.setGeometry(QtCore.QRect(510, 125, 130, 30))
        self.save_param_label.setText('Save parameters')
        self.save_param_label.setFont(font)

        self.save_curroi = QtWidgets.QCheckBox(self.tab_2)
        self.save_curroi.setGeometry(QtCore.QRect(675, 160, 20, 20))
        self.save_curroi.setChecked(True)
        self.save_curroi_label = QtWidgets.QLabel(self.tab_2)
        self.save_curroi_label.setGeometry(QtCore.QRect(510, 160, 130, 30))
        self.save_curroi_label.setText('Only current ROI')
        self.save_curroi_label.setFont(font)

        # TAB_3 -------------------------------------------------------------------------------------------------TAB_3
        self.instruction = QtWidgets.QLabel(self.tab_3)
        self.instruction.setGeometry(QtCore.QRect(20, 15, 260, 30))
        self.instruction.setText('Select sessions to average:')
        font.setPointSize(10)
        self.instruction.setFont(font)

        self.aver_ss = QtWidgets.QPushButton(self.tab_3)
        self.aver_ss.setGeometry(QtCore.QRect(300, 20, 141, 31))
        self.aver_ss.setFont(font)
        self.aver_ss.setObjectName("aver_ss")
        self.aver_ss.clicked.connect(lambda: self.average())

        self.split_ss = QtWidgets.QPushButton(self.tab_3)
        self.split_ss.setGeometry(QtCore.QRect(300, 60, 141, 31))
        self.split_ss.setFont(font)
        self.split_ss.setObjectName("split_ss")
        self.split_ss.clicked.connect(lambda: self.separate())

        self.message_tab3 = QtWidgets.QLabel(self.tab_3)
        self.message_tab3.setGeometry(QtCore.QRect(300, 150, 450, 50))
        self.message_tab3.setFont(font)
        self.message_tab3.setWordWrap(True)

        # TREE ---------------------------------------------------------------------------------------------------TREE
        self.tree = QtWidgets.QTreeWidget(self.tab_3)
        self.tree.setGeometry(QtCore.QRect(20, 15, 270, 190))
        self.tree.setHeaderHidden(True)
        self.tree.show()
        # ---------------------------------------------------------------------------------------------------------

        self.message = QtWidgets.QLabel(self.tab)
        self.message.setGeometry(QtCore.QRect(280, 150, 470, 50))
        self.message.setText('')
        font.setPointSize(10)
        self.message.setFont(font)

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 798, 26))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def search(self):  # Searching the data files
        # NB! This automatically calls OnClick function!
        self.res = None  # Otherwise, new data won't be recalculated
        self.align_values = None
        self.folder = str(QFileDialog.getExistingDirectory(None, "Select Directory")).replace('/', '\\')
        folder_name = self.folder.split('\\')[-1]
        # Now, it's needed to define what does this folder actually contains:
        if os.path.isdir(self.folder):  # If directory exists
            self.pbar.setFormat('Analyzing hierarchy...')
            self.pbar.setValue(0)  # Displaying progress of computation
            QApplication.processEvents()  # Updating progress bar
            all_subf = list(os.walk(self.folder))  # List of all subdirectories
            num_steps = []  # List for num  of steps from folder with 12 subfolders to mother folder (self.folder)
            for i in range(len(all_subf)):
                if len(all_subf[i][1]) == 12:  # this folder has 12 subfolders
                    # Splitting the path, inverting the path and getting idx of mother folder in it:
                    num_steps.append(list(reversed(all_subf[i][0].split('\\'))).index(folder_name))
                    num_tiffs = []
                    for fold_ori in range(12):
                        tiffs = glob.glob(f'{all_subf[i][0]}\\{all_subf[i][1][fold_ori]}\\*.tif')
                        num_tiffs.append(len(tiffs))
                    if num_tiffs.count(num_tiffs[0]) != len(num_tiffs):
                        print(f"Folder {all_subf[i][0]} has different number of frames per ori:\n{num_tiffs}")
            if num_steps.count(num_steps[0]) == len(num_steps):  # If all elements are identical
                self.ch_idx = num_steps[0]
                self.file_path.setText(self.folder)
                self.made = False  # to prevent displaying during data change
                self.message.setText("")
                self.pbar.setFormat('Ready')
                QApplication.processEvents()  # Updating progress bar
                print(f"Consensus child index is: {self.ch_idx}")

                # Loading the file with shift and rotation values, if exists:
                if os.path.isfile(self.folder + f'\\Align_{folder_name}.txt'):
                    print("File with align coordinates exists")
                    self.align_values = pd.read_table(self.folder + f'\\Align_{folder_name}.txt',
                                                      delimiter="\t").to_numpy(dtype=object)
                    self.align_values = np.char.mod('%s', self.align_values)  # All elements to string
                else:
                    print(f"File with align coordinates {self.folder}\\Align_{folder_name}.txt was nor found")
            else:
                self.message.setText(f"Wrong data structure")
        else:
            self.message.setText(f"Can not find directory {self.folder}")
            self.file_path.setText("")
            self.made = False

    def calculate(self):  # Perform main computations
        # В данной версии (24.07) тело функции calculate лишь заполняет "метаданные" об иерархической структуре
        # данных в переменную self.data, ориентируясь на которую, последующие функции -
        # self.compute_ori (считывает файлы текущей сессии и собирает стек)
        # self.align (выравнивает текущую сессию с первой - референсной по макс. проекции)
        # self.get_shift (в случае нажатия кнопки, применяет найденное смещение ко всему стеку)
        # self.compute_squares (совершает анализ "по квадратам" как в старых версиях программы)
        # - анализируют каждую сессию по очереди.
        np.seterr(divide='ignore', invalid='ignore')  # Ignoring by zero division
        self.get_parameters()
        self.folder = self.file_path.toPlainText()
        if len(self.folder) > 0:
            plt.close('all')
            self.row_sess = []  # row data for al sessions [12ori frames x y]
            self.max_proj = []  # max projections of all sessions
            if self.ch_idx == 0:  # folder contains 12 ori
                # Search for day with same date in self.data:
                self.subfolders = [f.path for f in os.scandir(self.folder) if f.is_dir()]
                cr_date = get_date(self.subfolders)
                same_day = None
                for day_k, day_v in self.data.items():
                    if isinstance(day_v, dict):
                        for sess in day_v.values():
                            if sess['date'] == cr_date:
                                same_day = day_k
                sess_name = self.folder.split('\\')[-1]
                self.par['name'] = sess_name
                if same_day is not None:
                    print(f'Appending to the day {same_day}')
                    self.data[f'{same_day}'][f'{sess_name}'] = {}
                    self.data[f'{same_day}'][f'{sess_name}']['path'] = self.folder
                    self.data[f'{same_day}'][f'{sess_name}']['date'] = cr_date
                    self.par['day'] = same_day
                    self.data[f'{same_day}'][f'{sess_name}']['param'] = self.par
                else:  # If no data on this day is present, create a new day
                    self.data[f'{str(cr_date)}'] = {}
                    self.data[f'{str(cr_date)}'][f'{sess_name}'] = {}
                    self.data[f'{str(cr_date)}'][f'{sess_name}']['path'] = self.folder
                    self.data[f'{str(cr_date)}'][f'{sess_name}']['date'] = cr_date
                    self.par['day'] = str(cr_date)
                    self.data[f'{str(cr_date)}'][f'{sess_name}']['param'] = copy.deepcopy(self.par)

            elif self.ch_idx == 1:  # folder contains N sessions with 12 ori
                day_name = self.folder.split("\\")[-1]
                self.data[f'{day_name}'] = {}
                for sess in [f.path for f in os.scandir(self.folder) if f.is_dir() and len(next(os.walk(f))[1]) == 12]:
                    sess_name = sess.split("\\")[-1]
                    self.par['day'] = day_name
                    self.par['name'] = sess_name
                    self.data[f'{day_name}'][f'{sess_name}'] = {}
                    self.res = None
                    self.data[f'{day_name}'][f'{sess_name}']['path'] = sess
                    self.subfolders = [f.path for f in os.scandir(sess) if f.is_dir()]
                    cr_date = get_date(self.subfolders)
                    self.data[f'{day_name}'][f'{sess_name}']['date'] = cr_date
                    self.data[f'{day_name}'][f'{sess_name}']['param'] = copy.deepcopy(self.par)

            elif self.ch_idx == 2:  # folder contains M days with N sessions with 12 ori
                for day in [f.path for f in os.scandir(self.folder) if f.is_dir()]:
                    day_name = day.split("\\")[-1]
                    self.data[f'{day_name}'] = {}
                    print(f'Day: {day}')
                    for sess in [f.path for f in os.scandir(day) if f.is_dir() and len(next(os.walk(f))[1]) == 12]:
                        self.res = None
                        sess_name = sess.split("\\")[-1]
                        self.par['day'] = day_name
                        self.par['name'] = sess_name
                        self.data[f'{day_name}'][f'{sess_name}'] = {}
                        self.data[f'{day_name}'][f'{sess_name}']['path'] = sess
                        self.subfolders = [f.path for f in os.scandir(sess) if f.is_dir()]
                        cr_date = get_date(self.subfolders)
                        self.data[f'{day_name}'][f'{sess_name}']['date'] = cr_date
                        self.data[f'{day_name}'][f'{sess_name}']['param'] = copy.deepcopy(self.par)

            #  Defining general number of sessions to display progress bar
            self.num_sess = sum(len(v) for v in self.data.values())
            print(f"Number of sessions: {self.num_sess}")
            self.compute_ori()
        else:
            self.message.setText("Please, select a folder with data files")

    def compute_ori(self):
        print("Compute_ori")
        self.pbar.setFormat('Reading data...')
        self.pbar.setValue(int(round(100 * self.count / self.num_sess)))  # Displaying progress of computation
        QApplication.processEvents()  # Updating progress bar
        # Переопределение self.subfolders здесь как раз обеспечивает рекурсивную обработку данных (вместо цикла)
        day_name = list(self.data.keys())[self.cur_day]
        self.sess_name = list(self.data[day_name].keys())[self.cur_sess]
        sess_path = self.data[day_name][self.sess_name]['path']
        self.subfolders = [f.path for f in os.scandir(sess_path) if f.is_dir()]
        print(f"  cur_day: {self.cur_day} ({day_name}), cur_sess: {self.cur_sess} ({self.sess_name})")

        # Multiprocessing:
        stack12ori = []
        subfolders = self.subfolders
        filtr = self.filter  # redefining to local var because multiprocessing can't iter PyQt objects (self.)
        with multiprocessing.Pool() as pool:  # Reading tiffs
            items = [(subfolders[ori], ori, filtr) for ori in range(12)]
            for stack_arr, ori in pool.starmap(read_tiffs, items):
                stack12ori.append(stack_arr)
        self.stack12ori = np.array(stack12ori)  # [12 ori, frames, x, y]

        if self.stack12ori.shape[2] % self.wind_size != 0:  # If shape isn't dividable, change square size
            self.wind_size = find_devisor(self.stack12ori.shape[2], self.wind_size)
            self.sq.setText(f"{self.wind_size}")
        self.h = int(self.stack12ori.shape[2] / self.wind_size)  # Height of an image
        self.w = int(self.stack12ori.shape[3] / self.wind_size)  # Width of an image
        self.trial_len = int(self.stack12ori.shape[1])  # Number of frames

        # Performing alignment:
        if self.ref_stack is None:  # Make the first session a reference session
            self.ref_stack = np.array(np.max(np.max(self.stack12ori, axis=0), axis=0))
            self.ref_name = self.sess_name
            print(f"  Reference stack, shape: {self.ref_stack.shape}")
            self.compute_squares(self.stack12ori)
        else:
            self.align(self.ref_stack, self.stack12ori)

    # ALIGNMENT FUNCTIONS ------------------------------------------------------------------------- ALIGNMENT FUNCTIONS
    # Establish plot with overlapping images and interactive interface (sliders):
    def align(self, reference, image1):  # image1: [12 ori, frames, x, y]
        print("  Aligning...")
        self.pbar.setFormat('Shifting the stack...')
        self.pbar.setValue(int(round(100 * self.count / self.num_sess)))  # Displaying progress of computation
        QApplication.processEvents()  # Updating progress bar
        plt.ion()

        self.fig_ali = plt.figure()
        self.ax_ali = self.fig_ali.add_subplot(111)

        self.fig_ali.subplots_adjust(bottom=0.35)
        self.image1 = np.max(image1, axis=0)  # self.image1: [frames, x, y]
        self.image1 = np.array(np.max(self.image1, axis=0)) / np.max(self.image1)  # self.image1: [x, y], normalized
        self.reference = np.array(reference / np.max(reference))  # Reference normalization
        (h, w) = reference.shape
        self.h_low = int(round((2 ** 0.5 - 1) * h / 2))
        self.h_high = int(round((2 ** 0.5 + 1) * h / 2))
        self.w_low = int(round((2 ** 0.5 - 1) * w / 2))
        self.w_high = int(round((2 ** 0.5 + 1) * w / 2))

        self.shift = {'x': 0, 'y': 0, 'deg': 0}

        self.img_ref = np.zeros((int(round(2 ** 0.5 * h)), int(round(2 ** 0.5 * w)), 3))  # RGB
        self.img_ref[self.h_low:self.h_high, self.w_low:self.w_high, 0] = self.reference  # Red channel
        self.img_res = self.img_ref
        self.img_res[self.h_low:self.h_high, self.w_low:self.w_high, 1] = self.image1  # Green channel

        self.im = self.ax_ali.imshow(self.img_res)  # --------------------------------------- SET IMAGE
        self.ax_ali.set_xlim(self.w_low, self.w_high)
        self.ax_ali.set_ylim(self.h_high, self.h_low)

        # Slider for degree of rotation
        axsl = plt.axes((0.14, 0.18, 0.63, 0.03))
        self.sl_rot = Slider(
            ax=axsl,
            label='',
            valmin=-180,
            valmax=180,
            valinit=0,
            valstep=1,  # np.array(1, len(self.ori_list) + 1, 1),
            orientation="horizontal",
            color="lightgrey",
            initcolor="lightgrey",
            handle_style={'facecolor': 'white', 'edgecolor': '.25', 'size': 10}
        )
        axsl.add_artist(axsl.xaxis)
        sl_xticks = np.arange(-180, 181, 30)
        axsl.set_xticks(sl_xticks)
        self.fig_ali.text(.87, .18, "Rotation", ha='center')

        # Slider for y
        self.fig_ali.subplots_adjust(bottom=0.25)
        axsl = plt.axes((0.14, 0.04, 0.63, 0.03))
        self.sl_y = Slider(
            ax=axsl,
            label='',
            valmin=-h,
            valmax=h,
            valinit=0,
            valstep=1,  # np.array(1, len(self.ori_list) + 1, 1),
            orientation="horizontal",
            color="lightgrey",
            initcolor="lightgrey",
            handle_style={'facecolor': 'white', 'edgecolor': '.25', 'size': 11}
        )
        axsl.add_artist(axsl.xaxis)
        axsl.locator_params(axis='x', nbins=12)  # Setting defined number of ticks
        self.fig_ali.text(.85, .1, "X", ha='center', fontsize=12)

        # Slider for x
        self.fig_ali.subplots_adjust(bottom=0.25)
        axsl = plt.axes((0.14, 0.1, 0.63, 0.03))
        self.sl_x = Slider(
            ax=axsl,
            label='',
            valmin=-w,
            valmax=w,
            valinit=0,
            valstep=1,  # np.array(1, len(self.ori_list) + 1, 1),
            orientation="horizontal",
            color="lightgrey",
            initcolor="lightgrey",
            handle_style={'facecolor': 'white', 'edgecolor': '.25', 'size': 11}
        )
        axsl.add_artist(axsl.xaxis)
        axsl.locator_params(axis='x', nbins=12)  # Setting defined number of ticks
        self.fig_ali.text(.85, .01, "Y", ha='center', fontsize=12)

        self.sl_rot.on_changed(self.display_ali)  # Calling function to change rotation
        self.sl_y.on_changed(self.display_ali)  # Calling function to change X
        self.sl_x.on_changed(self.display_ali)  # Calling function to change Y

        self.fig_ali.subplots_adjust(right=0.80)
        axes = plt.axes((0.82, 0.7, 0.15, 0.05))
        self.bnappl = Button(axes, 'Apply')  # Button for applying current shift to whole stack
        self.bnappl.on_clicked(self.get_shift)

        axes = plt.axes((0.82, 0.3, 0.15, 0.05))
        self.bnskip = Button(axes, 'Skip')  # Button to skip the stage of alignment and continue without any shift
        self.bnskip.on_clicked(self.skip)

        self.fig_ali.canvas.mpl_connect('key_press_event', self.onclick_key)
        self.fig_ali.canvas.mpl_connect('scroll_event', self.onclick_scroll)

        # Specifying title:
        axes = plt.axes((0.3, 0.9, 0.3, 0.05))
        axes.axis('off')
        axes.text(0, 0, f'{self.ref_name}', color='red', horizontalalignment='center', fontsize=12)
        axes = plt.axes((0.45, 0.9, 0.1, 0.05))
        axes.axis('off')
        axes.text(0, 0, 'vs.', horizontalalignment='center', fontsize=12)
        axes = plt.axes((0.6, 0.9, 0.3, 0.05))
        axes.axis('off')
        axes.text(0, 0, f'{self.sess_name}', color='green', horizontalalignment='center', fontsize=12)

        mngr = plt.get_current_fig_manager()
        mngr.window.setGeometry(round(self.screen_width / 2) - 300, round(self.screen_height / 2) - 350, 600, 700)

        if self.align_values is not None:  # Use values from txt file if present
            day_name = list(self.data.keys())[self.cur_day]
            idx = np.where((self.align_values[:, 0] == day_name) & (self.align_values[:, 1] == self.sess_name))[0]
            if len(idx) > 0:  # If align values for current trial are present
                sess_shift = self.align_values[idx, :][0]
                self.sl_x.set_val(int(sess_shift[2]))
                self.sl_y.set_val(int(sess_shift[3]))
                self.sl_rot.set_val(int(sess_shift[4]))
                print(f"  Values were found: {self.shift}")

    def display_ali(self, val):  # Retrieve shift values from sliders and Update the plot
        self.shift['deg'] = int(self.sl_rot.val)
        self.shift['x'] = int(self.sl_x.val)
        self.shift['y'] = int(self.sl_y.val)
        x = self.shift['x']
        y = self.shift['y']
        (h, w) = self.reference.shape
        img1 = np.zeros((int(round(2 ** 0.5 * h)), int(round(2 ** 0.5 * w))))
        img1[self.h_low:self.h_high, self.w_low:self.w_high] = self.image1

        img1 = rotate(img1, angle=self.shift['deg'], reshape=False, order=0)  # Rotating initial image
        img1 = sh(img1, (y, x), order=0)  # Shifting initial image ([y, x] because height is 0-axis and width is 1-st.

        self.img_res = self.img_ref
        self.img_res[:, :, 1] = img1
        self.im.set_data(self.img_res)  # Updating plot
        self.fig_ali.canvas.draw()

    def skip(self, val):
        # Shape of self.image1_init: [12 ori, frames x, y]
        self.compute_squares(self.stack12ori)

    def get_shift(self, val):  # Return the output
        print(f"    GET SHIFT: {self.shift}")
        plt.close(self.fig_ali)
        deg = self.shift['deg']
        x = self.shift['x']
        y = self.shift['y']

        # Filling the data variable:
        day_name = list(self.data.keys())[self.cur_day]
        sess_name = list(self.data[day_name].keys())[self.cur_sess]
        self.data[f'{day_name}'][f'{sess_name}']['param']['Shift_Rot'] = copy.deepcopy(deg)
        self.data[f'{day_name}'][f'{sess_name}']['param']['Shift_X'] = copy.deepcopy(x)
        self.data[f'{day_name}'][f'{sess_name}']['param']['Shift_Y'] = copy.deepcopy(y)

        (h, w) = self.ref_stack.shape
        h_low = int(round((2 ** 0.5 - 1) * h / 2))
        h_high = int(round((2 ** 0.5 + 1) * h / 2))
        w_low = int(round((2 ** 0.5 - 1) * w / 2))
        w_high = int(round((2 ** 0.5 + 1) * w / 2))
        # Dimensions: [ori, frames, x, y]
        res = np.zeros(
            (self.stack12ori.shape[0], self.stack12ori.shape[1], int(round(2 ** 0.5 * h)), int(round(2 ** 0.5 * w))))
        res[:, :, h_low:h_high, w_low:w_high] = self.stack12ori

        with multiprocessing.Pool() as pool:  # Shifting the whole stack
            items = [(res[ori, :, :, :], ori, deg, x, y) for ori in range(12)]
            for result, ori in pool.starmap(shift_rot, items):
                res[ori, :, :, :] = result
        res = res[:, :, h_low:h_high, w_low:w_high]  # res: [12 ori, frames y, x]
        # print()
        self.compute_squares(res)

    def compute_squares(self, res):
        print("      Compute_squares")
        self.pbar.setFormat('Computing squares...')
        self.pbar.setValue(int(round(100 * self.count / self.num_sess)))  # Displaying progress of computation
        QApplication.processEvents()

        plt.close(self.fig_ali)
        if self.res is None:
            # Creating res variable for collecting traces of each ROI for each orientation
            self.res = np.zeros((self.h, self.w, len(self.ori_list), self.trial_len))
            self.matrix = np.zeros((self.h, self.w, len(self.ori_list) + 1))
            self.matrix_180 = np.zeros((self.h, self.w, 6))

        # param = [h, w, trial_len, filt_time, wind_size, st_lat_fr, windbrdr]
        param = [self.h, self.w, self.trial_len, self.filt_time, self.wind_size, self.st_lat_fr, self.windbrdr]

        with multiprocessing.Pool() as pool:  # Computing squares using multiprocessing
            items = [(res[ori, :, :, :], ori, param) for ori in range(12)]
            for result, matrix, ori in pool.starmap(comp_sq, items):
                self.res[:, :, ori, :] = result  # [Y, X, 12 ori, frames]
                self.matrix[:, :, ori] = matrix

        # Sorting variables by orientation + filling matrix180 variable:
        self.res_sort = np.zeros(self.res.shape)
        self.matrix_sort = np.zeros(self.matrix.shape)
        matrix180 = np.zeros(self.matrix.shape)  # only 1st 6 val are numbers, other - NaNs
        for ori in range(len(self.ori_list)):
            init_idx = self.initial_ori.index(self.ori_list[ori])  # Idx of curr ori in initial sequence
            idx180 = ori % 6  # Remainder is the desired idx for 180-degree plot
            self.matrix_sort[:, :, ori] = self.matrix[:, :, init_idx]
            self.res_sort[:, :, ori, :] = self.res[:, :, init_idx, :]
            matrix180[:, :, idx180] += self.matrix_sort[:, :, ori]  # Here matrix_sort at cur ori is already sorted
        self.matrix_sort[:, :, -1] = self.matrix_sort[:, :, 0]  # Add a row = to the 1st to close the curve
        matrix180 /= 2  # Because there were 2 directions for each ori (getting average over them)
        # matrix180[:, :, 6] = matrix180[:, :, 0]  # Copying 1st value to better outlook
        matrix180[:, :, 6:] = np.nan  # Filling the residuary values with NaNs

        # Computing vector sum of responses to all ori to get preferred ori-2:

        # Removing negative values by adding const (min among ori-s) to all values in each ROI:
        m = np.min(self.matrix_sort[:, :, :-1], axis=2)
        m[np.where(m > 0)] = 0  # Not to modify ROI-s where there's no negative values to any ori
        m = np.broadcast_to(m[..., None], self.matrix_sort.shape[:2] + (12,))
        matrix_corr = copy.deepcopy(self.matrix_sort[:, :, :-1]) - m  # minus negative values = plussing
        print("      neg before: ", np.sum(self.matrix_sort < 0), "neg after: ", np.sum(matrix_corr < 0))

        # Vertical vector of complex exp of all angles (doubled for ori index), in radians:
        ori_exp = np.exp(1j * 2 * self.rad[:-1]).T
        dir_exp = np.exp(1j * self.rad[:-1]).T  # For direction selectivity index angles aren't doubled

        # Normalized vector sum of responses to all ori for each ROI (shape: [X, Y]):
        sum_vector_ori = np.sum(self.mult_along_axis(matrix_corr, ori_exp, 2), axis=2) / np.sum(matrix_corr, axis=2)
        sum_vector_dir = np.sum(self.mult_along_axis(matrix_corr, dir_exp, 2), axis=2) / np.sum(matrix_corr, axis=2)

        # Remainder of division [angle % 2*Pi] converts angels from [-Pi;Pi] to [0;2*Pi]
        self.pref_ori = np.mod(np.arctan2(sum_vector_ori.imag, sum_vector_ori.real), 2 * np.pi) / 2
        self.pref_dir = np.mod(np.arctan2(sum_vector_dir.imag, sum_vector_dir.real), 2 * np.pi)  # arctan2 -> [-Pi; Pi]

        # Len of sum vector reflexes orientation selectivity:
        matrix_OI = 100 * ((sum_vector_ori.real ** 2 + sum_vector_ori.imag ** 2) ** 0.5)
        matrix_DI = 100 * ((sum_vector_dir.real ** 2 + sum_vector_dir.imag ** 2) ** 0.5)

        # Computing target selectivity index:
        # TI = 100*(Rtarg + Rnull - (Rorth+ + Rorth-)) / (Rtarg + Rnull + Rorth+ + Rorth-)
        targ_idx = self.ori_list.index(int(self.targ_ori.toPlainText()))  # index of target ori in ori list
        null_idx = (targ_idx + 6) % 12  # index of opposite ori to target ori in ori list
        orth_pos_idx = (targ_idx + 3) % 12  # index of higher orthogonal ori to target ori in ori list
        orth_neg_idx = (targ_idx - 3) % 12  # index of lower orthogonal ori to target ori in ori list
        targ = matrix_corr[:, :, targ_idx]
        null = matrix_corr[:, :, null_idx]
        orth_pos = matrix_corr[:, :, orth_pos_idx]
        orth_neg = matrix_corr[:, :, orth_neg_idx]
        matrix_TI = 100 * (targ + null - (orth_pos + orth_neg)) / (targ + null + orth_pos + orth_neg)

        # Filling the data variable:
        day_name = list(self.data.keys())[self.cur_day]
        sess_name = list(self.data[day_name].keys())[self.cur_sess]
        self.data[f'{day_name}'][f'{sess_name}']['res'] = np.array(self.res_sort)  # [X, Y, 12ori, frames]
        self.data[f'{day_name}'][f'{sess_name}']['matrix'] = np.array(self.matrix_sort)  # [X, Y, 13ori] !
        self.data[f'{day_name}'][f'{sess_name}']['matrix_corr'] = np.array(matrix_corr)  # [X, Y, 12ori]
        self.data[f'{day_name}'][f'{sess_name}']['matrix180'] = np.array(matrix180)  # [X, Y, 6ori]
        maps = np.zeros((self.h, self.w, 5))  # 'maps': [X, Y, [pref_ori, pref_dir, OI, DI, TI]]
        maps[:, :, 0] = self.pref_ori  # [X, Y]
        maps[:, :, 1] = self.pref_dir  # [X, Y]
        maps[:, :, 2] = matrix_OI  # [X, Y]
        maps[:, :, 3] = matrix_DI  # [X, Y]
        maps[:, :, 4] = matrix_TI  # [X, Y]
        self.data[f'{day_name}'][f'{sess_name}']['maps'] = np.array(maps)

        self.count += 1
        self.pbar.setValue(int(round(100 * self.count / self.num_sess)))  # Displaying progress of computation
        QApplication.processEvents()

        # Calling recursive processing over dict self.data:
        if self.cur_sess + 2 <= len(self.data[day_name].keys()):
            self.cur_sess += 1
            self.compute_ori()
        else:
            if self.cur_day + 2 <= len(self.data.keys()):
                self.cur_day += 1
                self.cur_sess = 0
                self.compute_ori()
            else:  # i.e. if all sessions in folder were computed:
                self.pbar.setFormat('')
                self.populateTree()
                print(self.rem_bottom(self.data, 1))
                # Set thrshld to min value, i.e. no initial masking by intensity
                self.mask_resp = np.min(self.matrix_sort[~np.isnan(self.matrix_sort)])
                self.make_figures()
                self.made = True  # Informs that computations are made
                self.display()  # Calling display() function to visualize data
                self.save_align()  # Save shift and rotation values for all probes in txt.file

    # VISUALIZATION FUNCTIONS ----------------------------------------------------------------- VISUALIZATION FUNCTIONS
    def make_figures(self):
        checked = self.get_checked_sess()
        if len(checked) < self.cursess + 1:
            self.cursess = len(checked) - 1
        cursess = checked[self.cursess]

        # Preparing for displaying plots:
        # For orientation preference:
        plt.ion()
        self.fig1 = plt.figure()
        self.ax = self.fig1.add_subplot(111, polar=True, theta_offset=np.pi / 2, theta_direction=-1)
        self.ax.set_thetagrids(np.arange(0.0, 360.0, 30.0))
        # Creating button for switching between 360 & 180-degree plotting mode:
        self.fig1.subplots_adjust(right=0.70)
        axes = plt.axes((0.76, 0.75, 0.2, 0.05))
        self.btn180 = Button(axes, '360' + u'\N{DEGREE SIGN}' + '/180' + u'\N{DEGREE SIGN}')
        self.btn180.on_clicked(self.switch180)
        mngr = plt.get_current_fig_manager()
        mngr.window.setGeometry(20, 45, 500, 550)

        # For response map and current cell position:
        plt.ion()
        self.fig2 = plt.figure()
        self.ax2 = self.fig2.add_subplot(111)
        self.cmap = colormaps["gray"]
        self.cmap.set_bad(color='yellow')
        whole_matrix = np.array(self.data[cursess[0]][cursess[1]]['matrix'])
        min_val = np.min(whole_matrix[~np.isnan(whole_matrix)])  # min value ignoring NaN-s
        max_val = np.max(whole_matrix[~np.isnan(whole_matrix)])  # max value ignoring NaN-s
        proj = whole_matrix[:, :, self.curr_ori - 1]
        proj[np.where(proj < self.mask_resp)] = min_val  # Mask values < than thresh
        self.im2 = self.ax2.imshow(proj,
                                   interpolation="none",
                                   cmap=self.cmap,
                                   vmin=min_val,
                                   vmax=max_val)
        cbaxes = self.fig2.add_axes((0.8, 0.25, 0.03, 0.637))
        self.cbar2 = self.fig2.colorbar(self.im2, cax=cbaxes, label="DeltaF/F", orientation="vertical")
        self.cbar2.mappable.set_clim(vmin=min_val, vmax=max_val)

        self.xtiks = np.arange(1, self.w, 2)
        self.xlabs = list(map(str, self.xtiks + 1))
        self.ytiks = np.arange(1, self.h, 2)
        self.ylabs = list(map(str, self.ytiks + 1))

        # Creating slider for ori selection:
        self.fig2.subplots_adjust(bottom=0.25)
        axsl = plt.axes((0.14, 0.15, 0.63, 0.03))
        self.sl_ori = Slider(
            ax=axsl,
            label='',
            valmin=0,
            valmax=330,
            valinit=self.ori_list[self.curr_ori - 1],
            valstep=30,  # np.array(1, len(self.ori_list) + 1, 1),
            orientation="horizontal",
            color="lightgrey",
            initcolor="lightgrey",
            handle_style={'facecolor': 'white', 'edgecolor': '.25', 'size': 10}
        )
        axsl.add_artist(axsl.xaxis)
        sl_xticks = np.arange(0, 331, 30)
        axsl.set_xticks(sl_xticks)
        self.fig2.text(.88, .15, "Orientation", ha='center')
        self.sl_ori.valtext.set_visible(False)  # Hide current value label
        self.sl_ori.on_changed(self.change_ori)  # Calling function to display data for new ori

        # Creating slider for ori thresholding intensity:
        self.fig2.subplots_adjust(right=0.75)
        axsl2 = plt.axes((0.76, 0.25, 0.03, 0.637))
        self.sl_thr = Slider(
            ax=axsl2,
            label='',
            valmin=min_val,
            valmax=max_val,
            valinit=self.mask_resp,
            valstep=0.001,  # np.array(1, len(self.ori_list) + 1, 1),
            orientation="vertical",
            color="black",
            handle_style={'facecolor': 'white', 'edgecolor': '.25', 'size': 10}
        )  # Slider doesn't need ticks since they're identical to colorbar
        self.sl_thr.valtext.set_visible(False)  # Hide current value label
        self.sl_thr.on_changed(self.change_thresh)  # Calling function to display data for new ori

        # Creating slider for session selection:
        self.fig2.subplots_adjust(bottom=0.25)
        self.axsl3 = plt.axes((0.14, 0.05, 0.63, 0.03))
        self.sl_sess = Slider(
            ax=self.axsl3,
            label='',
            valmin=1,
            valmax=1,
            valinit=1,
            valstep=1,
            orientation="horizontal",
            color="lightgrey",
            initcolor="lightgrey",
            handle_style={'facecolor': 'white', 'edgecolor': '.25', 'size': 10}
        )
        self.axsl3.add_artist(self.axsl3.xaxis)
        sl_xticks = np.arange(1, 2, 1)
        self.axsl3.set_xticks(sl_xticks)
        self.fig2.text(.88, .06, "Trial", ha='center')
        self.sl_sess.valtext.set_visible(False)  # Hide current value label
        self.sl_sess.on_changed(self.change_sess)  # Calling function to display data for new probe

        mngr = plt.get_current_fig_manager()
        mngr.window.setGeometry(530, 45, 520, 550)
        self.fig2.canvas.mpl_connect('button_press_event', self.onclick)  # Picking cell directly from the image

    def display(self):  # Updating plots for pref. ori, curr. cell & activity trace
        self.message.setText("")
        if self.made:
            # Displaying orientation preference:
            checked = self.get_checked_sess()
            self.ax.clear()
            self.ax.plot(self.rads, self.radius0, 'black', linewidth=0.5)  # Plotting the circle at 0
            titl_act = ''
            for chs in checked:
                print(f"Checked: {chs}")
                session = copy.deepcopy(self.data[chs[0]][chs[1]])
                matrix = session[self.disp180][self.curr_roi[0], self.curr_roi[1], :]  # self.disp180's the key for dict
                p = self.ax.plot(self.rad, matrix, label=chs[1])
                last_color = p[-1].get_color()  # Color of last plotted curve
                # 'maps': [X, Y, [pref_ori, pref_dir, OI, DI, TI]]
                oi = round(session['maps'][self.curr_roi[0], self.curr_roi[1], 2], 1)
                di = round(session['maps'][self.curr_roi[0], self.curr_roi[1], 3], 1)
                if session.get('res_std') is not None:  # i.e., if cur sess is averaged one:
                    # Adding STD area to ori preference plot:
                    matrix_std = session['matrix_std']
                    matrix_std_h = matrix + matrix_std[self.curr_roi[0], self.curr_roi[1], :]  # Higher boarder
                    matrix_std_l = matrix - matrix_std[self.curr_roi[0], self.curr_roi[1], :]  # Lower boarder
                    light_color = self.lightness(last_color, 1.5)  # Light version of last col for std interval
                    self.ax.plot(self.rad, matrix_std_h, color=light_color, alpha=0.8)  # '#88aafc'
                    self.ax.plot(self.rad, matrix_std_l, color=light_color, alpha=0.8)
                    self.ax.fill_between(self.rad, matrix_std_l, matrix_std_h, facecolor=light_color, alpha=0.5)
                    # Changing titles:
                    titl_ori = f"Tuning curve. Mean +/- STD. ROI [{self.curr_roi[0] + 1};" \
                               f"{self.curr_roi[1] + 1}]\nOI {oi}% DI {di}%"
                    titl_act = f"ROI [{self.curr_roi[0] + 1};{self.curr_roi[1] + 1}]. Mean +/- STD"
                else:
                    titl_ori = f"Tuning curve. ROI [{self.curr_roi[0] + 1};{self.curr_roi[1] + 1}]" \
                               f"\nOI {oi}% DI {di}%"
                    titl_act = f"ROI [{self.curr_roi[0] + 1};{self.curr_roi[1] + 1}]"
                y_lim = self.ax.get_ylim()
                # 'maps': [X, Y, [pref_ori, pref_dir, OI, DI, TI]]
                if self.disp180 == 'matrix180':  # Displaying preferred ori:
                    self.ax.vlines(session['maps'][self.curr_roi[0], self.curr_roi[1], 0], y_lim[0], y_lim[1],
                                   linestyles="dotted", colors=last_color)
                else:  # Displaying preferred direction:
                    self.ax.vlines(session['maps'][self.curr_roi[0], self.curr_roi[1], 1], y_lim[0], y_lim[1],
                                   linestyles="dotted", colors=last_color)
                # self.ax.set_ylim(y_lim)

                self.ax.set_title(label=titl_ori, fontsize=13, y=1.09)
            self.fig1.canvas.draw()
            self.ax.set_thetagrids(np.arange(0.0, 360.0, 30.0))
            self.fig1.subplots_adjust(top=0.9)
            self.fig1.legend(loc='outside right upper')

            # Lightning current cell
            self.ax2.clear()
            if len(checked) < self.cursess + 1:
                self.cursess = len(checked) - 1
            cursess = checked[self.cursess]  # Current trial among all checked is selected with slider self.sl_sess
            whole_matrix = np.array(self.data[cursess[0]][cursess[1]]['matrix'])
            min_val = np.min(whole_matrix[~np.isnan(whole_matrix)])  # min value ignorring NaN-s
            max_val = np.max(whole_matrix[~np.isnan(whole_matrix)])  # max value ignorring NaN-s
            proj = np.array(whole_matrix[:, :, self.curr_ori - 1])
            rect = patches.Rectangle((self.curr_roi[1] - 0.5, self.curr_roi[0] - 0.5), 1, 1, linewidth=2,
                                     edgecolor='r', facecolor="none")
            # Specifying axis ticks:
            self.ax2.set_xticks(self.xtiks, labels=self.xlabs)
            self.ax2.set_yticks(self.ytiks, labels=self.ylabs)

            self.ax2.imshow(proj,
                            interpolation="none",
                            cmap=self.cmap,
                            vmin=self.mask_resp,
                            vmax=max_val
                            )
            self.ax2.add_patch(rect)
            self.cbar2.mappable.set_clim(vmin=min_val, vmax=max_val)  # Updating color bar limits
            cursess_nam = f"{cursess[0]}/{cursess[1]}"
            titl_map_resp = f"ROI: [{self.curr_roi[0] + 1}; {self.curr_roi[1] + 1}] {cursess_nam}"
            self.ax2.set_title(label=titl_map_resp, fontsize=13)

            # Displaying activity trace:
            if self.activity:
                self.ax4.clear()
                trial_x = range(self.trial_len)  # Values of x-axes
                trial_y = None
                max_values = []  # List of max values of all traces
                min_values = []
                res = np.array(self.data[cursess[0]][cursess[1]]['res'])
                pref_ori = whole_matrix[self.curr_roi[0], self.curr_roi[1], :-1].argmax(axis=0) + 1
                if self.check.get_status()[pref_ori - 1] is False:
                    self.check.eventson = False  # Not to call add_ori func, which calls display again
                    self.check.set_active(pref_ori - 1)  # In any case, display trace for preferred ori
                for ori in [i for i, val in enumerate(self.check.get_status()) if
                            val is True]:  # in indices of checked ori
                    trial_y = res[self.curr_roi[0], self.curr_roi[1], ori, :]  # Here ori is idx!
                    max_values.append(max(trial_y))
                    min_values.append(min(trial_y))
                    self.ax4.plot(trial_x, trial_y, self.clrs[ori + 2])  # +2 because 1st color is black, 2nd is white
                    if self.data[cursess[0]][cursess[1]].get('res_std') is not None:  # i.e. if cur sess's averaged one
                        # Adding STD area to ori preference plot:
                        res_std = np.array(self.data[cursess[0]][cursess[1]]['res_std'])
                        # Adding STD area to activity traces plot:
                        trial_y_std_h = trial_y + res_std[self.curr_roi[0], self.curr_roi[1], ori, :]
                        trial_y_std_l = trial_y - res_std[self.curr_roi[0], self.curr_roi[1], ori, :]
                        light_color = self.lightness(self.clrs[ori + 2], 1.5)
                        self.ax4.plot(trial_x, trial_y_std_h, color=light_color, alpha=0.4)
                        self.ax4.plot(trial_x, trial_y_std_l, color=light_color, alpha=0.4)
                        self.ax4.fill_between(trial_x, trial_y_std_l, trial_y_std_h, facecolor=light_color, alpha=0.6)
                self.ax4.hlines(y=0, xmin=0, xmax=self.trial_len, linewidth=0.5, color='black')  # Mark zero
                self.ax4.set_xlim(0, self.trial_len)
                self.ax4.set(xlabel="Time, frames", ylabel="Delta F/F")
                # Painting background reflexing poststim interval:
                ymin, ymax = self.ax4.get_ylim()
                y = np.arange(ymin, ymax, 0.001)  # Limits of colored area
                self.ax4.fill_betweenx(y, self.st_lat_fr + 1, round(self.windbrdr), facecolor='green', alpha=0.3)
                self.ax4.set_ylim(ymin, ymax)
                self.ax4.set_title(label=titl_act, fontsize=13)
                self.check.eventson = True  # Activating callback from slider controlling poststim interval
        else:
            if len(self.folder) > 0:
                self.message.setText("Please, press [Compute] at first")
            else:
                self.message.setText("Please, select a folder and press [Compute]")

    def disp_activ(self):  # Displaying activity trace plot
        if self.made:
            checked = self.get_checked_sess()
            cursess = checked[self.cursess]
            matrix_sort = np.array(self.data[cursess[0]][cursess[1]]['matrix'])

            # For activity trace of the cell:
            plt.ion()
            self.fig4 = plt.figure()
            self.ax4 = self.fig4.add_subplot(111)
            pref_ori = matrix_sort[self.curr_roi[0], self.curr_roi[1], :].argmax(axis=0) + 1

            # Creating check boxes:
            self.fig4.subplots_adjust(right=0.85)
            rax = self.fig4.add_axes((0.87, 0.2, 0.07, 0.68))
            labels = [str(int(ori)) for ori in self.ori_list]
            visibility = [pref_ori == ori for ori in range(1, len(self.ori_list) + 1)]  # Gives bool list of
            # whether each ori is present in preferred ori list
            self.check = CheckButtons(rax, labels, visibility)
            self.check.set_check_props({'facecolor': self.clrs[2:]})
            self.check.set_label_props({'color': self.clrs[2:]})  # 2: because 1st & 2nd colors are black and white
            self.check.set_frame_props({'edgecolor': self.clrs[2:]})
            self.check.on_clicked(self.add_ori)

            # Creating slider for upper boarder of poststim interval:
            self.fig4.subplots_adjust(bottom=0.3)
            axsl = plt.axes((0.12, 0.04, 0.73, 0.03))
            self.sl = Slider(
                ax=axsl,
                label='End',
                valmin=0,
                valmax=self.trial_len,
                valinit=self.windbrdr,
                valstep=1,
                orientation="horizontal",
                color="lightgrey",
                initcolor="lightgrey",
                handle_style={'facecolor': 'white', 'edgecolor': '.25', 'size': 10}
            )
            self.sl.on_changed(self.poststim)  # Calling function to recompute data with new poststim interval

            # Creating slider for lower boarder of poststim interval:
            axsl2 = plt.axes((0.12, 0.1, 0.73, 0.03))
            self.sl2 = Slider(
                ax=axsl2,
                label='Start',
                valmin=1,
                valmax=self.trial_len,
                valinit=self.st_lat_fr,
                valstep=1,
                orientation="horizontal",
                color="lightgrey",
                initcolor="lightgrey",
                handle_style={'facecolor': 'white', 'edgecolor': '.25', 'size': 10}
            )
            self.sl2.on_changed(self.poststim)  # Calling function to recompute data with new poststim interval

            # Specify place where a figure appears and its geometry: (dist from left, dist from top, width, height)
            mngr = plt.get_current_fig_manager()
            mngr.window.setGeometry(20, 645, 640, 320)
            self.activity = True  # Informs that activity plot is made
            self.display()  # Updating plot
        else:
            if len(self.folder) > 0:
                self.message.setText("Please, press [Compute] at first")
            else:
                self.message.setText("Please, select a folder and press [Compute]")

    def poststim(self, arg):  # Changing the higher boarder of poststimulus interval
        np.seterr(divide='ignore', invalid='ignore')
        if self.sl.val - self.sl2.val < 2:  # Upper limit must be greater than lower limit
            self.sl.set_val(self.sl2.val + 2)
        self.windbrdr = round(self.sl.val)  # Upper boarder of poststim interval
        self.st_lat_fr = round(self.sl2.val)  # Lower boarder of poststim interval
        # Recalculating self.matrix_sort variable with updated self.windbrdr (higher boarder of poststim interval):
        for ori in range(len(self.ori_list)):
            for i in range(self.h):
                for j in range(self.w):
                    trace_delta = self.res_sort[i, j, ori, :]
                    val = np.mean(trace_delta[self.st_lat_fr:self.windbrdr])  # It's already normalized by prestim
                    # since intensity was converted into deltaF/F
                    self.matrix_sort[i, j, ori] = val
        self.matrix_sort[:, :, -1] = self.matrix_sort[:, :, 0]  # Adding a row equal to the 1st to close the curve
        if not self.aver:
            sess_idx = self.sess_names.index(self.name)
            self.sess_matrix[sess_idx] = self.matrix_sort  # Updating data in session list
        self.display()  # Updating plots

    def show(self):  # Display projection of all cells with coloring showing their preferred orientation
        if self.made:
            checked = self.get_checked_sess()
            cursess = checked[self.cursess]
            if self.disp180 == 'matrix':  # 'maps': [X, Y, [pref_ori, pref_dir, OI, DI, TI]]
                proj = 180 * self.data[cursess[0]][cursess[1]]['maps'][:, :, 1] / np.pi
                type_idx = 'direction'
            else:
                proj = 180 * self.data[cursess[0]][cursess[1]]['maps'][:, :, 0] / np.pi
                type_idx = 'orientation'
            max_proj = np.max(self.data[cursess[0]][cursess[1]]['matrix'], axis=2)  # Proj of max response among ori-s
            proj[np.where(max_proj < self.mask_resp)] = -30  # Filtering the ORIs with too low response
            self.fig3 = plt.figure()
            self.ax3 = self.fig3.add_subplot(111)
            colormap = colors.ListedColormap(self.clrs)
            # Specifying axis ticks:
            self.ax3.set_xticks(self.xtiks, labels=self.xlabs)
            self.ax3.set_yticks(self.ytiks, labels=self.ylabs)
            im = self.ax3.imshow(proj, interpolation="nearest", origin="upper", cmap=colormap)
            cbar = self.fig3.colorbar(im, label="Preferred orientation", orientation="vertical")
            cbar.mappable.set_clim(vmin=-75, vmax=345)
            cbar.set_ticks([-30] + self.ori_list)
            cbar.set_ticklabels(["No resp"] + [str(int(i)) for i in self.ori_list])
            self.ax3.set_title(label=f"{cursess[0]}/{cursess[1]}\nPreferred {type_idx}",
                               fontsize=13)
            # Specify place where a figure appears and its geometry: (dist from left, dist from top, width, height)
            mngr = plt.get_current_fig_manager()
            mngr.window.setGeometry(1300, 480, 520, 520)
            self.fig3.canvas.mpl_connect('button_press_event', self.onclick)  # Picking cell directly from the image
            self.gen_img_made = True
        else:
            if len(self.folder) > 0:
                self.message.setText("Please, press [Compute] at first")
            else:
                self.message.setText("Please, select a folder and press [Compute]")

    def upd_gen_img(self):
        if self.gen_img_made:  # Updating general image with new matrix_sort
            projec = np.zeros((self.h, self.w))
            max_proj = np.max(self.matrix_sort, axis=2)  # Projection of max response among orientations
            for ii in range(self.h):
                for jj in range(self.w):
                    # Getting idx of ori caused max response and finding corresponding ori (in degrees) in ori list:
                    print(f"i: {ii}, j: {jj}, argmax: {self.matrix_sort[ii, jj, :-1].argmax(axis=0)}")
                    projec[ii, jj] = self.ori_list[self.matrix_sort[ii, jj, :-1].argmax(axis=0)]
            colormap = colors.ListedColormap(self.clrs)
            projec[np.where(max_proj < self.mask_resp)] = -30  # Filtering the ORIs with too low response
            self.ax3.imshow(projec, interpolation="nearest", origin="upper", cmap=colormap)

    def ori_index(self):  # Display projection of all cells with coloring showing their orientation selectivity index
        if self.made:
            plt.ion()
            checked = self.get_checked_sess()
            cursess = checked[self.cursess]
            max_proj = np.max(self.data[cursess[0]][cursess[1]]['matrix'], axis=2)  # Proj of max response among ori-s
            print(f"cursess: {cursess}")
            # proj = self.data[cursess[0]][cursess[1]]['matrix_OI']
            proj = np.array(
                self.data[cursess[0]][cursess[1]]['maps'][:, :, 2])  # [X, Y, [pref_ori, pref_dir, OI, DI, TI]]
            proj[np.where(max_proj < self.mask_resp)] = np.min(proj[~np.isnan(proj)])  # FilterORIs with too low resp
            self.figOI = plt.figure()
            self.axOI = self.figOI.add_subplot(111)
            # Adding button to change OI/DI:
            self.figOI.subplots_adjust(bottom=0.25)
            axes = plt.axes([0.4, 0.1, 0.15, 0.05])
            self.btnOI = Button(axes, 'OI/DI')
            self.btnOI.on_clicked(self.changeOI)

            # Specifying axis ticks:
            self.axOI.set_xticks(self.xtiks, labels=self.xlabs)
            self.axOI.set_yticks(self.ytiks, labels=self.ylabs)
            self.imOI = self.axOI.imshow(proj,
                                         interpolation="none",
                                         origin="upper",
                                         cmap=self.cmap
                                         )
            cbar = self.figOI.colorbar(self.imOI, label="Selectivity, %", orientation="vertical")
            # cbar.mappable.set_clim(vmin=0, vmax=np.max(proj[~np.isnan(proj)]))
            self.axOI.set_title(label=f"Orientation selectivity\n{cursess[0]}/{cursess[1]}",
                                fontsize=13)
            # Specify place where a figure appears and its geometry: (dist from left, dist from top, width, height)
            mngr = plt.get_current_fig_manager()
            mngr.window.setGeometry(1280, 500, 520, 520)
            self.figOI.canvas.mpl_connect('button_press_event', self.onclick)  # Picking cell directly from the image
        else:
            if len(self.folder) > 0:
                self.message.setText("Please, press [Compute] at first")
            else:
                self.message.setText("Please, select a folder and press [Compute]")

    def average(self):
        print("Average")
        checked = self.get_checked_sess()
        print(f"  Checked sessions: {checked}")
        selected_days = [chs[0] for chs in checked]
        selected_sess = [chs[1] for chs in checked]
        selected_res = []
        selected_m = []
        if self.made and len(checked) > 1:
            self.aver = True
            plt.close('all')  # Closing all figures
            for chs in checked:
                selected_res.append(self.data[chs[0]][chs[1]]['res'])  # "res_sort"s of selected sessions
                selected_m.append(self.data[chs[0]][chs[1]]['matrix'])  # "matrix_sort"s of selected sessions
            res_arr = np.array(selected_res)
            res_sort = np.mean(res_arr, axis=0)  # Averaging traces over sessions
            res_std = np.std(res_arr, axis=0)  # Evaluating standard deviation for traces
            matrix_arr = np.array(selected_m)
            matrix_sort = np.mean(matrix_arr, axis=0)  # Averaging ori preference over sessions
            matrix_std = np.std(matrix_arr, axis=0)  # Evaluating standard deviation for traces

            matrix180 = np.zeros(matrix_sort.shape)  # only 1st 6 val are numbers, other - NaNs
            for ori in range(len(self.ori_list)):
                idx180 = ori % 6  # Remainder is the desired idx for 180-degree plot
                matrix180[:, :, idx180] += matrix_sort[:, :, ori]  # Here matrix_sort at cur ori is already sorted
            matrix180 /= 2  # Because there were 2 directions for each ori (getting average over them)
            matrix180[:, :, 6:] = np.nan  # Filling the residuary values with NaNs

            # Computing vector sum of responses to all ori to get preferred ori-2:
            # Removing negative values by adding const (min among ori-s) to all values in each ROI:
            m = np.min(matrix_sort[:, :, :-1], axis=2)
            m[np.where(m > 0)] = 0  # Not to modify ROI-s where there's no negative values to any ori
            m = np.broadcast_to(m[..., None], matrix_sort.shape[:2] + (12,))
            matrix_corr = copy.deepcopy(matrix_sort[:, :, :-1]) - m  # minus negative values = plussing
            print(f"  Average: neg before: {np.sum(matrix_sort < 0)} neg after: {np.sum(matrix_corr < 0)}")

            # Vertical vector of complex exp of all angles (doubled for ori index), in radians:
            ori_exp = np.exp(1j * 2 * self.rad[:-1]).T
            dir_exp = np.exp(1j * self.rad[:-1]).T  # For direction selectivity idx angles aren't doubled

            # Normalized vector sum of responses to all ori for each ROI:
            sum_vector_ori = np.sum(self.mult_along_axis(matrix_corr, ori_exp, 2), axis=2) / np.sum(
                matrix_corr, axis=2)  # [X, Y]
            sum_vector_dir = np.sum(self.mult_along_axis(matrix_corr, dir_exp, 2), axis=2) / np.sum(
                matrix_corr, axis=2)  # [X, Y]

            # Remainder of division [angle % 2*Pi] converts angels from [-Pi;Pi] to [0;2*Pi]
            pref_ori = np.mod(np.arctan2(sum_vector_ori.imag, sum_vector_ori.real), 2 * np.pi) / 2
            pref_dir = np.mod(np.arctan2(sum_vector_dir.imag, sum_vector_dir.real), 2 * np.pi)  # arctan2->[-Pi;Pi]

            # Len of sum vector reflexes orientation selectivity (shape: [X, Y]):
            matrix_OI = 100 * ((sum_vector_ori.real ** 2 + sum_vector_ori.imag ** 2) ** 0.5)
            matrix_DI = 100 * ((sum_vector_dir.real ** 2 + sum_vector_dir.imag ** 2) ** 0.5)

            # Computing target selectivity index:
            # TI = 100*(Rtarg + Rnull - (Rorth+ + Rorth-)) / (Rtarg + Rnull + Rorth+ + Rorth-)
            targ_idx = self.ori_list.index(int(self.targ_ori.toPlainText()))  # index of target ori in ori list
            null_idx = (targ_idx + 6) % 12  # index of opposite ori to target ori in ori list
            orth_pos_idx = (targ_idx + 3) % 12  # index of higher orthogonal ori to target ori in ori list
            orth_neg_idx = (targ_idx - 3) % 12  # index of lower orthogonal ori to target ori in ori list
            targ = matrix_corr[:, :, targ_idx]
            null = matrix_corr[:, :, null_idx]
            orth_pos = matrix_corr[:, :, orth_pos_idx]
            orth_neg = matrix_corr[:, :, orth_neg_idx]
            matrix_TI = 100 * (targ + null - (orth_pos + orth_neg)) / (targ + null + orth_pos + orth_neg)

            label = ', '.join(selected_sess)
            self.message_tab3.setText(f"Current session: Averaged {label}")
            self.name = "Averaged"

            # Filling self.data with new session (averaged)
            if selected_days.count(selected_days[0]) == len(selected_days):  # i.e., all sessions are from same day
                all_sess = self.data[selected_days[0]].keys()
                num_aver = len([s for s in all_sess if "Averaged" in s]) + 1  # How many averaged sessions in this day
                new_day_name = selected_days[0]
                new_name = 'Averaged' + str(num_aver)
            else:
                num_aver = len([s for s in self.data.keys() if "Averaged" in s]) + 1  # How many averaged sessions
                new_day_name = 'Averaged' + str(num_aver)
                new_name = 'Averaged'
                self.data[new_day_name] = {}  # Creating new day

            self.data[new_day_name][new_name] = {}  # Creating new session
            self.data[new_day_name][new_name]['res'] = np.array(res_sort)
            self.data[new_day_name][new_name]['res_std'] = np.array(res_std)
            self.data[new_day_name][new_name]['matrix'] = np.array(matrix_sort)
            self.data[new_day_name][new_name]['matrix180'] = np.array(matrix180)
            self.data[new_day_name][new_name]['aver_over'] = str(label)
            self.data[new_day_name][new_name]['matrix_corr'] = np.array(matrix_corr)
            self.data[new_day_name][new_name]['matrix_std'] = np.array(matrix_std)
            maps = np.zeros((self.h, self.w, 5))  # 'maps': [X, Y, [pref_ori, pref_dir, OI, DI, TI]]
            maps[:, :, 0] = pref_ori  # [X, Y]
            maps[:, :, 1] = pref_dir  # [X, Y]
            maps[:, :, 2] = matrix_OI  # [X, Y]
            maps[:, :, 3] = matrix_DI  # [X, Y]
            maps[:, :, 4] = matrix_TI  # [X, Y]
            self.data[new_day_name][new_name]['maps'] = np.array(maps)

            self.populateTree()
            self.make_figures()
            self.display()
        else:
            self.message_tab3.setText("Needs at least two sessions")

    def separate(self):
        print("Separate")
        if self.made:
            # Updating slider for sessions:
            checked = self.get_checked_sess()
            self.sl_sess.valmax = len(checked)
            self.sl_sess.ax.set_xlim(1, len(checked))
            sl_xticks = np.arange(1, len(checked) + 1, 1)
            self.axsl3.set_xticks(sl_xticks)

            # Updating slider for threshold (color bar will be updated in display() method)
            checked = self.get_checked_sess()
            if len(checked) < self.cursess + 1:
                self.cursess = len(checked) - 1
            cursess = checked[self.cursess]
            whole_matrix = np.array(self.data[cursess[0]][cursess[1]]['matrix'])
            min_val = np.min(whole_matrix[~np.isnan(whole_matrix)])  # min value ignoring NaN-s
            max_val = np.max(whole_matrix[~np.isnan(whole_matrix)])  # max value ignoring NaN-s
            self.sl_thr.valmin = min_val
            self.sl_thr.valmax = max_val
            self.sl_thr.ax.set_ylim(self.sl_thr.valmin, self.sl_thr.valmax)

            # Check if current thres exceeds current amplitude limits and correct if it does:
            if self.mask_resp < min_val:
                self.sl_thr.set_val(min_val)
            elif self.mask_resp > max_val:
                self.sl_thr.set_val(max_val)

            # cursess = checked[int(self.sl_sess.val) - 1]
            # self.message_tab3.setText(f"Current session: {cursess[0]}/{cursess[1]}")
            self.message_tab3.setText("")
            self.aver = False
            self.display()

    def statistic(self):  # Precomputing statistics for selected days and trials --> self.p_values, stat_aver, stat_resp
        checked = self.get_checked_sess()
        # self.checked_days = list(set([x[0] for x in checked]))
        self.checked_days = list(dict.fromkeys([x[0] for x in checked]))  # Remove duplicates, preserve order
        if self.made and len(self.checked_days) > 1:  # Days comparison makes sense only if there's 2 days or more
            print('Statistic')
            from scipy.stats import mannwhitneyu as mu
            # self.p_values: [X, Y, IdxType, Day1, Day2]   np.ones -> initially no significant values
            self.p_values = np.ones((self.h, self.w, 5, len(self.checked_days), len(self.checked_days)))
            self.stat_aver = np.zeros((2, self.h, self.w, 6, len(self.checked_days)))  # [[aver,STD], X,Y,IdxType,Days]
            self.stat_resp = np.zeros((2, self.h, self.w, 13, len(self.checked_days)))  # [[aver,STD], X,Y,12ori, Days]
            garnd_aver = np.zeros((self.h, self.w, len(checked)))  # Resp to pref ori in all days and trials
            # 'maps': [X, Y, [pref_ori, pref_dir, OI, DI, TI]]
            trial_count = 0
            for day1 in range(len(self.checked_days)):
                day1_name = self.checked_days[day1]
                print(f"Day1: {day1} {day1_name}")
                trials_day1 = [trial[1] for trial in checked if trial[0] == day1_name]  # [['day', 'trial'], [...]]
                day1_data = np.zeros((self.h, self.w, 6, len(trials_day1)))  # [X, Y, IdxType, trials]
                day1_resp = np.zeros((self.h, self.w, 13, len(trials_day1)))  # [X, Y, 12ori, trials] - tuning curve
                for trial in range(len(trials_day1)):
                    trial1_name = trials_day1[trial]
                    day1_data[:, :, 1:, trial] = copy.deepcopy(self.data[day1_name][trial1_name]['maps'])
                    day1_resp[:, :, :, trial] = copy.deepcopy(self.data[day1_name][trial1_name]['matrix'])
                    max_resp = np.max(self.data[day1_name][trial1_name]['matrix'], axis=2)  # matrix: [X, Y, 13ori]
                    day1_data[:, :, 0, trial] = np.array(max_resp)  # Saving resp ampl to pref ori
                    garnd_aver[:, :, trial_count] = np.array(max_resp)
                    trial_count += 1

                self.stat_aver[0, :, :, :, day1] = np.average(day1_data, axis=3)  # Aver over trls for each idx & ROI
                self.stat_aver[1, :, :, :, day1] = np.std(day1_data, axis=3)  # STD over trls for each idx & ROI

                self.stat_resp[0, :, :, :, day1] = np.average(day1_resp, axis=3)  # Aver over trls for each ori & ROI
                self.stat_resp[1, :, :, :, day1] = np.std(day1_resp, axis=3)  # STD over trls for each ori & ROI

                for day2 in range(day1 + 1, len(self.checked_days)):
                    day2_name = self.checked_days[day2]
                    print(f"    Day2: {day2} {day2_name}")
                    trials_day2 = [trial[1] for trial in checked if trial[0] == day2_name]
                    day2_data = np.zeros((self.h, self.w, 5, len(trials_day2)))  # [X, Y, IdxType, trials]
                    for trial in range(len(trials_day2)):
                        trial_name = trials_day2[trial]
                        day2_data[:, :, :, trial] = copy.deepcopy(self.data[day2_name][trial_name]['maps'])
                    for i in range(5):  # 'maps': [X, Y, [pref_ori, pref_dir, OI, DI, TI]]
                        s, self.p_values[:, :, i, day1, day2] = mu(day1_data[:, :, i, :], day2_data[:, :, i, :], axis=2)
            self.grand_aver = np.average(garnd_aver, axis=2)  # Average resp to pref ori over all days and trials

            plt.ion()
            self.fig_stat, self.ax_stat = plt.subplots(2, 3)  # For plotting curves
            self.fig_stat.tight_layout()  # Prevents plots overlap
            self.fig_pval, self.ax_pval = plt.subplots(2, 3)  # For tables of p-values
            self.fig_pval.tight_layout()
            # self.ax_stat[1, 2].set_visible(False)  # There's no 6th index to plot

            # For tuning curves of each day:
            days_per_row = 4  # Rounding to bigger:
            num_row = int(len(self.checked_days) / days_per_row) + (len(self.checked_days) % days_per_row > 0)
            num_col = days_per_row if len(self.checked_days) >= days_per_row else len(self.checked_days)
            self.fig_tun, ax_tun = plt.subplots(num_row, num_col, subplot_kw=dict(polar=True))
            if len(self.checked_days) > days_per_row:
                self.ax_tun = [ax_tun[i, j] for i in range(num_row) for j in range(num_col)]
            else:  # When there's only 1 row (< 4 days), two indicies [i, j] provoke IndexError.
                self.ax_tun = [ax_tun[j] for j in range(num_col)]
            for i in range(len(self.checked_days), len(self.ax_tun)):
                self.ax_tun[i].set_visible(False)  # Mask empty slots
            self.fig_tun.tight_layout()

            # For average resp to pref ori over all days and trials:
            self.fig_gr_aver, self.ax_gr_aver = plt.subplots()
            self.fig_gr_aver.subplots_adjust(right=0.75)
            self.fig_gr_aver.canvas.mpl_connect('button_press_event', self.onclick_stat)  # Pick cell directly from img
            proj = np.array(self.grand_aver)
            min_val = np.min(proj[~np.isnan(proj)])  # min value ignorring NaN-s
            max_val = np.max(proj[~np.isnan(proj)])  # max value ignorring NaN-s
            print(f"Grand_aver  min: {min_val}  max: {max_val}")
            im = self.ax_gr_aver.imshow(proj,  # Initial plotting is needed for color bar setting
                                        interpolation="none",
                                        cmap=self.cmap,
                                        vmin=min_val,
                                        vmax=max_val
                                        )
            cbaxes = self.fig_gr_aver.add_axes([0.8, 0.15, 0.03, 0.7])
            self.cb_dr_avr = self.fig_gr_aver.colorbar(im, cax=cbaxes, label="DeltaF/F", orientation="vertical")

            # Creating slider for ori thresholding intensity:
            self.fig_gr_aver.subplots_adjust(right=0.75)
            axsl2 = plt.axes((0.76, 0.15, 0.03, 0.7))
            self.sl_gr_aver = Slider(
                ax=axsl2,
                label='',
                valmin=min_val,
                valmax=max_val,
                valinit=self.ampl_thrs,
                valstep=0.001,  # np.array(1, len(self.ori_list) + 1, 1),
                orientation="vertical",
                color="black",
                handle_style={'facecolor': 'white', 'edgecolor': '.25', 'size': 10}
            )  # Slider doesn't need ticks since they're identical to colorbar
            self.sl_gr_aver.valtext.set_visible(False)  # Hide current value label
            self.sl_gr_aver.on_changed(self.change_thrs_gr_aver)  # Calling function to display data for new ori

            mngr = plt.get_current_fig_manager()
            mngr.window.setGeometry(20, 500, 520, 520)

            self.plot_statistics()
        else:
            self.message.setText("Needs at least 2 days with 2 trials selected")

    def plot_statistics(self):  # Plot statistics for selected days and trials
        print("Plot_statistics")
        if self.p_values is not None:
            day_lbl = list(self.checked_days)
            title_list = ["Response amplitude to pref. ori", "Preferred orientation, deg", "Preferred direction, deg",
                          "Orientation selectivity index, %", "Direction selectivity index, %",
                          "Target selectivity index, %"]

            # self.stat_aver: [[Aver, STD], X, Y, IdxType, Days]
            # self.p_values: [X, Y, IdxType, Day1, Day2]
            # IdxType: [pref_ori, pref_dir, OI, DI, TI]

            # Plotting indexes dynamics and p-value tables
            c = 0
            for i in range(2):
                for j in range(3):
                    self.ax_stat[i, j].clear()  # Clearing all subplots with curves
                    self.ax_pval[i, j].clear()  # Clearing all subplots with p-values

                    # Plotting dynamics of each index over days (5-c = index of desired idx type in stat_aver)
                    if 5 - c in [1, 2]:  # if pref. ori and dir are to be plotted - convert to degree:
                        val = np.rad2deg(self.stat_aver[0, self.cur_roi_stat[0], self.cur_roi_stat[1], 5 - c, :])
                        err = np.rad2deg(self.stat_aver[1, self.cur_roi_stat[0], self.cur_roi_stat[1], 5 - c, :])
                    else:
                        val = self.stat_aver[0, self.cur_roi_stat[0], self.cur_roi_stat[1], 5 - c, :]
                        err = self.stat_aver[1, self.cur_roi_stat[0], self.cur_roi_stat[1], 5 - c, :]
                    self.ax_stat[i, j].plot(day_lbl, val)
                    self.ax_stat[i, j].errorbar(day_lbl,  # Plot error bars (+/-STD)
                                                val,
                                                yerr=err,
                                                capsize=4,
                                                elinewidth=0.8,
                                                ecolor='black')
                    self.ax_stat[i, j].set_title(title_list[5 - c], fontsize=8)

                    if i + j != 3:  # 5th subplot is not needed, there is only 5 indexes (Pr.ori, Pr.dir, OI, DI, TI)
                        # Plotting p-values of each index between days (5 - c = index of desired idx type in stat_aver)
                        p_value = np.round(self.p_values[self.cur_roi_stat[0], self.cur_roi_stat[1], 4 - c, :, :],
                                           decimals=4)  # [Day1, Day2]
                        colors = np.full(p_value.shape, 'w', dtype='object')
                        lower_triang = np.tril_indices(p_value.shape[0])  # Get ind-s of lower triangle of matrix
                        colors[np.where(p_value < 0.05)] = 'y'  # Color significant values
                        # Cells in lower triangle of the matrix weren't modified and contain ones -> weren't colored
                        p_value = np.array(p_value, dtype=object)  # Convert p-val matrix to object (enable str vals)
                        p_value[lower_triang] = ''  # Assign empty str to all lower triangl cells of p-val matrix
                        table = self.ax_pval[i, j].table(cellText=p_value,
                                                         cellColours=colors,
                                                         rowLabels=day_lbl,
                                                         colWidths=[0.2] * p_value.shape[1],
                                                         colLabels=day_lbl,
                                                         loc='center')
                        table.auto_set_font_size(False)
                        table.set_fontsize(8)
                        # table.scale(1, 1.5)
                        table.auto_set_column_width([i for i in range(p_value.shape[0])])
                        self.ax_pval[i, j].axis('off')
                        self.ax_pval[i, j].set_title(title_list[5 - c])
                    c += 1

            # Plotting tuning curves:  self.stat_resp: [[aver,STD], X,Y, 12ori, Days]
            for day in range(len(self.checked_days)):
                self.ax_tun[day].clear()
                matrix = self.stat_resp[0, self.cur_roi_stat[0], self.cur_roi_stat[1], :, day]
                self.ax_tun[day].plot(self.rads, self.radius0, 'black', linewidth=0.5)  # Plotting the circle at 0
                p = self.ax_tun[day].plot(self.rad, matrix)
                last_color = p[-1].get_color()  # Color of last plotted curve
                oi = round(self.stat_aver[0, self.cur_roi_stat[0], self.cur_roi_stat[1], 2, day], 1)
                di = round(self.stat_aver[0, self.cur_roi_stat[0], self.cur_roi_stat[1], 3, day], 1)

                matrix_std = self.stat_resp[1, self.cur_roi_stat[0], self.cur_roi_stat[1], :, day]
                matrix_std_h = matrix + matrix_std  # Higher boarder
                matrix_std_l = matrix - matrix_std  # Lower boarder
                light_color = self.lightness(last_color, 1.5)  # Light version of last col for std interval
                self.ax_tun[day].plot(self.rad, matrix_std_h, color=light_color, alpha=0.8)  # '#88aafc'
                self.ax_tun[day].plot(self.rad, matrix_std_l, color=light_color, alpha=0.8)
                self.ax_tun[day].fill_between(self.rad, matrix_std_l, matrix_std_h, facecolor=light_color, alpha=0.5)
                self.ax_tun[day].set_thetagrids(np.arange(0.0, 360.0, 30.0))
                titl_ori = f"{self.checked_days[day]}\nOI {oi}% DI {di}%"
                self.ax_tun[day].set_title(label=titl_ori, fontsize=9)
            self.fig_tun.suptitle(f'Mean+/-STD ROI [{self.cur_roi_stat[0] + 1};{self.cur_roi_stat[1] + 1}]',
                                  fontsize=12)

            self.plot_grand_aver()

    def plot_grand_aver(self):  # Plotting average resp to pref ori over all days and trials
        print("Plot_grand_aver")
        self.ax_gr_aver.clear()
        proj = np.array(self.grand_aver)
        min_val = np.min(proj[~np.isnan(proj)])  # min value ignorring NaN-s
        max_val = np.max(proj[~np.isnan(proj)])  # max value ignorring NaN-s

        # Check if current thres exceeds current amplitude limits and correct if it does:
        if self.ampl_thrs < min_val:
            self.sl_gr_aver.set_val(min_val)
        elif self.ampl_thrs > max_val:
            self.sl_gr_aver.set_val(max_val)

        self.ax_gr_aver.imshow(proj,
                               interpolation="none",
                               cmap=self.cmap,
                               vmin=self.ampl_thrs,
                               vmax=max_val
                               )
        rect = patches.Rectangle((self.cur_roi_stat[1] - 0.5, self.cur_roi_stat[0] - 0.5), 1, 1, linewidth=2,
                                 edgecolor='r', facecolor="none")
        self.ax_gr_aver.add_patch(rect)

        self.ax_gr_aver.set_xticks(self.xtiks, labels=self.xlabs)  # Specifying ticks and labels
        self.ax_gr_aver.set_yticks(self.ytiks, labels=self.ylabs)
        over_thres = np.count_nonzero(proj > self.ampl_thrs)
        titl = f"Average response to pref. ori\n Thrsh={np.round(self.ampl_thrs, decimals=3)}; {over_thres} ROI respond"
        self.ax_gr_aver.set_title(label=titl, fontsize=13)

        # The last table [1,2] is for % of ROI responding over curr threshold, that reveal signif. change over days
        self.ax_pval[1, 2].clear()
        tab_count = np.zeros((1, 5))  # Output table
        for idx in range(5):  # self.p_values: [X, Y, IdxType, Day1, Day2]
            p_cur_idx = np.array(self.p_values[:, :, idx, :, :])  # p-values for all ROI, day comb. for curr. index
            any_chng = np.zeros((self.h, self.w))
            # Count ROIs which have at least 1 significant p-value in any days combination
            # p_cur_idx contains ones in lower triangle and diagonal cells - they don't influence counting val<0.05
            any_chng[np.where(np.count_nonzero(p_cur_idx < 0.05, axis=(2, 3)) > 0)] = 1
            any_chng[np.where(proj <= self.ampl_thrs)] = 0  # Do not include ROI with resp < threshold
            any_chng[np.where(np.isnan(proj))] = 0  # Not include ROI which aren't present in any day (filled with None)
            sign_chng = np.count_nonzero(any_chng)
            tab_count[0, idx] = round(100 * sign_chng / over_thres, 1)  # % of responding ROI that reveal signif. change
        idx_list = ["Pref. ori.", "Pref. dir.", "OI", "DI", "TI"]
        table = self.ax_pval[1, 2].table(cellText=tab_count,
                                         rowLabels='%',
                                         colWidths=[0.2] * 5,
                                         colLabels=idx_list,
                                         loc='center')
        table.auto_set_font_size(False)
        table.set_fontsize(8)
        table.auto_set_column_width([i for i in range(5)])
        self.ax_pval[1, 2].axis('off')
        self.ax_pval[1, 2].set_title(f'% of ROI-s with signif. change.\n'
                                     f'in all responding ROI, thres. = {np.round(self.ampl_thrs, decimals=3)}')

    # TREE FUNCTIONS ----------------------------------------------------------------------------------- TREE FUNCTIONS
    def rem_bottom(self, d, level):  # Remove last levels (specified by 2nd parameter) of the given nested dictionary
        if not isinstance(d, dict):
            return d
        elif level == 0:
            return list(d)
        else:
            return {k: self.rem_bottom(v, level - 1) for k, v in d.items()}

    def populateTree(self):  # Function to fill populate treeview with a dictionary
        print("PopulateTree")
        self.tree.clear()  # Otherwise all data would be duplicated
        for day_k, day_v in self.data.items():
            item = QTreeWidgetItem(self.tree)
            item.setText(0, f'{day_k}')
            item.setFlags(item.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
            for sess_k in day_v:
                subitem = QTreeWidgetItem(item)
                subitem.setText(0, f'{sess_k}')
                subitem.setFlags(subitem.flags() | Qt.ItemIsUserCheckable)
                subitem.setCheckState(0, Qt.Unchecked)
        # Set last session checked for displaying:
        root = self.tree.invisibleRootItem()
        day_count = root.childCount()
        last_day = root.child(day_count - 1)
        sess_count = last_day.childCount()
        last_sess = last_day.child(sess_count - 1)
        last_sess.setCheckState(0, Qt.Checked)
        self.tree.expandAll()

    def get_checked_sess(self):
        checked = []

        def recurse(parent):
            for i in range(parent.childCount()):
                child = parent.child(i)
                grand_hildren = child.childCount()
                if grand_hildren > 0:
                    recurse(child)
                else:
                    if child.checkState(0) == Qt.Checked:
                        checked.append((parent.text(0), child.text(0)))

        recurse(self.tree.invisibleRootItem())
        return checked

    # SAVE METHOD ---------------------------------------------------------------------------------------- SAVE METHOD
    def savef(self):  # Saving res variable to Excel
        if len(self.folder) > 0 and self.made:
            checked = self.get_checked_sess()
            print(f"Checked sessions: {checked}")
            res_list = []
            matrix_list = []
            matrix180_list = []
            matrix_corr_list = []
            pref_ori_list = []
            param_list = []
            columns = []
            for chs in checked:
                session = self.data[chs[0]][chs[1]]
                # Saving data from res variable (traces for all ori-s):
                if self.save_res.checkState() == 2:
                    if self.save_curroi.checkState() == 2:  # Saving data only for current ROI
                        for ori in range(len(self.ori_list)):
                            # res_sort = [X Y Ori Frame] -> [Ori Frame]
                            value = list(session['res'][self.curr_roi[0], self.curr_roi[1], ori, :])
                            res_list.append([chs[0] + '/' + chs[1],
                                             self.curr_roi[0] + 1,
                                             self.curr_roi[1] + 1,
                                             self.ori_list[ori],
                                             'val'] +
                                            value)
                            if session.get('res_std') is not None:  # i.e., if cur sess is averaged one:
                                value = list(session['res_std'][self.curr_roi[0], self.curr_roi[1], ori, :])
                                res_list.append([chs[0] + '/' + chs[1],
                                                 self.curr_roi[0] + 1,
                                                 self.curr_roi[1] + 1,
                                                 self.ori_list[ori],
                                                 'val'] +
                                                value)
                    else:  # Saving data for all ROIs
                        for ori in range(len(self.ori_list)):
                            for idx, val in np.ndenumerate(session['res'][:, :, ori, 0]):  # [X Y Ori Frame]
                                value = list(session['res'][idx[0], idx[1], ori, :])
                                res_list.append([chs[0] + '/' + chs[1],
                                                 idx[0] + 1,
                                                 idx[1] + 1,
                                                 self.ori_list[ori],
                                                 'val'] +
                                                value)
                                if session.get('res_std') is not None:  # i.e., if cur sess is averaged one:
                                    value = list(session['res_std'][idx[0], idx[1], ori, :])
                                    res_list.append([chs[0] + '/' + chs[1],
                                                     idx[0] + 1,
                                                     idx[1] + 1,
                                                     self.ori_list[ori],
                                                     'std'] +
                                                    value)
                # Saving data from matrix variable (response amplitude for all ori-s):
                if self.save_matrix.checkState() == 2:
                    if self.save_curroi.checkState() == 2:  # Saving data only for current ROI
                        value = list(session['matrix'][self.curr_roi[0], self.curr_roi[1], :-1])
                        matrix_list.append([chs[0] + '/' + chs[1],
                                            self.curr_roi[0] + 1,
                                            self.curr_roi[1] + 1,
                                            'val'] +
                                           value)
                        value180 = list(session['matrix180'][self.curr_roi[0], self.curr_roi[1], :6])
                        matrix180_list.append([chs[0] + '/' + chs[1],
                                               self.curr_roi[0] + 1,
                                               self.curr_roi[1] + 1] +
                                              value180)
                        value_corr = list(session['matrix_corr'][self.curr_roi[0], self.curr_roi[1], :])
                        matrix_corr_list.append([chs[0] + '/' + chs[1],
                                                 self.curr_roi[0] + 1,
                                                 self.curr_roi[1] + 1,
                                                 'val'] +
                                                value_corr)
                        if session.get('matrix_std') is not None:  # i.e., if cur sess is averaged one:
                            value = list(session['matrix_std'][self.curr_roi[0], self.curr_roi[1], :-1])
                            matrix_list.append([chs[0] + '/' + chs[1],
                                                self.curr_roi[0] + 1,
                                                self.curr_roi[1] + 1,
                                                'std'] +
                                               value)
                    else:  # Saving data for all ROIs
                        for idx2, val in np.ndenumerate(session['matrix'][:, :, 0]):
                            value = list(session['matrix'][idx2[0], idx2[1], :-1])
                            matrix_list.append([chs[0] + '/' + chs[1],
                                                idx2[0] + 1,
                                                idx2[1] + 1,
                                                'val'] +
                                               value)
                            value180 = list(session['matrix180'][idx2[0], idx2[1], :6])
                            matrix180_list.append([chs[0] + '/' + chs[1],
                                                   idx2[0] + 1,
                                                   idx2[1] + 1] +
                                                  value180)
                            value_corr = list(session['matrix_corr'][idx2[0], idx2[1], :])
                            matrix_corr_list.append([chs[0] + '/' + chs[1],
                                                     idx2[0] + 1,
                                                     idx2[1] + 1,
                                                     'val'] +
                                                    value_corr)
                            if session.get('matrix_std') is not None:  # i.e., if cur sess is averaged one:
                                value = list(session['matrix_std'][idx2[0], idx2[1], :-1])
                                matrix_list.append([chs[0] + '/' + chs[1],
                                                    idx2[0] + 1,
                                                    idx2[1] + 1,
                                                    'val'] +
                                                   value)
                # Saving inf values of preferred ori-s & directions, OI, DI & TI
                # 'maps': [X, Y, [pref_ori, pref_dir, OI, DI, TI]]
                if self.save_map.checkState() == 2:
                    if self.save_curroi.checkState() == 2:  # Saving data only for current ROI
                        pref_ori_list.append([chs[0] + '/' + chs[1],
                                              self.curr_roi[0] + 1,
                                              self.curr_roi[1] + 1,
                                              np.rad2deg(session['maps'][self.curr_roi[0], self.curr_roi[1], 0]),
                                              np.rad2deg(session['maps'][self.curr_roi[0], self.curr_roi[1], 1]),
                                              session['maps'][self.curr_roi[0], self.curr_roi[1], 2],
                                              session['maps'][self.curr_roi[0], self.curr_roi[1], 3],
                                              session['maps'][self.curr_roi[0], self.curr_roi[1], 4]
                                              ])
                    else:
                        for idx2, val in np.ndenumerate(session['maps'][:, :, 0]):
                            pref_ori_list.append([chs[0] + '/' + chs[1],
                                                  idx2[0] + 1,
                                                  idx2[1] + 1,
                                                  np.rad2deg(val),
                                                  np.rad2deg(session['maps'][idx2[0], idx2[1], 1]),
                                                  session['maps'][idx2[0], idx2[1], 2],
                                                  session['maps'][idx2[0], idx2[1], 3],
                                                  session['maps'][idx2[0], idx2[1], 4]
                                                  ])
                # Saving parameters
                if self.save_param.checkState() == 2:
                    columns = list(self.par.keys())  # Parameters' names are identical for all sessions
                    if session.get('matrix_std') is not None:  # i.e., if cur sess is averaged one:
                        param_list.append([chs[0], chs[1]] + [session['aver_over']] + ['' for i in range(8)])
                    else:
                        param_list.append(session['param'].values())

            if self.save_curroi.checkState() == 2:  # Saving data only for current ROI
                add_name = f"_ROI[{self.curr_roi[0] + 1}-{self.curr_roi[1] + 1}]"
            else:
                add_name = ''
            with pd.ExcelWriter(f'{self.folder}\\Results{add_name}.xlsx') as writer:
                if len(res_list) > 0:
                    df_res = pd.DataFrame(res_list, columns=['Probe', 'X', 'Y', 'Ori',
                                                             'DType'] + [str(i) for i in range(1, self.trial_len + 1)])
                    df_res.to_excel(writer, sheet_name="Traces", index=False)
                if len(matrix_list) > 0:
                    df_matrix = pd.DataFrame(matrix_list,
                                             columns=['Probe', 'X', 'Y', 'DType'] + [str(i) for i in self.ori_list])
                    df_matrix.to_excel(writer, sheet_name="Tuning", index=False)
                    df_matrix180 = pd.DataFrame(matrix180_list,
                                                columns=['Probe', 'X', 'Y'] + [str(i) for i in self.ori_list[:6]])
                    df_matrix180.to_excel(writer, sheet_name="Tuning 180", index=False)
                    df_matrix_corr = pd.DataFrame(matrix_corr_list, columns=['Probe', 'X', 'Y', 'DType']
                                                                            + [str(i) for i in self.ori_list])
                    df_matrix_corr.to_excel(writer, sheet_name="Tuning corr", index=False)
                if len(pref_ori_list) > 0:
                    df_pref_ori = pd.DataFrame(pref_ori_list, columns=['Probe', 'X', 'Y', 'Pref. ori, deg',
                                                                       'Pref. dir, deg', 'OI, %', 'DI, %', 'TI, %'])
                    df_pref_ori.to_excel(writer, sheet_name="Preferred ori", index=False)
                if len(param_list) > 0:
                    df_param = pd.DataFrame(param_list, columns=columns)
                    df_param.to_excel(writer, sheet_name="Parameters", index=False)

            save_list = [self.save_res, self.save_matrix, self.save_map, self.save_param]
            if len([i for i in range(4) if save_list[i].checkState() == 2]) == 0:
                self.message.setText(f"Please, select items to save at Parameters tab")
            else:
                self.message.setText(f"Data were saved to\n{self.folder}")
        else:
            self.message.setText("Please, select folder with files")

    def save_stacks(self):
        from tifffile import imwrite
        print("Save stacks")
        # Saving tiff hyperstacks for responses to all ori-s for each of selected sessions
        if len(self.folder) > 0 and self.made:
            checked = self.get_checked_sess()
            pre_day = ''
            stack = []
            for chs in checked:
                if pre_day == chs[0]:  # if previous day is same to the current one, append matrix to this day
                    print(f"  Adding trial {chs[1]} to day {chs[0]}")
                    stack.append(self.data[chs[0]][chs[1]]['matrix'][:, :, :-1])
                    pre_day = chs[0]
                else:  # if previous day is different, save current stack and create new one
                    print(f"  New day {chs[0]}")
                    if len(stack) > 0:
                        stack = np.array(stack, dtype='float32')  # [X, Y, 12ori, Trials]
                        stack = np.swapaxes(stack, 2, 3)  # [trial, Y, ori, X]
                        stack = np.swapaxes(stack, 1, 2)  # [trial, ori, Y, X]
                        imwrite(f'{self.folder}\\Day_{pre_day}.tif', stack, imagej=True)
                        print(f"  Day {pre_day} was saved to {self.folder}")
                    stack = [self.data[chs[0]][chs[1]]['matrix'][:, :, :-1]]
                    print(f"  Adding trial {chs[1]} to day {chs[0]}")
                    pre_day = chs[0]
            stack = np.array(stack, dtype='float32')  # [trial, Y, X, ori]
            stack = np.swapaxes(stack, 2, 3)  # [trial, Y, ori, X]
            stack = np.swapaxes(stack, 1, 2)  # [trial, ori, Y, X]
            imwrite(f'{self.folder}\\Day_{pre_day}.tif', stack, imagej=True)
            print(f"  Day {pre_day} was saved to {self.folder}")
            self.message.setText(f"TIFFs were saved to\n{self.folder}")

    def save_align(self):
        folder_name = self.folder.split('\\')[-1]
        with open(self.folder + f'\\Align_{folder_name}.txt', 'w') as file:
            if len(self.folder) > 0 and self.made:
                file.write("Day\tName\tShift_X\tShift_Y\tShift_Rot\n")
                for day_key, day_val in self.data.items():
                    for trial_key in day_val.keys():
                        param = self.data[day_key][trial_key]['param']
                        x = param['Shift_X']
                        y = param['Shift_Y']
                        deg = param['Shift_Rot']
                        file.write(f"{day_key}\t{trial_key}\t{x}\t{y}\t{deg}\n")
        print(f"File with align values was saved to: {self.folder}")

    # AUXILIARY METHODS ----------------------------------------------------------------------------- AUXILIARY METHODS

    def location_on_the_screen(self):
        ag = QDesktopWidget().availableGeometry()
        widget = MainWindow.geometry()
        self.screen_width = ag.width()
        self.screen_height = ag.height()
        x = self.screen_width - widget.width() - 20
        y = 10  # 2 * ag.height() - sg.height() - widget.height()
        MainWindow.move(x, y)

    def get_parameters(self):
        frrate = self.fps.toPlainText()
        stim_lat = self.st_lat.toPlainText()
        duration = self.dur.toPlainText()
        sq = self.sq.toPlainText()
        fXY = self.filtXY.toPlainText()
        fT = self.filtT.toPlainText()
        if frrate.isdigit():
            self.frrate = float(frrate)
        if stim_lat.isdigit():
            self.stim_lat = float(stim_lat)
        if duration.isdigit():
            self.duration = float(duration)
        if sq.isdigit():
            self.wind_size = int(sq)
        if fXY.isdigit():
            self.filter = int(fXY)
        if fT.isdigit():
            self.filt_time = int(fT)
        self.st_lat_fr = round(self.frrate * self.stim_lat)  # frame of stimulation onset
        self.windbrdr = round(self.frrate * (
                self.stim_lat + self.duration))  # End of stimulation
        self.par = {'day': '',
                    'name': '',
                    'averaged': False,
                    'frrate': self.frrate,
                    'Latency_of_stimulation_onset(s)': self.stim_lat,
                    'Latency_of_stimulation_onset(frames)': self.st_lat_fr,
                    'Stimulation_duration(s)': self.duration,
                    'Stimulation_end(frames)': self.windbrdr,
                    'Size_of_square_ROI(px)': self.wind_size,
                    'filterXY(px)': self.filter,
                    'filterT(frames)': self.filt_time,
                    'Shift_X': 0,
                    'Shift_Y': 0,
                    'Shift_Rot': 0
                    }

    def change_ori(self, arg):
        print(f"Change ori: {int(self.sl_ori.val)}")
        self.curr_ori = int(self.sl_ori.val)
        self.curr_ori = self.ori_list.index(self.curr_ori) + 1
        self.display()

    def change_sess(self, val):
        self.cursess = int(self.sl_sess.val) - 1

        # Updating slider for threshold (color bar will be updated in display() method)
        checked = self.get_checked_sess()
        if len(checked) < self.cursess + 1:
            self.cursess = len(checked) - 1
        cursess = checked[self.cursess]
        whole_matrix = np.array(self.data[cursess[0]][cursess[1]]['matrix'])
        min_val = np.min(whole_matrix[~np.isnan(whole_matrix)])  # min value ignoring NaN-s
        max_val = np.max(whole_matrix[~np.isnan(whole_matrix)])  # max value ignoring NaN-s
        self.sl_thr.valmin = min_val
        self.sl_thr.valmax = max_val
        print(f"sl_thr  min: {self.sl_thr.valmin}  max: {self.sl_thr.valmax}")

        self.display()

    def change_thresh(self, val):  # Change threshold of response intensity to mask lower values
        self.mask_resp = float(self.sl_thr.val)
        self.display()

    def change_thrs_gr_aver(self, val):  # Change threshold of response intensity to mask lower values
        self.ampl_thrs = float(self.sl_gr_aver.val)
        if self.p_values is not None:
            self.plot_grand_aver()

    def onclick(self, event):  # Changing current ROI with clicking on projection (response map or general image)
        ix, iy = event.xdata, event.ydata
        x, y = event.x, event.y
        if ix is not None and iy is not None:
            if (63 < x < 386) and (129 < y < 447):  # This prevents considering slider and colorbar as area of an image
                self.curr_roi = (round(iy), round(ix))
                if self.activity:
                    selected = [ind for ind, box in enumerate(self.check.get_status()) if box is True]
                    [self.check.set_active(j) for j in selected]  # Clearing the set of ori-s to display on trace plot
                self.display()
                if self.p_values is not None:
                    self.plot_statistics()  # If statistics was calculated, update plots for new ROI

    def onclick_stat(self, event):  # Changing current ROI with clicking on grand aver fig. (statistics figures)
        # This separate function prevents plotting tuning curves for numerous selected trials on fig 1
        # while running statistic analysis. i.e. enables independent ROI selection for statistics fig-s & other fig-s.
        ix, iy = event.xdata, event.ydata
        x, y = event.x, event.y
        if ix is not None and iy is not None:
            if 63 < x < 386:  # This prevents considering slider and colorbar as area of an image
                self.cur_roi_stat = (round(iy), round(ix))
                if self.p_values is not None:
                    self.plot_statistics()  # If statistics was calculated, update plots for new ROI

    def onclick_key(self, event):  # Shifting alignment image with arrow buttons
        key = event.key
        if key == 'down':
            if self.sl_y.val + 1 <= self.sl_y.valmax:
                self.sl_y.set_val(self.sl_y.val + 1)
        elif key == 'up':
            if self.sl_y.val - 1 >= self.sl_y.valmin:
                self.sl_y.set_val(self.sl_y.val - 1)
        elif key == 'right':
            if self.sl_x.val + 1 <= self.sl_x.valmax:
                self.sl_x.set_val(self.sl_x.val + 1)
        elif key == 'left':
            if self.sl_x.val - 1 >= self.sl_x.valmin:
                self.sl_x.set_val(self.sl_x.val - 1)
        elif event.key == 'enter':
            self.get_shift(0)  # Apply current values pressing 'Enter'

    def onclick_scroll(self, event):  # Rotating alignment image with mouse wheel scroll
        increment = 1 if event.button == 'up' else -1
        if self.sl_rot.valmin < self.sl_rot.val + increment < self.sl_rot.valmax:
            self.sl_rot.set_val(self.sl_rot.val + increment)

    def switch180(self, val):  # Switch between 360- & 180-degree representation of tuning curve
        print("Switch180")
        if self.disp180 == 'matrix':
            self.disp180 = 'matrix180'
        else:
            self.disp180 = 'matrix'
        self.display()

    def changeOI(self, val):
        print("Switch180")
        checked = self.get_checked_sess()
        cursess = checked[self.cursess]
        max_proj = np.max(self.data[cursess[0]][cursess[1]]['matrix'], axis=2)  # Proj of max response among ori-s
        if self.dispOI == 'OI':
            proj = self.data[cursess[0]][cursess[1]]['maps'][:, :, 2]  # [X, Y, [pref_ori, pref_dir, OI, DI, TI]]
            proj[np.where(max_proj < self.mask_resp)] = np.min(proj[~np.isnan(proj)])  # Filter ORIs with too low resp
            self.axOI.set_title(label=f"Orientation selectivity\n{cursess[0]}/{cursess[1]}",
                                fontsize=13)
            self.dispOI = 'DI'
        else:
            proj = self.data[cursess[0]][cursess[1]]['maps'][:, :, 3]  # [X, Y, [pref_ori, pref_dir, OI, DI, TI]]
            proj[np.where(max_proj < self.mask_resp)] = -30  # Filtering the ORIs with too low response
            self.axOI.set_title(label=f"Direction selectivity\n{cursess[0]}/{cursess[1]}",
                                fontsize=13)
            self.dispOI = 'OI'
        self.imOI.set_data(proj)

    def add_ori(self, label):  # Updating plots with new set of ori-s to display at activity traces plot
        self.display()  # This func is necessary cause on_click gives 2 arg-s while display() receives only 1

    def mult_along_axis(self, a, b, axis):
        # ensure we're working with Numpy arrays
        a = np.array(a)
        b = np.array(b)
        # shape check
        if axis >= a.ndim:
            raise AxisError(axis, a.ndim)
        if a.shape[axis] != b.size:
            raise ValueError(
                f"Length of 'A' along the axis {axis} is {a.shape[axis]} while B.size is {b.size}"
            )
        # np.broadcast_to puts the new axis as the last axis, so
        # we swap the given axis with the last one, to determine the
        # corresponding array shape. np.swapaxes only returns a view
        # of the supplied array, so no data is copied unnecessarily.
        shape = np.swapaxes(a, a.ndim - 1, axis).shape
        # Broadcast to an array with the shape as above. Again,
        # no data is copied, we only get a new look at the existing data.
        b_brc = np.broadcast_to(b, shape)
        # Swap back the axes. As before, this only changes our "point of view".
        b_brc = np.swapaxes(b_brc, a.ndim - 1, axis)
        return a * b_brc

    def lightness(self, color, amount=0.5):
        c = colorsys.rgb_to_hls(*colors.to_rgb(color))
        return colorsys.hls_to_rgb(c[0], max(0, min(1, amount * c[1])), c[2])

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Visualization of calcium activity"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("MainWindow", "Main"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), _translate("MainWindow", "Parameters"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_3), _translate("MainWindow", "Multi Sessions"))
        self.select_btn.setText(_translate("MainWindow", "Select folder"))
        self.aver_ss.setText(_translate("MainWindow", "Plot average"))
        self.split_ss.setText(_translate("MainWindow", "Plot separately"))
        self.go.setText(_translate("MainWindow", "Compute"))
        self.save.setText(_translate("MainWindow", "Excel"))
        self.save_tif.setText(_translate("MainWindow", "TIFF"))
        self.show_btn.setText(_translate("MainWindow", "Preferred ori"))
        self.MU_btn.setText(_translate("MainWindow", "MU"))
        self.OI_btn.setText(_translate("MainWindow", "OI"))
        self.disp_act.setText(_translate("MainWindow", "Activity trace"))


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    ui.location_on_the_screen()
    MainWindow.show()
    sys.exit(app.exec_())
