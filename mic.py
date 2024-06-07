import os
import threading
import keyboard
import winsound
import sys
import json
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from PyQt5.QtWidgets import QApplication, QSystemTrayIcon, QMenu, QAction, QMessageBox, QDialog, QVBoxLayout, QLabel, QKeySequenceEdit
from PyQt5.QtGui import QIcon, QCursor
from PyQt5.QtCore import Qt
import winshell

# Проверка на наличие sys._MEIPASS для определения пути к иконкам
if hasattr(sys, '_MEIPASS'):
    base_path = sys._MEIPASS
else:
    base_path = os.path.abspath(".")

# Определение путей к иконкам
mic_off_icon_path = os.path.join(base_path, "mic_off_icon.png")
mic_on_icon_path = os.path.join(base_path, "mic_on_icon.png")

# Файл для хранения настроек
SETTINGS_FILE = "settings.json"

# Функция для загрузки настроек из файла
def load_settings():
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, "r") as file:
            return json.load(file)
    return {"hotkey": "F14", "autostart": False}

# Функция для сохранения настроек в файл
def save_settings(settings):
    with open(SETTINGS_FILE, "w") as file:
        json.dump(settings, file)

# Загружаем настройки
settings = load_settings()
current_hotkey = settings.get("hotkey", "F14")
autostart_enabled = settings.get("autostart", False)

# Получаем доступ к микрофону
devices = AudioUtilities.GetMicrophone()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

# Функция для воспроизведения звука из стандартной библиотеки Windows 10 в отдельном потоке
def play_sound(sound):
    sound_path = os.path.join("C:\\Windows\\Media", sound)
    winsound.PlaySound(sound_path, winsound.SND_FILENAME | winsound.SND_ASYNC)

# Функция для включения/выключения микрофона
def toggle_microphone():
    if volume.GetMute() == 1:
        volume.SetMute(0, None)
        tray_icon.setIcon(QIcon(mic_on_icon_path))  # Путь к иконке микрофона включенного
        tray_icon.showMessage("Микрофон ВКЛЮЧЕН", "Микрофон был включен")  # Сообщение в системном трее
        threading.Thread(target=play_sound, args=("Windows Hardware Insert.wav",)).start()
    else:
        volume.SetMute(1, None)
        tray_icon.setIcon(QIcon(mic_off_icon_path))  # Путь к иконке микрофона выключенного
        tray_icon.showMessage("Микрофон ВЫКЛЮЧЕН", "Микрофон был выключен")  # Сообщение в системном трее
        threading.Thread(target=play_sound, args=("Windows Hardware Fail.wav",)).start()

# Функция для изменения клавиши горячей клавиши
def change_hotkey():
    global current_hotkey
    dialog = QDialog()
    dialog.setWindowTitle("Изменить горячую клавишу")
    layout = QVBoxLayout()
    label = QLabel("Нажмите клавишу или сочетание клавиш, которые хотите использовать для управления микрофоном")
    layout.addWidget(label)

    key_sequence_edit = QKeySequenceEdit()
    layout.addWidget(key_sequence_edit)

    def on_key_sequence_changed(key_sequence):
        pressed_key = key_sequence.toString()
        if pressed_key:
            global current_hotkey
            keyboard.remove_hotkey(current_hotkey)
            keyboard.add_hotkey(pressed_key, toggle_microphone)
            current_hotkey = pressed_key
            settings['hotkey'] = current_hotkey
            save_settings(settings)
            dialog.hide()  # Скрываем диалоговое окно вместо закрытия

    key_sequence_edit.keySequenceChanged.connect(on_key_sequence_changed)
    dialog.setLayout(layout)
    dialog.exec_()

# Обновленная функция для изменения автозапуска приложения вместе с ОС Windows
def toggle_autostart():
    global autostart_enabled
    app_name = "MicrophoneToggleApp"
    app_path = os.path.abspath(sys.argv[0])
    startup_path = winshell.startup()

    shortcut_path = os.path.join(startup_path, f"{app_name}.lnk")
    
    if autostart_enabled:
        # Отключаем автозапуск
        autostart_enabled = False
        settings['autostart'] = autostart_enabled
        save_settings(settings)
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)
    else:
        # Включаем автозапуск
        autostart_enabled = True
        settings['autostart'] = autostart_enabled
        save_settings(settings)
        with winshell.shortcut(shortcut_path) as shortcut:
            shortcut.path = app_path
            shortcut.description = "Microphone Toggle Application"
            shortcut.working_directory = os.path.dirname(app_path)
            shortcut.icon_location = (app_path, 0)

    # Обновляем сообщение в трее
    status = "Вкл." if autostart_enabled else "Выкл."
    tray_icon.showMessage("Настройка автозапуска", f"Автозапуск {status}")
    toggle_autostart_action.setText(f"Включить/отключить автозапуск (Текущий: {status})")

# Обработчик события нажатия кнопки мыши на значке системного трея
def tray_left_clicked(reason):
    if reason == QSystemTrayIcon.Trigger:
        toggle_microphone()
    elif reason == QSystemTrayIcon.Context:
        tray_menu.exec_(QCursor.pos())  # Показываем контекстное меню в позиции курсора

# Функция для отображения информации о программе с HTML-разметкой
def show_about_dialog():
    dialog = QDialog()
    dialog.setWindowTitle("О программе")
    dialog.setFixedSize(350, 150)  # Устанавливаем фиксированный размер окна

    layout = QVBoxLayout()
    
    about_message = (
        "<html>"
        "<p>Это приложение включает и отключает микрофон в системе</p>"
        "<p>по нажатию заданной горячей клавиши или сочетания клавиш.</p>"
        "<hr> </hr>" 
        "<p>Программу создал FarEvil.</p>"
        "<p>Ссылка для связи с разработчиком: <a href='https://linktr.ee/FarEvil'>https://linktr.ee/FarEvil</a></p>"
        "</html>"
    )
    
    label = QLabel(about_message)
    label.setOpenExternalLinks(True)
    layout.addWidget(label)
    dialog.setLayout(layout)

    # Скрываем окно при закрытии вместо завершения приложения
    def close_event(event):
        event.ignore()
        dialog.hide()

    dialog.closeEvent = close_event
    dialog.setWindowFlags(dialog.windowFlags() & ~Qt.WindowContextHelpButtonHint)  # Убираем кнопку "Справка"
    dialog.exec_()

# Создаем объект приложения Qt
app = QApplication(sys.argv)

# Создаем системный трей
tray_icon = QSystemTrayIcon(QIcon(mic_off_icon_path), app)
tray_icon.setToolTip("Microphone Toggle App")

# Создаем контекстное меню для системного трея
tray_menu = QMenu()
toggle_action = QAction("Включить/выключить микрофон", tray_icon)
settings_menu = tray_menu.addMenu("Настройки")  # Пункт "Настройки"
change_hotkey_action = QAction("Изменить горячую клавишу", tray_icon)
toggle_autostart_action = QAction(f"Включить/отключить автозапуск (Текущий: {'Вкл.' if autostart_enabled else 'Выкл.'})", tray_icon)
about_action = QAction("О программе", tray_icon)
quit_action = QAction("Выход", tray_icon)
tray_menu.addAction(toggle_action)
tray_menu.addMenu(settings_menu)  # Добавляем подменю "Настройки"
settings_menu.addAction(change_hotkey_action)
settings_menu.addAction(toggle_autostart_action)
tray_menu.addAction(about_action)
tray_menu.addAction(quit_action)

# Обработка действий контекстного меню
toggle_action.triggered.connect(toggle_microphone)
change_hotkey_action.triggered.connect(change_hotkey)
toggle_autostart_action.triggered.connect(toggle_autostart)
about_action.triggered.connect(show_about_dialog)
quit_action.triggered.connect(app.quit)

# Привязываем обработчик события к значку системного трея
tray_icon.activated.connect(tray_left_clicked)

# Устанавливаем значок в зависимости от состояния микрофона при запуске
if volume.GetMute() == 1:
    tray_icon.setIcon(QIcon(mic_off_icon_path))
else:
    tray_icon.setIcon(QIcon(mic_on_icon_path))

# Показываем системный трей
tray_icon.show()

# Назначаем горячую клавишу для переключения микрофона
keyboard.add_hotkey(current_hotkey, toggle_microphone)

# Запускаем приложение Qt
sys.exit(app.exec_())