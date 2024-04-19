Для создания исполняемого файла (exe) из этого скрипта мы можем воспользоваться библиотекой PyInstaller. Вот как это сделать:

1. Установите PyInstaller, если у вас его еще нет. Вы можете установить его с помощью pip:
pip install pyinstaller

2. Сохраните ваш скрипт в файле с расширением .py, например, equipment_specification_generator.py.

3. Откройте командную строку (на Windows) или терминал (на macOS и Linux).

4. Перейдите в каталог, содержащий ваш скрипт, используя команду cd, например:
cd путь_к_каталогу_с_вашим_скриптом

5. Затем запустите следующую команду в командной строке для создания исполняемого файла:
pyinstaller --onefile equipment_specification_generator.py
