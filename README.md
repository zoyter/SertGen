# Генератор сертификатов SertGen

## Назначение

Приложение **SertGen** предназначено для массового создания сертификатов участников конференций, семинаров и т.п.
Для этого используется следующая схема работы:

1. Подготваливается шаблон сертификата и помещается в папку **03.templates**.
   
2. Подготоваливается/экспортируется список участников мероприятия в формате XLS и помещается в папку **01.input**. При этом обязательно следует соблюдать следующие условия:

    - первый столбец - это фамилия/имя участника мероприятия;

    - второй столбец - это название мероприятия;
  
    - третий столбце - это дата проведения мероприятия в текстовом формате;

    - четвертый столбце - это уникальный номер сертификата.

3. Запускается программа (файл main.py). Если у вас установлена операционная система Windows, то дополнительно надо указать место размещения программы Inkscape.

4. Выбираете шаблон сертифката из списка.

5. Выбираете табличный файл. При необходимости можно внести изменения.

6. Запускаете процесс генерации сертификатов. При этом параллельно информция о созданных сертификатах помещается в базу данных.

7. Просматриваете результат работы в папке **02.output**

## Системные требования

Работа приложения проверялась в **Python 3.9.1** под управлением операционной системы Windows 10 и Xubuntu 20.04.

Для корректной работы приложения необходимо установить следующие библиотеки:

- PyQt5 (интерфейс приложения)
  ```cmd
  pip install PyQt5
  ```
- xlrd (работа с табличными файлами XLS)
  ```cmd
  pip install xlrd
  ```
- для **Linux**
  
  - CairoSVG (создание PDF из SVG)
    ```cmd
    pip install CairoSVG
    ```
- для **Windows**
  - необходимо установить векторный редактор [Inkscape](https://inkscape.org/release/1.0.2/windows/32-bit/)
  рекомендуем устанавливать портативную версию


Смотри файл [requirements.txt](requirements.txt).

## Порядок сборки проекта в автономное приложение

**Запуск приложения:**

Для запуска приложения неоходимо установить все недостающие компоненты. После этого запустить файл **main.py**

**Известные проблемы:**

- Убедитесь, что в установлен python последней версии
- При необходимости установите недостающие модули

**Сборка приложения**

Для сборки автономного приложения требуется установить модуль **pyinstaller** и выполнить команду:  
```cmd
pyinstaller --onefile --noconsole main.py
```
 
После этого необходимо получившийся файл **main.exe** из каталога **dist** поместить в каталог проекта (в ту  папку, где находятся: 01.input, 02.output, 03.templates, 04.help, img, tmp, ui, db.py, db.sqlite)


## Что еще можно доделать

Много чего хотелось бы добавить, например:

- возможность сопоставления полей табличного файла с полями базы данных
- примитивный редактор SVG
- интерактивное размещение полей на сертификате


## P.S.

Персональный проект в рамках подготовки преподавателей второго года обучения ЯЛ.
