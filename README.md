# Task Monitor

Ведет протокол:

* имени процесса
* Process ID
* имени родительского процесса
* Process ID родительского процесса
* имя пользователя / домен, от чьего имени запущен процесс
* Путь и аргументы командной строки запущенного/завершенного процесса.

Протокол формируется в файлы форматов Plain Text и CSV.

Ограничение: не засекает процессы, время жизни которых менее 1 сек.

## Использование
1. Скачать архив и распаковать его на рабочий стол.
2. Запустить файл **Process_Monitor2.vbs**.
3. Произвести нужные операции (запуск/завершение отслеживаемых процессов).
4. Запустить еще раз **Process_Monitor2.vbs **- это остановит процесс мониторинга.
Рядом с этим скриптом будут созданы 2 файла:

* Processes_дата_время.log
* Processes_дата_время.csv

5. Запакуйте их в архив и прикрепите к своему сообщению по запросу Консультанта.
