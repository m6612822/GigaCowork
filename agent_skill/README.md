# Установка и запуск навыка

1. Распакуйте `ful-logistics-distributor.zip` в каталог навыков GigaCowork.
2. Убедитесь, что доступен Python 3 и библиотека `openpyxl` (также допускается `xlrd`, но скрипту она не требуется).
3. Приложите три Excel-файла: выгрузку ФУЛ, реестр Cloud и список сотрудников.
4. Запустите навык. Для прямого запуска скрипта используйте:

```bash
python scripts/process_ful.py --ful ful.xlsx --cloud cloud.xlsx --employees employees.xlsx --output-dir output
```

Результат: `output/result.xlsx`. Скрипт ничего не отправляет во внешние системы.
