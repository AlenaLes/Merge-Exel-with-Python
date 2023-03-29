# Объединение Exel таблиц с помощью Python

## Краткое описание
Данная задача была выполнена в рамках реальной оптимизации процесса для упрощения ежедневной работы. Данные, представленные в Exel-файлах, обезличены и созданы вручную в связи с политикой конфеденциальности компании.

Необходимо объединять два файла Exel и осуществлять небольшое форматирование таблиц. Данная работа ранее осуществлялась ежедневно вручную с помощью Exel, однако, в связи с большим объемом данных, данная задача занимала слишком много времени (в частоности обработка данных самим Exel). 
Для экономии времени и трудовых затрат написан код для автоматического решения данной задачи.

Исходные данные можно посмотреть в файле "статистика-first", итоговые данные, которые получили на выходе - "статистика".

Стэк:

- Python
- JupiterHub
- Exel

## Библиотеки для работы с Exel

```
#pip install pandas
#pip install openpyxl

import pandas as pd
import numpy as np
import os
import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook
```

Загружаем Exel-файлы и проверяем, что нужные нам листы открываются верно. Теперь можно делать объединение с помощью merge и Сразу же уберем ненужные столбцы.

```
Alpha_join = file_domena.merge(file_fio, how = 'inner', left_on = 'USERS', right_on = 'domena_login')
Alpha_join = Alpha_join.drop(columns=['Подразделение 2', 'Подразделение 3', 'Должность', 'domena_login', 'domenb_login'])
Alpha_join.head()
```
![image](https://user-images.githubusercontent.com/100629361/228647182-e0218914-1bc5-4bb0-b285-9823f295327f.png)

Завершив объединение таблиц и создав новые датафрэймы с обновленными данными можем приступить к обновлению первоначальной таблицы.
```
# Беоем исходный файл для добавления обновленных листов
wb = openpyxl.load_workbook('статистика.xlsx') 
```

Заменяем старые листы в Exel-файле новыми датафрэймами.
```
with pd.ExcelWriter('статистика.xlsx', mode='a', if_sheet_exists= 'replace') as writer:  
    Alpha_join.to_excel(writer, sheet_name='domena',index = False)
    Sigma_join.to_excel(writer, sheet_name='domenb',index = False)
```

На выходе получаем Exel-файл с тем же наименованием, но с обновленными данными.
