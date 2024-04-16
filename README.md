## YellowQuery: надстройка Excel для запросов к 1С

YellowQuery позволяет писать формулы в Excel для получения данных из 1С. Поддерживается встроенный язык запросов 1С:Предприятие 8.

### Идея

Большой помехой для эффективной работы с данными 1С в Excel является невозможность хранения в ячейках листа данных ссылочного типа. Эта надстройка сделана, чтобы проверить концепцию хранения таких данных в виде [навигационных ссылок](https://tinyurl.com/ytps9xyt).

### Функции

Для использования в формулах доступны следующие функции.

#### Функция YQ

Функция возвращает результат выполнения запроса на языке 1С:Предприятие 8. Если результат запроса не содержит строк или содержит одну колонку и строку, результат выводится в виде обычного значения, в остальных случаях - в виде массива. Начиная с Excel версии  2021, массив автоматически [переносится](https://support.microsoft.com/en-us/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531) в смежные ячейки. В ранних версиях Excel область ячеек для вывода массива задается [вручную](https://support.microsoft.com/en-us/office/guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7).

##### Синтаксис

**YQ(Base; QueryText[; Params])**

Например:

    =YQ("e1c://filev/o/Tests/ib";"выбрать первые 1 спр.наименование из справочник.контрагенты как спр где спр.ИНН = &ИНН";"ИНН";"1234567890")
    =YQ("e1c://server/MyServer/MyBase";"выбрать спр.ссылка, спр.наименование из справочник.пользователи как спр")

| Имя аргумента | Описание |
|-----|-----|
| Base (обязательный) | Навигационная ссылка информационной базы 1С, к которой выполняется запрос. Например, для файловой базы - e1c://filev/с/Path/To/My/Base, для клиент-серверного варианта - e1c://server/MyServer/MyBase |
| QueryText (обязательный) | Текст запроса на языке 1С. Например, выбрать первые 10 спр.наименование из справочник.пользователи как спр |
| Params | Произвольное число параметров запроса. Параметры задаются парами имя;значение.|

#### Функция REFP

Функция получает представление ссылки.

##### Синтаксис

**REFP(Reference)**

Например:

    =REFP("e1c://filev/o/Tests/ib#e1cib/data/Справочник.Контрагенты?ref=979a50ebf628462e11eedfb52fb803ac")

| Имя аргумента | Описание |
|-----|-----|
| Reference (обязательный) | Внешняя навигационная ссылка объекта, для которого требуется получить представление |

#### Функция REFA

Функция получает значение реквизита ссылки.

##### Синтаксис

**REFA(Reference; AttributeName)**

Например:

    =REFA("e1c://filev/o/Tests/ib#e1cib/data/Справочник.Контрагенты?ref=979a50ebf628462e11eedfb52fb803ac";"Код")

| Имя аргумента | Описание |
|-----|-----|
| Reference (обязательный) | Внешняя навигационная ссылка объекта, значение реквизита которого требуется получить |
| AttributeName (обязательный) | Имя реквизита объекта, значение которого требуется получить |

### Ограничения

- Для удобной работы функции **YQ** рекомендуется Excel 2021 из-за поддержки [формул динамического массива](https://support.microsoft.com/en-us/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531). В более ранних версиях Excel потребуется [вручную](https://support.microsoft.com/en-us/office/guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7) задавать область ячеек для вывода результата формулы.
- Требуется установка расширения конфигурации.
- Вызовы 1С из Excel выполняются через [Automation Server](https://ru.wikipedia.org/wiki/Microsoft_OLE_Automation) 1С:Предприятия 8.3 в режиме тонкого клиента (V83C.Application).
- У пользователя 1С должна быть включена аутентификация операционной системы, должны быть права на запуск тонкого клиента и Automation.

### Минимальные системные требования

- Excel 2007, рекомендуется Excel 2021.
- Технологическая платформа 1С:Предприятие 8.3.12

### Установка

- [Установить](https://tinyurl.com/5f9pt6ez) расширение конфигурации 1С YellowQuery.cfe
- [Установить](https://support.microsoft.com/en-us/office/add-or-remove-add-ins-in-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460) надстройку Excel YellowQuery.xlam
- Предоставить права доступа пользователю 1С на запуск тонкого клиента и Automation.
- Настроить пользователю аутентификацию операционной системы.

### Сборка

- [Установить](https://github.com/vanessa-opensource/vanessa-runner) vanessa-runner
- Выполнить Build.cmd

### Тестирование

- Выполнить сборку.
- [Скачать](https://releases.1c.ru/project/Platform83) демонстрационную информационную базу (файл DT) технологической платформы 1С 8.3, распаковать файл 1cv8.dt в каталог Tests.
- Инициализировать окружение Tests\Init.cmd. Будет добавлен диск O: и развернута демонстрационная база в каталоге O:\Tests\ib
- Открыть файл o:\Tests\Test.xlsx. При успешном выполнении тестов, в ячейке B2 должно стоять PASS.