## YellowQuery: надстройка Excel для запросов к 1С

YellowQuery позволяет писать формулы в Excel для получения данных из 1С. Поддерживается встроенный язык запросов 1С:Предприятие 8.

<img src="https://github.com/mcarrowd/YellowQuery/assets/4023864/ea602231-87f2-4422-a7b5-a1b45bee14f8 " data-canonical-src="https://github.com/mcarrowd/YellowQuery/4023864/ea602231-87f2-4422-a7b5-a1b45bee14f8 " width="600" height="450" />

### Идея

Большой помехой для эффективной работы с данными 1С в Excel является невозможность хранения в ячейках листа данных ссылочного типа. Эта надстройка сделана, чтобы проверить концепцию хранения таких данных в виде [навигационных ссылок](https://tinyurl.com/ytps9xyt).

### Функции

Описание функций приведено в [документации](https://github.com/mcarrowd/YellowQuery/wiki).

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