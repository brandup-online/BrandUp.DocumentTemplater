# BrandUp.WordGenerator

**BrandUp.WordGenerator** — фреймворк для создания .docx файлов из шаблонов.

[![Build Status](https://dev.azure.com/brandup/BrandUp%20Core/_apis/build/status%2FBrandUp%2Fbrandup-document-templater?branchName=master)](https://dev.azure.com/brandup/BrandUp%20Core/_build/latest?definitionId=72&branchName=master)

## Подготовка шаблона

Сначала нужно добавить вкладку "Разработчик" в ленту.
> На вкладке Файл перейдите в параметры > Настроить ленту.
> В разделе Настройка ленты в списке Основные вкладки установите флажок Разработчик.

Для того что бы создать поле в которое будут записаны данные, необходимо создать элемент управления и добавить в него тег с командой. Что бы создать элемент управления, на вкладке **Разработчик** в группе **Элементы управления**, слева выберите **Элемент управления содержимым**. Чтобы задать тег элемента управления, кликните на  элемент упраления, затем на вкладке **Разработчик** в группе **Элементы управления** нажмите кнопку **Свойства** и в поле **Тег** введите команду. В элементе управления можно создавать вложенные элементы управления, создавая сложную структуру документа.

## Комманды

Общий формат команды: `{commandName(Param1, Param2,... ParamN)}` где:
`commandName` - не регистрочувствительное имя команды, которая будет выполнена в элементе управления;
`Param1, Param2,... ParamN` - параметры команды, если это название свойства модели данных, то оно должно в точности повторять имя свойства в программе.

### Контекст данных

Контекст данных - это объект из которого команда получает данные. Контекст можно изменять с помощью команд.

### Типы команд

Команды разделяются по видам действия, которые они выполняют в документе:

- `None` - Команда не выполняет никаких действий в документе. Как правило служит для изменения контекста данных.
- `Content` - Команда заполняет элемент управления контентом.
- `List` - Команда клонирует элемент управления согласно количеству элементов а контексте и выполняет вложенные команды в элементе.

### Команды по умолчанию

Команды, которые поддерживаются изначально:

`setcontextofproperty(modelProperty)` - тип `None`. Устанавливает объект контекст данных.

- `modelProperty` - обязательный параметр. Контекст данных, который будет установлен.

`prop(modelProperty, format)` - тип `Content`. Команда записывает в элемент упраления данные.

- `modelProperty`- необязательный параметр. Свойство контекста, из которого будут взяты данные. Если свойства не указанно, то в элемент управление будет записан контекст (например, если контекст является объектом простого типа)
- `format` - необязательный параметр. [Формат](https://learn.microsoft.com/en-us/dotnet/standard/base-types/formatting-types) который будет применён к свойству при обращении его в строку.

`datetimenow(format)` - тип `Content`. устанавливает текущие дату и время.

- `format` - необязательный параметр. Формат даты.

`foreach(modelProperty)` - тип `List`. Создает контекст списка, для записи коллекции во вложенные элементы управления.

- `modelProperty` - необязательный параметр. свойство контекста данных, из которого будет взята коллекция. Контекст должен быть типа `IEnumerable`. Если параметер не установлен то коллекция будет создана из объекта контекста.
