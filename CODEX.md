# SpecPdfGenerator — контекст для Codex

## Структура проекта

```text
SpecPdfGenerator/
├── Program.cs                         # Основная логика формирования PDF
├── SpecPdfGenerator.csproj            # .NET 8 Windows-проект
├── README.md                          # Пользовательская документация
├── INTEGRATION.md                     # Интеграция с Excel
├── CODEX.md                           # Контекст и правила для Codex
├── ExcelAddin/
│   ├── Build-ExcelAddin.ps1           # Создание/обновление XLAM
│   ├── SpecPdfGeneratorAddin.bas      # VBA callbacks и запуск генератора
│   ├── customUI.xml                   # Используемый Ribbon Office 2007
│   ├── customUI14.xml                 # Устаревший источник; не упаковывать
│   ├── SpecPdfGenerator.xlam          # Готовая Excel-надстройка
│   └── README.md
├── Installer/
│   ├── SpecPdfGenerator.iss           # Сценарий Inno Setup
│   ├── Output/
│   │   └── SpecPdfGenerator-Setup.exe # Собранный installer
│   └── README.md
├── bin/
│   ├── Debug/                         # Отладочная сборка
│   └── Release/
│       └── net8.0-windows/win-x64/publish/
│           └── SpecPdfGenerator.exe   # Канонический publish-output
├── obj/                               # Промежуточные файлы .NET
├── tmp/
│   └── pdfs/                          # Временные PDF
└── tools/                             # Вспомогательные инструменты
```

Ресурсы Ribbon лежат уровнем выше папки проекта (`..`): `импорт-pdf-32.png` и `настройки-32.png`.

## Назначение

`SpecPdfGenerator` формирует PDF-спецификацию из активной книги Excel.
Основной exe публикуется в:

`bin\Release\net8.0-windows\win-x64\publish\SpecPdfGenerator.exe`

Не запускать `SpecPdfGenerator.exe` самостоятельно: пользователь проверяет его вручную.

## Excel надстройка

- Исходники: `ExcelAddin\`.
- Итоговый файл: `ExcelAddin\SpecPdfGenerator.xlam`.
- VBA-модуль: `ExcelAddin\SpecPdfGeneratorAddin.bas` хранится в UTF-8.
- При полной сборке `Build-ExcelAddin.ps1` делает временную копию модуля в Windows-1251 перед `VBComponents.Import()`. Не убирать это преобразование: иначе русские MsgBox отображаются как mojibake.
- Для полной пересборки XLAM нужен закрытый Excel и включённая опция Excel «Доверять доступ к объектной модели проектов VBA».
- Если требуется только обновить Ribbon или иконки без переимпорта VBA, использовать:

  `powershell -ExecutionPolicy Bypass -File .\ExcelAddin\Build-ExcelAddin.ps1 -UpdateRibbonOnly`

## Ribbon

Используется только Office 2007 Custom UI:

- XML: `ExcelAddin\customUI.xml`.
- Namespace: `http://schemas.microsoft.com/office/2006/01/customui`.
- Корневой элемент должен иметь `onLoad="RibbonOnLoad"`.
- В XLAM должна быть часть `customUI/customUI.xml`; `customUI/customUI14.xml` не используется.
- В `[Content_Types].xml` требуется Override:

  `PartName="/customUI/customUI.xml"`, `ContentType="application/xml"`.

- В `_rels/.rels` должна быть ровно одна Ribbon relationship:

  `Id="rIdSpecPdfGeneratorRibbon"`,
  `Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"`,
  `Target="customUI/customUI.xml"`.

Текущая вкладка: `Teplov`; группа: `Спецификация`.

Кнопки `Сформировать PDF` и `Настройки` обе имеют `size="large"` и используют пользовательские иконки 32×32.

## Иконки

Исходные PNG лежат уровнем выше каталога проекта:

- `..\импорт-pdf-32.png`
- `..\настройки-32.png`

Скрипт сборки упаковывает их в XLAM как:

- `customUI/images/import-pdf-32.png`
- `customUI/images/configure-generator-32.png`

Связи изображений находятся в `customUI/_rels/customUI.xml.rels` и имеют ID `importPdfImage` и `configureGeneratorImage`.

## Установщик

- Скрипт: `Installer\SpecPdfGenerator.iss`.
- Итог: `Installer\Output\SpecPdfGenerator-Setup.exe`.
- XLAM устанавливается в `%APPDATA%\Microsoft\Excel\XLSTART\SpecPdfGenerator.xlam`.
- Через `Excel\Options\OPEN*` надстройка больше не регистрируется. Installer очищает только старые точные записи, созданные прежними версиями.
- `ExePath` в `HKCU\Software\VB and VBA Program Settings\SpecPdfGenerator\Settings` сохраняется при обновлении, если путь действителен.
- Перед сборкой installer закрыть Excel.

Сборка installer:

`& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" "Installer\SpecPdfGenerator.iss"`

После изменений XLAM сначала проверить его как ZIP: наличие VBA, `customUI/customUI.xml`, image relationships и отсутствие `customUI14.xml`.
