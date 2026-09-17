# Установщик

1. Соберите генератор и XLAM:

```powershell
dotnet publish .\SpecPdfGenerator.csproj -c Release -r win-x64 --self-contained true
powershell -ExecutionPolicy Bypass -File .\ExcelAddin\Build-ExcelAddin.ps1
```

2. Откройте `Installer\SpecPdfGenerator.iss` в Inno Setup Compiler и скомпилируйте его.

Установщик кладёт генератор и надстройку в `Program Files\SpecPdfGenerator`, регистрирует XLAM для текущего пользователя Excel и сохраняет стандартный путь к EXE. При обновлении выбранный пользователем путь не перезаписывается, если он существует.
