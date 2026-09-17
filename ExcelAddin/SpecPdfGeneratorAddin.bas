Attribute VB_Name = "SpecPdfGeneratorAddin"
Option Explicit

Private Const SettingsApplication As String = "SpecPdfGenerator"
Private Const SettingsSection As String = "Settings"
Private Const SettingsKey As String = "ExePath"

Public Sub RibbonOnLoad(ByVal ribbon As IRibbonUI)
End Sub

Public Sub GeneratePdf(ByVal control As IRibbonControl)
    Dim workbook As Workbook
    Dim generatorPath As String
    Dim outputPath As String
    Dim shell As Object
    Dim exitCode As Long

    On Error GoTo Failed

    Set workbook = Application.ActiveWorkbook
    If workbook Is Nothing Or workbook Is ThisWorkbook Then
        MsgBox "Откройте сохранённую книгу Excel со спецификацией.", vbExclamation, "Спецификация"
        Exit Sub
    End If

    If Len(workbook.Path) = 0 Then
        MsgBox "Сначала сохраните файл Excel.", vbExclamation, "Спецификация"
        Exit Sub
    End If

    If Not workbook.Saved Then
        workbook.Save
        If Not workbook.Saved Then
            MsgBox "Не удалось сохранить текущую книгу. Формирование PDF отменено.", vbExclamation, "Спецификация"
            Exit Sub
        End If
    End If

    generatorPath = GetGeneratorPath()
    If Len(generatorPath) = 0 Then Exit Sub

    outputPath = BuildPdfPath(workbook.FullName)
    Set shell = CreateObject("WScript.Shell")
    exitCode = shell.Run(Quote(generatorPath) & " --input " & Quote(workbook.FullName) & " --output " & Quote(outputPath) & " --owner-hwnd " & CStr(Application.Hwnd), 1, True)

    If exitCode = 0 Then
        MsgBox "PDF успешно сформирован:" & vbCrLf & outputPath, vbInformation, "Спецификация"
    Else
        MsgBox "Ошибка формирования PDF. Код: " & CStr(exitCode), vbExclamation, "Спецификация"
    End If
    Exit Sub

Failed:
    MsgBox "Не удалось сформировать PDF: " & Err.Description, vbExclamation, "Спецификация"
End Sub

Public Sub ConfigureGenerator(ByVal control As IRibbonControl)
    Dim generatorPath As String
    generatorPath = PickGeneratorPath()
    If Len(generatorPath) = 0 Then Exit Sub

    SaveGeneratorPath generatorPath
    MsgBox "Путь к генератору сохранён:" & vbCrLf & generatorPath, vbInformation, "Спецификация"
End Sub

Private Function GetGeneratorPath() As String
    Dim generatorPath As String

    ' The current generator is distributed next to the XLAM in XLSTART.
    ' Prefer it to an old path remembered by earlier versions of the add-in.
    generatorPath = ThisWorkbook.Path & Application.PathSeparator & "SpecPdfGeneratorRuntime" & Application.PathSeparator & "SpecPdfGenerator.exe"
    If IsValidGenerator(generatorPath) Then
        GetGeneratorPath = generatorPath
        Exit Function
    End If

    generatorPath = GetSetting(SettingsApplication, SettingsSection, SettingsKey, vbNullString)
    If IsValidGenerator(generatorPath) Then
        GetGeneratorPath = generatorPath
        Exit Function
    End If

    generatorPath = PickGeneratorPath()
    If Len(generatorPath) = 0 Then Exit Function
    SaveSetting SettingsApplication, SettingsSection, SettingsKey, generatorPath
    GetGeneratorPath = generatorPath
End Function

Public Function PickGeneratorPath() As String
    Dim picker As FileDialog
    Set picker = Application.FileDialog(msoFileDialogFilePicker)
    With picker
        .Title = "Выберите SpecPdfGenerator.exe"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Исполняемые файлы", "*.exe"
        If .Show <> -1 Then Exit Function
        If Not IsValidGenerator(.SelectedItems(1)) Then
            MsgBox "Выберите существующий файл SpecPdfGenerator.exe.", vbExclamation, "Спецификация"
            Exit Function
        End If
        PickGeneratorPath = .SelectedItems(1)
    End With
End Function

Public Sub SaveGeneratorPath(ByVal generatorPath As String)
    SaveSetting SettingsApplication, SettingsSection, SettingsKey, generatorPath
End Sub

Private Function IsValidGenerator(ByVal filePath As String) As Boolean
    If Len(filePath) = 0 Or Len(Dir$(filePath)) = 0 Then Exit Function
    IsValidGenerator = (LCase$(Dir$(filePath)) = "specpdfgenerator.exe")
End Function

Private Function BuildPdfPath(ByVal workbookPath As String) As String
    Dim dotPosition As Long
    dotPosition = InStrRev(workbookPath, ".")
    If dotPosition > InStrRev(workbookPath, "\") Then
        BuildPdfPath = Left$(workbookPath, dotPosition - 1) & ".pdf"
    Else
        BuildPdfPath = workbookPath & ".pdf"
    End If
End Function

Private Function Quote(ByVal value As String) As String
    Quote = Chr$(34) & value & Chr$(34)
End Function
