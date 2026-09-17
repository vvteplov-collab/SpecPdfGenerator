using System.Globalization;
using System.Diagnostics;
using System.Drawing.Text;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows.Forms;
using System.Xml.Linq;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace SpecPdfGenerator;

internal static class Program
{
    private const string FirstTemplateSheet = "Первый лист";
    private const string NextTemplateSheet = "Второй лист";
    private const string TempSheetPrefix = "__PDF_TEMP_";

    private const int RowsPerPage = 20;
    private const int TemplateRows = 46;
    private const int DataStartRow = 11;
    private const int PagesPerPdfBatch = 25;
    private const int MaxParallelExcelInstances = 2;
    private const string TopPageNumberRange = "CF4:CG4";

    // Excel constants (late binding: Microsoft.Office.Interop.Excel package is not required).
    private const int XlCalculationManual = -4135;
    private const int XlTypePdf = 0;
    // The workbook contains vector WMF objects rather than raster images, so the
    // minimum PDF quality preserves its visual result while reducing publishing work.
    private const int XlQualityMinimum = 1;
    private const int XlPaperA3 = 8;
    private const int XlLandscape = 2;
    private const int XlSheetVisible = -1;

    private static readonly string[] HeaderTokens =
    {
        "позиция",
        "наименование и техническая характеристика",
        "тип, марка",
        "код оборудования",
        "завод-изготовитель",
        "единица",
        "коли",
        "масса",
        "примеч"
    };

    private static readonly string SettingsFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "SpecPdfGenerator",
        "settings.json");

    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length == 1 && args[0].Equals("--version", StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine("SpecPdfGenerator 1.1");
            return 0;
        }

        if (!TryParseArguments(args, out string? inputArgument, out string? outputArgument,
                out bool preserveFormatting, out bool replaceFontFamily,
                out string fontName, out bool italic, out bool autoWrap, out IntPtr ownerWindowHandle))
        {
            PrintUsage();
            return 2;
        }

        string input = Path.GetFullPath(inputArgument!);
        if (!File.Exists(input))
        {
            Console.Error.WriteLine($"Файл не найден: {input}");
            return 3;
        }

        string output = outputArgument is not null
            ? Path.GetFullPath(outputArgument)
            : Path.Combine(Path.GetDirectoryName(input)!, Path.GetFileNameWithoutExtension(input) + ".pdf");

        Application.EnableVisualStyles();
        using var optionsWindow = new GenerationOptionsWindow(preserveFormatting, replaceFontFamily, fontName, italic, autoWrap);
        DialogResult optionsResult = ownerWindowHandle != IntPtr.Zero
            ? optionsWindow.ShowDialog(new ExternalWindow(ownerWindowHandle))
            : optionsWindow.ShowDialog();
        if (optionsResult != DialogResult.OK)
            return 0;
        preserveFormatting = optionsWindow.PreserveFormatting;
        replaceFontFamily = optionsWindow.ReplaceFontFamily;
        fontName = optionsWindow.FontName;
        italic = optionsWindow.Italic;
        autoWrap = optionsWindow.AutoWrap;
        SaveSettings(new GeneratorSettings(preserveFormatting, replaceFontFamily, fontName, italic, autoWrap));

        using var cancellation = new CancellationTokenSource();
        var stopSync = new object();
        Action? forceStopExcel = null;
        var progressWindow = new ProgressWindow(() =>
        {
            cancellation.Cancel();
            Action? stop;
            lock (stopSync) stop = forceStopExcel;
            stop?.Invoke();
        });
        var worker = new Thread(() =>
        {
            try
            {
                string savedOutput = Generate(input, output, preserveFormatting, replaceFontFamily, fontName, italic, autoWrap, cancellation.Token, progressWindow.Report, stop =>
                {
                    lock (stopSync) forceStopExcel = stop;
                });
                Console.WriteLine($"Готово: {savedOutput}");
                progressWindow.FinishSuccess(savedOutput);
            }
            catch (OperationCanceledException)
            {
                progressWindow.FinishCancelled();
            }
            catch (Exception) when (cancellation.IsCancellationRequested)
            {
                progressWindow.FinishCancelled();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("ОШИБКА:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
                progressWindow.FinishError(ex.Message);
            }
        });
        worker.SetApartmentState(ApartmentState.STA);
        worker.IsBackground = true;
        // Start only after the form has a window handle, so progress callbacks
        // are always marshalled to the UI thread.
        progressWindow.Shown += (_, _) => worker.Start();

        if (ownerWindowHandle != IntPtr.Zero)
            progressWindow.ShowDialog(new ExternalWindow(ownerWindowHandle));
        else
            progressWindow.ShowDialog();
        return progressWindow.ExitCode;
    }

    private static bool TryParseArguments(
        string[] args,
        out string? input,
        out string? output,
        out bool preserveFormatting,
        out bool replaceFontFamily,
        out string fontName,
        out bool italic,
        out bool autoWrap,
        out IntPtr ownerWindowHandle)
    {
        GeneratorSettings saved = LoadSettings();
        input = null;
        output = null;
        preserveFormatting = saved.PreserveFormatting;
        replaceFontFamily = saved.ReplaceFontFamily;
        fontName = saved.FontName;
        italic = saved.Italic;
        autoWrap = saved.AutoWrap;
        ownerWindowHandle = IntPtr.Zero;
        for (int index = 0; index < args.Length; index++)
        {
            string argument = args[index].Trim().Trim('"');
            if (argument.Equals("--input", StringComparison.OrdinalIgnoreCase))
            {
                if (++index >= args.Length || input is not null) return false;
                input = args[index].Trim().Trim('"');
            }
            else if (argument.Equals("--output", StringComparison.OrdinalIgnoreCase))
            {
                if (++index >= args.Length || output is not null) return false;
                output = args[index].Trim().Trim('"');
            }
            else if (argument.Equals("--inherit-font", StringComparison.OrdinalIgnoreCase))
            {
                preserveFormatting = true;
                replaceFontFamily = false;
            }
            else if (argument.Equals("--inherit-font-family", StringComparison.OrdinalIgnoreCase))
            {
                replaceFontFamily = false;
            }
            else if (argument.Equals("--inherit-font-style", StringComparison.OrdinalIgnoreCase))
            {
                preserveFormatting = true;
            }
            else if (argument.Equals("--inherit-font-weight", StringComparison.OrdinalIgnoreCase))
            {
                preserveFormatting = true;
            }
            else if (argument.Equals("--inherit-font-italic", StringComparison.OrdinalIgnoreCase))
            {
                preserveFormatting = true;
            }
            else if (argument.Equals("--font", StringComparison.OrdinalIgnoreCase))
            {
                if (++index >= args.Length) return false;
                fontName = args[index].Trim().Trim('"');
                if (string.IsNullOrWhiteSpace(fontName)) return false;
            }
            else if (argument.Equals("--no-italic", StringComparison.OrdinalIgnoreCase))
            {
                italic = false;
            }
            else if (argument.Equals("--auto-wrap", StringComparison.OrdinalIgnoreCase))
            {
                autoWrap = true;
            }
            else if (argument.Equals("--owner-hwnd", StringComparison.OrdinalIgnoreCase))
            {
                if (++index >= args.Length || ownerWindowHandle != IntPtr.Zero ||
                    !long.TryParse(args[index], NumberStyles.Integer, CultureInfo.InvariantCulture, out long handleValue) || handleValue == 0)
                    return false;
                ownerWindowHandle = new IntPtr(handleValue);
            }
            else if (argument.StartsWith("--", StringComparison.Ordinal))
            {
                return false;
            }
            else if (input is null)
            {
                input = argument;
            }
            else if (output is null)
            {
                output = argument;
            }
            else
            {
                return false;
            }
        }

        return !string.IsNullOrWhiteSpace(input);
    }

    private static GeneratorSettings LoadSettings()
    {
        try
        {
            if (!File.Exists(SettingsFilePath)) return new GeneratorSettings();
            GeneratorSettings saved = JsonSerializer.Deserialize<GeneratorSettings>(File.ReadAllText(SettingsFilePath)) ?? new GeneratorSettings();
            return saved with
            {
                PreserveFormatting = saved.PreserveFormatting || saved.InheritFontStyle || saved.InheritFontWeight || saved.InheritFontItalic,
                ReplaceFontFamily = saved.ReplaceFontFamily && !saved.InheritFontFamily
            };
        }
        catch
        {
            return new GeneratorSettings();
        }
    }

    private static void SaveSettings(GeneratorSettings settings)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsFilePath)!);
            File.WriteAllText(SettingsFilePath, JsonSerializer.Serialize(settings));
        }
        catch
        {
            // A read-only profile must not prevent PDF creation.
        }
    }

    private static void PrintUsage()
    {
        Console.WriteLine("Перетащи .xls/.xlsx/.xlsm файл на SpecPdfGenerator.exe");
        Console.WriteLine("или запусти: SpecPdfGenerator.exe --input \"C:\\Project\\Спецификация.xlsm\" [--output \"C:\\Project\\Спецификация.pdf\"] [--font \"ISOCPEUR\"]");
    }

    private static string Generate(
        string input,
        string output,
        bool preserveFormatting,
        bool replaceFontFamily,
        string fontName,
        bool italic,
        bool autoWrap,
        CancellationToken cancellationToken,
        Action<string, int, int>? reportProgress,
        Action<Action?> setForceStopExcel)
    {
        cancellationToken.ThrowIfCancellationRequested();
        Type? excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType is null)
            throw new InvalidOperationException("Microsoft Excel не найден. Нужен установленный настольный Excel.");

        dynamic? excel = null;
        dynamic? wb = null;
        string? preparedInput = null;

        try
        {
            reportProgress?.Invoke("Запускаю Excel…", 0, 0);
            excel = Activator.CreateInstance(excelType)!;
            cancellationToken.ThrowIfCancellationRequested();
            dynamic capturedExcel = excel;
            setForceStopExcel(() => StopExcel(capturedExcel));
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.ScreenUpdating = false;
            excel.EnableEvents = false;

            reportProgress?.Invoke("Открываю Excel…", 0, 0);
            Console.WriteLine("Открываю Excel...");
            // Excel can display a blocking "Name conflict" dialog while opening
            // workbooks copied from other files.  The conflicting name is usually
            // the built-in print area, which this generator replaces anyway.
            preparedInput = PrepareInputWithoutPrintAreas(input);
            reportProgress?.Invoke("Открываю книгу со спецификацией…", 0, 0);
            wb = excel.Workbooks.Open(preparedInput, UpdateLinks: 0, ReadOnly: true);
            cancellationToken.ThrowIfCancellationRequested();

            dynamic firstTemplate = GetSheet(wb, FirstTemplateSheet)
                ?? throw new InvalidOperationException($"Нет листа-шаблона '{FirstTemplateSheet}'.");
            dynamic nextTemplate = GetSheet(wb, NextTemplateSheet)
                ?? throw new InvalidOperationException($"Нет листа-шаблона '{NextTemplateSheet}'.");
            excel.Calculation = XlCalculationManual;

            reportProgress?.Invoke("Ищу вкладки спецификации…", 0, 0);
            var sourceSheets = FindSourceSheets(wb);
            if (sourceSheets.Count == 0)
                throw new InvalidOperationException("Не найдено ни одной исходной вкладки спецификации.");

            // Do not pass a lambda over `dynamic` values to LINQ: the compiler
            // would treat it as a dynamically dispatched operation.
            var sourceSheetNames = new List<string>(sourceSheets.Count);
            foreach (dynamic sourceSheet in sourceSheets)
                sourceSheetNames.Add((string)sourceSheet.Name);

            Console.WriteLine("Исходные вкладки: " + string.Join(" → ", sourceSheetNames));

            reportProgress?.Invoke($"Читаю вкладки: 0 из {sourceSheets.Count}", 0, sourceSheets.Count);

            // Reading Font/Interior through Excel COM for every populated cell is
            // extremely expensive.  For OOXML workbooks (.xlsx/.xlsm/.xltx/.xltm)
            // read only the cell style ids from the package, then ask Excel for the
            // actual visual properties once per unique style.  This keeps Excel's
            // own interpretation of fills/underline intact while eliminating
            // thousands of COM round-trips.
            WorkbookCellStyleIndex? sourceStyleIndex = preserveFormatting
                ? WorkbookCellStyleIndex.TryLoad(preparedInput)
                : null;
            var styleFormattingCache = new Dictionary<int, CellFont>();
            Console.WriteLine(sourceStyleIndex is null
                ? "Форматирование: совместимый COM-режим."
                : "Форматирование: быстрый кэш по Excel style ID.");

            var sections = new List<SpecificationSection>(sourceSheets.Count);
            for (int sheetIndex = 0; sheetIndex < sourceSheets.Count; sheetIndex++)
            {
                dynamic sheet = sourceSheets[sheetIndex];
                cancellationToken.ThrowIfCancellationRequested();
                string sheetName = (string)sheet.Name;
                reportProgress?.Invoke($"Читаю вкладку: {sheetName} ({sheetIndex + 1} из {sourceSheets.Count})", sheetIndex, sourceSheets.Count);
                var fontOptions = new FontReadOptions(preserveFormatting);
                var part = ReadSpecificationRows(sheet, fontOptions, sourceStyleIndex, styleFormattingCache);
                Console.WriteLine($"  {sheet.Name}: {part.Count} строк");
                if (part.Count > 0 && part[0].IsHeading &&
                    string.Equals(Normalize(part[0].Fields[1]), Normalize(sheetName), StringComparison.Ordinal))
                    part.RemoveAt(0);
                sections.Add(new SpecificationSection(sheetName, part));
                reportProgress?.Invoke($"Прочитано вкладок: {sheetIndex + 1} из {sourceSheets.Count}", sheetIndex + 1, sourceSheets.Count);
            }

            if (preserveFormatting && sourceStyleIndex is not null)
                Console.WriteLine($"Уникальных исходных стилей прочитано через Excel: {styleFormattingCache.Count}");

            if (sections.All(section => section.Rows.Count == 0))
                throw new InvalidOperationException("В исходных вкладках не найдено строк спецификации.");

            int rowCount = sections.Sum(section => section.Rows.Count);
            Console.WriteLine($"Всего позиций/заголовков: {rowCount}");
            return GenerateNativePdf(firstTemplate, nextTemplate, sections, output, preserveFormatting, replaceFontFamily, fontName, italic, autoWrap, cancellationToken, reportProgress);
        }
        finally
        {
            setForceStopExcel(null);
            if (wb is not null)
            {
                try { wb.Close(SaveChanges: false); } catch { }
            }

            if (excel is not null)
            {
                try { excel.Quit(); } catch { }
            }

            ReleaseCom(wb);
            ReleaseCom(excel);
            if (preparedInput is not null &&
                !string.Equals(preparedInput, input, StringComparison.OrdinalIgnoreCase))
            {
                try { File.Delete(preparedInput); } catch { }
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private static List<dynamic> FindSourceSheets(dynamic wb)
    {
        var result = new List<dynamic>();

        // Worksheets are indexed in the same order as Excel tabs (left to right).
        for (int i = 1; i <= wb.Worksheets.Count; i++)
        {
            dynamic s = wb.Worksheets[i];
            if (IsVisibleSourceSheet(s))
                result.Add(s);
        }

        return result;
    }

    private static bool IsVisibleSourceSheet(dynamic sheet)
    {
        string name = ((string)sheet.Name).Trim();
        if (name.Equals(FirstTemplateSheet, StringComparison.OrdinalIgnoreCase) ||
            name.Equals(NextTemplateSheet, StringComparison.OrdinalIgnoreCase))
            return false;

        // Hidden worksheets are not Excel tabs and are intentionally excluded.
        try
        {
            if ((int)sheet.Visible != XlSheetVisible)
                return false;
        }
        catch { }

        return true;
    }

    private static List<SpecRow> ReadSpecificationRows(
        dynamic sheet,
        FontReadOptions fontOptions,
        WorkbookCellStyleIndex? styleIndex,
        IDictionary<int, CellFont> styleFormattingCache)
    {
        dynamic used = sheet.UsedRange;
        int firstRow = used.Row;
        int firstCol = used.Column;
        object?[,]? values = used.Value2 as object[,];
        if (values is null) return new List<SpecRow>();

        int rowCount = values.GetLength(0);
        int colCount = values.GetLength(1);

        var header = FindHeaders(values, firstRow, firstCol);
        if (header is null)
        {
            Console.WriteLine($"  {sheet.Name}: шапка не найдена, читаю данные как экспортную таблицу.");
            return ReadHeaderlessRows(sheet, values, firstRow, firstCol, fontOptions, styleIndex, styleFormattingCache);
        }

        var result = new List<SpecRow>();
        int localStart = header.HeaderRow - firstRow + 2; // row after header anchor; header itself may span several rows
        localStart = Math.Max(1, localStart);
        FontColumnDefaults[] fontDefaults = fontOptions.Any
            ? ReadColumnFontDefaults(sheet, header.Columns, firstRow + localStart - 1, firstRow + rowCount - 1, fontOptions)
            : [];

        for (int lr = localStart; lr <= rowCount; lr++)
        {
            int absoluteRow = firstRow + lr - 1;
            var fields = new string[9];
            for (int i = 0; i < 9; i++)
            {
                int lc = header.Columns[i] - firstCol + 1;
                if (lc >= 1 && lc <= colCount)
                    fields[i] = ToText(values[lr, lc]);
                else
                    fields[i] = "";
            }

            if (IsHeaderOrNumberRow(fields))
                continue;

            bool blank = fields.All(string.IsNullOrWhiteSpace);
            if (blank)
            {
                // We do not need empty separators in the final compact PDF. Section headings themselves are preserved.
                continue;
            }

            result.Add(new SpecRow(fields, fontOptions.Any
                ? ReadFonts(sheet, absoluteRow, header.Columns, fields, fontOptions, fontDefaults, styleIndex, styleFormattingCache)
                : null));
        }

        return result;
    }

    // Some exports contain the nine specification fields but no GOST header row.
    // Their first nine used columns are: position, name, type, code, manufacturer,
    // unit, quantity, mass, note. Empty rows are ignored; one-cell rows become
    // section headings through SpecRow.IsHeading.
    private static List<SpecRow> ReadHeaderlessRows(
        dynamic sheet,
        object?[,] values,
        int firstRow,
        int firstCol,
        FontReadOptions fontOptions,
        WorkbookCellStyleIndex? styleIndex,
        IDictionary<int, CellFont> styleFormattingCache)
    {
        int rowCount = values.GetLength(0);
        int colCount = values.GetLength(1);
        var result = new List<SpecRow>();
        int[] columns = Enumerable.Range(0, 9).Select(index => firstCol + index).ToArray();
        FontColumnDefaults[] fontDefaults = fontOptions.Any
            ? ReadColumnFontDefaults(sheet, columns, firstRow, firstRow + rowCount - 1, fontOptions)
            : [];

        for (int row = 1; row <= rowCount; row++)
        {
            var fields = new string[9];
            for (int col = 0; col < fields.Length && col < colCount; col++)
                fields[col] = ToText(values[row, col + 1]);

            if (!fields.All(string.IsNullOrWhiteSpace))
            {
                result.Add(new SpecRow(fields, fontOptions.Any
                    ? ReadFonts(sheet, firstRow + row - 1, columns, fields, fontOptions, fontDefaults, styleIndex, styleFormattingCache)
                    : null));
            }
        }

        return result;
    }

    private static CellFont[] ReadFonts(
        dynamic sheet,
        int row,
        IReadOnlyList<int> columns,
        IReadOnlyList<string> fields,
        FontReadOptions options,
        IReadOnlyList<FontColumnDefaults> defaults,
        WorkbookCellStyleIndex? styleIndex,
        IDictionary<int, CellFont> styleFormattingCache)
    {
        var fonts = new CellFont[columns.Count];
        string sheetName = (string)sheet.Name;

        for (int index = 0; index < columns.Count; index++)
        {
            // A blank value is not drawn into the PDF, so querying its COM
            // formatting only slows large specifications down. The template
            // already supplies the background for blank table cells.
            if (string.IsNullOrWhiteSpace(fields[index]))
                continue;

            try
            {
                int column = columns[index];

                // Fast path: in OOXML files the style id is stored directly in
                // worksheet XML.  Styles are workbook-global, so cells sharing a
                // style id also share the formatting properties used here.
                if (styleIndex is not null &&
                    styleIndex.TryGetStyleId(sheetName, row, column, out int styleId))
                {
                    if (!styleFormattingCache.TryGetValue(styleId, out CellFont cached))
                    {
                        cached = ReadCellFont(sheet.Cells[row, column], options);
                        styleFormattingCache[styleId] = cached;
                    }
                    fonts[index] = cached;
                    continue;
                }

                // Fallback for legacy/non-OOXML workbooks.  Uniform column
                // properties are reused instead of requesting the same COM
                // property for every cell.
                FontColumnDefaults columnDefaults = defaults[index];
                bool needsCellFont =
                    (options.Weight && columnDefaults.Bold is null) ||
                    (options.Italic && columnDefaults.Italic is null) ||
                    (options.Underline && columnDefaults.Underline is null);
                bool needsCellFill = options.Fill && !columnDefaults.HasUniformFill;
                dynamic? cell = needsCellFont || needsCellFill ? sheet.Cells[row, column] : null;
                dynamic? font = needsCellFont ? cell!.Font : null;

                bool bold = options.Weight &&
                    (columnDefaults.Bold ?? Convert.ToBoolean(font!.Bold, CultureInfo.InvariantCulture));
                bool italic = options.Italic &&
                    (columnDefaults.Italic ?? Convert.ToBoolean(font!.Italic, CultureInfo.InvariantCulture));
                bool underline = options.Underline &&
                    (columnDefaults.Underline ?? ReadUnderline(font!.Underline));
                XColor? fill = !options.Fill
                    ? null
                    : columnDefaults.HasUniformFill
                        ? columnDefaults.Fill
                        : ReadFillColor(cell!);

                fonts[index] = new CellFont(bold, italic, underline, fill);
            }
            catch
            {
                fonts[index] = new CellFont(false, false, false, null);
            }
        }

        return fonts;
    }

    private static CellFont ReadCellFont(dynamic cell, FontReadOptions options)
    {
        dynamic font = cell.Font;
        bool bold = options.Weight && Convert.ToBoolean(font.Bold, CultureInfo.InvariantCulture);
        bool italic = options.Italic && Convert.ToBoolean(font.Italic, CultureInfo.InvariantCulture);
        bool underline = options.Underline && ReadUnderline(font.Underline);
        XColor? fill = options.Fill ? ReadFillColor(cell) : null;
        return new CellFont(bold, italic, underline, fill);
    }

    private static FontColumnDefaults[] ReadColumnFontDefaults(
        dynamic sheet,
        IReadOnlyList<int> columns,
        int firstRow,
        int lastRow,
        FontReadOptions options)
    {
        var defaults = new FontColumnDefaults[columns.Count];
        for (int index = 0; index < columns.Count; index++)
        {
            try
            {
                dynamic range = sheet.Range[sheet.Cells[firstRow, columns[index]], sheet.Cells[lastRow, columns[index]]];
                dynamic font = range.Font;
                bool hasUniformFill = TryReadUniformFillColor(range, out XColor? fill);
                defaults[index] = new FontColumnDefaults(
                    options.Weight ? ReadUniformBool(font.Bold) : false,
                    options.Italic ? ReadUniformBool(font.Italic) : false,
                    options.Underline ? ReadUniformUnderline(font.Underline) : false,
                    hasUniformFill,
                    fill);
            }
            catch
            {
                defaults[index] = new FontColumnDefaults(null, null, null, false, null);
            }
        }
        return defaults;
    }

    private static bool? ReadUniformBool(object? value) =>
        value is null or DBNull ? null : Convert.ToBoolean(value, CultureInfo.InvariantCulture);

    private static bool? ReadUniformUnderline(object? value)
    {
        if (value is null or DBNull) return null;
        return ReadUnderline(value);
    }

    private static bool ReadUnderline(object value) =>
        value is bool underline ? underline : Convert.ToInt32(value, CultureInfo.InvariantCulture) != -4142; // xlUnderlineStyleNone

    private static XColor? ReadFillColor(dynamic cell)
    {
        // Excel Interior.Color is an OLE color value: BBGGRR. -4142 means no fill.
        object? pattern = cell.Interior.Pattern;
        if (pattern is null or DBNull || Convert.ToInt32(pattern, CultureInfo.InvariantCulture) == -4142)
            return null;
        object? value = cell.Interior.Color;
        if (value is null or DBNull) return null;
        int oleColor = Convert.ToInt32(value, CultureInfo.InvariantCulture);
        if (oleColor < 0 || oleColor == -4142) return null;
        return XColor.FromArgb(oleColor & 0xFF, (oleColor >> 8) & 0xFF, (oleColor >> 16) & 0xFF);
    }

    private static bool TryReadUniformFillColor(dynamic range, out XColor? fill)
    {
        fill = null;
        try
        {
            dynamic interior = range.Interior;
            object? pattern = interior.Pattern;
            if (pattern is null or DBNull) return false; // Different fills in the range.
            if (Convert.ToInt32(pattern, CultureInfo.InvariantCulture) == -4142)
                return true; // Uniform: no fill.
            object? value = interior.Color;
            if (value is null or DBNull) return false;
            int oleColor = Convert.ToInt32(value, CultureInfo.InvariantCulture);
            if (oleColor < 0 || oleColor == -4142) return false;
            fill = XColor.FromArgb(oleColor & 0xFF, (oleColor >> 8) & 0xFF, (oleColor >> 16) & 0xFF);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static HeaderMap? FindHeaders(object?[,] values, int firstRow, int firstCol)
    {
        int rows = Math.Min(values.GetLength(0), 80);
        int cols = Math.Min(values.GetLength(1), 240);
        int[] foundCols = new int[9];
        int headerRow = -1;

        for (int r = 1; r <= rows; r++)
        {
            var rowFound = new int[9];
            for (int c = 1; c <= cols; c++)
            {
                string text = Normalize(values[r, c]);
                if (text.Length == 0) continue;

                for (int h = 0; h < HeaderTokens.Length; h++)
                {
                    if (rowFound[h] == 0 && text.Contains(HeaderTokens[h]))
                        rowFound[h] = firstCol + c - 1;
                }
            }

            // Header labels in the supplied GOST form are anchored on one row (merged vertically afterwards).
            if (rowFound.Count(x => x > 0) >= 7 && rowFound[0] > 0 && rowFound[1] > 0)
            {
                // Fill any missing columns by looking in neighboring rows and then by monotonic inference.
                Array.Copy(rowFound, foundCols, 9);
                headerRow = firstRow + r - 1;
                break;
            }
        }

        if (headerRow < 0) return null;

        // Search +/-5 rows for missing header anchors.
        for (int h = 0; h < 9; h++)
        {
            if (foundCols[h] > 0) continue;
            for (int r = Math.Max(1, headerRow - firstRow - 4); r <= Math.Min(rows, headerRow - firstRow + 6); r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    string text = Normalize(values[r, c]);
                    if (text.Contains(HeaderTokens[h]))
                    {
                        foundCols[h] = firstCol + c - 1;
                        break;
                    }
                }
                if (foundCols[h] > 0) break;
            }
        }

        if (foundCols.Any(x => x <= 0))
            return null;

        return new HeaderMap(headerRow, foundCols);
    }

    private static bool IsHeaderOrNumberRow(string[] f)
    {
        string joined = Normalize(string.Join(" | ", f));
        if (joined.Contains("наименование и техническая характеристика") || joined.Contains("завод-изготовитель"))
            return true;

        int numeric = 0;
        for (int i = 0; i < 9; i++)
        {
            if (f[i].Trim() == (i + 1).ToString(CultureInfo.InvariantCulture))
                numeric++;
        }
        return numeric >= 7;
    }

    private static List<List<SpecRow?>> Paginate(List<SpecRow> rows, int capacity)
    {
        var pages = new List<List<SpecRow?>>();
        var current = new List<SpecRow?>(capacity);

        foreach (var row in rows)
        {
            bool heading = row.IsHeading;

            // Do not leave a section/system heading as the last line on an A3 page.
            if (heading && current.Count >= capacity - 1)
            {
                while (current.Count < capacity) current.Add(null);
                pages.Add(current);
                current = new List<SpecRow?>(capacity);
            }

            if (current.Count >= capacity)
            {
                pages.Add(current);
                current = new List<SpecRow?>(capacity);
            }

            current.Add(row);
        }

        if (current.Count > 0)
        {
            while (current.Count < capacity) current.Add(null);
            pages.Add(current);
        }

        return pages;
    }

    private static List<List<SpecRow?>> PaginateSections(
        IReadOnlyList<SpecificationSection> sections,
        int capacity)
    {
        var pages = new List<List<SpecRow?>>();
        foreach (var section in sections)
        {
            // Every worksheet gets its own A3 page. Rows 1 and 3 remain blank,
            // row 2 is the worksheet title, and the actual table starts at row 4.
            var current = new List<SpecRow?>(capacity)
            {
                null,
                SpecRow.SheetTitle(section.Name),
                null
            };

            for (int rowIndex = 0; rowIndex < section.Rows.Count; rowIndex++)
            {
                SpecRow row = section.Rows[rowIndex];
                if (row.IsHeading && current.Count >= capacity - 1)
                {
                    while (current.Count < capacity) current.Add(null);
                    pages.Add(current);
                    current = new List<SpecRow?>(capacity);
                }

                // Keep all physical lines of an auto-wrapped name on one page.
                if (row.RowSpan > 1 && row.RowSpan <= capacity && current.Count + row.RowSpan > capacity)
                {
                    while (current.Count < capacity) current.Add(null);
                    pages.Add(current);
                    current = new List<SpecRow?>(capacity);
                }

                if (current.Count >= capacity)
                {
                    pages.Add(current);
                    current = new List<SpecRow?>(capacity);
                }

                current.Add(row);
            }

            while (current.Count < capacity) current.Add(null);
            pages.Add(current);
        }
        return pages;
    }

    private static void FillIndexColumn(dynamic sheet, int pageIndex)
    {
        object[,] keys = new object[RowsPerPage, 1];
        int firstKey = pageIndex * RowsPerPage + 1;
        for (int i = 0; i < RowsPerPage; i++) keys[i, 0] = firstKey + i;

        int r1 = DataStartRow;
        int r2 = r1 + RowsPerPage - 1;
        sheet.Range[$"B{r1}:B{r2}"].Value2 = keys;
    }

    // The templates formerly used VLOOKUP formulas referring to a helper worksheet.
    // Put the generator's in-memory data straight into the copied page instead.
    private static void FillDataRows(dynamic sheet, IReadOnlyList<SpecRow?> page, int offset)
    {
        string[] columns = ["G", "K", "AK", "AW", "BD", "BM", "BQ", "BU", "BZ"];
        int firstRow = offset + DataStartRow;
        int lastRow = firstRow + RowsPerPage - 1;

        for (int field = 0; field < columns.Length; field++)
        {
            object[,] values = new object[RowsPerPage, 1];
            for (int row = 0; row < RowsPerPage; row++)
                values[row, 0] = page[row]?.Fields[field] ?? "";

            sheet.Range[$"{columns[field]}{firstRow}:{columns[field]}{lastRow}"].Value2 = values;
        }
    }

    private static string PrepareInputWithoutPrintAreas(string input)
    {
        string extension = Path.GetExtension(input);
        if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".xltx", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".xltm", StringComparison.OrdinalIgnoreCase))
            return input;

        string temporaryPath = Path.Combine(
            Path.GetTempPath(),
            "SpecPdfGenerator_" + Guid.NewGuid().ToString("N") + extension);

        try
        {
            File.Copy(input, temporaryPath);
            using var archive = ZipFile.Open(temporaryPath, ZipArchiveMode.Update);
            ZipArchiveEntry? workbookEntry = archive.GetEntry("xl/workbook.xml");
            if (workbookEntry is null)
                return temporaryPath;

            XDocument document;
            using (Stream stream = workbookEntry.Open())
                document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);

            var printAreaNames = document
                .Descendants()
                .Where(element => element.Name.LocalName == "definedName" &&
                    (string.Equals((string?)element.Attribute("name"), "_xlnm.Print_Area", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals((string?)element.Attribute("name"), "Print_Area", StringComparison.OrdinalIgnoreCase)))
                .ToList();

            if (printAreaNames.Count == 0)
                return temporaryPath;

            foreach (XElement printAreaName in printAreaNames)
                printAreaName.Remove();

            workbookEntry.Delete();
            ZipArchiveEntry rewrittenEntry = archive.CreateEntry("xl/workbook.xml", CompressionLevel.Optimal);
            using Stream outputStream = rewrittenEntry.Open();
            document.Save(outputStream, SaveOptions.DisableFormatting);
            Console.WriteLine($"Удалены конфликтующие области печати: {printAreaNames.Count}.");
            return temporaryPath;
        }
        catch
        {
            try { File.Delete(temporaryPath); } catch { }
            throw;
        }
    }

    private static void ConfigurePrintPage(dynamic sheet)
    {
        dynamic ps = sheet.PageSetup;
        ps.PrintArea = "$C$3:$CH$46";
        ps.PaperSize = XlPaperA3;
        ps.Orientation = XlLandscape;
        ps.Zoom = 91;
        ps.LeftMargin = 0;
        ps.RightMargin = 0;
        ps.TopMargin = 0;
        ps.BottomMargin = 0;
        ps.HeaderMargin = 0;
        ps.FooterMargin = 0;
    }

    private static string MakeTempPageName(int pageNumber)
    {
        return TempSheetPrefix + pageNumber.ToString("000", CultureInfo.InvariantCulture);
    }

    private static void HideNonGeneratedSheets(dynamic wb, HashSet<string> generatedPageNames)
    {
        foreach (dynamic sheet in wb.Worksheets)
        {
            string name = (string)sheet.Name;
            if (!generatedPageNames.Contains(name))
                sheet.Visible = 0; // xlSheetHidden
        }
    }

    private static bool TryGetPdf24DocTool(out string path)
    {
        path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "PDF24", "pdf24-DocTool.exe");
        if (File.Exists(path)) return true;

        path = string.Empty;
        return false;
    }

    private static string GenerateNativePdf(
        dynamic firstTemplate,
        dynamic nextTemplate,
        IReadOnlyList<SpecificationSection> sections,
        string output,
        bool preserveFormatting,
        bool replaceFontFamily,
        string fontName,
        bool italic,
        bool autoWrap,
        CancellationToken cancellationToken,
        Action<string, int, int>? reportProgress)
    {
        string temporaryDirectory = Path.Combine(Path.GetTempPath(), "SpecPdfGenerator_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(temporaryDirectory);

        try
        {
            // The first template holds the starting sheet number. It continues
            // across every generated page, including pages based on the second template.
            int? topPageNumberBase = ReadTopPageNumber(firstTemplate);
            if (topPageNumberBase.HasValue)
            {
                ClearTopPageNumber(firstTemplate);
                ClearTopPageNumber(nextTemplate);
            }

            ConfigurePrintPage(firstTemplate);
            ConfigurePrintPage(nextTemplate);
            string firstBackground = Path.Combine(temporaryDirectory, "first.pdf");
            string nextBackground = Path.Combine(temporaryDirectory, "next.pdf");
            reportProgress?.Invoke("Экспортирую шаблон первого листа…", 0, 0);
            firstTemplate.ExportAsFixedFormat(Type: XlTypePdf, Filename: firstBackground, Quality: XlQualityMinimum,
                IncludeDocProperties: false, IgnorePrintAreas: false, OpenAfterPublish: false);
            reportProgress?.Invoke("Экспортирую шаблон последующих листов…", 0, 0);
            nextTemplate.ExportAsFixedFormat(Type: XlTypePdf, Filename: nextBackground, Quality: XlQualityMinimum,
                IncludeDocProperties: false, IgnorePrintAreas: false, OpenAfterPublish: false);

            reportProgress?.Invoke("Считываю размеры таблицы…", 0, 0);
            var firstLayout = TemplateLayout.Create(firstTemplate, firstBackground, true, topPageNumberBase);
            var nextLayout = TemplateLayout.Create(nextTemplate, nextBackground, false, topPageNumberBase);
            IReadOnlyList<SpecificationSection> outputSections = sections;
            if (autoWrap)
            {
                reportProgress?.Invoke("Рассчитываю автоперенос наименований…", 0, 0);
                using var measurementDocument = new PdfDocument();
                var measurementPage = measurementDocument.AddPage();
                using var measurementGraphics = XGraphics.FromPdfPage(measurementPage);
                outputSections = firstLayout.ExpandSectionsForAutoWrap(
                    sections,
                    measurementGraphics,
                    preserveFormatting,
                    replaceFontFamily,
                    fontName,
                    italic);
            }
            reportProgress?.Invoke("Разбиваю спецификацию на страницы…", 0, 0);
            var pages = PaginateSections(outputSections, RowsPerPage);
            Console.WriteLine($"Листов А3: {pages.Count}");
            reportProgress?.Invoke($"Подготавливаю страницы: 0 из {pages.Count}", 0, pages.Count);
            using var document = new PdfDocument();
            var firstForm = XPdfForm.FromFile(firstBackground);
            var nextForm = XPdfForm.FromFile(nextBackground);
            for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                bool isFirstPage = pageIndex == 0;
                var layout = isFirstPage ? firstLayout : nextLayout;
                var background = isFirstPage ? firstForm : nextForm;
                var page = document.AddPage();
                page.Width = XUnit.FromPoint(background.PointWidth);
                page.Height = XUnit.FromPoint(background.PointHeight);

                using var graphics = XGraphics.FromPdfPage(page);
                graphics.DrawImage(background, 0, 0, background.PointWidth, background.PointHeight);
                layout.DrawRows(graphics, pages[pageIndex], preserveFormatting, replaceFontFamily, fontName, italic);
                layout.DrawStamp(graphics, pageIndex + 1, pages.Count, pageIndex);

                reportProgress?.Invoke($"Собираю PDF: {pageIndex + 1} из {pages.Count}", pageIndex + 1, pages.Count);
            }

            cancellationToken.ThrowIfCancellationRequested();
            string savedOutput = PrepareOutputPath(output);
            document.Save(savedOutput);
            return savedOutput;
        }
        finally
        {
            if (Directory.Exists(temporaryDirectory))
                Directory.Delete(temporaryDirectory, recursive: true);
        }
    }

    private static string PrepareOutputPath(string requestedOutput)
    {
        string directory = Path.GetDirectoryName(requestedOutput)!;
        Directory.CreateDirectory(directory);
        if (!File.Exists(requestedOutput)) return requestedOutput;

        try
        {
            File.Delete(requestedOutput);
            return requestedOutput;
        }
        catch (IOException)
        {
            string fileName = Path.GetFileNameWithoutExtension(requestedOutput);
            string extension = Path.GetExtension(requestedOutput);
            for (int copyNumber = 1; ; copyNumber++)
            {
                string copyPath = Path.Combine(directory, $"{fileName} ({copyNumber}){extension}");
                if (!File.Exists(copyPath)) return copyPath;
            }
        }
    }

    private static int? ReadTopPageNumber(dynamic sheet)
    {
        foreach (dynamic cell in sheet.Range[TopPageNumberRange].Cells)
        {
            object? value = cell.Value2;
            if (value is null || value is DBNull) continue;

            if (value is double number && Math.Abs(number - Math.Round(number)) < 0.000001)
                return Convert.ToInt32(Math.Round(number));

            if (int.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture),
                    NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed))
                return parsed;
        }

        return null;
    }

    private static void ClearTopPageNumber(dynamic sheet)
    {
        sheet.Range[TopPageNumberRange].ClearContents();
    }

    private static void GeneratePdfInParallel(
        string input,
        string output,
        List<List<SpecRow?>> pages,
        string pdf24DocTool,
        CancellationToken cancellationToken,
        Action<string, int, int>? reportProgress)
    {
        string outputDirectory = Path.GetDirectoryName(output)!;
        string temporaryDirectory = Path.Combine(outputDirectory, ".SpecPdfGenerator_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(temporaryDirectory);

        try
        {
            var batches = new List<PdfBatch>();
            for (int firstPage = 0; firstPage < pages.Count; firstPage += PagesPerPdfBatch)
            {
                int count = Math.Min(PagesPerPdfBatch, pages.Count - firstPage);
                batches.Add(new PdfBatch(
                    batches.Count,
                    firstPage,
                    pages.GetRange(firstPage, count),
                    Path.Combine(temporaryDirectory, $"part_{batches.Count:000}.pdf")));
            }

            int completedPages = 0;
            using var slots = new SemaphoreSlim(MaxParallelExcelInstances);
            var tasks = batches.Select(batch => Task.Run(async () =>
            {
                await slots.WaitAsync(cancellationToken).ConfigureAwait(false);
                try
                {
                    await RunOnStaThreadAsync(() => CreatePdfPart(input, batch, pages.Count, cancellationToken)).ConfigureAwait(false);
                    int completed = Interlocked.Add(ref completedPages, batch.Pages.Count);
                    reportProgress?.Invoke($"Подготовлено страниц: {completed} из {pages.Count}", completed, pages.Count);
                }
                finally
                {
                    slots.Release();
                }
            })).ToArray();

            Task.WhenAll(tasks).GetAwaiter().GetResult();
            cancellationToken.ThrowIfCancellationRequested();

            reportProgress?.Invoke("Объединяю PDF-пакеты…", 0, 0);
            JoinPdfParts(pdf24DocTool, batches.Select(batch => batch.OutputPath), output, cancellationToken);
        }
        finally
        {
            // This directory is created exclusively for the current run.
            if (Directory.Exists(temporaryDirectory))
                Directory.Delete(temporaryDirectory, recursive: true);
        }
    }

    private static Task RunOnStaThreadAsync(Action action)
    {
        var completion = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            try
            {
                action();
                completion.SetResult();
            }
            catch (Exception ex)
            {
                completion.SetException(ex);
            }
        })
        {
            IsBackground = true
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        return completion.Task;
    }

    private static void CreatePdfPart(string input, PdfBatch batch, int totalPageCount, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        Type? excelType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException("Microsoft Excel не найден.");

        dynamic? excel = null;
        dynamic? wb = null;
        using var cancellationRegistration = cancellationToken.Register(() =>
        {
            if (excel is not null) StopExcel(excel);
        });

        try
        {
            excel = Activator.CreateInstance(excelType)!;
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.ScreenUpdating = false;
            excel.EnableEvents = false;
            wb = excel.Workbooks.Open(input, UpdateLinks: 0, ReadOnly: true);

            dynamic firstTemplate = GetSheet(wb, FirstTemplateSheet)
                ?? throw new InvalidOperationException($"Нет листа-шаблона '{FirstTemplateSheet}'.");
            dynamic nextTemplate = GetSheet(wb, NextTemplateSheet)
                ?? throw new InvalidOperationException($"Нет листа-шаблона '{NextTemplateSheet}'.");
            ConfigurePrintPage(firstTemplate);
            ConfigurePrintPage(nextTemplate);

            string firstDocNo = Convert.ToString(firstTemplate.Range["BJ35"].Value2, CultureInfo.InvariantCulture) ?? "";
            string nextDocNo = Convert.ToString(nextTemplate.Range["BJ43"].Value2, CultureInfo.InvariantCulture) ?? firstDocNo;
            var generatedPageNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int localPage = 0; localPage < batch.Pages.Count; localPage++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int globalPage = batch.FirstPage + localPage;
                bool isFirstPage = globalPage == 0;
                dynamic template = isFirstPage ? firstTemplate : nextTemplate;
                template.Copy(After: wb.Sheets[wb.Sheets.Count]);
                dynamic page = wb.ActiveSheet;
                page.Name = MakeTempPageName(localPage + 1);
                generatedPageNames.Add((string)page.Name);
                FillIndexColumn(page, globalPage);
                FillDataRows(page, batch.Pages[localPage], offset: 0);

                if (isFirstPage)
                {
                    page.Range["CA41"].Value2 = 1;
                    page.Range["CD41"].Value2 = totalPageCount;
                    page.Range["BJ35"].Value2 = firstDocNo;
                }
                else
                {
                    page.Range["BJ43"].Value2 = nextDocNo;
                    page.Range["CE44"].Value2 = globalPage + 1;
                }
            }

            HideNonGeneratedSheets(wb, generatedPageNames);
            if (File.Exists(batch.OutputPath)) File.Delete(batch.OutputPath);
            wb.ExportAsFixedFormat(
                Type: XlTypePdf,
                Filename: batch.OutputPath,
                Quality: XlQualityMinimum,
                IncludeDocProperties: false,
                IgnorePrintAreas: false,
                OpenAfterPublish: false);
            cancellationToken.ThrowIfCancellationRequested();
        }
        finally
        {
            if (wb is not null)
            {
                try { wb.Close(SaveChanges: false); } catch { }
            }
            if (excel is not null)
            {
                try { excel.Quit(); } catch { }
            }
            ReleaseCom(wb);
            ReleaseCom(excel);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private static void JoinPdfParts(string pdf24DocTool, IEnumerable<string> partPaths, string output, CancellationToken cancellationToken)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(output)!);
        if (File.Exists(output)) File.Delete(output);

        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = pdf24DocTool,
                UseShellExecute = false,
                CreateNoWindow = true
            }
        };
        process.StartInfo.ArgumentList.Add("-join");
        process.StartInfo.ArgumentList.Add("-noProgress");
        process.StartInfo.ArgumentList.Add("-outputFile");
        process.StartInfo.ArgumentList.Add(output);
        foreach (string partPath in partPaths.OrderBy(path => path, StringComparer.OrdinalIgnoreCase))
            process.StartInfo.ArgumentList.Add(partPath);

        process.Start();
        try
        {
            while (!process.WaitForExit(250))
                cancellationToken.ThrowIfCancellationRequested();
        }
        catch (OperationCanceledException)
        {
            if (!process.HasExited) process.Kill(entireProcessTree: true);
            throw;
        }

        if (process.ExitCode != 0 || !File.Exists(output))
            throw new InvalidOperationException("PDF24 не смог объединить сформированные части PDF.");
    }

    private static dynamic? GetSheet(dynamic wb, string name)
    {
        try { return wb.Worksheets[name]; }
        catch { return null; }
    }

    private static string Normalize(object? value)
    {
        if (value is null) return "";
        return Convert.ToString(value, CultureInfo.InvariantCulture)?
            .Replace('\r', ' ')
            .Replace('\n', ' ')
            .Trim()
            .ToLowerInvariant() ?? "";
    }

    private static string ToText(object? value)
    {
        if (value is null) return "";
        if (value is double d)
            return d.ToString("0.##########", CultureInfo.CurrentCulture);
        return Convert.ToString(value, CultureInfo.CurrentCulture)?.Trim() ?? "";
    }

    private static void ReleaseCom(object? obj)
    {
        if (obj is null) return;
        try
        {
            if (Marshal.IsComObject(obj)) Marshal.FinalReleaseComObject(obj);
        }
        catch { }
    }

    private static void StopExcel(dynamic excel)
    {
        try
        {
            int hwnd = Convert.ToInt32(excel.Hwnd, CultureInfo.InvariantCulture);
            GetWindowThreadProcessId(new IntPtr(hwnd), out uint processId);
            if (processId != 0)
            {
                using var process = Process.GetProcessById((int)processId);
                if (!process.HasExited)
                    process.Kill(entireProcessTree: true);
                return;
            }
        }
        catch { }

        try { excel.Quit(); } catch { }
    }

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    private sealed class WorkbookCellStyleIndex
    {
        private readonly Dictionary<string, WorksheetStyleMap> _sheets;

        private WorkbookCellStyleIndex(Dictionary<string, WorksheetStyleMap> sheets)
        {
            _sheets = sheets;
        }

        public static WorkbookCellStyleIndex? TryLoad(string workbookPath)
        {
            string extension = Path.GetExtension(workbookPath);
            if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) &&
                !extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase) &&
                !extension.Equals(".xltx", StringComparison.OrdinalIgnoreCase) &&
                !extension.Equals(".xltm", StringComparison.OrdinalIgnoreCase))
                return null;

            try
            {
                using var archive = ZipFile.OpenRead(workbookPath);
                ZipArchiveEntry? workbookEntry = archive.GetEntry("xl/workbook.xml");
                ZipArchiveEntry? relationshipsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels");
                if (workbookEntry is null || relationshipsEntry is null)
                    return null;

                XDocument workbook;
                XDocument relationships;
                using (Stream stream = workbookEntry.Open())
                    workbook = XDocument.Load(stream);
                using (Stream stream = relationshipsEntry.Open())
                    relationships = XDocument.Load(stream);

                XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                XNamespace officeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                XNamespace packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

                var targets = relationships
                    .Root?
                    .Elements(packageRelationships + "Relationship")
                    .Where(element => !string.Equals((string?)element.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase))
                    .Where(element => !string.IsNullOrWhiteSpace((string?)element.Attribute("Id")) &&
                                      !string.IsNullOrWhiteSpace((string?)element.Attribute("Target")))
                    .ToDictionary(
                        element => (string)element.Attribute("Id")!,
                        element => NormalizePartPath("xl/workbook.xml", (string)element.Attribute("Target")!),
                        StringComparer.Ordinal);

                var sheets = new Dictionary<string, WorksheetStyleMap>(StringComparer.OrdinalIgnoreCase);
                foreach (XElement sheet in workbook.Descendants(spreadsheet + "sheet"))
                {
                    string? name = (string?)sheet.Attribute("name");
                    string? relationshipId = (string?)sheet.Attribute(officeRelationships + "id");
                    if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(relationshipId) ||
                        !targets.TryGetValue(relationshipId, out string? worksheetPart))
                        continue;

                    ZipArchiveEntry? worksheetEntry = archive.GetEntry(worksheetPart);
                    if (worksheetEntry is null)
                        continue;

                    using Stream stream = worksheetEntry.Open();
                    XDocument worksheet = XDocument.Load(stream);
                    sheets[name] = WorksheetStyleMap.Create(worksheet, spreadsheet);
                }

                return sheets.Count == 0 ? null : new WorkbookCellStyleIndex(sheets);
            }
            catch
            {
                // Any unusual/legacy workbook falls back to the existing Excel
                // COM formatting path rather than failing PDF generation.
                return null;
            }
        }

        public bool TryGetStyleId(string sheetName, int row, int column, out int styleId)
        {
            if (_sheets.TryGetValue(sheetName, out WorksheetStyleMap? map))
            {
                styleId = map.GetStyleId(row, column);
                return true;
            }

            styleId = 0;
            return false;
        }

        private static string NormalizePartPath(string basePart, string target)
        {
            if (target.StartsWith('/'))
                return target.TrimStart('/');

            var parts = new List<string>();
            string baseDirectory = basePart.Contains('/')
                ? basePart[..(basePart.LastIndexOf('/') + 1)]
                : "";

            foreach (string part in (baseDirectory + target).Replace('\\', '/').Split('/'))
            {
                if (part.Length == 0 || part == ".")
                    continue;
                if (part == "..")
                {
                    if (parts.Count > 0) parts.RemoveAt(parts.Count - 1);
                    continue;
                }
                parts.Add(part);
            }
            return string.Join('/', parts);
        }

        private sealed class WorksheetStyleMap
        {
            private readonly Dictionary<long, int> _cellStyles;
            private readonly Dictionary<int, int> _rowStyles;
            private readonly Dictionary<int, int> _columnStyles;

            private WorksheetStyleMap(
                Dictionary<long, int> cellStyles,
                Dictionary<int, int> rowStyles,
                Dictionary<int, int> columnStyles)
            {
                _cellStyles = cellStyles;
                _rowStyles = rowStyles;
                _columnStyles = columnStyles;
            }

            public static WorksheetStyleMap Create(XDocument worksheet, XNamespace spreadsheet)
            {
                var cellStyles = new Dictionary<long, int>();
                var rowStyles = new Dictionary<int, int>();
                var columnStyles = new Dictionary<int, int>();

                XElement? cols = worksheet.Root?.Element(spreadsheet + "cols");
                if (cols is not null)
                {
                    foreach (XElement col in cols.Elements(spreadsheet + "col"))
                    {
                        if (!TryReadInt(col.Attribute("style"), out int styleId) ||
                            !TryReadInt(col.Attribute("min"), out int min) ||
                            !TryReadInt(col.Attribute("max"), out int max))
                            continue;

                        min = Math.Max(1, min);
                        max = Math.Min(16384, max);
                        for (int column = min; column <= max; column++)
                            columnStyles[column] = styleId;
                    }
                }

                XElement? sheetData = worksheet.Root?.Element(spreadsheet + "sheetData");
                if (sheetData is not null)
                {
                    foreach (XElement rowElement in sheetData.Elements(spreadsheet + "row"))
                    {
                        int rowNumber = TryReadInt(rowElement.Attribute("r"), out int parsedRow) ? parsedRow : 0;
                        if (rowNumber > 0 && TryReadInt(rowElement.Attribute("s"), out int rowStyle))
                            rowStyles[rowNumber] = rowStyle;

                        foreach (XElement cell in rowElement.Elements(spreadsheet + "c"))
                        {
                            if (!TryReadInt(cell.Attribute("s"), out int styleId))
                                continue;

                            string? reference = (string?)cell.Attribute("r");
                            if (!TryParseCellReference(reference, out int cellRow, out int cellColumn))
                                continue;

                            cellStyles[MakeCellKey(cellRow, cellColumn)] = styleId;
                        }
                    }
                }

                return new WorksheetStyleMap(cellStyles, rowStyles, columnStyles);
            }

            public int GetStyleId(int row, int column)
            {
                if (_cellStyles.TryGetValue(MakeCellKey(row, column), out int styleId))
                    return styleId;
                if (_rowStyles.TryGetValue(row, out styleId))
                    return styleId;
                if (_columnStyles.TryGetValue(column, out styleId))
                    return styleId;
                return 0;
            }

            private static long MakeCellKey(int row, int column) => ((long)row << 20) | (uint)column;

            private static bool TryReadInt(XAttribute? attribute, out int value) =>
                int.TryParse((string?)attribute, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);

            private static bool TryParseCellReference(string? reference, out int row, out int column)
            {
                row = 0;
                column = 0;
                if (string.IsNullOrWhiteSpace(reference))
                    return false;

                int index = 0;
                while (index < reference.Length && char.IsLetter(reference[index]))
                {
                    column = checked(column * 26 + (char.ToUpperInvariant(reference[index]) - 'A' + 1));
                    index++;
                }

                if (column <= 0 || index >= reference.Length)
                    return false;

                return int.TryParse(reference.AsSpan(index), NumberStyles.Integer, CultureInfo.InvariantCulture, out row) && row > 0;
            }
        }
    }

    private sealed record HeaderMap(int HeaderRow, int[] Columns);
    private sealed record SpecificationSection(string Name, List<SpecRow> Rows);
    private sealed record PdfBatch(int Index, int FirstPage, List<List<SpecRow?>> Pages, string OutputPath);

    private sealed class TemplateLayout
    {
        private static readonly string[] DataColumns = ["G", "K", "AK", "AW", "BD", "BM", "BQ", "BU", "BZ"];
        // Excel centres this print area inside the PDF printer's non-printable edge.
        // These are PDF points measured from the background exported by the same Excel call.
        // They keep the overlay aligned with the already rendered Excel grid.
        private const double ExportInsetX = 14.768;
        private const double ExportInsetY = 12.095;
        private readonly TextSlot[,] _rows = new TextSlot[RowsPerPage, DataColumns.Length];
        private readonly TextSlot? _documentNumber;
        private readonly TextSlot? _pageNumber;
        private readonly TextSlot? _totalPages;
        private readonly TextSlot? _topPageNumber;
        private readonly int? _topPageNumberBase;

        private TemplateLayout(
            TextSlot? documentNumber,
            TextSlot? pageNumber,
            TextSlot? totalPages,
            TextSlot? topPageNumber,
            int? topPageNumberBase)
        {
            _documentNumber = documentNumber;
            _pageNumber = pageNumber;
            _totalPages = totalPages;
            _topPageNumber = topPageNumber;
            _topPageNumberBase = topPageNumberBase;
        }

        public static TemplateLayout Create(
            dynamic sheet,
            string backgroundPdf,
            bool isFirstPage,
            int? topPageNumberBase)
        {
            using var background = PdfReader.Open(backgroundPdf, PdfDocumentOpenMode.Import);
            var pdfPage = background.Pages[0];
            dynamic printArea = sheet.Range["C3:CH46"];
            double printWidth = Convert.ToDouble(printArea.Width, CultureInfo.InvariantCulture);
            double printHeight = Convert.ToDouble(printArea.Height, CultureInfo.InvariantCulture);
            double scaleX = (pdfPage.Width.Point - ExportInsetX * 2) / printWidth;
            double scaleY = (pdfPage.Height.Point - ExportInsetY * 2) / printHeight;
            var layout = new TemplateLayout(
                CaptureSlot(sheet, isFirstPage ? "BJ35" : "BJ43", printArea, scaleX, scaleY, ExportInsetX, ExportInsetY),
                CaptureSlot(sheet, isFirstPage ? "CA41" : "CE44", printArea, scaleX, scaleY, ExportInsetX, ExportInsetY),
                isFirstPage ? CaptureSlot(sheet, "CD41", printArea, scaleX, scaleY, ExportInsetX, ExportInsetY) : null,
                topPageNumberBase.HasValue
                    ? CaptureSlot(sheet, "CF4", printArea, scaleX, scaleY, ExportInsetX, ExportInsetY)
                    : null,
                topPageNumberBase);

            for (int row = 0; row < RowsPerPage; row++)
            {
                for (int field = 0; field < DataColumns.Length; field++)
                    layout._rows[row, field] = CaptureSlot(
                        sheet,
                        $"{DataColumns[field]}{DataStartRow + row}",
                        printArea,
                        scaleX,
                        scaleY,
                        ExportInsetX,
                        ExportInsetY,
                        field == 1 ? 3.0 : 0.0)!;
            }

            return layout;
        }

        public List<SpecificationSection> ExpandSectionsForAutoWrap(
            IReadOnlyList<SpecificationSection> sections,
            XGraphics graphics,
            bool preserveFormatting,
            bool replaceFontFamily,
            string fontName,
            bool italic)
        {
            var result = new List<SpecificationSection>(sections.Count);
            foreach (var section in sections)
            {
                var expandedRows = new List<SpecRow>();
                foreach (var row in section.Rows)
                {
                    CellFont font = row.Fonts[1];
                    IReadOnlyList<string> lines = _rows[0, 1].GetWrappedLines(
                        graphics,
                        row.Fields[1],
                        fontNameOverride: replaceFontFamily ? fontName : null,
                        boldOverride: preserveFormatting ? font.Bold : null,
                        italicOverride: replaceFontFamily && italic ? true : preserveFormatting ? font.Italic : replaceFontFamily ? false : null,
                        underlineOverride: preserveFormatting ? font.Underline : null);

                    if (lines.Count <= 1)
                    {
                        expandedRows.Add(row);
                        continue;
                    }

                    expandedRows.Add(row.WithNameLine(lines[0], lines.Count));
                    for (int lineIndex = 1; lineIndex < lines.Count; lineIndex++)
                        expandedRows.Add(row.WithNameLine(lines[lineIndex], continuation: true));
                }
                result.Add(new SpecificationSection(section.Name, expandedRows));
            }
            return result;
        }

        public void DrawRows(
            XGraphics graphics,
            IReadOnlyList<SpecRow?> page,
            bool preserveFormatting,
            bool replaceFontFamily,
            string fontName,
            bool italic)
        {
            for (int row = 0; row < RowsPerPage; row++)
            {
                if (page[row] is not { } record) continue;
                if (record.IsSheetTitle)
                {
                    _rows[row, 1].Draw(graphics, record.Fields[1], forceCenter: true, forceBold: true, underline: true,
                        fontNameOverride: replaceFontFamily ? fontName : null,
                        italicOverride: replaceFontFamily && italic ? true : null);
                    continue;
                }
                for (int field = 0; field < DataColumns.Length; field++)
                {
                    CellFont font = record.Fonts[field];
                    if (preserveFormatting && font.Fill is XColor fill)
                        _rows[row, field].Fill(graphics, fill);
                    _rows[row, field].Draw(
                        graphics,
                        record.Fields[field],
                        fontNameOverride: replaceFontFamily ? fontName : null,
                        boldOverride: preserveFormatting ? font.Bold : null,
                        italicOverride: replaceFontFamily && italic ? true : preserveFormatting ? font.Italic : replaceFontFamily ? false : null,
                        underlineOverride: preserveFormatting ? font.Underline : null);
                }
            }
        }

        public void DrawStamp(XGraphics graphics, int pageNumber, int totalPages, int topPageIncrement)
        {
            _pageNumber?.Draw(graphics, pageNumber.ToString(CultureInfo.InvariantCulture));
            _totalPages?.Draw(graphics, totalPages.ToString(CultureInfo.InvariantCulture));
            if (_topPageNumberBase is int baseNumber)
                _topPageNumber?.Draw(graphics, (baseNumber + topPageIncrement).ToString(CultureInfo.InvariantCulture));
        }

        private static TextSlot? CaptureSlot(
            dynamic sheet,
            string address,
            dynamic printArea,
            double scaleX,
            double scaleY,
            double offsetX,
            double offsetY,
            double leftInset = 0)
        {
            try
            {
                dynamic cell = sheet.Range[address];
                dynamic area = cell.MergeCells ? cell.MergeArea : cell;
                double x = offsetX + (Convert.ToDouble(area.Left, CultureInfo.InvariantCulture) - Convert.ToDouble(printArea.Left, CultureInfo.InvariantCulture)) * scaleX;
                double y = offsetY + (Convert.ToDouble(area.Top, CultureInfo.InvariantCulture) - Convert.ToDouble(printArea.Top, CultureInfo.InvariantCulture)) * scaleY;
                double width = Convert.ToDouble(area.Width, CultureInfo.InvariantCulture) * scaleX;
                double height = Convert.ToDouble(area.Height, CultureInfo.InvariantCulture) * scaleY;
                dynamic font = cell.Font;
                string fontName = "ISOCPEUR";
                double fontSize = Convert.ToDouble(font.Size, CultureInfo.InvariantCulture) * scaleY;
                bool bold = false;
                bool italic = true;
                int horizontalAlignment = Convert.ToInt32(cell.HorizontalAlignment, CultureInfo.InvariantCulture);
                int verticalAlignment = Convert.ToInt32(cell.VerticalAlignment, CultureInfo.InvariantCulture);
                bool wrapText = Convert.ToBoolean(cell.WrapText, CultureInfo.InvariantCulture);
                int indentLevel = Convert.ToInt32(cell.IndentLevel, CultureInfo.InvariantCulture);
                return new TextSlot(x, y, width, height, fontName, fontSize, bold, italic,
                    horizontalAlignment, verticalAlignment, wrapText, indentLevel, leftInset);
            }
            catch
            {
                return null;
            }
        }
    }

    private sealed record TextSlot(
        double X,
        double Y,
        double Width,
        double Height,
        string FontName,
        double FontSize,
        bool Bold,
        bool Italic,
        int HorizontalAlignment,
        int VerticalAlignment,
        bool WrapText,
        int IndentLevel,
        double LeftInset)
    {
        public TextSlot WithBounds(double y, double height) => this with { Y = y, Height = height };

        public void Fill(XGraphics graphics, XColor color)
        {
            if (Width <= 0 || Height <= 0) return;
            // The template PDF already contains the cell borders. Keep a narrow
            // margin so the overlay fill never paints over those grid lines.
            // Reduce the fill height by 10%, keeping the template grid clear.
            const double baseBorderInset = 2.0;
            const double fillScale = 0.9486832981; // sqrt(0.90)
            double baseWidth = Math.Max(0, Width - baseBorderInset * 2);
            double baseHeight = Math.Max(0, Height - baseBorderInset * 2);
            // The wide horizontal gap made the coloured source cell look
            // detached from its text. Keep the narrow border margin sideways;
            // only the vertical direction is reduced by 10%.
            double fillWidth = baseWidth;
            double fillHeight = baseHeight * fillScale;
            graphics.DrawRectangle(new XSolidBrush(color), new XRect(
                X + baseBorderInset,
                Y + baseBorderInset + (baseHeight - fillHeight) / 2,
                fillWidth,
                fillHeight));
        }

        public double RequiredHeight(XGraphics graphics, string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0;
            var font = CreateFont();
            return GetLines(graphics, font, text).Count * font.GetHeight() + 2.5;
        }

        public IReadOnlyList<string> GetWrappedLines(
            XGraphics graphics,
            string text,
            string? fontNameOverride = null,
            bool? boldOverride = null,
            bool? italicOverride = null,
            bool? underlineOverride = null)
        {
            if (string.IsNullOrWhiteSpace(text)) return [""];
            var font = CreateFont(fontNameOverride: fontNameOverride, boldOverride: boldOverride,
                italicOverride: italicOverride, underlineOverride: underlineOverride);
            return GetLines(graphics, font, text, forceWrap: true);
        }

        public void Draw(
            XGraphics graphics,
            string text,
            bool forceCenter = false,
            bool forceBold = false,
            bool underline = false,
            string? fontNameOverride = null,
            bool? boldOverride = null,
            bool? italicOverride = null,
            bool? underlineOverride = null)
        {
            if (string.IsNullOrWhiteSpace(text) || Width <= 0 || Height <= 0) return;
            var font = CreateFont(forceBold, underline, fontNameOverride, boldOverride, italicOverride, underlineOverride);
            var lines = GetLines(graphics, font, text);
            const double cellPadding = 1.25;
            double indent = IndentLevel * font.GetHeight() * 0.4;
            double textX = X + cellPadding + indent + LeftInset;
            double textWidth = Math.Max(0, Width - cellPadding * 2 - indent - LeftInset);
            double lineHeight = font.GetHeight();
            int visibleLines = Math.Min(lines.Count, Math.Max(1, (int)Math.Floor(Height / lineHeight)));
            double textHeight = lineHeight * visibleLines;
            // Excel constants: xlVAlignCenter = -4108, xlVAlignBottom = -4107.
            double top = VerticalAlignment == -4108 ? Y + (Height - textHeight) / 2 :
                VerticalAlignment == -4107 ? Y + Height - textHeight : Y;
            var format = forceCenter ? XStringFormats.TopCenter : HorizontalAlignment switch
            {
                -4108 => XStringFormats.TopCenter,
                -4152 => XStringFormats.TopRight,
                _ => XStringFormats.TopLeft
            };

            var state = graphics.Save();
            graphics.IntersectClip(new XRect(X, Y, Width, Height));
            try
            {
                for (int index = 0; index < visibleLines; index++)
                    graphics.DrawString(lines[index], font, XBrushes.Black,
                        new XRect(textX, top + index * lineHeight, textWidth, lineHeight), format);
            }
            finally
            {
                graphics.Restore(state);
            }
        }

        private XFont CreateFont(
            bool forceBold = false,
            bool underline = false,
            string? fontNameOverride = null,
            bool? boldOverride = null,
            bool? italicOverride = null,
            bool? underlineOverride = null)
        {
            bool bold = forceBold || (boldOverride ?? Bold);
            bool italic = italicOverride ?? Italic;
            bool isUnderlined = underline || (underlineOverride ?? false);
            var style = bold && italic ? XFontStyleEx.BoldItalic :
                bold ? XFontStyleEx.Bold : italic ? XFontStyleEx.Italic : XFontStyleEx.Regular;
            if (isUnderlined) style |= XFontStyleEx.Underline;
            string fontName = string.IsNullOrWhiteSpace(fontNameOverride) ? FontName : fontNameOverride;
            return new XFont(fontName, Math.Max(1, FontSize), style);
        }

        private List<string> GetLines(XGraphics graphics, XFont font, string text, bool forceWrap = false)
        {
            const double cellPadding = 1.25;
            double indent = IndentLevel * font.GetHeight() * 0.4;
            double textWidth = Math.Max(0, Width - cellPadding * 2 - indent - LeftInset);
            string normalized = KeepShortWordsWithNext(text);
            return WrapText || forceWrap ? Wrap(graphics, font, normalized, textWidth) : [normalized.Replace("\r", "").Replace("\n", " ")];
        }

        // Excel keeps a single-letter word with the following value. It avoids
        // visual orphans such as "A," alone at the end of a description line.
        private static string KeepShortWordsWithNext(string text)
        {
            // Keep a classification letter with both its description and value:
            // "герметичности A, Ø125" is one visual group.
            string grouped = System.Text.RegularExpressions.Regex.Replace(
                text,
                @"(\S+) ([A-Za-zА-Яа-яЁё])([,.]?) (?=\S)",
                "$1\u00A0$2$3\u00A0");
            return System.Text.RegularExpressions.Regex.Replace(
                grouped,
                @"(?<!\S)([A-Za-zА-Яа-яЁё])([,.]?) (?=\S)",
                "$1$2\u00A0");
        }

        private static List<string> Wrap(XGraphics graphics, XFont font, string text, double width)
        {
            var lines = new List<string>();
            foreach (string sourceLine in text.Replace("\r", "").Split('\n'))
            {
                string current = "";
                foreach (string word in sourceLine.Split(' ', StringSplitOptions.RemoveEmptyEntries))
                {
                    string candidate = current.Length == 0 ? word : current + " " + word;
                    if (current.Length > 0 && graphics.MeasureString(candidate, font).Width > width)
                    {
                        lines.Add(current);
                        current = word;
                    }
                    else if (current.Length == 0 && graphics.MeasureString(word, font).Width > width)
                    {
                        foreach (string part in SplitLongWord(graphics, font, word, width))
                            lines.Add(part);
                    }
                    else current = candidate;
                }
                if (current.Length > 0 || sourceLine.Length == 0)
                    lines.Add(current);
            }
            return lines.Count == 0 ? [""] : lines;
        }

        private static IEnumerable<string> SplitLongWord(XGraphics graphics, XFont font, string word, double width)
        {
            string part = "";
            foreach (char character in word)
            {
                string candidate = part + character;
                if (part.Length > 0 && graphics.MeasureString(candidate, font).Width > width)
                {
                    yield return part;
                    part = character.ToString();
                }
                else part = candidate;
            }
            if (part.Length > 0) yield return part;
        }
    }

    private sealed class SpecRow
    {
        public string[] Fields { get; }
        public CellFont[] Fonts { get; }
        public bool IsSheetTitle { get; }
        public bool IsContinuation { get; }
        public int RowSpan { get; }

        public SpecRow(
            string[] fields,
            CellFont[]? fonts = null,
            bool isSheetTitle = false,
            bool isContinuation = false,
            int rowSpan = 1)
        {
            Fields = fields;
            Fonts = fonts is { Length: 9 } ? fonts : new CellFont[9];
            IsSheetTitle = isSheetTitle;
            IsContinuation = isContinuation;
            RowSpan = Math.Max(1, rowSpan);
        }

        public static SpecRow SheetTitle(string name)
        {
            var fields = new string[9];
            fields[1] = name;
            return new SpecRow(fields, isSheetTitle: true);
        }

        public SpecRow WithNameLine(string nameLine, int rowSpan = 1, bool continuation = false)
        {
            var fields = continuation ? new string[9] : (string[])Fields.Clone();
            fields[1] = nameLine;
            return new SpecRow(fields, Fonts, IsSheetTitle, continuation, rowSpan);
        }

        public bool IsHeading
        {
            get
            {
                if (IsContinuation) return false;
                if (string.IsNullOrWhiteSpace(Fields[1])) return false;
                for (int i = 0; i < Fields.Length; i++)
                {
                    if (i == 1) continue;
                    if (!string.IsNullOrWhiteSpace(Fields[i])) return false;
                }
                return true;
            }
        }
    }

    private readonly record struct CellFont(bool Bold, bool Italic, bool Underline, XColor? Fill);

    private readonly record struct FontReadOptions(bool PreserveFormatting)
    {
        public bool Any => PreserveFormatting;
        public bool Underline => PreserveFormatting;
        public bool Weight => PreserveFormatting;
        public bool Italic => PreserveFormatting;
        public bool Fill => PreserveFormatting;
    }

    private readonly record struct FontColumnDefaults(bool? Bold, bool? Italic, bool? Underline, bool HasUniformFill, XColor? Fill);

    private sealed record GeneratorSettings(
        bool PreserveFormatting = false,
        bool ReplaceFontFamily = true,
        string FontName = "ISOCPEUR",
        bool Italic = true,
        bool AutoWrap = false)
    {
        // Kept only to migrate settings.json saved by versions before the new window.
        public bool InheritFontFamily { get; init; }
        public bool InheritFontStyle { get; init; }
        public bool InheritFontWeight { get; init; }
        public bool InheritFontItalic { get; init; }
    }

    private sealed class ExternalWindow(IntPtr handle) : IWin32Window
    {
        public IntPtr Handle => handle;
    }

    private sealed class GenerationOptionsWindow : Form
    {
        private readonly CheckBox _preserveFormatting = new()
        {
            AutoSize = true,
            Left = 20,
            Top = 20,
            Text = "Сохранить исходное форматирование"
        };
        private readonly CheckBox _replaceFontFamily = new()
        {
            AutoSize = true,
            Left = 20,
            Top = 48,
            Text = "Заменить тип шрифта"
        };
        private readonly CheckBox _autoWrap = new()
        {
            AutoSize = true,
            Left = 20,
            Top = 76,
            Text = "Включить автоперенос"
        };
        private readonly CheckBox _italic = new()
        {
            AutoSize = true,
            Left = 92,
            Top = 166,
            Text = "Курсив"
        };
        private readonly ComboBox _fontFamily = new()
        {
            AutoCompleteMode = AutoCompleteMode.SuggestAppend,
            AutoCompleteSource = AutoCompleteSource.ListItems,
            DropDownStyle = ComboBoxStyle.DropDown,
            Left = 155,
            Top = 136,
            Width = 205
        };

        public bool PreserveFormatting => _preserveFormatting.Checked;
        public bool ReplaceFontFamily => _replaceFontFamily.Checked;
        public bool AutoWrap => _autoWrap.Checked;
        public string FontName => string.IsNullOrWhiteSpace(_fontFamily.Text) ? "ISOCPEUR" : _fontFamily.Text.Trim();
        public bool Italic => _italic.Checked;

        public GenerationOptionsWindow(bool preserveFormatting, bool replaceFontFamily, string fontName, bool italic, bool autoWrap)
        {
            Text = "Создание PDF";
            ClientSize = new Size(390, 248);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;

            var fontLabel = new Label
            {
                AutoSize = true,
                Left = 20,
                Top = 140,
                Text = "Заменить шрифт на:"
            };
            foreach (FontFamily family in new InstalledFontCollection().Families.OrderBy(family => family.Name, StringComparer.CurrentCultureIgnoreCase))
                _fontFamily.Items.Add(family.Name);

            _preserveFormatting.Checked = preserveFormatting;
            _replaceFontFamily.Checked = replaceFontFamily;
            _italic.Checked = italic;
            _autoWrap.Checked = autoWrap;
            _fontFamily.Text = fontName;
            _replaceFontFamily.CheckedChanged += (_, _) => UpdateFontControls();
            UpdateFontControls();
            var createButton = new Button
            {
                DialogResult = DialogResult.OK,
                Left = 180,
                Top = 206,
                Width = 120,
                Height = 30,
                Text = "Создать PDF"
            };
            var cancelButton = new Button
            {
                DialogResult = DialogResult.Cancel,
                Left = 310,
                Top = 206,
                Width = 70,
                Height = 30,
                Text = "Отмена"
            };

            Controls.AddRange([_preserveFormatting, _replaceFontFamily, _autoWrap, _italic, fontLabel, _fontFamily, createButton, cancelButton]);
            AcceptButton = createButton;
            CancelButton = cancelButton;
        }

        private void UpdateFontControls()
        {
            _fontFamily.Enabled = _replaceFontFamily.Checked;
            _italic.Enabled = _replaceFontFamily.Checked;
        }
    }

    private sealed class ProgressWindow : Form
    {
        private readonly Label _status = new()
        {
            AutoSize = false,
            Left = 20,
            Top = 18,
            Width = 520,
            Height = 52,
            Text = "Подготовка…"
        };

        private readonly ProgressBar _progress = new()
        {
            Left = 20,
            Top = 78,
            Width = 520,
            Height = 24,
            Style = ProgressBarStyle.Marquee,
            MarqueeAnimationSpeed = 30
        };

        private readonly Button _close = new()
        {
            Left = 440,
            Top = 118,
            Width = 100,
            Height = 30,
            Text = "Отмена"
        };

        private readonly Action _cancel;
        private bool _finished;

        public int ExitCode { get; private set; } = 1;

        public ProgressWindow(Action cancel)
        {
            _cancel = cancel;
            Text = "Формирование PDF-спецификации";
            ClientSize = new System.Drawing.Size(560, 165);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ControlBox = false;
            TopMost = true;
            StartPosition = FormStartPosition.CenterScreen;

            _close.Click += (_, _) =>
            {
                if (_finished)
                {
                    Close();
                    return;
                }

                _close.Enabled = false;
                _status.Text = "Останавливаю Excel и отменяю формирование…";
                _cancel();
                Close();
            };
            Controls.AddRange([_status, _progress, _close]);
        }

        public void Report(string status, int completed, int total)
        {
            if (IsDisposed) return;
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => Report(status, completed, total)));
                return;
            }

            _status.Text = status;
            if (total <= 0)
            {
                _progress.Style = ProgressBarStyle.Marquee;
                return;
            }

            _progress.Style = ProgressBarStyle.Continuous;
            _progress.Minimum = 0;
            _progress.Maximum = Math.Max(1, total);
            _progress.Value = Math.Clamp(completed, 0, _progress.Maximum);
        }

        public void FinishSuccess(string output)
        {
            if (IsDisposed)
                return;

            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => FinishSuccess(output)));
                return;
            }

            Finish($"Готово. PDF сохранён:\n{output}", success: true);
            Close();
        }

        public void FinishError(string error)
        {
            Finish($"Не удалось сформировать PDF:\n{error}", success: false);
        }

        public void FinishCancelled()
        {
            Finish("Формирование отменено.", success: false);
        }

        private void Finish(string status, bool success)
        {
            if (IsDisposed) return;
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => Finish(status, success)));
                return;
            }

            ExitCode = success ? 0 : 1;
            _finished = true;
            _status.Text = status;
            _progress.Style = ProgressBarStyle.Continuous;
            _progress.Maximum = 100;
            _progress.Value = success ? 100 : 0;
            _close.Text = "Закрыть";
            _close.Enabled = true;
            ControlBox = true;
        }
    }
}
