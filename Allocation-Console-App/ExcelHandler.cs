using Allocation_Console_App.Entities;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Text.RegularExpressions;
using File = System.IO.File;
using Group = Allocation_Console_App.Entities.Group;

namespace Allocation_Console_App
{
    public class ExcelHandler
    {
        public AllocationContext ReadExcelFileNPOI(string filePath, AllocationContext context)
        {
            if (File.Exists(filePath))
            {
                IWorkbook workbook;
                using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    if (filePath.EndsWith(".xlsx"))
                    {
                        workbook = new XSSFWorkbook(file);
                    }
                    else if (filePath.EndsWith(".xls"))
                    {
                        workbook = new HSSFWorkbook(file);
                    }
                    else
                    {
                        Console.WriteLine("The specified file is not an Excel file.");
                        return null;
                    }
                }
                var sourceFolder = Path.GetDirectoryName(filePath);
                if (sourceFolder != null)
                {
                    context.SourceDirectory = sourceFolder;
                }
                List<Group> groups = new();
                ISheet sheet = workbook.GetSheetAt(0);
                for (int row = 0; row <= sheet.LastRowNum; row++)
                {
                    IRow excelRow = sheet.GetRow(row);
                    if (excelRow != null)
                    {
                        Group group = new Group();
                        bool skip = false; // Falls in einer Reihe wichtige Informationen nicht vorhanden sind, soll diese übersprungen werden.
                        for (int col = 0; col < excelRow.LastCellNum && !skip; col++)
                        {
                            ICell cell = excelRow.GetCell(col);
                            if (cell != null)
                            {
                                string? value = cell.ToString();
                                string numStr = Regex.Match(value, @"\d+").Value;
                                if (value != null)
                                {
                                    switch (col)
                                    {
                                        case 0: // Besteller ID
                                            if (!string.IsNullOrEmpty(numStr))
                                            {
                                                group.GroupId = int.Parse(numStr);
                                            }
                                            break;
                                        case 1: // Besteller Name
                                            group.Name = value;
                                            break;
                                        case 2: // Freitag
                                            if (!string.IsNullOrEmpty(numStr))
                                            {
                                                group.Day = "Freitag";
                                                group.Size = int.Parse(numStr);
                                            }
                                            break;
                                        case 3: // Samstag
                                            if (!string.IsNullOrEmpty(numStr))
                                            {
                                                group.Day = "Samstag";
                                                group.Size = int.Parse(numStr);
                                            }
                                            break;
                                        case 4: // Egal, Freitag oder Samstag
                                            break;
                                        case 5: // Freikarten
                                            break;
                                        case 6: // Tisch-Nummer, wird aber zugewiesen
                                            break;
                                        case 7: // Gewichtung/Priorität
                                            if (!string.IsNullOrEmpty(numStr))
                                            {
                                                group.Score = int.Parse(numStr);
                                            }
                                            break;
                                    }
                                }
                            }
                            else
                            {
                                if (col == 0)
                                {
                                    skip = true;
                                }
                            }
                        }
                        if (!skip)
                        {
                            groups.Add(group);
                        }
                    }
                }
                workbook.Close();
                List<Group> cleanedGroupsList = new List<Group>();
                foreach (var group in groups)
                {
                    if (group.GroupId != 0)
                    {
                        cleanedGroupsList.Add(group);
                    }
                }
                //Add colorIndex for formating cell later.
                int colorIndex = 1;
                foreach (var group in cleanedGroupsList)
                {
                    group.ColorIndex = colorIndex++;
                    if (colorIndex >= 9)
                    {
                        colorIndex = 1;
                    }
                }
                context.RawGroupData = cleanedGroupsList;
                return context;
            }
            else
            {
                Console.WriteLine("Der angegebene Pfad exsistiert nicht!");
            }
            return context;
        }

        public string WriteExcelFileNPOI(AllocationContext context, bool isTest = false)
        {
            IWorkbook workbook = new XSSFWorkbook();
            if (isTest)
            {
                if (context.TheaterFriday == null)
                {
                    return string.Empty;
                }
            }
            else
            {
                if (context.TheaterFriday == null || context.TheaterSaturday == null)
                {
                    return string.Empty;
                }
            }
            // Erzeuge Mappe für Freitag
            ISheet sheet = workbook.CreateSheet("AllocationErgebnisFriday");
            for (int i = 0; i < (context.TheaterFriday[0].Seats.Count / 2) + 1; i++)
            {
                sheet.CreateRow(i);
            }
            int colCount = 0;
            foreach (var table in context.TheaterFriday)
            {
                int rowCount = 0;
                for (int i = 0; i < table.Seats.Count; i++)
                {
                    if (i % 2 == 0)
                    {
                        table.Seats[i].col = colCount;
                        table.Seats[i].row = rowCount;

                        rowCount++;
                    }
                }
                colCount++;
                rowCount = 0;
                for (int i = 0; i < table.Seats.Count; i++)
                {
                    if (i % 2 == 1)
                    {
                        table.Seats[i].col = colCount;
                        table.Seats[i].row = rowCount;

                        rowCount++;
                    }
                }
                colCount += 2;
            }
            // Setze die Values.
            foreach (var table in context.TheaterFriday)
            {
                foreach (var seat in table.Seats)
                {
                    ICell cell = sheet.GetRow(seat.row).CreateCell(seat.col);
                    ICellStyle style = workbook.CreateCellStyle();
                    style.FillPattern = FillPattern.SolidForeground;
                    style.VerticalAlignment = VerticalAlignment.Center;
                    style.Alignment = HorizontalAlignment.Center;
                    if (seat.Owner != null)
                    {
                        cell.SetCellValue($"{seat.SeatNumber} Owner: {seat.Owner.GroupId}");
                        style.FillForegroundColor = GetNextColor(seat.Owner.ColorIndex);
                    }
                    else
                    {
                        cell.SetCellValue($"{seat.SeatNumber} Owner: FREE");
                        style.FillForegroundColor = HSSFColor.Grey25Percent.Index;
                    }
                    cell.CellStyle = style;
                }
            }
            // Passe die Höhe und Breite der Reihen und Spalten an. Für bessere Lesbarkeit.
            for (int row = 0; row <= sheet.LastRowNum; row++)
            {
                IRow excelRow = sheet.GetRow(row);
                if (excelRow != null)
                {
                    excelRow.Height = 50 * 20;
                    for (int col = 0; col < excelRow.LastCellNum; col++)
                    {
                        sheet.SetColumnWidth(col, 256 * 20);
                    }
                }
            }

            //Samstags wird nicht erzeugt in TestFile
            if (!isTest)
            {
                // Erzeuge Mappe für Samstag
                sheet = workbook.CreateSheet("AllocationErgebnisSaturday");
                for (int i = 0; i < (context.TheaterSaturday[0].Seats.Count / 2) + 1; i++)
                {
                    sheet.CreateRow(i);
                }
                colCount = 0;
                foreach (var table in context.TheaterSaturday)
                {
                    int rowCount = 0;
                    for (int i = 0; i < table.Seats.Count; i++)
                    {
                        if (i % 2 == 0)
                        {
                            table.Seats[i].col = colCount;
                            table.Seats[i].row = rowCount;
                            rowCount++;
                        }
                        //Console.WriteLine($"Table {table.TableNumber}: 1 : SeatNumber: {i} of {table.Seats.Count}");
                    }
                    colCount++;
                    rowCount = 0;
                    for (int i = 0; i < table.Seats.Count; i++)
                    {
                        if (i % 2 == 1)
                        {
                            table.Seats[i].col = colCount;
                            table.Seats[i].row = rowCount;

                            rowCount++;
                        }
                        //Console.WriteLine($"Table {table.TableNumber}: 2 : SeatNumber: {i} of {table.Seats.Count}");
                    }
                    colCount += 2;
                }

                // Setze die Values.
                foreach (var table in context.TheaterSaturday)
                {
                    foreach (var seat in table.Seats)
                    {
                        ICell cell = sheet.GetRow(seat.row).CreateCell(seat.col);

                        ICellStyle style = workbook.CreateCellStyle();
                        style.FillPattern = FillPattern.SolidForeground;
                        style.VerticalAlignment = VerticalAlignment.Center;
                        style.Alignment = HorizontalAlignment.Center;
                        if (seat.Owner != null)
                        {
                            cell.SetCellValue($"{seat.SeatNumber} Owner: {seat.Owner.GroupId}");
                            style.FillForegroundColor = GetNextColor(seat.Owner.ColorIndex);
                            IFont font = workbook.CreateFont();
                            font.Boldweight = (short)FontBoldWeight.Bold;

                            style.SetFont(font);
                        }
                        else
                        {
                            cell.SetCellValue($"{seat.SeatNumber} Owner: FREE");
                            style.FillForegroundColor = HSSFColor.Grey25Percent.Index;
                        }
                        cell.CellStyle = style;
                    }
                }

                // Passe die Höhe und Breite der Reihen und Spalten an. Für bessere Lesbarkeit.
                for (int row = 0; row <= sheet.LastRowNum; row++)
                {
                    IRow excelRow = sheet.GetRow(row);
                    if (excelRow != null)
                    {
                        excelRow.Height = 50 * 20;
                        for (int col = 0; col < excelRow.LastCellNum; col++)
                        {
                            sheet.SetColumnWidth(col, 256 * 20);
                        }
                    }
                }
            }

            //Unterscheide zwischen Test-Datei und richtiger Eingabe-Datei
            string fileName;
            if (isTest)
            {
                fileName = "Test_AllocationResult.xlsx";
            }
            else
            {
                fileName = "AllocationResult.xlsx";
            }
            using (FileStream fs = new FileStream(Path.Combine(context.SourceDirectory, fileName), FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
            workbook.Close();
            return Path.Combine(context.SourceDirectory, fileName);
        }

        public string WriteRawDataToFileNPOI(String path, AllocationContext context)
        {
            IWorkbook workbook;
            if (File.Exists(path) || context.RawGroupData == null)
            {
                using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    if (path.EndsWith(".xlsx"))
                    {
                        workbook = new XSSFWorkbook(file);
                    }
                    else if (path.EndsWith(".xls"))
                    {
                        workbook = new HSSFWorkbook(file);
                    }
                    else
                    {
                        Console.WriteLine("The specified file is not an Excel file.");
                        return string.Empty;
                    }
                }
            }
            else
            {
                return string.Empty;
            }
            // Erzeuge Mappe für Freitag
            ISheet sheet = workbook.CreateSheet("Grunddaten");
            sheet.CreateRow(0);
            for (int i = 1; i < context.RawGroupData.Count; i++)
            {
                sheet.CreateRow(i);
            }
            IRow row = sheet.GetRow(0);
            ICell cell;
            cell = row.CreateCell(0);
            cell.SetCellValue($"ID");
            cell = row.CreateCell(1);
            cell.SetCellValue($"Name");
            cell = row.CreateCell(2);
            cell.SetCellValue($"Tag");
            cell = row.CreateCell(3);
            cell.SetCellValue($"Anzahl");
            cell = row.CreateCell(4);
            cell.SetCellValue($"Priorität");
            for (int i = 1; i < context.RawGroupData.Count; i++)
            {
                var group = context.RawGroupData[i];
                row = sheet.GetRow(i);
                cell = row.CreateCell(0);
                cell.SetCellValue($"{group.GroupId}");
                cell = row.CreateCell(1);
                cell.SetCellValue($"{group.Name}");
                cell = row.CreateCell(2);
                cell.SetCellValue($"{group.Day}");
                cell = row.CreateCell(3);
                cell.SetCellValue($"{group.Size}");
                cell = row.CreateCell(4);
                cell.SetCellValue($"{group.Score}");
            }
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
            workbook.Close();
            return path;
        }

        public string WriteResultsToFileNPOI(String path, AllocationContext context, bool isFriday = true)
        {
            IWorkbook workbook;
            if (File.Exists(path))
            {
                using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    if (path.EndsWith(".xlsx"))
                    {
                        workbook = new XSSFWorkbook(file);
                    }
                    else if (path.EndsWith(".xls"))
                    {
                        workbook = new HSSFWorkbook(file);
                    }
                    else
                    {
                        Console.WriteLine("The specified file is not an Excel file.");
                        return string.Empty;
                    }
                }
            }
            else
            {
                return string.Empty;
            }
            ISheet sheet;
            if (isFriday)
            {
                sheet = workbook.CreateSheet("Ergebnis-Listen-Freitag");
            }
            else
            {
                sheet = workbook.CreateSheet("Ergebnis-Listen-Samstag");
            }
            int amountOfNeededRows = 0;
            if (context.OrderedGroupDataFriday != null && context.OrderedGroupDataFriday.Count > amountOfNeededRows)
            {
                amountOfNeededRows = context.OrderedGroupDataFriday.Count;
            }
            if (context.OrderedGroupDataSaturday != null && context.OrderedGroupDataSaturday.Count > amountOfNeededRows)
            {
                amountOfNeededRows = context.OrderedGroupDataSaturday.Count;
            }
            if (context.AllocationResultsFriday != null && context.AllocationResultsFriday.Count > amountOfNeededRows)
            {
                amountOfNeededRows = context.AllocationResultsFriday.Count;
            }
            if (context.AllocationResultsSaturday != null && context.AllocationResultsSaturday.Count > amountOfNeededRows)
            {
                amountOfNeededRows = context.AllocationResultsSaturday.Count;
            }
            if (context.EvaluationResultsFriday != null && context.EvaluationResultsFriday.Count > amountOfNeededRows)
            {
                amountOfNeededRows = context.EvaluationResultsFriday.Count;
            }
            amountOfNeededRows += 2;
            for (int i = 0; i < amountOfNeededRows; i++)
            {
                sheet.CreateRow(i);
            }
            if (isFriday)
            {
                // Sortierte Freitags Liste schreiben
                if (context.OrderedGroupDataFriday != null)
                {
                    ICell cell;
                    IRow row = sheet.GetRow(0);
                    cell = row.CreateCell(0);
                    cell.SetCellValue($"Sortiert Freitag");
                    row = sheet.GetRow(1);
                    cell = row.CreateCell(0);
                    cell.SetCellValue($"ID");
                    cell = row.CreateCell(1);
                    cell.SetCellValue($"Name");
                    cell = row.CreateCell(2);
                    cell.SetCellValue($"Tag");
                    cell = row.CreateCell(3);
                    cell.SetCellValue($"Anzahl");
                    cell = row.CreateCell(4);
                    cell.SetCellValue($"Priorität");
                    for (int i = 0; i < context.OrderedGroupDataFriday.Count; i++)
                    {
                        var group = context.OrderedGroupDataFriday[i];
                        row = sheet.GetRow(i + 2);
                        cell = row.CreateCell(0);
                        cell.SetCellValue($"{group.GroupId}");
                        cell = row.CreateCell(1);
                        cell.SetCellValue($"{group.Name}");
                        cell = row.CreateCell(2);
                        cell.SetCellValue($"{group.Day}");
                        cell = row.CreateCell(3);
                        cell.SetCellValue($"{group.Size}");
                        cell = row.CreateCell(4);
                        cell.SetCellValue($"{group.Score}");
                    }
                }
                // Zuweisung Eregebnis Freitag schreiben
                if (context.AllocationResultsFriday != null)
                {
                    context.AllocationResultsFriday = context.AllocationResultsFriday.OrderByDescending(x => x.AverageScore).ToList();
                    ICell cell;
                    IRow row = sheet.GetRow(0);
                    cell = row.CreateCell(6);
                    cell.SetCellValue($"Zuweisungs Ergebnisse Freitag");
                    row = sheet.GetRow(1);
                    cell = row.CreateCell(6);
                    cell.SetCellValue($"Gruppe");
                    cell = row.CreateCell(7);
                    cell.SetCellValue($"Größe");
                    cell = row.CreateCell(8);
                    cell.SetCellValue($"Hauptsitz");
                    cell = row.CreateCell(9);
                    cell.SetCellValue($"Wertung");
                    cell = row.CreateCell(10);
                    cell.SetCellValue($"Priorität");
                    for (int i = 0; i < context.AllocationResultsFriday.Count; i++)
                    {
                        var result = context.AllocationResultsFriday[i];
                        row = sheet.GetRow(i + 2);
                        cell = row.CreateCell(6);
                        cell.SetCellValue($"{result.Parent.GroupId}");
                        cell = row.CreateCell(7);
                        cell.SetCellValue($"{result.OwnedSeats.Count}");
                        cell = row.CreateCell(8);
                        cell.SetCellValue($"{result.MainSeat.SeatNumber}");
                        cell = row.CreateCell(9);
                        cell.SetCellValue($"{result.AverageScore}");
                        cell = row.CreateCell(10);
                        cell.SetCellValue($"{result.Parent.Score}");
                    }
                }
                // Evaluierungs Ergebnisse
                if (context.EvaluationResultsFriday != null)
                {
                    context.EvaluationResultsFriday = context.EvaluationResultsFriday.OrderByDescending(x => x.AverageScore).ToList();
                    ICell cell;
                    IRow row = sheet.GetRow(0);
                    cell = row.CreateCell(12);
                    cell.SetCellValue($"Evaluierungs Ergebnisse");
                    row = sheet.GetRow(1);
                    cell = row.CreateCell(12);
                    cell.SetCellValue($"Gruppe");
                    cell = row.CreateCell(13);
                    cell.SetCellValue($"Größe");
                    cell = row.CreateCell(14);
                    cell.SetCellValue($"Hauptsitz");
                    cell = row.CreateCell(15);
                    cell.SetCellValue($"Wertung");
                    cell = row.CreateCell(16);
                    cell.SetCellValue($"Priorität");
                    for (int i = 0; i < context.EvaluationResultsFriday.Count; i++)
                    {
                        var result = context.EvaluationResultsFriday[i];
                        row = sheet.GetRow(i + 2);
                        cell = row.CreateCell(12);
                        cell.SetCellValue($"{result.Parent.GroupId}");
                        cell = row.CreateCell(13);
                        cell.SetCellValue($"{result.OwnedSeats.Count}");
                        cell = row.CreateCell(14);
                        cell.SetCellValue($"{result.MainSeat.SeatNumber}");
                        cell = row.CreateCell(15);
                        cell.SetCellValue($"{result.AverageScore}");
                        cell = row.CreateCell(16);
                        cell.SetCellValue($"{result.Parent.Score}");
                    }
                }
            }
            else
            {
                // Sortierte Samstags Liste schreiben
                if (context.OrderedGroupDataSaturday != null)
                {
                    ICell cell;
                    IRow row = sheet.GetRow(0);
                    cell = row.CreateCell(0);
                    cell.SetCellValue($"Sortiert Samstag");
                    row = sheet.GetRow(1);
                    cell = row.CreateCell(0);
                    cell.SetCellValue($"ID");
                    cell = row.CreateCell(1);
                    cell.SetCellValue($"Name");
                    cell = row.CreateCell(2);
                    cell.SetCellValue($"Tag");
                    cell = row.CreateCell(3);
                    cell.SetCellValue($"Anzahl");
                    cell = row.CreateCell(4);
                    cell.SetCellValue($"Priorität");
                    for (int i = 0; i < context.OrderedGroupDataSaturday.Count; i++)
                    {
                        var group = context.OrderedGroupDataSaturday[i];
                        row = sheet.GetRow(i + 2);
                        cell = row.CreateCell(0);
                        cell.SetCellValue($"{group.GroupId}");
                        cell = row.CreateCell(1);
                        cell.SetCellValue($"{group.Name}");
                        cell = row.CreateCell(2);
                        cell.SetCellValue($"{group.Day}");
                        cell = row.CreateCell(3);
                        cell.SetCellValue($"{group.Size}");
                        cell = row.CreateCell(4);
                        cell.SetCellValue($"{group.Score}");
                    }
                }

                // Zuweisung Eregebnis Samstag schreiben
                if (context.AllocationResultsSaturday != null)
                {
                    context.AllocationResultsSaturday = context.AllocationResultsSaturday.OrderByDescending(x => x.AverageScore).ToList();
                    ICell cell;
                    IRow row = sheet.GetRow(0);
                    cell = row.CreateCell(6);
                    cell.SetCellValue($"Zuweisungs Ergebnisse Samstag");
                    row = sheet.GetRow(1);
                    cell = row.CreateCell(6);
                    cell.SetCellValue($"Gruppe");
                    cell = row.CreateCell(7);
                    cell.SetCellValue($"Größe");
                    cell = row.CreateCell(8);
                    cell.SetCellValue($"Hauptsitz");
                    cell = row.CreateCell(9);
                    cell.SetCellValue($"Wertung");
                    cell = row.CreateCell(10);
                    cell.SetCellValue($"Priorität");
                    for (int i = 0; i < context.AllocationResultsSaturday.Count; i++)
                    {
                        var result = context.AllocationResultsSaturday[i];
                        row = sheet.GetRow(i + 2);
                        cell = row.CreateCell(6);
                        cell.SetCellValue($"{result.Parent.GroupId}");
                        cell = row.CreateCell(7);
                        cell.SetCellValue($"{result.OwnedSeats.Count}");
                        cell = row.CreateCell(8);
                        cell.SetCellValue($"{result.MainSeat.SeatNumber}");
                        cell = row.CreateCell(9);
                        cell.SetCellValue($"{result.AverageScore}");
                        cell = row.CreateCell(10);
                        cell.SetCellValue($"{result.Parent.Score}");
                    }
                }
                // Evaluierungs Ergebnisse
                if (context.EvaluationResultsSaturday != null)
                {
                    context.EvaluationResultsSaturday = context.EvaluationResultsSaturday.OrderByDescending(x => x.AverageScore).ToList();
                    ICell cell;
                    IRow row = sheet.GetRow(0);
                    cell = row.CreateCell(12);
                    cell.SetCellValue($"Evaluierungs Ergebnisse");
                    row = sheet.GetRow(1);
                    cell = row.CreateCell(12);
                    cell.SetCellValue($"Gruppe");
                    cell = row.CreateCell(13);
                    cell.SetCellValue($"Größe");
                    cell = row.CreateCell(14);
                    cell.SetCellValue($"Hauptsitz");
                    cell = row.CreateCell(15);
                    cell.SetCellValue($"Wertung");
                    cell = row.CreateCell(16);
                    cell.SetCellValue($"Priorität");
                    for (int i = 0; i < context.EvaluationResultsSaturday.Count; i++)
                    {
                        var result = context.EvaluationResultsSaturday[i];
                        row = sheet.GetRow(i + 2);
                        cell = row.CreateCell(12);
                        cell.SetCellValue($"{result.Parent.GroupId}");
                        cell = row.CreateCell(13);
                        cell.SetCellValue($"{result.OwnedSeats.Count}");
                        cell = row.CreateCell(14);
                        cell.SetCellValue($"{result.MainSeat.SeatNumber}");
                        cell = row.CreateCell(15);
                        cell.SetCellValue($"{result.AverageScore}");
                        cell = row.CreateCell(16);
                        cell.SetCellValue($"{result.Parent.Score}");
                    }
                }
            }
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
            workbook.Close();

            return path;
        }

        public short GetNextColor(int index)
        {
            switch (index)
            {
                case 1:
                    return HSSFColor.Yellow.Index;
                    break;
                case 2:
                    return HSSFColor.Red.Index;
                    break;
                case 3:
                    return HSSFColor.Aqua.Index;
                    break;
                case 4:
                    return HSSFColor.Green.Index;
                    break;
                case 5:
                    return HSSFColor.Brown.Index;
                    break;
                case 6:
                    return HSSFColor.Pink.Index;
                    break;
                case 7:
                    return HSSFColor.Orange.Index;
                    break;
                case 8:
                    return HSSFColor.Plum.Index;
                    break;
                case 9:
                    return HSSFColor.SeaGreen.Index;
                    break;
                default:
                    return HSSFColor.Grey25Percent.Index;
                    break;
            }
        }
    }
}
