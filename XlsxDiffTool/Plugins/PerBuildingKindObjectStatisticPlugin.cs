using OfficeOpenXml;
using XlsxDiffEngine;
using XlsxDiffTool.Models;

namespace XlsxDiffTool.Plugins;

public class PerBuildingKindObjectStatisticPlugin : IPlugin
{
    public string Name => "PerBuildingKindObjectStatisticPlugin";

    public string Tooltip => "Modifies the excel files of the object statistics for better compersion";

    public void OnExcelPackageLoading(DiffConfigModel diffConfigModel, ExcelPackage excelPackage)
    {
        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];
        if (worksheet.Cells[7, 1].Text != "Statsitik") { return; }
        int sourceRow = 14;
        int destinationRow = 7;
        worksheet.Cells[destinationRow, 1, destinationRow, worksheet.Dimension.End.Column].Value =
            worksheet.Cells[sourceRow, 1, sourceRow, worksheet.Dimension.End.Column].Value;
        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
        {
            worksheet.Cells[destinationRow, col].StyleID = worksheet.Cells[sourceRow, col].StyleID;
        }
        worksheet.DeleteRow(sourceRow);
        worksheet.Cells[8, 1].Value = "Min";
        worksheet.Cells[9, 1].Value = "Von";
        worksheet.Cells[10, 1].Value = "Mittel";
        worksheet.Cells[11, 1].Value = "Bis";
        worksheet.Cells[12, 1].Value = "Max";
        var temp = worksheet.Cells[8, 5].Value;
        Move(worksheet, 11, 11);
        Move(worksheet, 10, 8);
        Move(worksheet, 9, 5);
        Move(worksheet, 8, 8);
        worksheet.Cells[11, 2].Value = temp;
        worksheet.Cells[8, 3, 11, 3].Value = null;
        worksheet.Cells[8, 4, 11, 4].Value = null;
        worksheet.Cells[8, 6, 11, 6].Value = null;
        foreach (var ws in excelPackage.Workbook.Worksheets)
        {
            ws.Cells[7, 1].Value = "Kennwert | Objekt";
        }
    }

    private static void Move(ExcelWorksheet worksheet, int srcRow, int dstColumn)
    {
        worksheet.Cells[8, dstColumn].Value = worksheet.Cells[srcRow, 2].Value;
        worksheet.Cells[9, dstColumn].Value = worksheet.Cells[srcRow, 3].Value;
        worksheet.Cells[10, dstColumn].Value = worksheet.Cells[srcRow, 4].Value;
        worksheet.Cells[11, dstColumn].Value = worksheet.Cells[srcRow, 5].Value;
        worksheet.Cells[12, dstColumn].Value = worksheet.Cells[srcRow, 6].Value;
    }

    public Task OnDiffCreation(DiffConfigModel diffConfigModel, ExcelDiffBuilder excelDiffBuilder)
    {
        return Task.CompletedTask;
    }

    public void OnExcelPackageSaving(DiffConfigModel diffConfigModel, ExcelPackage excelPackage)
    {
        return;
    }
}
