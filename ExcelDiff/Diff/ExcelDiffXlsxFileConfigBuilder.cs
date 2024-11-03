﻿using OfficeOpenXml;

namespace ExcelDiffEngine;

public class ExcelDiffXlsxFileConfigBuilder
{
    private XlsxFileInfo oldFile = new("Unspecified.xlsx");
    private XlsxFileInfo newFile = new("Unspecified.xlsx");

    public ExcelDiffXlsxFileConfigBuilder SetOldFile(string filePath, Action<ExcelPackage>? callback = null)
    {
        oldFile = oldFile with { FileInfo = new(filePath), Callback = callback };
        return this;
    }

    public ExcelDiffXlsxFileConfigBuilder SetNewFile(string filePath, Action<ExcelPackage>? callback = null)
    {
        newFile = newFile with { FileInfo = new(filePath), Callback = callback };
        return this;
    }

    public ExcelDiffXlsxFileConfigBuilder SetMergedWorksheetName(string mergedWorksheetName)
    {
        oldFile = oldFile with { MergedWorksheetName = mergedWorksheetName };
        newFile = newFile with { MergedWorksheetName = mergedWorksheetName };
        return this;
    }

    public ExcelDiffXlsxFileConfigBuilder SetDocumentName(string documentName)
    {
        oldFile = oldFile with { DocumentName = documentName };
        newFile = newFile with { DocumentName = documentName };
        return this;
    }

    public ExcelDiffXlsxFileConfigBuilder AddOldFileWorksheetInfo(string worksheetName, int fromRow = 1, int fromColumn = 1, int? toRow = null, int? toColumn = null)
    {
        oldFile = oldFile with
        {
            WorksheetInfos = [.. oldFile.WorksheetInfos, new() { Name = worksheetName, FromRow = fromRow, FromColumn = fromColumn, ToRow = toRow, ToColumn = toColumn }]
        };
        return this;
    }

    public ExcelDiffXlsxFileConfigBuilder AddNewFileWorksheetInfo(string worksheetName, int fromRow = 1, int fromColumn = 1, int? toRow = null, int? toColumn = null)
    {
        newFile = newFile with
        {
            WorksheetInfos = [.. newFile.WorksheetInfos, new() { Name = worksheetName, FromRow = fromRow, FromColumn = fromColumn, ToRow = toRow, ToColumn = toColumn }]
        };
        return this;
    }

    public ExcelDiffXlsxFileConfigBuilder AddWorksheetInfo(string worksheetName, int fromRow = 1, int fromColumn = 1, int? toRow = null, int? toColumn = null)
    {
        _ = AddOldFileWorksheetInfo(worksheetName, fromRow, fromColumn, toRow, toColumn);
        _ = AddNewFileWorksheetInfo(worksheetName, fromRow, fromColumn, toRow, toColumn);
        return this;
    }

    public ExcelDiffXlsxFileConfigBuilder SetDataArea(int fromRow = 1, int fromColumn = 1, int? toRow = null, int? toColumn = null)
    {
        oldFile = oldFile with { FromRow = fromRow, FromColumn = fromColumn, ToRow = toRow, ToColumn = toColumn };
        newFile = newFile with { FromRow = fromRow, FromColumn = fromColumn, ToRow = toRow, ToColumn = toColumn };
        return this;
    }

    internal (XlsxFileInfo oldFile, XlsxFileInfo newFile) Build()
    {
        if (oldFile.MergedWorksheetName is null && newFile.MergedWorksheetName is null)
        {
            oldFile = oldFile with { MergedWorksheetName = Path.GetFileNameWithoutExtension(newFile.FileInfo.Name) };
            newFile = newFile with { MergedWorksheetName = Path.GetFileNameWithoutExtension(newFile.FileInfo.Name) };
        }
        return (oldFile, newFile);
    }

}