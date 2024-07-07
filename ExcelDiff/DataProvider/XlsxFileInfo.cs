﻿namespace ExcelDiffEngine;

public record XlsxWorksheetInfo
{
    public string Name { get; set; } = "";
    public int FromRow { get; init; } = 1;
    public int FromColumn { get; init; } = 1;
    public int? ToRow { get; init; }
    public int? ToColumn { get; init; }
}

public record class XlsxFileInfo
{
    public XlsxFileInfo(string fileName)
    {
        FileInfo = new(fileName);
    }

    public XlsxFileInfo(FileInfo fileInfo)
    {
        FileInfo = fileInfo;
    }

    public FileInfo FileInfo { get; init; }

    public int FromRow { get; init; } = 1;
    public int FromColumn { get; init; } = 1;
    public int? ToRow { get; init; }
    public int? ToColumn { get; init; }

    public string? MergedWorksheetName { get; init; }

    public IReadOnlyList<XlsxWorksheetInfo> WorksheetInfos { get; init; } = [];
}