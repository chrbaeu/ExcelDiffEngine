﻿using OfficeOpenXml;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelDiffEngine;

internal sealed class ModificationRuleHandler
{
    private readonly List<ModificationRule> columnNameRules = [];
    private readonly List<(Regex regex, ModificationRule)> regexRules = [];
    private readonly StringComparer stringComparer;

    internal ModificationRuleHandler(IReadOnlyList<ModificationRule> rules, bool ignoreCase)
    {
        RegexOptions options = ignoreCase ? RegexOptions.IgnoreCase : RegexOptions.None;
        foreach (var rule in rules)
        {
            if (rule.Match[0] == '@')
            {
                regexRules.Add((new Regex(rule.Match[1..], options), rule));
            }
            else
            {
                columnNameRules.Add(rule);
            }
        }
        stringComparer = ignoreCase ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
    }

    internal void ApplyRules(ExcelRange excelCell, string columnName, DataKind dataKind)
    {
        foreach (var (regex, rule) in regexRules)
        {
            if (regex.IsMatch(columnName)) { ApplyRule(excelCell, rule, dataKind); }
        }
        foreach (var rule in columnNameRules)
        {
            if (stringComparer.Equals(rule.Match, columnName)) { ApplyRule(excelCell, rule, dataKind); }
        }
    }

    private void ApplyRule(ExcelRange excelCell, ModificationRule rule, DataKind dataKind)
    {
        if (rule.Target != DataKind.All && rule.Target != dataKind) { return; }
        switch (rule.Type)
        {
            case ':':
                excelCell.Style.Numberformat.Format = rule.Value;
                break;
            case '*':
                excelCell.Value = (double?)excelCell.Value * double.Parse(rule.Value, CultureInfo.InvariantCulture);
                break;
            case '=':
                excelCell.Formula = rule.Value.Replace("{#}", (((double?)excelCell.Value) ?? 0).ToString(CultureInfo.InvariantCulture));
                excelCell.Calculate();
                break;
            default:
                break;
        }
    }
}