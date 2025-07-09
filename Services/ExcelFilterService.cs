// -----------------------------------------------------------------------------
// ExcelFilterService.cs  –  10-Jul-2025
// Splits RL_NEW_PAYABLES_TLC_TM.XLS into P1 (numeric IDINVC) & P2 (alpha-mixed),
// preserving tables, macros *and* styles by cloning them into the template WB.
// -----------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using Microsoft.AspNetCore.Hosting;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ApFilterWebApp.Services
{
    public class ExcelFilterService
    {
        private readonly IWebHostEnvironment _env;

        private const string DefaultSourcePath   = @"\\fs01\Accounting\AP\RL_NEW_PAYABLES_TLC_TM.XLS";
        private const string DefaultOutputFolder = @"\\fs01\Accounting\AP\TM_AP_EXPORT\";
        private const string TemplateRelative    = "Templates/RL_NEW_PAYABLES_TLC_TM.XLS";

        public ExcelFilterService(IWebHostEnvironment env) => _env = env;

        // ---------------------------------------------------------------------
        // PUBLIC ENTRY
        // ---------------------------------------------------------------------
        public void Process(Stream? uploadedSource, string? destFolder)
        {
            var outputFolder = string.IsNullOrWhiteSpace(destFolder)
                ? DefaultOutputFolder
                : destFolder;
            Directory.CreateDirectory(outputFolder);

            // 1. load full source workbook
            var srcWb = LoadSource(uploadedSource);

            // 2. split rows
            var invSheet = srcWb.GetSheet("Invoices");
            var detSheet = srcWb.GetSheet("Invoice_Details");
            var invHdr   = invSheet.GetRow(0);
            var detHdr   = detSheet.GetRow(0);

            var invP1 = ExtractRows(invSheet, invHdr, numericOnly: true,  out var cntP1);
            var invP2 = ExtractRows(invSheet, invHdr, numericOnly: false, out var cntP2);

            var detP1 = FilterDetails(detSheet, detHdr, cntP1);
            var detP2 = FilterDetails(detSheet, detHdr, cntP2);

            // 3. build + save two files
            BuildAndSave("RL_NEW_PAYABLES_TLC_TM_P1.xls", invP1, detP1, outputFolder);
            BuildAndSave("RL_NEW_PAYABLES_TLC_TM_P2.xls", invP2, detP2, outputFolder);
        }

        // ---------------------------------------------------------------------
        // INTERNALS
        // ---------------------------------------------------------------------
        private HSSFWorkbook LoadSource(Stream? uploaded)
        {
            if (uploaded != null) return new HSSFWorkbook(uploaded);

            if (File.Exists(DefaultSourcePath))
                using (var fs = File.OpenRead(DefaultSourcePath))
                    return new HSSFWorkbook(fs);

            throw new FileNotFoundException("Source not found", DefaultSourcePath);
        }

        private void BuildAndSave(
            string          fileName,
            List<IRow>      invRows,
            List<IRow>      detRows,
            string          outputFolder)
        {
            var templatePath = Path.Combine(_env.ContentRootPath, TemplateRelative);
            using var tplFs  = File.OpenRead(templatePath);
            var destWb       = new HSSFWorkbook(tplFs);

            RefreshSheet(destWb.GetSheet("Invoices"),        invRows);
            RefreshSheet(destWb.GetSheet("Invoice_Details"), detRows);

            SaveWorkbookWithRetry(destWb, Path.Combine(outputFolder, fileName));
        }

        // ---------- STYLE-SAFE SHEET REFRESH ---------------------------------
        private static void RefreshSheet(ISheet sheet, List<IRow> newRows)
        {
            var destWb  = (HSSFWorkbook)sheet.Workbook;
            var srcWb   = (HSSFWorkbook)newRows[0].Sheet.Workbook;

            // caches so we clone each style / font only once
            var styleMap = new Dictionary<short, ICellStyle>();
            var fontMap  = new Dictionary<short, short>();          // srcFontIdx → destFontIdx

            // 1. wipe everything below the header
            for (int r = sheet.LastRowNum; r > 0; r--) sheet.RemoveRow(sheet.GetRow(r));

            // 2. copy fresh data
            for (int i = 1; i < newRows.Count; i++)
            {
                var srcRow = newRows[i];
                var dstRow = sheet.CreateRow(i);
                dstRow.Height = srcRow.Height;

                for (int c = 0; c < srcRow.LastCellNum; c++)
                {
                    var sCell = srcRow.GetCell(c);
                    if (sCell == null) continue;

                    var dCell = dstRow.CreateCell(c, sCell.CellType);

                    // ---- value
                    switch (sCell.CellType)
                    {
                        case CellType.String:  dCell.SetCellValue(sCell.StringCellValue);  break;
                        case CellType.Numeric: dCell.SetCellValue(sCell.NumericCellValue); break;
                        case CellType.Boolean: dCell.SetCellValue(sCell.BooleanCellValue); break;
                        case CellType.Formula: dCell.SetCellFormula(sCell.CellFormula);    break;
                        default:               dCell.SetCellValue(sCell.ToString());       break;
                    }

                    // ---- style
                    var sStyleIdx = sCell.CellStyle.Index;
                    if (!styleMap.TryGetValue(sStyleIdx, out var dStyle))
                    {
                        dStyle = destWb.CreateCellStyle();
                        dStyle.CloneStyleFrom(sCell.CellStyle);   // basic formats

                        // clone font if needed
                        var sFontIdx = sCell.CellStyle.FontIndex;
                        if (!fontMap.TryGetValue(sFontIdx, out var dFontIdx))
                        {
                            var sFont = srcWb.GetFontAt(sFontIdx);
                            var dFont = destWb.CreateFont();
                            dFont.Boldweight   = sFont.Boldweight;
                            dFont.Color        = sFont.Color;
                            dFont.FontHeight   = sFont.FontHeight;
                            dFont.FontName     = sFont.FontName;
                            dFont.IsItalic     = sFont.IsItalic;
                            dFont.Underline    = sFont.Underline;
                            dFontIdx           = dFont.Index;
                            fontMap[sFontIdx]  = dFontIdx;
                        }
                        dStyle.SetFont(destWb.GetFontAt(dFontIdx));
                        styleMap[sStyleIdx] = dStyle;
                    }
                    dCell.CellStyle = dStyle;
                }
            }
        }

        // ---------- SPLIT HELPERS -------------------------------------------
        private static List<IRow> ExtractRows(
            ISheet sheet, IRow header, bool numericOnly, out HashSet<string> cntSet)
        {
            int idCol  = GetCol(header, "IDINVC");
            int cntCol = GetCol(header, "CNTITEM");

            var rows   = new List<IRow> { header };
            cntSet     = new HashSet<string>();

            for (int r = 1; r <= sheet.LastRowNum; r++)
            {
                var row   = sheet.GetRow(r);
                var idVal = row?.GetCell(idCol)?.ToString() ?? "";
                bool ok   = numericOnly ? Regex.IsMatch(idVal, @"^\d+$")
                                        : !Regex.IsMatch(idVal, @"^\d+$");

                if (ok && row != null)
                {
                    rows.Add(row);
                    cntSet.Add(row.GetCell(cntCol)?.ToString() ?? "");
                }
            }
            return rows;
        }

        private static List<IRow> FilterDetails(
            ISheet sheet, IRow header, HashSet<string> allowedCnt)
        {
            int cntCol = GetCol(header, "CNTITEM");
            var rows   = new List<IRow> { header };

            for (int r = 1; r <= sheet.LastRowNum; r++)
            {
                var row = sheet.GetRow(r);
                if (row != null &&
                    allowedCnt.Contains(row.GetCell(cntCol)?.ToString() ?? ""))
                    rows.Add(row);
            }
            return rows;
        }

        private static int GetCol(IRow header, string name) =>
            header.Cells.Select((c, i) => (c.StringCellValue, i))
                        .First(p => p.StringCellValue == name).i;

        private static void SaveWorkbookWithRetry(HSSFWorkbook wb, string path)
        {
            const int max = 5;
            for (int attempt = 1; attempt <= max; attempt++)
            {
                try
                {
                    if (File.Exists(path)) File.Delete(path);
                    using var fs = new FileStream(path, FileMode.Create,
                                                  FileAccess.Write, FileShare.None);
                    wb.Write(fs);
                    return;
                }
                catch (IOException) when (attempt < max) { Thread.Sleep(1000); }
            }
            throw new IOException($"Unable to save '{path}' after {max} attempts.");
        }
    }
}
