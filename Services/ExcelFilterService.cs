// Created by James Fallouh
// Date: 2025-07-07
// Meta: Split RL_NEW_PAYABLES_TLC_TM.XLS into P1 (digits-only IDINVC) & P2 (others),
//       preserving ALL sheets + formatting, and retrying saves to avoid corruption.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Microsoft.AspNetCore.Hosting;

namespace ApFilterWebApp.Services
{
    public class ExcelFilterService
    {
        private readonly IWebHostEnvironment _env;
        private const string DefaultSourcePath    = @"\\fs01\Accounting\AP\RL_NEW_PAYABLES_TLC_TM.XLS";
        private const string DefaultOutputFolder  = @"\\fs01\Accounting\AP\TM_AP_EXPORT\";
        private const string TemplateRelativePath = "Templates/RL_NEW_PAYABLES_TLC_TM.XLS";

        public ExcelFilterService(IWebHostEnvironment env)
            => _env = env;

        public void Process(Stream? uploadedSource, string? destFolder)
        {
            // 1. Prepare folders & streams
            var outputFolder = string.IsNullOrWhiteSpace(destFolder)
                ? DefaultOutputFolder
                : destFolder;
            Directory.CreateDirectory(outputFolder);

            // 2. Load source workbook
            HSSFWorkbook srcWb;
            if (uploadedSource != null)
            {
                srcWb = new HSSFWorkbook(uploadedSource);
            }
            else if (File.Exists(DefaultSourcePath))
            {
                using var fs = File.OpenRead(DefaultSourcePath);
                srcWb = new HSSFWorkbook(fs);
            }
            else
            {
                throw new FileNotFoundException("Source not found", DefaultSourcePath);
            }

            // 3. Extract rows for P1 & P2
            var invSheet  = srcWb.GetSheet("Invoices");
            var detSheet  = srcWb.GetSheet("Invoice_Details");
            var invHdr    = invSheet.GetRow(0);
            var detHdr    = detSheet.GetRow(0);

            var invP1 = ExtractRows(invSheet, invHdr, numericOnly: true,  out var cntP1);
            var invP2 = ExtractRows(invSheet, invHdr, numericOnly: false, out var cntP2);

            var detP1 = FilterDetails(detSheet, detHdr, cntP1);
            var detP2 = FilterDetails(detSheet, detHdr, cntP2);

            // 4. Build P1 workbook and save
            var wbP1 = CreateWorkbook(srcWb, invP1, invHdr, detP1, detHdr);
            SaveWorkbookWithRetry(wbP1, Path.Combine(outputFolder, "RL_NEW_PAYABLES_TLC_TM_P1.xls"));

            // 5. Build P2 workbook and save
            var wbP2 = CreateWorkbook(srcWb, invP2, invHdr, detP2, detHdr);
            SaveWorkbookWithRetry(wbP2, Path.Combine(outputFolder, "RL_NEW_PAYABLES_TLC_TM_P2.xls"));
        }

        private static List<IRow> ExtractRows(
            ISheet sheet,
            IRow  header,
            bool  numericOnly,
            out HashSet<string> outCnt)
        {
            int idCol  = header.Cells
                         .Select((c,i)=>(c.StringCellValue,i))
                         .First(p => p.StringCellValue=="IDINVC").i;
            int cntCol = header.Cells
                         .Select((c,i)=>(c.StringCellValue,i))
                         .First(p => p.StringCellValue=="CNTITEM").i;

            var rows   = new List<IRow> { header };
            outCnt     = new HashSet<string>();

            for (int r = 1; r <= sheet.LastRowNum; r++)
            {
                var row    = sheet.GetRow(r);
                var idVal  = row?.GetCell(idCol)?.ToString() ?? "";
                bool isNum = Regex.IsMatch(idVal, @"^\d+$");

                if ((numericOnly && isNum) || (!numericOnly && !isNum))
                {
                    rows.Add(row!);
                    outCnt.Add(row!.GetCell(cntCol)?.ToString() ?? "");
                }
            }

            return rows;
        }

        private static List<IRow> FilterDetails(
            ISheet sheet,
            IRow  header,
            HashSet<string> allowedCnt)
        {
            int cntCol = header.Cells
                         .Select((c,i)=>(c.StringCellValue,i))
                         .First(p => p.StringCellValue=="CNTITEM").i;

            var rows = new List<IRow> { header };
            for (int r = 1; r <= sheet.LastRowNum; r++)
            {
                var row    = sheet.GetRow(r);
                var cntVal = row?.GetCell(cntCol)?.ToString() ?? "";
                if (allowedCnt.Contains(cntVal))
                    rows.Add(row!);
            }
            return rows;
        }

        private static HSSFWorkbook CreateWorkbook(
            HSSFWorkbook   srcWb,
            List<IRow>     invRows,
            IRow           invHdr,
            List<IRow>     detRows,
            IRow           detHdr)
        {
            var destWb   = new HSSFWorkbook();
            CopySheet(destWb, srcWb, "Invoices",        invRows, invHdr);
            CopySheet(destWb, srcWb, "Invoice_Details", detRows, detHdr);

            // copy every other sheet unchanged
            foreach (ISheet src in srcWb)
            {
                if (src.SheetName is "Invoices" or "Invoice_Details") continue;

                var allRows = Enumerable
                              .Range(0, src.LastRowNum+1)
                              .Select(r => src.GetRow(r))
                              .Where(r => r != null)
                              .Cast<IRow>()
                              .ToList();

                CopySheet(destWb, srcWb, src.SheetName, allRows, allRows[0]);
            }

            return destWb;
        }

        private static void CopySheet(
            HSSFWorkbook      destWb,
            HSSFWorkbook      srcWb,
            string            name,
            List<IRow>        srcRows,
            IRow              srcHeader)
        {
            var destSheet = destWb.CreateSheet(name);
            var styleMap  = new Dictionary<short,ICellStyle>();
            var fontMap   = new Dictionary<short,short>();

            // copy column widths
            for (int c = 0; c < srcHeader.LastCellNum; c++)
                destSheet.SetColumnWidth(c, srcHeader.Sheet.GetColumnWidth(c));

            for (int r = 0; r < srcRows.Count; r++)
            {
                var sRow = srcRows[r];
                var dRow = destSheet.CreateRow(r);
                dRow.Height = sRow.Height;

                for (int c = 0; c < srcHeader.LastCellNum; c++)
                {
                    var sc = sRow.GetCell(c);
                    if (sc == null) continue;

                    var dc = dRow.CreateCell(c, sc.CellType);
                    switch (sc.CellType)
                    {
                        case CellType.String:  dc.SetCellValue(sc.StringCellValue);  break;
                        case CellType.Numeric: dc.SetCellValue(sc.NumericCellValue); break;
                        case CellType.Boolean: dc.SetCellValue(sc.BooleanCellValue);break;
                        case CellType.Formula: dc.SetCellFormula(sc.CellFormula);   break;
                        default:               dc.SetCellValue(sc.ToString());      break;
                    }

                    var si = sc.CellStyle.Index;
                    if (!styleMap.TryGetValue(si, out var dstStyle))
                    {
                        dstStyle = destWb.CreateCellStyle();
                        dstStyle.CloneStyleFrom(sc.CellStyle);

                        var sfi = sc.CellStyle.FontIndex;
                        if (!fontMap.TryGetValue(sfi, out var dfi))
                        {
                            var sf = srcWb.GetFontAt(sfi);
                            var df = destWb.CreateFont();
                            df.Boldweight    = sf.Boldweight;
                            df.Color         = sf.Color;
                            df.FontHeight    = sf.FontHeight;
                            df.FontName      = sf.FontName;
                            df.IsItalic      = sf.IsItalic;
                            df.Underline     = sf.Underline;
                            dfi = df.Index;
                            fontMap[sfi] = dfi;
                        }
                        dstStyle.SetFont(destWb.GetFontAt(dfi));
                        styleMap[si] = dstStyle;
                    }
                    dc.CellStyle = dstStyle;
                }
            }
        }

        private static void SaveWorkbookWithRetry(HSSFWorkbook wb, string path)
        {
            const int maxAttempts = 5;
            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    if (File.Exists(path)) File.Delete(path);
                    using var fs = new FileStream(path,
                                                  FileMode.Create,
                                                  FileAccess.Write,
                                                  FileShare.None);
                    wb.Write(fs);
                    return;
                }
                catch (IOException) when (attempt < maxAttempts)
                {
                    Thread.Sleep(1000);
                }
            }
            throw new IOException($"Could not save '{path}' after multiple attempts.");
        }
    }
}