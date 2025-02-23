using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using WindowsFormsApp1.Class;
using Application = Microsoft.Office.Interop.Word.Application;


namespace PrintKernel
{
    class WordBase10 : IDisposable
    {
        public string LastError { get; set; }
        public Application wordApp10 { get; set; }
        public Document aDoc { get; set; }
        object _nullobj = System.Reflection.Missing.Value;
        object missing = System.Reflection.Missing.Value;
        object unit = WdUnits.wdStory;

        public void Init()
        {
            LastError = null;
        }

        /// <summary>
        /// 匯出作業
        /// </summary>
        /// <param name="pSourceFileName">WORD範本路徑</param>
        /// <param name="pTargetFilePath">暫存路徑</param>
        /// <param name="pDt">報表繫節資料(DataTable)</param>
        public void Export(String pTemplateFileName, String pTargetFileName, System.Data.DataTable pDt, List<TableFormat> formatDt, System.Data.DataTable dataDt, String pFileType, String pIsTempFile)
        {
            //try
            //{
            if (pDt != null && pDt.Rows.Count > 0)
            {
                foreach (DataRow dr in pDt.Rows)
                {
                    string folderPath = System.Windows.Forms.Application.StartupPath + @"\" + "ECS";
                    Console.WriteLine(string.Format("{0} 從範本複製一份到新檔案(ECS)", DateTime.Now));
                    Random rand = new Random();
                    //從範本複製一份到暫存檔案
                    string TempFileName = Guid.NewGuid().ToString() + ".doc";
                    CreateFile(pTemplateFileName, TempFileName, folderPath);

                    //替換圖檔
                    Console.WriteLine(string.Format("{0} 替換圖檔", DateTime.Now));
                    AddImage(dr);

                    #region"20240919表格修改"
                    // 這裡處理表格格式 
                    if (formatDt != null)
                    {
                        Console.WriteLine(string.Format("{0} 處理表格格式", DateTime.Now));
                        HandleTableFormatting(formatDt);
                    }


                    //處理完格式再往表格裡面到資料
                    if (dataDt != null)
                    {
                        Console.WriteLine(string.Format("{0} 填入表格資料", DateTime.Now));
                        FillTableData(dataDt);
                    }


                    // 處理合併欄位
                    if (formatDt != null)
                    {
                        Console.WriteLine(string.Format("{0} 合併欄位", DateTime.Now));
                        MergeTableCells(formatDt);
                    }
                    #endregion

                    //替換文字
                    Console.WriteLine(string.Format("{0} 替換文字", DateTime.Now));
                    ReplaceWordTxtFromDatatable(dr);

                    ConvertToSuperscript2();

                    string OutputPath = ConfigurationManager.AppSettings["LISFilePath"];

                    if (pIsTempFile.ToUpper() == "Y")
                    {
                        //word 放到網站主機 D:\temp 資料夾
                        OutputPath = ConfigurationManager.AppSettings["LISTempFilePath"];    //D:\temp
                    }
                    else
                    {
                        //pdf 存到網站主機 D:\IN 資料夾
                        OutputPath = ConfigurationManager.AppSettings["LISFilePath"];        //D:\IN
                    }

                    Console.WriteLine(string.Format("{0} {1} 產出檔案中...", DateTime.Now, OutputPath));
                    string TempFile = string.Format("{0}\\{1}", folderPath, TempFileName);
                    string outPutFullFileName = OutputPath + pTargetFileName;
                    CheckFolder(Path.GetDirectoryName(outPutFullFileName));

                    switch (pFileType)
                    {
                        case "doc":
                            Close();
                            File.Copy(TempFile, outPutFullFileName, true);
                            break;

                        case "pdf":
                            SaveToPdf(outPutFullFileName);
                            break;
                    }

                    File.Delete(TempFile);

                    Console.WriteLine(string.Format("{0} 完成", DateTime.Now));

                }
            }
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message);
            //}
        }


        #region"20240919表格修改"
        public void HandleTableFormatting(List<TableFormat> formatDt)
        {
            if (formatDt == null || formatDt.Count == 0)
                return;

            foreach (var tableFormat in formatDt)
            {
                if (tableFormat.IsInsert == 1)
                {
                    Table table = aDoc.Tables[tableFormat.TableNum];
                    //Row row = table.Rows[tableFormat.RowNum];
                    Row row = table.Rows.Add(table.Rows[tableFormat.RowNum]);

                    // 設定行高
                    row.Height = Convert.ToSingle(tableFormat.RowHeight * 28.35); // 高度以毫米為單位, 1cm = 28.35pt

                    //row.HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightExactly;
                    row.HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAuto;

                }
                //else
                //{
                //    Table table = aDoc.Tables[tableFormat.TableNum];
                //    Row row = table.Rows[tableFormat.RowNum];
                //    //Row row = table.Rows.Add(table.Rows[tableFormat.RowNum]);

                //    // 設定行高
                //    row.Height = Convert.ToSingle(tableFormat.RowHeight * 28.35); // 高度以毫米為單位, 1cm = 28.35pt

                //    //row.HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightExactly;
                //    row.HeightRule = Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAuto;
                //}

            }
            foreach (var tableFormat in formatDt.OrderByDescending(c => c.RowNum))
            {
                int i = 0;
                foreach (var cellFormat in tableFormat.Cells.OrderByDescending(c => c.CellNum))
                {
                    i++;
                    if (tableFormat.RowNum == 0 && cellFormat.CellNum == 0)
                    {
                        continue;
                    }
                    Cell cell;

                    if (i >= 1)
                    {
                        cell = aDoc.Content.Tables[tableFormat.TableNum].Cell(tableFormat.RowNum, cellFormat.CellNum);
                    }
                    else
                    {
                        cell = aDoc.Content.Tables[tableFormat.TableNum].Cell(tableFormat.RowNum, cellFormat.CellNum);
                    }


                    if (cellFormat.SplitType == "Vertical")
                    {
                        cell.Range.Select();
                        cell.Split(1, 2); // 分成2列
                        // 設置新單元格內部邊框樣式
                        Cell newCell1 = aDoc.Content.Tables[tableFormat.TableNum].Cell(tableFormat.RowNum, cellFormat.CellNum);
                        Cell newCell2 = aDoc.Content.Tables[tableFormat.TableNum].Cell(tableFormat.RowNum, cellFormat.CellNum + 1);
                        if (cellFormat.SplitShareLine)
                        {
                            SetSharedOuterBorders(newCell1, newCell2, isVertical: true);
                        }
                    }
                    else if (cellFormat.SplitType == "Horizontal")
                    {
                        cell.Range.Select();
                        cell.Split(2, 1); // 分成2行
                        // 設置新單元格內部邊框樣式
                        Cell newCell1 = aDoc.Content.Tables[tableFormat.TableNum].Cell(tableFormat.RowNum, cellFormat.CellNum);
                        Cell newCell2 = aDoc.Content.Tables[tableFormat.TableNum].Cell(tableFormat.RowNum + 1, cellFormat.CellNum);
                        if (cellFormat.SplitShareLine)
                        {
                            SetSharedOuterBorders(newCell1, newCell2, isVertical: false);
                        }

                    }
                }
            }
        }
        public void SetSharedOuterBorders(Cell cell1, Cell cell2, bool isVertical)
        {
            if (isVertical)
            {
                cell1.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleNone;
                cell2.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleNone;
            }
            else
            {
                cell1.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
                cell2.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
            }
        }

        public void FillTableData(System.Data.DataTable dataDt)
        {
            if (dataDt == null || dataDt.Rows.Count == 0)
                return;


            foreach (DataRow row in dataDt.Rows)
            {
                int tableIndex = Convert.ToInt32(row["TableNum"]);
                int rowIndex = Convert.ToInt32(row["RowNum"]);
                int cellIndex = Convert.ToInt32(row["CellNum"]);
                string cellData = row["CellData"].ToString();

                Table table = aDoc.Tables[tableIndex];
                table.Cell(rowIndex, cellIndex).Range.Text = cellData;
            }
        }


        public void MergeTableCells(List<TableFormat> formatDt)
        {
            if (formatDt == null || formatDt.Count == 0)
                return;

            foreach (var tableFormat in formatDt.OrderByDescending(c => c.RowNum))
            {
                Table table = aDoc.Tables[tableFormat.TableNum];
                foreach (var cellFormat in tableFormat.Cells.OrderByDescending(c => c.CellNum))
                {
                    // Skip if MergeDirection is empty or MergeCount is 0
                    if (string.IsNullOrEmpty(cellFormat.MergeDirection) || cellFormat.MergeCount <= 1)
                        continue;

                    try
                    {
                        if (cellFormat.MergeDirection == "Horizontal")
                        {
                            // Ensure the end cell exists before attempting to merge
                            int endCellIndex = cellFormat.CellNum + cellFormat.MergeCount - 1;
                            if (endCellIndex <= table.Columns.Count)
                            {
                                Cell startCell = table.Cell(tableFormat.RowNum, cellFormat.CellNum);
                                Cell endCell = table.Cell(tableFormat.RowNum, endCellIndex);
                                startCell.Merge(endCell);
                            }
                            else
                            {
                                Console.WriteLine("Cannot merge horizontally: End cell index " + endCellIndex + " exceeds column count " + table.Columns.Count);
                            }
                        }
                        else if (cellFormat.MergeDirection == "Vertical")
                        {
                            // Ensure the end cell exists before attempting to merge
                            int endRowIndex = tableFormat.RowNum + cellFormat.MergeCount - 1;
                            if (endRowIndex <= table.Rows.Count)
                            {
                                Cell startCell = table.Cell(tableFormat.RowNum, cellFormat.CellNum);
                                Cell endCell = table.Cell(endRowIndex, cellFormat.CellNum);
                                startCell.Merge(endCell);
                            }
                            else
                            {
                                Console.WriteLine("Cannot merge vertically: End row index " + endRowIndex + " exceeds row count " + table.Rows.Count);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error merging cells : " + ex.Message);
                    }
                }
            }
        }
        #endregion


        public void ConvertToSuperscript2()
        {
            try
            {
                GotoTheBegining(); // 確保從文件開頭開始搜尋
                this.wordApp10.Selection.Find.ClearFormatting();

                // 設定搜尋條件
                this.wordApp10.Selection.Find.Text = "#@*@#"; // 使用萬用字元找到 #@...@#
                this.wordApp10.Selection.Find.MatchWildcards = true; // 啟用萬用字元

                // 反覆搜尋，直到找不到符合條件的內容
                while (this.wordApp10.Selection.Find.Execute())
                {
                    this.wordApp10.Selection.Select(); // 選取找到的文字

                    string selectedText = this.wordApp10.Selection.Text;

                    // 確保格式符合 "#@...@#"
                    if (selectedText.StartsWith("#@") && selectedText.EndsWith("@#"))
                    {
                        // 取得實際內容（去掉 "#@" 和 "@#"）
                        string processedText = selectedText.Substring(2, selectedText.Length - 4);

                        // 直接修改選取範圍的文字內容
                        this.wordApp10.Selection.Text = processedText;

                        // 重新選取剛剛替換的文字
                        this.wordApp10.Selection.MoveLeft(WdUnits.wdCharacter, processedText.Length, WdMovementType.wdExtend);

                        // 設定上標
                        this.wordApp10.Selection.Font.Superscript = 1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生錯誤：" + ex.Message);
            }
            finally
            {
                GotoTheBegining(); // 完成後將游標移回文件開頭
                this.wordApp10.Selection.Find.ClearFormatting();
                this.wordApp10.Selection.Find.MatchWildcards = false; // 關閉萬用字元模式
            }
        }


        /// <summary>
        /// 保存為PDF
        /// </summary>
        /// <param name="pfileName"></param>
        public void SaveToPdf(string pfileName)
        {
            //try
            //{
            //要匯出為PDF格式，也可以選擇wdExportFormatXPS要匯出為XPS
            const WdExportFormat exportFormat = WdExportFormat.wdExportFormatPDF;
            //轉換完後是否要開啟完成檔，這要選false，不然檔案開在Server端
            const bool openAfterExport = false;
            //wdExportOptimizeForPrint較高品質，wdExportOptimizeForOnScreen，較低品質
            const WdExportOptimizeFor optimizeFor = WdExportOptimizeFor.wdExportOptimizeForPrint;
            //轉換範圍為全部頁數
            const WdExportRange range = WdExportRange.wdExportAllDocument;
            const WdExportItem item = WdExportItem.wdExportDocumentWithMarkup;
            const bool includeDocProps = false;
            const bool keepIrm = false;
            const WdExportCreateBookmarks createBookmarks = WdExportCreateBookmarks.wdExportCreateHeadingBookmarks;
            const bool docStructureTags = false;
            const bool bitmapMissingFonts = false;
            const bool useIso190051 = false;

            aDoc.ExportAsFixedFormat(pfileName,
                                     exportFormat,
                                     openAfterExport,
                                     optimizeFor,
                                     range,
                                     0,
                                     0,
                                     item,
                                     includeDocProps,
                                     keepIrm,
                                     createBookmarks,
                                     docStructureTags,
                                     bitmapMissingFonts,
                                     useIso190051
                                    );
            //}
            //catch (Exception)
            //{

            //}
            //finally
            //{
            Close();
            //}
        }

        /// <summary>
        /// 從範本複製一份到新檔案
        /// </summary>
        /// <param name="pTemplateFileName">範本檔名</param>
        /// <param name="pTargetPath">暫存WORD檔案資料夾</param>
        /// <param name="IsVisible">是否顯示處理過程</param>
        public void CreateFile(String pTemplateFileName, string pTempFileName, string folderPath, bool IsVisible = false)
        {
            //try
            //{
            //LIS網站Word範本 => 建立暫存檔
            
            CreateTempFile(pTemplateFileName, pTempFileName, folderPath);

            //是否顯示處理過程
            wordApp10 = new Application { Visible = IsVisible };

            //開啟檔案
            string OutputFile = string.Format("{0}\\{1}", folderPath, pTempFileName);
            aDoc = wordApp10.Documents.Open(OutputFile, ReadOnly: false, Visible: IsVisible);

            //啟用
            aDoc.Activate();
            //}
            //catch (Exception ex)
            //{
            //}
        }

        public void CreateTempFile(string pFileName, string pTempFileName, string folderPath)
        {
            //try
            //{
            // 檢查資料夾是否存在
            if (!Directory.Exists(folderPath))
            {
                // 如果不存在，創建資料夾
                Directory.CreateDirectory(folderPath);
            }

            string OutputFile = string.Format("{0}\\{1}", folderPath, pTempFileName);
            string url = ConfigurationManager.AppSettings["LISWordUrl"];
            //從LIS網站抓Word範本
            string UrlFilePath = string.Format("{0}/{1}", url, pFileName);
            new WebClient().DownloadFile(UrlFilePath, OutputFile);
            //}
            //catch (Exception ex)
            //{
            //    InsertLog("", ex.Message);
            //}
        }

        // 轉到文檔開頭(非必要少用，耗效能)
        public void GotoTheBegining()
        {
            try
            {
                wordApp10.Selection.HomeKey(unit, missing);
            }
            catch (Exception)
            {
            }
        }

        //保存為PDF文件
        public void SaveAsPDF(String paramExportFilePath)
        {
            try
            {
                //要匯出為PDF格式，也可以選擇wdExportFormatXPS要匯出為XPS
                WdExportFormat paramExportFormat = WdExportFormat.wdExportFormatPDF;
                //轉換完後是否要開啟完成檔，這要選false，不然檔案開在Server端
                Boolean paramOpenAfterExport = false;
                //wdExportOptimizeForPrint較高品質，wdExportOptimizeForOnScreen，較低品質
                WdExportOptimizeFor paramExportOptimizeFor = WdExportOptimizeFor.wdExportOptimizeForPrint;
                //轉換範圍為全部頁數
                WdExportRange paramExportRange = WdExportRange.wdExportAllDocument;
                Int32 paramStartPage = 0;
                Int32 paramEndPage = 0;
                WdExportItem paramExportItem = WdExportItem.wdExportDocumentContent;
                Boolean paramIncludeDocProps = true;
                Boolean paramKeepIRM = true;
                WdExportCreateBookmarks paramCreateBookmarks = WdExportCreateBookmarks.wdExportCreateWordBookmarks;
                Boolean paramDocStructureTags = true;
                Boolean paramBitmapMissingFonts = true;
                Boolean paramUseISO19005_1 = true;
                CheckFolder(Path.GetDirectoryName(paramExportFilePath));

                aDoc.ExportAsFixedFormat(paramExportFilePath, paramExportFormat, paramOpenAfterExport, paramExportOptimizeFor, paramExportRange, paramStartPage, paramEndPage, paramExportItem, paramIncludeDocProps, paramKeepIRM, paramCreateBookmarks, paramDocStructureTags, paramBitmapMissingFonts, paramUseISO19005_1, missing);

            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// 替換文字
        /// </summary>
        /// <param name="pDt">資料來源(Datatable)</param>
        public void ReplaceWordTxtFromDatatable(System.Data.DataRow pDr)
        {
            //try
            //{
            for (int i = 0; i <= pDr.Table.Columns.Count - 1; i++)
            {
                String findText = pDr.Table.Columns[i].ColumnName.ToString();
                String replaceWithText = pDr[findText].ToString();

                //文字替換(內文文字)
                Replace("@" + findText + "@", replaceWithText);

                //文字替換(頁首文字)
                FindReplaceHeaderTxt("@" + findText + "@", replaceWithText);

                //文字替換(文字方塊)
                FindReplaceTextFrame("@" + findText + "@", replaceWithText);
            }
            //}
            //catch (Exception)
            //{
            //}
        }

        /// <summary>
        /// 替換word中的文字
        /// </summary>
        /// <param name=”strOld”>查詢的文字</param>
        /// <param name=”strNew”>替換的文字</param>
        public void Replace(Object findText, Object replaceText)
        {
            try
            {
                string[] replacetxt_ary = replaceText.ToString().Split(new string[] { "|||" }, StringSplitOptions.None);
                if (replacetxt_ary.Length == 1 || (replacetxt_ary.Length > 1 && replacetxt_ary[1] == "Color"))
                {
                    //options
                    object matchCase = false;
                    object matchWholeWord = true;
                    object matchWildCards = false;
                    object matchSoundsLike = false;
                    object matchAllWordForms = false;
                    object forward = true;
                    object format = true; //需開啟
                    object matchKashida = false;
                    object matchDiacritics = false;
                    object matchAlefHamza = false;
                    object matchControl = false;
                    object read_only = false;
                    object visible = true;
                    object replace = WdReplace.wdReplaceAll;
                    object wrap = WdFindWrap.wdFindContinue;

                    //替換全域性Document
                    wordApp10.Selection.Find.ClearFormatting();
                    wordApp10.Selection.Find.Replacement.ClearFormatting();
                    wordApp10.Selection.Find.Text = findText.ToString();
                    replaceText = replacetxt_ary[0];
                    wordApp10.Selection.Find.Replacement.Text = replaceText.ToString();

                    //依據傳入參數是否帶格式|||Color 來判斷是否需修改字體顏色
                    //例如: 疑似三倍體|||Color|||RED，代表疑似三倍體文字要用紅色顯示
                    if (replacetxt_ary.Length > 2)
                    {
                        if (replacetxt_ary[1] == "Color")
                        {
                            switch (replacetxt_ary[2])
                            {
                                case "RED":
                                    wordApp10.Selection.Find.Replacement.Font.ColorIndex = WdColorIndex.wdRed;
                                    break;

                                case "BLUE":
                                    wordApp10.Selection.Find.Replacement.Font.ColorIndex = WdColorIndex.wdBlue;
                                    break;

                                case "LIGHTBLUE":
                                    wordApp10.Selection.Find.Replacement.Font.Color = (WdColor)((192 << 16) + (112 << 8) + 0); // 使用 RGB 值 (R=0, G=112, B=192)
                                    break;
                            }
                        }
                    }

                    wordApp10.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                            ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceText, ref replace,
                            ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
                }
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// 替換頁首文字
        /// </summary>
        /// <param name="findText"></param>
        /// <param name="replaceText"></param>
        private void FindReplaceHeaderTxt(object findText, object replaceText)
        {
            try
            {
                string[] replacetxt_ary = replaceText.ToString().Split(new string[] { "|||" }, StringSplitOptions.None);
                if (replacetxt_ary.Length > 1 && replacetxt_ary[1] == "Header")
                {
                    replaceText = replacetxt_ary[0];
                    object m = System.Type.Missing;
                    wordApp10.ActiveDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Find.Execute(
                       ref findText,
                       ref m, ref m, ref m, ref m, ref m, ref m, ref m, ref m,
                       ref replaceText,
                       ref m, ref m, ref m, ref m, ref m);
                }
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// 替換文字方塊
        /// </summary>
        /// <param name="findText"></param>
        /// <param name="replaceText"></param>
        private void FindReplaceTextFrame(string findText, string replaceText)
        {
            try
            {
                string[] replacetxt_ary = replaceText.ToString().Split(new string[] { "|||" }, StringSplitOptions.None);
                if (replacetxt_ary.Length > 1 && replacetxt_ary[1] == "TextFrame")
                {
                    Shapes shapes = aDoc.Shapes;
                    foreach (Shape shape in shapes)
                    {
                        if (shape.TextFrame.HasText == -1)
                        {
                            var initialText = shape.TextFrame.TextRange.Text;
                            var resultingText = initialText.Replace(findText, replacetxt_ary[0]);
                            shape.TextFrame.TextRange.Text = resultingText;
                        }
                    }
                }
            }
            catch (Exception)
            {
            }
        }
        /// <summary>
        /// 替換Header文字方塊中文字為圖片
        /// </summary>
        /// <param name="FindStr"></param>
        /// <param name="replacePic"></param>
        /// <param name="W"></param>
        /// <param name="H"></param>
        public void SearchReplacePicInHeaderAndTextBox(string FindStr, object replacePic, object W, object H)
        {
            //try
            //{
            // 取得當前文檔的所有節
            foreach (Microsoft.Office.Interop.Word.Section section in this.wordApp10.ActiveDocument.Sections)
            {
                // 取得節的主要頁首範圍
                Range headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                // 清除查找格式化
                headerRange.Find.ClearFormatting();

                // 設置查找條件並執行查找（替換Header中的特定字元）
                while (headerRange.Find.Execute(FindStr))
                {
                    // 保存找到的文字範圍
                    Range foundRange = headerRange.Duplicate;

                    // 插入圖片在文字範圍的後面
                    Range insertRange = foundRange.Duplicate;
                    insertRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    InlineShape inlineShape = insertRange.InlineShapes.AddPicture(
                        FileName: replacePic.ToString(),
                        LinkToFile: false,
                        SaveWithDocument: true
                    );

                    // 設定圖片寬高
                    inlineShape.Width = Convert.ToInt16(W);
                    inlineShape.Height = Convert.ToInt16(H);

                    // 僅清除文字內容，保留圖片
                    foundRange.Text = "";
                }

                // 遍歷Header中的所有Shape
                foreach (Shape shape in section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ShapeRange)
                {
                    // 確認這個Shape是文字方塊
                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoTextBox)
                    {
                        // 取得文字方塊中的文字範圍
                        Range textBoxRange = shape.TextFrame.TextRange;

                        // 清除查找格式化
                        textBoxRange.Find.ClearFormatting();

                        // 設置查找條件並執行查找
                        while (textBoxRange.Find.Execute(FindStr))
                        {
                            // 保存找到的文字範圍
                            Range foundTextBoxRange = textBoxRange.Duplicate;

                            // 插入圖片在文字範圍的後面
                            Range insertTextBoxRange = foundTextBoxRange.Duplicate;
                            insertTextBoxRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                            InlineShape inlineShapeInTextBox = insertTextBoxRange.InlineShapes.AddPicture(
                                FileName: replacePic.ToString(),
                                LinkToFile: false,
                                SaveWithDocument: true
                            );

                            // 設定圖片寬高
                            inlineShapeInTextBox.Width = Convert.ToInt16(W);
                            inlineShapeInTextBox.Height = Convert.ToInt16(H);

                            // 僅清除文字內容，保留圖片
                            foundTextBoxRange.Text = "";
                        }
                    }
                }
            }

            // 確保在頁首或頁尾模式時返回到主文件
            if (wordApp10.ActiveWindow.View.Type == WdViewType.wdPrintView)
            {
                wordApp10.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekMainDocument;
            }
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("發生錯誤: " + ex.Message);
            //}
        }
        /// <summary>
        /// 替換頁首文字和頁首文字方塊中的特定字元為文字
        /// </summary>
        /// <param name="findText"></param>
        /// <param name="replaceText"></param>
        private void FindReplaceHeaderTxtAndTextBox(object findText, object replaceText)
        {
            try
            {

                object m = System.Type.Missing;

                // 替換頁首範圍中的特定字元
                wordApp10.ActiveDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Find.Execute(
                   ref findText,
                   ref m, ref m, ref m, ref m, ref m, ref m, ref m, ref m,
                   ref replaceText,
                   ref m, ref m, ref m, ref m, ref m);

                // 遍歷Header中的所有Shape (文字方塊)
                foreach (Shape shape in wordApp10.ActiveDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ShapeRange)
                {
                    // 判斷是否為文字方塊
                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoTextBox)
                    {
                        // 取得文字方塊中的文字範圍
                        Range textBoxRange = shape.TextFrame.TextRange;

                        // 清除查找格式化
                        textBoxRange.Find.ClearFormatting();

                        // 設定查找條件並執行查找
                        bool foundInTextBox = textBoxRange.Find.Execute(
                            ref findText,
                            ref m, ref m, ref m, ref m, ref m, ref m, ref m, ref m,
                            ref replaceText,
                            ref m, ref m, ref m, ref m, ref m);

                        if (foundInTextBox)
                        {
                            textBoxRange.Text = replaceText.ToString(); // 將找到的文字替換為新文字
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("發生錯誤: " + ex.Message);
            }
        }
        /// <summary>
        /// 新增圖片
        /// </summary>
        /// <param name="pDr">資料來源</param>
        public void AddImage(System.Data.DataRow pDr)
        {
            //try
            //{
            for (int i = 0; i <= pDr.Table.Columns.Count - 1; i++)
            {
                String findText = pDr.Table.Columns[i].ColumnName.ToString().Trim();
                String replaceWithText = pDr[findText].ToString().Trim();

                string[] replacetxt_ary = replaceWithText.ToString().Split(new string[] { "|||" }, StringSplitOptions.None);
                if (replacetxt_ary.Length > 3)
                {
                    if (replacetxt_ary[1] == "Image")
                    {
                        string url = ConfigurationManager.AppSettings["LISImageUrl"];

                        string image_path = string.Format("{0}/{1}", url, replacetxt_ary[0]);

                        if (image_path.Length > 0)
                        {
                            float Width = Convert.ToInt16(replacetxt_ary[2]);
                            float Height = Convert.ToInt16(replacetxt_ary[3]);
                            var sel = wordApp10.Selection;

                            SearchReplacePic("@" + findText + "@", image_path, Width, Height);
                        }
                    }
                    if (replacetxt_ary[1] == "Pictures")
                    {
                        string url = ConfigurationManager.AppSettings["LISPicturesUrl"];

                        string image_path = string.Format("{0}/{1}", url, replacetxt_ary[0]);

                        if (image_path.Length > 0)
                        {
                            float Width = Convert.ToInt16(replacetxt_ary[2]);
                            float Height = Convert.ToInt16(replacetxt_ary[3]);
                            var sel = wordApp10.Selection;

                            SearchReplacePic("@" + findText + "@", image_path, Width, Height);
                        }
                    }
                }
                else if (replacetxt_ary.Length == 2 && replacetxt_ary[1] == "Image")
                {
                    string url = ConfigurationManager.AppSettings["LISImageUrl"];
                    string Path = ConfigurationManager.AppSettings["GGALISImageFilePath"];
                    Path = string.Format("{0}{1}", Path, replacetxt_ary[0]);
                    string image_path = string.Format("{0}/{1}", url, replacetxt_ary[0]);
                    image_path = image_path.Replace(@"\", "/");
                    image_path = image_path.Replace(@"//", "/");

                    if (File.Exists(Path))
                    {
                        using (var img = Image.FromFile(Path))
                        {
                            // 獲取圖片的寬度和高度
                            float width = img.Width;
                            float height = img.Height;
                            var sel = wordApp10.Selection;

                            SearchReplacePic("@" + findText + "@", Path, width, height);
                            //SearchReplaceTextFramePic("@" + findText + "@", Path);
                        }
                    }
                    else //沒找到圖就直接替換為""
                    {
                        //文字替換(內文文字)
                        Replace("@" + findText + "@", "");
                        //文字替換(頁首文字)
                        FindReplaceHeaderTxt("@" + findText + "@", "");
                        //文字替換(文字方塊)
                        FindReplaceTextFrame("@" + findText + "@", "");
                    }
                }
                else if (replacetxt_ary.Length == 2 && replacetxt_ary[1] == "Pictures")
                {
                    string url = ConfigurationManager.AppSettings["LISPicturesUrl"];
                    string Path = ConfigurationManager.AppSettings["GGALISPicturesFilePath"];
                    Path = string.Format("{0}{1}", Path, replacetxt_ary[0]);
                    string image_path = string.Format("{0}/{1}", url, replacetxt_ary[0]);
                    image_path = image_path.Replace(@"\", "/");
                    image_path = image_path.Replace(@"//", "/");

                    if (File.Exists(Path))
                    {
                        using (var img = Image.FromFile(Path))
                        {
                            // 獲取圖片的寬度和高度
                            float width = img.Width;
                            float height = img.Height;
                            var sel = wordApp10.Selection;

                            SearchReplacePic("@" + findText + "@", Path, width, height);
                            //SearchReplaceTextFramePic("@" + findText + "@", Path);
                        }
                    }
                    else //沒找到圖就直接替換為""
                    {
                        //文字替換(內文文字)
                        Replace("@" + findText + "@", "");
                        //文字替換(頁首文字)
                        FindReplaceHeaderTxt("@" + findText + "@", "");
                        //文字替換(文字方塊)
                        FindReplaceTextFrame("@" + findText + "@", "");
                    }
                }
                else if (replacetxt_ary.Length == 2 && replacetxt_ary[1] == "H_IMG")
                {
                    string url = ConfigurationManager.AppSettings["LISImageUrl"];
                    string Path = ConfigurationManager.AppSettings["GGALISImageFilePath"];
                    Path = string.Format("{0}{1}", Path, replacetxt_ary[0]);
                    string image_path = string.Format("{0}/{1}", url, replacetxt_ary[0]);
                    image_path = image_path.Replace(@"\", "/");
                    image_path = image_path.Replace(@"//", "/");

                    if (File.Exists(Path))
                    {
                        using (var img = System.Drawing.Image.FromFile(Path))
                        {
                            // 獲取圖片的寬度和高度
                            float width = img.Width;
                            float height = img.Height;
                            var sel = wordApp10.Selection;

                            SearchReplacePicInHeader("@" + findText + "@", Path, width, height);
                            //SearchReplaceTextFramePic("@" + findText + "@", Path);
                        }
                    }
                    else //沒找到圖就直接替換為""
                    {
                        //文字替換(內文文字)
                        Replace("@" + findText + "@", "");
                        //文字替換(頁首文字)
                        FindReplaceHeaderTxt("@" + findText + "@", "");
                        //文字替換(文字方塊)
                        FindReplaceTextFrame("@" + findText + "@", "");
                    }
                }
                else if (replacetxt_ary.Length == 2 && replacetxt_ary[1] == "H_IMG_TF")//替換Header中文字方塊內文字為圖片
                {
                    string Path = ConfigurationManager.AppSettings["GGALISImageFilePath"];
                    Path = string.Format("{0}{1}", Path, replacetxt_ary[0]);



                    if (!string.IsNullOrEmpty(replacetxt_ary[0]) && File.Exists(Path))
                    {
                        using (var img = System.Drawing.Image.FromFile(Path))
                        {
                            // 獲取圖片 的寬度和高度
                            float width = img.Width;
                            float height = img.Height;

                            // 設置限制的寬度和高度（6cm 和 3cm，轉換為點）
                            const float maxWidthCm = 6.0f;
                            const float maxHeightCm = 3.0f;
                            const float cmToPoints = 28.35f; // 1cm ≈ 28.35 points
                            float maxWidth = maxWidthCm * cmToPoints;   // 最大寬度（點）
                            float maxHeight = maxHeightCm * cmToPoints; // 最大高度（點）

                            // 計算圖片的等比例縮放
                            float aspectRatio = width / height; // 寬高比

                            // 計算等比縮小後的寬度和高度
                            if (width > maxWidth || height > maxHeight)
                            {
                                if (width / maxWidth > height / maxHeight)
                                {
                                    // 如果寬度超出比例更多，按寬度縮放
                                    width = maxWidth;
                                    height = maxWidth / aspectRatio;
                                }
                                else
                                {
                                    // 如果高度超出比例更多，按高度縮放
                                    height = maxHeight;
                                    width = maxHeight * aspectRatio;
                                }
                            }

                            var sel = wordApp10.Selection;

                            //SearchReplacePicInHeaderAndTextBox("@" + findText + "@", Path, 60, 60);
                            SearchReplacePicInHeaderAndTextBox("@" + findText + "@", Path, width, height);
                        }


                    }
                    else
                    {
                        //文字替換(內文文字)
                        Replace("@" + findText + "@", "");
                        //文字替換(頁首文字)
                        FindReplaceHeaderTxt("@" + findText + "@", "");
                        //文字替換(文字方塊)
                        FindReplaceTextFrame("@" + findText + "@", "");
                        //文字替換(頁首文字方塊)
                        FindReplaceHeaderTxtAndTextBox("@" + findText + "@", "");
                    }

                }
            }
            //}
            //catch (Exception ex)
            //{
            //    InsertLog("", ex.Message);
            //}
        }

        //替換頁首文字為圖片
        /// <param name="FindStr"></param>
        /// <param name="replacePic"></param>
        /// <param name="W"></param>
        /// <param name="H"></param>
        public void SearchReplacePicInHeader(string FindStr, object replacePic, object W, object H)
        {
            try
            {
                // 取得當前文檔的所有節
                foreach (Section section in this.wordApp10.ActiveDocument.Sections)
                {
                    // 取得節的主要頁首範圍
                    Range headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                    // 清除查找格式化
                    headerRange.Find.ClearFormatting();

                    // 設置查找條件並執行查找
                    bool found = headerRange.Find.Execute(FindStr);

                    if (found)
                    {
                        headerRange.Select();

                        InlineShape inlineShape = this.wordApp10.Selection.InlineShapes.AddPicture(
                            FileName: replacePic.ToString(),
                            LinkToFile: false,
                            SaveWithDocument: true
                        );

                        inlineShape.Width = Convert.ToInt16(W);
                        inlineShape.Height = Convert.ToInt16(H);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("發生錯誤: " + ex.Message);
            }
        }

        /// <summary>
        ///指定圖檔取代搜尋到的文字 
        /// </summary>
        /// <param name="FindStr"></param>
        /// <param name="replacePic"></param>
        /// <param name="W"></param>
        /// <param name="H"></param>
        public void SearchReplacePic(string FindStr, object replacePic, object W, object H)
        {
            try
            {
                GotoTheBegining();
                this.wordApp10.Selection.Find.ClearFormatting();
                if ((this.wordApp10.Selection.Find.Execute(FindStr) == true))
                {
                    this.wordApp10.Selection.Select();


                    object linkToFile = true;
                    object saveWithDocument = true;

                    InlineShape Inlineshape = this.wordApp10.Selection.InlineShapes.AddPicture(
                                             FileName: replacePic.ToString(),
                                             LinkToFile: false,
                                             SaveWithDocument: true
                                           );

                    Inlineshape.Width = Convert.ToInt16(W);
                    Inlineshape.Height = Convert.ToInt16(H);
                }
            }
            catch (Exception)
            {
            }
        }


        public void SearchReplaceTextFramePic(string findStr, string replacePicPath)
        {
            try
            {
                foreach (Shape shape in aDoc.Shapes)
                {
                    if (shape.TextFrame.HasText == -1 && shape.TextFrame.TextRange.Text.Contains(findStr))
                    {
                        //if (shape.TextFrame.HasText == 1 && shape.TextFrame.TextRange.Text.Contains(findStr))
                        shape.TextFrame.TextRange.Text = "";
                        if (File.Exists(replacePicPath))
                        {
                            InlineShape inlineShape = shape.TextFrame.TextRange.InlineShapes.AddPicture(
                                                       FileName: replacePicPath,
                                                       LinkToFile: false,
                                                       SaveWithDocument: true
                                                   );
                        }
                        else
                        {
                            shape.TextFrame.TextRange.Text = shape.TextFrame.TextRange.Text.Replace(findStr, "");
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("發生錯誤: " + ex.Message);
            }
        }


        public void Save(string pfilePath, string pfileName)
        {
            try
            {
                CheckFolder(pfilePath);
                aDoc.SaveAs2(string.Format(@"{0}\{1}", pfilePath, pfileName));
            }
            catch (Exception ex)
            {
                InsertLog("", ex.Message);
            }
        }

        private void CheckFolder(string pPath)
        {
            bool exists = System.IO.Directory.Exists(pPath);

            if (!exists)
                System.IO.Directory.CreateDirectory(pPath);
        }

        public void Close()
        {
            //try
            //{
            if (aDoc != null)
            {
                aDoc.Close();
                Marshal.FinalReleaseComObject(aDoc);
            }
            if (wordApp10 != null)
            {
                //try
                //{
                object dontSave = WdSaveOptions.wdDoNotSaveChanges;
                ((_Application)wordApp10).Quit(ref dontSave);
                //}
                //finally
                //{
                Marshal.FinalReleaseComObject(wordApp10);
                //}
            }
            aDoc = null;
            wordApp10 = null;
            //GC.Collect();
            //}
            //catch (Exception ex)
            //{
            //    InsertLog("", ex.Message);
            //}
        }

        public void Dispose()
        {
            //確實關閉Word Application
            try
            {
                object dontSave = WdSaveOptions.wdDoNotSaveChanges;
                ((_Application)wordApp10).Quit(ref dontSave);
            }
            finally
            {
                Marshal.FinalReleaseComObject(wordApp10);
            }
        }

        public void InsertLog(string QueueID, string pMsg)
        {
            try
            {
                Console.WriteLine(pMsg);
                string LogPath = string.Format(@"{0}\ErrorLog", System.Windows.Forms.Application.StartupPath);
                CheckFolder(LogPath);
                string logFile = string.Format(@"{0}\{1}.txt", LogPath, DateTime.Now.ToString("yyyy-MM-dd"));
                using (StreamWriter sw = (File.Exists(logFile)) ? File.AppendText(logFile) : File.CreateText(logFile))
                {
                    sw.WriteLine("{0}  QueueID:{1}  錯誤:{2}", DateTime.Now, QueueID, pMsg);
                }
            }
            catch (Exception)
            {
            }
        }
    }
}

