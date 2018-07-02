using TsvData2Excel.src;
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace TsvData2Excel
{
    class ExcelReadWriter : IDisposable
    {
        private Excel.Application mExcel = new Excel.Application
        {
            Visible = false
        };
        // ブック
        private Excel.Workbook mWorkbook;
        // シート
        private Excel.Worksheet mWorkSheet;

        // 新規で編集可能な行番号
        private int newRow = -1;

        /// <summary>
        /// コンストラクタ。
        /// </summary>
        public ExcelReadWriter(string filePath, string sheetName)
        {
            // ブックのオープン
            this.mWorkbook = this.mExcel.Workbooks.Open(System.IO.Path.GetFullPath(filePath));

            // シートの取得
            this.mWorkSheet = this.mWorkbook.Sheets[Int32.Parse(sheetName)];
        }

        /// <summary>
        /// 指定したセルに値を書き込みます。
        /// </summary>
        /// <param name="col">列レター</param>
        /// <param name="row">行番号</param>
        /// <param name="val">書き込む値（文字列）</param>
        public void WriteCell(string col, string row, string val)
        {
            this.mWorkSheet.Range[col + row].Value = val;
        }

        /// <summary>
        /// 指定したセルに値を書き込みます
        /// </summary>
        /// <param name="range">書き込み対象Range</param>
        /// <param name="val">書き込む値</param>
        public void WriteCell(Excel.Range range, string val)
        {
            range.Value = val;
        }

        /// <summary>
        /// 新規行に書き込みます。
        /// </summary>
        /// <param name="mappingList"></param>
        /// <param name="row"></param>
        public void WriteNewRow(IList<Mapping> mappingList, Dictionary<string, string> row)
        {
            foreach (Mapping mapping in mappingList)
            {
                // TSVの内容を連結
                string value = ConcatDictStr(row, mapping);
                // 対象の列すべてに対して更新をかける
                foreach (string key in mapping.Xlsx)
                {
                    WriteCell(key, this.newRow.ToString(), value);
                }
            }
            OutputLog("add Line：" + this.newRow);
        }

        /// <summary>
        /// 指定行を更新します。
        /// </summary>
        /// <param name="mappingList"></param>
        /// <param name="row"></param>
        /// <param name="rowNum"></param>
        public void WriteRow(IList<Mapping> mappingList, Dictionary<string, string> row, int rowNum)
        {
            foreach (Mapping mapping in mappingList)
            {
                // TSVの内容を連結
                string value = ConcatDictStr(row, mapping);
                // 対象の列すべてに対して更新をかける
                foreach (string key in mapping.Xlsx)
                {
                    WriteCell(key, rowNum.ToString(), value);
                }
            }
            OutputLog("update Line：" + rowNum);
        }

        /// <summary>
        /// 使用されている最後のセルを返します。
        /// </summary>
        /// <returns></returns>
        public Excel.Range GetLastCell()
        {
            return this.mWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
        }

        /// <summary>
        /// 指定した列全体から文字検索を行い、Rangeオブジェクトを返却します。
        /// </summary>
        /// <param name="column">検索対象の列</param>
        /// <param name="key">検索文字列</param>
        /// <returns>検索結果Rangeオブジェクト</returns>
        public Excel.Range SearchByColumn(string column, string key)
        {
            return this.mWorkSheet.Range[column + "1"].EntireColumn.Find(key);
        }

        /// <summary>
        /// 指定した列の値以外がすべて空になっている最初の行を返却します。
        /// </summary>
        /// <param name="filledColumn">指定列（配列）</param>
        /// <param name="endOfColumn">検索する列番号の最大値</param>
        /// <returns>検索結果Rangeオブジェクト</returns>
        public void SearchWritableRow(IList<string> filledColumn, string endOfColumn)
        {
            foreach (Excel.Range row in mWorkSheet.UsedRange.Rows)
            {
                bool flag = false;
                foreach (Excel.Range cell in mWorkSheet.Range["A" + row.Row, endOfColumn + row.Row])
                {
                    if (filledColumn.Contains(GetColumnLetter(cell.Column)))
                    {
                        continue;
                    }
                    if (cell.Value2 != null && cell.Value2.ToString() != "")
                    {
                        flag = true;
                        break;
                    }
                }
                if (!flag)
                {
                    this.newRow = row.Row;
                    break;
                }
            }
        }

        /// <summary>
        /// エクセルファイルを保存し、アプリケーションのオブジェクトを開放します。
        /// </summary>
        public void Dispose()
        {
            //開いているファイルのバックアップ保存
            DateTime dt = DateTime.Now;
            try
            {
                this.mWorkbook.SaveAs(this.mWorkbook.Path + "\backup" + dt.ToString("yyyyMMddHHmmss") + ".xlsx.bak");
                OutputLog("Save a backup file to" + this.mWorkbook.Path + "\backup" + dt.ToString("yyyyMMddHHmmss") + ".xlsx.bak");
            }
            catch (System.Runtime.InteropServices.ExternalException e)
            {
                this.mWorkbook.SaveAs(dt.ToString("yyyyMMddHHmmss") + ".xlsx.bak");
                OutputLog("Save a backup file to Document Folder");
            }


            // 開いているファイルの上書き保存
            this.mWorkbook.Save();
            this.mWorkbook.Close();
            this.mExcel.DisplayAlerts = false;
            this.mExcel.Quit();

            // オブジェクトの解放
            System.Runtime.InteropServices.Marshal.ReleaseComObject(this.mWorkSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(this.mWorkbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(this.mExcel);
        }

        private static string GetColumnLetter(int column)
        {
            if (column < 1) return String.Empty;
            return GetColumnLetter((column - 1) / 26) + (char)('A' + (column - 1) % 26);
        }

        /// <summary>
        /// DictionaryのStringをMappingに基づいて連結します
        /// </summary>
        /// <param name="row"></param>
        /// <param name="mapping"></param>
        /// <returns></returns>
        private static string ConcatDictStr(Dictionary<string, string> row, Mapping mapping)
        {
            string value = "";
            foreach (string key in mapping.Tsv)
            {
                if (value.Length != 0)
                {
                    value += mapping.SplitChar;
                }
                value += row[key];
            }

            return value;
        }

        private static void OutputLog(string v)
        {
            Console.WriteLine(String.Format("Log： {0}", v));
        }
    }
}
