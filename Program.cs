using TsvData2Excel.src;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace TsvData2Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("How to use：");
                Console.WriteLine("tsv2excel <tsv-file> <excel-file>");
                Console.WriteLine("<tsv-file> ... reading target tsvfile");
                Console.WriteLine("<excel-file> ... writing target excel file(.xlsx)");
                return;
            };

            string tsvFilePath = args[0];
            string xlsFilePath = args[1];
            foreach (string arg in args)
            {
                if (!File.Exists(arg))
                {
                    OutputError(String.Format("{0}が存在しません。", arg));
                    return;
                }
            }

            // 設定ファイルの読み込み
            Config config = ConfigSerializer.Serialize(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\config.json");

            // 読み込んだ情報を格納するDictionalyのリスト
            IList<Dictionary<string, string>> updateRows = ReadRows(tsvFilePath, config);

            // エクセルファイル(シート)の読み込み
            ExcelReadWriter excelReadWriter = new ExcelReadWriter();
            try
            {
                excelReadWriter.Open(xlsFilePath, config.TargetSheet);

                // 更新行があれば更新を行う。なければ新規行追加
                foreach (Dictionary<string, string> row in updateRows)
                {
                    // ID列の値を取得
                    string idStrVal = row[config.Identifier.Tsv];
                    // ID列の値と等しい行がシート内に存在するかどうかを確認
                    Excel.Range result = excelReadWriter.SearchByColumn(config.Identifier.Xlsx, idStrVal);
                    if (result == null)
                    {
                        excelReadWriter.SearchWritableRow(config.FilledColumn, config.EndOfColumn);
                        excelReadWriter.WriteNewRow(config.Mapping, row);
                    }
                    else
                    {
                        excelReadWriter.WriteRow(config.Mapping, row, result.Row);
                    }
                }
            }
            finally
            {
                excelReadWriter.Close();
            }

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

        private static IList<Dictionary<string, string>> ReadRows(string tsvFilePath, Config config)
        {
            // 読み込んだ情報を格納するDictionalyのリスト
            IList<Dictionary<string, string>> updateRows = new List<Dictionary<string, string>>();

            // tsvファイルの読み込み
            TextFieldParser parser = new TextFieldParser(tsvFilePath, Encoding.GetEncoding("UTF-8"));
            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters(config.SplitChar);

            if (!parser.EndOfData)
            {
                // header情報の読み込み
                string[] header = parser.ReadFields();
                // 行情報をヘッダと対応させたDictionaryを作成
                while (!parser.EndOfData)
                {
                    string[] row = parser.ReadFields();
                    Dictionary<string, string> rowDict = new Dictionary<string, string>();
                    for (int i = 0; i < header.Length; i++)
                    {
                        rowDict.Add(header[i], row[i]);
                    }
                    updateRows.Add(rowDict);
                }
            }

            return updateRows;
        }

        private static void OutputError(string v)
        {
            Console.WriteLine(String.Format("Error： {0}", v));
        }
    }
}
