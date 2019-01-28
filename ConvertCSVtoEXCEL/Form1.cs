using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ConvertCSVtoEXCEL
{
    public partial class Form1 : Form
    {
        // 読み込むCSVファイル
        string sourceFile;

        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 変換ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            toolStripStatusLabel1.Text = "処理中…";
            toolStripProgressBar1.Value = 0;

            // Progressクラスのインスタンスを生成
            var p = new Progress<int>(ShowProgress);

            // 実行時間の計測
            DateTime start = DateTime.Now;

            // 時間のかかる処理を別スレッドで開始
            string result = await Task.Run(() => DoWork(p));

            // 実行時間の表示
            DateTime end = DateTime.Now;
            var elapsed = end - start;
            string message = String.Format("実行時間: {0}", elapsed.ToString(@"hh\:mm\:ss"));
            Debug.WriteLine(message);
            textBox2.Text = message;

            // 処理結果の表示
            toolStripStatusLabel1.Text = result;
            toolStripProgressBar1.Value = 100;
            button1.Enabled = true;
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            // ドラッグされたファイル
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            sourceFile = files[0];

            var extension = Path.GetExtension(sourceFile);

            if (extension == ".csv" || extension == ".CSV")
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            // ドラッグ＆ドロップされたファイル
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            sourceFile = files[0];

            var extension = Path.GetExtension(sourceFile);

            if (extension == ".csv" || extension == ".CSV")
            {
                textBox1.Text = Path.GetFileName(sourceFile);
                textBox2.Text = string.Empty;
            }
        }

        // 進捗を表示するメソッド
        private void ShowProgress(int percent)
        {
            toolStripStatusLabel1.Text = percent + "％";
            toolStripProgressBar1.Value = percent;
        }

        // 時間のかかる処理を行うメソッド
        private string DoWork(IProgress<int> progress)
        {
            var allLines = File.ReadAllLines(sourceFile, Encoding.GetEncoding("shift-jis"));

            string directoryName = Path.GetDirectoryName(sourceFile);
            string fileName = Path.GetFileNameWithoutExtension(sourceFile) + ".xlsx";
            string filePath = Path.Combine(directoryName, fileName);

            IWorkbook book;

            // ブック作成
            book = new XSSFWorkbook();

            // シート無しのExcelファイルは保存はできるが、開くとエラーが発生する。
            book.CreateSheet("Sheet1");

            using (FileStream fs = File.Create(filePath))
            {
                book.Write(fs);
                fs.Close();
            }

            // シート設定
            ISheet sheet = book.GetSheetAt(0);
            sheet.AutoSizeColumn(0);

            // 文字列に書式変更
            var style = book.CreateCellStyle();
            style.DataFormat = book.CreateDataFormat().GetFormat("@");

            int n = allLines.Length;
            string oneRow;

            int row = 0;
            foreach (var oneLine in allLines)
            {
                oneRow = oneLine.Replace("\"", "");
                var items = oneRow.Split(',');

                try
                {
                    var dataRow = sheet.CreateRow(row);

                    int column = 0;
                    foreach (var item in items)
                    {
                        dataRow.CreateCell(column);
                        dataRow.Cells[column].SetCellValue(item);
                        dataRow.Cells[column].CellStyle = style;

                        // 幅の自動調整（時間がかかる）
                        //sheet.AutoSizeColumn(column);

                        column++;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }

                row++;

                Debug.WriteLine(String.Format("行: {0}/{1}", row, allLines.Length));

                // 進捗率
                int percentage = row * 100 / n;
                progress.Report(percentage);

                // 10000行ごとに保存
                //if (row % 10000 == 0)
                //{
                //    try
                //    {
                //        // Excelファイルを保存
                //        using (FileStream fs = File.Create(filePath))
                //        {
                //            book.Write(fs);
                //            fs.Close();
                //        }

                //        sheet = book.GetSheetAt(0);
                //    }
                //    catch (Exception ex)
                //    {
                //        Debug.WriteLine(ex);
                //    }
                //}

                // 20000行で終了
                //if (row % 20000 == 0)
                //{
                //    break;
                //}
            }

            // Excelファイルを保存
            using (FileStream fs = File.Create(filePath))
            {
                book.Write(fs);
                fs.Close();
            }

            // このメソッドからの戻り値
            return "完了";
        }
    }
}
