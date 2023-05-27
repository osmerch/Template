using System;
using Microsoft.Win32;
using System.IO;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Text.Json;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using JsonProperty = Newtonsoft.Json.Serialization.JsonProperty;


namespace Template_4337
{
    /// <summary>
    /// Логика взаимодействия для Shalimova_4337.xaml
    /// </summary>
    public partial class Shalimova_4337 : Window
    {
        public Shalimova_4337()
        {
            InitializeComponent();
        }
        private void btn1_Click(object sender, RoutedEventArgs e)
        {
            dynamic xlApp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls; *xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };

            if (!(ofd.ShowDialog() == true))
            {
                return;
            }
            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();

            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            using (WorkDDbEntities db = new WorkDDbEntities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    db.Workers.Add(new Workers()
                    {
                        ID = list[i, 0],
                        Post = list[i, 1],
                        FIO = list[i,2],
                        Login = list[i, 3],
                        Password = list[i, 4],
                        LastEnter = DateTime.ParseExact(list[i, 5].ToString(), "dd:MM:yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture),
                        EnterType = list[i, 6]
                    }); ;
                }
                db.SaveChanges();
                MessageBox.Show("Данные добавлены");
                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit();
                GC.Collect();
            }
        }

        private void btn3_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "файл JSON (Spisok.json)|*.json",
                Title = "Выберите файл данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            using (StreamReader reader = new StreamReader(ofd.FileName))
            {
                var settings = new JsonSerializerSettings
                {
                    ContractResolver = ShouldDeserializeContractResolver.Instance
                };
                List<Workers> workers = JsonConvert.DeserializeObject<List<Workers>>(await reader.ReadToEndAsync(), settings);
                using (WorkDDbEntities db = new WorkDDbEntities())
                {
                    db.Workers.RemoveRange(db.Workers);
                    foreach (var c in workers)
                    {
                        db.Workers.Add(c);
                    }
                    db.SaveChanges();
                }
                MessageBox.Show("Объекты добавлены в БД");
            }
        }

        private void btn2_Click(object sender, RoutedEventArgs e)
        {
           List<Workers> allWorkers;
            List<EnterTypeTb> allTypes;
            using (WorkDDbEntities db = new WorkDDbEntities())
            {
                allWorkers = db.Workers.ToList();
                allTypes = db.EnterTypeTb.ToList();
            }

            var app = new Excel.Application();
            app.SheetsInNewWorkbook = allTypes.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            for (int i = 0; i < allTypes.Count(); i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = allTypes[i].EnterType;
                worksheet.Cells[1][startRowIndex + 1] = "ID";
                worksheet.Cells[2][startRowIndex + 1] = "Должность";
                worksheet.Cells[3][startRowIndex + 1] = "Логин";
                startRowIndex++;
                var categ = allWorkers.GroupBy(s => s.EnterType).ToList();
                foreach (var c in categ)
                {
                    if (c.Key == allTypes[i].EnterType)
                    {
                        Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][1]];
                        headerRange.Merge();
                        headerRange.Value = allTypes[i].EnterType;
                        headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        headerRange.Font.Italic = true;
                        startRowIndex++;

                        foreach (Workers c1 in allWorkers)
                        {
                            if (c1.EnterType == c.Key)
                            {
                                worksheet.Cells[1][startRowIndex] = c1.ID;
                                worksheet.Cells[2][startRowIndex] = c1.Post;
                                worksheet.Cells[3][startRowIndex] = c1.Login;
                                startRowIndex++;

                            }
                        }
                        worksheet.Cells[1][startRowIndex].Formula = $"=СЧЁТ(A3:A{startRowIndex - 1})";
                        worksheet.Cells[1][startRowIndex].Font.Bold = true;
                    }
                    else
                    {
                        continue;
                    }
                }
                Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[5][startRowIndex - 1]];
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                worksheet.Columns.AutoFit();
            }
            app.Visible = true;
        }

        private void btn4_Click(object sender, RoutedEventArgs e)
        {
            List<Workers> allWorkers;
            List<EnterTypeTb> allTypes;
            using (WorkDDbEntities db = new WorkDDbEntities())
            {
                allWorkers = db.Workers.ToList();
                allTypes = db.EnterTypeTb.ToList();
                var categ = allWorkers.GroupBy(s => s.EnterType).ToList();
                var app = new Word.Application();
                Word.Document doc = app.Documents.Add();
                foreach (var c in categ)
                {
                    Word.Paragraph paragraph = doc.Paragraphs.Add();
                    Word.Range range = paragraph.Range;
                    range.Text = allTypes.Where(g => g.EnterType == c.Key).FirstOrDefault().EnterType;
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();
                    Word.Paragraph tableParagraph = doc.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table WorkersTable = doc.Tables.Add(tableRange, c.Count() + 1, 5);
                    WorkersTable.Borders.InsideLineStyle = WorkersTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    WorkersTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    Word.Range cellRange;
                    cellRange = WorkersTable.Cell(1, 1).Range;
                    cellRange.Text = "ID";
                    cellRange = WorkersTable.Cell(1, 2).Range;
                    cellRange.Text = "Должность";
                    cellRange = WorkersTable.Cell(1, 3).Range;
                    cellRange.Text = "Логин";
                    WorkersTable.Rows[1].Range.Bold = 1;
                    WorkersTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    int i = 1;

                    foreach (var current in c)
                    {
                        cellRange = WorkersTable.Cell(i + 1, 1).Range;
                        cellRange.Text = current.ID.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = WorkersTable.Cell(i + 1, 2).Range;
                        cellRange.Text = current.Post;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = WorkersTable.Cell(i + 1, 3).Range;
                        cellRange.Text = current.Login.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        i++;
                    }
                    Word.Paragraph countWorkersParagraph = doc.Paragraphs.Add();
                    Word.Range CountWorkersRange = countWorkersParagraph.Range;
                    CountWorkersRange.Text = $"Количесвто работников с данным типом входа - {c.Count()}";
                    CountWorkersRange.Font.Color = Word.WdColor.wdColorDarkRed;
                    CountWorkersRange.InsertParagraphAfter();
                    doc.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                }
                app.Visible = true;
            }
        }
    }
}
