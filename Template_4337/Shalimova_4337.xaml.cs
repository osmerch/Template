using Microsoft.Win32;
using System;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;


namespace Template_4337
{
    /// <summary>
    /// Логика взаимодействия для Shalimova_4337.xaml
    /// </summary>
    /// 
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

            if(!(ofd.ShowDialog() == true)){
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
                for (int i = 1; i < _rows;  i++)
                {
                    db.Workers.Add(new Workers()
                    {
                        ID = list[i, 0],
                        Post = list[i, 1],
                        Login = list[i, 3],
                        EnterType = list[i, 4]
                    }); ;
                }
                db.SaveChanges();
                MessageBox.Show("Данные добавлены");
                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit();
                GC.Collect();
            }
        }

        private void btn2_Click(object sender, RoutedEventArgs e)
        {
            List<Workers> allWorkers;
            List<EnterTypeTb> allTypes;
            using(WorkDDbEntities db = new WorkDDbEntities())
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
                foreach(var c in categ)
                {
                    if(c.Key == allTypes[i].EnterType)
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
    }
}
