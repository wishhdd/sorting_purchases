using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Сортировки_закупки
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }
        object[,] srcArr_ob;//1 - строка, 2 - столбец
        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            progressBar1.Value = 0;
            Excel.Application APExcel = new Microsoft.Office.Interop.Excel.Application();
            APExcel.DisplayAlerts = false;
            APExcel.Visible = true;
            APExcel.Workbooks.Open(openFileDialog1.FileName, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);
            Excel.Worksheet Excelsheet_ob = (Excel.Worksheet)APExcel.Sheets[1]; //определяем рабочий лист
            srcArr_ob = (object[,])Excelsheet_ob.UsedRange.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);//забираем весь лист
            APExcel.Workbooks.Close();//закрываем книгу
            APExcel.Quit();
            progressBar1.Value = 100;
            textBox1.Text = "Забрал данные...";
            if (dateTimePicker1.Value.Date >= dateTimePicker2.Value)
            {
                MessageBox.Show("Дата начала должна быть меньше или равна дате конца");
            }
            else
            {

                //MessageBox.Show(srcArr_ob.GetUpperBound(0).ToString());

                var obekti = new List<string>();
                var FPSS = new List<string>();

                // ============== будем бегать два раза

                #region // ============== щас бежим по цеху

                obekti.Clear();
                FPSS.Clear();
                dataGridView1.Rows.Clear();
                dataGridView1.Columns.Clear();
                dataGridView1.Columns.Add("", "");
                dataGridView1.Columns.Add("", "");
                dataGridView1.Columns.Add("", "");
                dataGridView1.Columns.Add("", "");
                textBox1.Text = "Собираю выборку для цеха...";
                progressBar1.Value = 2;
                progressBar1.Maximum = srcArr_ob.GetUpperBound(0) * 2 + 1;
                dataGridView1.Rows.Add(2);
                dataGridView1[0, 0].Value = "Поставщик";
                dataGridView1[1, 0].Value = "Профиль";
                dataGridView1[2, 0].Value = "Сечение";
                dataGridView1[3, 0].Value = "Сталь";

                for (int i = 2; i <= srcArr_ob.GetUpperBound(0); i++)
                {
                    //MessageBox.Show(Convert.ToDateTime(srcArr_ob[i, 1]).Date.ToString() + srcArr_ob[i, 5].ToString());
                    if (Convert.ToDateTime(srcArr_ob[i, 1]).Date >= dateTimePicker1.Value.Date && Convert.ToDateTime(srcArr_ob[i, 1]).Date <= dateTimePicker2.Value.Date)
                    {
                        if (Convert.ToString(srcArr_ob[i, 2]) == "В цех")//проверяем назначение
                        {
                            if (!obekti.Contains(Convert.ToString(srcArr_ob[i, 3])) && Convert.ToString(srcArr_ob[i, 3]) != "")//Собрал объекты
                            {
                                obekti.Add(Convert.ToString(srcArr_ob[i, 3]));
                            }

                            if (!FPSS.Contains(Convert.ToString(srcArr_ob[i, 8]) + "\t" + Convert.ToString(srcArr_ob[i, 4]) + "\t" + Convert.ToString(srcArr_ob[i, 5]) + "\t" + Convert.ToString(srcArr_ob[i, 6])) && (Convert.ToString(srcArr_ob[i, 8]) + Convert.ToString(srcArr_ob[i, 4]) + Convert.ToString(srcArr_ob[i, 5]) + Convert.ToString(srcArr_ob[i, 6])) != "")//Собрал профили, сечения, фирмы, стали
                            {
                                FPSS.Add(Convert.ToString(srcArr_ob[i, 8]) + "\t" + Convert.ToString(srcArr_ob[i, 4]) + "\t" + Convert.ToString(srcArr_ob[i, 5]) + "\t" + Convert.ToString(srcArr_ob[i, 6]));
                                // MessageBox.Show(Convert.ToString(srcArr_ob[i, 8]) + "\t" + Convert.ToString(srcArr_ob[i, 4]) + "\t" + Convert.ToString(srcArr_ob[i, 5]) + "\t" + Convert.ToString(srcArr_ob[i, 6]));
                            }
                        }
                    }
                    progressBar1.PerformStep();
                }
                obekti.Sort();
                FPSS.Sort();
                //MessageBox.Show(Convert.ToString(obekti.Count));
                for (int q = 0; q < obekti.Count; q++)//Выплюнули объекты
                {
                    //MessageBox.Show(q.ToString() + "/" + obekti.Count.ToString());
                    dataGridView1.Columns.Add("", "");
                    dataGridView1.Columns.Add("", "");
                    dataGridView1.Columns.Add("", "");
                    dataGridView1[dataGridView1.Columns.Count - 3, 0].Value = obekti[q];
                    dataGridView1[dataGridView1.Columns.Count - 2, 0].Value = obekti[q];
                    dataGridView1[dataGridView1.Columns.Count - 1, 0].Value = obekti[q];
                    dataGridView1[dataGridView1.Columns.Count - 3, 1].Value = "масса, кг";
                    dataGridView1[dataGridView1.Columns.Count - 2, 1].Value = "Средняя цена";
                    dataGridView1[dataGridView1.Columns.Count - 1, 1].Value = "Сумма";
                }
                //  string Firma = "";
                for (int q = 0; q < FPSS.Count; q++)//Выплюнули профили и фирмы
                {
                    //MessageBox.Show(q.ToString() + "/" + obekti.Count.ToString());
                    dataGridView1.Rows.Add();
                    string[] w = FPSS[q].Split('\t');
                    // if (Firma != w[0] && q > 0)
                    // {
                    //      dataGridView1.Rows.Add();
                    //      dataGridView1[0, dataGridView1.RowCount - 2].Value = "Итого:";
                    //      itogo_ps.Add(dataGridView1.RowCount - 2);
                    //
                    //    }
                    //    Firma = w[0];
                    dataGridView1[0, dataGridView1.RowCount - 1].Value = w[0];
                    dataGridView1[1, dataGridView1.RowCount - 1].Value = w[1];
                    dataGridView1[2, dataGridView1.RowCount - 1].Value = w[2];
                    dataGridView1[3, dataGridView1.RowCount - 1].Value = w[3];


                    //MessageBox.Show(w[0]);
                }
                dataGridView1.Rows.Add();
                dataGridView1[0, dataGridView1.RowCount - 1].Value = "Итого:";
                //itogo_ps.Add(dataGridView1.RowCount - 1);

                //Заполняем матрицу цифрами
                for (int i = 2; i <= srcArr_ob.GetUpperBound(0); i++)
                {
                    if (Convert.ToDateTime(srcArr_ob[i, 1]).Date >= dateTimePicker1.Value.Date && 
                        Convert.ToDateTime(srcArr_ob[i, 1]).Date <= dateTimePicker2.Value.Date)//проверяем совпадение по дате
                    {
                        for (int w = 2; w < dataGridView1.Rows.Count; w++)
                        {
                            if (Convert.ToString(dataGridView1[0, w].Value) + Convert.ToString(dataGridView1[1, w].Value) + Convert.ToString(dataGridView1[2, w].Value)
                                + Convert.ToString(dataGridView1[3, w].Value) == Convert.ToString(srcArr_ob[i, 8]) + Convert.ToString(srcArr_ob[i, 4])
                                    + Convert.ToString(srcArr_ob[i, 5]) + Convert.ToString(srcArr_ob[i, 6]))//проверяем совпадения профилей
                            {
                                for (int r = 4; r < dataGridView1.ColumnCount; r++)
                                {
                                    if (Convert.ToString(dataGridView1[r, 0].Value) == Convert.ToString(srcArr_ob[i, 3]))//MessageBox.Show("есть совпадения по продовцу профилю и марки стали");
                                    {
                                        dataGridView1[r, w].Value = Convert.ToDouble(dataGridView1[r, w].Value) + Convert.ToDouble(srcArr_ob[i, 7]);
                                        dataGridView1[r + 2, w].Value = Convert.ToDouble(dataGridView1[r + 2, w].Value) + Convert.ToDouble(srcArr_ob[i, 11]);
                                        //MessageBox.Show("");
                                        break;
                                    }
                                }
                            }
                        }
                    }

                    progressBar1.PerformStep();

                }

                //считаем среднюю цену металла и итого
                for (int i = 2; i < dataGridView1.Rows.Count; i++)
                {
                    for (int w = 5; w < dataGridView1.ColumnCount; w = w + 3)
                    {
                        dataGridView1[w, i].Value = Convert.ToDouble(dataGridView1[w + 1, i].Value) * 1000 / Convert.ToDouble(dataGridView1[w - 1, i].Value);
                        double num = Convert.ToDouble(dataGridView1[w, i].Value);
                        if (num != (Double)num) // Проверка то что число - это число.
                        {
                            dataGridView1[w, i].Value = 0;
                        }
                    }
                }
                //Собираем столбец "всего"
                dataGridView1.Columns.Add("", "");
                dataGridView1.Columns.Add("", "");
                dataGridView1.Columns.Add("", "");
                dataGridView1[dataGridView1.ColumnCount - 3, 0].Value = "Всего";
                dataGridView1[dataGridView1.ColumnCount - 2, 0].Value = "Всего";
                dataGridView1[dataGridView1.ColumnCount - 1, 0].Value = "Всего";
                dataGridView1[dataGridView1.Columns.Count - 3, 1].Value = "масса, кг";
                dataGridView1[dataGridView1.Columns.Count - 2, 1].Value = "Средняя цена";
                dataGridView1[dataGridView1.Columns.Count - 1, 1].Value = "Сумма";

                for (int i = 2; i < dataGridView1.RowCount - 1; i++)
                {
                    for (int x = 4; x < dataGridView1.ColumnCount - 3; x++)
                    {
                        if (dataGridView1[x, 1].Value.ToString() == "масса, кг")
                        {
                            dataGridView1[dataGridView1.Columns.Count - 3, i].Value = Convert.ToDouble(dataGridView1[dataGridView1.Columns.Count - 3, i].Value) + Convert.ToDouble(dataGridView1[x, i].Value);
                        }
                        if (dataGridView1[x, 1].Value.ToString() == "Сумма")
                        {
                            dataGridView1[dataGridView1.Columns.Count - 1, i].Value = Convert.ToDouble(dataGridView1[dataGridView1.Columns.Count - 1, i].Value) + Convert.ToDouble(dataGridView1[x, i].Value);
                        }
                        dataGridView1[dataGridView1.Columns.Count - 2, i].Value = Convert.ToDouble(dataGridView1[dataGridView1.Columns.Count - 1, i].Value) * 1000 / Convert.ToDouble(dataGridView1[dataGridView1.Columns.Count - 3, i].Value);
                    }

                }

                for (int i = 4; i < dataGridView1.ColumnCount; i++)
                {
                    dataGridView1.Update();
                    for (int x = 2; x < dataGridView1.RowCount - 1; x++)
                    {
                        if (dataGridView1[i, 1].Value.ToString() == "масса, кг")
                        {
                            dataGridView1[i, dataGridView1.RowCount - 1].Value = Convert.ToDouble(dataGridView1[i, dataGridView1.RowCount - 1].Value) + Convert.ToDouble(dataGridView1[i, x].Value);
                        }
                        if (dataGridView1[i, 1].Value.ToString() == "Сумма")
                        {
                            dataGridView1[i, dataGridView1.RowCount - 1].Value = Convert.ToDouble(dataGridView1[i, dataGridView1.RowCount - 1].Value) + Convert.ToDouble(dataGridView1[i, x].Value);
                        }
                    }

                }
                for (int i = 4; i < dataGridView1.ColumnCount; i++)
                {
                    if (dataGridView1[i, 1].Value.ToString() == "Средняя цена")
                    {
                        // MessageBox.Show(dataGridView1[i + 1, dataGridView1.RowCount - 1].Value.ToString() + "  " + dataGridView1[i - 1, dataGridView1.RowCount - 1].Value.ToString());
                        dataGridView1[i, dataGridView1.RowCount - 1].Value = Convert.ToDouble(dataGridView1[i + 1, dataGridView1.RowCount - 1].Value) * 1000 / Convert.ToDouble(dataGridView1[i - 1, dataGridView1.RowCount - 1].Value);
                    }
                }
                if (double.IsNaN(Convert.ToDouble(dataGridView1[dataGridView1.ColumnCount - 2, dataGridView1.RowCount - 1].Value))) // Проверка итоговая цена - это число
                {
                    dataGridView1[dataGridView1.ColumnCount - 2, dataGridView1.RowCount - 1].Value = "какая то ирунда";
                    //MessageBox.Show("!!!");
                }


                dateTimePicker1.Enabled = false;
                dateTimePicker2.Enabled = false;
                progressBar1.PerformStep();
                

                //пробуем сделать всё на одну кнопку :)
                
                progressBar1.Value = progressBar1.Maximum * 3 / 4;
                textBox1.Text = "Ждем экселя раз...";
                //MessageBox.Show("");
                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
                //Книга.
                ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
                //Таблица.
                ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        ExcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                    }
                }
                //форматируем таблицу
                string[] vsS ={ "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", 
"M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", 
"AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW",
"AX", "AY", "AZ" , 
"BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW",
"BX", "BY", "BZ",
"CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW",
"CX", "CY", "CZ",
"DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW",
"DX", "DY", "DZ",
"EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW",
"EX", "EY", "EZ",
"FA", "FB", "FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN", "FO", "FP", "FQ", "FR", "FS", "FT", "FU", "FV", "FW",
"FX", "FY", "FZ"};
                ExcelWorkSheet.get_Range("A1", "A1").Value = "Закупка металлопроката  в цех для произвдства металлоконструкций за перюд с " + dateTimePicker1.Value.Date.ToString().Replace(" 0:00:00", "") + " по " + dateTimePicker2.Value.Date.ToString().Replace(" 0:00:00", "");
                ExcelWorkSheet.get_Range(vsS[0] + "1", vsS[dataGridView1.ColumnCount - 1] + "1").MergeCells = true;
                ExcelWorkSheet.get_Range(vsS[0] + "1", vsS[dataGridView1.ColumnCount - 1] + "1").Font.Bold = true;
                ExcelWorkSheet.get_Range(vsS[0] + "1", vsS[dataGridView1.ColumnCount - 1] + "1").Font.Size = 18;
                ExcelWorkSheet.get_Range(vsS[0] + "1", vsS[dataGridView1.ColumnCount - 1] + "1").HorizontalAlignment = Excel.Constants.xlCenter;
                ExcelWorkSheet.get_Range(vsS[0] + "1", vsS[dataGridView1.ColumnCount - 1] + "1").VerticalAlignment = Excel.Constants.xlCenter;
                ExcelWorkSheet.get_Range("A2", "A3").MergeCells = true;
                ExcelWorkSheet.get_Range("A2", "A3").Font.Size = 14;
                ExcelWorkSheet.get_Range("A2", "A3").HorizontalAlignment = Excel.Constants.xlCenter;
                ExcelWorkSheet.get_Range("A2", "A3").VerticalAlignment = Excel.Constants.xlCenter;

                ExcelWorkSheet.get_Range("B2", "B3").MergeCells = true;
                ExcelWorkSheet.get_Range("B2", "B3").Font.Size = 14;
                ExcelWorkSheet.get_Range("B2", "B3").HorizontalAlignment = Excel.Constants.xlCenter;
                ExcelWorkSheet.get_Range("B2", "B3").VerticalAlignment = Excel.Constants.xlCenter;

                ExcelWorkSheet.get_Range("C2", "C3").MergeCells = true;
                ExcelWorkSheet.get_Range("C2", "C3").Font.Size = 14;
                ExcelWorkSheet.get_Range("C2", "C3").HorizontalAlignment = Excel.Constants.xlCenter;
                ExcelWorkSheet.get_Range("C2", "C3").VerticalAlignment = Excel.Constants.xlCenter;

                ExcelWorkSheet.get_Range("D2", "D3").MergeCells = true;
                ExcelWorkSheet.get_Range("D2", "D3").Font.Size = 14;
                ExcelWorkSheet.get_Range("D2", "D3").HorizontalAlignment = Excel.Constants.xlCenter;
                ExcelWorkSheet.get_Range("D2", "D3").VerticalAlignment = Excel.Constants.xlCenter;

                ExcelWorkSheet.get_Range("A" + (dataGridView1.RowCount + 1).ToString(), "D" + (dataGridView1.RowCount + 1).ToString()).MergeCells = true;
                ExcelWorkSheet.get_Range("A" + (dataGridView1.RowCount + 1).ToString(), "D" + (dataGridView1.RowCount + 1).ToString()).HorizontalAlignment = Excel.Constants.xlRight;
                ExcelWorkSheet.get_Range("A" + (dataGridView1.RowCount + 1).ToString(), "D" + (dataGridView1.RowCount + 1).ToString()).VerticalAlignment = Excel.Constants.xlCenter;

                ExcelWorkSheet.Rows[(dataGridView1.RowCount + 1).ToString(), Type.Missing].Font.Size = 14;
                ExcelWorkSheet.Rows[(dataGridView1.RowCount + 1).ToString(), Type.Missing].Font.Bold = true;

                ExcelWorkSheet.Rows["3", Type.Missing].HorizontalAlignment = Excel.Constants.xlCenter;
                ExcelWorkSheet.Rows["3", Type.Missing].VerticalAlignment = Excel.Constants.xlCenter;

                ExcelWorkSheet.Rows["2", Type.Missing].RowHeight = "50";
                //ExcelWorkSheet.get_Range("A2", "A3").EntireColumn.AutoFit();

                //Устанавливаем типы для столбцов

                for (int j = 4; j < dataGridView1.ColumnCount; j = j + 3)
                {
                    ExcelWorkSheet.Columns[vsS[j], Type.Missing].NumberFormat = "# ###";
                    ExcelWorkSheet.Columns[vsS[j + 1], Type.Missing].Style = "Currency";
                    ExcelWorkSheet.Columns[vsS[j + 2], Type.Missing].Style = "Currency";
                }

                //ExcelWorkSheet.get_Range("D2", "D3").Style = Excel.Style("Currency");


                //Обьединяме объекты
                ExcelApp.DisplayAlerts = false;
                for (int j = 4; j < dataGridView1.ColumnCount; j = j + 3)
                {
                    ExcelWorkSheet.get_Range(vsS[j] + "2", vsS[j + 2] + "2").MergeCells = true;
                    ExcelWorkSheet.get_Range(vsS[j] + "2", vsS[j + 2] + "2").HorizontalAlignment = Excel.Constants.xlCenter;
                    ExcelWorkSheet.get_Range(vsS[j] + "2", vsS[j + 2] + "2").VerticalAlignment = Excel.Constants.xlCenter;
                    ExcelWorkSheet.get_Range(vsS[j] + "2", vsS[j + 2] + "2").WrapText = "true";
                    ExcelWorkSheet.get_Range(vsS[j] + "2", vsS[j + 2] + "2").Font.Size = 12;
                }
                //Устанавливаем автоширину
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    ExcelWorkSheet.Columns[vsS[j], Type.Missing].EntireColumn.AutoFit();
                }

                //Рисуем границы
                ExcelWorkSheet.get_Range("A1", vsS[dataGridView1.ColumnCount - 1].ToString() + (dataGridView1.RowCount + 1).ToString()).Borders[Excel.XlBordersIndex.xlInsideHorizontal]
                    .LineStyle = Excel.XlLineStyle.xlContinuous;
                ExcelWorkSheet.get_Range("A1", vsS[dataGridView1.ColumnCount - 1].ToString() + (dataGridView1.RowCount + 1).ToString()).Borders[Excel.XlBordersIndex.xlInsideVertical]
        .LineStyle = Excel.XlLineStyle.xlContinuous;
                ExcelWorkSheet.get_Range("A1", vsS[dataGridView1.ColumnCount - 1].ToString() + (dataGridView1.RowCount + 1).ToString()).Borders[Excel.XlBordersIndex.xlEdgeBottom]
    .LineStyle = Excel.XlLineStyle.xlContinuous;
                ExcelWorkSheet.get_Range("A1", vsS[dataGridView1.ColumnCount - 1].ToString() + (dataGridView1.RowCount + 1).ToString()).Borders[Excel.XlBordersIndex.xlEdgeTop]
    .LineStyle = Excel.XlLineStyle.xlContinuous;
                ExcelWorkSheet.get_Range("A1", vsS[dataGridView1.ColumnCount - 1].ToString() + (dataGridView1.RowCount + 1).ToString()).Borders[Excel.XlBordersIndex.xlEdgeLeft]
    .LineStyle = Excel.XlLineStyle.xlContinuous;
                ExcelWorkSheet.get_Range("A1", vsS[dataGridView1.ColumnCount - 1].ToString() + (dataGridView1.RowCount + 1).ToString()).Borders[Excel.XlBordersIndex.xlEdgeRight]
    .LineStyle = Excel.XlLineStyle.xlContinuous;

                //Делаем культурный вид
                ExcelApp.ActiveWindow.SplitRow = 3;
                ExcelApp.ActiveWindow.SplitColumn = 4;
                ExcelApp.ActiveWindow.FreezePanes = true;



                //Вызываем нашу созданную эксельку.
                //ExcelApp.DisplayAlerts = true;
                //ExcelApp.Visible = true;
                //ExcelApp.UserControl = true;
                progressBar1.Value = progressBar1.Maximum;
                textBox1.Text = "ГОТОВО!!!";

                #endregion
              

                #region // ============== щас бежим по транзиту

                obekti.Clear();
                FPSS.Clear();
                dataGridView1.Rows.Clear();
                dataGridView1.Columns.Clear();
                textBox1.Text = "Собираю выборку для транзита...";
                progressBar1.Value = 2;
                progressBar1.Maximum = srcArr_ob.GetUpperBound(0) * 2 + 1;
                dataGridView1.Columns.Add("", "");
                dataGridView1.Columns.Add("", "");
                dataGridView1.Columns.Add("", "");
                dataGridView1.Columns.Add("", "");
                dataGridView1.Rows.Add(2);
                dataGridView1[0, 0].Value = "Поставщик";
                dataGridView1[1, 0].Value = "Профиль";
                dataGridView1[2, 0].Value = "Сечение";
                dataGridView1[3, 0].Value = "Сталь";

                for (int i = 2; i <= srcArr_ob.GetUpperBound(0); i++)
                {
                    //MessageBox.Show(Convert.ToDateTime(srcArr_ob[i, 1]).Date.ToString() + srcArr_ob[i, 5].ToString());
                    if (Convert.ToDateTime(srcArr_ob[i, 1]).Date >= dateTimePicker1.Value.Date && Convert.ToDateTime(srcArr_ob[i, 1]).Date <= dateTimePicker2.Value.Date)
                    {
                        if (Convert.ToString(srcArr_ob[i, 2]) == "Транзит")//проверяем назначение
                        {
                            if (!obekti.Contains(Convert.ToString(srcArr_ob[i, 3])) && Convert.ToString(srcArr_ob[i, 3]) != "")//Собрал объекты
                            {
                                obekti.Add(Convert.ToString(srcArr_ob[i, 3]));
                            }

                            if (!FPSS.Contains(Convert.ToString(srcArr_ob[i, 8]) + "\t" + Convert.ToString(srcArr_ob[i, 4]) + "\t" + Convert.ToString(srcArr_ob[i, 5]) + "\t" + Convert.ToString(srcArr_ob[i, 6])) && (Convert.ToString(srcArr_ob[i, 8]) + Convert.ToString(srcArr_ob[i, 4]) + Convert.ToString(srcArr_ob[i, 5]) + Convert.ToString(srcArr_ob[i, 6])) != "")//Собрал профили, сечения, фирмы, стали
                            {
                                FPSS.Add(Convert.ToString(srcArr_ob[i, 8]) + "\t" + Convert.ToString(srcArr_ob[i, 4]) + "\t" + Convert.ToString(srcArr_ob[i, 5]) + "\t" + Convert.ToString(srcArr_ob[i, 6]));
                                // MessageBox.Show(Convert.ToString(srcArr_ob[i, 8]) + "\t" + Convert.ToString(srcArr_ob[i, 4]) + "\t" + Convert.ToString(srcArr_ob[i, 5]) + "\t" + Convert.ToString(srcArr_ob[i, 6]));
                            }
                        }
                    }
                    progressBar1.PerformStep();
                }
                obekti.Sort();
                FPSS.Sort();
                //MessageBox.Show(Convert.ToString(obekti.Count));
                for (int q = 0; q < obekti.Count; q++)//Выплюнули объекты
                {
                    //MessageBox.Show(q.ToString() + "/" + obekti.Count.ToString());
                    dataGridView1.Columns.Add("", "");
                    dataGridView1.Columns.Add("", "");
                    dataGridView1.Columns.Add("", "");
                    dataGridView1[dataGridView1.Columns.Count - 3, 0].Value = obekti[q];
                    dataGridView1[dataGridView1.Columns.Count - 2, 0].Value = obekti[q];
                    dataGridView1[dataGridView1.Columns.Count - 1, 0].Value = obekti[q];
                    dataGridView1[dataGridView1.Columns.Count - 3, 1].Value = "масса, кг";
                    dataGridView1[dataGridView1.Columns.Count - 2, 1].Value = "Средняя цена";
                    dataGridView1[dataGridView1.Columns.Count - 1, 1].Value = "Сумма";
                }
                //  string Firma = "";
                for (int q = 0; q < FPSS.Count; q++)//Выплюнули профили и фирмы
                {
                    //MessageBox.Show(q.ToString() + "/" + obekti.Count.ToString());
                    dataGridView1.Rows.Add();
                    string[] w = FPSS[q].Split('\t');
                    // if (Firma != w[0] && q > 0)
                    // {
                    //      dataGridView1.Rows.Add();
                    //      dataGridView1[0, dataGridView1.RowCount - 2].Value = "Итого:";
                    //      itogo_ps.Add(dataGridView1.RowCount - 2);
                    //
                    //    }
                    //    Firma = w[0];
                    dataGridView1[0, dataGridView1.RowCount - 1].Value = w[0];
                    dataGridView1[1, dataGridView1.RowCount - 1].Value = w[1];
                    dataGridView1[2, dataGridView1.RowCount - 1].Value = w[2];
                    dataGridView1[3, dataGridView1.RowCount - 1].Value = w[3];


                    //MessageBox.Show(w[0]);
                }
                dataGridView1.Rows.Add();
                dataGridView1[0, dataGridView1.RowCount - 1].Value = "Итого:";
                //itogo_ps.Add(dataGridView1.RowCount - 1);

                //Заполняем матрицу цифрами
                for (int i = 2; i <= srcArr_ob.GetUpperBound(0); i++)
                {
                    if (Convert.ToDateTime(srcArr_ob[i, 1]).Date >= dateTimePicker1.Value.Date &&
    Convert.ToDateTime(srcArr_ob[i, 1]).Date <= dateTimePicker2.Value.Date)//проверяем совпадение по дате
                    {
                        for (int w = 2; w < dataGridView1.Rows.Count; w++)
                        {
                            if (Convert.ToString(dataGridView1[0, w].Value) + Convert.ToString(dataGridView1[1, w].Value) + Convert.ToString(dataGridView1[2, w].Value)
                                + Convert.ToString(dataGridView1[3, w].Value) == Convert.ToString(srcArr_ob[i, 8]) + Convert.ToString(srcArr_ob[i, 4])
                                    + Convert.ToString(srcArr_ob[i, 5]) + Convert.ToString(srcArr_ob[i, 6]))//проверяем совпадения профилей
                            {
                                for (int r = 4; r < dataGridView1.ColumnCount; r++)
                                {
                                    if (Convert.ToString(dataGridView1[r, 0].Value) == Convert.ToString(srcArr_ob[i, 3]))//MessageBox.Show("есть совпадения по продовцу профилю и марки стали");
                                    {
                                        dataGridView1[r, w].Value = Convert.ToDouble(dataGridView1[r, w].Value) + Convert.ToDouble(srcArr_ob[i, 7]);
                                        dataGridView1[r + 2, w].Value = Convert.ToDouble(dataGridView1[r + 2, w].Value) + Convert.ToDouble(srcArr_ob[i, 11]);
                                        //MessageBox.Show("");
                                        break;
                                    }
                                }
                            }
                        }
                    }

                    progressBar1.PerformStep();

                }

                //считаем среднюю цену металла и итого
                for (int i = 2; i < dataGridView1.Rows.Count; i++)
                {
                    for (int w = 5; w < dataGridView1.ColumnCount; w = w + 3)
                    {
                        dataGridView1[w, i].Value = Convert.ToDouble(dataGridView1[w + 1, i].Value) * 1000 / Convert.ToDouble(dataGridView1[w - 1, i].Value);
                        double num = Convert.ToDouble(dataGridView1[w, i].Value);
                        if (num != (Double)num) // Проверка то что число - это число.
                        {
                            dataGridView1[w, i].Value = 0;
                        }
                    }
                }
                //Собираем столбец "всего"
                dataGridView1.Columns.Add("", "");
                dataGridView1.Columns.Add("", "");
                dataGridView1.Columns.Add("", "");
                dataGridView1[dataGridView1.ColumnCount - 3, 0].Value = "Всего";
                dataGridView1[dataGridView1.ColumnCount - 2, 0].Value = "Всего";
                dataGridView1[dataGridView1.ColumnCount - 1, 0].Value = "Всего";
                dataGridView1[dataGridView1.Columns.Count - 3, 1].Value = "масса, кг";
                dataGridView1[dataGridView1.Columns.Count - 2, 1].Value = "Средняя цена";
                dataGridView1[dataGridView1.Columns.Count - 1, 1].Value = "Сумма";

                for (int i = 2; i < dataGridView1.RowCount - 1; i++)
                {
                    for (int x = 4; x < dataGridView1.ColumnCount - 3; x++)
                    {
                        if (dataGridView1[x, 1].Value.ToString() == "масса, кг")
                        {
                            dataGridView1[dataGridView1.Columns.Count - 3, i].Value = Convert.ToDouble(dataGridView1[dataGridView1.Columns.Count - 3, i].Value) + Convert.ToDouble(dataGridView1[x, i].Value);
                        }
                        if (dataGridView1[x, 1].Value.ToString() == "Сумма")
                        {
                            dataGridView1[dataGridView1.Columns.Count - 1, i].Value = Convert.ToDouble(dataGridView1[dataGridView1.Columns.Count - 1, i].Value) + Convert.ToDouble(dataGridView1[x, i].Value);
                        }
                        dataGridView1[dataGridView1.Columns.Count - 2, i].Value = Convert.ToDouble(dataGridView1[dataGridView1.Columns.Count - 1, i].Value) * 1000 / Convert.ToDouble(dataGridView1[dataGridView1.Columns.Count - 3, i].Value);
                    }

                }

                for (int i = 4; i < dataGridView1.ColumnCount; i++)
                {
                    dataGridView1.Update();
                    for (int x = 2; x < dataGridView1.RowCount - 1; x++)
                    {
                        if (dataGridView1[i, 1].Value.ToString() == "масса, кг")
                        {
                            dataGridView1[i, dataGridView1.RowCount - 1].Value = Convert.ToDouble(dataGridView1[i, dataGridView1.RowCount - 1].Value) + Convert.ToDouble(dataGridView1[i, x].Value);
                        }
                        if (dataGridView1[i, 1].Value.ToString() == "Сумма")
                        {
                            dataGridView1[i, dataGridView1.RowCount - 1].Value = Convert.ToDouble(dataGridView1[i, dataGridView1.RowCount - 1].Value) + Convert.ToDouble(dataGridView1[i, x].Value);
                        }
                    }

                }
                for (int i = 4; i < dataGridView1.ColumnCount; i++)
                {
                    if (dataGridView1[i, 1].Value.ToString() == "Средняя цена")
                    {
                        // MessageBox.Show(dataGridView1[i + 1, dataGridView1.RowCount - 1].Value.ToString() + "  " + dataGridView1[i - 1, dataGridView1.RowCount - 1].Value.ToString());
                        dataGridView1[i, dataGridView1.RowCount - 1].Value = Convert.ToDouble(dataGridView1[i + 1, dataGridView1.RowCount - 1].Value) * 1000 / Convert.ToDouble(dataGridView1[i - 1, dataGridView1.RowCount - 1].Value);
                    }
                }
                //MessageBox.Show("!");
                if (double.IsNaN(Convert.ToDouble(dataGridView1[dataGridView1.ColumnCount-2, dataGridView1.RowCount - 1].Value))) // Проверка итоговая цена - это число
                {
                    dataGridView1[dataGridView1.ColumnCount-2, dataGridView1.RowCount - 1].Value = "какая то ирунда";
                    //MessageBox.Show("!!!");
                }


                dateTimePicker1.Enabled = false;
                dateTimePicker2.Enabled = false;
                progressBar1.PerformStep();

                
                //пробуем сделать всё на одну кнопку :)

                progressBar1.Value = progressBar1.Maximum * 3 / 4;
                textBox1.Text = "Ждем экселя два...";
                //MessageBox.Show("");
                //Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                //Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
                //Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
                //Книга.
                ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
                //Таблица.
                ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        ExcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                    }
                }
                //форматируем таблицу
                //string[] vsS ={ "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", 
//"M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", 
//"AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW",
//"AX", "AY", "AZ" , 
//"BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW",
//"BX", "BY", "BZ",
//"CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW",
//"CX", "CY", "CZ",
//"DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW",
//"DX", "DY", "DZ",
//"EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW",
//"EX", "EY", "EZ",
//"FA", "FB", "FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN", "FO", "FP", "FQ", "FR", "FS", "FT", "FU", "FV", "FW",
//"FX", "FY", "FZ"};
                ExcelWorkSheet.get_Range("A1", "A1").Value = "Закупка металлопроката для транзита за перюд с " + dateTimePicker1.Value.Date.ToString().Replace(" 0:00:00", "") + " по " + dateTimePicker2.Value.Date.ToString().Replace(" 0:00:00", "");
                ExcelWorkSheet.get_Range(vsS[0] + "1", vsS[dataGridView1.ColumnCount - 1] + "1").MergeCells = true;
                ExcelWorkSheet.get_Range(vsS[0] + "1", vsS[dataGridView1.ColumnCount - 1] + "1").Font.Bold = true;
                ExcelWorkSheet.get_Range(vsS[0] + "1", vsS[dataGridView1.ColumnCount - 1] + "1").Font.Size = 18;
                ExcelWorkSheet.get_Range(vsS[0] + "1", vsS[dataGridView1.ColumnCount - 1] + "1").HorizontalAlignment = Excel.Constants.xlCenter;
                ExcelWorkSheet.get_Range(vsS[0] + "1", vsS[dataGridView1.ColumnCount - 1] + "1").VerticalAlignment = Excel.Constants.xlCenter;
                ExcelWorkSheet.get_Range("A2", "A3").MergeCells = true;
                ExcelWorkSheet.get_Range("A2", "A3").Font.Size = 14;
                ExcelWorkSheet.get_Range("A2", "A3").HorizontalAlignment = Excel.Constants.xlCenter;
                ExcelWorkSheet.get_Range("A2", "A3").VerticalAlignment = Excel.Constants.xlCenter;

                ExcelWorkSheet.get_Range("B2", "B3").MergeCells = true;
                ExcelWorkSheet.get_Range("B2", "B3").Font.Size = 14;
                ExcelWorkSheet.get_Range("B2", "B3").HorizontalAlignment = Excel.Constants.xlCenter;
                ExcelWorkSheet.get_Range("B2", "B3").VerticalAlignment = Excel.Constants.xlCenter;

                ExcelWorkSheet.get_Range("C2", "C3").MergeCells = true;
                ExcelWorkSheet.get_Range("C2", "C3").Font.Size = 14;
                ExcelWorkSheet.get_Range("C2", "C3").HorizontalAlignment = Excel.Constants.xlCenter;
                ExcelWorkSheet.get_Range("C2", "C3").VerticalAlignment = Excel.Constants.xlCenter;

                ExcelWorkSheet.get_Range("D2", "D3").MergeCells = true;
                ExcelWorkSheet.get_Range("D2", "D3").Font.Size = 14;
                ExcelWorkSheet.get_Range("D2", "D3").HorizontalAlignment = Excel.Constants.xlCenter;
                ExcelWorkSheet.get_Range("D2", "D3").VerticalAlignment = Excel.Constants.xlCenter;

                ExcelWorkSheet.get_Range("A" + (dataGridView1.RowCount + 1).ToString(), "D" + (dataGridView1.RowCount + 1).ToString()).MergeCells = true;
                ExcelWorkSheet.get_Range("A" + (dataGridView1.RowCount + 1).ToString(), "D" + (dataGridView1.RowCount + 1).ToString()).HorizontalAlignment = Excel.Constants.xlRight;
                ExcelWorkSheet.get_Range("A" + (dataGridView1.RowCount + 1).ToString(), "D" + (dataGridView1.RowCount + 1).ToString()).VerticalAlignment = Excel.Constants.xlCenter;

                ExcelWorkSheet.Rows[(dataGridView1.RowCount + 1).ToString(), Type.Missing].Font.Size = 14;
                ExcelWorkSheet.Rows[(dataGridView1.RowCount + 1).ToString(), Type.Missing].Font.Bold = true;

                ExcelWorkSheet.Rows["3", Type.Missing].HorizontalAlignment = Excel.Constants.xlCenter;
                ExcelWorkSheet.Rows["3", Type.Missing].VerticalAlignment = Excel.Constants.xlCenter;

                ExcelWorkSheet.Rows["2", Type.Missing].RowHeight = "50";
                //ExcelWorkSheet.get_Range("A2", "A3").EntireColumn.AutoFit();

                //Устанавливаем типы для столбцов

                for (int j = 4; j < dataGridView1.ColumnCount; j = j + 3)
                {
                    ExcelWorkSheet.Columns[vsS[j], Type.Missing].NumberFormat = "# ###";
                    ExcelWorkSheet.Columns[vsS[j + 1], Type.Missing].Style = "Currency";
                    ExcelWorkSheet.Columns[vsS[j + 2], Type.Missing].Style = "Currency";
                }

                //ExcelWorkSheet.get_Range("D2", "D3").Style = Excel.Style("Currency");


                //Обьединяме объекты
                ExcelApp.DisplayAlerts = false;
                for (int j = 4; j < dataGridView1.ColumnCount; j = j + 3)
                {
                    ExcelWorkSheet.get_Range(vsS[j] + "2", vsS[j + 2] + "2").MergeCells = true;
                    ExcelWorkSheet.get_Range(vsS[j] + "2", vsS[j + 2] + "2").HorizontalAlignment = Excel.Constants.xlCenter;
                    ExcelWorkSheet.get_Range(vsS[j] + "2", vsS[j + 2] + "2").VerticalAlignment = Excel.Constants.xlCenter;
                    ExcelWorkSheet.get_Range(vsS[j] + "2", vsS[j + 2] + "2").WrapText = "true";
                    ExcelWorkSheet.get_Range(vsS[j] + "2", vsS[j + 2] + "2").Font.Size = 12;
                }
                //Устанавливаем автоширину
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    ExcelWorkSheet.Columns[vsS[j], Type.Missing].EntireColumn.AutoFit();
                }

                //Рисуем границы
                ExcelWorkSheet.get_Range("A1", vsS[dataGridView1.ColumnCount - 1].ToString() + (dataGridView1.RowCount + 1).ToString()).Borders[Excel.XlBordersIndex.xlInsideHorizontal]
                    .LineStyle = Excel.XlLineStyle.xlContinuous;
                ExcelWorkSheet.get_Range("A1", vsS[dataGridView1.ColumnCount - 1].ToString() + (dataGridView1.RowCount + 1).ToString()).Borders[Excel.XlBordersIndex.xlInsideVertical]
        .LineStyle = Excel.XlLineStyle.xlContinuous;
                ExcelWorkSheet.get_Range("A1", vsS[dataGridView1.ColumnCount - 1].ToString() + (dataGridView1.RowCount + 1).ToString()).Borders[Excel.XlBordersIndex.xlEdgeBottom]
    .LineStyle = Excel.XlLineStyle.xlContinuous;
                ExcelWorkSheet.get_Range("A1", vsS[dataGridView1.ColumnCount - 1].ToString() + (dataGridView1.RowCount + 1).ToString()).Borders[Excel.XlBordersIndex.xlEdgeTop]
    .LineStyle = Excel.XlLineStyle.xlContinuous;
                ExcelWorkSheet.get_Range("A1", vsS[dataGridView1.ColumnCount - 1].ToString() + (dataGridView1.RowCount + 1).ToString()).Borders[Excel.XlBordersIndex.xlEdgeLeft]
    .LineStyle = Excel.XlLineStyle.xlContinuous;
                ExcelWorkSheet.get_Range("A1", vsS[dataGridView1.ColumnCount - 1].ToString() + (dataGridView1.RowCount + 1).ToString()).Borders[Excel.XlBordersIndex.xlEdgeRight]
    .LineStyle = Excel.XlLineStyle.xlContinuous;

                //Делаем культурный вид
                ExcelApp.ActiveWindow.SplitRow = 3;
                ExcelApp.ActiveWindow.SplitColumn = 4;
                ExcelApp.ActiveWindow.FreezePanes = true;



                //Вызываем нашу созданную эксельку.
                ExcelApp.DisplayAlerts = true;
                ExcelApp.Visible = true;
                ExcelApp.UserControl = true;
                progressBar1.Value = progressBar1.Maximum;
                textBox1.Text = "ГОТОВО!!!";
                
                #endregion
                // ============== 
                MessageBox.Show("Обязательно проверь контрольную сумму по массе и сумме денег.\nЕсли она не совпадает - проверь диапазан дат. \nЕсли всеравно не совпадает обратись к Константину. \n \tПрограмму можно закрыть.","Скажи спасибо Константину",MessageBoxButtons.OK,MessageBoxIcon.Information);
        
                //Код дописывать не ниже этого коментария
                Application.Exit();
                 
            }
            // MessageBox.Show("Строк в файле " + srcArr_ob.GetUpperBound(0).ToString()); - проерка что все строки взял
        }
    }

}