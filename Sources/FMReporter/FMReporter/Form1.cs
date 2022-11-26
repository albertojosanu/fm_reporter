using Microsoft.Data.SqlClient;
using System;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace FMReporter
{
    public partial class Form1 : Form
    {
        Int64[,] sumV;// = new Int64[5, 2];
        Int64[,] sumM;// = new Int64[5, 2];
        List<String> productName = new List<String>();
        Int16 AsnCount;
        Int16 type; // Выбор для смены: 1 - указывается только дата (с условием на выбранную дату),
                    // 2 - указывается дата и время начала смены и конца смены

        public Form1()
        {
            InitializeComponent();
            initSizes();
            sumV = new Int64[AsnCount, 2];
            sumM = new Int64[AsnCount, 2];
        }

        private void button1_Click(object sender, EventArgs e)
        {
            clearArrays();
            DateTime dt = dateTimePicker1.Value;
            DateTime dt2 = dateTimePicker2.Value;
            try
            {
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                builder.DataSource = "localhost";
                builder.UserID = "sa";
                builder.Password = "tm162#bcdef";
                builder.InitialCatalog = "OZNA_ASN";

                using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
                {
                    for (int i = 0; i < AsnCount; i++)
                    {
                        for (int req = 0; req < 2; req++)
                        {
                            String sql;
                            if (req == 0)   //Первый запрос - сработает, если наливы были в этот день
                            {
                                if (type == 1)
                                    sql = "SELECT PV1, PV2, PM1, PM2, ASN FROM MEAS WHERE (ASN=" + (i + 1).ToString() + ")"
                                    + " AND (DT1<=CAST('" + dt.ToString("yyyyMMdd") + " 23:59:59.999' AS DATETIME))"
                                    + " AND (DT1>=CAST('" + dt.ToString("yyyyMMdd") + " 00:00:00.000' AS DATETIME))"
                                    + " AND (PV1 IS NOT NULL) AND (PV2 IS NOT NULL) AND (PM1 IS NOT NULL) AND (PM2 IS NOT NULL) "
                                    + " ORDER BY DT1 DESC";
                                else
                                    sql = "SELECT PV1, PV2, PM1, PM2, ASN FROM MEAS WHERE (ASN=" + (i + 1).ToString() + ")"
                                    + " AND (DT1<=CAST('" + dt2.ToString("yyyyMMdd HH:mm:ss") + "' AS DATETIME))"
                                    + " AND (DT1>=CAST('" + dt.ToString("yyyyMMdd HH:mm:ss") + "' AS DATETIME))"
                                    + " AND (PV1 IS NOT NULL) AND (PV2 IS NOT NULL) AND (PM1 IS NOT NULL) AND (PM2 IS NOT NULL) "
                                    + " ORDER BY DT1 DESC";
                            }
                            else            //Второй запрос - отправится в случае, если наливов в этот день не было
                            {
                                sql = "SELECT TOP(1) PV2, PM2, ASN FROM MEAS"
                                        + " WHERE (ASN=" + (i + 1).ToString() + ")"
                                        + " AND (DT1<=CAST('" + dt.ToString("yyyyMMdd") + " 23:59:59.999' AS DATETIME))"
                                        + " AND (PV2 IS NOT NULL) AND (PM2 IS NOT NULL) "
                                        + " ORDER BY DT1 DESC";
                            }

                            using (SqlCommand command = new SqlCommand(sql, connection))
                            {
                                connection.Open();
                                using (SqlDataReader reader = command.ExecuteReader())
                                {
                                    if (req == 0)
                                    {
                                        if (reader.Read())
                                        {
                                            if (!reader.IsDBNull(0))
                                                sumV[i, 0] = Convert.ToInt64(reader.GetDouble(0) * 1000);
                                            if (!reader.IsDBNull(2))
                                                sumM[i, 0] = reader.GetInt64(2);
                                            if (!reader.IsDBNull(1))
                                                sumV[i, 1] = Convert.ToInt64(reader.GetDouble(1) * 1000);
                                            if (!reader.IsDBNull(3))
                                                sumM[i, 1] = reader.GetInt64(3);
                                            while (reader.Read())
                                            {
                                                if (!reader.IsDBNull(0))
                                                    sumV[i, 0] = Convert.ToInt64(reader.GetDouble(0) * 1000);
                                                if (!reader.IsDBNull(2))
                                                    sumM[i, 0] = reader.GetInt64(2);
                                            }
                                            reader.Close();
                                            connection.Close();
                                            break;  //Чтобы не отправлять второй запрос
                                        }
                                    }
                                    else
                                    {
                                        if (reader.Read())
                                        {
                                            if (!reader.IsDBNull(0))
                                                sumV[i, 0] = Convert.ToInt64(reader.GetDouble(0) * 1000);
                                            if (!reader.IsDBNull(1))
                                                sumM[i, 0] = reader.GetInt64(1);
                                            sumV[i, 1] = sumV[i, 0];
                                            sumM[i, 1] = sumM[i, 0];
                                        }
                                    }
                                    reader.Close();
                                    connection.Close();
                                }

                            }
                        }
                    }
                }
            }
            catch (SqlException se)
            {
                MessageBox.Show("Ошибка при выполнении запроса! " + se.ToString());
            }

            //String tmp = dt.ToString("dd.MM.yyyy");
            if (type == 1)
            {
                dataGridView1[0, 0].Value = dateTimePicker1.Value.ToString("dd.MM.yyyy");
            } 
            else
            {
                dataGridView1[0, 0].Value = dateTimePicker1.Value.ToString("dd.MM.yyyy HH:mm:ss");
                dataGridView1[0, 1].Value = dateTimePicker2.Value.ToString("dd.MM.yyyy HH:mm:ss");
            }
            

            for (int i = 0; i < AsnCount; i++)
            {
                //dataGridView1[0, i].Value = tmp;
                //dataGridView1[0, (i * 2)].Value = tmp;

                dataGridView1[4, (i * 2)].Value = sumV[i, 0].ToString();
                dataGridView1[5, (i * 2)].Value = sumV[i, 1].ToString();

                dataGridView1[4, (i * 2) + 1].Value = sumM[i, 0].ToString();
                dataGridView1[5, (i * 2) + 1].Value = sumM[i, 1].ToString();
            }
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            dataGridView1.ColumnCount = 6;
            dataGridView1.RowCount = AsnCount * 2;
            dataGridView1.ColumnHeadersVisible = true;

            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            dataGridView1.Columns[0].Width = 110;
            dataGridView1.Columns[1].Width = 120;
            dataGridView1.Columns[2].Width = 50;
            dataGridView1.Columns[3].Width = 20;
            dataGridView1.Columns[4].Width = 140;
            dataGridView1.Columns[5].Width = 140;

            if (type == 1)
            {
                dataGridView1.Columns[0].Name = "Дата";
                dataGridView1.Columns[4].Name = "Показание счетчиков на начало смены";
                dataGridView1.Columns[5].Name = "Показание счетчиков на конец смены";
            } 
            else
            {
                dataGridView1.Columns[0].Name = "Временной промежуток";
                dataGridView1.Columns[4].Name = "Показание счетчиков на начало временного промежутка";
                dataGridView1.Columns[5].Name = "Показание счетчиков на конец временного промежутка";
            }
            dataGridView1.Columns[1].Name = "Наименование нефтепродукта";
            dataGridView1.Columns[2].Name = "№ АСН";
            dataGridView1.Columns[3].Name = "";


            for (int i = 0; i < AsnCount; i++)
            {
                dataGridView1[1, (i * 2)].Value = productName[i];
                dataGridView1[1, (i * 2) + 1].Value = productName[i];

                dataGridView1[2, (i * 2)].Value = i + 1;
                dataGridView1[2, (i * 2) + 1].Value = i + 1;

                dataGridView1[3, (i * 2)].Value = "V";
                dataGridView1[3, (i * 2) + 1].Value = "M";
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            exportToExcel();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            //if (dateTimePicker1.Value >= (DateTime.Now.AddDays(1)))
            //{
            //    dateTimePicker1.Value = DateTime.Now;
            //}
        }

        private void clearArrays()
        {
            for (int i = 0; i < AsnCount; i++)
            {
                for (int j = 0; j < 2; j++)
                {
                    sumV[i, j] = 0;
                    sumM[i, j] = 0;
                }
            }
        }

        private void exportToExcel()
        {
            // Без принудительной сборки мусора, стакаются Excel-процессы, хз почему.
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Excel.Application oXL = new Excel.Application();
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range _excelCells1;
            Excel.Range rng;

            try
            {
                oXL.Visible = true;

                //Get a new workbook.
                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                rng = (Excel.Range)oSheet.Range[oSheet.Cells[1, 1], oSheet.Cells[7 + (AsnCount * 2), 6]];
                rng.Cells.Font.Name = "Times New Roman";
                rng.Cells.Font.Size = 12;


                //Titles
                if (type == 1)
                {
                    oSheet.Cells[1, 1] = "Дата";
                    oSheet.Cells[2, 5] = "На начало дня";
                    oSheet.Cells[2, 6] = "На конец дня";
                } 
                else
                {
                    oSheet.Cells[1, 1] = "Временной промежуток";
                    oSheet.Cells[2, 5] = "На начало временного промежутка";
                    oSheet.Cells[2, 6] = "На конец временного промежутка";
                }



                oSheet.Cells[1, 2].Style.WrapText = true;
                

                oSheet.Columns[1].ColumnWidth = 14;

                _excelCells1 = (Excel.Range)oSheet.get_Range("B1", "B2").Cells;
                _excelCells1.Merge(Type.Missing);
                oSheet.Cells[1, 2] = "Наименование нефтепродукта";
                oSheet.Columns[2].ColumnWidth = 20;

                _excelCells1 = (Excel.Range)oSheet.get_Range("C1", "C2").Cells;
                _excelCells1.Merge(Type.Missing);
                oSheet.Cells[1, 3] = "№ АСН";


                _excelCells1 = (Excel.Range)oSheet.get_Range("D1", "D2").Cells;
                _excelCells1.Merge(Type.Missing);

                _excelCells1 = (Excel.Range)oSheet.get_Range("E1", "F1").Cells;
                _excelCells1.Merge(Type.Missing);
                oSheet.Cells[1, 5] = "Показания счетчиков";
                oSheet.Columns[5].ColumnWidth = 15;
                oSheet.Columns[6].ColumnWidth = 15;

                

                oSheet.get_Range("A1", "F2").Font.Bold = true;
                oSheet.get_Range("A1", "F1").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("A1", "F1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                oSheet.get_Range("A2", "A2").VerticalAlignment = Excel.XlVAlign.xlVAlignTop;

                //Data
                rng = (Excel.Range)oSheet.Range[oSheet.Cells[1, 3], oSheet.Cells[2 + (AsnCount * 2), 4]];
                rng.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                rng = (Excel.Range)oSheet.Range[oSheet.Cells[2, 1], oSheet.Cells[2 + (AsnCount * 2), 1]];
                rng.Merge(Type.Missing);
               
                if (type == 1)
                {
                    oSheet.Cells[2, 1] = dataGridView1[0, 0].Value;
                }
                else
                {
                    oSheet.Cells[2, 1] = dataGridView1[0, 0].Value + "\n — \n" + dataGridView1[0, 1].Value;
                }
                


                for (int i = 0; i < (AsnCount * 2); i++)
                {
                    for (int j = 0; j < 5; j++)
                    {
                        oSheet.Cells[3 + i, 2 + j] = dataGridView1[1 + j, 0 + i].Value;
                    }
                }

                rng = (Excel.Range)oSheet.Range[oSheet.Cells[1, 1], oSheet.Cells[2, 6]];
                rng.Borders.ColorIndex = 0;
                rng.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                rng.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick;

                rng = (Excel.Range)oSheet.Range[oSheet.Cells[3, 2], oSheet.Cells[2 + (AsnCount * 2), 6]];
                rng.Borders.ColorIndex = 0;
                rng.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;


                rng = (Excel.Range)oSheet.Range[oSheet.Cells[1, 1], oSheet.Cells[2 + (AsnCount * 2), 6]];
                rng.BorderAround(Excel.XlLineStyle.xlContinuous,
                        Excel.XlBorderWeight.xlThick,
                        Excel.XlColorIndex.xlColorIndexNone,
                        0);

                for (int i = 0; i < AsnCount; ++i)
                {
                    rng = (Excel.Range)oSheet.Range[oSheet.Cells[3 + (i * 2), 2], oSheet.Cells[4 + (i * 2), 6]];
                    rng.BorderAround(Excel.XlLineStyle.xlContinuous,
                        Excel.XlBorderWeight.xlThick,
                        Excel.XlColorIndex.xlColorIndexNone,
                        0);
                }


                //Signatures
                rng = (Excel.Range)oSheet.Range[oSheet.Cells[5 + (AsnCount * 2), 1], oSheet.Cells[5 + (AsnCount * 2), 2]];
                rng.Merge(Type.Missing);
                rng = (Excel.Range)oSheet.Range[oSheet.Cells[7 + (AsnCount * 2), 1], oSheet.Cells[7 + (AsnCount * 2), 2]];
                rng.Merge(Type.Missing);
                oSheet.Cells[5 + (AsnCount * 2), 1] = "Товарный оператор";
                oSheet.Cells[7 + (AsnCount * 2), 1] = "Оператор по учету";
                for(int i = 0; i < 2; ++i)
                {
                    rng = (Excel.Range)oSheet.Range[oSheet.Cells[5 + (i * 2) + (AsnCount * 2), 3], oSheet.Cells[5 + (i * 2) + (AsnCount * 2), 4]];
                    rng.Merge(Type.Missing);
                    oSheet.Cells[5 + (i * 2) + (AsnCount * 2), 3] = "_______________";
                    rng = (Excel.Range)oSheet.Range[oSheet.Cells[5 + (i * 2) + (AsnCount * 2), 5], oSheet.Cells[5 + (i * 2) + (AsnCount * 2), 6]];
                    rng.Merge(Type.Missing);
                    oSheet.Cells[5 + (i * 2) + (AsnCount * 2), 5] = "___________________________";
                }

                rng = (Excel.Range)oSheet.Range[oSheet.Cells[5 + (AsnCount * 2), 1], oSheet.Cells[7 + (AsnCount * 2), 2]];
                rng.Font.Bold = true;
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }


        private void initSizes()
        {
            productName.Clear();
            //string path = @"..\test.cfg";
            string path = "config.cfg";
            try
            {
                using (StreamReader sr = new StreamReader(path, System.Text.Encoding.Default))
                {
                    //считываем тип выбора смены
                    string line = sr.ReadLine();
                    Regex regexCount = new Regex(@"(\d+).+$");
                    type = Convert.ToInt16(regexCount.Match(line).Groups[1].ToString());

                    //считываем количество АСН
                    line = sr.ReadLine();
                    AsnCount = Convert.ToInt16(regexCount.Match(line).Groups[1].ToString());
                    Console.WriteLine(regexCount.Match(line));
                    visualsInit();

                    line = sr.ReadLine();
                    while (line != null)
                    {
                        productName.Add(line.Trim());
                        line = sr.ReadLine();
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Error", e.ToString());
                return;
            }

        } 

        private void visualsInit()
        {
            //видимость и форматы визуалкам
            DateTime dt1;
            dt1 = DateTime.Now;
            string tmp = "";
            if (dt1.Day <= 9)
                tmp += "0";
            tmp += dt1.Day.ToString() + "/";
            if (dt1.Month <= 9)
                tmp += "0";
            tmp += dt1.Month.ToString() + "/" + dt1.Year.ToString();

            if (type == 1)
            {
                dateTimePicker1.Format = DateTimePickerFormat.Short;
                dateTimePicker1.Visible = true;
                dateTimePicker2.Visible = false;
                label1.Text = "Выбор даты";
            }
            else
            {
                dateTimePicker1.Format = DateTimePickerFormat.Custom;
                dateTimePicker1.CustomFormat = @"dd/MM/yyyy HH:mm";
                dateTimePicker1.Value = DateTime.ParseExact(tmp + " 08:00", @"dd/MM/yyyy HH:mm", null); ;

                dateTimePicker2.Format = DateTimePickerFormat.Custom;
                dateTimePicker2.CustomFormat = @"dd/MM/yyyy HH:mm";
                dateTimePicker2.Value = DateTime.ParseExact(tmp + " 18:00", @"dd/MM/yyyy HH:mm", null);
                dateTimePicker1.Visible = true;
                dateTimePicker2.Visible = true;
                label1.Text = "Выбор временного промежутка";
            }
        }
    }
}
