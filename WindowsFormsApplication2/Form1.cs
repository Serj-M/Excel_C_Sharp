using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace WindowsFormsApplication2
{
    public partial class Form1 : Form
    {
        private Microsoft.Office.Interop.Excel.Application ObjExcel;
        private Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
        private Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
        string directoryFile = "";
        string path;
        string num_file = "00";
        string num_file_previons = "00";
        int i_exc = 1;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        //************ИМПОРТ*************

        private void button1_Click(object sender, EventArgs e) 
        {
            dataGridView1.Rows.Clear();
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Файл Excel|*.XLSX;*.XLS";
            openDialog.ShowDialog();

            try
            {
                ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                //Книга.
                ObjWorkBook = ObjExcel.Workbooks.Open(openDialog.FileName);
                //Таблица.
                ObjWorkSheet = ObjExcel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                Microsoft.Office.Interop.Excel.Range rg = null;
                //ObjExcel.Visible = true;

                // directoryFileDodelat - хранит директорию и имя файла без расширения (.xls или .xlsx). Понадобиться для экспорта в туже директорию 
                directoryFile = openDialog.FileName;
                //MessageBox.Show(directoryFile);
                
                Int32 row = 1;
                Int32 more_Renge = 0;
                dataGridView1.Rows.Clear();
                List<String> arr = new List<string>();
                // пока не конец файла формируем датаГрид
                while (ObjWorkSheet.get_Range("b" + row, "b" + row).Value != null)
                {
                    // Читаем данные из ячейки
                    rg = ObjWorkSheet.get_Range("a" + row, "c" + row);
                    foreach (Microsoft.Office.Interop.Excel.Range item in rg)
                    {
                        try
                        {
                            arr.Add(item.Value.ToString().Trim());
                        }
                        catch { arr.Add(""); }
                    }
                    dataGridView1.Rows.Add(arr[0], arr[1], arr[2]/*, arr[3], arr[4], arr[5]*/);
                    for (int i = 1; i < dataGridView1.RowCount; i++) //цикл по всему датагриду
                    {
                        dataGridView1[3, i].Value = dataGridView1[2, i].Value;
                        dataGridView1[4, i].Value = "KSM_Agent";
                        dataGridView1[5, i].Value = "1qaz@WSX3edc";
                    }
                    arr.Clear();
                    //more_Renge = 0;
                    row++;
                }
                //проверка разрыва в Excel документе (поиск пустых строк в заполненом деопазоне)
                while ((ObjWorkSheet.get_Range("a" + row, "a" + row).Value != null) || (more_Renge < 5))
                {
                    if (ObjWorkSheet.get_Range("a" + row, "a" + row).Value == null) // проверка пустых сторк (more_Renge < 5)
                    {
                        more_Renge++;
                        row++;
                        continue;
                    }
                    else
                    {
                        if (more_Renge > 0)
                            MessageBox.Show("Исходный файл офрмлен не правильно. Не все данные из файла загружены! Существуют разрывы");
                    }
                    break;
                }

                MessageBox.Show("Файл успешно считан!", "Считывания excel файла", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch /*(Exception ex)*/ {/* MessageBox.Show("Ошибка: " + ex.Message, "Ошибка при считывании excel файла", MessageBoxButtons.OK, MessageBoxIcon.Error); */}
            finally
            {
                try
                {
                    ObjWorkBook.Close(false, "", null);
                    // Закрытие приложения Excel.
                    ObjExcel.Quit();
                    ObjWorkBook = null;
                    ObjWorkSheet = null;
                    ObjExcel = null;
                    GC.Collect();
                }
                catch
                {
                }
            }

            this.Text = this.Text + " - " + openDialog.SafeFileName;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                for (int i = 1; i < dataGridView1.RowCount; i++) //цикл по всему датагриду
                {
                    num_file_previons = num_file;
                    if (dataGridView1[0, i].Value != "") // если найден новый регион
                    {
                        i_exc = 1;
                        path = directoryFile.Remove(directoryFile.LastIndexOf(@"\")); // путь к файлу
                        num_file = dataGridView1[0, i].Value.ToString().Substring(0, 2); // определяется имя файла по номеру региона
                        if (i > 1) // если это не начало
                        {
                            ObjWorkBook.SaveAs(path + "\\" + num_file_previons + ".xlsx");
                            ObjWorkBook.Close();
                        }
                        ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                        //Книга.
                        ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
                        //Таблица.
                        ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

                        DataGridViewRow row = dataGridView1.Rows[i]; // строки

                        for (int j = 1; j < row.Cells.Count; j++) //цикл по ячейкам строки
                        {
                            ObjExcel.Cells[1, 1] = "desc";
                            ObjExcel.Cells[1, 2] = "ip";
                            ObjExcel.Cells[1, 3] = "name";
                            ObjExcel.Cells[1, 4] = "login";
                            ObjExcel.Cells[1, 5] = "pass";
                            ObjExcel.Cells[i_exc + 1, j] = row.Cells[j].Value;
                        }
                        i_exc++;
                    }
                    else
                    {
                        DataGridViewRow row = dataGridView1.Rows[i]; // строки

                        for (int j = 1; j < row.Cells.Count; j++) //цикл по ячейкам строки
                        {
                            ObjExcel.Cells[i_exc + 1, j] = row.Cells[j].Value;
                        }
                        i_exc++;
                        if (i == dataGridView1.RowCount-2)
                        {
                            ObjWorkBook.SaveAs(path + "\\" + num_file_previons + ".xlsx"); // сохранение для последнего региона в списке
                            ObjWorkBook.Close();
                        }
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
            finally
            {
                {
                    //ObjWorkBook.Close();
                    // Закрытие приложения Excel.
                    ObjExcel.Quit();
                    ObjWorkBook = null;
                    ObjWorkSheet = null;
                    ObjExcel = null;
                    GC.Collect();
                }
            }
        }
    }
}
