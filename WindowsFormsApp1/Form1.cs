using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;


namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private System.Windows.Forms.TextBox[] textBoxes;
        public Form1()
        {
            InitializeComponent();

            textBoxes = new System.Windows.Forms.TextBox[] { textBox4, textBox9, textBox6, textBox3, textBox7, textBox11, textBox10, textBox13, textBox12, textBox8 };
            ExcelPackage.License.SetNonCommercialPersonal("Daniil");
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }



        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str = comboBox1.SelectedItem.ToString();

            switch (str)
            {
                case "ООО \"СтройМаш\"":
                    textBoxOrgan.Text = "00001";
                    break;
                case "ООО \"ТвояМашинПочинилз\"":
                    textBoxOrgan.Text = "00002";
                    break;
                case "ООО \"КредитВКредит\"":
                    textBoxOrgan.Text = "00003";
                    break;
                default:
                    textBox1.Text = "";
                    break;
            }
        }


        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str = comboBox3.SelectedItem.ToString();

            switch (str)
            {
                case "Продукты пищевые":
                    textBoxDeyet.Text = "00001";
                    break;
                case "Напитки":
                    textBoxDeyet.Text = "00002";
                    break;
                case "Изделия табачные":
                    textBoxDeyet.Text = "00003";
                    break;
                default:
                    textBoxDeyet.Text = "";
                    break;
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str = comboBox4.SelectedItem.ToString();

            switch (str)
            {
                case "Производство пищевых продуктов":
                    textBoxOper.Text = "00001";
                    break;
                case "Производство напитков":
                    textBoxOper.Text = "00002";
                    break;
                case "Производство табачных изделий":
                    textBoxOper.Text = "00003";
                    break;
                default:
                    textBoxOper.Text = "";
                    break;
            }
        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            int rowIndex = e.RowIndex + 1;
            dataGridView1.Rows[e.RowIndex].Cells["Num"].Value = rowIndex.ToString();
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            // 2 столбец
            if (e.ColumnIndex == 1 && e.RowIndex >= 0)
            {
                string item = dataGridView1.Rows[e.RowIndex].Cells[1].Value as string;

                if (!string.IsNullOrEmpty(item))
                {
                    if (item == "Плов")
                    {
                        dataGridView1.Rows[e.RowIndex].Cells[2].Value = "1001";
                        dataGridView1.Rows[e.RowIndex].Cells[3].Value = "шт";
                        dataGridView1.Rows[e.RowIndex].Cells[4].Value = "796";
                        dataGridView1.Rows[e.RowIndex].Cells[5].Value = "1";
                        dataGridView1.Rows[e.RowIndex].Cells[6].Value = "250";

                    }
                    if (item == "Борщ")
                    {
                        dataGridView1.Rows[e.RowIndex].Cells[2].Value = "1002";
                        dataGridView1.Rows[e.RowIndex].Cells[3].Value = "шт";
                        dataGridView1.Rows[e.RowIndex].Cells[4].Value = "796";
                        dataGridView1.Rows[e.RowIndex].Cells[5].Value = "0.5";
                        dataGridView1.Rows[e.RowIndex].Cells[6].Value = "50";
                    }

                    if (item == "Гречка")
                    {
                        dataGridView1.Rows[e.RowIndex].Cells[2].Value = "1003";
                        dataGridView1.Rows[e.RowIndex].Cells[3].Value = "шт";
                        dataGridView1.Rows[e.RowIndex].Cells[4].Value = "796";
                        dataGridView1.Rows[e.RowIndex].Cells[5].Value = "0.87";
                        dataGridView1.Rows[e.RowIndex].Cells[6].Value = "150";
                    }


                }
            }
            if (e.RowIndex >= 0)
            {
                if (dataGridView1.Rows[e.RowIndex].Cells[1].Value != null)
                {
                    if (e.ColumnIndex == 7 && e.RowIndex >= 0)
                    {
                        string item = dataGridView1.Rows[e.RowIndex].Cells[7].Value as string;
                        int sum;
                        int cost = 0;
                        int num = 0;

                        if (dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() == "Плов")
                        {
                            cost = 250;

                        }
                        if (dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() == "Борщ")
                        {
                            cost = 50;
                        }

                        if (dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() == "Гречка")
                        {
                            cost = 150;
                        }


                        if (!string.IsNullOrEmpty(item))
                        {
                            if (int.TryParse(item, out sum))
                            {
                                if (int.TryParse(dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString(), out num))
                                {
                                    sum = num * cost;
                                    dataGridView1.Rows[e.RowIndex].Cells[8].Value = sum.ToString();
                                }
                            }
                        }



                    }

                    if (e.ColumnIndex == 9 && e.RowIndex >= 0)
                    {
                        string item = dataGridView1.Rows[e.RowIndex].Cells[9].Value as string;
                        int sum;
                        int cost = 0;
                        int num = 0;

                        if (dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() == "Плов")
                        {
                            cost = 250;

                        }
                        if (dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() == "Борщ")
                        {
                            cost = 50;
                        }

                        if (dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() == "Гречка")
                        {
                            cost = 150;
                        }


                        if (!string.IsNullOrEmpty(item))
                        {
                            if (int.TryParse(item, out sum))
                            {
                                if (int.TryParse(dataGridView1.Rows[e.RowIndex].Cells[9].Value.ToString(), out num))
                                {
                                    sum = num * cost;
                                    dataGridView1.Rows[e.RowIndex].Cells[10].Value = sum.ToString();
                                }
                            }
                        }
                    }

                    if (e.ColumnIndex == 11 && e.RowIndex >= 0)
                    {
                        string item = dataGridView1.Rows[e.RowIndex].Cells[11].Value as string;
                        int sum;
                        int cost = 0;
                        int num = 0;

                        if (dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() == "Плов")
                        {
                            cost = 250;

                        }
                        if (dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() == "Борщ")
                        {
                            cost = 50;
                        }

                        if (dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() == "Гречка")
                        {
                            cost = 150;
                        }


                        if (!string.IsNullOrEmpty(item))
                        {
                            if (int.TryParse(item, out sum))
                            {
                                if (int.TryParse(dataGridView1.Rows[e.RowIndex].Cells[11].Value.ToString(), out num))
                                {
                                    sum = num * cost;
                                    dataGridView1.Rows[e.RowIndex].Cells[12].Value = sum.ToString();
                                }
                            }
                        }
                    }

                    if (e.ColumnIndex == 13 && e.RowIndex >= 0)
                    {
                        string item = dataGridView1.Rows[e.RowIndex].Cells[13].Value as string;
                        int sum;
                        int cost = 0;
                        int num = 0;

                        if (dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() == "Плов")
                        {
                            cost = 250;

                        }
                        if (dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() == "Борщ")
                        {
                            cost = 50;
                        }

                        if (dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() == "Гречка")
                        {
                            cost = 150;
                        }


                        if (!string.IsNullOrEmpty(item))
                        {
                            if (int.TryParse(item, out sum))
                            {
                                if (int.TryParse(dataGridView1.Rows[e.RowIndex].Cells[13].Value.ToString(), out num))
                                {
                                    sum = num * cost;
                                    dataGridView1.Rows[e.RowIndex].Cells[14].Value = sum.ToString();
                                }
                            }
                        }
                    }

                    if (e.ColumnIndex == 15 && e.RowIndex >= 0)
                    {
                        string item = dataGridView1.Rows[e.RowIndex].Cells[15].Value as string;
                        int sum;
                        int cost = 0;
                        int num = 0;

                        if (dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() == "Плов")
                        {
                            cost = 250;

                        }
                        if (dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() == "Борщ")
                        {
                            cost = 50;
                        }

                        if (dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() == "Гречка")
                        {
                            cost = 150;
                        }


                        if (!string.IsNullOrEmpty(item))
                        {
                            if (int.TryParse(item, out sum))
                            {
                                if (int.TryParse(dataGridView1.Rows[e.RowIndex].Cells[15].Value.ToString(), out num))
                                {
                                    sum = num * cost;
                                    dataGridView1.Rows[e.RowIndex].Cells[16].Value = sum.ToString();
                                }
                            }
                        }
                    }
                }
            }

            // Считаем Итого
            if (textBoxes != null)
            {
                for (int i = 0; i < 10; i++)
                {
                    decimal totalSum = 0;

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        // Убеждаемся, что строка не является строкой для добавления новых строк
                        if (!row.IsNewRow)
                        {
                            string strValue = row.Cells[i + 7].Value as string;

                            decimal value;
                            if (decimal.TryParse(strValue, out value))
                            {
                                // Добавляем значение к сумме
                                totalSum += value;
                            }
                        }
                    }

                    textBoxes[i].Text = totalSum.ToString();
                }
            }






        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Form2 form2 = new Form2();
            form2.StartPosition = FormStartPosition.CenterScreen;
            form2.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string relativeTemplatePath = "Temp.xlsx"; // или "Templates\\Template.xlsx", если файл в подпапке Templates
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, relativeTemplatePath);

            if (!File.Exists(templatePath))
            {
                MessageBox.Show($"Файл шаблона не найден по пути: {templatePath}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(templatePath)))
                {
                    // Получаем доступ к нужному листу (по имени или индексу)
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["Str1"];

                    // Начинаем заполнять файл

                    worksheet.Cells["A7"].Value = comboBox1.Text;
                    worksheet.Cells["CA7"].Value = textBoxOrgan.Text;
                    worksheet.Cells["A9"].Value = comboBox2.Text;
                    worksheet.Cells["CA10"].Value = comboBox1.Text;
                    worksheet.Cells["A7"].Value = comboBox1.Text;


                    // 5. Сохраняем изменения в новый файл
                    excelPackage.SaveAs(new FileInfo("Out.XLS"));

                    MessageBox.Show($"Данные успешно экспортированы в Excel файл. ", "Экспорт завершен", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Excel: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
