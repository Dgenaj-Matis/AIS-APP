using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using GemBox.Spreadsheet;
using System.IO;
using System.Diagnostics;

namespace Kursovaya_rabota
{
    public partial class Form1 : Form
    {
        DataSet ds;
        MySqlDataAdapter dataAdapter;
        public static string ConnectString;
        string CommandText = "select * from аис_увп.сотрудник";
        
        public Form1(string login, string pass)
        {
            InitializeComponent();
            string ConnectString = "Database = аис_увп; Data Source = localhost; User Id=" + login + "; Password =" + pass + ";";
            Form1.ConnectString = ConnectString;
            this.StartPosition = FormStartPosition.CenterParent;
            this.dataGridView1.BackgroundColor = Color.CornflowerBlue;

            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;  //устанавливаем полное выделение строки 
            button5.Enabled = false;
            button1.Enabled = false;

            using (MySqlConnection connection = new MySqlConnection(ConnectString))
            {
                connection.Open();
                dataAdapter = new MySqlDataAdapter(CommandText, connection);

                ds = new DataSet();

                MySqlCommandBuilder bulder = new MySqlCommandBuilder(dataAdapter);
                dataAdapter.UpdateCommand = bulder.GetUpdateCommand();
                dataAdapter.InsertCommand = bulder.GetInsertCommand();
                dataAdapter.DeleteCommand = bulder.GetDeleteCommand();
                dataAdapter.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
            }
        }


        private void Button2_Click(object sender, EventArgs e) //Удаление выделенной строки 
        {
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                dataGridView1.Rows.Remove(row);
            }
        }

        private void Button3_Click(object sender, EventArgs e) //Сохранение данных в БД 
        {
                using (MySqlConnection connection = new MySqlConnection(ConnectString))
                {
                dataAdapter.Update(ds.Tables[0]);
                }
        }

        int numberRows = 0;

        private void Button4_Click(object sender, EventArgs e) //Запись в Calc
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile ef = new ExcelFile();
            ExcelWorksheet ws = ef.Worksheets.Add("Writing");
            MySqlConnection connection = new MySqlConnection(ConnectString);
            try
            {
                connection.Open();
                MySqlCommand command = new MySqlCommand(CommandText, connection);
                MySqlDataReader reader = command.ExecuteReader();
                int i = 0, j;
                while (reader.Read())
                {
                    for (j = 0; j < 9; j++)
                    {
                        ws.Cells[i, j].Value = reader[j];
                    }
                    i++;
                }
                ws.Cells.GetSubrangeAbsolute(0, 0, i - 1, 9).Sort(false).By(0, false).Apply();
                ef.Save("Премии.ods");
                connection.Close();
                MessageBox.Show("Данные успешно записаны в файл и отсортированы. Файл находится в Проекте.");
                Process.Start("Премии.ods");

                numberRows = i;

                if (numberRows > 1) { button5.Enabled = true; } //Если строк больше 1, используется форма Т-11а
                else { button1.Enabled = true; }                //Обратно - Т-11
            }
            catch
            {
                MessageBox.Show("Ошибка");
            }
        }

        string s = DateTime.Now.ToString("dd MMMM yyyy");   //Вывод даты заполнения документа

        private void Button5_Click(object sender, EventArgs e) //Запись в .odt файл формы Т-11а
        {
            try
            {
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
                ExcelFile ef = ExcelFile.Load("Премии.ods");
                ExcelWorksheet ws = ef.Worksheets[0];

                string getStringOutput, tmp;
                int j;
                File.Delete("Т-11а.odt");

                StreamWriter writer1 = new StreamWriter("Т-11а.odt", true);     //Шапка документа
                writer1.WriteLine("\t\t\t\t\t\t\t Унифицированная форма №Т-11а\n" + "\t\t\t\t\t\t\t Утверждена Постановлением Госкомстата России\n" + "\t\t\t\t\t\t\t от 05.01.2004 №1");
                writer1.WriteLine("\t\t\t\t\t\t\t\t\t\t Код\n" + "\t\t\t\t\t\t\t Форма   по ОКУД\t0301027\n" + " _______________________________________________________ \t по ОКПО \n");
                writer1.WriteLine("\t\t\t\t\t\t\t\t Номер документа \t Дата составления\n");
                writer1.WriteLine("\t\t\t\t\t       ПРИКАЗ  \t\t\t\t\t    " + s);
                writer1.WriteLine("\t\t\t\t\t (распоряжение)");
                writer1.WriteLine("\t\t\t\t      о поощрении работников\n");
                writer1.WriteLine(" ______________________________________________________________________________________________ \n");
                writer1.WriteLine(" ______________________________________________________________________________________________ \n");
                writer1.WriteLine(" ______________________________________________________________________________________________ \n");
                writer1.Close();
                
                for (int i = 0; i < numberRows; i++)
                {
                    getStringOutput = "";
                    for (int k = 0; k < 9; k++)
                    {
                        getStringOutput += ws.Cells[i, k].Value.ToString();
                        tmp = ws.Cells[i, k].Value.ToString();
                        switch (k)
                        {
                            case 0:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += "";
                                    }
                                    break;
                                }
                            case 1:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += "";
                                    }
                                    break;
                                }
                            case 2:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += "";
                                    }
                                    break;
                                }
                            case 3:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += "";
                                    }
                                    break;
                                }
                            case 4:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += "";
                                    }
                                    break;
                                }
                            case 5:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += "";
                                    }
                                    break;
                                }
                            case 6:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += "";
                                    }
                                    break;
                                }
                            case 7:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += "";
                                    }
                                    break;
                                }
                            case 8:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += "";
                                    }
                                    break;
                                }
                            case 9:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += "";
                                    }
                                    break;
                                }
                        }
                        getStringOutput += "    ";
                    }

                        StreamWriter writer = new StreamWriter("Т-11а.odt", true);
                        writer.WriteLine(getStringOutput + "\n");
                        writer.Close();
                }

                StreamWriter writer2 = new StreamWriter("Т-11а.odt", true);   //Конец документа
                writer2.WriteLine();
                writer2.WriteLine();
                writer2.WriteLine();
                writer2.WriteLine("Основание: представление");
                writer2.WriteLine(" ______________________________________________________________________________________________ \n");
                writer2.WriteLine(" ______________________________________________________________________________________________ \n");
                writer2.WriteLine();
                writer2.WriteLine();
                writer2.WriteLine("Руководитель организации _________________________________ ____________ _________________");
                writer2.Close();
                
                MessageBox.Show("Данные успешно записаны в файл для печати. Файл находится в Проекте.");
                Process.Start("Т-11а.odt");
            }
            catch
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void button1_Click(object sender, EventArgs e) //Запись в .odt файл формы Т-11
        {
            try
            {
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
                ExcelFile ef = ExcelFile.Load("Премии.ods");
                ExcelWorksheet ws = ef.Worksheets[0];

                string getStringOutput, tmp;
                int j;
                File.Delete("Т-11.odt");

                StreamWriter writer1 = new StreamWriter("Т-11.odt", true);
                writer1.WriteLine("\t\t\t\t\t\t\t Унифицированная форма №Т-11\n" + "\t\t\t\t\t\t\t Утверждена Постановлением Госкомстата России\n" + "\t\t\t\t\t\t\t от 05.01.2004 №1");
                writer1.WriteLine("\t\t\t\t\t\t\t\t\t\t Код\n" + "\t\t\t\t\t\t\t Форма   по ОКУД\t0301026\n" + " _______________________________________________________ \t по ОКПО \n");
                writer1.WriteLine("\t\t\t\t\t\t\t\t Номер документа \t Дата составления\n");
                writer1.WriteLine("\t\t\t\t\t       ПРИКАЗ  \t\t\t\t\t    " + s);
                writer1.WriteLine("\t\t\t\t\t (распоряжение)");
                writer1.WriteLine("\t\t\t\t      о поощрении работника\n");
                writer1.Close();

                getStringOutput = "";
                for (int i = 0; i < numberRows; i++)
                {
                    for (int k = 0; k < 9; k++)
                    {
                        getStringOutput += ws.Cells[i, k].Value.ToString();
                        tmp = ws.Cells[i, k].Value.ToString();
                        switch (k)
                        {
                            case 0:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += "";
                                    }
                                    break;
                                }
                            case 1:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += "";
                                    }
                                    break;
                                }
                            case 2:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += "";
                                    }
                                    break;
                                }
                            case 3:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += "";
                                    }
                                    break;
                                }
                            case 4:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += "";
                                    }
                                    break;
                                }
                            case 5:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += "";
                                    }
                                    break;
                                }
                            case 6:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += "";
                                    }
                                    break;
                                }
                            case 7:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += "";
                                    }
                                    break;
                                }
                            case 8:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += "";
                                    }
                                    break;
                                }
                            case 9:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += "";
                                    }
                                    break;
                                }
                        }
                        getStringOutput += "\n";
                    }
                    StreamWriter writer = new StreamWriter("Т-11.odt", true);
                    writer.WriteLine("\t\t" + getStringOutput);
                    writer.Close();
                }                   

                StreamWriter writer2 = new StreamWriter("Т-11.odt", true);  //Конец документа
                writer2.WriteLine(" В сумме ______________________________________________________________________________________ \n");
                writer2.WriteLine(" ________________________________________________________________ руб. ___________________ коп. \n");
                writer2.WriteLine();
                writer2.WriteLine();
                writer2.WriteLine();
                writer2.WriteLine("Основание: представление");
                writer2.WriteLine(" ______________________________________________________________________________________________ \n");
                writer2.WriteLine(" ______________________________________________________________________________________________ \n");
                writer2.WriteLine();
                writer2.WriteLine();
                writer2.WriteLine("Руководитель организации _________________________________ ____________ _________________");
                writer2.WriteLine("С приказом (распоряжением) работник ознакомлен _______________ '___' ____________ 20__ г.");
                writer2.Close();
                
                MessageBox.Show("Данные успешно записаны в файл для печати. Файл находится в Проекте.");
                Process.Start("Т-11.odt");
            }
            catch
            {
                MessageBox.Show("Ошибка");
            }
        }

        //КАЛЬКУЛЯТОР
        //Цифры
        private void button15_Click(object sender, EventArgs e)
        {
            textBox5.Text = textBox5.Text + 0;
        }
        private void button12_Click(object sender, EventArgs e)
        {
            textBox5.Text = textBox5.Text + 1;
        }
        private void button13_Click(object sender, EventArgs e)
        {
            textBox5.Text = textBox5.Text + 2;
        }
        private void button14_Click(object sender, EventArgs e)
        {
            textBox5.Text = textBox5.Text + 3;
        }
       private void button9_Click(object sender, EventArgs e)
        {
            textBox5.Text = textBox5.Text + 4;
        }
        private void button10_Click(object sender, EventArgs e)
        {
            textBox5.Text = textBox5.Text + 5;
        }
        private void button11_Click(object sender, EventArgs e)
        {
            textBox5.Text = textBox5.Text + 6;
        }
        private void button6_Click(object sender, EventArgs e)
        {
            textBox5.Text = textBox5.Text + 7;
        }
        private void button7_Click(object sender, EventArgs e)
        {
            textBox5.Text = textBox5.Text + 8;
        }
        private void button8_Click(object sender, EventArgs e)
        {
            textBox5.Text = textBox5.Text + 9;
        }

        private void button16_Click(object sender, EventArgs e) //Кнопка C
        {
             textBox5.Text = "";
             label6.Text = "";
        }
        private void button18_Click(object sender, EventArgs e) //Умножение
        {
            a = float.Parse(textBox5.Text);
            textBox5.Clear();
            count = 1;
            label6.Text = a.ToString() + "*";
        }
        private void button19_Click(object sender, EventArgs e) //Деление
        {
            a = float.Parse(textBox5.Text);
            textBox5.Clear();
            count = 2;
            label6.Text = a.ToString() + "/";
        }
        private void button17_Click_1(object sender, EventArgs e) //Удаление 1 символа
        {
            int lenght = textBox5.Text.Length - 1;
            string text = textBox5.Text;
            textBox5.Clear();
            for (int i = 0; i < lenght; i++)
            {
                textBox5.Text = textBox5.Text + text[i];
            }
        }       
        private void button20_Click(object sender, EventArgs e) //Равно
        {
            calculate();
            label6.Text = "";
        }
        //РАСЧЕТЫ
        float a, b;
        int count;
        private string login;
        private string pass;

        private void calculate()
        {
            switch (count)
            {
                case 1:
                    b = a * float.Parse(textBox5.Text);
                    textBox5.Text = b.ToString();
                    break;
                case 2:
                    if (float.Parse(textBox5.Text) != 0)
                    {
                        b = a / float.Parse(textBox5.Text);
                        textBox5.Text = b.ToString();
                    }
                    else { MessageBox.Show("И зачем оно тебе?"); }
                    break;
                case 3:
                    b = a * float.Parse(textBox5.Text);
                    b /= 100;
                    b = b - b * 13/100;
                    textBox5.Text = b.ToString();
                    break;
                case 4:
                    b = a * 3;
                    b = b * float.Parse(textBox5.Text);
                    b /= 100;
                    b = b - b * 13 / 100;
                    textBox5.Text = b.ToString();
                    break;
                case 5:
                    b = a * 12;
                    b = b * float.Parse(textBox5.Text);
                    b /= 100;
                    b = b - b * 13 / 100;
                    textBox5.Text = b.ToString();
                    break;
                default:
                    break;
            }
        }

        //РАСЧЁТ ПРЕМИЙ
        private void button21_Click(object sender, EventArgs e) //Eжемесячная премия
        {
            a = float.Parse(textBox5.Text);
            textBox5.Clear();
            count = 3;
            label6.Text = a.ToString() + "*";
        }
        private void button22_Click(object sender, EventArgs e) //Ежеквартальная премия
        {
            a = float.Parse(textBox5.Text);
            textBox5.Clear();
            count = 4;
            label6.Text = a.ToString() + "*";
        }
        private void button23_Click(object sender, EventArgs e) //Годная премия
        {
            a = float.Parse(textBox5.Text);
            textBox5.Clear();
            count = 5;
            label6.Text = a.ToString() + "*";
        }

        //СПРАВОЧНИКИ
        private void премииЗаВыслугуЛетToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Стаж, лет \t Размер премии, в %" +
                "\n" +
                "\n 1-5 \t\t\t 10" +
                "\n 5-10 \t\t\t 15" +
                "\n 10-15 \t\t\t 20" +
                "\n 15 и более \t\t 30");
        }
        private void премииЗаГостайнуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Степень секретности \t\t Размер премии, в %" +
                "\n" +
                "\n 'Секретно' без проверок \t\t\t 5-10" +
                "\n 'Секретно' с проверками \t\t\t 30-50" +
                "\n Совершенно секретно \t\t\t 30-50" +
                "\n Особой важности \t\t\t 50-70");
        }
        private void премииЗаОсобоВажныеЗаданияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Группа должности \t Размер премии, в %" +
                "\n" +
                "\n Младшая \t\t\t 0-60" +
                "\n Старшая \t\t\t 10-65" +
                "\n Ведущая \t\t\t 20-70" +
                "\n Главная \t\t\t 30-80" +
                "\n Высшая \t\t\t\t 50-130");
        }

        //МАНУАЛ
        private void работаСБазойДанныхToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("   ВСЕ поля ОБЯЗАТЕЛЬНО должны быть заполнены!\n" +
                "   Если какой-то из перечисленных далее столбцов не будет учитываться при расчёте премий, нужно в соответствующей строке поставить значение '0': " +
                "'Ежемесячная премия', 'Ежеквартальная премия', 'Ежегодная премия', 'Премия за фактически отработанное время'.\n" +
                "   Для того чтобы удалить строку, нужно выбрать её, нажав на нужном уровне ячейку крайнего левого столбика, затем нажать кнопку 'Удалить'.\n" +
                "   После каждого внесённого изменения для их сохранения необходимо нажать кнопку 'Сохранить'.");
        }
        private void выводФормToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("   Прежде чем вывести формы, необходимо сначала вывести табличный документ 'Премии.ods', нажав кнопку 'Вывод данных в таблицу'.\n" +
                "   После нажатия откроется табличный документ, его можно закрыть.\n" +
                "   Далее необходимо нажать на 'Т-11а' или 'Т-11'. После нажатия форма автоматически откроется.\n" +
                "   В зависимости от количества внесённых сотрудников программой будет автоматически определяться какая из форм должна использоваться.\n");
        }
        private void расчетПремийToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("   Для того чтобы расчитать премию (ежемесячную, ежеквартальную, годовую) необходимо:\n" +
                "   1) ввести в поле калькулятора оклад сотрудника;\n" +
                "   2) нажать соответствующую с рассчитываемой премией клавишу ('Ежемесячная', 'Ежеквартальная', 'Годовая');\n" +
                "   3) ввести в поле калькулятора процент премии, предусмотренный локальным актом;\n" +
                "   4) нажать кнопку '='.\n");
        }

        
    }
}
