using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Kursovaya_rabota
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string login = textBox1.Text;
            string pass = textBox2.Text;
            string ConnectString = "Database = аис_увп; Data Source = localhost; User Id=" + login + "; Password =" + pass + ";";
            using (MySqlConnection connection = new MySqlConnection(ConnectString))
            {

                try
                {
                    connection.Open();
                    Form1 f1 = new Form1(login, pass);
                    f1.Show();
                }
                catch (Exception)
                {
                    string message = "Логин или пароль не верны";
                    string caption = "Ошибка,блет";
                    MessageBoxButtons buttons = MessageBoxButtons.OK;
                    DialogResult result;

                    // Displays the MessageBox.
                    result = MessageBox.Show(message, caption, buttons);
                    if (result == System.Windows.Forms.DialogResult.Yes)
                    {
                        // Closes the parent form.
                        this.Close();
                    }
                }
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
