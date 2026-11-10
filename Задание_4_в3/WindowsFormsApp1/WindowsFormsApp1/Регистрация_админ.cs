using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Регистрация_админ : Form
    {
        public Регистрация_админ()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.Owner != null)
            {

                this.Owner.Show();
            }


            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string inputName = textBox1.Text.Trim();
            string inputPassword = textBox2.Text.Trim();

            if (string.IsNullOrEmpty(inputName) || string.IsNullOrEmpty(inputPassword))
            {
                MessageBox.Show("Пожалуйста, введите имя и пароль!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // 1. Создаем экземпляр DataSet библиотеки
                WindowsFormsApp1.БД_БиблеотекаDataSet dataSet = new WindowsFormsApp1.БД_БиблеотекаDataSet();

                // 2. Создаем адаптер именно для таблицы Kod_Vhod
                var adapter = new WindowsFormsApp1.БД_БиблеотекаDataSetTableAdapters.Kod_VhodTableAdapter();

                // 3. Заполняем таблицу данными из базы
                adapter.Fill(dataSet.Kod_Vhod);

                bool isAuthorized = false;

                // 4. Перебираем строки таблицы авторизации
                foreach (WindowsFormsApp1.БД_БиблеотекаDataSet.Kod_VhodRow row in dataSet.Kod_Vhod.Rows)
                {
                    // Проверяем логин и пароль по точным названиям столбцов из вашей БД_Библиотека
                    if (row.Имя_администратора.ToString() == inputName && row.Код_входа.ToString() == inputPassword)
                    {
                        isAuthorized = true;
                        break;
                    }
                }

                if (isAuthorized)
                {
                    MessageBox.Show("Авторизация успешна!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // Открываем вашу форму меню (убедитесь, что класс формы называется именно МенюАдмина или скорректируйте название)
                    Меню_Админ adminMenu = new Меню_Админ();
                    adminMenu.Show();

                    this.Hide();
                }
                else
                {
                    MessageBox.Show("Неверное имя пользователя или пароль.", "Ошибка доступа", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при работе с DataSet: " + ex.Message, "Ошибка БД", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            // 1. Проверяем заполнение полей
            if (string.IsNullOrWhiteSpace(textBox1.Text) || string.IsNullOrWhiteSpace(textBox2.Text))
            {
                MessageBox.Show("Пожалуйста, заполните все поля!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 2. Строка подключения для Access (поддерживает сетевые пути)
            // Внимание: проверьте расширение вашей БД. Если файл старый (.mdb), замените ACE.OLEDB.12.0 на JET.OLEDB.4.0
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\192.168.3.250\Veda\3 курс\ИСП 31\Отчеты ИСП 31 УП ПМ.03\Ильиных\Практика\Задание_4_в3\WindowsFormsApp1\WindowsFormsApp1\bin\x64\Debug\БД_Библеотека.accdb;";

            // 3. SQL-запрос (Пароль — зарезервированное слово в Access, поэтому имя столбца лучше взять в квадратные скобки)
            // Замените Имя_Пользователя и [Пароль] на точные названия столбцов из вашей таблицы Kod_Vhod
            string query = "INSERT INTO Kod_Vhod (Имя_администратора, [Код_входа]) VALUES (?, ?)";

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                using (OleDbCommand command = new OleDbCommand(query, connection))
                {
                    // В OleDb Access важен порядок добавления параметров, а не их имена!
                    command.Parameters.AddWithValue("?", textBox1.Text.Trim());
                    command.Parameters.AddWithValue("?", textBox2.Text.Trim());

                    try
                    {
                        connection.Open();
                        int result = command.ExecuteNonQuery();

                        if (result > 0)
                        {
                            MessageBox.Show("Администратор успешно зарегистрирован!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            textBox1.Clear();
                            textBox2.Clear();
                        }
                        else
                        {
                            MessageBox.Show("Не удалось сохранить данные.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка при работе с базой данных Access: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
    }
}
