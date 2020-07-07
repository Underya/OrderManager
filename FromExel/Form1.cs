//Автор: Матвеев Никита Иванонивч
//Проект использует набор библиотек Microsoft Office, для работы с Word, Exel и Outlouck 

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Windows.Forms;

namespace FromExel
{
    public partial class Form1 : Form
    {
        /// <summary>
        /// Файл конфигурации
        /// </summary>
        SetupConfig conf;

        /// <summary>
        /// Для сохранения заказов
        /// </summary>
        saveData saveData = null;

        /// <summary>
        /// Список контактов
        /// </summary>
        Contacts contacts;

        /// <summary>
        /// Класс для работы со списком заказов 
        /// </summary>
        OrderWatch orderWatch;

        /// <summary>
        /// Константа, которая указывает количество строчек в файле заказа
        /// По ней выводится количество строк для ввода на экран 
        /// </summary>
        int countRow;

        public Form1()
        {
            //Инициализация компонентов, созданных Visual Studio
            InitializeComponent();

            //Загрузка начальных параметров
            try
            {
                //Включение EXEL приложения
                ExelManager.Start();

                //Создание объекта, который предоставляет настройки из файла
                conf = new SetupConfig();

                //Получение параметров, записаных в файл
                ReceiveConfFromFile();

                //Объект, с помощью которого происходит сохранение заказа
                saveData = new ZakazManager(conf);

                //Добавление строчек на экран таблицы
                addRowInTable(countRow);

                orderWatch = new OrderWatch(conf["orderscore"]);
            }
            //Внутренняя ошибка
            catch (ExceptionError e2)
            {
                ShowError(e2);
                Close();
                //Закрытие приложения на всякий случай
                ExelManager.Quit();
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message);
                Close();
                ExelManager.Quit();
            }
        }

        /// <summary>
        /// Добавление строчек в табличку, которую видят пользователи
        /// </summary>
        /// <param name="CountRow"></param>
        private void addRowInTable(int CountRow)
        {
            //Добавление числа строк в тбалицу согласно параметру
            dataGridView1.Rows.Add(CountRow);
            //Заполнение начальных данных каждой строки
            for (int i = 1; i <= CountRow; i++)
            {
                //Получение I строки
                DataGridViewRow row = dataGridView1.Rows[i - 1];
                //Заполнение номера строки
                row.Cells[0].Value = i;
                //Заполнение возможных форматов
                DataGridViewComboBoxCell cellB = (DataGridViewComboBoxCell)row.Cells[2];
                Resourse.Format.SetFormatToCell(cellB);
                //Заполение других ячеек формы пустыми строка для избежания ошибки
                row.Cells[3].Value = "";
                row.Cells[4].Value = "";
                row.Cells[5].Value = "";

            }
        }

        /// <summary>
        /// Получение конфигурации, которая записанна в файле
        /// </summary>
        private void ReceiveConfFromFile()
        {
            
            //Инциализация форматов
            Resourse.Format.inital(conf["attname"]);
            //Добавление контактов
            contacts = new Contacts(conf["addres"]);
            //Добавление строк в таблицу
            countRow = int.Parse(conf["countwork"]);
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// Метод для вывода ошибки на экран
        /// </summary>
        /// <param name="error"></param>
        void ShowError(ExceptionError error)
        {
            MessageBox.Show(error.innerNumber + ':' + error.innerText + ".\r\n" + "Location:" + error.location);
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// Вывод текста для пользователя
        /// </summary>
        /// <param name="textLabel"></param>
        private void SetLabel(string textLabel)
        {
            StatusLabel.Text = textLabel;
        }

        /// <summary>
        /// Проверка, все ли поля заказчика заполнены
        /// </summary>
        /// <returns></returns>
        private bool CheckClientFieldFill()
        {
            // Проверка, все ли небходимые поля заполнены
            if (textBox1.Text == "" || textBox2.Text == "") return false;
            return true;
        }

        /// <summary>
        /// Формирование заказа
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormZak_Click(object sender, EventArgs e)
        {
            //Проверка, все ли небходимые поля заполнены
            if (CheckClientFieldFill())
            {
                MessageBox.Show("У заказчика не все поля заполнены");
                return;
            }

            //Установка значка в стиле формирование заказа
            SetLabel("Формирование \r\n заявки");

            try
            {
                //Начало работы с формою
                saveData.StartWork();
                //Указания клиента
                saveData.SetClient(textBox1.Text, textBox2.Text);
                //Указание руководства
                if (radioButton1.Checked) saveData.SetMain(1); else saveData.SetMain(2);
                //Добавление данных из таблицы
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    
                    string name = "", form = "";
                    int count = -1, exem = -1, sum = -1;
                    //Формирование из каждой строки таблицы, с которой работает пользователь
                    //Параметров с информацией о заказе, которые перейдут в метод
                    if (row.Cells[1].Value != null) name = row.Cells[1].Value.ToString();
                    if (row.Cells[2].Value != null) form = row.Cells[2].Value.ToString();
                    if (row.Cells[3].Value != null) if (row.Cells[3].Value.ToString() != "") count = int.Parse(row.Cells[3].Value.ToString());
                    if (row.Cells[4].Value != null) if (row.Cells[4].Value.ToString() != "") exem = int.Parse(row.Cells[4].Value.ToString());
                    if (row.Cells[5].Value != null) if (row.Cells[5].Value.ToString() != "") sum = int.Parse(row.Cells[5].Value.ToString());
                    saveData.AddWork(name, form, count, exem, sum);
                }


                //Установка номера
                saveData.SetNumberOrder(saveData.GetLastFileNumber());

                //Конец работы, заявка сформирована
                saveData.EndWork();
                //Сохрание заявки в таблицу
                orderWatch.addOrder(saveData.GetLastFileNumber(), saveData.GetLastFileName());

                //Попытка отправить письмо
                TrySendMail();
                               
                MessageBox.Show("Заявка успешно сформирован");
                //Скрыть метку
                SetLabel("");
            }
            catch (ExceptionError err)
            {
                //Скрыть метку
                SetLabel("");
                ShowError(err);
                //Закрытие текущей книги
                saveData.Cancel();
            }
            catch(Exception err)
            {
                //Скрыть метку
                SetLabel("");
                MessageBox.Show(err.Message);
                saveData.Cancel();
            }
        }

        /// <summary>
        /// Попытка отправить письмо
        /// </summary>
        private void TrySendMail()
        {
            //Если указан соотвествующий флаг
            //Отправление письма по почте
            if (checkBox3.Checked)
            {
                SetLabel("Управление \r\n почтой");

                //Если поле пустое
                if (textBox_email.Text == "")
                {
                    MessageBox.Show("Письмо не отправлено, так как не указан адрес");

                }

                //Полный путь к отправляемому файлу
                string fullPath = conf["savepath"] + saveData.GetLastFileNumber() + " " + saveData.GetLastFileName() + ".xlsx";

                //Если поле не пустое, отправка письма
                MailManager mm = new MailManager();
                mm.SendMail(textBox_email.Text, fullPath);
            }
        }

        /// <summary>
        /// очищение всех заполненных полей
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.Cells[1].Value = null;
                row.Cells[2].Value = null;
                row.Cells[3].Value = null;
                row.Cells[4].Value = null;
                row.Cells[5].Value = null;
            }
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "0";
            checkBox1.Checked = false;
            checkBox2.Checked = false;
        }

        /// <summary>
        /// Вывод окна для выбора контактов
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click_1(object sender, EventArgs e)
        {
            //иниициализация и показ формы
            SelectClient sc = new SelectClient();
            sc.setContacts(contacts);
            sc.ShowDialog();
            //Если был выбран пользователь
            if(sc.DialogResult == DialogResult.OK)
            {
                //Заполнение соответсвующих полей в форме
                textBox2.Text = sc.SelectName;
                textBox1.Text = sc.SelectOtdel;
                textBox_email.Text = sc.SelectEmail;
            }
           
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// Рассчёт числка листов в столбце калькулятора
        /// </summary>
        void Calculate()
        {
            //попытка разобрать все имеющиеся числа
            int countInEx = 0;
            if (textBox3.Text == "") return;
            //Если не удалось получить результат
            if(!int.TryParse(textBox3.Text, out countInEx))
            {
                MessageBox.Show("Не правильное значение");
                textBox3.Text = "0";
                textBox6.Text = "0";
                return;
            }

            int countEx = 0;
            if (textBox4.Text == "") return;
            //Если не удалось получить результат
            if (!int.TryParse(textBox4.Text, out countEx))
            {
                MessageBox.Show("Не правильное значение");
                textBox4.Text = "0";
                textBox6.Text = "0";
                return;
            }

            int Dop = 0;
            if (textBox5.Text == "") return;
            if (!int.TryParse(textBox5.Text, out Dop))
            {
                MessageBox.Show("Не правильное значение");
                textBox5.Text = "0";
                textBox6.Text = (countEx * countInEx).ToString();
                return;
            }

            //Сумма 
            int summ = countEx * countInEx;

            //Действия в зависимости от чек боксов
            if (checkBox2.Checked || checkBox1.Checked)
            {
                //Если оба
                if (checkBox1.Checked && checkBox2.Checked)
                {
                    int temp = summ / 2 + summ % 2;
                    summ = temp / 2 + temp % 2;

                } else summ = summ / 2 + summ % 2;
            }

            summ += Dop;

            //Если во все поля введены числа
            textBox6.Text = summ.ToString();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            Calculate();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            Calculate();
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            Calculate();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            Calculate();
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            Calculate();
        }

        /// <summary>
        /// Обработка нажайти клавиш
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyData)
            {
                //Выход из формы
                case Keys.Escape:

                    DialogResult res = MessageBox.Show("Закрыть программу?", "Выход", MessageBoxButtons.YesNo);
                    if (res == DialogResult.Yes) Close();
                    break;

                    //Открытие окна с выбором клиентов
                case Keys.F1:
                    button1.PerformClick();
                    break;

                    //Переключение ответсвенного
                case Keys.F2:
                    if (radioButton1.Checked) radioButton2.Checked = true; else radioButton1.Checked = true;
                    break;

                    //Переход к окну калькулятора
                case Keys.F3:
                    textBox3.Focus();
                    break;

                    //Выбор и заполнение первой ПУСТОЙ строки заказа
                case Keys.F4:
                    {

                        //Поиск пустой строки
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            //Проверка на пустоту
                            if (row.Cells[3].Value != null) if (row.Cells[3].Value.ToString() != "") continue;
                            if (row.Cells[4].Value != null) if (row.Cells[4].Value.ToString() != "") continue;
                            if (row.Cells[5].Value != null) if (row.Cells[5].Value.ToString() != "") continue;

                            //Если нашли, то занесение значений из калькуятора
                            row.Cells[3].Value = textBox3.Text;
                            row.Cells[4].Value = textBox4.Text;
                            row.Cells[5].Value = textBox6.Text;

                            //Очищение калькулятора
                            textBox3.Text = "";
                            textBox4.Text = "";
                            textBox5.Text = "0";
                            textBox6.Text = "0";

                            checkBox1.Checked = false;
                            checkBox2.Checked = false;

                            //Установка фокуса на ввод названия работы
                            dataGridView1.Focus();
                            dataGridView1.CurrentCell = row.Cells[1];

                            //Конец поиска
                            break;

                        }

                        break;
                    }

                //Клавиши для хождение по калькулятору
                case Keys.Up:
                    {
                        //В зависимости от того, какой элемент в фокусе, выбор другого элемента
                        if (textBox4.Focused) { textBox3.Focus(); break; }
                        if (textBox5.Focused) { textBox4.Focus(); break; }
                        if (textBox6.Focused) textBox5.Focus();
                        
                        break;
                    }

                //Клавиши для хождение по калькулятору
                case Keys.Down:
                    {
                        if (textBox3.Focused) { textBox4.Focus(); break; }
                        if (textBox4.Focused) { textBox5.Focus(); break; }
                        if (textBox5.Focused) { checkBox1.Focus(); }
                        break;
                    }

                    //Ввод email
                case Keys.F1 | Keys.Control:
                    {
                        //Смена флага
                        //Переход фокуса на флажок
                        checkBox3.Checked = !checkBox3.Checked;
                        checkBox3.Focus();
                        //Если поле ввода пустое, то оно получает фокус
                        if (textBox_email.Text == "") textBox_email.Focus();
                        break;
                    }

                    //Очистка всех полей
                case Keys.Control | Keys.F4:
                    {
                        button2.PerformClick();
                        break;
                    }

                    //Формирование заказа
                case Keys.Control | Keys.Enter:
                    {
                        FormZak.PerformClick();
                        break;
                    }

            }
        }

        /// <summary>
        /// Вывод папки с заказами
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void заказыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start(conf["savepath"]);
        }

        /// <summary>
        /// Старый цвет элементов управления
        /// </summary>
        Color chekBoxOldColor;

        private void checkBox1_Enter(object sender, EventArgs e)
        {
            chekBoxOldColor = checkBox1.BackColor;
            checkBox1.BackColor = Color.Red;
        }

        private void checkBox1_Leave(object sender, EventArgs e)
        {
            checkBox1.BackColor = chekBoxOldColor;
        }

        private void checkBox2_Enter(object sender, EventArgs e)
        {
            chekBoxOldColor = checkBox2.BackColor;
            checkBox2.BackColor = Color.Red;
        }

        private void checkBox2_Leave(object sender, EventArgs e)
        {
            checkBox2.BackColor = chekBoxOldColor;
        }

        private void историяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OrderScoreFormcs histor = new OrderScoreFormcs();
            histor.initial(conf["orderscore"], conf["savepath"]);
            histor.ShowDialog();
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// Закрытие формы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Закрытие EXEL приложения
            ExelManager.Stop();
        }

        /// <summary>
        /// Открыть папку с файлами настроек
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void настройкиToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void параметрыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start(conf["attdirector"]);
        }

        private void изменитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("config.ini");
        }
    }
}
