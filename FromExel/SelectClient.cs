using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FromExel
{
    public partial class SelectClient : Form
    {
        /// <summary>
        /// Набор контактов
        /// </summary>
        Contacts cons;
        
        /// <summary>
        /// Имя, выбранного в таблице контакта
        /// </summary>
        public string SelectName = "";

        /// <summary>
        /// Отедл, выбранного в таблице контка
        /// </summary>
        public string SelectOtdel = "";

        /// <summary>
        /// email выбранного пользователя
        /// </summary>
        public string SelectEmail = "";

        public SelectClient()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Инициализация формы по контактам
        /// </summary>
        /// <param name="cs"></param>
        public void setContacts(Contacts cs)
        {
            //Цикл по всем контактам
            for(int i = 0, countCont = cs.count; i < countCont; i++)
            {

                //Добавление новой строки в форму
                dataGridView1.Rows.Add();
                //Получение новой строки
                DataGridViewRow row = dataGridView1.Rows[i];
                //Добавление данных
                Contacts.Contact c = cs[i];
                row.Cells[0].Value = c.Name;
                row.Cells[1].Value = c.Otdel;
                row.Cells[2].Value = c.email;

            }

            //Созранения указателя на набор контктов
            cons = cs;
                    
        }

        /// <summary>
        /// Филтрация строк в таблице в зависимости от заполнения полей
        /// </summary>
        void Filter()
        {

            //цикл по всем элементам таблицы и выбор первого подходящего
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                //Получение имени контакта
                string name = row.Cells[0].Value.ToString();
                //Проверка, подходит ли по шаблону
                if (textBox1.Text != null)
                    if(textBox1.Text != "")
                        if (!name.ToLower().StartsWith(textBox1.Text.ToLower())) continue;
                //Получение подразделения
                string otdel = row.Cells[1].Value.ToString();
                if (textBox2.Text != null)
                    if (textBox2.Text != "")
                    {
                        
                        if (!otdel.ToLower().StartsWith(textBox2.Text.ToLower())) continue;
                    }
                //Если совпадает
                foreach (DataGridViewRow row2 in dataGridView1.Rows) if (row2.Selected) row2.Selected = false;
                row.Selected = true;
                dataGridView1.FirstDisplayedScrollingRowIndex = row.Index;

            }


        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            Filter();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            Filter();
        }

        /// <summary>
        /// Закрытие формы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        /// <summary>
        /// Сохранение данных об выбранном элементе и закрытие формы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            //Поиск выбранной строки
            foreach(DataGridViewRow row in dataGridView1.Rows)
            {
                //Если найдена выбранная строка
                if (row.Selected)
                {
                    //Получуние из неё полей
                    SelectName = row.Cells[0].Value.ToString();
                    SelectOtdel = row.Cells[1].Value.ToString();
                    SelectEmail = row.Cells[2].Value.ToString();
                    //Конец цикла
                    break;
                }
            }

            //Сообщение об результате и заркытие формы
            DialogResult = DialogResult.OK;
            Close();
        }

        /// <summary>
        /// Обработка нажитий клавиш
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SelectClient_KeyPress(object sender, KeyPressEventArgs e)
        {
            switch (e.KeyChar)
            {

            }
        }

        /// <summary>
        /// Обработка нажитий клавиш
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SelectClient_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyData)
            {
                case Keys.Escape:
                    //Закрытие формы
                    button2.PerformClick();
                    break;

                case Keys.Enter:
                    //Выбор элемента в форме
                    button1.PerformClick();
                    break;
            }
        }
    }
}
