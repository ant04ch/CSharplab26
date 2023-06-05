using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace OOP26
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            label1.Text= "Надайте інформацію про себе:";
            label2.Text= "ПІБ:";
            label3.Text = "Ваш номер телефону:";
            label4.Text = "Ваша пошта:";
            label5.Text = "Місто та поштовий індекс:";
            label6.Text = "Вулиця де проживаєте:";
            label7.Text = "Освіта:";
            label8.Text = "Ваша ціль:";
            label9.Text = "Процесійні досягнення:";
            label10.Text = "Навички:";
            label11.Text = "Досвід роботи:";
            label12.Text = "Дата:";
            button1.Text = "Заповнити Резюме";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string templateFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "resume.dotx");

            // Переконайтеся, що файл існує
            if (File.Exists(templateFilePath))
            {
                string newFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "resume_filled.dotx");

                // Перевірка на заповненість полів
                if (string.IsNullOrWhiteSpace(textBox1.Text) ||
                    string.IsNullOrWhiteSpace(textBox2.Text) ||
                    string.IsNullOrWhiteSpace(textBox3.Text) ||
                    string.IsNullOrWhiteSpace(textBox4.Text) ||
                     string.IsNullOrWhiteSpace(textBox5.Text)||
                      string.IsNullOrWhiteSpace(textBox6.Text)||
                       string.IsNullOrWhiteSpace(textBox7.Text)||
                        string.IsNullOrWhiteSpace(textBox8.Text)||
                         string.IsNullOrWhiteSpace(textBox9.Text)||
                          string.IsNullOrWhiteSpace(textBox10.Text))
                {
                    MessageBox.Show("Будь ласка, заповніть всі поля!");
                    return;
                }

                try
                {
                    // Створюємо копію шаблонного файлу
                    File.Copy(templateFilePath, newFilePath, true);

                    // Відкриваємо новий файл в Word
                    Word.Application wordApp = new Word.Application();
                    Word.Document document = wordApp.Documents.Open(newFilePath);

                    // Заміна значень тегів на вміст відповідних текстових полів
                    ReplaceTagWithText(document, "<FIO>", textBox1.Text);
                    ReplaceTagWithText(document, "<phone>", textBox2.Text);
                    ReplaceTagWithText(document, "<mail>", textBox3.Text);
                    ReplaceTagWithText(document, "<city>", textBox4.Text);
                    ReplaceTagWithText(document, "<street>", textBox5.Text);
                    ReplaceTagWithText(document, "<education>", textBox6.Text);
                    ReplaceTagWithText(document, "<object>", textBox7.Text);
                    ReplaceTagWithText(document, "<achievements>", textBox8.Text);
                    ReplaceTagWithText(document, "<skills>", textBox9.Text);
                    ReplaceTagWithText(document, "<workskill>", textBox10.Text);
                    ReplaceTagWithText(document, "<date>", dateTimePicker1.Text);

                    // Збереження нового файлу
                    document.SaveAs(newFilePath);

                    // Закриття документа Word
                    document.Close();
                    wordApp.Quit();

                    MessageBox.Show("Резюме успішно заповнено! Новий файл: " + newFilePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Виникла помилка при заповненні резюме: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Файл 'resume.dotx' не знайдено!");
            }
        }

        private void ReplaceTagWithText(Word.Document document, string tag, string text)
        {
            Word.Find find = document.Content.Find;
            find.Text = tag;
            find.Replacement.Text = text;
            find.Execute(Replace: Word.WdReplace.wdReplaceAll);
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!((e.KeyChar >= '0') && (e.KeyChar <= '9')))
            {
                // цифра
                return;
            }
            e.Handled = true;
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0') && (e.KeyChar <= '9'))
            {
                // цифра
                return;
            }
            if (e.KeyChar == '(')
            {
                e.KeyChar = '(';
                return;
            }
            if (e.KeyChar == ')')
            {
                e.KeyChar = ')';
                return;
            }
            if (e.KeyChar == '+')
            {
                e.KeyChar = '+';
                return;
            }
            if (Char.IsControl(e.KeyChar))
            {
                // <Enter>, <Backspace>, <Esc>
                if (e.KeyChar == (char)Keys.Enter)
                    // натиснута клавіша <Enter>
                    // встановити курсор на кнопку OK
                    textBox2.Focus();
                return;
            }
            // інші символи заборонені
            e.Handled = true;
        }// Захист для поля

        private void button2_Click(object sender, EventArgs e)
        {
            string newFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "resume_filled.dotx");
            Word.Application wordApp = new Word.Application();
            Word.Document document = wordApp.Documents.Open(newFilePath);
            document.Close();
            wordApp.Quit();
        }
    }
}


