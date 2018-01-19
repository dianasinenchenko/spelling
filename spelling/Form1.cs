using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace spelling
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var Ворд1 = new Microsoft.Office.Interop.Word.Application();
            Ворд1.Visible = false;
            // Открываем новый документ MS Word:
            Ворд1.Documents.Add();
            // Копируем содержимое текстового окна в документ
            Ворд1.Selection.Text = textBox1.Text;
            // Непосредственная проверка орфографии:
            Ворд1.ActiveDocument.CheckSpelling();
            // Копируем результат назад в текстовое поле
            textBox1.Text = Ворд1.Selection.Text;
            Ворд1.Documents.Close(false);
            // Закрыть документ Word без сохранения:
            Ворд1.Quit();
            Ворд1 = null;






        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Clear(); button1.Text = "Проверка орфографии";



        }



    }
}
