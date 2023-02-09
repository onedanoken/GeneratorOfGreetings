using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace generatorForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Идёт генерация...\nНажмите ОК, чтобы продолжить", "Генерация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            int code = WorkWithFiles.Start();
            switch (code)
            {
                case 0:
                    MessageBox.Show("Генерация успешно завершена!", "Генерация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
                case 1:
                    MessageBox.Show("Ошибка генерации!\nПроблема с конфигурационным файлом, проверьте его", "Генерация", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                case 2:
                    MessageBox.Show("Ошибка генерации!\nНе хватает поздравлений для генерации триад.\nДополните список поздравлений или добавьте новую группу.", "Генерация", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                case 3:
                    MessageBox.Show("Ошибка генерации!\nВозникла проблема с работой с документами.\nПроверьте наличие шаблона или выходной папки.", "Генерация", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            } 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string path = @"C:\СЯиТП\лаба 1\out";
            if (!Directory.Exists(path))
            {
                MessageBox.Show("Директория ещё не создана.\nДля создания директории произведите генерацию поздравлений.",
                    "Директория поздравлений", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Process.Start("explorer.exe", path);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string path = @"C:\СЯиТП\лаба 1\out";
            if (!Directory.Exists(path))
            {
                MessageBox.Show("Директория ещё не создана.\nДля создания директории произведите генерацию поздравлений.",
                    "Директория поздравлений", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                var latestFile = new DirectoryInfo(path)
                    .GetFiles()
                    .OrderByDescending(f => f.LastWriteTime)
                    .First();
                Word.Application result = new Microsoft.Office.Interop.Word.Application();
                result.Documents.Open(latestFile.FullName);
                result.Visible = true;
                

                return;
            }
            catch 
            {
                MessageBox.Show("Сгенерированных поздравлений ещё не существует.\nЗапустите генерацию поздравлений.",
                    "Файл с поздравлениями", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
    }
}