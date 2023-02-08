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
            WorkWithFiles.Start();
            MessageBox.Show("Генерация успешно завершена!", "Генерация", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }
    }
}