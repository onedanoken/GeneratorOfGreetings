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
            MessageBox.Show("��� ���������...\n������� ��, ����� ����������", "���������", MessageBoxButtons.OK, MessageBoxIcon.Information);
            WorkWithFiles.Start();
            MessageBox.Show("��������� ������� ���������!", "���������", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }
    }
}