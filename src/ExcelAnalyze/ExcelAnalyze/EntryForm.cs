using System;
using System.Windows.Forms;
using ExcelAnalyze.Logic;

namespace ExcelAnalyze
{
    public partial class EntryForm : Form
    {
        public EntryForm()
        {
            InitializeComponent();
        }

        private void EntryForm_Load(object sender, EventArgs e) { }

        private void searchButton_Click(object sender, EventArgs e)
        {
            var filePath = string.Empty;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = Descriptor.Constant.FileSearchFilter;
                openFileDialog.Title = Descriptor.Constant.FileSearchTitle;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = openFileDialog.FileName;
                }
            }
            pathTextBox.Text = filePath;
        }

        private async void getSize_Click(object sender, EventArgs e)
        {
            try
            {
                resultRichTextBox.Text = Descriptor.Message.ProcessStarted;
                IExcelFileService service = new ExcelFileServiceInterop();
                resultRichTextBox.Text = await service.GetWorkBookSizesAsync(pathTextBox.Text, passwordTextBox.Text);
            }
            catch (Exception ex)
            {
                Helper.ShowError(ex);
            }
        }
    }
}
