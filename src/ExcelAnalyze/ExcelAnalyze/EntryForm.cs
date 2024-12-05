﻿using System;
using System.Windows.Forms;
using ExcelAnalyze.Logic;

namespace ExcelAnalyze
{
    public partial class EntryForm : Form
    {
        private const string FileSearchTitle = "Select an Excel File";
        private const string FileSearchFilter = "Excel Files|*.xls;*.xlsx;*.xlsm";

        public EntryForm()
        {
            InitializeComponent();
        }

        private void EntryForm_Load(object sender, EventArgs e) { }

        private void searchButton_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = FileSearchFilter;
                openFileDialog.Title = FileSearchTitle;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = openFileDialog.FileName;
                }
            }
            pathTextBox.Text = filePath;
        }

        private void getSize_Click(object sender, EventArgs e)
        {
            resultRichTextBox.Text = Descriptor.Message.ProcessStarted;
            IExcelFileService service = new ExcelFileServiceString();
            resultRichTextBox.Text = service.GetWorkBookSizes(pathTextBox.Text, passwordTextBox.Text);
        }
    }
}