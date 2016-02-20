using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace ExcelTestMerge
{
    public partial class mainForm : Form
    {
        private const string ResultXlsx = "!Результат.xlsx";

        public mainForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            
            
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void selectFolderButton_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() != DialogResult.OK)
                return;

            var testFilesFolder = new DirectoryInfo(folderBrowserDialog.SelectedPath);
            var resultFilePath = Path.Combine(folderBrowserDialog.SelectedPath, ResultXlsx);

            var testcheckXlsx = File.Exists(resultFilePath) ? resultFilePath : "TestCheck.xlsx";
            var xlWorkbook = new XLWorkbook(testcheckXlsx);
            var data = xlWorkbook.Worksheet("Данные");

            data.Range($"B3", $"BJ102").Clear(XLClearOptions.Contents);
            
            var testFiles = testFilesFolder.GetFiles("*.xlsx").ToList();
            testFiles.RemoveAll(f => f.Name.Equals(ResultXlsx, StringComparison.InvariantCultureIgnoreCase));

            progressBar.Maximum = testFiles.Count;
            progressBar.Value = 0;

            var startIndex = 3;

            var list = new List<TestData>();

            foreach (var testFile in testFiles)
            {
                try
                {
                    list.Add(GetData(testFile));
                }
                catch (Exception)
                {
                }
                
                progressBar.Value++;
            }

            foreach (var testData in list.OrderBy(t => t.Student))
            {
                data.Cell($"B{startIndex}").Value = testData.Student;

                var cells = data.Range($"C{startIndex}", $"BJ{startIndex}").Cells().ToList();

                for (var i = 0; i < cells.Count; i++)
                    cells[i].SetValue(testData.Values[i]);

                startIndex++;
            }

            
            xlWorkbook.SaveAs(resultFilePath);

            System.Diagnostics.Process.Start(resultFilePath);
        }

        private TestData GetData(FileInfo file)
        {
            var testData = new TestData();

            var testDataBook = new XLWorkbook(file.FullName);
            var testSheet = testDataBook.Worksheet("Тест");

            testData.Student = testSheet.Cell("F1").GetValue<string>();
            
            testData.Values.AddRange(testSheet.Range("B4", "U4").Cells().Select(c => c.GetValue<string>()));
            testData.Values.AddRange(testSheet.Range("B7", "U7").Cells().Select(c => c.GetValue<string>()));
            testData.Values.AddRange(testSheet.Range("B10", "U10").Cells().Select(c => c.GetValue<string>()));

            return testData;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (saveTestingTemplateDialog.ShowDialog() != DialogResult.OK)
                return;

            File.Copy("TestingTemplate.xlsx", saveTestingTemplateDialog.FileName, true);
        }
    }

    class TestData
    {
        public string Student { get; set; }
        public List<string> Values { get; set; }

        public TestData()
        {
            Values = new List<string>();
        }
    }
}
