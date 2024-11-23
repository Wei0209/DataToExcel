using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
namespace WindowsFormsApp3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string name = textBox1.Text;
            string account = textBox2.Text;
            string password1 = textBox3.Text;
            string password2 = textBox4.Text;

            if (password1 != password2)
            {
                MessageBox.Show("密碼不一致，請重新輸入。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            string filepath = @"C:\Users\rvl224\source\repos\WindowsFormsApp3\excel\UserData.xlsx";

            if (!File.Exists(filepath))
            {
                
                CreateExcelFile(filepath);
            }
            addDataToExcel(filepath, name, account, password1);
            MessageBox.Show("資料儲存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void CreateExcelFile(string filePath)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                // 建立 Workbook 和 Worksheet
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet()
                {
                    Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "UserData"
                };
                sheets.Append(sheet);

                // 建立標題列
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                Row headerRow = new Row();
                headerRow.Append(
                    new Cell() { CellValue = new CellValue("Name"), DataType = CellValues.String },
                    new Cell() { CellValue = new CellValue("Account"), DataType = CellValues.String },
                    new Cell() { CellValue = new CellValue("Password"), DataType = CellValues.String }
                );
                sheetData.Append(headerRow);

                workbookPart.Workbook.Save();
            }
        }
        private void addDataToExcel(string filepath, string name, string account, string password1)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filepath, true))
            {
                WorksheetPart worksheetPart = document.WorkbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                Row newRow = new Row();
                newRow.Append(
                    new Cell() { CellValue = new CellValue(name), DataType = CellValues.String },
                    new Cell() { CellValue = new CellValue(account), DataType = CellValues.String },
                    new Cell() { CellValue = new CellValue(password1), DataType = CellValues.String }
                );
                sheetData.Append(newRow);

                worksheetPart.Worksheet.Save();
            }
        }
    }
}
