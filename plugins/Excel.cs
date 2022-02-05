using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExcelInventory
{
    class Excel
    {
        _Application excel = new _Excel.Application();
        Workbook wb;
        List<Worksheet> wsList = new List<Worksheet>();

        string path = "";

        public Excel(string path)
        {
            try
            {
                this.path = path;
                wb = excel.Workbooks.Open(path, ReadOnly: false);
                foreach (Worksheet sheet in wb.Worksheets)
                {
                    wsList.Add(sheet);
                }
            }
            catch (Exception e)
            {
                wb.Close();
                Console.WriteLine("Error: " + e);
            }
        }

        public List<string> getSheetsNames()
        {
            try
            {
                List<string> sheetNamesList = new List<string>();
                foreach (Worksheet sheet in wsList)
                {
                    sheetNamesList.Add(sheet.Name);
                }
                return sheetNamesList;
            }
            catch (Exception e)
            {
                wb.Close();
                Console.WriteLine("Error: " + e);
                return null;
            }
        }
        
        public void addImgToCell(int sheet, int row, int col, string img)
        {
            try
            {
                const float ImageSize = 100;
                Range oRange = (Range)wsList[sheet].Cells[row, col];
                oRange.Rows.RowHeight = ImageSize + 10;
                oRange.Columns.ColumnWidth = 19.5;

                float Left = (float)(double)oRange.Left + 5;
                float Top = (float)(double)oRange.Top + 5;
                wsList[sheet].Shapes.AddPicture
                    (
                        img,
                        Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                        Left,
                        Top,
                        ImageSize,
                        ImageSize
                    );
            }
            catch (Exception e)
            {
                wb.Close();
                Console.WriteLine("Error: " + e);
            }
        }

        public void saveExcelFile()
        {
            try
            {
                wb.Save();
            }
            catch (Exception e)
            {
                wb.Close();
                Console.WriteLine("Error: " + e);
            }
        }

        public void saveAsExcelFile(string path)
        {
            try
            {
                wb.SaveAs(path, ReadOnlyRecommended: false);
            }
            catch (Exception e)
            {
                wb.Close();
                Console.WriteLine("Error: " + e);
            }
        }
        
        public void closeExcelApp()
        {
            wb.Close();
        }

        public string readCell(int sheet, int row, int col)
        {
            try
            {
                if (wsList[sheet].Cells[row, col].Value2 != null)
                {
                    return wsList[sheet].Cells[row, col].Value2.ToString();
                }
                else
                {
                    return "";
                }
            } 
            catch (Exception e)
            {
                wb.Close();
                Console.WriteLine("Error: " + e);
                return null;
            }
        }
    }
}
