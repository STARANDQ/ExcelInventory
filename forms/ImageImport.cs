using System;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ExcelInventory
{
    public partial class ImageImport : Form
    {
        Excel excel = null;
        List<string> imgList;
        List<string> imgNamesList;

        public ImageImport()
        {
            InitializeComponent();
            comboBox1.SelectedIndex = 0;
        }

        public void OpenExcel(string path)
        {
            excel = new Excel(path);
            List<string> sheetNamesList = new List<string>(excel.getSheetsNames());
            foreach (string sheetName in sheetNamesList)
            {
                comboBox1.Items.Add(sheetName);
            }
            comboBox1.Enabled = true;
            textBox1.Enabled = true;
            comboBox1.SelectedIndex = 0;
        }

        private void clearAll()
        {
            if(excel != null)
            {
                excel.closeExcelApp();
            }
            openFileDialog1.FileName = String.Empty;
            for (int i = 0; i < openFileDialog2.FileNames.Length; i++)
            {
                openFileDialog2.FileNames[0] = String.Empty;
            }
            button1.Text = "Select File";
            excel = null;
            imgList = null;
            imgNamesList = null;
            comboBox1.Items.Clear();
            comboBox1.Items.Add("None");
            comboBox1.SelectedIndex = 0;
            comboBox1.Enabled = false;
            textBox1.Text = null;
            textBox1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
        }

        private void scannerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            clearAll();
            this.Hide();
            Scanner form1 = new Scanner();
            form1.Show();
        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            clearAll();
            System.Windows.Forms.Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Directory.GetCurrentDirectory() + "\\assets\\importFiles";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                button1.Text = openFileDialog1.SafeFileName;
                button4.Enabled = true;
                comboBox1.Items.Clear();
                OpenExcel(openFileDialog1.FileName);
            }
            else
                MessageBox.Show("Select a file", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (new Regex(@"^[A-Z]{1}$").IsMatch(textBox1.Text))
            {
                button2.Enabled = true;
            } else
            {
                button2.Enabled = false;
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsDigit(e.KeyChar) || new Regex(@"^[А-Яа-я]{1}$").IsMatch(e.KeyChar.ToString()))
                e.Handled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog2.InitialDirectory = Directory.GetCurrentDirectory() + "\\assets\\img";
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                imgList = new List<string>(openFileDialog2.FileNames);
                imgNamesList = new List<string>();
                foreach (string imgName in openFileDialog2.SafeFileNames)
                {
                    imgNamesList.Add(Regex.Replace(imgName, @"\..+$", "", RegexOptions.Multiline));
                }
                button3.Enabled = imgNamesList.Count != 0 ? true : false;
            } else
                MessageBox.Show("Select files", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int row = 1;
            int col = char.ToUpper(char.Parse(textBox1.Text)) - 64;
            List<string> idList = new List<string>();

            bool isCellEmpty = false;
            bool isEmpty = true;
            while (!isCellEmpty)
            {
                if (excel.readCell(comboBox1.SelectedIndex, row, col) != "")
                {
                    idList.Add(excel.readCell(comboBox1.SelectedIndex, row, col));
                    isEmpty = false;

                } else
                    isCellEmpty = true;

                row++;
            }

            if(!isEmpty)
            {
                int index = 0;
                foreach (string img in imgList)
                {
                    for (int i = 0; i < idList.Count; i++)
                    {
                        if (idList[i] == imgNamesList[index])
                        {
                            excel.addImgToCell(comboBox1.SelectedIndex, i + 1, col + 1, img);
                        }
                    }
                    index++;
                }

                saveFileDialog1.InitialDirectory = Directory.GetCurrentDirectory() + "\\assets\\exportFiles";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    excel.saveAsExcelFile(saveFileDialog1.FileName);
                    clearAll();
                }
                else
                {
                    MessageBox.Show("Save File", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            } else
                MessageBox.Show("Coloumn dosn't have ids", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            clearAll();
        }
    }
}
