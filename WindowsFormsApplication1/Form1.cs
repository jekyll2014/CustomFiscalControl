﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.IO.Ports;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        DataTable CommandDatabase = new DataTable();
        DataTable ErrorsDatabase = new DataTable();
        DataTable ResultDatabase = new DataTable();

        string SourceFile = "default.txt";

        int SerialtimeOut = 3000;

        public class ResultColumns
        {
            public static int Description { get; set; } = 0;
            public static int Value { get; set; } = 1;
            public static int Type { get; set; } = 2;
            public static int Length { get; set; } = 3;
            public static int Raw { get; set; } = 4;
        }

        public Form1()
        {
            InitializeComponent();
            listBox_code.Items.Add("");
            listBox_code.SelectedIndex = 0;
            commandsCSV_ToolStripTextBox.Text = CustomFiscalControl.Properties.Settings.Default.CommandsDatabaseFile;
            errorsCSV_toolStripTextBox.Text = CustomFiscalControl.Properties.Settings.Default.ErrorsDatabaseFile;
            ReadCsv(commandsCSV_ToolStripTextBox.Text, CommandDatabase);
            for (int i = 0; i < CommandDatabase.Rows.Count; i++) CommandDatabase.Rows[i][0] = Accessory.CheckHexString(CommandDatabase.Rows[i][0].ToString());
            dataGridView_commands.DataSource = CommandDatabase;

            dataGridView_result.DataSource = ResultDatabase;
            dataGridView_commands.ReadOnly = true;
            ResultDatabase.Columns.Add("Desc");
            ResultDatabase.Columns.Add("Value");
            ResultDatabase.Columns.Add("Type");
            ResultDatabase.Columns.Add("Length");
            ResultDatabase.Columns.Add("Raw");

            //ParseEscPos.Init(listBox_code.Items[0].ToString(), CommandDatabase);
            ParseEscPos.commandDataBase = CommandDatabase;
            for (int i = 0; i < dataGridView_commands.Columns.Count; i++) dataGridView_commands.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            for (int i = 0; i < dataGridView_result.Columns.Count; i++) dataGridView_result.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView_result.Columns[ResultColumns.Description].ReadOnly = true;
            dataGridView_result.Columns[ResultColumns.Value].ReadOnly = false;
            dataGridView_result.Columns[ResultColumns.Type].ReadOnly = true;
            dataGridView_result.Columns[ResultColumns.Length].ReadOnly = true;
            dataGridView_result.Columns[ResultColumns.Raw].ReadOnly = false;
            ReadCsv(CustomFiscalControl.Properties.Settings.Default.ErrorsDatabaseFile, ErrorsDatabase);
            SerialPopulate();
            toolStripTextBox_TimeOut.Text = SerialtimeOut.ToString();
            listBox_code.ContextMenuStrip = contextMenuStrip_code;
            dataGridView_commands.ContextMenuStrip = contextMenuStrip_dataBase;
        }

        public void ReadCsv(string fileName, DataTable table)
        {
            table.Clear();
            table.Columns.Clear();
            FileStream inputFile;
            try
            {
                inputFile = File.OpenRead(fileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening file:" + fileName + " : " + ex.Message);
                return;
            }

            //read headers
            StringBuilder inputStr = new StringBuilder();
            int c = inputFile.ReadByte();
            while (c != '\r' && c != '\n' && c != -1)
            {
                byte[] b = new byte[1];
                b[0] = (byte)c;
                inputStr.Append(Encoding.GetEncoding(CustomFiscalControl.Properties.Settings.Default.CodePage).GetString(b));
                c = inputFile.ReadByte();
            }

            //create and count columns and read headers
            int colNum = 0;
            if (inputStr.Length != 0)
            {
                string[] cells = inputStr.ToString().Split(CustomFiscalControl.Properties.Settings.Default.CSVdelimiter);
                colNum = cells.Length - 1;
                for (int i = 0; i < colNum; i++)
                {
                    table.Columns.Add(cells[i]);
                }
            }

            //read CSV content string by string
            while (c != -1)
            {
                int i = 0;
                c = 0;
                inputStr.Length = 0;
                while (i < colNum && c != -1 /*&& c != '\r' && c != '\n'*/)
                {
                    c = inputFile.ReadByte();
                    byte[] b = new byte[1];
                    b[0] = (byte)c;
                    if (c == CustomFiscalControl.Properties.Settings.Default.CSVdelimiter) i++;
                    if (c != -1) inputStr.Append(Encoding.GetEncoding(CustomFiscalControl.Properties.Settings.Default.CodePage).GetString(b));
                }
                while (c != '\r' && c != '\n' && c != -1) c = inputFile.ReadByte();
                if (inputStr.ToString().Replace(CustomFiscalControl.Properties.Settings.Default.CSVdelimiter, ' ').Trim().TrimStart('\r').TrimStart('\n').TrimEnd('\n').TrimEnd('\r') != "")
                {
                    string[] cells = inputStr.ToString().Split(CustomFiscalControl.Properties.Settings.Default.CSVdelimiter);

                    DataRow row = table.NewRow();
                    for (i = 0; i < cells.Length - 1; i++)
                    {
                        row[i] = cells[i].TrimStart('\r').TrimStart('\n').TrimEnd('\n').TrimEnd('\r');
                    }
                    table.Rows.Add(row);
                }
            }
            inputFile.Close();
        }

        private void button_find_Click(object sender, EventArgs e)
        {
            if (listBox_code.SelectedIndex == -1) return;
            ResultDatabase.Clear();
            textBox_search.Clear();
            ParseEscPos.sourceData = listBox_code.SelectedItem.ToString();
            int lineNum = -1;
            if (sender == findThisToolStripMenuItem && dataGridView_commands.CurrentCell != null) lineNum = dataGridView_commands.CurrentCell.RowIndex;
            if (ParseEscPos.FindCommand(0, lineNum))
            {
                ParseEscPos.FindCommandParameter();
                dataGridView_commands.CurrentCell = dataGridView_commands.Rows[ParseEscPos.commandDbLineNum].Cells[ParseEscPos.CSVColumns.CommandName];
                DataRow row = ResultDatabase.NewRow();
                if (ParseEscPos.itIsReply) row[ResultColumns.Value] = "[REPLY] " + ParseEscPos.commandName;
                else row[ResultColumns.Value] = "[COMMAND] " + ParseEscPos.commandName;
                row[ResultColumns.Raw] = ParseEscPos.commandName;
                if (ParseEscPos.crcFailed) row[ResultColumns.Description] += "!!!CRC FAILED!!! ";
                if (ParseEscPos.lengthIncorrect) row[ResultColumns.Description] += "!!!FRAME LENGTH INCORRECT!!! ";
                row[ResultColumns.Description] += ParseEscPos.commandDesc;

                ResultDatabase.Rows.Add(row);
                for (int i = 0; i < ParseEscPos.commandParamDesc.Count; i++)
                {
                    row = ResultDatabase.NewRow();
                    row[ResultColumns.Value] = ParseEscPos.commandParamValue[i];
                    row[ResultColumns.Type] = ParseEscPos.commandParamType[i];
                    row[ResultColumns.Length] = ParseEscPos.commandParamSizeDefined[i];
                    row[ResultColumns.Raw] = ParseEscPos.commandParamRAWValue[i];
                    row[ResultColumns.Description] = ParseEscPos.commandParamDesc[i];
                    if (ParseEscPos.commandParamType[i].ToLower() == ParseEscPos.dataTypes.Error) row[ResultColumns.Description] += ": " + GetErrorDesc(int.Parse(ParseEscPos.commandParamValue[i]));
                    ResultDatabase.Rows.Add(row);
                    if (ParseEscPos.commandParamType[i].ToLower() == ParseEscPos.dataTypes.Bitfield)  //add bitfield display
                    {
                        byte b = byte.Parse(ParseEscPos.commandParamValue[i]);
                        for (int i1 = 0; i1 < 8; i1++)
                        {
                            row = ResultDatabase.NewRow();
                            row[ResultColumns.Value] = (Accessory.GetBit(b, (byte)i1) ? (byte)1 : (byte)0).ToString();
                            row[ResultColumns.Type] = "bit" + i1.ToString();
                            row[ResultColumns.Description] = dataGridView_commands.Rows[ParseEscPos.commandParamDbLineNum[i] + i1 + 1].Cells[ParseEscPos.CSVColumns.CommandDescription].Value;
                            ResultDatabase.Rows.Add(row);
                        }
                    }
                }
            }
            else  //no command found. consider it's a string
            {
                DataRow row = ResultDatabase.NewRow();
                int i = 3;
                while (!ParseEscPos.FindCommand(0 + i) && 0 + i < listBox_code.SelectedItem.ToString().Length) //looking for a non-parseable part end
                {
                    i += 3;
                }
                ParseEscPos.commandName = "";
                row[ResultColumns.Value] += "";
                row[ResultColumns.Value] += "\"" + (String)listBox_code.SelectedItem.ToString() + "\"";
                dataGridView_commands.CurrentCell = dataGridView_commands.Rows[0].Cells[0];
                //dataGridView_commands.FirstDisplayedCell = dataGridView_commands.CurrentCell;
                //dataGridView_commands.Refresh();
                if (Accessory.PrintableHex(listBox_code.SelectedItem.ToString())) row[ResultColumns.Description] = "\"" + Encoding.GetEncoding(CustomFiscalControl.Properties.Settings.Default.CodePage).GetString(Accessory.ConvertHexToByteArray(listBox_code.SelectedItem.ToString())) + "\"";
            }
        }

        private void button_next_Click(object sender, EventArgs e)
        {
            if (listBox_code.SelectedIndex == -1)
            {
                if (listBox_code.Items.Count == 0) return;
                else listBox_code.SelectedIndex = 0;
            }
            if (listBox_code.SelectedIndex < listBox_code.Items.Count - 1) listBox_code.SelectedIndex++;
            button_find_Click(this, EventArgs.Empty);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void saveBinFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog.FileName = SourceFile;
            saveFileDialog.Title = "Save BIN file";
            saveFileDialog.DefaultExt = "bin";
            saveFileDialog.Filter = "BIN files|*.bin|PRN files|*.prn|All files|*.*";
            saveFileDialog.ShowDialog();
        }

        private void saveHexFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog.FileName = SourceFile;
            saveFileDialog.Title = "Save HEX file";
            saveFileDialog.DefaultExt = "txt";
            saveFileDialog.Filter = "Text files|*.txt|HEX files|*.hex|All files|*.*";
            saveFileDialog.ShowDialog();
        }

        private void saveCSVToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog.FileName = CustomFiscalControl.Properties.Settings.Default.CommandsDatabaseFile;
            saveFileDialog.Title = "Save CSV database";
            saveFileDialog.DefaultExt = "csv";
            saveFileDialog.Filter = "CSV files|*.csv|All files|*.*";
            saveFileDialog.ShowDialog();
        }

        private void saveFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (saveFileDialog.Title == "Save HEX file")
            {
                File.WriteAllText(saveFileDialog.FileName, "");
                foreach (string s in listBox_code.Items) File.AppendAllText(saveFileDialog.FileName, s + "\r\n", Encoding.GetEncoding(CustomFiscalControl.Properties.Settings.Default.CodePage));
            }
            else if (saveFileDialog.Title == "Save CSV database")
            {
                int columnCount = dataGridView_commands.ColumnCount;
                StringBuilder output = new StringBuilder();
                for (int i = 0; i < columnCount; i++)
                {
                    output.Append(dataGridView_commands.Columns[i].Name.ToString());
                    output.Append(";");
                }
                output.Append("\r\n");
                for (int i = 0; i < dataGridView_commands.RowCount; i++)
                {
                    for (int j = 0; j < columnCount; j++)
                    {
                        output.Append(dataGridView_commands.Rows[i].Cells[j].Value.ToString());
                        output.Append(";");
                    }
                    output.Append("\r\n");
                }
                try
                {
                    File.WriteAllText(saveFileDialog.FileName, output.ToString(), Encoding.GetEncoding(CustomFiscalControl.Properties.Settings.Default.CodePage));
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error writing to file " + saveFileDialog.FileName + ": " + ex.Message);
                }
            }

        }

        private void loadBinToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog.FileName = "";
            openFileDialog.Title = "Open BIN file";
            openFileDialog.DefaultExt = "bin";
            openFileDialog.Filter = "BIN files|*.bin|PRN files|*.prn|All files|*.*";
            openFileDialog.ShowDialog();
        }

        private void loadHexToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog.FileName = "";
            openFileDialog.Title = "Open HEX file";
            openFileDialog.DefaultExt = "txt";
            openFileDialog.Filter = "HEX files|*.hex|Text files|*.txt|All files|*.*";
            openFileDialog.ShowDialog();
        }

        private void LoadCSVToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog.FileName = "";
            openFileDialog.Title = "Open command CSV database";
            openFileDialog.DefaultExt = "csv";
            openFileDialog.Filter = "CSV files|*.csv|All files|*.*";
            openFileDialog.ShowDialog();
        }

        private void LoadErrorsCSV_toolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog.FileName = "";
            openFileDialog.Title = "Open errors CSV database";
            openFileDialog.DefaultExt = "csv";
            openFileDialog.Filter = "CSV files|*.csv|All files|*.*";
            openFileDialog.ShowDialog();
        }

        private void openFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (openFileDialog.Title == "Open HEX file") //hex text read
            {
                SourceFile = openFileDialog.FileName;
                listBox_code.Items.Clear();
                try
                {
                    foreach (string s in File.ReadAllLines(SourceFile)) listBox_code.Items.Add(Accessory.CheckHexString(s));
                }
                catch (Exception ex)
                {
                    MessageBox.Show("\r\nError reading file " + SourceFile + ": " + ex.Message);
                }
                //Form1.ActiveForm.Text += " " + SourceFile;
                //sourceData.Clear();
                //sourceData.AddRange(Accessory.ConvertHexToByteArray(textBox_code.Text));
                listBox_code.SelectedIndex = 0;
                //ParseEscPos.Init(listBox_code.Items[0].ToString(), CommandDatabase);
            }
            else if (openFileDialog.Title == "Open command CSV database") //hex text read
            {
                CommandDatabase = new DataTable();
                ReadCsv(openFileDialog.FileName, CommandDatabase);
                for (int i = 0; i < CommandDatabase.Rows.Count; i++) CommandDatabase.Rows[i][0] = Accessory.CheckHexString(CommandDatabase.Rows[i][0].ToString());
                dataGridView_commands.DataSource = CommandDatabase;
                ParseEscPos.commandDataBase = CommandDatabase;
            }
            else if (openFileDialog.Title == "Open errors CSV database") //hex text read
            {
                CommandDatabase = new DataTable();
                ReadCsv(openFileDialog.FileName, ErrorsDatabase);
            }

        }

        private void defaultCSVToolStripTextBox_Leave(object sender, EventArgs e)
        {
            if (commandsCSV_ToolStripTextBox.Text != CustomFiscalControl.Properties.Settings.Default.CommandsDatabaseFile)
            {
                CustomFiscalControl.Properties.Settings.Default.CommandsDatabaseFile = commandsCSV_ToolStripTextBox.Text;
                CustomFiscalControl.Properties.Settings.Default.Save();
            }
        }

        private void errorsCSV_toolStripTextBox_Leave(object sender, EventArgs e)
        {
            if (errorsCSV_toolStripTextBox.Text != CustomFiscalControl.Properties.Settings.Default.ErrorsDatabaseFile)
            {
                CustomFiscalControl.Properties.Settings.Default.ErrorsDatabaseFile = errorsCSV_toolStripTextBox.Text;
                CustomFiscalControl.Properties.Settings.Default.Save();
            }
        }

        private void enableDatabaseEditToolStripMenuItem_Click(object sender, EventArgs e)
        {
            enableDatabaseEditToolStripMenuItem.Checked = !enableDatabaseEditToolStripMenuItem.Checked;
            dataGridView_commands.ReadOnly = !enableDatabaseEditToolStripMenuItem.Checked;
        }

        private string GetErrorDesc(int errNum)
        {
            for (int i = 0; i < ErrorsDatabase.Rows.Count; i++)
            {
                if (int.Parse(ErrorsDatabase.Rows[i][0].ToString()) == errNum) return ErrorsDatabase.Rows[i][1].ToString();
            }
            return "!!!Unknown error!!!";
        }

        public static bool PrintableHex(string str)
        {
            for (int i = 0; i < str.Length; i += 3)
            {
                if (!byte.TryParse(str.Substring(i, 3), NumberStyles.HexNumber, null, out byte n)) return false;
                else if (n < 32 && n != 0) return false;
            }
            return true;
        }

        private void dataGridView_result_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            this.dataGridView_result.CellValueChanged -= new DataGridViewCellEventHandler(this.dataGridView_result_CellValueChanged);
            if (dataGridView_result.CurrentCell.ColumnIndex == ResultColumns.Value)
            {
                if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value.ToString() == ParseEscPos.dataTypes.Bitfield)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value = ParseEscPos.BitfieldToRaw(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value.ToString());
                    int n = 0;
                    int.TryParse(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value.ToString(), out n);
                    int i = dataGridView_result.CurrentRow.Index;
                    for (int i1 = 0; i1 < 8; i1++)
                    {
                        dataGridView_result.Rows[i + 1 + i1].Cells[ResultColumns.Value].Value = Convert.ToInt32(Accessory.GetBit((byte)n, (byte)i1)).ToString();
                    }
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value.ToString() == ParseEscPos.dataTypes.Data)
                {
                    int n = 0;
                    if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length].Value.ToString() == "?") n = dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value.ToString().Length;
                    else int.TryParse(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length].Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value = ParseEscPos.DataToRaw(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value.ToString(), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value.ToString() == ParseEscPos.dataTypes.Error)
                {
                    int n = 0;
                    int.TryParse(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length].Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value = ParseEscPos.ErrorToRaw(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value.ToString(), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value.ToString() == ParseEscPos.dataTypes.Money)
                {
                    int n = 0;
                    int.TryParse(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length].Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value = ParseEscPos.MoneyToRaw(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value.ToString(), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value.ToString() == ParseEscPos.dataTypes.Number)
                {
                    int n = 0;
                    int.TryParse(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length].Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value = ParseEscPos.NumberToRaw(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value.ToString(), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value.ToString() == ParseEscPos.dataTypes.Password)
                {
                    int n = 0;
                    int.TryParse(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length].Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value = ParseEscPos.PasswordToRaw(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value.ToString(), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value.ToString() == ParseEscPos.dataTypes.Quantity)
                {
                    int n = 0;
                    int.TryParse(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length].Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value = ParseEscPos.QuantityToRaw(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value.ToString(), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value.ToString() == ParseEscPos.dataTypes.String)
                {
                    int n = 0;
                    if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length].Value.ToString() == "?") n = dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value.ToString().Length;
                    else int.TryParse(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length].Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value = ParseEscPos.StringToRaw(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value.ToString(), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value.ToString() == ParseEscPos.dataTypes.PrefData)
                {
                    int n = 0;
                    if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length].Value.ToString() == "?") n = dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value.ToString().Length;
                    else int.TryParse(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length].Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value = ParseEscPos.PrefDataToRaw(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value.ToString(), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value.ToString() == ParseEscPos.dataTypes.TLVData)
                {
                    int n = 0;
                    if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length].Value.ToString() == "?") n = dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value.ToString().Length;
                    else int.TryParse(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length].Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value = ParseEscPos.TLVDataToRaw(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value.ToString(), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value.ToString().StartsWith("bit"))
                {
                    int i = dataGridView_result.CurrentCell.RowIndex - 1;
                    while (dataGridView_result.Rows[i].Cells[ResultColumns.Type].Value.ToString() != ParseEscPos.dataTypes.Bitfield) i--;
                    //collect bits to int
                    byte n = 0;
                    for (int i1 = 0; i1 < 8; i1++)
                    {
                        if (dataGridView_result.Rows[i + i1 + 1].Cells[ResultColumns.Value].Value.ToString().Trim() == "1") n += (byte)Math.Pow(2, i1);
                    }
                    dataGridView_result.Rows[i].Cells[ResultColumns.Value].Value = n.ToString();
                    dataGridView_result.Rows[i].Cells[ResultColumns.Raw].Value = ParseEscPos.BitfieldToRaw(dataGridView_result.Rows[i].Cells[ResultColumns.Value].Value.ToString());
                }
            }
            else if (dataGridView_result.CurrentCell.ColumnIndex == ResultColumns.Raw)
            {
                dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value = Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString());
                if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value.ToString() == ParseEscPos.dataTypes.Password)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value = Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString());
                    double l = ParseEscPos.RawToPassword(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString());
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value = l.ToString();
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value.ToString() == ParseEscPos.dataTypes.String)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value = Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString());
                    int n = 0;
                    if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length].Value.ToString() == "?") n = dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString().Length / 3;
                    else int.TryParse(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length].Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value = ParseEscPos.RawToString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString(), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value.ToString() == ParseEscPos.dataTypes.PrefData)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value = Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString());
                    int n = 0;
                    if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length].Value.ToString() == "?") n = dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString().Length / 3;
                    else int.TryParse(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length].Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value = ParseEscPos.RawToPrefData(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString(), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value.ToString() == ParseEscPos.dataTypes.TLVData)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value = Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString());
                    int n = 0;
                    if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length].Value.ToString() == "?") n = dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString().Length / 3;
                    else int.TryParse(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length].Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value = ParseEscPos.RawToTLVData(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString(), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value.ToString() == ParseEscPos.dataTypes.Number)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value = Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString());
                    double l = ParseEscPos.RawToNumber(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString());
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value = l.ToString();
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value.ToString() == ParseEscPos.dataTypes.Money)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value = Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString());
                    double l = ParseEscPos.RawToMoney(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString());
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value = l.ToString();
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value.ToString() == ParseEscPos.dataTypes.Quantity)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value = Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString());
                    double l = ParseEscPos.RawToQuantity(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString());
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value = l.ToString();
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value.ToString() == ParseEscPos.dataTypes.Error)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value = Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString());
                    double l = ParseEscPos.RawToError(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString());
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value = l.ToString();
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value.ToString() == ParseEscPos.dataTypes.Data)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value = Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString());
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value = ParseEscPos.RawToData(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString());
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value.ToString() == ParseEscPos.dataTypes.Bitfield)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value = Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString());
                    double l = ParseEscPos.RawToBitfield(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString());
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value = l.ToString();
                    int n = 0;
                    int.TryParse(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value.ToString(), out n);
                    int i = dataGridView_result.CurrentRow.Index;
                    for (int i1 = 0; i1 < 8; i1++)
                    {
                        dataGridView_result.Rows[i + 1 + i1].Cells[ResultColumns.Value].Value = (Accessory.GetBit((byte)n, (byte)i1) ? (byte)1 : (byte)0).ToString();
                    }
                }
            }
            this.dataGridView_result.CellValueChanged += new DataGridViewCellEventHandler(this.dataGridView_result_CellValueChanged);
        }

        private void dataGridView_commands_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            button_newCommand_Click(this, EventArgs.Empty);
        }

        private string collectCommand()
        {
            string data = "";
            for (int i = 0; i < ResultDatabase.Rows.Count; i++) data += ResultDatabase.Rows[i][ResultColumns.Raw].ToString();
            byte[] length = new byte[2];
            length[1] = (byte)((data.Length / 3) / 256);
            length[0] = (byte)((data.Length / 3) - length[1]);
            data = Accessory.ConvertByteArrayToHex(length) + data;
            byte[] dataByte = Accessory.ConvertHexToByteArray(data);
            return ("01 " + data + Accessory.ConvertByteToHex(ParseEscPos.Q3xf_CRC(dataByte, dataByte.Length)));
        }

        private void button_add_Click(object sender, EventArgs e)
        {
            listBox_code.Items.Add(collectCommand());
        }

        private void button_replace_Click(object sender, EventArgs e)
        {
            if (listBox_code.SelectedIndex == -1) return;
            listBox_code.Items[listBox_code.SelectedIndex] = collectCommand();
        }

        private void button_insert_Click(object sender, EventArgs e)
        {
            if (listBox_code.SelectedIndex == -1) return;
            listBox_code.Items.Insert(listBox_code.SelectedIndex, collectCommand());
            listBox_code.SelectedIndex--;
        }

        private void button_remove_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < listBox_code.Items.Count; i++)
                if (listBox_code.Items[i].ToString().StartsWith("06 "))
                {
                    listBox_code.Items.RemoveAt(i);
                    i--;
                }
        }

        private void button_clear_Click(object sender, EventArgs e)
        {
            listBox_code.Items.Clear();
        }

        private void textBox_search_TextChanged(object sender, EventArgs e)
        {
            dataGridView_commands.CurrentCell = null;
            DataGridViewRow row;
            if (textBox_search.Text != "")
            {
                for (int i = 0; i < dataGridView_commands.RowCount; i++)
                {
                    row = dataGridView_commands.Rows[i];
                    if (dataGridView_commands.Rows[i].Cells[ParseEscPos.CSVColumns.CommandName].Value.ToString() != "")
                    {
                        if (dataGridView_commands.Rows[i].Cells[ParseEscPos.CSVColumns.CommandDescription].Value.ToString().ToLower().Contains(textBox_search.Text.ToLower()))
                        {
                            row.Visible = true;
                            i++;
                            while (i < dataGridView_commands.RowCount && dataGridView_commands.Rows[i].Cells[ParseEscPos.CSVColumns.CommandName].Value.ToString() == "")
                            {
                                row = dataGridView_commands.Rows[i];
                                row.Visible = true;
                                i++;
                            }
                            i--;
                        }
                        else row.Visible = false;
                    }
                    else row.Visible = false;
                }
            }
            else
            {
                for (int i = 0; i < dataGridView_commands.RowCount; i++)
                {
                    row = dataGridView_commands.Rows[i];
                    row.Visible = true;
                }
            }
        }

        private void button_newCommand_Click(object sender, EventArgs e)
        {
            //restore 
            ParseEscPos.CSVColumns.CommandParameterSize = 1;
            ParseEscPos.CSVColumns.CommandParameterType = 2;
            ParseEscPos.CSVColumns.CommandParameterValue = 3;
            ParseEscPos.CSVColumns.CommandDescription = 4;
            ParseEscPos.itIsReply = false;

            //dataGridView_commands_CellDoubleClick(this, new DataGridViewCellEventArgs(this.dataGridView_commands.CurrentCell.ColumnIndex, this.dataGridView_commands.CurrentRow.Index));
            if (dataGridView_commands.Rows[dataGridView_commands.CurrentCell.RowIndex].Cells[ParseEscPos.CSVColumns.CommandName].Value.ToString() != "")
            {
                int currentRow = dataGridView_commands.CurrentCell.RowIndex;
                ResultDatabase.Clear();
                DataRow row = ResultDatabase.NewRow();
                row[ResultColumns.Value] = dataGridView_commands.Rows[currentRow].Cells[ParseEscPos.CSVColumns.CommandName].Value.ToString();
                row[ResultColumns.Raw] = row[ResultColumns.Value];
                row[ResultColumns.Description] = dataGridView_commands.Rows[currentRow].Cells[ParseEscPos.CSVColumns.CommandDescription].Value.ToString();
                ResultDatabase.Rows.Add(row);

                int i = currentRow + 1;
                while (i < dataGridView_commands.Rows.Count && dataGridView_commands.Rows[i].Cells[ParseEscPos.CSVColumns.CommandName].Value.ToString() == "")
                {
                    if (dataGridView_commands.Rows[i].Cells[ParseEscPos.CSVColumns.CommandParameterSize].Value.ToString() != "")
                    {
                        row = ResultDatabase.NewRow();
                        row[ResultColumns.Type] = dataGridView_commands.Rows[i].Cells[ParseEscPos.CSVColumns.CommandParameterType].Value.ToString();
                        row[ResultColumns.Length] = dataGridView_commands.Rows[i].Cells[ParseEscPos.CSVColumns.CommandParameterSize].Value.ToString();
                        row[ResultColumns.Description] = dataGridView_commands.Rows[i].Cells[ParseEscPos.CSVColumns.CommandDescription].Value.ToString();
                        if (row[ResultColumns.Type].ToString() == ParseEscPos.dataTypes.Password)
                        {
                            row[ResultColumns.Value] = textBox_password.Text;
                            int n = 0;
                            int.TryParse(row[ResultColumns.Length].ToString(), out n);
                            row[ResultColumns.Raw] = ParseEscPos.PasswordToRaw(textBox_password.Text, n);
                        }
                        else
                        {
                            row[ResultColumns.Value] = "";
                            row[ResultColumns.Raw] = "";
                        }
                        ResultDatabase.Rows.Add(row);
                        if (row[ResultColumns.Type].ToString() == ParseEscPos.dataTypes.Bitfield)  //decode bitfield
                        {
                            for (int i1 = 0; i1 < 8; i1++)
                            {
                                row = ResultDatabase.NewRow();
                                row[ResultColumns.Value] = "0";
                                row[ResultColumns.Type] = "bit" + i1.ToString();
                                row[ResultColumns.Description] = dataGridView_commands.Rows[i + i1 + 1].Cells[ParseEscPos.CSVColumns.CommandDescription].Value.ToString();
                                ResultDatabase.Rows.Add(row);
                            }
                        }
                    }
                    i++;
                }
            }
        }

        private void listBox_code_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            button_find_Click(this, EventArgs.Empty);
        }

        private void listBox_code_KeyDown(object sender, KeyEventArgs e)
        {
            if (sender != listBox_code) return;

            if (listBox_code.SelectedIndex == -1) return;

            if (e.Control && e.KeyCode == Keys.C && listBox_code.SelectedItem.ToString() != "") Clipboard.SetText(listBox_code.SelectedItem.ToString());
            else if (e.Control && e.KeyCode == Keys.Insert && listBox_code.SelectedItem.ToString() != "") Clipboard.SetText(listBox_code.SelectedItem.ToString());
            else if (e.Control && e.KeyCode == Keys.V && Accessory.GetStringFormat(Clipboard.GetText()) == 16)
            {
                listBox_code.Items[listBox_code.SelectedIndex] = Accessory.CheckHexString(Clipboard.GetText());
            }
            else if (e.Shift && e.KeyCode == Keys.Insert && Accessory.GetStringFormat(Clipboard.GetText()) == 16)
            {
                listBox_code.Items[listBox_code.SelectedIndex] = Accessory.CheckHexString(Clipboard.GetText());
            }
            else if (e.KeyCode == Keys.Delete && listBox_code.SelectedItem.ToString() != "") listBox_code.Items.RemoveAt(listBox_code.SelectedIndex);
            else if (e.Control && e.KeyCode == Keys.P) button_find_Click(this, EventArgs.Empty);
            else if (e.Control && e.KeyCode == Keys.S && button_Send.Enabled) button_Send_Click(this, EventArgs.Empty);
        }

        private void toolStripMenuItem_Connect_Click(object sender, EventArgs e)
        {
            if (toolStripMenuItem_Connect.Checked)
            {
                try
                {
                    SerialPort1.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error closing port " + SerialPort1.PortName + ": " + ex.Message);
                }
                toolStripComboBox_PortName.Enabled = true;
                toolStripComboBox_PortSpeed.Enabled = true;
                toolStripComboBox_PortHandshake.Enabled = true;
                toolStripComboBox_PortDataBits.Enabled = true;
                toolStripComboBox_PortParity.Enabled = true;
                toolStripComboBox_PortStopBits.Enabled = true;
                button_Send.Enabled = false;
                button_SendAll.Enabled = false;
                sendToolStripMenuItem.Enabled = false;
                toolStripMenuItem_Connect.Text = "Connect";
                toolStripMenuItem_Connect.Checked = false;
            }
            else
            {
                if (toolStripComboBox_PortName.SelectedIndex != 0)
                {
                    toolStripComboBox_PortName.Enabled = false;
                    toolStripComboBox_PortSpeed.Enabled = false;
                    toolStripComboBox_PortHandshake.Enabled = false;
                    toolStripComboBox_PortDataBits.Enabled = false;
                    toolStripComboBox_PortParity.Enabled = false;
                    toolStripComboBox_PortStopBits.Enabled = false;

                    SerialPort1.PortName = toolStripComboBox_PortName.Text;
                    SerialPort1.BaudRate = Convert.ToInt32(toolStripComboBox_PortSpeed.Text);
                    SerialPort1.DataBits = Convert.ToUInt16(toolStripComboBox_PortDataBits.Text);
                    SerialPort1.Handshake = (Handshake)Enum.Parse(typeof(Handshake), toolStripComboBox_PortHandshake.Text);
                    SerialPort1.Parity = (Parity)Enum.Parse(typeof(Parity), toolStripComboBox_PortParity.Text);
                    SerialPort1.StopBits = (StopBits)Enum.Parse(typeof(StopBits), toolStripComboBox_PortStopBits.Text);
                    //SerialPort1.ReadTimeout = CustomFiscalControl.Properties.Settings.Default.ReceiveTimeOut;
                    //SerialPort1.WriteTimeout = CustomFiscalControl.Properties.Settings.Default.SendTimeOut;
                    SerialPort1.ReadBufferSize = 8192;
                    try
                    {
                        SerialPort1.Open();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error opening port " + SerialPort1.PortName + ": " + ex.Message);
                        toolStripComboBox_PortName.Enabled = true;
                        toolStripComboBox_PortSpeed.Enabled = true;
                        toolStripComboBox_PortHandshake.Enabled = true;
                        toolStripComboBox_PortDataBits.Enabled = true;
                        toolStripComboBox_PortParity.Enabled = true;
                        toolStripComboBox_PortStopBits.Enabled = true;
                        return;
                    }
                    toolStripMenuItem_Connect.Text = "Disconnect";
                    toolStripMenuItem_Connect.Checked = true;
                    button_Send.Enabled = true;
                    button_SendAll.Enabled = true;
                    sendToolStripMenuItem.Enabled = true;
                }
            }
        }

        private void button_Send_Click(object sender, EventArgs e)
        {
            if (listBox_code.SelectedIndex == -1) return;
            if (listBox_code.SelectedItem.ToString() == "") return;

            if (listBox_code.SelectedItem.ToString().Substring(0, 2) == "01")
            {
                byte[] _txBytes = Accessory.ConvertHexToByteArray(listBox_code.SelectedItem.ToString());
                try
                {
                    SerialPort1.Write(_txBytes, 0, _txBytes.Length);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error sending to port " + SerialPort1.PortName + ": " + ex.Message);
                }
                List<byte> _rxBytes = new List<byte>();
                bool _timeout = false;
                bool _nackOK = false;
                bool _frameOK = false;
                bool _lengthOK = false;
                int _frameLength = 0;
                DateTime startTime = DateTime.UtcNow;
                try
                {
                    while (!_timeout)
                    {
                        int c = -1;
                        if (!_nackOK)
                        {
                            if (SerialPort1.BytesToRead > 0) c = SerialPort1.ReadByte();
                            if (c == 06)
                            {
                                _rxBytes.Add((byte)c);
                                _nackOK = true;
                            }
                        }
                        else if (_nackOK && !_frameOK)
                        {
                            if (SerialPort1.BytesToRead > 0) c = SerialPort1.ReadByte();
                            if (c == 01)
                            {
                                _rxBytes.Add((byte)c);
                                _frameOK = true;
                            }
                        }
                        else if (_frameOK && !_lengthOK)
                        {
                            if (SerialPort1.BytesToRead >= 2)
                            {
                                c = SerialPort1.ReadByte();
                                _rxBytes.Add((byte)c);
                                _frameLength = c;
                                c = SerialPort1.ReadByte();
                                _rxBytes.Add((byte)c);
                                _frameLength = _frameLength + c * 256;
                                _lengthOK = true;
                            }
                        }
                        else if (_lengthOK)
                        {
                            if (SerialPort1.BytesToRead > 0) c = SerialPort1.ReadByte();
                            if (c != -1) _rxBytes.Add((byte)c);
                        }
                        else SerialPort1.ReadByte();
                        if (_rxBytes.Count > _frameLength + 4) _timeout = true;
                        if (SerialPort1.BytesToRead == 0 && DateTime.UtcNow.Subtract(startTime).TotalMilliseconds > SerialtimeOut) _timeout = true;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error reading port " + SerialPort1.PortName + ": " + ex.Message);
                }
                if (_rxBytes.Count > 0)
                {
                    string data = Accessory.ConvertByteArrayToHex(_rxBytes.ToArray());
                    if (listBox_code.SelectedIndex + 1 >= listBox_code.Items.Count) listBox_code.Items.Add(data);
                    else if (listBox_code.Items[listBox_code.SelectedIndex + 1].ToString().Length > 2 && listBox_code.Items[listBox_code.SelectedIndex + 1].ToString().Substring(0, 2) == "06") listBox_code.Items[listBox_code.SelectedIndex + 1] = data;
                    else listBox_code.Items.Insert(listBox_code.SelectedIndex + 1, data);
                }
            }
        }

        private void button_SendAll_Click(object sender, EventArgs e)
        {
            if (listBox_code.SelectedIndex == -1) listBox_code.SelectedIndex = 0;
            for (int i = listBox_code.SelectedIndex; i < listBox_code.Items.Count; i++)
            {
                listBox_code.SelectedIndex = i;
                if (listBox_code.SelectedItem.ToString().Length > 2)
                {
                    if (listBox_code.SelectedItem.ToString().Substring(0, 2) == "06")
                    {
                        if (ToolStripMenuItem_stopOnErrorReplied.Checked && listBox_code.SelectedItem.ToString().Length > 20 && listBox_code.SelectedItem.ToString().Substring(5 * 3, 5) != "00 00") return;
                    }
                    else if (listBox_code.SelectedItem.ToString().Substring(0, 2) == "15") return;
                    else button_Send_Click(button_SendAll, EventArgs.Empty);
                }
            }
        }

        void SerialPopulate()
        {
            toolStripComboBox_PortName.Items.Clear();
            toolStripComboBox_PortHandshake.Items.Clear();
            toolStripComboBox_PortParity.Items.Clear();
            toolStripComboBox_PortStopBits.Items.Clear();
            //Serial settings populate
            toolStripComboBox_PortName.Items.Add("-None-");
            //Add ports
            foreach (string s in SerialPort.GetPortNames())
            {
                toolStripComboBox_PortName.Items.Add(s);
            }
            //Add handshake methods
            foreach (string s in Enum.GetNames(typeof(Handshake)))
            {
                toolStripComboBox_PortHandshake.Items.Add(s);
            }
            //Add parity
            foreach (string s in Enum.GetNames(typeof(Parity)))
            {
                toolStripComboBox_PortParity.Items.Add(s);
            }
            //Add stopbits
            foreach (string s in Enum.GetNames(typeof(StopBits)))
            {
                toolStripComboBox_PortStopBits.Items.Add(s);
            }
            toolStripComboBox_PortName.SelectedIndex = toolStripComboBox_PortName.Items.Count - 1;
            if (toolStripComboBox_PortName.Items.Count == 1)
            {
                toolStripMenuItem_Connect.Enabled = false;
            }
            toolStripComboBox_PortSpeed.SelectedIndex = 0;
            toolStripComboBox_PortHandshake.SelectedIndex = 0;
            toolStripComboBox_PortDataBits.SelectedIndex = 0;
            toolStripComboBox_PortParity.SelectedIndex = 0;
            toolStripComboBox_PortStopBits.SelectedIndex = 1;
            if (toolStripComboBox_PortName.SelectedIndex == 0) toolStripMenuItem_Connect.Enabled = false;
            else toolStripMenuItem_Connect.Enabled = true;
        }

        private void toolStripTextBox_TimeOut_TextChanged(object sender, EventArgs e)
        {
            if (!int.TryParse(toolStripTextBox_TimeOut.Text, out SerialtimeOut)) toolStripTextBox_TimeOut.Text = "1000";
        }

        private void ToolStripMenuItem_stopOnErrorReplied_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem_stopOnErrorReplied.Checked = !ToolStripMenuItem_stopOnErrorReplied.Checked;
        }

        private void listBox_code_MouseUp(object sender, MouseEventArgs e)
        {
            int index = this.listBox_code.IndexFromPoint(e.Location);
            if (index != ListBox.NoMatches && e.Button == MouseButtons.Right)
            {
                listBox_code.SelectedIndex = index;
                contextMenuStrip_code.Visible = true;
            }
            else
            {
                contextMenuStrip_code.Visible = false;
            }
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listBox_code.SelectedItem.ToString() != "") Clipboard.SetText(listBox_code.SelectedItem.ToString());
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listBox_code.SelectedIndex == -1) return;
            listBox_code.Items.RemoveAt(listBox_code.SelectedIndex);
        }

        private void parseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button_find_Click(this, EventArgs.Empty);
        }

        private void sendToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button_Send_Click(this, EventArgs.Empty);
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Accessory.GetStringFormat(Clipboard.GetText()) == 16)
            {
                listBox_code.Items[listBox_code.SelectedIndex] = Accessory.CheckHexString(Clipboard.GetText());
            }
        }

        private void newCommandToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button_newCommand_Click(this, EventArgs.Empty);
        }

        private void findThisToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button_find_Click(findThisToolStripMenuItem, EventArgs.Empty);
        }

        private void dataGridView_commands_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && e.RowIndex >= 0) dataGridView_commands.CurrentCell = dataGridView_commands.Rows[e.RowIndex].Cells[e.ColumnIndex];
        }

        private void COMPortToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SerialPopulate();
        }
    }
}