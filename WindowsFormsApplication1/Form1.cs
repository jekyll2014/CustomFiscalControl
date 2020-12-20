using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.IO.Ports;
using System.Text;
using System.Windows.Forms;
using CustomFiscalControl.Properties;

namespace CustomFiscalControl
{
    public partial class Form1 : Form
    {
        private DataTable _commandDatabase = new DataTable();
        private readonly DataTable _errorsDatabase = new DataTable();
        private readonly DataTable _resultDatabase = new DataTable();

        private string _sourceFile = "default.txt";

        private int _serialtimeOut = 3000;

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
            commandsCSV_ToolStripTextBox.Text = Settings.Default.CommandsDatabaseFile;
            errorsCSV_toolStripTextBox.Text = Settings.Default.ErrorsDatabaseFile;
            ReadCsv(commandsCSV_ToolStripTextBox.Text, _commandDatabase);
            for (var i = 0; i < _commandDatabase.Rows.Count; i++)
                _commandDatabase.Rows[i][0] = Accessory.CheckHexString(_commandDatabase.Rows[i][0].ToString());
            dataGridView_commands.DataSource = _commandDatabase;

            dataGridView_result.DataSource = _resultDatabase;
            dataGridView_commands.ReadOnly = true;
            _resultDatabase.Columns.Add("Desc");
            _resultDatabase.Columns.Add("Value");
            _resultDatabase.Columns.Add("Type");
            _resultDatabase.Columns.Add("Length");
            _resultDatabase.Columns.Add("Raw");

            //ParseEscPos.Init(listBox_code.Items[0].ToString(), CommandDatabase);
            ParseEscPos.CommandDataBase = _commandDatabase;
            for (var i = 0; i < dataGridView_commands.Columns.Count; i++)
                dataGridView_commands.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            for (var i = 0; i < dataGridView_result.Columns.Count; i++)
                dataGridView_result.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView_result.Columns[ResultColumns.Description].ReadOnly = true;
            dataGridView_result.Columns[ResultColumns.Value].ReadOnly = false;
            dataGridView_result.Columns[ResultColumns.Type].ReadOnly = true;
            dataGridView_result.Columns[ResultColumns.Length].ReadOnly = true;
            dataGridView_result.Columns[ResultColumns.Raw].ReadOnly = false;
            ReadCsv(Settings.Default.ErrorsDatabaseFile, _errorsDatabase);
            SerialPopulate();
            toolStripTextBox_TimeOut.Text = _serialtimeOut.ToString();
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
            var inputStr = new StringBuilder();
            var c = inputFile.ReadByte();
            while (c != '\r' && c != '\n' && c != -1)
            {
                var b = new byte[1];
                b[0] = (byte) c;
                inputStr.Append(Encoding.GetEncoding(Settings.Default.CodePage).GetString(b));
                c = inputFile.ReadByte();
            }

            //create and count columns and read headers
            var colNum = 0;
            if (inputStr.Length != 0)
            {
                var cells = inputStr.ToString().Split(Settings.Default.CSVdelimiter);
                colNum = cells.Length - 1;
                for (var i = 0; i < colNum; i++) table.Columns.Add(cells[i]);
            }

            //read CSV content string by string
            while (c != -1)
            {
                var i = 0;
                c = 0;
                inputStr.Length = 0;
                while (i < colNum && c != -1 /*&& c != '\r' && c != '\n'*/)
                {
                    c = inputFile.ReadByte();
                    var b = new byte[1];
                    b[0] = (byte) c;
                    if (c == Settings.Default.CSVdelimiter) i++;
                    if (c != -1) inputStr.Append(Encoding.GetEncoding(Settings.Default.CodePage).GetString(b));
                }

                while (c != '\r' && c != '\n' && c != -1) c = inputFile.ReadByte();
                if (inputStr.ToString().Replace(Settings.Default.CSVdelimiter, ' ').Trim().TrimStart('\r')
                    .TrimStart('\n').TrimEnd('\n').TrimEnd('\r') != "")
                {
                    var cells = inputStr.ToString().Split(Settings.Default.CSVdelimiter);

                    var row = table.NewRow();
                    for (i = 0; i < cells.Length - 1; i++)
                        row[i] = cells[i].TrimStart('\r').TrimStart('\n').TrimEnd('\n').TrimEnd('\r');
                    table.Rows.Add(row);
                }
            }

            inputFile.Close();
        }

        private void Button_find_Click(object sender, EventArgs e)
        {
            if (listBox_code.SelectedIndex == -1) return;
            _resultDatabase.Clear();
            textBox_search.Clear();
            ParseEscPos.SourceData.Clear();
            ParseEscPos.SourceData.AddRange(Accessory.ConvertHexToByteArray(listBox_code.SelectedItem.ToString()));
            var lineNum = -1;
            if (sender == findThisToolStripMenuItem && dataGridView_commands.CurrentCell != null)
                lineNum = dataGridView_commands.CurrentCell.RowIndex;
            if (ParseEscPos.FindCommand(0, lineNum))
            {
                ParseEscPos.FindCommandParameter();
                dataGridView_commands.CurrentCell = dataGridView_commands.Rows[ParseEscPos.CommandDbLineNum]
                    .Cells[ParseEscPos.CSVColumns.CommandNameColumn];
                var row = _resultDatabase.NewRow();
                if (ParseEscPos.ItIsReply) row[ResultColumns.Value] = "[REPLY] " + ParseEscPos.CommandName;
                else row[ResultColumns.Value] = "[COMMAND] " + ParseEscPos.CommandName;
                row[ResultColumns.Raw] = ParseEscPos.CommandName;
                if (ParseEscPos.CrcFailed) row[ResultColumns.Description] += "!!!CRC FAILED!!! ";
                if (ParseEscPos.LengthIncorrect) row[ResultColumns.Description] += "!!!FRAME LENGTH INCORRECT!!! ";
                row[ResultColumns.Description] += ParseEscPos.CommandDesc;

                _resultDatabase.Rows.Add(row);
                for (var i = 0; i < ParseEscPos.CommandParamDesc.Count; i++)
                {
                    row = _resultDatabase.NewRow();
                    row[ResultColumns.Value] = ParseEscPos.CommandParamValue[i];
                    row[ResultColumns.Type] = ParseEscPos.CommandParamType[i];
                    row[ResultColumns.Length] = ParseEscPos.CommandParamSizeDefined[i];
                    row[ResultColumns.Raw] =
                        Accessory.ConvertByteArrayToHex(ParseEscPos.CommandParamRawValue[i].ToArray());
                    row[ResultColumns.Description] = ParseEscPos.CommandParamDesc[i];
                    if (ParseEscPos.CommandParamType[i].ToLower() == ParseEscPos.DataTypes.Error)
                        row[ResultColumns.Description] +=
                            ": " + GetErrorDesc(int.Parse(ParseEscPos.CommandParamValue[i]));
                    _resultDatabase.Rows.Add(row);
                    if (ParseEscPos.CommandParamType[i].ToLower() == ParseEscPos.DataTypes.Bitfield
                    ) //add bitfield display
                    {
                        var b = byte.Parse(ParseEscPos.CommandParamValue[i]);
                        for (var i1 = 0; i1 < 8; i1++)
                        {
                            row = _resultDatabase.NewRow();
                            row[ResultColumns.Value] =
                                (Accessory.GetBit(b, (byte) i1) ? (byte) 1 : (byte) 0).ToString();
                            row[ResultColumns.Type] = "bit" + i1;
                            row[ResultColumns.Description] = dataGridView_commands
                                .Rows[ParseEscPos.CommandParamDbLineNum[i] + i1 + 1]
                                .Cells[ParseEscPos.CSVColumns.CommandDescriptionColumn].Value;
                            _resultDatabase.Rows.Add(row);
                        }
                    }
                }
            }
            else //no command found. consider it's a string
            {
                var row = _resultDatabase.NewRow();
                var i = 3;
                while (!ParseEscPos.FindCommand(0 + i / 3) && 0 + i < listBox_code.SelectedItem.ToString().Length
                ) //looking for a non-parseable part end
                    i += 3;
                ParseEscPos.CommandName = "";
                row[ResultColumns.Value] = "\"" + listBox_code.SelectedItem.ToString() + "\"";
                dataGridView_commands.CurrentCell = dataGridView_commands.Rows[0].Cells[0];
                //dataGridView_commands.FirstDisplayedCell = dataGridView_commands.CurrentCell;
                //dataGridView_commands.Refresh();
                if (Accessory.PrintableHex(listBox_code.SelectedItem.ToString()))
                    row[ResultColumns.Description] = "\"" + Encoding.GetEncoding(Settings.Default.CodePage)
                        .GetString(Accessory.ConvertHexToByteArray(listBox_code.SelectedItem.ToString())) + "\"";
            }
        }

        private void Button_next_Click(object sender, EventArgs e)
        {
            if (listBox_code.SelectedIndex == -1)
            {
                if (listBox_code.Items.Count == 0) return;
                listBox_code.SelectedIndex = 0;
            }

            if (listBox_code.SelectedIndex < listBox_code.Items.Count - 1) listBox_code.SelectedIndex++;
            Button_find_Click(this, EventArgs.Empty);
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void SaveBinFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog.FileName = _sourceFile;
            saveFileDialog.Title = "Save BIN file";
            saveFileDialog.DefaultExt = "bin";
            saveFileDialog.Filter = "BIN files|*.bin|PRN files|*.prn|All files|*.*";
            saveFileDialog.ShowDialog();
        }

        private void SaveHexFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog.FileName = _sourceFile;
            saveFileDialog.Title = "Save HEX file";
            saveFileDialog.DefaultExt = "txt";
            saveFileDialog.Filter = "Text files|*.txt|HEX files|*.hex|All files|*.*";
            saveFileDialog.ShowDialog();
        }

        private void SaveCSVToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog.FileName = Settings.Default.CommandsDatabaseFile;
            saveFileDialog.Title = "Save CSV database";
            saveFileDialog.DefaultExt = "csv";
            saveFileDialog.Filter = "CSV files|*.csv|All files|*.*";
            saveFileDialog.ShowDialog();
        }

        private void SaveFileDialog_FileOk(object sender, CancelEventArgs e)
        {
            if (saveFileDialog.Title == "Save HEX file")
            {
                File.WriteAllText(saveFileDialog.FileName, "");
                foreach (string s in listBox_code.Items)
                    File.AppendAllText(saveFileDialog.FileName, s + "\r\n",
                        Encoding.GetEncoding(Settings.Default.CodePage));
            }
            else if (saveFileDialog.Title == "Save CSV database")
            {
                var columnCount = dataGridView_commands.ColumnCount;
                var output = new StringBuilder();
                for (var i = 0; i < columnCount; i++)
                {
                    output.Append(dataGridView_commands.Columns[i].Name);
                    output.Append(";");
                }

                output.Append("\r\n");
                for (var i = 0; i < dataGridView_commands.RowCount; i++)
                {
                    for (var j = 0; j < columnCount; j++)
                    {
                        output.Append(dataGridView_commands.Rows[i].Cells[j].Value);
                        output.Append(";");
                    }

                    output.Append("\r\n");
                }

                try
                {
                    File.WriteAllText(saveFileDialog.FileName, output.ToString(),
                        Encoding.GetEncoding(Settings.Default.CodePage));
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error writing to file " + saveFileDialog.FileName + ": " + ex.Message);
                }
            }
        }

        private void LoadBinToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog.FileName = "";
            openFileDialog.Title = "Open BIN file";
            openFileDialog.DefaultExt = "bin";
            openFileDialog.Filter = "BIN files|*.bin|PRN files|*.prn|All files|*.*";
            openFileDialog.ShowDialog();
        }

        private void LoadHexToolStripMenuItem_Click(object sender, EventArgs e)
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

        private void OpenFileDialog_FileOk(object sender, CancelEventArgs e)
        {
            if (openFileDialog.Title == "Open HEX file") //hex text read
            {
                _sourceFile = openFileDialog.FileName;
                listBox_code.Items.Clear();
                try
                {
                    foreach (var s in File.ReadAllLines(_sourceFile))
                        listBox_code.Items.Add(Accessory.CheckHexString(s));
                }
                catch (Exception ex)
                {
                    MessageBox.Show("\r\nError reading file " + _sourceFile + ": " + ex.Message);
                }

                //Form1.ActiveForm.Text += " " + SourceFile;
                //sourceData.Clear();
                //sourceData.AddRange(Accessory.ConvertHexToByteArray(textBox_code.Text));
                listBox_code.SelectedIndex = 0;
                //ParseEscPos.Init(listBox_code.Items[0].ToString(), CommandDatabase);
            }
            else if (openFileDialog.Title == "Open command CSV database") //hex text read
            {
                _commandDatabase = new DataTable();
                ReadCsv(openFileDialog.FileName, _commandDatabase);
                for (var i = 0; i < _commandDatabase.Rows.Count; i++)
                    _commandDatabase.Rows[i][0] = Accessory.CheckHexString(_commandDatabase.Rows[i][0].ToString());
                dataGridView_commands.DataSource = _commandDatabase;
                ParseEscPos.CommandDataBase = _commandDatabase;
            }
            else if (openFileDialog.Title == "Open errors CSV database") //hex text read
            {
                _commandDatabase = new DataTable();
                ReadCsv(openFileDialog.FileName, _errorsDatabase);
            }
        }

        private void DefaultCSVToolStripTextBox_Leave(object sender, EventArgs e)
        {
            if (commandsCSV_ToolStripTextBox.Text != Settings.Default.CommandsDatabaseFile)
            {
                Settings.Default.CommandsDatabaseFile = commandsCSV_ToolStripTextBox.Text;
                Settings.Default.Save();
            }
        }

        private void ErrorsCSV_toolStripTextBox_Leave(object sender, EventArgs e)
        {
            if (errorsCSV_toolStripTextBox.Text != Settings.Default.ErrorsDatabaseFile)
            {
                Settings.Default.ErrorsDatabaseFile = errorsCSV_toolStripTextBox.Text;
                Settings.Default.Save();
            }
        }

        private void EnableDatabaseEditToolStripMenuItem_Click(object sender, EventArgs e)
        {
            enableDatabaseEditToolStripMenuItem.Checked = !enableDatabaseEditToolStripMenuItem.Checked;
            dataGridView_commands.ReadOnly = !enableDatabaseEditToolStripMenuItem.Checked;
        }

        private string GetErrorDesc(int errNum)
        {
            for (var i = 0; i < _errorsDatabase.Rows.Count; i++)
                if (int.Parse(_errorsDatabase.Rows[i][0].ToString()) == errNum)
                    return _errorsDatabase.Rows[i][1].ToString();
            return "!!!Unknown error!!!";
        }

        private void DataGridView_result_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView_result.CellValueChanged -= DataGridView_result_CellValueChanged;
            if (dataGridView_result.CurrentCell.ColumnIndex == ResultColumns.Value)
            {
                if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value
                    .ToString() == ParseEscPos.DataTypes.Bitfield)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value =
                        ParseEscPos.BitfieldToRaw(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                            .Cells[ResultColumns.Value].Value.ToString());
                    var n = 0;
                    int.TryParse(
                        dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value]
                            .Value.ToString(), out n);
                    var i = dataGridView_result.CurrentRow.Index;
                    for (var i1 = 0; i1 < 8; i1++)
                        dataGridView_result.Rows[i + 1 + i1].Cells[ResultColumns.Value].Value =
                            Convert.ToInt32(Accessory.GetBit((byte) n, (byte) i1)).ToString();
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type]
                    .Value.ToString() == ParseEscPos.DataTypes.Data)
                {
                    var n = 0;
                    if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length]
                            .Value.ToString() ==
                        "?")
                        n = dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                            .Cells[ResultColumns.Value].Value.ToString().Length;
                    else
                        int.TryParse(
                            dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                                .Cells[ResultColumns.Length].Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value =
                        ParseEscPos.DataToRaw(
                            dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                                .Cells[ResultColumns.Value].Value.ToString(), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type]
                    .Value.ToString() == ParseEscPos.DataTypes.Error)
                {
                    var n = 0;
                    int.TryParse(
                        dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length]
                            .Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value =
                        ParseEscPos.ErrorToRaw(
                            dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                                .Cells[ResultColumns.Value].Value.ToString(), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type]
                    .Value.ToString() == ParseEscPos.DataTypes.Money)
                {
                    var n = 0;
                    int.TryParse(
                        dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length]
                            .Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value =
                        ParseEscPos.MoneyToRaw(
                            dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                                .Cells[ResultColumns.Value].Value.ToString(), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type]
                    .Value.ToString() == ParseEscPos.DataTypes.Number)
                {
                    var n = 0;
                    int.TryParse(
                        dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length]
                            .Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value =
                        ParseEscPos.NumberToRaw(
                            dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                                .Cells[ResultColumns.Value].Value.ToString(), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type]
                    .Value.ToString() == ParseEscPos.DataTypes.Password)
                {
                    var n = 0;
                    int.TryParse(
                        dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length]
                            .Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value =
                        ParseEscPos.PasswordToRaw(
                            dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                                .Cells[ResultColumns.Value].Value.ToString(), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type]
                    .Value.ToString() == ParseEscPos.DataTypes.Quantity)
                {
                    var n = 0;
                    int.TryParse(
                        dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length]
                            .Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value =
                        ParseEscPos.QuantityToRaw(
                            dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                                .Cells[ResultColumns.Value].Value.ToString(), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type]
                    .Value.ToString() == ParseEscPos.DataTypes.String)
                {
                    var n = 0;
                    if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length]
                            .Value.ToString() ==
                        "?")
                        n = dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                            .Cells[ResultColumns.Value].Value.ToString().Length;
                    else
                        int.TryParse(
                            dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                                .Cells[ResultColumns.Length].Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value =
                        ParseEscPos.StringToRaw(
                            dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                                .Cells[ResultColumns.Value].Value.ToString(), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type]
                    .Value.ToString() == ParseEscPos.DataTypes.PrefData)
                {
                    var n = 0;
                    if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length]
                            .Value.ToString() ==
                        "?")
                        n = dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                            .Cells[ResultColumns.Value].Value.ToString().Length;
                    else
                        int.TryParse(
                            dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                                .Cells[ResultColumns.Length].Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value =
                        ParseEscPos.PrefDataToRaw(
                            dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                                .Cells[ResultColumns.Value].Value.ToString(), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type]
                    .Value.ToString() == ParseEscPos.DataTypes.TLVData)
                {
                    var n = 0;
                    if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length]
                            .Value.ToString() ==
                        "?")
                        n = dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                            .Cells[ResultColumns.Value].Value.ToString().Length;
                    else
                        int.TryParse(
                            dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                                .Cells[ResultColumns.Length].Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value =
                        ParseEscPos.TLVDataToRaw(
                            dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                                .Cells[ResultColumns.Value].Value.ToString(), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type]
                    .Value.ToString().StartsWith("bit"))
                {
                    var i = dataGridView_result.CurrentCell.RowIndex - 1;
                    while (dataGridView_result.Rows[i].Cells[ResultColumns.Type].Value.ToString() !=
                           ParseEscPos.DataTypes.Bitfield) i--;
                    //collect bits to int
                    byte n = 0;
                    for (var i1 = 0; i1 < 8; i1++)
                        if (dataGridView_result.Rows[i + i1 + 1].Cells[ResultColumns.Value].Value.ToString().Trim() ==
                            "1")
                            n += (byte) Math.Pow(2, i1);
                    dataGridView_result.Rows[i].Cells[ResultColumns.Value].Value = n.ToString();
                    dataGridView_result.Rows[i].Cells[ResultColumns.Raw].Value =
                        ParseEscPos.BitfieldToRaw(dataGridView_result.Rows[i].Cells[ResultColumns.Value].Value
                            .ToString());
                }
            }
            else if (dataGridView_result.CurrentCell.ColumnIndex == ResultColumns.Raw)
            {
                dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value =
                    Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                        .Cells[ResultColumns.Raw].Value.ToString());
                if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type].Value
                    .ToString() == ParseEscPos.DataTypes.Password)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value =
                        Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                            .Cells[ResultColumns.Raw].Value.ToString());
                    var l = ParseEscPos.RawToPassword(Accessory.ConvertHexToByteArray(dataGridView_result
                        .Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString()));
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value]
                        .Value = l.ToString();
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type]
                    .Value.ToString() == ParseEscPos.DataTypes.String)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value =
                        Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                            .Cells[ResultColumns.Raw].Value.ToString());
                    var n = 0;
                    if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length]
                            .Value.ToString() ==
                        "?")
                        n = dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw]
                            .Value.ToString().Length / 3;
                    else
                        int.TryParse(
                            dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                                .Cells[ResultColumns.Length].Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value
                        = ParseEscPos.RawToString(
                            Accessory.ConvertHexToByteArray(dataGridView_result
                                .Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value
                                .ToString()), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type]
                    .Value.ToString() == ParseEscPos.DataTypes.PrefData)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value =
                        Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                            .Cells[ResultColumns.Raw].Value.ToString());
                    var n = 0;
                    if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length]
                            .Value.ToString() ==
                        "?")
                        n = dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw]
                            .Value.ToString().Length / 3;
                    else
                        int.TryParse(
                            dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                                .Cells[ResultColumns.Length].Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value
                        = ParseEscPos.RawToPrefData(
                            Accessory.ConvertHexToByteArray(dataGridView_result
                                .Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value
                                .ToString()), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type]
                    .Value.ToString() == ParseEscPos.DataTypes.TLVData)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value =
                        Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                            .Cells[ResultColumns.Raw].Value.ToString());
                    var n = 0;
                    if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Length]
                            .Value.ToString() ==
                        "?")
                        n = dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw]
                            .Value.ToString().Length / 3;
                    else
                        int.TryParse(
                            dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                                .Cells[ResultColumns.Length].Value.ToString(), out n);
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value
                        = ParseEscPos.RawToTLVData(
                            Accessory.ConvertHexToByteArray(dataGridView_result
                                .Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value
                                .ToString()), n);
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type]
                    .Value.ToString() == ParseEscPos.DataTypes.Number)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value =
                        Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                            .Cells[ResultColumns.Raw].Value.ToString());
                    var l = ParseEscPos.RawToNumber(Accessory.ConvertHexToByteArray(dataGridView_result
                        .Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString()));
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value]
                        .Value = l.ToString();
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type]
                    .Value.ToString() == ParseEscPos.DataTypes.Money)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value =
                        Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                            .Cells[ResultColumns.Raw].Value.ToString());
                    var l = ParseEscPos.RawToMoney(Accessory.ConvertHexToByteArray(dataGridView_result
                        .Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString()));
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value]
                        .Value = l.ToString();
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type]
                    .Value.ToString() == ParseEscPos.DataTypes.Quantity)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value =
                        Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                            .Cells[ResultColumns.Raw].Value.ToString());
                    var l = ParseEscPos.RawToQuantity(Accessory.ConvertHexToByteArray(dataGridView_result
                        .Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString()));
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value]
                        .Value = l.ToString();
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type]
                    .Value.ToString() == ParseEscPos.DataTypes.Error)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value =
                        Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                            .Cells[ResultColumns.Raw].Value.ToString());
                    var l = ParseEscPos.RawToError(Accessory.ConvertHexToByteArray(dataGridView_result
                        .Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString()));
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value]
                        .Value = l.ToString();
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type]
                    .Value.ToString() == ParseEscPos.DataTypes.Data)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value =
                        Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                            .Cells[ResultColumns.Raw].Value.ToString());
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value].Value
                        = ParseEscPos.RawToData(Accessory.ConvertHexToByteArray(dataGridView_result
                            .Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString()));
                }
                else if (dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Type]
                    .Value.ToString() == ParseEscPos.DataTypes.Bitfield)
                {
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value =
                        Accessory.CheckHexString(dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex]
                            .Cells[ResultColumns.Raw].Value.ToString());
                    var l = ParseEscPos.RawToBitfield(Accessory.ConvertHexToByte(dataGridView_result
                        .Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Raw].Value.ToString()));
                    dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value]
                        .Value = l.ToString();
                    var n = 0;
                    int.TryParse(
                        dataGridView_result.Rows[dataGridView_result.CurrentCell.RowIndex].Cells[ResultColumns.Value]
                            .Value.ToString(), out n);
                    var i = dataGridView_result.CurrentRow.Index;
                    for (var i1 = 0; i1 < 8; i1++)
                        dataGridView_result.Rows[i + 1 + i1].Cells[ResultColumns.Value].Value =
                            (Accessory.GetBit((byte) n, (byte) i1) ? (byte) 1 : (byte) 0).ToString();
                }
            }

            dataGridView_result.CellValueChanged += DataGridView_result_CellValueChanged;
        }

        private void DataGridView_commands_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Button_newCommand_Click(this, EventArgs.Empty);
        }

        private string CollectCommand()
        {
            var data = "";
            for (var i = 0; i < _resultDatabase.Rows.Count; i++)
                data += _resultDatabase.Rows[i][ResultColumns.Raw].ToString();
            var length = new byte[2];
            length[1] = (byte) (data.Length / 3 / 256);
            length[0] = (byte) (data.Length / 3 - length[1]);
            data = Accessory.ConvertByteArrayToHex(length) + data;
            var dataByte = Accessory.ConvertHexToByteArray(data);
            return "01 " + data + Accessory.ConvertByteToHex(ParseEscPos.Q3xf_CRC(dataByte, dataByte.Length));
        }

        private void Button_add_Click(object sender, EventArgs e)
        {
            listBox_code.Items.Add(CollectCommand());
            listBox_code.SelectedIndex = listBox_code.Items.Count - 1;
        }

        private void Button_replace_Click(object sender, EventArgs e)
        {
            if (listBox_code.SelectedIndex == -1) return;
            listBox_code.Items[listBox_code.SelectedIndex] = CollectCommand();
        }

        private void Button_insert_Click(object sender, EventArgs e)
        {
            if (listBox_code.SelectedIndex == -1) return;
            listBox_code.Items.Insert(listBox_code.SelectedIndex, CollectCommand());
            listBox_code.SelectedIndex--;
        }

        private void Button_remove_Click(object sender, EventArgs e)
        {
            for (var i = 0; i < listBox_code.Items.Count; i++)
                if (listBox_code.Items[i].ToString().StartsWith("06 "))
                {
                    listBox_code.Items.RemoveAt(i);
                    i--;
                }
        }

        private void Button_clear_Click(object sender, EventArgs e)
        {
            listBox_code.Items.Clear();
        }

        private void TextBox_search_TextChanged(object sender, EventArgs e)
        {
            dataGridView_commands.CurrentCell = null;
            DataGridViewRow row;
            if (textBox_search.Text != "")
                for (var i = 0; i < dataGridView_commands.RowCount; i++)
                {
                    row = dataGridView_commands.Rows[i];
                    if (dataGridView_commands.Rows[i].Cells[ParseEscPos.CSVColumns.CommandNameColumn].Value
                        .ToString() != "")
                    {
                        if (dataGridView_commands.Rows[i].Cells[ParseEscPos.CSVColumns.CommandDescriptionColumn].Value
                            .ToString().ToLower().Contains(textBox_search.Text.ToLower()))
                        {
                            row.Visible = true;
                            i++;
                            while (i < dataGridView_commands.RowCount && dataGridView_commands.Rows[i]
                                .Cells[ParseEscPos.CSVColumns.CommandNameColumn].Value.ToString() == "")
                            {
                                row = dataGridView_commands.Rows[i];
                                row.Visible = true;
                                i++;
                            }

                            i--;
                        }
                        else
                        {
                            row.Visible = false;
                        }
                    }
                    else
                    {
                        row.Visible = false;
                    }
                }
            else
                for (var i = 0; i < dataGridView_commands.RowCount; i++)
                {
                    row = dataGridView_commands.Rows[i];
                    row.Visible = true;
                }
        }

        private void Button_newCommand_Click(object sender, EventArgs e)
        {
            //restore 
            ParseEscPos.CSVColumns.CommandParameterSizeColumn = 1;
            ParseEscPos.CSVColumns.CommandParameterTypeColumn = 2;
            ParseEscPos.CSVColumns.CommandParameterValueColumn = 3;
            ParseEscPos.CSVColumns.CommandDescriptionColumn = 4;
            ParseEscPos.ItIsReply = false;

            //dataGridView_commands_CellDoubleClick(this, new DataGridViewCellEventArgs(this.dataGridView_commands.CurrentCell.ColumnIndex, this.dataGridView_commands.CurrentRow.Index));
            if (dataGridView_commands.Rows[dataGridView_commands.CurrentCell.RowIndex]
                .Cells[ParseEscPos.CSVColumns.CommandNameColumn].Value.ToString() != "")
            {
                var currentRow = dataGridView_commands.CurrentCell.RowIndex;
                _resultDatabase.Clear();
                var row = _resultDatabase.NewRow();
                row[ResultColumns.Value] = dataGridView_commands.Rows[currentRow]
                    .Cells[ParseEscPos.CSVColumns.CommandNameColumn].Value.ToString();
                row[ResultColumns.Raw] = row[ResultColumns.Value];
                row[ResultColumns.Description] = dataGridView_commands.Rows[currentRow]
                    .Cells[ParseEscPos.CSVColumns.CommandDescriptionColumn].Value.ToString();
                _resultDatabase.Rows.Add(row);

                var i = currentRow + 1;
                while (i < dataGridView_commands.Rows.Count && dataGridView_commands.Rows[i]
                    .Cells[ParseEscPos.CSVColumns.CommandNameColumn].Value.ToString() == "")
                {
                    if (dataGridView_commands.Rows[i].Cells[ParseEscPos.CSVColumns.CommandParameterSizeColumn].Value
                        .ToString() != "")
                    {
                        row = _resultDatabase.NewRow();
                        row[ResultColumns.Type] = dataGridView_commands.Rows[i]
                            .Cells[ParseEscPos.CSVColumns.CommandParameterTypeColumn].Value.ToString();
                        row[ResultColumns.Length] = dataGridView_commands.Rows[i]
                            .Cells[ParseEscPos.CSVColumns.CommandParameterSizeColumn].Value.ToString();
                        row[ResultColumns.Description] = dataGridView_commands.Rows[i]
                            .Cells[ParseEscPos.CSVColumns.CommandDescriptionColumn].Value.ToString();
                        if (row[ResultColumns.Type].ToString() == ParseEscPos.DataTypes.Password)
                        {
                            row[ResultColumns.Value] = textBox_password.Text;
                            var n = 0;
                            int.TryParse(row[ResultColumns.Length].ToString(), out n);
                            row[ResultColumns.Raw] = ParseEscPos.PasswordToRaw(textBox_password.Text, n);
                        }
                        else
                        {
                            row[ResultColumns.Value] = "";
                            row[ResultColumns.Raw] = "";
                        }

                        _resultDatabase.Rows.Add(row);
                        if (row[ResultColumns.Type].ToString() == ParseEscPos.DataTypes.Bitfield) //decode bitfield
                            for (var i1 = 0; i1 < 8; i1++)
                            {
                                row = _resultDatabase.NewRow();
                                row[ResultColumns.Value] = "0";
                                row[ResultColumns.Type] = "bit" + i1;
                                row[ResultColumns.Description] = dataGridView_commands.Rows[i + i1 + 1]
                                    .Cells[ParseEscPos.CSVColumns.CommandDescriptionColumn].Value.ToString();
                                _resultDatabase.Rows.Add(row);
                            }
                    }

                    i++;
                }
            }
        }

        private void ListBox_code_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Button_find_Click(this, EventArgs.Empty);
        }

        private void ListBox_code_KeyDown(object sender, KeyEventArgs e)
        {
            if (sender != listBox_code) return;

            if (listBox_code.SelectedIndex == -1) return;
            //Ctrl-C - copy string to clipboard
            if (e.Control && e.KeyCode == Keys.C && listBox_code.SelectedItem.ToString() != "")
            {
                Clipboard.SetText(listBox_code.SelectedItem.ToString());
            }
            //Ctrl-Ins - copy string to clipboard
            else if (e.Control && e.KeyCode == Keys.Insert && listBox_code.SelectedItem.ToString() != "")
            {
                Clipboard.SetText(listBox_code.SelectedItem.ToString());
            }
            //Ctrl-V - insert string from clipboard
            else if (e.Control && e.KeyCode == Keys.V && Accessory.GetStringFormat(Clipboard.GetText()) == 16)
            {
                listBox_code.Items[listBox_code.SelectedIndex] = Accessory.CheckHexString(Clipboard.GetText());
            }
            //Shift-Ins - insert string from clipboard
            else if (e.Shift && e.KeyCode == Keys.Insert && Accessory.GetStringFormat(Clipboard.GetText()) == 16)
            {
                listBox_code.Items[listBox_code.SelectedIndex] = Accessory.CheckHexString(Clipboard.GetText());
            }
            //DEL - delete string
            else if (e.KeyCode == Keys.Delete)
            {
                var i = listBox_code.SelectedIndex;
                listBox_code.Items.RemoveAt(listBox_code.SelectedIndex);
                if (i >= listBox_code.Items.Count) i = listBox_code.Items.Count - 1;
                listBox_code.SelectedIndex = i;
            }
            //Ctrl-P - parse string
            else if (e.Control && e.KeyCode == Keys.P)
            {
                Button_find_Click(this, EventArgs.Empty);
            }
            //Ctrl-S - send string to device
            else if (e.Control && e.KeyCode == Keys.S && button_Send.Enabled)
            {
                Button_Send_Click(this, EventArgs.Empty);
            }
        }

        private void ToolStripMenuItem_Connect_Click(object sender, EventArgs e)
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
                    SerialPort1.Handshake =
                        (Handshake) Enum.Parse(typeof(Handshake), toolStripComboBox_PortHandshake.Text);
                    SerialPort1.Parity = (Parity) Enum.Parse(typeof(Parity), toolStripComboBox_PortParity.Text);
                    SerialPort1.StopBits = (StopBits) Enum.Parse(typeof(StopBits), toolStripComboBox_PortStopBits.Text);
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

        private void Button_Send_Click(object sender, EventArgs e)
        {
            if (listBox_code.SelectedIndex == -1) return;
            if (listBox_code.SelectedItem.ToString() == "") return;

            if (listBox_code.SelectedItem.ToString().Substring(0, 2) == "01")
            {
                var _txBytes = Accessory.ConvertHexToByteArray(listBox_code.SelectedItem.ToString());
                try
                {
                    SerialPort1.Write(_txBytes, 0, _txBytes.Length);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error sending to port " + SerialPort1.PortName + ": " + ex.Message);
                }

                var _rxBytes = new List<byte>();
                var _timeout = false;
                var _nackOK = false;
                var _frameOK = false;
                var _lengthOK = false;
                var _frameLength = 0;
                var startTime = DateTime.UtcNow;
                try
                {
                    //!!! rework byte reading direct to array
                    while (!_timeout)
                    {
                        var c = -1;
                        if (!_nackOK)
                        {
                            if (SerialPort1.BytesToRead > 0) c = SerialPort1.ReadByte();
                            if (c == 06)
                            {
                                _rxBytes.Add((byte) c);
                                _nackOK = true;
                            }
                        }
                        else if (!_frameOK)
                        {
                            if (SerialPort1.BytesToRead > 0) c = SerialPort1.ReadByte();
                            if (c == 01)
                            {
                                _rxBytes.Add((byte) c);
                                _frameOK = true;
                            }
                        }
                        else if (!_lengthOK)
                        {
                            if (SerialPort1.BytesToRead >= 2)
                            {
                                c = SerialPort1.ReadByte();
                                _rxBytes.Add((byte) c);
                                _frameLength = c;
                                c = SerialPort1.ReadByte();
                                _rxBytes.Add((byte) c);
                                _frameLength = _frameLength + c * 256;
                                _lengthOK = true;
                            }
                        }
                        else if (_lengthOK)
                        {
                            if (SerialPort1.BytesToRead > 0) _rxBytes.Add((byte) SerialPort1.ReadByte());
                        }
                        else
                        {
                            SerialPort1.ReadByte();
                        }

                        if (_rxBytes.Count > _frameLength + 4) _timeout = true;
                        if (SerialPort1.BytesToRead == 0 &&
                            DateTime.UtcNow.Subtract(startTime).TotalMilliseconds > _serialtimeOut) _timeout = true;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error reading port " + SerialPort1.PortName + ": " + ex.Message);
                }

                if (_rxBytes.Count > _frameLength && _frameOK && _lengthOK ||
                    showIncorrectRepliesToolStripMenuItem.Checked)
                {
                    var data = Accessory.ConvertByteArrayToHex(_rxBytes.ToArray());
                    if (listBox_code.SelectedIndex + 1 >= listBox_code.Items.Count) listBox_code.Items.Add(data);
                    else if (listBox_code.Items[listBox_code.SelectedIndex + 1].ToString().Length > 2 &&
                             listBox_code.Items[listBox_code.SelectedIndex + 1].ToString().Substring(0, 2) == "06")
                        listBox_code.Items[listBox_code.SelectedIndex + 1] = data;
                    else listBox_code.Items.Insert(listBox_code.SelectedIndex + 1, data);
                    if (autoParseReplyToolStripMenuItem.Checked) Button_next_Click(this, EventArgs.Empty);
                }
            }
        }

        private void Button_SendAll_Click(object sender, EventArgs e)
        {
            if (listBox_code.SelectedIndex == -1) listBox_code.SelectedIndex = 0;
            for (var i = listBox_code.SelectedIndex; i < listBox_code.Items.Count; i++)
            {
                listBox_code.SelectedIndex = i;
                if (listBox_code.SelectedItem.ToString().Length > 2)
                {
                    if (listBox_code.SelectedItem.ToString().Substring(0, 2) == "06")
                    {
                        if (ToolStripMenuItem_stopOnErrorReplied.Checked &&
                            listBox_code.SelectedItem.ToString().Length > 20 &&
                            listBox_code.SelectedItem.ToString().Substring(5 * 3, 5) != "00 00") return;
                    }
                    else if (listBox_code.SelectedItem.ToString().Substring(0, 2) == "15")
                    {
                        return;
                    }
                    else
                    {
                        Button_Send_Click(button_SendAll, EventArgs.Empty);
                    }
                }
            }
        }

        private void SerialPopulate()
        {
            toolStripComboBox_PortName.Items.Clear();
            toolStripComboBox_PortHandshake.Items.Clear();
            toolStripComboBox_PortParity.Items.Clear();
            toolStripComboBox_PortStopBits.Items.Clear();
            //Serial settings populate
            toolStripComboBox_PortName.Items.Add("-None-");
            //Add ports
            foreach (var s in SerialPort.GetPortNames()) toolStripComboBox_PortName.Items.Add(s);
            //Add handshake methods
            foreach (var s in Enum.GetNames(typeof(Handshake))) toolStripComboBox_PortHandshake.Items.Add(s);
            //Add parity
            foreach (var s in Enum.GetNames(typeof(Parity))) toolStripComboBox_PortParity.Items.Add(s);
            //Add stopbits
            foreach (var s in Enum.GetNames(typeof(StopBits))) toolStripComboBox_PortStopBits.Items.Add(s);
            toolStripComboBox_PortName.SelectedIndex = toolStripComboBox_PortName.Items.Count - 1;
            if (toolStripComboBox_PortName.Items.Count == 1) toolStripMenuItem_Connect.Enabled = false;
            toolStripComboBox_PortSpeed.SelectedIndex = 0;
            toolStripComboBox_PortHandshake.SelectedIndex = 0;
            toolStripComboBox_PortDataBits.SelectedIndex = 0;
            toolStripComboBox_PortParity.SelectedIndex = 0;
            toolStripComboBox_PortStopBits.SelectedIndex = 1;
            if (toolStripComboBox_PortName.SelectedIndex == 0) toolStripMenuItem_Connect.Enabled = false;
            else toolStripMenuItem_Connect.Enabled = true;
        }

        private void ToolStripTextBox_TimeOut_TextChanged(object sender, EventArgs e)
        {
            if (!int.TryParse(toolStripTextBox_TimeOut.Text, out _serialtimeOut))
                toolStripTextBox_TimeOut.Text = "1000";
        }

        private void ToolStripMenuItem_stopOnErrorReplied_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem_stopOnErrorReplied.Checked = !ToolStripMenuItem_stopOnErrorReplied.Checked;
        }

        private void ListBox_code_MouseUp(object sender, MouseEventArgs e)
        {
            var index = listBox_code.IndexFromPoint(e.Location);
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

        private void CopyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listBox_code.SelectedItem.ToString() != "") Clipboard.SetText(listBox_code.SelectedItem.ToString());
        }

        private void DeleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listBox_code.SelectedIndex == -1) return;
            listBox_code.Items.RemoveAt(listBox_code.SelectedIndex);
        }

        private void ParseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Button_find_Click(this, EventArgs.Empty);
        }

        private void SendToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Button_Send_Click(this, EventArgs.Empty);
        }

        private void PasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Accessory.GetStringFormat(Clipboard.GetText()) == 16)
                listBox_code.Items[listBox_code.SelectedIndex] = Accessory.CheckHexString(Clipboard.GetText());
        }

        private void NewCommandToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Button_newCommand_Click(this, EventArgs.Empty);
        }

        private void FindThisToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Button_find_Click(findThisToolStripMenuItem, EventArgs.Empty);
        }

        private void DataGridView_commands_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && e.RowIndex >= 0)
                dataGridView_commands.CurrentCell = dataGridView_commands.Rows[e.RowIndex].Cells[e.ColumnIndex];
        }

        private void COMPortToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SerialPopulate();
        }

        private void showIncorrectRepliesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            showIncorrectRepliesToolStripMenuItem.Checked = !showIncorrectRepliesToolStripMenuItem.Checked;
        }

        private void autoParseReplyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            autoParseReplyToolStripMenuItem.Checked = !autoParseReplyToolStripMenuItem.Checked;
        }
    }
}