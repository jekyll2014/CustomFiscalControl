using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using CustomFiscalControl.Properties;

public class ParseEscPos
{
    //INTERFACES
    //source of the data to parce
    //public static string sourceData = ""; //in Init()
    public static List<byte> SourceData = new List<byte>(); //in Init()

    //source of the command description (DataTable)
    public static DataTable CommandDataBase = new DataTable(); //in Init()

    //INTERNAL VARIABLES
    //Command list preselected
    //private static Dictionary<int, string> _commandList = new Dictionary<int, string>(); //in Init()

    private const byte AckSign = 0x06;
    private const byte NackSign = 0x15;
    private const byte FrameStartSign = 0x01;
    public static bool ItIsReply;
    public static bool ItIsReplyNack;
    public static bool CrcFailed;
    public static bool LengthIncorrect;

    //RESULT VALUES
    public static int CommandFrameLength;

    //place of the frame start in the text
    public static int CommandFramePosition; //in findCommand()

    //Command text
    public static string CommandName; //in findCommand()

    //Command desc
    public static string CommandDesc; //in findCommand()

    //string number of the command found
    public static int CommandDbLineNum; //in findCommand()

    //height of the command
    public static int CommandDbHeight; //in findCommand()

    //string number of the command found
    public static List<int> CommandParamDbLineNum = new List<int>(); //in findCommand()

    //list of command parameters real sizes
    public static List<int> CommandParamSize = new List<int>(); //in findCommand()

    //list of command parameters sizes defined in the database
    public static List<string> CommandParamSizeDefined = new List<string>(); //in findCommand()

    //command parameter description
    public static List<string> CommandParamDesc = new List<string>(); //in findCommand()

    //command parameter type
    public static List<string> CommandParamType = new List<string>(); //in findCommand()

    //command parameter RAW value
    public static List<List<byte>> CommandParamRawValue = new List<List<byte>>(); //in findCommand()

    //command parameter value
    public static List<string> CommandParamValue = new List<string>(); //in findCommand()

    //Length of command+parameters text
    public static int CommandBlockLength;

    public class CSVColumns
    {
        public static int CommandNameColumn { get; set; } = 0;
        public static int CommandParameterSizeColumn { get; set; } = 1;
        public static int CommandParameterTypeColumn { get; set; } = 2;
        public static int CommandParameterValueColumn { get; set; } = 3;
        public static int CommandDescriptionColumn { get; set; } = 4;
        public static int ReplyParameterSizeColumn { get; set; } = 5;
        public static int ReplyParameterTypeColumn { get; set; } = 6;
        public static int ReplyParameterValueColumn { get; set; } = 7;
        public static int ReplyDescriptionColumn { get; set; } = 8;
    }

    public class DataTypes
    {
        public static string Password { get; set; } = "password";
        public static string String { get; set; } = "string";
        public static string Number { get; set; } = "number";
        public static string Money { get; set; } = "money";
        public static string Quantity { get; set; } = "quantity";
        public static string Error { get; set; } = "error#";
        public static string Data { get; set; } = "data";
        public static string PrefData { get; set; } = "prefdata";
        public static string TLVData { get; set; } = "tlvdata";
        public static string Bitfield { get; set; } = "bitfield";
    }

    //lineNum = -1 - искать во всех командах
    //lineNum = x - искать в команде на определенной стоке базы
    public static bool FindCommand(int _pos, int lineNum = -1)
    {
        //reset all result values
        ClearCommand();

        if (SourceData.Count < _pos + 1) return false;
        //check if it's a command or reply
        if (SourceData[_pos] == FrameStartSign)
        {
            CSVColumns.CommandParameterSizeColumn = 1;
            CSVColumns.CommandParameterTypeColumn = 2;
            CSVColumns.CommandParameterValueColumn = 3;
            CSVColumns.CommandDescriptionColumn = 4;
            ItIsReply = false;
            if (SourceData.Count < _pos + 3) return false;
        }
        else if (SourceData[_pos] == AckSign || SourceData[_pos] == NackSign)
        {
            ItIsReply = true;
            if (SourceData[_pos] == NackSign) ItIsReplyNack = true;
            CSVColumns.CommandParameterSizeColumn = 5;
            CSVColumns.CommandParameterTypeColumn = 6;
            CSVColumns.CommandParameterValueColumn = 7;
            CSVColumns.CommandDescriptionColumn = 8;
            if (SourceData.Count < _pos + 4) return false;
            _pos++;
        }
        else
        {
            return false;
        }

        //select data frame
        CommandFrameLength = 0;
        if (SourceData[_pos] == FrameStartSign)
        {
            CommandFrameLength = SourceData[_pos + 1];
            CommandFrameLength = CommandFrameLength + SourceData[_pos + 2] * 256;
            _pos += 3;
        }
        else
        {
            return false;
        }

        //check if "commandFrameLength" less than "sourcedata". note the last byte of "sourcedata" is CRC.
        if (SourceData.Count < _pos + CommandFrameLength + 1)
        {
            CommandFrameLength = SourceData.Count - _pos;
            LengthIncorrect = true;
        }

        //find command
        if (SourceData.Count < _pos + 1) return false; //check if it doesn't go over the last symbol
        var i = 0;
        if (lineNum != -1) i = lineNum;
        for (; i < CommandDataBase.Rows.Count; i++)
            if (CommandDataBase.Rows[i][CSVColumns.CommandNameColumn].ToString() != "")
                if (SourceData[_pos] ==
                    Accessory.ConvertHexToByte(CommandDataBase.Rows[i][CSVColumns.CommandNameColumn].ToString())
                ) //if command matches
                    if (lineNum < 0 || lineNum == i) //if string matches
                    {
                        CommandName = CommandDataBase.Rows[i][CSVColumns.CommandNameColumn].ToString();
                        CommandDbLineNum = i;
                        CommandDesc = CommandDataBase.Rows[i][CSVColumns.CommandDescriptionColumn].ToString();
                        CommandFramePosition = _pos;
                        //get CRC of the frame
                        //check length of sourceData
                        int calculatedCRC = Q3xf_CRC(SourceData.GetRange(_pos - 2, CommandFrameLength + 2).ToArray(),
                            CommandFrameLength + 2);
                        int sentCRC = SourceData[_pos + CommandFrameLength];
                        if (calculatedCRC != sentCRC) CrcFailed = true;
                        else CrcFailed = false;
                        //check command height - how many rows are occupated
                        var i1 = 0;
                        while (CommandDbLineNum + i1 + 1 < CommandDataBase.Rows.Count &&
                               CommandDataBase.Rows[CommandDbLineNum + i1 + 1][CSVColumns.CommandNameColumn].ToString() ==
                               "") i1++;
                        CommandDbHeight = i1;
                        return true;
                    }

        return false;
    }

    public static bool FindCommandParameter()
    {
        ClearCommandParameters();
        //collect parameters from database
        var _stopSearch = CommandDbLineNum + 1;
        while (_stopSearch < CommandDataBase.Rows.Count &&
               CommandDataBase.Rows[_stopSearch][CSVColumns.CommandNameColumn].ToString() == "") _stopSearch++;
        for (var i = CommandDbLineNum + 1; i < _stopSearch; i++)
            if (CommandDataBase.Rows[i][CSVColumns.CommandParameterSizeColumn].ToString() != "")
            {
                CommandParamDbLineNum.Add(i);
                CommandParamSizeDefined.Add(CommandDataBase.Rows[i][CSVColumns.CommandParameterSizeColumn].ToString());
                if (CommandParamSizeDefined.Last() == "?")
                {
                    CommandParamSize.Add(CommandFrameLength - 1);
                    for (var i1 = 0; i1 < CommandParamSize.Count - 1; i1++)
                        CommandParamSize[CommandParamSize.Count - 1] -= CommandParamSize[i1];
                    if (CommandParamSize[CommandParamSize.Count - 1] < 0)
                        CommandParamSize[CommandParamSize.Count - 1] = 0;
                }
                else
                {
                    var v = 0;
                    int.TryParse(CommandParamSizeDefined.Last(), out v);
                    CommandParamSize.Add(v);
                }

                CommandParamDesc.Add(CommandDataBase.Rows[i][CSVColumns.CommandDescriptionColumn].ToString());
                CommandParamType.Add(CommandDataBase.Rows[i][CSVColumns.CommandParameterTypeColumn].ToString());
            }

        //recalculate "?" according to the size of parameters after it.
        for (var i = 0; i < CommandParamSizeDefined.Count - 1; i++)
            if (CommandParamSizeDefined[i] == "?")
            {
                for (var i1 = i + 1; i1 < CommandParamSize.Count; i1++) CommandParamSize[i] -= CommandParamSize[i1];
                i = CommandParamSizeDefined.Count;
            }

        var commandParamPosition = CommandFramePosition + 1;
        //process each parameter
        for (var parameter = 0; parameter < CommandParamDbLineNum.Count; parameter++)
        {
            //collect predefined RAW values
            var predefinedParamsRaw = new List<string>();
            var j = CommandParamDbLineNum[parameter] + 1;
            while (j < CommandDataBase.Rows.Count &&
                   CommandDataBase.Rows[j][CSVColumns.CommandParameterValueColumn].ToString() != "")
            {
                predefinedParamsRaw.Add(CommandDataBase.Rows[j][CSVColumns.CommandParameterValueColumn].ToString());
                j++;
            }

            //Calculate predefined params
            var predefinedParamsVal = new List<int>();
            foreach (var formula in predefinedParamsRaw)
            {
                var val = 0;
                if (!int.TryParse(formula.Trim(), out val)) val = 0;
                predefinedParamsVal.Add(val);
            }

            //get parameter from text
            var errFlag = false; //Error in parameter found
            var errMessage = "";

            var _prmType = CommandDataBase.Rows[CommandParamDbLineNum[parameter]][CSVColumns.CommandParameterTypeColumn]
                .ToString().ToLower();
            if (parameter != 0) commandParamPosition = commandParamPosition + CommandParamSize[parameter - 1];
            var _raw = new List<byte>();
            var _val = "";

            if (_prmType == DataTypes.Password)
            {
                double l = 0;
                if (commandParamPosition + CommandParamSize[parameter] <= SourceData.Count - 1)
                {
                    _raw = SourceData.GetRange(commandParamPosition, CommandParamSize[parameter]);
                    l = RawToPassword(_raw.ToArray());
                    _val = l.ToString().Substring(0, 2) + "/" + l.ToString().Substring(2);
                }
                else
                {
                    errFlag = true;
                    errMessage = "!!!ERR: Out of data bound!!!";
                    if (commandParamPosition <= SourceData.Count - 1)
                        _raw = SourceData.GetRange(commandParamPosition, SourceData.Count - 1 - commandParamPosition);
                }
            }
            else if (_prmType == DataTypes.String)
            {
                if (commandParamPosition + CommandParamSize[parameter] <= SourceData.Count - 1)
                {
                    _raw = SourceData.GetRange(commandParamPosition, CommandParamSize[parameter]);
                    _val = RawToString(_raw.ToArray(), CommandParamSize[parameter]);
                }
                else
                {
                    errFlag = true;
                    errMessage = "!!!ERR: Out of data bound!!!";
                    if (commandParamPosition <= SourceData.Count - 1)
                        _raw = SourceData.GetRange(commandParamPosition, SourceData.Count - 1 - commandParamPosition);
                }
            }
            else if (_prmType == DataTypes.Number)
            {
                double l = 0;
                if (commandParamPosition + CommandParamSize[parameter] <= SourceData.Count - 1)
                {
                    _raw = SourceData.GetRange(commandParamPosition, CommandParamSize[parameter]);
                    l = RawToNumber(_raw.ToArray());
                    _val = l.ToString();
                }
                else
                {
                    errFlag = true;
                    errMessage = "!!!ERR: Out of data bound!!!";
                    if (commandParamPosition <= SourceData.Count - 1)
                        _raw = SourceData.GetRange(commandParamPosition, SourceData.Count - 1 - commandParamPosition);
                }
            }
            else if (_prmType == DataTypes.Money)
            {
                double l = 0;
                if (commandParamPosition + CommandParamSize[parameter] <= SourceData.Count - 1)
                {
                    _raw = SourceData.GetRange(commandParamPosition, CommandParamSize[parameter]);
                    l = RawToMoney(_raw.ToArray());
                    _val = l.ToString();
                }
                else
                {
                    errFlag = true;
                    errMessage = "!!!ERR: Out of data bound!!!";
                    if (commandParamPosition <= SourceData.Count - 1)
                        _raw = SourceData.GetRange(commandParamPosition, SourceData.Count - 1 - commandParamPosition);
                }
            }
            else if (_prmType == DataTypes.Quantity)
            {
                double l = 0;
                if (commandParamPosition + CommandParamSize[parameter] <= SourceData.Count - 1)
                {
                    _raw = SourceData.GetRange(commandParamPosition, CommandParamSize[parameter]);
                    l = RawToQuantity(_raw.ToArray());
                    _val = l.ToString();
                }
                else
                {
                    errFlag = true;
                    errMessage = "!!!ERR: Out of data bound!!!";
                    if (commandParamPosition <= SourceData.Count - 1)
                        _raw = SourceData.GetRange(commandParamPosition, SourceData.Count - 1 - commandParamPosition);
                }
            }
            else if (_prmType == DataTypes.Error)
            {
                double l = 0;
                if (commandParamPosition + CommandParamSize[parameter] <= SourceData.Count - 1)
                {
                    _raw = SourceData.GetRange(commandParamPosition, CommandParamSize[parameter]);
                    l = RawToError(_raw.ToArray());
                    _val = l.ToString();
                    if (l != 0 && CommandFrameLength == 3 && parameter == 0 &&
                        commandParamPosition + CommandParamSize[parameter] == SourceData.Count - 1)
                    {
                        if (CommandParamDbLineNum.Count > 1)
                            CommandParamDbLineNum.RemoveRange(1, CommandParamDbLineNum.Count - parameter - 1);
                        if (CommandParamSize.Count > 1)
                            CommandParamSize.RemoveRange(1, CommandParamSize.Count - parameter - 1);
                        if (CommandParamSizeDefined.Count > 1)
                            CommandParamSizeDefined.RemoveRange(1, CommandParamSizeDefined.Count - parameter - 1);
                        if (CommandParamDesc.Count > 1)
                            CommandParamDesc.RemoveRange(1, CommandParamDesc.Count - parameter - 1);
                        if (CommandParamType.Count > 1)
                            CommandParamType.RemoveRange(1, CommandParamType.Count - parameter - 1);
                    }
                }
                else
                {
                    errFlag = true;
                    errMessage = "!!!ERR: Out of data bound!!!";
                    if (commandParamPosition <= SourceData.Count - 1)
                        _raw = SourceData.GetRange(commandParamPosition, SourceData.Count - 1 - commandParamPosition);
                }
            }
            else if (_prmType == DataTypes.Data)
            {
                if (commandParamPosition + CommandParamSize[parameter] <= SourceData.Count - 1)
                {
                    _raw = SourceData.GetRange(commandParamPosition, CommandParamSize[parameter]);
                    _val = RawToData(_raw.ToArray());
                }
                else
                {
                    errFlag = true;
                    errMessage = "!!!ERR: Out of data bound!!!";
                    if (commandParamPosition <= SourceData.Count - 1)
                        _raw = SourceData.GetRange(commandParamPosition, SourceData.Count - 1 - commandParamPosition);
                }
            }
            else if (_prmType == DataTypes.PrefData)
            {
                var prefLength = 0;
                //get gata length
                if (commandParamPosition + 2 <= SourceData.Count - 1)
                    prefLength = (int) RawToNumber(SourceData.GetRange(commandParamPosition, 2).ToArray());
                //check if the size is correct
                if (prefLength + 2 > CommandParamSize[parameter])
                {
                    prefLength = CommandParamSize[parameter] - 2;
                    errFlag = true;
                    errMessage = "!!!ERR: Out of data bound!!!";
                }
                else
                {
                    CommandParamSize[parameter] = prefLength + 2;
                }

                //get data
                if (commandParamPosition + CommandParamSize[parameter] <= SourceData.Count - 1)
                {
                    _raw = SourceData.GetRange(commandParamPosition, CommandParamSize[parameter]);
                    //_val = "[" + prefLength.ToString() + "]" + Accessory.ConvertHexToString(_raw.Substring(6), CustomFiscalControl.Properties.Settings.Default.CodePage);
                    _val = RawToPrefData(_raw.ToArray(), CommandParamSize[parameter]);
                }
                else
                {
                    errFlag = true;
                    errMessage = "!!!ERR: Out of data bound!!!";
                    if (commandParamPosition <= SourceData.Count - 1)
                        _raw = SourceData.GetRange(commandParamPosition, SourceData.Count - 1 - commandParamPosition);
                }
            }
            else if (_prmType == DataTypes.TLVData)
            {
                var TlvType = 0;
                var TlvLength = 0;
                if (CommandParamSize[parameter] > 0)
                {
                    if (commandParamPosition + 4 <= SourceData.Count - 1)
                    {
                        //get type of parameter
                        TlvType = (int) RawToNumber(SourceData.GetRange(commandParamPosition, 2).ToArray());
                        //get gata length
                        TlvLength = (int) RawToNumber(SourceData.GetRange(commandParamPosition + 2, 2).ToArray());
                    }

                    //check if the size is correct
                    if (TlvLength + 4 > CommandParamSize[parameter])
                    {
                        TlvLength = CommandParamSize[parameter] - 4;
                        errFlag = true;
                        errMessage = "!!!ERR: Out of data bound!!!";
                    }
                    else
                    {
                        CommandParamSize[parameter] = TlvLength + 4;
                    }

                    //get data
                    if (commandParamPosition + CommandParamSize[parameter] <= SourceData.Count - 1)
                    {
                        _raw = SourceData.GetRange(commandParamPosition, CommandParamSize[parameter]);
                        //_val = "[" + TlvType.ToString() + "]" + "[" + TlvLength.ToString() + "]" + Accessory.ConvertHexToString(_raw.Substring(12), CustomFiscalControl.Properties.Settings.Default.CodePage);
                        _val = RawToTLVData(_raw.ToArray(), CommandParamSize[parameter]);
                    }
                    else
                    {
                        errFlag = true;
                        errMessage = "!!!ERR: Out of data bound!!!";
                        if (commandParamPosition <= SourceData.Count - 1)
                            _raw = SourceData.GetRange(commandParamPosition,
                                SourceData.Count - 1 - commandParamPosition);
                    }
                }
            }
            else if (_prmType == DataTypes.Bitfield)
            {
                double l = 0;
                if (commandParamPosition + CommandParamSize[parameter] <= SourceData.Count - 1)
                {
                    _raw = SourceData.GetRange(commandParamPosition, CommandParamSize[parameter]);
                    l = RawToBitfield(_raw[0]);
                    _val = l.ToString();
                }
                else
                {
                    errFlag = true;
                    errMessage = "!!!ERR: Out of data bound!!!";
                    if (commandParamPosition <= SourceData.Count - 1)
                        _raw = SourceData.GetRange(commandParamPosition, SourceData.Count - 1 - commandParamPosition);
                }
            }
            else
            {
                //flag = true;
                errFlag = true;
                errMessage = "!!!ERR: Incorrect parameter type!!!";
                if (commandParamPosition + CommandParamSize[parameter] <= SourceData.Count - 1)
                {
                    _raw = SourceData.GetRange(commandParamPosition, CommandParamSize[parameter]);
                    //_val = Accessory.ConvertHexToString(_raw, CustomFiscalControl.Properties.Settings.Default.CodePage);
                }
                else
                {
                    errFlag = true;
                    errMessage = "!!!ERR: Out of data bound!!!";
                    if (commandParamPosition <= SourceData.Count - 1)
                        _raw = SourceData.GetRange(commandParamPosition, SourceData.Count - 1 - commandParamPosition);
                }
            }

            CommandParamRawValue.Add(_raw);
            CommandParamValue.Add(_val);

            var predefinedFound =
                false; //Matching predefined parameter found and it's number is in "predefinedParameterMatched"
            if (errFlag) CommandParamDesc[parameter] += errMessage + "\r\n";

            //compare parameter value with predefined values to get proper description
            var predefinedParameterMatched = 0;
            for (var i1 = 0; i1 < predefinedParamsVal.Count; i1++)
                if (CommandParamValue[parameter] == predefinedParamsVal[i1].ToString())
                {
                    predefinedFound = true;
                    predefinedParameterMatched = i1;
                }

            CommandParamDesc[parameter] += "\r\n";
            if (CommandParamDbLineNum[parameter] + predefinedParameterMatched + 1 <
                CommandDbLineNum + CommandDbHeight && predefinedFound)
                CommandParamDesc[parameter] +=
                    CommandDataBase.Rows[CommandParamDbLineNum[parameter] + predefinedParameterMatched + 1][
                        CSVColumns.CommandDescriptionColumn].ToString();
        }

        ResultLength();
        return true;
    }

    internal static void ClearCommand()
    {
        ItIsReply = false;
        ItIsReplyNack = false;
        CrcFailed = false;
        LengthIncorrect = false;
        CommandFramePosition = -1;
        CommandDbLineNum = -1;
        CommandDbHeight = -1;
        CommandName = "";
        CommandDesc = "";

        CommandParamSize.Clear();
        CommandParamDesc.Clear();
        CommandParamType.Clear();
        CommandParamValue.Clear();
        CommandParamRawValue.Clear();
        CommandParamDbLineNum.Clear();
        CommandBlockLength = 0;
    }

    internal static void ClearCommandParameters()
    {
        CommandParamSize.Clear();
        CommandParamDesc.Clear();
        CommandParamType.Clear();
        CommandParamValue.Clear();
        CommandParamRawValue.Clear();
        CommandParamDbLineNum.Clear();
        CommandParamSizeDefined.Clear();
        CommandBlockLength = 0;
    }

    internal static int ResultLength() //Calc "CommandBlockLength" - length of command text in source text field
    {
        //FrameStart[1] + DataLength[2] + data + CRC
        CommandBlockLength = 3 + CommandFrameLength + 1;
        if (ItIsReply) CommandBlockLength += 3;
        return CommandBlockLength;
    }

    public static byte Q3xf_CRC(byte[] data, int length)
    {
        ushort sum = 0;
        for (var i = 0; i < length; i++) sum += data[i];
        var sumH = (byte) (sum / 256);
        var sumL = (byte) (sum - sumH * 256);
        return (byte) (sumH ^ sumL);
    }

    public static string RawToString(byte[] b, int n)
    {
        var outStr = Encoding.GetEncoding(Settings.Default.CodePage).GetString(b);
        if (outStr.Length > n) outStr = outStr.Substring(0, n);
        return outStr;
    }

    public static string RawToPrefData(byte[] b, int n)
    {
        var s = new List<byte>();
        s.AddRange(b);
        if (s.Count < 2) return "";
        if (s.Count > n + 2) s = s.GetRange(0, n + 2);
        var outStr = "";
        var strLength = (int) RawToNumber(s.GetRange(0, 2).ToArray());
        outStr = "[" + strLength + "]";
        if (s.Count == 2 + strLength)
        {
            var b1 = s.GetRange(2, s.Count - 2).ToArray();
            if (Accessory.PrintableByteArray(b1))
                outStr += "\"" + Encoding.GetEncoding(Settings.Default.CodePage).GetString(b1) + "\"";
            else outStr += "[" + Accessory.ConvertByteArrayToHex(b1) + "]";
        }
        else
        {
            outStr += "INCORRECT LENGTH";
        }

        return outStr;
    }

    // !!! check TLV actual data layout
    public static string RawToTLVData(byte[] b, int n)
    {
        var s = new List<byte>();
        s.AddRange(b);
        if (s.Count < 4) return "";
        if (s.Count > n + 4) s = s.GetRange(0, n + 4);
        var outStr = "";
        var tlvType = (int) RawToNumber(s.GetRange(0, 2).ToArray());
        outStr = "[" + tlvType + "]";
        var strLength = (int) RawToNumber(s.GetRange(2, 2).ToArray());
        outStr += "[" + strLength + "]";
        if (s.Count == 4 + strLength)
        {
            var b1 = s.GetRange(2, s.Count - 2).ToArray();
            if (Accessory.PrintableByteArray(b1))
                outStr += "\"" + Encoding.GetEncoding(Settings.Default.CodePage)
                    .GetString(s.GetRange(4, s.Count - 4).ToArray()) + "\"";
            else outStr += "[" + Accessory.ConvertByteArrayToHex(b1) + "]";
        }
        else
        {
            outStr += "INCORRECT LENGTH";
        }

        return outStr;
    }

    public static double RawToPassword(byte[] b)
    {
        double l = 0;
        for (var n = 0; n < b.Length; n++) l += b[n] * Math.Pow(256, n);
        return l;
    }

    public static double RawToNumber(byte[] b)
    {
        double l = 0;
        for (var n = 0; n < b.Length; n++)
            if (n == b.Length - 1 && Accessory.GetBit(b[n], 7))
            {
                l += b[n] * Math.Pow(256, n);
                l = l - Math.Pow(2, b.Length * 8);
            }
            else
            {
                l += b[n] * Math.Pow(256, n);
            }

        return l;
    }

    public static double RawToMoney(byte[] b)
    {
        double l = 0;
        for (var n = 0; n < b.Length; n++)
            if (n == b.Length - 1 && Accessory.GetBit(b[n], 7))
            {
                l += b[n] * Math.Pow(256, n);
                l = l - Math.Pow(2, b.Length * 8);
            }
            else
            {
                l += b[n] * Math.Pow(256, n);
            }

        return l / 100;
    }

    public static double RawToQuantity(byte[] b)
    {
        double l = 0;
        for (var n = 0; n < b.Length; n++)
            if (n == b.Length - 1 && Accessory.GetBit(b[n], 7))
            {
                l += b[n] * Math.Pow(256, n);
                l = l - Math.Pow(2, b.Length * 8);
            }
            else
            {
                l += b[n] * Math.Pow(256, n);
            }

        return l / 1000;
    }

    public static double RawToError(byte[] b)
    {
        double l = 0;
        for (var n = 0; n < b.Length; n++) l += b[n] * Math.Pow(256, n);
        return l;
    }

    public static string RawToData(byte[] b)
    {
        if (Accessory.PrintableByteArray(b))
            return "\"" + Encoding.GetEncoding(Settings.Default.CodePage).GetString(b) + "\"";
        return "[" + Accessory.ConvertByteArrayToHex(b) + "]";
    }

    public static double RawToBitfield(byte b)
    {
        return b;
    }

    public static string StringToRaw(string s, int n)
    {
        //while (s.Length < n) s += "\0";
        //return Accessory.ConvertStringToHex(s, CustomFiscalControl.Properties.Settings.Default.CodePage).Substring(0, n * 3);
        var outStr = Accessory.ConvertStringToHex(s.Substring(1, s.Length - 2), Settings.Default.CodePage);
        if (outStr.Length > n * 3) outStr = outStr.Substring(0, n * 3);
        while (outStr.Length < n * 3) outStr += "00 ";
        return outStr;
    }

    // !!! incorrect layout
    public static string PrefDataToRaw(string s, int n)
    {
        if (s.Length > n - 2) s = s.Substring(0, n - 2);
        var outStr = NumberToRaw(s.Length.ToString(), 2);
        outStr += Accessory.ConvertStringToHex(s, Settings.Default.CodePage);
        return outStr;
    }

    // !!! incorrect layout
    public static string TLVDataToRaw(string s, int n)
    {
        if (!(s.Contains('[') && s.Contains(']'))) return "";
        if (s.Length < 3) return "";
        var outStr = "";
        var tlvType = -1;
        int.TryParse(s.Substring(0, s.IndexOf(']')).Replace("[", ""), out tlvType);
        s = s.Substring(s.IndexOf(']') + 1);

        if (n > s.Length) n = s.Length;
        outStr = ErrorToRaw(n.ToString(), 2);
        outStr += Accessory.ConvertStringToHex(s, Settings.Default.CodePage).Substring(0, n * 3);
        return outStr;
    }

    public static string PasswordToRaw(string s, int n)
    {
        s = s.Replace(" ", "").Replace("/", "");
        long l = 0;
        if (s != "") long.TryParse(s, out l);
        var str = "";
        for (var i = 0; i < n; i++) str += Accessory.ConvertByteToHex((byte) (l / Math.Pow(256, i)));
        return str;
    }

    public static string NumberToRaw(string s, int n)
    {
        double d = 0;
        if (s != "") double.TryParse(s, out d);
        if (d < 0) d = Math.Pow(2, n * 8) - Math.Abs(d);
        var b = new byte[n];
        for (var i = n - 1; i >= 0; i--)
        {
            b[i] += (byte) (d / Math.Pow(256, i));
            d -= b[i] * Math.Pow(256, i);
        }

        var str = "";
        for (var i = 0; i < n; i++) str += Accessory.ConvertByteToHex(b[i]);
        return str;
    }

    public static string MoneyToRaw(string s, int n)
    {
        double d = 0;
        if (s != "") double.TryParse(s, out d);
        d *= 100;
        if (d < 0) d = Math.Pow(2, n * 8) - Math.Abs(d);
        var b = new byte[n];
        for (var i = n - 1; i >= 0; i--)
        {
            b[i] += (byte) (d / Math.Pow(256, i));
            d -= b[i] * Math.Pow(256, i);
        }

        var str = "";
        for (var i = 0; i < n; i++) str += Accessory.ConvertByteToHex(b[i]);
        return str;
    }

    public static string QuantityToRaw(string s, int n)
    {
        double d = 0;
        if (s != "") double.TryParse(s, out d);
        d *= 1000;
        if (d < 0) d = Math.Pow(2, n * 8) - Math.Abs(d);
        var b = new byte[n];
        for (var i = n - 1; i >= 0; i--)
        {
            b[i] += (byte) (d / Math.Pow(256, i));
            d -= b[i] * Math.Pow(256, i);
        }

        var str = "";
        for (var i = 0; i < n; i++) str += Accessory.ConvertByteToHex(b[i]);
        return str;
    }

    public static string ErrorToRaw(string s, int n)
    {
        long l = 0;
        if (s != "") long.TryParse(s, out l);
        var str = "";
        for (var i = 0; i < n; i++) str += Accessory.ConvertByteToHex((byte) (l / Math.Pow(256, i)));
        return str;
    }

    public static string DataToRaw(string s, int n)
    {
        var outStr = "";
        if (s.Substring(0, 1) == "[") outStr = s.Substring(1, s.Length - 2);
        else if (s.Substring(0, 1) == "\"")
            outStr = Accessory.ConvertStringToHex(s.Substring(1, s.Length - 2), Settings.Default.CodePage);
        else return "";
        if (outStr.Length > n * 3) outStr = outStr.Substring(0, n * 3);
        while (outStr.Length < n * 3) outStr += "00 ";
        return outStr;
    }

    public static string BitfieldToRaw(string s)
    {
        byte l = 0;
        if (s != "") byte.TryParse(s, out l);
        var str = "";
        str += Accessory.ConvertByteToHex(l);
        return str;
    }
}