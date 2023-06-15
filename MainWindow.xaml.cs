using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using WinForms = System.Windows.Forms;
using System.IO;
using OfficeOpenXml;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Windows.Forms;
using MessageBox = System.Windows.Forms.MessageBox;
using System.Threading;
using System.Windows.Markup;

namespace HeaderGenerator
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        string folder;
        int plane;
        private System.Action myUpdate = () => { };
        private System.Action[] actions;
        private SqlConnection Connection = null;
        private SqlDataAdapter DataAdapter = null;
        private DataTable table = null;
        string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
        public void LRbits(string bits, ref int Rbit, ref int Lbit)
        {
            int bit = 0;
            int i, j = 0;
            bool flug = false;
            int len = bits.Length;
            for (i = 0; i < len; i++)
            {
                if ((bits[i] != '.' && bits[i] != '…') && i < len)
                {
                    bit = (int)(Math.Pow(10, j) * bit + (int)Char.GetNumericValue(bits[i]));
                    j++;
                }
                else if (bits[i] == '.' && i < len)
                {
                    if (!flug)
                    {
                        flug = true;
                        Lbit = bit;
                        bit = 0;
                        j = 0;

                    }
                }
                else if (bits[i] == '…' && i < len)
                {
                    if (!flug)
                    {
                        flug = true;
                        Lbit = bit;
                        bit = 0;
                        j = 0;

                    }
                }

                if (i == len - 1)
                {
                    if (Lbit == 0)
                    {
                        Lbit = bit;
                        Rbit = 0;
                    }
                    else
                    {
                        Rbit = bit;
                        bit = 0;
                        j = 0;
                    }
                }
            }
        }
        private void С70parsing()
        {

            string[,] excelTable;
            int totalRows = 0;
            int totalColums = 0;


            string[] dirs = Directory.GetDirectories(folder, "*", SearchOption.TopDirectoryOnly);
            foreach (string dir in dirs)
            {
                string[] files = Directory.GetFiles(dir);
                foreach (string filename in files)
                {
                    string newfilename = filename;
                    newfilename = newfilename.Remove(0, dir.Length + 1);
                    if (newfilename.StartsWith("sig_") && newfilename.EndsWith("_mkio.xlsx"))
                    {


                        LogText.Text += "МКИО  :: Извлекаю данные из " + newfilename;

                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        ExcelPackage excelFile = new ExcelPackage(filename);
                        for (int l = 0; l < 2; l++)
                        {
                            ExcelWorksheet worksheet = excelFile.Workbook.Worksheets[l];
                            totalRows = worksheet.Dimension.End.Row;
                            totalColums = worksheet.Dimension.End.Column;

                            excelTable = new string[totalRows, totalColums];

                            for (int rowIndex = 1; rowIndex <= totalRows; rowIndex++)
                            {
                                IEnumerable<string> row = worksheet.Cells[rowIndex, 1, rowIndex, totalColums].Select(c => c.Value == null ? "" : c.Value.ToString());

                                List<string> list = row.ToList<string>();

                                for (int i = 0; i < list.Count; i++)
                                {
                                    excelTable[rowIndex - 1, i] = Convert.ToString(list[i]);
                                   
                                }
                            }
                            string temp1 = "", temp2 = "";

                            string MessageName = "";
                            string LineDB = "";
                            string SubLine = "";
                            string WordNumbers = "";
                            string SubAdress = "";
                            string NumWord = "";
                            string Coment = "";
                            string NameParam = "";
                            string Bits = "";
                            int Rbit = 0;
                            int Lbit = 0;
                            string PriemnikRUS = "";
                            string PriemnikENG = "";
                            string sqlExpression = "";
                            string sqlExpression2 = "";

                            using (SqlConnection connection = new SqlConnection(connectionString))

                            {
                                connection.Open();

                                MessageName = excelTable[4, 2].ToString() + "_TO_" + excelTable[4, 3].ToString();
                                LineDB = "МКИО".ToString();
                                SubLine = excelTable[9, 3].ToString();
                                WordNumbers = excelTable[7, 2].ToString();
                                if (l == 0) SubAdress = excelTable[5, 3].ToString();
                                else if (l == 1) SubAdress = excelTable[5, 2].ToString();

                                PriemnikRUS = excelTable[3, 2].ToString();
                                PriemnikENG = excelTable[4, 3].ToString();

                                for (int i = 0; i < totalRows - 1; i++)
                                {
                                    for (int j = 0; j < totalColums; j++)
                                    {
                                        if (i > 10 && j == 0 && (excelTable[i + 1, j] == "" || excelTable[i + 1, j] == " ") && (excelTable[i, j] != "" || excelTable[i, j] == " "))
                                        {
                                            temp1 = excelTable[i, 0];
                                            temp2 = excelTable[i, 12];
                                        }
                                        if (i > 10 && j == 0 && (excelTable[i, j] == "" || excelTable[i, j] == " ") && excelTable[i, 2] != "")
                                        {
                                            excelTable[i, 0] = temp1;
                                            excelTable[i, 12] = temp2;
                                        }
                                        //Console.Write(excelTable[i, j] + " ");
                                    }
                                    if (i > 10 && excelTable[i, 9].ToString() != null && excelTable[i, 9].ToString() != " " && excelTable[i, 9].ToString() != "")
                                    {

                                        NumWord = excelTable[i, 0].ToString();
                                        Coment = excelTable[i, 1].ToString();
                                        NameParam = excelTable[i, 2].ToString();
                                        Bits = excelTable[i, 9].ToString();

                                        LRbits(Bits, ref Rbit, ref Lbit);
                                        sqlExpression = "INSERT INTO Words (MessageName, Line, SubLine, WordNumbers, SubAdress, NumWord,Coment, NameParam, LBit, RBit, PriemnikRUS, PriemnikENG) VALUES ('" + MessageName + "', '" + LineDB + "', '" + SubLine;
                                        sqlExpression2 = sqlExpression + "', " + WordNumbers + ", " + SubAdress + ", " + NumWord + ", '" + Coment + "', '" + NameParam + "', " + Lbit + ", " + Rbit + ", '" + PriemnikRUS + "', '" + PriemnikENG + "')";

                                        Connection = new SqlConnection(connectionString);
                                        Connection.Open();
                                        SqlCommand createCommand = new SqlCommand(sqlExpression2, Connection);
                                        createCommand.ExecuteNonQuery();

                                        Connection.Close();

                                        Console.Write(sqlExpression2);
                                        Console.Write("\n");
                                    }
                                }
                            }
                        }

                        Console.WriteLine("Подключение закрыто...");
                        LogText.Text += "    Готово" + "\n";
                    }
                    else if (newfilename.StartsWith("sig_") && newfilename.EndsWith("_dpk.xlsx"))
                    {
                        LogText.Text += "ДПК      :: Извлекаю данные из " + newfilename + "\n";
                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        ExcelPackage excelFile = new ExcelPackage(filename);
                        ExcelWorksheet worksheet = excelFile.Workbook.Worksheets[0];
                        totalRows = worksheet.Dimension.End.Row;
                        totalColums = worksheet.Dimension.End.Column;

                        excelTable = new string[totalRows, totalColums];

                        for (int rowIndex = 1; rowIndex <= totalRows; rowIndex++)
                        {
                            IEnumerable<string> row = worksheet.Cells[rowIndex, 1, rowIndex, totalColums].Select(c => c.Value == null ? "" : c.Value.ToString());

                            List<string> list = row.ToList<string>();

                            for (int i = 0; i < list.Count; i++)
                            {
                                excelTable[rowIndex - 1, i] = Convert.ToString(list[i]);
                               
                            }
                        }
                        string temp1 = "", temp2 = "";

                        string MessageName = "";
                        string LineDB = "";
                        string SubLine = "";
                        string WordNumbers = "";
                        string SubAdress = "";
                        string NumWord = "";
                        string Coment = "";
                        string NameParam = "";
                        string Bits = "";
                        int Rbit = 0;
                        int Lbit = 0;
                        string PriemnikRUS = "";
                        string PriemnikENG = "";
                        string sqlExpression = "";
                        string sqlExpression2 = "";

                        using (SqlConnection connection = new SqlConnection(connectionString))

                        {
                            connection.Open();

                            MessageName = excelTable[4, 1].ToString() + "_TO_" + excelTable[4, 2].ToString();
                            LineDB = "ДПК".ToString();
                            SubLine = excelTable[6, 2].ToString();
                            WordNumbers = "0";
                            SubAdress = "0";
                            PriemnikRUS = excelTable[3, 2].ToString();
                            PriemnikENG = excelTable[4, 2].ToString();

                            for (int i = 0; i < totalRows - 1; i++)
                            {
                                for (int j = 0; j < totalColums; j++)
                                {
                                    if (i > 8 && j == 0 && (excelTable[i + 1, j] == "" || excelTable[i + 1, j] == " ") && (excelTable[i, j] != "" || excelTable[i, j] == " "))
                                    {
                                        temp1 = excelTable[i, 0];
                                        temp2 = excelTable[i, 9];
                                    }
                                    if (i > 8 && j == 0 && (excelTable[i, j] == "" || excelTable[i, j] == " ") && excelTable[i, 2] != "")
                                    {
                                        excelTable[i, 0] = temp1;
                                        excelTable[i, 9] = temp2;
                                    }
                                    //Console.Write(excelTable[i, j] + " ");
                                }
                                if (i > 8 && excelTable[i, 8].ToString() != null && excelTable[i, 8].ToString() != " " && excelTable[i, 8].ToString() != "")
                                {

                                    NumWord = excelTable[i, 0].ToString();
                                    Coment = excelTable[i, 1].ToString();
                                    NameParam = excelTable[i, 2].ToString();
                                    Bits = excelTable[i, 8].ToString();

                                    LRbits(Bits, ref Rbit, ref Lbit);
                                    sqlExpression = "INSERT INTO Words (MessageName, Line, SubLine, WordNumbers, SubAdress, NumWord,Coment, NameParam, LBit, RBit, PriemnikRUS, PriemnikENG) VALUES ('" + MessageName + "', '" + LineDB + "', '" + SubLine;
                                    sqlExpression2 = sqlExpression + "', " + WordNumbers + ", " + SubAdress + ", " + NumWord + ", '" + Coment + "', '" + NameParam + "', " + Lbit + ", " + Rbit + ", '" + PriemnikRUS + "', '" + PriemnikENG + "')";

                                    Connection = new SqlConnection(connectionString);
                                    Connection.Open();
                                    SqlCommand createCommand = new SqlCommand(sqlExpression2, Connection);
                                    createCommand.ExecuteNonQuery();

                                    Connection.Close();

                                    Console.Write(sqlExpression2);
                                    Console.Write("\n");
                                }
                            }
                        }
                        Console.WriteLine("Подключение закрыто...");
                        LogText.Text += " Готово" + "\n";
                    }
                }
            }
        }

        private void Cy57parsing()
        {

            string[,] excelTable;
            int totalRows = 0;
            int totalColums = 0;

            string[] files = Directory.GetFiles(folder);
            foreach (string filename in files)
            {
                string newfilename = filename;
                newfilename = newfilename.Remove(0, folder.Length + 1);
                if (newfilename.StartsWith("sig_") && newfilename.EndsWith("_мкио.xlsx"))
                {
                    LogText.Text += "МКИО  :: Извлекаю данные из " + newfilename;
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    ExcelPackage excelFile = new ExcelPackage(filename);
                    for (int l = 0; l < 2; l++)
                    {
                        ExcelWorksheet worksheet = excelFile.Workbook.Worksheets[l];
                        totalRows = worksheet.Dimension.End.Row;
                        totalColums = worksheet.Dimension.End.Column;

                        excelTable = new string[totalRows, totalColums];

                        for (int rowIndex = 1; rowIndex <= totalRows; rowIndex++)
                        {
                            IEnumerable<string> row = worksheet.Cells[rowIndex, 1, rowIndex, totalColums].Select(c => c.Value == null ? "" : c.Value.ToString());

                            List<string> list = row.ToList<string>();

                            for (int i = 0; i < list.Count; i++)
                            {
                                excelTable[rowIndex - 1, i] = Convert.ToString(list[i]);
                                
                            }
                        }
                        string temp1 = "", temp2 = "";

                        string MessageName = "";
                        string LineDB = "";
                        string SubLine = "";
                        string WordNumbers = "";
                        string SubAdress = "";
                        string NumWord = "";
                        string Coment = "";
                        string NameParam = "";
                        string Bits = "";
                        int Rbit = 0;
                        int Lbit = 0;
                        string PriemnikRUS = "";
                        string PriemnikENG = "";
                        string sqlExpression = "";
                        string sqlExpression2 = "";

                        using (SqlConnection connection = new SqlConnection(connectionString))

                        {
                            connection.Open();

                            MessageName = excelTable[4, 2].ToString() + "_TO_" + excelTable[4, 3].ToString();
                            LineDB = "МКИО".ToString();
                            SubLine = excelTable[9, 3].ToString();
                            WordNumbers = excelTable[7, 2].ToString();
                            if (l == 0) SubAdress = excelTable[5, 3].ToString();
                            else if (l == 1) SubAdress = excelTable[5, 2].ToString();

                            PriemnikRUS = excelTable[3, 2].ToString();
                            PriemnikENG = excelTable[4, 3].ToString();

                            for (int i = 0; i < totalRows - 1; i++)
                            {
                                for (int j = 0; j < totalColums; j++)
                                {
                                    if (i > 10 && j == 0 && (excelTable[i + 1, j] == "" || excelTable[i + 1, j] == " ") && (excelTable[i, j] != "" || excelTable[i, j] == " "))
                                    {
                                        temp1 = excelTable[i, 0];
                                        temp2 = excelTable[i, 12];
                                    }
                                    if (i > 10 && j == 0 && (excelTable[i, j] == "" || excelTable[i, j] == " ") && excelTable[i, 2] != "")
                                    {
                                        excelTable[i, 0] = temp1;
                                        excelTable[i, 12] = temp2;
                                    }
                                    //Console.Write(excelTable[i, j] + " ");
                                }
                                if (i > 10 && excelTable[i, 9].ToString() != null && excelTable[i, 9].ToString() != " " && excelTable[i, 9].ToString() != "")
                                {

                                    NumWord = excelTable[i, 0].ToString();
                                    Coment = excelTable[i, 1].ToString();
                                    NameParam = excelTable[i, 2].ToString();
                                    Bits = excelTable[i, 9].ToString();

                                    LRbits(Bits, ref Rbit, ref Lbit);
                                    sqlExpression = "INSERT INTO Words (MessageName, Line, SubLine, WordNumbers, SubAdress, NumWord,Coment, NameParam, LBit, RBit, PriemnikRUS, PriemnikENG) VALUES ('" + MessageName + "', '" + LineDB + "', '" + SubLine;
                                    sqlExpression2 = sqlExpression + "', " + WordNumbers + ", " + SubAdress + ", " + NumWord + ", '" + Coment + "', '" + NameParam + "', " + Lbit + ", " + Rbit + ", '" + PriemnikRUS + "', '" + PriemnikENG + "')";

                                    Connection = new SqlConnection(connectionString);
                                    Connection.Open();
                                    SqlCommand createCommand = new SqlCommand(sqlExpression2, Connection);
                                    createCommand.ExecuteNonQuery();

                                    Connection.Close();

                                    Console.Write(sqlExpression2);
                                    Console.Write("\n");
                                }
                            }
                        }
                    }

                    Console.WriteLine("Подключение закрыто...");
                    LogText.Text += "    Готово" + "\n";
                }
                else if (newfilename.StartsWith("sig_") && newfilename.EndsWith("_дпк.xlsx"))
                {
                    LogText.Text += "ДПК      :: Извлекаю данные из " + newfilename + "\n";
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    ExcelPackage excelFile = new ExcelPackage(filename);
                    ExcelWorksheet worksheet = excelFile.Workbook.Worksheets[0];
                    totalRows = worksheet.Dimension.End.Row;
                    totalColums = worksheet.Dimension.End.Column;

                    excelTable = new string[totalRows, totalColums];

                    for (int rowIndex = 1; rowIndex <= totalRows; rowIndex++)
                    {
                        IEnumerable<string> row = worksheet.Cells[rowIndex, 1, rowIndex, totalColums].Select(c => c.Value == null ? "" : c.Value.ToString());

                        List<string> list = row.ToList<string>();

                        for (int i = 0; i < list.Count; i++)
                        {
                            excelTable[rowIndex - 1, i] = Convert.ToString(list[i]);
  
                        }
                    }
                    string temp1 = "", temp2 = "";

                    string MessageName = "";
                    string LineDB = "";
                    string SubLine = "";
                    string WordNumbers = "";
                    string SubAdress = "";
                    string NumWord = "";
                    string Coment = "";
                    string NameParam = "";
                    string Bits = "";
                    int Rbit = 0;
                    int Lbit = 0;
                    string PriemnikRUS = "";
                    string PriemnikENG = "";
                    string sqlExpression = "";
                    string sqlExpression2 = "";

                    using (SqlConnection connection = new SqlConnection(connectionString))

                    {
                        connection.Open();

                        MessageName = excelTable[4, 1].ToString() + "_TO_" + excelTable[4, 2].ToString();
                        LineDB = "ДПК".ToString();
                        SubLine = excelTable[6, 2].ToString();
                        WordNumbers = "0";
                        SubAdress = "0";
                        PriemnikRUS = excelTable[3, 2].ToString();
                        PriemnikENG = excelTable[4, 2].ToString();

                        for (int i = 0; i < totalRows - 1; i++)
                        {
                            for (int j = 0; j < totalColums; j++)
                            {
                                if (i > 8 && j == 0 && (excelTable[i + 1, j] == "" || excelTable[i + 1, j] == " ") && (excelTable[i, j] != "" || excelTable[i, j] == " "))
                                {
                                    temp1 = excelTable[i, 0];
                                    temp2 = excelTable[i, 9];
                                }
                                if (i > 8 && j == 0 && (excelTable[i, j] == "" || excelTable[i, j] == " ") && excelTable[i, 2] != "")
                                {
                                    excelTable[i, 0] = temp1;
                                    excelTable[i, 9] = temp2;
                                }
                                //Console.Write(excelTable[i, j] + " ");
                            }
                            if (i > 8 && excelTable[i, 8].ToString() != null && excelTable[i, 8].ToString() != " " && excelTable[i, 8].ToString() != "")
                            {

                                NumWord = excelTable[i, 0].ToString();
                                Coment = excelTable[i, 1].ToString();
                                NameParam = excelTable[i, 2].ToString();
                                Bits = excelTable[i, 8].ToString();

                                LRbits(Bits, ref Rbit, ref Lbit);
                                sqlExpression = "INSERT INTO Words (MessageName, Line, SubLine, WordNumbers, SubAdress, NumWord,Coment, NameParam, LBit, RBit, PriemnikRUS, PriemnikENG) VALUES ('" + MessageName + "', '" + LineDB + "', '" + SubLine;
                                sqlExpression2 = sqlExpression + "', " + WordNumbers + ", " + SubAdress + ", " + NumWord + ", '" + Coment + "', '" + NameParam + "', " + Lbit + ", " + Rbit + ", '" + PriemnikRUS + "', '" + PriemnikENG + "')";

                                Connection = new SqlConnection(connectionString);
                                Connection.Open();
                                SqlCommand createCommand = new SqlCommand(sqlExpression2, Connection);
                                createCommand.ExecuteNonQuery();

                                Connection.Close();

                                Console.Write(sqlExpression2);
                                Console.Write("\n");
                            }
                        }
                    }
                    Console.WriteLine("Подключение закрыто...");
                    LogText.Text += " Готово" + "\n";

                }
                
            }
        }

        private void T50parsing()
        {

            string[,] excelTable;
            int totalRows = 0;
            int totalColums = 0;


            string[] dirs = Directory.GetDirectories(folder, "*", SearchOption.TopDirectoryOnly);
            foreach (string dir in dirs)
            {
                string[] files = Directory.GetFiles(dir);
                foreach (string filename in files)
                {
                    string newfilename = filename;
                    newfilename = newfilename.Remove(0, dir.Length + 1);
                    if (newfilename.StartsWith("sig_") && newfilename.EndsWith("_mkio.xlsx"))
                    {


                        LogText.Text += "МКИО  :: Извлекаю данные из " + newfilename;

                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        ExcelPackage excelFile = new ExcelPackage(filename);
                        for (int l = 0; l < 2; l++)
                        {
                            ExcelWorksheet worksheet = excelFile.Workbook.Worksheets[l];
                            totalRows = worksheet.Dimension.End.Row;
                            totalColums = worksheet.Dimension.End.Column;

                            excelTable = new string[totalRows, totalColums];

                            for (int rowIndex = 1; rowIndex <= totalRows; rowIndex++)
                            {
                                IEnumerable<string> row = worksheet.Cells[rowIndex, 1, rowIndex, totalColums].Select(c => c.Value == null ? "" : c.Value.ToString());

                                List<string> list = row.ToList<string>();

                                for (int i = 0; i < list.Count; i++)
                                {
                                    excelTable[rowIndex - 1, i] = Convert.ToString(list[i]);

                                }
                            }
                            string temp1 = "", temp2 = "";

                            string MessageName = "";
                            string LineDB = "";
                            string SubLine = "";
                            string WordNumbers = "";
                            string SubAdress = "";
                            string NumWord = "";
                            string Coment = "";
                            string NameParam = "";
                            string Bits = "";
                            int Rbit = 0;
                            int Lbit = 0;
                            string PriemnikRUS = "";
                            string PriemnikENG = "";
                            string sqlExpression = "";
                            string sqlExpression2 = "";

                            using (SqlConnection connection = new SqlConnection(connectionString))

                            {
                                connection.Open();

                                MessageName = excelTable[2, 2].ToString() + "_TO_" + excelTable[4, 2].ToString();
                                LineDB = "МКИО".ToString();
                                SubLine = excelTable[5, 2].ToString();
                                WordNumbers = "0";
                                SubAdress = "0";
                                PriemnikRUS = excelTable[4, 1].ToString();
                                PriemnikENG = excelTable[4, 2].ToString();

                                for (int i = 0; i < totalRows - 1; i++)
                                {
                                    for (int j = 0; j < totalColums; j++)
                                    {
                                        if (i > 10 && j == 0 && (excelTable[i + 1, j] == "" || excelTable[i + 1, j] == " ") && (excelTable[i, j] != "" || excelTable[i, j] == " "))
                                        {
                                            temp1 = excelTable[i, 0];
                                            temp2 = excelTable[i, 12];
                                        }
                                        if (i > 10 && j == 0 && (excelTable[i, j] == "" || excelTable[i, j] == " ") && excelTable[i, 2] != "")
                                        {
                                            excelTable[i, 0] = temp1;
                                            excelTable[i, 12] = temp2;
                                        }
                                        Console.Write(excelTable[i, j] + " ");
                                    }
                                    if (i > 10 && excelTable[i, 9].ToString() != null && excelTable[i, 9].ToString() != " " && excelTable[i, 9].ToString() != "")
                                    {

                                        NumWord = excelTable[i, 0].ToString();
                                        Coment = excelTable[i, 1].ToString();
                                        NameParam = excelTable[i, 2].ToString();
                                        Bits = excelTable[i, 9].ToString();

                                        LRbits(Bits, ref Rbit, ref Lbit);
                                        sqlExpression = "INSERT INTO Words (MessageName, Line, SubLine, WordNumbers, SubAdress, NumWord,Coment, NameParam, LBit, RBit, PriemnikRUS, PriemnikENG) VALUES ('" + MessageName + "', '" + LineDB + "', '" + SubLine;
                                        sqlExpression2 = sqlExpression + "', " + WordNumbers + ", " + SubAdress + ", " + NumWord + ", '" + Coment + "', '" + NameParam + "', " + Lbit + ", " + Rbit + ", '" + PriemnikRUS + "', '" + PriemnikENG + "')";

                                        Connection = new SqlConnection(connectionString);
                                        Connection.Open();
                                        SqlCommand createCommand = new SqlCommand(sqlExpression2, Connection);
                                        createCommand.ExecuteNonQuery();

                                        Connection.Close();

                                        Console.Write(sqlExpression2);
                                        Console.Write("\n");
                                    }
                                }
                            }
                        }

                        Console.WriteLine("Подключение закрыто...");
                        LogText.Text += "    Готово" + "\n";
                    }
                    else if (newfilename.StartsWith("sig_") && newfilename.EndsWith("_dpk.xls"))
                    {
                        LogText.Text += "ДПК      :: Извлекаю данные из " + newfilename + "\n";
                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        ExcelPackage excelFile = new ExcelPackage(filename);
                        ExcelWorksheet worksheet = excelFile.Workbook.Worksheets[0];
                        totalRows = worksheet.Dimension.End.Row;
                        totalColums = worksheet.Dimension.End.Column;

                        excelTable = new string[totalRows, totalColums];

                        for (int rowIndex = 1; rowIndex <= totalRows; rowIndex++)
                        {
                            IEnumerable<string> row = worksheet.Cells[rowIndex, 1, rowIndex, totalColums].Select(c => c.Value == null ? "" : c.Value.ToString());

                            List<string> list = row.ToList<string>();

                            for (int i = 0; i < list.Count; i++)
                            {
                                excelTable[rowIndex - 1, i] = Convert.ToString(list[i]);

                            }
                        }
                        string temp1 = "", temp2 = "";

                        string MessageName = "";
                        string LineDB = "";
                        string SubLine = "";
                        string WordNumbers = "";
                        string SubAdress = "";
                        string NumWord = "";
                        string Coment = "";
                        string NameParam = "";
                        string Bits = "";
                        int Rbit = 0;
                        int Lbit = 0;
                        string PriemnikRUS = "";
                        string PriemnikENG = "";
                        string sqlExpression = "";
                        string sqlExpression2 = "";

                        using (SqlConnection connection = new SqlConnection(connectionString))

                        {
                            connection.Open();

                            MessageName = excelTable[2, 3].ToString() + "_TO_" + excelTable[5, 3].ToString();
                            LineDB = "ДПК".ToString();
                            SubLine = excelTable[6, 3].ToString();
                            WordNumbers = "0";
                            SubAdress = "0";
                            PriemnikRUS = excelTable[5, 2].ToString();
                            PriemnikENG = excelTable[5, 3].ToString();

                            for (int i = 0; i < totalRows - 1; i++)
                            {
                                for (int j = 0; j < totalColums; j++)
                                {
                                    if (i > 8 && j == 0 && (excelTable[i + 1, j] == "" || excelTable[i + 1, j] == " ") && (excelTable[i, j] != "" || excelTable[i, j] == " "))
                                    {
                                        temp1 = excelTable[i, 0];
                                        temp2 = excelTable[i, 13];
                                    }
                                    if (i > 8 && j == 0 && (excelTable[i, j] == "" || excelTable[i, j] == " ") && excelTable[i, 2] != "")
                                    {
                                        excelTable[i, 0] = temp1;
                                        excelTable[i, 13] = temp2;
                                    }
                                    //Console.Write(excelTable[i, j] + " ");
                                }
                                if (i > 8 && excelTable[i, 8].ToString() != null && excelTable[i, 8].ToString() != " " && excelTable[i, 8].ToString() != "")
                                {

                                    NumWord = excelTable[i, 0].ToString();
                                    Coment = excelTable[i, 1].ToString();
                                    NameParam = excelTable[i, 2].ToString();
                                    Bits = excelTable[i, 9].ToString();

                                    LRbits(Bits, ref Rbit, ref Lbit);
                                    sqlExpression = "INSERT INTO Words (MessageName, Line, SubLine, WordNumbers, SubAdress, NumWord,Coment, NameParam, LBit, RBit, PriemnikRUS, PriemnikENG) VALUES ('" + MessageName + "', '" + LineDB + "', '" + SubLine;
                                    sqlExpression2 = sqlExpression + "', " + WordNumbers + ", " + SubAdress + ", " + NumWord + ", '" + Coment + "', '" + NameParam + "', " + Lbit + ", " + Rbit + ", '" + PriemnikRUS + "', '" + PriemnikENG + "')";

                                    Connection = new SqlConnection(connectionString);
                                    Connection.Open();
                                    SqlCommand createCommand = new SqlCommand(sqlExpression2, Connection);
                                    createCommand.ExecuteNonQuery();

                                    Connection.Close();

                                    Console.Write(sqlExpression2);
                                    Console.Write("\n");
                                }
                            }
                        }
                        Console.WriteLine("Подключение закрыто...");
                        LogText.Text += " Готово" + "\n";
                    }
                }
            }
        }

        private void Cy35parsing()
        {

            string[,] excelTable;
            int totalRows = 0;
            int totalColums = 0;


            string[] dirs = Directory.GetDirectories(folder, "*", SearchOption.TopDirectoryOnly);
            foreach (string dir in dirs)
            {
                string[] files = Directory.GetFiles(dir);
                foreach (string filename in files)
                {
                    string newfilename = filename;
                    newfilename = newfilename.Remove(0, dir.Length + 1);
                    if (newfilename.StartsWith("sig_") && newfilename.EndsWith("_mkio.xlsx"))
                    {


                        LogText.Text += "МКИО  :: Извлекаю данные из " + newfilename;

                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        ExcelPackage excelFile = new ExcelPackage(filename);
                        for (int l = 0; l < 2; l++)
                        {
                            ExcelWorksheet worksheet = excelFile.Workbook.Worksheets[l];
                            totalRows = worksheet.Dimension.End.Row;
                            totalColums = worksheet.Dimension.End.Column;

                            excelTable = new string[totalRows, totalColums];

                            for (int rowIndex = 1; rowIndex <= totalRows; rowIndex++)
                            {
                                IEnumerable<string> row = worksheet.Cells[rowIndex, 1, rowIndex, totalColums].Select(c => c.Value == null ? "" : c.Value.ToString());

                                List<string> list = row.ToList<string>();

                                for (int i = 0; i < list.Count; i++)
                                {
                                    excelTable[rowIndex - 1, i] = Convert.ToString(list[i]);

                                }
                            }
                            string temp1 = "", temp2 = "";

                            string MessageName = "";
                            string LineDB = "";
                            string SubLine = "";
                            string WordNumbers = "";
                            string SubAdress = "";
                            string NumWord = "";
                            string Coment = "";
                            string NameParam = "";
                            string Bits = "";
                            int Rbit = 0;
                            int Lbit = 0;
                            string PriemnikRUS = "";
                            string PriemnikENG = "";
                            string sqlExpression = "";
                            string sqlExpression2 = "";

                            using (SqlConnection connection = new SqlConnection(connectionString))

                            {
                                connection.Open();

                                MessageName = excelTable[2, 2].ToString() + "_TO_" + excelTable[4, 2].ToString();
                                LineDB = "МКИО".ToString();
                                SubLine = excelTable[5, 2].ToString();
                                WordNumbers = "0";
                                SubAdress = "0";
                                PriemnikRUS = excelTable[4, 1].ToString();
                                PriemnikENG = excelTable[4, 2].ToString();

                                for (int i = 0; i < totalRows - 1; i++)
                                {
                                    for (int j = 0; j < totalColums; j++)
                                    {
                                        if (i > 10 && j == 0 && (excelTable[i + 1, j] == "" || excelTable[i + 1, j] == " ") && (excelTable[i, j] != "" || excelTable[i, j] == " "))
                                        {
                                            temp1 = excelTable[i, 0];
                                            temp2 = excelTable[i, 12];
                                        }
                                        if (i > 10 && j == 0 && (excelTable[i, j] == "" || excelTable[i, j] == " ") && excelTable[i, 2] != "")
                                        {
                                            excelTable[i, 0] = temp1;
                                            excelTable[i, 12] = temp2;
                                        }
                                        Console.Write(excelTable[i, j] + " ");
                                    }
                                    if (i > 10 && excelTable[i, 9].ToString() != null && excelTable[i, 9].ToString() != " " && excelTable[i, 9].ToString() != "")
                                    {

                                        NumWord = excelTable[i, 0].ToString();
                                        Coment = excelTable[i, 1].ToString();
                                        NameParam = excelTable[i, 2].ToString();
                                        Bits = excelTable[i, 9].ToString();

                                        LRbits(Bits, ref Rbit, ref Lbit);
                                        sqlExpression = "INSERT INTO Words (MessageName, Line, SubLine, WordNumbers, SubAdress, NumWord,Coment, NameParam, LBit, RBit, PriemnikRUS, PriemnikENG) VALUES ('" + MessageName + "', '" + LineDB + "', '" + SubLine;
                                        sqlExpression2 = sqlExpression + "', " + WordNumbers + ", " + SubAdress + ", " + NumWord + ", '" + Coment + "', '" + NameParam + "', " + Lbit + ", " + Rbit + ", '" + PriemnikRUS + "', '" + PriemnikENG + "')";

                                        Connection = new SqlConnection(connectionString);
                                        Connection.Open();
                                        SqlCommand createCommand = new SqlCommand(sqlExpression2, Connection);
                                        createCommand.ExecuteNonQuery();

                                        Connection.Close();

                                        Console.Write(sqlExpression2);
                                        Console.Write("\n");
                                    }
                                }
                            }
                        }

                        Console.WriteLine("Подключение закрыто...");
                        LogText.Text += "    Готово" + "\n";
                    }
                    else if (newfilename.StartsWith("sig_") && newfilename.EndsWith("_dpk.xls"))
                    {
                        LogText.Text += "ДПК      :: Извлекаю данные из " + newfilename + "\n";
                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        ExcelPackage excelFile = new ExcelPackage(filename);
                        ExcelWorksheet worksheet = excelFile.Workbook.Worksheets[0];
                        totalRows = worksheet.Dimension.End.Row;
                        totalColums = worksheet.Dimension.End.Column;

                        excelTable = new string[totalRows, totalColums];

                        for (int rowIndex = 1; rowIndex <= totalRows; rowIndex++)
                        {
                            IEnumerable<string> row = worksheet.Cells[rowIndex, 1, rowIndex, totalColums].Select(c => c.Value == null ? "" : c.Value.ToString());

                            List<string> list = row.ToList<string>();

                            for (int i = 0; i < list.Count; i++)
                            {
                                excelTable[rowIndex - 1, i] = Convert.ToString(list[i]);

                            }
                        }
                        string temp1 = "", temp2 = "";

                        string MessageName = "";
                        string LineDB = "";
                        string SubLine = "";
                        string WordNumbers = "";
                        string SubAdress = "";
                        string NumWord = "";
                        string Coment = "";
                        string NameParam = "";
                        string Bits = "";
                        int Rbit = 0;
                        int Lbit = 0;
                        string PriemnikRUS = "";
                        string PriemnikENG = "";
                        string sqlExpression = "";
                        string sqlExpression2 = "";

                        using (SqlConnection connection = new SqlConnection(connectionString))

                        {
                            connection.Open();

                            MessageName = excelTable[2, 3].ToString() + "_TO_" + excelTable[5, 3].ToString();
                            LineDB = "ДПК".ToString();
                            SubLine = excelTable[6, 3].ToString();
                            WordNumbers = "0";
                            SubAdress = "0";
                            PriemnikRUS = excelTable[5, 2].ToString();
                            PriemnikENG = excelTable[5, 3].ToString();

                            for (int i = 0; i < totalRows - 1; i++)
                            {
                                for (int j = 0; j < totalColums; j++)
                                {
                                    if (i > 8 && j == 0 && (excelTable[i + 1, j] == "" || excelTable[i + 1, j] == " ") && (excelTable[i, j] != "" || excelTable[i, j] == " "))
                                    {
                                        temp1 = excelTable[i, 0];
                                        temp2 = excelTable[i, 13];
                                    }
                                    if (i > 8 && j == 0 && (excelTable[i, j] == "" || excelTable[i, j] == " ") && excelTable[i, 2] != "")
                                    {
                                        excelTable[i, 0] = temp1;
                                        excelTable[i, 13] = temp2;
                                    }
                                    //Console.Write(excelTable[i, j] + " ");
                                }
                                if (i > 8 && excelTable[i, 8].ToString() != null && excelTable[i, 8].ToString() != " " && excelTable[i, 8].ToString() != "")
                                {

                                    NumWord = excelTable[i, 0].ToString();
                                    Coment = excelTable[i, 1].ToString();
                                    NameParam = excelTable[i, 2].ToString();
                                    Bits = excelTable[i, 9].ToString();

                                    LRbits(Bits, ref Rbit, ref Lbit);
                                    sqlExpression = "INSERT INTO Words (MessageName, Line, SubLine, WordNumbers, SubAdress, NumWord,Coment, NameParam, LBit, RBit, PriemnikRUS, PriemnikENG) VALUES ('" + MessageName + "', '" + LineDB + "', '" + SubLine;
                                    sqlExpression2 = sqlExpression + "', " + WordNumbers + ", " + SubAdress + ", " + NumWord + ", '" + Coment + "', '" + NameParam + "', " + Lbit + ", " + Rbit + ", '" + PriemnikRUS + "', '" + PriemnikENG + "')";

                                    Connection = new SqlConnection(connectionString);
                                    Connection.Open();
                                    SqlCommand createCommand = new SqlCommand(sqlExpression2, Connection);
                                    createCommand.ExecuteNonQuery();

                                    Connection.Close();

                                    Console.Write(sqlExpression2);
                                    Console.Write("\n");
                                }
                            }
                        }
                        Console.WriteLine("Подключение закрыто...");
                        LogText.Text += " Готово" + "\n";
                    }
                }
            }
        }



        public MainWindow()
        {
            InitializeComponent();
            DownloadBD.IsEnabled = false;
            UpdateBD.IsEnabled = false;
            TreeItem.IsEnabled = false;
            SearchItem.IsEnabled = false;
            TrashItem.IsEnabled = false;

            actions = new System.Action[] { С70parsing, T50parsing, Cy57parsing, Cy35parsing };
            Connection = new SqlConnection(connectionString);
            Connection.Open();
            string cmd = "SELECT COUNT(*) FROM Words";
            SqlCommand createCommand = new SqlCommand(cmd, Connection);
            int result = (int)createCommand.ExecuteScalar();
            if (result > 0)
            {
                DialogResult dialogResult = MessageBox.Show("В приложении имеется загруженная база данных!\nПродолжить работу с ней?", "Обнаружена БД", MessageBoxButtons.YesNo);
                if (dialogResult == System.Windows.Forms.DialogResult.Yes)
                {
                    DownLoadTree();
                    DownLoadTable();
                    UpdateBD.IsEnabled = true;
                    TreeItem.IsEnabled = true;
                    SearchItem.IsEnabled = true;
                    TrashItem.IsEnabled = true;
                }
                else if (dialogResult == System.Windows.Forms.DialogResult.No)
                {
                    cmd = "TRUNCATE TABLE Words";
                    createCommand = new SqlCommand(cmd, Connection);
                    createCommand.ExecuteNonQuery();
                    Thread.Sleep(10000); // засыпаем на 10 секунд 
                }
            }
            Connection.Close();

        }

        private void ComboBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            switch (ComboBox1.SelectedIndex)
            {
                case 0:
                    LogText.Text += "Выбранный самолёт - С-70" + "\n" + "2. Нажмите \"Загрузить БД\" и выберите корневую папку со СТИС С-70" + "\n";
                    plane = 0;
                    DownloadBD.IsEnabled = true;
                    UpdateBD.IsEnabled = false;
                    break;
                case 1:
                    LogText.Text += "Выбранный самолёт - T50" + "\n" + "2. Нажмите \"Загрузить БД\" и выберите корневую папку со СТИС Т50" + "\n";
                    plane = 1;
                    DownloadBD.IsEnabled = true;
                    UpdateBD.IsEnabled = false;
                    break;
                case 2:
                    LogText.Text += "Выбранный самолёт - Cу57" + "\n" + "2. Нажмите \"Загрузить БД\" и выберите корневую папку со СТИС Су-57" + "\n";
                    plane = 2;
                    DownloadBD.IsEnabled = true;
                    UpdateBD.IsEnabled = false;
                    break;
                case 3:
                    LogText.Text += "Выбранный самолёт - Су35" + "\n" + "2. Нажмите \"Загрузить БД\" и выберите корневую папку со СТИС Су-35" + "\n";
                    plane = 3;
                    DownloadBD.IsEnabled = true;
                    UpdateBD.IsEnabled = false;
                    break;
                default:
                    break;
            }
        }

        private void DownloadBD_Click(object sender, RoutedEventArgs e)
        {
            WinForms.FolderBrowserDialog dialog = new WinForms.FolderBrowserDialog();
            WinForms.DialogResult result = dialog.ShowDialog();

            if (result == WinForms.DialogResult.OK)
            {
                DownloadBD.IsEnabled = false;

                folder = dialog.SelectedPath;
                LogText.Text += ("Загружаю БД из ").ToString() + folder + Environment.NewLine;
                myUpdate = actions[plane];
                myUpdate();
                LogText.Text += "3. Данные извлечены и обработаны, можно переходить на другие вкладки" + "\n";
                UpdateBD.IsEnabled = true;
                TreeItem.IsEnabled = true;
                SearchItem.IsEnabled = true;
                TrashItem.IsEnabled = true;
                
                DownLoadTree();
            }
            else
            {
                LogText.Text += "!БД не выбрана!" + "\n";
                result = 0;
            }

        }

        private void UpdateBD_Click(object sender, RoutedEventArgs e)
        {
            LogText.Text += ("Обновляю БД из ").ToString() + folder + Environment.NewLine;
            Connection = new SqlConnection(connectionString);
            Connection.Open();
            string cmd = "TRUNCATE TABLE Words";
            SqlCommand createCommand = new SqlCommand(cmd, Connection);

            createCommand.ExecuteNonQuery();
            Thread.Sleep(10000); // засыпаем на 10 секунд
            myUpdate = actions[plane];
            myUpdate();
            
            DownLoadTree();
            LogText.Text += "3. Данные обновлены, можно переходить на другие вкладки" + "\n";
        }

        private void DownLoadTree()
        {
            Connection = new SqlConnection(connectionString);
            Connection.Open();
            string cmd = "SELECT  DISTINCT Line AS Канал, SubLine AS Линия, PriemnikEng AS Приёмник\r\nFROM Words";
            SqlCommand createCommand = new SqlCommand(cmd, Connection);
            createCommand.ExecuteNonQuery();

            DataAdapter = new SqlDataAdapter(createCommand);
            table = new DataTable("Words");
            DataAdapter.Fill(table);
            dataGridView2.ItemsSource = table.DefaultView;
            Connection.Close();

        }
        private void DownLoadTable()
        {
            Connection = new SqlConnection(connectionString);
            Connection.Open();
            string cmd = "SELECT Line AS 'Канал', SubLine AS 'Линия', MessageName AS 'Слово', Coment AS 'Обозначение сигнала', NameParam AS 'Имя сигнала' FROM Words WHERE NewTable = 1";
            SqlCommand createCommand = new SqlCommand(cmd, Connection);
            createCommand.ExecuteNonQuery();

            DataAdapter = new SqlDataAdapter(createCommand);
            table = new DataTable("Words");
            DataAdapter.Fill(table);
            dataGridView6.ItemsSource = table.DefaultView;
            Connection.Close();
        }
        private void addOyToTrash_Click(object sender, RoutedEventArgs e)
        {
            DataRowView row = (DataRowView)dataGridView2.SelectedItems[0];
            var canal = row.Row.ItemArray[0].ToString();
            var line = row.Row.ItemArray[1].ToString();
            var priemnik = row.Row.ItemArray[2].ToString();
            Connection = new SqlConnection(connectionString);
            Connection.Open();
            string cmd = "UPDATE Words \r\nSET NewTable = 1\r\nWHERE Line = '" + canal + "' AND SubLine = '" + line + "' AND PriemnikEng = '" + priemnik + "'";
            SqlCommand createCommand = new SqlCommand(cmd, Connection);
            createCommand.ExecuteNonQuery();
            Connection.Close();
        }

        private void addWordToTrash_Click(object sender, RoutedEventArgs e)
        {
            DataRowView row_2 = (DataRowView)dataGridView3.SelectedItems[0];
            
            var message = row_2.Row.ItemArray[0].ToString();
            var subadress = row_2.Row.ItemArray[1].ToString();
            var numword = row_2.Row.ItemArray[2].ToString();
            var priemnikrus = row_2.Row.ItemArray[3].ToString();
            var priemnikeng = row_2.Row.ItemArray[4].ToString();
            Connection = new SqlConnection(connectionString);
            Connection.Open();
            string cmd = "UPDATE Words \r\nSET NewTable = 1\r\nWHERE MessageName = '" + message + "' AND SubAdress = '" + subadress + "' AND WordNumbers = '" + numword + "' AND PriemnikRus = '" + priemnikrus + "' AND PriemnikEng = '" + priemnikeng + "'";
            SqlCommand createCommand = new SqlCommand(cmd, Connection);
            createCommand.ExecuteNonQuery();
            Connection.Close();
        }

        private void addParamToTrash_Click(object sender, RoutedEventArgs e)
        {
            DataRowView row_3 = (DataRowView)dataGridView1.SelectedItems[0];

            var numword = row_3.Row.ItemArray[0].ToString();
            var nameparam = row_3.Row.ItemArray[1].ToString();
            var obaznach = row_3.Row.ItemArray[2].ToString();
            var lbit = row_3.Row.ItemArray[3].ToString();
            var rbit = row_3.Row.ItemArray[4].ToString();

            Connection = new SqlConnection(connectionString);
            Connection.Open();
            string cmd = "UPDATE Words \r\nSET NewTable = 1\r\nWHERE NumWord = '" + numword + "' AND NameParam = '" + nameparam + "' AND Coment = '" + obaznach + "' AND LBit = '" + lbit + "' AND RBit = '" + rbit + "'";
            SqlCommand createCommand = new SqlCommand(cmd, Connection);
            createCommand.ExecuteNonQuery();
            Connection.Close();

        }

        private void Selector_OnSelectionChanged_Tab_conrol(object sender, SelectionChangedEventArgs e)
        {
            if(product.SelectedIndex == 3)
            {
                DownLoadTable();
            }
        }

        private void dataGridView2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Connection = new SqlConnection(connectionString);
            Connection.Open();

            DataRowView row = (DataRowView)dataGridView2.SelectedItems[0];
            string canal = row.Row.ItemArray[0].ToString();
            string line = row.Row.ItemArray[1].ToString();
            string priemnik = row.Row.ItemArray[2].ToString();

            string cmd = "SELECT DISTINCT MessageName AS 'Имя сообщения', SubAdress AS 'Подадрес', WordNumbers AS 'Количество слов', PriemnikRus AS 'Приёмник', PriemnikEng AS 'Приёмнкик (англ)' FROM Words WHERE Line = '" + canal + "' AND SubLine = '" + line + "' AND PriemnikEng = '" + priemnik + "'";
            SqlCommand createCommand = new SqlCommand(cmd, Connection);
            createCommand.ExecuteNonQuery();

            DataAdapter = new SqlDataAdapter(createCommand);
            table = new DataTable("Words");
            DataAdapter.Fill(table);
            dataGridView3.ItemsSource = table.DefaultView;
            Connection.Close();
        }

        private void dataGridView3_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dataGridView3.CurrentColumn != null)
            {
                Connection = new SqlConnection(connectionString);
                Connection.Open();
                DataRowView row = (DataRowView)dataGridView3.SelectedItems[0];
                string mes = row.Row.ItemArray[0].ToString();

                string cmd = "SELECT DISTINCT NumWord AS 'Номер слова', NameParam AS 'Имя сигнала', Coment AS 'Обозначение сигнала', LBit AS 'Левый бит', RBit AS 'Правый бит' FROM Words WHERE MessageName = '" + mes + "'";

                SqlCommand createCommand = new SqlCommand(cmd, Connection);
                createCommand.ExecuteNonQuery();

                DataAdapter = new SqlDataAdapter(createCommand);
                table = new DataTable("Words");
                DataAdapter.Fill(table);
                dataGridView1.ItemsSource = table.DefaultView;
                Connection.Close();
            }
        }
        private void dataGridView4_SelectionChanged(object sender, SelectionChangedEventArgs e)
        { }

            private void GenerateTk_Click(object sender, RoutedEventArgs e)
        {
            WinForms.FolderBrowserDialog dialog = new WinForms.FolderBrowserDialog();
            WinForms.DialogResult result = dialog.ShowDialog();
            if (result == WinForms.DialogResult.OK)
            {
                string folder_2 = dialog.SelectedPath + "\\res.txt";
                StreamWriter sw = new StreamWriter(folder_2);
                Connection = new SqlConnection(connectionString);
                Connection.Open();
                SqlCommand thisCommand = Connection.CreateCommand();
                thisCommand.CommandText = "SELECT NameParam, MessageName, NumWord, LBit, RBit, Coment FROM Words WHERE NewTable = 1 ";
                SqlDataReader thisReader = thisCommand.ExecuteReader();
                string res = string.Empty;
                while (thisReader.Read())
                {
                    sw.WriteLine("BITFIELD(" + thisReader["NameParam"] + ", " + thisReader["MessageName"] + "[" + thisReader["NumWord"] + "], " + thisReader["LBit"] + ", " + thisReader["RBit"] + "); //" + thisReader["Coment"]);
                }
                thisReader.Close();
                Connection.Close();
                sw.Close();
                MessageBox.Show("Результат записан в " + folder_2);
            }
            else
            {
                MessageBox.Show("!Путь не выбран!"); 
            }

           

        }

        private void UpdateTrash_Click(object sender, RoutedEventArgs e)
        {
            DownLoadTable();
        }

       

        private void TBSearch_Button_Click(object sender, RoutedEventArgs e)
        {
            string str = TBSearch.Text.ToString();
            Connection = new SqlConnection(connectionString);
            Connection.Open();
            string cmd = "SELECT * FROM Words WHERE NameParam = '" + TBSearch.Text.ToString() + "'";
            SqlCommand createCommand = new SqlCommand(cmd, Connection);
            createCommand.ExecuteNonQuery();

            DataAdapter = new SqlDataAdapter(createCommand);
            table = new DataTable("Words");
            DataAdapter.Fill(table);
            dataGridView5.ItemsSource = table.DefaultView;
            Connection.Close();
        }

        private void DeleteTrash_Click(object sender, RoutedEventArgs e)
        {
            Connection = new SqlConnection(connectionString);
            Connection.Open();
            string cmd = "UPDATE Words \r\nSET NewTable = 0\r\nWHERE NewTable = 1";
            SqlCommand createCommand = new SqlCommand(cmd, Connection);
            createCommand.ExecuteNonQuery();
            Connection.Close();
            DownLoadTable();
        }

        private void deleteFromTrash_2_Click(object sender, RoutedEventArgs e)
        {        
            DataRowView row_10 = (DataRowView)dataGridView6.SelectedItems[0];

            var canal = row_10.Row.ItemArray[0].ToString();
            var subline = row_10.Row.ItemArray[1].ToString();
            var mesname = row_10.Row.ItemArray[2].ToString();
            var coment = row_10.Row.ItemArray[3].ToString();
            var nameparam = row_10.Row.ItemArray[4].ToString();

            Connection = new SqlConnection(connectionString);
            Connection.Open();
            string cmd = "UPDATE Words \r\nSET NewTable = 0\r\nWHERE Line = '" + canal + "' AND SubLine = '" + subline + "' AND MessageName = '" + mesname + "' AND Coment = '" + coment + "' AND NameParam = '" + nameparam + "'";

            SqlCommand createCommand = new SqlCommand(cmd, Connection);
            createCommand.ExecuteNonQuery();
            Connection.Close();
            DownLoadTable();
        }
        private void addToTrash_Click(object sender, RoutedEventArgs e)
        {
            DataRowView row_3 = (DataRowView)dataGridView5.SelectedItems[0];

            var numword = row_3.Row.ItemArray[6].ToString();
            var nameparam = row_3.Row.ItemArray[8].ToString();
            var obaznach = row_3.Row.ItemArray[7].ToString();
            var lbit = row_3.Row.ItemArray[9].ToString();
            var rbit = row_3.Row.ItemArray[10].ToString();

            Connection = new SqlConnection(connectionString);
            Connection.Open();
            string cmd = "UPDATE Words \r\nSET NewTable = 1\r\nWHERE NumWord = '" + numword + "' AND NameParam = '" + nameparam + "' AND Coment = '" + obaznach + "' AND LBit = '" + lbit + "' AND RBit = '" + rbit + "'";
            SqlCommand createCommand = new SqlCommand(cmd, Connection);
            createCommand.ExecuteNonQuery();
            Connection.Close();
        }
        private void deleteFromTrash_Click(object sender, RoutedEventArgs e)
        {
            Connection = new SqlConnection(connectionString);
            Connection.Open();

            DataRowView row_4 = (DataRowView)dataGridView5.SelectedItems[0];

            var ID = row_4.Row.ItemArray[0].ToString();
            string cmd = "UPDATE Words \r\nSET NewTable = 0\r\nWHERE ID = '" + ID + "'";

            SqlCommand createCommand = new SqlCommand(cmd, Connection);
            createCommand.ExecuteNonQuery();
            Connection.Close();
            DownLoadTable();
        }
    }
}
