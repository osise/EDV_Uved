using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
using Telerik.WinControls.UI;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
using System.IO;


namespace TelerikWinFormsApp14
{
    public partial class RadForm1 : RadForm
    {
        Word._Application application;

        Object missingObj = System.Reflection.Missing.Value;
        Object trueObj = true;
        Object falseObj = false;
        string fullName = Path.Combine(Environment.ExpandEnvironmentVariables("%temp%"), "Template.doc");
        DataSet ds = new DataSet();
        List<string> ListOfSnils = new List<string>();

        string FIO;
        string fa;
        string im;
        string ot;
        string birthdate;
        string address;
        string numresh;
        string datresh;
        string sroks;
        string srokpo;
        string category;
        string snils;
        string currentday;

        public RadForm1()
        {
            RadMessageBox.SetThemeName("Fluent");
            InitializeComponent();
            try
            {
                string subPath = @"C:\Uved\"; // your code goes here

                bool exists = System.IO.Directory.Exists(subPath);

                if (!exists)
                { System.IO.Directory.CreateDirectory(subPath); }
                File.WriteAllText(fullName, NVP.Properties.Resources.Template, Encoding.Default) ;
               
                 
            }
            catch (Exception)
            {
              
            }
           
        }

            private void Replword(string stub, string text, Word.Document wordDoc)
            {
            try
            {
                var range = wordDoc.Content;
                range.Find.ClearFormatting();
                range.Find.Execute(FindText: stub, ReplaceWith: text);

            }
            catch (Exception)
            {
                RadMessageBox.Show("Ошибка доступа к шаблону. Обратитесь к разработчику");
                throw;
            }
             
            }
            public void SelectIdFromDB2(int i)
            {
                ds.Clear();
            

                using (OleDbConnection connection = new OleDbConnection
                        (
                        "Provider=IBMDADB2; " +
                        "UID = db2admin; " +
                        "PWD = Iw9o5JBW; " +
                        "DATABASE = ROS; " +
                        "HOSTNAME = 10.55.0.184; " +
                        "PORT = 50000; " +
                        "PROTOCOL = TCPIP;"
                        ))
                {

                    try
                    {
                        connection.Open();
                        // if (connection.State == ConnectionState.Open)
                        // { MessageBox.Show("1"); }
                        
                        // if (connection.State == ConnectionState.Open)
                        //{ MessageBox.Show("1"); }


                        OleDbDataAdapter adapter = new OleDbDataAdapter();
                        OleDbCommand command;

                        command = new OleDbCommand("select M.FA, M.IM, M.OT, M.RDAT, M.ADR_INDEX, M.ADRFAKT, p.SROKS, p.SROKPO, p.NRESH, p.dresh, L.T " +
                        "from GSP.GSP p " +
                        "LEFT JOIN PF.MAN M ON M.ID = P.ID " +
                        "left join GSP.LGKAT L ON p.L1 = L.K " +
                        "WHERE M.NPERS = ? " +
                        "AND (p.DRESH = (SELECT MAX(p1.DRESH) FROM GSP.GSP p1 WHERE M.ID = p1.ID GROUP BY p1.ID) " +
                        "OR p.DRESH is NULL)", connection);
                        command.Parameters.Add("M.NPERS", OleDbType.VarChar).Value = ListOfSnils[i];
                        adapter.SelectCommand = command;
                        adapter.Fill(ds);

                        fa = ds.Tables[0].Rows[0].Field<string>("FA");
                        im = ds.Tables[0].Rows[0].Field<string>("IM");
                        ot = ds.Tables[0].Rows[0].Field<string>("OT");
                        FIO = fa + " " + im + " " + ot;
                        currentday = DateTime.Now.Date.ToString("dd.MM.yyyy");
                        snils = ListOfSnils[i];
                        birthdate = ds.Tables[0].Rows[0].Field<DateTime>("RDAT").ToString("dd.MM.yyyy") + " г.р.";
                        address = ds.Tables[0].Rows[0].Field<string>("ADRFAKT");
                        numresh = ds.Tables[0].Rows[0].Field<string>("NRESH");

                        if (ds.Tables[0].Rows[0].Field<DateTime?>("DRESH") == null) { datresh = "Дата решения отсутствует"; }
                        else datresh = ds.Tables[0].Rows[0].Field<DateTime>("DRESH").ToString("dd.MM.yyyy");

                        if (ds.Tables[0].Rows[0].Field<DateTime?>("SROKS") == null) { sroks = "Дата назначения отсутствует"; }
                        else sroks = ds.Tables[0].Rows[0].Field<DateTime>("SROKS").ToString("dd.MM.yyyy");

                        if (ds.Tables[0].Rows[0].Field<DateTime?>("SROKPO") == null) { srokpo = "бессрочно"; }
                        else srokpo = ds.Tables[0].Rows[0].Field<DateTime>("SROKPO").ToString("dd.MM.yyyy");

                        category = ds.Tables[0].Rows[0].Field<string>("T");

                    
                     //radRichTextEditor1.Text = FIO + "\n" + birthdate + " года рождения" + "\n" + snils + "\n" + address + "\n" + numresh + " " + datresh + "\n" + sroks + " " + srokpo + "\n" + category;
                       // string ass = radRichTextEditor1.Text;
                   
                        connection.Dispose();
                        connection.Close();
                        // if (connection.State == ConnectionState.Closed)
                        // { MessageBox.Show("2"); }
                    }

                    catch (Exception)
                    {
                        RadMessageBox.Show("Ошибка доступа к базе НВП. Проверьте подключение к локальной сети");
                    }
                    InsertData();
            }
            }

        public void InsertData()
            { 
            application = new Word.Application();
            var wordDoc = application.Documents.Open(fullName);
                try
                {
                    Replword("{FIO}", FIO, wordDoc);
                    Replword("{DATER}", birthdate, wordDoc);
                    Replword("{ADDR}", address, wordDoc);
                    Replword("{SNILS}", snils, wordDoc);
                    Replword("{DATEC}", currentday, wordDoc);
                    Replword("{NAME}", im, wordDoc);
                    Replword("{OTCH}", ot, wordDoc);
                    Replword("{NRESH}", numresh, wordDoc);
                    Replword("{RESHDATE}", datresh, wordDoc);
                    Replword("{SROKS}", sroks, wordDoc);
                    Replword("{SROKPO}", srokpo, wordDoc);
                    Replword("{KATEG}", category, wordDoc);
                }
                catch (Exception)
                {
                    RadMessageBox.Show("Ошибка вставки данных в шаблон");
                    throw;
                }

                try
                {
                
               
                 wordDoc.SaveAs(@"C:\Uved\Уведомление о назначении ЕДВ " + fa + " " + im + " " + ot + ".doc"); 
                    wordDoc.Close();
                    application.Quit();

                //MessageBox.Show("Готово");
                ds.Clear();
                }
                catch (Exception)
                {
                    RadMessageBox.Show("Ошибка сохранения файла");
                    throw;
                }
     
            }

            private static string ExctraxtIni(string s)
            {
                var inits = Regex.Match(s, @"(\w+)\s+(\w+)\s+(\w+)").Groups;
                return string.Format("{0} {1}. {2}.", inits[1], inits[2].Value[0], inits[3].Value[0]);
            }

        private void radButton2_Click(object sender, EventArgs e)
        {
            try
            {
                string ass = radRichTextEditor1.Text;
                ass = ass.Replace(System.Environment.NewLine, string.Empty);
                if (string.IsNullOrEmpty(ass) || Regex.IsMatch(ass, @"[a-zA-Zа-яА-Я]"))
                { RadMessageBox.Show("Данные отсутствуют или содержат некорректные символы"); }
                else
                {
                    ListOfSnils = Enumerable.Range(0, ass.Length / 14).Select(i => ass.Substring(i * 14, 14)).ToList();
                    radRichTextEditor1.Text = string.Empty;
                    for (int i = 0; i < ListOfSnils.Count; i++)
                    {
                        SelectIdFromDB2(i);
                        radRichTextEditor1.Text += FIO + " - уведомление создано" + "\r\n";
                        radLabel2.Refresh();
                        radLabel2.Text = Convert.ToString(i+1) + "/" + Convert.ToString(ListOfSnils.Count);
                    }
                    radRichTextEditor1.Text += "Работа окончена" + "\r\n";
                }
            }
            catch (Exception)
            {
                RadMessageBox.Show("Данные некорректны");
                application.Quit();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(application);
                throw;
            }
            
        }

        private void radButton1_Click(object sender, EventArgs e)
        {
            radRichTextEditor1.Text = string.Empty;
        }

        private void radButton3_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer", "C:\\Uved");
        }

        private void radRichTextEditor1_TextChanged(object sender, EventArgs e)
        {

        }

        private void radButton5_Click(object sender, EventArgs e)
        {
            RadMessageBox.Show(" Разработчик: Осипов Сергей \r\n ОПФР по Курганской области 2021г. \r\n Техподдержка: тел. 11-79");
        }
    }
}
