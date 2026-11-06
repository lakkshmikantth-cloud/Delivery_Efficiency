using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;
using System.Reflection;
using System.IO;
using System.Globalization;
using System.Data.SqlClient;
using System.Data;
using Microsoft.Office.Interop.Excel;

namespace IndentStatus_Automail
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());
            try
            {
                AA();
                string Attachment =  System.Windows.Forms.Application.StartupPath + "\\DeliveryEfficiency_Details_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
                MailMessage message = new MailMessage();
                SmtpClient smtp = new SmtpClient();
                string ToMailsList = "";
                string CcMailList = "";
                string SMTPHost = "smtp.gmail.com";
                string FromMail = "order@elmeasure.com";
                string Password = "cgju gymh cppd udgd"; //"uoix ikzn glmb stvz";//"Elecon47$";
                //string MailSubject = "Delivery Efficiency Report as On Month:" + DateTime.Now.ToMonthName() + "";
                string MailSubject = "Delivery Efficiency Report_Meters as On Month:" + DateTime.Now.AddMonths(-1).ToString("MMMM") + "";
                string EStext = "";
                message.From = new MailAddress(FromMail);
                SqlConnection con11 = new SqlConnection("Data Source=192.168.4.150;Initial Catalog=SAP_ELM;User ID=sa;Password=Elnetsrv123*");
                SqlDataAdapter da11 = new SqlDataAdapter();
                System.Data.DataTable dt1 = new System.Data.DataTable();
                SqlCommand cmd11 = new SqlCommand("SELECT T0.[U_Email] , T0.[U_CcMail] FROM [dbo].[@AUTO_EMAIL_RPT]  T0 WHERE cast(T0.[U_rptname]as varchar (max)) ='DeliveryEfficiencyAsOnMonth_AutoMail'", con11);

                da11.SelectCommand = cmd11;
                da11.Fill(dt1);
                // LIVE
               ToMailsList = dt1.Rows[0]["U_Email"].ToString();

               CcMailList = dt1.Rows[0]["U_CcMail"].ToString();

                //TEST
                //ToMailsList = "shantha@elmeasure.com";
               // ToMailsList = "sangayya@elmeasure.com";
                //ToMailsList = "ganeshbabu@elmeasure.com";
                //ToMailsList = "bharani.ns@elmeasure.com,";
                //ToMailsList = "sangayya@elmeasure.com,shantha@elmeasure.com";
              //  CcMailList = "srilakshmi@elmeasure.in,lakshmikanth@elmeasure.in";


                string TotalMailsList = ToMailsList;
                string MailsList = TotalMailsList.TrimEnd(',');
                string[] Multi = MailsList.Split(',');
                foreach (string MultiMailID in Multi)
                {
                    message.To.Add(new MailAddress(MultiMailID));
                }
                string CCEmailList = CcMailList;
                if (CcMailList != "")
                {
                    string CcList = CCEmailList.TrimEnd(',').TrimStart(',');
                    string[] Multi1 = CcList.Split(',');
                    foreach (string MultiMailID1 in Multi1)
                    {
                        string p = MultiMailID1.Replace("\"", "");
                        message.CC.Add(new MailAddress(p));
                    }
                }
                string textBody = " <table border=" + 1 + " cellpadding=" + 0 + " cellspacing=" + 0 + " width = " + 400 + "><tr bgcolor='#4da6ff'><td><b>Column 1</b></td> <td> <b> Column 2</b> </td></tr>";

                textBody += "</table>";
                message.Subject = MailSubject;
                //message.Body = Message;
                Assembly assembly = Assembly.GetExecutingAssembly();//HTMLPage3
                // StreamReader reader = new StreamReader(System.Windows.Forms.Application.StartupPath + "\\" + "IndentStatus.htm");
                StreamReader reader = new StreamReader(System.Windows.Forms.Application.StartupPath + "\\" + "HTMLPage3.htm");

                string readFile = reader.ReadToEnd();
                string myString = "";
                myString = readFile;
                CultureInfo hindi = new CultureInfo("hi-IN");
                SqlConnection con = new SqlConnection("Data Source=192.168.4.150;Initial Catalog=SAP_ELM;User ID=sa;Password=Elnetsrv123*");
                string sp = "DeliveryEfficiencyAsOnMonth_AutoMail";
                SqlCommand cmdProc = new SqlCommand(sp, con);
                cmdProc.CommandType = CommandType.StoredProcedure;

                cmdProc.CommandTimeout = 36000;
                SqlDataAdapter daELM1 = new SqlDataAdapter(cmdProc);
                CultureInfo cultureInfo = new CultureInfo("en-IN");
                System.Data.DataTable dtELM = new System.Data.DataTable();
                //WriteDataTableToExcel(dtELM, "AAA", System.Windows.Forms.Application.StartupPath, "TEST");
                daELM1.SelectCommand .CommandTimeout = 36000;
                daELM1.Fill(dtELM);

                //decimal parsed = decimal.Parse(dtELM.Rows[0]["#A3"].ToString().Trim(), CultureInfo.InvariantCulture);

                string TLI = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#TLI"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#TLI"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");

                string IOA = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#IOA"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#IOA"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");


                string SOA = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#SOA"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#SOA"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                
                string NBOI = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#NBOI"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#NBOI"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");


                string UAH = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#UAH"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#UAH"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");

                
                string CAT1 = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#CAT1"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#CAT1"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                string ACT1 = string.Format(hindi, "{0:c}", decimal.Parse(dtELM.Rows[0]["#ACT1"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#ACT1"].ToString().Trim(), CultureInfo.InvariantCulture)).Replace("₹", "");
                string L3CT1 = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#L3CT1"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#L3CT1"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");

                string G3CT1 = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#G3CT1"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#G3CT1"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                string G130CT1 = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#G130CT1"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#G130CT1"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                



                string ACTD1 = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#ACTD1"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#ACTD1"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");

                string CAT2 = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#CAT2"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#CAT2"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                string ACT2 = string.Format(hindi, "{0:c}", decimal.Parse(dtELM.Rows[0]["#ACT2"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#ACT2"].ToString().Trim(), CultureInfo.InvariantCulture)).Replace("₹", "");
                string L7CT2 = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#L7CT2"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#L7CT2"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");

                string G7CT2 = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#G7CT2"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#G7CT2"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                string G130CT2 = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#G130CT2"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#G130CT2"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                


                string ACTD2 = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#ACTD2"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#ACTD2"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");


                string OCAT = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#OCAT"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#OCAT"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                string AOCT = string.Format(hindi, "{0:c}", decimal.Parse(dtELM.Rows[0]["#AOCT"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#AOCT"].ToString().Trim(), CultureInfo.InvariantCulture)).Replace("₹", "");
                string L20OCT = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#L20OCT"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#L20OCT"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");

                string G20OCT = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#G20OCT"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#G20OCT"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                string G130OCT = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#G130OCT"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#G130OCT"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                
                string ACTDO = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#ACTDO"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#ACTDO"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");



                string CATNS = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#CATNS"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#CATNS"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                string CTNSA = string.Format(hindi, "{0:c}", decimal.Parse(dtELM.Rows[0]["#CTNSA"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#CTNSA"].ToString().Trim(), CultureInfo.InvariantCulture)).Replace("₹", "");
                string CANSL20 = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#CANSL20"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#CANSL20"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");

                string CSG20 = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#CSG20"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#CSG20"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                string CNG130 = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#CNG130"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#CNG130"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                

                string SEFF = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#SEFF"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#SEFF"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                string PDIEF = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#PDIEF"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#PDIEF"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                string PDEFF = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#PDEFF"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#PDEFF"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                string DWSD = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#DWSD"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#DWSD"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");


                string DWOD = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#DWOD"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#DWOD"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                string DWTD = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#DWTD"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#DWTD"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                string RSTH = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#RSTH"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#RSTH"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");

                string BBAC = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#BBAC"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#BBAC"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                string PBNA = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#PBNA"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#PBNA"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                string CRLT = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#CRLT"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#CRLT"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                string RVDC = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#RVDC"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#RVDC"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                string ACLI = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#ACLI"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#ACLI"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");
                string CSATNDO = string.Format(hindi, "{0:c}", dtELM.Rows[0]["#CSATNDO"].ToString().Trim() == "" ? "0" : dtELM.Rows[0]["#CSATNDO"].ToString().Trim(), CultureInfo.InvariantCulture).Replace("₹", "");



                myString = myString.Replace("#DATE", DateTime.Now.AddMonths(-1).ToString("MMMM")+"-"+DateTime.Now.Year);
                //myString = myString.Replace("#DATE", DateTime.Now.AddMonths(-1).ToString("MMMM") + "-" + "2022");
                myString = myString.Replace("#TLI", TLI);
                myString = myString.Replace("#IOA", IOA);
                myString = myString.Replace("#DWOD", DWOD);
                myString = myString.Replace("#DWSD", DWSD);
                myString = myString.Replace("#DWTD", DWTD);
                myString = myString.Replace("#RSTH", RSTH);

                myString = myString.Replace("#BBAC", BBAC);
                myString = myString.Replace("#PBNA", PBNA);
                myString = myString.Replace("#CRLT", CRLT);
                myString = myString.Replace("#RVDC", RVDC);
                myString = myString.Replace("#ACLI", ACLI);
                myString = myString.Replace("#SOA", SOA);
                myString = myString.Replace("#NBOI", NBOI);
                myString = myString.Replace("#UAH", UAH);
                myString = myString.Replace("#CSATNDO", CSATNDO);
                
                myString = myString.Replace("#CAT1", CAT1);
                myString = myString.Replace("#ACT1", ACT1);
                myString = myString.Replace("#L3CT1", L3CT1);
                myString = myString.Replace("#G3CT1", G3CT1);
                myString = myString.Replace("#G130CT1", G130CT1);
                myString = myString.Replace("#ACTD1", ACTD1);
                myString = myString.Replace("#CAT2", CAT2);
                myString = myString.Replace("#ACT2", ACT2);
                myString = myString.Replace("#L7CT2", L7CT2);
                myString = myString.Replace("#G7CT2", G7CT2);
                myString = myString.Replace("#ACTD2", ACTD2);
                myString = myString.Replace("#G130CT2", G130CT2);
                myString = myString.Replace("#OCAT", OCAT);
                myString = myString.Replace("#AOCT", AOCT);
                myString = myString.Replace("#L20OCT", L20OCT);
                myString = myString.Replace("#G20OCT", G20OCT);
                myString = myString.Replace("#G130OCT", G130OCT);
                myString = myString.Replace("#ACTDO", ACTDO);
                myString = myString.Replace("#CATNS", CATNS);
                myString = myString.Replace("#CTNSA", CTNSA);
                myString = myString.Replace("#CANSL20", CANSL20);
                myString = myString.Replace("#CSG20", CSG20);
                myString = myString.Replace("#CNG130", CNG130);


                myString = myString.Replace("#SEFF", SEFF);
                myString = myString.Replace("#PDIEF", PDIEF);
                myString = myString.Replace("#PDEFF", PDEFF);

                message.Body = myString;

                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(Attachment);
                message.Attachments.Add(attachment);

                message.IsBodyHtml = true;

                message.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess;
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.Host = SMTPHost;
                smtp.Port = 587;
                smtp.EnableSsl = true;
                smtp.UseDefaultCredentials = false;

                smtp.Timeout = 100000;
                smtp.Credentials = new NetworkCredential(FromMail, Password);
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate(object s,
                System.Security.Cryptography.X509Certificates.X509Certificate certificate,
                System.Security.Cryptography.X509Certificates.X509Chain chain,
                System.Net.Security.SslPolicyErrors sslPolicyErrors)
                {
                    return true;
                };
                message.Priority = MailPriority.High;
                smtp.Send(message);
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.InnerException.ToString());
            }
            System.Windows.Forms.Application.Exit();
        }


        public static void AA()
        {
            try
            {
                System.Data.DataTable dt = new System.Data.DataTable();
                System.Data.DataTable dt1 = new System.Data.DataTable();
                System.Data.DataTable dt2 = new System.Data.DataTable();

                object missing = Type.Missing;
                SqlConnection con = new SqlConnection("Data Source=192.168.4.150;Initial Catalog=SAP_ELM;User ID=sa;Password=Elnetsrv123*");
                string sp = "DeliveryEfficiency_ELM";
                SqlCommand cmdProc = new SqlCommand(sp, con);
                cmdProc.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter daELM1 = new SqlDataAdapter(cmdProc);


                daELM1.SelectCommand.CommandTimeout = 600;

                daELM1.Fill(dt);
          

                if (dt == null || dt.Columns.Count == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");
                Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook oWB = oXL.Workbooks.Add(missing);
                Microsoft.Office.Interop.Excel.Worksheet oSheet = oWB.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                Microsoft.Office.Interop.Excel.Range excelCellrange;
                oSheet.Name = "DeliveryEfficiency-ELM";

                for (var i = 0; i < dt.Columns.Count; i++)
                {
                    oSheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;
                }
                for (var i = 0; i < dt.Rows.Count; i++)
                {
                    for (var j = 0; j < dt.Columns.Count; j++)
                    {
                        oSheet.Cells[i + 2, j + 1] = dt.Rows[i][j];
                    }
                }

                

                oXL.Visible = false;
                excelCellrange = oSheet.Range[oSheet.Cells[1, 1], oSheet.Cells[1, dt.Columns.Count]];
                excelCellrange.EntireColumn.AutoFit();
                

                excelCellrange = oSheet.Range[oSheet.Cells[1, 1], oSheet.Cells[1, dt.Columns.Count]];
                FormattingExcelCells(excelCellrange, "#FFFF00", System.Drawing.Color.Black, true);
                excelCellrange.EntireRow.Font.Bold = true;

                Range rg = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[1, 8];
                rg.EntireColumn.NumberFormat = "DD-MM-YYYY";

                Range rg6 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[1, 9];
                rg6.EntireColumn.NumberFormat = "DD-MM-YYYY";


                Range rg1 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[1, 10];
                rg1.EntireColumn.NumberFormat = "DD-MM-YYYY";

                Range rg7 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[1, 30];
                rg7.EntireColumn.NumberFormat = "DD-MM-YYYY";

                Range rg8 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[1, 31];
                rg8.EntireColumn.NumberFormat = "DD-MM-YYYY";

                Range rg9 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[1, 44];
                rg9.EntireColumn.NumberFormat = "DD-MM-YYYY";

                //Range rg10 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[1, 46];
                //rg10.EntireColumn.NumberFormat = "DD-MM-YYYY";

                Range rg11 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[1, 47];
                rg11.EntireColumn.NumberFormat = "DD-MM-YYYY";

                Range rg12 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[1, 48];
                rg12.EntireColumn.NumberFormat = "DD-MM-YYYY";

                Range rg13 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[1, 49];
                rg13.EntireColumn.NumberFormat = "DD-MM-YYYY";

                Range rg14 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[1, 50];
                rg14.EntireColumn.NumberFormat = "DD-MM-YYYY";

                Range rg15 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[1, 51];
                rg15.EntireColumn.NumberFormat = "DD-MM-YYYY";


                oSheet.Activate();
                oXL.ActiveWindow.SplitRow = 1;
                oXL.ActiveWindow.FreezePanes = true;
                
                oWB.SaveAs(System.Windows.Forms.Application.StartupPath + "\\DeliveryEfficiency_Details_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx");
                oWB.Close();
                oXL.Quit();

            }
            catch(Exception ex)
            {


            }
        }
        public static void FormattingExcelCells(Microsoft.Office.Interop.Excel.Range range, string HTMLcolorCode, System.Drawing.Color fontColor, bool IsFontbool)
        {
            range.Interior.Color = System.Drawing.ColorTranslator.FromHtml(HTMLcolorCode);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(fontColor);
            if (IsFontbool == true)
            {
                range.Font.Bold = IsFontbool;
            }
        }
        public static void DeleteRows(Workbook book, int n)
        {
            foreach (Microsoft.Office.Interop.Excel.Worksheet workSheet in book.Sheets)
            {
                Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A1", "A" + n);
                Microsoft.Office.Interop.Excel.Range row = range.EntireRow;
                row.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
            }
        }
        public static string ToMonthName(this DateTime dateTime)
        {
            return CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(dateTime.Month);
        }
    }
}
