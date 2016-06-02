using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Mail;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Text.RegularExpressions;
using System.Web;
using ConsoleApplication1.RSExecution2005;
using System.Web.Services.Protocols;
using System.DirectoryServices.AccountManagement;
using System.IO.Compression;
using System.DirectoryServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            //InsertUsingStoredProcedure();
            //ExtractPDFPath();
            //ParsingAndCasting();
            //RegistrationStub();
            //StringManipulation();
            //GetCurrentTimeZone();
            //StringSplitAndPrint();
            //StoreFilesIntoDB();
            //ReadFileFromDB();
            //SendEmail();
            //CreateiTextSharpPDF();
            //SendSSRSPDFEmail();
            //printEmailList();
            ActiveDirectory();

            //ArchiveLogFiles();

            Console.ReadLine();
        }


        private static void ArchiveLogFiles()
        {
            const string zipFilePath = @"C:\work\Temp\testwork\Err";

            string websiteSuffix = "_CEDA.iMIS.WebSite.txt";
            string adminPanelSuffix = "_Evocate.EvoCMS.AdminPanel.txt";
            string webserviceSuffix = "_Evocate.EvoCMS.Service.txt";

            using (FileStream zipFileToOpen = new FileStream(zipFilePath, FileMode.Open))
            using (ZipArchive archive = new ZipArchive(zipFileToOpen, ZipArchiveMode.Update))
            {
                //ZipArchiveEntry readMeEntry = archive.CreateEntry()

                foreach (var zipArchiveEntry in archive.Entries)
                {
                    Console.WriteLine("Fullname of the zip archive entry: {0}", zipArchiveEntry.FullName);
                }
            }

        }

        private static void ActiveDirectory()
        {
            //using(PrincipalContext pc = new PrincipalContext(ContextType.Domain, "CEDA"))
            //{
            //    bool isValid = pc.ValidateCredentials("preetham", "adfdsas");
            //    UserPrincipal user = UserPrincipal.Current;
            //    if(user != null)
            //    {
            //        Console.WriteLine(user.LastBadPasswordAttempt);
            //    }
            //}

            Console.WriteLine(System.Net.Dns.GetHostName());
            Console.WriteLine(System.Environment.MachineName);

            // Create and write to Excel sheet
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;


            using (var context = new PrincipalContext(ContextType.Domain, "ad.ceda.com.au"))
            {
                using (var searcher = new PrincipalSearcher(new UserPrincipal(context)))
                {
                    //Start Excel and get Application object.
                    oXL = new Microsoft.Office.Interop.Excel.Application();
                    oXL.Visible = true;

                    //Get a new workbook.
                    oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                    oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                    //Add table headers going cell by cell.
                    oSheet.Cells[1, 1] = "First Name";
                    oSheet.Cells[1, 2] = "Last Name";
                    oSheet.Cells[1, 3] = "SAM account name";
                    oSheet.Cells[1, 4] = "User principal name";
                    oSheet.Cells[1, 5] = "Title";
                    oSheet.Cells[1, 6] = "Location";

                    //Format A1:D1 as bold, vertical alignment = center.
                    oSheet.get_Range("A1", "F1").Font.Bold = true;
                    oSheet.get_Range("A1", "F1").VerticalAlignment =
                        Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;



                    int counter = 2;

                    foreach (var result in searcher.FindAll())
                    {
                        DirectoryEntry de = result.GetUnderlyingObject() as DirectoryEntry;

                        if (de.Properties["givenName"].Value == null || de.Properties["sn"].Value == null)
                            continue;

                        if (de.Properties["sn"].Value != null)
                        {
                            if (de.NativeGuid != null)
                            {
                                int flags = (int)de.Properties["userAccountControl"].Value;
                                if (Convert.ToBoolean(flags & 0x0002))
                                    continue;
                            }
                        }

                        Console.WriteLine("First Name: " + de.Properties["givenName"].Value);
                        oSheet.get_Range("A" + counter).Value2 = de.Properties["givenName"].Value;
                        Console.WriteLine("Last Name : " + de.Properties["sn"].Value);
                        
                        oSheet.get_Range("B" + counter).Value2 = de.Properties["sn"].Value;
                        Console.WriteLine("SAM account name   : " + de.Properties["samAccountName"].Value);
                        oSheet.get_Range("C" + counter).Value2 = de.Properties["samAccountName"].Value;
                        Console.WriteLine("User principal name: " + de.Properties["userPrincipalName"].Value);
                        oSheet.get_Range("D" + counter).Value2 = de.Properties["userPrincipalName"].Value;
                        Console.WriteLine("Title: " + de.Properties["title"].Value);
                        oSheet.get_Range("E" + counter).Value2 = de.Properties["title"].Value;
                        Console.WriteLine("Location: " + de.Properties["l"].Value);
                        oSheet.get_Range("F" + counter).Value2 = de.Properties["l"].Value;
                        Console.WriteLine();
                        counter++;
                    }

                    //AutoFit columns A:D.
                    oRng = oSheet.get_Range("A1", "F1");
                    oRng.EntireColumn.AutoFit();

                    oXL.Visible = false;
                    oXL.UserControl = false;
                    oWB.SaveAs("c:\\work\\test506.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    oWB.Close();
                }
            }
        }

        private static void StringSplitAndPrint()
        {
            String informzMailings = ConfigurationManager.AppSettings["Informz.Mailings"].ToString();
            var strArray = informzMailings.Split(',');

            foreach (var str in strArray)
            {
                Console.WriteLine(str);
            }
        }

        private static void printEmailList()
        {
            Console.WriteLine("Event confirmation emails");

            string emailList = ConfigurationManager.AppSettings["BCCList"].ToString();
            string[] emails = emailList.Split(';');

            foreach (var email in emails)
            {
                Console.WriteLine(email);
            }

            Console.ReadLine();
        }

        private static void SendSSRSPDFEmail()
        {
            String connString = ConfigurationManager.ConnectionStrings["LocalDB"].ToString();
            SqlConnection sqlConnection = new SqlConnection(connString);

            using (SqlCommand cmd = new SqlCommand("select top 5 * from MailEventInvoice where sentflag = @flag", sqlConnection))
            {
                cmd.Parameters.AddWithValue("flag", false);
                sqlConnection.Open();
                string contactId = "", eventCode = "", coordinatorId = "";

                using (SqlDataReader reader = cmd.ExecuteReader(System.Data.CommandBehavior.Default))
                {
                    while (reader.Read())
                    {
                        contactId = reader["BT_ID"].ToString();
                        eventCode = reader["EventCode"].ToString();
                        coordinatorId = reader["EventCoordinatorId"].ToString();
                    }
                }

                ReportExecutionService rs = new ReportExecutionService();
                rs.Credentials = new NetworkCredential("WEBSSRS", "#Rep0rt%e3", "CEDA");
                rs.Url = "http://vicsqldev/reportserver/ReportExecution2005.asmx";

                // Render arguments
                byte[] result = null;
                string reportPath = "/CEDA Meeting Invoice";
                string format = "PDF";
                string historyID = null;
                string devInfo = @"<DeviceInfo><Toolbar>False</Toolbar></DeviceInfo>";

                // Prepare report parameter.
                RSExecution2005.ParameterValue[] parameters = new RSExecution2005.ParameterValue[2];
                parameters[0] = new RSExecution2005.ParameterValue();
                parameters[0].Name = "ContactId";
                parameters[0].Value = contactId;
                parameters[1] = new RSExecution2005.ParameterValue();
                parameters[1].Name = "meeting";
                parameters[1].Value = eventCode;

                RSExecution2005.DataSourceCredentials[] credentials = new RSExecution2005.DataSourceCredentials[3];
                credentials[0] = new RSExecution2005.DataSourceCredentials();
                credentials[0].UserName = "imis20_reports";
                credentials[1] = new RSExecution2005.DataSourceCredentials();
                credentials[1].Password = "!iMis99";
                credentials[2] = new RSExecution2005.DataSourceCredentials();
                credentials[2].DataSourceName = "iMISDevReports";


                string encoding;
                string mimeType;
                string extension;
                RSExecution2005.Warning[] warnings = null;
                string[] streamIDs = null;

                ExecutionInfo execInfo = new ExecutionInfo();
                ExecutionHeader execHeader = new ExecutionHeader();

                rs.ExecutionHeaderValue = execHeader;

                execInfo = rs.LoadReport(reportPath, historyID);

                rs.SetExecutionParameters(parameters, "en-us");
                String SessionId = rs.ExecutionHeaderValue.ExecutionID;

                Console.WriteLine("SessionID: {0}", rs.ExecutionHeaderValue.ExecutionID);


                try
                {
                    //result = rs.Render(format, devInfo, out extension, out encoding, out mimeType, out warnings, out streamIDs);
                    result = rs.Render(format, devInfo, out extension, out mimeType, out encoding, out warnings, out streamIDs);

                    execInfo = rs.GetExecutionInfo();

                    Console.WriteLine("Execution date and time: {0}", execInfo.ExecutionDateTime);


                }
                catch (SoapException e)
                {
                    Console.WriteLine(e.Detail.OuterXml);
                }
                // Write the contents of the report to an MHTML file.
                try
                {
                    FileStream stream = new FileStream(@"C:\Work\Temp\testpdf.pdf", FileMode.Create);
                    //FileStream stream = File.Create("invoice.pdf", result.Length);
                    Console.WriteLine("File created.");
                    stream.Write(result, 0, result.Length);
                    Console.WriteLine("Result written to the file.");
                    stream.Close();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }

                OutlookCalendarAppointment(eventCode, coordinatorId);
            }

        }


        private static void OutlookCalendarAppointment(string eventCode, string coordinatorId)
        {
            try
            {
                // Check timezone for outlook calendar settings

                string eventTZ = "AUS Eastern Standard Time";
                string eventCodeState = eventCode.ToLower();
                eventCodeState = eventCodeState.Substring(0, 1);

                switch (eventCodeState)
                {
                    case "s": // SA
                        eventTZ = "Cen. Australia Standard Time";
                        break;
                    case "w": // WA
                        eventTZ = "W. Australia Standard Time";
                        break;
                    default:
                        eventTZ = "AUS Eastern Standard Time";
                        break;
                }


                // Local time zone info. to be used to enable correct calendar entry for people registering from any TZ (in / outside AUS)
                TimeZoneInfo localTimeZone = TimeZoneInfo.FindSystemTimeZoneById(TimeZone.CurrentTimeZone.StandardName);
                TimeZoneInfo eventTimeZone = TimeZoneInfo.FindSystemTimeZoneById(eventTZ);

                String webPageUrl = "www.ceda.com.au";
                String sData = File.ReadAllText(@"C:\inetpub\wwwroot\EvoCMS.Live\Admin\FOLDERS\Service\Templates\OutlookCalendarAppointment.ics");

                using (SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["DataSource.iMIS.Connection"].ToString()))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("select meeting, title, description, begin_date, end_date, address_1, address_2, address_3, city, state_province, zip, country, notes, contact_id from meet_master where meeting = @meeting", cn))
                    {
                        cmd.Parameters.AddWithValue("meeting", eventCode);

                        using (SqlDataReader reader = cmd.ExecuteReader(System.Data.CommandBehavior.Default))
                        {
                            while (reader.Read())
                            {
                                // Parse the string data.

                                string venue = (string.IsNullOrEmpty(reader["address_1"].ToString()) ? "" : reader["address_1"].ToString())
                                    + (string.IsNullOrEmpty(reader["address_2"].ToString()) ? "" : (", " + reader["address_2"].ToString()))
                                    + (string.IsNullOrEmpty(reader["address_3"].ToString()) ? "" : (", " + reader["address_3"].ToString()))
                                    + (string.IsNullOrEmpty(reader["city"].ToString()) ? "" : (", " + reader["city"].ToString()))
                                    + (string.IsNullOrEmpty(reader["state_province"].ToString()) ? "" : (", " + reader["state_province"].ToString()))
                                    + (string.IsNullOrEmpty(reader["zip"].ToString()) ? "" : (", " + reader["zip"].ToString()))
                                    + (string.IsNullOrEmpty(reader["country"].ToString()) ? "" : (", " + reader["country"].ToString()));

                                DateTime start = new DateTime();
                                DateTime end = new DateTime();

                                var startDate = DateTime.TryParse(reader["begin_date"].ToString(), out start) ? DateTime.Parse(reader["begin_date"].ToString()) : DateTime.Now;
                                var endDate = DateTime.TryParse(reader["end_date"].ToString(), out end) ? DateTime.Parse(reader["end_date"].ToString()) : DateTime.Now;

                                DateTime eventStartTime = TimeZoneInfo.ConvertTimeToUtc(startDate, eventTimeZone);
                                DateTime eventEndTime = TimeZoneInfo.ConvertTimeToUtc(endDate, eventTimeZone);
                                String EventRelativeUrl = String.Format("{0}/{1}/{2}?EventCode={3}",
                                    "/event/eventdetails",
                                    endDate.ToString("yyyy/M"),
                                    eventCode.ToLower(),
                                    eventCode);

                                sData = sData.Replace("__SUBJECT__", "CEDA event | " + reader["title"].ToString());
                                sData = sData.Replace("__BODY__", reader["description"].ToString());
                                sData = sData.Replace("__TITLE__", reader["title"].ToString());
                                sData = sData.Replace("__DESCRIPTION__", reader["description"].ToString());
                                sData = sData.Replace("__EVENTURL__", webPageUrl + "//" + EventRelativeUrl);
                                sData = sData.Replace("__POLICY__", reader["notes"].ToString());
                                sData = sData.Replace("__START_DATE__", eventStartTime.ToString("yyyyMMddTHHmm00"));
                                sData = sData.Replace("__END_DATE__", eventEndTime.ToString("yyyyMMddTHHmm00"));
                                sData = sData.Replace("__LOCATION__", venue.Replace("\n", " ").Replace(",", " ").Replace("&nbsp;", " "));
                                sData = sData.Replace("__GUID__", Guid.NewGuid().ToString());

                            }
                        }
                    }
                }

                using (SqlConnection cn2 = new SqlConnection(ConfigurationManager.ConnectionStrings["DataSource.iMIS.Connection"].ToString()))
                {
                    cn2.Open();

                    using (SqlCommand cmd2 = new SqlCommand("select full_name, work_phone, email from name where id = @id", cn2))
                    {
                        cmd2.Parameters.AddWithValue("id", coordinatorId);

                        using (SqlDataReader reader = cmd2.ExecuteReader(System.Data.CommandBehavior.Default))
                        {
                            while (reader.Read())
                            {
                                // Cater for events coordinator
                                Boolean hasAnEventCoordinator = string.IsNullOrEmpty(coordinatorId) ? false : true;
                                if (hasAnEventCoordinator)
                                {
                                    // Get coordinator contact details
                                    sData = sData.Replace("__COORDINATORNAME__", reader["full_name"].ToString());
                                    sData = sData.Replace("__COORDINATORPHONE__", reader["work_phone"].ToString());
                                    sData = sData.Replace("__COORDINATOREMAIL__", reader["email"].ToString());
                                }
                                else
                                {
                                    sData = sData.Replace("__COORDINATORNAME__", "CEDA");
                                    sData = sData.Replace("__COORDINATORPHONE__", "03 9662 3544");
                                    sData = sData.Replace("__COORDINATOREMAIL__", "info@ceda.com.au");
                                }
                            }
                        }
                    }
                }

                // Create an attachment from a stream.
                MemoryStream objMS = new MemoryStream(new UTF8Encoding(true).GetBytes(sData));

                FileStream stream = new FileStream(@"C:\Work\Temp\cal.ics", FileMode.Create);
                //FileStream stream = File.Create("invoice.pdf", result.Length);
                Console.WriteLine("File created.");
                stream.Write(new UTF8Encoding(true).GetBytes(sData), 0, new UTF8Encoding(true).GetBytes(sData).Length);
                Console.WriteLine("Result written to the file.");
                stream.Close();

                //objAttachment = new Attachment(objMS, "OutlookCalendarAppointment.ics", "text/calendar; method=REQUEST");
            }
            finally
            {

            }
        }

        private static void CreateiTextSharpPDF()
        {
            Console.WriteLine("Creating iTextSharp PDF document...");

            System.IO.FileStream fs = new FileStream(@"C:\work\Temp\Test1.pdf", FileMode.Create);
            Document document = new Document(PageSize.A4, 25, 25, 30, 30);
            PdfWriter writer = PdfWriter.GetInstance(document, fs);

            document.AddAuthor("Preetham KH - CEDA");
            document.AddCreator("iTextSharp");
            document.AddKeywords("invoice receipt event");
            document.AddSubject("Event Invoice/Receipt");
            document.AddTitle("CEDA Event Invoice");

            String connString = ConfigurationManager.ConnectionStrings["LocalDB"].ToString();
            SqlConnection sqlConnection = new SqlConnection(connString);
            sqlConnection.Open();

            FileStream fStream = File.OpenRead("C:\\Work\\_2331567_Invoice.pdf");
            byte[] contents = new byte[fStream.Length];
            fStream.Read(contents, 0, (int)fStream.Length);
            fStream.Close();

            document.Open();
            // Add a simple and wellknown phrase to the document in a flow layout manner
            document.Add(new Paragraph(System.Text.Encoding.UTF8.GetString(contents)));
            // Close the document
            document.Close();
            // Close the writer instance
            writer.Close();
            // Always close open filehandles explicity
            fs.Close();
        }

        private static void SendEmail()
        {
            //Console.WriteLine("Starting sendmail functionality");

            try
            {
                string orderNumbers = "";

                using (SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["LocalDBCEDA"].ToString()))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("select top 1 * from invoicepdfpath where sentflag = @flag", cn))
                    {
                        cmd.Parameters.AddWithValue("flag", false);
                        using (SqlDataReader reader = cmd.ExecuteReader(System.Data.CommandBehavior.Default))
                        {
                            while (reader.Read())
                            {
                                String commonName = reader[3].ToString().Substring(reader[3].ToString().IndexOf('_'), reader[3].ToString().LastIndexOf('_') - (reader[3].ToString().IndexOf('_') - 1));

                                FileStream pdfStream = File.OpenRead(@"C:\inetpub\wwwroot\EvoCMS.Live\Admin\FOLDERS\Service\Files\Invoices\" + commonName + "Invoice.pdf");
                                FileStream calStream = File.OpenRead(@"C:\inetpub\wwwroot\EvoCMS.Live\Admin\FOLDERS\Service\Files\Invoices\" + commonName + "OutlookCalendarAppointment.ics");

                                Attachment pdfAttachment = new Attachment(pdfStream, "Invoice.pdf");
                                Attachment calendarAttachment = new Attachment(calStream, "OutlookCalendarAppointment.ics", "text/calendar; method=REQUEST");

                                using (SmtpClient client = new SmtpClient())
                                {
                                    Int32 iPort = 0;
                                    Int32.TryParse(ConfigurationManager.AppSettings["Email.Port"].ToString(), out iPort);
                                    if (iPort > 0)
                                        client.Port = iPort;

                                    client.Host = ConfigurationManager.AppSettings["Email.Host"].ToString();
                                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                                    client.EnableSsl = Boolean.Parse(ConfigurationManager.AppSettings["Email.EnableSSL"].ToString());
                                    //client.Port = Int32.Parse(ConfigurationManager.AppSettings["Email.Port"].ToString());

                                    String username = ConfigurationManager.AppSettings["Email.Username"].ToString();
                                    String password = ConfigurationManager.AppSettings["Email.Password"].ToString();

                                    if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
                                    {
                                        client.UseDefaultCredentials = false;
                                        client.Credentials = new NetworkCredential(username, password);
                                    }
                                    else
                                        client.UseDefaultCredentials = true;

                                    using (MailMessage message = new MailMessage())
                                    {
                                        //Recipient list
                                        String[] toList = reader[7].ToString().Split(';');
                                        String[] ccList = reader[8].ToString().Split(';');
                                        String[] bccList = reader[9].ToString().Split(';');
                                        message.From = new MailAddress("noreply@ceda.com.au");

                                        if (toList.Length > 0)
                                        {
                                            foreach (String to in toList)
                                            {
                                                message.To.Add(to);
                                            }
                                        }

                                        if (bccList.Length > 0)
                                        {
                                            foreach (String bcc in bccList)
                                            {
                                                message.Bcc.Add(bcc);
                                            }
                                        }
                                        message.Subject = reader[12].ToString();
                                        message.IsBodyHtml = true;
                                        message.BodyEncoding = Encoding.GetEncoding(1254);
                                        message.Attachments.Add(pdfAttachment);
                                        message.Attachments.Add(calendarAttachment);
                                        message.Body = reader[11].ToString();

                                        client.Send(message);

                                        // Mail sent if no exception is thrown
                                        if (string.IsNullOrEmpty(orderNumbers))
                                        {
                                            orderNumbers = reader[1].ToString();
                                        }
                                        else
                                            orderNumbers += "," + reader[1].ToString();

                                        //Console.WriteLine("Message sent...");
                                    }
                                }
                            }
                        }
                    }
                }

                // UPDATE Flag field in the InvoicePDFPath table to 1 after sending email
                // UPDATE InvoicePDFPath SET sentFlag = '1' where ID in ()

                using (SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["LocalDBCEDA"].ToString()))
                {
                    cn.Open();
                    string[] recordIds = orderNumbers.Split(',');

                    foreach (string record in recordIds)
                    {
                        using (SqlCommand cmd = new SqlCommand("update invoicepdfpath set sentflag = @sentflag, emailsenttime = @emailsenttime where order_number = '" + record + "'", cn))
                        {
                            //cmd.Parameters.AddWithValue("ordernumbers", record);
                            cmd.Parameters.AddWithValue("emailsenttime", DateTime.Now);
                            cmd.Parameters.AddWithValue("sentflag", true);
                            Console.WriteLine(cmd.ToString());
                            var result = cmd.ExecuteNonQuery();

                            //Console.WriteLine("Updated sentflags in the database");
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private static void ReadFileFromDB()
        {
            string ToSaveFileTo = "C:\\work\\Temp\\Test.pdf";

            using (SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["LocalDB"].ToString()))
            {
                cn.Open();
                using (SqlCommand cmd = new SqlCommand("select PDFFile from storeFiles  where ID='" + "1" + "' ", cn))
                {
                    using (SqlDataReader dr = cmd.ExecuteReader(System.Data.CommandBehavior.Default))
                    {
                        if (dr.Read())
                        {

                            byte[] fileData = (byte[])dr.GetValue(0);
                            using (System.IO.FileStream fs = new System.IO.FileStream(ToSaveFileTo, System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite))
                            {
                                using (System.IO.BinaryWriter bw = new System.IO.BinaryWriter(fs))
                                {
                                    bw.Write(fileData);
                                    bw.Close();
                                }
                            }
                        }

                        dr.Close();
                    }
                }
            }

        }

        private static void StoreFilesIntoDB()
        {
            Console.WriteLine("Storing file into database");

            String connString = ConfigurationManager.ConnectionStrings["LocalDB"].ToString();
            SqlConnection sqlConnection = new SqlConnection(connString);
            sqlConnection.Open();

            FileStream fStream = File.OpenRead("C:\\Work\\_2331567_Invoice.pdf");
            byte[] contents = new byte[fStream.Length];
            fStream.Read(contents, 0, (int)fStream.Length);
            fStream.Close();
            using (SqlCommand cmd = new SqlCommand("insert into storeFiles " + "(PDFFile)values(@data)", sqlConnection))
            {
                cmd.Parameters.Add("@data", contents);
                cmd.ExecuteNonQuery();
                Console.WriteLine("Pdf File Saved into the database");
            }
        }

        private static void GetCurrentTimeZone()
        {
            Console.WriteLine("Getting current timezone");

            TimeZone localZone = TimeZone.CurrentTimeZone;
            DateTime currentDate = DateTime.Now;

            Console.WriteLine("LocalZone Standard name: " + localZone.StandardName);
            Console.WriteLine("LocalZone DaylightName  name: " + localZone.DaylightName);
            Console.WriteLine("LocalZone IsDaylightSavingTime name: " + localZone.IsDaylightSavingTime(currentDate));
            Console.WriteLine("LocalZone ToUniversalTime name: " + localZone.ToUniversalTime(currentDate));
            Console.WriteLine("LocalZone GetUtcOffset name: " + localZone.GetUtcOffset(currentDate));
            Console.WriteLine("LocalZone GetDaylightChanges name: " + localZone.GetDaylightChanges(currentDate.Year));
            Console.WriteLine("DateTime" + currentDate);

            string ausEast = "AUS Eastern Standard Time"; // Rest of Australia
            string ausWest = "W. Australia Standard Time"; // WA
            string ausCent = "Cen. Australia Standard Time"; // SA
            string ausQld = "E. Australia Standard Time"; // QLD

            TimeZoneInfo waZone = TimeZoneInfo.FindSystemTimeZoneById(ausWest);
            TimeZoneInfo saZone = TimeZoneInfo.FindSystemTimeZoneById(ausCent);
            TimeZoneInfo ausZone = TimeZoneInfo.FindSystemTimeZoneById(ausEast);
            TimeZoneInfo qldZone = TimeZoneInfo.FindSystemTimeZoneById(ausQld);

            Console.WriteLine("\n\nLocal Timezone: " + localZone.StandardName);
            //Console.WriteLine("Local Time AEST: " + TimeZoneInfo.ConvertTime(currentDate, ausZone).TimeOfDay);
            var localTime = string.Format("{0:hh}", currentDate) + string.Format("{0:mm}", currentDate) + string.Format("{0:ss}", "00");
            var waTime = TimeZoneInfo.ConvertTime(currentDate, waZone);
            var saTime = TimeZoneInfo.ConvertTime(currentDate, saZone);
            var qldTime = TimeZoneInfo.ConvertTime(currentDate, qldZone);

            Console.WriteLine("Local Time AEST: " + localTime);
            Console.WriteLine("Converted Time in WA: " + string.Format("{0:hh}", waTime) + string.Format("{0:mm}", waTime) + string.Format("{0:ss}", "00"));
            Console.WriteLine("Converted Time in SA: " + string.Format("{0:hh}", saTime) + string.Format("{0:mm}", saTime) + string.Format("{0:ss}", "00"));
            Console.WriteLine("Convertd Time in QLD: " + string.Format("{0:hh}", qldTime) + string.Format("{0:mm}", qldTime) + string.Format("{0:ss}", "00"));

            Console.WriteLine("\n\nLocal Date: " + currentDate.ToString("yyyyMMdd"));
            Console.WriteLine("WA Date: " + waTime.ToString("yyyyMMdd"));
            Console.WriteLine("SA Date: " + saTime.ToString("yyyyMMdd"));
            Console.WriteLine("QLD Date: " + qldTime.ToString("yyyyMMdd"));

            Console.WriteLine("\n\nFinal Local: " + currentDate.ToString("yyyyMMdd") + "T" + localTime);
            Console.WriteLine("Final WA: " + waTime.ToString("yyyyMMdd") + "T" + string.Format("{0:hh}", waTime) + string.Format("{0:mm}", waTime) + string.Format("{0:ss}", "00"));
            Console.WriteLine("Final SA: " + saTime.ToString("yyyyMMdd") + "T" + string.Format("{0:hh}", saTime) + string.Format("{0:mm}", saTime) + string.Format("{0:ss}", "00"));
            Console.WriteLine("Final QLD: " + qldTime.ToString("yyyyMMdd") + "T" + string.Format("{0:hh}", qldTime) + string.Format("{0:mm}", qldTime) + string.Format("{0:ss}", "00"));

            Console.WriteLine("**************");
            TimeZoneInfo testZone = TimeZoneInfo.FindSystemTimeZoneById(ausWest);
            var test = TimeZoneInfo.ConvertTimeToUtc(new DateTime(2015, 09, 04, 11, 45, 00), testZone);
            Console.WriteLine("TimeZone Offset: " + TimeZone.CurrentTimeZone.ToLocalTime(test));

            Console.ReadLine();
        }

        private static void ParsingAndCasting()
        {
            Console.WriteLine(Int32.Parse("194452.00"));
        }

        private static void ExtractPDFPath()
        {
            string currentPDFPath = "C:\\inetpub\\wwwroot\\EvoCMS.Live\\Admin\\FOLDERS\\Service\\Temp\\_3848120_Invoice.pdf";

            // strip from \\FOLDERS, replaces\\ with \
            int lengthCurrentPDFPath = currentPDFPath.Length;
            Console.WriteLine("Length of current pdf path: " + lengthCurrentPDFPath);
            var substring = currentPDFPath.Substring(58, lengthCurrentPDFPath - 58);
            string address = "http://adminpanel.ceda.com.au/folders/service/temp/" + substring;

            Console.WriteLine(address);
        }

        private static void InsertUsingStoredProcedure()
        {
            String connString = ConfigurationManager.ConnectionStrings["LocalDB"].ToString();
            //SqlConnection sqlConnection = new SqlConnection(connString);

            using (SqlConnection sqlConnection = new SqlConnection(connString))
            {
                SqlCommand sqlCommand = new SqlCommand("sp_SaveEventInvoicePDF", sqlConnection)
                {
                    CommandType = CommandType.StoredProcedure
                };

                sqlCommand.Parameters.AddWithValue("ordernumber", 12345);
                sqlCommand.Parameters.AddWithValue("InvoicePDFPath", @"C:\inetpub\wwwroot\EvoCMS.Live\Admin\FOLDERS\Service\Temp\_359826_Invoice.pdf");
                sqlCommand.Parameters.AddWithValue("timestamp", DateTime.Now);

                sqlConnection.Open();

                var result = sqlCommand.ExecuteNonQuery();

                if (result != 0)
                {
                    Console.WriteLine("Inserted row into the table");
                }
            }
        }


        // Method to retrieve data from stored proc
        private static void RetrieveDataFromStoredProc(SqlConnection sqlConnection)
        {
            SqlCommand sqlCommand = new SqlCommand("sp_GetRegistrantClass_Rego", sqlConnection)
            {
                CommandType = CommandType.StoredProcedure
            };
            sqlCommand.Parameters.AddWithValue("@ContactID", 100000);

            var outParameter1 = new SqlParameter
            {
                ParameterName = "@MemberType",
                SqlDbType = SqlDbType.VarChar,
                Direction = ParameterDirection.Output,
                Size = 15

            };
            sqlCommand.Parameters.Add(outParameter1);

            var outParameter2 = new SqlParameter
            {
                ParameterName = "@ISMember",
                SqlDbType = SqlDbType.Bit,
                Direction = ParameterDirection.Output
            };
            sqlCommand.Parameters.Add(outParameter2);

            sqlConnection.Open();

            sqlCommand.ExecuteNonQuery();

            Console.WriteLine(outParameter1.Value.ToString());
            Console.WriteLine(outParameter2.Value.ToString());

            using (SqlDataReader reader = sqlCommand.ExecuteReader())
            {
                while (reader.Read())
                {
                    Console.WriteLine(reader[0]);
                }
            }
        }



        // Method to retrieve data from views
        private static void RetrieveDataFromView(SqlConnection sqlConnection)
        {
            SqlCommand sqlCommand = new SqlCommand("select * from TestView", sqlConnection);
            sqlConnection.Open();
            using (SqlDataReader reader = sqlCommand.ExecuteReader())
            {
                while (reader.Read())
                {
                    Console.WriteLine(reader[0] + ":" + reader[1] + ":" + reader[2]);
                }
            }
        }



        // Some string manipulation try-outs
        private static void StringManipulation()
        {
            String path = "//adminpanel.ceda.com.au/folders/service/temp/_100920151544399_Invoice.pdf";

            Console.WriteLine(path.Substring(0, path.IndexOf('_')));
            Console.WriteLine(path.Substring(path.IndexOf('_'), path.LastIndexOf('_') - path.IndexOf('_') + 1));

            StringBuilder sb = new StringBuilder();
            sb.Append("test");

            Console.WriteLine(sb);

            Console.WriteLine("------- Regex --------");
            //string eventCode = "N150922";
            string eventCode2 = "N150922b";

            string result = Regex.Match(eventCode2, @"\d+").Value;
            string year = string.Concat("20", result.Substring(0, 2));
            string month = string.Concat(result.Substring(2, 2));
            Console.WriteLine("/event/eventdetails/" + year + "/" + month + "/" + eventCode2 + "?EventCode=" + eventCode2);
            //Console.WriteLine("Digit Only string: " + result);

            Console.ReadLine();

            //var evt = "n150922b-234234";
            //Console.WriteLine(evt.IndexOf('-'));
            //Console.WriteLine(evt.Substring(evt.IndexOf('-') + 1));

            //var stringEnd = "com.au";
            ////Tests
            //var test = "http://www.ceda.com.au/abcde";
            //var finalURL = "";
            //var testlength = test.Length;

            //if (test.Contains("http") && test.Contains("com.au"))
            //{
            //    var index = test.IndexOf("com.au");
            //    var tester = test.Substring(index + 6, test.Length - (index + 6));
            //    Console.WriteLine(index);
            //    Console.WriteLine(index + 6);
            //    Console.WriteLine(index + 1);
            //    Console.WriteLine(testlength);
            //    finalURL = "http://www.staging.ceda.com.au" + tester;
            //    Console.WriteLine(finalURL);
            //}


            //if (!test.Contains("http"))
            //{
            //    Console.WriteLine("This doesnt get printed");
            //}

            //if (sHref.Contains("http") && sHref.Contains("com.au"))
            //{
            //    var index = sHref.IndexOf("com.au", StringComparison.Ordinal);
            //    var finalHref = "http://staging.ceda.com.au" +
            //                    sHref.Substring(index + 6, sHref.Length - (index + 6));

            //    sHref = finalHref;
            //}
        }


        private static void RegistrationStub()
        {
            //Console.WriteLine("Testing the CRegistration Class");

            ////CContact cContact = new CContact();

            ////var username = "manager";
            ////var password = "1960adec";
            //var eventID = "V150825";
            //var registrantClass = "ME";
            //var eventCode = "100004";

            //var connectionString = "server=vicsqldev;database=iMIS;uid=Web_Imis;pwd=x!@ctly9;";

            ////CStaffUser cStaff = new CStaffUser();
            ////cStaff.UserName = username;
            ////cStaff.cl

            //CRegistration cRegistration = new CRegistration(Admin, eventID, eventCode, false, false)
            //{
            //    RegistrantClass =  registrantClass
            //};

            ////setting the member type
            //cRegistration.SourceCode = "INTERNET";

        }
    }
}