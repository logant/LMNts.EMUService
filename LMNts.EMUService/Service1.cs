using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.DataVisualization.Charting;
using System.Xml;
using Opto22.Emu;
using System.Net.Mime;

namespace LMNts.EMUService
{
    public partial class Service1 : ServiceBase
    {
        // Global Variables
        string ip4W;
        string ip4E;
        string ip5W;
        string ip5E;
        string vizURL;
        string unsubscribe;
        string dbServer;
        string dbName;
        string dbTableName;
        string dbUser;
        string dbPassword;
        string mailClient;
        int prevYears;
        int optoMMPPort;
        int commTimeout;
        int scanTime;
        bool ipv6;
        double warningEmail = 1;

        LinkedResource chartEMU0Img;
        LinkedResource chartEMU1Img;
        LinkedResource chartEMU2Img;
        LinkedResource chartEMU3Img;

        string typHTML = "<td style=\"font-size:14px;\">#VALUE#</td>";
        string colHTML = "<col width=\"200\"/>";
        string totHTML = "<td style=\"font-size:14px; font-weight:bold\">#VALUE# kW</td>";

        System.Timers.Timer timerUpdateData;

        Opto22Emu emu0 = null;
        Opto22Emu emu1 = null;
        Opto22Emu emu2 = null;
        Opto22Emu emu3 = null;
        SqlConnection connection;
        Int32 mailSent = 0;

        public Service1()
        {
            InitializeComponent();

            // create the event log to track problems with the service
            if (!EventLog.SourceExists("LMNtsEMUSource"))
            {
                EventLog.CreateEventSource("LMNtsEMUSource", "EMULog");
            }
            eventLog1.Source = "LMNtsEMUSource";
            eventLog1.Log = "EMULog";

            // Settings are stored in an XML file to allow easy changes.
            // Read the XML file and update the settings as necessary
            XmlDocument xmlDoc = new XmlDocument();
            try
            {
                FileInfo fi = new FileInfo(typeof(Service1).Assembly.Location);
                string xmlPath = fi.Directory + "\\Settings.xml";
                XmlTextReader reader = new XmlTextReader(xmlPath);
                xmlDoc.Load(reader);
                if (xmlDoc != null)
                {
                    XmlNodeList settingsNodes = xmlDoc.SelectNodes("EMU/Setting");
                    foreach (XmlNode setting in settingsNodes)
                    {
                        switch (setting.Attributes["Name"].Value)
                        {
                            case "VizURL":
                                Properties.Settings.Default.vizURL = setting.InnerText;
                                Properties.Settings.Default.Save();
                                break;
                            case "EMUIP0":
                                Properties.Settings.Default.emuIP0 = setting.InnerText;
                                Properties.Settings.Default.Save();
                                break;
                            case "EMUIP1":
                                Properties.Settings.Default.emuIP1 = setting.InnerText;
                                Properties.Settings.Default.Save();
                                break;
                            case "EMUIP2":
                                Properties.Settings.Default.emuIP2 = setting.InnerText;
                                Properties.Settings.Default.Save();
                                break;
                            case "EMUIP3":
                                Properties.Settings.Default.emuIP3 = setting.InnerText;
                                Properties.Settings.Default.Save();
                                break;
                            case "Unsubscribe":
                                Properties.Settings.Default.unsubscribe = setting.InnerText;
                                Properties.Settings.Default.Save();
                                break;
                            case "DBServer":
                                Properties.Settings.Default.dbServer = setting.InnerText;
                                Properties.Settings.Default.Save();
                                break;
                            case "DBName":
                                Properties.Settings.Default.dbName = setting.InnerText;
                                Properties.Settings.Default.Save();
                                break;
                            case "DBTableName":
                                Properties.Settings.Default.dbTableName = setting.InnerText;
                                Properties.Settings.Default.Save();
                                break;
                            case "DBUser":
                                Properties.Settings.Default.dbUser = setting.InnerText;
                                Properties.Settings.Default.Save();
                                break;
                            case "DBPassword":
                                Properties.Settings.Default.dbPassword = setting.InnerText;
                                Properties.Settings.Default.Save();
                                break;
                            case "MailClient":
                                Properties.Settings.Default.mailClient = setting.InnerText;
                                Properties.Settings.Default.Save();
                                break;
                            case "OptoMMPPort":
                                Properties.Settings.Default.optoMMPPort = Int32.Parse(setting.InnerText);
                                Properties.Settings.Default.Save();
                                break;
                            case "Timeout":
                                Properties.Settings.Default.commTimeout = Int32.Parse(setting.InnerText);
                                Properties.Settings.Default.Save();
                                break;
                            case "ScanTime":
                                Properties.Settings.Default.scanTime = Int32.Parse(setting.InnerText);
                                Properties.Settings.Default.Save();
                                break;
                            case "IPV6":
                                Properties.Settings.Default.ipv6 = bool.Parse(setting.InnerText);
                                Properties.Settings.Default.Save();
                                break;
                            case "PreviousYears":
                                try
                                {
                                    Properties.Settings.Default.prevYears = int.Parse(setting.InnerText);
                                }
                                catch
                                {
                                    Properties.Settings.Default.prevYears = 2;
                                }
                                Properties.Settings.Default.Save();
                                break;
                            default:
                                break;
                        }
                    }
                    reader.Close();
                }
                else
                {
                    string errorMsg = DateTime.Now + " - Could not find settings file\n";
                    eventLog1.WriteEntry(errorMsg);
                }

                // Apply setting values to variables
                ip4E = Properties.Settings.Default.emuIP0;
                ip4W = Properties.Settings.Default.emuIP1;
                ip5E = Properties.Settings.Default.emuIP2;
                ip5W = Properties.Settings.Default.emuIP3;
                vizURL = Properties.Settings.Default.vizURL;
                unsubscribe = Properties.Settings.Default.unsubscribe;
                dbServer = Properties.Settings.Default.dbServer;
                dbName = Properties.Settings.Default.dbName;
                dbTableName = Properties.Settings.Default.dbTableName;
                dbUser = Properties.Settings.Default.dbUser;
                dbPassword = Properties.Settings.Default.dbPassword;
                mailClient = Properties.Settings.Default.mailClient;
                optoMMPPort = Properties.Settings.Default.optoMMPPort;
                commTimeout = Properties.Settings.Default.commTimeout;
                scanTime = Properties.Settings.Default.scanTime;
                ipv6 = Properties.Settings.Default.ipv6;
                prevYears = Properties.Settings.Default.prevYears;

            }
            catch (Exception ex)
            {
                string errorMsg = DateTime.Now + "\n - " + ex.Message + "\n\n";
                eventLog1.WriteEntry(errorMsg);
            }
        }

        protected override void OnStart(string[] args)
        {
            eventLog1.WriteEntry("Service Started");

            try
            {
                //this chunk of code was in the buttonConnect_Click method...
                if ((emu0 != null) && (emu1 != null) && (emu2 != null) && (emu3 != null))
                {
                    timerUpdateData.Stop();
                    emu0.Stop();
                    emu1.Stop();
                    emu2.Stop();
                    emu3.Stop();
                    emu0 = null;
                    emu1 = null;
                    emu2 = null;
                    emu3 = null;
                }
                emu0 = new Opto22Emu(ip4W, optoMMPPort, ipv6);
                emu1 = new Opto22Emu(ip4E, optoMMPPort, ipv6);
                emu2 = new Opto22Emu(ip5W, optoMMPPort, ipv6);
                emu3 = new Opto22Emu(ip5E, optoMMPPort, ipv6);
                SQLConnect();

                // start the timer event
                timerUpdateData = new System.Timers.Timer(scanTime);
                timerUpdateData.Elapsed += new System.Timers.ElapsedEventHandler(timerUpdateData_Tick_1);
                timerUpdateData.Enabled = true;
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry(ex.Message);
            }
        }

        protected override void OnStop()
        {
            eventLog1.WriteEntry("Service Stopped");
        }

        private void SQLConnect()
        {
            //eventLog1.WriteEntry("In the SQLConnect method");
            // First trash the connection if it already exists and has failed
            if (connection != null)
            {
                try
                {
                    connection.Close();
                    connection.Dispose();
                    connection = null;
                }
                catch
                {
                    try
                    {
                        connection.Dispose();
                        connection = null;
                    }
                    catch { }
                }
            }

            // Create the connection using values extracted from the XML settings file.
            string connStr = string.Format("server={0};user={1};database={2};password={3};", dbServer, dbUser, dbName, dbPassword);
            connection = new SqlConnection(connStr);
        }

        private void timerUpdateData_Tick_1(object sender, EventArgs e)
        {
            // Get the total true power value for this point in time.
            float powerEMU0 = emu0.CurrentValue(Opto22Emu.eInstantaneousF32Values.v1_CT_Total_TruePower);
            float powerEMU1 = emu1.CurrentValue(Opto22Emu.eInstantaneousF32Values.v1_CT_Total_TruePower);
            float powerEMU2 = emu2.CurrentValue(Opto22Emu.eInstantaneousF32Values.v1_CT_Total_TruePower);
            float powerEMU3 = emu3.CurrentValue(Opto22Emu.eInstantaneousF32Values.v1_CT_Total_TruePower);
            string now = DateTime.Now.ToString("yyyy-MM-dd H:mm:ss");

            // Write the total true power data to the SQL Database
            string insertCmd;
            SqlCommand InsertCommand;
            try
            {
                connection.Open();

                insertCmd = "INSERT INTO " + dbTableName + " (Location, Timestamp, Power)" +
                                   " VALUES ('EMU0','" + now + "', " + powerEMU0 + ");";
                InsertCommand = connection.CreateCommand();
                InsertCommand.CommandText = insertCmd;
                InsertCommand.ExecuteNonQuery();

                insertCmd = "INSERT INTO " + dbTableName + " (Location, Timestamp, Power)" +
                            " VALUES ('EMU1','" + now + "'," + powerEMU1 + ");";
                InsertCommand = connection.CreateCommand();
                InsertCommand.CommandText = insertCmd;
                InsertCommand.ExecuteNonQuery();

                insertCmd = "INSERT INTO " + dbTableName + " (Location, Timestamp, Power)" +
                            " VALUES ('EMU2','" + now + "'," + powerEMU2 + ");";
                InsertCommand = connection.CreateCommand();
                InsertCommand.CommandText = insertCmd;
                InsertCommand.ExecuteNonQuery();

                insertCmd = "INSERT INTO " + dbTableName + " (Location, Timestamp, Power)" +
                        " VALUES ('EMU3','" + now + "'," + powerEMU3 + ");";
                InsertCommand = connection.CreateCommand();
                InsertCommand.CommandText = insertCmd;
                InsertCommand.ExecuteNonQuery();
                connection.Close();
            }
            catch (Exception ex)
            {
                try
                {
                    // SQL connection may have failed.  Try to recreate the db connection and write the data
                    SQLConnect();
                    connection.Open();

                    insertCmd = "INSERT INTO " + dbTableName + " (Location, Timestamp, Power)" +
                                   " VALUES ('EMU0','" + now + "', " + powerEMU0 + ");";
                    InsertCommand = connection.CreateCommand();
                    InsertCommand.CommandText = insertCmd;
                    InsertCommand.ExecuteNonQuery();

                    insertCmd = "INSERT INTO " + dbTableName + " (Location, Timestamp, Power)" +
                                " VALUES ('EMU1','" + now + "'," + powerEMU1 + ");";
                    InsertCommand = connection.CreateCommand();
                    InsertCommand.CommandText = insertCmd;
                    InsertCommand.ExecuteNonQuery();

                    insertCmd = "INSERT INTO " + dbTableName + " (Location, Timestamp, Power)" +
                                " VALUES ('EMU2','" + now + "'," + powerEMU2 + ");";
                    InsertCommand = connection.CreateCommand();
                    InsertCommand.CommandText = insertCmd;
                    InsertCommand.ExecuteNonQuery();

                    insertCmd = "INSERT INTO " + dbTableName + " (Location, Timestamp, Power)" +
                            " VALUES ('EMU3','" + now + "'," + powerEMU3 + ");";
                    InsertCommand = connection.CreateCommand();
                    InsertCommand.CommandText = insertCmd;
                    InsertCommand.ExecuteNonQuery();
                    connection.Close();
                }
                catch
                {
                    // Failed to write.
                    eventLog1.WriteEntry("Error: " + ex.Message);
                }
            }

            //  If the time is 10:10, collect the daily averages from the previous day and send out an email.
            string curTime = string.Format("{0:HH:mm}", DateTime.Now);
            if (string.Equals(curTime, "10:10") && mailSent != 1)
            {
                List<double> dayEMU0 = new List<double>();
                List<double> nightEMU0 = new List<double>();
                List<double> dayEMU1 = new List<double>();
                List<double> nightEMU1 = new List<double>();
                List<double> dayEMU2 = new List<double>();
                List<double> nightEMU2 = new List<double>();
                List<double> dayEMU3 = new List<double>();
                List<double> nightEMU3 = new List<double>();
                List<string> yearsList = new List<string>();
                try
                {
                    connection.Open();
                    //nPower - Night time power usage goes from 7pm - 7am
                    DateTime StartTime = DateTime.Now.AddHours(-15);
                    StartTime = StartTime.AddMinutes(-10);
                    DateTime EndTime = DateTime.Now.AddHours(-3);
                    EndTime = EndTime.AddMinutes(-10);
                    string yearTitle = "\t\t\t" + StartTime.Date.ToString("d");
                    yearsList.Add(StartTime.ToString("d"));
                    string dayNote = "\t\t\t" + StartTime.DayOfWeek;
                    string neTime = EndTime.ToString("yyyy-MM-dd H:mm:ss");
                    string nsTime = StartTime.ToString("yyyy-MM-dd H:mm:ss");

                    // Get current night time data from yesterday with EMU0
                    insertCmd = "SELECT AVG(Power) AS AvgPower FROM " + dbTableName + " WHERE (Timestamp >= '" + nsTime + "') AND (Timestamp < '" + neTime + "') AND (Location = '4W')";
                    InsertCommand = connection.CreateCommand();
                    InsertCommand.CommandText = insertCmd;
                    double nPowerEMU0;

                    try
                    {
                        nPowerEMU0 = Convert.ToDouble(InsertCommand.ExecuteScalar());
                    }
                    catch
                    {
                        nPowerEMU0 = 0;
                    }
                    nightEMU0.Add(nPowerEMU0);

                    // Get night time data for previous years on EMU0 as necessary.
                    string nOldPowerEMU0 = string.Empty;
                    if (prevYears > 0)
                    {
                        for (int i = 1; i <= prevYears; i++)
                        {
                            DateTime sTime = StartTime.AddYears(-i);
                            DateTime eTime = EndTime.AddYears(-i);
                            sTime.AddDays(5);
                            eTime.AddDays(5);
                            string oldneTime;
                            string oldnsTime;
                            if (DateTime.IsLeapYear(sTime.Year))
                            {
                                sTime.AddDays((i - 1) + 2);
                                eTime.AddDays((i - 1) + 2);
                                yearsList.Add(sTime.AddDays((i - 1) + 2).ToString("d"));
                                oldneTime = eTime.AddDays((i - 1) + 2).ToString("yyyy-MM-dd H:mm:ss");
                                oldnsTime = sTime.AddDays((i - 1) + 2).ToString("yyyy-MM-dd H:mm:ss");
                            }
                            else
                            {
                                sTime.AddDays((i - 1) + 1);
                                eTime.AddDays((i - 1) + 1);
                                yearsList.Add(sTime.AddDays((i - 1) + 1).ToString("d"));
                                oldneTime = eTime.AddDays((i - 1) + 1).ToString("yyyy-MM-dd H:mm:ss");
                                oldnsTime = sTime.AddDays((i - 1) + 1).ToString("yyyy-MM-dd H:mm:ss");
                            }
                            string year = sTime.Date.ToString("d");
                            
                            if (sTime.DayOfWeek.ToString().Length < 8)
                            {
                                dayNote = dayNote + "\t/\t" + sTime.DayOfWeek + "\t";
                            }
                            else
                            {
                                dayNote = dayNote + "\t/\t" + sTime.DayOfWeek;
                            }
                            yearTitle = yearTitle + "\t/\t" + year;
                            insertCmd = "SELECT AVG(Power) AS AvgPower FROM " + dbTableName + " WHERE (Timestamp >= '" + oldnsTime + "') AND (Timestamp < '" + oldneTime + "') AND (Location = '4W')";
                            InsertCommand = connection.CreateCommand();
                            InsertCommand.CommandText = insertCmd;
                            try
                            {
                                double old4W = Convert.ToDouble(InsertCommand.ExecuteScalar());
                                nightEMU0.Add(old4W);
                                string value = string.Format("{0:N3}", old4W) + " kW";
                                if (value.Length < 8)
                                {
                                    nOldPowerEMU0 += "\t/\t" + string.Format("{0:N3}", old4W) + " kW\t";
                                }
                                else
                                {
                                    nOldPowerEMU0 += "\t/\t" + string.Format("{0:N3}", old4W) + " kW";
                                }
                            }
                            catch
                            {
                                nOldPowerEMU0 += "\t/\t0 kW\t";
                                nightEMU0.Add(0.0);
                            }
                        }
                    }

                    // Get current night time data from yesterday on EMU1
                    insertCmd = "SELECT AVG(Power) AS AvgPower FROM " + dbTableName + " WHERE (Timestamp >= '" + nsTime + "') AND (Timestamp < '" + neTime + "') AND (Location = '4E')";
                    InsertCommand = connection.CreateCommand();
                    InsertCommand.CommandText = insertCmd;
                    double nPowerEMU1;
                    try
                    {
                        nPowerEMU1 = Convert.ToDouble(InsertCommand.ExecuteScalar());
                    }
                    catch
                    {
                        nPowerEMU1 = 0;
                    }
                    nightEMU1.Add(nPowerEMU1);

                    // Get night time data for previous years on EMU1 as necessary.
                    string nOldPowerEMU1 = string.Empty;
                    if (prevYears > 0)
                    {
                        for (int i = 1; i <= prevYears; i++)
                        {
                            DateTime sTime = StartTime.AddYears(-i);
                            DateTime eTime = EndTime.AddYears(-i);
                            string oldneTime;
                            string oldnsTime;
                            if (DateTime.IsLeapYear(sTime.Year))
                            {
                                sTime.AddDays((i - 1) + 2);
                                eTime.AddDays((i - 1) + 2);
                                oldneTime = eTime.AddDays((i - 1) + 2).ToString("yyyy-MM-dd H:mm:ss");
                                oldnsTime = sTime.AddDays((i - 1) + 2).ToString("yyyy-MM-dd H:mm:ss");
                            }
                            else
                            {
                                sTime.AddDays((i - 1) + 1);
                                eTime.AddDays((i - 1) + 1);
                                oldneTime = eTime.AddDays((i - 1) + 1).ToString("yyyy-MM-dd H:mm:ss");
                                oldnsTime = sTime.AddDays((i - 1) + 1).ToString("yyyy-MM-dd H:mm:ss");
                            }
                            string year = sTime.Date.ToString("d");
                            insertCmd = "SELECT AVG(Power) AS AvgPower FROM " + dbTableName + " WHERE (Timestamp >= '" + oldnsTime + "') AND (Timestamp < '" + oldneTime + "') AND (Location = '4E')";
                            InsertCommand = connection.CreateCommand();
                            InsertCommand.CommandText = insertCmd;
                            try
                            {
                                double old4E = Convert.ToDouble(InsertCommand.ExecuteScalar());
                                nightEMU1.Add(old4E);
                                string value = string.Format("{0:N3}", old4E) + " kW";
                                if (value.Length < 8)
                                {
                                    nOldPowerEMU1 += "\t/\t" + string.Format("{0:N3}", old4E) + " kW\t";
                                }
                                else
                                {
                                    nOldPowerEMU1 += "\t/\t" + string.Format("{0:N3}", old4E) + " kW";
                                }
                            }
                            catch
                            {
                                nOldPowerEMU1 += "\t/\t0 kW\t";
                                nightEMU1.Add(0.0);
                            }
                        }
                    }

                    // Get current night time data from yesterday on EMU2
                    insertCmd = "SELECT AVG(Power) AS AvgPower FROM " + dbTableName + " WHERE (Timestamp >= '" + nsTime + "') AND (Timestamp < '" + neTime + "') AND (Location = '5W')";
                    InsertCommand = connection.CreateCommand();
                    InsertCommand.CommandText = insertCmd;
                    double nPowerEMU2;
                    try
                    {
                        nPowerEMU2 = Convert.ToDouble(InsertCommand.ExecuteScalar());
                    }
                    catch
                    {
                        nPowerEMU2 = 0;
                    }
                    nightEMU2.Add(nPowerEMU2);

                    // Get night time data for previous years on EMU2 as necessary.
                    string nOldPowerEMU2 = string.Empty;
                    if (prevYears > 0)
                    {
                        for (int i = 1; i <= prevYears; i++)
                        {
                            DateTime sTime = StartTime.AddYears(-i);
                            DateTime eTime = EndTime.AddYears(-i);
                            string oldneTime;
                            string oldnsTime;
                            if (DateTime.IsLeapYear(sTime.Year))
                            {
                                sTime.AddDays((i - 1) + 2);
                                eTime.AddDays((i - 1) + 2);
                                oldneTime = eTime.AddDays((i - 1) + 2).ToString("yyyy-MM-dd H:mm:ss");
                                oldnsTime = sTime.AddDays((i - 1) + 2).ToString("yyyy-MM-dd H:mm:ss");
                            }
                            else
                            {
                                sTime.AddDays((i - 1) + 1);
                                eTime.AddDays((i - 1) + 1);
                                oldneTime = eTime.AddDays((i - 1) + 1).ToString("yyyy-MM-dd H:mm:ss");
                                oldnsTime = sTime.AddDays((i - 1) + 1).ToString("yyyy-MM-dd H:mm:ss");
                            }
                            string year = sTime.Date.ToString("d");
                            insertCmd = "SELECT AVG(Power) AS AvgPower FROM " + dbTableName + " WHERE (Timestamp >= '" + oldnsTime + "') AND (Timestamp < '" + oldneTime + "') AND (Location = '5W')";
                            InsertCommand = connection.CreateCommand();
                            InsertCommand.CommandText = insertCmd;
                            try
                            {
                                double old5W = Convert.ToDouble(InsertCommand.ExecuteScalar());
                                nightEMU2.Add(old5W);
                                string value = string.Format("{0:N3}", old5W) + " kW";
                                if (value.Length < 8)
                                {
                                    nOldPowerEMU2 += "\t/\t" + string.Format("{0:N3}", old5W) + " kW\t";
                                }
                                else
                                {
                                    nOldPowerEMU2 += "\t/\t" + string.Format("{0:N3}", old5W) + " kW";
                                }
                            }
                            catch
                            {
                                nOldPowerEMU2 += "\t/\t0 kW\t";
                                nightEMU2.Add(0.0);
                            }
                        }
                    }

                    // Get current night time data from yesterday on EMU3
                    insertCmd = "SELECT AVG(Power) AS AvgPower FROM " + dbTableName + " WHERE (Timestamp >= '" + nsTime + "') AND (Timestamp < '" + neTime + "') AND (Location = '5E')";
                    InsertCommand = connection.CreateCommand();
                    InsertCommand.CommandText = insertCmd;
                    double nPowerEMU3;
                    try
                    {
                        nPowerEMU3 = Convert.ToDouble(InsertCommand.ExecuteScalar());
                    }
                    catch
                    {
                        nPowerEMU3 = 0;
                    }
                    nightEMU3.Add(nPowerEMU3);

                    // Get night time data for previous years on EMU3 as necessary.
                    string nOldPowerEMU3 = string.Empty;
                    if (prevYears > 0)
                    {
                        for (int i = 1; i <= prevYears; i++)
                        {
                            DateTime sTime = StartTime.AddYears(-i);
                            DateTime eTime = EndTime.AddYears(-i);
                            string oldneTime;
                            string oldnsTime;
                            if (DateTime.IsLeapYear(sTime.Year))
                            {
                                sTime.AddDays((i - 1) + 2);
                                eTime.AddDays((i - 1) + 2);
                                oldneTime = eTime.AddDays((i - 1) + 2).ToString("yyyy-MM-dd H:mm:ss");
                                oldnsTime = sTime.AddDays((i - 1) + 2).ToString("yyyy-MM-dd H:mm:ss");
                            }
                            else
                            {
                                sTime.AddDays((i - 1) + 1);
                                eTime.AddDays((i - 1) + 1);
                                oldneTime = eTime.AddDays((i - 1) + 1).ToString("yyyy-MM-dd H:mm:ss");
                                oldnsTime = sTime.AddDays((i - 1) + 1).ToString("yyyy-MM-dd H:mm:ss");
                            }
                            string year = sTime.Date.ToString("d");
                            insertCmd = "SELECT AVG(Power) AS AvgPower FROM " + dbTableName + " WHERE (Timestamp >= '" + oldnsTime + "') AND (Timestamp < '" + oldneTime + "') AND (Location = '5E')";
                            InsertCommand = connection.CreateCommand();
                            InsertCommand.CommandText = insertCmd;
                            try
                            {
                                double old5E = Convert.ToDouble(InsertCommand.ExecuteScalar());
                                nightEMU3.Add(old5E);
                                string value = string.Format("{0:N3}", old5E) + " kW";
                                if (value.Length < 8)
                                {
                                    nOldPowerEMU3 += "\t/\t" + string.Format("{0:N3}", old5E) + " kW\t";
                                }
                                else
                                {
                                    nOldPowerEMU3 += "\t/\t" + string.Format("{0:N3}", old5E) + " kW";
                                }
                            }
                            catch
                            {
                                nOldPowerEMU3 += "\t/\t0 kW\t";
                                nightEMU3.Add(0.0);
                            }
                        }
                    }

                    //dPower - Day power.  Days are from 7am - 7pm.
                    StartTime = DateTime.Now.AddHours(-27);
                    StartTime = StartTime.AddMinutes(-10);
                    EndTime = DateTime.Now.AddHours(-15);
                    EndTime = EndTime.AddMinutes(-10);

                    string deTime = EndTime.ToString("yyyy-MM-dd H:mm:ss");
                    string dsTime = StartTime.ToString("yyyy-MM-dd H:mm:ss");

                    // Get current day time data from yesterday on EMU0
                    insertCmd = "SELECT AVG(Power) AS AvgPower FROM " + dbTableName + " WHERE (Timestamp >= '" + dsTime + "') AND (Timestamp < '" + deTime + "')  AND (Location = '4W')";
                    InsertCommand = connection.CreateCommand();
                    InsertCommand.CommandText = insertCmd;
                    double dPowerEMU0;
                    try
                    {
                        dPowerEMU0 = Convert.ToDouble(InsertCommand.ExecuteScalar());
                    }
                    catch
                    {
                        dPowerEMU0 = 0;
                    }
                    dayEMU0.Add(dPowerEMU0);

                    // Get day time data for previous years on EMU0 as necessary.
                    string dOldPowerEMU0 = string.Empty;
                    if (prevYears > 0)
                    {
                        for (int i = 1; i <= prevYears; i++)
                        {
                            DateTime sTime = StartTime.AddYears(-i);
                            DateTime eTime = EndTime.AddYears(-i);

                            string oldneTime;
                            string oldnsTime;
                            if (DateTime.IsLeapYear(sTime.Year))
                            {
                                sTime.AddDays((i - 1) + 2);
                                eTime.AddDays((i - 1) + 2);
                                oldneTime = eTime.AddDays((i - 1) + 2).ToString("yyyy-MM-dd H:mm:ss");
                                oldnsTime = sTime.AddDays((i - 1) + 2).ToString("yyyy-MM-dd H:mm:ss");
                            }
                            else
                            {
                                sTime.AddDays((i - 1) + 1);
                                eTime.AddDays((i - 1) + 1);
                                oldneTime = eTime.AddDays((i - 1) + 1).ToString("yyyy-MM-dd H:mm:ss");
                                oldnsTime = sTime.AddDays((i - 1) + 1).ToString("yyyy-MM-dd H:mm:ss");
                            }
                            string year = sTime.Date.ToString("d");
                            insertCmd = "SELECT AVG(Power) AS AvgPower FROM " + dbTableName + " WHERE (Timestamp >= '" + oldnsTime + "') AND (Timestamp < '" + oldneTime + "') AND (Location = '4W')";
                            InsertCommand = connection.CreateCommand();
                            InsertCommand.CommandText = insertCmd;
                            try
                            {
                                double old4W = Convert.ToDouble(InsertCommand.ExecuteScalar());
                                dayEMU0.Add(old4W);
                                string value = string.Format("{0:N3}", old4W) + " kW";
                                if (value.Length < 8)
                                {
                                    dOldPowerEMU0 += "\t/\t" + string.Format("{0:N3}", old4W) + " kW\t";
                                }
                                else
                                {
                                    dOldPowerEMU0 += "\t/\t" + string.Format("{0:N3}", old4W) + " kW";
                                }
                            }
                            catch
                            {
                                dOldPowerEMU0 += "\t/\t0 kW\t";
                                dayEMU0.Add(0.0);
                            }
                        }
                    }

                    // Get current day time data from yesterday on EMU1
                    insertCmd = "SELECT AVG(Power) AS AvgPower FROM " + dbTableName + " WHERE (Timestamp >= '" + dsTime + "') AND (Timestamp < '" + deTime + "')  AND (Location = '4E')";
                    InsertCommand = connection.CreateCommand();
                    InsertCommand.CommandText = insertCmd;
                    double dPowerEMU1;
                    try
                    {
                        dPowerEMU1 = Convert.ToDouble(InsertCommand.ExecuteScalar());
                    }
                    catch
                    {
                        dPowerEMU1 = 0;
                    }
                    dayEMU1.Add(dPowerEMU1);

                    // Get day time data for previous years on EMU1 as necessary.
                    string dOldPowerEMU1 = string.Empty;
                    if (prevYears > 0)
                    {
                        for (int i = 1; i <= prevYears; i++)
                        {
                            DateTime sTime = StartTime.AddYears(-i);
                            DateTime eTime = EndTime.AddYears(-i);
                            string oldneTime;
                            string oldnsTime;
                            if (DateTime.IsLeapYear(sTime.Year))
                            {
                                sTime.AddDays((i - 1) + 2);
                                eTime.AddDays((i - 1) + 2);
                                oldneTime = eTime.AddDays((i - 1) + 2).ToString("yyyy-MM-dd H:mm:ss");
                                oldnsTime = sTime.AddDays((i - 1) + 2).ToString("yyyy-MM-dd H:mm:ss");
                            }
                            else
                            {
                                sTime.AddDays((i - 1) + 1);
                                eTime.AddDays((i - 1) + 1);
                                oldneTime = eTime.AddDays((i - 1) + 1).ToString("yyyy-MM-dd H:mm:ss");
                                oldnsTime = sTime.AddDays((i - 1) + 1).ToString("yyyy-MM-dd H:mm:ss");
                            }

                            string year = sTime.Date.ToString("d");
                            insertCmd = "SELECT AVG(Power) AS AvgPower FROM " + dbTableName + " WHERE (Timestamp >= '" + oldnsTime + "') AND (Timestamp < '" + oldneTime + "') AND (Location = '4E')";
                            InsertCommand = connection.CreateCommand();
                            InsertCommand.CommandText = insertCmd;
                            try
                            {
                                double old4E = Convert.ToDouble(InsertCommand.ExecuteScalar());
                                dayEMU1.Add(old4E);
                                string value = string.Format("{0:N3}", old4E) + " kW";
                                if (value.Length < 8)
                                {
                                    dOldPowerEMU1 += "\t/\t" + string.Format("{0:N3}", old4E) + " kW\t";
                                }
                                else
                                {
                                    dOldPowerEMU1 += "\t/\t" + string.Format("{0:N3}", old4E) + " kW";
                                }
                            }
                            catch
                            {
                                dOldPowerEMU1 += "\t/\t0 kW\t";
                                dayEMU1.Add(0.0);
                            }
                        }
                    }

                    // Get current day time data from yesterday on EMU2
                    insertCmd = "SELECT AVG(Power) AS AvgPower FROM " + dbTableName + " WHERE (Timestamp >= '" + dsTime + "') AND (Timestamp < '" + deTime + "')  AND (Location = '5W')";
                    InsertCommand = connection.CreateCommand();
                    InsertCommand.CommandText = insertCmd;
                    double dPowerEMU2;
                    try
                    {
                        dPowerEMU2 = Convert.ToDouble(InsertCommand.ExecuteScalar());
                    }
                    catch
                    {
                        dPowerEMU2 = 0;
                    }
                    dayEMU2.Add(dPowerEMU2);

                    // Get day time data for previous years on EMU2 as necessary.
                    string dOldPowerEMU2 = string.Empty;
                    if (prevYears > 0)
                    {
                        for (int i = 1; i <= prevYears; i++)
                        {
                            DateTime sTime = StartTime.AddYears(-i);
                            DateTime eTime = EndTime.AddYears(-i);
                            string oldneTime;
                            string oldnsTime;
                            if (DateTime.IsLeapYear(sTime.Year))
                            {
                                sTime.AddDays((i - 1) + 2);
                                eTime.AddDays((i - 1) + 2);
                                oldneTime = eTime.AddDays((i - 1) + 2).ToString("yyyy-MM-dd H:mm:ss");
                                oldnsTime = sTime.AddDays((i - 1) + 2).ToString("yyyy-MM-dd H:mm:ss");
                            }
                            else
                            {
                                sTime.AddDays((i - 1) + 1);
                                eTime.AddDays((i - 1) + 1);
                                oldneTime = eTime.AddDays((i - 1) + 1).ToString("yyyy-MM-dd H:mm:ss");
                                oldnsTime = sTime.AddDays((i - 1) + 1).ToString("yyyy-MM-dd H:mm:ss");
                            }
                            string year = sTime.Date.ToString("d");
                            insertCmd = "SELECT AVG(Power) AS AvgPower FROM " + dbTableName + " WHERE (Timestamp >= '" + oldnsTime + "') AND (Timestamp < '" + oldneTime + "') AND (Location = '5W')";
                            InsertCommand = connection.CreateCommand();
                            InsertCommand.CommandText = insertCmd;
                            try
                            {
                                double old5W = Convert.ToDouble(InsertCommand.ExecuteScalar());
                                dayEMU2.Add(old5W);
                                string value = string.Format("{0:N3}", old5W) + " kW";
                                if (value.Length < 8)
                                {
                                    dOldPowerEMU2 += "\t/\t" + string.Format("{0:N3}", old5W) + " kW\t";
                                }
                                else
                                {
                                    dOldPowerEMU2 += "\t/\t" + string.Format("{0:N3}", old5W) + " kW";
                                }
                            }
                            catch
                            {
                                dOldPowerEMU2 += "\t/\t0 kW\t";
                                dayEMU2.Add(0.0);
                            }
                        }
                    }

                    // Get current day time data from yesterday on EMU3
                    insertCmd = "SELECT AVG(Power) AS AvgPower FROM " + dbTableName + " WHERE (Timestamp >= '" + dsTime + "') AND (Timestamp < '" + deTime + "')  AND (Location = '5E')";
                    InsertCommand = connection.CreateCommand();
                    InsertCommand.CommandText = insertCmd;
                    double dPowerEMU3;
                    try
                    {
                        dPowerEMU3 = Convert.ToDouble(InsertCommand.ExecuteScalar());
                    }
                    catch
                    {
                        dPowerEMU3 = 0;
                    }
                    dayEMU3.Add(dPowerEMU3);

                    // Get day time data for previous years on EMU3 as necessary.
                    string dOldPowerEMU3 = string.Empty;
                    if (prevYears > 0)
                    {
                        for (int i = 1; i <= prevYears; i++)
                        {
                            DateTime sTime = StartTime.AddYears(-i);
                            DateTime eTime = EndTime.AddYears(-i);
                            string oldneTime;
                            string oldnsTime;
                            if (DateTime.IsLeapYear(sTime.Year))
                            {
                                sTime.AddDays((i - 1) + 2);
                                eTime.AddDays((i - 1) + 2);
                                oldneTime = eTime.AddDays((i - 1) + 2).ToString("yyyy-MM-dd H:mm:ss");
                                oldnsTime = sTime.AddDays((i - 1) + 2).ToString("yyyy-MM-dd H:mm:ss");
                            }
                            else
                            {
                                sTime.AddDays((i - 1) + 1);
                                eTime.AddDays((i - 1) + 1);
                                oldneTime = eTime.AddDays((i - 1) + 1).ToString("yyyy-MM-dd H:mm:ss");
                                oldnsTime = sTime.AddDays((i - 1) + 1).ToString("yyyy-MM-dd H:mm:ss");
                            }
                            string year = sTime.Date.ToString("d");
                            insertCmd = "SELECT AVG(Power) AS AvgPower FROM " + dbTableName + " WHERE (Timestamp >= '" + oldnsTime + "') AND (Timestamp < '" + oldneTime + "') AND (Location = '5E')";
                            InsertCommand = connection.CreateCommand();
                            InsertCommand.CommandText = insertCmd;
                            try
                            {
                                double old5E = Convert.ToDouble(InsertCommand.ExecuteScalar());
                                dayEMU3.Add(old5E);
                                string value = string.Format("{0:N3}", old5E) + " kW";
                                if (value.Length < 8)
                                {
                                    dOldPowerEMU3 += "\t/\t" + string.Format("{0:N3}", old5E) + " kW\t";
                                }
                                else
                                {
                                    dOldPowerEMU3 += "\t/\t" + string.Format("{0:N3}", old5E) + " kW";
                                }
                            }
                            catch
                            {
                                dOldPowerEMU3 += "\t/\t0 kW\t";
                                dayEMU3.Add(0.0);
                            }
                        }
                    }
                    // All data has been aquired.  Close the SQL connection
                    connection.Close();

                    // Generate charts for the email
                    GenerateChart(dayEMU0, nightEMU0, dayEMU1, nightEMU1, dayEMU2, nightEMU2, dayEMU3, nightEMU3, yearsList);

                    // Generate HTML Email.  Most of the email is based on predefined HTML files with keywords swapped out for the actual data.
                    string mainBody = Properties.Resources.MainEMail.ToString();

                    // Start replacing variables from the main body.
                    mainBody = mainBody.Replace("#DAY#", StartTime.DayOfWeek.ToString());
                    mainBody = mainBody.Replace("#CURRENTDATE#", yearsList[0]);
                    mainBody = mainBody.Replace("#DEMU0#", string.Format("{0:N3}", dayEMU0[0]));
                    mainBody = mainBody.Replace("#NEMU0#", string.Format("{0:N3}", nightEMU0[0]));
                    mainBody = mainBody.Replace("#DEMU1", string.Format("{0:N3}", dayEMU1[0]));
                    mainBody = mainBody.Replace("#NEMU1#", string.Format("{0:N3}", nightEMU1[0]));
                    mainBody = mainBody.Replace("#DEMU2#", string.Format("{0:N3}", dayEMU2[0]));
                    mainBody = mainBody.Replace("#NEMU2#", string.Format("{0:N3}", nightEMU2[0]));
                    mainBody = mainBody.Replace("#DEMU3#", string.Format("{0:N3}", dayEMU3[0]));
                    mainBody = mainBody.Replace("#NEMU3#", string.Format("{0:N3}", nightEMU3[0]));
                    mainBody = mainBody.Replace("#TOTAL#", string.Format("{0:N3}", (dayEMU0[0] + nightEMU0[0] + dayEMU1[0] + nightEMU1[0] + dayEMU2[0] + nightEMU2[0] + dayEMU3[0] + nightEMU3[0])));
                    mainBody = mainBody.Replace("#DATAVIEW#", vizURL);
                    mainBody = mainBody.Replace("#UNSUBSCRIBE#", unsubscribe);

                    // Showing multiple years of data is optional and could just show a simple chart with the most current data.
                    if (yearsList.Count > 1)
                    {
                        string addYearData = string.Empty;
                        string addCols = string.Empty;
                        string addDates = string.Empty;
                        string dEMU0Prev = string.Empty;
                        string nEMU0Prev = string.Empty;
                        string dEMU1Prev = string.Empty;
                        string nEMU1Prev = string.Empty;
                        string dEMU2Prev = string.Empty;
                        string nEMU2Prev = string.Empty;
                        string dEMU3Prev = string.Empty;
                        string nEMU3Prev = string.Empty;
                        string totPrev = string.Empty;
                        for (int i = 1; i < yearsList.Count; i++)
                        {
                            addCols += colHTML;
                            addDates += typHTML.Replace("#VALUE#", yearsList[i]);
                            dEMU0Prev += typHTML.Replace("#VALUE#", string.Format("{0:N3}", dayEMU0[i]) + "kW");
                            nEMU0Prev += typHTML.Replace("#VALUE#", string.Format("{0:N3}", nightEMU0[i]) + "kW");
                            dEMU1Prev += typHTML.Replace("#VALUE#", string.Format("{0:N3}", dayEMU1[i]) + "kW");
                            nEMU1Prev += typHTML.Replace("#VALUE#", string.Format("{0:N3}", nightEMU1[i]) + "kW");
                            dEMU2Prev += typHTML.Replace("#VALUE#", string.Format("{0:N3}", dayEMU2[i]) + "kW");
                            nEMU2Prev += typHTML.Replace("#VALUE#", string.Format("{0:N3}", nightEMU2[i]) + "kW");
                            dEMU3Prev += typHTML.Replace("#VALUE#", string.Format("{0:N3}", dayEMU3[i]) + "kW");
                            nEMU3Prev += typHTML.Replace("#VALUE#", string.Format("{0:N3}", nightEMU3[i]) + "kW");
                            totPrev += totHTML.Replace("#VALUE#", string.Format("{0:N3}", (dayEMU0[i] + nightEMU0[i] + dayEMU1[i] + nightEMU1[i] + dayEMU2[i] + nightEMU2[i] + dayEMU3[i] + nightEMU3[i])));
                        }

                        mainBody = mainBody.Replace("#PREVCOLWIDTHS#", addCols);
                        mainBody = mainBody.Replace("#PREVDATES#", addDates);
                        mainBody = mainBody.Replace("#EMU0DAYPREV#", dEMU0Prev);
                        mainBody = mainBody.Replace("#EMU0NIGHTPREV#", nEMU0Prev);
                        mainBody = mainBody.Replace("#EMU1DAYPREV#", dEMU1Prev);
                        mainBody = mainBody.Replace("#EMU1NIGHTPREV#", nEMU1Prev);
                        mainBody = mainBody.Replace("#EMU2DAYPREV#", dEMU2Prev);
                        mainBody = mainBody.Replace("#EMU2NIGHTPREV#", nEMU2Prev);
                        mainBody = mainBody.Replace("#EMU3DAYPREV#", dEMU3Prev);
                        mainBody = mainBody.Replace("#EMU3NIGHTPREV#", nEMU3Prev);
                        mainBody = mainBody.Replace("#TOTALSPREV#", totPrev);
                    }
                    else
                    {
                        mainBody = mainBody.Replace("#ADDYEARS#", "");
                    }

                    SendEmail("EMUGroup@MyCompany.com", "EmailSentFrom@MyCompany.com", "[LMN EMU] Daily Energy Usage Email Digest", mainBody);

                    mailSent = 1;

                }
                catch (Exception ex)
                {
                    eventLog1.WriteEntry(ex.Message);
                }
            }
            else if (string.Equals(curTime, "10:11") || string.Equals(curTime, "10:12"))
            {
                // The mailSent int is meant to ensure the email only gets sent once.  After the time is up reset the value
                // and delete the charts so they can be created anew the next day.
                mailSent = 0;

                chartEMU0Img.Dispose();
                chartEMU1Img.Dispose();
                chartEMU3Img.Dispose();
                chartEMU2Img.Dispose();

                string path = @"C:\temp\";
                List<string> fileNames = new List<string>();
                fileNames.Add("EMU0Chart.jpg");
                fileNames.Add("EMU1Chart.jpg");
                fileNames.Add("EMU2Chart.jpg");
                fileNames.Add("EMU3Chart.jpg");

                for (int i = 0; i < 4; i++)
                {
                    try
                    {
                        System.IO.File.Delete(path + fileNames[i]);
                    }
                    catch
                    { }
                }
            }
        }

        private void SendEmail(string sendTo, string sendFrom, string subject, string message)
        {
            AlternateView avHtml = AlternateView.CreateAlternateViewFromString(message, null, System.Net.Mime.MediaTypeNames.Text.Html);
            chartEMU0Img.ContentId = "Chart4W";
            chartEMU1Img.ContentId = "Chart4E";
            chartEMU2Img.ContentId = "Chart5W";
            chartEMU3Img.ContentId = "Chart5E";

            // May be better to embed the resources into the appliation rather than rely on them to be deployed.
            // The addresses where the files reside is the final resting place of the program when it's installed.
            LinkedResource chartYAxisImg = new LinkedResource("C:\\Program Files\\EMU Monitor\\YAxis.png", MediaTypeNames.Image.Jpeg);
            chartYAxisImg.ContentId = "YAxis";
            LinkedResource emuImg = new LinkedResource("C:\\Program Files\\EMU Monitor\\EMU Logo.png", MediaTypeNames.Image.Jpeg);
            emuImg.ContentId = "LMNtsEMU";

            avHtml.LinkedResources.Add(chartEMU0Img);
            avHtml.LinkedResources.Add(chartEMU1Img);
            avHtml.LinkedResources.Add(chartEMU2Img);
            avHtml.LinkedResources.Add(chartEMU3Img);
            avHtml.LinkedResources.Add(chartYAxisImg);
            avHtml.LinkedResources.Add(emuImg);

            System.Net.Mail.MailMessage mailMessage = new System.Net.Mail.MailMessage();
            mailMessage.To.Add(sendTo);
            mailMessage.From = new System.Net.Mail.MailAddress(sendFrom);
            mailMessage.Subject = subject;
            mailMessage.AlternateViews.Add(avHtml);
            mailMessage.IsBodyHtml = true;
            System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient(mailClient);
            smtp.Send(mailMessage);
        }

        private void GenerateChart(List<double> dEMU0, List<double> nEMU0, List<double> dEMU1, List<double> nEMU1, List<double> dEMU2, List<double> nEMU2, List<double> dEMU3, List<double> nEMU3, List<string> year)
        {
            // The charts that are shown at the bottom of the email are all .Net generated and then images
            // are created from them to be used within the email.
            Chart chartEMU0;
            Chart chartEMU1;
            Chart chartEMU2;
            Chart chartEMU3;

            List<string> trimmedYears = new List<string>();
            foreach (string s in year)
            {
                string subString = s.Substring(s.Length - Math.Min(4, s.Length));
                trimmedYears.Add(subString);
            }

            // EMU 0
            chartEMU0 = new Chart();
            chartEMU0.Size = new Size(150, 300);
            chartEMU0.Titles.Add("EMU 0");

            ChartArea ca = new ChartArea();
            ca.InnerPlotPosition = new ElementPosition(-3, 0, 106, 90);
            ca.BorderWidth = 0;
            ca.AxisX.MajorGrid.LineColor = Color.FromArgb(0, 0, 0, 0);
            ca.AxisY.MajorGrid.LineColor = Color.LightGray;
            ca.AxisY.MinorGrid.LineColor = Color.LightGray;
            ca.AxisX.LabelStyle.Font = new Font("Arial", 10);
            ca.AxisX.IsLabelAutoFit = false;
            ca.AxisX.MajorTickMark.Enabled = false;
            ca.AxisX.LabelStyle.Angle = -90;
            ca.AxisY.LabelStyle.Enabled = false;
            ca.AxisY.MajorTickMark.Enabled = false;
            ca.AxisY.Maximum = 50;
            chartEMU0.ChartAreas.Add(ca);

            // EMU0 Day
            Series dEMU0Series = new Series();
            dEMU0Series.Name = "EMU0 Day";
            dEMU0Series.ChartType = SeriesChartType.Column;
            dEMU0Series.XValueType = ChartValueType.String;
            dEMU0Series.YValueType = ChartValueType.Double;
            chartEMU0.Series.Add(dEMU0Series);

            double[] d4Wvalues = dEMU0.ToArray();
            string[] years = trimmedYears.ToArray();
            chartEMU0.Series["EMU0 Day"].Points.DataBindXY(years, d4Wvalues);
            chartEMU0.Series["EMU0 Day"].Points[0].Color = Color.FromArgb(255, 47, 36, 191);
            try
            {
                chartEMU0.Series["EMU0 Day"].Points[1].Color = Color.FromArgb(255, 99, 91, 207);
            }
            catch { }
            try
            {
                chartEMU0.Series["EMU0 Day"].Points[2].Color = Color.FromArgb(255, 151, 145, 223);
            }
            catch { }

            // EMU0 Night
            Series nEMU0Series = new Series();
            nEMU0Series.Name = "EMU0 Night";
            nEMU0Series.ChartType = SeriesChartType.Column;
            nEMU0Series.XValueType = ChartValueType.String;
            nEMU0Series.YValueType = ChartValueType.Double;
            chartEMU0.Series.Add(nEMU0Series);

            double[] n4Wvalues = nEMU0.ToArray();
            chartEMU0.Series["EMU0 Night"].Points.DataBindXY(years, n4Wvalues);
            chartEMU0.Series["EMU0 Night"].Points[0].Color = Color.FromArgb(255, 31, 24, 127);
            try
            {
                chartEMU0.Series["EMU0 Night"].Points[1].Color = Color.FromArgb(255, 87, 82, 159);
            }
            catch { }
            try
            {
                chartEMU0.Series["EMU0 Night"].Points[2].Color = Color.FromArgb(255, 143, 139, 191);
            }
            catch { }

            // EMU 1
            chartEMU1 = new Chart();
            chartEMU1.Size = new Size(150, 300);
            chartEMU1.Titles.Add("EMU 1");
            chartEMU1.BorderSkin.BorderWidth = 0;

            ca = new ChartArea();

            ca.InnerPlotPosition = new ElementPosition(-3, 0, 106, 90);
            ca.BorderWidth = 0;
            ca.AxisX.MajorGrid.LineColor = Color.FromArgb(0, 0, 0, 0);
            ca.AxisY.MajorGrid.LineColor = Color.LightGray;
            ca.AxisY.MinorGrid.LineColor = Color.LightGray;
            ca.AxisX.LabelStyle.Font = new Font("Arial", 10);
            ca.AxisX.IsLabelAutoFit = false;
            ca.AxisX.LabelStyle.Angle = -90;
            ca.AxisX.MajorTickMark.Enabled = false;
            ca.AxisY.LabelStyle.Enabled = false;
            ca.AxisY.LineColor = Color.FromArgb(0, 0, 0, 0);
            ca.AxisY.MajorTickMark.Enabled = false;
            ca.AxisY.Maximum = 50;

            chartEMU1.ChartAreas.Add(ca);

            // EMU1 Day
            Series dEMU1Series = new Series();
            dEMU1Series.Name = "EMU1 Day";
            dEMU1Series.ChartType = SeriesChartType.Column;
            dEMU1Series.XValueType = ChartValueType.String;
            dEMU1Series.YValueType = ChartValueType.Double;
            chartEMU1.Series.Add(dEMU1Series);

            double[] d4EVals = dEMU1.ToArray();
            chartEMU1.Series["EMU1 Day"].Points.DataBindXY(years, d4EVals);
            chartEMU1.Series["EMU1 Day"].Points[0].Color = Color.FromArgb(255, 191, 32, 17);
            try
            {
                chartEMU1.Series["EMU1 Day"].Points[1].Color = Color.FromArgb(255, 207, 88, 77);
            }
            catch { }
            try
            {
                chartEMU1.Series["EMU1 Day"].Points[2].Color = Color.FromArgb(255, 223, 143, 136);
            }
            catch { }

            // EMU1 Day
            Series nEMU1Series = new Series();
            nEMU1Series.Name = "EMU1 Night";
            nEMU1Series.ChartType = SeriesChartType.Column;
            nEMU1Series.XValueType = ChartValueType.String;
            nEMU1Series.YValueType = ChartValueType.Double;
            chartEMU1.Series.Add(nEMU1Series);

            double[] n4EVals = nEMU1.ToArray();
            chartEMU1.Series["EMU1 Night"].Points.DataBindXY(years, n4EVals);
            chartEMU1.Series["EMU1 Night"].Points[0].Color = Color.FromArgb(255, 127, 21, 11);
            try
            {
                chartEMU1.Series["EMU1 Night"].Points[1].Color = Color.FromArgb(255, 159, 80, 72);
            }
            catch { }
            try
            {
                chartEMU1.Series["EMU1 Night"].Points[2].Color = Color.FromArgb(255, 191, 138, 133);
            }
            catch { }

            // EMU 2
            chartEMU2 = new Chart();
            chartEMU2.Size = new Size(150, 300);
            chartEMU2.Titles.Add("EMU 2");

            ca = new ChartArea();

            ca.InnerPlotPosition = new ElementPosition(-3, 0, 106, 90);
            ca.BorderWidth = 0;
            ca.AxisX.MajorGrid.LineColor = Color.FromArgb(0, 0, 0, 0);
            ca.AxisY.MajorGrid.LineColor = Color.LightGray;
            ca.AxisY.MinorGrid.LineColor = Color.LightGray;
            ca.AxisX.LabelStyle.Font = new Font("Arial", 10);
            ca.AxisX.IsLabelAutoFit = false;
            ca.AxisX.LabelStyle.Angle = -90;
            ca.AxisX.MajorTickMark.Enabled = false;
            ca.AxisY.LabelStyle.Enabled = false;
            ca.AxisY.LineColor = Color.FromArgb(0, 0, 0, 0);
            ca.AxisY.MajorTickMark.Enabled = false;
            ca.AxisY.Maximum = 50;

            chartEMU2.ChartAreas.Add(ca);

            // EMU2 Day
            Series dEMU2Series = new Series();
            dEMU2Series.Name = "EMU2 Day";
            dEMU2Series.ChartType = SeriesChartType.Column;
            dEMU2Series.XValueType = ChartValueType.String;
            dEMU2Series.YValueType = ChartValueType.Double;
            chartEMU2.Series.Add(dEMU2Series);

            double[] dEMU2Vals = dEMU2.ToArray();
            chartEMU2.Series["EMU2 Day"].Points.DataBindXY(years, dEMU2Vals);
            chartEMU2.Series["EMU2 Day"].Points[0].Color = Color.FromArgb(255, 191, 102, 6);
            try
            {
                chartEMU2.Series["EMU2 Day"].Points[1].Color = Color.FromArgb(255, 207, 140, 68);
            }
            catch { }
            try
            {
                chartEMU2.Series["EMU2 Day"].Points[2].Color = Color.FromArgb(255, 223, 178, 130);
            }
            catch { }

            // EMU2 Night
            Series nEMU2Series = new Series();
            nEMU2Series.Name = "EMU2 Night";
            nEMU2Series.ChartType = SeriesChartType.Column;
            nEMU2Series.XValueType = ChartValueType.String;
            nEMU2Series.YValueType = ChartValueType.Double;
            chartEMU2.Series.Add(nEMU2Series);

            double[] nEMU2Vals = nEMU2.ToArray();
            chartEMU2.Series["EMU2 Night"].Points.DataBindXY(years, nEMU2Vals);
            chartEMU2.Series["EMU2 Night"].Points[0].Color = Color.FromArgb(255, 127, 68, 4);
            try
            {
                chartEMU2.Series["EMU2 Night"].Points[1].Color = Color.FromArgb(255, 159, 115, 67);
            }
            catch { }
            try
            {
                chartEMU2.Series["EMU2 Night"].Points[2].Color = Color.FromArgb(255, 191, 161, 129);
            }
            catch { }

            // EMU 3
            chartEMU3 = new Chart();
            chartEMU3.Size = new Size(150, 300);
            chartEMU3.Titles.Add("EMU 3");

            ca = new ChartArea();

            ca.InnerPlotPosition = new ElementPosition(-3, 0, 106, 90);
            ca.BorderWidth = 0;
            ca.AxisX.MajorGrid.LineColor = Color.FromArgb(0, 0, 0, 0);
            ca.AxisY.MajorGrid.LineColor = Color.LightGray;
            ca.AxisY.MinorGrid.LineColor = Color.LightGray;
            ca.AxisX.LabelStyle.Font = new Font("Arial", 10);
            ca.AxisX.IsLabelAutoFit = false;
            ca.AxisX.LabelStyle.Angle = -90;
            ca.AxisX.MajorTickMark.Enabled = false;
            ca.AxisY.LabelStyle.Enabled = false;
            ca.AxisY.LineColor = Color.FromArgb(0, 0, 0, 0);
            ca.AxisY.MajorTickMark.Enabled = false;
            ca.AxisY.Maximum = 50;

            chartEMU3.ChartAreas.Add(ca);

            // 5E Day
            Series dEMU3Series = new Series();
            dEMU3Series.Name = "EMU3 Day";
            dEMU3Series.ChartType = SeriesChartType.Column;
            dEMU3Series.XValueType = ChartValueType.String;
            dEMU3Series.YValueType = ChartValueType.Double;
            chartEMU3.Series.Add(dEMU3Series);

            double[] dEMU3Vals = dEMU3.ToArray();
            chartEMU3.Series["EMU3 Day"].Points.DataBindXY(years, dEMU3Vals);
            chartEMU3.Series["EMU3 Day"].Points[0].Color = Color.FromArgb(255, 36, 191, 54);
            try
            {
                chartEMU3.Series["EMU3 Day"].Points[1].Color = Color.FromArgb(255, 79, 178, 58);
            }
            catch { }
            try
            {
                chartEMU3.Series["EMU3 Day"].Points[2].Color = Color.FromArgb(255, 129, 185, 92);
            }
            catch { };
            // EMU3 Night
            Series nEMU3Series = new Series();
            nEMU3Series.Name = "EMU3 Night";
            nEMU3Series.ChartType = SeriesChartType.Column;
            nEMU3Series.XValueType = ChartValueType.String;
            nEMU3Series.YValueType = ChartValueType.Double;
            chartEMU3.Series.Add(nEMU3Series);

            double[] nEMU3Vals = nEMU3.ToArray();
            chartEMU3.Series["EMU3 Night"].Points.DataBindXY(years, nEMU3Vals);
            chartEMU3.Series["EMU3 Night"].Points[0].Color = Color.FromArgb(255, 24, 127, 36);
            try
            {
                chartEMU3.Series["EMU3 Night"].Points[1].Color = Color.FromArgb(255, 58, 124, 44);
            }
            catch { }
            try
            {
                chartEMU3.Series["EMU3 Night"].Points[2].Color = Color.FromArgb(255, 107, 144, 82);
            }
            catch { }

            // Save out the charts as images which will then be embedded in the email. 
            chartEMU0.SaveImage(@"C:\temp\EMU0Chart.jpg", ChartImageFormat.Jpeg);
            chartEMU1.SaveImage(@"C:\temp\EMU1Chart.jpg", ChartImageFormat.Jpeg);
            chartEMU2.SaveImage(@"C:\temp\EMU2Chart.jpg", ChartImageFormat.Jpeg);
            chartEMU3.SaveImage(@"C:\temp\EMU3Chart.jpg", ChartImageFormat.Jpeg);

            chartEMU0Img = new LinkedResource(@"C:\temp\EMU0Chart.jpg", MediaTypeNames.Image.Jpeg);
            chartEMU1Img = new LinkedResource(@"C:\temp\EMU1Chart.jpg", MediaTypeNames.Image.Jpeg);
            chartEMU2Img = new LinkedResource(@"C:\temp\EMU2Chart.jpg", MediaTypeNames.Image.Jpeg);
            chartEMU3Img = new LinkedResource(@"C:\temp\EMU3Chart.jpg", MediaTypeNames.Image.Jpeg);
        }
    }
}
