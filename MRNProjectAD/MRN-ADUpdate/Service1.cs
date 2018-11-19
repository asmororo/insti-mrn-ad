using Progress.Open4GL.Proxy;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

using System.Timers;
using System.Configuration;
using activedll.StrongTypesNS;
using System.DirectoryServices;

namespace MRN_ADUpdate
{
    public partial class Service1 : ServiceBase
    {
        private const string dominio = "mrnptr.com.br";
        private const string adUser = "gfi77";
        private const string adPassword = "Brasil1802";

        private Timer timer = new Timer();
        private string cnAppServer = ConfigurationManager.AppSettings["cnAppServer"].ToString();
        private string appUser = ConfigurationManager.AppSettings["appUser"].ToString();
        private string appPassword = ConfigurationManager.AppSettings["appPassword"].ToString();
        private string appServerInfo = ConfigurationManager.AppSettings["appServerInfo"].ToString();

        public Service1()
        {
            InitializeComponent();
        }

        private Connection AppConection()
        {
            try
            {
                Connection cn = new Connection(cnAppServer, appUser, appPassword, appServerInfo);
                cn.SessionModel = Convert.ToInt32(ConfigurationManager.AppSettings["SessionMode"]);
                //cn.TcpClientRetry = Convert.ToInt32(ConfigurationManager.AppSettings["appTcpClientRetry"]);
                //cn.SocketTimeout = Convert.ToInt32(ConfigurationManager.AppSettings["appSocketTimeout"]);
                //cn.TcpClientRetryInterval = Convert.ToInt32(ConfigurationManager.AppSettings["appTcpClientRetryInterval"]);
                //cn.RequestWaitTimeout = Convert.ToInt32(ConfigurationManager.AppSettings["appRequestWaitTimeout"]);
               
                return cn;
            }
            catch
            {
                return null;
            }
        }

        public void espend001()
        {
            ttParamDataTable ttParam = new ttParamDataTable();
            ttPendenteDataTable ttPendente = new ttPendenteDataTable();
            ttRetDataTable ttRet = new ttRetDataTable();

            Connection cn = new Connection(cnAppServer, appUser, appPassword, appServerInfo);

            DirectoryEntry entry = new DirectoryEntry("LDAP://" + dominio, adUser, adPassword);

            try
            {
                cn = AppConection();
                if (cn == null)
                    throw new Exception("Erro de conexão com AppServer");

                ttParam.AddttParamRow(Convert.ToDateTime("01/01/1900"), Convert.ToDateTime("01/01/2200"), " ", "ZZZZZZZZZZZZZZZZ", " ", "ZZZZZZZZZZZZZZZZ", " ", "ZZZZZZZZZZZZZZZZ");

                activedll.active activeMrn = new activedll.active(cn);

                string cret = activeMrn.espend001(ttParam, out ttPendente, out ttRet);
                
                if (cret != "OK")
                {
                    foreach (DataRow item in ttRet.Rows)
                    {
                        WriteToFile($"Erro: {item["_c_DESC"].ToString()}");
                    }
                }
                else
                {
                    object nativeObject = entry.NativeObject;

                    DirectorySearcher search = new DirectorySearcher(entry);

                    foreach (DataRow row in ttPendente.Rows)
                    {
                        search.Filter = $"(SAMAccountName= { row["usuario_ad"] })";

                        SearchResult result = search.FindOne();
                        DirectoryEntry userEntry = result.GetDirectoryEntry();

                        //int old_UAC = (int)userEntry.Properties["userAccountControl"][0];

                        // AD user account disable flag
                        //int ADS_UF_ACCOUNTDISABLE = 2;

                        // To disable an ad user account, we need to set the disable bit/flag:
                        //userEntry.Properties["userAccountControl"][0] = (old_UAC | ADS_UF_ACCOUNTDISABLE);
                        //userEntry.CommitChanges();

                        WriteToFile($"Usuário: {row["usuario_ad"]}");
                    }

                    WriteToFile("OK");
                }
            }
            catch (Exception ex)
            {
                WriteToFile(ex.Message);
            }
            finally
            {
                cn.ReleaseConnection();
                cn.Dispose();
            }
        }

        protected override void OnStart(string[] args)
        {
            WriteToFile("Service is started at " + DateTime.Now);
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            timer.Interval = 50000;
            timer.Enabled = true;
        }

        protected override void OnStop()
        {
            WriteToFile("Service is stopped at " + DateTime.Now);
        }

        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            WriteToFile("Service is recall at " + DateTime.Now);

            try
            {
                //DirectoryEntry entry = new DirectoryEntry("LDAP://" + dominio, adUser, adPassword);
                //object nativeObject = entry.NativeObject;

                //DirectorySearcher search = new DirectorySearcher(entry);
                //search.Filter = $"(SAMAccountName=diego.santos)";
                //search.PropertiesToLoad.Add("mail");

                //SearchResult result = search.FindOne();

                //WriteToFile($"Email AD: {result.Properties["mail"][0].ToString()} ");

                espend001();
            }
            catch (Exception ex)
            {
                WriteToFile(ex.Message);
            }

            //espend001();
        }

        public void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + @"\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            string filePath = AppDomain.CurrentDomain.BaseDirectory + @"\Logs\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filePath))
            {
                //Create a file to write to.
                using (StreamWriter sw = File.CreateText(filePath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filePath))
                {
                    sw.WriteLine(Message);
                }
            }
        }
    }
}
