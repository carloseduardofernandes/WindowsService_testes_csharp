using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;

namespace WindowsService
{
    public partial class Service1 : ServiceBase
    {
        //public static EventLog _EventLog { get; set; }//static public para usar de fora
        private System.Timers.Timer timer13;
        private System.Timers.Timer timer342;

        private Thread threadMain;

        public Service1()
        {
            InitializeComponent();

           // Inicializar();
        }

        protected override void OnStart(string[] args)
        {
            Inicializar();
        }

        protected override void OnStop()
        {
            base.OnStop();
            //eventLog.WriteEntry("In OnStop.");
        }

        public void Inicializar()
        {
            try
            {
                //le excell e pega registros forma string par ainsert com aspas simples e virgula

                /*   var listA = new List<string>();
                 var stringQuery = "";

                 using (var reader = new StreamReader(@"C:\\Users\\caefernandes\\Downloads\\Base Radar de Proximidade Call Center.csv"))
                 {
                     while (!reader.EndOfStream)
                     {
                         var line = reader.ReadLine();
                         var values = line.Split(';');

                         listA.Add($"('{values[0]}')");
                     }
                 }

                 var qtd = listA.Count;

                 stringQuery = string.Join(",", listA);

                var fileName = string.Format("C:\\Users\\caefernandes\\Downloads\\Base Radar de Proximidade Call Center.xlsx");
                  var connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", fileName);

                  var adapter = new OleDbDataAdapter("SELECT * FROM [planilha1$]", connectionString);
                  var ds = new DataSet();

                  adapter.Fill(ds, "anyNameHere");

                  DataTable data = ds.Tables["anyNameHere"];

                  var path = Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory);
                  */
                //_EventLog = new System.Diagnostics.EventLog();

                //if (!System.Diagnostics.EventLog.SourceExists("LogServico"))
                //{
                //    System.Diagnostics.EventLog.CreateEventSource(
                //        "LogServico", "LogServico");
                //}

                //EventLog.Source = "LogServico";
                //EventLog.Log = "LogServico";

                //_EventLog = EventLog;

                //EventLog.WriteEntry("In OnStart.");

                //timer13 = new System.Timers.Timer();
                //timer13.Interval = 60000; // 60 seconds
                //timer13.Elapsed += new ElapsedEventHandler((sender, e) => Executa13(sender, e));
                //timer13.AutoReset = false;//desativa o auto fire, para nao correr risco de "encavalar"
                //timer13.Enabled = true;

                //timer342 = new System.Timers.Timer();
                //timer342.Interval = 60000; // 60 seconds
                //timer342.Elapsed += new ElapsedEventHandler(this.Executa342);
                //timer342.AutoReset = false;//desativa o auto fire, para nao correr risco de "encavalar"
                //timer342.Enabled = true;

                //threadMain = new Thread(new ThreadStart(ThreadMain));
                //threadMain.SetApartmentState(ApartmentState.STA);
                //threadMain.Start();
            }
            catch (Exception e)
            {
                throw e;
                // eventLog.WriteEntry("Erro ao inicializar o serviço:" + e.Message);
            }
        }

        [STAThread]
        private void ThreadMain()
        {
            Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.InvariantCulture;

            //somente para manter o serviço rodando o tempo todo
            while (true)
            {
                try
                {

                }
                catch (Exception e)
                {
                   // eventLog.WriteEntry("Erro no serviço:" + e.Message);
                }
                finally
                {
                    GC.Collect();
                    Thread.Sleep(2000);
                }
            }
        }

        private void Executa13(object sender, ElapsedEventArgs args)
        {
            var thread = new Thread(Executa13Thread);
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            timer13.Stop();//para o timer
        }

        [STAThread]
        private void Executa13Thread(object state)
        {
            try
            {
                //eventLog.WriteEntry("HORA DE RODAR A 13");
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                Thread.Sleep(2000);
                timer13.Start();//inicia novamente somente apos finalizar as tarefas acima
            }
        }

        private void Executa342(object sender, ElapsedEventArgs args)
        {
            var thread = new Thread(Executa342Thread);
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            timer342.Stop();//para o timer
        }

        [STAThread]
        private void Executa342Thread(object state)
        {
            try
            {
               // eventLog.WriteEntry("HORA DE RODAR A 342");
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                Thread.Sleep(2000);
                timer342.Start();//inicia novamente somente apos finalizar as tarefas acima
            }
        }
    }
}
