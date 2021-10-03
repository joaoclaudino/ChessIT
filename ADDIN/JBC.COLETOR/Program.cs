using Castle.Core.Logging;
using JBC.Coletor.Controller;
using JBC.Framework.Attribute;
using JBC.Framework.DAO;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
//using System.Windows.Forms;

namespace JBC.Coletor
{
    [AddIn(
        Name = "JBC.Coletor"
        , Description = "JBC.Coletor - Claudino Soluções em Software"
        , Namespace = "JBC"
        , InitMethod = "CheckInit"
        )]
    class Program
    {
        public SAPbobsCOM.Company oCompany { get; set; }
        public SAPbouiCOM.Application oApplication { get; set; }

        public static SAPbobsCOM.Company oCompanyS { get; set; }
        public static SAPbouiCOM.Application oApplicationS { get; set; }

        public BusinessOneDAO oB1DAO { get; set; }
        public ILogger Log { get; set; }

        static void Main(string[] args)
        {

        }

        public void CheckInit()
        {
            oCompanyS = oCompany;
            oApplicationS = oApplication;
            Form sboForm = oApplication.Forms.GetFormByTypeAndCount(169, 1);

            sboForm.Freeze(true);
            try
            {

                var assembly = Assembly.GetExecutingAssembly();
                var resourceName = "JBC.Coletor.xml.Menu.xml";
                //string response = System.Resources.GetString(resourceName);
                using (System.IO.Stream stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                {
                    using (System.IO.FileStream fileStream = new System.IO.FileStream(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Menu.xml"), System.IO.FileMode.Create))
                    {
                        for (int i = 0; i < stream.Length; i++)
                        {
                            fileStream.WriteByte((byte)stream.ReadByte());
                        }
                        fileStream.Close();
                    }
                }

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Menu.xml");

                string xml = xmlDoc.InnerXml.ToString();
                oApplication.LoadBatchActions(ref xml);
            }
            catch (Exception exception)
            {
                oApplication.StatusBar.SetText("Erro ao criar menu: " + exception.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                Environment.Exit(-1);
            }
            finally
            {
                sboForm.Freeze(false);
                sboForm.Update();
            }
            //MessageBox.Show("123");

            //ThreadStart threadDelegate = new ThreadStart(this.processo);
            //Thread newThread = new Thread(threadDelegate);
            //newThread.Start();



            oApplication.AppEvent += HandleAppEvent;
            oApplication.MenuEvent += HandleMenuEvent;
            oApplication.ItemEvent += HandleFormLoadEvent;
            oApplication.ItemEvent += HandleButtonClickEvent;
        }

        private void HandleAppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    {
                        System.Windows.Forms.Application.Exit();
                    }
                    break;
            }
        }

        private void HandleMenuEvent(ref MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;

            if (pVal.MenuUID.Equals("FrmGeraOS") && !pVal.BeforeAction)
            {
                try
                {


                    var resourceName = "JBC.Coletor.SrfFiles.FrmGeraOS.srf";
                    //string response = System.Resources.GetString(resourceName);
                    using (System.IO.Stream stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                    {
                        using (System.IO.FileStream fileStream = new System.IO.FileStream(System.IO.Path.Combine(System.Environment.CurrentDirectory, "FrmGeraOS.srf"), System.IO.FileMode.Create))
                        {
                            for (int i = 0; i < stream.Length; i++)
                            {
                                fileStream.WriteByte((byte)stream.ReadByte());
                            }
                            fileStream.Close();
                        }
                    }
                    string srfPath = System.Environment.CurrentDirectory + "\\FrmGeraOS.srf";

                    if (File.Exists(srfPath) == false)
                    {
                        throw new Exception("Arquivo SRF não encontrado. Verifique a instalação do addOn.");
                    }

                    string xml = File.ReadAllText(srfPath);

                    string formUID = GerarFormUID("FrmGeraOS");

                    xml = xml.Replace("uid=\"FrmGeraOS\"", string.Format("uid=\"{0}\"", formUID));

                    //#if DEBUG
                    //                    xml = xml.Replace("from dummy", "");
                    //#endif

                    oApplication.LoadBatchActions(ref xml);
                }
                catch (Exception exception)
                {
                    oApplication.StatusBar.SetText(exception.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }

            if (pVal.MenuUID.Equals("FrmCnfIntegraBalanca") && !pVal.BeforeAction)
            {
                try
                {
                    var resourceName = "JBC.Coletor.SrfFiles.FrmCnfIntegraBalanca.srf";
                    //string response = System.Resources.GetString(resourceName);
                    using (System.IO.Stream stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                    {
                        using (System.IO.FileStream fileStream = new System.IO.FileStream(System.IO.Path.Combine(System.Environment.CurrentDirectory, "FrmCnfIntegraBalanca.srf"), System.IO.FileMode.Create))
                        {
                            for (int i = 0; i < stream.Length; i++)
                            {
                                fileStream.WriteByte((byte)stream.ReadByte());
                            }
                            fileStream.Close();
                        }
                    }

                    string srfPath = System.Environment.CurrentDirectory + "\\FrmCnfIntegraBalanca.srf";

                    if (File.Exists(srfPath) == false)
                    {
                        throw new Exception("Arquivo SRF não encontrado. Verifique a instalação do addOn.");
                    }

                    string xml = File.ReadAllText(srfPath);

                    string formUID = GerarFormUID("FrmCnfIntegraBalanca");

                    xml = xml.Replace("uid=\"FrmCnfIntegraBalanca\"", string.Format("uid=\"{0}\"", formUID));

#if DEBUG
                    xml = xml.Replace("from dummy", "");
#endif

                    oApplication.LoadBatchActions(ref xml);
                }
                catch (Exception exception)
                {
                    oApplication.StatusBar.SetText(exception.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
        }

        private void HandleFormLoadEvent(string formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;

            if (pVal.EventType == BoEventTypes.et_FORM_LOAD && !pVal.BeforeAction)
            {
                if (pVal.FormTypeEx == "FrmGeraOS")
                {
                    Form form = oApplication.Forms.Item(pVal.FormUID);

                    new View.GeraOSView(form);
                }

                if (pVal.FormTypeEx == "FrmCnfIntegraBalanca")
                {
                    Form form = oApplication.Forms.Item(pVal.FormUID);

                    new View.CnfIntegraBalancaView(form);
                }

                if (pVal.FormTypeEx == "FrmNotasGeradas")
                {
                    Form form = oApplication.Forms.Item(pVal.FormUID);

                    new View.NotasGeradasView(form, m_NotasGeradas);
                }

                if (pVal.FormTypeEx == "139") //Pedido de Venda (OS)
                {
                    Form form = oApplication.Forms.Item(pVal.FormUID);

                    Item _stParPeso = form.Items.Add("stParPeso", BoFormItemTypes.it_STATIC);
                    _stParPeso.Top = form.Items.Item("1980000501").Top - 70;
                    _stParPeso.Left = form.Items.Item("2001").Left + 450;
                    _stParPeso.Width = 100;
                    _stParPeso.LinkTo = "2001";

                    StaticText stParPeso = (StaticText)_stParPeso.Specific;
                    stParPeso.Caption = "Parâmetro Pesagem";

                    Item _cbParPeso = form.Items.Add("cbParPeso", BoFormItemTypes.it_COMBO_BOX);
                    _cbParPeso.Top = _stParPeso.Top + 20;
                    _cbParPeso.Left = _stParPeso.Left;
                    _cbParPeso.Width = 100;
                    _cbParPeso.LinkTo = "2001";
                    _cbParPeso.DisplayDesc = true;

                    ComboBox cbParPeso = (ComboBox)_cbParPeso.Specific;
                    cbParPeso.ValidValues.Add("0", "Peso Bruto");
                    cbParPeso.ValidValues.Add("1", "Tara");

                    cbParPeso.Select("0");

                    Item _btCapPeso = form.Items.Add("btCapPeso", BoFormItemTypes.it_BUTTON);
                    _btCapPeso.Top = _cbParPeso.Top + 20;
                    _btCapPeso.Left = _cbParPeso.Left;
                    _btCapPeso.Width = 100;
                    _btCapPeso.LinkTo = "2001";

                    Button btEtiq = (Button)_btCapPeso.Specific;
                    btEtiq.Caption = "Capturar Peso";

                    form.Mode = BoFormMode.fm_OK_MODE;
                }
            }
        }

        private void HandleButtonClickEvent(string formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;



            //Pedido de Venda (OS)
            if (pVal.EventType == BoEventTypes.et_CLICK && pVal.FormTypeEx == "139" && pVal.ItemUID == "btCapPeso" && !pVal.BeforeAction)
            {
                bool sucesso = false;
                string mensagem = "";

                oApplication.StatusBar.SetText("Lendo peso...", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);

                ProgressBar progressBar = oApplication.StatusBar.CreateProgressBar("Lendo peso...", 100, false);

                string ip = "";
                string porta = "";
                string diretorio = "";
                string nomeArquivo = "";

                try
                {
                    SAPbobsCOM.Recordset recordSet = null;
                    try
                    {
                        string query = @"SELECT * FROM ""@INTCFG"" ";

                        recordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        recordSet.DoQuery(query);

                        if (!recordSet.EoF)
                        {
                            ip = recordSet.Fields.Item("U_IP").Value.ToString();
                            porta = recordSet.Fields.Item("U_Porta").Value.ToString();
                            diretorio = recordSet.Fields.Item("U_Diretorio").Value.ToString();
                            nomeArquivo = recordSet.Fields.Item("U_NomeArquivo").Value.ToString();
                        }
                        else throw new Exception("Dados para conexão não configurados");
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Erro ao buscar dados de conexão: " + ex.Message);
                    }
                    finally
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(recordSet);
                        GC.Collect();
                    }

                    progressBar.Value = 10;

                    Form form = oApplication.Forms.Item(pVal.FormUID);

                    Matrix matrix = (Matrix)form.Items.Item("38").Specific;

                    string parametroPesagem = ((ComboBox)form.Items.Item("cbParPeso").Specific).Selected.Value;

                    string hora = DateTime.Now.ToString("HH:mm:ss");

                    progressBar.Value = 40;

                    string caminho = System.IO.Path.Combine(diretorio, nomeArquivo);

                    try
                    {
                        string arquivo = Helper.ApiHelper.GetArquivo(ip, porta, caminho);

                        if (arquivo == "ARQUIVO NÃO ENCONTRADO")
                        {
                            throw new Exception("Falha na comunicaçao com a balança capital: arquivo não encontrado");
                        }
                        else
                        {
                            progressBar.Value = 80;

                            StringBuilder sb = new StringBuilder(arquivo);

                            bool linhaEncontrada = false;

                            string[] lines = arquivo.Split(
                                new[] { @"\r\n\" },
                                StringSplitOptions.RemoveEmptyEntries
                            );

                            foreach (string line in lines)
                            {
                                string clearedLine = line.Replace("\\", "").Replace(" ", "").Replace("\"", "").Trim();

                                if (clearedLine.Substring(6, 8) == hora)
                                {
                                    linhaEncontrada = true;

                                    string peso = clearedLine.Substring(0, 6);

                                    if (peso.Equals(string.Empty))
                                        throw new Exception("Não há peso tara definido para OS em questão");

                                    string passo = "0";

                                    try
                                    {

                                        string um = "";

                                        try
                                        {
                                            um = ((EditText)matrix.Columns.Item("212").Cells.Item(1).Specific).String.ToUpper();
                                        }
                                        catch { }

                                        passo = "1";

                                        string sPesoBruto = ((EditText)form.Items.Item("U_PesoBruto").Specific).String;
                                        string sTara = ((EditText)form.Items.Item("U_Tara").Specific).String;

                                        passo = "2";

                                        double pesoBruto = double.Parse((sPesoBruto.Contains(",") ? sPesoBruto.Replace(".", "").Replace(",", ".") : sPesoBruto), System.Globalization.CultureInfo.InvariantCulture); ;
                                        double tara = double.Parse((sTara.Contains(",") ? sTara.Replace(".", "").Replace(",", ".") : sTara), System.Globalization.CultureInfo.InvariantCulture); ;

                                        passo = "3";

                                        if (parametroPesagem.Equals("0"))
                                        {
                                            passo = "4";

                                            pesoBruto = double.Parse((peso.Contains(",") ? peso.Replace(".", "").Replace(",", ".") : peso), System.Globalization.CultureInfo.InvariantCulture);

                                            ((EditText)form.Items.Item("U_PesoBruto").Specific).String = pesoBruto.ToString();
                                        }
                                        else
                                        {
                                            passo = "5";

                                            tara = double.Parse((peso.Contains(",") ? peso.Replace(".", "").Replace(",", ".") : peso), System.Globalization.CultureInfo.InvariantCulture);

                                            ((EditText)form.Items.Item("U_Tara").Specific).String = tara.ToString();
                                        }

                                        passo = "6";

                                        if (tara > 0)
                                        {
                                            passo = "7";

                                            double pesoLiquido = pesoBruto - tara;

                                            ((EditText)form.Items.Item("U_PesoLiq").Specific).String = pesoLiquido.ToString();

                                            if (um == "TONELADAS")
                                            {
                                                passo = "8";

                                                ((EditText)matrix.Columns.Item("11").Cells.Item(1).Specific).String = (pesoLiquido / 1000).ToString();
                                            }
                                        }

                                        passo = "9";

                                        ((EditText)form.Items.Item("U_HoraEntradaOS").Specific).String = hora;
                                        ((ComboBox)form.Items.Item("U_Situacao").Specific).Select("11");
                                        ((ComboBox)form.Items.Item("U_Status").Specific).Select("P");
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception("Erro ao preencher campos em tela (etapa " + passo + ") [ " + peso + "] :" + ex.Message);
                                    }

                                    sucesso = true;

                                    break;

                                }
                            }

                            if (!linhaEncontrada)
                                throw new Exception("Não há peso registrado na balança para o horário " + hora);
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Falha na comunicaçao com a balança capital: " + ex.Message);
                    }

                    progressBar.Value = 100;

                }
                catch (Exception ex)
                {
                    mensagem = ex.Message;
                }
                finally
                {
                    progressBar.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(progressBar);
                    GC.Collect();
                }

                if (sucesso)
                {
                    oApplication.StatusBar.SetText("Operação completada com êxito", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
                }
                else
                {
                    oApplication.StatusBar.SetText(mensagem);
                }
            }
        }

        public static List<int> m_NotasGeradas = new List<int>();

        public static void OpenNotasGeradasView(List<int> notas)
        {
            try
            {
                m_NotasGeradas = notas;

                var resourceName = "JBC.Coletor.SrfFiles.FrmNotasGeradas.srf";
                //string response = System.Resources.GetString(resourceName);
                using (System.IO.Stream stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                {
                    using (System.IO.FileStream fileStream = new System.IO.FileStream(System.IO.Path.Combine(System.Environment.CurrentDirectory, "FrmNotasGeradas.srf"), System.IO.FileMode.Create))
                    {
                        for (int i = 0; i < stream.Length; i++)
                        {
                            fileStream.WriteByte((byte)stream.ReadByte());
                        }
                        fileStream.Close();
                    }
                }

                string srfPath = System.Environment.CurrentDirectory + "\\FrmNotasGeradas.srf";

                if (File.Exists(srfPath) == false)
                {
                    throw new Exception("Arquivo SRF não encontrado. Verifique a instalação do addOn.");
                }

                string xml = File.ReadAllText(srfPath);

                string formUID = GerarFormUID("FrmNotasGeradas");

                xml = xml.Replace("uid=\"FrmNotasGeradas\"", string.Format("uid=\"{0}\"", formUID));

#if DEBUG
                xml = xml.Replace("from dummy", "");
#endif

                oApplicationS.LoadBatchActions(ref xml);
            }
            catch (Exception exception)
            {
                oApplicationS.StatusBar.SetText(exception.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        private static int ExecuteCommand(string command, int timeout)
        {
            var processInfo = new ProcessStartInfo("cmd.exe", "/C " + command)
            {
                CreateNoWindow = true,
                UseShellExecute = false,
                WorkingDirectory = "C:\\",
            };

            var process = Process.Start(processInfo);
            process.WaitForExit(timeout);
            var exitCode = process.ExitCode;
            process.Close();
            return exitCode;
        }

        private  static string GerarFormUID(string formType)
        {
            string result = string.Empty;

            int count = 0;

            bool next = true;

            while (next)
            {
                count++;

                try
                {
                    oApplicationS.Forms.GetForm(formType, count);
                }
                catch
                {
                    next = false;
                }
            }

            result = string.Format("Frm{0}-{1}", count, new Random().Next(999));

            return result;
        }

        public static void LimparObjeto(Object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);

            }
            catch { }
            try
            {
                obj = null;
            }
            catch { }
            GC.Collect();
            GC.WaitForFullGCComplete();

        }

        //public void processo()
        //{
        //    MainController mainController = new MainController(oApplication, oCompany);
        //}
    }
}
