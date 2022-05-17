using System;
using System.Linq;
using System.IO;
using System.Xml.Linq;
using DIAPI = SAPbobsCOM;
using System.Windows.Forms;

namespace ModelodeAprov.Controller
{
   public class ConectaSAP
    {


        public static DIAPI.Company  oCompany= new DIAPI.Company();

        public static bool ConectaSap( string user , string password)
        {

            string sXmlFileName = Path.GetFileName("config.xml");
            var xmlconfigfile = XDocument.Load(sXmlFileName);
            var xmlconfig = from d in xmlconfigfile.Root.Descendants("config")
                            select new
                            {
                                Licenseserver = d.Element("Licenseserver").Value,
                                ServerSAP = d.Element("ServerSAP").Value,
                                userDB = d.Element("userDB").Value,
                                passwordDB = d.Element("passwordDB").Value,
                                tiposerver = d.Element("tiposerver").Value,
                                Empresa = d.Element("Empresa").Value,
                                UserSAP = user, // d.Element("UserSAP").Value,
                                SenhaSAP =password //d.Element("SenhaSAP").Value
                            };


            foreach (var DadosConfig in xmlconfig)
            {
                oCompany.SLDServer = DadosConfig.Licenseserver;
                oCompany.Server = DadosConfig.ServerSAP;
                oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Portuguese_Br;
                if(DadosConfig.tiposerver == "dst_MSSQL2017")
                {
                    oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2017;

                }
                if(DadosConfig.tiposerver == "dst_MSSQL2014")
                {
                    oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;

                }

             
                oCompany.DbUserName = DadosConfig.userDB;
                oCompany.DbPassword = DadosConfig.passwordDB;
                oCompany.CompanyDB = DadosConfig.Empresa;
                oCompany.UserName = DadosConfig.UserSAP;
                oCompany.Password = DadosConfig.SenhaSAP;
               
            }

            int RetValSAP;
            
            RetValSAP = oCompany.Connect();
            if (RetValSAP != 0)
            {
                string ErrMsg = oCompany.GetLastErrorDescription();
                MessageBox.Show(ErrMsg);
                Escrevelog(ErrMsg);
            }
            else
            {
                Escrevelog("Conectado no Sap");
            }
            return oCompany.Connected;
        }

        public static void Escrevelog(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }


    }
}
