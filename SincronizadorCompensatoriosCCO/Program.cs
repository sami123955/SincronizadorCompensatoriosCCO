using SincronizadorCompensatoriosCCO.Modelo;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Web;


namespace SincronizadorCompensatoriosCCO
{
    class Program
    {

        private struct _Constantes
        {
            public const int Id_Motivo = 0;
            public const int MotivoDelCompensatorio = 1;
            public const int NombreDeUsuario = 2;
            public const int NumeroDeDias = 3;
            public const int ResponsableCarga = 4;
            public const int Observacion = 5;
            public const int FechaCompensnatorio = 6;
            public const String RegistroExcel = "Registro Excel";
            public const String MotivosCompensatoriosPorUsuario = "Motivos Compensatorio Por Usuario";
        }

        public bool ValidarRegistroNoVacio(JArray jsonArray)
        {
            if (jsonArray[_Constantes.Id_Motivo].Count() > 0
                || jsonArray[_Constantes.MotivoDelCompensatorio].Count() > 0
                || jsonArray[_Constantes.NombreDeUsuario].Count() > 0
                 || jsonArray[_Constantes.NumeroDeDias].Count() > 0
                 || jsonArray[_Constantes.ResponsableCarga].Count() > 0
                 || jsonArray[_Constantes.Observacion].Count() > 0
                 || jsonArray[_Constantes.FechaCompensnatorio].Count() > 0
                )
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public bool ValidarArchivo(JArray jsonArray)
        {
            int CantidadRegistros = jsonArray.Count();

            bool ArchivoCorrecto = true;
            for (int i = 0; i < CantidadRegistros; i++)
            {
                if (this.ValidarRegistroNoVacio((JArray)jsonArray[i]))
                {
                    if (jsonArray[i][_Constantes.Id_Motivo].Count() <= 0)
                    {
                        ArchivoCorrecto = false;
                        break;
                    }
                    if (jsonArray[i][_Constantes.MotivoDelCompensatorio].Count() <= 0)
                    {
                        ArchivoCorrecto = false;
                        break;
                    }
                    if (jsonArray[i][_Constantes.NombreDeUsuario].Count() <= 0)
                    {
                        ArchivoCorrecto = false;
                        break;
                    }
                    if (jsonArray[i][_Constantes.NumeroDeDias].Count() <= 0)
                    {
                        ArchivoCorrecto = false;
                        break;
                    }
                    if (jsonArray[i][_Constantes.ResponsableCarga].Count() <= 0)
                    {
                        ArchivoCorrecto = false;
                        break;
                    }
                    if (jsonArray[i][_Constantes.Observacion].Count() <= 0)
                    {
                        ArchivoCorrecto = false;
                        break;
                    }
                    if (jsonArray[i][_Constantes.FechaCompensnatorio].Count() <= 0)
                    {
                        ArchivoCorrecto = false;
                        break;
                    }
                }
                else
                {
                    break;
                }

            }
            return ArchivoCorrecto;
        }

        public string ReuestExcel(ListItem item, int ValorFilaInicio, int ValorFilaFin, SharePointOnlineCredentials cred)
        {
            string urlServicio = "https://sistemasinteligentesenred.sharepoint.com/sites/Intranet/GestionCompensatorios/_vti_bin/ExcelRest.aspx/"
                    + "CargasExcel/"
                    + HttpUtility.UrlEncode(item.FieldValues["FileLeafRef"].ToString())
                    + "/Model/Ranges(" + HttpUtility.UrlEncode("'A" + ValorFilaInicio + "|G" + ValorFilaFin + "'") + ")?$format=json";
            String url = urlServicio.Replace("+", " ");
            WebRequest req = (WebRequest)WebRequest.Create(url);

            req.Credentials = cred;
            req.Headers["X-FORMS_BASED_AUTH_ACCEPTED"] = "f";
            HttpWebResponse response = (HttpWebResponse)req.GetResponse();
            Stream receiveStream = response.GetResponseStream();
            StreamReader readStream = new StreamReader(receiveStream, Encoding.UTF8);
            string Json = readStream.ReadToEnd();
            response.Close();
            readStream.Close();
            return Json;
        }

        static void Main(string[] args)
        {

            /*Credenciales para traer la informacion desde el sharepoint*/
            string Usuario = "ssantamaria@conocimientocorporativo.com";
            string Clave = "Santamaria123";
            /*Variable de seguridad para poder loguiarse en el sharepoint*/
            SecureString SecurePassword = new SecureString();
            char[] VecClave = Clave.ToCharArray();
            /*recorro la clave y agrego la letra a la seguridad*/
            foreach (char letra in VecClave)
            {
                SecurePassword.AppendChar(Convert.ToChar(letra));
            }
            /*me logue con share point*/
            SharePointOnlineCredentials cred = new SharePointOnlineCredentials(Usuario, SecurePassword);
            /*Creo el contexto al sitio de sharepoint*/
            var URLSir = "https://conocimiento.sharepoint.com/teams/dev/GestionCompenstorios";
            //ClientContext context = new ClientContext(@"https://sistemasinteligentesenred.sharepoint.com/sites/Intranet/GestionCompensatorios");
            ClientContext context = new ClientContext(@"https://conocimiento.sharepoint.com/teams/dev/GestionCompenstorios");
            context.Credentials = cred;
            List listCargaExcel = context.Web.Lists.GetByTitle("Cargas Excel");
            //CamlQuery query = CamlQuery.CreateAllItemsQuery(100);






            /*Filtro los datos que voy a consultar*/
            CamlQuery query = new CamlQuery();
            query.ViewXml = "<View>" +
              "<Query>" +
                "<Where>" +
                  "<Eq>" +
                    "<FieldRef Name='Estado'/>" +
                    "<Value Type='Text'>Pendiente</Value>" +
                  "</Eq>" +
                "</Where>" +
              "</Query>" +
              "</View>";

            ListItemCollection itemsPedidos = listCargaExcel.GetItems(query);
            context.Load(itemsPedidos);
            context.ExecuteQuery();

            int ValorFilaInicio = 2;
            int ValorFilaFin = 0;
            bool TerminoArchivo = false;
            int NumeroFilasRecorer = 100;

            /*recorro las lista con todos los */
            foreach (ListItem item in itemsPedidos)
            {
                int valorCampo = 0;
                do
                {
                    valorCampo = 0;
                    ValorFilaFin = ValorFilaInicio + NumeroFilasRecorer;

                    Program objProgram = new Program();
                    string Json = objProgram.ReuestExcel(item, ValorFilaInicio, ValorFilaFin, cred);
                    JObject objJson = JObject.Parse(Json);
                    JArray objJARow = (JArray)objJson["rows"];
                    RegistrosExcel objRegistrosExcel = new RegistrosExcel(context, _Constantes.RegistroExcel);
                    MotivosCompensatorios objMotivosCompensatorios = new MotivosCompensatorios(context, _Constantes.MotivosCompensatoriosPorUsuario);
                    if (objProgram.ValidarArchivo(objJARow))
                    {
                        while ((valorCampo < NumeroFilasRecorer) && objProgram.ValidarRegistroNoVacio((JArray)objJARow[valorCampo]))
                        {
                            int Id_Motivo = (int)objJARow[valorCampo][_Constantes.Id_Motivo]["v"];
                            string MotivoDelCompensatorio = (string)objJARow[valorCampo][_Constantes.MotivoDelCompensatorio]["v"]; ;
                            string NombreDeUsuario = (string)objJARow[valorCampo][_Constantes.NombreDeUsuario]["v"];
                            decimal NumeroDeDias = Convert.ToDecimal(objJARow[valorCampo][_Constantes.NumeroDeDias]["v"]);
                            string ResponsableCarga = (string)objJARow[valorCampo][_Constantes.ResponsableCarga]["v"];
                            string Observacion = (string)objJARow[valorCampo][_Constantes.Observacion]["v"];
                            string FechaCompensnatorio = (string)objJARow[valorCampo][_Constantes.FechaCompensnatorio]["fv"];
                            int IdCargaExcel = (int)item.FieldValues["ID"];

                            objRegistrosExcel.GuardarRegistroExcel(Id_Motivo
                                                                   , MotivoDelCompensatorio
                                                                   , NombreDeUsuario
                                                                   , NumeroDeDias
                                                                   , ResponsableCarga
                                                                   , Observacion
                                                                   , IdCargaExcel
                                                                   , FechaCompensnatorio);
                            objMotivosCompensatorios.GuardarMotivosCompensatorios(Id_Motivo, MotivoDelCompensatorio, NombreDeUsuario, NumeroDeDias);

                            valorCampo++;
                            ValorFilaInicio++;
                        }
                        if (objJARow[valorCampo][0].Count() <= 0)
                        {
                            TerminoArchivo = true;
                            item["Estado"] = "Procesado";
                            item["Observacion"] = "";
                            item.Update();
                            context.ExecuteQuery();
                        }
                    }
                    else
                    {
                        item["Estado"] = "Procesado parcialmente";
                        item["Observacion"] = "El archivo tiene celdas vacias o no tiene el formato correcto";
                        item.Update();
                        context.ExecuteQuery();
                        TerminoArchivo = true;
                    }

                } while (!TerminoArchivo);


            }

        }
    }
}
