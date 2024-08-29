using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Integration_IROUTE.App
{
    public class Globals
    {
        public static String ShortName = "(EXX)";
        public static int continuar = -1;
        public static string Query = null;
        public static SAPbobsCOM.Recordset oRec = default(SAPbobsCOM.Recordset);
        public static SAPbouiCOM.Application SBO_Application;
        public static SAPbobsCOM.CompanyService oCmpSrv;
        public static SAPbouiCOM.EventFilters oFilters;
        public static SAPbouiCOM.EventFilter oFilter;
        public static SAPbobsCOM.Company oCompany;
        public static SAPbobsCOM.Company oCompanyMirror;
        public static int SAPVersion;
        public static string Addon = null;
        public static string version = null;
        public static string oldversion = "";
        public static bool Actual = false;
        public static int lRetCode;
        public static int sErrCode;
        public static string sErrMsg = null;
        public static string Error = null;

        public static string UrlRoute = null;
        public static string UserRoute = null;
        public static string PasswordRoute = null;

        public const string LinkedSystemObject = "1";
        public const string LinkedUDO = "2";
        public const string LinkedTable = "3";

        public static object Release(object objeto)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objeto);
            Query = null;
            GC.Collect();
            return null;
        }

        public static void ErrorMessage(String msg)
        {
            Globals.SBO_Application.SetStatusBarMessage(ShortName + ": " + msg, SAPbouiCOM.BoMessageTime.bmt_Short);
        }

        public static void InformationMessage(String msg)
        {
            Globals.SBO_Application.SetStatusBarMessage(ShortName + ": " + msg, SAPbouiCOM.BoMessageTime.bmt_Short, false);
        }

        public static void SuccessMessage(String msg)
        {
            Globals.SBO_Application.SetStatusBarMessage(ShortName + ": " + msg, SAPbouiCOM.BoMessageTime.bmt_Short, false);
        }

        public static void MessageBox(String msg)
        {
            Globals.SBO_Application.MessageBox(msg);
        }

        public static SAPbobsCOM.Recordset RunQuery(string Query)
        {
            try
            {
                oRec = (SAPbobsCOM.Recordset)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRec.DoQuery(Query);
                return oRec;
            }
            catch (Exception ex)
            {
                Globals.SBO_Application.MessageBox(ex.Message);
                return null;
            }
        }

        public static string LoadFromXML(ref string FileName)
        {
            System.Xml.XmlDocument oXmlDoc = null;
            string sPath = null;
            oXmlDoc = new System.Xml.XmlDocument();
            sPath = System.Windows.Forms.Application.StartupPath;
            oXmlDoc.Load(sPath + FileName);
            return (oXmlDoc.InnerXml);
        }

        public static string ObtenerSL()
        {
            try
            {
                Globals.Query = Properties.Resources.ObtenerSL;
                Globals.RunQuery(Globals.Query);
                Globals.oRec.MoveFirst();

                if (Globals.oRec.RecordCount == 0)
                    throw new Exception("No se encuentra configurado la IP para service layer, por favor verifique: Herramientas > Ventanas definidas por usuario > EXX_ADRG_CONF - Configuración Reclas Gasto \nRegistre la IP de service layer con el código 001");
                else
                {
                    string server = Globals.oRec.Fields.Item("U_EXX_CONF_VALOR").Value.ToString();
                    if (string.IsNullOrEmpty(server))
                        throw new Exception("No se encuentra configurado la IP para service layer, por favor verifique: Herramientas > Ventanas definidas por usuario > EXX_ADRG_CONF - Configuración Reclas Gasto \nRegistre la IP de service layer con el código 001");
                    else
                        return server;
                }
            }
            catch (Exception ex)
            {
                Globals.MessageBox(ex.Message);
                throw ex;
            }
            finally
            {
                Globals.Release(Globals.oRec);
            }
        }

        public static void ObtenerConfiguracion()
        {
            try
            {
                Globals.Query = Properties.Resources.ObtenerConfiguracion;
                Globals.RunQuery(Globals.Query);
                Globals.oRec.MoveFirst();

                if (Globals.oRec.RecordCount == 0)
                    throw new Exception("No se encuentra configurado los datos de IROUTE, por favor verifique: Herramientas > Ventanas definidas por usuario > EXX_AIIR_CONF - Configuración Reclas Gasto \nRegistre la IP, usuario y clave de IROUTE.");
                else
                {
                    string server = Globals.oRec.Fields.Item("URL").Value.ToString();
                    string usuario = Globals.oRec.Fields.Item("USUARIO").Value.ToString();
                    string clave = Globals.oRec.Fields.Item("CLAVE").Value.ToString();
                    if (string.IsNullOrEmpty(server)) throw new Exception("No se encuentra configurado la URL de I-ROUTE, por favor verifique: Herramientas > Ventanas definidas por usuario > EXX_AIIR_CONF - Configuración Reclas Gasto \nRegistre la IP de IROUTE con el código 002");
                    if (string.IsNullOrEmpty(usuario)) throw new Exception("No se encuentra configurado el usuario de I-ROUTE, por favor verifique: Herramientas > Ventanas definidas por usuario > EXX_AIIR_CONF - Configuración Reclas Gasto \nRegistre el usuario de IROUTE con el código 003");
                    if (string.IsNullOrEmpty(clave)) throw new Exception("No se encuentra configurado la contraseña de I-ROUTE, por favor verifique: Herramientas > Ventanas definidas por usuario > EXX_AIIR_CONF - Configuración Reclas Gasto \nRegistre la clave de IROUTE con el código 004");

                    Globals.UrlRoute = server;
                    Globals.UserRoute = usuario;
                    Globals.PasswordRoute = clave;
                }
            }
            catch (Exception ex)
            {
                Globals.MessageBox(ex.Message);
                throw ex;
            }
            finally
            {
                Globals.Release(Globals.oRec);
            }
        }
    }
}
