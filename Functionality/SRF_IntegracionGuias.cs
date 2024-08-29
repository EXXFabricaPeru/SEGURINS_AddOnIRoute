using Integration_IROUTE.App;
using Integration_IROUTE.Entity;
using Integration_IROUTE.Framework;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Integration_IROUTE.Functionality
{
    public class SRF_IntegracionGuias
    {
        public static void LoadFormEstr(string Frm)
        {
            SAPbouiCOM.Form oForm = default(SAPbouiCOM.Form);
            try
            {
                oForm = Globals.SBO_Application.Forms.Item(Frm);
                Globals.SBO_Application.MessageBox("El formulario ya se encuentra abierto.");
                oForm.Visible = true;
            }
            catch
            {
                SAPbouiCOM.FormCreationParams fcp = default(SAPbouiCOM.FormCreationParams);
                fcp = (SAPbouiCOM.FormCreationParams)Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                fcp.FormType = Frm;
                fcp.UniqueID = Frm;
                string FormName = "\\Views\\frmIntegracionGuias.srf";
                fcp.XmlData = Globals.LoadFromXML(ref FormName);
                oForm = Globals.SBO_Application.Forms.AddEx(fcp);

                oForm.Visible = true;
            }
        }

        public static void ItemPressed(ref ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.ActionSuccess)
                    switch (pVal.ItemUID)
                    {
                        case "btnBuscar":
                            Buscar(pVal, oForm, out BubbleEvent);
                            break;
                        case "btnEnviar":
                            EnviarIRoute(pVal, oForm, out BubbleEvent);
                            break;
                    }
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                BubbleEvent = false;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }

        private static void Buscar(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                oForm.Freeze(true);
                Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid1").Specific;
                SAPbouiCOM.DataTable oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_1");
                EditTextColumn etColumna;
                ComboBoxColumn cbColumna;

                ComboBox Estado = (ComboBox)oForm.Items.Item("cbEstado").Specific;
                if (Estado.Selected == null) throw new Exception("Primero seleccione un estado.");
                string EstadoFiltro = "'" + (Estado.Selected.Value == "TD" ? "PE','PR','EN','DE" : Estado.Selected.Value) + "'";

                Globals.Query = Integration_IROUTE.Properties.Resources.ListarGuías;
                Globals.Query = string.Format(Globals.Query, EstadoFiltro);
                oGrid.DataTable = oLista;
                oGrid.DataTable.Rows.Clear();
                oLista.Rows.Clear();
                oGrid.DataTable.ExecuteQuery(Globals.Query);
                oGrid.Columns.Item("Sel").Type = BoGridColumnType.gct_CheckBox;
                oGrid.Columns.Item("Sel").TitleObject.Caption = "Seleccionar";
                oGrid.Columns.Item("DocEntry").TitleObject.Caption = "N° Interno";
                oGrid.Columns.Item("DocEntry").Editable = false;
                etColumna = ((EditTextColumn)(oGrid.Columns.Item("DocEntry")));
                etColumna.LinkedObjectType = "15";
                oGrid.Columns.Item("DocNum").TitleObject.Caption = "Número";
                oGrid.Columns.Item("DocNum").Visible = false;
                oGrid.Columns.Item("FolioPref").TitleObject.Caption = "Serie";
                oGrid.Columns.Item("FolioPref").Editable = false;
                oGrid.Columns.Item("FolioNum").TitleObject.Caption = "Correlativo";
                oGrid.Columns.Item("FolioNum").Editable = false;
                oGrid.Columns.Item("CardCode").TitleObject.Caption = "Cliente";
                oGrid.Columns.Item("CardCode").Editable = false;
                etColumna = ((EditTextColumn)(oGrid.Columns.Item("CardCode")));
                etColumna.LinkedObjectType = "2";
                oGrid.Columns.Item("CardName").TitleObject.Caption = "Nombre";
                oGrid.Columns.Item("CardName").Editable = false;
                oGrid.Columns.Item("U_EXX_AIIR_EST").TitleObject.Caption = "Estado";
                oGrid.Columns.Item("U_EXX_AIIR_EST").Editable = false;
                oGrid.Columns.Item("U_EXX_AIIR_EST").Type = BoGridColumnType.gct_ComboBox;
                cbColumna = ((ComboBoxColumn)(oGrid.Columns.Item("U_EXX_AIIR_EST")));
                cbColumna.ValidValues.Add("PE", "Pendiente");
                cbColumna.ValidValues.Add("PR", "Programado");
                cbColumna.ValidValues.Add("EN", "Entregado");
                cbColumna.ValidValues.Add("DE", "Devuelto");
                cbColumna.DisplayType = BoComboDisplayType.cdt_Description;
                oGrid.Columns.Item("DocDate").TitleObject.Caption = "Fecha contabilización";
                oGrid.Columns.Item("DocDate").Editable = false;
                oGrid.Columns.Item("Comments").TitleObject.Caption = "Comentario";
                oGrid.Columns.Item("Comments").Editable = false;
                oGrid.Columns.Item("Mensaje").TitleObject.Caption = "Mensaje";
                oGrid.Columns.Item("Mensaje").Editable = false;
                if (Estado.Selected.Value == "PE")
                {
                    oGrid.Columns.Item("Sel").Editable = oGrid.Columns.Item("Sel").Editable = true;
                    oGrid.Columns.Item("Sel").Visible = true;
                    ((Button)oForm.Items.Item("btnEnviar").Specific).Item.Enabled = true;
                    ((Button)oForm.Items.Item("btnEnviar").Specific).Item.Visible = true;
                }
                else
                {
                    oGrid.Columns.Item("Sel").Editable = false;
                    oGrid.Columns.Item("Sel").Visible = false;
                    ((Button)oForm.Items.Item("btnEnviar").Specific).Item.Visible = false;
                }
                oGrid.AutoResizeColumns();
                oGrid.Columns.Item("Mensaje").Width = 150;

                string Path = AppDomain.CurrentDomain.BaseDirectory + "\\Images";
                Bitmap bmp;
                for (int index = 0; index < oGrid.Rows.Count; index++)
                {
                    bmp = new Bitmap(Path + @"\ImagesWhite.jpg");
                    oGrid.CommonSetting.SetRowBackColor(index + 1, ColorTranslator.ToOle(bmp.GetPixel(1, 1)));
                }
                oGrid.CommonSetting.FixedColumnsCount = 2;
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                BubbleEvent = false;
                throw ex;
            }
        }

        private static void EnviarIRoute(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                ((Button)oForm.Items.Item("btnEnviar").Specific).Item.Enabled = false;
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid1").Specific;
                string Path = AppDomain.CurrentDomain.BaseDirectory + "\\Images";
                Bitmap bmp;

                Globals.ObtenerConfiguracion();

                for (int index = 0; index < oGrid.Rows.Count; index++)
                {

                    if (oGrid.DataTable.GetValue("Sel", index).ToString().Trim() == "Y" &&
                        oGrid.DataTable.GetValue("DocEntry", index).ToString().Trim() != "")
                    {
                        try
                        {
                            string estado = string.Empty;
                            string mensaje = string.Empty;
                            string DocEntry = oGrid.DataTable.GetValue("DocEntry", index).ToString();

                            var rsp = SL.ObtenerGuia(DocEntry);
                            if (rsp.StatusCode != System.Net.HttpStatusCode.OK && rsp.StatusCode != System.Net.HttpStatusCode.Created && rsp.StatusCode != System.Net.HttpStatusCode.NoContent && rsp.StatusCode != 0)
                            {
                                mensaje = SapError.get_message(rsp.Content);
                                oGrid.DataTable.SetValue("Mensaje", index, mensaje.Length > 254 ? mensaje.Substring(0, 254) : mensaje);
                                bmp = new Bitmap(Path + @"\ImagesRed.jpg");
                                oGrid.CommonSetting.SetRowBackColor(index + 1, ColorTranslator.ToOle(bmp.GetPixel(0, 0)));
                            }
                            else
                            {
                                ODLN oGuia = JsonConvert.DeserializeObject<ODLN>(rsp.Content);

                                List<Entrega> oEntrega = new List<Entrega>();
                                oEntrega.Add(new Entrega
                                {
                                    nroTransporte = oGuia.DocNum.ToString(),
                                    codigoSociedad = "SEGURINDUSTRIA", // oGuia.Address,
                                    codigoAlmacen = null, //oGuia.Address,
                                    codigoViajeConsolidado = "viaje-" + oGuia.DocNum,
                                    fechaEstimadaInicioViaje = Convert.ToDateTime(oGuia.TaxDate).ToString("dd/MM/yyyy") + " 08:00",
                                    fechaEstimadaFinViaje = Convert.ToDateTime(oGuia.TaxDate).ToString("dd/MM/yyyy") + "15:00",
                                    nroPedido = oGuia.DocNum.ToString(),
                                    nroPedidoWeb = "web-" + oGuia.DocNum, //oGuia.Address,
                                    fechaEntrega = Convert.ToDateTime(oGuia.TaxDate).ToString("dd/MM/yyyy"),
                                    ventanaHorariaInicio = "09:00",
                                    ventanaHorariaFin = "15:00",
                                    horaCita = "11:00",
                                    tiempoVisita = 30,
                                    cantidadBultos = oGuia.DocumentLines.Count(),
                                    peso = oGuia.DocumentLines.Sum(x => x.Height1),
                                    prioridad = "Si", //oGuia.Address,
                                    fragil = "No", //oGuia.Address,
                                    adicional1 = "",
                                    adicional2 = "",
                                    debeCobrar = "No", //oGuia.Address,
                                    codRastreo = "tailoy-34BD5",
                                    viajeCourier = 1,
                                    Cliente = new cliente
                                    {
                                        codigoCliente = oGuia.CardCode,
                                        nombre = oGuia.CardName,
                                        telefono = "928374652",
                                        correo = "tester@comsatel.com.pe",
                                        codigoCanal = "CAN001",
                                        codigoSubcanal = "SUBCAN005",
                                        categoria = "GENERAL",
                                        codigoSociedad = "2100",
                                        Sede = new sede
                                        {
                                            codigo = oGuia.ShipToCode,
                                            nombre = oGuia.ShipToCode,
                                            direccion = oGuia.Address2,
                                            referencia = oGuia.Address2,
                                            longitud = "",
                                            latitud = "",
                                            departamento = "Lima",
                                            provincia = "Lima",
                                            distrito = "La Molina",
                                            Contacto = new contacto
                                            {
                                                nombre = "Julio",
                                                apellido = "Ramirez",
                                                docIdentidad = "81430367",
                                                correo = "desarrolladorDemo@comsatel.com.pe"
                                            }
                                        }
                                    },
                                    Guias = new List<guias> { new guias
                                    {
                                        nroGuia = oGuia.FolioPrefixString + "-" + oGuia.FolioNumber,
                                        formaPago = "EFECTIVO",
                                        detalleGuia = String.Join(", ", oGuia.DocumentLines.Select(x => x.ItemDescription))
                                    } }
                                });

                                var rsp2 = IROUTE.EnviarGuia(oEntrega);
                                if (rsp2.StatusCode != System.Net.HttpStatusCode.OK && rsp2.StatusCode != System.Net.HttpStatusCode.Created && rsp2.StatusCode != System.Net.HttpStatusCode.NoContent && rsp2.StatusCode != 0)
                                {
                                    mensaje = SapError.get_message(rsp2.Content);
                                    oGrid.DataTable.SetValue("Mensaje", index, mensaje.Length > 254 ? mensaje.Substring(0, 254) : mensaje);
                                    bmp = new Bitmap(Path + @"\ImagesRed.jpg");
                                    oGrid.CommonSetting.SetRowBackColor(index + 1, ColorTranslator.ToOle(bmp.GetPixel(0, 0)));
                                }
                                else
                                {
                                    //Enviar 
                                    estado = "RECIBIDO";

                                    if (estado == "RECIBIDO")
                                    {
                                        Globals.Query = "UPDATE ODLN set \"U_EXX_AIIR_EST\" = 'PR' WHERE \"DocEntry\" = " + DocEntry;
                                        Globals.RunQuery(Globals.Query);
                                        oGrid.DataTable.SetValue("Mensaje", index, "Sincronización exitosa");
                                        bmp = new Bitmap(Path + @"\ImagesGreen.jpg");
                                        oGrid.CommonSetting.SetRowBackColor(index + 1, ColorTranslator.ToOle(bmp.GetPixel(0, 0)));
                                    }
                                    else
                                    {
                                        oGrid.DataTable.SetValue("Mensaje", index, "Error de I-ROUTE");
                                        bmp = new Bitmap(Path + @"\ImagesRed.jpg");
                                        oGrid.CommonSetting.SetRowBackColor(index + 1, ColorTranslator.ToOle(bmp.GetPixel(0, 0)));
                                    }
                                }
                            }

                        }
                        catch (Exception ex)
                        {
                            oGrid.DataTable.SetValue("Mensaje", index, ex.Message.Length > 254 ? ex.Message.Substring(0, 254) : ex.Message);
                            bmp = new Bitmap(Path + @"\ImagesRed.jpg");
                            oGrid.CommonSetting.SetRowBackColor(index + 1, ColorTranslator.ToOle(bmp.GetPixel(0, 0)));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                throw ex;
            }
            finally
            {
                Globals.Release(Globals.oRec);
                GC.Collect();
            }
        }

        public static void ChooseFromList(ref ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.ActionSuccess)
                    switch (pVal.ItemUID)
                    {
                        case "etCardCode":
                            SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                            string sCFL_ID = oCFLEvento.ChooseFromListUID;
                            SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                            SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                            if (oDataTable != null)
                            {
                                ((SAPbouiCOM.EditText)oForm.Items.Item("etCardName").Specific).Value = oDataTable.GetValue("CardName", 0).ToString();
                                ((SAPbouiCOM.EditText)oForm.Items.Item("etCardCode").Specific).Value = oDataTable.GetValue("CardCode", 0).ToString();
                            }
                            break;
                    }
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                BubbleEvent = false;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
