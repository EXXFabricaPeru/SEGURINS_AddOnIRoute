using Integration_IROUTE.App;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Integration_IROUTE.DBStructure
{
    public class CreateStructure
    {
        public static void CreateStruct()
        {
            List<List<String>> EstadoValues = new List<List<String>>();
            agregarValorValido("PE", "Pendiente", ref EstadoValues);
            agregarValorValido("PR", "Programado", ref EstadoValues);
            agregarValorValido("EN", "Entregado", ref EstadoValues);
            agregarValorValido("DE", "Devuelto", ref EstadoValues);
            
            crearTabla("EXX_AIIR_CONF", "Conf - Integración IROUTE", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            crearCampo("EXX_AIIR_CONF", "EXX_CONF_VALOR", "Valor", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", "", false, null);

            crearCampo("ODLN", "EXX_AIIR_EST", "Estado IR", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 2, "PE", "", false, EstadoValues);

        }
        private static void agregarValorValido(String Value, String Description, ref List<List<String>> ValidValues)
        {
            List<String> ValidValue = new List<String>();
            ValidValue.Add(Value);
            ValidValue.Add(Description);
            ValidValues.Add(ValidValue);
        }

        private static bool crearTabla(string tabla, string nombretabla, SAPbobsCOM.BoUTBTableType tipo = SAPbobsCOM.BoUTBTableType.bott_NoObject)//CHG
        {
            SAPbobsCOM.UserTablesMD oTablaUser = (SAPbobsCOM.UserTablesMD)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
            try
            {
                if (!oTablaUser.GetByKey(tabla))
                {
                    oTablaUser.TableName = tabla;
                    oTablaUser.TableDescription = nombretabla;
                    oTablaUser.TableType = tipo;

                    int RetVal = oTablaUser.Add();
                    if ((RetVal != 0))
                    {
                        String errMsg;
                        int errCode;
                        Globals.oCompany.GetLastError(out errCode, out errMsg);
                        throw new Exception(errMsg);
                    }
                    else
                        return true;
                }
                else
                    return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oTablaUser);
            }
        }

        private static void crearCampo(string tabla, string campo, string descripcion, SAPbobsCOM.BoFieldTypes tipo, SAPbobsCOM.BoFldSubTypes subtipo, int tamaño,
                                       string ValorPorDefecto, string sLinkedTable, Boolean Mandatory, List<List<String>> ValidValues, string LinkedType = null)
        {
            int existeCampo = 0;

            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string cadena = "select \"FieldID\" from CUFD where (\"TableID\"='" + tabla + "' or \"TableID\"='@" + tabla + "') and \"AliasID\"='" + campo + "'";
            rs.DoQuery(cadena);

            existeCampo = rs.RecordCount;
            int FieldID = Convert.ToInt32(rs.Fields.Item(0).Value);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            rs = null;

            SAPbobsCOM.UserFieldsMD oCampo = (SAPbobsCOM.UserFieldsMD)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

            try
            {
                if (existeCampo == 0)//Crear
                {
                    oCampo.TableName = tabla;
                    oCampo.Name = campo;
                    oCampo.Description = descripcion;
                    oCampo.Type = tipo;
                    oCampo.SubType = subtipo;
                    oCampo.Mandatory = Mandatory ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;

                    if (tamaño > 0)
                    {
                        oCampo.EditSize = tamaño;
                    }

                    if (sLinkedTable.ToString() != "")
                    {
                        switch (LinkedType)
                        {
                            case Globals.LinkedSystemObject:
                                oCampo.LinkedSystemObject = (SAPbobsCOM.UDFLinkedSystemObjectTypesEnum)Convert.ToInt32(sLinkedTable);
                                break;
                            case Globals.LinkedUDO:
                                oCampo.LinkedUDO = sLinkedTable;
                                break;
                            case Globals.LinkedTable:
                                oCampo.LinkedTable = sLinkedTable;
                                break;
                        }
                    }

                    if (ValidValues != null)
                    {
                        foreach (List<String> ValidValue in ValidValues)
                        {
                            oCampo.ValidValues.Value = ValidValue[0];
                            oCampo.ValidValues.Description = ValidValue[1];
                            oCampo.ValidValues.Add();
                        }
                    }

                    if (ValorPorDefecto.ToString() != "")
                    {
                        oCampo.DefaultValue = ValorPorDefecto;
                    }

                    int RetVal = oCampo.Add();
                    if (RetVal != 0)
                    {
                        String errMsg;
                        int errCode;
                        Globals.oCompany.GetLastError(out errCode, out errMsg);
                        throw new Exception(errMsg);
                    }
                }
                //else//Actualizar
                //{
                //    oCampo.GetByKey("@" + tabla, FieldID);
                //    oCampo.Description = descripcion;
                //    if (ValidValues != null)
                //    {
                //        foreach (List<String> ValidValue in ValidValues)
                //        {
                //            Boolean Existe = false;
                //            for (int i = 0; i < oCampo.ValidValues.Count; i++)
                //            {
                //                oCampo.ValidValues.SetCurrentLine(i);
                //                if (oCampo.ValidValues.Value == ValidValue[0])
                //                    Existe = true;

                //            }

                //            if (!Existe)
                //            {
                //                oCampo.ValidValues.Value = ValidValue[0];
                //                oCampo.ValidValues.Description = ValidValue[1];
                //                oCampo.ValidValues.Add();
                //            }
                //        }
                //    }

                //    if (ValorPorDefecto.ToString() != "")
                //    {
                //        oCampo.DefaultValue = ValorPorDefecto;
                //    }

                //    int RetVal = oCampo.Update();
                //    if ((RetVal != 0))
                //    {
                //        String errMsg;
                //        int errCode;
                //        oCompany.GetLastError(out errCode, out errMsg);
                //        throw new Exception(errMsg);
                //    }
                //}
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCampo);
            }
        }
    }
}