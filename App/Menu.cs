using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Integration_IROUTE.App
{
    public static class Menu
    {
        public static void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = default(SAPbouiCOM.Menus);
            SAPbouiCOM.MenuItem oMenuItem = default(SAPbouiCOM.MenuItem);
            oMenus = Globals.SBO_Application.Menus;
            SAPbouiCOM.MenuCreationParams oCreationPackage = default(SAPbouiCOM.MenuCreationParams);
            oCreationPackage = (SAPbouiCOM.MenuCreationParams)Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
            try
            {
                oMenuItem = Globals.SBO_Application.Menus.Item("43520");
                oMenus = oMenuItem.SubMenus;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                oCreationPackage.UniqueID = "EXX_AIIR";
                oCreationPackage.String = "EXX - Integración I-ROUTE";
                oCreationPackage.Position = oMenuItem.SubMenus.Count + 1;

                //oCreationPackage.Image = ""//ruta iamgen

                if (!(oMenus.Exists("EXX_AIIR")))
                {
                    oMenus.AddEx(oCreationPackage);
                }

                oMenuItem = Globals.SBO_Application.Menus.Item("EXX_AIIR");
                oMenus = oMenuItem.SubMenus;

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "EXX_AIIR1";
                oCreationPackage.String = "EXX - Integración Guías";
                if (!(oMenus.Exists("EXX_AIIR1")))
                {
                    oMenus.AddEx(oCreationPackage);
                }
            }
            catch (Exception ex)
            {
                Globals.SBO_Application.SetStatusBarMessage(ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }
        }
    }
}
