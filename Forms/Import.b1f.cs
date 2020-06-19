using System;
using System.Collections.Generic;
using System.Threading;
using System.Xml;
using SalesOrdersImport.Controllers;
using SalesOrdersImport.Helpers;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using SalesOrdersImport.Models;
using System.Threading.Tasks;
using RSM.SAPB1.Support;

namespace SalesOrdersImport
{
    [FormAttribute("SalesOrdersImport.Form1", "Forms/Import.b1f")]
    class Import : UserFormBase
    {
        public Import()
        {
        }

        public BaseController Controller { get; set; }

        ExcelFileController excelFileController = new ExcelFileController();
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_1").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_4").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Item_6").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_7").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_8").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_9").Specific));
            this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            SAPbouiCOM.Framework.Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(this.SBO_Application_ItemEvent_ChooseFromList);
        }

        private SAPbouiCOM.StaticText StaticText0;

        private void OnCustomInitialize()
        {

        }

        private void SBO_Application_ItemEvent_ChooseFromList(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST)
            {

                IChooseFromListEvent oCFLEvento = null;
                oCFLEvento = ((IChooseFromListEvent)(pVal));
                string sCFL_ID = null;
                sCFL_ID = oCFLEvento.ChooseFromListUID;
                Form oForm = null;
                oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(FormUID);
                SAPbouiCOM.ChooseFromList oCFL = null;

                oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                
                SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
                SAPbouiCOM.Condition oCon = null;
                oCon = oCons.Count==0?oCons.Add():oCons.Item(0);
                oCon.Alias = "CardType";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;

                oCon.CondVal = Controller.OrderType == OrderType.Purchase ? "S" : "C";
                
                oCFL.SetConditions(oCons);
                

                if (oCFLEvento.BeforeAction == false)
                {
                    DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;
                    string val = null;
                    try
                    {
                        val = Convert.ToString(oDataTable.GetValue(0, 0));
                    }
                    catch (Exception ex)
                    {

                    }
                    if ((pVal.ItemUID == "Item_6") | (pVal.ItemUID == "Button"))
                    {
                        oForm.DataSources.UserDataSources.Item("UD_0").ValueEx = val;
                    }

                }
            }

            if ((FormUID == "CFL1") & (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD))
            {
                System.Windows.Forms.Application.Exit();
            }

        }

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;



        //folder dialog
        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            RSM.SAPB1.Support.SelectFileDialog dialog = new SelectFileDialog("", "", "", DialogType.OPEN);
            dialog.Open();
            EditText0.Value = dialog.SelectedFile;

            var sheets = excelFileController.ToExcelsSheetList(EditText0.Value);

            while (ComboBox0.ValidValues.Count > 0)
            {
                ComboBox0.ValidValues.Remove(0, BoSearchKey.psk_Index);
            }

            for (int i = 0; i < sheets.Count; i++)
            {
                ComboBox0.ValidValues.Add(i.ToString(), sheets[i]);
            }

            if (ComboBox0.ValidValues.Count > 0)
            {
                ComboBox0.Select(0, BoSearchKey.psk_Index);
            }

        }

        //cancel
        private void Button2_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Close();
        }


        private void Button0_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                Controller.StartImport();
            }
            catch (Exception e)
            {
                Notifications.ShowMessage(e.Message, BoStatusBarMessageType.smt_Error);
            }

        }
    }
}