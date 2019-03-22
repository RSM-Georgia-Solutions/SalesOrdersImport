﻿using System;
using System.Collections.Generic;
using System.Xml;
using SalesOrdersImport.Controllers;
using SalesOrdersImport.Helpers;
using SAPbouiCOM;
using SAPbouiCOM.Framework;

namespace SalesOrdersImport
{
    [FormAttribute("SalesOrdersImport.Form1", "Import.b1f")]
    class Form1 : UserFormBase
    {
        public Form1()
        {
        }


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

        private void addPath1(string value)
        {
            EditText0.Value = value;

            if (value != "")
            {
                var sheetNames = excelFileController.ToExcelsSheetList(EditText0.Value);
                Others.SetSheetNames(sheetNames, ComboBox0);
            }
        }

        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            ShowFolder newFolder = new ShowFolder();
            //add function to event 
            newFolder.currFunc += addPath1;
            //run method of folder class
            newFolder.loadFolder();
        }

        private void Button2_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Close();
        }
    }
}