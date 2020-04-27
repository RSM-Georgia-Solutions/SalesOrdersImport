using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace SalesOrdersImport
{
    [FormAttribute("SalesOrdersImport.PostedOrders", "PostedOrders.b1f")]
    class PostedOrders : UserFormBase
    {
        public List<string> OrderCodes { get; set; }
        //public PostedOrders(List<string> orderCodes)
        //{
        //    _orderCodes = orderCodes;
        //}
        public PostedOrders(List<string> orderCodes)
        {
            OrderCodes = orderCodes;
            string query = string.Empty;
            if (OrderCodes == null)
            {
                return;
            }
            foreach (string orderCode in OrderCodes)
            {
                query += $"SELECT  '{ orderCode}'   as [Order Code] union all ";
            }
            try
            {
                query = query.Remove(query.Length - 10, 10);

                Grid0.DataTable.ExecuteQuery(DiManager.QueryHanaTransalte($"{query}"));
                SAPbouiCOM.EditTextColumn oEditCol;
                oEditCol = ((SAPbouiCOM.EditTextColumn)(Grid0.Columns.Item("Order Code")));
                oEditCol.LinkedObjectType = "17";
                oEditCol.Editable = false;
            }
            catch (Exception)
            {

            }
        }
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_0").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Grid Grid0;

        private void OnCustomInitialize()
        {

        }
    }
}
