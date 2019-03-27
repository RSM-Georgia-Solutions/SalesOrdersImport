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
        List<string> _orderCodes;
        //public PostedOrders(List<string> orderCodes)
        //{
        //    _orderCodes = orderCodes;
        //}
        public PostedOrders()
        {

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
            string query = string.Empty;
            if (_orderCodes == null)
            {
                return;
            }
            foreach (string orderCode in _orderCodes)
            {
                query += orderCode + " as [Order Code],";
            }
            query = query.Remove(query.Length - 1, 1);
            Grid0.DataTable.ExecuteQuery(DiManager.QueryHanaTransalte($"SELECT {query}"));
        }
    }
}
