using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SalesOrdersImport.Models
{
    public class OrderModel
    {
        public string BpCode { get; set; }
        public OrderType OrderType { get; set; }
        public DateTime DeliveryDate { get; set; }
        public string Address { get; set; }
        public int AddressCode { get; set; }
        public int LineNum { get; set; }
        public string UadrCode { get; set; }
        public string OnlineOrderN { get; set; }

        public List<OrderRowModel> rows = new List<OrderRowModel>();

        public string Add()
        {
            SAPbobsCOM.Documents order = OrderType == OrderType.Sales ? (SAPbobsCOM.Documents)RSM.SAPB1.Support.Global.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders) : (SAPbobsCOM.Documents)RSM.SAPB1.Support.Global.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);
            order.CardCode = BpCode;
            order.DocDate = DateTime.Now;
            order.DocDueDate = DeliveryDate;
            order.Address2 = Address;
            order.UserFields.Fields.Item("U_RSM_UADR_CODE").Value = UadrCode;
            order.UserFields.Fields.Item("U_ONLN_ORDR_N").Value = OnlineOrderN;

            foreach (var row in rows)
            {
                order.Lines.ItemCode = row.ItemCode;
                order.Lines.Quantity = row.Quantity;
                order.Lines.Add();
            }

            var res = order.Add();
            if (res != 0)
            {
                throw new Exception(RSM.SAPB1.Support.Global.Company.GetLastErrorDescription());
            }
            else
            {
                return RSM.SAPB1.Support.Global.Company.GetNewObjectKey();
            }
        }
    }

    public class OrderRowModel
    {
        public string ItemCode { get; set; }
        public int Quantity { get; set; }

    }


}
