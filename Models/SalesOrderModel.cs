﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SalesOrdersImport.Models
{
    public class SalesOrderModel
    {
        public string BpCode { get; set; }
        public DateTime DeliveryDate { get; set; }
        public string Address { get; set; }
        public int AddressCode { get; set; }
        public int LineNum { get; set; }
        public string UadrCode { get; set; }
        public string OnlineOrderN { get; set; }

        public List<SalesOrderRowModel> rows = new List<SalesOrderRowModel>();

        public string Add()
        {
            SAPbobsCOM.Documents salesOrder = (SAPbobsCOM.Documents)DiManager.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
            salesOrder.CardCode = BpCode;
            salesOrder.DocDate = DateTime.Now;
            salesOrder.DocDueDate = DeliveryDate;
            salesOrder.Address2 = Address;
            salesOrder.UserFields.Fields.Item("U_RSM_UADR_CODE").Value = UadrCode;
            salesOrder.UserFields.Fields.Item("U_ONLN_ORDR_N").Value = OnlineOrderN;

            foreach (var row in rows)
            {
                salesOrder.Lines.ItemCode = row.ItemCode;
                salesOrder.Lines.Quantity = row.Quantity;
                salesOrder.Lines.Add();
            }

            var res = salesOrder.Add();
            if (res != 0)
            {
                throw new Exception(DiManager.Company.GetLastErrorDescription());
            }
            else
            {
                return DiManager.Company.GetNewObjectKey();
            }
        }
    }

    public class SalesOrderRowModel
    {
        public string ItemCode { get; set; }
        public int Quantity { get; set; }

    }


}
