using MoreLinq;
using SalesOrdersImport.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SalesOrdersImport.Controllers
{
    public class SalesOrderController
    {
        public static List<SalesOrderModel> parseDataTableToSalesOrder(string bpCode, DataTable data)
        {
            List<SalesOrderModel> salesOrderModels = new List<SalesOrderModel>();

            //IEnumerable<DataRow> documents;
            var documents = data.AsEnumerable().ToList().DistinctBy(c => c["Document Number"]).Select(c => c["Document Number"]);

            foreach (var item in documents)
            {
                SalesOrderModel salesOrder = new SalesOrderModel
                {
                    BpCode = bpCode,
                    DeliveryDate = DateTime.Parse(data.AsEnumerable().Where(c => c["Document Number"].ToString() == item.ToString()).First()["Delivery Date"].ToString()),
                    LineNum = int.Parse(item.ToString())
                };

                int AddressCodex = int.Parse(data.AsEnumerable().Where(c => c["Document Number"].ToString() == item.ToString()).First()["Address Code"].ToString());
                salesOrder.AddressCode = AddressCodex;

                DiManager.Recordset.DoQuery(DiManager.QueryHanaTransalte($"SELECT * FROM [@RSM_UADR] WHERE Code = {AddressCodex}"));
                if (DiManager.Recordset.EoF)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Address not Found");
                }
                else
                {
                    string Address = DiManager.Recordset.Fields.Item("Name").Value.ToString() + Environment.NewLine + DiManager.Recordset.Fields.Item("U_District").Value.ToString() + Environment.NewLine + DiManager.Recordset.Fields.Item("U_ID").Value.ToString() + Environment.NewLine + DiManager.Recordset.Fields.Item("U_Address").Value.ToString();
                    salesOrder.Address = Address;
                }
                foreach (var doc in data.AsEnumerable().Where(c => c["Document Number"].ToString() == item.ToString()))
                {
                    var LineNum = int.Parse(doc["Document Number"].ToString());
                    var Quantity = int.Parse(doc["Quantity"].ToString());
                    var ItemCode = doc["Item Code"].ToString();
                    var AddressCode = int.Parse(doc["Address Code"].ToString());
                    var DeliveryDate = DateTime.Parse(doc["Delivery Date"].ToString());

                    SalesOrderRowModel salesRow = new SalesOrderRowModel
                    {
                        ItemCode = ItemCode,
                        Quantity = Quantity
                    };
                    salesOrder.rows.Add(salesRow);
                }

                salesOrderModels.Add(salesOrder);

            }

            return salesOrderModels;
        }
    }
}
