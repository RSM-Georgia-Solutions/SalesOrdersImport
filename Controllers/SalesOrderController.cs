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
                    DeliveryDate = DateTime.Parse(data.AsEnumerable().First(c => c["Document Number"].ToString() == item.ToString())["Delivery Date"].ToString()),
                    LineNum = int.Parse(item.ToString())
                };

                int AddressCodex = int.Parse(data.AsEnumerable().Where(c => c["Document Number"].ToString() == item.ToString()).First()["Address Code"].ToString());
                salesOrder.AddressCode = AddressCodex;

                string OnlineOrderN = data.AsEnumerable().Where(c => c["Document Number"].ToString() == item.ToString()).First()["Online Order N"].ToString();
                salesOrder.OnlineOrderN = OnlineOrderN;

                DiManager.Recordset.DoQuery(DiManager.QueryHanaTransalte($"SELECT * FROM [@RSM_UADR] WHERE Code = {AddressCodex}"));
                if (DiManager.Recordset.EoF)
                {
                   // SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Address not Found");
                }
                else
                {
                    string Address = DiManager.Recordset.Fields.Item("Code").Value + Environment.NewLine + DiManager.Recordset.Fields.Item("Name").Value + Environment.NewLine + DiManager.Recordset.Fields.Item("U_District").Value + Environment.NewLine + DiManager.Recordset.Fields.Item("U_ID").Value + Environment.NewLine + DiManager.Recordset.Fields.Item("U_Address").Value; 
                    salesOrder.Address = Address;
                    salesOrder.UadrCode = DiManager.Recordset.Fields.Item("Code").Value.ToString();
                }
                foreach (var doc in data.AsEnumerable().Where(c => c["Document Number"].ToString() == item.ToString()))
                {
                    var LineNum = int.Parse(doc["Document Number"].ToString());
                    var x = doc["Quantity"].ToString();
                    var Quantity = int.Parse(doc["Quantity"].ToString());
                    var ItemCode = doc["Item Code"].ToString();
                    var AddressCode = int.Parse(doc["Address Code"].ToString());
                    var DeliveryDate = DateTime.Parse(doc["Delivery Date"].ToString());
                    var OnlnOrdrN = doc["Online Order N"].ToString();

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

        private List<string> postSalesOrders(List<SalesOrderModel> salesOrders, ProgressBar ProgressBar)
        {
            List<string> salesOrderCodes = new List<string>();
            try
            {
                ProgressBar = SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.CreateProgressBar("Creating Sales Order", salesOrders.Count, false);
            }
            catch (Exception e)
            {

            }

            foreach (var order in salesOrders)
            {
                try
                {
                    string err = order.Add();
                    salesOrderCodes.Add(err);
                }
                catch (Exception e)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(e.Message);
                    if (DiManager.Company.InTransaction)
                    {
                        DiManager.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    }
                    try
                    {
                        ProgressBar.Stop();
                    }
                    catch (Exception)
                    {
                    }
                    return new List<string>();
                }
                try
                {
                    ProgressBar.Value++;
                }
                catch (Exception)
                {

                }
            }
            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetSystemMessage("წარმატება", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
            try
            {
                ProgressBar.Stop();
            }
            catch (Exception e)
            {
            }
            return salesOrderCodes;
        }

        public void StartImport()
        {
            string bpCode = EditText2.Value;

            if (ComboBox0.Selected != null && EditText2.Value != "")
            {
                var data = excelFileController.ReadExcelFile(ComboBox0.Selected.Value, EditText0.Value);
                var salesOrders = SalesOrderController.parseDataTableToSalesOrder(bpCode, data);
                SAPbouiCOM.ProgressBar ProgressBar = null;

                if (DiManager.Company.InTransaction)
                {
                    DiManager.Company.StartTransaction();
                }

                List<string> salesOrderCodes = new List<string>();

                Task task = Task.Run(() => salesOrderCodes = postSalesOrders(salesOrders, ProgressBar));

                task.ConfigureAwait(true).GetAwaiter().OnCompleted(() => {
                    if (DiManager.Company.InTransaction)
                    {
                        DiManager.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetSystemMessage("წარმატება", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);

                    }

                    try
                    {
                        ProgressBar.Stop();
                    }
                    catch (Exception)
                    {

                    }
                    //PostedSalesOrders postedOrders = new PostedSalesOrders();
                    //postedOrders.Show();
                    PostedOrders postedOrders2 = new PostedOrders(salesOrderCodes);
                    postedOrders2.Show();
                });

            }
        }
    }
}
