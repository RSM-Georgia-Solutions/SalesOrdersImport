using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM;
using SalesOrdersImport.Models;
using System.Data;
using MoreLinq;

namespace SalesOrdersImport.Controllers
{
    public abstract class BaseController
    {
        public EditText ExcelFile { get { return (EditText)Form.Items.Item("Item_1").Specific; } }
        public ComboBox ExcelSheet { get { return (ComboBox)Form.Items.Item("Item_4").Specific; } }
        public EditText CardCode { get { return (EditText)Form.Items.Item("Item_6").Specific; } }
        public ExcelFileController excelFileController { get; set; }

        public IForm Form { get; set; }
        public SAPbobsCOM.Company oCompany { get; set; }
        public OrderType OrderType { get; set; }

        public BaseController(IForm Form, SAPbobsCOM.Company oCompany, ExcelFileController excelController)
        {
            this.Form = Form;
            this.oCompany = oCompany;
            excelFileController = excelController;
        }


        public void StartImport()
        {
            if (ExcelSheet.Selected != null && CardCode.Value != "")
            {
                var test = ExcelSheet.Value;
                var test2 = ExcelSheet.Selected.Description;
                var test3 = ExcelSheet.Selected.ToString();
                var test4 = ExcelSheet.Selected.Value.ToString();
                var data = excelFileController.ReadExcelFile(ExcelSheet.Selected.Description, ExcelFile.Value);

                var salesOrders = ParseFile(CardCode.Value, data, OrderType);
                SAPbouiCOM.ProgressBar ProgressBar = null;

                if (DiManager.Company.InTransaction)
                {
                    DiManager.Company.StartTransaction();
                }

                List<string> salesOrderCodes = new List<string>();

                Task task = Task.Run(() => salesOrderCodes = PostOrders(salesOrders, ProgressBar));

                task.ConfigureAwait(true).GetAwaiter().OnCompleted(() =>
                {
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
                    PostedOrders postedOrders2 = new PostedOrders(salesOrderCodes, OrderType);
                    postedOrders2.Show();
                });

            }
        }

        protected List<OrderModel> ParseFile(string bpCode, System.Data.DataTable data, OrderType orderType)
        {
            List<OrderModel> salesOrderModels = new List<OrderModel>();

            //IEnumerable<DataRow> documents;
            var documents = data.AsEnumerable().ToList().DistinctBy(c => c["Document Number"]).Select(c => c["Document Number"]);

            foreach (var item in documents)
            {
                OrderModel salesOrder = new OrderModel
                {
                    BpCode = bpCode,
                    DeliveryDate = DateTime.Parse(data.AsEnumerable().First(c => c["Document Number"].ToString() == item.ToString())["Delivery Date"].ToString()),
                    LineNum = int.Parse(item.ToString()),
                    OrderType = orderType
                };

                

                if (orderType == OrderType.Sales)
                {
                    int AddressCodex = int.Parse(data.AsEnumerable().Where(c => c["Document Number"].ToString() == item.ToString()).First()["Address Code"].ToString());
                    salesOrder.AddressCode = AddressCodex;

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

                    string OnlineOrderN = data.AsEnumerable().Where(c => c["Document Number"].ToString() == item.ToString()).First()["Online Order N"].ToString();
                    salesOrder.OnlineOrderN = OnlineOrderN;

                }


                
                foreach (var doc in data.AsEnumerable().Where(c => c["Document Number"].ToString() == item.ToString()))
                {
                    var Quantity = int.Parse(doc["Quantity"].ToString());
                    var ItemCode = doc["Item Code"].ToString();

                    OrderRowModel salesRow = new OrderRowModel
                    {
                        ItemCode = ItemCode,
                        Quantity = Quantity
                    };
                    salesOrder.rows.Add(salesRow);
                }

                salesOrderModels.Add(salesOrder);

            }

            return salesOrderModels;
        }///

        protected List<string> PostOrders(List<OrderModel> orders, ProgressBar ProgressBar)
        {
            List<string> salesOrderCodes = new List<string>();
            try
            {
                ProgressBar = SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.CreateProgressBar("Creating Order", orders.Count, false);
            }
            catch (Exception e)
            {

            }

            foreach (var order in orders)
            {
                try
                {
                    if(OrderType == OrderType.Purchase)
                    {
                        order.Address = "";
                        order.OnlineOrderN = "";
                        order.UadrCode = "";
                    }
                    
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
    }
}
