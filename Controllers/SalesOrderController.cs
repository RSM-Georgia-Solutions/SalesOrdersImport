using MoreLinq;
using SalesOrdersImport.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using SAPbouiCOM;

namespace SalesOrdersImport.Controllers
{
    public class SalesOrderController : BaseController
    {
        public SalesOrderController(IForm Form, SAPbobsCOM.Company oCompany, ExcelFileController excelController) : base(Form, oCompany, excelController)
        {
            OrderType = OrderType.Sales;
        }

    }
}
