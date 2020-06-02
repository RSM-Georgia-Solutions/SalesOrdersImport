﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using SAPbouiCOM;

namespace SalesOrdersImport.Controllers
{
    public class PurchaseOrderController : BaseController
    {
        public PurchaseOrderController(IForm Form, SAPbobsCOM.Company oCompany, ExcelFileController excelController) : base(Form, oCompany, excelController)
        {
            OrderType = OrderType.Purchase;
        }

    }
}