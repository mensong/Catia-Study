using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CatiaModule
{
    class Program
    {
        static void Main(string[] args)
        {
            // 获取当前活动ProductDocument
            ProductStructureTypeLib.ProductDocument prdDoc =
                (ProductStructureTypeLib.ProductDocument)CatiaHelper.ActiveDoc;
            // 创建一个ID为newProduct的Product
            var p = prdDoc.Product.Products.AddNewProduct("newProduct");

            //Catia.StartCommand("LoginCmd");

            //string iLibraryName = @"G:\CatiaMacro\VBA 项目2.catvba";
            //string iProgramName = @"模块1";
            //string iFunctionName = @"myfoo";
            //CatiaHelper.ExecuteVBA(iLibraryName, iProgramName, iFunctionName);

            //CatiaHelper.App.SystemService.Evaluate(
            //    "Function myfoo()\r\nMsgBox \"ok\"\r\nmyfoo = \"我是返回值\"\r\nEnd Function",
            //    INFITF.CATScriptLanguage.CATVBScriptLanguage, "myfoo", null);

            //CatiaHelper.ExecuteVBAScript("MsgBox \"ok\"");
            //CatiaHelper.ExecuteVBAScript("Function myfoo()\r\nMsgBox \"ok\"\r\nmyfoo = \"我是返回值\"\r\nEnd Function", "myfoo");


        }
    }
}
