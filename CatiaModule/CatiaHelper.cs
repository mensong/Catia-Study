using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CatiaModule
{
    public class CatiaHelper
    {
        private static INFITF.Application _catiaApp = null;
        public static INFITF.Application App 
        { 
            get 
            {
                if (_catiaApp != null)
                    return _catiaApp;
                _catiaApp = (INFITF.Application)Marshal.GetActiveObject("Catia.Application");
                return _catiaApp;
            }
        }

        /// <summary>
        /// ProductStructureTypeLib.ProductDocument
        /// DRAFTINGITF.DrawingDocument
        /// PPR.PPRDocument
        /// CATFunctSystem.FunctionalDocument
        /// CATMat.MaterialDocument
        /// ComponentsCatalogsTypeLib.CatalogDocument
        /// MECMOD.PartDocument
        /// PROCESSITF.ProcessDocument
        /// SAMITF.AnalysisDocument
        /// 这些可以在对象浏览器里搜索Document查得到
        /// </summary>
        public static INFITF.Document ActiveDoc
        {
            get 
            {
                return App.ActiveDocument;
            }
            set
            {
                value.Activate();
            }
        }

        public static dynamic ExecuteVBA(
            string iLibraryFilePath, 
            string iProgramName, 
            string iFunctionName, 
            IEnumerable<object> iParameters = null)
        {
            Array parameters = null;
            if (iParameters != null)
                parameters = iParameters.ToArray();
            var res = App.SystemService.ExecuteScript(
                ref iLibraryFilePath,
                INFITF.CatScriptLibraryType.catScriptLibraryTypeVBAProject,
                ref iProgramName,
                ref iFunctionName,
                parameters);
            return res;
        }

        public static dynamic ExecuteVBAScript(string script,
            string functionName = null, IEnumerable<object> iParameters = null)
        {
            if (string.IsNullOrEmpty(functionName))
            {//如果没有传入functionName，则构造一个以消除执行错误
                functionName = "ExecuteVBAScript_helper";
                script += "\r\nFunction ExecuteVBAScript_helper()\r\nEnd Function";
            }
            var res = App.SystemService.Evaluate(script,
                INFITF.CATScriptLanguage.CATVBScriptLanguage, functionName, null);
            return res;
        }

    }
}
