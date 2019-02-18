using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Project_Template_VSTO_VBA
{
    [System.Runtime.InteropServices.ComVisible(true)]
    public interface IDocumentUtilities
    {
        void BindDataToExcel(Range range, object[,] response);
        object CitiesinZA();
        ArrayList GetComplexArray();
        IFund GetComplexObject();
        ArrayList GetSimpleArray();
        List<IFund> RunReaderPages();
        ArrayList getListofObj();
        //IFund[] ListFund { get; set; }
        object getString(ref object gg);
    }
}