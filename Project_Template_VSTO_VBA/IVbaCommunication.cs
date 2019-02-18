using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.Collections;

namespace Project_Template_VSTO_VBA
{
    [System.Runtime.InteropServices.ComVisible(true)]
    public interface IVbaCommunication
    {
        void BindDataToExcel(Range range, object[,] response);
        string[] CitiesinZA();
        ArrayList GetComplexArray();
        IDataTransferObject GetComplexObject();
        ArrayList getListofObj();
        ArrayList GetSimpleArray();
        object getString(ref object gg);
        IDataTransferObject[] ListDataTransferObject();
        List<IDataTransferObject> RunReaderPages();
    }
}