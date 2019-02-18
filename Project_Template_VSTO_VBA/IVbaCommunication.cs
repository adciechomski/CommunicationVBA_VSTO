using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.Collections;

namespace Project_Template_VSTO_VBA
{
    [System.Runtime.InteropServices.ComVisible(true)]
    public interface IVbaCommunication
    {
        IDataTransferObject GetComplexObject();
        ArrayList getListofObj();
        ArrayList GetSimpleArray();
        string getString(string stringFromVBA);
        object objectToVSTO(ref object x);
        string[,] stringArrayToVSTO(ref string[,] response);
        Scripting.Dictionary passDictionary2VBA(string callName);
    }
}