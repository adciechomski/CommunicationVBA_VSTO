using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_Template_VSTO_VBA
{
    [System.Runtime.InteropServices.ComVisible(true)]
    [System.Runtime.InteropServices.ClassInterface(
System.Runtime.InteropServices.ClassInterfaceType.None)]
    public class VbaCommunication : IVbaCommunication
    {
        string naming { get; set; }
        public object getString(ref object gg)
        {
            return gg;
        }
        public string[] CitiesinZA()
        {
            naming = "123";
            string[] cities = new string[] { "Johannesburg", "Pretoria", "Cape Town" };
            return cities;
        }
        public void BindDataToExcel(Microsoft.Office.Interop.Excel.Range range, object[,] response)
        {
            int rows = response.GetLength(0);
            int cols = response.GetLength(1);
            int n = 0;
            Microsoft.Office.Interop.Excel.Range newRange = range.get_Offset(n, 0).get_Resize(rows - n, cols);
            newRange.Value = response;
        }
        public ArrayList GetSimpleArray()
        {
            ArrayList arr = new ArrayList();
            arr.Add(3);
            arr.Add(2);
            return arr;
        }
        public ArrayList GetComplexArray()
        {
            ArrayList arr = new ArrayList();
            arr.Add(new DataTransferObject() { Id = "2" });
            arr.Add(new DataTransferObject() { Id = "3" });
            return arr;
        }
        public IDataTransferObject GetComplexObject()
        {
            return new DataTransferObject() { Id = "2" };
        }
        public IDataTransferObject[] ListDataTransferObject()
        {
            IDataTransferObject[] iDataTransferObjectlist = new DataTransferObject[1];
            iDataTransferObjectlist[0] = GetComplexObject();
            //IListDataTransferObject iDataTransferObjectL = new ArrayList();
            return iDataTransferObjectlist;
        }
        public ArrayList getListofObj()
        {
            ArrayList arr = new ArrayList();
            arr.Add(new DataTransferObject() { Id = "2" });
            arr.Add(new DataTransferObject() { Id = "3" });
            return arr;
        }
        public List<IDataTransferObject> RunReaderPages()
        {
            /* 1. make a List of PdfReaderPage elements */
            List<IDataTransferObject> csharp_list = new List<IDataTransferObject>()
        {
            new DataTransferObject() { Id = "2" },
            new DataTransferObject() { Id = "3", Name = "adam" },
            new DataTransferObject() { Id = "2" }
        };
            /* 2. convert it into a vb collection */
            // var vb_coll = new Microsoft.VisualBasic.Collection();
            // csharp_list.ForEach(x => vb_coll.Add(x));
            /* 3. deliver it */
            //return vb_coll;
            return csharp_list;
        }
    }

    [System.Runtime.InteropServices.ComVisible(true)]
    public interface IDataTransferObject
    {
        string Id { get; set; }
        string Name { get; set; }
        string[,] arr { get; set; }
    }

    [System.Runtime.InteropServices.ComVisible(true)]
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public sealed class DataTransferObject : IDataTransferObject
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string[,] arr { get; set; }

        public DataTransferObject()
        {
            string[,] countries =  { { "Johannesburg", "Pretoria", "Cape Town" }, { "asd", "ada", "ola"} };
            this.arr = countries;
        }

        public DataTransferObject(string id, string name)
        {
            this.Id = id;
            this.Name = name;
        }
    }
}
