using Microsoft.VisualBasic;
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
    public class DocumentUtilities : IDocumentUtilities
    {
        string naming { get; set; }
        public object getString(ref object gg)
        {
            return gg;
        }
        public object CitiesinZA()
        {
            naming = "123";
            string[] cities = new string[] { "Johannesburg", "Pretoria", "Cape Town" };
            return this;// (object)cities;
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
            arr.Add(new Fund() { Id = "2" });
            arr.Add(new Fund() { Id = "3" });
            return arr;
        }
        public IFund GetComplexObject()
        {
            return new Fund() { Id = "2" };
        }
        public IFund[] ListFund()
        {
            IFund[] iFundlist = new Fund[1];
            iFundlist[0] = GetComplexObject();
            //IListFund iFundL = new ArrayList();
            return iFundlist;
        }
        public ArrayList getListofObj()
        {
            ArrayList arr = new ArrayList();
            arr.Add(new Fund() { Id = "2" });
            arr.Add(new Fund() { Id = "3" });
            return arr;
        }
        public List<IFund> RunReaderPages()
        {
            /* 1. make a List of PdfReaderPage elements */
            List<IFund> csharp_list = new List<IFund>()
        {
            new Fund() { Id = "2" },
            new Fund() { Id = "3" },
            new Fund() { Id = "2" }
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
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public class PdfReaderPage
    {
        public int Foo;
    }

    [System.Runtime.InteropServices.ComVisible(true)]
    public interface IFund
    {
        string Id { get; set; }
        string Name { get; set; }
    }

    [System.Runtime.InteropServices.ComVisible(true)]
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public sealed class Fund : IFund
    {
        public string Id { get; set; }
        public string Name { get; set; }

        public Fund()
        {
        }

        public Fund(string id, string name)
        {
            this.Id = id;
            this.Name = name;
        }
    }
    [System.Runtime.InteropServices.ComVisible(true)]
    public interface IListFund
    {
        IFund[] ListFund { get; set;}
    }
}
