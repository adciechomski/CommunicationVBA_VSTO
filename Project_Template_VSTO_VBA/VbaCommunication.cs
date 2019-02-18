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
        public string getString(string stringFromVBA)
        {
            stringFromVBA = stringFromVBA + " 123";
            return stringFromVBA;
        }
        public string[,] stringArrayToVSTO(ref string[,] response)
        {
            response[1, 0] = "adam:";
            return response;
        }
        public object objectToVSTO(ref object x)
        {
            object ii = x.GetType().InvokeMember("ii", System.Reflection.BindingFlags.GetProperty, null, x, null);
            object ss = x.GetType().InvokeMember("ss", System.Reflection.BindingFlags.GetProperty, null, x, null);

            int number = System.Convert.ToInt32(ii);
            string value = System.Convert.ToString(ss);

            //System.Windows.Forms.MessageBox.Show(number + value);
            //x.ii = 12;
            return x;
        }

        public ArrayList GetSimpleArray()
        {
            ArrayList arr = new ArrayList();
            arr.Add(3);
            arr.Add(2);
            return arr;
        }
        public IDataTransferObject GetComplexObject()
        {
            return new DataTransferObject() { Id = "2" };
        }
        public ArrayList getListofObj()
        {
            ArrayList arr = new ArrayList();
            arr.Add(new DataTransferObject() { Id = "2" });
            arr.Add(new DataTransferObject() { Id = "3" });
            return arr;
        }
        public Scripting.Dictionary passDictionary2VBA(string callName)
        {
            Scripting.Dictionary dict = new Scripting.Dictionary();
            dict.Add("Apples", new DataTransferObject() { Id = callName });
            dict.Add("Oranges", new DataTransferObject() { Id = callName });
            return dict;
        }
    }

    [System.Runtime.InteropServices.ComVisible(true)]
    public interface IDataTransferObject
    {
        string Id { get; set; }
        string Name { get; set; }
        string[,] arr { get; set; }
    }


}
