using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_Template_VSTO_VBA
{
    [System.Runtime.InteropServices.ComVisible(true)]
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public sealed class DataTransferObject : IDataTransferObject, IDataTransferObject1
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string[,] arr { get; set; }

        public DataTransferObject()
        {
            string[,] countries = { { "Johannesburg", "Pretoria", "Cape Town" }, { "asd", "ada", "ola" } };
            this.arr = countries;
        }

        public DataTransferObject(string id, string name)
        {
            this.Id = id;
            this.Name = name;
        }
    }
}
