using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.ReportingServices.RdlObjectModel;
using Microsoft.ReportingServices.RdlObjectModel.Serialization;

namespace ReportObjectModel
{
    class Program
    {
        static void Main()
        {
            string idef = @"C:\Reports\Sales by Product.rdl"; // input report in RDL 2008 format
            string odef = @"C:\Reports\Sales by Product1.rdl"; // output report in RDL 2008 format

            //Report report = null;
            //RdlSerializer serializer;

            //if (!File.Exists(idef)) return;
            //// deserialize from disk
            //using (FileStream fs = File.OpenRead(idef))
            //{
            //    serializer = new RdlSerializer();
            //    report = serializer.Deserialize(fs);
            //}

            //report.Author = "Teo Lachev";
            //report.Description = "RDL Object Demo";
            //// TODO: use and abuse RDL as you wish
            //// serialize to disk
            //using (FileStream os = new FileStream(odef, FileMode.Create))
            //{
            //    serializer.Serialize(os, report);
            //}
        }
    }
}