using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using Visio = Microsoft.Office.Interop.Visio;
using Microsoft.Office.Core;

namespace OfficeConvert
{
    public class VisioConverter : Converter
    {
        private Visio.Application app;
        private Visio.Documents docs;
        private Visio.Document doc;
        private Visio.Page page;
        // private Excel.Worksheet sheet;

        public VisioConverter()
        {                
        }

        public VisioConverter(string inputFile, string outputFile)
        {
            this.Convert(inputFile, outputFile);
        }

        public override void Convert(String inputFile, String outputFile)
        {
            Object nothing = Type.Missing;
            try
            {
                if (!File.Exists(inputFile))
                {
                    throw new ConvertException("File not Exists");
                }

                if (IsPasswordProtected(inputFile))
                {
                    throw new ConvertException("Password Exist");
                }

                app = new Visio.Application();
                docs = app.Documents; ;
                doc = docs.Open(inputFile);
                int pageCount = 0;
                List<string> pageNames = new List<string>();
                bool hasContent = false;
                foreach (Visio.Page vpage in doc.Pages)
                {
                    page = vpage;
                    pageNames.Add(page.Name);
                    pageCount++;    
                    if (pageCount > 0 && pageNames.Count > 0) 
                    {
                        hasContent = true;
                    }
                }

                if (!hasContent) throw new ConvertException("No Content");
                doc.ExportAsFixedFormat(Microsoft.Office.Interop.Visio.VisFixedFormatTypes.visFixedFormatPDF, outputFile, Visio.VisDocExIntent.visDocExIntentPrint, Visio.VisPrintOutRange.visPrintAll, 1, pageCount - 1, false, true, true, true, false, nothing);
            }
            catch (Exception e)
            {
                release();
                throw new ConvertException(e.Message);
            }

            release();
        }

        private void release()
        {
            if (page != null)
            {
                try
                {
                    releaseCOMObject(page);
                }
                catch (Exception e)
                {
                    Console.Error.WriteLine(e.Message + "\r\n" + e.ToString() + "\r\n" + e.StackTrace);
                }
            }

            if (doc != null)
            {
                try
                {
                    doc.Close();
                    releaseCOMObject(doc);
                }
                catch (Exception e)
                {
                    Console.Error.WriteLine(e.Message + "\r\n" + e.ToString() + "\r\n" + e.StackTrace);
                }
            }

            if (docs != null)
            {
                try
                {           
                    releaseCOMObject(docs);
                }
                catch (Exception e)
                {
                    Console.Error.WriteLine(e.Message + "\r\n" + e.ToString() + "\r\n" + e.StackTrace);
                }
            }

            if (app != null)
            {
                try
                {
                    app.Quit();
                    releaseCOMObject(app);
                }
                catch (Exception e)
                {
                    Console.Error.WriteLine(e.Message + "\r\n" + e.ToString() + "\r\n" + e.StackTrace);
                }
            }
        }
    }
}
