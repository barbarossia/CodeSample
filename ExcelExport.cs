using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Web;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Xsl;

namespace Common
{

    public abstract class ExcelExport<T> where T : class
    {
        protected const string spreadsheetML = @"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        protected const string relationSchema = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        protected const string workbookContentType = @"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        protected const string worksheetContentType = @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        protected const string styleContentType = @"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
        protected const string stringsContentType = @"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";


        public Template Template { get; set; }
        public IEnumerable<T> Data { get; set; }

        public ExcelExport()
        {

        }

        public string Export()
        {
            //WriteFromXml();
            return this.ExportCore();
        }

        protected virtual string ExportCore()
        {


            return GenerateExcelFile();

        }

        private void WriteFromXml()
        {
            DataContractSerializer s = new DataContractSerializer(typeof(List<T>));
            using (FileStream fs = File.Open(string.Concat("D:", "/", "text.xml"), FileMode.Create))
            {
                s.WriteObject(fs, Data);
            }
        }

        //protected abstract IEnumerable<T> GetData();

        private string GenerateExcelFile()
        {
            XNamespace aw = spreadsheetML;
            XNamespace r = relationSchema;
            // create the worksheet
            XDocument xmlStartPart = new XDocument(
                new XElement(aw + "workbook", new XAttribute(XNamespace.Xmlns + "r", relationSchema),
                    new XElement(aw + "sheets",
                        CreateElements(new[] { this.Template.SheetName }, aw, r)
                            ))
                );

            XDocument xmlStylePart = XDocument.Load(HttpContext.Current.Server.MapPath(this.Template.StylePath));
            string fileName = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".xlsx";
            // create the package (i.e., the document container)
            using (Package pkgOutputDoc = Package.Open(fileName, FileMode.Create, FileAccess.ReadWrite))
            {
                // save a temporary part to create the default application/xml content types
                Uri uriDefaultContentType = new Uri("/default.xml", UriKind.Relative);
                PackagePart partTemp = pkgOutputDoc.CreatePart(uriDefaultContentType, "application/xml");

                // save the main document part (workbook.xml)
                Uri uriStartPart = new Uri("/xl/workbook.xml", UriKind.Relative);
                PackagePart partWorkbookXML = pkgOutputDoc.CreatePart(uriStartPart, workbookContentType);
                using (StreamWriter streamStartPart = new StreamWriter(partWorkbookXML.GetStream(FileMode.Create, FileAccess.Write)))
                {
                    xmlStartPart.Save(streamStartPart);
                    streamStartPart.Close();
                    pkgOutputDoc.Flush();
                }

                // create the relationship parts
                pkgOutputDoc.CreateRelationship(uriStartPart, TargetMode.Internal, relationSchema + "/officeDocument", "rId1");

                Uri uriStylePart = new Uri("/xl/styles.xml", UriKind.Relative);
                PackagePart partStyleXML = pkgOutputDoc.CreatePart(uriStylePart, styleContentType);
                using (StreamWriter streamStylePart = new StreamWriter(partStyleXML.GetStream(FileMode.Create, FileAccess.Write)))
                {
                    xmlStylePart.Save(streamStylePart);
                    streamStylePart.Close();
                    pkgOutputDoc.Flush();
                }

                partWorkbookXML.CreateRelationship(uriStylePart, TargetMode.Internal, relationSchema + "/styles", "rId2");

                Uri uriWorksheet = new Uri("/xl/worksheets/sheet1.xml", UriKind.Relative);
                PackagePart partWorksheetXML = pkgOutputDoc.CreatePart(uriWorksheet, worksheetContentType);

                XDocument xmlWorksheet;
                Stream stream = CreateSheetData();

                stream.Position = 0;
                xmlWorksheet = XDocument.Load(stream);
                using (StreamWriter streamWorksheet = new StreamWriter(partWorksheetXML.GetStream(FileMode.Create, FileAccess.Write)))
                {
                    xmlWorksheet.Save(streamWorksheet);
                    streamWorksheet.Close();
                    pkgOutputDoc.Flush();
                }

                partWorkbookXML.CreateRelationship(uriWorksheet, TargetMode.Internal, relationSchema + "/worksheet", "rId1");

                // remove the temporary part that created the default xml content type
                pkgOutputDoc.DeletePart(uriDefaultContentType);

                // close the document
                pkgOutputDoc.Flush();
                pkgOutputDoc.Close();

                return fileName;
            }
        }

        protected Stream CreateSheetData()
        {
            using (MemoryStream input = new MemoryStream())
            {
                var dcs = new DataContractSerializer(Data.GetType()); // serialize 'p' in 'input'
                dcs.WriteObject(input, Data);

                input.Position = 0;

                //Create a new XslTransform object.
                XslCompiledTransform xslt = new XslCompiledTransform();
                XsltArgumentList argsList = CreateXsltArgumentList();
                var settings = new XsltSettings();
                settings.EnableScript = true;
                xslt.Load(HttpContext.Current.Server.MapPath(this.Template.XSLTPath), settings, null);

                XmlWriterSettings writerSettings = new XmlWriterSettings();

                MemoryStream stream = new MemoryStream();
                //Create an XmlTextWriter which outputs to the memory stream.
                using (XmlWriter writer = XmlWriter.Create(stream, writerSettings))
                {
                    //Transform the file and send the output to the memory stream.                   
                    xslt.Transform(XmlReader.Create(input), argsList, writer);
                }

                return stream;

            }

        }

        private XsltArgumentList CreateXsltArgumentList()
        {
            XsltArgumentList argsList = new XsltArgumentList();
            foreach (var arg in Template.XSLTArgs)
            {
                argsList.AddParam(arg.Key, "", arg.Value);
            }
            return argsList;
        }


        protected XElement[] CreateElements(string[] elementName, XNamespace aw, XNamespace r)
        {
            XElement[] list = null;
            // Set values of sheets
            if (elementName != null)
            {
                var length = elementName.Length;
                list = new XElement[length];
                for (int i = 0; i < length; i++)
                {
                    list[i] = new XElement(aw + "sheet",
                            new XAttribute("name", elementName[i]),
                            new XAttribute("sheetId", i + 1),
                            new XAttribute(r + "id", "rId" + (i + 1))
                            );
                }
            }

            return list;
        }



    }
    

    public struct Template
    {
        public string Style { get; set; }
        public string XSLT { get; set; }
        public string SheetName { get; set; }
        public string Location { get; set; }
        public Dictionary<string, string> XSLTArgs { get; set; }

        public string StylePath
        {
            get
            {
                return string.Concat(Location, "/", Style);
            }
        }
        public string XSLTPath
        {
            get
            {
                return string.Concat(Location, "/", XSLT);
            }
        }
    }
}
