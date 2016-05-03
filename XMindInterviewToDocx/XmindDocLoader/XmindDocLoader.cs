using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using XMindInterviewToDocx.BusinessObjects;

namespace XMindInterviewToDocx.XmindDocLoader
{
    class XmindDocLoader
    {
        XmlDocument xmlDoc;
        Interview interview;

        public XmindDocLoader(XmlDocument xmlDoc)
        {
            this.xmlDoc = xmlDoc;
        }

        public XmindDocLoader(string xmindDocPath)
        {
            xmlDoc = new XmlDocument();

            ZipArchive zipArchive;

            try
            {
                zipArchive = ZipFile.OpenRead(xmindDocPath);
            }
            catch(Exception ex)
            {
                throw; 
            }

            ZipArchiveEntry zipArchiveEntry = zipArchive.GetEntry("content.xml");
            xmlDoc.Load(zipArchiveEntry.Open());
        }

        public Interview GenerateInterview()
        {
            if(xmlDoc == null)
            {
                throw new Exception("XmlDocument not loaded.");
            }

            return new Interview(xmlDoc.FirstChild.InnerText);
        }

    }

}
