using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace XMindInterviewToDocx
{
    public class Interview
    {
        XmlDocument xmindDoc;
        public Interview(XmlDocument xmindDoc) {
            this.xmindDoc = xmindDoc;
        }

        public void GetTopics()
        { }

        public void GetRequirements()
        { }
    }
}
