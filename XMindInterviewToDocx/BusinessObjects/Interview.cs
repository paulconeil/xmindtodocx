using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMindInterviewToDocx.BusinessObjects
{
    class Interview
    {
        List<Topic> topics;
        string title;

        public Interview(string title)
        {
            this.title = title;
        }

        public void AddTopic(string topicValue)
        {
            if (topics == null)
                topics = new List<Topic>();
            topics.Add(new Topic(topicValue));
        }
    }
}
