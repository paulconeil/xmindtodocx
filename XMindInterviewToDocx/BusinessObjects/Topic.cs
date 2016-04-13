using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMindInterviewToDocx.BusinessObjects
{
    class Topic
    {
        List<Response> responses;
        string topicValue;

        public Topic(string topicValue)
        {
            this.topicValue = topicValue;
        }
        
        public void AddResponse(string response)
        {
            if(responses == null)
            {
                responses = new List<Response>();
            }
            responses.Add(new Response(response));
        }


    }
}
