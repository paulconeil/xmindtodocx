using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMindInterviewToDocx.BusinessObjects
{
    class Response
    {
        List<Response> followup;
        string response = "";


        public Response(string responseValue)
        {
            this.response = responseValue;
        }

        internal List<Response> Responses
        {
            get
            {
                return followup;
            }
        }
    }
}
