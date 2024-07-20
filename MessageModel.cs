using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tool_TrainingGPT.cs
{
    public class MessageModel
    {
        public string role { get; set; }
        public string content { get; set; }

        public MessageModel()
        {
            this.role = String.Empty;
            this.content = String.Empty;
        }
    }

    public class ListMessages
    {
        public List<MessageModel> messages { get; set; }
        public ListMessages()
        {
            this.messages = new List<MessageModel>();
        }
    }
}
