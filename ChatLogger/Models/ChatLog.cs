using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChatLogger.Models
{
    public class ChatLog
    {
        public int ID { get; set; }
        public string Time { get; set; }
        public string Name { get; set; }
        public string Message { get; set; }
        public string Date { get; set; }
        //public bool isEmoticon { get; set; }
    }
}
