using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChatLogger.Models
{
    public class Invites
    {
        public int ID { get; set; }
        public string Invitee { get; set; }
        public string Invited { get; set; }
        public string Date { get; set; }
    }
}
