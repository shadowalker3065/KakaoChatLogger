using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChatLogger.Models
{
    /// <summary>
    /// Chat record class
    /// </summary>
    public class ChatLog
    {
        /// <summary>
        /// The ID for the chat message
        /// </summary>
        public int ID { get; set; }

        /// <summary>
        /// Time when the message was sent
        /// </summary>
        public string Time { get; set; }

        /// <summary>
        /// Poster of the message
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// The message itself
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// Date posted
        /// </summary>
        public string Date { get; set; }
    }
}
