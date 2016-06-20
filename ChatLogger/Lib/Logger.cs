using ChatLogger.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ChatLogger.Lib
{
    public class Logger
    {
        /// <summary>
        /// The current date for the class at the time of the loop. Updated per new date by private setCurrentDate()
        /// </summary>
        public string currentDate { get; set; }

        /// <summary>
        /// The list of invited people 
        /// </summary>
        public List<Invites> inviteRecord { get; set; }

        /// <summary>
        ///  Collection of announcements
        /// </summary>
        public List<Announcements> announcementRecord { get; set; }

        /// <summary>
        /// Leaver record
        /// </summary>
        public List<Leave> leaverRecord { get; set; }

        // The invite ID
        private int inviteID;

        /// <summary>
        /// Initiate logger with empty lists
        /// </summary>
        public Logger()
        {
            inviteID = 1;
            inviteRecord = new List<Invites>();
            leaverRecord = new List<Leave>();
            announcementRecord = new List<Announcements>();
        }

        /// <summary>
        /// Update the leaver's list
        /// </summary>
        /// <param name="rawr">Leaver string. The string from the .txt file</param>
        public void updateLeaver(string rawr)
        {
            char[] delimiterChars = { ' ' };
            string[] words = rawr.Split(delimiterChars);
            var leaver = new Leave();
            leaver.Name = words[0];
            leaver.Date = currentDate;
        }

        /// <summary>
        /// New addition to the chatlist.
        /// </summary>
        /// <param name="log">Item to be added</param>
        /// <param name="CurrentList">Ref variable, current list</param>
        public void AddNewLog(ChatLog log, ref List<ChatLog> CurrentList)
        {
            log.Date = currentDate;
            CurrentList.Add(log);
        }

        /// <summary>
        /// Update the current date for the iteration
        /// TODO: Requires update on the date verification system
        /// </summary>
        /// <param name="date">New date</param>
        public void setCurrentDate(string date)
        {
            currentDate = Regex.Replace(date, "[^a-zA-Z,0-9 ]", "");
            currentDate = currentDate.Substring(1);
        }

        /// <summary>
        /// Add new addition to the announcement list
        /// </summary>
        /// <param name="newAnnouncement">New announcement to be added</param>
        public void setAnnouncement(string newAnnouncement)
        {
            char[] delimiterChars = { '[', ']' };

            string[] words = newAnnouncement.Split(delimiterChars);

            var announcement = new Announcements();
            announcement.Announcement = words[6];
            announcement.Date = currentDate;
            announcement.Time = words[3];
            announcement.Name = words[1];
            announcementRecord.Add(announcement);
        }

        /// <summary>
        /// Add a new invite record to the list
        /// </summary>
        /// <param name="invite">Invite to be added</param>
        public void setNewInvite(string invite)
        {
            int start = invite.IndexOf("invited");
            string invitee = invite.Substring(0, start - 1);

            string[] namesArray = invite.Substring(start + 7).Split(',');
            List<string> namesList = new List<string>(namesArray.Length);
            namesList.AddRange(namesArray);
            namesList.Reverse();
            namesList = namesList.Distinct().ToList();

            foreach (var name in namesList)
            {
                Invites invite_obj = new Invites();
                invite_obj.Invitee = invitee;
                invite_obj.Invited = name;
                invite_obj.ID = inviteID++;
                invite_obj.Date = currentDate;
                inviteRecord.Add(invite_obj);
            }

        }
    }
}
