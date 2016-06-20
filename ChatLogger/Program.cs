using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using MiscUtil.IO;
using ChatLogger.Models;
using ChatLogger.Lib;
using ChatLogger.MiscUtil;

namespace ChatLogger
{
    class Program
    {
        static void Main(string[] args)
        {
            #region Init
            // Init the ChatRecord that will record all the chat messages. Probably should add it to Logger() instead in Logger class.
            List<ChatLog> chatRecord = new List<ChatLog>();

            // Inits the Logger class as well as the associated lists
            Logger log = new Logger();
            #endregion

            try
            {
                // TODO: Allow the user to specify where their local Kakao chat record file is. Through Console.ReadLine maybe? 
                // read the txt file.
                using (StreamReader sr = new StreamReader(@"E:\Hobby\KakaoExport.txt"))
                {
                    String line = sr.ReadToEnd();

                    #region Define ID
                    // Message ID is the number of messages passed, this is excluding invites, announcements, leavers, etc.
                    int messageID = 0;

                    // Loop # defines the amount of times the system has read the .txt file, or the "line #" of the file.
                    int loopNumber = 0;
                    #endregion

                    #region Loop through read the file
                    // Start loop on reading through each line of the .txt file
                    foreach (string rawr in new LineReader(() => new StringReader(line)))
                    {
                        loopNumber++;

                        // defines an emtpy line where the string literally contains nothing. 
                        if (rawr == "")
                            continue;

                        // gets the first character of the input string
                        char firstCharacter = rawr[0];

                        #region Kakao messages
                        // tracks if the passed line was a message
                        if (firstCharacter == '[')
                        {
                            messageID++;
                            updateChatLog(ref chatRecord, messageID, log, rawr);
                            Console.WriteLine("Saved new message. Loop # is " + loopNumber);
                        }
                        #endregion

                        #region Update date
                        // In most cases, Kakao denotes the dates with the format ;---------------Week, Date, Month, Year------------------', which satisfies this argument
                        else if (firstCharacter == '-')
                        {
                            // TODO: Include date verifier. Currently does not verify if the data passed was a date or not. If it was not a date, it should call the updateChatLog(ref List<ChatLog> chatRecord, int messageID, Logger log, string rawr) function
                            // However, this requires a nested if statement... sigh
                            log.setCurrentDate(rawr);
                            Console.WriteLine("Saved new dates.");
                        }
                        #endregion

                        #region Invite & Leavers
                        // If the string failes to satisfy any of the above statements, this is when we detect the invite/leavers
                        else if (rawr.Contains("invited"))
                        {
                            log.setNewInvite(rawr);
                            Console.WriteLine("Saved new invites.");
                        }
                        else if (rawr.Contains("left"))
                        {
                            // include left code here
                            log.updateLeaver(rawr);
                            Console.WriteLine("Oh no! Someone left.");
                        }
                        #endregion

                        #region None of the above
                        else
                        {
                            ChatLog chat = chatRecord.Last();
                            chat.Message += " / " + rawr;
                            Console.WriteLine("Updated message: " + messageID);
                        }
                        #endregion
                    }
                    #endregion
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("The file could not be read:");
                Console.WriteLine(e.Message);
                Console.WriteLine("Last completed operation: " + chatRecord.Last().Message);
            }
            finally
            {
                #region Write after reading
                // The reading of the .txt file has been completed. How would you like to save it?
                Console.WriteLine("Reading complete. What would you like to do?");
                Console.WriteLine("'C': Chat log to excel");
                Console.WriteLine("'I': Invitet log to excel");
                Console.WriteLine("'A': Announcement log to excel");
                Console.WriteLine("'L': Leaver log to excel");

                string decision = Console.ReadLine();

                switch(decision)
                {
                    case "C":
                        Excelutlity.ExportChatLogToExcel(chatRecord);
                        break;
                    case "I":
                        Excelutlity.ExportInvitesToExcel(log.inviteRecord);
                        break;
                    case "A":
                        Excelutlity.ExportAnnouncementToExcel(log.announcementRecord);
                        break;
                    case "L":
                        Excelutlity.ExportLeaversToExcel(log.leaverRecord);
                        break;
                    default:
                        Console.WriteLine("Error! Defaulting to all logs...");
                        Excelutlity.ExportChatLogToExcel(chatRecord);
                        Excelutlity.ExportInvitesToExcel(log.inviteRecord);
                        Excelutlity.ExportAnnouncementToExcel(log.announcementRecord);
                        Excelutlity.ExportLeaversToExcel(log.leaverRecord);
                        break;
                }
                #endregion

                Console.WriteLine("Completed");
                GC.Collect();
            }
        }
        /// <summary>
        /// Updates the master chat log list
        /// </summary>
        /// <param name="chatRecord">Chat record that needs to be updated</param>
        /// <param name="messageID">Message ID</param>
        /// <param name="log">The ref class</param>
        /// <param name="rawr">Input new message</param>
        private static void updateChatLog(ref List<ChatLog> chatRecord, int messageID, Logger log, string rawr)
        {
            #region Announcement
            // Check if the string passed was an announcement, do not save as a message.
            if (rawr.Contains("공지") || rawr.Contains("Announcement"))
            {
                // TODO: include announcement code here, currently placeholder that only increases the text table
                //announcements.Add(new Announcements { announcement = rawr });
                log.setAnnouncement(rawr);
                Console.WriteLine("New announcement!");
            }
            #endregion

            #region Real message
            else
            {
                char[] delimiterChars = { '[', ']' };

                // split the line of text that's in [name] [time] "message" format to an array
                string[] words = rawr.Split(delimiterChars);
                string name = words[1];
                string time = words[3];

                try
                {
                    string message = words[4];
                    ChatLog currentLog = new ChatLog { ID = messageID, Name = name, Time = time, Message = message };
                    log.AddNewLog(currentLog, ref chatRecord);
                }
                catch
                {
                    string error = "An error cocured";
                }
            }
            #endregion 
        }
    }
}
