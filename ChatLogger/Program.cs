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
            List<ChatLog> chatRecord = new List<ChatLog>();
            //List<Announcements> announcements = new List<Announcements>();
            Logger log = new Logger();
            // this is a change.
            try
            {
                using (StreamReader sr = new StreamReader(@"E:\Hobby\KakaoExport.txt"))
                {
                    String line = sr.ReadToEnd();
                    int messageID = 0;
                    int loopNumber = 0;

                    foreach (string rawr in new LineReader(() => new StringReader(line)))
                    {
                        loopNumber++;
                        if (rawr == "")
                            continue;

                        char firstCharacter = rawr[0];

                        if (firstCharacter == 0)
                        {
                            return;
                        }
                        else if (firstCharacter == '[')
                        {
                            messageID++;
                            updateChatLog(ref chatRecord, messageID, log, rawr);
                            Console.WriteLine("Saved new message. Loop # is " + loopNumber);
                        }
                        else if (firstCharacter == '-')
                        {
                            // TODO: Include date verifier
                            log.setCurrentDate(rawr);
                            Console.WriteLine("Saved new dates.");
                        }
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
                        else
                        {
                            ChatLog chat = chatRecord.Last();
                            chat.Message += " / " + rawr;
                            Console.WriteLine("Updated message: " + messageID);
                        }
                    }
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
                // The reading of the .txt file has been completed. How would you like to save it?
                Console.WriteLine("Reading complete. What would you like to do?");
                Console.WriteLine("'L': Chat log to excel");
                Console.WriteLine("'I': Invitet log to excel");
                Console.WriteLine("'A': Announcement log to excel");

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

                Console.WriteLine("Completed");
                GC.Collect();
            }
        }
        /// <summary>
        /// Updates the master chat log
        /// </summary>
        /// <param name="chatRecord">Chat record that needs to be updated</param>
        /// <param name="messageID">Message ID</param>
        /// <param name="log">The ref class</param>
        /// <param name="rawr">Input new message</param>
        private static void updateChatLog(ref List<ChatLog> chatRecord, int messageID, Logger log, string rawr)
        {

            if (rawr.Contains("공지") || rawr.Contains("Announcement"))
            {
                // TODO: include announcement code here, currently placeholder that only increases the text table
                //announcements.Add(new Announcements { announcement = rawr });
                log.setAnnouncement(rawr);
                Console.WriteLine("New announcement!");
            }

            else
            {
                int start = rawr.IndexOf('[');
                int end = rawr.IndexOf(']');

                char[] delimiterChars = { '[', ']' };

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
        }
    }
}
