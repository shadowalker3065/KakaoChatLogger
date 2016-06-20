using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using ChatLogger.Models;

namespace ChatLogger.MiscUtil
{
    public class Excelutlity
    {
        public static void ExportChatLogToExcel(List<ChatLog> chat)
        {
            // creating Excel Application
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();

            // creating new WorkBook within Excel application
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);

            // creating new Excelsheet in workbook
            Microsoft.Office.Interop.Excel._Worksheet workSheet = null;

            Sheets xlSheets = null;

            xlSheets = workbook.Sheets as Sheets;
            app.Visible = true;

            workSheet = workbook.Worksheets[1];
            workSheet.Name = "Chat Record Sheet";

            // I created Application and Worksheet objects before try/catch,
            // so that i can close them in finnaly block.
            // It's IMPORTANT to release these COM objects!!
            try
            {
                // ------------------------------------------------
                // Creation of header cells
                // ------------------------------------------------
                workSheet.Cells[1, "A"] = "Name";
                workSheet.Cells[1, "B"] = "Date";
                workSheet.Cells[1, "C"] = "Time";
                workSheet.Cells[1, "D"] = "Message";

                // ------------------------------------------------
                // Populate sheet with some real data from "cars" list
                // ------------------------------------------------
                int row = 2; // start row (in row 1 are header cells)
                foreach (ChatLog record in chat)
                {
                    workSheet.Cells[row, "A"] = record.Name;
                    workSheet.Cells[row, "B"] = record.Date;
                    workSheet.Cells[row, "C"] = record.Time;
                    workSheet.Cells[row, "D"] = record.Message;

                    row++;
                }

                Console.WriteLine("Completed chat history! Now processing invites...");
                // Apply some predefined styles for data to look nicely :)
                workSheet.Range["A1"].AutoFormat(Microsoft.Office.Interop.Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic1);

                // Define filename
                string fileName = string.Format(@"{0}\ExcelData.xlsx", Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

                // Save this data as a file
                workSheet.SaveAs(fileName);

                // Display SUCCESS message
                //MessageBox.Show(string.Format("The file '{0}' is saved successfully!", fileName));
            }
            catch (Exception exception)
            {
                Console.WriteLine("There was an errror!");
                Console.WriteLine("Error Message: " + exception);
               // MessageBox.Show("Exception",
                //"There was a PROBLEM saving Excel file!\n" + exception.Message,
                //MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Quit Excel application
                app.Quit();

                // Release COM objects (very important!)
                if (app != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);

                if (workSheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet);

                // Empty variables
                app = null;
                workSheet = null;

                // Force garbage collector cleaning
                GC.Collect();
            }
        }

        public static void ExportInvitesToExcel(List<Invites> invites)
        {
            // creating Excel Application
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();

            // creating new WorkBook within Excel application
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);

            // creating new Excelsheet in workbook
            Microsoft.Office.Interop.Excel._Worksheet workSheet = null;

            Sheets xlSheets = null;

            xlSheets = workbook.Sheets as Sheets;
            app.Visible = true;

            workSheet = workbook.Worksheets[1];
            workSheet.Name = "Invite Record Sheet";

            // I created Application and Worksheet objects before try/catch,
            // so that i can close them in finnaly block.
            // It's IMPORTANT to release these COM objects!!
            try
            {
                // ------------------------------------------------
                // Creation of header cells
                // ------------------------------------------------
                workSheet.Cells[1, "A"] = "Date";
                workSheet.Cells[1, "B"] = "Invitee";
                workSheet.Cells[1, "C"] = "Invited";

                // ------------------------------------------------
                // Populate sheet with some real data from "cars" list
                // ------------------------------------------------
                int row = 2; // start row (in row 1 are header cells)
                foreach (Invites record in invites)
                {
                    workSheet.Cells[row, "A"] = record.Date;
                    workSheet.Cells[row, "B"] = record.Invitee;
                    workSheet.Cells[row, "C"] = record.Invited;

                    row++;
                }

                // Apply some predefined styles for data to look nicely :)
                workSheet.Range["A1"].AutoFormat(Microsoft.Office.Interop.Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic1);

                // Define filename
                string fileName = string.Format(@"{0}\ExcelData_inviteRecord.xlsx", Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

                // Save this data as a file
                workSheet.SaveAs(fileName);

                // Display SUCCESS message
                //MessageBox.Show(string.Format("The file '{0}' is saved successfully!", fileName));
            }
            catch (Exception exception)
            {
                Console.WriteLine("There was an errror!");
                Console.WriteLine("Error Message: " + exception);
                // MessageBox.Show("Exception",
                //"There was a PROBLEM saving Excel file!\n" + exception.Message,
                //MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Quit Excel application
                app.Quit();

                // Release COM objects (very important!)
                if (app != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);

                if (workSheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet);

                // Empty variables
                app = null;
                workSheet = null;

                // Force garbage collector cleaning
                GC.Collect();
            }
        }

        public static void ExportAnnouncementToExcel(List<Announcements> announcement)
        {
            // creating Excel Application
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();

            // creating new WorkBook within Excel application
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);

            // creating new Excelsheet in workbook
            Microsoft.Office.Interop.Excel._Worksheet workSheet = null;

            Sheets xlSheets = null;

            xlSheets = workbook.Sheets as Sheets;
            app.Visible = true;

            workSheet = workbook.Worksheets[1];
            workSheet.Name = "Announcement Record Sheet";

            // I created Application and Worksheet objects before try/catch,
            // so that i can close them in finnaly block.
            // It's IMPORTANT to release these COM objects!!
            try
            {
                // ------------------------------------------------
                // Creation of header cells
                // ------------------------------------------------
                workSheet.Cells[1, "A"] = "Name";
                workSheet.Cells[1, "B"] = "Date";
                workSheet.Cells[1, "C"] = "Time";
                workSheet.Cells[1, "D"] = "Announcement";

                // ------------------------------------------------
                // Populate sheet with some real data from "cars" list
                // ------------------------------------------------
                int row = 2; // start row (in row 1 are header cells)
                foreach (Announcements record in announcement)
                {
                    workSheet.Cells[row, "A"] = record.Name;
                    workSheet.Cells[row, "B"] = record.Date;
                    workSheet.Cells[row, "C"] = record.Time;
                    workSheet.Cells[row, "D"] = record.Announcement;

                    row++;
                }

                // Apply some predefined styles for data to look nicely :)
                workSheet.Range["A1"].AutoFormat(Microsoft.Office.Interop.Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic1);

                // Define filename
                string fileName = string.Format(@"{0}\ExcelData_announcementRecord.xlsx", Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

                // Save this data as a file
                workSheet.SaveAs(fileName);

                // Display SUCCESS message
                //MessageBox.Show(string.Format("The file '{0}' is saved successfully!", fileName));
            }
            catch (Exception exception)
            {
                Console.WriteLine("There was an errror!");
                Console.WriteLine("Error Message: " + exception);
                // MessageBox.Show("Exception",
                //"There was a PROBLEM saving Excel file!\n" + exception.Message,
                //MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Quit Excel application
                app.Quit();

                // Release COM objects (very important!)
                if (app != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);

                if (workSheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet);

                // Empty variables
                app = null;
                workSheet = null;

                // Force garbage collector cleaning
                GC.Collect();
            }
        }

        public static void ExportLeaversToExcel(List<Leave> leaver)
        {
            // creating Excel Application
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();

            // creating new WorkBook within Excel application
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);

            // creating new Excelsheet in workbook
            Microsoft.Office.Interop.Excel._Worksheet workSheet = null;

            Sheets xlSheets = null;

            xlSheets = workbook.Sheets as Sheets;
            app.Visible = true;

            workSheet = workbook.Worksheets[1];
            workSheet.Name = "Announcement Record Sheet";

            // I created Application and Worksheet objects before try/catch,
            // so that i can close them in finnaly block.
            // It's IMPORTANT to release these COM objects!!
            try
            {
                // ------------------------------------------------
                // Creation of header cells
                // ------------------------------------------------
                workSheet.Cells[1, "A"] = "Name";
                workSheet.Cells[1, "B"] = "Date";

                // ------------------------------------------------
                // Populate sheet with some real data from "cars" list
                // ------------------------------------------------
                int row = 2; // start row (in row 1 are header cells)
                foreach (Leave record in leaver)
                {
                    workSheet.Cells[row, "A"] = record.Name;
                    workSheet.Cells[row, "B"] = record.Date;

                    row++;
                }

                // Apply some predefined styles for data to look nicely :)
                workSheet.Range["A1"].AutoFormat(Microsoft.Office.Interop.Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic1);

                // Define filename
                string fileName = string.Format(@"{0}\ExcelData_leaverRecord.xlsx", Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

                // Save this data as a file
                workSheet.SaveAs(fileName);

                // Display SUCCESS message
                //MessageBox.Show(string.Format("The file '{0}' is saved successfully!", fileName));
            }
            catch (Exception exception)
            {
                Console.WriteLine("There was an errror!");
                Console.WriteLine("Error Message: " + exception);
                // MessageBox.Show("Exception",
                //"There was a PROBLEM saving Excel file!\n" + exception.Message,
                //MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Quit Excel application
                app.Quit();

                // Release COM objects (very important!)
                if (app != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);

                if (workSheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet);

                // Empty variables
                app = null;
                workSheet = null;

                // Force garbage collector cleaning
                GC.Collect();
            }
        }
    }
}
