using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

// Made by Guy Kaplan (23.3.20)

namespace P4V_get_CLs_info
{
    class Program
    {
        private static readonly List<string> months = new List<string>(new string[]
            { "jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"});

        static void Main(string[] args)
        {
            var textSB = new StringBuilder($"{Environment.NewLine}");
            textSB.AppendLine("***** Written by Guy Kaplan *****");
            textSB.AppendLine("This program displays info of all P4V CLs from a given month");
            textSB.AppendLine("for <stream_name> ONLY - and saves the data table in an Excel file.");
            textSB.AppendLine("Please enter month and year in this format: <month> <year>");
            textSB.AppendLine("Example: Mar 2020");
            Console.WriteLine(textSB.ToString());
            string userInput = Console.ReadLine();

            while (!IsUserInputValid(userInput))
            {
                Console.WriteLine("Wrong input. Please enter valid month and year: (i.e Mar 2020)");
                userInput = Console.ReadLine();
            }

            string[] input = userInput.Split(' ');
            string month = input[0];
            string year = input[1];

            // Create Excel file with all CLs info
            CreateExcelSheetAndTable(month, year);

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }//end of Main

        /// <summary>
        /// Checks whether the user input string represents a valid month,
        /// by the template: "-month- -year-". Example: "Mar 2020"
        /// </summary>
        /// <param name="userInput">The input line the user entered. Needs to be a valid month</param>
        /// <returns>True if the input string matches the template. False otherwise</returns>
        static bool IsUserInputValid(string userInput)
        {
            // Make sure userInput has 8 chars and ONLY 1 space in it
            if (userInput.Length != 8 || userInput.Count(char.IsWhiteSpace) != 1)
            {
                return false;
            }

            string[] input = userInput.Split(' ');
            string month = input[0];
            string yearStr = input[1];

            // Month has to be 3 chars long, and Year 4 chars
            if (month.Length != 3 || yearStr.Length != 4)
            {
                return false;
            }

            // Check if given year is a number
            if (!int.TryParse(yearStr, out int year))
            {
                return false;
            }

            // Check if month is valid
            if (!months.Contains(month.ToLower()))
            {
                return false;
            }

            // Check if year is valid
            if (year > DateTime.Today.Year)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Executes the given cmd command and returns its cmd output
        /// </summary>
        /// <param name="cmdToRun">The cmd command to run</param>
        /// <returns>The output of the execution of the given command</returns>
        static string GetCmdCommandOutput(string cmdToRun)
        {
            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo()
            {
                WindowStyle = ProcessWindowStyle.Hidden,
                FileName = "cmd.exe",
                Arguments = $"/c {cmdToRun}",
                UseShellExecute = false,
                RedirectStandardOutput = true
            };
            process.StartInfo = startInfo;
            process.Start();

            // Get the output into a string
            return process.StandardOutput.ReadToEnd();
        }

        /// <summary>
        /// <para>Gets a string and two string separators and returns the string between these separators.</para>
        /// <para>If one of the separators is missing from the original string, returns an empty string.</para>
        /// </summary>
        /// <param name="token">The original string</param>
        /// <param name="first">The first string separator</param>
        /// <param name="second">The second string separator</param>
        /// <returns>
        /// The string between the two given separators.
        /// If one of the separators is missing from the original string, returns an empty string.
        /// </returns>
        static string GetStringBetween(string token, string first, string second)
        {
            if (!token.Contains(first)) return string.Empty;

            var afterFirst = token.Split(new[] { first }, StringSplitOptions.None)[1];

            if (!afterFirst.Contains(second)) return string.Empty;

            var result = afterFirst.Split(new[] { second }, StringSplitOptions.None)[0];

            return result;
        }

        /// <summary>
        /// Creates an Excel document with one sheet and one table with the P4V data
        /// of the given month and year
        /// </summary>
        /// <param name="month">The given month to show the data from</param>
        /// <param name="year">The given year to show the data from</param>
        static void CreateExcelSheetAndTable(string month, string year)
        {
            string FILE_PATH = $@"C:\temp\<stream_name>-CLs-for-{month}-{year}.xlsx";
            object missing = Type.Missing;

            // Excel commands:

            Excel.Application oXL = new Excel.Application
            {
                Visible = false
            };
            Excel.Workbook oWB = oXL.Workbooks.Add(missing);

            // create sheet
            Excel.Worksheet sheet = oWB.ActiveSheet as Excel.Worksheet;
            sheet.Name = $"{month} {year}";

            AddDataToSheet(sheet, month, year);

            // delete prev file if exists - to prevent overwrite request popup msg
            if (File.Exists(FILE_PATH))
            {
                File.Delete(FILE_PATH);
            }
            oWB.SaveAs(FILE_PATH, Excel.XlFileFormat.xlOpenXMLWorkbook,
                missing, missing, missing, missing,
                Excel.XlSaveAsAccessMode.xlNoChange,
                missing, missing, missing, missing, missing);
            oWB.Close(missing, missing, missing);
            oXL.UserControl = true;
            oXL.Quit();
            Console.WriteLine($"Results file created in {FILE_PATH}");

        }//end of CreateExcelSheetsAndTables

        /// <summary>
        /// This function adds the data form the given month and year to the Excel sheet
        /// </summary>
        /// <param name="sheet">The sheet object</param>
        /// <param name="monthStr">The given month</param>
        /// <param name="year">The given year</param>
        static void AddDataToSheet(Excel.Worksheet sheet, string monthStr, string year)
        {
            int month = GetMonthNumber(monthStr);
            // cmd P4V commands
            string cmdToRun = $"p4 changes -l -s submitted //...@{year}/{month}/1:00:01:00,{year}/{month}/31:23:59:59";
            const string GET_ALL_STREAM_NAME_CLIENTS_CMD = "p4 clients -S <stream_name>";
            const string STREAM_NAME_FILE_PREFIX = "<stream_name>/";
            var listOfAllClientsLines = new List<string>();
            var listOfAllClients = new List<string>();
            var listOfAllCLsLines = new List<string>();
            var listOfAllCLs = new List<string>();
            var listOfCurrClFiles = new List<string>();
            int numOfCLs = 0;

            // get data

            // **** get all desired Stream clients ****
            Console.WriteLine();
            Console.WriteLine($"Getting all {monthStr} {year} <stream_name> clients (workspaces)...");

            // Get the output of the cmd
            string result = GetCmdCommandOutput(GET_ALL_STREAM_NAME_CLIENTS_CMD);

            //listOfAllClientsLines = result.Split('\n').ToList();
            listOfAllClientsLines = result.Split(new string[] { "Client " }, StringSplitOptions.None).ToList();

            // create list of Clients = Workspaces (from list)
            foreach (var clientLine in listOfAllClientsLines)
            {
                string[] tokens = clientLine.Split(' ');
                listOfAllClients.Add(tokens[0]); // add cliet to list
            }

            // **** get all CLs info ****
            Console.WriteLine();
            Console.WriteLine($"Getting P4V CLs info...{Environment.NewLine}");

            // Get the output of the cmd
            result = GetCmdCommandOutput(cmdToRun);

            //listOfAllCLsLines = result.Split('\n').ToList();
            listOfAllCLsLines = result.Split(new string[] { "\nChange " }, StringSplitOptions.None).ToList();
            // remove first element (empty string) if needed
            if (string.IsNullOrEmpty(listOfAllCLsLines.ElementAt(0)))
            {
                listOfAllCLsLines.RemoveAt(0);
            }

            // set curr Excel sheet

            // add column names and settings (row 1)
            sheet.Cells[1, 1] = "CL#";

            sheet.Cells[1, 2] = "User name";

            sheet.Cells[1, 3] = "Date submitted";

            sheet.Cells[1, 4] = "Description";
            // set width to 60 and set the column to ‘wrap - text’
            sheet.Columns[4].ColumnWidth = 60;
            sheet.Columns[4].WrapText = true;

            sheet.Cells[1, 5] = "List of files";
            sheet.Columns[5].ColumnWidth = 120;
            sheet.Columns[5].WrapText = true;

            sheet.Cells[1, 6] = "Status";

            // add ------- seperators under each column name (row 2)
            sheet.Cells[2, 1] = "-----";
            sheet.Cells[2, 2] = "---------------";
            sheet.Cells[2, 3] = "-------------------";
            sheet.Cells[2, 4] = "-----------------";
            sheet.Cells[2, 5] = "---------------";
            sheet.Cells[2, 6] = "---------";

            int currRowNumber = 3;

            // reverse the list so it will be from the 1st day of the month
            listOfAllCLsLines.Reverse();

            // print curr Stream CLs table
            foreach (var clLine in listOfAllCLsLines)
            {
                string[] tokens = clLine.Split(' ');
                // make sure line is valid
                if (tokens.Length >= 5 && int.TryParse(tokens[0], out int n))
                {
                    string currClClientName = string.Empty;

                    // get string between '@' and '\r'
                    currClClientName = tokens[4].Trim().Split('@')[1].Split('\r')[0];

                    if (listOfAllClients.Contains(currClClientName)) // CL is in <stream_name>
                    {
                        sheet.Cells[currRowNumber, 1] = tokens[0]; // CL#
                        sheet.Cells[currRowNumber, 2] = tokens[4].Substring(0, tokens[4].IndexOf('@')); // User name
                        sheet.Cells[currRowNumber, 3] = tokens[2]; // Date submitted

                        // Description
                        string description = GetStringBetween(clLine, $"@{currClClientName}", "Bug #:")
                            .Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ').Replace(',', ' ').Trim();
                        if (string.IsNullOrEmpty(description))
                        {
                            description = GetStringBetween(clLine, $"@{currClClientName}", "lastreview=")
                            .Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ').Replace(',', ' ').Trim();
                        }
                        if (string.IsNullOrEmpty(description))
                        {
                            description = GetStringBetween(clLine, "\n", "Bug #:")
                            .Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ').Replace(',', ' ').Trim();
                        }
                        if (string.IsNullOrEmpty(description))
                        {
                            description = clLine.Substring(clLine.IndexOf('@') + 1).Replace(currClClientName, "")
                            .Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ').Replace(',', ' ').Trim();
                        }
                        sheet.Cells[currRowNumber, 4] = description; // Description

                        // List of files
                        string[] currClFiles = GetCmdCommandOutput($"p4 files @={tokens[0]}").Split('\n');
                        for (int i = 0; i < currClFiles.Length; i++)
                        {
                            currClFiles[i] = GetStringBetween(currClFiles[i], STREAM_NAME_FILE_PREFIX, " - "); // was: "#"
                        }
                        sheet.Cells[currRowNumber, 5] = currClFiles[0]; // List of files (display 1st file)

                        sheet.Cells[currRowNumber, 6] = string.Empty; // Status

                        // create a new row for each other file in the curr CL
                        for (int i = 1; i < currClFiles.Length; i++)
                        {
                            sheet.Cells[++currRowNumber, 5] = currClFiles[i]; // List of files (display curr file)
                        }

                        // FOR TESTING ONLY:
                        //string user = tokens[4].Substring(0, tokens[4].IndexOf('@'));
                        /*string clNum = tokens[0];
                        if (clNum.Equals("756692"))
                        {
                            Console.WriteLine(clLine);
                        }*/
                        numOfCLs++;
                        currRowNumber++;
                    }
                }
            }//end of foreach

            sheet.Columns.AutoFit(); // auto-fit all columns to their text
            sheet.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft; // align left all cells

            Console.WriteLine($"Number of CLs: {numOfCLs}");
            Console.WriteLine("==============================================");

            // check that the number of CLs is greater than 0
            if (numOfCLs == 0)
            {
                try
                {
                    throw new ArgumentNullException("There are NO CLs to display");
                }
                // make sure program won't close immediately, so user can read the Exception msg
                finally
                {
                    Console.WriteLine("Press any key to exit...");
                    Console.ReadKey();
                }
            }
        }// end of AddDataToSheet

        /// <summary>
        /// Converts a 3-chars month to its numeric value.
        /// Examples: Jan => 1 ; Feb => 2
        /// </summary>
        /// <param name="month">The given 3-chars month string</param>
        /// <returns>The given month's numeric value (1-12)</returns>
        public static int GetMonthNumber(string month)
        {
            if (month.Length != 3)
            {
                throw new ArgumentOutOfRangeException("Invalid month length (MUST be 3-chars long)");
            }
            switch (month.ToLower())
            {
                case "jan":
                    return 1;
                case "feb":
                    return 2;
                case "mar":
                    return 3;
                case "apr":
                    return 4;
                case "may":
                    return 5;
                case "jun":
                    return 6;
                case "jul":
                    return 7;
                case "aug":
                    return 8;
                case "sep":
                    return 9;
                case "oct":
                    return 10;
                case "nov":
                    return 11;
                case "dec":
                    return 12;
                default:
                    throw new ArgumentOutOfRangeException("Invalid month");
            }
        }
    }
}
