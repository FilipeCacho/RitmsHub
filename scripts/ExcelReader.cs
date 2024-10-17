using System;
using System.IO;
using System.Reflection;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace RitmsHub.Scripts
{
    public class ExcelReader
    {
        private static string filePath;
        private static ExcelPackage package;

        static ExcelReader()
        {
            InitializeExcelPackage();
        }

        public static int CurrentEnvironment { get; set; } = 3; // Default to PRD (3)

        private static void InitializeExcelPackage()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            if (!ExcelTemplateManager.CheckAndCreateExcelFile())
            {
                throw new FileNotFoundException("Failed to initialize Excel file.");
            }

            filePath = ExcelTemplateManager.GetExcelFilePath();
            package = new ExcelPackage(new FileInfo(filePath));
        }

        private static string FindExcelFileDirectory()
        {
            string currentDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            while (currentDirectory != null)
            {
                if (Path.GetFileName(currentDirectory).Equals("RitmsHub", StringComparison.OrdinalIgnoreCase))
                {
                    return currentDirectory;
                }
                currentDirectory = Directory.GetParent(currentDirectory)?.FullName;
            }
            throw new DirectoryNotFoundException("Could not find the RitmsHub directory containing the Excel file.");
        }

        public static (string url, string username, string password, string appid, string redirecturi, string authType, string loginPrompt, bool requireNewInstance) ReadLoginValues()
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Login"];
            if (worksheet == null)
            {
                throw new ArgumentException("The specified tab 'Login' in the worksheet does not exist");
            }

            // Adjust rowNumber to account for header row and correct environment mapping
            int rowNumber = CurrentEnvironment + 1;

            return (
                url: worksheet.Cells[$"A{rowNumber}"].Text,
                username: worksheet.Cells[$"B{rowNumber}"].Text,
                password: worksheet.Cells[$"C{rowNumber}"].Text,
                appid: worksheet.Cells[$"D{rowNumber}"].Text,
                redirecturi: worksheet.Cells[$"E{rowNumber}"].Text,
                authType: worksheet.Cells[$"F{rowNumber}"].Text,
                loginPrompt: worksheet.Cells[$"G{rowNumber}"].Text,
                requireNewInstance: bool.Parse(worksheet.Cells[$"H{rowNumber}"].Text)
            );
        }

        //iterate rows in the excel "Create Teams" worksheet
        public static List<TeamRow> ReadBuInfoExcel()
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Create Teams"];
            if (worksheet == null)
            {
                throw new ArgumentException("The specified tab 'Create Teams' in the worksheet does not exist");
            }

            List<TeamRow> rows = new List<TeamRow>();
            int rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++)
            {
                if (!IsRowEmpty(worksheet, row))
                {
                    string columnA = worksheet.Cells[row, 1].Text.Trim();
                    string columnB = ProcessSingleWordColumn(worksheet.Cells[row, 2].Text, "B", row);
                    string columnC = ProcessSingleWordColumn(worksheet.Cells[row, 3].Text, "C", row);
                    string columnD = ProcessSingleWordColumn(worksheet.Cells[row, 4].Text, "D", row);
                    string columnE = worksheet.Cells[row, 5].Text.Trim();

                    TeamRow teamRow = new TeamRow
                    {
                        ColumnA = columnA,
                        ColumnB = columnB,
                        ColumnC = columnC,
                        ColumnD = columnD,
                        ColumnE = columnE
                    };
                    rows.Add(teamRow);
                }
            }
            return rows;
        }

        private static string ProcessSingleWordColumn(string cellValue, string columnName, int rowNumber)
        {
            string processed = cellValue.Trim();

            if (string.IsNullOrWhiteSpace(processed))
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"Warning: Empty value in column {columnName} at row {rowNumber}. Using default value.");
                Console.ResetColor();
                return $"Default{columnName}";
            }

            if (processed.Contains(" "))
            {
                string[] words = processed.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine($"Warning: Multiple words found in column {columnName} at row {rowNumber}. Using only the first word.");
                Console.ResetColor();
                return words[0];
            }

            return processed;
        }

        private static bool IsRowEmpty(ExcelWorksheet worksheet, int row)
        {
            return string.IsNullOrWhiteSpace(worksheet.Cells[row, 1].Text) &&
                   string.IsNullOrWhiteSpace(worksheet.Cells[row, 2].Text) &&
                   string.IsNullOrWhiteSpace(worksheet.Cells[row, 3].Text) &&
                   string.IsNullOrWhiteSpace(worksheet.Cells[row, 4].Text) &&
                   string.IsNullOrWhiteSpace(worksheet.Cells[row, 5].Text);
        }

        public static List<TeamRow> ValidateTeamsToCreate()
        {
            List<TeamRow> validRows = new List<TeamRow>();
            List<TeamRow> createTeamData = ReadBuInfoExcel();

            // Iterate rows in the excel "Create Teams" worksheet
            for (int i = 0; i < createTeamData.Count; i++)
            {
                var row = createTeamData[i];

                //force letters to upper case to ensure correct team names
                row.ColumnA = row.ColumnA.ToUpper();
                row.ColumnB = row.ColumnB.ToUpper();
                row.ColumnC = row.ColumnC.ToUpper();

                bool isValid = true;
                string errorMessage = $"Row {i + 2}: ";

                string[] parts = row.ColumnA.Split('-');
                string countryCode = parts[1];

                // Validate EU teams
                if (CodesAndRoles.CountryCodeEU.Contains(countryCode))
                {
                    // Check if all cells have text
                    if (string.IsNullOrWhiteSpace(row.ColumnA) || string.IsNullOrWhiteSpace(row.ColumnB) ||
                        string.IsNullOrWhiteSpace(row.ColumnC) || string.IsNullOrWhiteSpace(row.ColumnD) || string.IsNullOrWhiteSpace(row.ColumnE))
                    {
                        errorMessage += "All cells from A-E of this row must contain text. ";
                        isValid = false;
                    }

                    // Validate Column A
                    if (!Regex.IsMatch(row.ColumnA, @"^\d-[A-Z]{2}-[A-Z0-9]{3}-\d{2}$"))
                    {
                        errorMessage += "Column A must be in the format '0-XX-XXX-00' where X is a letter and 0 is any digit. ";
                        isValid = false;
                    }

                    // Validate Column B
                    if (row.ColumnB.Length != 8)
                    {
                        errorMessage += "Column B must contain exactly 8 characters. ";
                        isValid = false;
                    }

                    // Validate Column C
                    if (!Regex.IsMatch(row.ColumnC, @"^ZP[A-Za-z0-9]$"))
                    {
                        errorMessage += "Column C must start with 'ZP' followed by a single character. ";
                        isValid = false;
                    }

                    // Validate Column D
                    if (row.ColumnD.Length != 4)
                    {
                        errorMessage += "Column D must contain exactly 4 characters. ";
                        isValid = false;
                    }

                    if (isValid)
                    {
                        validRows.Add(row);
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine(errorMessage);
                        Console.ResetColor();

                    }
                }
                // Placeholder logic for NA teams
                else if (CodesAndRoles.CountryCodeNA.Contains(countryCode))
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"Row {i + 2}: Found a BU with a US country code. No action taken.");
                    Console.ResetColor();
                    // Add NA-specific validation or processing logic here if needed in the future
                }
                // Handle invalid country codes, placeholder logic
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Row {i + 2}: Found a BU with an invalid country code. No action taken.");
                    Console.ResetColor();
                }
            }

            return validRows;
        }

        public static List<AssignTeamData> ReadAssignTeamsData()
        {
            InitializeExcelPackage();
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Assign Teams"];
            if (worksheet == null)
            {
                throw new ArgumentException("The specified tab 'Assign Teams' in the worksheet does not exist");
            }

            List<AssignTeamData> assignTeamDataList = new List<AssignTeamData>();
            int rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++) // Start from row 2 to skip headers
            {
                // Get the raw text from each cell
                string rawUsername = worksheet.Cells[row, 1].Text;
                string rawTeamName = worksheet.Cells[row, 2].Text;

                // Process username (column 1): remove all spaces
                string username = rawUsername.Replace(" ", "");

                // Process team name (column 2): trim start and end, replace multiple spaces with single space
                string teamName = string.Join(" ", rawTeamName.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));

                if (!string.IsNullOrWhiteSpace(username) && !string.IsNullOrWhiteSpace(teamName))
                {
                    assignTeamDataList.Add(new AssignTeamData
                    {
                        Username = username,
                        TeamName = teamName
                    });
                }
            }

            return assignTeamDataList;
        }
    }

  



    public class TeamRow
        {
            public string ColumnA { get; set; }
            public string ColumnB { get; set; }
            public string ColumnC { get; set; }
            public string ColumnD { get; set; }
            public string ColumnE { get; set; }
        }
    }
