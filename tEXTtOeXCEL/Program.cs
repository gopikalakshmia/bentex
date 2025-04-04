using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        var license = new EPPlusLicense();
        license.SetNonCommercialPersonal("YourName");

        string filePath = "Book1.xlsx";

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[1]; // Get the first worksheet
            int colCount = worksheet.Dimension.Columns;     // Total columns

            string note = "sample.txt";
            var result = new List<string[]>(); // List of string arrays
            var group = new List<string>(); // Temporary group to hold the lines

            var lines = File.ReadLines(note);
            foreach (var line in lines)
            {
                if (line.StartsWith("16"))
                {
                    if (group.Count > 0)
                    {
                        // Convert the group to an array and add it to result
                        result.Add(group.ToArray());
                        group.Clear(); // Start a new group
                    }
                }
                group.Add(line); // Add the line to the current group
            }

            // After the loop, make sure to add the last group if it contains data
            if (group.Count > 0 && group[0] == "16")
            {
                result.Add(group.ToArray());
            }

            // Display the result and write to the Excel file
            int row = 3; // Start writing data from row 3 in Excel
            foreach (var array in result)
            {
                string s = string.Join(", ", array);
                string[] splitArray = s.Split(',');

                if (splitArray[0] == "16") // Ensure you're processing lines that start with "16"
                {
                    int column = 1; // Start from the first column for each new group
                    for (int j = 0; j < splitArray.Length; j++)
                    {
                       if (splitArray[j].Trim() != " " && splitArray[j].Trim() != "88") // Ignore blank entries & in the array
                        {
                            for (int col = 1; col <= colCount; col++)
                            {
                                // Match columns by name and assign corresponding values
                                if (worksheet.Cells[1, col].Text == "Code")
                                {
                                    worksheet.Cells[row, column].Value = splitArray[j];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; // Move to the next column
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "Transaction Code_02")
                                {
                                    worksheet.Cells[row, column].Value = splitArray[j];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; // Move to the next column
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "Amount_03")
                                {
                                    worksheet.Cells[row, column].Value = splitArray[j];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; // Move to the next column
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "16_04")
                                {
                                    worksheet.Cells[row, column].Value = splitArray[j];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; // Move to the next column
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "16_05")
                                {
                                    worksheet.Cells[row, column].Value = splitArray[j];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; // Move to the next column
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "Batch" && splitArray[j].StartsWith("Batch"))
                                {
                                    worksheet.Cells[row, column].Value = splitArray[j].Split('=')[1];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; // Move to the next column
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "DFI BANK" && splitArray[j].StartsWith("DFI BANK"))
                                {
                                      worksheet.Cells[row, column].Value = splitArray[j].Split("=")[1];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; // Move to the next column
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "DFI ACCT" && splitArray[j].StartsWith("DFI ACCT"))
                                {
                                     worksheet.Cells[row, column].Value = splitArray[j].Split("=")[1];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; // Move to the next column
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "IND ID NO" && splitArray[j].StartsWith("IND ID NO"))
                                {
                                      worksheet.Cells[row, column].Value = splitArray[j].Split("=")[1];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; // Move to the next column
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "IND NAME" && splitArray[j].StartsWith("IND NAME"))
                                {
                                     worksheet.Cells[row, column].Value = splitArray[j].Split("=")[1];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; // Move to the next column
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "TRACE NO" && splitArray[j].StartsWith("TRACE NO"))
                                {
                                     worksheet.Cells[row, column].Value = splitArray[j].Split("=")[1];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; // Move to the next column
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "BATCH NUMBER" && splitArray[j].StartsWith("BATCH NUMBER"))
                                {
                                     worksheet.Cells[row, column].Value = splitArray[j].Split("=")[1];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; // Move to the next column
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "SETT BANKREF" && splitArray[j].StartsWith("SETT BANKREF"))
                                {
                                      worksheet.Cells[row, column].Value = splitArray[j].Split("=")[1];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; // Move to the next column
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "SETT CUSTREF" && splitArray[j].StartsWith("SETT CUSTREF"))
                                {
                                     worksheet.Cells[row, column].Value = splitArray[j].Split("=")[1];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; // Move to the next column
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "SETT AMOUNT" && splitArray[j].StartsWith("SETT AMOUNT"))
                                {
                                     worksheet.Cells[row, column].Value = splitArray[j].Split("=")[1];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; // Move to the next column
                                    break;
                                }
                            }
                        }
                    }
                    row++; // Increment row after processing the group
                }
            }

            // Save the changes to the Excel file
            package.Save();
        }
    }
}
