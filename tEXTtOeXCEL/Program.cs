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
            var worksheet = package.Workbook.Worksheets[1];
            int colCount = worksheet.Dimension.Columns;     

            string note = "sample.txt";
            var result = new List<string[]>(); 
            var group = new List<string>(); 

            var lines = File.ReadLines(note);
            foreach (var line in lines)
            {
                if (line.StartsWith("16"))
                {
                    if (group.Count > 0)
                    {
                        
                        result.Add(group.ToArray());
                        group.Clear();
                    }
                }
                group.Add(line); 
            }

        
            if (group.Count > 0 && group[0] == "16")
            {
                result.Add(group.ToArray());
            }

 
            int row = 3;
            foreach (var array in result)
            {
                string s = string.Join(", ", array);
                string[] splitArray = s.Split(',');

                if (splitArray[0] == "16") 
                {
                    int column = 1; 
                    for (int j = 0; j < splitArray.Length; j++)
                    {
                       if (splitArray[j].Trim() != " " && splitArray[j].Trim() != "88")
                        {
                            for (int col = 1; col <= colCount; col++)
                            {
 
                                if (worksheet.Cells[1, col].Text == "Code")
                                {
                                    worksheet.Cells[row, column].Value = splitArray[j];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; 
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "Transaction Code_02")
                                {
                                    worksheet.Cells[row, column].Value = splitArray[j];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; 
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "Amount_03")
                                {
                                    worksheet.Cells[row, column].Value = splitArray[j];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; 
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "16_04")
                                {
                                    worksheet.Cells[row, column].Value = splitArray[j];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; 
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "16_05")
                                {
                                    worksheet.Cells[row, column].Value = splitArray[j];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++;
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "Batch" && splitArray[j].StartsWith("Batch"))
                                {
                                    worksheet.Cells[row, column].Value = splitArray[j].Split('=')[1];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; 
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "DFI BANK" && splitArray[j].StartsWith("DFI BANK"))
                                {
                                      worksheet.Cells[row, column].Value = splitArray[j].Split("=")[1];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; 
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "DFI ACCT" && splitArray[j].StartsWith("DFI ACCT"))
                                {
                                     worksheet.Cells[row, column].Value = splitArray[j].Split("=")[1];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; 
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "IND ID NO" && splitArray[j].StartsWith("IND ID NO"))
                                {
                                      worksheet.Cells[row, column].Value = splitArray[j].Split("=")[1];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; 
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "IND NAME" && splitArray[j].StartsWith("IND NAME"))
                                {
                                     worksheet.Cells[row, column].Value = splitArray[j].Split("=")[1];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; 
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "TRACE NO" && splitArray[j].StartsWith("TRACE NO"))
                                {
                                     worksheet.Cells[row, column].Value = splitArray[j].Split("=")[1];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; 
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "BATCH NUMBER" && splitArray[j].StartsWith("BATCH NUMBER"))
                                {
                                     worksheet.Cells[row, column].Value = splitArray[j].Split("=")[1];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; 
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
                                    column++; 
                                    break;
                                }
                                if (worksheet.Cells[1, col].Text == "SETT AMOUNT" && splitArray[j].StartsWith("SETT AMOUNT"))
                                {
                                     worksheet.Cells[row, column].Value = splitArray[j].Split("=")[1];
                                    Console.WriteLine(row + " " + column + " " + splitArray[j]);
                                    column++; 
                                    break;
                                }
                            }
                        }
                    }
                    row++; 
                }
            }

           
            package.Save();
        }
    }
}
