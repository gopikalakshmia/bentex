
using System.Text.RegularExpressions;
using OfficeOpenXml;

class Header
{
    static void Main1(string[] args)
    {

        var license = new EPPlusLicense();
        license.SetNonCommercialPersonal("YourName");


        string filePath = "sample.txt";
        var excelFilepath = "output.xlsx";
        string headerPattern = @"^(0[1-4])\b";
        Regex headerRegex = new Regex(headerPattern);

        try
        {

            var lines = File.ReadLines(filePath);//read text file

            foreach (var line in lines)
            {
                // Console.WriteLine(line);
            }

            using (var package = new ExcelPackage(new FileInfo(excelFilepath)))
            {
                //Header worksheet
                var worksheet = package.Workbook.Worksheets.Add("Header");
                int col = 1;
                foreach (var line in lines)
                {
                    var dataArray = line.Split(',');
                    //header

                    if (headerRegex.IsMatch(dataArray[0]))
                    {
                        int row = 1;

                        // Console.WriteLine("true");
                        for (int i = 1; i <= dataArray.Length - 1; i++)
                        {
                            Console.WriteLine(row + " " + col);
                            if (dataArray[i] != "" && (dataArray[i] != "2/" && dataArray[i] != "/")) //considering / or 2/ end of the line
                            {
                                Console.WriteLine(dataArray[0] + "0" + i + " " + dataArray[i]);
                                worksheet.Cells[row, col].Value = dataArray[0] + "0" + i;
                                worksheet.Cells[row + 1, col].Value = dataArray[i];
                                col = col + 1;
                            }


                        }
                    }

                }
                package.Save();

            }
        }
        catch (Exception ex)
        {
            // Handle any exceptions
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

