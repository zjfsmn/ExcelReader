using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace ConsoleApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            string directoryPath = @"/Users/jingfei_zhang/Desktop/award"; // need to modify if the dirctory path is different;
            var subrecipientTotals = new Dictionary<string, decimal>();

            if (Directory.Exists(directoryPath))
            {
                Dictionary<string, decimal> subrecipientDictionary = new Dictionary<string, decimal>();

                foreach (string filePath in Directory.GetFiles(directoryPath, "*.xlsx"))
                {
                    FileInfo fileInfo = new FileInfo(filePath);

                    using (ExcelPackage package = new ExcelPackage(fileInfo))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                        int rowCount = worksheet.Dimension.Rows;

                        //  Console.WriteLine("File Name: " + fileInfo.Name);
                        // Console.WriteLine("Sheet Name: " + worksheet);
                        // get the row indexes of cells G and H
                        int rowA = 0;
                        int rowG = 0;
                        int rowH = 0;
                        for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
                        {
                            string cellValue = worksheet.Cells[i, 1].Text;
                            if (cellValue == "A.")
                                rowA = i;
                            else if (cellValue == "G.")
                                rowG = i;
                            else if (cellValue == "H.")
                                rowH = i;
                        }

                        // Console.WriteLine(fileInfo.Name + " G. is  " + rowG);
                        // Console.WriteLine(fileInfo.Name + " H. is  " + rowH);
                        int rowSubaward = 1;
                        // loop through the rows between G and H
                        for (int i = rowG + 1; i < rowH; i++)
                        {
                            // get the combined value of the second and third columns
                            string subrecipientValue = worksheet.Cells[i, 2].Text.ToString().TrimEnd() + " " + worksheet.Cells[i, 3].Text.ToString().TrimStart();


                            // check if the cell contains "Subaward" and get the subrecipient name
                            if (worksheet.Cells[i, 2].Text.StartsWith("Subaward"))
                            {
                                rowSubaward = i;
                                // Console.WriteLine(fileInfo.Name + "subrecipientname is  " + subrecipientValue);
                                string subrecipient = subrecipientValue.Substring(10).ToString().TrimEnd();
                                //  Console.WriteLine(subrecipient);
                                // Console.WriteLine("rowSubaward is " + rowSubaward);


                                // get the subaward amount for the subrecipient
                                // Find the cell with the value "Total" in the same row
                                for (int col = 4; col <= worksheet.Dimension.End.Column; col++)
                                {
                                    for (int row = 1; row <= rowA; row++)
                                    {
                                        if (worksheet.Cells[row, col].Value?.ToString() == "Total")
                                        {
                                            // Console.WriteLine("column is " + col + " row number is " + row);         
                                            decimal subaward = decimal.Parse(worksheet.Cells[rowSubaward, col].Value?.ToString() ?? "0");
                                            //  Console.WriteLine(worksheet.Cells[rowSubaward, col].Value);
                                            if (!string.IsNullOrEmpty(subrecipient))
                                            {
                                                if (!subrecipientTotals.ContainsKey(subrecipient))
                                                {
                                                    subrecipientTotals[subrecipient] = 0;
                                                }
                                                subrecipientTotals[subrecipient] += subaward;
                                                // Console.WriteLine($"{fileInfo.Name }: {subrecipient} - {subaward}");
                                                Console.WriteLine($"{fileInfo.Name }: {subrecipient} ");

                                            }

                                            break;
                                        }

                                    }

                                }

                            }
                        }

                    }
                }
            }
            else
            {
                Console.WriteLine("Directory not found.");
            }

            Console.WriteLine("Subrecipient total subaward amount:");
            if(subrecipientTotals.Keys.Any()){
                foreach (KeyValuePair<string, decimal> subrecipientTotal in subrecipientTotals)
                {
                    Console.WriteLine($"{subrecipientTotal.Key} : {subrecipientTotal.Value}");
                }
            }else{
                Console.WriteLine("No subaward.");
            }
           
        }

        public static List<string> ReadSubrecipients( String filePath){
            FileInfo fileInfo = new FileInfo(filePath);
            List<string> subrecipients = new List<string>();
            using (ExcelPackage package = new ExcelPackage(fileInfo))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                        int rowCount = worksheet.Dimension.Rows;

                        //  Console.WriteLine("File Name: " + fileInfo.Name);
                        // Console.WriteLine("Sheet Name: " + worksheet);
                        // get the row indexes of cells G and H
                      
                        int rowG = 0;
                        int rowH = 0;
                        for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
                        {
                            string cellValue = worksheet.Cells[i, 1].Text;
                          if (cellValue == "G.")
                                rowG = i;
                            else if (cellValue == "H.")
                                rowH = i;
                        }

                        // Console.WriteLine(fileInfo.Name + " G. is  " + rowG);
                        // Console.WriteLine(fileInfo.Name + " H. is  " + rowH);
                
                        // loop through the rows between G and H
                        for (int i = rowG + 1; i < rowH; i++)
                        {
                            // get the combined value of the second and third columns
                            string subrecipientValue = worksheet.Cells[i, 2].Text.ToString().TrimEnd() + " " + worksheet.Cells[i, 3].Text.ToString().TrimStart();


                            // check if the cell contains "Subaward" and get the subrecipient name
                            if (worksheet.Cells[i, 2].Text.StartsWith("Subaward"))
                            {
                                
                                // Console.WriteLine(fileInfo.Name + "subrecipientname is  " + subrecipientValue);
                                string subrecipient = subrecipientValue.Substring(10).ToString().TrimEnd();
                                subrecipients.Add(subrecipient);
                                //   Console.WriteLine(subrecipient);
                               



                            }
                        }

                    }
                    return subrecipients;

        }



    }
}
