using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;

class Program
{
    static void Main(string[] args)
    {
        ReadDoc();
    }

    private static void ReadDoc()
    {
        try
        {
            string document = @"C:\Users\aad4w\OneDrive\Desktop\Shawn\ReadExcel\FLOOR_LOCATION.xlsx";

            //Uses openxml sdk to open and read the given excel file
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(document, false))
            {
                //creates an object for the workbook part  
                WorkbookPart workbookPart = doc.WorkbookPart;
                Sheets thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();
                StringBuilder excelResult = new StringBuilder();

                //loop to go through the entire sheet 
                foreach (Sheet thesheet in thesheetcollection)
                {
                    excelResult.AppendLine(thesheet.Name);
                    excelResult.AppendLine("---------------------------");
                    //Gets the worksheet object by using the sheet id  
                    Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;

                    SheetData thesheetdata = (SheetData)theWorksheet.GetFirstChild<SheetData>();
                    foreach (Row thecurrentrow in thesheetdata)
                    {
                        foreach (Cell thecurrentcell in thecurrentrow)
                        {
                            //statement to take the integer value  
                            string currentcellvalue = string.Empty;
                            if (thecurrentcell.DataType != null)
                            {
                                if (thecurrentcell.DataType == CellValues.SharedString)
                                {
                                    int id;
                                    if (Int32.TryParse(thecurrentcell.InnerText, out id))
                                    {
                                        SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                        if (item.Text != null)
                                        {
                                            //code to take the string value  
                                            excelResult.Append(item.Text.Text + "\t\t");
                                        }
                                        else if (item.InnerText != null)
                                        {
                                            currentcellvalue = item.InnerText;
                                        }
                                        else if (item.InnerXml != null)
                                        {
                                            currentcellvalue = item.InnerXml;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                excelResult.Append(Convert.ToInt16(thecurrentcell.InnerText) + "\t\t");
                            }
                        }
                        excelResult.AppendLine();
                    }
                    excelResult.Append("");
                    Console.WriteLine(excelResult.ToString());
                    Console.ReadLine();
                }
            }
        }
        catch (Exception)
        {
            Console.WriteLine("Whoops there was an error :)");
        }
    }
}