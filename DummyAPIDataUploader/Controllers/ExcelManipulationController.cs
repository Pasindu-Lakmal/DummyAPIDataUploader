using Microsoft.AspNetCore.Mvc;
using Syncfusion.XlsIO;
using System.IO;
using Syncfusion.Drawing;
using Syncfusion.Office;
using Microsoft.OpenApi.Any;
using static System.Net.Mime.MediaTypeNames;
using static System.Net.WebRequestMethods;
using System.Text;
using System.Net;
using System.Collections.Generic;




// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace DummyAPIDataUploader.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelManipulationController : ControllerBase
    {
        private readonly IWebHostEnvironment _environment;
        private readonly string _uploadFolder;

        public ExcelManipulationController(IWebHostEnvironment environment)
        {
            _environment = environment;
            _uploadFolder = Path.Combine(_environment.ContentRootPath, "Uploads");
            if (!Directory.Exists(_uploadFolder))
            {
                Directory.CreateDirectory(_uploadFolder);
            }
        }

        //[HttpPost("upload")]
        //public async Task<IActionResult> UploadFile(IFormFile file)
        //{
        //    if (file == null || file.Length == 0)
        //    {
        //        return BadRequest("No file uploaded.");
        //    }

        //    var filePath = Path.Combine(_uploadFolder, file.FileName);

        //    using (var stream = new FileStream(filePath, FileMode.Create))
        //    {
        //        await file.CopyToAsync(stream);
        //    }

        //    return Ok(new { filePath });
        //}


        [HttpPost("upload")]
        public async Task<IActionResult> UploadFile(IFormFile newfile)
        {
            try
            {
                if (newfile == null || newfile.Length <= 0)
                {
                    return BadRequest("File not selected or empty.");
                }

                // Read the uploaded Excel file
                using (var stream = new MemoryStream())
                {
                    await newfile.CopyToAsync(stream);
                    stream.Position = 0;

                    // Initialize ExcelEngine
                    using (var excelEngine = new ExcelEngine())
                    {
                        // Instantiate Excel application object
                        IApplication application = excelEngine.Excel;

                        // Open the workbook
                        IWorkbook workbook = application.Workbooks.Open(stream);

                        // Assuming there's only one worksheet in the workbook
                        IWorksheet worksheet = workbook.Worksheets[0];

                        // Read the content of the Excel file
                        // You can perform any further processing here based on your requirements

                        // Example: Read content from a cell
                        //var cellValue = worksheet.Range["C3"].Value;

                        // Example: Read content from a range of cells
                        var rangeValues = worksheet.Range["A1:D10"];
                    
                        // Close the workbook
                        workbook.Close();
                        return Ok(rangeValues);
                    }
                }

                // Return success message or data as needed
                //return Ok("File uploaded and processed successfully.");
            }
            catch (Exception ex)
            {
                // Log the exception or handle it accordingly
                return StatusCode(StatusCodes.Status500InternalServerError, $"An error occurred: {ex.Message}");
            }
        }

        [HttpGet("read/{fileName}")]
        public IActionResult ReadFile(string fileName)
        {
            var filePath = Path.Combine(_uploadFolder, fileName);

            if (!System.IO.File.Exists(filePath))
            {
                return NotFound();
            }

            var fileContent = System.IO.File.ReadAllText(filePath);
            return Ok(new { content = fileContent });
        }

        //[HttpPost("ReadMultiSelect")]
        //public async Task<string> readMultiSelect([FromBody] IFormFile file)
        //{

        //}



        [HttpPost("CreateMultiselect")]
        public IActionResult CreateMultiSelect([FromBody] MultiselectDropDownDetails[] dropdownDetails)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                string cellString = "";

                foreach (var item in dropdownDetails)
                {
                    var upperCaseCellId = item.Cell.ToUpper();
                    cellString += upperCaseCellId+":"+upperCaseCellId+", ";
                   
                }
                cellString = cellString.Remove(cellString.Length - 2);

                string code1 = "Private Sub Worksheet_Change(ByVal Target As Range)\n\t" +
                                "Dim rngDropdown As Range\n\t" +
                                "Dim oldValue As String\n\t" +
                                "Dim newValue As String\n\t" +
                                "Dim DelimiterType As String\n\t" +
                                "Dim DelimiterCount As Integer\n\t" +
                                "Dim TargetType As Integer\n\t" +
                                "Dim i As Integer\n\t" +
                                "Dim arr() As String\n\t" +

                                "If Target.Count > 1 Then Exit Sub\n\t" +
                                "If Not Intersect(Target, Me.Range(\"";

                                            string code2 = "\")) Is Nothing Then\n\t\t" +
                                    "On Error Resume Next\n\t\t" +

                                    "Set rngDropdown = Me.Cells.SpecialCells(xlCellTypeAllValidation)\n\t\t" +
                                    "On Error GoTo exitError\n\t\t" +

                                    "If rngDropdown Is Nothing Then GoTo exitError\n\t\t" +

                                    "TargetType = 0\n\t\t" +
                                    "TargetType = Target.Validation.Type\n\t" +
                                    "If TargetType = 3 Then  ' is validation type is \"list\"\n\t" +
                                        "Application.ScreenUpdating = False\n\t" +
                                        "Application.EnableEvents = False\n\t" +
                                        "newValue = Target.Value\n\t" +
                                        "Application.Undo\n\t" +
                                        "oldValue = Target.Value\n\t" +
                                        "Target.Value = newValue\n\t" +
                                        "If oldValue <> \"\" Then\n\t" +
                                            "If newValue <> \"\" Then\n\t" +
                                                "If oldValue = newValue Then ' leave the value if there is only one in the list\n\t" +
                                                    "Target.Value = oldValue\n\t" +
                                                "ElseIf InStr(1, oldValue, newValue & \",\") > 0 Then\n\t" +
                                                    "Target.Value = Replace(oldValue, newValue & \",\", \"\")\n\t" +
                                                "ElseIf InStr(1, oldValue, \",\" & newValue) > 0 Then\n\t" +
                                                    "Target.Value = Replace(oldValue, \",\" & newValue, \"\")\n\t" +
                                                "Else\n\t" +
                                                    "Target.Value = oldValue & \",\" & newValue\n\t" +
                                                "End If\n\t" +
                                            "End If\n\t" +
                                        "End If\n\t" +
                                        "Application.EnableEvents = True\n\t" +
                                        "Application.ScreenUpdating = True\n\t" +
                                    "End If\n\t" +
                                "End If\n" +

                            "exitError:\n\t" +
                                "Application.EnableEvents = True\n" +
                            "End Sub";
                string macroCode = code1 + cellString + code2;
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Data Validation for List

                foreach (var item in dropdownDetails)
                {
                    var upperCaseCellId = item.Cell.ToUpper();
                    worksheet.Range[upperCaseCellId + "1"].Text = item.ColumnName;
                    worksheet.Range[upperCaseCellId + "1"].AutofitColumns();
                    
                    
                    for (int i = 2; i <= 5000; i++)
                    {
                        worksheet.Range[upperCaseCellId + i].AutofitColumns();
                        IDataValidation listValidations = worksheet.Range[upperCaseCellId + i].DataValidation;
                        listValidations.ListOfValues = item.DropDownValues;
                        listValidations.ErrorBoxText = "Choose the value from the list";
                        listValidations.ErrorBoxTitle = "ERROR";
                      
                    }
                }



                //IDataValidation listValidation =  worksheet.Range["C3"].DataValidation;
                //worksheet.Range["C1"].Text = "Data Validation List in C3";
                //worksheet.Range["C1"].AutofitColumns();
                //listValidation.ListOfValues = new string[] { "ListItem1", "ListItem2", "ListItem3" };

                //IDataValidation listValidation1 = worksheet.Range["B3"].DataValidation;
                //worksheet.Range["B1"].Text = "Data Validation List in B3";
                //worksheet.Range["B1"].AutofitColumns();
                //listValidation1.ListOfValues = new string[] { "Item1", "Item2", "Item3" };

                //Shows the error message
                //listValidation.ErrorBoxText = "Choose the value from the list";
                //listValidation.ErrorBoxTitle = "ERROR";
                //listValidation.PromptBoxText = "Data validation for list";
                //listValidation.IsPromptBoxVisible = true;
                //listValidation.ShowPromptBox = true;

                //Creating Vba project
                IVbaProject project = workbook.VbaProject;

                //Accessing vba modules collection
                IVbaModules vbaModules = project.Modules;

                // Accessing sheet module
                IVbaModule vbaModule = vbaModules[worksheet.CodeName];

                //Adding vba code to the module
                vbaModule.Code = macroCode;


                MemoryStream stream = new MemoryStream();
                workbook.SaveAs(stream);

                //Set the position as '0'.
                stream.Position = 0;

                //Download the Excel file in the browser
                FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");
                fileStreamResult.FileDownloadName = "Output.xlsm";
                return fileStreamResult;

            }


        }



    }
}






//[HttpPost("CreateMultiselectWithCellId")]
//public IActionResult CreateMultiSelectWithCellId([FromBody] MultiselectDropDownDetails[] dropDownDetails)
//{
//    using (ExcelEngine excelEngine = new ExcelEngine())
//    {
//        IApplication application = excelEngine.Excel;
//        application.DefaultVersion = ExcelVersion.Xlsx;
//        IWorkbook workbook = application.Workbooks.Create(1);
//        IWorksheet worksheet = workbook.Worksheets[0];

//        //Data Validation for List
//        IDataValidation listValidation = worksheet.Range["C3"].DataValidation;
//        worksheet.Range["C1"].Text = "Data Validation List in C3";
//        worksheet.Range["C1"].AutofitColumns();
//        listValidation.ListOfValues = new string[] { "ListItem1", "ListItem2", "ListItem3" };

//        //Shows the error message
//        listValidation.ErrorBoxText = "Choose the value from the list";
//        listValidation.ErrorBoxTitle = "ERROR";
//        listValidation.PromptBoxText = "Data validation for list";
//        listValidation.IsPromptBoxVisible = true;
//        listValidation.ShowPromptBox = true;

//        //Creating Vba project
//        IVbaProject project = workbook.VbaProject;

//        //Accessing vba modules collection
//        IVbaModules vbaModules = project.Modules;

//        // Accessing sheet module
//        IVbaModule vbaModule = vbaModules[worksheet.CodeName];




//        MemoryStream stream = new MemoryStream();
//        workbook.SaveAs(stream);

//        //Set the position as '0'.
//        stream.Position = 0;

//        //Download the Excel file in the browser
//        FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");
//        fileStreamResult.FileDownloadName = "Output.xlsm";
//        return fileStreamResult;

//    }


//}


//[HttpPost("CreateMultiValue")]
//public IActionResult CreateMultiValue([FromBody] string macrocode)
//{

//    using (ExcelEngine excelEngine = new ExcelEngine())
//    {
//        IApplication application = excelEngine.Excel;
//        application.DefaultVersion = ExcelVersion.Xlsx;
//        IWorkbook workbook = application.Workbooks.Create(1);
//        IWorksheet sheet = workbook.Worksheets[0];

//        IVbaProject project = workbook.VbaProject;

//        IVbaModules vbaModules = project.Modules;

//        IVbaModule vbaModule = vbaModules[sheet.CodeName];

//        vbaModule.Code = macrocode;

//        MemoryStream stream = new MemoryStream();
//        workbook.SaveAs(stream);

//        //Set the position as '0'.
//        stream.Position = 0;

//        //Download the Excel file in the browser
//        FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");
//        fileStreamResult.FileDownloadName = "Output.xlsm";
//        return fileStreamResult;
//    }




//}




// GET: api/<ExcelManipulationController>
//[HttpGet("CreateBasicExcel")]
//public IActionResult Get()
//{
//    //Adding Document to the workbook
//    //Create an instance of ExcelEngine
//    using (ExcelEngine excelEngine = new ExcelEngine())
//    {
//        IApplication application = excelEngine.Excel;
//        application.DefaultVersion = ExcelVersion.Xlsx;

//        //Create a workbook
//        IWorkbook workbook = application.Workbooks.Create(1);
//        IWorksheet worksheet = workbook.Worksheets[0];

//        //Adding a picture

//        //Disable gridlines in the worksheet
//        worksheet.IsGridLinesVisible = false;

//        //Enter values to the cells from A3 to A5
//        worksheet.Range["A3"].Text = "46036 Michigan Ave";
//        worksheet.Range["A4"].Text = "Canton, USA";
//        worksheet.Range["A5"].Text = "Phone: +1 231-231-2310";

//        //Make the text bold
//        worksheet.Range["A3:A5"].CellStyle.Font.Bold = true;

//        //Merge cells
//        worksheet.Range["D1:E1"].Merge();

//        //Enter text to the cell D1 and apply formatting.
//        worksheet.Range["D1"].Text = "INVOICE";
//        worksheet.Range["D1"].CellStyle.Font.Bold = true;
//        worksheet.Range["D1"].CellStyle.Font.RGBColor = Color.FromArgb(42, 118, 189);
//        worksheet.Range["D1"].CellStyle.Font.Size = 35;

//        //Apply alignment in the cell D1
//        worksheet.Range["D1"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight;
//        worksheet.Range["D1"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop;

//        //Enter values to the cells from D5 to E8
//        worksheet.Range["D5"].Text = "INVOICE#";
//        worksheet.Range["E5"].Text = "DATE";
//        worksheet.Range["D6"].Number = 1028;
//        worksheet.Range["E6"].Value = "12/31/2018";
//        worksheet.Range["D7"].Text = "CUSTOMER ID";
//        worksheet.Range["E7"].Text = "TERMS";
//        worksheet.Range["D8"].Number = 564;
//        worksheet.Range["E8"].Text = "Due Upon Receipt";

//        //Apply RGB backcolor to the cells from D5 to E8
//        worksheet.Range["D5:E5"].CellStyle.Color = Color.FromArgb(42, 118, 189);
//        worksheet.Range["D7:E7"].CellStyle.Color = Color.FromArgb(42, 118, 189);

//        //Apply known colors to the text in cells D5 to E8
//        worksheet.Range["D5:E5"].CellStyle.Font.Color = ExcelKnownColors.White;
//        worksheet.Range["D7:E7"].CellStyle.Font.Color = ExcelKnownColors.White;

//        //Make the text as bold from D5 to E8
//        worksheet.Range["D5:E8"].CellStyle.Font.Bold = true;

//        //Apply alignment to the cells from D5 to E8
//        worksheet.Range["D5:E8"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
//        worksheet.Range["D5:E5"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
//        worksheet.Range["D7:E7"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
//        worksheet.Range["D6:E6"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop;

//        //Enter value and applying formatting in the cell A7
//        worksheet.Range["A7"].Text = "  BILL TO";
//        worksheet.Range["A7"].CellStyle.Color = Color.FromArgb(42, 118, 189);
//        worksheet.Range["A7"].CellStyle.Font.Bold = true;
//        worksheet.Range["A7"].CellStyle.Font.Color = ExcelKnownColors.White;

//        //Apply alignment
//        worksheet.Range["A7"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
//        worksheet.Range["A7"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

//        //Enter values in the cells A8 to A12
//        worksheet.Range["A8"].Text = "Steyn";
//        worksheet.Range["A9"].Text = "Great Lakes Food Market";
//        worksheet.Range["A10"].Text = "20 Whitehall Rd";
//        worksheet.Range["A11"].Text = "North Muskegon,USA";
//        worksheet.Range["A12"].Text = "+1 231-654-0000";

//        //Create a Hyperlink for e-mail in the cell A13
//        IHyperLink hyperlink = worksheet.HyperLinks.Add(worksheet.Range["A13"]);
//        hyperlink.Type = ExcelHyperLinkType.Url;
//        hyperlink.Address = "Steyn@greatlakes.com";
//        hyperlink.ScreenTip = "Send Mail";

//        //Merge column A and B from row 15 to 22
//        worksheet.Range["A15:B15"].Merge();
//        worksheet.Range["A16:B16"].Merge();
//        worksheet.Range["A17:B17"].Merge();
//        worksheet.Range["A18:B18"].Merge();
//        worksheet.Range["A19:B19"].Merge();
//        worksheet.Range["A20:B20"].Merge();
//        worksheet.Range["A21:B21"].Merge();
//        worksheet.Range["A22:B22"].Merge();

//        //Enter details of products and prices
//        worksheet.Range["A15"].Text = "  DESCRIPTION";
//        worksheet.Range["C15"].Text = "QTY";
//        worksheet.Range["D15"].Text = "UNIT PRICE";
//        worksheet.Range["E15"].Text = "AMOUNT";
//        worksheet.Range["A16"].Text = "Cabrales Cheese";
//        worksheet.Range["A17"].Text = "Chocos";
//        worksheet.Range["A18"].Text = "Pasta";
//        worksheet.Range["A19"].Text = "Cereals";
//        worksheet.Range["A20"].Text = "Ice Cream";
//        worksheet.Range["C16"].Number = 3;
//        worksheet.Range["C17"].Number = 2;
//        worksheet.Range["C18"].Number = 1;
//        worksheet.Range["C19"].Number = 4;
//        worksheet.Range["C20"].Number = 3;
//        worksheet.Range["D16"].Number = 21;
//        worksheet.Range["D17"].Number = 54;
//        worksheet.Range["D18"].Number = 10;
//        worksheet.Range["D19"].Number = 20;
//        worksheet.Range["D20"].Number = 30;
//        worksheet.Range["D23"].Text = "Total";

//        //Apply number format
//        worksheet.Range["D16:E22"].NumberFormat = "$.00";
//        worksheet.Range["E23"].NumberFormat = "$.00";

//        //Apply incremental formula for column Amount by multiplying Qty and UnitPrice
//        application.EnableIncrementalFormula = true;
//        worksheet.Range["E16:E20"].Formula = "=C16*D16";

//        //Formula for Sum the total
//        worksheet.Range["E23"].Formula = "=SUM(E16:E22)";

//        //Apply borders
//        worksheet.Range["A16:E22"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
//        worksheet.Range["A16:E22"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
//        worksheet.Range["A16:E22"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].Color = ExcelKnownColors.Grey_25_percent;
//        worksheet.Range["A16:E22"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].Color = ExcelKnownColors.Grey_25_percent;
//        worksheet.Range["A23:E23"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
//        worksheet.Range["A23:E23"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
//        worksheet.Range["A23:E23"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].Color = ExcelKnownColors.Black;
//        worksheet.Range["A23:E23"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].Color = ExcelKnownColors.Black;

//        //Apply font setting for cells with product details
//        worksheet.Range["A3:E23"].CellStyle.Font.FontName = "Arial";
//        worksheet.Range["A3:E23"].CellStyle.Font.Size = 10;
//        worksheet.Range["A15:E15"].CellStyle.Font.Color = ExcelKnownColors.White;
//        worksheet.Range["A15:E15"].CellStyle.Font.Bold = true;
//        worksheet.Range["D23:E23"].CellStyle.Font.Bold = true;

//        //Apply cell color
//        worksheet.Range["A15:E15"].CellStyle.Color = Color.FromArgb(42, 118, 189);

//        //Apply alignment to cells with product details
//        worksheet.Range["A15"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
//        worksheet.Range["C15:C22"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
//        worksheet.Range["D15:E15"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;

//        //Apply row height and column width to look good
//        worksheet.Range["A1"].ColumnWidth = 36;
//        worksheet.Range["B1"].ColumnWidth = 11;
//        worksheet.Range["C1"].ColumnWidth = 8;
//        worksheet.Range["D1:E1"].ColumnWidth = 18;
//        worksheet.Range["A1"].RowHeight = 47;
//        worksheet.Range["A2"].RowHeight = 15;
//        worksheet.Range["A3:A4"].RowHeight = 15;
//        worksheet.Range["A5"].RowHeight = 18;
//        worksheet.Range["A6"].RowHeight = 29;
//        worksheet.Range["A7"].RowHeight = 18;
//        worksheet.Range["A8"].RowHeight = 15;
//        worksheet.Range["A9:A14"].RowHeight = 15;
//        worksheet.Range["A15:A23"].RowHeight = 18;

//        //Saving the Excel to the MemoryStream 
//        MemoryStream stream = new MemoryStream();
//        workbook.SaveAs(stream);

//        //Set the position as '0'.
//        stream.Position = 0;

//        //Download the Excel file in the browser
//        FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");
//        fileStreamResult.FileDownloadName = "Output.xlsx";
//        return fileStreamResult;
//    }
//}