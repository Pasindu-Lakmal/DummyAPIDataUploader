namespace DummyAPIDataUploader.Services.ExcelManipulationServices
{
    public interface IExcelManipulationService
    {
        //Task<List<UploadLogDetail>> readMacro();
        string getCreatedVBCode(string[] cellIDs);
    }
}
