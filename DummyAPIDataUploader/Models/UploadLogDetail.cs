namespace DummyAPIDataUploader.Models
{
    public class UploadLogDetail
    {
        public int Id { get; set; }

        public string UploadedBy { get; set; } = string.Empty;


        public DateTime UploadedDate { get; set; }
        public string FileName { get; set; } = string.Empty;

        public string FileType { get; set; } = string.Empty;

        public string StaffData { get; set; } = string.Empty;
    }
}
