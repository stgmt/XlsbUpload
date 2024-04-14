using System.ComponentModel.DataAnnotations;

namespace XlsbUpload.models
{
    internal class EmployeeTask
    {
        [Required]
        public string IdTask { get; set; } // Идентификатор задачи

        [Required]
        public string TIN { get; set; } // Табельный номер

    }
}
