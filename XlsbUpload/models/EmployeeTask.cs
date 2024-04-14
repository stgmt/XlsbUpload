using System.ComponentModel.DataAnnotations;

namespace XlsbUpload.models
{
    internal class EmployeeTask
    {
        [Required]
        public string IdTask { get; set; } // Идентификатор задачи

        [Required]
        public string TIN { get; set; } // Табельный номер

        [Required]
        public string FirstName { get; set; }

        [Required]
        public string LastName { get; set; } 


    }
}
