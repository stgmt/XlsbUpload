using System;
using System.ComponentModel.DataAnnotations;

namespace XlsbUpload.models
{
    internal class Employee
    {
        [Required]
        public string TIN { get; set; } // Табельный номер
        [Required]
        public string LastName { get; set; }
        [Required]
        public string FirstName { get; set; }
        [Required]
        public string MiddleName { get; set; }
        [Required]
        public DateTime DateOfBirth { get; set; }
        [Required]
        public string DepartmentId { get; set; }

    }
}
