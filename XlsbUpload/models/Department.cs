using System.ComponentModel.DataAnnotations;

namespace XlsbUpload.models
{
    internal class Department
    {
        [Required]
        public string IdDepartment { get; set; }

        [Required]
        public string DepartmentName { get; set; }
    }
}
