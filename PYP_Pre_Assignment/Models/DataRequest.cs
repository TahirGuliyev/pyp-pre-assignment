using System.ComponentModel.DataAnnotations;

namespace PYP_Pre_Assignment.Models
{
    public class DataRequest
    {
        [Required]
        public DateTime StartDate { get; set; }
        [Required]
        public DateTime EndDate { get; set; }
        [Required]
        [DataType(DataType.EmailAddress)]
        public string? AcceptorEmail { get; set; }
        [Required]
        public FilterEnum Filter { get; set; }
    }
}
