namespace ContactManger.Models
{
    using System.ComponentModel.DataAnnotations;
    public partial class Contact
    {
        [Required]
        public int Id { get; set; }
        [Required(ErrorMessage ="First Name Requried")]
        public string FirstName { get; set; }
        [Required(ErrorMessage = "Last Name Requried")]
        public string LastName { get; set; }
        [Required(ErrorMessage = "Phone No Requried")]
        public string PhoneNo { get; set; }
        [Required(ErrorMessage = "Email ID Requried")]
        public string EmailID { get; set; }
    }
}

//count 17