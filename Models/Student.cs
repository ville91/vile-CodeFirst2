using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace Vile_CodeFirst_2.Models
{
    public class Student
    {
        [Required]
        public int ID { get; set; }

        [Required]
        [MaxLength(80)]
        public string Name { get; set; }

        [Required]
        public int birthdate { get; set; }

    }
}