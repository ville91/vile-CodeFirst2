using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Vile_CodeFirst_2.DBContext
{
    public class StudentManagementDBContext : System.Data.Entity.DbContext
    {
        public StudentManagementDBContext() : base("StudentManagementDB")
        {

        }
        public System.Data.Entity.DbSet<Vile_CodeFirst_2.Models.Student> Students { get; set; }

    }

}