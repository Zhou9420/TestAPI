using System;
using System.Collections.Generic;

#nullable disable

namespace TestAPI.Models
{
    public partial class Admin
    {
        public int Id { get; set; }
        public int? LoginId { get; set; }
        public string LoginPwd { get; set; }
        public string Name { get; set; }
    }
}
