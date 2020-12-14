using System;
using System.Collections.Generic;

#nullable disable

namespace TestAPI.Models
{
    public partial class User
    {
        public User()
        {
            Orders = new HashSet<Order>();
        }

        public int Id { get; set; }
        public string Name { get; set; }
        public string Grader { get; set; }
        public string Phone { get; set; }
        public string Address { get; set; }
        public string Email { get; set; }

        public virtual ICollection<Order> Orders { get; set; }
    }
}
