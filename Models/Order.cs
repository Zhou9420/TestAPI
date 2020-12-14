using System;
using System.Collections.Generic;

#nullable disable

namespace TestAPI.Models
{
    public partial class Order
    {
        public Order()
        {
            OrderDetails = new HashSet<OrderDetail>();
        }

        public int Id { get; set; }
        public int? UserId { get; set; }
        public string Phone { get; set; }
        public string Address { get; set; }
        public string Amount { get; set; }

        public virtual User User { get; set; }
        public virtual ICollection<OrderDetail> OrderDetails { get; set; }
    }
}
