using System;
using System.Collections.Generic;

#nullable disable

namespace TestAPI.Models
{
    public partial class Product
    {
        public Product()
        {
            OrderDetails = new HashSet<OrderDetail>();
        }

        public int Id { get; set; }
        public string Name { get; set; }
        public int? CategoryId { get; set; }
        public decimal? Price { get; set; }
        public string Desc { get; set; }
        public int? Stock { get; set; }
        public DateTime? AddTime { get; set; }

        public virtual Category IdNavigation { get; set; }
        public virtual ICollection<OrderDetail> OrderDetails { get; set; }
    }
}
