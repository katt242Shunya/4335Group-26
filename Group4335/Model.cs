using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Group4335
{
    public class Order
    {
        public int Id { get; set; }
        public string OrderCode { get; set; }    
        public DateTime CreationDate { get; set; } 
        public string ClientCode { get; set; }     
        public string Services { get; set; }       
        public string Status { get; set; }        
    }
}
