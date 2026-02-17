using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Group4333
{
    public class Context : System.Data.Entity.DbContext
    {
        public Context() : base("name=Connection") { }
        public System.Data.Entity.DbSet<Employees> Employees { get; set; }
    }
}
