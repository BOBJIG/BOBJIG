using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NhibernateConsoleapplication
{
    class Emp
    {
        public virtual int Emp_id { get; set; }
        public virtual string Empname { get; set; }
        public virtual Fulladress FullAdress  { get; set; }
        public virtual ISet<Orders> OrderInformation { get; set; }
    }
}
