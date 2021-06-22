using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitOfWork
{
    interface IUnitOfWorkV1
    {
       

        void BeforeCommit();
        void BeforeRollBack();
        
        void AfterCommit();
        void AfterRollBack();
        
    }
}
