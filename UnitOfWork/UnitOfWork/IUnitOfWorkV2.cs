using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitOfWork
{
    interface IUnitOfWorkV2 : IUnitOfWork
    {

        event EventHandler   BeforeCommit;
        event EventHandler BeforeRollBack;
        
        event EventHandler AfterCommit;
        event EventHandler AfterRollBack;

        
    }
}
