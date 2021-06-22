using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitOfWork
{
   public interface IUnitOfWork
    {
        void BeginWork();
        void Commit();
        void RollBack();

       DbCommand CreateCommand();

    }
}
