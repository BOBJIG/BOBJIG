using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.Common;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitOfWork
{
    public  class EmployeeRepository:IDisposable
    {
        public readonly DbConnection _Connection;

        public EmployeeRepository()
        {
            _Connection = SqlClientFactory.Instance.CreateConnection();
            _Connection.ConnectionString = ConfigurationManager.ConnectionStrings["Default"].ConnectionString;
            _Connection.Open();
        }

        public void GetEmployee()
        {
            
        }


        public void AddEmployee(Employee emp)
        {
            var query = string.Format("INSERT INTO EMPLOYEE (NAME,ROLE,GENDER) VALUES" + "('{0}','{1}','{2}')", emp.Name,
                emp.Role, emp.Gender);


            using (var cmd=_Connection.CreateCommand())
            {
                cmd.CommandText = query;
                cmd.ExecuteNonQuery();
            }
        }


        public void Dispose()
        {
            if (_Connection != null)
            {
                _Connection.Close();
                _Connection.Dispose();
            }
        }
    }
}
