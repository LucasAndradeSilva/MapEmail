using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Text;

namespace IMapMail.Service
{
    public class DBService
    {
        public string GravaEmails(string Emails, IConfigurationRoot configuration)
        {
            var con = new SqlConnection(configuration.GetSection("ConnectionString").Value);
            var cmd = new SqlCommand("P_GuardaEmail", con) { CommandType = System.Data.CommandType.StoredProcedure };
            cmd.Parameters.AddWithValue("@json", Emails);

            try
            {
                con.Open();
                var reader = cmd.ExecuteReader();
                var retorno = reader.Read() ? reader[0].ToString() : string.Empty;
                return retorno;
            }
            catch (Exception exe)
            {
                return "";
            }
            finally
            {
                con.Close();
                cmd.Dispose();
            }
        }
    }
}
