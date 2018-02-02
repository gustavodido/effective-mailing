using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Data;
using Oracle.DataAccess.Client;

using Microsoft.Office.Interop.Excel;

namespace EffectiveMailing
{
    class Program
    {
    }

    static void CreateTable()
    {
        using (OracleConnection conn = new OracleConnection(Properties.Settings.Default.ConnectionString))
        {
            conn.Open();
            OracleCommand cmd = conn.CreateCommand();
            cmd.CommandText = "CREATE TABLE " + Properties.Settings.Default.TableName + " (DT_CARGA DATE, ARQUIVO VARCHAR2(255), ACAO NUMBER, OPERADORA VARCHAR2(255), ATRIBUTO VARCHAR2(100), VALOR VARCHAR(255)) NOLOGGING";
            cmd.CommandType = CommandType.Text;
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch(OracleException ex)
            {
                WriteLog(ex.Message);
                if (ex.Number != 955)
                    throw ex;
            }               
       }
    }
}
