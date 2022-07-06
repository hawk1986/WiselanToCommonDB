using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace DANCom
{
    
    static class DANClass
    {
        private const string MISSTRING = "Data Source=twtpe1sqm01;User ID=MISUser;Password=1234QWERasdf;Initial Catalog=GCMIS";

        static public void WriteToLog(string _PrgID, int _Step, string _Status, string _UserID, string _ComputerName, string _Msg)
        {
            if (_PrgID.Trim() == "" || _Status.Trim() == "" || _UserID.Trim() == "" || _Msg.Trim() == "") return;

            using (SqlConnection cn = new SqlConnection(MISSTRING))
            {
                cn.Open();
                SqlCommand cmd = new SqlCommand("INSERT INTO ProgramLog (ProgramID, ExecStep, ExecStatus, ExecUserID, ComputerName, ExecMsg) VALUES (@prgid, @step, @status, @userid, @computername, @msg)", cn);
                cmd.Parameters.Add("@prgid", SqlDbType.VarChar);
                cmd.Parameters.Add("@step", SqlDbType.Int);
                cmd.Parameters.Add("@status", SqlDbType.Char);
                cmd.Parameters.Add("@userid", SqlDbType.VarChar);
                cmd.Parameters.Add("@computername", SqlDbType.VarChar);
                cmd.Parameters.Add("@msg", SqlDbType.NVarChar);
                cmd.Parameters["@prgid"].Value = _PrgID;
                cmd.Parameters["@step"].Value = _Step;
                cmd.Parameters["@status"].Value = _Status;
                cmd.Parameters["@userid"].Value = _UserID;
                cmd.Parameters["@computername"].Value = _ComputerName;
                cmd.Parameters["@msg"].Value = _Msg;
                //cmd.ExecuteNonQuery();
            }
        }

        static public void SendMail(string sRecipients, string sSubject, string sBody)
        {
            using (SqlConnection cn = new SqlConnection(MISSTRING))
            {
                cn.Open();
                SqlCommand cmd = new SqlCommand("INSERT INTO MailLog (MailRecipients, MailSubject, MailBody) VALUES (@rec, @sub, @body)", cn);
                cmd.Parameters.Add(new SqlParameter("@rec", SqlDbType.VarChar));
                cmd.Parameters.Add(new SqlParameter("@sub", SqlDbType.VarChar));
                cmd.Parameters.Add(new SqlParameter("@body", SqlDbType.NVarChar));

                cmd.Parameters["@rec"].Value = sRecipients;
                cmd.Parameters["@sub"].Value = sSubject;
                cmd.Parameters["@body"].Value = sBody;

                cmd.ExecuteNonQuery();
            }
        }

    } 
}
