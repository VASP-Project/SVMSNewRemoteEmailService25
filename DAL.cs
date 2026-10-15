using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Email_Send_WinService
{
    public class DAL
    {
        string Dkey = "($h@r!(u!8*MW4oB1VmL5GIwBIjqFYQntHT0CMi2uEYAmBkwxpvsbLQ6KX1SCno9XQ==";
        string SqlconString;
        SqlConnection cn;

        //SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);

        SqlCommand cmd;
        DataSet ds;
        SqlDataAdapter da;
        public DAL()
        {

            SqlconString = EncryptDecryptPassword.DecryptText(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString, Dkey);
            cn = new SqlConnection(SqlconString);
        }

        private static List<T> ConvertDataTable<T>(DataTable dt)
        {
            List<T> data = new List<T>();
            foreach (DataRow row in dt.Rows)
            {
                T item = GetItem<T>(row);
                data.Add(item);
            }
            return data;
        }
        private static T GetItem<T>(DataRow dr)
        {
            Type temp = typeof(T);
            T obj = Activator.CreateInstance<T>();

            foreach (DataColumn column in dr.Table.Columns)
            {
                foreach (PropertyInfo pro in temp.GetProperties())
                {
                    if (pro.Name == column.ColumnName)
                        if (dr[column.ColumnName] != DBNull.Value)
                            pro.SetValue(obj, dr[column.ColumnName], null);
                        //pro.SetValue(obj, dr[column.ColumnName], null);
                        else
                        continue;
                }
            }
            return obj;
        }

        public List<EmailConfigModel> GetEmailConfigs(string applicationName)
        {
            try
            {
                //LogService.WriteErrorLog("GetEmailConfigs in DLL");
                List<EmailConfigModel> emailConfigs = new List<EmailConfigModel>();
                SqlDataAdapter da = new SqlDataAdapter("GetEmailConfigs", cn);
                da.SelectCommand.CommandType = CommandType.StoredProcedure;
                da.SelectCommand.Parameters.Add("@ApplicationName", SqlDbType.NVarChar).Value = applicationName;
                DataTable dt = new DataTable();
                da.Fill(dt);
                if (dt != null && dt.Rows != null && dt.Rows.Count > 0)
                    emailConfigs = ConvertDataTable<EmailConfigModel>(dt);

                return emailConfigs;
            }
            catch
            {
                cn.Close();
                throw;
            }
            finally
            {
                cn.Close();
            }
        }

        public List<EmailDataModel> GetEmailData(string applicationName)
        {
            try
            {
                //LogService.WriteErrorLog("GetEmailData in DLL");
                List<EmailDataModel> emailDatas = new List<EmailDataModel>();
                SqlDataAdapter da = new SqlDataAdapter("GetEmailData", cn);
                da.SelectCommand.CommandType = CommandType.StoredProcedure;
                da.SelectCommand.Parameters.Add("@ApplicationName", SqlDbType.NVarChar).Value = applicationName;
                DataTable dt = new DataTable();
                da.Fill(dt);
                if (dt != null && dt.Rows != null && dt.Rows.Count > 0)
                    emailDatas = ConvertDataTable<EmailDataModel>(dt);

                return emailDatas;
            }
            catch
            {
                cn.Close();
                throw;
            }
            finally
            {
                cn.Close();
            }
        }

        public bool UpdateMailSentStatus(int id)
        {
            try
            {
                //LogService.WriteErrorLog("UpdateMailSentStatus in DLL");
                cmd = new SqlCommand("UpdateMailSentStatus", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Id", id);
                cn.Open();
                int result = cmd.ExecuteNonQuery();
                cn.Close();
                return true;
            }
            catch (Exception ex)
            {
                LogService.WriteErrorLog(string.Format("Error found in the UpdateMailSentStatus exception is {0} ", ex.Message));
                return false;
            }
            finally { cn.Close(); }
        }

        public bool UpdateRowProcessStatus(int id)
        {
            try
            {
                //LogService.WriteErrorLog("UpdateRowProcessStatus in DLL");
                cmd = new SqlCommand("UpdateRowProcessStatus", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Id", id);
                cn.Open();
                int result = cmd.ExecuteNonQuery();
                cn.Close();
                return true;
            }
            catch (Exception ex)
            {
                LogService.WriteErrorLog(string.Format("Error found in the UpdateRowProcessStatus exception is {0} ", ex.Message));
                return false;
            }
            finally { cn.Close(); }
        }

        public DataTable GetPendingSchedulerInspectionNotifications()
        {
            DataTable dt = new DataTable();

            try
            {
                SqlDataAdapter da = new SqlDataAdapter(
                    "GetPendingSchedulerInspectionNotifications", cn);

                da.SelectCommand.CommandType =
                    CommandType.StoredProcedure;

                da.Fill(dt);

                return dt;
            }
            catch (Exception ex)
            {
                LogService.WriteErrorLog(
                    "Error found in the GetPendingSchedulerInspectionNotifications exception is "
                    + ex.Message);

                return dt;
            }
            finally
            {
                cn.Close();
            }
        }

        public bool UpdateAwsSlotOverdueNotificationSent(int inspectionId)
        {
            try
            {
                using (SqlCommand cmd = new SqlCommand(
                    "UpdateAwsSlotOverdueNotificationSent", cn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue(
                        "@InspectionId",
                        inspectionId);

                    if (cn.State != ConnectionState.Open)
                    {
                        cn.Open();
                    }

                    cmd.ExecuteNonQuery();

                    return true;
                }
            }
            catch (Exception ex)
            {
                LogService.WriteErrorLog(
                    "Error in UpdateAwsSlotOverdueNotificationSent: "
                    + ex);

                return false;
            }
            finally
            {
                cn.Close();
            }
        }


        public bool SendEmailUsingService(string ApplicationName, string toMail, string ccMail, string bccMail, string subject, string message, string attchementPath)
        {
            try
            {
                //LogService.WriteErrorLog("SendEmailUsingService in DLL");
                cmd = new SqlCommand("InsertEmailData", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@ApplicationName", SqlDbType.NVarChar).Value = ApplicationName;
                cmd.Parameters.AddWithValue("@ToMail", SqlDbType.NVarChar).Value = toMail;
                cmd.Parameters.AddWithValue("@CCMail", SqlDbType.NVarChar).Value = ccMail;
                cmd.Parameters.AddWithValue("@BCCMail", SqlDbType.NVarChar).Value = bccMail;
                cmd.Parameters.AddWithValue("@Subject", SqlDbType.NVarChar).Value = subject;
                cmd.Parameters.AddWithValue("@MailBody", SqlDbType.NVarChar).Value = message;
                cmd.Parameters.AddWithValue("@AttachmentPath", SqlDbType.NVarChar).Value = attchementPath;
                cn.Open();
                int result = cmd.ExecuteNonQuery();
                cn.Close();

                return true;

            }
            catch (Exception ex)
            {
                LogService.WriteErrorLog(string.Format("Error found in the SendEmailUsingService exception is {0} ", ex.Message));
                return false;

            }
            finally { cn.Close(); }

        }

        public DataTable GetServiceRunTime()
        {
            DataTable dtSP = new DataTable();
            try
            {
                //LogService.WriteErrorLog("GetServiceRunTime in DLL");
                bool result = false;
                SqlDataAdapter daSP = new SqlDataAdapter("GetServiceRunDetails", cn);
                daSP.SelectCommand.CommandType = CommandType.StoredProcedure;
                daSP.SelectCommand.Parameters.Add("@Flag", SqlDbType.NVarChar).Value = "S";
                daSP.SelectCommand.Parameters.Add("@Id", SqlDbType.Int).Value = 0;
                daSP.Fill(dtSP);

                return dtSP;
            }
            catch (Exception ex)
            {
                LogService.WriteErrorLog(string.Format("Error found in the GetServiceRunTime exception is {0} ", ex.Message));
                return dtSP;
            }
            finally { cn.Close(); }
        }

        public void InsertUpdateServiceRunTime(string flag, int Id)
        {

            try
            {
                //LogService.WriteErrorLog("InsertUpdateServiceRunTime in DLL");
                if (flag == "I")
                {
                    cmd = new SqlCommand("GetServiceRunDetails", cn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@Flag", "I");
                    cmd.Parameters.AddWithValue("@Id", 0);
                    cn.Open();
                    cmd.ExecuteNonQuery();
                    cn.Close();
                }
                else
                {
                    cmd = new SqlCommand("GetServiceRunDetails", cn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@Flag", "U");
                    cmd.Parameters.AddWithValue("@Id", Id);
                    cn.Open();
                    cmd.ExecuteNonQuery();
                    cn.Close();

                }
            }
            catch (Exception ex)
            {
                LogService.WriteErrorLog(string.Format("Error found in the InsertUpdateServiceRunTime exception is {0} ", ex.Message));

            }


        }

        public bool InsertAwsMissingInspectionNotificationLog(int schedulerInputId,string schedulerSlotId,DateTime scheduledEndDateTime, DateTime notificationTime)
        {
            try
            {
                using (SqlConnection connection =
                    new SqlConnection(SqlconString))
                {
                    connection.Open();

                    string query = @"
                INSERT INTO dbo.AwsMissingInspectionNotificationLog
                (
                    SchedulerInputId,
                    SchedulerSlotId,
                    ScheduledEndDateTime,
                    NotificationTime
                )
                SELECT
                    @SchedulerInputId,
                    @SchedulerSlotId,
                    @ScheduledEndDateTime,
                    @NotificationTime
                WHERE NOT EXISTS
                (
                    SELECT 1
                    FROM dbo.AwsMissingInspectionNotificationLog
                    WHERE SchedulerInputId = @SchedulerInputId
                      AND SchedulerSlotId = @SchedulerSlotId
                );";

                    using (SqlCommand command =
                        new SqlCommand(query, connection))
                    {
                        command.Parameters.Add(
                            "@SchedulerInputId", SqlDbType.Int)
                            .Value = schedulerInputId;

                        command.Parameters.Add(
                            "@SchedulerSlotId", SqlDbType.VarChar, 100)
                            .Value = schedulerSlotId;

                        command.Parameters.Add(
                            "@ScheduledEndDateTime", SqlDbType.DateTime)
                            .Value = scheduledEndDateTime;

                        command.Parameters.Add(
                            "@NotificationTime", SqlDbType.DateTime)
                            .Value = notificationTime;

                        return command.ExecuteNonQuery() == 1;
                    }
                }
            }
            catch (Exception ex)
            {
                LogService.WriteErrorLog(
                    "InsertAwsMissingInspectionNotificationLog Error: "
                    + ex);

                return false;
            }
        }
    }

}
