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
    public class DAL_SVMS
    {

        string Dkey = "($h@r!(u!8*MW4oB1VmL5GIwBIjqFYQntHT0CMi2uEYAmBkwxpvsbLQ6KX1SCno9XQ==";
        string SqlconString;
        SqlConnection cn;
        SqlCommand cmd;
        DataSet ds;
        SqlDataAdapter da;

        public DAL_SVMS()
        {

            SqlconString = EncryptDecryptPassword.DecryptText(ConfigurationManager.ConnectionStrings["ConnstrSVMS"].ConnectionString, Dkey);
            cn = new SqlConnection(SqlconString);
        }

        public bool SaveSecurityClearanceData(int SPRunMinutes)
        {
            try
            {

                //LogService.WriteErrorLog("SaveSecurityClearanceData in DLL");
                DataTable dtSP = new DataTable();
                bool result = false;
                SqlDataAdapter daSP = new SqlDataAdapter("GetPendingBadgeSPDetails", cn);
                daSP.SelectCommand.CommandType = CommandType.StoredProcedure;
                daSP.SelectCommand.Parameters.Add("@Flag", SqlDbType.NVarChar).Value = "S";
                daSP.SelectCommand.Parameters.Add("@SPRunDate", SqlDbType.DateTime).Value = System.DateTime.Now;
                daSP.SelectCommand.Parameters.Add("@Id", SqlDbType.Int).Value = 0;
                daSP.Fill(dtSP);
                if (dtSP.Rows.Count > 0)
                {
                    TimeSpan span = System.DateTime.Now.Subtract(Convert.ToDateTime(dtSP.Rows[0]["SPRunDate"]));
                    if (SPRunMinutes <= Convert.ToInt32(span.TotalMinutes))
                    {
                        result = UpdatePendingBadgeData();
                        if (result)
                        {
                            cmd = new SqlCommand("GetPendingBadgeSPDetails", cn);
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@Flag", "U");
                            cmd.Parameters.AddWithValue("@SPRunDate", System.DateTime.Now);
                            cmd.Parameters.AddWithValue("@Id", Convert.ToInt32(dtSP.Rows[0][0]));
                            cn.Open();
                            cmd.ExecuteNonQuery();
                            cn.Close();
                        }
                    }
                }
                else
                {
                    result = UpdatePendingBadgeData();
                    if (result)
                    {
                        cmd = new SqlCommand("GetPendingBadgeSPDetails", cn);
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@Flag", "I");
                        cmd.Parameters.AddWithValue("@SPRunDate", System.DateTime.Now);
                        cmd.Parameters.AddWithValue("@Id", 0);
                        cn.Open();
                        cmd.ExecuteNonQuery();
                        cn.Close();
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                cn.Close();
                throw;
                // return false;
            }
            finally { cn.Close(); }
        }


        private bool UpdatePendingBadgeData()
        {
            try
            {
                //LogService.WriteErrorLog("UpdatePendingBadgeData in DLL");

                SqlDataAdapter da = new SqlDataAdapter("RptPendingBadgeApplicants", cn);
                da.SelectCommand.CommandType = CommandType.StoredProcedure;
                da.SelectCommand.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = "";

                DataTable dt = new DataTable();
                da.Fill(dt);

                cn.Close();
                //if (dt != null && dt.Rows != null && dt.Rows.Count > 0)
                //    lstPending = ConvertDataTable<EmailDataModel>(dt);

                for (int c = 0; c < dt.Rows.Count; c++)
                {
                    //if (!string.IsNullOrEmpty(Convert.ToString(dt.Rows[c][5])))
                    //{
                    //DateTime dtClearance = (string.IsNullOrEmpty(Convert.ToString(dt.Rows[c][5])) ? "" : Convert.ToDateTime(dt.Rows[c][5].ToString()));
                    cmd = new SqlCommand("SaveSecurityClearanceDetails", cn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@BadgeKey", dt.Rows[c][0].ToString());
                    cmd.Parameters.AddWithValue("@CompanyName", dt.Rows[c][1].ToString());
                    cmd.Parameters.AddWithValue("@LastName", dt.Rows[c][2].ToString());
                    cmd.Parameters.AddWithValue("@FirstName", dt.Rows[c][3].ToString());
                    cmd.Parameters.AddWithValue("@FingerprintDate", Convert.ToDateTime(dt.Rows[c][4]));
                    //cmd.Parameters.AddWithValue("@ClearanceDate", Convert.ToDateTime(dt.Rows[c][5].ToString()));
                    cmd.Parameters.AddWithValue("@ClearanceDate", dt.Rows[c][5]);
                    cmd.Parameters.AddWithValue("@FlagUpdateClearance", "N");
                    cmd.Parameters.AddWithValue("@NotificationDate", DBNull.Value);
                    cmd.Parameters.AddWithValue("@DenyDate", DBNull.Value);
                    cmd.Parameters.AddWithValue("@CurrentDate", System.DateTime.Now);
                    cmd.Parameters.AddWithValue("@NotifyBy", "");
                    cmd.Parameters.AddWithValue("@DenyBy", "");
                    cn.Open();
                    int result = cmd.ExecuteNonQuery();
                    cn.Close();
                    // }

                }
                LogService.WriteErrorLog("Clearance Report Procedure data is updated");
                return true;
            }
            catch (Exception ex)
            {
                cn.Close();
                throw;
            }
        }



        public DataTable GetReminderClearanceData(string flag)
        {
            //LogService.WriteErrorLog("GetReminderClearanceData in DLL");
            DataTable dt = new DataTable();
            try
            {

                SqlDataAdapter da = new SqlDataAdapter("ReminderClearanceData", cn);
                da.SelectCommand.CommandType = CommandType.StoredProcedure;
                da.SelectCommand.Parameters.Add("@Flag", SqlDbType.NVarChar).Value = flag;
                da.SelectCommand.Parameters.Add("@Id", SqlDbType.Int).Value = 0;
                da.SelectCommand.Parameters.Add("@CurrentDate", SqlDbType.DateTime).Value = System.DateTime.Now;
                da.Fill(dt);
                cn.Close();
                return dt;
            }
            catch (Exception ex)
            {
                cn.Close();
                throw;

            }
            finally { cn.Close(); }
        }

        //public DataTable GetBadgeReportSettings()
        //{
        //    DataTable dt = new DataTable();
        //    try
        //    {

        //        SqlDataAdapter da = new SqlDataAdapter("ReminderClearanceData", cn);
        //        da.SelectCommand.CommandType = CommandType.StoredProcedure;
        //        da.SelectCommand.Parameters.Add("@Flag", SqlDbType.NVarChar).Value = "BR";
        //        da.SelectCommand.Parameters.Add("@Id", SqlDbType.Int).Value = 0;
        //        da.Fill(dt);
        //        cn.Close();
        //        return dt;
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error(string.Format("Error found in the GetBadgeReportSettings exception is {0} ", ex.Message));
        //        return dt;

        //    }
        //    finally { cn.Close(); }
        //}

        //public DataTable GetBadgeReportEmailSettings()
        //{
        //    DataTable dt = new DataTable();
        //    try
        //    {

        //        SqlDataAdapter da = new SqlDataAdapter("ReminderClearanceData", cn);
        //        da.SelectCommand.CommandType = CommandType.StoredProcedure;
        //        da.SelectCommand.Parameters.Add("@Flag", SqlDbType.NVarChar).Value = "BR";
        //        da.SelectCommand.Parameters.Add("@Id", SqlDbType.Int).Value = 0;
        //        da.Fill(dt);
        //        cn.Close();
        //        return dt;
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error(string.Format("Error found in the GetBadgeReportSettings exception is {0} ", ex.Message));
        //        return dt;

        //    }
        //    finally { cn.Close(); }
        //}
        public bool UpdateMailSentStatus(int id)
        {
            try
            {
                //LogService.WriteErrorLog("UpdateMailSentStatus in DLL");
                cmd = new SqlCommand("ReminderClearanceData", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Flag", "UC");
                cmd.Parameters.AddWithValue("@Id", id);
                cmd.Parameters.AddWithValue("@CurrentDate", System.DateTime.Now);
                cn.Open();
                int result = cmd.ExecuteNonQuery();
                cn.Close();
                return true;
            }
            catch (Exception ex)
            {
                cn.Close();
                throw;
            }
            finally { cn.Close(); }
        }

        public bool UpdateAuditMailSentStatus(int id)
        {
            try
            {
                //LogService.WriteErrorLog("UpdateMailSentStatus in DLL");
                cmd = new SqlCommand("ReminderAuditData", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Flag", "UC");
                cmd.Parameters.AddWithValue("@Id", id);
                cmd.Parameters.AddWithValue("@CurrentDate", System.DateTime.Now);
                cn.Open();
                int result = cmd.ExecuteNonQuery();
                cn.Close();
                return true;
            }
            catch (Exception ex)
            {
                cn.Close();
                throw;
            }
            finally { cn.Close(); }
        }

        public bool UpdateNovMailSentStatus(int id)
        {
            try
            {
                //LogService.WriteErrorLog("UpdateMailSentStatus in DLL");
                cmd = new SqlCommand("ReminderNov", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Flag", "UC");
                cmd.Parameters.AddWithValue("@Id", id);
                cmd.Parameters.AddWithValue("@CurrentDate", System.DateTime.Now);
                cn.Open();
                int result = cmd.ExecuteNonQuery();
                cn.Close();
                return true;
            }
            catch (Exception ex)
            {
                cn.Close();
                throw;
            }
            finally { cn.Close(); }
        }

        //AUDIT
        public bool SaveAuditData(int SPRunMinutes)
        {
            try
            {

                //LogService.WriteErrorLog("SaveSecurityClearanceData in DLL");
                DataTable dtSP = new DataTable();
                bool result = false;
                SqlDataAdapter daSP = new SqlDataAdapter("GetPendingBadgeSPDetails", cn);
                daSP.SelectCommand.CommandType = CommandType.StoredProcedure;
                daSP.SelectCommand.Parameters.Add("@Flag", SqlDbType.NVarChar).Value = "S";
                daSP.SelectCommand.Parameters.Add("@SPRunDate", SqlDbType.DateTime).Value = System.DateTime.Now;
                daSP.SelectCommand.Parameters.Add("@Id", SqlDbType.Int).Value = 0;
                daSP.Fill(dtSP);
                if (dtSP.Rows.Count > 0)
                {
                    TimeSpan span = System.DateTime.Now.Subtract(Convert.ToDateTime(dtSP.Rows[0]["SPRunDate"]));
                    if (SPRunMinutes <= Convert.ToInt32(span.TotalMinutes))
                    {
                        result = UpdateAuditData();
                        if (result)
                        {
                            cmd = new SqlCommand("GetPendingBadgeSPDetails", cn);
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@Flag", "U");
                            cmd.Parameters.AddWithValue("@SPRunDate", System.DateTime.Now);
                            cmd.Parameters.AddWithValue("@Id", Convert.ToInt32(dtSP.Rows[0][0]));
                            cn.Open();
                            cmd.ExecuteNonQuery();
                            cn.Close();
                        }
                    }
                }
                else
                {
                    result = UpdateAuditData();
                    if (result)
                    {
                        cmd = new SqlCommand("GetPendingBadgeSPDetails", cn);
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@Flag", "I");
                        cmd.Parameters.AddWithValue("@SPRunDate", System.DateTime.Now);
                        cmd.Parameters.AddWithValue("@Id", 0);
                        cn.Open();
                        cmd.ExecuteNonQuery();
                        cn.Close();
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                cn.Close();
                throw;
                // return false;
            }
            finally { cn.Close(); }
        }

        private bool UpdateAuditData()
        {
            try
            {
                //LogService.WriteErrorLog("UpdatePendingBadgeData in DLL");

                SqlDataAdapter da = new SqlDataAdapter("RptPendingBadgeApplicants", cn);
                da.SelectCommand.CommandType = CommandType.StoredProcedure;
                da.SelectCommand.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = "";

                DataTable dt = new DataTable();
                da.Fill(dt);

                cn.Close();
                //if (dt != null && dt.Rows != null && dt.Rows.Count > 0)
                //    lstPending = ConvertDataTable<EmailDataModel>(dt);

                for (int c = 0; c < dt.Rows.Count; c++)
                {
                    //if (!string.IsNullOrEmpty(Convert.ToString(dt.Rows[c][5])))
                    //{
                    //DateTime dtClearance = (string.IsNullOrEmpty(Convert.ToString(dt.Rows[c][5])) ? "" : Convert.ToDateTime(dt.Rows[c][5].ToString()));
                    cmd = new SqlCommand("SaveSecurityClearanceDetails", cn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@BadgeKey", dt.Rows[c][0].ToString());
                    cmd.Parameters.AddWithValue("@CompanyName", dt.Rows[c][1].ToString());
                    cmd.Parameters.AddWithValue("@LastName", dt.Rows[c][2].ToString());
                    cmd.Parameters.AddWithValue("@FirstName", dt.Rows[c][3].ToString());
                    cmd.Parameters.AddWithValue("@FingerprintDate", Convert.ToDateTime(dt.Rows[c][4]));
                    //cmd.Parameters.AddWithValue("@ClearanceDate", Convert.ToDateTime(dt.Rows[c][5].ToString()));
                    cmd.Parameters.AddWithValue("@ClearanceDate", dt.Rows[c][5]);
                    cmd.Parameters.AddWithValue("@FlagUpdateClearance", "N");
                    cmd.Parameters.AddWithValue("@NotificationDate", DBNull.Value);
                    cmd.Parameters.AddWithValue("@DenyDate",  DBNull.Value);
                    cmd.Parameters.AddWithValue("@CurrentDate", System.DateTime.Now);
                    cmd.Parameters.AddWithValue("@NotifyBy", "");
                    cmd.Parameters.AddWithValue("@DenyBy", "");
                    cn.Open();
                    int result = cmd.ExecuteNonQuery();
                    cn.Close();
                    // }

                }
                LogService.WriteErrorLog("Clearance Report Procedure data is updated");
                return true;
            }
            catch (Exception ex)
            {
                cn.Close();
                throw;
            }
        }

        public DataTable GetReminderAuditData(string flag)
        {
            //LogService.WriteErrorLog("GetReminderClearanceData in DLL");
            DataTable dt = new DataTable();
            try
            {

                SqlDataAdapter da = new SqlDataAdapter("ReminderAuditData", cn);
                da.SelectCommand.CommandType = CommandType.StoredProcedure;
                da.SelectCommand.Parameters.Add("@Flag", SqlDbType.NVarChar).Value = flag;
                da.SelectCommand.Parameters.Add("@Id", SqlDbType.Int).Value = 0;
                da.SelectCommand.Parameters.Add("@CurrentDate", SqlDbType.DateTime).Value = System.DateTime.Now;
                da.Fill(dt);
                cn.Close();
                return dt;
            }
            catch (Exception ex)
            {
                cn.Close();
                throw;

            }
            finally { cn.Close(); }
        }

        public DataTable GetReminderNovData(string flag)
        {
            //LogService.WriteErrorLog("GetReminderClearanceData in DLL");
            DataTable dt = new DataTable();
            try
            {

                SqlDataAdapter da = new SqlDataAdapter("ReminderNov", cn);
                da.SelectCommand.CommandType = CommandType.StoredProcedure;
                da.SelectCommand.Parameters.Add("@Flag", SqlDbType.NVarChar).Value = flag;
                da.SelectCommand.Parameters.Add("@Id", SqlDbType.Int).Value = 0;
                da.SelectCommand.Parameters.Add("@CurrentDate", SqlDbType.DateTime).Value = System.DateTime.Now;
                da.Fill(dt);
                cn.Close();
                return dt;
            }
            catch (Exception ex)
            {
                cn.Close();
                throw;

            }
            finally { cn.Close(); }
        }

        public DataTable GetCompaniesWithMissedAuditsData(string flag)
        {
            
            DataTable dt = new DataTable();
            try
            {

                SqlDataAdapter da = new SqlDataAdapter("GetCompaniesWithMissedAudits", cn);
                da.SelectCommand.CommandType = CommandType.StoredProcedure;
                da.SelectCommand.Parameters.Add("@Flag", SqlDbType.NVarChar).Value = flag;
                da.SelectCommand.Parameters.Add("@CompanyId", SqlDbType.Int).Value = 0;
                da.SelectCommand.Parameters.Add("@AuditDate", SqlDbType.DateTime).Value = System.DateTime.Now;
                da.Fill(dt);
                cn.Close();
                return dt;
            }
            catch (Exception ex)
            {
                cn.Close();
                throw;

            }
            finally { cn.Close(); }
        }

        public bool UpdateMisingAuditMailSentStatus(int companyId, DateTime auditDate)
        {
            try
            {
                using (SqlCommand cmd = new SqlCommand("GetCompaniesWithMissedAudits", cn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@Flag", "UC");
                    cmd.Parameters.AddWithValue("@CompanyId", companyId);
                    cmd.Parameters.AddWithValue("@AuditDate", auditDate);
                    cn.Open();
                    int result = cmd.ExecuteNonQuery();
                    cn.Close();
                    return result > 0;
                }
            }
            catch (Exception ex)
            {
                cn.Close();
                throw;
            }
            finally
            {
                cn.Close();
            }
        }

        public bool IsProhibitedMissingAuditEmailSentToday()
        {
            try
            {
                using (SqlCommand cmd = new SqlCommand("GetTodayProhibitedMissingAuditEmails", cn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cn.Open();

                    int count = Convert.ToInt32(cmd.ExecuteScalar());
                    return count > 0;
                }
            }
            catch (Exception ex)
            {
                LogService.WriteErrorLog("Error in IsProhibitedMissingAuditEmailSentToday: " + ex.Message);
                throw;
            }
            finally
            {
                if (cn.State == ConnectionState.Open)
                    cn.Close();
            }
        }


    }
}
