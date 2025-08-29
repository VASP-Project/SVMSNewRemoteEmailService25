using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Email_Send_WinService
{
    public static class LogService
    {
        public static void WriteErrorLog(Exception ex)
        {
            StreamWriter sw = null;
            try
            {
                string logFile = DateTime.Now.ToString("yyyyMMdd") + ".txt";
                if (!System.IO.File.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\"+logFile))
                {
                    System.IO.File.Create(AppDomain.CurrentDomain.BaseDirectory+"\\Logs\\"+logFile).Close();
                }

                //sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\LogFile.txt", true);

                sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "Logs\\" + logFile, true);
                sw.WriteLine(DateTime.Now.ToString() + ": " + ex.Source.ToString().Trim() + "; " + ex.Message.ToString().Trim());
                sw.Flush();
                sw.Close();



            }
            catch
            {
            }
        }

        public static void WriteErrorLog(string Message)
        {
            StreamWriter sw = null;
            try
            {
                string logFile = DateTime.Now.ToString("yyyyMMdd") + ".txt";
                if (!System.IO.File.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\"+logFile))
                {
                    System.IO.File.Create(AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\" + logFile).Close();
                }
                // sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\LogFile.txt", true);


                sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "Logs\\" + logFile, true);
                sw.WriteLine(DateTime.Now.ToString() + ": " + Message);
                sw.Flush();
                sw.Close();

            }
            catch(Exception ex)
            {
            }
        }
    }
}
