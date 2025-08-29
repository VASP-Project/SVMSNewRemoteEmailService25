using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Email_Send_WinService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new Service1()
            };
            ServiceBase.Run(ServicesToRun);

            // if debug mode
            //Service1 service = new Service1();
            //service.SendRemindeNovMail();
            //service.SendReminderMail();
            //service.SendRemindeAuditMail();
            //System.Threading.Thread.Sleep(System.Threading.Timeout.Infinite);


            //ServiceBase[] ServicesToRun = new ServiceBase[]
            //    {
            //                    new Service1()
            //    };

            //if (Environment.UserInteractive)
            //{
            //    RunInteractive(ServicesToRun);
            //}
            //else
            //{
            //    ServiceBase.Run(ServicesToRun);
            //}

        }


        //private static void RunInteractive(ServiceBase[] ServicesToRun)
        //{
        //    MethodInfo OnStartMethod = typeof(ServiceBase).GetMethod("OnStart",
        //    BindingFlags.Instance | BindingFlags.NonPublic);
        //    foreach (ServiceBase Service in ServicesToRun)
        //    {
        //        Console.Write("Starting {0}...", Service.ServiceName);
        //        OnStartMethod.Invoke(Service, new Object[] { new String[] { } });
        //        Console.Write("{0} Started", Service.ServiceName);
        //    }

        //    // Console.WriteLine("Press Any Key To Stop The Service {0}", Service.ServiceName);
        //    //Console.Read();
        //    //Console.WriteLine();

        //    MethodInfo OnStopMethod = typeof(ServiceBase).GetMethod("OnStop",
        //    BindingFlags.Instance | BindingFlags.NonPublic);
        //    foreach (ServiceBase Service in ServicesToRun)
        //    {
        //        Console.Write("Stopping {0}...", Service.ServiceName);
        //        OnStopMethod.Invoke(Service, null);
        //        Console.WriteLine("{0} Stopped", Service.ServiceName);
        //    }
        //    Thread.Sleep(1000);


        //}

    }
}
