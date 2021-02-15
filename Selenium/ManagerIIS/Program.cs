using Microsoft.Web.Administration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ManagerIIS
{
    public class Program
    {
        static void Main(string[] args)
        {

            ServerManager iisManager = new ServerManager();
            foreach(WorkerProcess w3wp in iisManager.WorkerProcesses) {
                Console.WriteLine("W3WP ({0})", w3wp.ProcessId);
            
                foreach (Request request in w3wp.GetRequests(0)) {
                    Console.WriteLine("{0} – {1},{2},{3}",
                                request.Url,
                                request.ClientIPAddr,
                                request.TimeElapsed,
                                request.TimeInState);
                }
            }
        }
    }
}
