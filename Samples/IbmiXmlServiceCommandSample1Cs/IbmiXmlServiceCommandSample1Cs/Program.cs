using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IbmiXmlserviceStd;

namespace IbmiXmlServiceCommandSample1Cs
{
    class Program
    {
        static void Main(string[] args)
        {
            String baseUrl = "http://1.1.1.1:30000/cgi-bin/xmlcgi.pgm";
            String ibmiUser = "USER";
            String ibmiPass = "PASS";
            bool useHttpCredentials = false;
            bool rtncmd = false;
            String clCmd = "SNDMSG MSG(TESTDOTNET) TOUSR(QSYSOPR)";

            XmlServicei _ibmi = new XmlServicei();

            // Set credential information
            _ibmi.SetUserInfoExt(baseUrl, ibmiUser, ibmiPass, useHttpCredentials);

            // Run CL command
            rtncmd = _ibmi.ExecuteCommand(clCmd);
            
            if (rtncmd)
            {
                Console.WriteLine("Command was successful.");
                Environment.Exit(0);  
            } else
            {
                Console.WriteLine("Command failed.");
                Environment.Exit(99);
            }

        }
    }
}
