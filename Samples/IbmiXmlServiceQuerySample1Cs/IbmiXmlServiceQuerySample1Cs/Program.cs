using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
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
            bool rtnquery = false;
            String strSql = "SELECT * FROM QIWS/QCUSTCDT33";
            String strOutputFile = "c:\\rjstemp\\qcustcdt.csv";
            String strRtnVal = "";

            // Create new IBMi instance
            XmlServicei _ibmi = new XmlServicei();

            // Run query and output results if successful
            try
            {

                // Set credential information
                _ibmi.SetUserInfoExt(baseUrl, ibmiUser, ibmiPass, useHttpCredentials);

                // Query database table to CSV return string
                strRtnVal = _ibmi.ExecuteSqlQueryToCsvString(strSql, strOutputFile);

                // If no data, bail out
                if (strRtnVal.Length <= 0 )
                {
                    throw new Exception("No data returned from query. Error: " + _ibmi.GetLastXmlResponse()); 
                }

                // Write return value to output file if data returned
                rtnquery = _ibmi.WriteStringToFile(strRtnVal, strOutputFile, false, true);
                
                if (rtnquery)
                {
                    Console.WriteLine("Query was written successfully to file " + strOutputFile);
                    Environment.ExitCode=0;
                }
                else
                {
                    Console.WriteLine("Query was not written to file. Error: " + _ibmi.GetLastError());
                    Environment.ExitCode = 99;
                }

            } catch(Exception ex)
            {
                Console.WriteLine("Query failed. Error: " + ex.Message);
                Environment.ExitCode = 99;
            } finally
            {
                // Kill IBMi service jobs
                _ibmi.KillService();

                // Exit with appropriate exit code
                Environment.Exit(Environment.ExitCode);
            }

        }
    }
}
