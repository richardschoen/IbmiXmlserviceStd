using System;
using System.Collections.Generic;
using System.Collections;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic;
using System.Xml;
using System.Data;
using System.Net;
using System.Web;
using System.Runtime.Serialization;

/// <summary>
/// This class is used to interface with existing IBM i database, program calls, CL commands, service programs and 
///  data queues via the XMLSERVICE service program. It also needs an Apache instance hooked up for XMLCGI calls
///  since this class uses the HTTP interface. 
///  
///  Return data is returned in a usable .Net DataTable format or you can process the raw XML responses yourself.
///  
///  This class should work with V5R4 and above of the OS400/IBM i operating system running XMLSERVICE.
///  
///  Requirement: The XMLSERVICE library must also exist and be compiled on the system, including the XMLCGI program. 
///  Apache instance must also be set up and running and configured for XMLCGI calls.
///  
///  Note: For appropriate security you should configure your Apache instance for SSL and for user Authentication. This
///  way all traffic is secured and there is an extra User/Password layer. To enable HTTP authentication use the 
///  UseHttpCredentials parameter and set it to True on the SetHttpUserInfo method.
///  
///  You can always refer to the Yips site for more install info on XMLSERVICE: 
///  http://yips.idevcloud.com/wiki/index.php/xmlservice/xmlserviceinstall 
///  </summary>
///  <remarks>Tested with XML Tookit library (XMLSERVICE) V1.9.7</remarks>
/// 

namespace IbmiXmlserviceStd
{

public class XmlServicei
{
    private bool _bXMLIsLoaded = false;
    private string _LastError;
    private DataTable _dtColumnDefinitions;
    private int _iColumnCount;
    private int _iRowCount;
    private DataSet _dsXmlResponseData;
    private DataTable _dtReturnData;
    private DataSet _dsCommandResponse;
    private DataTable _dtCommandResponse;
    private DataSet _dsProgramResponse;
    private DataTable _dtProgramResponse;
    private string _BaseURL = "";
    private string _Db2Parm = "db2=@@db2value&uid=@@uidvalue&pwd=@@pwdvalue&ipc=@@ipcvalue&ctl=@@ctlvalue&xmlin=@@xmlinvalue&xmlout=@@xmloutvalue";
    private string _User = "";
    private string _Password = "";
    private string _IPCINfo = "/tmp/xmlservicei";
    private string _IPCPersistance = "*sbmjob";
    private string _cDataOption = "*cdata";
    private string _DB2Info = "*LOCAL";
    private string _LastHTTPResponse = "";
    private string _LastXMLResponse = "";
    private int _iXmlResponseBufferSize = 500000;
    private int _HttpTimeout = 60000;
    private bool _UseHttpCredentials = false;
    private string _HttpUser = "";
    private string _HttpPassword = "";
    private string _LastDataQueueName = "";
    private string _LastDataQueueLibrary = "";
    private string _LastDataQueueDataSent = "";
    private string _LastDataQueueDataReceived = "";
    private int _LastDataQueueLengthReceived = 0;
    private string _CrLf = "\r\n";
    private bool _allowInvalidSslCertificates = false;
    private string _LastPostData = "";

        /// <summary>
        ///  Program call parameter structure
        ///  </summary>
        ///  <remarks></remarks>
        public struct PgmParmList
    {
        public string parmtype;
        public string parmvalue;
    }
    /// <summary>
    ///  Program call return structure
    ///  </summary>
    ///  <remarks></remarks>
    public struct RtnPgmCall
    {
        public bool success;
        public PgmParmList[] parms;
    }
    /// <summary>
    ///  Stored procedure parameter list
    ///  </summary>
    ///  <remarks></remarks>
    public struct ProcedureParmList
    {
        public string parmtype;
        public string parmvalue;
        public string parmname;
    }
    /// <summary>
    ///  Stored procedure return structure
    ///  </summary>
    ///  <remarks></remarks>
    public struct RtnProcedureCall
    {
        public bool success;
        public PgmParmList[] parms;
    }


    /// <summary>
    ///  Set HTTP timeout for XMLSERVICE requests
    ///  </summary>
    ///  <param name="iHttpTimeout">HTTP timeout in milliseconds</param>
    ///  <returns>True-Success, False-Fail</returns>
    ///  <remarks></remarks>
    public bool SetHttpTimeout(int iHttpTimeout)
    {
        try
        {
            _LastError = "";
            _HttpTimeout = iHttpTimeout;
            return true;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  Set base URL to XMLCGI program.
    ///  Example: http://1.1.1.1:30000/cgi-bin/xmlcgi.pgm
    ///  Set this value one time each time the class is instantiated.
    ///  </summary>
    ///  <param name="sBaseUrl">Base URL to set for path to XMLSERVICE and XMLCGI calls.</param>
    ///  <returns>True-Success, False-Fail</returns>
    ///  <remarks></remarks>
    public bool SetBaseURL(string sBaseUrl)
    {
        try
        {
            _LastError = "";
            _BaseURL = sBaseUrl;
            return true;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }
        /// <summary>
        ///  Set ALL base user info parameters for XMLCGI program calls in a single method call.
        ///  Set this value one time each time the class Is instantiated.
        ///  This is a convenience method to set all connection info in a single call.
        ///  </summary>
        ///  <param name="sBaseUrl">Base URL to set for path to XMLSERVICE and XMLCGI calls.</param>
        ///  <param name="sUser">IBM i User and HTTP Auth</param>
        ///  <param name="sPassword">IBM i Password and HTTP Auth</param>
        ///  <param name="UseHttpCredentials">Use Apache HTTP authentication credentials</param>
        ///  <param name="sDb2Info">DB2 server info. Default = *LOCAL for current DB2 server</param>
        ///  <param name="sIpcInfo">IPC info. Example: /tmp/xmlservicei</param>
        ///  <param name="persistjobs">True-Stateful XTOOLKIT jobs are started that must be ended eventually with KillService method. False-Stateless jobs. XMLSERVICE XTOOLKIT jobs end immediately after call completes.</param>
        ///  <param name="sHttpUser">Http Auth user. Only use if HTTP auth credentials are different than IBMi user info and web server auth enabled.</param>
        ///  <param name="sHttpPass">Http Auth password. Only use if HTTP auth credentials are different than IBMi user info and web server auth enabled.</param>
        ///  <param name="iSize">XML response buffer size. Default = 500000</param>
        ///  <param name="bEenableCdata">Enable wrapping returned query fields with CDATA tags automatically. True=Enable CData, False=Disable CData. Default=True</param>
        ///  <param name="allowInvalidSslCertificates">Optional Allow invallid certs. true=Yes, false=no. Default=false - certs must be valid.</param>
        ///  <returns>True-Success, False-Fail</returns>
        ///  <remarks></remarks>
        public bool SetUserInfo(string sBaseUrl, string sUser, string sPassword, bool UseHttpCredentials, string sIpcInfo = "/tmp/xmlservicei", bool persistJobs = true, string sDb2Info = "*LOCAL", int iHttpTimeout = 60000, string sHttpUser = "", string sHttpPass = "", int iSize = 500000, bool bEenableCdata = true, bool allowInvalidSslCertificates = false)
    {
        try
        {
            _LastError = "";

            // Set this before IPC info gets set since CData value is set there.
            if (SetCdataOption(bEenableCdata) == false)
                throw new Exception("Error setting CData option");

            if (SetBaseURL(sBaseUrl) == false)
                throw new Exception("Error setting base URL");

            if (SetUserInfo(sUser, sPassword, UseHttpCredentials) == false)
                throw new Exception("Error setting user info");

            if (SetDb2Info(sDb2Info) == false)
                throw new Exception("Error setting DB2 info");

            if (SetHttpTimeout(iHttpTimeout) == false)
                throw new Exception("Error setting HTTP timeout info");

            if (SetIpcInfo(sIpcInfo, persistJobs) == false)
                throw new Exception("Error setting IPC info");

            // Change HTTP auth user and password to be other than default IBMi credentials
            if (sHttpPass.Trim() != "" & sHttpUser.Trim() != "")
            {
                if (SetHttpUserInfo(sHttpUser, sHttpPass, UseHttpCredentials))
                    throw new Exception("Error setting HTTP user info");
            }

            if (SetXmlResponseBufferSize(iSize) == false)
                throw new Exception("Error setting XML response buffer size");

            // Set allow invalid certificates
            _allowInvalidSslCertificates = allowInvalidSslCertificates;

                return true;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }

    /// <summary>
    ///  Set base user info for XMLCGI program calls.
    ///  Set this value one time each time the class Is instantiated.
    ///  This sets the IBM i user login info And also sets the default
    ///  HTTP auth user credentials for the Apache server if HTTP authentication 
    ///  Is enabled on the Apache server. 
    ///  </summary>
    ///  <param name="sUser">IBM i User</param>
    ///  <param name="sPassword">IBM i Password</param>
    ///  <param name="UseHttpCredentials">Use Apache HTTP authentication credentials.</param>
    ///  <returns>True-Success, False-Fail</returns>
    ///  <remarks></remarks>
    public bool SetUserInfo(string sUser, string sPassword, bool UseHttpCredentials)
    {
        try
        {
            _LastError = "";
            // Set IBM i user info
            _User = sUser;
            _Password = sPassword;
            // Set IBM i apache authentication user info default'
            _HttpUser = sUser;
            _HttpPassword = sPassword;
            // Set use Apache HTTP authentication flag
            _UseHttpCredentials = UseHttpCredentials;
            return true;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  Set HTTP Apache authenticated user credential info for XMLCGI program calls.
    ///  Set this value one time each time the class Is instantiated.
    ///  Also make sure you call SetUserInfo first which sets the the IBM i user And 
    ///  the default HTTP user And password to the same User And Password as SetUserInfo. 
    ///  The IBM i user profile And password can be overridden with SetHttpUserInfo 
    ///  if the Apache authentication user Is Not the same as the IBM i user profile 
    ///  Or it uses an authorization list Or LDAP user for HTTP authentication
    ///  if HTTP auth user Is different than IBM i user.
    ///  </summary>
    ///  <param name="sUser">IBM i Apache HTTP Server Web Site Auth User</param>
    ///  <param name="sPassword">IBM i Apache HTTP Server Web Site Auth Password</param>
    ///  <returns>True-Success, False-Fail</returns>
    ///  <remarks></remarks>
    public bool SetHttpUserInfo(string sUser, string sPassword, bool UseHttpCredentials)
    {
        try
        {
            _LastError = "";
            _HttpUser = sUser;
            _HttpPassword = sPassword;
            // Set use Apache HTTP authentication flag
            _UseHttpCredentials = UseHttpCredentials;
            return true;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }

    /// <summary>
    ///  Set IPC info and stateful or stateless persistence info.
    ///  The class defaults to /tmp/xmlservicei so this call not needed unless you want to use custom IPC info
    ///  </summary>
    ///  <param name="sIpcInfo">IPC info. Use a unique path for each persistent job. Example: /tmp/xmlservicei</param>
    ///  <param name="persistjobs">True-Stateful XTOOLKIT jobs are started that must be ended eventually with KillService method. False-Stateless jobs. XMLSERVICE XTOOLKIT jobs end immediately after call completes.</param>
    ///  <returns>True-Success, False-Fail</returns>
    ///  <remarks></remarks>
    public bool SetIpcInfo(string sIpcInfo = "/tmp/xmlservicei", bool persistJobs = true)
    {
        try
        {
            _LastError = "";
            // Set IPC info only if a valur is passed. 
            if (sIpcInfo != "")
                _IPCINfo = sIpcInfo;
            // Read here for more on persistence.
            // http://yips.idevcloud.com/wiki/index.php/XMLService/XMLSERVICEConnect

            if (persistJobs)
                _IPCPersistance = "*sbmjob " + _cDataOption;  // Enable CDATA on control option if set
            else
                _IPCPersistance = "*here " + _cDataOption;// Enable CDATA on control option if set

            return true;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  Set Cdata status. Cdata will be enabled by default.
    ///  The class defaults to *cdata enabled so this call not needed unless you want to turn CDATA on or off.
    ///  </summary>
    ///  <param name="enableCdata">Enable wrapping returned query fields with CDATA tags automatically. True=Enable CData, False=Disable CData. Default=True</param>
    ///  <returns>True-Success, False-Fail</returns>
    ///  <remarks></remarks>
    public bool SetCdataOption(bool enableCdata = true)
    {
        try
        {
            _LastError = "";

            if (enableCdata)
                _cDataOption = "*cdata";
            else
                // *cdata(off/on) switch doesn't seem to be needed so we just eliminate *cdata keyword for disabled.
                _cDataOption = "";

            return true;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  Set DB2 system info. Defaults to *LOCAL to access local system database
    ///  </summary>
    ///  <param name="sDB2Info">DB2 server info. Default = *LOCAL for current DB2 server</param>
    ///  <returns>True-Success, False-Fail</returns>
    ///  <remarks></remarks>
    public bool SetDb2Info(string sDB2Info = "*LOCAL")
    {
        try
        {
            _LastError = "";
            _DB2Info = sDB2Info;
            return true;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  Set XML response buffer size 
    ///  </summary>
    ///  <param name="iSize">XML response buffer size. Default = 500000</param>
    ///  <returns>True-Success, False-Fail</returns>
    ///  <remarks></remarks>
    public bool SetXmlResponseBufferSize(int iSize = 500000)
    {
        try
        {
            _LastError = "";
            _iXmlResponseBufferSize = iSize;
            return true;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }



    /// <summary>
    ///  This function gets the XML Datatable of data that was loaded with LoadPgmCallDataSetFromXmlString 
    ///  </summary>
    ///  <returns>Data table of data or nothing if no data set</returns>
    ///  <remarks></remarks>
    public DataTable GetProgramResponseDataTable()
    {
        try
        {
            _LastError = "";

            if (_bXMLIsLoaded == false)
                throw new Exception("No XML data is currently loaded. Call LoadPgmCallDataTableFromXmlResponseString to load an XML file first.");

            return _dtProgramResponse;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return null;
        }
    }
    /// <summary>
    ///  This function gets the XML Datatable of data that was loaded with LoadCmdCallDataSetFromXmlString 
    ///  </summary>
    ///  <returns>Data table of data or nothing if no data set</returns>
    ///  <remarks></remarks>
    public DataTable GetCommandResponseDataTable()
    {
        try
        {
            _LastError = "";

            if (_bXMLIsLoaded == false)
                throw new Exception("No XML data is currently loaded. Call LoadCmdCallDataTableFromXmlResponseString to load an XML file first.");

            return _dtCommandResponse;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return null;
        }
    }
    /// <summary>
    ///  This function gets the XML DataSet of data that was loaded with LoadDataSetFromXMLFile
    ///  Only use this DataSet if you want to access the raw XML response data without further processing.
    ///  </summary>
    ///  <returns>DataSet of XML response data or nothing if no data set</returns>
    ///  <remarks></remarks>
    public DataSet GetXmlResponseDataSet()
    {
        try
        {
            _LastError = "";

            if (_bXMLIsLoaded == false)
                throw new Exception("No XML data is currently loaded. Call LoadDataTableFromXmlResponseFile or LoadDataTableFromXmlResponseString to load an XML file first.");

            return _dsXmlResponseData;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return null;
        }
    }
    /// <summary>
    ///  Returns last error message string
    ///  </summary>
    ///  <returns>Last error message string</returns>
    ///  <remarks></remarks>
    public string GetLastError()
    {
        try
        {
            return _LastError;
        }
        catch (Exception)
        {
            return "";
        }
    }
        /// <summary>
        ///  Returns last request that was posted
        ///  </summary>
        ///  <returns>Last post message string</returns>
        ///  <remarks></remarks>
        public string GetLastPostData()
        {
            try
            {
                return _LastPostData;
            }
            catch (Exception)
            {
                return "";
            }
        }
        /// <summary>
        ///  Returns XML response message string from last XMLSERVICE call.
        ///  </summary>
        ///  <returns>Last XML response message string</returns>
        ///  <remarks></remarks>
        public string GetLastXmlResponse()
    {
        try
        {
            return _LastXMLResponse;
        }
        catch (Exception)
        {
            return "";
        }
    }
    /// <summary>
    ///  This function gets the DataTable of data loaded from XML with LoadDataSetFromXMLFile and returns as a CSV string
    ///  </summary>
    ///  <param name="sFieldSepchar">Field delimiter/separator. Default = Comma</param>
    ///  <param name="sFieldDataDelimChar">Field data delimiter character. Default = double quotes.</param>
    ///  <returns>CSV string from DataTable</returns>
    public string GetQueryResultsDataTableToCsvString(string sFieldSepchar = ",", string sFieldDataDelimChar = "\"")
    {
        try
        {
            _LastError = "";

            //string sHeadings = "";
            //string sBody = "";
            StringBuilder sCsvData = new StringBuilder();

            // first write a line with the columns name
            string sep = "";
            System.Text.StringBuilder builder = new System.Text.StringBuilder();
            foreach (DataColumn col in _dtReturnData.Columns)
            {
                builder.Append(sep).Append(col.ColumnName);
                sep = sFieldSepchar;
            }
            sCsvData.AppendLine(builder.ToString());

            // then write all the rows
            foreach (DataRow row in _dtReturnData.Rows)
            {
                sep = "";
                builder = new System.Text.StringBuilder();

                foreach (DataColumn col in _dtReturnData.Columns)
                {
                    builder.Append(sep);
                    builder.Append(sFieldDataDelimChar).Append(row[col.ColumnName]).Append(sFieldDataDelimChar);
                    sep = sFieldSepchar;
                }
                sCsvData.AppendLine(builder.ToString());
            }

            // Return CSV output
            return sCsvData.ToString();
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return "";
        }
    }
    /// <summary>
    ///  This function gets the DataTable of XML data loaded from the last query with LoadDataSetFromXMLFile and returns as a CSV file
    ///  </summary>
    ///  <param name="sOutputFile">Output CSV file</param>
    ///  <param name="sFieldSepchar">Field delimiter/separator. Default = Comma</param>
    ///  <param name="sFieldDataDelimChar">Field data delimiter character. Default = double quotes.</param>
    ///  <param name="replace">Replace output file True=Replace file,False=Do not replace</param>
    ///  <returns>True-CSV file written successfully, False-Failure writing CSV output file.</returns>
    public bool GetQueryResultsDataTableToCsvFile(string sOutputFile, string sFieldSepchar = ",", string sFieldDataDelimChar = "\"", bool replace = false)
    {
        string sCsvWork;

        try
        {
            _LastError = "";

            // Delete existing file if replacing
            if (File.Exists(sOutputFile))
            {
                if (replace)
                    File.Delete(sOutputFile);
                else
                    throw new Exception("Output file " + sOutputFile + " already exists and replace not selected.");
            }

            // Get data and output
            using (System.IO.StreamWriter writer = new System.IO.StreamWriter(sOutputFile))
            {

                // Get CSV string
                sCsvWork = GetQueryResultsDataTableToCsvString(sFieldSepchar, sFieldDataDelimChar);

                // Write out CSV data
                writer.Write(sCsvWork);

                // Flush final output and close
                writer.Flush();
                writer.Close();

                return true;
            }
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }

    /// <summary>
    ///  This function gets the DataTable of data loaded from XML with LoadDataSetFromXMLFile and returns as a XML string
    ///  </summary>
    ///  <param name="sTableName">Table name. Default = "Table1"</param>
    ///  <param name="bWriteSchema">Write XML schema in return data</param>
    ///  <returns>XML string from data table</returns>
    public string GetQueryResultsDataTableToXmlString(string sTableName = "Table1", bool bWriteSchema = false)
    {
        string sRtnXml = "";

        try
        {
            _LastError = "";

            // if table not set, default to Table1
            if (sTableName.Trim() == "")
                sTableName = "Table1";

            // Export results to XML
            if (_dtReturnData == null == false)
            {
                StringBuilder SB = new StringBuilder();
                System.IO.StringWriter SW = new System.IO.StringWriter(SB);
                _dtReturnData = GetQueryResultsDataTable();
                _dtReturnData.TableName = sTableName;
                // Write XMl with or without schema info
                if (bWriteSchema)
                    _dtReturnData.WriteXml(SW, System.Data.XmlWriteMode.WriteSchema);
                else
                    _dtReturnData.WriteXml(SW);
                sRtnXml = SW.ToString();
                SW.Close();
                return sRtnXml;
            }
            else
                throw new Exception("No data available. Error: " + GetLastError());
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return "";
        }
    }
    /// <summary>
    ///  This function gets the DataTable of XML data loaded from the last query with LoadDataSetFromXMLFile and returns as a CSV file
    ///  </summary>
    ///  <param name="sOutputFile">Output CSV file</param>
    ///  <param name="sTableName">Table name. Default = "Table1"</param>
    ///  <param name="bWriteSchema">Write XML schema in return data</param>
    ///  <param name="replace">Replace output file True=Replace file,False=Do not replace</param>
    ///  <returns>True-XML file written successfully, False-Failure writing XML output file.</returns>
    public bool GetQueryResultsDataTableToXmlFile(string sOutputFile, string sTableName = "Table1", bool bWriteSchema = false, bool replace = false)
    {
        string sXmlWork;

        try
        {
            _LastError = "";

            // Delete existing file if replacing
            if (File.Exists(sOutputFile))
            {
                if (replace)
                    File.Delete(sOutputFile);
                else
                    throw new Exception("Output file " + sOutputFile + " already exists and replace not selected.");
            }

            // Get data and output 
            using (System.IO.StreamWriter writer = new System.IO.StreamWriter(sOutputFile))
            {

                // Get XML string
                sXmlWork = GetQueryResultsDataTableToXmlString(sTableName, bWriteSchema);

                // Write out CSV data
                writer.Write(sXmlWork);

                // Flush final output and close
                writer.Flush();
                writer.Close();

                return true;
            }
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  This function gets the DataTable of data loaded from XML with LoadDataSetFromXMLFile and returns as a JSON string
    ///  </summary>
    ///  <returns>CSV string from DataTable</returns>
    public string GetQueryResultsDataTableToJsonString(bool debugInfo = false)
    {

        // TODO - Use Newtonsoft JSON to convert to JSON

        string sJsonData = "";
        JsonHelper oJsonHelper = new JsonHelper();

        try
        {
            _LastError = "";

            // If data table is blank, bail
            if (_dtReturnData == null)
                throw new Exception("Data table is Nothing. No data available.");

            // Convert DataTable to JSON
            sJsonData = oJsonHelper.DataTableToJsonWithStringBuilder(_dtReturnData, debugInfo);

            // Return JSON output
            return sJsonData.ToString();
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return "";
        }
    }
    /// <summary>
    ///  This function gets the DataTable of XML data loaded from the last query with LoadDataSetFromXMLFile and returns as a JSON file
    ///  </summary>
    ///  <param name="sOutputFile">Output JSON file</param>
    ///  <param name="replace">Replace output file True=Replace file,False=Do not replace</param>
    ///  <returns>True-JSON file written successfully, False-Failure writing JSON output file.</returns>
    public bool GetQueryResultsDataTableToJsonFile(string sOutputFile, bool replace = false)
    {
        string sJsonWork;

        try
        {
            _LastError = "";

            // Delete existing file if replacing
            if (File.Exists(sOutputFile))
            {
                if (replace)
                    File.Delete(sOutputFile);
                else
                    throw new Exception("Output file " + sOutputFile + " already exists and replace not selected.");
            }

            // Get data and output 
            using (System.IO.StreamWriter writer = new System.IO.StreamWriter(sOutputFile))
            {

                // Get JSON string
                sJsonWork = GetQueryResultsDataTableToJsonString();

                // Write out JSON data
                writer.Write(sJsonWork);

                // Flush final output and close
                writer.Flush();
                writer.Close();

                return true;
            }
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  This function gets the DataTable containing records from the last query response data loaded from XML.
    ///  </summary>
    ///  <returns>Data table of data or nothing if no data set</returns>
    ///  <remarks></remarks>
    public DataTable GetQueryResultsDataTable()
    {
        try
        {
            _LastError = "";

            if (_bXMLIsLoaded == false)
                throw new Exception("No XML data is currently loaded. Call LoadDataSetFromXMLFile to load an XML file first.");

            return _dtReturnData;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return null;
        }
    }

    /// <summary>
    ///  Perform HTTP Get with selected URL to a string
    ///  </summary>
    ///  <param name="sURL">URL to Get</param>
    ///  <returns>String return value</returns>
    public string GetUrlToString(string sURL)
    {
        try
        {
            _LastError = "";

            System.Net.HttpWebRequest fr;
            Uri targetURI = new Uri(sURL);
            string s1 = "";
            fr = (System.Net.HttpWebRequest)HttpWebRequest.Create(targetURI);
            fr.Method = "GET";
            if ((fr.GetResponse().ContentLength > 0))
            {
                System.IO.StreamReader str = new System.IO.StreamReader(fr.GetResponse().GetResponseStream());
                s1 = str.ReadToEnd();
                str.Close();
            }

            return s1;
        }
        catch (WebException ex)
        {
            _LastError = ex.Message;
            return "";
        }
    }



    /// <summary>
    ///  This function runs an SQL INSERT, UPDATE, DELETE or other action query against the DB2 database with selected SQL statement.
    ///  </summary>
    ///  <param name="sSQL">SQL INSERT, UPDATE and DELETE. Select is not allowed </param>
    ///  <param name="sQueryResultOutputFile">PC output file for XML response data</param>
    ///  <returns>True - Query service call succeeded, False - Query service call failed</returns>
    ///  <remarks>Note: Committment control is disabled via the commit='none' option so journaling is not used at the moment on any files you plan to modify via INSERT/UPDATE/DELETE</remarks>
    public bool ExecuteSqlNonQuery(string sSQL, string sQueryResultOutputFile = "")
    {
        string sdb2parm = _Db2Parm;
        string sRtnXML = "";
        bool rtnexecute;
        string sSuccessValue = "<execute stmt='stmt1'>@@CHR10<success>+++ success stmt1</success>@@CHR10</execute>";
        string sSuccessValue2 = "+++ success stmt1";
        string sSqlWork = sSQL;
        // <success>         +++ success stmt1   </success>
        try
        {
            _LastError = "";
            sRtnXML = "";
            _LastHTTPResponse = "";
            _LastXMLResponse = "";

            // Clear data sets
            _dsXmlResponseData = null;
            _dtReturnData = null;

            // Replace success values with correct ascii values. The return data contains a line feed
            sSuccessValue = sSuccessValue.Replace("@@CHR10", _CrLf);

            // Set up SQL with CDATA if enabled
            if (_cDataOption.Trim().ToLower() == "*cdata")
                sSqlWork = "<![CDATA[" + sSqlWork + "]]>";

            // SQL query base XML string. This version disables committment control.
            string sXMLIN = "<?xml version='1.0'?>" + _CrLf + "<?xml-stylesheet type='text/xsl' href='/DemoXslt.xsl'?>" + _CrLf + "<script>" + _CrLf + "<sql>" + _CrLf + "<options options='noauto' commit='none' error='fast'/>" + _CrLf + "</sql>" + _CrLf + "<sql>" + _CrLf + "<connect conn='myconn' options='noauto' error='fast'/>" + _CrLf + "</sql>" + _CrLf + "<sql>" + _CrLf + "<prepare conn='myconn'>" + sSqlWork + "</prepare>" + _CrLf + "</sql>" + _CrLf + "<sql>" + _CrLf + "<execute/>" + _CrLf + "</sql>" + _CrLf + "<sql>" + _CrLf + "<describe desc='col'/>" + _CrLf + "</sql>" + _CrLf + "<sql>" + _CrLf + "<fetch block='all' desc='on'/>" + _CrLf + "</sql>" + _CrLf + "<sql>" + _CrLf + "<free/>" + _CrLf + "</sql>" + _CrLf + "</script>";

            string sXMLOUT = _iXmlResponseBufferSize.ToString();  // Buffer size

            _LastError = "";

            // Removed - Can't block action queries
            // If sSQL.ToUpper.Trim.StartsWith("INSERT") = False And sSQL.ToUpper.Trim.StartsWith("UPDATE") = False And sSQL.ToUpper.Trim.StartsWith("DELETE") = False Then
            // Throw New Exception("Only SQL INSERT, UPDATE and DELETE actions are allowed.")
            // End If

            // Replace parms into XML string
            sdb2parm = sdb2parm.Replace("@@db2value", _DB2Info);
            sdb2parm = sdb2parm.Replace("@@uidvalue", _User);
            sdb2parm = sdb2parm.Replace("@@pwdvalue", _Password);
            sdb2parm = sdb2parm.Replace("@@ipcvalue", _IPCINfo);
            sdb2parm = sdb2parm.Replace("@@ctlvalue", _IPCPersistance);
            sdb2parm = sdb2parm.Replace("@@xmlinvalue", sXMLIN);
            sdb2parm = sdb2parm.Replace("@@xmloutvalue", sXMLOUT);

            // Execute request 
            sRtnXML = ExecuteHttpPostRequest(_BaseURL, "POST", sdb2parm, _HttpTimeout, _UseHttpCredentials,_allowInvalidSslCertificates);

            // Bail out if HTTPRequest failure
            if (sRtnXML.StartsWith("ERROR"))
                throw new Exception(sRtnXML);

            // Save last response
            _LastXMLResponse = sRtnXML;

            // Check for SQL action success in XML response data
            if (sRtnXML.Contains(sSuccessValue) | sRtnXML.Contains(sSuccessValue2))
                rtnexecute = true;
            else
            {
                _LastError = "SQL error occured. Please use GetLastXmlResponse to review the last XML info to determine the cause.";
                rtnexecute = false;
            }

            return rtnexecute;
        }
        catch (Exception ex)
        {
            _LastXMLResponse = sRtnXML;
            _LastError = ex.Message;
            return false;
        }
    }

    /// <summary>
    ///  This function queries the DB2 database with selected SQL statement, returns the XML response
    ///  and then loads the internal DataTable object with the returned records.
    ///  The internal results DataTable can be accessed by and of the GetDataTable* methods.
    ///  </summary>
    ///  <param name="sSQL">SQL Select. INSERT, UPDATE and DELETE not allowed </param>
    ///  <param name="sQueryResultOutputFile">Optional PC output file for XML response data. Otherwise data set is created from memory.</param>
    ///  <returns>True- Query service call succeeded, False-Query service call failed</returns>
    ///  <remarks></remarks>
    public bool ExecuteSqlQuery(string sSQL, string sQueryResultOutputFile = "")
    {
        string sdb2parm = _Db2Parm;
        string sRtnXML = "";
        bool rtnload;
        string sSQLWork = sSQL;

        try
        {
            _LastError = "";
            sRtnXML = "";
            _LastHTTPResponse = "";
            _LastXMLResponse = "";

            // Clear data sets
            _dsXmlResponseData = null;
            _dtReturnData = null;

            // Set up SQL with CDATA if enabled
            if (_cDataOption.Trim().ToLower() == "*cdata")
                sSQLWork = "<![CDATA[" + sSQLWork + "]]>";

            // SQL query base XML string
            string sXMLIN = "<?xml version='1.0'?>" + _CrLf + "<?xml-stylesheet type='text/xsl' href='/DemoXslt.xsl'?>" + _CrLf + "<script>" + _CrLf + "<sql>" + _CrLf + "<options options='noauto' autocommit='off'/>" + _CrLf + "</sql>" + _CrLf + "<sql>" + _CrLf + "<connect conn='myconn' options='noauto'/>" + _CrLf + "</sql>" + _CrLf + "<sql>" + _CrLf + "<prepare conn='myconn'>" + sSQLWork.Trim() + "</prepare>" + _CrLf + "</sql>" + _CrLf + "<sql>" + _CrLf + "<execute/>" + _CrLf + "</sql>" + _CrLf + "<sql>" + _CrLf + "<describe desc='col'/>" + _CrLf + "</sql>" + _CrLf + "<sql>" + _CrLf + "<fetch block='all' desc='on'/>" + _CrLf + "</sql>" + _CrLf + "<sql>" + _CrLf + "<free/>" + _CrLf + "</sql>" + _CrLf + "</script>";

            string sXMLOUT = _iXmlResponseBufferSize.ToString(); // Buffer size

            if (sSQL.ToUpper().Contains("UPDATE ") | sSQL.ToUpper().Contains("DELETE ") | sSQL.ToUpper().Contains("INSERT "))
                throw new Exception("Only SQL selection queries are supported.");

            if (sSQL.ToUpper().StartsWith("UPDATE") | sSQL.ToUpper().StartsWith("DELETE") | sSQL.ToUpper().StartsWith("INSERT"))
                throw new Exception("SQL statement cannot start with INSERT, UPDATE or DELETE.");

            if (sSQL.ToUpper().StartsWith("SELECT") == false && sSQL.ToUpper().StartsWith("CALL") == false)
                throw new Exception("SQL selection must start with SELECT or CALL");

            // Replace parms into XML string
            sdb2parm = sdb2parm.Replace("@@db2value", _DB2Info);
            sdb2parm = sdb2parm.Replace("@@uidvalue", _User);
            sdb2parm = sdb2parm.Replace("@@pwdvalue", _Password);
            sdb2parm = sdb2parm.Replace("@@ipcvalue", _IPCINfo);
            sdb2parm = sdb2parm.Replace("@@ctlvalue", _IPCPersistance);
            sdb2parm = sdb2parm.Replace("@@xmlinvalue", sXMLIN);
            sdb2parm = sdb2parm.Replace("@@xmloutvalue", sXMLOUT);

            // Execute request 
            sRtnXML = ExecuteHttpPostRequest(_BaseURL, "POST", sdb2parm, _HttpTimeout, _UseHttpCredentials, _allowInvalidSslCertificates);

            // Bail out if HTTPRequest failure
            if (sRtnXML.StartsWith("ERROR"))
                throw new Exception(sRtnXML);

            // Save XML results to file and reload XML as a DataTable
            if (sQueryResultOutputFile.Trim() != "")
            {

                // Write results to output file if specified
                WriteStringToFile(sRtnXML, sQueryResultOutputFile);

                // Load DataTable from XML file
                rtnload = LoadDataTableFromXmlResponseFile(sQueryResultOutputFile);
            }
            else

                // Load DataTable from XML response
                rtnload = LoadDataTableFromXmlResponseString(sRtnXML);

            // Save last response
            _LastXMLResponse = sRtnXML;

            if (rtnload == true)
                return true;
            else
                return false;
        }
        catch (Exception ex)
        {
            _LastXMLResponse = sRtnXML;
            _LastError = ex.Message;
            return false;
        }
    }

    /// <summary>
    ///  This function queries the DB2 database with selected SQL statement and returns results to data table all in one step 
    ///  </summary>
    ///  <param name="sSQL">SQL Select. INSERT, UPDATE and DELETE not allowed </param>
    ///  <param name="sQueryResultOutputFile">Optional PC output file for XML response data. Otherwise data set is created from memory.</param>
    ///  <returns>DataTable with results of query or Nothing</returns>
    ///  <remarks></remarks>
    public DataTable ExecuteSqlQueryToDataTable(string sSQL, string sQueryResultOutputFile = "")
    {
        string sdb2parm = _Db2Parm;
        bool rtnquery;

        try
        {
            _LastError = "";
            _LastHTTPResponse = "";
            _LastXMLResponse = "";

            // Run query and load internal results DataTable
            rtnquery = ExecuteSqlQuery(sSQL, sQueryResultOutputFile);

            // Return results DataTable
            if (rtnquery)
                return GetQueryResultsDataTable();
            else
                throw new Exception("Query failed. Error: " + GetLastError());
        }
        catch (Exception ex)
        {
            _LastXMLResponse = GetLastXmlResponse();
            _LastError = ex.Message;
            return null;
        }
    }

    /// <summary>
    ///  This function queries the DB2 database with selected SQL statement and returns results to generic list all in one step.
    ///  Column names can optionally be returned in first row of generic list.
    ///  </summary>
    ///  <param name="sSQL">SQL Select. INSERT, UPDATE and DELETE not allowed </param>
    ///  <param name="sQueryResultOutputFile">Optional PC output file for XML response data. Otherwise data set is created from memory.</param>
    ///  <param name="firstRowColumnNames">Optional - Return first row as column names. False=No column names, True=Return column names. Default=False</param>
    ///  <returns>Generic List object with results of query or Nothing on error</returns>
    ///  <remarks></remarks>
    public List<List<object>> ExecuteSqlQueryToList(string sSQL, string sQueryResultOutputFile = "", bool firstRowColumnNames = false)
    {
        string sdb2parm = _Db2Parm;
        DataTable dtTemp;

        try
        {
            _LastError = "";
            _LastHTTPResponse = "";
            _LastXMLResponse = "";

            // Run query to DataTable
            dtTemp = ExecuteSqlQueryToDataTable(sSQL, sQueryResultOutputFile);

            // If no data table returned, bail out now. Error info will have already been set.
            if (dtTemp == null)
                return null;
            else
                // Export to list and return 
                return ExportDataTableToList(dtTemp, firstRowColumnNames);
        }
        catch (Exception ex)
        {
            _LastXMLResponse = GetLastXmlResponse();
            _LastError = ex.Message;
            return null;
        }
    }

    /// <summary>
    ///  This function queries the DB2 database with selected SQL statement and returns results to XML dataset stream all in one step 
    ///  </summary>
    ///  <param name="sSQL">SQL Select. INSERT, UPDATE and DELETE not allowed </param>
    ///  <param name="sOutputFile">Optional PC output file for XML response data. Otherwise data set is created from memory.</param>
    ///  <returns>XML string</returns>
    ///  <remarks></remarks>
    public string ExecuteSqlQueryToXmlString(string sSQL, string sOutputFile = "", string sTableName = "Table1", bool bWriteSchema = false)
    {
        string sdb2parm = _Db2Parm;
        string sRtnXML = "";
        bool rtnquery;
        DataTable _dt;

        try
        {
            _LastError = "";
            sRtnXML = "";
            _LastHTTPResponse = "";
            _LastXMLResponse = "";

            // if table not set, default to Table1
            if (sTableName.Trim() == "")
                sTableName = "Table1";

            // Run query and load internal results DataTable
            rtnquery = ExecuteSqlQuery(sSQL, sOutputFile);

            // Export DataTable results to XML
            if (rtnquery)
            {
                StringBuilder SB = new StringBuilder();
                System.IO.StringWriter SW = new System.IO.StringWriter(SB);
                _dt = GetQueryResultsDataTable();
                _dt.TableName = sTableName;
                // Write XMl with or without schema info
                if (bWriteSchema)
                    _dt.WriteXml(SW, System.Data.XmlWriteMode.WriteSchema);
                else
                    _dt.WriteXml(SW);
                sRtnXML = SW.ToString();
                SW.Close();
                return sRtnXML;
            }
            else
                throw new Exception("Query failed. Error: " + GetLastError());
        }
        catch (Exception ex)
        {
            _LastXMLResponse = GetLastXmlResponse();
            _LastError = ex.Message;
            return "";
        }
    }

    /// <summary>
    ///  This function queries the DB2 database with selected SQL statement and returns results to XML file all in one step 
    ///  </summary>
    ///  <param name="sSQL">SQL Select. INSERT, UPDATE and DELETE not allowed </param>
    ///  <param name="sOutputFile">Optional PC output file for XML response data. Otherwise data set is created from memory.</param>
    ///  <returns>True - Query service call succeeded, False - Query service call failed</returns>
    ///  <remarks></remarks>
    public bool ExecuteSqlQueryToXmlFile(string sSQL, string sXmlOutputFile, bool replace = false, string sOutputFile = "", string sTableName = "Table1", bool bWriteSchema = false)
    {
        string sdb2parm = _Db2Parm;
        bool rtnquery;

        try
        {
            _LastError = "";
            _LastHTTPResponse = "";
            _LastXMLResponse = "";

            // Run query and load internal results DataTable
            rtnquery = ExecuteSqlQuery(sSQL, sOutputFile);

            // Export results to XML file
            if (rtnquery)
                return GetQueryResultsDataTableToXmlFile(sXmlOutputFile, sTableName, bWriteSchema, replace);
            else
                throw new Exception("Query failed. Error: " + GetLastError());
        }
        catch (Exception ex)
        {
            _LastXMLResponse = GetLastXmlResponse();
            _LastError = ex.Message;
            return false;
        }
    }

    /// <summary>
    ///  This function queries the DB2 database with selected SQL statement and returns results to Csv string in one step
    ///  </summary>
    ///  <param name="sSQL">SQL Select. INSERT, UPDATE and DELETE not allowed </param>
    ///  <param name="sQueryResultOutputFile">Optional PC output file for XML response data. Otherwise data set is created from memory.</param>
    ///  <param name="sFieldSepchar">Field delimiter/separator. Default = Comma</param>
    ///  <param name="sFieldDataDelimChar">Field data delimiter character. Default = double quotes.</param>
    ///  <returns>CSV string</returns>
    public string ExecuteSqlQueryToCsvString(string sSQL, string sQueryResultOutputFile = "", string sFieldSepchar = ",", string sFieldDataDelimChar = "\"")
    {
        string sdb2parm = _Db2Parm;
        bool rtnquery;

        try
        {
            _LastError = "";
            _LastHTTPResponse = "";
            _LastXMLResponse = "";

            // Run query and load internal results DataTable
            rtnquery = ExecuteSqlQuery(sSQL, sQueryResultOutputFile);

            // Export results to CSV string
            if (rtnquery)
                return GetQueryResultsDataTableToCsvString(sFieldSepchar, sFieldDataDelimChar);
            else
                throw new Exception("Query failed. Error: " + GetLastError());
        }
        catch (Exception ex)
        {
            _LastXMLResponse = GetLastXmlResponse();
            _LastError = ex.Message;
            return "";
        }
    }

    /// <summary>
    ///  This function queries the DB2 database with selected SQL statement and returns results to Csv file in one step
    ///  </summary>
    ///  <param name="sSQL">SQL Select. INSERT, UPDATE and DELETE not allowed </param>
    ///  <param name="sCsvOutputFile">Output CSV file</param>
    ///  <param name="replace">Replace output file True=Replace file,False=Do not replace</param>
    ///  <param name="sQueryResultOutputFile">Optional PC output file for XML response data. Otherwise data set is created from memory.</param>
    ///  <param name="sFieldSepchar">Field delimiter/separator. Default = Comma</param>
    ///  <param name="sFieldDataDelimChar">Field data delimiter character. Default = double quotes.</param>
    ///  <returns>True-Query service call succeeded, False-Query service call failed</returns>
    public bool ExecuteSqlQueryToCsvFile(string sSQL, string sCsvOutputFile, bool replace = false, string sQueryResultOutputFile = "", string sFieldSepchar = ",", string sFieldDataDelimChar = "\"")
    {
        string sdb2parm = _Db2Parm;
        bool rtnquery;

        try
        {
            _LastError = "";
            _LastHTTPResponse = "";
            _LastXMLResponse = "";

            // Run query and load internal results DataTable
            rtnquery = ExecuteSqlQuery(sSQL, sQueryResultOutputFile);

            // Export results to CSV file
            if (rtnquery)
                return GetQueryResultsDataTableToCsvFile(sCsvOutputFile, sFieldSepchar, sFieldDataDelimChar, replace);
            else
                throw new Exception("Query failed. Error: " + GetLastError());
        }
        catch (Exception ex)
        {
            _LastXMLResponse = GetLastXmlResponse();
            _LastError = ex.Message;
            return false;
        }
    }


    /// <summary>
    ///  This function queries the DB2 database with selected SQL statement and returns results to JSON string in one step
    ///  </summary>
    ///  <param name="sSQL">SQL Select. INSERT, UPDATE and DELETE not allowed </param>
    ///  <param name="sQueryResultOutputFile">Optional PC output file for XML response data. Otherwise data set is created from memory.</param>
    ///  <returns>JOSN string</returns>
    public string ExecuteSqlQueryToJsonString(string sSQL, string sQueryResultOutputFile = "")
    {
        string sdb2parm = _Db2Parm;
        bool rtnquery;

        try
        {
            _LastError = "";
            _LastHTTPResponse = "";
            _LastXMLResponse = "";

            // Run query and load internal results DataTable
            rtnquery = ExecuteSqlQuery(sSQL, sQueryResultOutputFile);

            // Export results to CSV string
            if (rtnquery)
                return GetQueryResultsDataTableToJsonString();
            else
                throw new Exception("Query failed. Error: " + GetLastError());
        }
        catch (Exception ex)
        {
            _LastXMLResponse = GetLastXmlResponse();
            _LastError = ex.Message;
            return "";
        }
    }

    /// <summary>
    ///  This function queries the DB2 database with selected SQL statement and returns results to JSON file in one step
    ///  </summary>
    ///  <param name="sSQL">SQL Select. INSERT, UPDATE and DELETE not allowed </param>
    ///  <param name="sJsonOutputFile">Output JSON file</param>
    ///  <param name="replace">Replace output file True=Replace file,False=Do not replace</param>
    ///  <param name="sQueryResultOutputFile">Optional PC output file for XML response data. Otherwise data set is created from memory.</param>
    ///  <returns>True-Query service call succeeded, False-Query service call failed</returns>
    public bool ExecuteSqlQueryToJsonFile(string sSQL, string sJsonOutputFile, bool replace = false, string sQueryResultOutputFile = "")
    {
        string sdb2parm = _Db2Parm;
        bool rtnquery;

        try
        {
            _LastError = "";
            _LastHTTPResponse = "";
            _LastXMLResponse = "";

            // Run query and load internal results DataTable
            rtnquery = ExecuteSqlQuery(sSQL, sQueryResultOutputFile);

            // Export results to JSON file
            if (rtnquery)
                return GetQueryResultsDataTableToJsonFile(sJsonOutputFile, replace);
            else
                throw new Exception("Query failed. Error: " + GetLastError());
        }
        catch (Exception ex)
        {
            _LastXMLResponse = GetLastXmlResponse();
            _LastError = ex.Message;
            return false;
        }
    }

    /// <summary>
    ///  This function runs the specified IBM i CL command line. The CL command can be a regular program call or a SBMJOB type of command to submit a job.
    ///  </summary>
    ///  <param name="sCommandString">CL command line to execute</param>
    ///  <returns>True - Command call succeeded, False - Command call failed</returns>
    ///  <remarks></remarks>
    public bool ExecuteCommand(string sCommandString)
    {
        string sdb2parm = _Db2Parm;
        string sRtnXML = "";

        // Clear data set
        _dsCommandResponse = null;
        _dtCommandResponse = null;

        try
        {
            _LastError = "";
            sRtnXML = "";
            _LastHTTPResponse = "";
            _LastXMLResponse = "";

            // Command call base XML string 
            string sXMLIN = "<?xml version='1.0'?>" + _CrLf + "<?xml-stylesheet type='text/xsl' href='/DemoXslt.xsl'?>" + _CrLf + "<script>" + _CrLf + "<cmd>" + sCommandString + "</cmd>" + _CrLf + "</script>";

            string sXMLOUT = _iXmlResponseBufferSize.ToString();  // "32768" 'Original Buffer size

            // Replace parms into XML string
            sdb2parm = sdb2parm.Replace("@@db2value", _DB2Info);
            sdb2parm = sdb2parm.Replace("@@uidvalue", _User);
            sdb2parm = sdb2parm.Replace("@@pwdvalue", _Password);
            sdb2parm = sdb2parm.Replace("@@ipcvalue", _IPCINfo);
            sdb2parm = sdb2parm.Replace("@@ctlvalue", _IPCPersistance);
            sdb2parm = sdb2parm.Replace("@@xmlinvalue", sXMLIN);
            sdb2parm = sdb2parm.Replace("@@xmloutvalue", sXMLOUT);

            // Execute request
            sRtnXML = ExecuteHttpPostRequest(_BaseURL, "POST", sdb2parm, _HttpTimeout, _UseHttpCredentials, _allowInvalidSslCertificates);

            // Bail out if +++ success not returned in XML response
            if (sRtnXML.Contains("+++ success"))
            {
                _LastXMLResponse = sRtnXML;

                // Load data table with program response parms
                if (LoadCmdCallDataTableFromXmlResponseString(GetLastXmlResponse()))
                    return true;
                else
                {
                    throw new Exception("Errors loading command call response XML string.");
                }
            }
            else
            {
                _LastXMLResponse = sRtnXML;

                // Load data table with program response parms
                if (LoadCmdCallDataTableFromXmlResponseString(GetLastXmlResponse()))
                    return false; // Actual command call failed. Return false
                else
                {
                    throw new Exception("Errors loading command call response XML string.");
                }

                //throw new Exception(sRtnXML);
            }
        }
        catch (Exception ex)
        {
            _LastXMLResponse = sRtnXML;
            _LastError = ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  This function runs the specified IBM i program call. 
    ///  You will need to pass in an array list of parameters.
    ///  </summary>
    ///  <param name="sProgram">IBM i program name</param>
    ///  <param name="sLibrary">IBM i program library</param>
    ///  <param name="aParmList">IBM i program parm list array</param>
    ///  <returns>True-Program call succeeded, False-Program call failed</returns>
    ///  <remarks></remarks>
    public bool ExecuteProgram(string sProgram, string sLibrary, ArrayList aParmList)
    {
        string sdb2parm = _Db2Parm;
        string sRtnXML = "";

        try
        {
            _LastError = "";
            sRtnXML = "";
            _LastHTTPResponse = "";
            _LastXMLResponse = "";

            // Clear data set
            _dsProgramResponse = null;
            _dtProgramResponse = null;

            // Program call base XML string 
            string sXMLIN = "<?xml version='1.0'?>" + _CrLf + "<pgm name='@@pgmname' lib='@@pgmlibrary'>" + _CrLf;
            foreach (PgmParmList pm in aParmList)
                sXMLIN = sXMLIN + "<parm><data type='" + pm.parmtype + "'>" + pm.parmvalue + "</data></parm>" + _CrLf;
            sXMLIN = sXMLIN + "</pgm>" + _CrLf;

            string sXMLOUT = _iXmlResponseBufferSize.ToString(); // "32768" 'Buffer size

            // Replace core program info in program string
            sXMLIN = sXMLIN.Replace("@@pgmname", sProgram);
            sXMLIN = sXMLIN.Replace("@@pgmlibrary", sLibrary);

            // Replace parms into XML string
            sdb2parm = sdb2parm.Replace("@@db2value", _DB2Info);
            sdb2parm = sdb2parm.Replace("@@uidvalue", _User);
            sdb2parm = sdb2parm.Replace("@@pwdvalue", _Password);
            sdb2parm = sdb2parm.Replace("@@ipcvalue", _IPCINfo);
            sdb2parm = sdb2parm.Replace("@@ctlvalue", _IPCPersistance);
            sdb2parm = sdb2parm.Replace("@@xmlinvalue", sXMLIN);
            sdb2parm = sdb2parm.Replace("@@xmloutvalue", sXMLOUT);

            // Execute request
            sRtnXML = ExecuteHttpPostRequest(_BaseURL, "POST", sdb2parm, _HttpTimeout, _UseHttpCredentials, _allowInvalidSslCertificates);

            // Bail out if +++ success not returned in XML response
            if (sRtnXML.Contains("+++ success"))
            {
                _LastXMLResponse = sRtnXML;

                // Load data table with program response parms
                if (LoadPgmCallDataTableFromXmlResponseString(GetLastXmlResponse()))
                    return true;
                else
                    throw new Exception("Errors loading program call response XML string.");
            }
            else
                throw new Exception(sRtnXML);
        }
        catch (Exception ex)
        {
            _LastXMLResponse = sRtnXML;
            _LastError = ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  This function runs the specified IBM i service program subprocedure call. 
    ///  http://www.statususer.org/pdf/20130212Presentation.pdf
    ///  http://174.79.32.155/wiki/pmwiki.php/XMLSERVICE/XMLSERVICESamples
    ///   <pgm name='ZZSRV' lib='XMLSERVICE' func='ZZTIMEUSA'>
    ///  <parm io='both'>
    ///  <data type='8A'>09:45 AM</data>
    ///  </parm>
    ///  <return>
    ///  <data type='8A'>nada</data>
    ///  </return>
    ///  </pgm>
    ///  </summary>
    ///  <param name="sServiceProgram">IBM i service program name</param>
    ///  <param name="sLibrary">IBM i program library</param>
    ///  <param name="sProcedure">IBM i subprocedure</param>
    ///  <param name="aParmList">IBM i program parm list array</param>
    ///  <param name="rtnParmList">IBM i program return parm list array</param>
    ///  <returns>True-Service program call succeeded, False-Service program failed</returns>
    ///  <remarks></remarks>
    public bool ExecuteProgramProcedure(string sServiceProgram, string sLibrary, string sProcedure, ArrayList aParmList, ArrayList rtnParmList)
    {
        string sdb2parm = _Db2Parm;
        string sRtnXML = "";

        try
        {
            _LastError = "";
            sRtnXML = "";
            _LastHTTPResponse = "";
            _LastXMLResponse = "";

            // Clear data set
            _dsProgramResponse = null;
            _dtProgramResponse = null;

            // Program call base XML string 
            string sXMLIN = "<?xml version='1.0'?>" + _CrLf + "<script>" + "<pgm name='@@pgmname' lib='@@pgmlibrary' func='@@pgmfunction'>" + _CrLf;

            foreach (ProcedureParmList pm in aParmList)
                sXMLIN = sXMLIN + "<parm io='input'><data type='" + pm.parmtype.Trim() + "'>" + pm.parmvalue + "</data></parm>" + _CrLf;
            // If return parm list passed, set it up
            if (rtnParmList.Count > 0)
            {
                sXMLIN = sXMLIN + "<return>" + _CrLf;
                foreach (ProcedureParmList pm in rtnParmList)
                    sXMLIN = sXMLIN + "<data type='" + pm.parmtype.Trim() + "'>" + pm.parmvalue + "</data>" + _CrLf;
                sXMLIN = sXMLIN + "</return>";
            }
            sXMLIN = sXMLIN + "</pgm>" + _CrLf;
            sXMLIN = sXMLIN + "</script>" + _CrLf;

            string sXMLOUT = _iXmlResponseBufferSize.ToString(); // "32768" 'Buffer size

            // Replace core program info in program string
            sXMLIN = sXMLIN.Replace("@@pgmname", sServiceProgram);
            sXMLIN = sXMLIN.Replace("@@pgmlibrary", sLibrary);
            sXMLIN = sXMLIN.Replace("@@pgmfunction", sProcedure);

            // ' ''TEST subprocedure call
            // sXMLIN = "<?xml version='1.0'?>" & vbCrLf & _
            // "<script>"
            // 'sXMLIN = ""
            /// sXMLIN = "<cmd comment='addlible'>ADDLIBLE LIB(XMLDOTNETI) POSITION(*FIRST)</cmd>" & vbCrLf
            // sXMLIN = sXMLIN & "<pgm name='XMLSVC001' lib='XMLDOTNETI' func='SAMPLE1'>" & vbCrLf
            // sXMLIN = sXMLIN & "<parm io='input'>" & vbCrLf
            // 'sXMLIN = sXMLIN & "<data type='65535A' varying='on'>Test</data>" & vbCrLf
            // sXMLIN = sXMLIN & "<data type='200A'>Test</data>" & vbCrLf
            // sXMLIN = sXMLIN & "</parm>" & vbCrLf
            // sXMLIN = sXMLIN & "<return>" & vbCrLf
            // sXMLIN = sXMLIN & "<data type='200A'></data>" & vbCrLf
            // sXMLIN = sXMLIN & "</return>" & vbCrLf
            // sXMLIN = sXMLIN & "</pgm>" & vbCrLf
            // sXMLIN = sXMLIN & "</script>" & vbCrLf

            // Replace parms into XML string
            sdb2parm = sdb2parm.Replace("@@db2value", _DB2Info);
            sdb2parm = sdb2parm.Replace("@@uidvalue", _User);
            sdb2parm = sdb2parm.Replace("@@pwdvalue", _Password);
            sdb2parm = sdb2parm.Replace("@@ipcvalue", _IPCINfo);
            sdb2parm = sdb2parm.Replace("@@ctlvalue", _IPCPersistance);
            sdb2parm = sdb2parm.Replace("@@xmlinvalue", sXMLIN);
            sdb2parm = sdb2parm.Replace("@@xmloutvalue", sXMLOUT);

            // Execute request
            sRtnXML = ExecuteHttpPostRequest(_BaseURL, "POST", sdb2parm, _HttpTimeout, _UseHttpCredentials, _allowInvalidSslCertificates);

            // Bail out if +++ success not returned in XML response
            if (sRtnXML.Contains("+++ success"))
            {
                _LastXMLResponse = sRtnXML;

                // Load data table with program response parms
                if (LoadPgmCallDataTableFromXmlResponseString(GetLastXmlResponse()))
                    return true;
                else
                {
                    throw new Exception("Errors loading program call response XML string.");
                }
            }
            else
            {
                throw new Exception(sRtnXML);
            }

        }
        catch (Exception ex)
        {
            _LastXMLResponse = sRtnXML;
            _LastError = ex.Message;
            return false;
        }
    }



    /// <summary>
    ///  Send data to sequential data queue. Data queue must already exist.
    ///  </summary>
    ///  <param name="sDataQueue">Data queue name</param>
    ///  <param name="sDataQueueLibrary">Data queue library</param>
    ///  <param name="sData">String data to send</param>
    ///  <param name="iDataLength">Length of data to send. Default=0 which will auto-calculate length.</param>
    ///  <returns></returns>
    ///  <remarks></remarks>
    public bool SendDataQueueSequential(string sDataQueue, string sDataQueueLibrary, string sData, int iDataLength = 0)
    {
        bool rtnbool;

        try
        {

            // Execute the program with parms
            ArrayList arrParms = new ArrayList();
            XmlServicei.PgmParmList pDqName = new XmlServicei.PgmParmList();
            XmlServicei.PgmParmList pDqLib = new XmlServicei.PgmParmList();
            XmlServicei.PgmParmList pLength = new XmlServicei.PgmParmList();
            XmlServicei.PgmParmList pData = new XmlServicei.PgmParmList();
            //Following parm not currently used
            //XmlServicei.PgmParmList pWaitTime = new XmlServicei.PgmParmList();

            // Reset last data queue operation info
            _LastError = "";
            _LastDataQueueName = sDataQueue;
            _LastDataQueueLibrary = sDataQueueLibrary;
            _LastDataQueueLengthReceived = iDataLength;
            _LastDataQueueDataReceived = "";
            _LastDataQueueDataSent = "";

            // Set data queue
            pDqName.parmtype = "10A";
            pDqName.parmvalue = sDataQueue;
            // Set data queue library
            pDqLib.parmtype = "10A";
            pDqLib.parmvalue = sDataQueueLibrary;
            // Set length if passed
            pLength.parmtype = "5p0";
            if (iDataLength > 0)
                pLength.parmvalue = iDataLength.ToString();
            else
                pLength.parmvalue = sData.TrimEnd().Length.ToString();
            // Set data to be passed
            pData.parmtype = sData.TrimEnd().Length + "A";
            pData.parmvalue = sData.TrimEnd();
            // Add parms to parm array
            arrParms.Add(pDqName);
            arrParms.Add(pDqLib);
            arrParms.Add(pLength);
            arrParms.Add(pData);

            // Call program to send the data queue entry
            rtnbool = ExecuteProgram("QSNDDTAQ", "QSYS", arrParms);

            if (rtnbool == true)
            {
                _LastError = "Send to data queue was Successful. XML Response: " + _CrLf + GetLastXmlResponse();

                // Set last data queue operation info
                _LastDataQueueDataSent = pData.parmvalue;

                return true;
            }
            else
            {
                _LastError = "Send to data queue failed. XML Response: " + _CrLf + GetLastXmlResponse();
                return false;
            }
        }
        catch (Exception ex)
        {
            _LastError = "Error occurred sending to data queue. Error: " + ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  Receive data from sequential data queue. Data queue must already exist.
    ///  </summary>
    ///  <param name="sDataQueue">Data queue name</param>
    ///  <param name="sDataQueueLibrary">Data queue library</param>
    ///  <param name="iDataLength">Length of data to send. Default=0 which will auto-calculate length.</param>
    ///  <param name="iWaitTime">Length of time to wait for a response. Default=10 seconds.</param>
    ///  <returns>Returns data or string starting with"EXCEPTION" if errors occurred</returns>
    ///  <remarks></remarks>
    public string ReceiveDataQueueSequential(string sDataQueue, string sDataQueueLibrary, int iDataLength, int iWaitTime = 10)
    {
        bool rtnbool;
        string sData = "";

        try
        {

            // Execute the program with parms
            ArrayList arrParms = new ArrayList();
            XmlServicei.PgmParmList pDqName = new XmlServicei.PgmParmList();
            XmlServicei.PgmParmList pDqLib = new XmlServicei.PgmParmList();
            XmlServicei.PgmParmList pLength = new XmlServicei.PgmParmList();
            XmlServicei.PgmParmList pData = new XmlServicei.PgmParmList();
            XmlServicei.PgmParmList pWaitTime = new XmlServicei.PgmParmList();

            // Reset last data queue operation info
            _LastError = "";
            _LastDataQueueName = sDataQueue;
            _LastDataQueueLibrary = sDataQueueLibrary;
            _LastDataQueueLengthReceived = iDataLength;
            _LastDataQueueDataReceived = "";
            _LastDataQueueDataSent = "";

            pDqName.parmtype = "10A";
            pDqName.parmvalue = sDataQueue;
            pDqLib.parmtype = "10A";
            pDqLib.parmvalue = sDataQueueLibrary;
            pLength.parmtype = "5p0";
            pLength.parmvalue = iDataLength.ToString();
            pData.parmtype = iDataLength + "A";
            pData.parmvalue = sData.TrimEnd();
            pWaitTime.parmtype = "5p0";
            pWaitTime.parmvalue = iWaitTime.ToString();
            arrParms.Add(pDqName);
            arrParms.Add(pDqLib);
            arrParms.Add(pLength);
            arrParms.Add(pData);
            arrParms.Add(pWaitTime);

            // Call program to receive the data queue entry
            rtnbool = ExecuteProgram("QRCVDTAQ", "QSYS", arrParms);

            if (rtnbool == true)
                _LastError = "Receive from data queue was Successful. XML Response: " + _CrLf + GetLastXmlResponse();
            else
                _LastError = "Receive from data queue failed. XML Response: " + _CrLf + GetLastXmlResponse();

            // Get return parms to datatable and extract return data
            if (LoadPgmCallDataTableFromXmlResponseString(GetLastXmlResponse()))
            {

                // Make sure data returned
                if (_dtProgramResponse.Rows.Count != 5)
                    throw new Exception("It appears no parms were returned from receive data queue call.");

                DataRow _drParm = null;

                // Parm2 data row contains length of returned data
                _drParm = _dtProgramResponse.Rows[2];
                _LastDataQueueLengthReceived = Convert.ToInt32(_drParm[1]);

                // Parm3 data row contains returned data
                _drParm = _dtProgramResponse.Rows[3];

                // Return data queue return data value
                if (_drParm[0].ToString() == "Parm3")
                {

                    // Set last data queue operation info
                    _LastDataQueueName = sDataQueue;
                    _LastDataQueueLibrary = sDataQueueLibrary;
                    _LastDataQueueDataReceived = _drParm[1].ToString();
                    return _drParm[1].ToString();
                }
                else
                    throw new Exception("It appears that Parm3 - data queue value is missing.");
            }
            else
                throw new Exception("Errors loading program call response XML string.");
        }
        catch (Exception ex)
        {
            _LastError = "Error occurred receiving from data queue. Error: " + ex.Message;
            return "EXCEPTION " + _LastError;
        }
    }
    /// <summary>
    ///  Get last data queue name
    ///  </summary>
    ///  <returns>Last data queue name used</returns>
    ///  <remarks></remarks>
    public string GetLastDataQueueName()
    {
        try
        {
            return _LastDataQueueName;
        }
        catch (Exception)
        {
            return "";
        }
    }
    /// <summary>
    ///  Get last data queue library
    ///  </summary>
    ///  <returns>Last data queue library used</returns>
    ///  <remarks></remarks>
    public string GetLastDataQueueLibrary()
    {
        try
        {
            return _LastDataQueueLibrary;
        }
        catch (Exception)
        {
            return "";
        }
    }
    /// <summary>
    ///  Get last data buffer sent to data queue
    ///  </summary>
    ///  <returns>Last data buffer sent to data queue.</returns>
    ///  <remarks></remarks>
    public string GetLastDataQueueDataSent()
    {
        try
        {
            return _LastDataQueueDataSent;
        }
        catch (Exception)
        {
            return "";
        }
    }
    /// <summary>
    ///  Get last data buffer received from data queue
    ///  </summary>
    ///  <returns></returns>
    ///  <remarks>Last data buffer received from data queue.</remarks>
    public string GetLastDataQueueDataReceived()
    {
        try
        {
            return _LastDataQueueDataReceived;
        }
        catch (Exception)
        {
            return "";
        }
    }
    /// <summary>
    ///  Get last data queue length received
    ///  </summary>
    ///  <returns>Length of data returned or 0 if none.</returns>
    ///  <remarks></remarks>
    public int GetLastDataQueueLengthReceived()
    {
        try
        {
            return _LastDataQueueLengthReceived;
        }
        catch (Exception)
        {
            return 0;
        }
    }



    /// <summary>
    ///  This kills the currently active IPC service instance on IBM i system.
    ///  These jobs are all marked as XTOOLKIT and can be seen by running 
    ///  the following CL command: WRKACTJOB JOB(XTOOLKIT)
    ///  Note: If you will be doing a lot of work, you can leave the job instantiated. Otherwise kill the XTOOLKIT job 
    ///  if you're done using it. 
    ///  </summary>
    ///  <returns>True - kill succeeded. Return XML will always be blank so we will always return true most likely.</returns>
    ///  <remarks></remarks>
    public bool KillService()
    {
        string sdb2parm = _Db2Parm;
        string sRtnXML = "";

        try
        {

            // Kill service call base XML string 
            string sXMLIN = "<?xml version='1.0'?>" + _CrLf + "<?xml-stylesheet type='text/xsl' href='/DemoXslt.xsl'?>" + _CrLf + "<script>" + _CrLf + "</script>";

            string sXMLOUT = _iXmlResponseBufferSize.ToString();

            _LastError = "";

            sdb2parm = sdb2parm.Replace("@@db2value", _DB2Info);
            sdb2parm = sdb2parm.Replace("@@uidvalue", _User);
            sdb2parm = sdb2parm.Replace("@@pwdvalue", _Password);
            sdb2parm = sdb2parm.Replace("@@ipcvalue", _IPCINfo);
            sdb2parm = sdb2parm.Replace("@@ctlvalue", "*immed");
            sdb2parm = sdb2parm.Replace("@@xmlinvalue", sXMLIN);
            sdb2parm = sdb2parm.Replace("@@xmloutvalue", sXMLOUT);
            sdb2parm = sdb2parm + "&submit=*immed end (kill job)";

            // Execute request
            sRtnXML = ExecuteHttpPostRequest(_BaseURL, "POST", sdb2parm, _HttpTimeout, _UseHttpCredentials, _allowInvalidSslCertificates);

            // Bail out if HTTPRequest failure
            if (sRtnXML.StartsWith("ERROR"))
                throw new Exception(sRtnXML);

            // Execute XMLSERVICE POST request to run command
            sRtnXML = ExecuteHttpPostRequest(_BaseURL, "POST", sdb2parm, _HttpTimeout, _UseHttpCredentials, _allowInvalidSslCertificates);

            // Bail out if +++ success not returned in XML response
            if (sRtnXML == "")
            {
                _LastXMLResponse = sRtnXML;
                return true;
            }
            else
                throw new Exception(sRtnXML);
        }
        catch (Exception ex)
        {
            _LastXMLResponse = sRtnXML;
            _LastError = ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  This kills the temporary IPC directories in IFS directory /tmp or other specified location
    ///  It's a good idea to periodically clean up these directories either via this function or
    ///  via a job on IBM i that does a RMDIR.  Ex: RMDIR  '/tmp/xmlservicei*'
    ///  </summary>
    ///  <param name="sIfsDirPath">IFS directory path to remove. Wildcards are OK.  Ex: /tmp/xmlservicei*'</param>
    ///  <returns>True - IFS directory kill succeeded. False - IFS directory kill failed.</returns>
    ///  <remarks></remarks>
    public bool KillIpcIfsDirectories(string sIfsDirPath)
    {
        bool rtn = false;

        try
        {
            _LastXMLResponse = "";
            _LastError = "";
            _LastHTTPResponse = "";

            // Erase directories if path passed
            if (sIfsDirPath.Trim() != "")
                rtn = ExecuteCommand("RMDIR DIR('" + sIfsDirPath.Trim() + "')");
            else
                throw new Exception("No IFS directory path specified.");

            return rtn;
        }
        catch (Exception ex)
        {
            // _LastXMLResponse = ' This value should already be set
            _LastError = ex.Message;
            return false;
        }
    }



    /// <summary>
    ///  Write text string to file
    ///  </summary>
    ///  <param name="sTextString">Text string value</param>
    ///  <param name="sOutputFile">Output file</param>
    ///  <param name="bAppend">True=Append</param>
    ///  <param name="bReplace">True=Replace file before writing</param>
    ///  <returns>True-Success, False-Failure to write</returns>
    ///  <remarks></remarks>
    public bool WriteStringToFile(string sTextString, string sOutputFile, bool bAppend = false, bool bReplace = true)
    {
        // Write text string to output file
        try
        {
            _LastError = "";

            if (System.IO.File.Exists(sOutputFile) == true)
            {
                if (bReplace == true)
                    System.IO.File.Delete(sOutputFile);
            }

            using (System.IO.StreamWriter oWriter = new System.IO.StreamWriter(sOutputFile, bAppend))
            {
                oWriter.Write(sTextString);
                oWriter.Flush();
                oWriter.Close();
            }

            return true;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  Read text file contents and return as string
    ///  </summary>
    ///  <param name="sInputFile">Input file name</param>
    ///  <returns>Contents of file as string or blanks on error</returns>
    ///  <remarks></remarks>
    public string ReadTextFile(string sInputFile)
    {
        try
        {
            _LastError = "";

            if (System.IO.File.Exists(sInputFile) == false)
                throw new Exception(sInputFile + " does not exist.");

            using (System.IO.StreamReader oReader = new System.IO.StreamReader(sInputFile, true))
            {
                string sWork = "";

                // Read all text
                sWork = oReader.ReadToEnd();

                oReader.Close();

                return sWork;
            }
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return "";
        }
    }


    /// <summary>
    ///  Create library for temp files. Library name is TMP
    ///  </summary>
    ///  <returns>True-Library created, False-Library exists or library not created.</returns>
    public bool CreateTempLibrary()
    {
        try
        {
            if (CheckObjectExists("TMP", "QSYS", "*LIB") == false)
                return ExecuteCommand("CRTLIB LIB(TMP) TYPE(*PROD) TEXT('Temp Files') AUT(*ALL) CRTAUT(*ALL)");
            else
                throw new Exception("Library TMP already exists.");
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  Create IBM i library if it does not exist
    ///  </summary> 
    ///  <param name="sLibraryName">Library name</param>
    ///  <param name="sLibraryType">Library type. Default - *PROD</param>
    ///  <param name="sLibraryText">Library text description</param>
    ///  <param name="sLibraryAut">Library authority. Default - *ALL</param>
    ///  <param name="sLibraryCrtAut">Create authority. Default - *ALL</param>
    ///  <returns>True-Library created, False-Library exists or library not created.</returns>
    public bool CreateLibrary(string sLibraryName, string sLibraryType = "*PROD", string sLibraryText = "", string sLibraryAut = "*ALL", string sLibraryCrtAut = "*ALL")
    {
        try
        {
            string sCmd = string.Format("CRTLIB LIB({0}) TYPE({1}) TEXT('{2}') AUT({3}) CRTAUT({4})", sLibraryName.Trim().ToUpper(), sLibraryType.Trim().ToUpper(), sLibraryText.Trim(), sLibraryAut.Trim().ToUpper(), sLibraryCrtAut.Trim().ToUpper());

            if (CheckObjectExists(sLibraryName.Trim().ToUpper(), "QSYS", "*LIB") == false)
                return ExecuteCommand(sCmd);
            else
                throw new Exception(string.Format("Library {0} already exists.", sLibraryName.Trim().ToUpper()));
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  Check for IBM i object existence
    ///  </summary>
    ///  <param name="sObjectName">Object name</param>
    ///  <param name="sObjectLibrary">Object library name</param>
    ///  <param name="sObjectType">Object type.</param>
    ///  <returns>True-Object exists,False-Object does not exist.</returns>
    public bool CheckObjectExists(string sObjectName, string sObjectLibrary, string sObjectType)
    {
        try
        {
            if (sObjectName.Trim() == "" | sObjectLibrary.Trim() == "" | sObjectType.Trim() == "")
                throw new Exception("Object name, library and type are all required fields.");

            string sCmd = string.Format("CHKOBJ OBJ({0}/{1}) OBJTYPE({2}) MBR(*NONE) AUT(*NONE)", sObjectLibrary.Trim().ToUpper(), sObjectName.Trim().ToUpper(), sObjectType.Trim().ToUpper());
            return ExecuteCommand(sCmd);
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  Delete IBM i library if it exists
    ///  </summary> 
    ///  <param name="sLibraryName">Library name</param>
    ///  <returns>True-Library deleted, False-Library does not exists or library not deleted.</returns>
    public bool DeleteLibrary(string sLibraryName)
    {
        try
        {
            string sCmd = string.Format("DLTLIB LIB({0})", sLibraryName.Trim().ToUpper());

            if (CheckObjectExists(sLibraryName.Trim().ToUpper(), "QSYS", "*LIB") == true)
                return ExecuteCommand(sCmd);
            else
                throw new Exception(string.Format("Library {0} does not exist.", sLibraryName.Trim().ToUpper()));
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  Clear IBM i library if it exists
    ///  </summary> 
    ///  <param name="sLibraryName">Library name</param>
    ///  <returns>True-Library cleared, False-Library does not exists or library not cleared.</returns>
    public bool ClearLibrary(string sLibraryName)
    {
        try
        {
            string sCmd = string.Format("CLRLIB LIB({0})", sLibraryName.Trim().ToUpper());

            if (CheckObjectExists(sLibraryName.Trim().ToUpper(), "QSYS", "*LIB") == true)
                return ExecuteCommand(sCmd);
            else
                throw new Exception(string.Format("Library {0} does not exist.", sLibraryName.Trim().ToUpper()));
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  Create fixed length flat physical file
    ///  </summary> 
    ///  <param name="sFileName">File name</param>
    ///  <param name="sFileLibrary">File library</param>
    ///  <param name="sLibraryTextDescription">File description</param>
    ///  <param name="iRecordLength">Record length. Default=400</param>
    ///  <returns>True-Success, False-Error</returns>
    public bool CreatePhysicalFileFixed(string sFileName, string sFileLibrary, string sLibraryTextDescription = "", int iRecordLength = 400)
    {
        try
        {
            string sCmd = string.Format(" CRTPF FILE({0}/{1}) RCDLEN({2}) TEXT('{3}') OPTION(*NOSRC *NOLIST) MAXMBRS(*NOMAX) SIZE(*NOMAX)", sFileLibrary.Trim().ToUpper(), sFileName.Trim().ToUpper(), iRecordLength, sLibraryTextDescription.Trim());

            if (CheckObjectExists(sFileName.Trim().ToUpper(), sFileLibrary.Trim().ToUpper(), "*FILE") == false)
                return ExecuteCommand(sCmd);
            else
                throw new Exception(string.Format("File {0} in library {1} already exists.", sFileName.Trim().ToUpper(), sFileLibrary.Trim().ToUpper()));
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  Create fixed length flat physical file using SQL with long or short name. 
    ///  Data field name will be RECORD.
    ///  </summary> 
    ///  <param name="sTableName">SQL table name. Up to 30 characters.</param>
    ///  <param name="sTableLibrary">SQL table schema/library</param>
    ///  <param name="iRecordLength">Record length. Default=400</param>
    ///  <returns>True-Success, False-Error</returns>
    public bool CreateSqlTableFixed(string sTableName, string sTableLibrary, int iRecordLength = 400)
    {
        bool rtncmd = true;

        try
        {

            // Table creation command
            string sSqlCmd = string.Format("CREATE TABLE {0}/{1} (RECORD CHAR ({2}) NOT NULL WITH DEFAULT)", sTableLibrary.Trim().ToUpper(), sTableName.Trim().ToUpper(), iRecordLength);

            // Call SQL create command
            rtncmd = ExecuteSqlNonQuery(sSqlCmd);

            if (rtncmd)
                _LastError = string.Format("Table {0} was created in library {1}.", sTableName.Trim().ToUpper(), sTableLibrary.Trim().ToUpper());
            else
                _LastError = string.Format("Errors occurred. It's possible table {0} in library {1} already exists.", sTableName.Trim().ToUpper(), sTableLibrary.Trim().ToUpper());

            return rtncmd;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  Insert record into fixed length flat physical file using SQL.
    ///  Data field name will be RECORD.
    ///  </summary> 
    ///  <param name="sTableName">SQL table name. Up to 30 characters.</param>
    ///  <param name="sTableLibrary">SQL table schema/library</param>
    ///  <param name="sRecordData">Single field data record</param>
    ///  <returns>True-Success, False-Error</returns>
    public bool InsertSqlTableFixed(string sTableName, string sTableLibrary, string sRecordData)
    {
        bool rtncmd = true;

        try
        {

            // Double up any single quoted records
            sRecordData = sRecordData.Replace("'", "''");

            string sSqlCmd = string.Format("INSERT INTO {0}/{1} (RECORD) VALUES('{2}')", sTableLibrary.Trim().ToUpper(), sTableName.Trim().ToUpper(), sRecordData);

            // Call SQL command
            rtncmd = ExecuteSqlNonQuery(sSqlCmd);

            if (rtncmd)
                _LastError = string.Format("Record inserted to Table {0} in library {1}.", sTableName.Trim().ToUpper(), sTableLibrary.Trim().ToUpper());
            else
                _LastError = string.Format("Errors occurred. It's possible table {0} in library {1} does not exist or there were unpaired single quotes.", sTableName.Trim().ToUpper(), sTableLibrary.Trim().ToUpper());

            return rtncmd;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  Delete table using SQL DROP TABLE action
    ///  </summary> 
    ///  <param name="sTableName">SQL table name. Up to 30 characters.</param>
    ///  <param name="sTableLibrary">SQL table schema/library</param>
    ///  <returns>True-Success, False-Error</returns>
    public bool DeleteSqlTable(string sTableName, string sTableLibrary)
    {
        bool rtncmd = true;

        try
        {

            // SQL command
            string sSqlCmd = string.Format("DROP TABLE {0}/{1}", sTableLibrary.Trim().ToUpper(), sTableName.Trim().ToUpper());

            // Call SQL command
            rtncmd = ExecuteSqlNonQuery(sSqlCmd);

            if (rtncmd)
                _LastError = string.Format("Table {0} was deleted from library {1}.", sTableName.Trim().ToUpper(), sTableLibrary.Trim().ToUpper());
            else
                _LastError = string.Format("Errors occurred. It's possible table {0} in library {1} does not exist.", sTableName.Trim().ToUpper(), sTableLibrary.Trim().ToUpper());

            return rtncmd;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  Clear table by Deleting all records from table using SQL DELETE action
    ///  </summary> 
    ///  <param name="sTableName">SQL table name. Up to 30 characters.</param>
    ///  <param name="sTableLibrary">SQL table schema/library</param>
    ///  <returns>True-Success, False-Error</returns>
    public bool ClearSqlTable(string sTableName, string sTableLibrary)
    {
        bool rtncmd = true;

        try
        {

            // SQL command
            string sSqlCmd = string.Format("DELETE FROM {0}/{1}", sTableLibrary.Trim().ToUpper(), sTableName.Trim().ToUpper());

            // Call SQL command
            rtncmd = ExecuteSqlNonQuery(sSqlCmd);

            if (rtncmd)
                _LastError = string.Format("Records were deleted from Table {0} in library {1}.", sTableName.Trim().ToUpper(), sTableLibrary.Trim().ToUpper());
            else
                _LastError = string.Format("Errors occurred. It's possible table {0} in library {1} does not exist.", sTableName.Trim().ToUpper(), sTableLibrary.Trim().ToUpper());

            return rtncmd;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }
    /// <summary>
    ///  Does SQL table exist ? We check SYSTABLES in QSYS2 for table existence.
    ///  Note: Only works for DDS defined or SQL defined tables. Flat files created with CRTPF will not show up.
    ///  </summary> 
    ///  <param name="sTableName">SQL table name. Up to 30 characters.</param>
    ///  <param name="sTableLibrary">SQL table schema/library</param>
    ///  <returns>True-Exists, False-Does not exist or rrror</returns>
    public bool CheckSqlTableExists(string sTableName, string sTableLibrary)
    {
        DataTable dtWork;
        string sql = "";

        try
        {

            // Build table check for SQL table
            sql = string.Format("SELECT COUNT(*) as TABLECOUNT From QSYS2/SYSTABLES WHERE TABLE_SCHEMA='{0}' and TABLE_NAME='{1}'", sTableLibrary.Trim(), sTableName.Trim());

            // Run the table check query
            dtWork = ExecuteSqlQueryToDataTable(sql);

            if (dtWork == null)
                throw new Exception("SQL error occurred.");
            else
                // Should only ever get a single count result row
                if (dtWork.Rows.Count == 1)
            {
                // Check the count to see if we found the table using SYSTABLES
                if (Convert.ToInt32(dtWork.Rows[0]["TABLECOUNT"]) > 0)
                {
                    _LastError = string.Format("Table {0} in library {1} exists.", sTableName.Trim().ToUpper(), sTableLibrary.Trim().ToUpper());
                    return true;
                }
                else
                {
                    _LastError = string.Format("Table {0} in library {1} does not exist or is possibly a flat file created with CRTPF so not in SYSTABLES.", sTableName.Trim().ToUpper(), sTableLibrary.Trim().ToUpper());
                    return false;
                }
            }
            else
            {
                _LastError = string.Format("Error occurred. Only 1 count row expected.", sTableName.Trim().ToUpper(), sTableLibrary.Trim().ToUpper());
                return false;
            }
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return false;
        }
    }



        /// <summary>
        ///  Make an XMLSERVICE HTTP POST request with selected URL and get response
        ///  </summary>
        ///  <param name="URL">URL where XMLSERVICE is set up.</param>
        ///  <param name="method">POST or GET</param>
        ///  <param name="POSTdata">Data to post ont he request</param>
        ///  <param name="iTimeout">Optional HTTP request timeout. Default = 60000 milliseconds</param>
        ///  <param name="iUseHttpCredentials">Optional Use network credentials for web server auth. 0=No, 1=Yes Default = 0</param>
        ///  <param name="allowInvalidSslCertificates">Optional Allow invallid certs. true=Yes, false=no. Default=false - certs must be valid.</param>
        ///  <returns>XML response or error string starting with "ERROR" </returns>
        ///  <remarks></remarks>
        private string ExecuteHttpPostRequest(string URL, string method, string POSTdata, int iTimeout = 60000, bool iUseHttpCredentials = false, bool allowInvalidSslCertificates = false)
        {
        string responseData = "";
        try
        {
            _LastHTTPResponse = "";
            // Save last post data
            _LastPostData = POSTdata;

            // Set TLS mode to 1.2 if not set and also set allow invalid certificates setting value.   
            if (System.Net.ServicePointManager.SecurityProtocol != System.Net.SecurityProtocolType.Tls12)
            {

                // Set TLS 1.2 protocal 
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

                // If enabled, this callback allows us to ignore invalid certificates. 
                // https://stackoverflow.com/questions/2675133/c-sharp-ignore-certificate-errors
                if (allowInvalidSslCertificates)
                {
                    System.Net.ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;
                }
            }

            System.Net.HttpWebRequest hwrequest = (System.Net.HttpWebRequest) System.Net.WebRequest.Create(URL);
            hwrequest.Accept = "*/*";
            hwrequest.AllowAutoRedirect = true;
            // If specific use Http web server authentication
            if (iUseHttpCredentials)
                hwrequest.Credentials = new NetworkCredential(_HttpUser, _HttpPassword);
            hwrequest.UserAgent = "XmlServicei/0.1";
            hwrequest.Timeout = iTimeout;
            hwrequest.Method = method;
            if (hwrequest.Method == "POST")
            {
                hwrequest.ContentType = "application/x-www-form-urlencoded";

                // Percent encoding fix to resolve issues with CL commands or SQL statements containing % signs
                // We convert pct signs to %25 before encoding the URL.
                // This should compensate for XMLCGI not handling percent signs in calls - 8/14/2016
                POSTdata = POSTdata.Replace("%", "%25");
                // Now encode the URL post data before posting - 8/14/2016
                POSTdata = EncodeUrl(POSTdata);

                System.Text.ASCIIEncoding encoding = new System.Text.ASCIIEncoding(); // Use UTF8Encoding for XML requests
                byte[] postByteArray = encoding.GetBytes(POSTdata);
                hwrequest.ContentLength = postByteArray.Length;
                System.IO.Stream postStream = hwrequest.GetRequestStream();
                postStream.Write(postByteArray, 0, postByteArray.Length);
                postStream.Close();
            }
            System.Net.HttpWebResponse hwresponse = (System.Net.HttpWebResponse) hwrequest.GetResponse();
            _LastHTTPResponse = hwresponse.StatusCode + " " + hwresponse.StatusDescription;
            if (hwresponse.StatusCode == System.Net.HttpStatusCode.OK)
            {
                System.IO.StreamReader responseStream = new System.IO.StreamReader(hwresponse.GetResponseStream());
                responseData = responseStream.ReadToEnd();
            }
            hwresponse.Close();
        }
        catch (Exception e)
        {
            responseData = "ERROR - An HTTP error occurred: " + e.Message;
        }
        return responseData;
    }

    /// <summary>
    ///  Encode URL string
    ///  </summary>
    ///  <param name="sURL">URL to encode</param>
    ///  <returns>Encoded URL string</returns>
    ///  <remarks></remarks>
    private string EncodeUrl(string sURL)
    {
        try
        {
            _LastError = "";
            sURL = System.Web.HttpUtility.UrlEncode(sURL);
            return sURL;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            // Return original URL
            return sURL;
        }
    }

    /// <summary>
    ///  This function loads an XML file that contains a data stream returned from the XMLSERVICE service program
    ///  into an internal XML response DataSet that can be retreived via the GetDataSet function.
    ///  The results are processed into a DataTable which can then be 
    ///  </summary>
    ///  <param name="sXMLFile">XML File</param>
    ///  <returns>True-successfully loaded XML file, False-failed to load XML file</returns>
    ///  <remarks></remarks>
    private bool LoadDataTableFromXmlResponseFile(string sXMLFile)
    {
        try
        {
            _LastError = "";

            // Bail if no XML file
            if (System.IO.File.Exists(sXMLFile) == false)
                throw new Exception("XML file " + sXMLFile + " does not exist. Process cancelled.");

            // Load XML data into a dataset  
            _dsXmlResponseData = new DataSet();
            _dsXmlResponseData.ReadXml(sXMLFile);

            // Extract returned SQL column definitions from XML file
            _dtColumnDefinitions = _dsXmlResponseData.Tables["col"];
            _iColumnCount = _dsXmlResponseData.Tables["col"].Rows.Count;
            _iRowCount = _dsXmlResponseData.Tables["data"].Rows.Count / _iColumnCount;

            // Start new DataTable for XML resultset returned by XMLSERVICE from IBM i
            _dtReturnData = new DataTable();
            // Set the table name
            _dtReturnData.TableName = "Table1";

            foreach (DataRow dr1 in _dsXmlResponseData.Tables["col"].Rows)
                _dtReturnData.Columns.Add(dr1[0].ToString(), Type.GetType("System.String"));

            int colct=0;
            DataRow row2 = null;
            foreach (DataRow dr1 in _dsXmlResponseData.Tables["data"].Rows)
            {

                // If first column field, create new row
                if (colct == 0)
                    row2 = _dtReturnData.NewRow();
                else if (colct == _iColumnCount)
                {
                    // All 
                    _dtReturnData.Rows.Add(row2);
                    row2 = _dtReturnData.NewRow();
                    colct = 0;
                }

                // Set column value for row
                row2[dr1[0].ToString()] = dr1[1].ToString();
                colct = colct + 1;
            }

            // Add last row to dataset
            _dtReturnData.Rows.Add(row2);

            _LastError = _iRowCount + " rows were returned from XML file " + sXMLFile;

            _bXMLIsLoaded = true;
            return true;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            _bXMLIsLoaded = false;
            return false;
        }
    }
    /// <summary>
    ///  This function loads an XML string that contains a data stream returned from the XMLSERVICE service program.
    ///  The XML data is loaded into a DataSet and then the data is loaded into a DataTable.
    ///  </summary>
    ///  <param name="sXMLString">XML data string returned from query</param>
    ///  <returns>True-successfully loaded XML file, False-failed to load XML file</returns>
    ///  <remarks></remarks>
    private bool LoadDataTableFromXmlResponseString(string sXMLString)
    {
        try
        {
            _LastError = "";

            System.IO.StringReader rdr = new System.IO.StringReader(sXMLString);

            // Load XML response data into a temporary work dataset  
            _dsXmlResponseData = new DataSet();
            _dsXmlResponseData.ReadXml(rdr);

            // Extract returned SQL column definitions from XML file
            _dtColumnDefinitions = _dsXmlResponseData.Tables["col"];
            _iColumnCount = _dsXmlResponseData.Tables["col"].Rows.Count;
            _iRowCount = _dsXmlResponseData.Tables["data"].Rows.Count / _iColumnCount;

            // Start new DataTable for XML resultset returned by XMLSERVICE from IBM i
            _dtReturnData = new DataTable();
            // Set table name
            _dtReturnData.TableName = "Table1";
            foreach (DataRow dr1 in _dsXmlResponseData.Tables["col"].Rows)
                _dtReturnData.Columns.Add(dr1[0].ToString(), Type.GetType("System.String"));

            int colct=0;
            DataRow row2 = null;
            foreach (DataRow dr1 in _dsXmlResponseData.Tables["data"].Rows)
            {

                // If first column field, create new row
                if (colct == 0)
                    row2 = _dtReturnData.NewRow();
                else if (colct == _iColumnCount)
                {
                    // All 
                    _dtReturnData.Rows.Add(row2);
                    row2 = _dtReturnData.NewRow();
                    colct = 0;
                }

                // Set column value for row
                row2[dr1[0].ToString()] = dr1[1].ToString();
                colct = colct + 1;
            }

            // Add last row to dataset
            _dtReturnData.Rows.Add(row2);

            _LastError = _iRowCount + " rows were returned from XML string.";

            _bXMLIsLoaded = true;
            return true;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            _bXMLIsLoaded = false;
            return false;
        }
    }
    /// <summary>
    ///  This function loads an XML string that contains a data stream returned from the XMLSERVICE service program
    ///  </summary>
    ///  <param name="sXMLString">XML File</param>
    ///  <returns>True-successfully loaded XML file, False-failed to load XML file</returns>
    ///  <remarks></remarks>
    private bool LoadPgmCallDataTableFromXmlResponseString(string sXMLString)
    {
        try
        {
            _LastError = "";

            System.IO.StringReader rdr = new System.IO.StringReader(sXMLString);

            // Load XML data into a dataset  
            _dsProgramResponse = new DataSet();
            _dsProgramResponse.ReadXml(rdr);

            // Extract returned SQL column definitions from XML file
            // 'dtColumnDefinitions = mDS1.Tables("col")
            _iRowCount = _dsProgramResponse.Tables["data"].Rows.Count;

            // Start new DataTable for XML resultset returned by XMLSERVICE from IBM i
            _dtProgramResponse = new DataTable();
            // Set the table name
            _dtProgramResponse.TableName = "Table1";
            _dtProgramResponse.Columns.Add("parmtype", Type.GetType("System.String"));
            _dtProgramResponse.Columns.Add("parmvalue", Type.GetType("System.String"));

            DataRow row2 = null;
            int ict = 0;
            foreach (DataRow dr1 in _dsProgramResponse.Tables["data"].Rows)
            {
                row2 = _dtProgramResponse.NewRow();
                row2[0] = "Parm" + ict; // Set unique ordinal zero based parm name 
                row2[1] = dr1[1];
                _dtProgramResponse.Rows.Add(row2);
                ict += 1;
            }

            _LastError = _iRowCount + " rows were returned from XML string.";

            _bXMLIsLoaded = true;
            return true;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            _bXMLIsLoaded = false;
            return false;
        }
    }
    /// <summary>
    ///  This function loads an XML string that contains a data stream returned from the XMLSERVICE service program command calls
    ///  </summary>
    ///  <param name="sXMLString">XML File</param>
    ///  <returns>True-successfully loaded XML file, False-failed to load XML file</returns>
    ///  <remarks></remarks>
    private bool LoadCmdCallDataTableFromXmlResponseString(string sXMLString)
    {
        int iTableCount = 0;

        try
        {
            _LastError = "";

            System.IO.StringReader rdr = new System.IO.StringReader(sXMLString);

            // Load XML data into a dataset  
            _dsCommandResponse = new DataSet();
            _dsCommandResponse.ReadXml(rdr);

            // Get number of tables
            iTableCount = _dsCommandResponse.Tables.Count;

            // If only 1 table, command ran successfully

            if (iTableCount == 1)
            {

                // Extract returned SQL column definitions from XML file
                // 'dtColumnDefinitions = mDS1.Tables("col")
                _iRowCount = _dsCommandResponse.Tables["cmd"].Rows.Count;

                // Start new DataTable for XML resultset returned by XMLSERVICE from IBM i
                _dtCommandResponse = new DataTable();
                // Set the table name
                _dtCommandResponse.TableName = "Table1";
                _dtCommandResponse.Columns.Add("parmtype", Type.GetType("System.String"));
                _dtCommandResponse.Columns.Add("parmvalue", Type.GetType("System.String"));

                DataRow row2 = null;
                int ict = 0;
                foreach (DataRow dr1 in _dsCommandResponse.Tables["cmd"].Rows)
                {
                    row2 = _dtCommandResponse.NewRow();
                    row2[0] = "Msg" + ict; // Set unique ordinal zero based parm name 
                    row2[1] = dr1[0];
                    _dtCommandResponse.Rows.Add(row2);
                    ict += 1;
                }
            }
            else if (iTableCount > 1)
            {

                // Extract returned SQL column definitions from XML file
                // 'dtColumnDefinitions = mDS1.Tables("col")
                _iRowCount = _dsCommandResponse.Tables["joblogrec"].Rows.Count;

                // Start new DataTable for XML resultset returned by XMLSERVICE from IBM i
                _dtCommandResponse = new DataTable();
                // Set the table name
                _dtCommandResponse.TableName = "Table1";
                _dtCommandResponse.Columns.Add("parmtype", Type.GetType("System.String"));
                _dtCommandResponse.Columns.Add("parmvalue", Type.GetType("System.String"));

                DataRow row2 = null;
                int ict = 0;
                foreach (DataRow dr1 in _dsCommandResponse.Tables["joblogrec"].Rows)
                {
                    row2 = _dtCommandResponse.NewRow();
                    row2[0] = "Msg" + ict; // Set unique ordinal zero based parm name 
                    row2[1] = dr1[0].ToString() + " - " + dr1[2].ToString() + " - " + dr1[1].ToString();
                    _dtCommandResponse.Rows.Add(row2);
                    ict += 1;
                }
            }

            _LastError = _iRowCount + " rows were returned from XML string.";

            _bXMLIsLoaded = true;
            return true;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            _bXMLIsLoaded = false;
            return false;
        }
    }
    /// <summary>
    ///  Export DataTable Row List to Generic List and optionally include column names.
    ///  </summary>
    ///  <param name="dtTemp">DataTable Object</param>
    ///  <param name="firstRowColumnNames">Optional - Return first row as column names. False=No column names, True=Return column names. Default=False</param>
    ///  <returns>List object</returns>
    private List<List<object>> ExportDataTableToList(DataTable dtTemp, bool firstRowColumnNames = false)
    {
        List<List<object>> result = new List<List<object>>();
        List<object> values = new List<object>();

        try
        {
            _LastError = "";

            // Include first row as columns
            if (firstRowColumnNames)
            {
                foreach (DataColumn column in dtTemp.Columns)
                    values.Add(column.ColumnName);
                result.Add(values);
            }

            // Output all the data now
            foreach (DataRow row in dtTemp.Rows)
            {
                values = new List<object>();
                foreach (DataColumn column in dtTemp.Columns)
                {
                    if (row.IsNull(column))
                        values.Add(null);
                    else
                        values.Add(row[column]);
                }
                result.Add(values);
            }
            return result;
        }
        catch (Exception ex)
        {
            _LastError = ex.Message;
            return null;
        }
    }
}

}
