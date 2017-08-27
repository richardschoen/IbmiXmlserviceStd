Imports System.Xml
Imports System.Data
Imports System.Net
Imports System.Web
Imports System
Imports System.IO
Imports System.Text
Imports System.Runtime.Serialization

''' <summary>
''' This class is used to interface with existing IBM i database, program calls, CL commands, service programs and 
''' data queues via the XMLSERVICE service program. It also needs an Apache instance hooked up for XMLCGI calls
''' since this class uses the HTTP interface. 
''' 
''' Return data is returned in a usable .Net DataTable format or you can process the raw XML responses yourself.
''' 
''' This class should work with V5R4 and above of the OS400/IBM i operating system running XMLSERVICE.
''' 
''' Requirement: The XMLSERVICE library must also exist and be compiled on the system, including the XMLCGI program. 
''' Apache instance must also be set up and running and configured for XMLCGI calls.
''' 
''' Note: For appropriate security you should configure your Apache instance for SSL and for user Authentication. This
''' way all traffic is secured and there is an extra User/Password layer. To enable HTTP authentication use the 
''' UseHttpCredentials parameter and set it to True on the SetHttpUserInfo method.
''' 
''' You can always refer to the Yips site for more install info on XMLSERVICE: 
''' http://yips.idevcloud.com/wiki/index.php/xmlservice/xmlserviceinstall 
''' </summary>
''' <remarks>Tested with XML Tookit library (XMLSERVICE) V1.9.7</remarks>
Public Class XmlServicei

#Region "PrivateVariables"

    Private _bXMLIsLoaded As Boolean = False
    Private _LastError As String
    Private _dtColumnDefinitions As DataTable
    Private _iColumnCount As Integer
    Private _iRowCount As Integer
    Private _dsXmlResponseData As DataSet
    Private _dtReturnData As DataTable
    Private _dsCommandResponse As DataSet
    Private _dtCommandResponse As DataTable
    Private _dsProgramResponse As DataSet
    Private _dtProgramResponse As DataTable
    Private _BaseURL As String = ""
    Private _DB2Parm As String = "db2=@@db2value&uid=@@uidvalue&pwd=@@pwdvalue&ipc=@@ipcvalue&ctl=@@ctlvalue&xmlin=@@xmlinvalue&xmlout=@@xmloutvalue"
    Private _User As String = ""
    Private _Password As String = ""
    Private _IPCINfo As String = "/tmp/xmlservicei"
    Private _IPCPersistance As String = "*sbmjob"
    Private _DB2Info As String = "*LOCAL"
    Private _LastHTTPResponse As String = ""
    Private _LastXMLResponse As String = ""
    Private _iXmlResponseBufferSize As Integer = 500000
    Private _HttpTimeout As Integer = 60000
    Private _UseHttpCredentials As Boolean = False
    Private _HttpUser As String = ""
    Private _HttpPassword As String = ""
    Private _LastDataQueueName As String = ""
    Private _LastDataQueueLibrary As String = ""
    Private _LastDataQueueDataSent As String = ""
    Private _LastDataQueueDataReceived As String = ""
    Private _LastDataQueueLengthReceived As Integer = 0
#End Region

#Region "Structures"

    ''' <summary>
    ''' Program call parameter structure
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure PgmParmList
        Dim parmtype As String
        Dim parmvalue As String
    End Structure
    ''' <summary>
    ''' Program call return structure
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure RtnPgmCall
        Dim success As Boolean
        Dim parms() As PgmParmList
    End Structure
    ''' <summary>
    ''' Stored procedure parameter list
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure ProcedureParmList
        Dim parmtype As String
        Dim parmvalue As String
        Dim parmname As String
    End Structure
    ''' <summary>
    ''' Stored procedure return structure
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure RtnProcedureCall
        Dim success As Boolean
        Dim parms() As PgmParmList
    End Structure
#End Region

#Region "SetMethods"

    ''' <summary>
    ''' Set HTTP timeout for XMLSERVICE requests
    ''' </summary>
    ''' <param name="iHttpTimeout">HTTP timeout in milliseconds</param>
    ''' <returns>True-Success, False-Fail</returns>
    ''' <remarks></remarks>
    Public Function SetHttpTimeout(iHttpTimeout As Integer) As Boolean
        Try
            _LastError = ""
            _HttpTimeout = iHttpTimeout
            Return True
        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Set base URL to XMLCGI program.
    ''' Example: http://1.1.1.1:30000/cgi-bin/xmlcgi.pgm
    ''' Set this value one time each time the class is instantiated.
    ''' </summary>
    ''' <param name="sBaseUrl">Base URL to set for path to XMLSERVICE and XMLCGI calls.</param>
    ''' <returns>True-Success, False-Fail</returns>
    ''' <remarks></remarks>
    Public Function SetBaseURL(sBaseUrl As String) As Boolean
        Try
            _LastError = ""
            _BaseURL = sBaseUrl
            Return True
        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Set ALL base user info parameters for XMLCGI program calls in a single method call.
    ''' Set this value one time each time the class Is instantiated.
    ''' This is a convenience method to set all connection info in a single call.
    ''' </summary>
    ''' <param name="sBaseUrl">Base URL to set for path to XMLSERVICE and XMLCGI calls.</param>
    ''' <param name="sUser">IBM i User and HTTP Auth</param>
    ''' <param name="sPassword">IBM i Password and HTTP Auth</param>
    ''' <param name="UseHttpCredentials">Use Apache HTTP authentication credentials</param>
    ''' <param name="sDb2Info">DB2 server info. Default = *LOCAL for current DB2 server</param>
    ''' <param name="sIpcInfo">IPC info. Example: /tmp/xmlservicei</param>
    ''' <param name="persistjobs">True-Stateful XTOOLKIT jobs are started that must be ended eventually with KillService method. False-Stateless jobs. XMLSERVICE XTOOLKIT jobs end immediately after call completes.</param>
    ''' <param name="sHttpUser">Http Auth user. Only use if HTTP auth credentials are different than IBMi user info and web server auth enabled.</param>
    ''' <param name="sHttpPass">Http Auth password. Only use if HTTP auth credentials are different than IBMi user info and web server auth enabled.</param>
    ''' <param name="iSize">XML response buffer size. Default = 500000</param>
    ''' <returns>True-Success, False-Fail</returns>
    ''' <remarks></remarks>
    Public Function SetUserInfoExt(sBaseUrl As String, sUser As String, sPassword As String, UseHttpCredentials As Boolean, Optional sIpcInfo As String = "/tmp/xmlservicei", Optional persistJobs As Boolean = True, Optional sDb2Info As String = "*LOCAL", Optional iHttpTimeout As Integer = 60000, Optional sHttpUser As String = "", Optional sHttpPass As String = "", Optional iSize As Integer = 500000) As Boolean

        Try

            _LastError = ""

            If SetBaseURL(sBaseUrl) = False Then
                Throw New Exception("Error setting base URL")
            End If

            If SetUserInfo(sUser, sPassword, UseHttpCredentials) = False Then
                Throw New Exception("Error setting user info")
            End If

            If SetDb2Info(sDb2Info) = False Then
                Throw New Exception("Error setting DB2 info")
            End If

            If SetHttpTimeout(iHttpTimeout) = False Then
                Throw New Exception("Error setting HTTP timeout info")
            End If

            If SetIpcInfo(sIpcInfo, persistJobs) = False Then
                Throw New Exception("Error setting IPC info")
            End If

            'Change HTTP auth user and password to be other than default IBMi credentials
            If sHttpPass.Trim <> "" And sHttpUser.Trim <> "" Then
                If SetHttpUserInfo(sHttpUser, sHttpPass, UseHttpCredentials) Then
                    Throw New Exception("Error setting HTTP user info")
                End If
            End If

            If SetXmlResponseBufferSize(iSize) = False Then
                Throw New Exception("Error setting XML response buffer size")
            End If

            Return True

        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Set base user info for XMLCGI program calls.
    ''' Set this value one time each time the class Is instantiated.
    ''' This sets the IBM i user login info And also sets the default
    ''' HTTP auth user credentials for the Apache server if HTTP authentication 
    ''' Is enabled on the Apache server. 
    ''' </summary>
    ''' <param name="sUser">IBM i User</param>
    ''' <param name="sPassword">IBM i Password</param>
    ''' <param name="UseHttpCredentials">Use Apache HTTP authentication credentials</param>
    ''' <returns>True-Success, False-Fail</returns>
    ''' <remarks></remarks>
    Public Function SetUserInfo(sUser As String, sPassword As String, UseHttpCredentials As Boolean) As Boolean
        Try
            _LastError = ""
            'Set IBM i user info
            _User = sUser
            _Password = sPassword
            ' Set IBM i apache authentication user info default'
            _HttpUser = sUser
            _HttpPassword = sPassword
            'Set use Apache HTTP authentication flag
            _UseHttpCredentials = UseHttpCredentials
            Return True
        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Set HTTP Apache authenticated user credential info for XMLCGI program calls.
    ''' Set this value one time each time the class Is instantiated.
    ''' Also make sure you call SetUserInfo first which sets the the IBM i user And 
    ''' the default HTTP user And password to the same User And Password as SetUserInfo. 
    ''' The IBM i user profile And password can be overridden with SetHttpUserInfo 
    ''' if the Apache authentication user Is Not the same as the IBM i user profile 
    ''' Or it uses an authorization list Or LDAP user for HTTP authentication
    ''' if HTTP auth user Is different than IBM i user.
    ''' </summary>
    ''' <param name="sUser">IBM i Apache HTTP Server Web Site Auth User</param>
    ''' <param name="sPassword">IBM i Apache HTTP Server Web Site Auth Password</param>
    ''' <returns>True-Success, False-Fail</returns>
    ''' <remarks></remarks>
    Public Function SetHttpUserInfo(sUser As String, sPassword As String, UseHttpCredentials As Boolean) As Boolean
        Try
            _LastError = ""
            _HttpUser = sUser
            _HttpPassword = sPassword
            'Set use Apache HTTP authentication flag
            _UseHttpCredentials = UseHttpCredentials
            Return True
        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Set IPC info and stateful or stateless persistence info.
    ''' The class defaults to /tmp/xmlservicei so this call not needed unless you want to use custom IPC info
    ''' </summary>
    ''' <param name="sIpcInfo">IPC info. Use a unique path for each persistent job. Example: /tmp/xmlservicei</param>
    ''' <param name="persistjobs">True-Stateful XTOOLKIT jobs are started that must be ended eventually with KillService method. False-Stateless jobs. XMLSERVICE XTOOLKIT jobs end immediately after call completes.</param>
    ''' <returns>True-Success, False-Fail</returns>
    ''' <remarks></remarks>
    Public Function SetIpcInfo(Optional sIpcInfo As String = "/tmp/xmlservicei", Optional persistJobs As Boolean = True) As Boolean
        Try
            _LastError = ""
            'Set IPC info only if a valur is passed. 
            If sIpcInfo <> "" Then
                _IPCINfo = sIpcInfo
            End If
            'Read here for more on persistence.
            'http://yips.idevcloud.com/wiki/index.php/XMLService/XMLSERVICEConnect

            If persistJobs Then 'XMLSERVICE jobs will persist for stateful processes
                _IPCPersistance = "*sbmjob"
            Else 'XMLSERVICE jobs will not persist for stateless processes
                _IPCPersistance = "*here"
            End If

            Return True
        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Set DB2 system info. Defaults to *LOCAL to access local system database
    ''' </summary>
    ''' <param name="sDB2Info">DB2 server info. Default = *LOCAL for current DB2 server</param>
    ''' <returns>True-Success, False-Fail</returns>
    ''' <remarks></remarks>
    Public Function SetDb2Info(Optional sDB2Info As String = "*LOCAL") As Boolean
        Try
            _LastError = ""
            _DB2Info = sDB2Info
            Return True
        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Set XML response buffer size 
    ''' </summary>
    ''' <param name="iSize">XML response buffer size. Default = 500000</param>
    ''' <returns>True-Success, False-Fail</returns>
    ''' <remarks></remarks>
    Public Function SetXmlResponseBufferSize(Optional iSize As Integer = 500000) As Boolean
        Try
            _LastError = ""
            _iXmlResponseBufferSize = iSize
            Return True
        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function

#End Region

#Region "GetMethods"

    ''' <summary>
    ''' This function gets the XML Datatable of data that was loaded with LoadPgmCallDataSetFromXmlString 
    ''' </summary>
    ''' <returns>Data table of data or nothing if no data set</returns>
    ''' <remarks></remarks>
    Public Function GetProgramResponseDataTable() As DataTable
        Try
            _LastError = ""

            If _bXMLIsLoaded = False Then
                Throw New Exception("No XML data is currently loaded. Call LoadPgmCallDataTableFromXmlResponseString to load an XML file first.")
            End If

            Return _dtProgramResponse
        Catch ex As Exception
            _LastError = ex.Message
            Return Nothing
        End Try
    End Function
    ''' <summary>
    ''' This function gets the XML Datatable of data that was loaded with LoadCmdCallDataSetFromXmlString 
    ''' </summary>
    ''' <returns>Data table of data or nothing if no data set</returns>
    ''' <remarks></remarks>
    Public Function GetCommandResponseDataTable() As DataTable
        Try
            _LastError = ""

            If _bXMLIsLoaded = False Then
                Throw New Exception("No XML data is currently loaded. Call LoadCmdCallDataTableFromXmlResponseString to load an XML file first.")
            End If

            Return _dtCommandResponse
        Catch ex As Exception
            _LastError = ex.Message
            Return Nothing
        End Try
    End Function
    ''' <summary>
    ''' This function gets the XML DataSet of data that was loaded with LoadDataSetFromXMLFile
    ''' Only use this DataSet if you want to access the raw XML response data without further processing.
    ''' </summary>
    ''' <returns>DataSet of XML response data or nothing if no data set</returns>
    ''' <remarks></remarks>
    Public Function GetXmlResponseDataSet() As DataSet
        Try
            _LastError = ""

            If _bXMLIsLoaded = False Then
                Throw New Exception("No XML data is currently loaded. Call LoadDataTableFromXmlResponseFile or LoadDataTableFromXmlResponseString to load an XML file first.")
            End If

            Return _dsXmlResponseData
        Catch ex As Exception
            _LastError = ex.Message
            Return Nothing
        End Try
    End Function
    ''' <summary>
    ''' Returns last error message string
    ''' </summary>
    ''' <returns>Last error message string</returns>
    ''' <remarks></remarks>
    Public Function GetLastError() As String
        Try
            Return _LastError
        Catch ex As Exception
            Return ""
        End Try
    End Function
    ''' <summary>
    ''' Returns XML response message string from last XMLSERVICE call.
    ''' </summary>
    ''' <returns>Last XML response message string</returns>
    ''' <remarks></remarks>
    Public Function GetLastXmlResponse() As String
        Try
            Return _LastXMLResponse
        Catch ex As Exception
            Return ""
        End Try
    End Function
    ''' <summary>
    ''' This function gets the DataTable of data loaded from XML with LoadDataSetFromXMLFile and returns as a CSV string
    ''' </summary>
    ''' <param name="sFieldSepchar">Field delimiter/separator. Default = Comma</param>
    ''' <param name="sFieldDataDelimChar">Field data delimiter character. Default = double quotes.</param>
    ''' <returns>CSV string from DataTable</returns>
    Public Function GetQueryResultsDataTableToCsvString(Optional sFieldSepchar As String = ",", Optional sFieldDataDelimChar As String = """") As String

        Try
            _LastError = ""

            Dim sHeadings As String = ""
            Dim sBody As String = ""
            Dim sCsvData As New StringBuilder

            ' first write a line with the columns name
            Dim sep As String = ""
            Dim builder As New System.Text.StringBuilder
            For Each col As DataColumn In _dtReturnData.Columns
                builder.Append(sep).Append(col.ColumnName)
                sep = sFieldSepchar
            Next
            sCsvData.AppendLine(builder.ToString())

            ' then write all the rows
            For Each row As DataRow In _dtReturnData.Rows
                sep = ""
                builder = New System.Text.StringBuilder

                For Each col As DataColumn In _dtReturnData.Columns
                    builder.Append(sep)
                    builder.Append(sFieldDataDelimChar).Append(row(col.ColumnName)).Append(sFieldDataDelimChar)
                    sep = sFieldSepchar
                Next
                sCsvData.AppendLine(builder.ToString())
            Next

            'Return CSV output
            Return sCsvData.ToString

        Catch ex As Exception
            _LastError = ex.Message
            Return ""
        End Try
    End Function
    ''' <summary>
    ''' This function gets the DataTable of XML data loaded from the last query with LoadDataSetFromXMLFile and returns as a CSV file
    ''' </summary>
    ''' <param name="sOutputFile">Output CSV file</param>
    ''' <param name="sFieldSepchar">Field delimiter/separator. Default = Comma</param>
    ''' <param name="sFieldDataDelimChar">Field data delimiter character. Default = double quotes.</param>
    ''' <param name="replace">Replace output file True=Replace file,False=Do not replace</param>
    ''' <returns>True-CSV file written successfully, False-Failure writing CSV output file.</returns>
    Public Function GetQueryResultsDataTableToCsvFile(ByVal sOutputFile As String, Optional sFieldSepchar As String = ",", Optional sFieldDataDelimChar As String = """", Optional replace As Boolean = False) As Boolean

        Dim sCsvWork As String

        Try
            _LastError = ""

            'Delete existing file if replacing
            If File.Exists(sOutputFile) Then
                If replace Then
                    File.Delete(sOutputFile)
                Else
                    Throw New Exception("Output file " & sOutputFile & " already exists and replace not selected.")
                End If
            End If

            'Get data and output
            Using writer As System.IO.StreamWriter = New System.IO.StreamWriter(sOutputFile)

                'Get CSV string
                sCsvWork = GetQueryResultsDataTableToCsvString(sFieldSepchar, sFieldDataDelimChar)

                'Write out CSV data
                writer.Write(sCsvWork)

                'Flush final output and close
                writer.Flush()
                writer.Close()

                Return True
            End Using
        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function

    ''' <summary>
    ''' This function gets the DataTable of data loaded from XML with LoadDataSetFromXMLFile and returns as a XML string
    ''' </summary>
    ''' <param name="sTableName">Table name. Default = "Table1"</param>
    ''' <param name="bWriteSchema">Write XML schema in return data</param>
    ''' <returns>XML string from data table</returns>
    Public Function GetQueryResultsDataTableToXmlString(Optional sTableName As String = "Table1", Optional bWriteSchema As Boolean = False) As String

        Dim sRtnXml As String = ""

        Try
            _LastError = ""

            'if table not set, default to Table1
            If sTableName.Trim = "" Then sTableName = "Table1"

            'Export results to XML
            If _dtReturnData Is Nothing = False Then
                Dim SB As New StringBuilder
                Dim SW As New IO.StringWriter(SB)
                _dtReturnData = GetQueryResultsDataTable()
                _dtReturnData.TableName = sTableName
                'Write XMl with or without schema info
                If bWriteSchema Then
                    _dtReturnData.WriteXml(SW, Data.XmlWriteMode.WriteSchema)
                Else
                    _dtReturnData.WriteXml(SW)
                End If
                sRtnXml = SW.ToString
                SW.Close()
                Return sRtnXml
            Else
                Throw New Exception("No data available. Error: " & GetLastError())
            End If

        Catch ex As Exception
            _LastError = ex.Message
            Return ""
        End Try
    End Function
    ''' <summary>
    ''' This function gets the DataTable of XML data loaded from the last query with LoadDataSetFromXMLFile and returns as a CSV file
    ''' </summary>
    ''' <param name="sOutputFile">Output CSV file</param>
    ''' <param name="sTableName">Table name. Default = "Table1"</param>
    ''' <param name="bWriteSchema">Write XML schema in return data</param>
    ''' <param name="replace">Replace output file True=Replace file,False=Do not replace</param>
    ''' <returns>True-XML file written successfully, False-Failure writing XML output file.</returns>
    Public Function GetQueryResultsDataTableToXmlFile(ByVal sOutputFile As String, Optional sTableName As String = "Table1", Optional bWriteSchema As Boolean = False, Optional replace As Boolean = False) As Boolean

        Dim sXmlWork As String

        Try
            _LastError = ""

            'Delete existing file if replacing
            If File.Exists(sOutputFile) Then
                If replace Then
                    File.Delete(sOutputFile)
                Else
                    Throw New Exception("Output file " & sOutputFile & " already exists and replace not selected.")
                End If
            End If

            'Get data and output 
            Using writer As System.IO.StreamWriter = New System.IO.StreamWriter(sOutputFile)

                'Get XML string
                sXmlWork = GetQueryResultsDataTableToXmlString(sTableName, bWriteSchema)

                'Write out CSV data
                writer.Write(sXmlWork)

                'Flush final output and close
                writer.Flush()
                writer.Close()

                Return True
            End Using

        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function
    ''' <summary>
    ''' This function gets the DataTable of data loaded from XML with LoadDataSetFromXMLFile and returns as a JSON string
    ''' </summary>
    ''' <returns>CSV string from DataTable</returns>
    Public Function GetQueryResultsDataTableToJsonString(Optional debugInfo As Boolean = False) As String

        'TODO - Use Newtonsoft JSON to convert to JSON

        Dim sJsonData As String = ""
        Dim oJsonHelper As New JsonHelper

        Try
            _LastError = ""

            'If data table is blank, bail
            If _dtReturnData Is Nothing Then
                Throw New Exception("Data table is Nothing. No data available.")
            End If

            'Convert DataTable to JSON
            oJsonHelper.DataTableToJsonWithStringBuilder(_dtReturnData, debugInfo)

            'Return JSON output
            Return sJsonData.ToString

        Catch ex As Exception
            _LastError = ex.Message
            Return ""
        End Try
    End Function
    ''' <summary>
    ''' This function gets the DataTable of XML data loaded from the last query with LoadDataSetFromXMLFile and returns as a JSON file
    ''' </summary>
    ''' <param name="sOutputFile">Output JSON file</param>
    ''' <param name="replace">Replace output file True=Replace file,False=Do not replace</param>
    ''' <returns>True-JSON file written successfully, False-Failure writing JSON output file.</returns>
    Public Function GetQueryResultsDataTableToJsonFile(ByVal sOutputFile As String, Optional replace As Boolean = False) As Boolean

        Dim sJsonWork As String

        Try
            _LastError = ""

            'Delete existing file if replacing
            If File.Exists(sOutputFile) Then
                If replace Then
                    File.Delete(sOutputFile)
                Else
                    Throw New Exception("Output file " & sOutputFile & " already exists and replace not selected.")
                End If
            End If

            'Get data and output 
            Using writer As System.IO.StreamWriter = New System.IO.StreamWriter(sOutputFile)

                'Get JSON string
                sJsonWork = GetQueryResultsDataTableToJsonString()

                'Write out JSON data
                writer.Write(sJsonWork)

                'Flush final output and close
                writer.Flush()
                writer.Close()

                Return True
            End Using

        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function
    ''' <summary>
    ''' This function gets the DataTable containing records from the last query response data loaded from XML.
    ''' </summary>
    ''' <returns>Data table of data or nothing if no data set</returns>
    ''' <remarks></remarks>
    Public Function GetQueryResultsDataTable() As DataTable
        Try
            _LastError = ""

            If _bXMLIsLoaded = False Then
                Throw New Exception("No XML data is currently loaded. Call LoadDataSetFromXMLFile to load an XML file first.")
            End If

            Return _dtReturnData
        Catch ex As Exception
            _LastError = ex.Message
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Perform HTTP Get with selected URL to a string
    ''' </summary>
    ''' <param name="sURL">URL to Get</param>
    ''' <returns>String return value</returns>
    Public Function GetUrlToString(sURL As String) As String
        Try
            _LastError = ""

            Dim fr As System.Net.HttpWebRequest
            Dim targetURI As New Uri(sURL)
            Dim s1 As String = ""
            fr = DirectCast(HttpWebRequest.Create(targetURI), System.Net.HttpWebRequest)
            fr.Method = "GET"
            If (fr.GetResponse().ContentLength > 0) Then
                Dim str As New System.IO.StreamReader(fr.GetResponse().GetResponseStream())
                s1 = str.ReadToEnd
                str.Close()
            End If

            Return s1

        Catch ex As System.Net.WebException
            _LastError = ex.Message
            Return ""
        End Try

    End Function

#End Region

#Region "ExecuteMethods"

    ''' <summary>
    ''' This function runs an SQL INSERT, UPDATE, DELETE or other action query against the DB2 database with selected SQL statement.
    ''' </summary>
    ''' <param name="sSQL">SQL INSERT, UPDATE and DELETE. Select is not allowed </param>
    ''' <param name="sQueryResultOutputFile">PC output file for XML response data</param>
    ''' <returns>True - Query service call succeeded, False - Query service call failed</returns>
    ''' <remarks>Note: Committment control is disabled via the commit='none' option so journaling is not used at the moment on any files you plan to modify via INSERT/UPDATE/DELETE</remarks>
    Public Function ExecuteSqlNonQuery(sSQL As String, Optional sQueryResultOutputFile As String = "") As Boolean

        Dim sdb2parm As String = _DB2Parm
        Dim sRtnXML As String = ""
        Dim rtnexecute As Boolean
        Dim sSuccessValue As String = "<execute stmt='stmt1'>@@CHR10<success>+++ success stmt1</success>@@CHR10</execute>"
        Dim rtnload As Boolean = False

        Try

            _LastError = ""
            sRtnXML = ""
            _LastHTTPResponse = ""
            _LastXMLResponse = ""

            'Clear data sets
            _dsXmlResponseData = Nothing
            _dtReturnData = Nothing

            'Replace success values with correct ascii values. The return data contains a line feed
            sSuccessValue = sSuccessValue.Replace("@@CHR10", vbLf)

            'SQL query base XML string. This version disables committment control.
            Dim sXMLIN As String = "<?xml version='1.0'?>" & vbCrLf &
                 "<?xml-stylesheet type='text/xsl' href='/DemoXslt.xsl'?>" & vbCrLf &
                    "<script>" & vbCrLf &
                    "<sql>" & vbCrLf &
                    "<options options='noauto' commit='none' error='fast'/>" & vbCrLf &
                    "</sql>" & vbCrLf &
                    "<sql>" & vbCrLf &
                    "<connect conn='myconn' options='noauto' error='fast'/>" & vbCrLf &
                    "</sql>" & vbCrLf &
                    "<sql>" & vbCrLf &
                    "<prepare conn='myconn'>" & sSQL & "</prepare>" & vbCrLf &
                    "</sql>" & vbCrLf &
                    "<sql>" & vbCrLf &
                    "<execute/>" & vbCrLf &
                    "</sql>" & vbCrLf &
                    "<sql>" & vbCrLf &
                    "<describe desc='col'/>" & vbCrLf &
                    "</sql>" & vbCrLf &
                    "<sql>" & vbCrLf &
                    "<fetch block='all' desc='on'/>" & vbCrLf &
                    "</sql>" & vbCrLf &
                    "<sql>" & vbCrLf &
                    "<free/>" & vbCrLf &
                    "</sql>" & vbCrLf &
                    "</script>"

            Dim sXMLOUT As String = _iXmlResponseBufferSize  'Buffer size

            _LastError = ""

            'Removed - Can't block action queries
            'If sSQL.ToUpper.Trim.StartsWith("INSERT") = False And sSQL.ToUpper.Trim.StartsWith("UPDATE") = False And sSQL.ToUpper.Trim.StartsWith("DELETE") = False Then
            'Throw New Exception("Only SQL INSERT, UPDATE and DELETE actions are allowed.")
            'End If

            'Replace parms into XML string
            sdb2parm = sdb2parm.Replace("@@db2value", _DB2Info)
            sdb2parm = sdb2parm.Replace("@@uidvalue", _User)
            sdb2parm = sdb2parm.Replace("@@pwdvalue", _Password)
            sdb2parm = sdb2parm.Replace("@@ipcvalue", _IPCINfo)
            sdb2parm = sdb2parm.Replace("@@ctlvalue", _IPCPersistance)
            sdb2parm = sdb2parm.Replace("@@xmlinvalue", sXMLIN)
            sdb2parm = sdb2parm.Replace("@@xmloutvalue", sXMLOUT)

            'Execute request 
            sRtnXML = ExecuteHttpPostRequest(_BaseURL, "POST", sdb2parm, _HttpTimeout, _UseHttpCredentials)

            'Bail out if HTTPRequest failure
            If sRtnXML.StartsWith("ERROR") Then
                Throw New Exception(sRtnXML)
            End If

            'Save last response
            _LastXMLResponse = sRtnXML

            'Check for SQL action success in XML response data
            If sRtnXML.IndexOf(sSuccessValue) > 0 Then
                rtnexecute = True
            Else
                _LastError = "SQL error occured. Please use GetLastXmlResponse to review the last XML info to determine the cause."
                rtnexecute = False
            End If

            Return rtnexecute

        Catch ex As Exception
            _LastXMLResponse = sRtnXML
            _LastError = ex.Message
            Return False
        End Try
    End Function

    ''' <summary>
    ''' This function queries the DB2 database with selected SQL statement, returns the XML response
    ''' and then loads the internal DataTable object with the returned records.
    ''' The internal results DataTable can be accessed by and of the GetDataTable* methods.
    ''' </summary>
    ''' <param name="sSQL">SQL Select. INSERT, UPDATE and DELETE not allowed </param>
    ''' <param name="sQueryResultOutputFile">Optional PC output file for XML response data. Otherwise data set is created from memory.</param>
    ''' <returns>True- Query service call succeeded, False-Query service call failed</returns>
    ''' <remarks></remarks>
    Public Function ExecuteSqlQuery(sSQL As String, Optional sQueryResultOutputFile As String = "") As Boolean

        Dim sdb2parm As String = _DB2Parm
        Dim sRtnXML As String = ""
        Dim rtnload As Boolean

        Try

            _LastError = ""
            sRtnXML = ""
            _LastHTTPResponse = ""
            _LastXMLResponse = ""

            'Clear data sets
            _dsXmlResponseData = Nothing
            _dtReturnData = Nothing

            'SQL query base XML string
            Dim sXMLIN As String = "<?xml version='1.0'?>" & vbCrLf &
                 "<?xml-stylesheet type='text/xsl' href='/DemoXslt.xsl'?>" & vbCrLf &
                    "<script>" & vbCrLf &
                    "<sql>" & vbCrLf &
                    "<options options='noauto' autocommit='off'/>" & vbCrLf &
                    "</sql>" & vbCrLf &
                    "<sql>" & vbCrLf &
                    "<connect conn='myconn' options='noauto'/>" & vbCrLf &
                    "</sql>" & vbCrLf &
                    "<sql>" & vbCrLf &
                    "<prepare conn='myconn'>" & sSQL.Trim & "</prepare>" & vbCrLf &
                    "</sql>" & vbCrLf &
                    "<sql>" & vbCrLf &
                    "<execute/>" & vbCrLf &
                    "</sql>" & vbCrLf &
                    "<sql>" & vbCrLf &
                    "<describe desc='col'/>" & vbCrLf &
                    "</sql>" & vbCrLf &
                    "<sql>" & vbCrLf &
                    "<fetch block='all' desc='on'/>" & vbCrLf &
                    "</sql>" & vbCrLf &
                    "<sql>" & vbCrLf &
                    "<free/>" & vbCrLf &
                    "</sql>" & vbCrLf &
                    "</script>"

            Dim sXMLOUT As String = _iXmlResponseBufferSize 'Buffer size

            If sSQL.ToUpper.Contains("UPDATE ") Or sSQL.ToUpper.Contains("DELETE ") Or sSQL.ToUpper.Contains("INSERT ") Then
                Throw New Exception("Only SQL selection queries are supported.")
            End If

            If sSQL.ToUpper.StartsWith("UPDATE") Or sSQL.ToUpper.StartsWith("DELETE") Or sSQL.ToUpper.StartsWith("INSERT") Then
                Throw New Exception("SQL statement cannot start with INSERT, UPDATE or DELETE.")
            End If

            If sSQL.ToUpper.StartsWith("SELECT") = False Then
                Throw New Exception("SQL selection must start with SELECT")
            End If

            'Replace parms into XML string
            sdb2parm = sdb2parm.Replace("@@db2value", _DB2Info)
            sdb2parm = sdb2parm.Replace("@@uidvalue", _User)
            sdb2parm = sdb2parm.Replace("@@pwdvalue", _Password)
            sdb2parm = sdb2parm.Replace("@@ipcvalue", _IPCINfo)
            sdb2parm = sdb2parm.Replace("@@ctlvalue", _IPCPersistance)
            sdb2parm = sdb2parm.Replace("@@xmlinvalue", sXMLIN)
            sdb2parm = sdb2parm.Replace("@@xmloutvalue", sXMLOUT)

            'Execute request 
            sRtnXML = ExecuteHttpPostRequest(_BaseURL, "POST", sdb2parm, _HttpTimeout, _UseHttpCredentials)

            'Bail out if HTTPRequest failure
            If sRtnXML.StartsWith("ERROR") Then
                Throw New Exception(sRtnXML)
            End If

            'Save XML results to file and reload XML as a DataTable
            If sQueryResultOutputFile.Trim <> "" Then

                'Write results to output file if specified
                WriteStringToFile(sRtnXML, sQueryResultOutputFile)

                'Load DataTable from XML file
                rtnload = LoadDataTableFromXmlResponseFile(sQueryResultOutputFile)

            Else 'Load XML results to DataTable from memory

                'Load DataTable from XML response
                rtnload = LoadDataTableFromXmlResponseString(sRtnXML)

            End If

            'Save last response
            _LastXMLResponse = sRtnXML

            If rtnload = True Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            _LastXMLResponse = sRtnXML
            _LastError = ex.Message
            Return False
        End Try
    End Function

    ''' <summary>
    ''' This function queries the DB2 database with selected SQL statement and returns results to data table all in one step 
    ''' </summary>
    ''' <param name="sSQL">SQL Select. INSERT, UPDATE and DELETE not allowed </param>
    ''' <param name="sQueryResultOutputFile">Optional PC output file for XML response data. Otherwise data set is created from memory.</param>
    ''' <returns>DataTable with results of query or Nothing</returns>
    ''' <remarks></remarks>
    Public Function ExecuteSqlQueryToDataTable(sSQL As String, Optional sQueryResultOutputFile As String = "") As DataTable

        Dim sdb2parm As String = _DB2Parm
        Dim sRtnXML As String = ""
        Dim rtnquery As Boolean

        Try

            _LastError = ""
            sRtnXML = ""
            _LastHTTPResponse = ""
            _LastXMLResponse = ""

            'Run query and load internal results DataTable
            rtnquery = ExecuteSqlQuery(sSQL, sQueryResultOutputFile)

            'Return results DataTable
            If rtnquery Then
                Return GetQueryResultsDataTable()
            Else
                Throw New Exception("Query failed. Error: " & GetLastError())
            End If

        Catch ex As Exception
            _LastXMLResponse = GetLastXmlResponse()
            _LastError = ex.Message
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' This function queries the DB2 database with selected SQL statement and returns results to XML dataset stream all in one step 
    ''' </summary>
    ''' <param name="sSQL">SQL Select. INSERT, UPDATE and DELETE not allowed </param>
    ''' <param name="sOutputFile">Optional PC output file for XML response data. Otherwise data set is created from memory.</param>
    ''' <returns>XML string</returns>
    ''' <remarks></remarks>
    Public Function ExecuteSqlQueryToXmlString(sSQL As String, Optional sOutputFile As String = "", Optional sTableName As String = "Table1", Optional bWriteSchema As Boolean = False) As String

        Dim sdb2parm As String = _DB2Parm
        Dim sRtnXML As String = ""
        Dim rtnquery As Boolean
        Dim _dt As DataTable

        Try

            _LastError = ""
            sRtnXML = ""
            _LastHTTPResponse = ""
            _LastXMLResponse = ""

            'if table not set, default to Table1
            If sTableName.Trim = "" Then sTableName = "Table1"

            'Run query and load internal results DataTable
            rtnquery = ExecuteSqlQuery(sSQL, sOutputFile)

            'Export DataTable results to XML
            If rtnquery Then
                Dim SB As New StringBuilder
                Dim SW As New IO.StringWriter(SB)
                _dt = GetQueryResultsDataTable()
                _dt.TableName = sTableName
                'Write XMl with or without schema info
                If bWriteSchema Then
                    _dt.WriteXml(SW, Data.XmlWriteMode.WriteSchema)
                Else
                    _dt.WriteXml(SW)
                End If
                sRtnXML = SW.ToString
                SW.Close()
                Return sRtnXML
            Else
                Throw New Exception("Query failed. Error: " & GetLastError())
            End If

        Catch ex As Exception
            _LastXMLResponse = GetLastXmlResponse()
            _LastError = ex.Message
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' This function queries the DB2 database with selected SQL statement and returns results to XML file all in one step 
    ''' </summary>
    ''' <param name="sSQL">SQL Select. INSERT, UPDATE and DELETE not allowed </param>
    ''' <param name="sOutputFile">Optional PC output file for XML response data. Otherwise data set is created from memory.</param>
    ''' <returns>True - Query service call succeeded, False - Query service call failed</returns>
    ''' <remarks></remarks>
    Public Function ExecuteSqlQueryToXmlFile(sSQL As String, sXmlOutputFile As String, Optional replace As Boolean = False, Optional sOutputFile As String = "", Optional sTableName As String = "Table1", Optional bWriteSchema As Boolean = False) As Boolean

        Dim sdb2parm As String = _DB2Parm
        Dim sRtnXML As String = ""
        Dim rtnquery As Boolean

        Try

            _LastError = ""
            sRtnXML = ""
            _LastHTTPResponse = ""
            _LastXMLResponse = ""

            'Run query and load internal results DataTable
            rtnquery = ExecuteSqlQuery(sSQL, sOutputFile)

            'Export results to XML file
            If rtnquery Then
                Return GetQueryResultsDataTableToXmlFile(sXmlOutputFile, sTableName, bWriteSchema, replace)
            Else
                Throw New Exception("Query failed. Error: " & GetLastError())
            End If

        Catch ex As Exception
            _LastXMLResponse = GetLastXmlResponse()
            _LastError = ex.Message
            Return False
        End Try
    End Function

    ''' <summary>
    ''' This function queries the DB2 database with selected SQL statement and returns results to Csv string in one step
    ''' </summary>
    ''' <param name="sSQL">SQL Select. INSERT, UPDATE and DELETE not allowed </param>
    ''' <param name="sQueryResultOutputFile">Optional PC output file for XML response data. Otherwise data set is created from memory.</param>
    ''' <param name="sFieldSepchar">Field delimiter/separator. Default = Comma</param>
    ''' <param name="sFieldDataDelimChar">Field data delimiter character. Default = double quotes.</param>
    ''' <returns>CSV string</returns>
    Public Function ExecuteSqlQueryToCsvString(sSQL As String, Optional sQueryResultOutputFile As String = "", Optional sFieldSepchar As String = ",", Optional sFieldDataDelimChar As String = """") As String

        Dim sdb2parm As String = _DB2Parm
        Dim sRtnXML As String = ""
        Dim rtnquery As Boolean

        Try

            _LastError = ""
            sRtnXML = ""
            _LastHTTPResponse = ""
            _LastXMLResponse = ""

            'Run query and load internal results DataTable
            rtnquery = ExecuteSqlQuery(sSQL, sQueryResultOutputFile)

            'Export results to CSV string
            If rtnquery Then
                Return GetQueryResultsDataTableToCsvString(sFieldSepchar, sFieldDataDelimChar)
            Else
                Throw New Exception("Query failed. Error: " & GetLastError())
            End If

        Catch ex As Exception
            _LastXMLResponse = GetLastXmlResponse()
            _LastError = ex.Message
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' This function queries the DB2 database with selected SQL statement and returns results to Csv file in one step
    ''' </summary>
    ''' <param name="sSQL">SQL Select. INSERT, UPDATE and DELETE not allowed </param>
    ''' <param name="sCsvOutputFile">Output CSV file</param>
    ''' <param name="replace">Replace output file True=Replace file,False=Do not replace</param>
    ''' <param name="sQueryResultOutputFile">Optional PC output file for XML response data. Otherwise data set is created from memory.</param>
    ''' <param name="sFieldSepchar">Field delimiter/separator. Default = Comma</param>
    ''' <param name="sFieldDataDelimChar">Field data delimiter character. Default = double quotes.</param>
    ''' <returns>True-Query service call succeeded, False-Query service call failed</returns>
    Public Function ExecuteSqlQueryToCsvFile(sSQL As String, sCsvOutputFile As String, Optional replace As Boolean = False, Optional sQueryResultOutputFile As String = "", Optional sFieldSepchar As String = ",", Optional sFieldDataDelimChar As String = """") As Boolean

        Dim sdb2parm As String = _DB2Parm
        Dim sRtnXML As String = ""
        Dim rtnquery As Boolean

        Try

            _LastError = ""
            sRtnXML = ""
            _LastHTTPResponse = ""
            _LastXMLResponse = ""

            'Run query and load internal results DataTable
            rtnquery = ExecuteSqlQuery(sSQL, sQueryResultOutputFile)

            'Export results to CSV file
            If rtnquery Then
                Return GetQueryResultsDataTableToCsvFile(sCsvOutputFile, replace, sFieldSepchar, sFieldDataDelimChar)
            Else
                Throw New Exception("Query failed. Error: " & GetLastError())
            End If

        Catch ex As Exception
            _LastXMLResponse = GetLastXmlResponse()
            _LastError = ex.Message
            Return False
        End Try
    End Function


    ''' <summary>
    ''' This function queries the DB2 database with selected SQL statement and returns results to JSON string in one step
    ''' </summary>
    ''' <param name="sSQL">SQL Select. INSERT, UPDATE and DELETE not allowed </param>
    ''' <param name="sQueryResultOutputFile">Optional PC output file for XML response data. Otherwise data set is created from memory.</param>
    ''' <returns>JOSN string</returns>
    Public Function ExecuteSqlQueryToJsonString(sSQL As String, Optional sQueryResultOutputFile As String = "") As String

        Dim sdb2parm As String = _DB2Parm
        Dim sRtnXML As String = ""
        Dim rtnquery As Boolean

        Try

            _LastError = ""
            sRtnXML = ""
            _LastHTTPResponse = ""
            _LastXMLResponse = ""

            'Run query and load internal results DataTable
            rtnquery = ExecuteSqlQuery(sSQL, sQueryResultOutputFile)

            'Export results to CSV string
            If rtnquery Then
                Return GetQueryResultsDataTableToJsonString()
            Else
                Throw New Exception("Query failed. Error: " & GetLastError())
            End If

        Catch ex As Exception
            _LastXMLResponse = GetLastXmlResponse()
            _LastError = ex.Message
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' This function queries the DB2 database with selected SQL statement and returns results to JSON file in one step
    ''' </summary>
    ''' <param name="sSQL">SQL Select. INSERT, UPDATE and DELETE not allowed </param>
    ''' <param name="sJsonOutputFile">Output JSON file</param>
    ''' <param name="replace">Replace output file True=Replace file,False=Do not replace</param>
    ''' <param name="sQueryResultOutputFile">Optional PC output file for XML response data. Otherwise data set is created from memory.</param>
    ''' <returns>True-Query service call succeeded, False-Query service call failed</returns>
    Public Function ExecuteSqlQueryToJsonFile(sSQL As String, sJsonOutputFile As String, Optional replace As Boolean = False, Optional sQueryResultOutputFile As String = "") As Boolean

        Dim sdb2parm As String = _DB2Parm
        Dim sRtnXML As String = ""
        Dim rtnquery As Boolean

        Try

            _LastError = ""
            sRtnXML = ""
            _LastHTTPResponse = ""
            _LastXMLResponse = ""

            'Run query and load internal results DataTable
            rtnquery = ExecuteSqlQuery(sSQL, sQueryResultOutputFile)

            'Export results to JSON file
            If rtnquery Then
                Return GetQueryResultsDataTableToJsonFile(sJsonOutputFile, replace)
            Else
                Throw New Exception("Query failed. Error: " & GetLastError())
            End If

        Catch ex As Exception
            _LastXMLResponse = GetLastXmlResponse()
            _LastError = ex.Message
            Return False
        End Try
    End Function

    ''' <summary>
    ''' This function runs the specified IBM i CL command line. The CL command can be a regular program call or a SBMJOB type of command to submit a job.
    ''' </summary>
    ''' <param name="sCommandString">CL command line to execute</param>
    ''' <returns>True - Command call succeeded, False - Command call failed</returns>
    ''' <remarks></remarks>
    Public Function ExecuteCommand(sCommandString As String) As Boolean

        Dim sdb2parm As String = _DB2Parm
        Dim sRtnXML As String = ""

        'Clear data set
        _dsCommandResponse = Nothing
        _dtCommandResponse = Nothing

        Try

            _LastError = ""
            sRtnXML = ""
            _LastHTTPResponse = ""
            _LastXMLResponse = ""

            'Command call base XML string 
            Dim sXMLIN As String = "<?xml version='1.0'?>" & vbCrLf &
          "<?xml-stylesheet type='text/xsl' href='/DemoXslt.xsl'?>" & vbCrLf &
                    "<script>" & vbCrLf &
                    "<cmd>" & sCommandString & "</cmd>" & vbCrLf &
                    "</script>"

            Dim sXMLOUT As String = _iXmlResponseBufferSize  '"32768" 'Original Buffer size

            'Replace parms into XML string
            sdb2parm = sdb2parm.Replace("@@db2value", _DB2Info)
            sdb2parm = sdb2parm.Replace("@@uidvalue", _User)
            sdb2parm = sdb2parm.Replace("@@pwdvalue", _Password)
            sdb2parm = sdb2parm.Replace("@@ipcvalue", _IPCINfo)
            sdb2parm = sdb2parm.Replace("@@ctlvalue", _IPCPersistance)
            sdb2parm = sdb2parm.Replace("@@xmlinvalue", sXMLIN)
            sdb2parm = sdb2parm.Replace("@@xmloutvalue", sXMLOUT)

            'Execute request
            sRtnXML = ExecuteHttpPostRequest(_BaseURL, "POST", sdb2parm, _HttpTimeout, _UseHttpCredentials)

            'Bail out if +++ success not returned in XML response
            If sRtnXML.Contains("+++ success") Then

                _LastXMLResponse = sRtnXML

                'Load data table with program response parms
                If LoadCmdCallDataTableFromXmlResponseString(GetLastXmlResponse) Then
                    Return True
                Else
                    Throw New Exception("Errors loading command call response XML string.")
                End If

                Return True
            Else
                _LastXMLResponse = sRtnXML

                'Load data table with program response parms
                If LoadCmdCallDataTableFromXmlResponseString(GetLastXmlResponse) Then
                    Return False 'Actual command call failed. Return false
                Else
                    Throw New Exception("Errors loading command call response XML string.")
                End If

                Throw New Exception(sRtnXML)

            End If

        Catch ex As Exception
            _LastXMLResponse = sRtnXML
            _LastError = ex.Message
            Return False
        End Try
    End Function
    ''' <summary>
    ''' This function runs the specified IBM i program call. 
    ''' You will need to pass in an array list of parameters.
    ''' </summary>
    ''' <param name="sProgram">IBM i program name</param>
    ''' <param name="sLibrary">IBM i program library</param>
    ''' <param name="aParmList">IBM i program parm list array</param>
    ''' <returns>True-Program call succeeded, False-Program call failed</returns>
    ''' <remarks></remarks>
    Public Function ExecuteProgram(sProgram As String, sLibrary As String, aParmList As ArrayList) As Boolean

        Dim sdb2parm As String = _DB2Parm
        Dim sRtnXML As String = ""

        Try

            _LastError = ""
            sRtnXML = ""
            _LastHTTPResponse = ""
            _LastXMLResponse = ""

            'Clear data set
            _dsProgramResponse = Nothing
            _dtProgramResponse = Nothing

            'Program call base XML string 
            Dim sXMLIN As String = "<?xml version='1.0'?>" & vbCrLf &
            "<pgm name='@@pgmname' lib='@@pgmlibrary'>" & vbCrLf
            For Each pm As PgmParmList In aParmList
                sXMLIN = sXMLIN & "<parm><data type='" & pm.parmtype & "'>" & pm.parmvalue & "</data></parm>" & vbCrLf
            Next
            sXMLIN = sXMLIN & "</pgm>" & vbCrLf

            Dim sXMLOUT As String = _iXmlResponseBufferSize '"32768" 'Buffer size

            'Replace core program info in program string
            sXMLIN = sXMLIN.Replace("@@pgmname", sProgram)
            sXMLIN = sXMLIN.Replace("@@pgmlibrary", sLibrary)

            'Replace parms into XML string
            sdb2parm = sdb2parm.Replace("@@db2value", _DB2Info)
            sdb2parm = sdb2parm.Replace("@@uidvalue", _User)
            sdb2parm = sdb2parm.Replace("@@pwdvalue", _Password)
            sdb2parm = sdb2parm.Replace("@@ipcvalue", _IPCINfo)
            sdb2parm = sdb2parm.Replace("@@ctlvalue", _IPCPersistance)
            sdb2parm = sdb2parm.Replace("@@xmlinvalue", sXMLIN)
            sdb2parm = sdb2parm.Replace("@@xmloutvalue", sXMLOUT)

            'Execute request
            sRtnXML = ExecuteHttpPostRequest(_BaseURL, "POST", sdb2parm, _HttpTimeout, _UseHttpCredentials)

            'Bail out if +++ success not returned in XML response
            If sRtnXML.Contains("+++ success") Then
                _LastXMLResponse = sRtnXML

                'Load data table with program response parms
                If LoadPgmCallDataTableFromXmlResponseString(GetLastXmlResponse) Then
                    Return True
                Else
                    Throw New Exception("Errors loading program call response XML string.")
                End If
            Else
                Throw New Exception(sRtnXML)
            End If

        Catch ex As Exception
            _LastXMLResponse = sRtnXML
            _LastError = ex.Message
            Return False
        End Try
    End Function
    ''' <summary>
    ''' This function runs the specified IBM i service program subprocedure call. 
    ''' http://www.statususer.org/pdf/20130212Presentation.pdf
    ''' http://174.79.32.155/wiki/pmwiki.php/XMLSERVICE/XMLSERVICESamples
    '''  <pgm name='ZZSRV' lib='XMLSERVICE' func='ZZTIMEUSA'>
    ''' <parm io='both'>
    ''' <data type='8A'>09:45 AM</data>
    ''' </parm>
    ''' <return>
    ''' <data type='8A'>nada</data>
    ''' </return>
    ''' </pgm>
    ''' </summary>
    ''' <param name="sServiceProgram">IBM i service program name</param>
    ''' <param name="sLibrary">IBM i program library</param>
    ''' <param name="sProcedure">IBM i subprocedure</param>
    ''' <param name="aParmList">IBM i program parm list array</param>
    ''' <returns>True-Service program call succeeded, False-Service program failed</returns>
    ''' <remarks></remarks>
    Public Function ExecuteProgramProcedure(sServiceProgram As String, sLibrary As String, sProcedure As String, aParmList As ArrayList, rtnParmList As ArrayList) As Boolean

        Dim sdb2parm As String = _DB2Parm
        Dim sRtnXML As String = ""

        Try

            _LastError = ""
            sRtnXML = ""
            _LastHTTPResponse = ""
            _LastXMLResponse = ""

            'Clear data set
            _dsProgramResponse = Nothing
            _dtProgramResponse = Nothing

            'Program call base XML string 
            Dim sXMLIN As String = "<?xml version='1.0'?>" & vbCrLf &
            "<script>" &
            "<pgm name='@@pgmname' lib='@@pgmlibrary' func='@@pgmfunction'>" & vbCrLf

            For Each pm As ProcedureParmList In aParmList
                sXMLIN = sXMLIN & "<parm io='input'><data type='" & pm.parmtype.Trim & "'>" & pm.parmvalue & "</data></parm>" & vbCrLf
            Next
            'If return parm list passed, set it up
            If rtnParmList.Count > 0 Then
                sXMLIN = sXMLIN & "<return>" & vbCrLf
                For Each pm As ProcedureParmList In rtnParmList
                    sXMLIN = sXMLIN & "<data type='" & pm.parmtype.Trim & "'>" & pm.parmvalue & "</data>" & vbCrLf
                Next
                sXMLIN = sXMLIN & "</return>"
            End If
            sXMLIN = sXMLIN & "</pgm>" & vbCrLf
            sXMLIN = sXMLIN & "</script>" & vbCrLf

            Dim sXMLOUT As String = _iXmlResponseBufferSize '"32768" 'Buffer size

            'Replace core program info in program string
            sXMLIN = sXMLIN.Replace("@@pgmname", sServiceProgram)
            sXMLIN = sXMLIN.Replace("@@pgmlibrary", sLibrary)
            sXMLIN = sXMLIN.Replace("@@pgmfunction", sProcedure)

            '' ''TEST subprocedure call
            'sXMLIN = "<?xml version='1.0'?>" & vbCrLf & _
            '"<script>"
            ''sXMLIN = ""
            '''sXMLIN = "<cmd comment='addlible'>ADDLIBLE LIB(XMLDOTNETI) POSITION(*FIRST)</cmd>" & vbCrLf
            'sXMLIN = sXMLIN & "<pgm name='XMLSVC001' lib='XMLDOTNETI' func='SAMPLE1'>" & vbCrLf
            'sXMLIN = sXMLIN & "<parm io='input'>" & vbCrLf
            ''sXMLIN = sXMLIN & "<data type='65535A' varying='on'>Test</data>" & vbCrLf
            'sXMLIN = sXMLIN & "<data type='200A'>Test</data>" & vbCrLf
            'sXMLIN = sXMLIN & "</parm>" & vbCrLf
            'sXMLIN = sXMLIN & "<return>" & vbCrLf
            'sXMLIN = sXMLIN & "<data type='200A'></data>" & vbCrLf
            'sXMLIN = sXMLIN & "</return>" & vbCrLf
            'sXMLIN = sXMLIN & "</pgm>" & vbCrLf
            'sXMLIN = sXMLIN & "</script>" & vbCrLf

            'Replace parms into XML string
            sdb2parm = sdb2parm.Replace("@@db2value", _DB2Info)
            sdb2parm = sdb2parm.Replace("@@uidvalue", _User)
            sdb2parm = sdb2parm.Replace("@@pwdvalue", _Password)
            sdb2parm = sdb2parm.Replace("@@ipcvalue", _IPCINfo)
            sdb2parm = sdb2parm.Replace("@@ctlvalue", _IPCPersistance)
            sdb2parm = sdb2parm.Replace("@@xmlinvalue", sXMLIN)
            sdb2parm = sdb2parm.Replace("@@xmloutvalue", sXMLOUT)

            'Execute request
            sRtnXML = ExecuteHttpPostRequest(_BaseURL, "POST", sdb2parm, _HttpTimeout, _UseHttpCredentials)

            'Bail out if +++ success not returned in XML response
            If sRtnXML.Contains("+++ success") Then
                _LastXMLResponse = sRtnXML

                'Load data table with program response parms
                If LoadPgmCallDataTableFromXmlResponseString(GetLastXmlResponse) Then
                    Return True
                Else
                    Throw New Exception("Errors loading program call response XML string.")
                End If

                Return True
            Else
                Throw New Exception(sRtnXML)
            End If

            _LastXMLResponse = sRtnXML
            Return True

        Catch ex As Exception
            _LastXMLResponse = sRtnXML
            _LastError = ex.Message
            Return False
        End Try
    End Function

#End Region

#Region "DataQueueMethods"

    ''' <summary>
    ''' Send data to sequential data queue. Data queue must already exist.
    ''' </summary>
    ''' <param name="sDataQueue">Data queue name</param>
    ''' <param name="sDataQueueLibrary">Data queue library</param>
    ''' <param name="sData">String data to send</param>
    ''' <param name="iDataLength">Length of data to send. Default=0 which will auto-calculate length.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SendDataQueueSequential(sDataQueue As String, sDataQueueLibrary As String, sData As String, Optional iDataLength As Integer = 0) As Boolean

        Dim rtnbool As Boolean

        Try

            'Execute the program with parms
            Dim arrParms As New ArrayList
            Dim pDqName As New XmlServicei.PgmParmList
            Dim pDqLib As New XmlServicei.PgmParmList
            Dim pLength As New XmlServicei.PgmParmList
            Dim pData As New XmlServicei.PgmParmList
            Dim pWaitTime As New XmlServicei.PgmParmList

            'Reset last data queue operation info
            _LastError = ""
            _LastDataQueueName = sDataQueue
            _LastDataQueueLibrary = sDataQueueLibrary
            _LastDataQueueLengthReceived = iDataLength
            _LastDataQueueDataReceived = ""
            _LastDataQueueDataSent = ""

            'Set data queue
            pDqName.parmtype = "10A"
            pDqName.parmvalue = sDataQueue
            'Set data queue library
            pDqLib.parmtype = "10A"
            pDqLib.parmvalue = sDataQueueLibrary
            'Set length if passed
            pLength.parmtype = "5p0"
            If iDataLength > 0 Then
                pLength.parmvalue = iDataLength
            Else
                pLength.parmvalue = sData.TrimEnd.Length
            End If
            'Set data to be passed
            pData.parmtype = sData.TrimEnd.Length & "A"
            pData.parmvalue = sData.TrimEnd
            'Add parms to parm array
            arrParms.Add(pDqName)
            arrParms.Add(pDqLib)
            arrParms.Add(pLength)
            arrParms.Add(pData)

            'Call program to send the data queue entry
            rtnbool = ExecuteProgram("QSNDDTAQ", "QSYS", arrParms)

            If rtnbool = True Then
                _LastError = "Send to data queue was Successful. XML Response: " & vbCrLf & GetLastXmlResponse()

                'Set last data queue operation info
                _LastDataQueueDataSent = pData.parmvalue

                Return True
            Else
                _LastError = "Send to data queue failed. XML Response: " & vbCrLf & GetLastXmlResponse()
                Return False
            End If

        Catch ex As Exception
            _LastError = "Error occurred sending to data queue. Error: " & ex.Message
            Return False
        End Try


    End Function
    ''' <summary>
    ''' Receive data from sequential data queue. Data queue must already exist.
    ''' </summary>
    ''' <param name="sDataQueue">Data queue name</param>
    ''' <param name="sDataQueueLibrary">Data queue library</param>
    ''' <param name="iDataLength">Length of data to send. Default=0 which will auto-calculate length.</param>
    ''' <param name="iWaitTime">Length of time to wait for a response. Default=10 seconds.</param>
    ''' <returns>Returns data or string starting with"EXCEPTION" if errors occurred</returns>
    ''' <remarks></remarks>
    Public Function ReceiveDataQueueSequential(sDataQueue As String, sDataQueueLibrary As String, iDataLength As Integer, Optional iWaitTime As Integer = 10) As String

        Dim rtnbool As Boolean
        Dim sData As String = ""

        Try

            'Execute the program with parms
            Dim arrParms As New ArrayList
            Dim pDqName As New XmlServicei.PgmParmList
            Dim pDqLib As New XmlServicei.PgmParmList
            Dim pLength As New XmlServicei.PgmParmList
            Dim pData As New XmlServicei.PgmParmList
            Dim pWaitTime As New XmlServicei.PgmParmList

            'Reset last data queue operation info
            _LastError = ""
            _LastDataQueueName = sDataQueue
            _LastDataQueueLibrary = sDataQueueLibrary
            _LastDataQueueLengthReceived = iDataLength
            _LastDataQueueDataReceived = ""
            _LastDataQueueDataSent = ""

            pDqName.parmtype = "10A"
            pDqName.parmvalue = sDataQueue
            pDqLib.parmtype = "10A"
            pDqLib.parmvalue = sDataQueueLibrary
            pLength.parmtype = "5p0"
            pLength.parmvalue = iDataLength
            pData.parmtype = iDataLength & "A"
            pData.parmvalue = sData.TrimEnd
            pWaitTime.parmtype = "5p0"
            pWaitTime.parmvalue = iWaitTime
            arrParms.Add(pDqName)
            arrParms.Add(pDqLib)
            arrParms.Add(pLength)
            arrParms.Add(pData)
            arrParms.Add(pWaitTime)

            'Call program to receive the data queue entry
            rtnbool = ExecuteProgram("QRCVDTAQ", "QSYS", arrParms)

            If rtnbool = True Then
                _LastError = "Receive from data queue was Successful. XML Response: " & vbCrLf & GetLastXmlResponse()
            Else
                _LastError = "Receive from data queue failed. XML Response: " & vbCrLf & GetLastXmlResponse()
            End If

            'Get return parms to datatable and extract return data
            If LoadPgmCallDataTableFromXmlResponseString(GetLastXmlResponse) Then

                'Make sure data returned
                If _dtProgramResponse.Rows.Count <> 5 Then
                    Throw New Exception("It appears no parms were returned from receive data queue call.")
                End If

                Dim _drParm As DataRow = Nothing

                'Parm2 data row contains length of returned data
                _drParm = _dtProgramResponse.Rows(2)
                _LastDataQueueLengthReceived = Convert.ToInt32(_drParm(1))

                'Parm3 data row contains returned data
                _drParm = _dtProgramResponse.Rows(3)

                'Return data queue return data value
                If _drParm(0).ToString = "Parm3" Then

                    'Set last data queue operation info
                    _LastDataQueueName = sDataQueue
                    _LastDataQueueLibrary = sDataQueueLibrary
                    _LastDataQueueDataReceived = _drParm(1)
                    Return _drParm(1)
                Else
                    Throw New Exception("It appears that Parm3 - data queue value is missing.")
                End If
            Else
                Throw New Exception("Errors loading program call response XML string.")
            End If
        Catch ex As Exception
            _LastError = "Error occurred receiving from data queue. Error: " & ex.Message
            Return "EXCEPTION " & _LastError
        End Try

    End Function
    ''' <summary>
    ''' Get last data queue name
    ''' </summary>
    ''' <returns>Last data queue name used</returns>
    ''' <remarks></remarks>
    Public Function GetLastDataQueueName() As String
        Try
            Return _LastDataQueueName
        Catch ex As Exception
            Return ""
        End Try
    End Function
    ''' <summary>
    ''' Get last data queue library
    ''' </summary>
    ''' <returns>Last data queue library used</returns>
    ''' <remarks></remarks>
    Public Function GetLastDataQueueLibrary() As String
        Try
            Return _LastDataQueueLibrary
        Catch ex As Exception
            Return ""
        End Try
    End Function
    ''' <summary>
    ''' Get last data buffer sent to data queue
    ''' </summary>
    ''' <returns>Last data buffer sent to data queue.</returns>
    ''' <remarks></remarks>
    Public Function GetLastDataQueueDataSent() As String
        Try
            Return _LastDataQueueDataSent
        Catch ex As Exception
            Return ""
        End Try
    End Function
    ''' <summary>
    ''' Get last data buffer received from data queue
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Last data buffer received from data queue.</remarks>
    Public Function GetLastDataQueueDataReceived() As String
        Try
            Return _LastDataQueueDataReceived
        Catch ex As Exception
            Return ""
        End Try
    End Function
    ''' <summary>
    ''' Get last data queue length received
    ''' </summary>
    ''' <returns>Length of data returned or 0 if none.</returns>
    ''' <remarks></remarks>
    Public Function GetLastDataQueueLengthReceived() As Integer
        Try
            Return _LastDataQueueLengthReceived
        Catch ex As Exception
            Return 0
        End Try
    End Function

#End Region

#Region "MiscHelperMethods"

    ''' <summary>
    ''' This kills the currently active IPC service instance on IBM i system.
    ''' These jobs are all marked as XTOOLKIT and can be seen by running 
    ''' the following CL command: WRKACTJOB JOB(XTOOLKIT)
    ''' Note: If you will be doing a lot of work, you can leave the job instantiated. Otherwise kill the XTOOLKIT job 
    ''' if you're done using it. 
    ''' </summary>
    ''' <returns>True - kill succeeded. Return XML will always be blank so we will always return true most likely.</returns>
    ''' <remarks></remarks>
    Public Function KillService() As Boolean

        Dim sdb2parm As String = _DB2Parm
        Dim sRtnXML As String = ""
        Dim rtnload As Boolean = False

        Try

            'Kill service call base XML string 
            Dim sXMLIN As String = "<?xml version='1.0'?>" & vbCrLf &
          "<?xml-stylesheet type='text/xsl' href='/DemoXslt.xsl'?>" & vbCrLf &
                    "<script>" & vbCrLf &
                    "</script>"

            Dim sXMLOUT As String = _iXmlResponseBufferSize

            _LastError = ""

            sdb2parm = sdb2parm.Replace("@@db2value", _DB2Info)
            sdb2parm = sdb2parm.Replace("@@uidvalue", _User)
            sdb2parm = sdb2parm.Replace("@@pwdvalue", _Password)
            sdb2parm = sdb2parm.Replace("@@ipcvalue", _IPCINfo)
            sdb2parm = sdb2parm.Replace("@@ctlvalue", "*immed")
            sdb2parm = sdb2parm.Replace("@@xmlinvalue", sXMLIN)
            sdb2parm = sdb2parm.Replace("@@xmloutvalue", sXMLOUT)
            sdb2parm = sdb2parm & "&submit=*immed end (kill job)"

            'Execute request
            sRtnXML = ExecuteHttpPostRequest(_BaseURL, "POST", sdb2parm, _HttpTimeout, _UseHttpCredentials)

            'Bail out if HTTPRequest failure
            If sRtnXML.StartsWith("ERROR") Then
                Throw New Exception(sRtnXML)
            End If

            'Execute XMLSERVICE POST request to run command
            sRtnXML = ExecuteHttpPostRequest(_BaseURL, "POST", sdb2parm, _HttpTimeout, _UseHttpCredentials)

            'Bail out if +++ success not returned in XML response
            If sRtnXML = "" Then
                _LastXMLResponse = sRtnXML
                Return True
            Else
                Throw New Exception(sRtnXML)
            End If

        Catch ex As Exception
            _LastXMLResponse = sRtnXML
            _LastError = ex.Message
            Return False
        End Try
    End Function
    ''' <summary>
    ''' This kills the temporary IPC directories in IFS directory /tmp or other specified location
    ''' It's a good idea to periodically clean up these directories either via this function or
    ''' via a job on IBM i that does a RMDIR.  Ex: RMDIR  '/tmp/xmlservicei*'
    ''' </summary>
    ''' <param name="sIfsDirPath">IFS directory path to remove. Wildcards are OK.  Ex: /tmp/xmlservicei*'</param>
    ''' <returns>True - IFS directory kill succeeded. False - IFS directory kill failed.</returns>
    ''' <remarks></remarks>
    Public Function KillIpcIfsDirectories(sIfsDirPath As String) As Boolean

        Dim rtn As Boolean = False

        Try

            _LastXMLResponse = ""
            _LastError = ""
            _LastHTTPResponse = ""

            'Erase directories if path passed
            If sIfsDirPath.Trim <> "" Then
                rtn = ExecuteCommand("RMDIR DIR('" & sIfsDirPath.Trim & "')")
            Else
                Throw New Exception("No IFS directory path specified.")
            End If

            Return rtn

        Catch ex As Exception
            '_LastXMLResponse = ' This value should already be set
            _LastError = ex.Message
            Return False
        End Try
    End Function



    ''' <summary>
    ''' Write text string to file
    ''' </summary>
    ''' <param name="sTextString">Text string value</param>
    ''' <param name="sOutputFile">Output file</param>
    ''' <param name="bAppend">True=Append</param>
    ''' <param name="bReplace">True=Replace file before writing</param>
    ''' <returns>True-Success, False-Failure to write</returns>
    ''' <remarks></remarks>
    Public Function WriteStringToFile(ByVal sTextString As String, ByVal sOutputFile As String, Optional bAppend As Boolean = False, Optional bReplace As Boolean = True) As Boolean
        'Write text string to output file
        Try
            _LastError = ""

            If System.IO.File.Exists(sOutputFile) = True Then
                If bReplace = True Then
                    System.IO.File.Delete(sOutputFile)
                End If
            End If

            Using oWriter As New System.IO.StreamWriter(sOutputFile, bAppend)

                oWriter.Write(sTextString)
                oWriter.Flush()
                oWriter.Close()

            End Using

            Return True

        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Read text file contents and return as string
    ''' </summary>
    ''' <param name="sInputFile">Input file name</param>
    ''' <returns>Contents of file as string or blanks on error</returns>
    ''' <remarks></remarks>
    Public Function ReadTextFile(ByVal sInputFile As String) As String
        Try

            _LastError = ""

            If IO.File.Exists(sInputFile) = False Then
                Throw New Exception(sInputFile & " does not exist.")
            End If

            Using oReader As New System.IO.StreamReader(sInputFile, True)
                Dim sWork As String = ""

                'Read all text
                sWork = oReader.ReadToEnd()

                oReader.Close()

                Return sWork

            End Using
        Catch ex As Exception
            _LastError = ex.Message
            Return ""
        End Try
    End Function

#End Region

#Region "IbmiHelperMethods"
    ''' <summary>
    ''' Create library for temp files. Library name is TMP
    ''' </summary>
    ''' <returns>True-Library created, False-Library exists or library not created.</returns>   
    Function CreateTempLibrary() As Boolean
        Try
            If CheckObjectExists("TMP", "QSYS", "*LIB") = False Then
                Return ExecuteCommand("CRTLIB LIB(TMP) TYPE(*PROD) TEXT('Temp Files') AUT(*ALL) CRTAUT(*ALL)")
            Else
                Throw New Exception("Library TMP already exists.")
            End If
        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Create IBM i library if it does not exist
    ''' </summary> 
    ''' <param name="sLibraryName">Library name</param>
    ''' <param name="sLibraryType">Library type. Default - *PROD</param>
    ''' <param name="sLibraryText">Library text description</param>
    ''' <param name="sLibraryAut">Library authority. Default - *ALL</param>
    ''' <param name="sLibraryCrtAut">Create authority. Default - *ALL</param>
    ''' <returns>True-Library created, False-Library exists or library not created.</returns>   
    Function CreateLibrary(sLibraryName As String, Optional sLibraryType As String = "*PROD", Optional sLibraryText As String = "", Optional sLibraryAut As String = "*ALL", Optional sLibraryCrtAut As String = "*ALL") As Boolean
        Try
            Dim sCmd As String = String.Format("CRTLIB LIB({0}) TYPE({1}) TEXT('{2}') AUT({3}) CRTAUT({4})", sLibraryName.Trim.ToUpper, sLibraryType.Trim.ToUpper, sLibraryText.Trim, sLibraryAut.Trim.ToUpper, sLibraryCrtAut.Trim.ToUpper)

            If CheckObjectExists(sLibraryName.Trim.ToUpper, "QSYS", "*LIB") = False Then
                Return ExecuteCommand(sCmd)
            Else
                Throw New Exception(String.Format("Library {0} already exists.", sLibraryName.Trim.ToUpper))
            End If
        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Check for IBM i object existence
    ''' </summary>
    ''' <param name="sObjectName">Object name</param>
    ''' <param name="sObjectLibrary">Object library name</param>
    ''' <param name="sObjectType">Object type.</param>
    ''' <returns>True-Object exists,False-Object does not exist.</returns>
    Function CheckObjectExists(sObjectName As String, sObjectLibrary As String, sObjectType As String) As Boolean
        Try

            If sObjectName.Trim = "" Or sObjectLibrary.Trim = "" Or sObjectType.Trim = "" Then
                Throw New Exception("Object name, library and type are all required fields.")
            End If

            Dim sCmd As String = String.Format("CHKOBJ OBJ({0}/{1}) OBJTYPE({2}) MBR(*NONE) AUT(*NONE)", sObjectLibrary.Trim.ToUpper, sObjectName.Trim.ToUpper, sObjectType.Trim.ToUpper)
            Return ExecuteCommand(sCmd)
        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Delete IBM i library if it exists
    ''' </summary> 
    ''' <param name="sLibraryName">Library name</param>
    ''' <returns>True-Library deleted, False-Library does not exists or library not deleted.</returns>   
    Function DeleteLibrary(sLibraryName As String) As Boolean
        Try
            Dim sCmd As String = String.Format("DLTLIB LIB({0})", sLibraryName.Trim.ToUpper)

            If CheckObjectExists(sLibraryName.Trim.ToUpper, "QSYS", "*LIB") = True Then
                Return ExecuteCommand(sCmd)
            Else
                Throw New Exception(String.Format("Library {0} does not exist.", sLibraryName.Trim.ToUpper))
            End If
        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Clear IBM i library if it exists
    ''' </summary> 
    ''' <param name="sLibraryName">Library name</param>
    ''' <returns>True-Library cleared, False-Library does not exists or library not cleared.</returns>   
    Function ClearLibrary(sLibraryName As String) As Boolean
        Try
            Dim sCmd As String = String.Format("CLRLIB LIB({0})", sLibraryName.Trim.ToUpper)

            If CheckObjectExists(sLibraryName.Trim.ToUpper, "QSYS", "*LIB") = True Then
                Return ExecuteCommand(sCmd)
            Else
                Throw New Exception(String.Format("Library {0} does not exist.", sLibraryName.Trim.ToUpper))
            End If
        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Create fixed length flat physical file
    ''' </summary> 
    ''' <param name="sFileName">File name</param>
    ''' <param name="sFileLibrary">File library</param>
    ''' <param name="sLibraryTextDescription">File description</param>
    ''' <param name="iRecordLength">Record length. Default=400</param>
    ''' <returns>True-Success, False-Error</returns>
    Function CreatePhysicalFileFixed(sFileName As String, sFileLibrary As String, Optional sLibraryTextDescription As String = "", Optional iRecordLength As Integer = 400) As Boolean
        Try
            Dim sCmd As String = String.Format(" CRTPF FILE({0}/{1}) RCDLEN({2}) TEXT('{3}') OPTION(*NOSRC *NOLIST) MAXMBRS(*NOMAX) SIZE(*NOMAX)", sFileLibrary.Trim.ToUpper, sFileName.Trim.ToUpper, iRecordLength, sLibraryTextDescription.Trim)

            If CheckObjectExists(sFileName.Trim.ToUpper, sFileLibrary.Trim.ToUpper, "*FILE") = False Then
                Return ExecuteCommand(sCmd)
            Else
                Throw New Exception(String.Format("File {0} in library {1} already exists.", sFileName.Trim.ToUpper, sFileLibrary.Trim.ToUpper))
            End If
        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Create fixed length flat physical file using SQL with long or short name. 
    ''' Data field name will be RECORD.
    ''' </summary> 
    ''' <param name="sTableName">SQL table name. Up to 30 characters.</param>
    ''' <param name="sTableLibrary">SQL table schema/library</param>
    ''' <param name="iRecordLength">Record length. Default=400</param>
    ''' <returns>True-Success, False-Error</returns>
    Function CreateSqlTableFixed(sTableName As String, sTableLibrary As String, Optional iRecordLength As Integer = 400) As Boolean

        Dim rtncmd As Boolean = True

        Try

            'Table creation command
            Dim sSqlCmd As String = String.Format("CREATE TABLE {0}/{1} (RECORD CHAR ({2}) NOT NULL WITH DEFAULT)", sTableLibrary.Trim.ToUpper, sTableName.Trim.ToUpper, iRecordLength)

            'Call SQL create command
            rtncmd = ExecuteSqlNonQuery(sSqlCmd)

            If rtncmd Then
                _LastError = String.Format("Table {0} was created in library {1}.", sTableName.Trim.ToUpper, sTableLibrary.Trim.ToUpper)
            Else
                _LastError = String.Format("Errors occurred. It's possible table {0} in library {1} already exists.", sTableName.Trim.ToUpper, sTableLibrary.Trim.ToUpper)
            End If

            Return rtncmd

        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Insert record into fixed length flat physical file using SQL.
    ''' Data field name will be RECORD.
    ''' </summary> 
    ''' <param name="sTableName">SQL table name. Up to 30 characters.</param>
    ''' <param name="sTableLibrary">SQL table schema/library</param>
    ''' <param name="sRecordData">Single field data record</param>
    ''' <returns>True-Success, False-Error</returns>
    Function InsertSqlTableFixed(sTableName As String, sTableLibrary As String, sRecordData As String) As Boolean

        Dim rtncmd As Boolean = True

        Try

            'Double up any single quoted records
            sRecordData = sRecordData.Replace("'", "''")

            Dim sSqlCmd As String = String.Format("INSERT INTO {0}/{1} (RECORD) VALUES('{2}')", sTableLibrary.Trim.ToUpper, sTableName.Trim.ToUpper, sRecordData)

            'Call SQL command
            rtncmd = ExecuteSqlNonQuery(sSqlCmd)

            If rtncmd Then
                _LastError = String.Format("Record inserted to Table {0} in library {1}.", sTableName.Trim.ToUpper, sTableLibrary.Trim.ToUpper)
            Else
                _LastError = String.Format("Errors occurred. It's possible table {0} in library {1} does not exist or there were unpaired single quotes.", sTableName.Trim.ToUpper, sTableLibrary.Trim.ToUpper)
            End If

            Return rtncmd

        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Delete table using SQL DROP TABLE action
    ''' </summary> 
    ''' <param name="sTableName">SQL table name. Up to 30 characters.</param>
    ''' <param name="sTableLibrary">SQL table schema/library</param>
    ''' <returns>True-Success, False-Error</returns>
    Function DeleteSqlTable(sTableName As String, sTableLibrary As String) As Boolean

        Dim rtncmd As Boolean = True

        Try

            'SQL command
            Dim sSqlCmd As String = String.Format("DROP TABLE {0}/{1}", sTableLibrary.Trim.ToUpper, sTableName.Trim.ToUpper)

            'Call SQL command
            rtncmd = ExecuteSqlNonQuery(sSqlCmd)

            If rtncmd Then
                _LastError = String.Format("Table {0} was deleted from library {1}.", sTableName.Trim.ToUpper, sTableLibrary.Trim.ToUpper)
            Else
                _LastError = String.Format("Errors occurred. It's possible table {0} in library {1} does not exist.", sTableName.Trim.ToUpper, sTableLibrary.Trim.ToUpper)
            End If

            Return rtncmd

        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Clear table by Deleting all records from table using SQL DELETE action
    ''' </summary> 
    ''' <param name="sTableName">SQL table name. Up to 30 characters.</param>
    ''' <param name="sTableLibrary">SQL table schema/library</param>
    ''' <returns>True-Success, False-Error</returns>
    Function ClearSqlTable(sTableName As String, sTableLibrary As String) As Boolean

        Dim rtncmd As Boolean = True

        Try

            'SQL command
            Dim sSqlCmd As String = String.Format("DELETE FROM {0}/{1}", sTableLibrary.Trim.ToUpper, sTableName.Trim.ToUpper)

            'Call SQL command
            rtncmd = ExecuteSqlNonQuery(sSqlCmd)

            If rtncmd Then
                _LastError = String.Format("Records were deleted from Table {0} in library {1}.", sTableName.Trim.ToUpper, sTableLibrary.Trim.ToUpper)
            Else
                _LastError = String.Format("Errors occurred. It's possible table {0} in library {1} does not exist.", sTableName.Trim.ToUpper, sTableLibrary.Trim.ToUpper)
            End If

            Return rtncmd

        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Does SQL table exist ? We check SYSTABLES in QSYS2 for table existence.
    ''' Note: Only works for DDS defined or SQL defined tables. Flat files created with CRTPF will not show up.
    ''' </summary> 
    ''' <param name="sTableName">SQL table name. Up to 30 characters.</param>
    ''' <param name="sTableLibrary">SQL table schema/library</param>
    ''' <returns>True-Exists, False-Does not exist or rrror</returns>
    Function CheckSqlTableExists(sTableName As String, sTableLibrary As String) As Boolean

        Dim rtncmd As Boolean = True
        Dim dtWork As DataTable
        Dim sql As String = ""

        Try

            'Build table check for SQL table
            sql = String.Format("SELECT COUNT(*) as TABLECOUNT From QSYS2/SYSTABLES WHERE TABLE_SCHEMA='{0}' and TABLE_NAME='{1}'", sTableLibrary.Trim, sTableName.Trim)

            'Run the table check query
            dtWork = ExecuteSqlQueryToDataTable(sql)

            If dtWork Is Nothing Then
                Throw New Exception("SQL error occurred.")
            Else
                'Should only ever get a single count result row
                If dtWork.Rows.Count = 1 Then
                    'Check the count to see if we found the table usinf SYSTABLES
                    If Convert.ToInt32(dtWork.Rows(0).Item("TABLECOUNT")) > 0 Then
                        _LastError = String.Format("Table {0} in library {1} exists.", sTableName.Trim.ToUpper, sTableLibrary.Trim.ToUpper)
                        Return True
                    Else
                        _LastError = String.Format("Table {0} in library {1} does not exist or is possibly a flat file created with CRTPF so not in SYSTABLES.", sTableName.Trim.ToUpper, sTableLibrary.Trim.ToUpper)
                        Return False
                    End If
                Else
                    _LastError = String.Format("Error occurred. Only 1 count row expected.", sTableName.Trim.ToUpper, sTableLibrary.Trim.ToUpper)
                    Return False
                End If
            End If

        Catch ex As Exception
            _LastError = ex.Message
            Return False
        End Try
    End Function
#End Region

#Region "PrivateMethods"


    ''' <summary>
    ''' Make an XMLSERVICE HTTP POST request with selected URL and get response
    ''' </summary>
    ''' <param name="URL">URL where XMLSERVICE is set up.</param>
    ''' <param name="method">POST or GET</param>
    ''' <param name="POSTdata">Data to post ont he request</param>
    ''' <param name="iTimeout">Optional HTTP request timeout. Default = 60000 milliseconds</param>
    ''' <param name="iUseHttpCredentials">Optional Use network credentials for web server auth. 0=No, 1=Yes Default = 0</param>
    ''' <returns>XML response or error string starting with "ERROR" </returns>
    ''' <remarks></remarks>
    Private Function ExecuteHttpPostRequest(URL As String, method As String, POSTdata As String, Optional iTimeout As Integer = 60000, Optional iUseHttpCredentials As Boolean = False) As String
        Dim responseData As String = ""
        Try
            _LastHTTPResponse = ""
            Dim hwrequest As Net.HttpWebRequest = Net.WebRequest.Create(URL)
            hwrequest.Accept = "*/*"
            hwrequest.AllowAutoRedirect = True
            'If specific use Http web server authentication
            If iUseHttpCredentials Then
                hwrequest.Credentials = New NetworkCredential(_HttpUser, _HttpPassword)
            End If
            hwrequest.UserAgent = "XmlServicei/0.1"
            hwrequest.Timeout = iTimeout
            hwrequest.Method = method
            If hwrequest.Method = "POST" Then
                hwrequest.ContentType = "application/x-www-form-urlencoded"

                'Percent encoding fix to resolve issues with CL commands or SQL statements containing % signs
                'We convert pct signs to %25 before encoding the URL.
                'This should compensate for XMLCGI not handling percent signs in calls - 8/14/2016
                POSTdata = POSTdata.Replace("%", "%25")
                'Now encode the URL post data before posting - 8/14/2016
                POSTdata = EncodeUrl(POSTdata)

                Dim encoding As New Text.ASCIIEncoding() 'Use UTF8Encoding for XML requests
                Dim postByteArray() As Byte = encoding.GetBytes(POSTdata)
                hwrequest.ContentLength = postByteArray.Length
                Dim postStream As IO.Stream = hwrequest.GetRequestStream()
                postStream.Write(postByteArray, 0, postByteArray.Length)
                postStream.Close()
            End If
            Dim hwresponse As Net.HttpWebResponse = hwrequest.GetResponse()
            _LastHTTPResponse = hwresponse.StatusCode & " " & hwresponse.StatusDescription
            If hwresponse.StatusCode = Net.HttpStatusCode.OK Then
                Dim responseStream As IO.StreamReader =
                  New IO.StreamReader(hwresponse.GetResponseStream())
                responseData = responseStream.ReadToEnd()
            End If
            hwresponse.Close()
        Catch e As Exception
            responseData = "ERROR - An HTTP error occurred: " & e.Message
        End Try
        Return responseData

    End Function

    ''' <summary>
    ''' Encode URL string
    ''' </summary>
    ''' <param name="sURL">URL to encode</param>
    ''' <returns>Encoded URL string</returns>
    ''' <remarks></remarks>
    Private Function EncodeUrl(sURL As String) As String
        Try
            _LastError = ""
            sURL = System.Web.HttpUtility.UrlEncode(sURL)
            Return sURL
        Catch ex As Exception
            _LastError = ex.Message
            'Return original URL
            Return sURL
        End Try
    End Function

    ''' <summary>
    ''' This function loads an XML file that contains a data stream returned from the XMLSERVICE service program
    ''' into an internal XML response DataSet that can be retreived via the GetDataSet function.
    ''' The results are processed into a DataTable which can then be 
    ''' </summary>
    ''' <param name="sXMLFile">XML File</param>
    ''' <returns>True-successfully loaded XML file, False-failed to load XML file</returns>
    ''' <remarks></remarks>
    Private Function LoadDataTableFromXmlResponseFile(sXMLFile As String) As Boolean

        Try

            _LastError = ""

            'Bail if no XML file
            If System.IO.File.Exists(sXMLFile) = False Then
                Throw New Exception("XML file " & sXMLFile & " does not exist. Process cancelled.")
            End If

            ' Load XML data into a dataset  
            _dsXmlResponseData = New DataSet()
            _dsXmlResponseData.ReadXml(sXMLFile)

            'Extract returned SQL column definitions from XML file
            _dtColumnDefinitions = _dsXmlResponseData.Tables("col")
            _iColumnCount = _dsXmlResponseData.Tables("col").Rows.Count
            _iRowCount = _dsXmlResponseData.Tables("data").Rows.Count / _iColumnCount

            'Start new DataTable for XML resultset returned by XMLSERVICE from IBM i
            _dtReturnData = New DataTable
            For Each dr1 As DataRow In _dsXmlResponseData.Tables("col").Rows
                _dtReturnData.Columns.Add(dr1(0), Type.GetType("System.String"))
            Next

            Dim colct As Integer
            Dim row2 As DataRow = Nothing
            For Each dr1 As DataRow In _dsXmlResponseData.Tables("data").Rows

                'If first column field, create new row
                If colct = 0 Then
                    row2 = _dtReturnData.NewRow
                ElseIf colct = _iColumnCount Then
                    'All 
                    _dtReturnData.Rows.Add(row2)
                    row2 = _dtReturnData.NewRow
                    colct = 0
                End If

                'Set column value for row
                row2(dr1(0).ToString) = dr1(1).ToString
                colct = colct + 1
            Next

            'Add last row to dataset
            _dtReturnData.Rows.Add(row2)

            _LastError = _iRowCount & " rows were returned from XML file " & sXMLFile

            _bXMLIsLoaded = True
            Return True

        Catch ex As Exception
            _LastError = ex.Message
            _bXMLIsLoaded = False
            Return False
        End Try
    End Function
    ''' <summary>
    ''' This function loads an XML string that contains a data stream returned from the XMLSERVICE service program.
    ''' The XML data is loaded into a DataSet and then the data is loaded into a DataTable.
    ''' </summary>
    ''' <param name="sXMLString">XML data string returned from query</param>
    ''' <returns>True-successfully loaded XML file, False-failed to load XML file</returns>
    ''' <remarks></remarks>
    Private Function LoadDataTableFromXmlResponseString(sXMLString As String) As Boolean

        Try

            _LastError = ""

            Dim rdr As New System.IO.StringReader(sXMLString)

            ' Load XML response data into a temporary work dataset  
            _dsXmlResponseData = New DataSet()
            _dsXmlResponseData.ReadXml(rdr)

            'Extract returned SQL column definitions from XML file
            _dtColumnDefinitions = _dsXmlResponseData.Tables("col")
            _iColumnCount = _dsXmlResponseData.Tables("col").Rows.Count
            _iRowCount = _dsXmlResponseData.Tables("data").Rows.Count / _iColumnCount

            'Start new DataTable for XML resultset returned by XMLSERVICE from IBM i
            _dtReturnData = New DataTable
            For Each dr1 As DataRow In _dsXmlResponseData.Tables("col").Rows
                _dtReturnData.Columns.Add(dr1(0), Type.GetType("System.String"))
            Next

            Dim colct As Integer
            Dim row2 As DataRow = Nothing
            For Each dr1 As DataRow In _dsXmlResponseData.Tables("data").Rows

                'If first column field, create new row
                If colct = 0 Then
                    row2 = _dtReturnData.NewRow
                ElseIf colct = _iColumnCount Then
                    'All 
                    _dtReturnData.Rows.Add(row2)
                    row2 = _dtReturnData.NewRow
                    colct = 0
                End If

                'Set column value for row
                row2(dr1(0).ToString) = dr1(1).ToString
                colct = colct + 1
            Next

            'Add last row to dataset
            _dtReturnData.Rows.Add(row2)

            _LastError = _iRowCount & " rows were returned from XML string."

            _bXMLIsLoaded = True
            Return True

        Catch ex As Exception
            _LastError = ex.Message
            _bXMLIsLoaded = False
            Return False
        End Try
    End Function
    ''' <summary>
    ''' This function loads an XML string that contains a data stream returned from the XMLSERVICE service program
    ''' </summary>
    ''' <param name="sXMLString">XML File</param>
    ''' <returns>True-successfully loaded XML file, False-failed to load XML file</returns>
    ''' <remarks></remarks>
    Private Function LoadPgmCallDataTableFromXmlResponseString(sXMLString As String) As Boolean

        Try

            _LastError = ""

            Dim rdr As New System.IO.StringReader(sXMLString)

            ' Load XML data into a dataset  
            _dsProgramResponse = New DataSet()
            _dsProgramResponse.ReadXml(rdr)

            'Extract returned SQL column definitions from XML file
            ''dtColumnDefinitions = mDS1.Tables("col")
            _iRowCount = _dsProgramResponse.Tables("data").Rows.Count

            'Start new DataTable for XML resultset returned by XMLSERVICE from IBM i
            _dtProgramResponse = New DataTable
            _dtProgramResponse.Columns.Add("parmtype", Type.GetType("System.String"))
            _dtProgramResponse.Columns.Add("parmvalue", Type.GetType("System.String"))

            Dim row2 As DataRow = Nothing
            Dim ict As Integer = 0
            For Each dr1 As DataRow In _dsProgramResponse.Tables("data").Rows
                row2 = _dtProgramResponse.NewRow
                row2(0) = "Parm" & ict 'Set unique ordinal zero based parm name 
                row2(1) = dr1(1)
                _dtProgramResponse.Rows.Add(row2)
                ict += 1
            Next

            _LastError = _iRowCount & " rows were returned from XML string."

            _bXMLIsLoaded = True
            Return True

        Catch ex As Exception
            _LastError = ex.Message
            _bXMLIsLoaded = False
            Return False
        End Try
    End Function
    ''' <summary>
    ''' This function loads an XML string that contains a data stream returned from the XMLSERVICE service program command calls
    ''' </summary>
    ''' <param name="sXMLString">XML File</param>
    ''' <returns>True-successfully loaded XML file, False-failed to load XML file</returns>
    ''' <remarks></remarks>
    Private Function LoadCmdCallDataTableFromXmlResponseString(sXMLString As String) As Boolean

        Dim iTableCount As Integer = 0

        Try

            _LastError = ""

            Dim rdr As New System.IO.StringReader(sXMLString)

            ' Load XML data into a dataset  
            _dsCommandResponse = New DataSet()
            _dsCommandResponse.ReadXml(rdr)

            'Get number of tables
            iTableCount = _dsCommandResponse.Tables.Count

            'If only 1 table, command ran successfully

            If iTableCount = 1 Then

                'Extract returned SQL column definitions from XML file
                ''dtColumnDefinitions = mDS1.Tables("col")
                _iRowCount = _dsCommandResponse.Tables("cmd").Rows.Count

                'Start new DataTable for XML resultset returned by XMLSERVICE from IBM i
                _dtCommandResponse = New DataTable
                _dtCommandResponse.Columns.Add("parmtype", Type.GetType("System.String"))
                _dtCommandResponse.Columns.Add("parmvalue", Type.GetType("System.String"))

                Dim row2 As DataRow = Nothing
                Dim ict As Integer = 0
                For Each dr1 As DataRow In _dsCommandResponse.Tables("cmd").Rows
                    row2 = _dtCommandResponse.NewRow
                    row2(0) = "Msg" & ict 'Set unique ordinal zero based parm name 
                    row2(1) = dr1(0)
                    _dtCommandResponse.Rows.Add(row2)
                    ict += 1
                Next

            ElseIf iTableCount > 1 Then 'More than 1 table.  Errors probably occurred on command call

                'Extract returned SQL column definitions from XML file
                ''dtColumnDefinitions = mDS1.Tables("col")
                _iRowCount = _dsCommandResponse.Tables("joblogrec").Rows.Count

                'Start new DataTable for XML resultset returned by XMLSERVICE from IBM i
                _dtCommandResponse = New DataTable
                _dtCommandResponse.Columns.Add("parmtype", Type.GetType("System.String"))
                _dtCommandResponse.Columns.Add("parmvalue", Type.GetType("System.String"))

                Dim row2 As DataRow = Nothing
                Dim ict As Integer = 0
                For Each dr1 As DataRow In _dsCommandResponse.Tables("joblogrec").Rows
                    row2 = _dtCommandResponse.NewRow
                    row2(0) = "Msg" & ict 'Set unique ordinal zero based parm name 
                    row2(1) = dr1(0).ToString & " - " & dr1(2).ToString & " - " & dr1(1).ToString
                    _dtCommandResponse.Rows.Add(row2)
                    ict += 1
                Next

            End If

            _LastError = _iRowCount & " rows were returned from XML string."

            _bXMLIsLoaded = True
            Return True

        Catch ex As Exception
            _LastError = ex.Message
            _bXMLIsLoaded = False
            Return False
        End Try
    End Function
#End Region

End Class
