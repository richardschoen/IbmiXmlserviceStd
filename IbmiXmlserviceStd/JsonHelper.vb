Imports System.Collections.Generic
Imports System.Linq
Imports System.Web
Imports System.Runtime.Serialization.Json
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Data
Imports Newtonsoft.Json
'Imports System.Web.script.Serialization

''' <summary>
''' JSON Serialization and Deserialization Assistant Class
''' Source from: https://gist.github.com/monk8800/3760559
''' Note: Normally we would use the Newtonsoft JSON API, but this class removes any extra dependencies.
''' </summary>
Public Class JsonHelper

    Private _lasterror As String = ""

    ''' <summary>
    ''' Get last error
    ''' </summary>
    ''' <returns></returns>
    Public Function GetLastError() As String
        Return _lasterror
    End Function

    ''' <summary>
    ''' JSON Serialization
    ''' </summary>
    Public Function JsonSerializer(Of T)(ByVal obj As T) As String
        Dim ser As New DataContractJsonSerializer(GetType(T))
        Dim ms As New MemoryStream()
        ser.WriteObject(ms, obj)
        Dim jsonString As String = Encoding.UTF8.GetString(ms.ToArray())
        ms.Close()
        'Replace Json Date String                                         
        Dim p As String = "\\/Date\((\d+)\+\d+\)\\/"
        Dim matchEvaluator As New MatchEvaluator(AddressOf ConvertJsonDateToDateString)
        Dim reg As New Regex(p)
        jsonString = reg.Replace(jsonString, matchEvaluator)
        Return jsonString
    End Function

    ''' <summary>
    ''' JSON Deserialization
    ''' </summary>
    Public Function JsonDeserialize(Of T)(ByVal jsonString As String) As T
        'Convert "yyyy-MM-dd HH:mm:ss" String as "\/Date(1319266795390+0800)\/"
        Dim p As String = "\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2}"
        Dim matchEvaluator As New MatchEvaluator(AddressOf ConvertDateStringToJsonDate)
        Dim reg As New Regex(p)
        jsonString = reg.Replace(jsonString, matchEvaluator)
        Dim ser As New DataContractJsonSerializer(GetType(T))
        Dim ms As New MemoryStream(Encoding.UTF8.GetBytes(jsonString))
        Dim obj As T = DirectCast(ser.ReadObject(ms), T)
        Return obj
    End Function

    ''' <summary>
    ''' Convert Serialization Time /Date(1319266795390+0800) as String
    ''' </summary>
    Private Function ConvertJsonDateToDateString(ByVal m As Match) As String
        Dim result As String = String.Empty
        Dim dt As New DateTime(1970, 1, 1)
        dt = dt.AddMilliseconds(Long.Parse(m.Groups(1).Value))
        dt = dt.ToLocalTime()
        result = dt.ToString("yyyy-MM-dd HH:mm:ss")
        Return result
    End Function

    ''' <summary>
    ''' Convert Date String as Json Time
    ''' </summary>
    Private Function ConvertDateStringToJsonDate(ByVal m As Match) As String
        Dim result As String = String.Empty
        Dim dt As DateTime = DateTime.Parse(m.Groups(0).Value)
        dt = dt.ToUniversalTime()
        Dim ts As TimeSpan = dt - DateTime.Parse("1970-01-01")
        result = String.Format("\/Date({0}+0800)\/", ts.TotalMilliseconds)
        Return result
    End Function
    ''' <summary>
    ''' Convert DataTable object to Json String
    ''' Source: 'https://stackoverflow.com/questions/21648064/vb-net-datatable-serialize-to-json
    ''' </summary>
    ''' <param name="dtWork">DataTable object</param>
    ''' <returns>Serialized JSON DataTable as string</returns>
    Public Function ConvertDataTableToJson(ByVal dtWork As DataTable) As String
        Try
            _lasterror = ""

            'TODO - This code depracated until MS ports: System.Web.script.Serialization to .Net Core

            ''Check DataTable to make sure it has data
            'If dtWork Is Nothing Then
            '    Throw New Exception("Data table is Nothing. No data available to serialize.")
            'End If

            ''Serialize to JSON and return 
            'Return New JavaScriptSerializer().Serialize(From dr As DataRow In dtWork.Rows Select dtWork.Columns.Cast(Of DataColumn)().ToDictionary(Function(col) col.ColumnName, Function(col) dr(col)))

            Return ""

        Catch ex As Exception
            _lasterror = ex.Message
            Return ""
        End Try
    End Function
    ''' <summary>
    ''' Serialize DataTable with Newtonsoft.JSON
    ''' Sample from:
    ''' http://www.c-sharpcorner.com/UploadFile/9bff34/3-ways-to-convert-datatable-to-json-string-in-Asp-Net-C-Sharp/
    ''' </summary>
    ''' <param name="table">DataTable</param>
    ''' <returns>JSON string</returns>
    Public Function DataTableToJSONWithJSONNet(table As DataTable) As String
        Dim JSONString As String = String.Empty
        Try
            'TODO - Removed Newtonsoft dependency. If Newtensoft added, you can enable this function again
            Return ""

            'JSONString = JsonConvert.SerializeObject(table)
            'Return JSONString
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Convert DataTable to json string using StringBuilder
    ''' </summary>
    ''' <param name="table">DataTable input</param>
    ''' <param name="debugINfo">True-Write debug info in response JSON. No debug info in error response</param>
    ''' <returns>DataTable results as JSON string</returns>
    Public Function DataTableToJsonWithStringBuilder(table As DataTable, Optional debugInfo As Boolean = False) As String

        Dim jsonString As StringBuilder = New StringBuilder()

        Try
            'Convert table rows to JSON
            If table.Rows.Count > 0 Then
                jsonString.Append("[")
                For i As Integer = 0 To table.Rows.Count - 1
                    jsonString.Append("{")
                    For j As Integer = 0 To table.Columns.Count - 1
                        If j < table.Columns.Count - 1 Then
                            jsonString.Append("""" + table.Columns(j).ColumnName.ToString() + """:" + """" + table.Rows(i)(j).ToString() + """,")
                        ElseIf j = table.Columns.Count - 1 Then
                            jsonString.Append("""" + table.Columns(j).ColumnName.ToString() + """:" + """" + table.Rows(i)(j).ToString() + """")
                        End If
                    Next
                    If i = table.Rows.Count - 1 Then
                        jsonString.Append("}")
                    Else
                        jsonString.Append("},")
                    End If
                Next
                jsonString.Append("]")

                'Return the JSON result
                Return jsonString.ToString()
            Else ' No data
                Return "[{""noresults"":""No results found""}]"
            End If

        Catch ex As Exception
            If debugInfo Then
                Return "[{""noresults"":""" & ex.Message & """}]"
            Else
                Return "[{""noresults"":""Exception occurred""}]"
            End If
        End Try

    End Function

End Class
