<%@Language=VBScript%>
<%
Option Explicit
Response.ContentType = "application/json"

Dim healthStatus, dbStatus, connectionString, objWSH
Dim startTime, endTime, responseTime
Dim errorMessage, errorNumber

' Initialize variables
healthStatus = "healthy"
dbStatus = "connected"
errorMessage = ""
errorNumber = 0
startTime = Timer

' Get connection string from environment variable
Set objWSH = CreateObject("WScript.Shell")
connectionString = objWSH.Environment("PROCESS")("ConnectionString")

' Check if connection string exists
If IsEmpty(connectionString) Or connectionString = "" Then
    healthStatus = "unhealthy"
    dbStatus = "configuration_error"
    errorMessage = "Connection string not found in environment variables"
    errorNumber = -1
Else
    ' Test database connection
    Call TestDatabaseConnection()
End If

endTime = Timer
responseTime = Round((endTime - startTime) * 1000, 2) ' Convert to milliseconds

' Clean up
Set objWSH = Nothing

' Output JSON response
Response.Write "{"
Response.Write """status"": """ & healthStatus & ""","
Response.Write """timestamp"": """ & Now() & ""","
Response.Write """responseTime"": " & responseTime & ","
Response.Write """database"": {"
Response.Write """status"": """ & dbStatus & ""","
Response.Write """connectionString"": """ & Left(connectionString, 20) & "..."""
If errorNumber <> 0 Then
    Response.Write ",""error"": {"
    Response.Write """number"": " & errorNumber & ","
    Response.Write """message"": """ & Replace(errorMessage, """", "\""") & """"
    Response.Write "}"
End If
Response.Write "},"
Response.Write """server"": {"
Response.Write """name"": """ & Request.ServerVariables("SERVER_NAME") & ""","
Response.Write """software"": """ & Request.ServerVariables("SERVER_SOFTWARE") & ""","
Response.Write """time"": """ & Request.ServerVariables("DATE_GMT") & """"
Response.Write "}"
Response.Write "}"

Sub TestDatabaseConnection()
    On Error Resume Next
    
    Dim objConn, testQuery, rs
    
    ' Create connection object
    Set objConn = CreateObject("ADODB.Connection")
    objConn.ConnectionString = connectionString
    objConn.ConnectionTimeout = 10 ' 10 second timeout
    objConn.CommandTimeout = 15 ' 15 second command timeout
    
    ' Attempt to open connection
    objConn.Open
    
    If Err.Number <> 0 Then
        healthStatus = "unhealthy"
        dbStatus = "connection_failed"
        errorMessage = Err.Description
        errorNumber = Err.Number
        Err.Clear
        Exit Sub
    End If
    
    ' Test a simple query to ensure database is responsive
    testQuery = "SELECT 1 as HealthCheck"
    Set rs = objConn.Execute(testQuery)
    
    If Err.Number <> 0 Then
        healthStatus = "unhealthy"
        dbStatus = "query_failed"
        errorMessage = Err.Description
        errorNumber = Err.Number
        Err.Clear
    Else
        ' Verify we got expected result
        If Not rs.EOF Then
            If rs.Fields("HealthCheck").Value = 1 Then
                dbStatus = "connected"
            Else
                healthStatus = "unhealthy"
                dbStatus = "unexpected_result"
                errorMessage = "Health check query returned unexpected result"
                errorNumber = -2
            End If
        Else
            healthStatus = "unhealthy"
            dbStatus = "no_result"
            errorMessage = "Health check query returned no results"
            errorNumber = -3
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    ' Clean up connection
    objConn.Close
    Set objConn = Nothing
    
    On Error GoTo 0
End Sub
%>