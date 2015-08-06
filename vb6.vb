Public Const adEmpty As Integer = 0     ' No value
Public Const adSmallInt As Integer = 2     ' A 2-byte signed integer.
Public Const adInteger As Integer = 3     ' A 4-byte signed integer.
Public Const adSingle As Integer = 4     ' A single-precision floating-point value.
Public Const adDouble As Integer = 5     ' A double-precision floating-point value.
Public Const adCurrency As Integer = 6     ' A currency value
Public Const adDate As Integer = 7     ' The number of days since December 30, 1899 + the fraction of a day.
Public Const adBSTR As Integer = 8     ' A null-terminated character string.
Public Const adIDispatch As Integer = 9     ' A pointer to an IDispatch interface on a COM object. Note: Currently not supported by ADO.
Public Const adError As Integer = 10     ' A 32-bit error code
Public Const adBoolean As Integer = 11     ' A boolean value.
Public Const adVariant As Integer = 12     ' An Automation Variant. Note: Currently not supported by ADO.
Public Const adIUnknown As Integer = 13     ' A pointer to an IUnknown interface on a COM object. Note: Currently not supported by ADO.
Public Const adDecimal As Integer = 14     ' An exact numeric value with a fixed precision and scale.
Public Const adTinyInt As Integer = 16     ' A 1-byte signed integer.
Public Const adUnsignedTinyInt As Integer = 17     ' A 1-byte unsigned integer.
Public Const adUnsignedSmallInt As Integer = 18     ' A 2-byte unsigned integer.
Public Const adUnsignedInt As Integer = 19     ' A 4-byte unsigned integer.
Public Const adBigInt As Integer = 20     ' An 8-byte signed integer.
Public Const adUnsignedBigInt As Integer = 21     ' An 8-byte unsigned integer.
Public Const adFileTime As Integer = 64     ' The number of 100-nanosecond intervals since January 1,1601
Public Const adGUID As Integer = 72     ' A globally unique identifier (GUID)
Public Const adBinary As Integer = 128     ' A binary value.
Public Const adChar As Integer = 129     ' A string value.
Public Const adWChar As Integer = 130     ' A null-terminated Unicode character string.
Public Const adNumeric As Integer = 131     ' An exact numeric value with a fixed precision and scale.
Public Const adUserDefined As Integer = 132     ' A user-defined variable.
Public Const adDBDate As Integer = 133     ' A date value (yyyymmdd).
Public Const adDBTime As Integer = 134     ' A time value (hhmmss).
Public Const adDBTimeStamp As Integer = 135     ' A date/time stamp (yyyymmddhhmmss plus a fraction in billionths).
Public Const adChapter As Integer = 136     ' A 4-byte chapter value that identifies rows in a child rowset
Public Const adPropVariant As Integer = 138     ' An Automation PROPVARIANT.
Public Const adVarNumeric As Integer = 139     ' A numeric value (Parameter object only).
Public Const adVarChar As Integer = 200     ' A string value (Parameter object only).
Public Const adLongVarChar As Integer = 201     ' A long string value.
Public Const adVarWChar As Integer = 202     ' A null-terminated Unicode character string.
Public Const adLongVarWChar As Integer = 203     ' A long null-terminated Unicode string value.
Public Const adVarBinary As Integer = 204     ' A binary value (Parameter object only).
Public Const adLongVarBinary As Integer = 205     ' A long binary value.
Public Const AdArray As Integer = 8192     ' A flag value combined with another data type constant. Indicates an array of that other data type.

Public Const adParamUnknown As Integer = 0     ' Direction unknown
Public Const adParamInput As Integer = 1     ' Default. Input parameter
Public Const adParamInputOutput As Integer = 3     ' Input and output parameter
Public Const adParamOutput As Integer = 2     ' Output parameter
Public Const adParamReturnValue As Integer = 4     ' Return value

Public Const adCmdUnspecified As Integer = -1     ' Does not specify the command type argument.
Public Const adCmdText As Integer = 1     ' Evaluates CommandText as a textual definition of a command or stored procedure call.
Public Const adCmdTable As Integer = 2     ' Evaluates CommandText as a table name whose columns are all returned by an internally generated SQL query.
Public Const adCmdStoredProc As Integer = 4     ' Evaluates CommandText as a stored procedure name.
Public Const adCmdUnknown As Integer = 8     ' Default. Indicates that the type of command in the CommandText property is not known.
Public Const adCmdFile As Integer = 256     ' Evaluates CommandText as the file name of a persistently stored Recordset. Used with Recordset.Open or Requery only.
Public Const adCmdTableDirect As Integer = 512     ' Evaluates CommandText as a table name whose columns are all returned. Used with Recordset.Open or Requery only. To use the Seek method, the Recordset must be opened with adCmdTableDirect. This value cannot be combined with the ExecuteOptionEnum value adAsyncExecute.

Public Const adAsyncExecute = 16
Public Const adAsyncFetch = 32
Public Const adAsyncFetchNonBlocking = 64
Public Const adExecuteNoRecords = 128
Public Const adExecuteStream = 1024

Public Const adPersistADTG = 0
Public Const adPersistXML = 1
Public Const adPersistADO = 1
Public Const adPersistProviderSpecific = 2

Public Const adModeUnknown As Integer = 0     ' Permissions have not been set or cannot be determined.
Public Const adModeRead As Integer = 1     ' Read-only.
Public Const adModeWrite As Integer = 2     ' Write-only.
Public Const adModeReadWrite As Integer = 3     ' Read/write.
Public Const adModeShareDenyRead As Integer = 4     ' Prevents others from opening a connection with read permissions.
Public Const adModeShareDenyWrite As Integer = 8     ' Prevents others from opening a connection with write permissions.
Public Const adModeShareExclusive As Integer = 12     ' Prevents others from opening a connection.
Public Const adModeShareDenyNone As Integer = 16     ' Allows others to open a connection with any permissions.
Public Const adModeRecursive As Integer = 4194304     ' Used with adModeShareDenyNone, adModeShareDenyWrite, or adModeShareDenyRead to set permissions on all sub-records of the current Record.

<Obsolete("VB6 IsNull is obsoleted")>
Shared Function IsNull(v)
    If (TypeOf v Is String) Then
        Return String.IsNullOrEmpty(v)
    End If
    Return (v Is Nothing)
End Function

<Obsolete("VB6 IsObject is obsoleted")>
Shared Function IsObject(v)
    Return ( Not( v Is Nothing ) )
End Function

<Obsolete("VB6 IsEmpty is obsoleted")>
Shared Function IsEmpty(v)
    Return String.IsNullOrEmpty(v)
End Function

<Obsolete("VB6 Consider using Native .NET Classes for a performance increase")>
Shared Function CreateObject(v)
    Return System.Activator.CreateInstance(System.Type.GetTypeFromProgID( v ))
End Function

<Obsolete("VB6 CCur is obsoleted")>
Shared Function CCur(v)
    Return v.ToString("c2")
End Function

' Wrapper for RegExp
Public Class RegExp
	<Obsolete("VB6 RegExp is obsoleted, use the .NET System.Text.RegularExpressions.Regex instead.")>
	Sub New()
	End Sub
End Class
