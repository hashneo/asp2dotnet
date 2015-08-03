<%
'############ Common Functions, used unchanged by many classes
'Takes an SQL String
'Calls other functions

'Takes an SQL Query
'Runs the Query and returns a recordset
 Function LoadRSFromDB(p_strSQL)
    dim rs, cmd

    Set rs = Server.CreateObject("adodb.Recordset")
    Set cmd = Server.CreateObject("adodb.Command")
    

    'Run the SQL
    cmd.ActiveConnection  = dbConnectionString
    cmd.CommandText = p_strSQL
    cmd.CommandType = adCmdText
    cmd.Prepared = true

    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenForwardOnly, adLockReadOnly

    if Err <> 0 then
        Err.Raise  Err.Number, "ADOHelper: RunSQLReturnRS", Err.Description
    end if

    
    ' Disconnect the recordsets and cleanup  
    'Set rs.ActiveConnection = Nothing  
    'Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set LoadRSFromDB = rs
End Function


 Function RunSQL(ByVal p_strSQL)
        ' Create the ADO objects
        Dim cmd
        Set cmd = Server.CreateObject("adodb.Command")

        cmd.ActiveConnection  = dbConnectionString
        cmd.ActiveConnection.BeginTrans
        cmd.CommandText = p_strSQL
        cmd.CommandType = adCmdText

        ' Execute the query without returning a recordset
        ' Specifying adExecuteNoRecords reduces overhead and improves performance
        cmd.Execute true, , adExecuteNoRecords
        cmd.ActiveConnection.CommitTrans

        if Err <> 0 then
            cmd.ActiveConnection.RollBackTrans
            Err.Raise  Err.Number, "ADOHelper: RunSQL", Err.Description
        end if

        ' Cleanup
        Set cmd.ActiveConnection = Nothing
        Set cmd = Nothing
End Function

 Function InsertRecord(tblName, strAutoFieldName, ArrFlds, ArrValues )
dim conn, rs, thisID   
    Set conn = Server.CreateObject ("ADODB.Connection")
    Set rs = Server.CreateObject ("ADODB.Recordset")

    conn.open dbConnectionString
    conn.BeginTrans
    rs.Open tblName, conn, adOpenKeyset, adLockOptimistic, adCmdTable

    rs.AddNew  ArrFlds, ArrValues
    rs.Update 

    thisID = rs(strAutoFieldName)

    rs.Close
    Set rs = Nothing

    conn.CommitTrans        
    conn.close
    Set conn = Nothing

    If Err.number = 0 Then
        InsertRecord = thisID
    End If        
End Function 


function SingleQuotes(pStringIn)
    if pStringIn = "" or isnull(pStringIn) then exit function
    Dim pStringModified
    pStringModified = Replace(pStringIn,"'","''")
    SingleQuotes =  pStringModified
end function

public function echo(p_STR)
    response.write p_Str
end function

public function die(p_STR)
    echo p_Str
    response.end
end function

public function echobr(p_STR)
    echo p_Str & "<br>" & vbCRLF
end function

public function htmlencode(p_STR)
    htmlencode = trim(server.htmlencode(p_Str & " "))
end function

Randomize 'Insure that the numbers are really random
    Function RandomString(p_NumChars)
    Dim n
    Dim tmpChar,tmpString
    for n = 0 to p_NumChars
        tmpChar = Chr(Int(32+( Rnd * (126-33))))
        'Random characters (letters, numbers, etc.)
        tmpString = tmpString & tmpChar
    next
    RandomString = tmpString
End Function

Const dbConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\vboo_pubs.mdb;"
'Const dbConnectionString = "vboo_pubs"

' -- ADO command types
Const   adCmdUnspecified    =   -1  '   Gibt nicht den Typ des Befehlsarguments an.
Const   adCmdText               =   1       '   Wertet CommandText als Textdefinition eines Befehls oder als Aufruf einer gespeicherten Prozedur aus.
Const   adCmdTable              =   2       '   Wertet CommandText als Tabellennamen aus, dessen Spalten vollzählig von einer intern generierten SQL-Abfrage zurückgegeben werden.
Const   adCmdStoredProc =   4       '   Wertet CommandText als Namen einer gespeicherten Prozedur aus.
Const   adCmdUnknown        =   8       '   Standard. Gibt an, dass der in der CommandText-Eigenschaft verwendete Befehlstyp unbekannt ist.
Const   adCmdFile                   =   256 '   Wertet CommandText als Dateinamen eines dauerhaft gespeicherten Recordset-Objekts aus.
Const   adCmdTableDirect    =   512 '   Wertet CommandText als Tabellennamen aus, dessen Spalten vollzählig zurückgegeben werden.


' -- ADO cursor types
Const   adOpenForwardOnly   =   0       '   Standard. Verwendet einen Cursor vom Typ Vorwärts. Dieser Cursor ist identisch mit einem statischen Cursor, mit dem Unterschied, dass ein Durchsuchen der Datensätze nur in Vorwärtsrichtung möglich ist. Dadurch wird die Leistung verbessert, wenn nur ein einziger Durchlauf durch ein Recordset-Objekt durchgeführt werden muss.
Const   adOpenDynamic           =   2       '   Verwendet einen dynamischen Cursor. Von anderen Benutzern vorgenommene Zusätze, Änderungen und Löschvorgänge können angezeigt werden. Alle Bewegungsarten durch das Recordset–Objekt sind zulässig, mit Ausnahme von Lesezeichen (sofern der Provider diese nicht unterstützt).
Const   adOpenKeyset                =   1       '   Verwendet einen Cursor von Typ Keyset. Ähnelt einem dynamischen Cursor, mit dem Unterschied, dass von anderen Benutzern hinzugefügte Datensätze nicht angezeigt werden können, obwohl ein Zugriff auf von anderen Benutzern gelöschte Datensätze von Ih' Recordset-Objekt aus nicht möglich ist. Von anderen Benutzern vorgenommene Datenänderungen können weiterhin angezeigt werden.
Const   adOpenStatic                =   3       '   Verwendet einen statischen Cursor. Eine statische Kopie einer Gruppe von Datensätzen, anhand derer Daten gesucht und Berichte entwickelt werden können. Von anderen Benutzern vorgenommene Zusätze, Änderungen oder Löschvorgänge können nicht angezeigt werden.
Const   adOpenUnspecified       =   -1  '   Gibt keinen Cursortyp an


' -- ADO cursor locations
const adUseServer        = 2 '# (Default)
const adUseClient        = 3

' -- ADO lock types
const adLockReadOnly        = 1
const adLockPessimistic     = 2
const adLockOptimistic      = 3
const adLockBatchOptimistic = 4

' -- DataTypeEnum Values
'Const      AdArray                     =       0x2000  '   Ein Flagwert, immer in Kombination mit einer anderen Datentypkonstanten, der ein Array dieses anderen Datentyps kennzeichnet.
Const       adBigInt                        =       20      '       Zeigt eine 8 Byte umfassende ganze Zahl mit Vorzeichen (DBTYPE_I8) an.
Const       adBinary                        =       128     '       Zeigt einen Binärwert (DBTYPE_BYTES) an.
Const       adBoolean                   =       11      '       Zeigt einen booleschen Wert (DBTYPE_BOOL) an.
Const       adBSTR                      =       8           '       Zeigt eine mit Nullzeichen endende Zeichenfolge (Unicode) (DBTYPE_BSTR) an.
Const       adChapter                   =       136     '       Zeigt einen 4 Byte langen Chapter-Wert zum Bezeichnen von Zeilen in einem untergeordneten Rowset-Objekt (DBTYPE_HCHAPTER) an.
Const       adChar                      =       129     '       Zeigt einen Zeichenfolgenwert (DBTYPE_STR) an.
Const       adCurrency              =       6           '       Zeigt einen Currency-Wert (DBTYPE_CY) an. Eine Währungsangabe (Currency) ist eine Festkommazahl mit vier Stellen hinter dem Komma. Sie wird als 8 Byte umfassende ganze Zahl mit Vorzeichen, skaliert mit 10.000, gespeichert.
Const       adDate                      =       7           '       Zeigt einen Datumswert (DBTYPE_DATE) an. Ein Datum wird als Wert vom Datentyp Double gespeichert, dessen Integeranteil die Anzahl der Tage seit dem 30. Dezember 1899 darstellt. Bei dem Bruchteil des Double handelt es sich um den Bruchteil des Tages.
Const       adDBDate                    =       133     '       Zeigt einen Datumswert (jjjjmmtt) (DBTYPE_DBDATE) an.
Const       adDBTime                    =       134     '       Zeigt einen Zeitwert (hhmmss) (DBTYPE_DBTIME) an.
Const       adDBTimeStamp       =       135     '       Zeigt einen Timestamp mit Datum und Zeit (jjjjmmtthhmmss plus ein Bruch in Milliardstel) (DBTYPE_DBTIMESTAMP) an.
Const       adDecimal                   =       14      '       Zeigt einen genauen numerischen Wert mit fester Genauigkeit und Skalierung (DBTYPE_DECIMAL) an.
Const       adDouble                    =       5           '       Zeigt einen Gleitkommawert doppelter Genauigkeit (DBTYPE_R8) an.
Const       adEmpty                 =       0           '       Gibt an, dass kein Wert vorhanden ist (DBTYPE_EMPTY).
Const       adError                     =       10      '       Zeigt einen 32-Bit-Fehlercode (DBTYPE_ERROR) an.
Const       adFileTime                  =       64      '       Zeigt einen 64 Bit umfassenden Wert, der die Anzahl der seit dem 1. Januar 1601 verstrichenen 100-Nanosekunden-Intervalle darstellt (DBTYPE_FILETIME) an.
Const       adGUID                      =       72      '       Zeigt einen global eindeutigen Bezeichner (Globally Unique Identifier, GUID) (DBTYPE_GUID) an.
Const       adIDispatch             =       9           '       Zeigt einen Zeiger auf eine IDispatch-Schnittstelle in einem COM-Objekt (DBTYPE_IDISPATCH) an.
Const       adInteger                   =       3           '       Zeigt eine 4 Byte umfassende ganze Zahl mit Vorzeichen (DBTYPE_I4) an.
Const       adIUnknown              =       13      '       Zeigt einen Zeiger auf eine IUnknown-Schnittstelle in einem COM-Objekt (DBTYPE_IUNKNOWN) an.
Const       adLongVarBinary     =       205     '       Zeigt einen Binärwert vom Datentyp Long (nur Parameter-Objekt) an.
Const       adLongVarChar       =       201     '       Zeigt einen Zeichenfolgenwert vom Datentyp Long (nur Parameter-Objekt) an.
Const       adLongVarWChar      =       203     '       Zeigt einen mit Nullzeichen endenden Unicode-Zeichenfolgenwert vom Datentyp Long (nur Parameter-Objekt) an.
Const       adNumeric               =       131     '       Zeigt einen genauen numerischen Wert mit fester Genauigkeit und Skalierung (DBTYPE_NUMERIC) an.
Const       adPropVariant           =       138     '       Zeigt einen PROPVARIANT-Wert für die Automatisierung (DBTYPE_PROP_VARIANT) an.
Const       adSingle                        =       4           '       Zeigt einen Gleitkommawert einfacher Genauigkeit (DBTYPE_R4) an.
Const       adSmallInt                  =       2           '       Zeigt eine 2 Byte umfassende Ganzzahl mit Vorzeichen (DBTYPE_I2) an.
Const       adTinyInt                   =       16      '       Zeigt eine 1 Byte umfassende Ganzzahl mit Vorzeichen (DBTYPE_I1) an.
Const       adUnsignedBigInt        =       21      '       Zeigt eine 8 Byte umfassende Ganzzahl ohne Vorzeichen (DBTYPE_UI8) an.
Const       adUnsignedInt           =       19      '       Zeigt eine 4 Byte umfassende Ganzzahl ohne Vorzeichen (DBTYPE_UI4) an.
Const       adUnsignedSmallInt=     18      '       Zeigt eine 2 Byte umfassende Ganzzahl ohne Vorzeichen (DBTYPE_UI2) an.
Const       adUnsignedTinyInt   =       17      '       Zeigt eine 1 Byte umfassende Ganzzahl ohne Vorzeichen (DBTYPE_UI1) an.
Const       adUserDefined           =       132     '       Zeigt eine benutzerdefinierte Variable (DBTYPE_UDT) an.
Const       adVarBinary             =       204     '       Zeigt einen Binärwert (nur Parameter-Objekt) an.
Const       adVarChar                   =       200     '       Zeigt einen Zeichenfolgenwert (nur Parameter-Objekt) an.
Const       adVariant                   =       12      '       Zeigt einen Wert vom Datentyp Variant für die Automatisierung (DBTYPE_VARIANT) an.
Const       adVarNumeric            =       139     '       Zeigt einen numerischen Wert (nur Parameter-Objekt) an.
Const       adVarWChar              =       202     '       Zeigt eine mit Nullzeichen endende Unicode-Zeichenfolge (nur Parameter-Objekt) an.
Const       adWChar                 =       130     '       Zeigt eine mit Nullzeichen endende Unicode-Zeichenfolge (DBTYPE_WSTR) an.

' -- ParameterDirectionEnum Values
const adParamUnknown = &H0000
const adParamInput = &H0001
const adParamOutput = &H0002
const adParamInputOutput = &H0003
const adParamReturnValue = &H0004

' -- ExecuteOptionEnum Values
Const adAsyncExecute = 16
Const adAsyncFetch = 32
Const adAsyncFetchNonBlocking = 64
Const adExecuteNoRecords = 128
Const adBookmarkCurrent = 0
%>