'Option Explicit
'-------------------------------------------------
'SQL connection declarations
    Const SQLhost = "***"
    Const SQLuser = "***"
    Const SQLPW = "***"
    Const SQLDB = "***"
'--------------------------------------------------
'-------------TEEEMP-------------------------------
''SQL connection declarations
'Const SQLhost = "***"
'Const SQLuser = "***"
'Const SQLPW = "***"
'Const SQLDB = "***"
'--------------------------------------------------



Public con As Object
'since this version maintains existing connection, it would be good to kill it when finishing your macros. For example on Userform close trigger:
'                Private Sub UserForm_Terminate()
'                On Error Resume Next
'                con.Close
'                Set con = Nothing
'                End Sub

Function GetColumns2String(InputRNG As Range)
Dim r As Range
Dim i As Long
Dim MySTR, SIngleSTR As String
For Each r In InputRNG.Cells
    SIngleSTR = CStr(r.Value)
    SIngleSTR = "[" & SIngleSTR & "]"
    MySTR = MySTR & SIngleSTR & ", "
Next r

MySTR = Left(MySTR, Len(MySTR) - 2)
GetColumns2String = MySTR


End Function


Function GetData2String(InputRNG As Range) As String
Dim r As Range
Dim i As Long
Dim MySTR, SIngleSTR As String
For Each r In InputRNG.Cells
    SIngleSTR = r.Value
    If r.Text = "" Then
    SIngleSTR = "NULL"
    ElseIf IsDate(r) Then
    SIngleSTR = "'" & Format(SIngleSTR, "yyyymmdd") & "'"
    Else
    SIngleSTR = Replace(SIngleSTR, "'", "''", , , vbTextCompare)  'important part for apostrophes
    SIngleSTR = Replace(SIngleSTR, "CHR(34)", "", , , vbTextCompare) 'removing the quotation marks in a cell
    SInssssssssssssssssssssssssssssssleSTR = "'" & SIngleSTR & "'"
    End If
    MySTR = MySTR & SIngleSTR & ", "
Next r

MySTR = Left(MySTR, Len(MySTR) - 2)
GetData2String = MySTR


End Function

'Get the data from MS SQL
Sub MS_SQLData2(SQL, Optional OutputRNG As Range, Optional TheHeader As Boolean)




            If Not con Is Nothing Then
           ' Debug.Print "existing con triggered, yes yes yes " & SQL
            GoTo skipcon:
            End If
resumecon:
    Dim rs          As Object
    Dim AccessFile  As String
    Dim i           As Integer
    Dim Constr As String
    'Disable screen flickering.
    Application.ScreenUpdating = False
    
    'Set the name of the table you want to retrieve the data.
    On Error Resume Next
    'Create the ADODB connection object.
    Set con = CreateObject("ADODB.connection")
    'Check if the object was created.
    If Err.Number <> 0 Then
        MsgBox "Connection was not created!", vbCritical, "Connection Error"
        Exit Sub
    End If
    On Error GoTo 0
'connection string forming------------------------------------------------
Constr = "DRIVER={SQL Server};SERVER="
    Constr = Constr & SQLhost & ";trusted_connection=no;DataBase=" & SQLDB
    Constr = Constr & ";UID=" & SQLuser & ";PWD=" & SQLPW
'-------------------------------------------------------------------------

'Open the connection.
    con.Open Constr
    On Error Resume Next
skipcon:
    On Error GoTo resumecon:
    'Create the ADODB recordset object.
    Set rs = CreateObject("ADODB.Recordset")
    'Check if the object was created.
    If Err.Number <> 0 Then
        'Error! Release the objects and exit.
        Set rs = Nothing
        Set con = Nothing
        'Display an error message to the user.
        'MsgBox "Recordset was not created!", vbCritical, "Recordset Error"
        Exit Sub
    End If
    On Error GoTo 0
         
    'Set thee cursor location.
    rs.CursorLocation = 3 'adUseClient on early  binding
    rs.CursorType = 1 'adOpenKeyset on early  binding
    
    'Open the recordset.
     con.CommandTimeout = 0
    rs.Open SQL, con
    
    'now for non select (mytrick =)
    If InStr(SQL, "SELECT") = 0 Then GoTo DondDoAnythingElse:
        'getting the geader
        If TheHeader = True Then
             'Copy the recordset headers.d
            For i = 0 To rs.Fields.Count - 1
              OutputRNG.Offset(-1, i) = rs.Fields(i).Name
            Next i
    End If
    
                                            'Check if the recordet is empty.
                                            On Error Resume Next
                                            
                                            If rs.EOF And rs.BOF Then

                                                'Close the recordet and the connection.
                                                rs.Close
DondDoAnythingElse:
                                                '###con.Close
                                                'Release the objects.
                                                Set rs = Nothing
                                                '###Set con = Nothing
                                                'Enable the screen.
                                                Application.ScreenUpdating = True
                                                'In case of an empty recordset display an error.
                                                                'MsgBox "There are no records in the recordset!", vbCritical, "No Records"
                                                Exit Sub
                                            End If

'
    'Write the query values in the sheet.
OutputRNG.CopyFromRecordset rs
    
    'Close the recordet and the connection.
    rs.Close
    '###con.Close
    
    'Release the objects.
    Set rs = Nothing
    '###Set con = Nothing
    
    
    'Enable the screen.
    Application.ScreenUpdating = True
End Sub
Function MS_SQLFunction(SQL) As Variant
'Debug.Print Now & "   runnin querry" & SQL
    'Declaring the necessary variables.
        'Dim con         As Object
            'con upgrade here
    Dim Constr As String
    
            If Not con Is Nothing Then
            'Debug.Print "existing Fun con con triggered, yes yes yes " & SQL
            GoTo skipcon:
            End If
resumecon:
    Dim rs          As Object
    Dim AccessFile  As String
    Dim i           As Integer
            
    'Disable screen flickering.
    Application.ScreenUpdating = False
    On Error Resume Next
    'Create the ADODB connection object.
    Set con = CreateObject("ADODB.connection")
    'Check if the object was created.
    If Err.Number <> 0 Then
        MsgBox "Connection was not created!", vbCritical, "Connection Error"
        Exit Function
    End If
    On Error GoTo 0
    'connection string forming------------------------------------------------
    Constr = "DRIVER={SQL Server};SERVER="
    Constr = Constr & SQLhost & ";trusted_connection=no;DataBase=" & SQLDB
    Constr = Constr & ";UID=" & SQLuser & ";PWD=" & SQLPW
'-------------------------------------------------------------------------

'Open the connection.
    con.Open Constr
    On Error Resume Next
skipcon:
        On Error GoTo resumecon:
    'Create the ADODB recordset object.
    Set rs = CreateObject("ADODB.Recordset")
    'Check if the object was created.
    If Err.Number <> 0 Then
        'Error! Release the objects and exit.
        Set rs = Nothing
        Set con = Nothing
        'Display an error message to the user.
        'MsgBox "Recordset was not created!", vbCritical, "Recordset Error"
        Exit Function
    End If
    On Error GoTo 0
         
    'Set thee cursor location.
    rs.CursorLocation = 3 'adUseClient on early  binding
    rs.CursorType = 1 'adOpenKeyset on early  binding
    
    'Open the recordset.
    con.CommandTimeout = 0
    rs.Open SQL, con
    MS_SQLFunction = rs.GetRows

DondDoAnythingElse:
                                                '###con.Close
                                                'Release the objects.
                                                Set rs = Nothing
                                               '### Set con = Nothing
                                                'Enable the screen.
                                                Application.ScreenUpdating = True
                                                'In case of an empty recordset display an error.
                                                                'MsgBox "There are no records in the recordset!", vbCritical, "No Records"
                                                Exit Function
                                            

'
    'Write the query values in the sheet.
OutputRNG.CopyFromRecordset rs
    
    'Close the recordet and the connection.
    rs.Close
    '###con.Close
    
    'Release the objects.
    Set rs = Nothing
    '###Set con = Nothing
    
    
    'Enable the screen.
    Application.ScreenUpdating = True
    
End Function


Sub Testit33()
'###MAS IMPART
Dim Theheaders, TheData, TheS
Dim InputRange, TempR, r As Range
Dim maxrow, maxcol, i, z As Long
Set InputRange = Selection
maxrow = InputRange.Rows.Count
maxcol = InputRange.Columns.Count
Set TempR = InputRange.Rows(1)
'TempR.Select
'MsgBox TempR.Address(External:=True)
'For Each r In TempR
'Debug.Print r.Value
'Next r
InputRange.Rows(1).Select
Theheaders = GetColumns2String(Selection)
        If maxcol = 1 Then
        MsgBox "no data here"
        Exit Sub
        End If
For i = 2 To maxrow
InputRange.Rows(i).Select
TheData = TheData & "(" & GetData2String(Selection) & ")"
If Not i = maxrow Then TheData = TheData & ", "
Next i
'Debug.Print Theheaders
'Debug.Print "____" & TheData

TheS = "INSERT INTO " & ComboTables.Value & "(" & Theheaders & ") Values " & TheData
Debug.Print TheS
End Sub

Sub Excel2SQL_Importer(ImportRange As Range, TheTable As String)
Dim i, z, m, maxrow, maxcol As Long
Dim HeadSTR, ContentSTR, RowStr, SIngleSTR, TheS As String
    maxcol = ImportRange.Columns.Count
    maxrow = ImportRange.Rows.Count
    HeadSTR = GetColumns2String(ImportRange.Rows(1))

For i = 2 To maxrow
    ContentSTR = ContentSTR & "(" & GetData2String(ImportRange.Rows(i)) & ")"
    If i Mod 1000 = 0 Or i = maxrow Then '-------
        TheS = "INSERT INTO " & TheTable & " (" & HeadSTR & ") VALUES " & ContentSTR
        ContentSTR = ""
       Debug.Print i
    
        MS_SQLData2 TheS
    Else
        ContentSTR = ContentSTR & ", "
    End If
Next i
'Debug.Print "Data imported"
End Sub
