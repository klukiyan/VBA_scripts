'THIS IS A UNIVERSAL MODULE THAT CONTAINS USEFUL UNIVERSAL MACROS AND FUNCTIONS THAT CAN BE USED TO HELP/SPEED UP PROJECT CREATION.
'IMPORT THIS MODULE INTO YOUR PROJECT AND USE THE BELOW MACROS AND ROUTINES
'DO NOT MODIFY OR EDIT ANY OF BELOW CONTENT.
'HERE IS THE FULL LIST: (YOU CAN COPY PASTE NAMES FROM HERE)

'Function AllFilePaths(YourFolder As String, IncludeSubFolders As Boolean, Optional WildString As String) As Variant
'Function FolderChoose() As String
'Function FolderExists(ByVal Path As String) As Boolean 'check if file or folder exists
'Function GetFExt(FilePath As String) As String
'Function GetLNPerson(uName As String) As Variant
'Function IsFileOpen(filename As String) as boolean
'Function LatestFilePath(folpath As String) As String
'Function ReadAllText(Target2Read As String) As String
'Function RealUNC(OldPath As String) As String
'Sub CreateText_UNI(UniPath As String, Optional uniContent As String)
'Sub OpenAFolder(folderfullpatH As String)
'Sub RemoveCaption(objForm As Object)
'Sub SendTheMail_UNI(vaRecipient, vaMsg, stSubject As Variant, server, database, stAttachment As String)
'Sub WriteText_UNI(ExistingPath As String, AddCOntent As String)

'details:
'Function AllFilePaths(YourFolder As String, IncludeSubFolders As Boolean, Optional WildString As String) As Variant
        'This function will return an array that conains paths to all the files in the specified folder "YourFolder"
        'If "IncludeSubFolders is set to true, then it will also return the paths to files from subfolders and their subfolders till the end of folder tree
        'If WildString contains something, then it will return files of particular type/file content.
        'WildString to be used same as in search function.
            'if you want to search only excel files of xlsx type, then wildstring should be "*.xlsx"
            'if any excel files then should be "*.xls*"
            'if all files containing word "Receipts" then "*Receipts*
            'if all receipts that are excel (xlsb format) then wildstring should be "*Receipts*.xlsb"
            

'Function FolderChoose() As String
'    Asks the user to selecte a FOLDER. Returns the path to selected folder.
'        If user cancels, returns empty string
'
'Function FolderExists(ByVal Path As String) As Boolean 'check if file or folder exists
'    Checks if the give path to file/or folder exists. If file or folder is present returns TRUE, otherwise FALSE
'
'Function GetFExt(FilePath As String) As String
'    Gets the extension of the given filename(or path) and returns as string (for example ".xlsx")
'
'Function GetLNPerson(uName As String) As Variant
'    Connects to Lotus notes and looks for first match of person by Name/Surname/Shortname/NameSurname/anything
'    Returns an Array where:
'            (0) = shortname
'            (1) = internet address
'            (2) = Name
'            (3) = Surname
'    For example to get my email I can use something like this:
'        GetLnPerson("Lukiyan")(1)
'        Result will be klukiyan@csc.com
'    If you want to access multiple fields, it is recommended to firstly put the whole area result and into antoher area,
'    and then target the array elements from it.
'    Because it will connect to lotus notes each time the function is used.
'    Let me know if you need more clarifications on this.
'
'Function IsFileOpen(filename As String) As Boolean
'        Checks if the excel workbook ("filename" - is full path) is open by another user. Returns TRUE OR FALSE

'Function LatestFilePath(folpath As String) As String
'    Returns a string which is a PATH to file from given folderpath which is latest modified.
'
'Function ReadAllText(Target2Read As String) As String
'    Reads ALL text from a given textfile path and returns the result as string value
'
'Function RealUNC(OldPath As String) As String
'    Converts any path to ful UNC path (x to \\cscdsppra...) [String input string output =) ]
'
'Sub CreateText_UNI(UniPath As String, Optional uniContent As String)
'    Creates a TEXT file under a given path
'    If optional Unicontent string is given, then also writes it into the newly created file.
'    WARNING: If a file already existed, it will OVERWRITE it
'
'Sub OpenAFolder(folderfullpatH As String)
'    Opens the given folder path as folder in WINDOWS explorer =)
'    just opens And that 's it
'
'Sub RemoveCaption(objForm As Object)
'    Removes the bluebar and blue borders from a userform.
'    Easiest is to put it into begining of "userform initialize" script. For example:
'    UserForm_Initialize
'    RemoveCaption Me
'        or can be Call Remove Caption (Me) parantheses depend same as with MsgBox funtion
'
'Sub SendTheMail_UNI(vaRecipient, vaMsg, stSubject As Variant, server, database, stAttachment As String)
'    This ine is clear. Sends the message with given conditions (supports attachments)
'
'Sub WriteText_UNI(ExistingPath As String, AddCOntent As String)
'    Adss a line, or many lines (string to be given) to an EXISTING text file. Text file has to be already craeted and exist. it will not overwrite it.
'    It will add the contents at the bottom
'        Practical example:
'        You want to adda line to a file, but you don't yet know whether it exists.
'        You use Folderexists function to check it, if it exists then you wrtieText_uni, if not, then use CreateText_Uni
'
'Combinations of above functions and macros and your macros can quickly bring powerful results.


Private Const RESOURCETYPE_ANY = &H0
Private Const RESOURCE_CONNECTED = &H1
Private Type NETRESOURCE
   dwScope As Long
   dwType As Long
   dwDisplayType As Long
   dwUsage As Long
   lpLocalName As Long
   lpRemoteName As Long
   lpComment As Long
   lpProvider As Long
End Type


Private Function DriveLetterToUNC(Optional DriveLetter As String = "C:") As String
   'converts a given drive letter to the mapped UNC of the local machine
   'eg DriveLetterToUNC("F:")
   '  returns "\\servername\drivename"
   '  or "F:" if not found

   Dim hEnum As Long
   Dim NetInfo(1023) As NETRESOURCE
   Dim entries As Long
   Dim nStatus As Long
   Dim LocalName As String
   Dim UNCName As String
   Dim i As Long
   Dim r As Long

   ' Begin the enumeration
   nStatus = WNetOpenEnum(RESOURCE_CONNECTED, RESOURCETYPE_ANY, _
      0&, ByVal 0&, hEnum)

   DriveLetterToUNC = DriveLetter

   'Check for success from open enum
   If ((nStatus = 0) And (hEnum <> 0)) Then
      ' Set number of entries
      entries = 1024

      ' Enumerate the resource
      nStatus = WNetEnumResource(hEnum, entries, NetInfo(0), _
         CLng(Len(NetInfo(0))) * 1024)

      ' Check for success
      If nStatus = 0 Then
         For i = 0 To entries - 1
            ' Get the local name
            LocalName = ""
            If NetInfo(i).lpLocalName <> 0 Then
               LocalName = Space(lstrlen(NetInfo(i).lpLocalName) + 1)
               r = lstrcpy(LocalName, NetInfo(i).lpLocalName)
            End If

            ' Strip null character from end
            If Len(LocalName) <> 0 Then
               LocalName = Left(LocalName, (Len(LocalName) - 1))
            End If

            If UCase$(LocalName) = UCase$(DriveLetter) Then
               ' Get the remote name
               UNCName = ""
               If NetInfo(i).lpRemoteName <> 0 Then
                  UNCName = Space(lstrlen(NetInfo(i).lpRemoteName) + 1)
                  r = lstrcpy(UNCName, NetInfo(i).lpRemoteName)
               End If

               ' Strip null character from end
               If Len(UNCName) <> 0 Then
                  UNCName = Left(UNCName, (Len(UNCName) - 1))
               End If

               ' Return the UNC path to drive
               DriveLetterToUNC = Trim(UNCName)

               ' Exit the loop
               Exit For
            End If
         Next i
      End If
   End If

   ' End enumeration
   nStatus = WNetCloseEnum(hEnum)
End Function
Function RealUNC(OldPath As String) As String
Dim drive As String
Dim FInalPath As String
Dim realdrive As String

drive = Left(OldPath, 2)
 realdrive = DriveLetterToUNC(drive)
 FInalPath = Mid(OldPath, 3, 10000)
 RealUNC = realdrive & FInalPath
End Function






Function GetLNPerson(uName As String) As Variant
    'this function will look into Lotus Notes contact database for a given name and return the first matching user as an array as follows
    '(0) = shortname
    '(1) = internet address
    '(2) = Name
    '(3) = Surname
    '4 parameters in total =)
    Dim MailServer1 As String
    Dim MailDBPath As String
    
    MailServer1 = "EMEA-ML18/SRV/CSC"
    MailDBPath = "names.nsf"
    'above are the data
    
    'checking what is there
   
    Dim nDb As Object, doc As Object, session As Object, vw As Object
    Set session = CreateObject("Notes.NotesSession")
    Set nDb = session.GetDatabase(MailServer1, MailDBPath)
    Set vw = nDb.GetView("($Users)")
    'OK LET"S START THE ROLL NOW
    
    
    Set doc = vw.GetDocumentByKey(uName)
   
    If Not (doc Is Nothing) Then
        GetLNPerson = Array(doc.getitemvalue("shortname")(0), doc.getitemvalue("internetaddress")(0), doc.getitemvalue("FirstName")(0), doc.getitemvalue("LastName")(0))
      Else
        uName = ""
      End If

    Set session = Nothing
    Set nDb = Nothing
    Set vw = Nothing
    Set doc = Nothing
    
End Function

'universal add a line to existing text file
Sub WriteText_UNI(ExistingPath As String, AddCOntent As String)
Dim objfile As Object
Dim objfso As Object
Set objfso = CreateObject("Scripting.FileSystemObject")
Set objfile = objfso.OpenTextFile(ExistingPath, 8) 'here comes the path

                    objfile.WriteLine AddCOntent  'here comes the CODE
objfile.Close
End Sub


Sub RemoveCaption(objForm As Object) 'removes frame around USERFORM
    Dim lStyle          As Long
    Dim hMenu           As Long
    Dim mhWndForm       As Long
    
    If Val(Application.Version) < 9 Then
        mhWndForm = FindWindow("ThunderXFrame", objForm.Caption) 'XL97
    Else
        mhWndForm = FindWindow("ThunderDFrame", objForm.Caption) 'XL2000+
    End If
    lStyle = GetWindowLong(mhWndForm, -16)
    lStyle = lStyle And Not &HC00000
    SetWindowLong mhWndForm, -16, lStyle
    DrawMenuBar mhWndForm
End Sub

Function FolderChoose() As String 'select a folder by user
ChDir ("\\cscdsppra001.emea.globalcsc.net\SSC-Prague\")

Dim diaFolder As FileDialog
  Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    diaFolder.Show
    
    If diaFolder.SelectedItems.Count > 0 Then
    folderpath = diaFolder.SelectedItems(1)
    Else
    folderpath = ""
    End If
    
    'If diaFolder.Show = True Then folderpath = diaFolder.SelectedItems(1) Else folderpath = ""
    If Len(folderpath) > 0 Then FolderChoose = folderpath & "\"
Set diaFolder = Nothing
End Function

Sub OpenAFolder(folderfullpatH As String) 'opens and activates a folder as a window :)
Shell "explorer.exe" & " " & folderfullpatH, vbNormalFocus
End Sub

Function FolderExists(ByVal Path As String) As Boolean 'check if file or folder exists
  On Error Resume Next
  FolderExists = Dir(Path, vbDirectory) <> ""
End Function


Sub CreateText_UNI(UniPath As String, Optional uniContent As String)
'UNIVERSAL MACRO FOR CREATING NEW TEXT FILES
'WARNING: IT WILL OVERWRITE THE EXISTING FILES
Dim objfile As Object
Dim objfso As Object
Set objfso = CreateObject("Scripting.FileSystemObject")
Set objfile = objfso.CreateTextFile(UniPath, True) 'here we have the path

                        If Len(uniContent) > 0 Then
                        objfile.WriteLine uniContent 'here comes the CODE
                        End If
objfile.Close
End Sub

Sub Reading_test(Target2Read) 'this one is temporary
Dim objfile As Object
Dim objfso As Object
Dim curline As String
Set objfso = CreateObject("Scripting.FileSystemObject")
Set objfile = objfso.OpenTextFile(Target2Read, 1)
Dim i As Long
i = 1
                Do Until objfile.AtEndOfStream
                            curline = objfile.ReadLine
                            If InStr(curline, "3") > 0 Then
                            'Debug.Print curline
                            End If
                            
                                'StrLine = objfile.ReadALL 'this one reads everything at once
                                'strLine = objfile.ReadLine 'this one reads line by line
                                        'code comes here
                Loop
objfile.Close
End Sub

Function ReadAllText(Target2Read As String) As String

Dim objfile As Object
Dim objfso As Object
Set objfso = CreateObject("Scripting.FileSystemObject")
Set objfile = objfso.OpenTextFile(Target2Read, 1)
        If objfile.AtEndOfStream Then
        ReadAllText = ""
        Else
        ReadAllText = objfile.ReadAll
        End If
objfile.Close
End Function

Function GetFExt(FilePath As String) As String
GetFExt = Mid(FilePath, InStrRev(FilePath, ".") + 1, 100)
End Function

Function LatestFilePath(folpath As String) As String
    Dim MyPath As String
    Dim MyFile As String
    Dim LatestFile As String
    Dim LatestDate As Date
    Dim LMD As Date
    
    'Specify the path to the folder
    MyPath = folpath
    
    'Make sure that the path ends in a backslash
    If Right(MyPath, 1) <> "\" Then
    MsgBox "Not a folder"
    LatestFilePath = ""
    Exit Function
    End If
    
    'Get the first Excel file from the folder
    MyFile = Dir(MyPath & "*.*", vbNormal)
    
    'If no files were found, exit the sub
    If Len(MyFile) = 0 Then
        'MsgBox "No files were found...", vbExclamation
        LatestFilePath = ""
        Exit Function
    End If
    
    'Loop through each Excel file in the folder
    Do While Len(MyFile) > 0
    
        'Assign the date/time of the current file to a variable
        LMD = FileDateTime(MyPath & MyFile)
        
        'If the date/time of the current file is greater than the latest
        'recorded date, assign its filename and date/time to variables
        If LMD > LatestDate Then
            LatestFile = MyFile
            LatestDate = LMD
        End If
        
        'Get the next Excel file from the folder
        MyFile = Dir
        
    Loop
    
    LatestFilePath = MyPath & LatestFile
End Function

Sub SendTheMail_UNI(vaRecipient, vaMsg, stSubject As Variant, server, database, stAttachment As String)

    Dim noSession As Object, noDatabase As Object, noDocument As Object
    Dim obAttachment As Object, EmbedObject As Object
     

     
    Const EMBED_ATTACHMENT As Long = 1454
     
'recipient
        On Error Resume Next 'GoTo SendMailError


         'Instantiate the Lotus Notes COM's  Objects.
        Set noSession = CreateObject("Notes.NotesSession")
        Set noDatabase = noSession.GetDatabase(server, database)
         'If Lotus Notes is not open then open the mail-part of it.
        If noDatabase.IsOpen = False Then noDatabase.OPENMAIL
         'Create the e-mail and the attachment.
        Set noDocument = noDatabase.CreateDocument
        Set obAttachment = noDocument.CreateRichTextItem("stAttachment")
        Set EmbedObject = obAttachment.EmbedObject(EMBED_ATTACHMENT, "", stAttachment)
         'Add values to the created e-mail main properties.
        With noDocument
            .Form = "Memo"
            .SendTo = vaRecipient
            .Subject = stSubject
            .Body = vaMsg
            .SaveMessageOnSend = True ' responsible for whether the message will stay in sent folder
            '.ReplyTo = "emea.gs@csc.com"   'OPTIONAL repsonsible for forced reply to (used with team mailboxes)
            '.Principal = "BATMAN" ' optional used for team maiboxes as well
        End With
         'Send the e-mail.
        With noDocument
            .PostedDate = Now() 'this command is repsonsible for inserting the date when the document will be sent
            .Send 0, vaRecipient 'this command can be .Save instead of Send, and it will save draft instead of sending
        End With
End Sub

Function IsFileOpen(filename As String) As Boolean
    Dim filenum As Integer, errnum As Integer

    On Error Resume Next   ' Turn error checking off.
    filenum = FreeFile()   ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open filename For Input Lock Read As #filenum
    Close filenum          ' Close the file.
    errnum = Err           ' Save the error number that occurred.
    On Error GoTo 0        ' Turn error checking back on.

    ' Check to see which error occurred.
    Select Case errnum

        ' No error occurred.
        ' File is NOT already open by another user.
        Case 0
         IsFileOpen = False

        ' Error number for "Permission Denied."
        ' File is already opened by another user.
        Case 70
            IsFileOpen = True

        ' Another error occurred.
        Case Else
            Error errnum
    End Select

End Function

Function AllFilePaths(YourFolder As String, IncludeSubFolders As Boolean, Optional WildString As String, Optional FolCreatedDays As Long) As Variant
Dim InnerAll() As String
Dim g As Variant
ReDim InnerAll(10000)
Dim i, z As Long
i = 0
  Dim fso As Object 'FileSystemObject
    Dim fldStart As Object 'Folder
    Dim fld As Object 'Folder
    Dim fl As Object 'File
    Dim Mask As String
                                    'hint:         If fl.Name Like Mask Then
    Set fso = CreateObject("scripting.FileSystemObject") ' late binding
    Set fldStart = fso.GetFolder(YourFolder)

        If Len(WildString) = 0 Then WildString = "*"
    For Each fl In fldStart.Files
        If fl.Name Like WildString And Not Left(fl.Name, 1) = "~" Then
            'Debug.Print fl.Path & "\" & fl.Name
            InnerAll(i) = fl.Path
            i = i + 1
        End If
    Next fl
    'what to do it fubfolders also included
            If IncludeSubFolders = True Then
                    'each subfolder will call itself
                    For Each fld In fldStart.subfolders
                            If FolCreatedDays > 0 Then
                                    If fld.DateLastModified > (Date - FolCreatedDays) Then
                                        g = AllFilePaths(fld.Path, True, WildString)
                                        If Len(g(0)) > 0 Then
                                            For z = 0 To UBound(g)
                                                InnerAll(i) = g(z)
                                                i = i + 1
                                            Next z
                                        End If
                                    End If
                            Else
                                     g = AllFilePaths(fld.Path, True, WildString)
                                        If Len(g(0)) > 0 Then
                                            For z = 0 To UBound(g)
                                                InnerAll(i) = g(z)
                                                i = i + 1
                                            Next z
                                        End If
                                    End If


                    Next fld
            End If
If i > 0 Then i = i - 1
ReDim Preserve InnerAll(0 To i)

AllFilePaths = InnerAll
End Function

Sub TestAbove()
Dim MyResults As Variant
Dim i As Long
MyResults = AllFilePaths("C:\Users\klukiyan\Desktop\", True, "*png")
For i = 0 To UBound(MyResults)
    Debug.Print (MyResults(i))
Next i
End Sub

Sub TestAbove2()
Debug.Print Now
Dim MyResults As Variant
Dim i As Long
MyResults = AllFilePaths("\\cscdsppra001.emea.globalcsc.net\projects\SSC-Prague\Finance\IC\Italy interco\7 RECONCIL\", True, "*FX Analysis*PD09*", 120)
For i = 0 To UBound(MyResults)
    Debug.Print (MyResults(i))
Next i
Debug.Print Now
End Sub
