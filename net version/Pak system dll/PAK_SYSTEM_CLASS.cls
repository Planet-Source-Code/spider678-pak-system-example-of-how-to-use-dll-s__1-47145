VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PAK_SYSTEM_CLASS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'######################################################################################
'######################################################################################
'######################################################################################
'#######################     PAK SYSTEM MADE BY SPIDER     ############################
'######################     spider_1027@btopenworld.com     ###########################
'######################################################################################
'######################################################################################
'######################################################################################

'This pak system has three functions that need to be called:
'OPENPAK(Filepath As String)
'ADD_FILE(qTITLE As String, Qfile() As Byte)
'EXTRACT_FILE(U As Long)
'to create a new pak file just specify a path and if it dont exist it will be created
'Note our pak system can not support files inside it over  1.95577740669251 gigabytes



Private Type PAK_FILE_SYSTEM
    Title As String
    Size As Long
End Type

Private Type ERROR_TYPES
    PAK_NOT_FOUND As String
    FILE_NOT_FOUND As String
    EMPTY_PAK As String
    NOT_VALID_PAK As String
    DONE_OPERATION As String
    NO_PAK_LOADED As String
    PATH_FILE_ACCESS_ERROR As String
    OBJECT_OUT_OF_BOUNDS As String
    FILE_ALREADY_EXISTS As String
End Type

Private A As ERROR_TYPES
Private PAK() As PAK_FILE_SYSTEM

Private PAK_FILEPATH As String
Private VALID_PAK_LOADED As Long
Private CHICKEN As String
Private STR As String
Private Pcount As Long

Public Function RETURN_TITLE(XX As Long) As String
'>>>> Make sure a valid pak is loaded
    If VALID_PAK_LOADED <> 1 Then
    RETURN_TITLE = A.NO_PAK_LOADED
    Exit Function
    End If

'>>>> Get title of specified index number
    If XX >= 0 And XX <= (Pcount - 1) Then
    RETURN_TITLE = PAK(XX).Title
    Else
    RETURN_TITLE = (A.OBJECT_OUT_OF_BOUNDS)
    End If
End Function

Public Function RETURN_SIZE(XX As Long) As String
'>>>> Make sure a valid pak is loaded
    If VALID_PAK_LOADED <> 1 Then
    RETURN_SIZE = (A.NO_PAK_LOADED)
    Exit Function
    End If

'>>>> Get size of specified file
    If XX >= 0 And XX <= Pcount Then
    RETURN_SIZE = PAK(XX).Size
    Else
    RETURN_SIZE = (A.OBJECT_OUT_OF_BOUNDS)
    End If
End Function
Public Function RETURN_NUM_LOADED() As String
RETURN_NUM_LOADED = 0
'>>>> Make sure a valid pak is loaded
    If VALID_PAK_LOADED <> 1 Then
    RETURN_NUM_LOADED = (A.NO_PAK_LOADED)
    Exit Function
    End If
'>>>>
    RETURN_NUM_LOADED = (Pcount - 1)
End Function

Public Function OPENPAK(Filepath As String) As String
VALID_PAK_LOADED = 0
Dim PrevFileSize As Long
Dim HEADER(254) As Byte
ReDim PAK(0)
Pcount = 0
'Form1.ListView1.ListItems.Clear
PAK_FILEPATH = Filepath

On Error GoTo WEGOTERROR
'>>>> Open file
Close #1
Open Filepath For Binary As #1
'>>>> If pak has no header start a new one
    If LOF(1) = 0 Then
    Call CREATE_PAK(PAK_FILEPATH)
    End If
    
Close #1
Open Filepath For Binary As #1

'>>>> Get header
    Get #1, , HEADER()

'>>>> Make sure header is valid
    If ARRAY2STRING(HEADER()) <> "HEADER_SPIDER_PAK_SYSTEM_GEN_001//:EN_GB_<>ENGINE_LOCK_CODE<79546-10346-99148-32094-61232>|<<¬EOF!This header file belongs to the spider pak system. spider_1027@btopenworld.com  !!!Warning!!! [modification of this will corrupt it.]  ACAXqaD £RWERDGUM&465m" Then
    OPENPAK = (A.NOT_VALID_PAK)
    Close #1
    Exit Function
    End If
    
VALID_PAK_LOADED = 1

'>>>> If Empty pak
    If LOF(1) = 255 Then
    OPENPAK = (A.EMPTY_PAK)
    Close #1
    Exit Function
    End If

'>>>> Loop until we run out of file headers to get
    Do Until Seek(1) + PrevFileSize >= LOF(1)
    If Seek(1) > LOF(1) Then Exit Function 'second check
'>>>> Get a sector
    Call GET_SECTOR
'>>>> Keep track of our previous file size
    PrevFileSize = PAK(Pcount - 1).Size
Loop

OPENPAK = (A.DONE_OPERATION)
Close #1

Exit Function

WEGOTERROR:
If Err.Number = 76 Then OPENPAK = (A.PAK_NOT_FOUND)
If Err.Number = 75 Then OPENPAK = (A.PATH_FILE_ACCESS_ERROR)
'MsgBox Err.Description


Close #1
End Function

Public Function ADD_FILE(qTITLE As String, FILE2ADD As String) As String
Dim Qfile() As Byte
Dim VV As Long
Dim GGG As String
STR = ""

On Error GoTo WEGOTERROR

'>>>> Make sure we dont get duplicate files
    For i = 0 To Pcount - 1
        If PAK(i).Title = CUT_SPACES(qTITLE) Then
        ADD_FILE = A.FILE_ALREADY_EXISTS
        Exit Function
        End If
    Next i

'>>> Get file adding into an array
    Open FILE2ADD For Binary As #9
        If LOF(9) = 0 Then
        ADD_FILE = (A.FILE_NOT_FOUND)
        Exit Function
        End If
    ReDim Qfile(LOF(9))
    Get #9, , Qfile()
    Close #9


If VALID_PAK_LOADED <> 1 Then
ADD_FILE = (A.NO_PAK_LOADED)
Exit Function
End If


Open PAK_FILEPATH For Binary As #1

'>>>> Our seek
    VV = LOF(1) + 1

'>>>> Our size according to what we adding
    CHICKEN = LONG2STR(UBound(Qfile) + 1) 'offset binary = 0 start
    STR = CHICKEN
    Put #1, VV, STR
    VV = VV + Len(STR)

'>>>> get our title and make sure our title is exactly 30 chars long
    STR = qTITLE
    GGG = FIXED_len_STR(STR, 255)
    Put #1, VV, GGG
    VV = VV + Len(GGG)
    
'>>>> Put our file
    Put #1, VV, Qfile

Close #1

'>>>> Refresh
    Call OPENPAK(PAK_FILEPATH)
    
'>>>> Error and messaging handling
    Exit Function
WEGOTERROR:
    If Err.Number = 76 Then ADD_FILE = (A.PAK_NOT_FOUND)
    If Err.Number = 75 Then ADD_FILE = (A.PATH_FILE_ACCESS_ERROR)
    MsgBox Err.Number
    MsgBox Err.Description
End Function

Public Function EXTRACT_FILE(Name2Find As String, Destpath As String) As String
Dim U As Long
Dim OLDlens As Long
Dim WWW() As Byte

'>>>> Make sure we have a valid pak open first
    If VALID_PAK_LOADED <> 1 Then
    EXTRACT_FILE = (A.NO_PAK_LOADED)
    Exit Function
    End If

'>>>> Find the file that we want out
    For i = 0 To Pcount - 1
        If PAK(i).Title = CUT_SPACES(Name2Find) Then
        U = i
        GoTo 111
        End If
    Next i
    EXTRACT_FILE = (A.FILE_NOT_FOUND)
    Exit Function

111
OLDlens = 256 'size of header and godnos
'U = U - 1 '???

'>>>> Get the seek point by adding up all previous file sizes
    For i = 0 To U
    OLDlens = OLDlens + PAK(i).Size + 255 + 4
    Next i

'>>>> ???? size adjusts or somethin
    OLDlens = OLDlens - PAK(U).Size - 1

'>>>> Make the array that the file will fill
    ReDim WWW(PAK(U).Size - 1)

'>>>> Open and get file into array
    Open PAK_FILEPATH For Binary As #1
    Get #1, OLDlens + 1, WWW

'>>>> Open destination file location and put it there
    'Kill (App.Path & "\extraction_temp_file.bmp")
    Open Destpath For Binary As #2 '& "\" & CUT_SPACES(Name2Find)
    Put #2, , WWW
    Close #2

'>>>> Close pak again
    Close #1

EXTRACT_FILE = (A.DONE_OPERATION)
End Function

'======================================================================================
'======================================================================================
'Private Functions  ===================================================================
'======================================================================================
'======================================================================================
Private Sub CREATE_PAK(FilePath2Create As String)
    Dim TEMPSTR As String
    Dim GG As String

'>>>> Delete file
    Close #1
    Open FilePath2Create For Binary As #1
    Close #1
    Kill FilePath2Create

'>>>> Create a new file with the pak header
    Open FilePath2Create For Binary As #1
    TEMPSTR = "HEADER_SPIDER_PAK_SYSTEM_GEN_001//:EN_GB_<>ENGINE_LOCK_CODE<79546-10346-99148-32094-61232>|<<¬EOF!This header file belongs to the spider pak system. spider_1027@btopenworld.com  !!!Warning!!! [modification of this will corrupt it.]  ACAXqaD £RWERDGUM&465m"

'>>>> Add a notes.txt file in our pak
    STR = STR & "This file has been automatically generated by spider's pak maker. (spider_1027@btopenworld.com)." & Chr(13) & Chr(10) & "The pak system is controlled by calls to the dll which handles all of the functions for you." & Chr(13) & Chr(10) & "The dll has the following functions"
    STR = STR & "ADD_FILE" & Chr(13) & Chr(10)
    STR = STR & "EXTRACT_FILE" & Chr(13) & Chr(10)
    STR = STR & "OPENPAK" & Chr(13) & Chr(10)
    STR = STR & "RETURN_NUM_LOADED" & Chr(13) & Chr(10)
    STR = STR & "RETURN_SIZE" & Chr(13) & Chr(10)
    STR = STR & "RETURN_TITLE" & Chr(13) & Chr(10)
    GG = STR
    TEMPSTR = TEMPSTR & LONG2STR(Len(GG))
    TEMPSTR = TEMPSTR & FIXED_len_STR("Notes.txt", 255)
    STR = GG
    STR = (TEMPSTR & STR)
    Put #1, , STR
    Close #1
End Sub
Private Function CUT_SPACES(LookSTR As String) As String
'>>>> Trunucate string to a string with no spaces
For i = Len(LookSTR) To 1 Step -1
    If Mid(LookSTR, i, 1) <> Chr(32) Then
    CUT_SPACES = Mid(LookSTR, 1, i)
    Exit Function
    End If
Next i
End Function

Private Function FIXED_len_STR(D As String, E As Long) As String
Dim ExtStr As String
'>>>> Convert any size string under or over a certain length the length you specify
If Len(D) <> E Then
    If Len(D) > E Then 'if over cut short
    ExtStr = RETURN_EXTENSION(D, E) 'get extension if there is one
    D = Mid(D, 1, E) 'take overlong out
    D = Mid(D, 1, Len(D) - Len(ExtStr)) & ExtStr 'add new filename with extension
    End If
    
    If Len(D) < E Then '> If under add spaces (chr 32)
        Do Until Len(D) = E
        D = D & Chr(32)
        Loop
    End If
End If

FIXED_len_STR = D
End Function
Private Function RETURN_EXTENSION(Mstr As String, MAX_LEN_EVER As Long) As String
Dim RTNstr As String
    
    For i = Len(Mstr) To 1 Step -1
'>>>> If a dot then get from there onwards as an extension
    If Mid(Mstr, i, 1) = "." Then
    RTNstr = Mid(Mstr, i)
    If Len(RTNstr) > MAX_LEN_EVER Then RTNstr = Mid(RTNstr, 1, MAX_LEN_EVER)
    RETURN_EXTENSION = RTNstr
    Exit Function
    End If
    Next i
'>>>> If no dot then no extension
    RETURN_EXTENSION = ""
End Function

Private Sub GET_SECTOR()
Dim Qsize As Long
Dim qTITLE As String

Dim Size(3) As Byte
Dim Title(254) As Byte
Dim YYYu As Long

'>>>> ??
    If Pcount > 0 Then YYYu = PAK(Pcount - 1).Size '}}

'>>>> Get size bytes and title
    Get #1, (Seek(1) + YYYu), Size() '}}
    Get #1, Seek(1), Title()

'>>>> Convert data to the right format and save to our virtual file system
    PAK(Pcount).Size = STR2LONG(ARRAY2STRING(Size))
    PAK(Pcount).Title = CUT_SPACES(ARRAY2STRING(Title))
'>>>> Dont add blank files(??)
    If Len(PAK(Pcount).Size) = 0 Then Exit Sub
'>>>> Add more space to our file system for our next file header
    Pcount = Pcount + 1
    ReDim Preserve PAK(Pcount)
End Sub

Private Function ARRAY2STRING(A() As Byte) As String
Dim Fstr As String
    For i = 0 To UBound(A)
    Fstr = Fstr & Chr(A(i))
    Next i
    ARRAY2STRING = Mid(Fstr, 1, Len(Fstr))
End Function

Private Function STR2LONG(tSTR As String) As Long
Dim A(3) As Long
Dim H As Long
H = 255

'>>>> Turn ascii values from string into numbers
    A(0) = Asc(Mid(tSTR, 1, 1)) '1-255
    A(1) = Asc(Mid(tSTR, 2, 1)) ' * H  'over 255 to 255*255
    A(2) = Asc(Mid(tSTR, 3, 1)) ' * H * H  '
    A(3) = Asc(Mid(tSTR, 4, 1)) ' * H * H * H '
'>>>> Compensate for not using 0 char in string
    A(0) = A(0) - 1
    A(1) = A(1) - 1
    A(2) = A(2) - 1
    A(3) = A(3) - 1
'>>>> Now multiply each value to its previous value
    A(1) = A(1) * H
    A(2) = A(2) * H * H
    A(3) = A(3) * H * H * H

'>>>> Add all values back together
    STR2LONG = A(0) + A(1) + A(2) + A(3)
End Function

Private Function LONG2STR(LL As Long) As String
Dim H As Long
Dim A(3) As Long
H = 255

'>>>> Turn a string into 4 bytes by seeing how much will go in each column
1 '>>>> the 255 * 255 * 255 column
    If LL > H * H * H Then
    LL = LL - H * H * H
    A(3) = A(3) + 1
    GoTo 1
End If
2 '>>>> the 255 * 255 column
    If LL > H * H Then
    LL = LL - H * H
    A(2) = A(2) + 1
    GoTo 2
End If
3 '>>>> the 255 column
    If LL > H Then
    LL = LL - H
    A(1) = A(1) + 1
    GoTo 3
End If

'>>>> The singles column (single values ie. 54,67,150 etc)
    A(0) = LL

'>>>> Compensate for not using the char 0
    For i = 0 To 3
    A(i) = A(i) + 1
    Next i
787
'>>>> Do our maths like manually (make sure no bigger than 256)
    If A(0) > 255 Then
    A(0) = 1
    A(1) = A(1) + 1
    GoTo 787
    End If
786
    If A(1) > 255 Then
    A(1) = 1
    A(2) = A(2) + 1
    GoTo 786
    End If
785
    If A(2) > 255 Then
    A(2) = 1
    A(3) = A(3) + 1
    GoTo 785
    End If

'>>>> An overflow error that we never seem to get? maybe long is a bit small for this?
    If A(3) > 255 Then
    'MsgBox "YOU ARE TAKING THE PISS ARENT YA THATS 1.5 GIG", vbCritical
    Exit Function
    End If

'>>>> Convert broken up longs into ascii
    STR = ""
    STR = STR & Chr(A(0))
    STR = STR & Chr(A(1))
    STR = STR & Chr(A(2))
    STR = STR & Chr(A(3))

LONG2STR = STR
End Function


Private Sub Class_Initialize()
    A.PAK_NOT_FOUND = "PAK_NOT_FOUND"
    A.FILE_NOT_FOUND = "FILE_NOT_FOUND"
    A.EMPTY_PAK = "EMPTY_PAK"
    A.NOT_VALID_PAK = "NOT_VALID_PAK"
    A.DONE_OPERATION = "DONE_OPERATION"
    A.NO_PAK_LOADED = "NO_PAK_LOADED"
    A.PATH_FILE_ACCESS_ERROR = "PATH_FILE_ACCESS_ERROR"
    A.OBJECT_OUT_OF_BOUNDS = "OBJECT_OUT_OF_BOUNDS"
    A.FILE_ALREADY_EXISTS = "FILE_ALREADY_EXISTS"
End Sub
'######################################################################################
'######################################################################################
'######################################################################################
'######################################################################################
'######################################################################################
'######################################################################################

