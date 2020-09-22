Attribute VB_Name = "modISO9660"
'####################################################
'####################ISO MODULE######################
'####################################################
'#################by RadeonMaster####################
'####################################################
'######I worked really hard on this code#############
'#######and it wasn't easy, to find out##############
'######how ISO9660 works.               #############
'#####PLZ VOTE! THX.                     ############
'#''''----------------------------------''''''''''''#
'####################################################


Private Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String
Const ChunkSize = 2048     'This is the normal size of a cd-rom sector in bytes

Public Function IsIso(ByVal FileName As String) As Boolean
    Dim Data As String
    Dim Buffer As String
    'If the file exists...
    If FileExists(FileName) Then
        FFile = FreeFile
        'open it, ...
        Open FileName For Binary As #FFile
        Do Until i = 32774
                Data = ""
                'calculate needed space for the variable, ...
                If 32774 - Loc(FFile) < ChunkSize Then     '(32774 because we want read the first 16 sectors (ChunkSize * 16 + 6 bytes of the header))
                    Data = String(32774 - Loc(FFile), 0)
                Else
                    Data = String(ChunkSize, 0)
                End If
                'get the chunk, ...
                Get #FFile, , Data
                'and add it to the buffer
                Buffer = Buffer & Data
                i = i + Len(Data)
                'Do a DoEvents, so it don't crash
                DoEvents
        'get the next 2048 byte
        Loop
        'If the 5 last chars are "CD001", it is an iso...
        If Right(LCase(Buffer), 5) = "cd001" Then
            IsIso = True
        Else
        'if not, return false
            IsIso = False
        End If
        'close the file
        Close #FFile
    End If
    'clear variables
    Buffer = ""
    Data = ""
End Function

Public Function OpenISO9660(ByVal FileName As String) As Variant
    Dim Data As String
    Dim Buffer As String
    Dim arrx(0 To 14) As String
    FFile = FreeFile
    Open FileName For Binary As #FFile
        'we can say, the first 33648 bytes are the general information block
        Do Until i = 33648
                Data = ""
                'calculate needed space for the variable, ...
                If 33648 - Loc(FFile) < ChunkSize Then     '(32648 because we want read all the main informations in the iso)
                    Data = String(33648 - Loc(FFile), 0)
                Else
                    Data = String(ChunkSize, 0)
                End If
                'get the chunk, ...
                Get #FFile, , Data
                'and add it to the buffer
                Buffer = Buffer & Data
                i = i + Len(Data)
                'Do a DoEvents, so it don't crash
                DoEvents
        'get the next 2048 byte
        Loop
        'Start stripping informations
        'get volume descriptor (title), the lenght of the vd can be max. 32 byte
        arrx(0) = Trim(Mid(Buffer, 32777, 32))
        'get system descriptor
        arrx(1) = Trim(Mid(Buffer, 32809, 32))
        'get num. of sectors
        arrx(2) = CVLLittleEndian(Left(Mid(Buffer, 32849, 8), 4))
        'get size of Path Table
        arrx(3) = ConvertSize(CVLLittleEndian(Left(Mid(Buffer, 32893, 8), 4)))
        'get volume set identifier
        arrx(4) = Trim(Mid(Buffer, 32959, 128))
        'get publisher identifier
        arrx(5) = Trim(Mid(Buffer, 33087, 128))
        'get data preparer identifier
        arrx(6) = Trim(Mid(Buffer, 33215, 128))
        'get application identifier
        arrx(7) = Trim(Mid(Buffer, 33343, 128))
        'get copyright file identifier
        arrx(8) = Trim(Mid(Buffer, 33471, 37))
        'get abstract file identifier
        arrx(9) = Trim(Mid(Buffer, 33508, 37))
        'get bibliographic file identifier
        arrx(10) = Trim(Mid(Buffer, 33545, 37))
        'get date and time of volume creation
        arrx(11) = MakeDateString(Trim(Mid(Buffer, 33582, 16)))
        'get date and time of most recent modification
        arrx(12) = MakeDateString(Trim(Mid(Buffer, 33599, 16)))
        'get date and time when volume expires
        arrx(13) = MakeDateString(Trim(Mid(Buffer, 33616, 16)))
        'get date and time when volume is effective
        arrx(14) = MakeDateString(Trim(Mid(Buffer, 33633, 16)))
        'close the file
        Close #FFile
    'clear variables
    Buffer = ""
    Data = ""
    'return the array
    OpenISO9660 = arrx
End Function

Private Function FileExists(FileName As String) As Boolean
    On Error GoTo ErrHandle
        'If file does Not exist, there will be an Error
        FFile = FreeFile
        Open FileName For Input As #FFile
        Close #FFile
        'no error, file exists
        FileExists = True
    Exit Function
ErrHandle:
        'error, file does Not exist
        FileExists = False
    Exit Function
End Function

'Convert Little Endian to Long
'needed for some values in an ISO
Private Function CVLLittleEndian(Y As String) As Long
    Dim temp As Double
    temp = Asc(Mid(Y, 1, 1)) + 256# * Asc(Mid(Y, 2, 1)) + 65536# * Asc(Mid(Y, 3, 1)) + 16777216# * Asc(Mid(Y, 4, 1))
    If temp > 2147483647# Then
        CVLLittleEndian = CLng(temp - 4294967296#)
    Else
        CVLLittleEndian = CLng(temp)
    End If
End Function

'This will generate a better to read date&time string from the datestring in the iso
'The date in an ISO looks like this: 1998112408020000
'This function will convert it to this: 1998, 11, 24 08:02:00
Private Function MakeDateString(ByVal sDate As String) As String
    On Error Resume Next
    Dim sYears As Integer
    Dim sMonths As Integer
    Dim sDays As Integer
    Dim sHours As Integer
    Dim sMin As Integer
    Dim sSec As Integer
    Dim sMilli As Integer
    'extract the years
    sYears = Left(sDate, 4)
    'extract the months
    sMonths = Mid(sDate, 5, 2)
    'extract the days
    sDays = Mid(sDate, 7, 2)
    'and so on...
    sHours = Mid(sDate, 9, 2)
    sMin = Mid(sDate, 11, 2)
    sSec = Mid(sDate, 13, 2)
    sMilli = Mid(sDate, 15, 2)
    'now generate a good to read string
    MakeDateString = sYears & ", " & sMonths & ", " & sDays & " " & sHours & ":" & sMin & ":" & sSec & ":" & sMilli
End Function

'This will convert bytes into KB or MB
Private Function ConvertSize(ByVal Amount As Long) As String
    Dim Buffer As String
    Dim Result As String
    'Create the buffer
    Buffer = Space$(255)
    'Format the ByteSize
    Result = StrFormatByteSize(Amount, Buffer, Len(Buffer))
    'convert the bytes into KB or MB
    If InStr(Result, vbNullChar) > 1 Then ConvertSize = Left$(Result, InStr(Result, vbNullChar) - 1)
End Function
