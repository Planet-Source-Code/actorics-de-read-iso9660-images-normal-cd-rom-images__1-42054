VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   0  'Kein
   Caption         =   "Open ISO 9660 Image Files"
   ClientHeight    =   825
   ClientLeft      =   -90
   ClientTop       =   -660
   ClientWidth     =   5235
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      BackColor       =   &H00DDDDDD&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   5145
      TabIndex        =   33
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton cmdExit 
         Caption         =   "X"
         Height          =   255
         Left            =   4920
         TabIndex        =   35
         Top             =   50
         Width           =   255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Open ISO 9660 Image Files"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   3735
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   4680
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Information"
      Height          =   3615
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   4695
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Volume Indentifier:"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label lblVolume 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1680
         TabIndex        =   31
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "System Identifier:"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label lblSystem 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1560
         TabIndex        =   29
         Top             =   480
         Width           =   45
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Num. of sectors:"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label lblNumSectors 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1560
         TabIndex        =   27
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Path Table Length:"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label lblPathTableLength 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1680
         TabIndex        =   25
         Top             =   960
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Volume Set Identifier:"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   1200
         Width           =   1500
      End
      Begin VB.Label lblVolumeSet 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1920
         TabIndex        =   23
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Publisher Identifier:"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblPublisher 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1680
         TabIndex        =   21
         Top             =   1440
         Width           =   45
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Data Preparer:"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   1680
         Width           =   1035
      End
      Begin VB.Label lblDataPrep 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1440
         TabIndex        =   19
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "App. Identifier:"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   1020
      End
      Begin VB.Label lblApp 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1440
         TabIndex        =   17
         Top             =   1920
         Width           =   45
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Copyright file:"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   2160
         Width           =   945
      End
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1440
         TabIndex        =   15
         Top             =   2160
         Width           =   45
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "abstract file:"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label lblAbstract 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1200
         TabIndex        =   13
         Top             =   2400
         Width           =   45
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Bibliographic file:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   2640
         Width           =   1185
      End
      Begin VB.Label lblBibliographic 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1560
         TabIndex        =   11
         Top             =   2640
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Date and Time of Creation:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   2880
         Width           =   1905
      End
      Begin VB.Label lblDateTimeCreation 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2280
         TabIndex        =   9
         Top             =   2880
         Width           =   45
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Date and Time when expires:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   3120
         Width           =   2070
      End
      Begin VB.Label lblDateTimeExpires 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2400
         TabIndex        =   7
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Date and Time when effective:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   3360
         Width           =   2190
      End
      Begin VB.Label lblDateTimeEffective 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2520
         TabIndex        =   5
         Top             =   3360
         Width           =   45
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Go!"
      Height          =   270
      Left            =   4080
      TabIndex        =   3
      Top             =   450
      Width           =   420
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   450
      Width           =   375
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   435
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   285
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code will open an ISO 9660 image file (*.iso) and get information about it.
'THIS WILL ONLY WORK WITH *.ISO! NO OTHER IMAGE FORMAT SUPPORTED!
'The OpenISO function will return the informations about the iso as an array.
'NO API!!!

Dim arrx As Variant

Private Sub cmdBrowse_Click()
    'select a file
    Dialog.DialogTitle = "Open ISO file"
    Dialog.Filter = "ISO files (*.iso)|*.iso|All Files (*.*)|*.*"
    'show open-dialog
    Dialog.ShowOpen
    'check if user selected a file
    If Dialog.FileName = vbNullString Then
        'if not then exit sub
        Exit Sub
    Else
        'if yes, get new filename
        txtFileName = Dialog.FileName
    End If
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdStart_Click()
    txtFileName.SetFocus
    'Check if it is an ISO
    If IsIso(txtFileName) Then
        'if yes get info
        arrx = OpenISO9660(txtFileName)
        'set captions of labels
        'system to use cd-rom
        lblSystem.Caption = arrx(0)
        'title of the cd-rom
        lblVolume.Caption = arrx(1)
        'number of sectors
        lblNumSectors.Caption = arrx(2)
        'path table length
        lblPathTableLength.Caption = arrx(3)
        'volume set identifier
        lblVolumeSet.Caption = arrx(4)
        'publisher identifier
        lblPublisher.Caption = arrx(5)
        'data preparer identifier
        lblDataPrep.Caption = arrx(6)
        'application (burning-software) identifier
        lblApp.Caption = arrx(7)
        'copyright file identifier
        lblCopyright.Caption = arrx(8)
        'abstract file identifier
        lblAbstract.Caption = arrx(9)
        'bibliographic file identifier
        lblBibliographic.Caption = arrx(10)
        'date and time of creation
        lblDateTimeCreation.Caption = arrx(11)
        'arrx(12) can be implemented like the others, but it isn't important
        'it just shows the date of the most recent modification
        
        'date and time when volume expires
        lblDateTimeExpires.Caption = arrx(13)
        'date and time when volume is effective
        lblDateTimeEffective.Caption = arrx(14)
        'show full form
        Me.Height = 5025
    Else
        'if it is not an ISO, tell the user
        MsgBox "Selected file is not an ISO!" & vbCrLf & "Check the following:" & vbCrLf & "   Does the file exist?" & vbCrLf & "   Is it a normal ISO?", vbExclamation, "Error"
    End If
End Sub

'This will draw a gray gradient
Private Sub DrawBackGround()
    Const intBLUESTART% = 255
    Const intBLUEEND% = 10
    Const intBANDHEIGHT% = 2
    Const intRed% = 150
    Const intGreen% = 150
    Const intBlue% = 255
    Const intBackRed% = 8
    Const intBackGreen% = 16
    Const intBackBlue% = 32
    Dim sngBlueCur As Single
    Dim sngBlueStep As Single
    Dim intFormHeight As Integer
    Dim intFormWidth As Integer
    Dim intY As Integer
    Dim iColor As Integer
    Dim iRed As Single, iBlue As Single, iGreen As Single
    'Get system values for height and width
    intFormHeight = ScaleHeight
    intFormWidth = ScaleWidth
        iColor = intBlue
    'Calculate step size and blue start value
    sngBlueStep = intBANDHEIGHT * (intBLUEEND - intBLUESTART) / Me.Height
    sngBlueCur = intBLUESTART
    'Paint colors!!!!!!!
    For intY = 0 To intFormHeight Step intBANDHEIGHT
        If iColor And intBlue Then iBlue = sngBlueCur
        If iColor And intRed Then iRed = sngBlueCur
        If iColor And intGreen Then iGreen = sngBlueCur
        If iColor And intBackBlue Then iBlue = 255 - sngBlueCur
        If iColor And intBackRed Then iRed = 255 - sngBlueCur
        If iColor And intBackGreen Then iGreen = 255 - sngBlueCur
        Line (-1, intY - 1)-(intFormWidth, intY + intBANDHEIGHT), RGB(iRed, iGreen, iBlue), BF
        sngBlueCur = sngBlueCur + sngBlueStep
    Next intY
End Sub

Private Sub Form_Load()
    'Resize the picturebox
    Picture1.Width = Me.Width - 30
    'Let the form stay on top
    StayOnTop Me, True
    'add cool flat borders to the controls
    FlatBorder Me.hwnd
    FlatBorder cmdStart.hwnd
    FlatBorder cmdBrowse.hwnd
    FlatBorder cmdExit.hwnd
    FlatBorder Frame1.hwnd
    FlatBorder txtFileName.hwnd
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        FormDrag Me
    End If
End Sub

Private Sub Form_Paint()
    'PAINT A COOL GRAY GRADIENT!
    DrawBackGround
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'don't let the form stay at top anymore
    StayOnTop Me, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'don't let the form stay at top anymore
    StayOnTop Me, False
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Drag the form
    If Button = 1 Then
        FormDrag Me
    End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Drag the form
    If Button = 1 Then
        FormDrag Me
    End If
End Sub
