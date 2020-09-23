VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1_file_encr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encrypt File"
   ClientHeight    =   2460
   ClientLeft      =   3675
   ClientTop       =   3480
   ClientWidth     =   4680
   Icon            =   "Form1_file_encr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1_en_ok 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Encrypt"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton cmdEncrypt 
         Caption         =   "&Encrypt"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtEPassword 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox txtEFilename 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmdEOpen 
         Caption         =   "&Open"
         Height          =   285
         Left            =   3120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog cd1 
         Left            =   -120
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "Form1_file_encr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim eSource As String, eDestination As String, ePassword As String
Dim dSource As String, dDestination As String, dPassword As String

Dim WithEvents EN As clsBinaryEncryptor
Attribute EN.VB_VarHelpID = -1

Private Sub cmdEncrypt_Click()
Dim fso As FileSystemObject
Dim ts As TextStream
Dim ts1 As TextStream
Dim str As String

Dim time1 As String
Dim date1 As String
time1 = Time
date1 = Date

Set fso = New FileSystemObject
Set ts = fso.OpenTextFile(Form1_top.Text2.Text, ForAppending, True)
ts.WriteBlankLines (nm)


str = "File Encrypted by the user at " + time1 + " on " + date1
ts.WriteLine (str)
Dim str1 As String
Set ts1 = fso.OpenTextFile(Form1_top.Text2.Text, ForReading, False)
str1 = ts1.ReadAll

Form1_top.Text1.Text = str1
ts.Close
ts1.Close
Set fso = Nothing

Form1_file_encr.pb1.Value = 100

If txtEPassword.Text = "" Then
GoTo err1:
End If
Dim b As Boolean
Screen.MousePointer = vbHourglass
    b = EN.EncryptFile(eSource, eDestination, IIf(txtEPassword.Text = "", "default", txtEPassword.Text))

If b = True Then
    MsgBox "The file is encrypted successfully." & vbCrLf & "Please save your password.", vbInformation, "Encrypt File"
Else
    MsgBox "Error occured while encrypting the file. Please contact the software developer.", vbCritical, "Encrypt File"
End If

Screen.MousePointer = 0

Exit Sub
err1:
MsgBox "Please specify a password", vbCritical, "Encrypt Database"
End Sub

Private Sub cmdEOpen_Click()
With cd1
    .CancelError = True
    .Filter = "All Files *.*|*.*"
    .Flags = cdlOFNFileMustExist
    
    On Error GoTo OpenError
    .DialogTitle = "Open a File to encrypt."
    .ShowOpen
    eSource = .Filename
    
    On Error GoTo saveerror
    .DialogTitle = "Save the encrypted File."
    .ShowSave
    txtEFilename.Text = .Filename
    eDestination = .Filename
    
    txtEPassword.SetFocus
    txtEPassword.SelStart = 0
    txtEPassword.SelLength = Len(txtEPassword.Text)
    
    cmdEncrypt.Enabled = True
    
End With

Exit Sub
OpenError:
saveerror:

End Sub

Private Sub Command1_en_ok_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set EN = New clsBinaryEncryptor
End Sub
