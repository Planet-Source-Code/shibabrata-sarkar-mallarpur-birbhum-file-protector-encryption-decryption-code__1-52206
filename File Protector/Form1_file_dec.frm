VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1_file_dec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Protector"
   ClientHeight    =   4245
   ClientLeft      =   3675
   ClientTop       =   2130
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1_file_dec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   4680
   Begin VB.Frame Frame1 
      Caption         =   "Encrypt"
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton cmdEOpen 
         Caption         =   "&Open"
         Height          =   285
         Left            =   3120
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtEFilename 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtEPassword 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   840
         Width           =   3975
      End
      Begin VB.CommandButton cmdEncrypt 
         Caption         =   "&Encrypt"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1_de_ok 
      Caption         =   "OK"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Decrypt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   4455
      Begin VB.CommandButton cmdDOpen 
         Caption         =   "Open"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtDFilename 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtDPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   840
         Width           =   3975
      End
      Begin VB.CommandButton cmdDecrypt 
         Caption         =   "&Decrypt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   1200
         Width           =   1575
      End
      Begin MSComDlg.CommonDialog cd1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "Form1_file_dec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'File Protector by Shibabrata Sarkar Shibabrata_Sarkar@rediffmail.com
'JCS Limited http://www.jayasoftwares.com

'If you are using VB6, use this:
Private Declare Function fCreateShellLink Lib "VB6STKIT.DLL" (ByVal _
        lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal _
        lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long

'NOTE: If you get any GPFs, use this one and not the one above:
'Or if you are using VB5 or earlier, use this instead:
'Private Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal _
 '    lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal _
 '    lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long
 
'To update windows Icon Cache
Private Declare Sub SHChangeNotify Lib "shell32" (ByVal wEventId As Long, _
                        ByVal uFlags As Long, ByVal dwItem1 As Long, _
                        ByVal dwItem2 As Long)

' A file type association has changed.
Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0

Dim flname As String



Dim eSource As String, eDestination As String, ePassword As String
Dim dSource As String, dDestination As String, dPassword As String

Dim WithEvents EN As clsBinaryEncryptor
Attribute EN.VB_VarHelpID = -1

Private Sub cmdDecrypt_Click()


'Set fso = New FileSystemObject
'Set ts = fso.OpenTextFile(Form1_top.Text2.Text, ForAppending, True)
'ts.WriteBlankLines (nm)

'str = "File Decrypted by the user at " + time1 + " on " + date1
'ts.WriteLine (str)
'Dim str1 As String
'Set ts1 = fso.OpenTextFile(Form1_top.Text2.Text, ForReading, False)
'str1 = ts1.ReadAll

'Form1_top.Text1.Text = str1
'ts.Close
'ts1.Close
'Set fso = Nothing

'frmMain_decrypt.pb2.Value = 100

If txtDPassword.Text = "" Then
GoTo err1
End If
Screen.MousePointer = vbHourglass
Dim b As Boolean
    b = EN.DecryptFile(dSource, dDestination, IIf(txtDPassword.Text = "", "default", txtDPassword.Text))

If b = True Then
    Dim m As Integer
    m = MsgBox("The file is decrypted successfully." & vbCrLf & "Do you want to view the file now?", vbInformation + vbYesNo, "Decrypt File")
    If m = vbYes Then Browser dDestination, Me.hwnd
Else
    MsgBox "Error occured while decrypting the file. Please contact the software developer.", vbCritical, "Decrypt File"
End If
Screen.MousePointer = 0
Exit Sub
err1:
MsgBox "Please specify a password", vbCritical, "Decrypt Database"
End Sub

Private Sub cmdDOpen_Click()
With cd1
    .CancelError = True
    .Filter = "File Protector Files *.fpr|*.fpr"
    .Flags = cdlOFNFileMustExist
    
    On Error GoTo OpenError
    .DialogTitle = "Open a File to Decrypt"
    .ShowOpen
    dSource = .Filename
    
    On Error GoTo saveerror
    .DialogTitle = "Save the Decrypted File (Specify Original Extension"
    .ShowSave
    txtDFilename.Text = .Filename
    dDestination = .Filename
    
    txtDPassword.SetFocus
    txtDPassword.SelStart = 0
    txtDPassword.SelLength = Len(txtDPassword.Text)
    
    cmdDecrypt.Enabled = True
    
End With

Exit Sub
OpenError:
saveerror:

End Sub

Private Sub cmdEncrypt_Click()


    Dim strString As String
    Dim lngDword As Long
    Dim Record As String


    
        
        'Command$ is the file you need To open!
         
        'Load the file
        'Open Command$ For Input As #1
        'Do While Not EOF(1)
            'Line Input #1, Record
            'txtFile = txtFile & Record & vbCrLf
        'Loop

        ''Add your file to the Recent file folder:
        'lReturn = fCreateShellLink("..\..\Recent", _
                Command$, Command$, "")

    


    'See if our file extension already exists:
    'If GetString(HKEY_CLASSES_ROOT, ".fpr", "Content Type") = "" Then
        'Nope - not added yet. Register the file type:
        
        'create an entry in the class key
        Call SaveString(HKEY_CLASSES_ROOT, ".fpr", "", "fprfile")
        'content type
        Call SaveString(HKEY_CLASSES_ROOT, ".fpr", "Content Type", "text/plain")
        'name
        Call SaveString(HKEY_CLASSES_ROOT, "fprfile", "", "Saitech File Protector File")
        'edit flags
        Call SaveDWord(HKEY_CLASSES_ROOT, "fprfile", "EditFlags", "0000")
        'file's icon (can be an icon file, or an icon located within a dll file)
        'in this example, I am using a resource icon in this exe, 0 (app icon).
        Call SaveString(HKEY_CLASSES_ROOT, "fprfile\DefaultIcon", "", App.Path & "\" & App.EXEName & ".exe,0")
        'Shell
        Call SaveString(HKEY_CLASSES_ROOT, "fprfile\Shell", "", "")
        'Shell Open
        Call SaveString(HKEY_CLASSES_ROOT, "fprfile\Shell\Open", "", "Open with File Protector")
        'Shell open command
        Call SaveString(HKEY_CLASSES_ROOT, "fprfile\Shell\Open\Command", "", App.Path & "\" & App.EXEName & ".exe %1")
        'Update the Windows Icon Cache to see our icon right away:
        SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0

   'End If
    




'Set fso = New FileSystemObject
'Set ts = fso.OpenTextFile(Form1_top.Text2.Text, ForAppending, True)
'ts.WriteBlankLines (nm)


'str = "File Encrypted by the user at " + time1 + " on " + date1
'ts.WriteLine (str)
'Dim str1 As String
'Set ts1 = fso.OpenTextFile(Form1_top.Text2.Text, ForReading, False)
'str1 = ts1.ReadAll

'Form1_top.Text1.Text = str1
'ts.Close
'ts1.Close
'Set fso = Nothing

'Form1_file_encr.pb1.Value = 100

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
    .DialogTitle = "Open a File to Encrypt"
    .ShowOpen
    eSource = .Filename
    
    On Error GoTo saveerror
    .DialogTitle = "Save the Encrypted File"
    .ShowSave
    
    flname = ".fpr"
    txtEFilename.Text = .Filename
    eDestination = .Filename + flname 'txtEFilename.Text + flname
    
    txtEPassword.SetFocus
    txtEPassword.SelStart = 0
    txtEPassword.SelLength = Len(txtDPassword.Text)
    
    cmdEncrypt.Enabled = True
    
End With

Exit Sub
OpenError:
saveerror:

End Sub

Private Sub Command1_de_ok_Click()
Unload Me

End Sub

Private Sub Form_Load()
  Form1_file_dec.Show
  txtEFilename.SetFocus


Set EN = New clsBinaryEncryptor
'Form1_file_dec.txtEFilename.SetFocus

If Command$ <> "%1" And Command$ <> "" Then
   
   
   
   
   
   Form1_file_dec.Show
   txtDPassword.SetFocus
   txtEFilename.Enabled = False
   cmdEOpen.Enabled = False
   txtEPassword.Enabled = False
   cmdEncrypt.Enabled = False
    
    With cd1
    '.CancelError = True
    '.Filter = "All Files *.*|*.*"
    '.Flags = cdlOFNFileMustExist
    
    'On Error GoTo OpenError
    '.DialogTitle = "Open a File to decrypt."
    '.ShowOpen
    dSource = Command$ '.Filename
    
    On Error GoTo saveerror
    .DialogTitle = "Save The Decrypted File (Specify Original Extension)"
    .ShowSave
    '.Filename = Command$
    txtDFilename.Text = .Filename
    dDestination = .Filename
    
    
    txtDPassword.SelStart = 0
    txtDPassword.SelLength = Len(txtDPassword.Text)
    
   
    
End With

Exit Sub
'OpenError:
saveerror:

End If


End Sub

