VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBinaryEncryptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event EncryptProgress(Progress As Long, ProgressMax As Long)
Public Event DecryptProgress(Progress As Long, ProgressMax As Long)

Sub SaveBinaryArray(ByVal Filename As String, WriteData() As Byte)

    Dim T As Integer
    T = FreeFile
    Open Filename For Binary Access Write As #T
        
            Put #T, , WriteData()
        
    Close #T
    
End Sub

Function ReadBinaryArray(ByVal Source As String)

    Dim bytBuf() As Byte
    Dim intN As Long
    
    Dim T As Integer
    T = FreeFile
    
    Open Source For Binary Access Read As #T
    
    Dim n As Long
    
    ReDim bytBuf(1 To LOF(T)) As Byte
    Get #T, , bytBuf()
    
    ReadBinaryArray = bytBuf()
    
    Close #T
    
End Function

Function EncryptFile(Source As String, Destination As String, password As String) As Boolean

On Error GoTo 1
    Dim ByteIn() As Byte, ByteOut() As Byte
    
    ByteIn() = ReadBinaryArray(Source)
    
    ReDim ByteOut(LBound(ByteIn) To UBound(ByteIn)) As Byte
    
    Dim i As Long, j As Long
    Dim k As Long
    Dim l As Long
    
    Dim PL As Integer
    PL = Len(password)
    
    Dim ChrBNow As Integer
    Dim PosNow As Integer
    
    Dim TempByte As Double
    
    'Decrypt
    Dim TempDByte As Double, ByteFinal As Integer
    
    For i = LBound(ByteIn) To UBound(ByteIn)
        
       PosNow = i Mod PL
        
        ChrBNow = AscB(Mid(password, PosNow + 1, 1)) Xor 17
        'ChrBNow = 1
        
        
        
        
         TempByte = ByteIn(i) Mod 256
         TempDByte = TempByte - ChrBNow
         
            If TempDByte < 0 Then
                
                 ByteFinal = 256 - Abs(TempDByte)
                    
                Else
                    
                    ByteFinal = TempDByte
                
                End If
                
        ByteOut(i) = ByteFinal
        
        
        
    Next i
    
   
  
   

SaveBinaryArray Destination, ByteOut

EncryptFile = True
Exit Function
1
EncryptFile = False

End Function

Function DecryptFile(Source As String, Destination As String, password As String) As Boolean

On Error GoTo 1
    Dim ByteIn() As Byte, ByteOut() As Byte
    
    ByteIn() = ReadBinaryArray(Source)
    
    ReDim ByteOut(LBound(ByteIn) To UBound(ByteIn)) As Byte
    
    Dim i As Long, j As Long
    
    Dim PL As Integer
    PL = Len(password)
    
    Dim ChrBNow As Integer
    Dim PosNow As Integer
    
    Dim TempByte As Integer, ByteFinal As Integer
    Dim finalbyte As Double
    For i = LBound(ByteIn) To UBound(ByteIn)
        
        PosNow = i Mod PL
        
        ChrBNow = AscB(Mid(password, PosNow + 1, 1)) Xor 17
        
        'TempByte = 256
        
       
           TempByte = ByteIn(i) + ChrBNow
           finalbyte = TempByte Mod 256
           
           'If finalbyte < 0 Then
                
                   ' ByteFinal = Abs(finalbyte)
                    
               ' Else
                    
                    'ByteFinal = finalbyte
                
                'End If
        
        ByteOut(i) = finalbyte
        
        
      
    Next i
    
    'frmMain_decrypt.pb2.Value = 100

SaveBinaryArray Destination, ByteOut

DecryptFile = True
Exit Function
1
DecryptFile = False

End Function
