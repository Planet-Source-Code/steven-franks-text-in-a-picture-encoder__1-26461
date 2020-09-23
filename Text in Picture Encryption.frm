VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text in Picture Encoder"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "Text in Picture Encryption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   257
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   393
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEncode 
      Caption         =   "Encode"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2760
      TabIndex        =   16
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdDecode 
      Caption         =   "Decode"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4320
      TabIndex        =   15
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Frame fraEncryption 
      Caption         =   "Encryption"
      Height          =   735
      Left            =   2760
      TabIndex        =   12
      Top             =   2400
      Width           =   3015
      Begin VB.VScrollBar CodeScroll 
         Height          =   285
         Left            =   2640
         Max             =   -1
         Min             =   1
         TabIndex        =   17
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtCodeNum 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Text            =   "0"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblCodeNum 
         Alignment       =   1  'Right Justify
         Caption         =   "Code Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraPreview 
      Caption         =   "Preview"
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Preview Image"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Image imgPreview 
         Height          =   1710
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame fraMethod 
      Caption         =   "Encoding Method"
      Height          =   1095
      Left            =   2760
      TabIndex        =   0
      Top             =   1200
      Width           =   3015
      Begin VB.OptionButton opnMethod 
         Caption         =   "Inverse"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2775
      End
      Begin VB.OptionButton opnMethod 
         Caption         =   "Normal"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   2775
      End
   End
   Begin MSComDlg.CommonDialog dlgFileOpen 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraFiles 
      Caption         =   "Files to use"
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txtBitmap 
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox txtTextFile 
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   720
         Width           =   3855
      End
      Begin VB.CommandButton cmdBMPBrowse 
         Caption         =   "..."
         Height          =   255
         Left            =   5160
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdTXTBrowse 
         Caption         =   "..."
         Height          =   255
         Left            =   5160
         TabIndex        =   6
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblTextFile 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Text File:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblBitmap 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bitmap File:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private EncodeType As Integer

Private Function Encrypt(ByVal Text As String, ByVal CodeNum As Integer) As String
    Dim Count As Integer
    Dim Temp As Integer
    For Count = 1 To Len(Text)
        Temp = Asc(Mid$(Text, Count, 1)) + CodeNum
        If Temp > 255 Then Temp = Temp - 255
        Encrypt = Encrypt & Chr$(Temp)
    Next Count
End Function

Private Function ConvertFromCode(ByVal Data As String) As Byte
    Dim Count As Integer
    Dim Temp As Byte
    'Decodes 4 bytes from the picture into a single byte of text
    For Count = 1 To 4
        Temp = Asc(Mid$(Data, Count, 1)) 'Get the next byte from the string of 4
        If (Temp And 2) = (2 * EncodeType) Then 'See if 2nd bit is 1
            ConvertFromCode = ConvertFromCode + 2 ^ (9 - (Count * 2)) 'if it is add to the return value
        End If
        If (Temp And 1) = EncodeType Then 'See if the 1st bit is 1
            ConvertFromCode = ConvertFromCode + 2 ^ (8 - (Count * 2)) 'if it is add to the return value
        End If
    Next Count
End Function

Private Function IntToBinStr(ByVal Num As Integer) As String
    Dim Count As Integer
    Dim Temp As Long
    Dim Temp2 As Byte
    For Count = 0 To 7
        Temp2 = (Num And 2 ^ Count)
        If Temp2 = 2 ^ Count Then
            Temp = Temp + 10 ^ Count
        End If
    Next Count
    IntToBinStr = Format$(Temp, "00000000")
End Function

Private Sub Merge(ByVal PicFile As String, ByVal TxtFile As String)
    Dim PicTemp As String * 1
    Dim TxtTemp As String
    Dim Temp As Integer
    Dim Temp2 As String
    Dim Count As Integer
    Dim Counter As Long
    Dim EncodingFactor As Integer
    'This is a very complecated subroutine, and may take some time to understand (it took me a long time to write it!)
    Counter = 56 'Set the Counter to the first location (bitmap header is 54 bytes, so I decided to start with the 56th byte)
    If EncodeType = 1 Then 'This is to do with the inverse and normal encoding methods
        EncodingFactor = 1
    Else
        EncodingFactor = -1
    End If
    Open PicFile For Binary As #1
        Open TxtFile For Input As #2
            Do While Not EOF(2)
                TxtTemp = Encrypt(Input$(1, 2), Val(txtCodeNum)) 'Input and encrypt a letter from the file
                Temp2 = IntToBinStr(Asc(TxtTemp))  'Get the binary string which represents the ascii code for that letter (after encryption)
                For Count = 0 To 3
                    Get #1, (Counter + Count), PicTemp
                    'Now the complecated part :(
                    'First bit 2
                    If Mid$(Temp2, (((Count + 1) * 2) - 1), 1) = "1" Then 'If the current bit (in the string above) is a 1
                        If (Asc(PicTemp) And 2) <> (2 * EncodeType) Then 'But the bit in this byte of the picture isn't what it needs to be (1 if normal, 0 if inverse)
                            PicTemp = Chr$(Asc(PicTemp) + (2 * EncodingFactor)) 'Set it to the correct value
                        End If
                    Else 'If the current bit is a 0 however
                        If (Asc(PicTemp) And 2) = (2 * EncodeType) Then 'And the bit in the picture byte is wrong (0 if normal, 1 if inverse)
                            PicTemp = Chr$(Asc(PicTemp) - (2 * EncodingFactor)) 'Set it to the correct value
                        End If
                    End If
                    'Then repeat for bit 1
                    If Mid$(Temp2, ((Count + 1) * 2), 1) = "1" Then
                        If (Asc(PicTemp) And 1) <> EncodeType Then
                            PicTemp = Chr$(Asc(PicTemp) + EncodingFactor)
                        End If
                    Else
                        If (Asc(PicTemp) And 1) = EncodeType Then
                            PicTemp = Chr$(Asc(PicTemp) - EncodingFactor)
                        End If
                    End If
                    Put #1, Counter + Count, PicTemp 'Now put the modified byte back into the picture file
                Next Count
                Counter = Counter + 4
            Loop
            'Now I chose to terminate the text with a chr$(0) since this is never in a plain text file
            Temp2 = IntToBinStr(Val(txtCodeNum)) 'If you encrypt chr$(0) using this algorithm (if you can call it that) you get chr$(the code number)
            'Now just repeat the above process to encode the null byte
            For Count = 0 To 3
                Get #1, Counter + Count, PicTemp
                If Mid$(Temp2, ((Count + 1) * 2) - 1, 1) = "1" Then
                    If (Asc(PicTemp) And 2) <> (2 * EncodeType) Then
                        PicTemp = Chr$(Asc(PicTemp) + (2 * EncodingFactor))
                    End If
                Else
                    If (Asc(PicTemp) And 2) = (2 * EncodeType) Then
                        PicTemp = Chr$(Asc(PicTemp) - (2 * EncodingFactor))
                    End If
                End If
                If Mid$(Temp2, ((Count + 1) * 2), 1) = "1" Then
                    If (Asc(PicTemp) And 1) <> EncodeType Then
                        PicTemp = Chr$(Asc(PicTemp) + EncodingFactor)
                    End If
                Else
                    If (Asc(PicTemp) And 1) = EncodeType Then
                        PicTemp = Chr$(Asc(PicTemp) - EncodingFactor)
                    End If
                End If
                Put #1, Counter + Count, PicTemp
            Next Count
        Close #2
    Close #1
End Sub

Private Sub Recover(ByVal PicFile As String, ByVal TxtFile As String)
    Dim Counter As Long
    Dim PicTemp As String * 4
    Dim TxtTemp As String
    Dim Temp As Integer
    Counter = 56 'Again Set the Counter at the start position
    Open PicFile For Binary As #1
        Open TxtFile For Output As #2
            Do
                Get #1, Counter, PicTemp 'Get 4 bytes from the picture file
                Temp = ConvertFromCode(PicTemp) - Val(txtCodeNum) 'Convert into ascii code for the letter and decrypt it (take away the code number, how original!)
                If Temp < 0 Then Temp = Temp + 255
                If Temp > 0 Then
                    TxtTemp = Chr$(Temp) 'If the ascii code isn't 0 then convert to the letter
                    Print #2, TxtTemp; 'Add to the text file
                End If
                Counter = Counter + 4
            Loop Until Temp = 0 'Exit when you find the terminator
        Close #2
    Close #1
End Sub

Private Sub cmdencode_Click()
    Dim PicFile As String
    Dim TxtFile As String
    Dim TimeCount As Single
    Dim Temp As Long
    Dim Temp2 As Single
    PicFile = txtBitmap.Text
    TxtFile = txtTextFile.Text
    If Len(PicFile) > 0 And Len(TxtFile) > 0 And Dir(PicFile) <> "" And Dir(TxtFile) <> "" Then
        If FileLen(PicFile) >= ((FileLen(TxtFile) + 1) * 4) + 56 Then
            frmMain.MousePointer = 11
            TimeCount = Timer
            Merge PicFile, TxtFile
            TimeCount = Timer - TimeCount
            frmMain.MousePointer = 0
            MsgBox "Text has been encoded into the image" & vbCrLf & "Encoding time: " & Format(TimeCount, "0.00") & " seconds", vbInformation, "Encoding Complete"
        Else
            Temp = ((FileLen(TxtFile) + 1) * 4) + 56
            Temp2 = Format(Temp / 1048576, "0.00")
            MsgBox "This picture file is not big enough to hold the text" & vbCrLf & "for this text file you will need a" & vbCrLf & "bitmap whose filesize is at least: " & Temp2 & "MB", vbCritical, "Error - Bitmap too small"
        End If
    End If
End Sub
Private Sub cmdBMPBrowse_Click()
    Dim Filename As String
    Dim TimeCount As Single
    dlgFileOpen.Filename = ""
    dlgFileOpen.Flags = cdlOFNHideReadOnly
    dlgFileOpen.Filter = "Windows Bitmap Files|*.bmp"
    dlgFileOpen.ShowOpen
    Filename = dlgFileOpen.Filename
    If Len(Filename) > 0 Then
        txtBitmap.Text = Filename
    End If
End Sub

Private Sub cmdTXTBrowse_Click()
    Dim Filename As String
    dlgFileOpen.Filename = ""
    dlgFileOpen.Flags = cdlOFNHideReadOnly
    dlgFileOpen.Filter = "Windows Text Files|*.txt"
    dlgFileOpen.ShowOpen
    Filename = dlgFileOpen.Filename
    If Len(Filename) > 0 Then
        txtTextFile.Text = Filename
    End If
End Sub

Private Sub cmdDecode_Click()
    Dim TxtFile As String
    Dim PicFile As String
    PicFile = txtBitmap.Text
    TxtFile = txtTextFile.Text
    If Len(TxtFile) > 0 And Len(PicFile) > 0 And Dir(PicFile) <> "" Then
        frmMain.MousePointer = 11
        TimeCount = Timer
        Recover PicFile, TxtFile
        TimeCount = Timer - TimeCount
        frmMain.MousePointer = 0
        MsgBox "Text has been decoded from the image" & vbCrLf & "Decoding time: " & Format(TimeCount, "0.00") & " seconds", vbInformation, "Decoding Complete"
    End If
End Sub

Private Sub cmdPreview_Click()
    If Dir(txtBitmap.Text) <> "" And LCase$(Right$(txtBitmap.Text, 3)) = "bmp" Then
        imgPreview.Picture = LoadPicture(txtBitmap.Text)
    End If
End Sub

Private Sub CodeScroll_Change()
    Dim Temp As Integer
    Temp = Val(txtCodeNum) + CodeScroll.Value
    If Temp > 200 Then Temp = 0
    If Temp < 0 Then Temp = 200
    txtCodeNum = Temp
    CodeScroll.Value = 0
End Sub

Private Sub Form_Load()
    Dim ShowWarning As Boolean
    frmWarning.INIFile = App.Path
    If Right$(frmWarning.INIFile, 1) <> "\" Then frmWarning.INIFile = frmWarning.INIFile & "\"
    frmWarning.INIFile = frmWarning.INIFile & "Settings.ini"
    ShowWarning = GetFromFile(frmWarning.INIFile, "Settings", "ShowWarning", True)
    If ShowWarning Then
        frmWarning.Show
        frmMain.Hide
    End If
    EncodeType = 1
End Sub

Private Sub opnMethod_Click(Index As Integer)
    EncodeType = 1 - Index
End Sub

Private Sub txtBitmap_Change()
    If Len(txtBitmap.Text) > 0 And Len(txtTextFile.Text) > 0 Then
        cmdEncode.Enabled = True
        cmdDecode.Enabled = True
        cmdPreview.Enabled = True
    Else
        cmdEncode.Enabled = False
        cmdDecode.Enabled = False
        cmdPreview.Enabled = False
    End If
End Sub

Private Sub txtCodeNum_Change()
    If Val(txtCodeNum) > 200 Then
        txtCodeNum = 200
    ElseIf Val(txtCodeNum) < 0 Then
        txtCodeNum = 0
    End If
End Sub

Private Sub txtTextFile_Change()
    If Len(txtBitmap.Text) > 0 And Len(txtTextFile.Text) > 0 Then
        cmdEncode.Enabled = True
        cmdDecode.Enabled = True
        cmdPreview.Enabled = True
    Else
        cmdEncode.Enabled = False
        cmdDecode.Enabled = False
        cmdPreview.Enabled = False
    End If
End Sub
