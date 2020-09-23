VERSION 5.00
Begin VB.Form frmWarning 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Important - Please Read!"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "Warning.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkWarning 
      Caption         =   "Don't tell me again"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtWarning 
      Height          =   2295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public INIFile As String

Private Sub cmdContinue_Click()
    If chkWarning.Value = 0 Then
        WriteToFile INIFile, "Settings", "ShowWarning", True
    Else
        WriteToFile INIFile, "Settings", "ShowWarning", False
    End If
    frmMain.Show
    Unload frmWarning
End Sub

Private Sub Form_Load()
    txtWarning.Text = "This Encryption method buries text into a bitmap picture file, this file must however be quite large, about 9 times the size of the text file. This may mean that you need to compress the image to transmit it, I would suggest you DON'T use JPEG compression, because it looses some of the detail of the image and on some images it could loose some of the text! Instead use a loss-less image compression format like PNG. This compression should be done in an art package like Paint Shop Pro or similar. I am NOT responsible if you compress the image and loose data!"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Show
End Sub


