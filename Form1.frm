VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "Serial"
      Height          =   315
      Left            =   3300
      TabIndex        =   2
      Top             =   120
      Width           =   1155
   End
   Begin VB.ListBox lstMain 
      Height          =   2400
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4335
   End
   Begin VB.ComboBox cbDrive 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim h As clsMainInfo
 
Private Sub cmdGo_Click()
 
    Dim hT As Long
    Dim uW() As Byte
    Dim dW() As Byte
    Dim pW() As Byte
    
    Set h = New clsMainInfo
    
    With h
        .CurrentDrive = Val(cbDrive.Text)
         lstMain.Clear
         lstMain.AddItem "Current drive: " & .CurrentDrive
         lstMain.AddItem ""
         lstMain.AddItem "Model number: " & .GetModelNumber
         lstMain.AddItem "Serial number: " & .GetSerialNumber
         lstMain.AddItem "Firmware Revision: " & .GetFirmwareRevision
    End With
    
    Set h = Nothing
    
End Sub
 
Private Sub Form_Load()
    cbDrive.AddItem 0
    cbDrive.AddItem 1
    cbDrive.AddItem 2
    cbDrive.AddItem 3
    cbDrive.ListIndex = 0
End Sub
 

