VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pointer Test"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   3450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test Pointers"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtVal 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblNewMemLoc 
      Caption         =   "Target Memory Location"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label lblNewVal 
      Caption         =   "Value of target after CopyMemory:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label lblMemLoc 
      Caption         =   "Source Memory Location:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   3255
   End
   Begin VB.Line lneSep 
      BorderWidth     =   3
      X1              =   120
      X2              =   3360
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblVal 
      Caption         =   "Value of X ="
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdTest_Click()
    Dim Num As Integer     'Source Variable
    Dim NewNum As Integer  'Target variable
    Dim SML As Long     'Source Memory Location
    Dim TML As Long     'Target Memory Location
    
    'Error Handling
    If Not IsNumeric(txtVal.Text) Then
        MsgBox "X not numeric!"
        Exit Sub
    End If
    If Val(txtVal.Text) > 32767 Then
        MsgBox "X must be less than 32767!"
        Exit Sub
    End If
    If Val(txtVal.Text) < -32768 Then
        MsgBox "X must be greater than -32768!"
        Exit Sub
    End If
    
    'Assign values
    Num = Val(txtVal.Text)
    NewNum = 0
    
    'Get Memory Locations
    SML = VarPtr(Num)
    TML = VarPtr(NewNum)
    
    'CopyMemory
    CopyMemory ByVal TML, ByVal SML, 2 'Or CopyMemory NewNum, Num, 2
    
    'Display Results
    lblMemLoc.Caption = "Source Memory Location: " & Str(SML)
    lblNewMemLoc.Caption = "Target Memory Location: " & Str(TML)
    lblNewVal.Caption = "Value of target after CopyMemory: " & Str(NewNum)
End Sub
