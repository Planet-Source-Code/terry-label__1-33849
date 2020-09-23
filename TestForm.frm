VERSION 5.00
Object = "{52D5F641-AFBC-11CF-A66F-444553540000}#2.0#0"; "FLXLABEL.ocx"
Begin VB.Form TestForm 
   BackColor       =   &H000000C0&
   Caption         =   "Label Control - Test Form"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin FLEXLABEL.Label3D Label3D1 
      Height          =   1515
      Left            =   120
      TabIndex        =   0
      Tag             =   "My First Custom Control"
      Top             =   120
      Width           =   7755
      _extentx        =   13679
      _extenty        =   3519
      font            =   "TestForm.frx":0000
      borderstyle     =   1
      forecolor       =   192
      caption         =   "Clean Sweep Imaging"
      effect          =   6
      backcolor       =   0
   End
End
Attribute VB_Name = "TestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label3D1_Click()
    MsgBox "My properties are " & vbCrLf & _
      "Caption = " & Label3D1.Caption & Chr$(13) & _
      "TextAlignment = " & Label3D1.TextAlignment & Chr$(13) & _
      "Effect = " & Label3D1.Effect
End Sub

