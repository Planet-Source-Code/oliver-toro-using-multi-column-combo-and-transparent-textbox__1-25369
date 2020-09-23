VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Multi-Column Combo and Transparent TextBox"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmMain.frx":014A
   ScaleHeight     =   2280
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSForms.Label Label4 
      Height          =   240
      Left            =   270
      TabIndex        =   8
      Top             =   1530
      Width           =   1005
      VariousPropertyBits=   8388627
      Caption         =   "Adress:"
      Size            =   "1773;423"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TextBox3 
      Height          =   330
      Left            =   1395
      TabIndex        =   3
      Top             =   1530
      Width           =   6135
      VariousPropertyBits=   746604563
      Size            =   "10821;582"
      SpecialEffect   =   6
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TextBox2 
      Height          =   330
      Left            =   1395
      TabIndex        =   2
      Top             =   1170
      Width           =   6135
      VariousPropertyBits=   746604563
      Size            =   "10821;582"
      SpecialEffect   =   6
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   330
      Left            =   1395
      TabIndex        =   1
      Top             =   810
      Width           =   2535
      VariousPropertyBits=   746604563
      Size            =   "4471;582"
      SpecialEffect   =   6
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   270
      TabIndex        =   7
      Top             =   1170
      Width           =   1095
      VariousPropertyBits=   8388627
      Caption         =   "Name:"
      Size            =   "1931;423"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   240
      Left            =   270
      TabIndex        =   6
      Top             =   810
      Width           =   1005
      VariousPropertyBits=   8388627
      Caption         =   "ID:"
      Size            =   "1773;423"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox CheckBox1 
      Height          =   330
      Left            =   225
      TabIndex        =   4
      Top             =   1845
      Width           =   1365
      VariousPropertyBits=   746596371
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2408;582"
      Value           =   "0"
      Caption         =   "State:"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   240
      Left            =   270
      TabIndex        =   5
      Top             =   90
      Width           =   1005
      VariousPropertyBits=   8388627
      Caption         =   "Find:"
      Size            =   "1773;423"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox ComboBox1 
      Height          =   330
      Left            =   1395
      TabIndex        =   0
      Top             =   90
      Width           =   6225
      VariousPropertyBits=   746608667
      DisplayStyle    =   3
      Size            =   "10980;582"
      ColumnCount     =   2
      ListRows        =   3
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "6350;705"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ---------------------------------------------------------------
' Oliver Toro
' olto22@hotmail.com
' ---------------------------------------------------------------
' FM20.dll (Microsoft Forms 2.0)
' Multicolumn Combo and Transparent textBox
' ---------------------------------------------------------------
' You can change listrows property for total rows to display.
' ColumnCount is total of columns,
' ColumnHeads for headers,
' ColumnWidth you can set in pt or cm,
' And try the behavior for Enter Key.
' ---------------------------------------------------------------


Sub Load_Combo()
     ComboBox1.ColumnCount = 2
     
     ComboBox1.AddItem "Cliente 1"
     ComboBox1.List(0, 1) = 1

     ComboBox1.AddItem "Cliente 2"
     ComboBox1.List(1, 1) = 2

     ComboBox1.AddItem "Cliente 3"
     ComboBox1.List(2, 1) = 3
     
     ComboBox1.TextColumn = 1
End Sub

Private Sub ComboBox1_LostFocus()
     If ComboBox1.MatchFound = True Then
          MsgBox "Searching data - Customer#:" & ComboBox1.Column(ComboBox1.BoundColumn) & "'"
          
          TextBox1.Text = 1
          TextBox2.Text = "Cliente 1"
          TextBox3.Text = "Direccion 1"
          CheckBox1.Value = True
     Else
          MsgBox "You must select a choice...", vbExclamation, "Error"
          ComboBox1.SetFocus
     End If
End Sub


Private Sub Form_Load()
     Call Load_Combo
End Sub

