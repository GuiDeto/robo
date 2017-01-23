VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "0"
   ClientHeight    =   8175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   3840
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   2055
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim xlApp As Object
    Dim xlWB As Object
    Dim nome, cpf As String
    
Private Sub Command1_Click()
    
    nome = xlWB.Application.Cells(2, 2).Value 'linha e coluna
    cpf = xlWB.Application.Cells(2, 1).Value 'linha e coluna
    obj = xlWB.Application.Cells(2, 13).Value & " - " & xlWB.Application.Cells(2, 24).Value
    prop = xlWB.Application.Cells(2, 25).Value & " / " & xlWB.Application.Cells(2, 26).Value & " a vista: " & xlWB.Application.Cells(2, 25).Value & " / " & _
    xlWB.Application.Cells(2, 27).Value & " em 24x / " & xlWB.Application.Cells(2, 29).Value & "em 36x"
    valor = xlWB.Application.Cells(2, 24).Value
    emailp = xlWB.Application.Cells(2, 9).Value
 
        
    Label1.Caption = "Nome: " & nome & " CPF: " & cpf & " obj: " & obj & " proposta: " & prop & " valor: " & valor & " e-mail: " & emailp
    
    Set xlWB = Nothing
    Set xlApp = Nothing
    'Exit Sub
    
End Sub

Private Sub Command2_Click()
 Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    Set xlApp = CreateObject("Excel.Application")
    'xlApp.Visible = True
    Set xlWB = xlApp.Workbooks.Open("d:\teste1.xls")
    xlWB.Sheets("Plan1").Select
End Sub

Private Sub Timer1_Timer()

 Form1.Caption = Form1.Caption + 1

End Sub
