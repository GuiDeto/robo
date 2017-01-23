VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Robo Detonix"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox sheetName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2400
      TabIndex        =   4
      Text            =   "Plan1"
      Top             =   1440
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Carregar Planilha"
      Height          =   435
      Left            =   5760
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2400
      TabIndex        =   2
      Text            =   "Coopmil31082016.xls"
      Top             =   840
      Width           =   6135
   End
   Begin VB.Timer Tempo 
      Interval        =   1
      Left            =   4440
      Top             =   240
   End
   Begin VB.CommandButton parar 
      Caption         =   "Parar"
      Height          =   435
      Left            =   7200
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   2400
      TabIndex        =   0
      Text            =   "3"
      Top             =   360
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   240
   End
   Begin VB.Label Label3 
      Caption         =   "Linha atual:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Nome da Sheet:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Endereço planilha:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function Getasynckeystate Lib "user32" Alias "GetAsyncKeyState" (ByVal VKEY As Long) As Integer
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long
    
    Dim Status As Integer
    Dim xlApp As Object
    Dim xlWB As Object
    Dim nome, cpf, obj, prop, emailp, valor, tel As String
    Dim i As Integer

Private Sub Command1_Click()
    Set xlApp = CreateObject("Excel.Application")
    'xlApp.Visible = True
    'MsgBox App.Path & Text2.text
    Set xlWB = xlApp.Workbooks.Open(App.Path & "\" & Text2.text)
    xlWB.Sheets(sheetName.text).Select
End Sub

Private Sub Tempo_Timer()
Dim teclaF1, teclaF2, Shift  As Long

teclaShift = Getasynckeystate(vbKeyShift)
teclaF1 = Getasynckeystate(vbKeyF1)
teclaF2 = Getasynckeystate(vbKeyF2)

If teclaF1 And teclaShift Then
     Timer1.Enabled = True
     Status = 1
ElseIf teclaF2 And teclaShift Then
    Timer1.Enabled = False
    Unload Me
    End
End If

End Sub

Public Function colaNome(Coluna As Integer)
     nome = xlWB.Application.Cells(Coluna, 2).Value 'linha e coluna
     Sendkeys "{TAB 2}"
     'Sendkeys "Nome: "
     Sendkeys (nome)
     'Text2.text = Text2.text & "Nome: " & nome & vbCrLf
End Function

Public Function colaCpf(Coluna As Integer)
     cpf = xlWB.Application.Cells(Coluna, 1).Value 'linha e coluna
     Sendkeys "{TAB}"
     'Sendkeys "cpf: "
     Sendkeys (cpf)
     'Text2.text = Text2.text & cpf & vbCrLf
End Function

Public Function colaObj(Coluna As Integer)
    
    obj = "Trata-se de " & xlWB.Application.Cells(Coluna, 13).Value & ", referente ao Contrato/Credito Nº " & xlWB.Application.Cells(Coluna, 12).Value & ", cujo valor atualizado encontra-se em: R$ " & xlWB.Application.Cells(Coluna, 24).Value & "."
    'Sendkeys "Obj: "
     Sendkeys (obj)
    'Text2.text = Text2.text & obj & vbCrLf
End Function

Public Function colaProp(Coluna As Integer)
     prop = "Propomos as seguintes formas de pagamento: A vista: R$ " & xlWB.Application.Cells(Coluna, 25).Value & ". R$ " & xlWB.Application.Cells(Coluna, 27).Value & " parcelado em ate 12x. R$ " & xlWB.Application.Cells(Coluna, 28).Value & " parcelado em até 24x. Ou R$ " & xlWB.Application.Cells(Coluna, 29).Value & " parcelado em ate 36x."
     Sendkeys "{TAB}"
     'Sendkeys "Proposta: "
     Sendkeys (prop)
    'Text2.text = Text2.text & prop & vbCrLf
End Function

Public Function colaValor(Coluna As Integer)
    valor = xlWB.Application.Cells(Coluna, 24).Value
    Sendkeys "{TAB}"
    'Sendkeys "valor: "
     Sendkeys (valor)
     Sendkeys "{TAB 2}"
    'Text2.text = Text2.text & valor & vbCrLf
End Function

Public Function colaEmail(Coluna As Integer)
     emailp = xlWB.Application.Cells(Coluna, 9).Value
     Sendkeys "{TAB}"
     'Sendkeys "email: "
     Sendkeys (emailp)
     'Text2.text = Text2.text & emailp & vbCrLf
End Function

Public Function colaTelefone(Coluna As Integer)
    tel = ""
    
    If Len(xlWB.Application.Cells(Coluna, 3).Value) <> 0 Then
    tel = tel & " {(}" & xlWB.Application.Cells(Coluna, 3).Value & "{)} " & xlWB.Application.Cells(Coluna, 4).Value
    End If
    If Len(xlWB.Application.Cells(Coluna, 5).Value) <> 0 Then
    tel = tel & " {(}" & xlWB.Application.Cells(Coluna, 5).Value & "{)} " & xlWB.Application.Cells(Coluna, 6).Value
    End If
    If Len(xlWB.Application.Cells(Coluna, 7).Value) <> 0 Then
    tel = tel & " {(}" & xlWB.Application.Cells(Coluna, 7).Value & "{)} " & xlWB.Application.Cells(Coluna, 8).Value
    End If
     
     Sendkeys "{TAB}"
     'Sendkeys "Telefone: "
     Sendkeys (tel)
     'Text2.text = Text2.text & emailp & vbCrLf
End Function

Private Sub parar_Click()
    Timer1.Enabled = False
    End
End Sub

Private Sub Timer1_Timer()
i = Text1.text
If Status = 1 Then
 Call colaObj(i)
 Status = 2
ElseIf Status = 2 Then
 Call colaProp(i)
 Status = 3
ElseIf Status = 3 Then
 Call colaValor(i)
 Status = 4
ElseIf Status = 4 Then
 Call colaCpf(i)
 Status = 5
ElseIf Status = 5 Then
 Call colaNome(i)
 Status = 6
ElseIf Status = 6 Then
 Call colaEmail(i)
 Status = 7
ElseIf Status = 7 Then
Call colaTelefone(i)
 Status = 8
ElseIf Status = 8 Then
 Timer1.Enabled = False
 
ElseIf Status = 9 Then
 'Text1.text = Text1.text + 1
End If
 
End Sub
