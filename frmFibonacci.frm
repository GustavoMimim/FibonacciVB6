VERSION 5.00
Begin VB.Form fibonacci 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fibonacci"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   5775
      Begin VB.CommandButton btnPararVerificacao 
         Caption         =   "Parar"
         Height          =   360
         Left            =   1320
         TabIndex        =   2
         Top             =   1560
         Width           =   990
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   2880
         TabIndex        =   5
         Top             =   1440
         Width           =   2775
         Begin VB.Label lblSequenciaAtual 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   0
            TabIndex        =   7
            Top             =   240
            Width           =   90
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sequencia atual: "
            Height          =   195
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   1245
         End
      End
      Begin VB.TextBox txtNumSequencia 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Text            =   "354224848179261915075"
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton btnVerificar 
         Caption         =   "Verificar"
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   1560
         Width           =   990
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Verifique se o número está na sequência de Fibonacci digitando ele a baixo!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5640
      End
   End
End
Attribute VB_Name = "fibonacci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pararExecucao As Boolean

Private Sub btnPararVerificacao_Click()

    pararExecucao = True

End Sub

Private Sub btnVerificar_Click()

On Error GoTo tratarErro

    Dim inputNumero          As Variant
    Dim sequenciaFibonacci() As Variant
    Dim sequenciaAtual       As Long

    inputNumero = CDec(txtNumSequencia.Text)

    If Not IsNumeric(inputNumero) Then

        MsgBox "Por favor, verifique o valor que foi digitado" & vbNewLine & "Só é permitido números inteiros!", vbCritical
        Exit Sub

    Else
    
        If inputNumero < 0 Then
    
            MsgBox "O número precisa ser positivo!", vbCritical
            Exit Sub
            
        End If
        
    End If

    ReDim sequenciaFibonacci(1)
    sequenciaFibonacci(0) = 0
    sequenciaFibonacci(1) = 1

    sequenciaAtual = 2

    Do While True

        If pararExecucao Then

            pararExecucao = False
            Exit Sub

        End If

        DoEvents

        lblSequenciaAtual.Caption = sequenciaAtual

        ReDim Preserve sequenciaFibonacci(sequenciaAtual)

        sequenciaFibonacci(sequenciaAtual) = CDec(sequenciaFibonacci(sequenciaAtual - 2)) + CDec(sequenciaFibonacci(sequenciaAtual - 1))

        If sequenciaFibonacci(sequenciaAtual) = inputNumero Then

            MsgBox "O valor digitado é " & sequenciaAtual & "° número da sequência!"
            Exit Sub

        ElseIf sequenciaFibonacci(sequenciaAtual) > inputNumero Then

            MsgBox "O número digitado NÃO está na sequência de Fibonacci!"
            Exit Sub

        End If
        
        sequenciaAtual = sequenciaAtual + 1

    Loop

Exit Sub

tratarErro:

    If MsgBox("Oops! Um erro inesperado aconteceu." & vbNewLine & "Ficarei muito feliz se me reportar esse problema na parte de issues no Github." & vbNewLine & "Deseja copiar o link do repositório na sua área de transferência?") = vbYes Then

        'Clipboard.SetText Text1.Text ' Poe o texto no ClipBoard
    
    End If

End Sub
