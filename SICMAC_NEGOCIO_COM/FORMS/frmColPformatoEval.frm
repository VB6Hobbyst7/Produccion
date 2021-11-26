VERSION 5.00
Begin VB.Form frmColPformatoEval 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formato de Evaluación"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   Icon            =   "frmColPformatoEval.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Evaluación"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtVerfica 
         Height          =   195
         Left            =   4200
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtIngresos 
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   480
         Width           =   2000
      End
      Begin VB.TextBox txtEgresos 
         Height          =   285
         Left            =   2160
         TabIndex        =   1
         Top             =   960
         Width           =   2000
      End
      Begin VB.Label Label1 
         Caption         =   "Total de Egresos: S/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1030
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Ingresos Brutos :  S/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   520
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmColPformatoEval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'Archivo:  frmColPformatoEval.frm
'ARLO   :  05/12/2017 ---SUBIDO DESDE LA 60
'Resumen:  Registra la evaluación del cliente
'******************************************************
Option Explicit
Dim fnInteres As Double
Dim fnCapital As Double
Dim fnCapacidadPago As Double
Dim fnTpoClinte As String

Public Sub Inicio(ByVal psInteres As Double, ByVal psCapital As Double, ByVal psTpoCliente As String)
    fnInteres = psInteres
    fnCapital = psCapital
    fnTpoClinte = psTpoCliente
    Me.txtIngresos.Text = ""
    Me.txtEgresos.Text = ""
    Me.Show 1
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
    Me.txtVerfica.Text = 0
End Sub

Private Sub cmdGrabar_Click()
    
If txtIngresos = "" Or txtEgresos = "" Then
    Exit Sub
End If
    If val(Replace(txtEgresos, ",", "")) > val(Replace(txtIngresos, ",", "")) Then
        MsgBox "El Total de Egresos, no debe ser mayor a los Ingresos Brutos", vbInformation, "Aviso"
        Exit Sub
    End If
    
    fnCapacidadPago = (txtIngresos - txtEgresos)
    
    If fnCapacidadPago = 0 Then
        MsgBox "El cliente no tiene la [Capacidad de Pago] suficiente para poder atenderlo. Vuelva a ingresar los datos de la Evaluación.", vbInformation, "Aviso"
        Exit Sub
    Else
        fnCapacidadPago = ((fnInteres + 0.003 * fnCapital) / (txtIngresos - txtEgresos)) * 100
    End If
        
    If (fnTpoClinte = 1) Then
        If (fnCapacidadPago) > 75 Then
            MsgBox "El cliente no tiene la [Capacidad de Pago] suficiente para poder atenderlo. Vuelva a ingresar los datos de la Evaluación.", vbInformation, "Aviso"
            Exit Sub
        End If
    ElseIf (fnTpoClinte = 2) Then
        If (fnCapacidadPago) > 85 Then
            MsgBox "El cliente no tiene la [Capacidad de Pago] suficiente para poder atenderlo. Vuelva a ingresar los datos de la Evaluación.", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    Me.txtVerfica.Text = 1
    Me.Hide

End Sub

Private Sub txtEgresos_LostFocus()
    txtEgresos.MaxLength = "13"
    txtEgresos.Text = Format(txtEgresos.Text, "#,#00.00")
End Sub

Private Sub txtIngresos_KeyPress(KeyAscii As Integer)
txtIngresos.MaxLength = "7"
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtIngresos.MaxLength = "13"
        txtIngresos.Text = Format(txtIngresos.Text, "#,#00.00")
        Me.txtEgresos.SetFocus
    ElseIf KeyAscii <> 8 Then
        If Not IsNumeric(Chr(KeyAscii)) Then
        Beep
        KeyAscii = 0
        End If
    End If
End Sub
Private Sub txtEgresos_KeyPress(KeyAscii As Integer)
txtEgresos.MaxLength = "7"
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtEgresos.MaxLength = "13"
        txtEgresos.Text = Format(txtEgresos.Text, "#,#00.00")
        Me.cmdGrabar.SetFocus
    ElseIf KeyAscii <> 8 Then
        If Not IsNumeric(Chr(KeyAscii)) Then
        Beep
        KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtIngresos_LostFocus()
    txtIngresos.MaxLength = "13"
    txtIngresos.Text = Format(txtIngresos.Text, "#,#00.00")
End Sub
