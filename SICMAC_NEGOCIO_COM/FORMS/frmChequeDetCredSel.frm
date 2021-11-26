VERSION 5.00
Begin VB.Form frmChequeDetCredSel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas de la Persona"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5640
   Icon            =   "frmChequeDetCredSel.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCuentas 
      Caption         =   "Cuentas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2505
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   5430
      Begin VB.ListBox lstCuentas 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   130
         TabIndex        =   3
         Top             =   300
         Width           =   5160
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1755
      TabIndex        =   1
      Top             =   2580
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2895
      TabIndex        =   0
      Top             =   2595
      Width           =   1000
   End
End
Attribute VB_Name = "frmChequeDetCredSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************
'** Nombre : frmChequeDetCredSel
'** Descripción : Selección cuentas crédito de la Persona segun TI-ERS126-2013
'** Creación : EJVG, 20131217 12:25:00 PM
'*****************************************************************************
Option Explicit
Dim fsNroCuenta As String

Public Function Inicio(ByVal psPersCod As String, ByVal pnMoneda As Moneda) As String
    CargaCuentas psPersCod, pnMoneda
    If lstCuentas.ListCount > 0 Then
        lstCuentas.ListIndex = 0
        Show 1
        Inicio = fsNroCuenta
    Else
        MsgBox "Cliente no posee cuentas", vbInformation, "Aviso"
        Exit Function
    End If
End Function
Private Sub CargaCuentas(ByVal psPersCod As String, ByVal pnMoneda As Moneda)
    Dim obj As New NCOMDocRec
    Dim oRs As New ADODB.Recordset
    Set oRs = obj.ListaCreditosxRecepcionCheque(psPersCod, pnMoneda)
    lstCuentas.Clear
    Do While Not oRs.EOF
        lstCuentas.AddItem oRs!cCtaCod & Space(2) & oRs!cRelacionDesc & Space(2) & Format(oRs!dVigencia, gsFormatoFechaView)
        oRs.MoveNext
    Loop
    Set oRs = Nothing
    Set obj = Nothing
End Sub
Private Sub cmdAceptar_Click()
    Dim sCta As String, sProd As String, sMon As String
    Dim sCadena As String, sRel As String
    Dim nPos As Integer
    sCta = ""
    sProd = ""
    sMon = ""
    sCadena = lstCuentas.List(lstCuentas.ListIndex)
    nPos = InStr(1, sCadena, Space(2), vbTextCompare)
      
    If nPos > 0 Then
        sCta = Mid(sCadena, 1, nPos - 1)
    End If
    fsNroCuenta = sCta
    Unload Me
End Sub
Private Sub cmdCancelar_Click()
    fsNroCuenta = ""
    Unload Me
End Sub

