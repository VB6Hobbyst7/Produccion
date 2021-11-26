VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmRepMovBGEN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de Movimientos de Boveda General"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   Icon            =   "frmRepMovBGEN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   720
      Left            =   15
      TabIndex        =   3
      Top             =   810
      Width           =   5400
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   405
         Left            =   3840
         TabIndex        =   5
         Top             =   180
         Width           =   1470
      End
      Begin VB.CommandButton CmdGenRep 
         Caption         =   "Generar Reporte"
         Height          =   405
         Left            =   105
         TabIndex        =   4
         Top             =   180
         Width           =   1470
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rango de Movimientos"
      Height          =   675
      Left            =   15
      TabIndex        =   0
      Top             =   75
      Width           =   5415
      Begin MSMask.MaskEdBox txtFechaFinal 
         Height          =   330
         Left            =   4065
         TabIndex        =   6
         Top             =   225
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaInicial 
         Height          =   300
         Left            =   1365
         TabIndex        =   7
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha de Final :"
         Height          =   225
         Left            =   2775
         TabIndex        =   2
         Top             =   278
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Inicio :"
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   278
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmRepMovBGEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ListadoDEMovimientosDEBGEN()
Dim P As Previo.clsPrevio
Dim R As ADODB.Recordset
Dim sSQL As String
Dim sCadRep As String
Dim nTotal As Double
Dim loConec As New DConecta

    sSQL = " ATM_ListadoDEMovimBGEN '" & Format(CDate(Me.txtFechaInicial.Text), "yyyy/mm/dd") & "','" & Format(CDate(Me.txtFechaFinal.Text), "yyyy/mm/dd") & "'"
    
    Set R = New ADODB.Recordset
    sCadRep = "."
    
    'Cabecera
    sCadRep = sCadRep & Space(5) & "CMAC MAYNAS S.A." & Space(50) & "FECHA : " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & Chr(10)
    sCadRep = sCadRep & Space(5) & "SIMACC-Tarjeta de Debito" & Space(42) & "Usuario : " & gsCodUser & Chr(10) & Chr(10)
    sCadRep = sCadRep & Space(40) & "LISTADO DE MOVIMIENTOS DE BOVEDA GENEARAL" & Chr(10) & Chr(10) & Chr(10)
    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
    sCadRep = sCadRep & Space(5) & "FECHA" & Space(15) & "DESCRIPCION" & Space(18) & "CANTIDAD" & Chr(10)
    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
    
    nTotal = 0
    
    'AbrirConexion
    loConec.AbreConexion
    R.Open sSQL, loConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    Do While Not R.EOF
    sCadRep = sCadRep & Space(5) & Format(R!dFecha, "dd/mm/yyyy") & Space(5) & Left(R!cDesc & Space(30), 25) & Right(Space(30) & Format(R!nCantidad, "#0.00"), 16) & Chr(10)
        nTotal = nTotal + 1
    R.MoveNext
    Loop
    R.Close
    'CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
    Set R = Nothing
    
    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
    sCadRep = sCadRep & Space(5) & "CANTIDAD : " & Space(5) & Right(Space(20) & Format(nTotal, "#0"), 5) & Chr(10)
    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
    
    
    Set P = New Previo.clsPrevio
    Call P.Show(sCadRep, "REPORTE")
    Set P = Nothing

End Sub

Private Sub CmdGenRep_Click()
    If Not IsDate(Me.txtFechaInicial.Text) Or Not IsDate(Me.txtFechaFinal.Text) Then
        MsgBox "Fecha Invalida", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Call ListadoDEMovimientosDEBGEN

End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub


Private Sub Form_Load()

    txtFechaInicial.Text = Format(Now, "dd/mm/yyyy")
    txtFechaFinal.Text = Format(Now, "dd/mm/yyyy")
    

End Sub
