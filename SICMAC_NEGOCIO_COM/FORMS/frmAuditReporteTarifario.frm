VERSION 5.00
Begin VB.Form frmAuditReporteTarifario 
   Caption         =   "Reporte del Tarifario de Gastos"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAuditReporteTarifario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros de Búsqueda"
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      Begin VB.CommandButton Command1 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox CboProd 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Producto:"
         Height          =   255
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmAuditReporteTarifario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    EjecutarReporte "Reporte Tarifario de Gastos", CboProd.ListIndex
End Sub

Private Sub EjecutarReporte(ByVal psDescOperacion As String, ByVal psProducto As String)
Dim loTarifario As COMNAuditoria.NCOMTarifario
Dim loPrevio As previo.clsprevio
Dim lscadimp As String   'cadena q forma
Dim lsmensaje As String

If gsCodAge = "" Then
    MsgBox "Usted no es usuario de Esta Agencia...Comuniquese con RRHH", vbInformation, "AVISO"
    Exit Sub
End If

Set loTarifario = New COMNAuditoria.NCOMTarifario

    loTarifario.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis

    lscadimp = loTarifario.ReporteTarifarioGastos(CboProd.Text, psProducto, lsmensaje, gImpresora)
    
Set loTarifario = Nothing


    If Len(Trim(lscadimp)) > 0 Then

        Set loPrevio = New previo.clsprevio
            loPrevio.Show Chr$(27) & Chr$(77) & lscadimp, psDescOperacion, True, , gImpresora
        Set loPrevio = Nothing

    Else
        MsgBox "No Existen Datos para el reporte ", vbInformation, "Aviso"
    End If
End Sub

Private Sub Form_Load()
    CargarProducto
End Sub

Sub CargarProducto()
    CboProd.AddItem "Todos", 0
    CboProd.AddItem "Gastos de Creditos", 1
    CboProd.AddItem "Gastos de Judiciales", 2
    CboProd.SelText = "Todos"
End Sub


