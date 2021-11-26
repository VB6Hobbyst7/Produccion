VERSION 5.00
Begin VB.Form frmUtilidadesLista 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Utilidades Por Persona"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   Icon            =   "frmUtilidadesLista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   315
      Left            =   4230
      TabIndex        =   3
      Top             =   3015
      Width           =   915
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   3240
      TabIndex        =   2
      Top             =   3015
      Width           =   915
   End
   Begin VB.Frame fraUtilidades 
      Caption         =   "Lista de Utilidades a Cobrar"
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
      Height          =   2895
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   5100
      Begin SICMACT.FlexEdit grdUtilidad 
         Height          =   2175
         Left            =   45
         TabIndex        =   1
         Top             =   630
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   3836
         Cols0           =   26
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   $"frmUtilidadesLista.frx":030A
         EncabezadosAnchos=   "500-1200-1200-1500-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-R-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         SelectionMode   =   1
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label txtPersona 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   90
         TabIndex        =   4
         Top             =   270
         Width           =   4920
      End
   End
End
Attribute VB_Name = "frmUtilidadesLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************************************************************
'* NOMBRE         : "frmUtilidadesLista"
'* DESCRIPCION    : Formulario que lista las utilidades de los ex trabajadores
'* CREACION       : RIRO, 20150528 10:00 AM
'************************************************************************************************************************************************

Option Explicit

Private bAceptar As Boolean
Private rs As ADODB.Recordset
'Private nImporte As Double
'Private nIdUtilidad As Long
'Private nPeriodo As Integer
Private oUtil As PagoUtilidades

'Public Sub Inicia(ByVal pRs As ADODB.Recordset, _
'                  ByRef pnImporte As Double, _
'                  ByRef pnIdUtilidad As Long, _
'                  ByVal psNomCliente As String, _
'                  ByRef pnPeriodo As Integer)

Public Sub Inicia(ByVal pRs As ADODB.Recordset, _
                  ByVal psNomCliente As String)

bAceptar = False
limpiaType
Set rs = pRs
Me.txtPersona.Caption = Trim(psNomCliente)
Me.Show 1
'pnImporte = nImporte
'pnIdUtilidad = nIdUtilidad
'pnPeriodo = nPeriodo

End Sub
Public Function getPagoUtilidades() As PagoUtilidades
    getPagoUtilidades = oUtil
End Function
Private Function limpiaType()
oUtil.nIdUtilidad = -1
oUtil.nPeriodo = -1
oUtil.sNombre = ""
oUtil.sDOI = ""
oUtil.sCargo = ""
oUtil.sArea = ""
oUtil.dFechaIngreso = "01/01/1900"
oUtil.nImporte = -1
oUtil.C09_ParticipAdistribuir = 0
oUtil.C10_DiasLaborTodosTrabAnio = 0
oUtil.C11_RemunPercibTodosTrabAnio = 0
oUtil.C12_MontDistribXdiasLabor = 0
oUtil.C13_MontDistribXremunPercib = 0
oUtil.C14_TotDiasEfectivLabor = 0
oUtil.C15_TotRemuneraciones = 0
oUtil.C16_ParticipXdiasLabor = 0
oUtil.C17_ParticipXremuneraciones = 0
oUtil.C18_TotParticipUtilidades = 0
oUtil.C19_RetencionImpuestoRenta = 0
oUtil.C20_TotalDescuento = 0
oUtil.C21_TotalPagar = 0
End Function

Private Function verificaGridVacio() As Boolean
    If grdUtilidad.Rows = 2 And _
       Trim(grdUtilidad.TextMatrix(1, 1)) = "" And _
       Trim(grdUtilidad.TextMatrix(1, 2)) = "" And _
       Trim(grdUtilidad.TextMatrix(1, 3)) = "" And _
       Trim(grdUtilidad.TextMatrix(1, 4)) = "" Then
               
       verificaGridVacio = True
    Else
       verificaGridVacio = False
    End If
End Function
Private Sub CmdAceptar_Click()
    bAceptar = True
    Dim nRow As Integer
    Dim nCol As Integer
    If verificaGridVacio Then
'        nImporte = -1
'        nIdUtilidad = -1
'        pnPeriodo = -1
        limpiaType
    Else
        'nPeriodo = CInt(IIf(IsNumeric(grdUtilidad.TextMatrix(grdUtilidad.row, 1)), grdUtilidad.TextMatrix(grdUtilidad.row, 1), 0))
        'nImporte = CDbl(IIf(IsNumeric(grdUtilidad.TextMatrix(grdUtilidad.row, 2)), grdUtilidad.TextMatrix(grdUtilidad.row, 2), 0))
        'nIdUtilidad = CLng(IIf(IsNumeric(grdUtilidad.TextMatrix(grdUtilidad.row, 4)), grdUtilidad.TextMatrix(grdUtilidad.row, 4), 0))
        
        oUtil.sNombre = grdUtilidad.TextMatrix(grdUtilidad.row, 9)
        oUtil.dFechaIngreso = grdUtilidad.TextMatrix(grdUtilidad.row, 10)
        
        oUtil.nPeriodo = grdUtilidad.TextMatrix(grdUtilidad.row, 1)
        oUtil.nImporte = grdUtilidad.TextMatrix(grdUtilidad.row, 2)
        oUtil.nIdUtilidad = grdUtilidad.TextMatrix(grdUtilidad.row, 4)
        oUtil.sMoneda = grdUtilidad.TextMatrix(grdUtilidad.row, 3)
        
        oUtil.nIdTrama = grdUtilidad.TextMatrix(grdUtilidad.row, 5)
        oUtil.sDOI = grdUtilidad.TextMatrix(grdUtilidad.row, 6)
        oUtil.sArea = grdUtilidad.TextMatrix(grdUtilidad.row, 7)
        oUtil.sCargo = grdUtilidad.TextMatrix(grdUtilidad.row, 8)
        oUtil.sNombre = grdUtilidad.TextMatrix(grdUtilidad.row, 9)
        oUtil.dFechaIngreso = grdUtilidad.TextMatrix(grdUtilidad.row, 10)
        
        oUtil.C09_ParticipAdistribuir = grdUtilidad.TextMatrix(grdUtilidad.row, 12)
        oUtil.C10_DiasLaborTodosTrabAnio = grdUtilidad.TextMatrix(grdUtilidad.row, 13)
        oUtil.C11_RemunPercibTodosTrabAnio = grdUtilidad.TextMatrix(grdUtilidad.row, 14)
        oUtil.C12_MontDistribXdiasLabor = grdUtilidad.TextMatrix(grdUtilidad.row, 15)
        oUtil.C13_MontDistribXremunPercib = grdUtilidad.TextMatrix(grdUtilidad.row, 16)
        oUtil.C14_TotDiasEfectivLabor = grdUtilidad.TextMatrix(grdUtilidad.row, 17)
        oUtil.C15_TotRemuneraciones = grdUtilidad.TextMatrix(grdUtilidad.row, 18)
        oUtil.C16_ParticipXdiasLabor = grdUtilidad.TextMatrix(grdUtilidad.row, 19)
        oUtil.C17_ParticipXremuneraciones = grdUtilidad.TextMatrix(grdUtilidad.row, 20)
        oUtil.C18_TotParticipUtilidades = grdUtilidad.TextMatrix(grdUtilidad.row, 21)
        oUtil.C19_RetencionImpuestoRenta = grdUtilidad.TextMatrix(grdUtilidad.row, 22)
        oUtil.C20_TotalDescuento = grdUtilidad.TextMatrix(grdUtilidad.row, 23)
        oUtil.C21_TotalPagar = grdUtilidad.TextMatrix(grdUtilidad.row, 24)
        oUtil.sCiudad = grdUtilidad.TextMatrix(grdUtilidad.row, 25)
        
    End If
    Unload Me
End Sub
Private Sub cmdSalir_Click()
bAceptar = False
Unload Me
End Sub
Private Sub Form_Activate()
Dim i As Integer
i = 1
If Not rs Is Nothing Then
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            grdUtilidad.AdicionaFila
            grdUtilidad.TextMatrix(i, 1) = rs("nPeriodo")
            grdUtilidad.TextMatrix(i, 2) = Format(rs("nImporte"), "#,###0.00")
            grdUtilidad.TextMatrix(i, 3) = rs("cMoneda")
            grdUtilidad.TextMatrix(i, 4) = rs("nIdUtilidades")
            grdUtilidad.TextMatrix(i, 5) = rs("nIdTrama")
            grdUtilidad.TextMatrix(i, 6) = rs("cDoi")
            grdUtilidad.TextMatrix(i, 7) = rs("cArea")
            grdUtilidad.TextMatrix(i, 8) = rs("cCargo")
            grdUtilidad.TextMatrix(i, 9) = rs("cPersNombre")
            grdUtilidad.TextMatrix(i, 10) = rs("dFechaIngreso")
            grdUtilidad.TextMatrix(i, 11) = rs("nEstado")
            grdUtilidad.TextMatrix(i, 12) = rs("C09_ParticipAdistribuir")
            grdUtilidad.TextMatrix(i, 13) = rs("C10_DiasLaborTodosTrabAnio")
            grdUtilidad.TextMatrix(i, 14) = rs("C11_RemunPercibTodosTrabAnio")
            grdUtilidad.TextMatrix(i, 15) = rs("C12_MontDistribXdiasLabor")
            grdUtilidad.TextMatrix(i, 16) = rs("C13_MontDistribXremunPercib")
            grdUtilidad.TextMatrix(i, 17) = rs("C14_TotDiasEfectivLabor")
            grdUtilidad.TextMatrix(i, 18) = rs("C15_TotRemuneraciones")
            grdUtilidad.TextMatrix(i, 19) = rs("C16_ParticipXdiasLabor")
            grdUtilidad.TextMatrix(i, 20) = rs("C17_ParticipXremuneraciones")
            grdUtilidad.TextMatrix(i, 21) = rs("C18_TotParticipUtilidades")
            grdUtilidad.TextMatrix(i, 22) = rs("C19_RetencionImpuestoRenta")
            grdUtilidad.TextMatrix(i, 23) = rs("C20_TotalDescuento")
            grdUtilidad.TextMatrix(i, 24) = rs("C21_TotalPagar")
            grdUtilidad.TextMatrix(i, 25) = rs("cCiudad")
            rs.MoveNext
            i = i + 1
        Loop
    End If
    grdUtilidad.row = 1
    grdUtilidad.col = 1
    grdUtilidad.SetFocus
    SendKeys "{RIGHT}"
    DoEvents
    cmdAceptar.Default = True
End If
End Sub
Private Sub Form_Load()
cmdSalir.Cancel = True
bAceptar = False
'nImporte = -1
'nIdUtilidad = -1
'nImporte = -1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not bAceptar Then
'        nImporte = -1
'        nIdUtilidad = -1
'        nImporte = -1
        limpiaType
    End If
End Sub
