VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMntRepColumnas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Reportes"
   ClientHeight    =   5850
   ClientLeft      =   1335
   ClientTop       =   1125
   ClientWidth     =   9555
   Icon            =   "frmMntRepColumnas.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   8220
      TabIndex        =   10
      Top             =   675
      Width           =   1215
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   8220
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   8220
      TabIndex        =   8
      Top             =   1890
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   8220
      TabIndex        =   7
      ToolTipText     =   "Elimina una operación"
      Top             =   1485
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   8220
      TabIndex        =   6
      ToolTipText     =   "Busca operación por su código"
      Top             =   270
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8220
      TabIndex        =   5
      ToolTipText     =   "Sale de mantenimiento de operaciones"
      Top             =   5310
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cuentas Contables"
      Height          =   2655
      Left            =   150
      TabIndex        =   4
      Top             =   3105
      Width           =   7995
      Begin VB.CommandButton cmdDesasgina 
         Caption         =   "&Quitar Cuenta"
         Height          =   375
         Left            =   1860
         TabIndex        =   3
         Top             =   2205
         Width           =   1620
      End
      Begin VB.CommandButton cmdAsignar 
         Caption         =   "&Asignar Cuenta"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   2205
         Width           =   1620
      End
      Begin MSDataGridLib.DataGrid dtgCuentas 
         Height          =   1920
         Left            =   135
         TabIndex        =   1
         Top             =   225
         Width           =   7740
         _ExtentX        =   13653
         _ExtentY        =   3387
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   2
         RowHeight       =   19
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "cCtaContCod"
            Caption         =   "Cta. Cont."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cCtaContDesc"
            Caption         =   "Descripcion de Cuenta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               DividerStyle    =   3
               ColumnWidth     =   1814.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   5355.213
            EndProperty
         EndProperty
      End
   End
   Begin MSDataGridLib.DataGrid dtgReportes 
      Height          =   2745
      Left            =   165
      TabIndex        =   0
      Top             =   255
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   4842
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   2
      RowHeight       =   19
      RowDividerStyle =   1
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "cOpeCod"
         Caption         =   "cCodRepo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "cDescCol"
         Caption         =   "Descripción"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "NnroCOL"
         Caption         =   "Nº Col"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "cTotal"
         Caption         =   "Totaliza"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "COPEDESC"
         Caption         =   "Operación"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3089.764
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   5054.74
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMntRepColumnas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sOpeCod As String
Dim sOpeDesc As String
Dim lnNroCol As Integer
Dim nPos As Long
Dim sSql As String
Dim rscta As ADODB.Recordset
Dim rsRep As ADODB.Recordset
Dim YaTrans As Boolean
Dim WithEvents oImp   As NContImprimir
Attribute oImp.VB_VarHelpID = -1
Dim oBarra  As clsProgressBar

Dim clsRepCtaCol As DRepCtaColumna

Private Sub cmdAsignar_Click()
Dim prs As ADODB.Recordset
On Error GoTo cmdAsignarErr
If rsRep.EOF Then
   MsgBox "No se encontraron datos de Reporte", vbInformation, "¡Aviso!"
   Exit Sub
End If
Dim oDesc As New ClassDescObjeto
oDesc.lbUltNivel = False
oDesc.ColCod = 0
oDesc.ColDesc = 1
Dim oCta As New DCtaCont
Set prs = oCta.CargaCtaCont()
Set oCta = Nothing
oDesc.ShowGrid prs, "Cuentas Contables"
RSClose prs
If oDesc.lbOk Then
   gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
   clsRepCtaCol.InsertaRepColumnaCta rsRep!cOpeCod, rsRep!nNroCol, oDesc.gsSelecCod, gsMovNro
   CargaCuentas
End If
Exit Sub
cmdAsignarErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub
Private Sub CargaCuentas()
If Not rsRep.EOF Then
   Set clsRepCtaCol = New DRepCtaColumna
   Set rscta = clsRepCtaCol.CargaRepColumnaCta(rsRep!cOpeCod, rsRep!nNroCol, , True)
   Set dtgCuentas.DataSource = rscta
End If
End Sub

Private Sub cmdBuscar_Click()
Dim clsBuscar As New ClassDescObjeto
clsBuscar.BuscarDato rsRep, 0, "Operación"
Set clsBuscar = Nothing
dtgReportes.SetFocus
End Sub

Private Sub cmdDesasgina_Click()
If rscta.EOF Then
   MsgBox "No existen cuentas para Eliminar...", vbInformation, "Aviso"
   Exit Sub
End If
If MsgBox("¿ Esta seguro que desea quitar la Cuenta Contable del Reporte ?", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
  clsRepCtaCol.EliminaRepColumnaCta rscta!cOpeCod, rscta!nNroCol, rscta!cCtaContCod
  CargaCuentas
End If
dtgReportes.SetFocus
End Sub

Private Sub cmdEliminar_Click()
Dim nRep As VbMsgBoxResult
On Error GoTo EliminarErr
If rscta.RecordCount > 0 Then
   nRep = MsgBox("Operación tiene Cuentas Contables Asignadas. ¿ Desea continuar Eliminación ? ", vbQuestion + vbYesNo, "¡Aviso!")
Else
   nRep = MsgBox("¿ Esta seguro de eliminar la operación ?", vbQuestion + vbYesNo, "¡Confirmación!")
End If
If nRep = vbYes Then
   clsRepCtaCol.EliminaRepColumna rsRep!cOpeCod, rsRep!nNroCol
   rsRep.Delete adAffectCurrent
End If
dtgReportes.SetFocus
Exit Sub
EliminarErr:
   dtgReportes.MarqueeStyle = dbgHighlightRow
   MsgBox TextErr(Err.Description), vbCritical, "¡Aviso!"
End Sub

Private Sub cmdImprimir_Click()
Dim sImpre As String
Set oImp = New NContImprimir
MousePointer = 11
oImp.Inicio gsNomCmac, gsCodAge, Format(gdFecSis, gsFormatoFechaView)
sImpre = oImp.ImprimeRepCtaCol(gnLinPage)
Set oImp = Nothing
MousePointer = 0
EnviaPrevio sImpre, "Reportes de Cuentas por Columnas", gnLinPage, False

End Sub



Private Sub cmdModificar_Click()
On Error GoTo ErrMod
Dim nPos
nPos = rsRep.Bookmark
glAceptar = False
frmMntRepColumnasNuevo.Inicio rsRep!cOpeCod, Trim(rsRep!cOpeDesc), Trim(Str(rsRep!nNroCol)), Trim(rsRep!cDescCol), Trim(rsRep!cTotal)
If glAceptar Then
   CargaReporte
   rsRep.Bookmark = nPos
End If
dtgReportes.SetFocus
Exit Sub
ErrMod:
 MsgBox TextErr(Err.Description), vbInformation, "Aviso"

End Sub

Private Sub cmdNuevo_Click()
glAceptar = False
frmMntRepColumnasNuevo.Inicio rsRep!cOpeCod, rsRep!cOpeDesc, clsRepCtaCol.MaxNroColumna(rsRep!cOpeCod) + 1, "", False
If glAceptar Then
   CargaReporte
   rsRep.Find "cOpeCod = '" & frmMntRepColumnasNuevo.pOpeCod & "'"
   rsRep.Find "nNroCol = " & frmMntRepColumnasNuevo.pNroCol & "", , adSearchForward, rsRep.Bookmark
End If
dtgReportes.SetFocus
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub dtgCuentas_GotFocus()
dtgCuentas.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub dtgCuentas_LostFocus()
dtgCuentas.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub dtgReportes_GotFocus()
dtgReportes.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub dtgReportes_LostFocus()
dtgReportes.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub dtgReportes_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
CargaCuentas
End Sub

Private Sub Form_Load()
frmMdiMain.Enabled = False
CentraForm Me
Set clsRepCtaCol = New DRepCtaColumna
CargaReporte
CargaCuentas
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
RSClose rscta
RSClose rsRep
Set clsRepCtaCol = Nothing
frmMdiMain.Enabled = True
End Sub

Private Sub CargaReporte()
Set rsRep = clsRepCtaCol.CargaRepColumna(, , , adLockOptimistic, True)
Set dtgReportes.DataSource = rsRep
End Sub

Private Sub oImp_BarraClose()
oBarra.CloseForm Me
End Sub

Private Sub oImp_BarraProgress(value As Variant, psTitulo As String, psSubTitulo As String, psTituloBarra As String, ColorLetras As ColorConstants)
oBarra.Progress value, psTitulo, psSubTitulo, psTituloBarra, ColorLetras
End Sub

Private Sub oImp_BarraShow(pnMax As Variant)
Set oBarra = New clsProgressBar
oBarra.ShowForm Me
oBarra.CaptionSyle = eCap_CaptionPercent
oBarra.Max = pnMax
End Sub


