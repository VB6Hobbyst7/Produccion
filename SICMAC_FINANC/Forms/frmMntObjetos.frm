VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMntObjetos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Objetos: Mantenimiento"
   ClientHeight    =   4500
   ClientLeft      =   1560
   ClientTop       =   3630
   ClientWidth     =   9585
   Icon            =   "frmMntObjetos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid dgObj 
      Height          =   3375
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   17
      RowDividerStyle =   4
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "cObjetoCod"
         Caption         =   "Código"
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
         DataField       =   "cObjetoDesc"
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
         DataField       =   "nObjetoNiv"
         Caption         =   "Nivel"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1785.26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5114.835
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   615.118
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8280
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   8280
      TabIndex        =   3
      Top             =   1020
      Width           =   1215
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   8280
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   8280
      TabIndex        =   5
      Top             =   1860
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   555
      Left            =   8280
      Picture         =   "frmMntObjetos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   615
      Left            =   8280
      Picture         =   "frmMntObjetos.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtObjDescrip 
      Height          =   795
      Left            =   60
      Locked          =   -1  'True
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3600
      Width           =   8100
   End
End
Attribute VB_Name = "frmMntObjetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql       As String
Dim nOrdenObj  As Integer
Dim grdActivo  As Boolean
Dim lbConsulta As Boolean
Dim rsObj      As ADODB.Recordset
Dim clsObjeto  As DObjeto
Dim WithEvents oImp As NContImprimir
Attribute oImp.VB_VarHelpID = -1
Dim oBarra     As clsProgressBar

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Public Sub Inicio(pbConsulta As Boolean)
lbConsulta = pbConsulta
Me.Show 0, frmMdiMain
End Sub

Private Sub ManejaBoton(plOpcion As Boolean)
cmdBuscar.Enabled = plOpcion
cmdNuevo.Enabled = plOpcion
cmdModificar.Enabled = plOpcion
cmdEliminar.Enabled = plOpcion
cmdImprimir.Enabled = plOpcion
dgObj.Enabled = plOpcion
End Sub
  
Private Sub cmdBuscar_Click()
Dim clsBuscar As New ClassDescObjeto
ManejaBoton False
clsBuscar.BuscarDato rsObj, nOrdenObj, "Objeto"
nOrdenObj = clsBuscar.gnOrdenBusca
Set clsBuscar = Nothing
ManejaBoton True
dgObj.SetFocus
End Sub

Private Sub cmdEliminar_Click()
On Error GoTo ErrDelete
If MsgBox(" ¿ Seguro de eliminar Objeto ? ", vbOKCancel, "Confirmación de Eliminación") = vbCancel Then
    dgObj.SetFocus
    Exit Sub
End If
If Not clsObjeto.ObjInstancia(rsObj!cObjetoCod) Then
   MsgBox "Objeto no es Ultimo Nivel. Imposible Eliminar", vbInformation, "¡Aviso!"
   Exit Sub
End If
clsObjeto.EliminaObjeto rsObj!cObjetoCod
            
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaMantObjetos
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", "Se Elimino el Objeto |Cod : " & rsObj!cObjetoCod & " |Descripción : " & rsObj!cObjetoDesc
            Set objPista = Nothing
            '*******
rsObj.Delete adAffectCurrent
dgObj.SetFocus
Exit Sub
ErrDelete:
  MsgBox TextErr(Err.Description), vbInformation, "¡Aviso del Eliminación!"
  dgObj.SetFocus
End Sub


Private Sub cmdImprimir_Click()
Dim lsImpre As String
If rsObj.EOF And rsObj.BOF Then
   MsgBox "No Existen Objetos Registrados", vbInformation, "Aviso"
   Exit Sub
End If
Set oImp = New NContImprimir

   lsImpre = oImp.ImprimeObjetos(gnLinPage)
Set oImp = Nothing
EnviaPrevio lsImpre, "Objetos: Reporte", gnLinPage, False
            
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaMantObjetos
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Imprimio los Objetos "
            Set objPista = Nothing
            '*******
dgObj.SetFocus
End Sub

Private Sub cmdModificar_Click()
Dim Pos As Variant
ManejaBoton False
frmMntObjetosNuevo.Inicia False, rsObj!cObjetoCod, rsObj!cObjetoDesc
If frmMntObjetosNuevo.OK Then
   RefrescaObj frmMntObjetosNuevo.cObjetoCod
End If
ManejaBoton True
dgObj.SetFocus
End Sub

Private Sub cmdNuevo_Click()
Dim sNewObj As String
frmMntObjetosNuevo.Inicia True, rsObj!cObjetoCod, ""
If frmMntObjetosNuevo.OK Then
   sNewObj = frmMntObjetosNuevo.cObjetoCod
   RefrescaObj sNewObj
End If
ManejaBoton True
dgObj.SetFocus
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub dgObj_GotFocus()
dgObj.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub dgObj_LostFocus()
dgObj.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub dgObj_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Not rsObj.EOF Then
   txtObjDescrip = rsObj!cObjetoDesc
End If
End Sub

Private Sub Form_Load()
Dim sCod As String
frmMdiMain.Enabled = False
nOrdenObj = 0
Set clsObjeto = New DObjeto
RefrescaObj
If lbConsulta Then
   cmdNuevo.Visible = False
   cmdEliminar.Visible = False
   cmdModificar.Visible = False
   cmdImprimir.Top = cmdNuevo.Top
End If
CentraForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set clsObjeto = Nothing
RSClose rsObj
frmMdiMain.Enabled = True
End Sub

Private Sub RefrescaObj(Optional psObjCod As String = "", Optional nPos As Integer = 1)
Set rsObj = clsObjeto.CargaObjeto(, , adLockOptimistic)
Set dgObj.DataSource = rsObj
If Not psObjCod = "" Then
   rsObj.Find "cObjetoCod = '" & psObjCod & "'"
End If
txtObjDescrip = rsObj!cObjetoDesc
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
