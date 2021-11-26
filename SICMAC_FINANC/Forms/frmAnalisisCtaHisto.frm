VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAnalisisCtaHisto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Análisis de Cuentas: Datos Históricos"
   ClientHeight    =   4500
   ClientLeft      =   1560
   ClientTop       =   3630
   ClientWidth     =   9825
   Icon            =   "frmAnalisisCtaHisto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   9825
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid dgAna 
      Height          =   4215
      Left            =   90
      TabIndex        =   5
      Top             =   90
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
      RowHeight       =   19
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "cCtaContCod"
         Caption         =   "Cuenta Contable"
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
         DataField       =   "cMovDesc"
         Caption         =   "Descripción"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Fecha"
         Caption         =   "Fecha"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "nMovImporte"
         Caption         =   "Importe"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "cMovNro"
         Caption         =   ""
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
         MarqueeStyle    =   4
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            DividerStyle    =   6
            ColumnAllowSizing=   -1  'True
            ColumnWidth     =   1544.882
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   6
            ColumnWidth     =   3825.071
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            DividerStyle    =   6
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column04 
            DividerStyle    =   6
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
            ColumnWidth     =   0
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8550
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   8550
      TabIndex        =   1
      Top             =   1020
      Width           =   1215
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   8550
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   8550
      TabIndex        =   3
      Top             =   1860
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   555
      Left            =   8550
      Picture         =   "frmAnalisisCtaHisto.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "frmAnalisisCtaHisto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nOrdenObj As Integer
Dim rsAna     As ADODB.Recordset
Dim clsAnal As NAnalisisCtas

Private Sub ManejaBoton(plOpcion As Boolean)
cmdBuscar.Enabled = plOpcion
cmdNuevo.Enabled = plOpcion
cmdModificar.Enabled = plOpcion
cmdEliminar.Enabled = plOpcion
dgAna.Enabled = plOpcion
End Sub
  
Private Sub cmdBuscar_Click()
Dim clsBuscar As New ClassDescObjeto
ManejaBoton False
clsBuscar.BuscarDato rsAna, nOrdenObj, "Cuenta Contable", 4, 2
nOrdenObj = clsBuscar.gnOrdenBusca
Set clsBuscar = Nothing
ManejaBoton True
dgAna.SetFocus
End Sub

Private Sub cmdEliminar_Click()
On Error GoTo ErrDelete

If MsgBox(" ¿ Seguro de Eliminar Dato Histórico ? ", vbQuestion + vbYesNo, "Confirmación de Eliminación") = vbNo Then
    dgAna.SetFocus
    Exit Sub
End If
clsAnal.EliminaPendienteHisto rsAna!cMovNro
rsAna.Delete adAffectCurrent
dgAna.SetFocus
Exit Sub
ErrDelete:
  MsgBox TextErr(Err.Description), vbInformation, "¡Aviso del Eliminación!"
  dgAna.SetFocus
End Sub

Private Sub cmdModificar_Click()
Dim Pos As Variant
If rsAna.EOF Then
   MsgBox "No existen datos para modificar", vbInformation, "¡Aviso!"
   Exit Sub
End If
ManejaBoton False
Pos = rsAna.Bookmark
gnMovNro = rsAna!nMovNro
gsMovNro = rsAna!cMovNro
gsGlosa = rsAna!cMovDesc
glAceptar = False
frmAnalisisCtaHistoDet.Inicio rsAna!cCtaContCod, rsAna!nMovImporte, rsAna!nMovEstado, rsAna!nMovFlag, False
If glAceptar Then
   RefrescaDatos gsMovNro
End If
ManejaBoton True
dgAna.SetFocus
End Sub

Private Sub cmdNuevo_Click()
Dim sNewObj As String
glAceptar = False
gsMovNro = ""
gsGlosa = ""
frmAnalisisCtaHistoDet.Inicio "", 0, gMovEstContabMovContable, gMovFlagVigente, True
If glAceptar Then
   RefrescaDatos gsMovNro
End If
ManejaBoton True
dgAna.SetFocus
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub dgAna_GotFocus()
dgAna.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub dgAna_LostFocus()
dgAna.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub Form_Load()
nOrdenObj = 0
Me.Caption = GetCaptionForm(Me.Caption, gsOpeCod, Me.Width)
Set clsAnal = New NAnalisisCtas
RefrescaDatos
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set clsAnal = Nothing
RSClose rsAna
End Sub

Private Sub RefrescaDatos(Optional psMovNro As String = "")
Set rsAna = clsAnal.CargaPendientesHisto(gsOpeCod, , , adLockOptimistic)
Set dgAna.DataSource = rsAna
If psMovNro <> "" Then
   rsAna.Find "cMovNro = '" & psMovNro & "'"
End If
End Sub
