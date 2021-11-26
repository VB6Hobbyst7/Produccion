VERSION 5.00
Begin VB.Form frmAjusteDepreciaDet 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11190
   Icon            =   "frmAjusteDepreciaDet.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraRubroCbo 
      Height          =   705
      Left            =   75
      TabIndex        =   2
      Top             =   -60
      Width           =   8565
      Begin VB.ComboBox cboRubro 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   7440
      End
      Begin VB.Label lblRubro 
         Caption         =   "Rubros :"
         Height          =   210
         Left            =   165
         TabIndex        =   4
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8640
      TabIndex        =   1
      Top             =   5175
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9945
      TabIndex        =   0
      Top             =   5175
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   4530
      Left            =   75
      TabIndex        =   6
      Top             =   570
      Width           =   11070
      Begin VB.CheckBox chkTodas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Todos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   465
         TabIndex        =   8
         Top             =   195
         Width           =   1020
      End
      Begin Sicmact.FlexEdit Flex 
         Height          =   3975
         Left            =   75
         TabIndex        =   7
         Top             =   465
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   7011
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-OK-Codigo-Serie-Descripción-Fecha-Valor Hist-Ubicacion-Meses Dep"
         EncabezadosAnchos=   "300-350-1000-1300-2300-1000-1200-2200-800"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-X-4-5-6-7-8"
         TextStyleFixed  =   3
         ListaControles  =   "0-4-0-0-0-2-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-L-L-R-R-L-R"
         FormatosEdit    =   "0-0-0-0-0-0-2-0-3"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         Appearance      =   0
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Label lblFecha 
      Caption         =   "01/01/2001"
      Height          =   225
      Left            =   10020
      TabIndex        =   5
      Top             =   240
      Width           =   1080
   End
End
Attribute VB_Name = "frmAjusteDepreciaDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oDep As DAjusteDeprecia
Dim lsCaption As String
Dim lbIngreso As Boolean

Public Sub Ini(psCaption As String, pbIngreso As Boolean)
    lsCaption = psCaption
    lbIngreso = pbIngreso
    Me.Show 1
End Sub

Private Sub cboRubro_Click()
    If cboRubro.ListIndex <> -1 Then
        Me.Flex.rsFlex = oDep.CargaActivosDepreciaDet(Right(Me.cboRubro.Text, 4), gdFecSis)
    End If
End Sub

Private Sub chkTodas_Click()
    Dim I As Integer
    If chkTodas.value = 1 Then
        For I = 1 To Me.Flex.Rows - 1
            Me.Flex.TextMatrix(I, 1) = 1
        Next I
    Else
        For I = 1 To Me.Flex.Rows - 1
            Me.Flex.TextMatrix(I, 1) = 0
        Next I
    End If
End Sub

Private Sub cmdAceptar_Click()
    
    If Not Valida Then
        MsgBox "Debe elejir por lo menos a un activo a ingresar.", vbInformation, "Aviso"
        Me.Flex.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Desea Grabr los Cambios ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    If lbIngreso Then
        oDep.InsertarBSDeprecia Right(Me.cboRubro.Text, 5), Me.Flex.GetRsNew
        Unload Me
    Else
        oDep.MantenimientoBSDeprecia Right(Me.cboRubro.Text, 5), Me.Flex.GetRsNew
        cboRubro_Click
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Caption = lsCaption
    Set oDep = New DAjusteDeprecia
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen
    

    fraRubroCbo.Visible = True
    CargaComboRubros
    If cboRubro.ListCount > 0 Then cboRubro.ListIndex = 0
    Me.lblFecha.Caption = Format(gdFecSis, gsFormatoFechaView)
    
    If lbIngreso Then
        Me.Flex.rsFlex = oAlmacen.GetActivosFijos("2")
    Else
        Me.Flex.rsFlex = oDep.CargaActivosDepreciaDet(Right(Me.cboRubro.Text, 4), gdFecSis)
    End If
End Sub

Private Sub CargaComboRubros()
    Dim sSql As String
    Dim R As ADODB.Recordset
    Set R = New ADODB.Recordset
    cboRubro.Clear
    Set R = oDep.CargaAjusteDeprecia(, adLockOptimistic)
    RSLlenaCombo R, cboRubro
    RSClose R
End Sub

Private Function Valida() As Boolean
    Dim I As Integer
    Dim lbBan As Boolean
    lbBan = False
    If lbIngreso Then
        For I = 1 To Me.Flex.Rows - 1
            If Me.Flex.TextMatrix(I, 1) = "." Then
                I = Me.Flex.Rows - 1
                lbBan = True
            End If
        Next I
        Valida = lbBan
    Else
        Valida = True
    End If
End Function
