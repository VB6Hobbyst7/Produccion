VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPigTarifario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Tarifario"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   Icon            =   "frmPigTarifario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   5745
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   8280
      Begin SICMACT.FlexEdit feTarifario 
         Height          =   3975
         Left            =   75
         TabIndex        =   16
         Top             =   180
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   7011
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Proceso-Material-Moneda-FechaIni-Valor"
         EncabezadosAnchos=   "500-2000-1900-1100-1250-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Frame Frame2 
         Height          =   900
         Left            =   75
         TabIndex        =   8
         Top             =   4110
         Width           =   8115
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            ItemData        =   "frmPigTarifario.frx":030A
            Left            =   3705
            List            =   "frmPigTarifario.frx":0314
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   495
            Width           =   1740
         End
         Begin MSDataListLib.DataCombo cboProceso 
            Height          =   315
            Left            =   90
            TabIndex        =   17
            Top             =   480
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin MSMask.MaskEdBox mskFecha 
            Height          =   300
            Left            =   5535
            TabIndex        =   9
            Top             =   495
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin SICMACT.EditMoney txtValor 
            Height          =   300
            Left            =   6750
            TabIndex        =   10
            Top             =   480
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin MSDataListLib.DataCombo cboMaterial 
            Height          =   315
            Left            =   1815
            TabIndex        =   18
            Top             =   480
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label Label5 
            Caption         =   "Valor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6765
            TabIndex        =   15
            Top             =   270
            Width           =   600
         End
         Begin VB.Label Label4 
            Caption         =   "Fec. Inicial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5520
            TabIndex        =   14
            Top             =   270
            Width           =   1110
         End
         Begin VB.Label Label3 
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3780
            TabIndex        =   13
            Top             =   270
            Width           =   720
         End
         Begin VB.Label Label2 
            Caption         =   "Material"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1875
            TabIndex        =   12
            Top             =   270
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo de Proceso"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   75
            TabIndex        =   11
            Top             =   255
            Width           =   1455
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   585
         Left            =   75
         ScaleHeight     =   525
         ScaleWidth      =   8070
         TabIndex        =   1
         Top             =   5070
         Width           =   8130
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            CausesValidation=   0   'False
            Height          =   360
            Left            =   6930
            TabIndex        =   7
            Top             =   105
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.CommandButton cmdSalir 
            Cancel          =   -1  'True
            Caption         =   "&Salir"
            CausesValidation=   0   'False
            Height          =   360
            Left            =   5640
            TabIndex        =   6
            Top             =   105
            Width           =   1100
         End
         Begin VB.CommandButton cmdGrabar 
            Caption         =   "&Grabar"
            Height          =   360
            Left            =   5610
            TabIndex        =   5
            Top             =   105
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.CommandButton cmdModificar 
            Caption         =   "&Modificar"
            Height          =   360
            Left            =   1215
            TabIndex        =   4
            Top             =   105
            Width           =   1100
         End
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   360
            Left            =   2310
            TabIndex        =   3
            Top             =   105
            Width           =   1100
         End
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "&Nuevo"
            Height          =   360
            Left            =   120
            TabIndex        =   2
            Top             =   105
            Width           =   1100
         End
      End
   End
End
Attribute VB_Name = "frmPigTarifario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cmdBoton As Integer

Private Sub CargaCombo(Combo As DataCombo, pnConstante As Integer)
Dim oPigFuncion As DPigFunciones
Dim rs As Recordset

    Set oPigFuncion = New DPigFunciones
    Set rs = oPigFuncion.GetConstante(pnConstante)
    Set oPigFuncion = Nothing

    Set Combo.RowSource = rs
    Combo.ListField = "cConsDescripcion"
    Combo.BoundColumn = "nConsValor"

    Set rs = Nothing

End Sub

Private Sub cboMaterial_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cboMoneda.SetFocus
End If
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    mskFecha.SetFocus
End If
End Sub

Private Sub cboProceso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cboMaterial.SetFocus
End If
End Sub

Private Sub cmdCancelar_Click()
    Call HabilitaControles(False)
End Sub

Private Sub cmdeliminar_Click()
Dim oGrabaTarifario As DPigActualizaBD
Dim oPigFunc As DPigFunciones
Dim lnProceso As Integer
Dim lnMaterial As Integer
Dim rs As Recordset


    If MsgBox("Desea Eliminar el Tarifario?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        If cboProceso.BoundText <> "" Then
            lnProceso = cboProceso.BoundText
        Else
            MsgBox "Seleccione un elemento a eliminar", vbInformation, "Aviso"
            Exit Sub
        End If
        If cboMaterial.BoundText <> "" Then
            lnMaterial = cboMaterial.BoundText
        Else
            MsgBox "Seleccione un elemento a eliminar", vbInformation, "Aviso"
            Exit Sub
        End If

        Set oGrabaTarifario = New DPigActualizaBD
        Call oGrabaTarifario.dDeleteColocPigPrecioMaterial(lnProceso, lnMaterial, cboMoneda.ListIndex + 1, mskFecha, False)
        Set oGrabaTarifario = Nothing
        
        Set oPigFunc = New DPigFunciones
        Set rs = oPigFunc.GetTarifario
        Set oPigFunc = Nothing
        Set feTarifario.Recordset = rs
        
        Set oPigFunc = Nothing
        
        Call HabilitaControles(False)
        
    End If
    
End Sub

Private Sub cmdGrabar_Click()
Dim oGrabaTarifario As DPigActualizaBD
Dim oPigFunc As DPigFunciones
Dim lnMaterial As Integer
Dim lnProceso As Integer
Dim rs As Recordset

    If Not ValidaGrabar Then
        Exit Sub
    End If
    lnMaterial = cboMaterial.BoundText
    lnProceso = cboProceso.BoundText
    
    Set oGrabaTarifario = New DPigActualizaBD
    If cmdBoton = 1 Then
        Call oGrabaTarifario.dInsertColocPigPrecioMaterial(lnMaterial, cboMoneda.ListIndex + 1, lnProceso, mskFecha, txtValor.Text, False)
    Else
        Call oGrabaTarifario.dUpdateColocPigPrecioMaterial(lnProceso, lnMaterial, cboMoneda.ListIndex + 1, mskFecha, _
                    txtValor.Text, IIf(feTarifario.TextMatrix(feTarifario.Row, 3) = "SOLES", 1, 2))
    End If
    Set oGrabaTarifario = Nothing
    
    Set oPigFunc = New DPigFunciones
    Set rs = oPigFunc.GetTarifario
    Set oPigFunc = Nothing
    Set feTarifario.Recordset = rs
    
    Set oPigFunc = Nothing
    
    Call HabilitaControles(False)

End Sub

Private Sub CmdModificar_Click()

    cmdBoton = 2
    Call HabilitaControles(True)
    Limpiar
    CargaDatos
    cboMaterial.Enabled = False
    cboProceso.Enabled = False
    mskFecha.Enabled = False
    cboMoneda.SetFocus
    
End Sub

Private Sub CargaDatos()

    cboProceso.Text = feTarifario.TextMatrix(feTarifario.Row, 1)
    cboMaterial.Text = feTarifario.TextMatrix(feTarifario.Row, 2)
    cboMoneda.ListIndex = IIf(feTarifario.TextMatrix(feTarifario.Row, 3) = "SOLES", 0, 1)
    mskFecha.Text = feTarifario.TextMatrix(feTarifario.Row, 4)
    txtValor = feTarifario.TextMatrix(feTarifario.Row, 5)
    
End Sub

Private Sub cmdNuevo_Click()

    cmdBoton = 1
    Call HabilitaControles(True)
    Limpiar
    cboMaterial.Enabled = True
    cboProceso.Enabled = True
    mskFecha.Enabled = True
    cboProceso.SetFocus
    
End Sub

Private Sub cmdsalir_Click()
    
    Unload Me
    
End Sub

Private Sub feTarifario_Click()
    
    cboProceso.Text = feTarifario.TextMatrix(feTarifario.Row, 1)
    cboMaterial.Text = feTarifario.TextMatrix(feTarifario.Row, 2)
    cboMoneda.ListIndex = IIf(feTarifario.TextMatrix(feTarifario.Row, 3) = "SOLES", 0, 1)
    mskFecha = feTarifario.TextMatrix(feTarifario.Row, 4)
    txtValor.Text = feTarifario.TextMatrix(feTarifario.Row, 5)
    
End Sub

Private Sub Form_Load()
Dim oPigFunc As DPigFunciones
Dim rs As Recordset
Dim i As Integer

Set oPigFunc = New DPigFunciones
Set rs = oPigFunc.GetTarifario
Set oPigFunc = Nothing
Set feTarifario.Recordset = rs

    Set rs = Nothing
    
    Call CargaCombo(cboProceso, gColocPigTipoProcesoTar)
    Call CargaCombo(cboMaterial, gColocPigMaterial)
    
    feTarifario.Height = 4845
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub mskFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsDate(mskFecha.Text) Then
        txtValor.SetFocus
    Else
        MsgBox "Fecha no válida", vbInformation, "Aviso"
        mskFecha.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub HabilitaControles(ByVal pbHabilita As Boolean)
    cboProceso.Enabled = pbHabilita
    cboMaterial.Enabled = pbHabilita
    cboMoneda.Enabled = pbHabilita
    mskFecha.Enabled = pbHabilita
    txtValor = pbHabilita
    feTarifario.Height = IIf(pbHabilita, 3975, 4845)
    CmdNuevo.Visible = Not pbHabilita
    cmdModificar.Visible = Not pbHabilita
    CmdEliminar.Visible = Not pbHabilita
    cmdSalir.Visible = Not pbHabilita
    feTarifario.Enabled = Not pbHabilita
    cmdGrabar.Visible = pbHabilita
    CmdCancelar.Visible = pbHabilita
End Sub

Private Sub Limpiar()

cboProceso.BoundText = -1
cboMaterial.BoundText = -1
cboMoneda.ListIndex = -1
mskFecha.Text = "__/__/____"
txtValor = 0

End Sub

Private Function ValidaGrabar() As Boolean

    ValidaGrabar = True
    
    If cboProceso.Text = "" Then
        MsgBox "Seleccione el Tipo de Proceso para el Tarifario", vbInformation, "Aviso"
        ValidaGrabar = False
        Exit Function
    End If
        
    If cboMaterial.Text = "" Then
        MsgBox "Seleccione el Tipo de Material para el Tarifario", vbInformation, "Aviso"
        ValidaGrabar = False
        Exit Function
    End If
        
    If cboMoneda.Text = "" Then
        MsgBox "Seleccione el Tipo de Moneda", vbInformation, "Aviso"
        ValidaGrabar = False
        Exit Function
    End If
    
    If mskFecha = "" Then
        MsgBox "Fecha no válida", vbInformation, "Aviso"
        ValidaGrabar = False
        Exit Function
    ElseIf Not IsDate(mskFecha) Then
        MsgBox "Fecha no válida", vbInformation, "Aviso"
        ValidaGrabar = False
        Exit Function
    End If
    
End Function

Private Sub TxtValor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdGrabar.SetFocus
End If
End Sub
