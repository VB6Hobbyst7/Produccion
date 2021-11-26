VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRemesaConfirmacionExt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EXTORNO DE CONFIRMACION DE REMESA"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8805
   Icon            =   "frmRemesaConfirmacionExt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBusqueda 
      Height          =   1095
      Left            =   80
      TabIndex        =   7
      Top             =   0
      Width           =   8655
      Begin VB.Frame fraOrigen 
         Caption         =   "Destino"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   5175
         Begin SICMACT.TxtBuscar txtAreaAgeCod 
            Height          =   300
            Left            =   915
            TabIndex        =   12
            Top             =   240
            Width           =   1125
            _extentx        =   1984
            _extenty        =   529
            appearance      =   1
            appearance      =   1
            font            =   "frmRemesaConfirmacionExt.frx":030A
            appearance      =   1
            stitulo         =   ""
         End
         Begin VB.Label lblAreaAgeDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2085
            TabIndex        =   14
            Top             =   240
            Width           =   2955
         End
         Begin VB.Label Label6 
            Caption         =   "Agencia :"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   270
            Width           =   735
         End
      End
      Begin VB.Frame fraFecha 
         Caption         =   "Fecha"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   735
         Left            =   5400
         TabIndex        =   9
         Top             =   240
         Width           =   1575
         Begin MSMask.MaskEdBox txtFecha 
            Height          =   330
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         Height          =   320
         Left            =   7440
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5100
      Left            =   80
      TabIndex        =   0
      Top             =   1080
      Width           =   8655
      Begin VB.Frame frmMotExtorno 
         Caption         =   "Motivos del Extorno"
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
         Height          =   2700
         Left            =   3360
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   2845
         Begin VB.CommandButton cmdExtContinuar 
            Caption         =   "&Continuar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   860
            TabIndex        =   18
            Top             =   2040
            Width           =   1095
         End
         Begin VB.TextBox txtDetExtorno 
            BackColor       =   &H00C0FFFF&
            Height          =   750
            Left            =   240
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   1200
            Width           =   2415
         End
         Begin VB.ComboBox cmbMotivos 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            ItemData        =   "frmRemesaConfirmacionExt.frx":0336
            Left            =   240
            List            =   "frmRemesaConfirmacionExt.frx":0338
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Detalles del Extorno"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblExtCmb 
            BackStyle       =   0  'Transparent
            Caption         =   "Motivo"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   320
         Left            =   7440
         TabIndex        =   4
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         Caption         =   "Todos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   780
      End
      Begin VB.TextBox txtGlosa 
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   685
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   4320
         Width           =   7215
      End
      Begin VB.CommandButton cmdExtornar 
         Caption         =   "&Extornar"
         Height          =   320
         Left            =   7440
         TabIndex        =   1
         Top             =   4330
         Width           =   1095
      End
      Begin SICMACT.FlexEdit fg 
         Height          =   3525
         Left            =   120
         TabIndex        =   5
         Top             =   525
         Width           =   8415
         _extentx        =   14843
         _extenty        =   6218
         cols0           =   10
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-Itm-Fecha-Origen-Moneda-Monto Hab.-Tipo Transporte-Empresa-nMovNro-cMovNro"
         encabezadosanchos=   "400-500-1750-2500-1000-1250-1500-1800-0-0"
         font            =   "frmRemesaConfirmacionExt.frx":033A
         font            =   "frmRemesaConfirmacionExt.frx":0366
         font            =   "frmRemesaConfirmacionExt.frx":0392
         font            =   "frmRemesaConfirmacionExt.frx":03BE
         font            =   "frmRemesaConfirmacionExt.frx":03EA
         fontfixed       =   "frmRemesaConfirmacionExt.frx":0416
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-1-X-X-X-X-X-X-X-X"
         textstylefixed  =   3
         listacontroles  =   "0-4-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-L-L-C-R-L-L-C-C"
         formatosedit    =   "0-0-0-0-0-2-0-0-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbpuntero       =   -1
         lbordenacol     =   -1
         colwidth0       =   405
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin VB.Label Label1 
         Caption         =   "Glosa :"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   4080
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmRemesaConfirmacionExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'** Nombre : frmRemesaConfirmacionExt
'** Descripción : Formulario para el extorno de confirmación de remesas
'** Creación : EJVG, 20140630 11:00:00 AM
'****************************************************************************************
Option Explicit
'CTI3
Dim Datos As Variant

Private Sub chkTodos_Click()
    Dim i As Long
    Dim lsCheck As String
    If fg.TextMatrix(1, 0) = "" Then
        chkTodos.value = 0
        Exit Sub
    End If
    lsCheck = IIf(chkTodos.value = 1, "1", "")
    For i = 1 To fg.Rows - 1
        fg.TextMatrix(i, 1) = lsCheck
    Next
End Sub


'******************************************************
'****CTI3 (ferimoro)     09102018

Sub limpDatExt()
'******CTI3 (ferimoro) 27092018
 frmMotExtorno.Visible = False
 fraOrigen.Enabled = True
 fraFecha.Enabled = True
 cmdProcesar.Enabled = True
 cmdExtornar.Enabled = True
 cmbMotivos.ListIndex = -1
 txtDetExtorno.Text = ""
 'cmbMotivos.SetFocus
End Sub

Private Sub cmdExtContinuar_Click()
    Dim oCaja As COMNCajaGeneral.NCOMCajaGeneral
    Dim lsMovNro As String
    Dim i As Integer
    Dim lbExito As Boolean
    Dim lsCadImpre As String
    
    On Error GoTo ErrCmdConfirmar
    
    '***CTI3 (FERIMORO)   02102018
    Dim DatosExtorna(1) As String

'***************CTI3  (ferimoro)  01102018
If cmbMotivos.ListIndex = -1 Or Len(txtDetExtorno.Text) <= 0 Then
    MsgBox "Debe ingresar el motivo y/o detalle del Extorno", vbInformation, "Aviso"
    Exit Sub
End If

    '***CTI3 (ferimoro)    02102018
    frmMotExtorno.Visible = False
    DatosExtorna(0) = cmbMotivos.Text
    DatosExtorna(1) = txtDetExtorno.Text

'    If Len(Trim(txtGlosa.Text)) = 0 Then
'        MsgBox "Ud. debe ingresar la glosa de extorno", vbInformation, "Aviso"
'        EnfocaControl txtGlosa
'        Exit Sub
'    End If
        
    If MsgBox("¿Esta seguro de extornar las confirmaciones de remesas seleccionadas?", vbYesNo + vbQuestion, "Aviso") = vbNo Then Limpiar: limpDatExt: Exit Sub
            
    Screen.MousePointer = 11
    cmdExtornar.Enabled = False
    Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
    lbExito = oCaja.ExtornaConfirmacionRemesa(Datos, gdFecSis, Right(gsCodAge, 2), gsCodUser, gsOpeCod, Trim(txtGlosa.Text), DatosExtorna)
    Screen.MousePointer = 0
    
    If lbExito Then
        MsgBox "Se ha extornado satisfactoriamente las confirmaciones seleccionados", vbInformation, "Aviso"
        Limpiar
        limpDatExt
    Else
        MsgBox "Ha sucedido un error al extornar los registros, si el error persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
    cmdExtornar.Enabled = True
    Set oCaja = Nothing
    Exit Sub
ErrCmdConfirmar:
    Screen.MousePointer = 0
    cmdExtornar.Enabled = True
    MsgBox err.Description, vbCritical, "Aviso"

End Sub

Private Sub cmdExtornar_Click()

Datos = DameListaMovimientos("")
    If UBound(Datos, 2) = 0 Then
        MsgBox "Ud. debe seleccionar al menos un registro para continuar", vbInformation, "Aviso"
        Exit Sub
    End If

'******CTI3 (ferimoro) 27092018
 frmMotExtorno.Visible = True
 fraOrigen.Enabled = False
 fraFecha.Enabled = False
 cmdProcesar.Enabled = False
 cmdExtornar.Enabled = False
 cmbMotivos.SetFocus
'******************************
End Sub
Private Sub cmdProcesar_Click()
    Dim oCaja As COMNCajaGeneral.NCOMCajaGeneral
    Dim rs As New ADODB.Recordset
    Dim fila As Long
    Dim lsMarca As String
    
    On Error GoTo ErrcmdProcesar
    If Not ValidaInterfaz Then Exit Sub
    
    Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
    Set rs = New ADODB.Recordset
    chkTodos.value = 0
    FormateaFlex fg
    
    Screen.MousePointer = 11
    Set rs = oCaja.ListaConfirmacionRemesaxExtorno(Right(txtAreaAgeCod, 2), CDate(txtFecha.Text))
    If Not rs.EOF Then
        lsMarca = "1"
        Do While Not rs.EOF
            fg.AdicionaFila
            fila = fg.row
            fg.TextMatrix(fila, 1) = lsMarca
            fg.TextMatrix(fila, 2) = Format(rs!dFecha, "dd/mm/yyyy hh:mm:ss AMPM")
            fg.TextMatrix(fila, 3) = rs!cOrigen
            fg.TextMatrix(fila, 4) = rs!cmoneda
            fg.TextMatrix(fila, 5) = Format(rs!nMovImporte, gsFormatoNumeroView)
            fg.TextMatrix(fila, 6) = rs!cTipoTransp
            fg.TextMatrix(fila, 7) = rs!cPersNombreTransp
            fg.TextMatrix(fila, 8) = rs!nMovNroConfirma
            fg.TextMatrix(fila, 9) = rs!cMovNroConfirma
            rs.MoveNext
        Loop
    Else
        lsMarca = "0"
        MsgBox "No se encontraron resultados", vbInformation, "Aviso"
    End If
    chkTodos.value = lsMarca
    RSClose rs
    Set oCaja = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrcmdProcesar:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdsalir_Click()
    Unload Me
End Sub
'******CTI3 (ferimoro) 18102018
Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

Set oCons = New COMDConstantes.DCOMConstantes

Set R = oCons.ObtenerConstanteExtornoMotivo

Set oCons = Nothing
Call Llenar_Combo_MotivoExtorno(R, cmbMotivos)

End Sub
Private Sub Form_Load()
    Limpiar
    Call CargaControles 'CTI3
End Sub
Private Sub Limpiar()
    txtAreaAgeCod.Text = "026" & Right(gsCodAge, 2)
    lblAreaAgeDesc.Caption = gsNomAge
    txtFecha.Text = Format(gdFecSis, gsFormatoFechaView)
    chkTodos.value = 0
    FormateaFlex fg
    txtGlosa.Text = ""
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        EnfocaControl cmdExtornar
    End If
End Sub
Private Function ValidaInterfaz() As Boolean
    Dim lsValFecha As String
    ValidaInterfaz = False
    If Len(txtAreaAgeCod.Text) <> 5 Then
        MsgBox "No se ha especificado la Agencia Destino", vbInformation, "Aviso"
        Exit Function
    End If
    lsValFecha = ValidaFecha(txtFecha.Text)
    If Len(lsValFecha) > 0 Then
        MsgBox lsValFecha, vbInformation, "Aviso"
        Exit Function
    End If
    ValidaInterfaz = True
End Function
Private Function DameListaMovimientos(ByVal psAgeCod As String) As Variant
    Dim fila As Long
    Dim lista As Variant
    Dim iLista As Integer
    
    ReDim lista(1 To 3, 0 To 0)
    If fg.TextMatrix(1, 0) <> "" Then
        For fila = 1 To fg.Rows - 1
            If fg.TextMatrix(fila, 1) = "." Then
                iLista = UBound(lista, 2) + 1
                ReDim Preserve lista(1 To 3, 0 To iLista)
                lista(1, iLista) = CInt(fg.TextMatrix(fila, 0)) 'Nro Fila flex
                lista(2, iLista) = CLng(fg.TextMatrix(fila, 8)) 'nMovNroRef
                lista(3, iLista) = fg.TextMatrix(fila, 9) 'cMovNroRef
            End If
        Next
    End If
    DameListaMovimientos = lista
End Function
