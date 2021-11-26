VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRemesaAgenciaToIFiAgenciaExt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EXTORNO DE REMESA A AGENCIA - DEVOLUCIÓN A BANCO"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8805
   Icon            =   "frmRemesaAgenciaToIFiAgenciaExt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5100
      Left            =   80
      TabIndex        =   8
      Top             =   1080
      Width           =   8655
      Begin VB.CommandButton cmdExtornar 
         Caption         =   "&Extornar"
         Height          =   320
         Left            =   7440
         TabIndex        =   12
         Top             =   4330
         Width           =   1095
      End
      Begin VB.TextBox txtGlosa 
         Height          =   685
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   4320
         Width           =   7215
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         Caption         =   "Todos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   780
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   320
         Left            =   7440
         TabIndex        =   9
         Top             =   4680
         Width           =   1095
      End
      Begin SICMACT.FlexEdit fg 
         Height          =   3525
         Left            =   120
         TabIndex        =   13
         Top             =   525
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   6218
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Itm-Fecha-Destino-Moneda-Monto Hab.-Tipo Transporte-Empresa-nMovNro-cMovNro"
         EncabezadosAnchos=   "400-500-1750-2500-1000-1250-1500-1800-0-0"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-4-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-L-C-R-L-L-C-C"
         FormatosEdit    =   "0-0-0-0-0-2-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         lbOrdenaCol     =   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label1 
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   4080
         Width           =   735
      End
   End
   Begin VB.Frame fraBusqueda 
      Height          =   1095
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         Height          =   320
         Left            =   7440
         TabIndex        =   7
         Top             =   480
         Width           =   1095
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
         TabIndex        =   5
         Top             =   240
         Width           =   1575
         Begin MSMask.MaskEdBox txtFecha 
            Height          =   330
            Left            =   240
            TabIndex        =   6
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
      Begin VB.Frame fraOrigen 
         Caption         =   "Origen"
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
         TabIndex        =   1
         Top             =   240
         Width           =   5175
         Begin SICMACT.TxtBuscar txtAreaAgeCod 
            Height          =   300
            Left            =   915
            TabIndex        =   2
            Top             =   240
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            sTitulo         =   ""
         End
         Begin VB.Label Label6 
            Caption         =   "Agencia :"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   270
            Width           =   735
         End
         Begin VB.Label lblAreaAgeDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2085
            TabIndex        =   3
            Top             =   240
            Width           =   2955
         End
      End
   End
End
Attribute VB_Name = "frmRemesaAgenciaToIFiAgenciaExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************************
'** Nombre : frmRemesaAgenciaToIFiAgenciaExt
'** Descripción : Formulario para el extorno de remesas a otras agencia o devolución a Banco
'** Creación : EJVG, 20140812 05:00:00 PM
'*******************************************************************************************
Option Explicit

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
Private Sub cmdExtornar_Click()
    Dim oCaja As COMNCajaGeneral.NCOMCajaGeneral
    Dim oMov As COMDMov.DCOMMov
    Dim Datos As Variant
    Dim lsMovNro As String
    Dim i As Integer
    Dim lbExito As Boolean
    Dim lsCadImpre As String
    
    On Error GoTo ErrCmdConfirmar
    
    Datos = DameListaMovimientos("")
    If UBound(Datos, 2) = 0 Then
        MsgBox "Ud. debe seleccionar al menos un registro para continuar", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(Trim(txtGlosa.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la glosa de extorno", vbInformation, "Aviso"
        EnfocaControl txtGlosa
        Exit Sub
    End If
    'Valida que no se hayan realizado confirmaciones de las remesas
    Set oMov = New COMDMov.DCOMMov
    For i = 1 To UBound(Datos, 2)
        If oMov.ExisteMovimientosDespues(Datos(2, i), True) Then
            MsgBox "El registro [" & Datos(1, i) & "] ya tienen movimientos luego de la remesa, no se puede continuar", vbExclamation, "Aviso"
            fg.row = Datos(1, i)
            fg.TopRow = Datos(1, i)
            fg.col = 2
            Set oMov = Nothing
            Exit Sub
        End If
    Next
    Set oMov = Nothing
        
    If MsgBox("¿Esta seguro de extornar las remesas seleccionadas?", vbYesNo + vbQuestion, "Aviso") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    cmdExtornar.Enabled = False
    Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
    lbExito = oCaja.ExtornarRemesa(Datos, gdFecSis, Right(gsCodAge, 2), gsCodUser, gsOpeCod, Trim(txtGlosa.Text))
    Screen.MousePointer = 0
    
    If lbExito Then
        MsgBox "Se ha extornado las remesas seleccionados", vbInformation, "Aviso"
        Limpiar
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
    Set rs = oCaja.ListaRemesaxExtorno(Right(txtAreaAgeCod, 2), CDate(txtFecha.Text))
    If Not rs.EOF Then
        lsMarca = "1"
        Do While Not rs.EOF
            fg.AdicionaFila
            fila = fg.row
            fg.TextMatrix(fila, 1) = lsMarca
            fg.TextMatrix(fila, 2) = Format(rs!dFecha, "dd/mm/yyyy hh:mm:ss AMPM")
            fg.TextMatrix(fila, 3) = rs!cDestino
            fg.TextMatrix(fila, 4) = rs!cMoneda
            fg.TextMatrix(fila, 5) = Format(rs!nMovImporte, gsFormatoNumeroView)
            fg.TextMatrix(fila, 6) = rs!cTipoTransp
            fg.TextMatrix(fila, 7) = rs!cPersNombreTransp
            fg.TextMatrix(fila, 8) = rs!nMovNro
            fg.TextMatrix(fila, 9) = rs!cMovNro
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
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Limpiar
End Sub
Private Sub Limpiar()
    txtAreaAgeCod.Text = "026" & Right(gsCodAge, 2)
    lblAreaAgeDesc.Caption = gsNomAge
    txtFecha.Text = Format(gdFecSis, gsFormatoFechaView)
    chkTodos.value = 0
    FormateaFlex fg
    txtGlosa.Text = ""
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
    Dim Lista As Variant
    Dim iLista As Integer
    
    ReDim Lista(1 To 3, 0 To 0)
    If fg.TextMatrix(1, 0) <> "" Then
        For fila = 1 To fg.Rows - 1
            If fg.TextMatrix(fila, 1) = "." Then
                iLista = UBound(Lista, 2) + 1
                ReDim Preserve Lista(1 To 3, 0 To iLista)
                Lista(1, iLista) = CInt(fg.TextMatrix(fila, 0)) 'Nro Fila flex
                Lista(2, iLista) = CLng(fg.TextMatrix(fila, 8)) 'nMovNroRef
                Lista(3, iLista) = fg.TextMatrix(fila, 9) 'cMovNroRef
            End If
        Next
    End If
    DameListaMovimientos = Lista
End Function
