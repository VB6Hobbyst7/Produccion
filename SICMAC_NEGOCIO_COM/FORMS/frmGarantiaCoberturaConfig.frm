VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGarantiaCoberturaConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Coberturas x Producto"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   Icon            =   "frmGarantiaCoberturaConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6900
      TabIndex        =   15
      ToolTipText     =   "Salir"
      Top             =   6970
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   75
      TabIndex        =   16
      Top             =   75
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   12091
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Créditos Directos Analista"
      TabPicture(0)   =   "frmGarantiaCoberturaConfig.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmbClienteTpo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdGuardar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdCancelar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdEditar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmbProducto"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraCobertura"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdMostrar"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Excepciones"
      TabPicture(1)   =   "frmGarantiaCoberturaConfig.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraCoberturaRapiFlashCampaña"
      Tab(1).Control(1)=   "cmdEditarEx"
      Tab(1).Control(2)=   "cmdCancelarEx"
      Tab(1).Control(3)=   "cmdGuardarEx"
      Tab(1).Control(4)=   "fraCoberturaRapiFlash"
      Tab(1).ControlCount=   5
      Begin VB.Frame fraCoberturaRapiFlashCampaña 
         Caption         =   "Coberturas RapiFlash con Campaña SuperFlash"
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
         Height          =   2055
         Left            =   -74880
         TabIndex        =   30
         Top             =   2640
         Width           =   7545
         Begin VB.TextBox txtRetiroIntAdelantadoCamp 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   9
            Top             =   840
            Width           =   720
         End
         Begin VB.TextBox txtRetiroIntMensualCamp 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   10
            Top             =   1200
            Width           =   720
         End
         Begin VB.TextBox txtRetiroIntFinalCamp 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   11
            Top             =   1560
            Width           =   720
         End
         Begin VB.Label Label16 
            Caption         =   "Cobertura de Plazo Fijo por forma de Retiro:"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label Label15 
            Caption         =   "Retiro de Interés Adelantado :"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label Label14 
            Caption         =   "%"
            Height          =   255
            Left            =   4080
            TabIndex        =   35
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label13 
            Caption         =   "Retiro de Interés Mensual :"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   1230
            Width           =   2295
         End
         Begin VB.Label Label12 
            Caption         =   "%"
            Height          =   255
            Left            =   4080
            TabIndex        =   33
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label Label11 
            Caption         =   "Retiro de Interés a Final de Plazo :"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   1590
            Width           =   2535
         End
         Begin VB.Label Label10 
            Caption         =   "%"
            Height          =   255
            Left            =   4080
            TabIndex        =   31
            Top             =   1560
            Width           =   255
         End
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "&Mostrar"
         Height          =   350
         Left            =   6765
         TabIndex        =   2
         Top             =   700
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Frame fraCobertura 
         Caption         =   "Coberturas por Garantías"
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
         Height          =   5175
         Left            =   120
         TabIndex        =   19
         Top             =   1150
         Width           =   7575
         Begin SICMACT.FlexEdit feCoberturas 
            Height          =   4575
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   7320
            _ExtentX        =   12912
            _ExtentY        =   8070
            Cols0           =   8
            HighLight       =   1
            AllowUserResizing=   1
            EncabezadosNombres=   "#-BienGarantiaID--Clasificación-Bien en Garantía-Cob. Pref.-Cob. No Pref.-Aux"
            EncabezadosAnchos=   "0-0-400-1800-2300-1200-1200-0"
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
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-2-X-X-5-6-X"
            ListaControles  =   "0-0-4-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-C-L-L-R-R-C"
            FormatosEdit    =   "0-0-0-0-0-2-2-0"
            CantDecimales   =   4
            TextArray0      =   "#"
            lbFlexDuplicados=   0   'False
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
         Begin VB.Label LblVariable 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "* La cobertura, es las veces del valor que requiere una garantía para cubrir un crédito."
            DragMode        =   1  'Automatic
            ForeColor       =   &H00000080&
            Height          =   270
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   4845
            Width           =   7300
         End
      End
      Begin VB.ComboBox cmbProducto 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   720
         Width           =   4380
      End
      Begin VB.CommandButton cmdEditarEx 
         Caption         =   "&Editar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73845
         TabIndex        =   13
         ToolTipText     =   "Editar"
         Top             =   6375
         Width           =   1000
      End
      Begin VB.CommandButton cmdCancelarEx 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -72810
         TabIndex        =   14
         ToolTipText     =   "Cancelar"
         Top             =   6375
         Width           =   1000
      End
      Begin VB.CommandButton cmdGuardarEx 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         TabIndex        =   12
         ToolTipText     =   "Guardar"
         Top             =   6375
         Width           =   1000
      End
      Begin VB.Frame fraCoberturaRapiFlash 
         Caption         =   "Coberturas RapiFlash"
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
         Height          =   2055
         Left            =   -74880
         TabIndex        =   21
         Top             =   480
         Width           =   7545
         Begin VB.TextBox txtRetiroIntFinal 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   8
            Top             =   1560
            Width           =   720
         End
         Begin VB.TextBox txtRetiroIntMensual 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   7
            Top             =   1200
            Width           =   720
         End
         Begin VB.TextBox txtRetiroIntAdelantado 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   6
            Top             =   840
            Width           =   720
         End
         Begin VB.Label Label9 
            Caption         =   "%"
            Height          =   255
            Left            =   4080
            TabIndex        =   28
            Top             =   1560
            Width           =   255
         End
         Begin VB.Label Label8 
            Caption         =   "Retiro de Interés a Final de Plazo :"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   1590
            Width           =   2535
         End
         Begin VB.Label Label7 
            Caption         =   "%"
            Height          =   255
            Left            =   4080
            TabIndex        =   26
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label Label6 
            Caption         =   "Retiro de Interés Mensual :"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   1230
            Width           =   2295
         End
         Begin VB.Label Label5 
            Caption         =   "%"
            Height          =   255
            Left            =   4080
            TabIndex        =   24
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   "Retiro de Interés Adelantado :"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label Label3 
            Caption         =   "Cobertura de Plazo Fijo por forma de Retiro:"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1150
         TabIndex        =   4
         ToolTipText     =   "Editar"
         Top             =   6375
         Width           =   1000
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2190
         TabIndex        =   5
         ToolTipText     =   "Cancelar"
         Top             =   6375
         Width           =   1000
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Guardar"
         Top             =   6375
         Width           =   1000
      End
      Begin VB.ComboBox cmbClienteTpo 
         Height          =   315
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   1860
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Cliente :"
         Height          =   255
         Left            =   4680
         TabIndex        =   18
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Producto :"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmGarantiaCoberturaConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************************************************
'** Nombre : frmGarantiaCoberturaConfig
'** Descripción : Para edición de la configuración de Coberturas de la Garantía por Producto creado segun TI-ERS063-2014
'** Creación : EJVG, 20150221 10:00:00 AM
'***********************************************************************************************************************
Option Explicit
Dim fsTpoProdCod As String
Dim fbCliPreferencial As Boolean

Dim fRsCoberturas As ADODB.Recordset
Dim fnRetiroIntAdelantado As Currency
Dim fnRetiroIntMensual As Currency
Dim fnRetiroIntFinal As Currency

Dim fnRetiroIntAdelantadoCampania As Currency
Dim fnRetiroIntMensualCampania As Currency
Dim fnRetiroIntFinalCampania As Currency

Dim fbEditar As Boolean
Dim fbFocoGrilla As Boolean

Private Sub cmdCancelar_Click()
    MostrarCobertura
    Habilitar False
End Sub
Private Sub cmdEditar_Click()
    Habilitar True
    EnfocaControl feCoberturas
    feCoberturas_RowColChange
End Sub
Private Sub Habilitar(ByVal pbHabilitar As Boolean)
    cmbProducto.Enabled = Not pbHabilitar
    cmbClienteTpo.Enabled = Not pbHabilitar
    'cmdMostrar.Enabled = Not pbHabilitar
    cmdEditar.Enabled = Not pbHabilitar
    cmdGuardar.Enabled = pbHabilitar
    cmdCancelar.Enabled = pbHabilitar
    feCoberturas.lbEditarFlex = pbHabilitar
    
    fbEditar = pbHabilitar
End Sub
Private Sub cmdCancelarEx_Click()
    MostrarExcepciones
    HabilitarEx False
End Sub
Private Sub cmdEditarEx_Click()
    HabilitarEx True
    EnfocaControl txtRetiroIntAdelantado
End Sub
Private Sub HabilitarEx(ByVal pbHabilitar As Boolean)
    txtRetiroIntAdelantado.Enabled = pbHabilitar
    txtRetiroIntMensual.Enabled = pbHabilitar
    txtRetiroIntFinal.Enabled = pbHabilitar
    
    txtRetiroIntAdelantadoCamp.Enabled = pbHabilitar
    txtRetiroIntMensualCamp.Enabled = pbHabilitar
    txtRetiroIntFinalCamp.Enabled = pbHabilitar
    
    cmdEditarEx.Enabled = Not pbHabilitar
    cmdGuardarEx.Enabled = pbHabilitar
    cmdCancelarEx.Enabled = pbHabilitar
End Sub
Private Sub cmdGuardarEx_Click()
    Dim oNGar As COMNCredito.NCOMGarantia
    
    On Error GoTo ErrGuardarEx
    
    If Not IsNumeric(txtRetiroIntAdelantado.Text) Then
        MsgBox "El monto de retiro de interés adelantado debe ser mayor a cero", vbInformation, "Aviso"
        EnfocaControl txtRetiroIntAdelantado
        Exit Sub
    Else
        If CCur(txtRetiroIntAdelantado.Text) <= 0 Then
            MsgBox "El monto de retiro de interés adelantado debe ser mayor a cero", vbInformation, "Aviso"
            EnfocaControl txtRetiroIntAdelantado
            Exit Sub
        ElseIf CCur(txtRetiroIntAdelantado.Text) > 100 Then
            MsgBox "El monto de retiro de interés adelantado no debe ser mayor a cien", vbInformation, "Aviso"
            EnfocaControl txtRetiroIntAdelantado
            Exit Sub
        End If
    End If
    If Not IsNumeric(txtRetiroIntMensual.Text) Then
        MsgBox "El monto de retiro de interés mensual debe ser mayor a cero", vbInformation, "Aviso"
        EnfocaControl txtRetiroIntMensual
        Exit Sub
    Else
        If CCur(txtRetiroIntMensual.Text) <= 0 Then
            MsgBox "El monto de retiro de interés mensual debe ser mayor a cero", vbInformation, "Aviso"
            EnfocaControl txtRetiroIntMensual
            Exit Sub
        ElseIf CCur(txtRetiroIntMensual.Text) > 100 Then
            MsgBox "El monto de retiro de interés mensual no debe ser mayor a cien", vbInformation, "Aviso"
            EnfocaControl txtRetiroIntMensual
            Exit Sub
        End If
    End If
    If Not IsNumeric(txtRetiroIntFinal.Text) Then
        MsgBox "El monto de retiro de interés a final del plazo debe ser mayor a cero", vbInformation, "Aviso"
        EnfocaControl txtRetiroIntFinal
        Exit Sub
    Else
        If CCur(txtRetiroIntFinal.Text) <= 0 Then
            MsgBox "El monto de retiro de interés a final del plazo debe ser mayor a cero", vbInformation, "Aviso"
            EnfocaControl txtRetiroIntFinal
            Exit Sub
        ElseIf CCur(txtRetiroIntFinal.Text) > 100 Then
            MsgBox "El monto de retiro de interés a final del plazo no debe ser mayor a cien", vbInformation, "Aviso"
            EnfocaControl txtRetiroIntFinal
            Exit Sub
        End If
    End If
    'Campañas
    If Not IsNumeric(txtRetiroIntAdelantadoCamp.Text) Then
        MsgBox "El monto de retiro de interés adelantado debe ser mayor a cero", vbInformation, "Aviso"
        EnfocaControl txtRetiroIntAdelantadoCamp
        Exit Sub
    Else
        If CCur(txtRetiroIntAdelantadoCamp.Text) <= 0 Then
            MsgBox "El monto de retiro de interés adelantado debe ser mayor a cero", vbInformation, "Aviso"
            EnfocaControl txtRetiroIntAdelantadoCamp
            Exit Sub
        ElseIf CCur(txtRetiroIntAdelantadoCamp.Text) > 100 Then
            MsgBox "El monto de retiro de interés adelantado no debe ser mayor a cien", vbInformation, "Aviso"
            EnfocaControl txtRetiroIntAdelantadoCamp
            Exit Sub
        End If
    End If
    If Not IsNumeric(txtRetiroIntMensualCamp.Text) Then
        MsgBox "El monto de retiro de interés mensual debe ser mayor a cero", vbInformation, "Aviso"
        EnfocaControl txtRetiroIntMensualCamp
        Exit Sub
    Else
        If CCur(txtRetiroIntMensualCamp.Text) <= 0 Then
            MsgBox "El monto de retiro de interés mensual debe ser mayor a cero", vbInformation, "Aviso"
            EnfocaControl txtRetiroIntMensualCamp
            Exit Sub
        ElseIf CCur(txtRetiroIntMensualCamp.Text) > 100 Then
            MsgBox "El monto de retiro de interés mensual no debe ser mayor a cien", vbInformation, "Aviso"
            EnfocaControl txtRetiroIntMensualCamp
            Exit Sub
        End If
    End If
    If Not IsNumeric(txtRetiroIntFinalCamp.Text) Then
        MsgBox "El monto de retiro de interés a final del plazo debe ser mayor a cero", vbInformation, "Aviso"
        EnfocaControl txtRetiroIntFinalCamp
        Exit Sub
    Else
        If CCur(txtRetiroIntFinalCamp.Text) <= 0 Then
            MsgBox "El monto de retiro de interés a final del plazo debe ser mayor a cero", vbInformation, "Aviso"
            EnfocaControl txtRetiroIntFinalCamp
            Exit Sub
        ElseIf CCur(txtRetiroIntFinalCamp.Text) > 100 Then
            MsgBox "El monto de retiro de interés a final del plazo no debe ser mayor a cien", vbInformation, "Aviso"
            EnfocaControl txtRetiroIntFinalCamp
            Exit Sub
        End If
    End If
    
    If MsgBox("¿Está seguro de guardar la configuración ingresada?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    cmdGuardarEx.Enabled = False
    Screen.MousePointer = 11
    
    Set oNGar = New COMNCredito.NCOMGarantia
    'oNGar.GrabarCoberturasConfigEX CCur(txtRetiroIntAdelantado.Text), CCur(txtRetiroIntMensual.Text), CCur(txtRetiroIntFinal.Text)
    oNGar.GrabarCoberturasConfigEX CCur(txtRetiroIntAdelantado.Text), CCur(txtRetiroIntMensual.Text), CCur(txtRetiroIntFinal.Text), CCur(txtRetiroIntAdelantadoCamp.Text), CCur(txtRetiroIntMensualCamp.Text), CCur(txtRetiroIntFinalCamp.Text)
    
    Screen.MousePointer = 0
    cmdGuardarEx.Enabled = True
    
    CargarExcepciones
    
    Exit Sub
ErrGuardarEx:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdMostrar_Click()
    Dim oDGar As COMDCredito.DCOMGarantia
    Dim i As Integer
    
    On Error GoTo ErrMostrar
    
    If cmbProducto.ListIndex = -1 Then
        Exit Sub
    End If
    If cmbClienteTpo.ListIndex = -1 Then
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    fsTpoProdCod = CInt(Trim(Right(cmbProducto.Text, 3)))
    fbCliPreferencial = IIf(CInt(Trim(Right(cmbClienteTpo.Text, 3))) = 1, True, False)
    
    Set oDGar = New COMDCredito.DCOMGarantia
    Set fRsCoberturas = New ADODB.Recordset

    Set fRsCoberturas = oDGar.RecuperaCoberturaConfig(fsTpoProdCod, fbCliPreferencial)

    MostrarCobertura
    Habilitar False
   
    Screen.MousePointer = 0
    Exit Sub
ErrMostrar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub MostrarCobertura()
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    
    Set rs = fRsCoberturas.Clone
    FormateaFlex feCoberturas
    If Not rs.EOF Then
        Do While Not rs.EOF
            feCoberturas.AdicionaFila
            i = feCoberturas.row
            feCoberturas.TextMatrix(i, 1) = rs!nBienGarantia
            feCoberturas.TextMatrix(i, 2) = rs!cCheck
            feCoberturas.TextMatrix(i, 3) = rs!cClasificacion
            feCoberturas.TextMatrix(i, 4) = rs!cBienGarantia
            feCoberturas.TextMatrix(i, 5) = Format(rs!nCoberturaPreferida, "#0.0000")
            feCoberturas.TextMatrix(i, 6) = Format(rs!nCoberturaNoPreferida, "#0.0000")
            rs.MoveNext
        Loop
    End If
    feCoberturas.row = 1
    feCoberturas.TopRow = 1
End Sub
Private Sub CargarExcepciones()
    Dim oDGar As New COMDCredito.DCOMGarantia
    Dim rs As New ADODB.Recordset
    
    Set rs = oDGar.RecuperaCoberturaConfigRapiFlash()
    Do While Not rs.EOF
        Select Case rs!nRetiroTpo
            Case eGarantiaTipoRetiroRapiFlash.nRetiroIntAdelantado
                fnRetiroIntAdelantado = rs!nRetiroMonto
            Case eGarantiaTipoRetiroRapiFlash.nRetiroIntMensual
                fnRetiroIntMensual = rs!nRetiroMonto
            Case eGarantiaTipoRetiroRapiFlash.nRetiroIntFinal
                fnRetiroIntFinal = rs!nRetiroMonto
        End Select
        rs.MoveNext
    Loop
    
    Set rs = oDGar.RecuperaCoberturaConfigRapiFlashCamp()
    Do While Not rs.EOF
        Select Case rs!nRetiroTpo
            Case eGarantiaTipoRetiroRapiFlash.nRetiroIntAdelantado
                fnRetiroIntAdelantadoCampania = rs!nRetiroMonto
            Case eGarantiaTipoRetiroRapiFlash.nRetiroIntMensual
                fnRetiroIntMensualCampania = rs!nRetiroMonto
            Case eGarantiaTipoRetiroRapiFlash.nRetiroIntFinal
                fnRetiroIntFinalCampania = rs!nRetiroMonto
        End Select
        rs.MoveNext
    Loop
    
    RSClose rs
    
    MostrarExcepciones
    HabilitarEx False
End Sub
Private Sub MostrarExcepciones()
    txtRetiroIntAdelantado.Text = Format(fnRetiroIntAdelantado, gsFormatoNumeroView)
    txtRetiroIntMensual.Text = Format(fnRetiroIntMensual, gsFormatoNumeroView)
    txtRetiroIntFinal.Text = Format(fnRetiroIntFinal, gsFormatoNumeroView)
    
    txtRetiroIntAdelantadoCamp.Text = Format(fnRetiroIntAdelantadoCampania, gsFormatoNumeroView)
    txtRetiroIntMensualCamp.Text = Format(fnRetiroIntMensualCampania, gsFormatoNumeroView)
    txtRetiroIntFinalCamp.Text = Format(fnRetiroIntFinalCampania, gsFormatoNumeroView)
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub feCoberturas_GotFocus()
    fbFocoGrilla = True
End Sub
Private Sub feCoberturas_LostFocus()
    fbFocoGrilla = False
End Sub
Private Sub feCoberturas_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(feCoberturas.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
    If pnCol = 5 Or pnCol = 6 Then
        If Not IsNumeric(feCoberturas.TextMatrix(pnRow, pnCol)) Then
            MsgBox "Ud. debe ingresar un monto mayor a cero", vbInformation, "Aviso"
            feCoberturas.row = pnRow
            feCoberturas.TopRow = pnRow
            feCoberturas.col = pnCol
            Cancel = False
            Exit Sub
        End If
    End If
End Sub
Private Sub feCoberturas_RowColChange()
    If fbEditar Then
        If feCoberturas.col = 5 Or feCoberturas.col = 6 Then
            If feCoberturas.TextMatrix(feCoberturas.row, 2) = "." Then
                feCoberturas.lbEditarFlex = True
            Else
                feCoberturas.lbEditarFlex = False
            End If
        Else
            feCoberturas.lbEditarFlex = True
        End If
    End If
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 11
    CargarControles
    Set fRsCoberturas = New ADODB.Recordset
    
    Call CambiaTamañoCombo(cmbProducto, 420)
    Screen.MousePointer = 0
End Sub
'Private Sub txtMonto_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(txtMonto, KeyAscii, 15)
'    If KeyAscii = 13 Then
'        EnfocaControl txtGlosa
'    End If
'End Sub
Private Sub Form_Unload(Cancel As Integer)
    RSClose fRsCoberturas
End Sub
Private Sub txtRetiroIntAdelantado_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtRetiroIntAdelantado, KeyAscii, 15)
    If KeyAscii = 13 Then
        EnfocaControl txtRetiroIntMensual
    End If
End Sub
Private Sub txtRetiroIntAdelantado_LostFocus()
    If IsNumeric(txtRetiroIntAdelantado.Text) Then
        If CCur(txtRetiroIntAdelantado.Text) > 100 Then
            MsgBox "El monto no debe ser mayor a 100%", vbInformation, "Aviso"
            EnfocaControl txtRetiroIntAdelantado
            Exit Sub
        End If
    End If
    txtRetiroIntAdelantado.Text = Format(txtRetiroIntAdelantado.Text, gsFormatoNumeroView)
End Sub
Private Sub txtRetiroIntAdelantadoCamp_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtRetiroIntAdelantadoCamp, KeyAscii, 15)
    If KeyAscii = 13 Then
        EnfocaControl txtRetiroIntMensualCamp
    End If
End Sub

Private Sub txtRetiroIntFinalCamp_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtRetiroIntFinalCamp, KeyAscii, 15)
    If KeyAscii = 13 Then
        EnfocaControl cmdGuardarEx
    End If
End Sub

Private Sub txtRetiroIntMensual_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtRetiroIntMensual, KeyAscii, 15)
    If KeyAscii = 13 Then
        EnfocaControl txtRetiroIntFinal
    End If
End Sub
Private Sub txtRetiroIntMensual_LostFocus()
    If IsNumeric(txtRetiroIntMensual.Text) Then
        If CCur(txtRetiroIntMensual.Text) > 100 Then
            MsgBox "El monto no debe ser mayor a 100%", vbInformation, "Aviso"
            EnfocaControl txtRetiroIntMensual
            Exit Sub
        End If
    End If
    txtRetiroIntMensual.Text = Format(txtRetiroIntMensual.Text, gsFormatoNumeroView)
End Sub
Private Sub txtRetiroIntFinal_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtRetiroIntFinal, KeyAscii, 15)
    If KeyAscii = 13 Then
        EnfocaControl txtRetiroIntAdelantadoCamp
    End If
End Sub
Private Sub txtRetiroIntFinal_LostFocus()
    If IsNumeric(txtRetiroIntFinal.Text) Then
        If CCur(txtRetiroIntFinal.Text) > 100 Then
            MsgBox "El monto no debe ser mayor a 100%", vbInformation, "Aviso"
            EnfocaControl txtRetiroIntFinal
            Exit Sub
        End If
    End If
    txtRetiroIntFinal.Text = Format(txtRetiroIntFinal.Text, gsFormatoNumeroView)
End Sub
Private Sub CargarControles()
    Dim oDGar As New COMDCredito.DCOMGarantia
    Dim rs As New ADODB.Recordset
    
    'Carga Sub Producto
    cmbProducto.Clear
    Set rs = oDGar.RecuperaProductoCoberturaConfig()
    Do While Not rs.EOF
        cmbProducto.AddItem rs!cdescripcion & Space(200) & rs!cCodigo
        rs.MoveNext
    Loop
    'Carga Tipo de Cliente
    cmbClienteTpo.Clear
    cmbClienteTpo.AddItem "PREFERENCIAL " & Space(200) & "1"
    cmbClienteTpo.AddItem "NO PREFERENCIAL " & Space(200) & "2"
    'Carga Excepciones
    CargarExcepciones
End Sub
Private Sub cmdGuardar_Click()
    Dim oNGar As New COMNCredito.NCOMGarantia
    Dim bExito As Boolean
    Dim i As Integer, Index As Integer
    Dim lvBienGarantia As Variant
    
    On Error GoTo ErrGuardar
    
    ReDim lvBienGarantia(3, 0)
    For i = 1 To feCoberturas.Rows - 1
        If feCoberturas.TextMatrix(i, 2) = "." Then
            If Not IsNumeric(feCoberturas.TextMatrix(i, 5)) Then
                MsgBox "Ud. debe ingresar un monto mayor o igual a 1.00", vbInformation, "Aviso"
                feCoberturas.row = i
                feCoberturas.TopRow = i
                feCoberturas.col = 5
                EnfocaControl feCoberturas
                feCoberturas_RowColChange
                Exit Sub
            Else
                If CCur(feCoberturas.TextMatrix(i, 5)) <= 0 Then
                    MsgBox "Ud. debe ingresar un monto mayor o igual a 1.00", vbInformation, "Aviso"
                    feCoberturas.row = i
                    feCoberturas.TopRow = i
                    feCoberturas.col = 5
                    EnfocaControl feCoberturas
                    feCoberturas_RowColChange
                    Exit Sub
                End If
            End If
            If Not IsNumeric(feCoberturas.TextMatrix(i, 6)) Then
                MsgBox "Ud. debe ingresar un monto mayor o igual a 1.00", vbInformation, "Aviso"
                feCoberturas.row = i
                feCoberturas.TopRow = i
                feCoberturas.col = 6
                EnfocaControl feCoberturas
                feCoberturas_RowColChange
                Exit Sub
            Else
                If CCur(feCoberturas.TextMatrix(i, 6)) <= 0 Then
                    MsgBox "Ud. debe ingresar un monto mayor o igual a 1.00", vbInformation, "Aviso"
                    feCoberturas.row = i
                    feCoberturas.TopRow = i
                    feCoberturas.col = 6
                    EnfocaControl feCoberturas
                    feCoberturas_RowColChange
                    Exit Sub
                End If
            End If
            Index = UBound(lvBienGarantia, 2) + 1
            ReDim Preserve lvBienGarantia(3, 0 To Index)
            lvBienGarantia(1, Index) = CInt(feCoberturas.TextMatrix(i, 1)) 'ID del Bien en Garantía
            lvBienGarantia(2, Index) = CCur(feCoberturas.TextMatrix(i, 5)) 'Monto Cobertura Preferida
            lvBienGarantia(3, Index) = CCur(feCoberturas.TextMatrix(i, 6)) 'Monto Cobertura NO Preferida
        End If
    Next
    
    If UBound(lvBienGarantia, 2) = 0 Then
        MsgBox "Ud. debe seleccionar al menos un registro para continuar", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("¿Está seguro de guardar la configuración ingresada?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    cmdGuardar.Enabled = False
    
    bExito = oNGar.GrabarCoberturasConfig(fsTpoProdCod, fbCliPreferencial, lvBienGarantia)
    
    cmdGuardar.Enabled = True
    Screen.MousePointer = 0
    
    If bExito Then
        MsgBox "Se ha grabado las coberturas satisfactoriamente", vbInformation, "Aviso"
        cmdMostrar_Click
    Else
        MsgBox "Ha sucedido un error al grabar las coberturas." & Chr(13) & Chr(13) & "Inténtelo nuevamente, si el error persiste comuníquelo con el Dpto. de TI", vbCritical, "Aviso"
    End If
    Exit Sub
ErrGuardar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If fbFocoGrilla Then
        If KeyCode = 86 And Shift = 2 Then
            KeyCode = 10
        End If
    End If
End Sub
Private Sub cmbProducto_Click()
    cmdMostrar_Click
End Sub
Private Sub cmbClienteTpo_Click()
    cmdMostrar_Click
End Sub
Private Sub txtRetiroIntMensualCamp_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtRetiroIntMensualCamp, KeyAscii, 15)
    If KeyAscii = 13 Then
        EnfocaControl txtRetiroIntFinalCamp
    End If
End Sub
