VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmColocLineaCredito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Línea de Crédito"
   ClientHeight    =   7575
   ClientLeft      =   1350
   ClientTop       =   615
   ClientWidth     =   9000
   Icon            =   "FrmColocCredLineaCredito.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   7560
      Left            =   15
      TabIndex        =   26
      Top             =   0
      Width           =   8940
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   450
         Left            =   7425
         TabIndex        =   22
         Top             =   1845
         Width           =   1410
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   450
         Left            =   7395
         TabIndex        =   23
         Top             =   6480
         Width           =   1410
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   450
         Left            =   7395
         TabIndex        =   25
         Top             =   7005
         Width           =   1410
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   450
         Left            =   7395
         TabIndex        =   24
         Top             =   7005
         Width           =   1410
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   450
         Left            =   7425
         TabIndex        =   21
         Top             =   1305
         Width           =   1410
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   450
         Left            =   7425
         TabIndex        =   20
         Top             =   765
         Width           =   1410
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   450
         Left            =   7425
         TabIndex        =   19
         Top             =   240
         Width           =   1410
      End
      Begin VB.Frame Frame3 
         Height          =   3960
         Left            =   120
         TabIndex        =   29
         Top             =   3540
         Width           =   7185
         Begin VB.ComboBox CmbEstado 
            Height          =   315
            Left            =   855
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   165
            Width           =   1590
         End
         Begin VB.CommandButton CmdTasaEliminar 
            Caption         =   "&Eliminar"
            Height          =   300
            Left            =   5940
            TabIndex        =   18
            Top             =   3555
            Width           =   1035
         End
         Begin VB.CommandButton CmdTasaEditar 
            Caption         =   "Editar"
            Height          =   300
            Left            =   5940
            TabIndex        =   17
            Top             =   3210
            Width           =   1035
         End
         Begin VB.CommandButton CmdTasaNuevo 
            Caption         =   "&Nuevo"
            Height          =   300
            Left            =   5940
            TabIndex        =   16
            Top             =   2865
            Width           =   1035
         End
         Begin SICMACT.FlexEdit FETasas 
            Height          =   1590
            Left            =   150
            TabIndex        =   15
            Top             =   2310
            Width           =   5685
            _ExtentX        =   10028
            _ExtentY        =   2805
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "-Tipo-Tasa Inicial-Tasa Final"
            EncabezadosAnchos=   "350-3200-900-900"
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
            ColumnasAEditar =   "X-1-2-3"
            ListaControles  =   "0-3-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-R"
            FormatosEdit    =   "0-0-2-2"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   285
         End
         Begin VB.Frame frmPlazos 
            Height          =   1560
            Left            =   105
            TabIndex        =   30
            Top             =   435
            Width           =   6975
            Begin VB.TextBox TxtSaldo 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2055
               TabIndex        =   14
               Top             =   1185
               Width           =   1395
            End
            Begin VB.ComboBox CmbLinea 
               Height          =   315
               Left            =   4980
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   870
               Width           =   1830
            End
            Begin VB.TextBox TxtMontomaxLinea 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2055
               TabIndex        =   12
               Top             =   870
               Width           =   1395
            End
            Begin VB.TextBox TxtMontoMin 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   4980
               TabIndex        =   11
               Top             =   540
               Width           =   930
            End
            Begin VB.TextBox TxtMontoMax 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2055
               TabIndex        =   10
               Top             =   540
               Width           =   930
            End
            Begin VB.TextBox TxtPlazoMin 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   4980
               TabIndex        =   9
               Top             =   210
               Width           =   930
            End
            Begin VB.TextBox TxtPlazoMax 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2055
               TabIndex        =   8
               Top             =   210
               Width           =   930
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Saldo de Linea :"
               Height          =   195
               Left            =   120
               TabIndex        =   39
               Top             =   1245
               Width           =   1155
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Linea :"
               Height          =   195
               Left            =   3780
               TabIndex        =   38
               Top             =   930
               Width           =   1065
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Monto Maximo de Linea :"
               Height          =   195
               Left            =   120
               TabIndex        =   37
               Top             =   930
               Width           =   1785
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Monto Minimo :"
               Height          =   195
               Left            =   3780
               TabIndex        =   36
               Top             =   600
               Width           =   1080
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Monto Maximo :"
               Height          =   195
               Left            =   120
               TabIndex        =   35
               Top             =   570
               Width           =   1125
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "dias"
               Height          =   195
               Left            =   5955
               TabIndex        =   34
               Top             =   255
               Width           =   285
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Plazo Minimo : "
               Height          =   195
               Left            =   3765
               TabIndex        =   33
               Top             =   270
               Width           =   1065
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "dias"
               Height          =   195
               Left            =   3045
               TabIndex        =   32
               Top             =   270
               Width           =   285
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Plazo Maximo :"
               Height          =   195
               Left            =   120
               TabIndex        =   31
               Top             =   240
               Width           =   1065
            End
         End
         Begin VB.CommandButton CmdTasaAceptar 
            Caption         =   "&Aceptar"
            Height          =   300
            Left            =   5940
            TabIndex        =   41
            Top             =   3210
            Width           =   1020
         End
         Begin VB.CommandButton CmdTasaCancelar 
            Caption         =   "&Cancelar"
            Height          =   300
            Left            =   5940
            TabIndex        =   42
            Top             =   3555
            Width           =   1035
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Estado :"
            Height          =   195
            Left            =   165
            TabIndex        =   46
            Top             =   210
            Width           =   630
         End
         Begin VB.Label Label18 
            Caption         =   "Tasas de Linea de Credito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   120
            TabIndex        =   40
            Top             =   2025
            Width           =   2850
         End
      End
      Begin VB.Frame fraCodigoLC 
         Caption         =   "Lineas de Credito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   135
         TabIndex        =   27
         Top             =   120
         Width           =   7200
         Begin VB.CommandButton CmdAtras 
            Caption         =   "<-"
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
            Left            =   150
            TabIndex        =   28
            Top             =   555
            Width           =   300
         End
         Begin MSDataGridLib.DataGrid DGLineas 
            Height          =   1455
            Left            =   120
            TabIndex        =   0
            Top             =   540
            Width           =   6945
            _ExtentX        =   12250
            _ExtentY        =   2566
            _Version        =   393216
            AllowUpdate     =   0   'False
            ColumnHeaders   =   -1  'True
            ForeColor       =   -2147483630
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
               DataField       =   "cLineaCred"
               Caption         =   "Codigo"
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
               DataField       =   "cDescripcion"
               Caption         =   "Descripcion"
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
                  ColumnWidth     =   1709.858
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   4800.189
               EndProperty
            EndProperty
         End
         Begin VB.ComboBox CmbPlazo 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2453
            Width           =   1770
         End
         Begin VB.ComboBox CmbMoneda 
            Height          =   315
            Left            =   1005
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   2453
            Width           =   1170
         End
         Begin VB.TextBox TxtCodLinea 
            Height          =   285
            Left            =   1005
            TabIndex        =   1
            Top             =   2123
            Width           =   1155
         End
         Begin VB.TextBox Txtdescrip 
            Height          =   300
            Left            =   2910
            TabIndex        =   2
            Top             =   2115
            Visible         =   0   'False
            Width           =   3690
         End
         Begin VB.ComboBox CmbPersFinan 
            Height          =   315
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2108
            Visible         =   0   'False
            Width           =   4020
         End
         Begin VB.ComboBox CmbProducto 
            Height          =   315
            Left            =   1005
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2865
            Width           =   3675
         End
         Begin VB.Label LblAyuda 
            BackColor       =   &H80000018&
            Height          =   240
            Left            =   180
            TabIndex        =   50
            Top             =   270
            Width           =   6810
         End
         Begin VB.Label LblTitulo 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   240
            Left            =   2955
            TabIndex        =   43
            Top             =   -30
            Width           =   780
         End
         Begin VB.Label LblDescLinea 
            Caption         =   "Descrip."
            Height          =   210
            Left            =   2280
            TabIndex        =   48
            Top             =   2160
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo :"
            Height          =   210
            Left            =   270
            TabIndex        =   47
            Top             =   2160
            Width           =   645
         End
         Begin VB.Label LblMoneda 
            Caption         =   "Moneda : "
            Height          =   210
            Left            =   255
            TabIndex        =   44
            Top             =   2505
            Width           =   780
         End
         Begin VB.Shape Shape1 
            Height          =   1275
            Left            =   120
            Top             =   2025
            Width           =   6945
         End
         Begin VB.Label LblPlazo 
            Caption         =   "Plazo : "
            Height          =   210
            Left            =   2295
            TabIndex        =   45
            Top             =   2505
            Width           =   570
         End
         Begin VB.Label Label2 
            Caption         =   "Producto :"
            Height          =   210
            Left            =   255
            TabIndex        =   49
            Top             =   2910
            Width           =   645
         End
      End
   End
End
Attribute VB_Name = "FrmColocLineaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oLineas As DLineaCredito
Dim R As ADODB.Recordset
Dim RTasas As ADODB.Recordset
Dim nTipoIngreso As Integer '1: Fondo; 2: Sub Fondo; 3:Plazo; 4:Producto; 5:Version
Dim psPersCod As String
Dim sEstructuraLinea As String
Dim sEstructuraLineaAnt As String
Dim bCargandoControles As Boolean
Dim KeyTmp As String
Dim CmdEjecutar As Integer
Dim CmdEjecutarTasas As Integer
Dim sAyudaAnt As String
Private Function ValidaDatos() As Boolean
    ValidaDatos = True
    Select Case nTipoIngreso
        Case 1 'Fondo
            If Len(Trim(TxtCodLinea.Text)) < 2 Then
                MsgBox "Falta Ingresar el Codigo de la Linea de Credito", vbInformation, "Aviso"
                ValidaDatos = False
                TxtCodLinea.SetFocus
                Exit Function
            End If
            If CmbPersFinan.ListIndex = -1 Then
                MsgBox "Ingrese la Fuente de Financiamiento de la Linea de Credito", vbInformation, "Aviso"
                ValidaDatos = False
                CmbPersFinan.SetFocus
                Exit Function
            End If
            If CmbEstado.ListIndex = -1 Then
                MsgBox "Seleccione el Estado de la Linea de Credito", vbInformation, "Aviso"
                ValidaDatos = False
                CmbEstado.SetFocus
                Exit Function
            End If
        Case 2 'Sub fondo
            If Len(Trim(TxtCodLinea.Text)) < 4 Then
                MsgBox "Falta Ingresar el Codigo de la Linea de Credito", vbInformation, "Aviso"
                ValidaDatos = False
                TxtCodLinea.SetFocus
                Exit Function
            End If
            If Len(Trim(Txtdescrip.Text)) = 0 Then
                MsgBox "Ingrese la Descripcion de la Linea de Credito", vbInformation, "Aviso"
                ValidaDatos = False
                Txtdescrip.SetFocus
                Exit Function
            End If
            If CmbMoneda.ListIndex = -1 Then
                MsgBox "Seleccione la Moneda", vbInformation, "Aviso"
                ValidaDatos = False
                CmbMoneda.SetFocus
                Exit Function
            End If
            If CmbEstado.ListIndex = -1 Then
                MsgBox "Seleccione el Estado de la Linea de Credito", vbInformation, "Aviso"
                ValidaDatos = False
                CmbEstado.SetFocus
                Exit Function
            End If
        Case 3 'Plazo
            If Len(Trim(Txtdescrip.Text)) = 0 Then
                MsgBox "Ingrese la Descripcion de la Linea de Credito", vbInformation, "Aviso"
                ValidaDatos = False
                Txtdescrip.SetFocus
                Exit Function
            End If
            If CmbPlazo.ListIndex = -1 Then
                MsgBox "Seleccione el Plazo de la Linea de Credito", vbInformation, "Aviso"
                ValidaDatos = False
                CmbPlazo.SetFocus
                Exit Function
            End If
            If CmbEstado.ListIndex = -1 Then
                MsgBox "Seleccione el Estado de la Linea de Credito", vbInformation, "Aviso"
                ValidaDatos = False
                CmbEstado.SetFocus
                Exit Function
            End If
        Case 4 'Producto
            If Len(Trim(Txtdescrip.Text)) = 0 Then
                MsgBox "Ingrese la Descripcion de la Linea de Credito", vbInformation, "Aviso"
                ValidaDatos = False
                Txtdescrip.SetFocus
                Exit Function
            End If
            If CmbProducto.ListIndex = -1 Then
                MsgBox "Seleccione el producto de la Linea de Credito", vbInformation, "Aviso"
                ValidaDatos = False
                CmbProducto.SetFocus
                Exit Function
            End If
            If CmbEstado.ListIndex = -1 Then
                MsgBox "Seleccione el Estado de la Linea de Credito", vbInformation, "Aviso"
                ValidaDatos = False
                CmbEstado.SetFocus
                Exit Function
            End If
        Case 5 'Version
            If Len(Trim(TxtCodLinea.Text)) < 11 Then
                MsgBox "Falta Ingresar el Codigo de la Linea de Credito", vbInformation, "Aviso"
                ValidaDatos = False
                TxtCodLinea.SetFocus
                Exit Function
            End If
            If CmbEstado.ListIndex = -1 Then
                MsgBox "Seleccione el Estado de la Linea de Credito", vbInformation, "Aviso"
                ValidaDatos = False
                CmbEstado.SetFocus
                Exit Function
            End If
            If Len(Trim(TxtPlazoMax.Text)) = 0 Then
                MsgBox "Ingrese el Plazo Maximo de la Linea de Credito", vbInformation, "Aviso"
                ValidaDatos = False
                TxtPlazoMax.SetFocus
                Exit Function
            End If
            If Len(Trim(TxtPlazoMin.Text)) = 0 Then
                MsgBox "Ingrese el Plazo Minimo de la Linea de Credito", vbInformation, "Aviso"
                ValidaDatos = False
                TxtPlazoMin.SetFocus
                Exit Function
            End If
            If Len(Trim(TxtMontoMax.Text)) = 0 Then
                MsgBox "Ingrese el Monto Maximo de de Prestamo", vbInformation, "Aviso"
                ValidaDatos = False
                TxtMontoMax.SetFocus
                Exit Function
            End If
            If Len(Trim(TxtMontoMin.Text)) = 0 Then
                MsgBox "Ingrese el Monto Minimo de Prestamo", vbInformation, "Aviso"
                ValidaDatos = False
                TxtMontoMin.SetFocus
                Exit Function
            End If
            If Len(Trim(TxtMontomaxLinea.Text)) = 0 Then
                MsgBox "Ingrese el Monto Maximo de Linea de Credito", vbInformation, "Aviso"
                ValidaDatos = False
                TxtMontomaxLinea.SetFocus
                Exit Function
            End If
            If Len(Trim(TxtSaldo.Text)) = 0 Then
                MsgBox "Ingrese el Saldo de la Linea de Credito", vbInformation, "Aviso"
                ValidaDatos = False
                TxtSaldo.SetFocus
                Exit Function
            End If
            If Len(Trim(Txtdescrip.Text)) = 0 Then
                MsgBox "Ingrese la Descripcion de la Linea de Credito", vbInformation, "Aviso"
                ValidaDatos = False
                Txtdescrip.SetFocus
                Exit Function
            End If
    End Select
    
End Function
Public Sub InicioConsulta()
    CmdNuevo.Enabled = False
    CmdEditar.Enabled = False
    CmdEliminar.Enabled = False
    Me.Show 1
End Sub
Public Sub InicioActualiza()
    Me.Show 1
End Sub
Private Sub HabilitarTasasLinea(ByVal pbHabilita As Boolean)
    fraCodigoLC.Enabled = pbHabilita
    CmbEstado.Enabled = pbHabilita
    frmPlazos.Enabled = pbHabilita
    CmdAceptar.Enabled = pbHabilita
    CmdCancelar.Enabled = pbHabilita
End Sub
Private Sub CargaControles()
Dim RTemp As ADODB.Recordset
Dim oIFinan As DInstFinanc
Dim oConstante As DConstante
Dim i As Integer

    bCargandoControles = True
    On Error GoTo ERRORCargaControles
    'Carga Combo de Flex Edit
    Set oConstante = New DConstante
    FETasas.CargaCombo oConstante.RecuperaConstantes(gColocLineaCredTasas)
    Set oConstante = Nothing
    'Carga Monedas
    Call CargaComboConstante(gMoneda, CmbMoneda)
    'Carga Plazos de Linea
    Call CargaComboConstante(gColocLineaCredPlazo, CmbPlazo)
    'Carga Estados de Linea
    CmbEstado.AddItem "ACTIVA" & Space(50) & "1"
    CmbEstado.AddItem "INACTIVA" & Space(50) & "0"
        
    'Carga Fondos de Linea de Credito
    Set oIFinan = New DInstFinanc
    Set RTemp = oIFinan.RecuperaIFinancieraPersCod
    CmbPersFinan.Clear
    Do While Not RTemp.EOF
        CmbPersFinan.AddItem RTemp!cPersNombre & Space(50) & RTemp!cPersCod
        RTemp.MoveNext
    Loop
    RTemp.Close
    Set RTemp = Nothing
    Set oIFinan = Nothing
    
    'Tipos de linea
    CmbLinea.Clear
    CmbLinea.AddItem ""
    
    'Carga Productos de Linea de Credito
    Call CargaComboConstante(gProducto, CmbProducto)
    Do While i < CmbProducto.ListCount
        If Mid(Trim(Right(CmbProducto.List(i), 10)), 1, 2) = "23" Then
            CmbProducto.RemoveItem (i)
            i = i - 1
        End If
        i = i + 1
    Loop
    bCargandoControles = False
    
    Exit Sub
    
ERRORCargaControles:
    MsgBox Err.Description, vbCritical, "Aviso"
    bCargandoControles = False
End Sub
Private Sub HabilitaIngreso(ByVal pbHabilita As Boolean, ByVal pTipoIng As Integer)
'1: Fondo; 2:Subfondo; 3: Plazo; 4:Producto; 5:Version
    DGLineas.Enabled = Not pbHabilita
    If pTipoIng = 1 Then
        Txtdescrip.Visible = Not pbHabilita
        LblDescLinea.Caption = "Fondo :"
        CmbPersFinan.Visible = pbHabilita
    Else
        Txtdescrip.Visible = pbHabilita
        CmbPersFinan.Visible = Not pbHabilita
        LblDescLinea.Caption = "Descrip. :"
    End If
    If pTipoIng = 4 Then
        CmbProducto.Enabled = pbHabilita
    Else
        CmbProducto.Enabled = False
    End If
    CmbEstado.Enabled = pbHabilita
    If pTipoIng = 5 Or pbHabilita = False Then
        TxtPlazoMin.Enabled = pbHabilita
        TxtPlazoMax.Enabled = pbHabilita
        TxtMontoMin.Enabled = pbHabilita
        TxtMontoMax.Enabled = pbHabilita
        TxtMontomaxLinea.Enabled = pbHabilita
        CmbLinea.Enabled = pbHabilita
        TxtSaldo.Enabled = pbHabilita
        CmdTasaNuevo.Enabled = pbHabilita
        CmdTasaEditar.Enabled = pbHabilita
        CmdTasaEliminar.Enabled = pbHabilita
        FETasas.lbEditarFlex = pbHabilita
    End If
    Select Case pTipoIng
        Case 1 'Fondo
            CmbMoneda.Enabled = False
            CmbPlazo.Enabled = False
            TxtCodLinea.Enabled = pbHabilita
            TxtCodLinea.MaxLength = 2
        Case 2 'Sub fondo
            CmbMoneda.Enabled = True
            CmbPlazo.Enabled = False
            TxtCodLinea.Enabled = pbHabilita
            TxtCodLinea.MaxLength = 4 'Mas la moneda
        Case 3 'Plazo
            CmbMoneda.Enabled = False
            CmbPlazo.Enabled = True
            TxtCodLinea.Enabled = False
            TxtCodLinea.MaxLength = 5 'FF + SF + M + Plazo
        Case 4 'Producto
            CmbMoneda.Enabled = False
            CmbPlazo.Enabled = False
            TxtCodLinea.Enabled = False
            TxtCodLinea.MaxLength = 6
        Case 5
            CmbMoneda.Enabled = False
            CmbPlazo.Enabled = False
            TxtCodLinea.Enabled = pbHabilita
            TxtCodLinea.MaxLength = 11
    End Select
        
    
    If pTipoIng = 5 Then
        CmdTasaNuevo.Enabled = pbHabilita
        CmdTasaEditar.Enabled = pbHabilita
        CmdTasaEliminar.Enabled = pbHabilita
        CmdTasaAceptar.Enabled = pbHabilita
        CmdTasaCancelar.Enabled = pbHabilita
    End If
    
    CmdNuevo.Enabled = Not pbHabilita
    CmdEditar.Enabled = Not pbHabilita
    CmdEliminar.Enabled = Not pbHabilita
    CmdImprimir.Enabled = Not pbHabilita
    CmdSalir.Visible = Not pbHabilita
    CmdAceptar.Visible = pbHabilita
    CmdCancelar.Visible = pbHabilita
    DGLineas.Height = IIf(pbHabilita, 1455, 2760)
End Sub
Private Sub LimpiaPantalla()
    Txtdescrip.Text = ""
    CmbEstado.ListIndex = -1
    TxtPlazoMax.Text = "0"
    TxtPlazoMin.Text = "0"
    TxtMontoMax.Text = "0.00"
    TxtMontoMin.Text = "0.00"
    TxtMontomaxLinea.Text = "0.00"
    CmbLinea.ListIndex = -1
    TxtSaldo.Text = "0.00"
    FETasas.TextMatrix(1, 0) = ""
    FETasas.TextMatrix(1, 1) = ""
    FETasas.TextMatrix(1, 2) = ""
    FETasas.TextMatrix(1, 3) = ""
    FETasas.Rows = 2
End Sub
Private Sub CargaLineasDetalles()
Dim oLinea As DLineaCredito
    FETasas.Rows = 2
    FETasas.TextMatrix(1, 0) = ""
    FETasas.TextMatrix(1, 1) = ""
    FETasas.TextMatrix(1, 2) = ""
    FETasas.TextMatrix(1, 3) = ""
    If Not R.BOF And Not R.EOF Then
        TxtCodLinea.Text = R!cLineacred
        Txtdescrip.Text = R!cDescripcion
        If nTipoIngreso >= 3 Then
            CmbMoneda.ListIndex = IndiceListaCombo(CmbMoneda, Mid(R!cLineacred, 5, 1))
        End If
        If nTipoIngreso >= 4 Or (nTipoIngreso >= 3 And CmdEjecutar = 2) Then
            CmbPlazo.ListIndex = IndiceListaCombo(CmbPlazo, Mid(R!cLineacred, 6, 1))
        End If
        CmbProducto.ListIndex = IndiceListaCombo(CmbProducto, Mid(R!cLineacred, 7, 3))
        CmbEstado.ListIndex = IndiceListaCombo(CmbEstado, Trim(Str(IIf(R!bEstado, 1, 0))))
        TxtPlazoMax.Text = Format(R!nPlazoMax, "#0")
        TxtPlazoMin.Text = Format(R!nPlazoMin, "#0")
        TxtMontoMin.Text = Format(R!nMontoMin, "#0.00")
        TxtMontoMax.Text = Format(R!nMontoMax, "#0.00")
        TxtMontomaxLinea.Text = Format(R!nMontoMaxLinea, "#0.00")
        CmbLinea.ListIndex = IndiceListaCombo(CmbLinea, Trim(IIf(IsNull(R!cTipo), "", R!cTipo)))
        TxtSaldo.Text = Format(R!nSaldo, "#0.00")
        
        'Carga Tasas de Linea
        Set oLinea = New DLineaCredito
        Set RTasas = oLinea.RecuperaLineasTasas(R!cLineacred)
        Do While Not RTasas.EOF
            FETasas.AdicionaFila
            FETasas.TextMatrix(RTasas.Bookmark, 1) = RTasas!cConsDescripcion & Space(50) & Trim(Str(CInt(Trim(RTasas!cColocLinCredTasaTpo))))
            FETasas.TextMatrix(RTasas.Bookmark, 2) = Format(RTasas!nTasaIni, "#0.00")
            FETasas.TextMatrix(RTasas.Bookmark, 3) = Format(RTasas!nTasafin, "#0.00")
            RTasas.MoveNext
        Loop
        RTasas.Close
        Set RTasas = Nothing
        If Len(Trim(FETasas.TextMatrix(1, 1))) > 0 Then
            FETasas.Row = 1
        End If
    End If
End Sub
Private Sub CargaLineas()
On Error GoTo ErroCargaLineas
    bCargandoControles = True
    If Not R Is Nothing Then
        R.Close
        Set R = Nothing
    End If
    Set oLineas = New DLineaCredito
    If nTipoIngreso = 1 Then
        Set R = oLineas.RecuperaLineasCredito(nTipoIngreso)
    Else
        Set R = oLineas.RecuperaLineasCredito(nTipoIngreso, sEstructuraLinea)
    End If
    If Not R.BOF And Not R.EOF Then
        If nTipoIngreso = 1 Then
            psPersCod = R!cPersCod
        End If
    End If
    Set DGLineas.DataSource = R
    DGLineas.Refresh
    If R.RecordCount > 0 Then
        R.Find "cLineaCred = '" & sEstructuraLineaAnt & "'"
        If R.EOF Then
            R.MoveFirst
        End If
    End If
    bCargandoControles = False
    Exit Sub
    
ErroCargaLineas:
    bCargandoControles = False
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub CargaDatos(ByVal psCodLinea As String)

On Error GoTo ERRORCargaDatos

    R.Find "cLineaCred = '" & psCodLinea & "'"
    If Not R.BOF And Not R.EOF Then
        CmbEstado.ListIndex = IndiceListaCombo(CmbEstado, R!bEstado)
        TxtPlazoMax.Text = Format(R!nPlazoMax, "#0.00")
        TxtPlazoMin.Text = Format(R!nPlazoMin, "#0.00")
        TxtMontoMax.Text = Format(R!nMontoMax, "#0.00")
        TxtMontoMin.Text = Format(R!nMontoMin, "#0.00")
        TxtMontomaxLinea.Text = Format(R!nMontoMaxLinea, "#0.00")
        CmbLinea.ListIndex = IndiceListaCombo(CmbLinea, R!cTipo)
        TxtSaldo.Text = Format(R!nSaldo, "#0.00")
        FETasas.Clear
        Set RTasas = oLineas.RecuperaLineasTasas(R!cLineacred)
        Do While Not R.EOF
            FETasas.AdicionaFila
            FETasas.TextMatrix(RTasas.Bookmark, 1) = RTasas!cConsDescripcion & Space(50) & RTasas!cColocLinCredTasaTpo
            FETasas.TextMatrix(RTasas.Bookmark, 2) = Format(RTasas!nTasaIni, "#0.00")
            FETasas.TextMatrix(RTasas.Bookmark, 3) = Format(RTasas!nTasafin, "#0.00")
            R.MoveNext
        Loop
        RTasas.Close
        Set RTasas = Nothing
    End If
    Exit Sub
    
ERRORCargaDatos:
    MsgBox Err.Description, vbInformation, "Aviso"
    
End Sub



Private Sub CmbEstado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtPlazoMax.Enabled Then
            TxtPlazoMax.SetFocus
        Else
            CmdAceptar.SetFocus
        End If
    End If
End Sub


Private Sub CmbLinea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtSaldo.SetFocus
    End If
End Sub

Private Sub CmbMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmbPlazo.Enabled Then
            CmbPlazo.SetFocus
        Else
            CmbEstado.SetFocus
        End If
    End If
End Sub



Private Sub CmbPlazo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmbEstado.SetFocus
    End If
End Sub

Private Sub CmbProducto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmbEstado.SetFocus
    End If
End Sub

Private Sub CmdAceptar_Click()
Dim oLinea As DLineaCredito
Dim i As Integer
Dim CadDes As String
Dim oLineaNeg As NLineaCredito
Dim nSaldoLinea As Double
On Error GoTo ErrorCmdAceptar_Click
    
    If Not ValidaDatos Then
        Exit Sub
    End If
    If CmdEjecutar = 1 Then
                If nTipoIngreso = 1 Then
                     CadDes = "FONDO " & Trim(Left(CmbPersFinan.Text, Len(CmbPersFinan.Text) - 13))
                     psPersCod = Right(CmbPersFinan.Text, 13)
                Else
                    TxtCodLinea.MaxLength = 0
                    CadDes = Trim(Txtdescrip.Text)
                    If nTipoIngreso = 2 Then
                        TxtCodLinea.Text = Trim(TxtCodLinea.Text) & Trim(Right(CmbMoneda.Text, 2))
                    Else
                        If nTipoIngreso = 3 Then
                            TxtCodLinea.Text = Trim(TxtCodLinea.Text) & Trim(Right(CmbPlazo.Text, 2))
                        Else
                            If nTipoIngreso = 4 Then
                                TxtCodLinea.Text = Trim(TxtCodLinea.Text) & Trim(Right(CmbProducto.Text, 10))
                            Else
                                If nTipoIngreso = 5 Then
                                    Set oLineaNeg = New NLineaCredito
                                    nSaldoLinea = oLineaNeg.ChequeaSaldoAColocarLineaCredito(Trim(TxtCodLinea.Text))
                                    Set oLineaNeg = Nothing
                                    If CDbl(TxtSaldo.Text) > nSaldoLinea Then
                                        MsgBox "El Saldo de la Linea es Mayor al Permitido, se Cambiara el Saldo de la Linea", vbInformation, "Aviso"
                                        TxtSaldo.Text = Format(nSaldoLinea, "#0.00")
                                        TxtCodLinea.MaxLength = 11
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                
                'Proceso de Grabacion
                Set oLinea = New DLineaCredito
                oLinea.IniciaGrabado
                Call oLinea.NuevaLineaCredito(Trim(TxtCodLinea.Text), CadDes, CInt(Trim(Right(CmbEstado.Text, 2))), CInt(TxtPlazoMax.Text), CInt(TxtPlazoMin.Text), CDbl(Format(TxtMontoMax.Text, "#0.00")), CDbl(Format(TxtMontoMin.Text, "#0.00")), CDbl(Format(TxtMontomaxLinea.Text, "#0.00")), Trim(Right(CmbLinea.Text, 2)), 0#, CDbl(Format(TxtSaldo.Text, "#0.00")), psPersCod)
                For i = 1 To FETasas.Rows - 1
                    If Len(Trim(FETasas.TextMatrix(i, 0))) > 0 Then
                        Call oLinea.NuevaLineaCreditoTasas(Trim(TxtCodLinea.Text), Right(FETasas.TextMatrix(i, 1), 2), Format(CDbl(FETasas.TextMatrix(i, 2)), "#0.00"), Format(CDbl(FETasas.TextMatrix(i, 3)), "#0.00"))
                    End If
                Next i
                'If nTipoIngreso <> 5 Then
                 Call oLinea.CreaSaldoLineaCredito(Trim(TxtCodLinea.Text), False)
                'End If
                oLinea.FinalizaGrabado
                Set oLinea = Nothing
                
        Else
            
            If nTipoIngreso = 5 Then
                Set oLineaNeg = New NLineaCredito
                nSaldoLinea = oLineaNeg.ChequeaSaldoAColocarLineaCredito(Trim(TxtCodLinea.Text))
                Set oLineaNeg = Nothing
                If CDbl(TxtSaldo.Text) > nSaldoLinea Then
                    MsgBox "El Saldo de la Linea es Mayor al Permitido, se Cambiara el Salod de la Linea", vbInformation, "Aviso"
                    TxtSaldo.Text = Format(nSaldoLinea, "#0.00")
                    TxtCodLinea.MaxLength = 11
                    Exit Sub
                End If
            End If
            
            Set oLinea = New DLineaCredito
            oLinea.IniciaGrabado
            Call oLinea.ActualizarLinea(Trim(TxtCodLinea.Text), Txtdescrip.Text, CInt(Trim(Right(CmbEstado.Text, 2))), CInt(TxtPlazoMax.Text), CInt(TxtPlazoMin.Text), CDbl(Format(TxtMontoMax.Text, "#0.00")), CDbl(Format(TxtMontoMin.Text, "#0.00")), CDbl(Format(TxtMontomaxLinea.Text, "#0.00")), Trim(Right(CmbLinea.Text, 2)), 0#, CDbl(Format(TxtSaldo.Text, "#0.00")), psPersCod)
            
            If Len(Trim(FETasas.TextMatrix(1, 0))) > 0 Then
                Call oLinea.EliminaTasasLinea(Trim(TxtCodLinea.Text))
                For i = 1 To FETasas.Rows - 1
                    If Len(Trim(FETasas.TextMatrix(i, 0))) > 0 Then
                        Call oLinea.ActualizarLineaTasas(Trim(TxtCodLinea.Text), Right(FETasas.TextMatrix(i, 1), 2), Format(CDbl(FETasas.TextMatrix(i, 2)), "#0.00"), Format(CDbl(FETasas.TextMatrix(i, 3)), "#0.00"))
                    End If
                Next i
            End If
            oLinea.FinalizaGrabado
            Set oLinea = Nothing
        End If
        Call HabilitaIngreso(False, nTipoIngreso)
        Call CargaLineas
        If DGLineas.Enabled Then
            DGLineas.SetFocus
        End If
    Exit Sub
    
ErrorCmdAceptar_Click:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub CmdAtras_Click()
    DGLineas.SetFocus
    nTipoIngreso = nTipoIngreso - 1
    If nTipoIngreso < 0 Then
        nTipoIngreso = 1
        CmdAtras.Enabled = False
        Exit Sub
    End If
    R.Close
    Set R = Nothing
    sEstructuraLineaAnt = sEstructuraLinea
    Select Case nTipoIngreso
        Case 1
            LblTitulo.Caption = "FONDO"
            CmdAtras.Enabled = False
        Case 2
            LblTitulo.Caption = "SUB FONDO"
            sEstructuraLinea = Mid(sEstructuraLinea, 1, 2)
        Case 3
            LblTitulo.Caption = "PLAZOS"
            sEstructuraLinea = Mid(sEstructuraLinea, 1, 5)
        Case 4
            LblTitulo.Caption = "PRODUCTO"
            sEstructuraLinea = Mid(sEstructuraLinea, 1, 6)
        Case 5
            LblTitulo.Caption = "VERSION"
            sEstructuraLinea = Mid(sEstructuraLinea, 1, 9)
    End Select
    LblAyuda.Caption = Trim(Right(sAyudaAnt, 50))
    If nTipoIngreso = 1 Then
        sAyudaAnt = ""
    Else
        sAyudaAnt = Mid(sAyudaAnt, 1, (nTipoIngreso - 1) * 50)
    End If
    
    Call CargaLineas
    Call CargaLineasDetalles
    
End Sub

Private Sub CmdCancelar_Click()
    Call LimpiaPantalla
    Call HabilitaIngreso(False, 1)
End Sub

Private Sub cmdEditar_Click()
    If R Is Nothing Then
        MsgBox "No existe linea a Editar", vbInformation, "Aviso"
        Exit Sub
    Else
        If R.RecordCount = 0 Then
            MsgBox "No existe linea a Editar", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    Call HabilitaIngreso(True, nTipoIngreso)
    Select Case nTipoIngreso
        Case 1
            TxtCodLinea.MaxLength = 2
        Case 2
            TxtCodLinea.MaxLength = 5
        Case 3
            TxtCodLinea.MaxLength = 6
        Case 4
            TxtCodLinea.MaxLength = 9
        Case 5
            TxtCodLinea.MaxLength = 11
    End Select
    Call CargaLineasDetalles
    CmdEjecutar = 2
End Sub

Private Sub CmdEliminar_Click()
Dim oLinea As DLineaCredito
    On Error GoTo ErrorCmdEliminar_Click
    If MsgBox("Se va a Eliminar la Linea de Credito, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set oLinea = New DLineaCredito
        oLinea.EliminaLineaCredito (R!cLineacred)
        R.Close
        Set R = Nothing
        Call CargaLineas
        Call CargaLineasDetalles
        DGLineas.SetFocus
    End If
    Exit Sub
ErrorCmdEliminar_Click:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub CmdNuevo_Click()
    Call LimpiaPantalla
    Call HabilitaIngreso(True, nTipoIngreso)
    CmdEjecutar = 1
    TxtCodLinea.Text = sEstructuraLinea
    If Txtdescrip.Enabled And Txtdescrip.Visible And nTipoIngreso <> 2 Then
        Txtdescrip.SetFocus
    Else
        If TxtCodLinea.Enabled Then
            TxtCodLinea.SetFocus
        Else
            If CmbMoneda.Enabled Then
                CmbMoneda.SetFocus
            Else
                If CmbPlazo.Enabled Then
                    CmbPlazo.SetFocus
                Else
                    If CmbProducto.Enabled Then
                        CmbProducto.SetFocus
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdTasaAceptar_Click()
    CmdTasaNuevo.Visible = True
    CmdTasaEditar.Visible = True
    CmdTasaEliminar.Visible = True
    CmdTasaAceptar.Visible = False
    CmdTasaAceptar.Enabled = False
    CmdTasaCancelar.Visible = False
    CmdTasaCancelar.Enabled = False
    Call HabilitarTasasLinea(True)
End Sub

Private Sub CmdTasaCancelar_Click()
    Call HabilitarTasasLinea(True)
    CmdTasaNuevo.Visible = True
    CmdTasaEditar.Visible = True
    CmdTasaEliminar.Visible = True
    CmdTasaAceptar.Visible = False
    CmdTasaAceptar.Enabled = False
    CmdTasaCancelar.Visible = False
    CmdTasaCancelar.Enabled = False
    CmbEstado.Enabled = True
    CmdAceptar.Enabled = True
    CmdCancelar.Enabled = True
    Call CargaLineasDetalles
End Sub

Private Sub CmdTasaEditar_Click()
    CmdEjecutarTasas = 2
    FETasas.lbEditarFlex = True
    CmdTasaNuevo.Visible = False
    CmdTasaEditar.Visible = False
    CmdTasaEliminar.Visible = False
    CmdTasaAceptar.Visible = True
    CmdTasaAceptar.Enabled = True
    CmdTasaCancelar.Visible = True
    CmdTasaCancelar.Enabled = True
    Call HabilitarTasasLinea(False)
    FETasas.SetFocus
End Sub

Private Sub CmdTasaEliminar_Click()
    CmdEjecutarTasas = 3
    If MsgBox("Desea Eliminar el Registro ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Call FETasas.EliminaFila(FETasas.Row)
    End If
End Sub

Private Sub CmdTasaNuevo_Click()
    CmdEjecutarTasas = 1
    FETasas.lbEditarFlex = True
    FETasas.AdicionaFila
    CmdTasaNuevo.Visible = False
    CmdTasaEditar.Visible = False
    CmdTasaEliminar.Visible = False
    CmdTasaAceptar.Visible = True
    CmdTasaAceptar.Enabled = True
    CmdTasaCancelar.Visible = True
    CmdTasaCancelar.Enabled = True
    Call HabilitarTasasLinea(False)
    FETasas.SetFocus
End Sub

Private Sub DGLineas_DblClick()
    nTipoIngreso = nTipoIngreso + 1
    CmdAtras.Enabled = True
    If nTipoIngreso > 5 Then
        nTipoIngreso = 5
        Exit Sub
    End If
    
    Select Case nTipoIngreso
        Case 1
            LblTitulo.Caption = "FONDO"
        Case 2
            LblTitulo.Caption = "SUB FONDO"
        Case 3
            LblTitulo.Caption = "PLAZO"
        Case 4
            LblTitulo.Caption = "PRODUCTO"
        Case 5
            LblTitulo.Caption = "VERSION"
    End Select
    If Not R Is Nothing Then
        If Not R.BOF And Not R.EOF Then
            sEstructuraLinea = R!cLineacred
        End If
    End If
    sAyudaAnt = sAyudaAnt & LblAyuda.Caption & Space(50 - Len(LblAyuda.Caption))
    LblAyuda.Caption = Trim(R!cDescripcion)
    sEstructuraLineaAnt = sEstructuraLinea
    Call CargaLineas
    Call CargaLineasDetalles
    
End Sub

Private Sub DGLineas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call DGLineas_DblClick
    Else
        If KeyAscii = 8 Then
            Call CmdAtras_Click
        End If
    End If
End Sub

Private Sub DGLineas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not bCargandoControles Then
        If nTipoIngreso = 1 Then
            psPersCod = R!cPersCod
        End If
    End If
    Call CargaLineasDetalles
End Sub

Private Sub Form_Load()
    nTipoIngreso = 1
    CmdEjecutar = -1
    CmdEjecutarTasas = -1
    Call HabilitaIngreso(False, nTipoIngreso)
    Call CargaLineas
    Call CargaControles
    CmdAtras.Enabled = False
    DGLineas.Height = 2760
    sEstructuraLinea = ""
    sEstructuraLineaAnt = ""
End Sub


Private Sub TxtCodLinea_Click()
    If Not bCargandoControles Then
        Select Case nTipoIngreso
            Case 2 'Sub Fondo
                If TxtCodLinea.SelStart < 2 Then
                    TxtCodLinea.SelStart = 2
                End If
            Case 5 'Version
                If TxtCodLinea.SelStart < 9 Then
                    TxtCodLinea.SelStart = 9
                End If
        End Select
    End If
End Sub

Private Sub TxtCodLinea_GotFocus()
    fEnfoque TxtCodLinea
End Sub

Private Sub TxtCodLinea_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not bCargandoControles Then
        If KeyCode = 46 Or KeyCode = 8 Then
        Select Case nTipoIngreso
            Case 2 'Sub Fondo
                If Len(TxtCodLinea.Text) = 2 Then
                    KeyTmp = TxtCodLinea.Text
                    KeyCode = 0
                End If
            Case 5 'Version
                If Len(TxtCodLinea.Text) = 9 Then
                    KeyTmp = TxtCodLinea.Text
                    KeyCode = 0
                End If
        End Select
        End If
        Select Case nTipoIngreso
            Case 2 'Sub Fondo
                If TxtCodLinea.SelStart < 2 Then
                    TxtCodLinea.SelStart = 2
                    KeyCode = 0
                End If
            Case 5 'Version
                If TxtCodLinea.SelStart < 9 Then
                    TxtCodLinea.SelStart = 9
                    KeyCode = 0
                End If
        End Select
    End If
End Sub

Private Sub TxtCodLinea_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        If Txtdescrip.Visible Then
            Txtdescrip.SetFocus
        Else
            CmbPersFinan.SetFocus
        End If
    End If
    Select Case nTipoIngreso
        Case 2 'Sub Fondo
            If TxtCodLinea.SelStart < 2 Then
                TxtCodLinea.SelStart = 2
                KeyAscii = 0
            End If
        Case 5 'Version
            If TxtCodLinea.SelStart < 9 Then
                TxtCodLinea.SelStart = 9
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub TxtCodLinea_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not bCargandoControles Then
        If KeyCode = 46 Or KeyCode = 8 Then
        Select Case nTipoIngreso
            Case 2 'Sub Fondo
                If Len(TxtCodLinea.Text) < 2 Then
                    TxtCodLinea.Text = KeyTmp
                    TxtCodLinea.SelStart = 2
                End If
            Case 5 'Version
                If Len(TxtCodLinea.Text) < 9 Then
                    TxtCodLinea.Text = KeyTmp
                    TxtCodLinea.SelStart = 9
                End If
        End Select
        End If
        Select Case nTipoIngreso
            Case 2 'Sub Fondo
                If TxtCodLinea.SelStart < 2 Then
                    TxtCodLinea.SelStart = 2
                End If
            Case 5 'Version
                If TxtCodLinea.SelStart < 9 Then
                    TxtCodLinea.SelStart = 9
                End If
        End Select
    End If
End Sub

Private Sub TxtDescrip_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        If CmbMoneda.Enabled Then
            CmbMoneda.SetFocus
        Else
            If CmbPlazo.Enabled Then
                CmbPlazo.SetFocus
            Else
                If CmbProducto.Enabled Then
                    CmbProducto.SetFocus
                Else
                    CmbEstado.SetFocus
                End If
            End If
        End If
    End If
End Sub


Private Sub TxtMontoMax_GotFocus()
    fEnfoque TxtMontoMax
End Sub

Private Sub TxtMontoMax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtMontoMin.SetFocus
    End If
End Sub


Private Sub TxtMontomaxLinea_GotFocus()
    fEnfoque TxtMontomaxLinea
End Sub

Private Sub TxtMontomaxLinea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmbLinea.SetFocus
    End If
End Sub

Private Sub TxtMontoMin_GotFocus()
    fEnfoque TxtMontoMin
End Sub

Private Sub TxtMontoMin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtMontomaxLinea.SetFocus
    End If
End Sub

Private Sub TxtPlazoMax_GotFocus()
    fEnfoque TxtPlazoMax
End Sub

Private Sub TxtPlazoMax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtPlazoMin.SetFocus
    End If
End Sub

Private Sub TxtPlazoMin_GotFocus()
    fEnfoque TxtPlazoMin
End Sub

Private Sub TxtPlazoMin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtMontoMax.SetFocus
    End If
End Sub

Private Sub TxtSaldo_GotFocus()
    fEnfoque TxtSaldo
End Sub

Private Sub TxtSaldo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdTasaNuevo.SetFocus
    End If
End Sub
