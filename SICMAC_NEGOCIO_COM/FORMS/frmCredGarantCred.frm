VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCredGarantCred 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Garantias de Credito"
   ClientHeight    =   7890
   ClientLeft      =   3210
   ClientTop       =   1620
   ClientWidth     =   9120
   Icon            =   "frmCredGarantCred.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6645
      TabIndex        =   35
      Top             =   7335
      Width           =   1170
   End
   Begin VB.CommandButton CmdSolicitud 
      Caption         =   "Soli&citud"
      Height          =   375
      Left            =   4275
      TabIndex        =   31
      Top             =   7335
      Width           =   1170
   End
   Begin VB.CommandButton CmdSugerencia 
      Caption         =   "S&ugerencia"
      Height          =   375
      Left            =   5460
      TabIndex        =   30
      Top             =   7335
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   "Garantias"
      Height          =   3540
      Left            =   120
      TabIndex        =   25
      Top             =   3720
      Width           =   8880
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   1275
         TabIndex        =   33
         Top             =   3120
         Width           =   1155
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   105
         TabIndex        =   32
         Top             =   3120
         Width           =   1155
      End
      Begin SICMACT.FlexEdit FEGarantCred 
         Height          =   1470
         Left            =   90
         TabIndex        =   26
         Top             =   1560
         Width           =   8685
         _extentx        =   15319
         _extenty        =   2593
         cols0           =   11
         highlight       =   1
         allowuserresizing=   3
         encabezadosnombres=   "-Garantia-Gravament-Valor Comercial-Realizacion-Disponible-Titular-Nro Docum-TipoDoc-cNumGarant-EstadoPFijo"
         encabezadosanchos=   "300-3800-1200-1500-1200-1200-3500-1200-0-0-1200"
         font            =   "frmCredGarantCred.frx":000C
         font            =   "frmCredGarantCred.frx":0034
         font            =   "frmCredGarantCred.frx":005C
         font            =   "frmCredGarantCred.frx":0084
         fontfixed       =   "frmCredGarantCred.frx":00AC
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         columnasaeditar =   "X-X-2-X-X-X-X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-R-R-R-R-L-L-L-C-L"
         formatosedit    =   "0-0-2-2-2-2-0-0-0-0-0"
         lbbuscaduplicadotext=   -1  'True
         colwidth0       =   300
         rowheight0      =   300
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   1275
         TabIndex        =   27
         Top             =   3120
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   3120
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton cmdcancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1275
         TabIndex        =   28
         Top             =   3120
         Visible         =   0   'False
         Width           =   1155
      End
      Begin MSDataGridLib.DataGrid DGGarantias 
         Height          =   1245
         Left            =   105
         TabIndex        =   34
         Top             =   240
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   2196
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "cTpoGarDescripcion"
            Caption         =   "Garantia"
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
         BeginProperty Column01 
            DataField       =   "nTasacion"
            Caption         =   "Valor  Comercial"
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
         BeginProperty Column02 
            DataField       =   "nRealizacion"
            Caption         =   "Realizacion"
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
         BeginProperty Column03 
            DataField       =   "nDisponible"
            Caption         =   "Disponible"
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
            DataField       =   "cMonedaDesc"
            Caption         =   "Moneda"
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
         BeginProperty Column05 
            DataField       =   "cNroDoc"
            Caption         =   "Nro Documento"
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
         BeginProperty Column06 
            DataField       =   "cEstadoGar"
            Caption         =   "Estado"
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
         BeginProperty Column07 
            DataField       =   "cPersCodEmisor"
            Caption         =   "Emisor"
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
         BeginProperty Column08 
            DataField       =   "cNumGarant"
            Caption         =   "cNumGarant"
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
         BeginProperty Column09 
            DataField       =   "PFestado"
            Caption         =   "Estado_PF"
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
            ScrollBars      =   3
            BeginProperty Column00 
               ColumnWidth     =   3555.213
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   1484.787
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column09 
            EndProperty
         EndProperty
      End
   End
   Begin SICMACT.ActXCodCta ActxCuenta 
      Height          =   420
      Left            =   75
      TabIndex        =   24
      Top             =   60
      Width           =   3585
      _extentx        =   6324
      _extenty        =   741
      texto           =   "Credito :"
      enabledcmac     =   -1  'True
      enabledcta      =   -1  'True
      enabledprod     =   -1  'True
      enabledage      =   -1  'True
      cmac            =   "112"
   End
   Begin VB.Frame FraCliente 
      Caption         =   "Datos del Cliente"
      ForeColor       =   &H00000000&
      Height          =   915
      Left            =   75
      TabIndex        =   17
      Top             =   525
      Width           =   8910
      Begin VB.Label Label1 
         Caption         =   "Código:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   225
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   210
         TabIndex        =   22
         Top             =   495
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Doc. Identidad:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5325
         TabIndex        =   21
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label LblCodCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1530
         TabIndex        =   20
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label LblNomCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1530
         TabIndex        =   19
         Top             =   525
         Width           =   6615
      End
      Begin VB.Label LblDIdent 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6555
         TabIndex        =   18
         Top             =   225
         Width           =   1575
      End
   End
   Begin VB.Frame FraCredito 
      Caption         =   "Datos del Crédito"
      ForeColor       =   &H00000000&
      Height          =   1650
      Left            =   75
      TabIndex        =   8
      Top             =   1485
      Width           =   8940
      Begin VB.Label LblTipoCred 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1500
         TabIndex        =   37
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label10 
         Caption         =   "Credito:"
         Height          =   270
         Left            =   195
         TabIndex        =   36
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label5 
         Caption         =   "Producto:"
         Height          =   270
         Left            =   195
         TabIndex        =   16
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label6 
         Caption         =   "Monto Solicitado:"
         Height          =   255
         Left            =   210
         TabIndex        =   15
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label Label7 
         Caption         =   "Moneda:"
         Height          =   255
         Left            =   5805
         TabIndex        =   14
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Destino:"
         Height          =   255
         Left            =   5790
         TabIndex        =   13
         Top             =   720
         Width           =   705
      End
      Begin VB.Label lblTipoProd 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1500
         TabIndex        =   12
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label LblMontoS 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1500
         TabIndex        =   11
         Top             =   1080
         Width           =   1590
      End
      Begin VB.Label LblDestinoC 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6795
         TabIndex        =   10
         Top             =   720
         Width           =   1995
      End
      Begin VB.Label lblMoneda 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6795
         TabIndex        =   9
         Top             =   1080
         Width           =   1980
      End
   End
   Begin VB.Frame FraGarante 
      Caption         =   "Persona Crédito"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   75
      TabIndex        =   3
      Top             =   3120
      Width           =   8955
      Begin VB.Label Label4 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Relación:"
         Height          =   255
         Left            =   5820
         TabIndex        =   6
         Top             =   225
         Width           =   735
      End
      Begin VB.Label LblNPersona 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1500
         TabIndex        =   5
         Top             =   240
         Width           =   4155
      End
      Begin VB.Label LblRelacion 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6825
         TabIndex        =   4
         Top             =   225
         Width           =   1935
      End
   End
   Begin VB.ComboBox CmbPersonas 
      Height          =   315
      Left            =   4740
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   4245
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7830
      TabIndex        =   1
      Top             =   7335
      Width           =   1170
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "E&xaminar"
      Height          =   375
      Left            =   3780
      TabIndex        =   0
      Top             =   90
      Width           =   900
   End
End
Attribute VB_Name = "frmCredGarantCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmCredGarantCred
'***     Descripcion:       Permite el Grabar el Monte de la Garantia del Credito
'***     Creado por:        NSSE
'***     Maquina:           07SIST_08
'***     Fecha-Tiempo:         11/06/2001 10:14:45 AM
'***     Ultima Modificacion: Creacion de la Opcion
'*****************************************************************************************

Option Explicit

Enum TIniciofrmCredGarantCred
    PorSolicitud = 1
    PorMenu = 2
End Enum

Dim RGarantias As ADODB.Recordset
Dim FilaAct As Integer
Dim nCmdEjecutar As Integer
Dim vInicio As TIniciofrmCredGarantCred
Dim vnGravadoAnt As Double

Dim fnProducto As Integer  ' Identifica si se grava garantias de un Credito (0) , CartaFianza(1)


Dim fnTipoCamb As Double    'solo se recupera una vez
Dim fnMontoGravado As Double 'solo se recupera una vez
Dim objPista As COMManejador.Pista

Private bLeasing As Boolean
Public Sub Inicioleasing(ByVal pInicio As TIniciofrmCredGarantCred, Optional ByVal psCtaCod As String = "", _
                Optional ByVal pnProducto As Integer = "0")
    
    vInicio = pInicio
    bLeasing = True
    fnProducto = pnProducto
    If vInicio = PorMenu Then
        CmdSolicitud.Enabled = False
        CmdSugerencia.Enabled = False
    Else
        CmdSolicitud.Enabled = True
        CmdSugerencia.Enabled = True
        ActxCuenta.NroCuenta = psCtaCod
        Call ActxCuenta_KeyPress(13)
    End If
    'ALPA 20120416
    If Mid(psCtaCod, 6, 3) = "515" Or Mid(psCtaCod, 6, 3) = "516" Then
        FraCredito.Caption = "Datos del Arrendamiento Financiero"
        Label10.Caption = "Operación"
        Me.Caption = "Operación Arrendamiento Financiero"
        FraGarante.Caption = "Persona AF"
        ActxCuenta.texto = "Operación"
    End If

    Me.Show 1
End Sub

Public Sub Inicio(ByVal pInicio As TIniciofrmCredGarantCred, Optional ByVal psCtaCod As String = "", _
                Optional ByVal pnProducto As Integer = "0")
    
    vInicio = pInicio
    bLeasing = False
    fnProducto = pnProducto
    If vInicio = PorMenu Then
        CmdSolicitud.Enabled = False
        CmdSugerencia.Enabled = False
    Else
        CmdSolicitud.Enabled = True
        CmdSugerencia.Enabled = True
        ActxCuenta.NroCuenta = psCtaCod
        Call ActxCuenta_KeyPress(13)
    End If
    Me.Show 1
End Sub

Private Sub LimpiarPantalla()
    ActxCuenta.NroCuenta = ""
    ActxCuenta.CMAC = gsCodCMAC
    ActxCuenta.Age = gsCodAge
    CmbPersonas.Clear
    CmbPersonas.ListIndex = -1
    LblCodCliente.Caption = ""
    LblDIdent.Caption = ""
    LblNomCliente.Caption = ""
    lblTipoProd.Caption = ""
    LblDestinoC.Caption = ""
    LblMontoS.Caption = ""
    lblmoneda.Caption = ""
    LblNPersona.Caption = ""
    LblRelacion.Caption = ""
    Set DGGarantias.DataSource = Nothing
    DGGarantias.Refresh
    Set RGarantias = Nothing
    Call LimpiaFlex(FEGarantCred)
    CmdNuevo.Enabled = False
    CmdEditar.Enabled = False
    cmdEliminar.Enabled = False
    FEGarantCred.row = 1
    FEGarantCred.BackColorRow (vbWhite)
End Sub
'*****************************************************************************************
'***     Rutina:           CargaPersonasGarantes
'***     Descripcion:       Carga en el Combo CmbPersonas los Titulares y Garantes
'***     Modificado por:        NSSE
'***     Fecha-Tiempo:         11/06/2001 10:39:39 AM
'*****************************************************************************************
Private Sub CargaPersonasGarantes(ByVal psCtaCod As String, _
                                    ByVal pRs As ADODB.Recordset)

'Dim oCredito As COMDCredito.DCOMCredito
'Dim R As ADODB.Recordset
    
    On Error GoTo ErrorCargaPersonasGarantes
    CmbPersonas.Clear
'    Set oCredito = New COMDCredito.DCOMCredito
'    Set R = oCredito.RecuperaRelacPers(psCtaCod)
    Do While Not pRs.EOF
        If pRs!nConsValor = gColRelPersTitular Or pRs!nConsValor = gColRelPersGarante Or pRs!nConsValor = gColRelPersCodeudor Or pRs!nConsValor = gColRelPersConyugue Or pRs!nConsValor = gColRelPersRepresCodeudor Then
            CmbPersonas.AddItem PstaNombre(pRs!cPersNombre, True) & Space(150 - Len(pRs!cPersNombre)) & pRs!cPersCod & Space(50 - Len(pRs!cPersCod)) & Trim(pRs!cConsDescripcion)
        End If
        pRs.MoveNext
    Loop
    'R.Close
    
    If CmbPersonas.ListCount > 0 Then
        CmbPersonas.ListIndex = 0
    Else
        CmbPersonas.ListIndex = -1
    End If
    
    Exit Sub

ErrorCargaPersonasGarantes:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
'Dim oCredito As COMDCredito.DCOMCredito
Dim oGarantia As COMDCredito.DCOMGarantia
'Dim loCartaF As COMDCartaFianza.DCOMCartaFianza
'Dim R As ADODB.Recordset

Dim rsDatosGarCred As ADODB.Recordset
Dim rsPersGarant As ADODB.Recordset
'Dim rsGarantias As ADODB.Recordset ' Ya hay una Variable Global
Dim rsGarCred As ADODB.Recordset
Dim rsGarCF As ADODB.Recordset


On Error GoTo ErrorCargaDatos
CargaDatos = True

Set oGarantia = New COMDCredito.DCOMGarantia
Call oGarantia.CargarDatosGarantCredito(ActxCuenta.NroCuenta, Trim(Mid(CmbPersonas.Text, 150, 50)), gdFecSis, rsDatosGarCred, rsPersGarant, RGarantias, rsGarCred, rsGarCF, fnTipoCamb, fnMontoGravado)
Set oGarantia = Nothing

'Carga de Datos del Credito
If fnProducto = 0 Then ' Creditos
    'Set oCredito = New COMDCredito.DCOMCredito
    'Set R = oCredito.RecuperaDatosGarantiaCred(ActxCuenta.NroCuenta)
    If Not rsDatosGarCred.BOF And Not rsDatosGarCred.EOF Then
        LblCodCliente.Caption = rsDatosGarCred!cPersCod
        LblDIdent.Caption = IIf(IsNull(rsDatosGarCred!DNI), "", rsDatosGarCred!DNI)
        LblNomCliente.Caption = PstaNombre(rsDatosGarCred!cPersNombre, True)
        lblTipoProd.Caption = Trim(rsDatosGarCred!cSTipoProdDescrip)
        lblTipoCred.Caption = Trim(rsDatosGarCred!cTipoProdDescrip)
        LblDestinoC.Caption = rsDatosGarCred!cDestinoDescripcion
        LblMontoS.Caption = Format(rsDatosGarCred!nMonto, "#0.00")
        lblmoneda.Caption = rsDatosGarCred!cMonedaDesc
    Else
        'R.Close
        'Set R = Nothing
        'Set oCredito = Nothing
        CargaDatos = False
        Exit Function
    End If
    'R.Close
    'Set R = Nothing
    
    Call CargaPersonasGarantes(ActxCuenta.NroCuenta, rsPersGarant)
    'CmbPersonas.ListIndex = IndiceListaCombo(CmbPersonas, Trim(Str(gColRelPersTitular)))
    CmbPersonas.ListIndex = 0
    
    LblNPersona.Caption = Mid(CmbPersonas.Text, 1, 150)
    LblRelacion.Caption = Trim(Right(CmbPersonas.Text, 15))
    
    'Set oGarantia = New COMDCredito.DCOMGarantia
    'Set RGarantias = oGarantia.RecuperaGarantiasPersona(Trim(Mid(CmbPersonas.Text, 150, 50)))
    Set DGGarantias.DataSource = RGarantias
    'Set oGarantia = Nothing
    DGGarantias.Refresh
    Call LimpiaFlex(FEGarantCred)
    'Set R = oCredito.RecuperaGarantiasCredito(ActxCuenta.NroCuenta)
    Do While Not rsGarCred.EOF
        FEGarantCred.AdicionaFila
        FEGarantCred.RowHeight(rsGarCred.Bookmark) = 280
        FEGarantCred.TextMatrix(rsGarCred.Bookmark, 1) = rsGarCred!cTpoGarDescripcion
        FEGarantCred.TextMatrix(rsGarCred.Bookmark, 2) = Format(rsGarCred!nGravado, "#,#0.00")
        FEGarantCred.TextMatrix(rsGarCred.Bookmark, 3) = Format(rsGarCred!nTasacion, "#,#0.00")
        FEGarantCred.TextMatrix(rsGarCred.Bookmark, 4) = Format(rsGarCred!nRealizacion, "#,#0.00")
        FEGarantCred.TextMatrix(rsGarCred.Bookmark, 5) = Format(rsGarCred!nPorGravar, "#,#0.00")
        FEGarantCred.TextMatrix(rsGarCred.Bookmark, 6) = Trim(rsGarCred!cPersNombre)
        FEGarantCred.TextMatrix(rsGarCred.Bookmark, 7) = Trim(rsGarCred!cNroDoc)
        FEGarantCred.TextMatrix(rsGarCred.Bookmark, 8) = Trim(rsGarCred!cTpoDoc)
        FEGarantCred.TextMatrix(rsGarCred.Bookmark, 9) = Trim(rsGarCred!cNumGarant)
        
        rsGarCred.MoveNext
    Loop
    'R.Close
    'Set R = Nothing
    'Set oCredito = Nothing
    Exit Function
ElseIf fnProducto = 1 Then ' Carta Fianza
    'Set loCartaF = New COMDCartaFianza.DCOMCartaFianza
    'Set R = loCartaF.RecuperaDatosGarantiaCF(ActxCuenta.NroCuenta)
    If Not rsGarCF.BOF And Not rsGarCF.EOF Then
        LblCodCliente.Caption = rsGarCF!cPersCod
        LblDIdent.Caption = IIf(IsNull(rsGarCF!DNI), "", rsGarCF!DNI)
        LblNomCliente.Caption = PstaNombre(rsGarCF!cPersNombre, True)
        lblTipoProd.Caption = Trim(rsGarCF!cTipoCredDescrip)
        LblDestinoC.Caption = rsGarCF!cDestinoDescripcion
        LblMontoS.Caption = Format(rsGarCF!nMonto, "#0.00")
        lblmoneda.Caption = rsGarCF!cMonedaDesc
    Else
        'R.Close
        'Set R = Nothing
        'Set loCartaF = Nothing
        CargaDatos = False
        Exit Function
    End If
    'R.Close
    'Set R = Nothing
    
    Call CargaPersonasGarantes(ActxCuenta.NroCuenta, rsPersGarant)
    'CmbPersonas.ListIndex = IndiceListaCombo(CmbPersonas, Trim(Str(gColRelPersTitular)))
    CmbPersonas.ListIndex = 0
    
    LblNPersona.Caption = Mid(CmbPersonas.Text, 1, 150)
    LblRelacion.Caption = Trim(Right(CmbPersonas.Text, 15))
    
    'Set oGarantia = New COMDCredito.DCOMGarantia
    'Set RGarantias = oGarantia.RecuperaGarantiasPersona(Trim(Mid(CmbPersonas.Text, 150, 50)))
    Set DGGarantias.DataSource = RGarantias
    'Set oGarantia = Nothing
    DGGarantias.Refresh
    Call LimpiaFlex(FEGarantCred)
    'Set oCredito = New COMDCredito.DCOMCredito
    'Set R = oCredito.RecuperaGarantiasCredito(ActxCuenta.NroCuenta)
    Do While Not rsGarCred.EOF
        FEGarantCred.AdicionaFila
        FEGarantCred.RowHeight(rsGarCred.Bookmark) = 280
        FEGarantCred.TextMatrix(rsGarCred.Bookmark, 1) = rsGarCred!cTpoGarDescripcion
        FEGarantCred.TextMatrix(rsGarCred.Bookmark, 2) = Format(rsGarCred!nGravado, "#,#0.00")
        FEGarantCred.TextMatrix(rsGarCred.Bookmark, 3) = Format(rsGarCred!nTasacion, "#,#0.00")
        FEGarantCred.TextMatrix(rsGarCred.Bookmark, 4) = Format(rsGarCred!nRealizacion, "#,#0.00")
        FEGarantCred.TextMatrix(rsGarCred.Bookmark, 5) = Format(rsGarCred!nPorGravar, "#,#0.00")
        FEGarantCred.TextMatrix(rsGarCred.Bookmark, 6) = Trim(rsGarCred!cPersNombre)
        FEGarantCred.TextMatrix(rsGarCred.Bookmark, 7) = Trim(rsGarCred!cNroDoc)
        FEGarantCred.TextMatrix(rsGarCred.Bookmark, 8) = Trim(rsGarCred!cTpoDoc)
        FEGarantCred.TextMatrix(rsGarCred.Bookmark, 9) = Trim(rsGarCred!cNumGarant)
        
        rsGarCred.MoveNext
    Loop
    'R.Close
    'Set R = Nothing
    'Set oCredito = Nothing
    Exit Function
End If
    
ErrorCargaDatos:
        MsgBox err.Description, vbCritical, "Aviso"
End Function

Private Sub HabilitaActualizarGarantia(ByVal pbHabilita As Boolean)
    'ActxCuenta.Enabled = Not pbHabilita
    DGGarantias.Enabled = Not pbHabilita
    CmdLimpiar.Enabled = Not pbHabilita
    cmdbuscar.Enabled = Not pbHabilita
    CmbPersonas.Enabled = Not pbHabilita
    CmdNuevo.Visible = Not pbHabilita
    'CmdEditar.Visible = Not pbHabilita
    cmdEliminar.Visible = Not pbHabilita
    cmdgrabar.Visible = pbHabilita
    cmdcancelar.Visible = pbHabilita
    CmdSolicitud.Enabled = Not pbHabilita
    CmdSugerencia.Enabled = Not pbHabilita
    CmdSalir.Enabled = Not pbHabilita
    FEGarantCred.lbEditarFlex = pbHabilita
    FEGarantCred.col = 2
    If pbHabilita Then
        Call FEGarantCred.BackColorRow(vbYellow, True)
    Else
        Call FEGarantCred.BackColorRow(vbWhite)
    End If
    FEGarantCred.SetFocus
    If vInicio = PorMenu Then
        CmdSolicitud.Enabled = False
        CmdSugerencia.Enabled = False
    End If
End Sub

Private Sub ActxCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not CargaDatos(ActxCuenta.NroCuenta) Then
            MsgBox "Credito No Existe", vbInformation, "Aviso"
            Call CmdLimpiar_Click
        Else
            ActxCuenta.Enabled = False
            CmdNuevo.Enabled = True
            CmdEditar.Enabled = True
            cmdEliminar.Enabled = True
        End If
    End If
End Sub


Private Sub CmbPersonas_Click()
Dim oGarantia As COMDCredito.DCOMGarantia
    
    On Error GoTo ErrorCmbPersonas_Click
    Set oGarantia = New COMDCredito.DCOMGarantia
    If Not Me.Visible Then
        Exit Sub
    End If
    Set RGarantias = oGarantia.RecuperaGarantiasPersona(Trim(Mid(CmbPersonas.Text, 150, 50)))
    Set DGGarantias.DataSource = RGarantias
    Set oGarantia = Nothing
    DGGarantias.Refresh
    
    If RGarantias.RecordCount <= 0 Then
        MsgBox "Persona No Tiene Garantias", vbInformation, "Aviso"
        CmdNuevo.Enabled = False
        CmdEditar.Enabled = False
        cmdEliminar.Enabled = False
    Else
        CmdNuevo.Enabled = True
        CmdEditar.Enabled = True
        cmdEliminar.Enabled = True
    End If
    
    Exit Sub

ErrorCmbPersonas_Click:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdBuscar_Click()

On Error GoTo ErrorCmdBuscar_Click
    Screen.MousePointer = 11
    If fnProducto = 0 Then
        ActxCuenta.NroCuenta = frmCredPersEstado.Inicio(Array(gColocEstSolic, gColocEstSug), "Creditos Solicitados", , , , gsCodAge)
        If Len(ActxCuenta.NroCuenta) = 18 Then
            Call ActxCuenta_KeyPress(13)
        Else
            Call CmdLimpiar_Click
        End If
        Exit Sub
    ElseIf fnProducto = 1 Then
        ActxCuenta.NroCuenta = frmCFPersEstado.Inicio(Array(gColocEstSolic, gColocEstSug), "Cartas Fianza")
        If Len(ActxCuenta.NroCuenta) = 18 Then
            Call ActxCuenta_KeyPress(13)
        Else
            Call CmdLimpiar_Click
        End If
        Exit Sub
    End If

ErrorCmdBuscar_Click:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdCancelar_Click()
    Call HabilitaActualizarGarantia(False)
    If nCmdEjecutar = 1 Then
        Call FEGarantCred.EliminaFila(FilaAct)
        'Call LimpiarPantalla
        Call HabilitaActualizarGarantia(False)
        cmdbuscar.SetFocus
    Else
        FEGarantCred.TextMatrix(FilaAct, 2) = Format(vnGravadoAnt, "#0.00")
        CmdEditar.SetFocus
    End If
    nCmdEjecutar = -1
    vnGravadoAnt = 0
    FilaAct = 0
End Sub

Private Sub CmdEditar_Click()
Dim oCredito As COMDCredito.DCOMCredito

    If Trim(FEGarantCred.TextMatrix(1, 0)) = "" Then
        MsgBox "No existen Garantias para Editar", vbInformation, "Aviso"
        Exit Sub
    End If

    Set oCredito = New COMDCredito.DCOMCredito
    If oCredito.GarantiaPerteneceACreditoAprobado(ActxCuenta.NroCuenta, FEGarantCred.TextMatrix(FEGarantCred.row, 7)) Then
        MsgBox "Garantia ya esta siendo Usada por un Credito Vigente", vbInformation, "Aviso"
        Exit Sub
        Set oCredito = Nothing
    End If
    Set oCredito = Nothing

    nCmdEjecutar = 2
    FilaAct = FEGarantCred.row
    Call HabilitaActualizarGarantia(True)
    vnGravadoAnt = CDbl(FEGarantCred.TextMatrix(FilaAct, 2))
    FEGarantCred.lbEditarFlex = True
End Sub

Private Sub cmdEliminar_Click()
Dim oCredito As COMDCredito.DCOMCredito
Dim oGarantia As COMDCredito.DCOMGarantia
    On Error GoTo ErrorcmdEliminar_Click
        
    If Trim(FEGarantCred.TextMatrix(FEGarantCred.row, 1)) = "" Then
        MsgBox "Seleccione una Garantia del Credito", vbInformation, "Aviso"
        Exit Sub
    End If
    Set oCredito = New COMDCredito.DCOMCredito
    If oCredito.GarantiaPerteneceACreditoAprobado(ActxCuenta.NroCuenta, FEGarantCred.TextMatrix(FEGarantCred.row, 9)) Then
        MsgBox "Garantia ya esta siendo Usada por un Credito Vigente", vbInformation, "Aviso"
        Exit Sub
        Set oCredito = Nothing
    End If
    
    Set oCredito = Nothing
    
    If Trim(FEGarantCred.TextMatrix(1, 0)) = "" Then
        MsgBox "No existen Garantias para Editar", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("Se Va ha Eliminar el Registro, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set oCredito = New COMDCredito.DCOMCredito
        
        'MAVM 07072010 ***
        If FEGarantCred.TextMatrix(FEGarantCred.row, 8) = "145" Then
            Dim loGrabarResc As COMNColoCPig.NCOMColPContrato
            Set loGrabarResc = New COMNColoCPig.NCOMColPContrato
            Call loGrabarResc.nRescataJoyaGarantia(Trim(FEGarantCred.TextMatrix(FEGarantCred.row, 7)), 0)
        End If
        '***
        
        Call oCredito.EliminarGarantia(ActxCuenta.NroCuenta, FEGarantCred.TextMatrix(FEGarantCred.row, 9), CDbl(FEGarantCred.TextMatrix(FEGarantCred.row, 2)))
        Set oCredito = Nothing
        Call FEGarantCred.EliminaFila(FEGarantCred.row)
        'Refresca Garantias
        Set oGarantia = New COMDCredito.DCOMGarantia
        Set RGarantias = oGarantia.RecuperaGarantiasPersona(Trim(Mid(CmbPersonas.Text, 150, 50)))
        Set DGGarantias.DataSource = RGarantias
        Set oGarantia = Nothing
        
    End If
    
    Exit Sub

ErrorcmdEliminar_Click:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdGrabar_Click()

'Dim oCredito As COMDCredito.DCOMCredito
'Dim oNegCredito As COMNCredito.NCOMCredito

Dim oGarantia As COMDCredito.DCOMGarantia
Dim oNegGarantia As COMNCredito.NCOMGarantia
Dim sValidar As String
Dim bExisteGarant As Boolean
'Dim oTipoCam As COMDConstSistema.NCOMTipoCambio
'Dim nTipoCamb As Double
'Dim nMontoGravado As Double

    On Error GoTo ErrorCmdGrabar_Click
'    nMontoGravado = 0

 'By Capi 12082008 para que valide estado de plazo fijo
        If Trim(FEGarantCred.TextMatrix(FEGarantCred.row, 10)) = "CANCELADA" Then
            MsgBox "Garantia Plazo Fijo no Validado: Cuenta Cancelada", vbInformation, "Aviso"
            Exit Sub
        End If
 '

    
    If CDbl(FEGarantCred.TextMatrix(FEGarantCred.row, 2)) <= 0 Then
        MsgBox "El Monto a Grabar debe ser mayor a Cero", vbInformation, "Aviso"
        FEGarantCred.col = 2
        Exit Sub
    End If
    
'    Set oTipoCam = New COMDConstSistema.NCOMTipoCambio
'    nTipoCamb = oTipoCam.EmiteTipoCambio(gdFecSis, TCFijoMes)
'    Set oTipoCam = Nothing
    
    'Verifica el Monto Gravado
'    Set oCredito = New COMDCredito.DCOMCredito
'    nMontoGravado = oCredito.RecuperaMontoGarantiaCredito(ActxCuenta.NroCuenta, gdFecSis)
    'Set oGarantia = New DGarantia
    'Set R = oGarantia.RecuperaGarantiaCreditoDatos(ActxCuenta.NroCuenta)
    'Do While Not R.EOF
    '    If CInt(Mid(ActxCuenta.NroCuenta, 9, 1)) <> R!nMoneda Then
    '        'De Dolares a Soles
    '        If CInt(Mid(ActxCuenta.NroCuenta, 9, 1)) = 1 Then
    '            nMontoGravado = nMontoGravado + CDbl(Format(R!nGravado * nTipoCamb, "#0.00"))
    '
    '        Else 'De Soles a Dolares
    '            nMontoGravado = nMontoGravado + CDbl(Format(R!nGravado / nTipoCamb, "#0.00"))
    '
    '        End If
    '    R.MoveNext
    'Loop
    'R.Close
    
    If CInt(Mid(ActxCuenta.NroCuenta, 9, 1)) <> RGarantias!nmoneda Then
        'De Dolares a Soles
        If CInt(Mid(ActxCuenta.NroCuenta, 9, 1)) = 1 Then
            If CDbl(Format(CDbl(Me.FEGarantCred.TextMatrix(FEGarantCred.row, 2)) * fnTipoCamb, "#0.00")) < (CDbl(Me.LblMontoS.Caption) - fnMontoGravado) Then
                If MsgBox("Monto de Garantia a Soles es : " & Format(CDbl(Me.FEGarantCred.TextMatrix(FEGarantCred.row, 2)) * fnTipoCamb, "#0.00") & Chr(10) & " y no Cubre el Prestamo, Desea Continuar ? ", vbInformation + vbYesNo, "Aviso") = vbNo Then
                    Exit Sub
                End If
            End If
        Else 'De Soles a Dolares
            If CDbl(Format(CDbl(Me.FEGarantCred.TextMatrix(FEGarantCred.row, 2)) / fnTipoCamb, "#0.00")) < (CDbl(Me.LblMontoS.Caption) - fnMontoGravado) Then
                If MsgBox("Monto de Garantia a Dolares es : " & Format(CDbl(Me.FEGarantCred.TextMatrix(FEGarantCred.row, 2)) / fnTipoCamb, "#0.00") & Chr(10) & " y no Cubre el Prestamo, Desea Continuar ? ", vbInformation + vbYesNo, "Aviso") = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    'ARCV 14-08-2006
    Else
        If CDbl(Format(CDbl(Me.FEGarantCred.TextMatrix(FEGarantCred.row, 2)), "#0.00")) < (CDbl(Me.LblMontoS.Caption) - fnMontoGravado) Then
            If MsgBox("El Monto de la Garantia no Cubre el Prestamo, Desea Continuar ? ", vbInformation + vbYesNo, "Aviso") = vbNo Then
                Exit Sub
            End If
        End If
    '---------------
    End If
    
    If MsgBox("Se va a Grabar la Garantia, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    Set oNegGarantia = New COMNCredito.NCOMGarantia
    
    Call oNegGarantia.GrabarDatos(nCmdEjecutar, sValidar, bExisteGarant, CDbl(FEGarantCred.TextMatrix(FilaAct, 2)), RGarantias!nDisponible, ActxCuenta.NroCuenta, _
                            RGarantias!cNumGarant, RGarantias!nmoneda, CDbl(FEGarantCred.TextMatrix(FilaAct, 2)), vnGravadoAnt)
    
    ''*** PEAC 20090126
    objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, Trim(RGarantias!cNumGarant), ActxCuenta.NroCuenta, gCodigoCuenta
    'WIOR 20150619 AGREGO Trim(RGarantias!cNumGarant)
    
        
    Set oNegGarantia = Nothing
    'Set oNegCredito = New COMNCredito.NCOMCredito
    'If oNegCredito.ValidaDatosGarantiaCredito(CDbl(FEGarantCred.TextMatrix(FilaAct, 2)), RGarantias!nDisponible) <> "" Then
    If sValidar <> "" And bExisteGarant = False Then
        MsgBox sValidar, vbInformation, "Aviso"
        'Set oNegCredito = Nothing
        FEGarantCred.col = 2
        FEGarantCred.SetFocus
        Exit Sub
    End If
    'Set oNegCredito = Nothing
    'Set oCredito = New COMDCredito.DCOMCredito
    'If nCmdEjecutar = 1 Then
    '    If Not oCredito.ExisteGarantia(ActxCuenta.NroCuenta, RGarantias!cNumGarant) Then
    '        Call oCredito.NuevaGarantia(ActxCuenta.NroCuenta, RGarantias!cNumGarant, RGarantias!nMoneda, CDbl(FEGarantCred.TextMatrix(FilaAct, 2)))
    '       FEGarantCred.TextMatrix(FEGarantCred.Row, 9) = RGarantias!cNumGarant
    '    Else
    '        MsgBox "Garantia Ya Existe", vbInformation, "Aviso"
    '        Set oCredito = Nothing
    '        Call cmdcancelar_Click
    '        Exit Sub
    '    End If
    'Else
    '    Call oCredito.ActualizaGarantias(ActxCuenta.NroCuenta, RGarantias!cNumGarant, RGarantias!nMoneda, CDbl(FEGarantCred.TextMatrix(FilaAct, 2)), vnGravadoAnt)
    'End If
    'Set oCredito = Nothing
    
    If nCmdEjecutar = 1 Then
        If Not bExisteGarant Then
            FEGarantCred.TextMatrix(FEGarantCred.row, 9) = RGarantias!cNumGarant
        Else
            MsgBox "Garantia Ya Existe", vbInformation, "Aviso"
            'Set oCredito = Nothing
            Call cmdCancelar_Click
            Exit Sub
        End If
    End If
    
    'MAVM 07072010 ***
    If RGarantias!cTpoDoc = "145" Then
        Dim loGrabarResc As COMNColoCPig.NCOMColPContrato
        Set loGrabarResc = New COMNColoCPig.NCOMColPContrato
        Call loGrabarResc.nRescataJoyaGarantia(Trim(FEGarantCred.TextMatrix(FilaAct, 7)), 1)
    End If
    '***

    'Refresca Garantias
    Set oGarantia = New COMDCredito.DCOMGarantia
    Set RGarantias = oGarantia.RecuperaGarantiasPersona(Trim(Mid(CmbPersonas.Text, 150, 50)))
    Set DGGarantias.DataSource = RGarantias
    Set oGarantia = Nothing
    
    Call HabilitaActualizarGarantia(False)
    nCmdEjecutar = -1
    vnGravadoAnt = 0
    'CmdEditar.SetFocus
    
    Exit Sub
ErrorCmdGrabar_Click:
    MsgBox err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub CmdLimpiar_Click()
    Call LimpiarPantalla
    ActxCuenta.Enabled = True
End Sub

Private Sub cmdNuevo_Click()
 Dim filtroCFplazoFijoHipoteca As Boolean
 filtroCFplazoFijoHipoteca = False
 
    If RGarantias.RecordCount <= 0 Then
        MsgBox "No existen Garantias para Gravar", vbInformation, "Aviso"
        Exit Sub
    End If
  DGGarantias.ColContaining (10)
  
    
    'MADM 20100824 - SOLO HIPOTECAS Y DPF
    If fnProducto = 1 Then
        If (CInt(RGarantias!cTpoDoc) = 17) Or (CInt(RGarantias!cTpoDoc) = 101 And (RGarantias!nTpoGarantia) = 1) Then
             filtroCFplazoFijoHipoteca = True
        End If
    End If
    'end madm
    
    If Not (RGarantias!cTpoDoc = "145" And fnProducto = 1) Then
        'MADM 20100824
        If (filtroCFplazoFijoHipoteca = True) Or (fnProducto <> 1) Then
            FEGarantCred.AdicionaFila
            FEGarantCred.TextMatrix(FEGarantCred.Rows - 1, 1) = RGarantias!cTpoGarDescripcion
            FEGarantCred.TextMatrix(FEGarantCred.Rows - 1, 2) = "0.00"
            FEGarantCred.TextMatrix(FEGarantCred.Rows - 1, 3) = Format(RGarantias!nTasacion, "#0.00")
            FEGarantCred.TextMatrix(FEGarantCred.Rows - 1, 4) = Format(RGarantias!nRealizacion, "#0.00")
            FEGarantCred.TextMatrix(FEGarantCred.Rows - 1, 5) = Format(RGarantias!nDisponible, "#0.00")
            FEGarantCred.TextMatrix(FEGarantCred.Rows - 1, 6) = Trim(Mid(CmbPersonas.Text, 1, 150))
            FEGarantCred.TextMatrix(FEGarantCred.Rows - 1, 7) = RGarantias!cNroDoc
            FEGarantCred.TextMatrix(FEGarantCred.Rows - 1, 8) = RGarantias!cTpoDoc
            FEGarantCred.TextMatrix(FEGarantCred.Rows - 1, 9) = RGarantias!cNumGarant 'ARCV 11-07-2006
            'By Capi 12082008
            FEGarantCred.TextMatrix(FEGarantCred.Rows - 1, 10) = RGarantias!PFEstado
            '
            FEGarantCred.col = 2
            nCmdEjecutar = 1
            FilaAct = FEGarantCred.Rows - 1
            FEGarantCred.row = FilaAct
            Call HabilitaActualizarGarantia(True)
            FEGarantCred.SetFocus
        Else
            MsgBox "Para Cartas Fianzas solo Puede Gravar garantias como depósitos a plazo fijo o hipotecas"
        End If
    Else
        MsgBox "No se permite enlazar con el Tipo de Credito", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
    Set frmCredSolicitud = Nothing
End Sub

Private Sub CmdSolicitud_Click()
    Unload Me
End Sub

Private Sub CmdSugerencia_Click()
    If fnProducto = 0 Then
        'ALPA 20120416
        'Call frmCredSugerencia.InicioCargaDatos(ActxCuenta.NroCuenta, bLeasing)
        If gnAgenciaCredEval = 0 Then 'JUEZ 20121219
            Call frmCredSugerencia.InicioCargaDatos(ActxCuenta.NroCuenta, bLeasing, True)
        Else
            Call frmCredSugerencia_NEW.InicioCargaDatos(ActxCuenta.NroCuenta, bLeasing, True)
        End If
        '***********************
    Else
        frmCFSugerencia.inicia (ActxCuenta.NroCuenta)
    End If
End Sub

Private Sub FEGarantCred_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 2 Then
        If Trim(FEGarantCred.TextMatrix(pnRow, pnCol)) = "" Then
            FEGarantCred.TextMatrix(pnRow, pnCol) = "0.00"
        End If
        If CDbl(FEGarantCred.TextMatrix(pnRow, pnCol)) < 0 Then
            FEGarantCred.TextMatrix(pnRow, pnCol) = "0.00"
            MsgBox "No se puede Ingresar Valores Negativos", vbInformation, "Aviso"
            FEGarantCred.row = pnRow
            FEGarantCred.col = pnCol - 1
            Exit Sub
        End If
        
        'ARCV 14-08-2006
        If CDbl(Format(CDbl(Me.FEGarantCred.TextMatrix(pnRow, 2)), "#0.00")) > CDbl(Format(CDbl(Me.FEGarantCred.TextMatrix(pnRow, 5)), "#0.00")) And (CDbl(Me.FEGarantCred.TextMatrix(pnRow, 2)) > 0) Then
            MsgBox "El monto ingresado no puede ser mayor al disponible", vbInformation, "Mensaje"
            FEGarantCred.TextMatrix(pnRow, pnCol) = "0.00"
            Exit Sub
        End If
        
        '*** PEAC 20080523 -------------------------

        If CInt(Mid(ActxCuenta.NroCuenta, 9, 1)) <> RGarantias!nmoneda Then
            'De garantia Dolares a credito Soles
            If CInt(Mid(ActxCuenta.NroCuenta, 9, 1)) = 1 Then
                If CDbl(Format(CDbl(Me.FEGarantCred.TextMatrix(FEGarantCred.row, 2)) * fnTipoCamb, "#0.00")) > CDbl(Format(CDbl(LblMontoS.Caption), "#0.00")) Then
                    '*** PEAC 20090414 - POR DIFERENCIA DE CAMBIO SE PERMITIRÀ HASTA 0.10 DE DIFERENCIA MAYOR PARA PROSEGUIR
                    If CDbl(Format(CDbl(Me.FEGarantCred.TextMatrix(FEGarantCred.row, 2)) * fnTipoCamb, "#0.00")) - CDbl(Format(CDbl(LblMontoS.Caption), "#0.00")) > 0.1 Then
                        MsgBox "T/C:" + CStr(fnTipoCamb) + " Monto de Garantia a Soles es : " & Format(CDbl(Me.FEGarantCred.TextMatrix(FEGarantCred.row, 2)) * fnTipoCamb, "#0.00") & Chr(10) & ", el cual es mayor al monto solicitado", vbInformation, "Mensaje"
                        FEGarantCred.TextMatrix(pnRow, pnCol) = "0.00"
                        Exit Sub
                    End If
                End If
            Else 'De garantia Soles a credito Dolares
                If CDbl(Format(CDbl(Me.FEGarantCred.TextMatrix(FEGarantCred.row, 2)) / fnTipoCamb, "#0.00")) > CDbl(Format(CDbl(LblMontoS.Caption), "#0.00")) Then
                    '*** PEAC 20090414 - POR DIFERENCIA DE CAMBIO SE PERMITIRÀ HASTA 0.10 DE DIFERENCIA MAYOR PARA PROSEGUIR
                    If CDbl(Format(CDbl(Me.FEGarantCred.TextMatrix(FEGarantCred.row, 2)) / fnTipoCamb, "#0.00")) - CDbl(Format(CDbl(LblMontoS.Caption), "#0.00")) > 0.1 Then
                        MsgBox "T/C:" + CStr(fnTipoCamb) + " Monto de Garantia a Dolares es : " & Format(CDbl(Me.FEGarantCred.TextMatrix(FEGarantCred.row, 2)) / fnTipoCamb, "#0.00") & Chr(10) & ", el cual es mayor al monto solicitado", vbInformation, "Mensaje"
                        FEGarantCred.TextMatrix(pnRow, pnCol) = "0.00"
                        Exit Sub
                    End If
                End If
            End If
        Else
            If CDbl(Format(CDbl(Me.FEGarantCred.TextMatrix(pnRow, 2)), "#0.00")) > CDbl(Format(CDbl(LblMontoS.Caption), "#0.00")) And (CDbl(Me.FEGarantCred.TextMatrix(pnRow, 2)) > 0) Then
                MsgBox "El monto ingresado no puede ser mayor al monto solicitado", vbInformation, "Mensaje"
                FEGarantCred.TextMatrix(pnRow, pnCol) = "0.00"
                Exit Sub
            End If
        End If
        '*** FIN PEAC 20080523 -------------------------
        
        '---------------------------
        'ARCV  11-07-2006
        Dim oGarant As COMDCredito.DCOMGarantia
        Set oGarant = New COMDCredito.DCOMGarantia
        Dim rs As ADODB.Recordset
        
        Set rs = oGarant.RecuperaDatosTablaValores_x_Garantia(FEGarantCred.TextMatrix(pnRow, 9))
        Set oGarant = Nothing
        
        If rs!valor > 0 Then
            If CDbl(FEGarantCred.TextMatrix(pnRow, pnCol)) <> CDbl(rs!valor) And CDbl(FEGarantCred.TextMatrix(pnRow, pnCol)) > 0 Then
                FEGarantCred.TextMatrix(pnRow, pnCol) = Format(rs!valor, "#0.00")
                MsgBox "El monto del Gravamen no puede ser distinto a " & Format(rs!valor, "#0.00"), vbInformation, "Aviso"
                FEGarantCred.row = pnRow
                FEGarantCred.col = pnCol - 1
                Exit Sub
            End If
        End If
        '----------------
    End If
End Sub

Private Sub FEGarantCred_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Call FEGarantCred_OnCellChange(pnRow, pnCol)
End Sub

Private Sub FEGarantCred_RowColChange()
    If FEGarantCred.row <> FilaAct And nCmdEjecutar <> -1 Then
        FEGarantCred.row = FilaAct
    End If
End Sub

Private Sub Form_Load()
    CentraForm Me
    ActxCuenta.CMAC = gsCodCMAC
    ActxCuenta.Age = gsCodAge
    CmdNuevo.Enabled = False
    CmdEditar.Enabled = False
    cmdEliminar.Enabled = False
    nCmdEjecutar = -1
    FilaAct = -1
    
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCredRegistrarGravamen
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
End Sub
