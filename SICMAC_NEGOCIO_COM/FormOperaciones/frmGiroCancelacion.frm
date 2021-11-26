VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGiroCancelacion 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   Icon            =   "frmGiroCancelacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClave 
      Caption         =   "Ingresar Clave de Seguridad"
      Height          =   375
      Left            =   1440
      TabIndex        =   31
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   225
      TabIndex        =   6
      Top             =   5985
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   5985
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   5985
      Width           =   1095
   End
   Begin VB.Frame fraCliente 
      Caption         =   "Remitente/Destinatario"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3360
      Left            =   105
      TabIndex        =   17
      Top             =   2520
      Width           =   7635
      Begin VB.TextBox txtIdAut 
         Height          =   330
         Left            =   6120
         TabIndex        =   29
         Top             =   195
         Visible         =   0   'False
         Width           =   1380
      End
      Begin TabDlg.SSTab tabRemDest 
         Height          =   2790
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   4921
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Remitente"
         TabPicture(0)   =   "frmGiroCancelacion.frx":030A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraRem"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Destinatario"
         TabPicture(1)   =   "frmGiroCancelacion.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraDest"
         Tab(1).ControlCount=   1
         Begin VB.Frame fraRem 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1995
            Left            =   180
            TabIndex        =   19
            Top             =   420
            Width           =   6975
            Begin VB.Label lblDireccion 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   345
               Left            =   1140
               TabIndex        =   27
               Top             =   1020
               Width           =   5475
            End
            Begin VB.Label lblNombre 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   345
               Left            =   1140
               TabIndex        =   26
               Top             =   660
               Width           =   5475
            End
            Begin VB.Label lblFecNac 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   350
               Left            =   4500
               TabIndex        =   25
               Top             =   300
               Width           =   2115
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Fec. Nac. :"
               Height          =   195
               Left            =   3600
               TabIndex        =   24
               Top             =   378
               Width           =   795
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Dirección :"
               Height          =   195
               Left            =   180
               TabIndex        =   23
               Top             =   1095
               Width           =   765
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Nombre :"
               Height          =   195
               Left            =   180
               TabIndex        =   22
               Top             =   735
               Width           =   645
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Doc. ID.:"
               Height          =   195
               Left            =   180
               TabIndex        =   21
               Top             =   378
               Width           =   645
            End
            Begin VB.Label lblDocID 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   345
               Left            =   1140
               TabIndex        =   20
               Top             =   300
               Width           =   1395
            End
         End
         Begin VB.Frame fraDest 
            Height          =   2115
            Left            =   -74880
            TabIndex        =   18
            Top             =   360
            Width           =   7155
            Begin SICMACT.FlexEdit grdDest 
               Height          =   1455
               Left            =   120
               TabIndex        =   28
               Top             =   180
               Width           =   6915
               _ExtentX        =   12197
               _ExtentY        =   2566
               Cols0           =   6
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "#-Nombre-Referencia-cPersCod-bRegistrado-Direccion"
               EncabezadosAnchos=   "350-3000-3400-0-0-0"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
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
               ColumnasAEditar =   "X-X-X-X-X-X"
               TextStyleFixed  =   3
               ListaControles  =   "0-0-0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-L-L-C-C-C"
               FormatosEdit    =   "0-0-0-0-0-0"
               TextArray0      =   "#"
               lbFlexDuplicados=   0   'False
               lbUltimaInstancia=   -1  'True
               ColWidth0       =   345
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
            Begin VB.CommandButton cmdRegistrar 
               Caption         =   "&Registrar Destinatario"
               Height          =   375
               Left            =   2580
               TabIndex        =   3
               Top             =   1680
               Width           =   1815
            End
         End
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Id Autorización"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4785
         TabIndex        =   30
         Top             =   315
         Visible         =   0   'False
         Width           =   1290
      End
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Datos Giro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2190
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7635
      Begin VB.CommandButton cmdExaminar 
         Caption         =   "&Examinar"
         Height          =   375
         Left            =   3840
         TabIndex        =   1
         Top             =   300
         Width           =   975
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   435
         Left            =   180
         TabIndex        =   0
         Top             =   300
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   767
         Texto           =   "Giro N°"
      End
      Begin VB.Label lblITF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   5220
         TabIndex        =   37
         Top             =   1200
         Width           =   1755
      End
      Begin VB.Label lblSimbolo 
         AutoSize        =   -1  'True
         Caption         =   "S/"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   7095
         TabIndex        =   36
         Top             =   1275
         Width           =   270
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ITF :"
         Height          =   195
         Left            =   4740
         TabIndex        =   35
         Top             =   1275
         Width           =   345
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   5220
         TabIndex        =   34
         Top             =   1605
         Width           =   1755
      End
      Begin VB.Label Label10 
         Caption         =   "IMPORTE PAGO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   33
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblSimbolo 
         AutoSize        =   -1  'True
         Caption         =   "S/"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   7080
         TabIndex        =   16
         Top             =   840
         Width           =   270
      End
      Begin VB.Label lblAgencia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1500
         TabIndex        =   15
         Top             =   795
         Width           =   2595
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ag. Destino :"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   870
         Width           =   915
      End
      Begin VB.Label lblMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   5220
         TabIndex        =   13
         Top             =   795
         Width           =   1755
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Left            =   4560
         TabIndex        =   12
         Top             =   870
         Width           =   540
      End
      Begin VB.Label lblTipo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1500
         TabIndex        =   11
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Giro :"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1635
         Width           =   735
      End
      Begin VB.Label lblApertura 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1500
         TabIndex        =   9
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Apertura :"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1275
         Width           =   1185
      End
   End
   Begin VB.PictureBox Tarjeta 
      Height          =   375
      Left            =   7200
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   32
      Top             =   5520
      Width           =   615
   End
End
Attribute VB_Name = "frmGiroCancelacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim sOperacion As String, sRemitente As String

'***************Variabres Agregadas********************
Dim Gtitular As String
Dim GAutNivel As String
Dim GAutMontoFinSol As Double
Dim GAutMontoFinDol As Double
Dim GMontoAprobado As Double, GNroID As Long, GPersCod As String, GClaveGiro As String
Dim fnPersRealizaOpe As Boolean 'WIOR 20121015
Dim fnPersPersoneriaGen As Integer 'WIOR 20121015
Dim fnPersCodRealiza As String 'WIOR 20121015
Dim fcPersCod As String 'WIOR 20130301
Dim fbValidaClave As Boolean  'RECO20140530 ERS008-2014
Dim fcAgeCodDest As String 'RECO20140530 ERS008-2014
Dim lsNumTarjGir As String
Public lnValPinPad As Integer

'Variables necesarias para banca por internet
Public lbEsBanca As Boolean

Private Function EsExoneradaLavadoDinero() As Boolean

    Dim i As Long
    Dim bExito As Boolean
    Dim clsExo As COMNCaptaServicios.NCOMCaptaServicios
    Dim sPersCod As String
    bExito = True
    Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
    For i = 1 To grdDest.Rows - 1
        sPersCod = grdDest.TextMatrix(i, 3)
        If Not clsExo.EsPersonaExoneradaLavadoDinero(sPersCod) Then
            bExito = False
            Exit For
        End If
    Next i
    Set clsExo = Nothing
    EsExoneradaLavadoDinero = bExito
End Function

Private Function IniciaLavDinero(poLavDinero As frmMovLavDinero) As String
    Dim i As Long
    Dim nRelacion As COMDConstantes.CaptacRelacPersona
    Dim sPersCod As String, sNombre As String, sDocId As String, sDireccion As String
    Dim nMonto As Double
    Dim sCuenta As String

    poLavDinero.TitPersLavDinero = sRemitente
    poLavDinero.TitPersLavDineroNom = lblNombre
    poLavDinero.TitPersLavDineroDir = lblDireccion
    poLavDinero.TitPersLavDineroDoc = lblDocID

    nMonto = CDbl(lblMonto)
    sCuenta = txtCuenta.NroCuenta

    'IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, True, False, nMonto, sCuenta, sOperacion)
End Function

Private Function DestinatarioRegistrado() As Long
    Dim i As Long, nFila As Long
    Dim bRegistrado As Boolean
    bRegistrado = True
    nFila = 0
    For i = 1 To grdDest.Rows - 1
        If grdDest.TextMatrix(i, 3) = "" Then
            bRegistrado = False
            nFila = i
            Exit For
        End If
    Next i
    DestinatarioRegistrado = nFila
End Function

Private Sub ClearScreen()
    grdDest.Rows = 2
    grdDest.FormaCabecera
    lblNombre = ""
    lblMonto = "0.00"
    lblITF = "" 'NAGL 20181030
    lblTotal = "0.00" 'NAGL 20181030
    lblTipo = ""
    lblApertura = ""
    LblAgencia = ""
    lblFecNac = ""
    lblDireccion = ""
    lblDocID = ""
    cmdClave.Enabled = False
    cmdGrabar.Enabled = False
    CmdCancelar.Enabled = False
    cmdRegistrar.Enabled = False
    fraDatos.Enabled = True
    FraCliente.Enabled = False
    txtCuenta.Cuenta = ""
    txtCuenta.Age = ""
    txtCuenta.Prod = gGiro
    txtCuenta.CMAC = gsCodCMAC
    txtCuenta.EnabledProd = False
    txtCuenta.EnabledAge = True
    txtCuenta.EnabledCta = True
    lblMonto.BackColor = &HFFFFFF
    'lblSimbolo.Caption = ""
    lblSimbolo(0).Caption = ""
    lblSimbolo(2).Caption = "" 'NAGL 20181030 Agregó el array
    sRemitente = ""
    fbValidaClave = False 'RECO20140530 ERS008-2014
End Sub

Private Sub CargaDatosGiro(ByVal sCuenta As String)
    Dim rsGiro As ADODB.Recordset
    Dim clsGiro As COMNCaptaServicios.NCOMCaptaServicios

    Dim nFila As Long
    Dim sDestinatario As String
    Set clsGiro = New COMNCaptaServicios.NCOMCaptaServicios
    Set rsGiro = clsGiro.GetGiroDatos(sCuenta)
    If Not (rsGiro.EOF And rsGiro.BOF) Then
        sDestinatario = ""
        LblAgencia = Trim(rsGiro("cAgencia"))
        lblMonto = Format$(rsGiro("nSaldo"), "#,##0.00")
        Call CalculaITFCargo 'NAGL Según RFC1807260001
        lblTipo = Trim(rsGiro("cTipo"))
        lblApertura = Format$(rsGiro("dPrdEstado"), "dd mmm yyyy")
        lblNombre = PstaNombre(Trim(rsGiro("cRemitente")), False)
        lblDireccion = Trim(rsGiro("cPersDireccDomicilio"))
        lblFecNac = Format$(rsGiro("dFecNac"), "dd mmm yyyy")
        lblDocID = Trim(rsGiro("cDocID"))
        sRemitente = Trim(rsGiro("cPersCod"))
        fcAgeCodDest = Trim(rsGiro("cAgenciaDest"))
        lnValPinPad = Trim(rsGiro("nClavePinPad"))
        Dim dlsMant As COMDCaptaGenerales.DCOMCaptaGenerales
        Set dlsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    
        If dlsMant.GetNroOPeradoras(gsCodAge) > 1 Then
            If sRemitente = gsCodPersUser Then
                MsgBox "Ud. No puede hacer operaciones con sus propias cuentas.", vbInformation, "Aviso"
                Unload Me
                Exit Sub
            End If
        End If
        Set dlsMant = Nothing
        
        Do While Not rsGiro.EOF
            If sDestinatario <> Trim(rsGiro("cDestinatario")) Then
                sDestinatario = Trim(rsGiro("cDestinatario"))
                If grdDest.TextMatrix(1, 1) <> "" Then grdDest.Rows = grdDest.Rows + 1
                nFila = grdDest.Rows - 1
                grdDest.TextMatrix(nFila, 0) = nFila
                grdDest.TextMatrix(nFila, 2) = IIf(IsNull(rsGiro("cDocID")), rsGiro("cReferencia"), Trim(rsGiro("cDocID")))
                If rsGiro("cCodDest") = "" Then
                    Dim j As Long
                    For j = 1 To grdDest.Cols - 1
                        grdDest.Col = j
                        grdDest.CellBackColor = &HFFFFC0
                    Next j
                    grdDest.TextMatrix(nFila, 1) = Trim(rsGiro("cDestinatario"))
                    grdDest.TextMatrix(nFila, 2) = Trim(rsGiro("cReferencia"))
                    cmdRegistrar.Enabled = True
                    grdDest.TextMatrix(nFila, 4) = "0"
                Else
                    grdDest.TextMatrix(nFila, 1) = PstaNombre(rsGiro("cDestinatario"), False)
                    grdDest.TextMatrix(nFila, 2) = Trim(rsGiro("cReferencia"))
                    grdDest.TextMatrix(nFila, 4) = "1"
                    'WIOR 20121107 **********************
                    fnPersPersoneriaGen = IIf(CInt(rsGiro("nPersoneria")) > 1, 2, 1)
                    'WIOR FIN ***************************
                End If
                grdDest.TextMatrix(nFila, 3) = rsGiro("cCodDest")
                grdDest.TextMatrix(nFila, 5) = rsGiro("cDireccion")
            End If
            rsGiro.MoveNext
        Loop
        fraDatos.Enabled = False
        FraCliente.Enabled = True
        cmdGrabar.Enabled = True
        CmdCancelar.Enabled = True
        'Verifica si el Giro tiene Clave
        'GClaveGiro = clsGiro.GetGiroSeguridad(sCuenta)
        
        'ANDE 20190620: Modificación obtener clave de giro para banca por internet
        GClaveGiro = clsGiro.GetGiroSeguridad(sCuenta, lbEsBanca)
        'END ANDE 20190620
        If GClaveGiro <> "" Then
            cmdClave.Enabled = True
        End If
        cmdGrabar.SetFocus
    Else
        MsgBox "Número de Giro no encontrado o Cancelado.", vbInformation, "SICMACM - Aviso"
        txtCuenta.Age = ""
        txtCuenta.Cuenta = ""
        txtCuenta.SetFocusAge
        sRemitente = ""
    End If
    Set clsGiro = Nothing
End Sub

Public Sub CalculaITFCargo()
Dim oITF As New COMDConstSistema.FCOMITF
Dim nRedondeoITF As Double
If gbITFAplica Then
    If lblMonto.Caption > gnITFMontoMin Then
        Me.lblITF.Caption = Format(fgITFCalculaImpuesto(lblMonto), "#,##0.00")
        nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
        If nRedondeoITF > 0 Then
            Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
        End If
        lblTotal.Caption = Format(CDbl(lblMonto.Caption) - CDbl(lblITF.Caption), "#,##0.00")
    Else
        Me.lblITF.Caption = "0.00"
        lblTotal.Caption = Format(CDbl(lblMonto.Caption), "0.00")
    End If
Else
        Me.lblITF.Caption = "0.00"
        lblTotal.Caption = Format(CDbl(lblMonto.Caption), "0.00")
End If
End Sub 'NAGL Según RFC1807260001

Private Sub cmdCancelar_Click()
    ClearScreen
    txtCuenta.SetFocus
End Sub

Private Sub cmdClave_Click()
    Dim sClaveSeg As String 'RECO20140530 ERS008-2014
    Dim Pinpads As New clases.clsPinpad 'added by marg20191120 hb
    If lnValPinPad <> 1 Then
        sClaveSeg = InputBox("Ingrese la clave de seguridad.", "SICMACM - Giros") 'RECO20140530 ERS008-2014
        'MsgBox "La clave de Seguridad es:" & Chr(13) & GClaveGiro, , "Clave Giro"'RECO20140530 ERS008-2014
    
    Else
        lsNumTarjGir = "4697100000000025"
        'sClaveSeg = Tarjeta.PedirPinEnc(lsNumTarjGir, gNMKPOS, gWKPOS, 1, gnPinPadPuerto) 'commented by ande-marg hb
        
        sClaveSeg = Pinpads.PedirPinEncNDig(lsNumTarjGir, gNMKPOS, gWKPOS, 1, gnPinPadPuerto, IIf(Not lbEsBanca, 4, 6)) 'added by marg-ande hb
        If sClaveSeg = "" Then
            MsgBox "No existe conexión con PinPad", vbInformation, "SICMACM - Operaciones"
        Else
            MsgBox "Clave de Seguridad Ingresada", vbInformation, "SICMACM - Operaciones"
        End If
    End If
    
    If sClaveSeg = GClaveGiro Then 'RECO20140530 ERS008-2014
        fbValidaClave = True 'RECO20140530 ERS008-2014
        cmdGrabar.Enabled = True
        cmdClave.Enabled = False
    Else
        fbValidaClave = False
    End If 'RECO20140530 ERS008-2014
End Sub

Private Sub cmdExaminar_Click()
frmGiroPendiente.Inicio frmGiroCancelacion 'RECO20140415 ERS008-214
'frmGiroPendiente.Show 'RECO20140415 ERS008-214
Dim sCuenta As String
Dim nmoneda As Moneda

sCuenta = txtCuenta.NroCuenta
If Len(sCuenta) = 18 Then
    txtCuenta.SetFocusCuenta
    nmoneda = CLng(Mid(sCuenta, 9, 1))
    If nmoneda = COMDConstantes.gMonedaExtranjera Then
        lblMonto.BackColor = &HC0FFC0
        lblSimbolo(0).Caption = "US$"
        lblSimbolo(2).Caption = "US$" 'NAGL 20181030 Agregó el array
    Else
        lblMonto.BackColor = &HFFFFFF
        'lblSimbolo.Caption = "S/"
        lblSimbolo(0).Caption = gcPEN_SIMBOLO
        lblSimbolo(2).Caption = gcPEN_SIMBOLO 'NAGL 20181030 Agregó el array
    End If
    SendKeys "{Enter}"
End If
End Sub

Private Sub cmdGrabar_Click()
'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
'fin Comprobacion si es RFIII

Dim fbPersonaReaOtros As Boolean 'WIOR 20130301
Dim fnCondicion As Integer 'WIOR 20130301
    Dim sCuenta As String, sMovNro As String, sPersLavDinero As String
    Dim nMonto As Double
    Dim nFila As Long
    Dim rsGiro As ADODB.Recordset
    Dim clsGiro As COMNCaptaServicios.NCOMCaptaServicios
    Dim ClsMov As COMNContabilidad.NCOMContFunciones
    Dim nMontoLavDinero As Double, nTC As Double
    Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim nmoneda As COMDConstantes.Moneda
    Dim nFicSal As String
    Dim lsBoleta As String
    'By Capi 14022008
    Dim loLavDinero As frmMovLavDinero
    Set loLavDinero = New frmMovLavDinero
    'End by
    'FRHU ERS077-2015 20151204
    Dim item As Integer
    For item = 1 To grdDest.Rows - 1
        If grdDest.TextMatrix(item, 3) <> "" Then
            Call VerSiClienteActualizoAutorizoSusDatos(grdDest.TextMatrix(item, 3), gServGiroCancEfec)
        End If
    Next item
    'FIN FRHU ERS077-2015
    nmoneda = CLng(Mid(txtCuenta.NroCuenta, 9, 1))
    nFila = DestinatarioRegistrado()
    If nFila > 0 Then
        MsgBox "El Destinatario tiene que ser registrado.", vbInformation, "Aviso"
        tabRemDest.Tab = 1
        grdDest.row = nFila
        cmdRegistrar.SetFocus
        Exit Sub
    End If
    'RECO20143005 ERS008-2014*********************
    If fbValidaClave = False Then
        fbValidaClave = False
        MsgBox "Clave de seguridad incorrecta o aún no fue ingresada", vbInformation, "Aviso"
        Exit Sub
    End If
    'If Me.txtCuenta.Age = gsCodAge Then
    If Me.txtCuenta.Age = gsCodAge And Not lbEsBanca Then
        MsgBox "No se puede cancelar giros de la misma agencia", vbInformation, "Aviso"
        Exit Sub
    End If
    'RECO FIN*************************************
    If MsgBox("¿Desea Grabar la Operación?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If

    'nMonto = CDbl(lblMonto)'Comentado by NAGL 20181030
    nMonto = CDbl(lblTotal) 'NAGL 20181030 Según RFC1807260001
    'Realiza la Validación para el Lavado de Dinero
    If Not EsExoneradaLavadoDinero() Then
        sPersLavDinero = ""
        Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
        nMontoLavDinero = clsLav.GetCapParametro(COMDConstantes.gMonOpeLavDineroME)
        Set clsLav = Nothing
    
        If nmoneda = COMDConstantes.gMonedaNacional Then
            Dim clsTC As COMDConstSistema.DCOMGeneral
            Set clsTC = New COMDConstSistema.DCOMGeneral
            nTC = clsTC.EmiteTipoCambio(gdFecSis, COMDConstantes.TCFijoDia)
            Set clsTC = Nothing
        Else
            nTC = 1
        End If
        If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
            'By Capi 1402208
            Call IniciaLavDinero(loLavDinero)
            'ALPA 20081009***********************************************************************************************************************
            'sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, sCuenta, Mid(Me.Caption, 15), True, "Servicios-Giro-Cancelacion", , , , , nmoneda)
            sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, sCuenta, Mid(Me.Caption, 13), True, "Servicios-Giro-Cancelacion", , , , , nmoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen, , gServGiroCancEfec) 'WIOR 20131106 AGREGO gServGiroCancEfec
            '************************************************************************************************************************************
            If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
            'End
        End If
    End If
    'WIOR 20130301 comento INICIO *********************************
    ''WIOR 20121015 *****************************************
    'If fnPersPersoneriaGen = 2 And loLavDinero.OrdPersLavDinero = "Exit" Then
    '    frmPersRealizaOperacion.Inicia "Giros", gPersRealizaGirosCanc
    '    fnPersRealizaOpe = frmPersRealizaOperacion.PersRegistrar
    '    fnPersCodRealiza = frmPersRealizaOperacion.PersCod
    '    If Not fnPersRealizaOpe Then
    '        MsgBox "Se va a proceder a Anular la Cancelacion del Giro."
    '        Exit Sub
    '    End If
    'Else
    '    fnPersCodRealiza = "Exit"
    'End If
    ''WIOR FIN **********************************************
    'WIOR 20130301 comento FIN *********************************
      'WIOR 20130301 Personas Sujetas a Procedimiento Reforzado *************************************
        fbPersonaReaOtros = False
        If (loLavDinero.OrdPersLavDinero = "Exit") Then
                
                Dim oPersonaSPR As UPersona_Cli
                Dim oPersonaU As COMDPersona.UCOMPersona
                Dim nTipoConBN As Integer
                Dim sConPersona As String
                Dim pbClienteReforzado As Boolean
                Dim rsAgeParam As Recordset
                Dim objCap As COMNCaptaGenerales.NCOMCaptaMovimiento
                Dim lnMontoX As Double, lnTC As Double
                Dim ObjTc As COMDConstSistema.NCOMTipoCambio
                
                
                Set oPersonaU = New COMDPersona.UCOMPersona
                Set oPersonaSPR = New UPersona_Cli
                
                fbPersonaReaOtros = False
                pbClienteReforzado = False
                fnCondicion = 0
                
                    oPersonaSPR.RecuperaPersona grdDest.TextMatrix(nFila + 1, 3)
                                        
                    If oPersonaSPR.Personeria = 1 Then
                        If oPersonaSPR.Nacionalidad <> "04028" Then
                            sConPersona = "Extranjera"
                            fnCondicion = 1
                            pbClienteReforzado = True
                        ElseIf oPersonaSPR.Residencia <> 1 Then
                            sConPersona = "No Residente"
                            fnCondicion = 2
                            pbClienteReforzado = True
                        ElseIf oPersonaSPR.RPeps = 1 Then
                            sConPersona = "PEPS"
                            fnCondicion = 4
                            pbClienteReforzado = True
                        ElseIf oPersonaU.ValidaEnListaNegativaCondicion(IIf(Trim(oPersonaSPR.ObtenerDNI) = "", oPersonaSPR.ObtenerNumeroDoc(0), oPersonaSPR.ObtenerDNI), oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                            If nTipoConBN = 1 Or nTipoConBN = 3 Then
                                sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                                fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                                pbClienteReforzado = True
                            End If
                        End If
                    Else
                        If oPersonaU.ValidaEnListaNegativaCondicion(oPersonaSPR.ObtenerDNI, oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                            If nTipoConBN = 1 Or nTipoConBN = 3 Then
                                sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                                fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                                pbClienteReforzado = True
                            End If
                        End If
                    End If
                    
                If pbClienteReforzado Then
                    MsgBox "El Cliente: " & Trim(grdDest.TextMatrix(nFila + 1, 1)) & " es un Cliente de Procedimiento Reforzado (Persona " & sConPersona & ")", vbInformation, "Aviso"
                    frmPersRealizaOpeGeneral.Inicia sOperacion & " (Persona " & sConPersona & ")", gServGiroCancEfec
                    fbPersonaReaOtros = frmPersRealizaOpeGeneral.PersRegistrar
                    
                    If Not fbPersonaReaOtros Then
                        MsgBox "Se va a proceder a Anular la Operacion ", vbInformation, "Aviso"
                        cmdGrabar.Enabled = True
                        Exit Sub
                    End If
                Else
                    fnCondicion = 0
                    lnMontoX = nMonto
                    pbClienteReforzado = False
                    
                    Set ObjTc = New COMDConstSistema.NCOMTipoCambio
                    lnTC = ObjTc.EmiteTipoCambio(gdFecSis, TCFijoDia)
                    Set ObjTc = Nothing
                
                
                    Set objCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
                    Set rsAgeParam = objCap.getCapAbonoAgeParam(gsCodAge)
                    Set objCap = Nothing
                    
                    If Mid(Trim(txtCuenta.NroCuenta), 9, 1) = 1 Then
                        lnMontoX = Round(lnMontoX / lnTC, 2)
                    End If
                
                    If Not (rsAgeParam.EOF And rsAgeParam.BOF) Then
                        If lnMontoX >= rsAgeParam!nMontoMin And lnMontoX <= rsAgeParam!nMontoMax Then
                            frmPersRealizaOpeGeneral.Inicia sOperacion, gServGiroCancEfec
                            fbPersonaReaOtros = frmPersRealizaOpeGeneral.PersRegistrar
                            If Not fbPersonaReaOtros Then
                                MsgBox "Se va a proceder a Anular la Operacion", vbInformation, "Aviso"
                                cmdGrabar.Enabled = True
                                Exit Sub
                            End If
                        End If
                    End If
                    
                End If
        End If
        'WIOR FIN ***************************************************************
    sCuenta = txtCuenta.NroCuenta

    Set rsGiro = grdDest.GetRsNew()
    Set ClsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set ClsMov = Nothing
    Set clsGiro = New COMNCaptaServicios.NCOMCaptaServicios
    'By Capi 18022008
    
    'clsGiro.ServGiroCancelacionEfectivo nMonto, sMovNro, rsGiro, sCuenta, gsNomAge, lblNombre.Caption, sLpt, nmoneda, sPersLavDinero, lsBoleta, gbImpTMU
    'ALPA 20081009*****************************************************
    clsGiro.ServGiroCancelacionEfectivo nMonto, sMovNro, rsGiro, sCuenta, gsNomAge, lblNombre.Caption, sLpt, nmoneda, sPersLavDinero, lsBoleta, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, gServGiroCancEfec, COMDConstantes.gITFGiroCancelEfect, CCur(Me.lblITF.Caption)
    'SE CAMBIO DE gsOpeCod POR gServGiroCancEfec EN ServGiroCancelacionEfectivo JHCU
    'NAGL - RFC1807260001 Agregó los siguientes parámetros gsOpeCod, COMDConstantes.gITFCobroEfectivo, CCur(Me.lblITF.Caption)
    'NAGL 20190307 Cambió de COMDConstantes.gITFCobroEfectivo A COMDConstantes.gITFGiroCancelEfect

    If gnMovNro > 0 Then
        'Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen)
        Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) ' JACA 20110224
    End If
    '******************************************************************
      If Trim(lsBoleta) <> "" Then
        'By Capi 20012008
        Dim lbok As Boolean
        lbok = True
        Do While lbok
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
            Print #nFicSal, lsBoleta
            Print #nFicSal, ""
            Close #nFicSal
            If MsgBox("Desea Reimprimir Boleta ??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                lbok = False
            End If
        Loop
    End If
    'WIOR 20130301 comento INICIO *********************************
    ''WIOR 20121015 ************************************
    'If fnPersRealizaOpe Then
    '    frmPersRealizaOperacion.InsertaPersonaRealizaOperacion gnMovNro, sCuenta, frmPersRealizaOperacion.PersTipoCliente, _
    '    frmPersRealizaOperacion.PersCod, frmPersRealizaOperacion.PersTipoDOI, frmPersRealizaOperacion.PersDOI, frmPersRealizaOperacion.PersNombre, _
    '    frmPersRealizaOperacion.TipoOperacion
    '
    '    fnPersRealizaOpe = False
    'End If
    'fnPersPersoneriaGen = 0
    ''WIOR FIN *****************************************
    'WIOR 20130301 comento FIN *********************************
    'WIOR 20130301 ************************************************************
    If fbPersonaReaOtros And gnMovNro > 0 Then
        frmPersRealizaOpeGeneral.InsertaPersonasOperacion gnMovNro, Trim(sCuenta), fnCondicion
        fbPersonaReaOtros = False
    End If
    'WIOR FIN *****************************************************************
    Set loLavDinero = Nothing
    Set clsGiro = Nothing

    ClearScreen
    
     'INICIO JHCU ENCUESTA 16-10-2019
    Encuestas gsCodUser, gsCodAge, "ERS0292019", gServGiroCancEfec
    'FIN
End Sub

Private Sub cmdRegistrar_Click()
Dim sCodCli As String
Dim nFila As Long
'By Capi 06122007
Dim X As Integer
Dim cNombre As String
Dim cCadenaTotal As String
Dim cCadenaParcial As String
Dim nEspacio As Integer
Dim nContiene As Integer
Dim bCoincide As Boolean

'End By

nFila = grdDest.row
If grdDest.TextMatrix(nFila, 3) <> "" Then
    MsgBox "Usuario Registrado. No es posible modificar sus datos", vbInformation
    Exit Sub
End If
    Dim clsPers As COMDPersona.UCOMPersona
    Set clsPers = frmBuscaPersona.Inicio(False)
    If Not clsPers Is Nothing Then
        If clsPers.sPersCod <> "" Then
            'By Capi 06122007 para validar cliente elegido
            bCoincide = True
            cNombre = Trim(clsPers.sPersNombre)
            cCadenaTotal = Trim(grdDest.TextMatrix(nFila, 1))
            nEspacio = 1
            nContiene = 1
            'By Capi 02092008 cuando el nombre que digiten en giros es mayor al que esta en Personas
            If Len(cCadenaTotal) > Len(cNombre) Then
                MsgBox "Cliente Elegido. No coincide con el registrado en la Agencia Origen"
                bCoincide = False
                nEspacio = 0
            End If
            '
            Do While nEspacio <> 0
                If nEspacio = 1 Then
                   nEspacio = InStr(nEspacio, cCadenaTotal, " ", vbTextCompare)
                Else
                   nEspacio = InStr(nEspacio + 1, cCadenaTotal, " ", vbTextCompare)
                End If
                If nEspacio = 0 Then
                    cCadenaParcial = Mid(cCadenaTotal, nContiene + 1, Len(Trim(clsPers.sPersNombre)) - nContiene - 1)
                Else
                    cCadenaParcial = Mid(cCadenaTotal, nContiene + 1, nEspacio - nContiene - 1)
                End If
                nContiene = InStr(1, cNombre, cCadenaParcial, vbTextCompare)
                If nContiene = 0 Or Null Then
                    MsgBox "Cliente Elegido. No coincide con el registrado en la Agencia Origen"
                    bCoincide = False
                    Exit Do
                End If
                nContiene = nEspacio
           Loop
           'End By
           If bCoincide = True Then
                grdDest.TextMatrix(nFila, 2) = clsPers.sPersIdnroDNI
                grdDest.TextMatrix(nFila, 1) = PstaNombre(Trim(clsPers.sPersNombre), False)
                grdDest.TextMatrix(nFila, 5) = Trim(clsPers.sPersDireccDomicilio)
                grdDest.TextMatrix(nFila, 3) = clsPers.sPersCod
           End If
        End If
    End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    'By Capi 05122007 para que no se pueda cambiar al destinatario
    cmdRegistrar.Enabled = False
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    Me.Caption = "Servicios - Giro - Cancelación"
    sOperacion = "SERV. GIRO CANCELACION"
    ClearScreen
    fnPersRealizaOpe = False 'WIOR 20121015
    fbValidaClave = False 'RECO20140530 ERS008-2014
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
'        'RECO20140530 ERS008-2014*******************************
'        If Me.txtCuenta.Age = gsCodAge Then
'            MsgBox "No se puede cancelar giros de la misma agencia", vbCritical, "Aviso"
'            Call ClearScreen
'            Exit Sub
'        End If
'        'RECO FIN***********************************************
        Dim EsPorBanca As Boolean
       
        'ANDE20190617: Validando si es un giro hecho por banca por internet
        Dim ServiciosLN As New COMNCaptaServicios.NCOMCaptaServicios
           
        EsPorBanca = ServiciosLN.GetGiroBanxaByInternet(txtCuenta.GetCuenta)
        lbEsBanca = EsPorBanca
        If Not EsPorBanca Then
              If Me.txtCuenta.Age = gsCodAge Then
                MsgBox "No se puede cancelar giros de la misma agencia", vbCritical, "Aviso"
                Call ClearScreen
                Exit Sub
            End If
        End If
        'end validación giro de banca por internet-------------
        
        Dim sCuenta As String, sMoneda As String
        sCuenta = txtCuenta.NroCuenta
        
        '*****Agregado by NAGL 20181030******
        sCuenta = txtCuenta.NroCuenta
        sMoneda = Mid(sCuenta, 9, 1)
        If sMoneda = "2" Then
            lblMonto.BackColor = &HC0FFC0
            lblSimbolo(0).Caption = "US$"
            lblSimbolo(2).Caption = "US$"
        Else
            lblMonto.BackColor = &HFFFFFF
            'lblSimbolo.Caption = "S/."
            lblSimbolo(0).Caption = gcPEN_SIMBOLO
            lblSimbolo(2).Caption = gcPEN_SIMBOLO
        End If '******NAGL 20181030*********
        
        CargaDatosGiro sCuenta
'        If fcAgeCodDest <> gsCodAge Then
'            MsgBox "Giro no corresponde a agencia", vbCritical, "Aviso"
'            Call ClearScreen
'            Exit Sub
'        End If

        'ANDE20190617: Validando si es un giro hecho por banca por internet--
        If Not EsPorBanca Then
            If fcAgeCodDest <> gsCodAge Then
                MsgBox "Giro no corresponde a agencia", vbCritical, "Aviso"
                Call ClearScreen
                Exit Sub
            End If
        End If
        'end---------------------------------------------------
    End If
End Sub

Private Sub TxtCuenta_KeyUp(KeyCode As Integer, Shift As Integer)
Dim sCuenta As String, sMoneda As String
sCuenta = txtCuenta.NroCuenta
sMoneda = Mid(sCuenta, 9, 1)
If sMoneda = "2" Then
    lblMonto.BackColor = &HC0FFC0
    'lblSimbolo.Caption = "US$"
    lblSimbolo(0).Caption = "US$"
    lblSimbolo(2).Caption = "US$" 'NAGL 20181030 Agregó el array
Else
    lblMonto.BackColor = &HFFFFFF
    'lblSimbolo.Caption = "S/."
    'lblSimbolo.Caption = gcPEN_SIMBOLO
    lblSimbolo(0).Caption = gcPEN_SIMBOLO
    lblSimbolo(2).Caption = gcPEN_SIMBOLO 'NAGL 20181030 Agregó el array
End If
End Sub

Private Sub txtIdAut_KeyPress(KeyAscii As Integer)
Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   Dim Nsoperacion As String
   
   Dim nmoneda As COMDConstantes.Moneda

nmoneda = CLng(Mid(txtCuenta.NroCuenta, 9, 1))

    Nsoperacion = COMDConstantes.gServGiroCancEfec
   If grdDest.TextMatrix(1, 3) <> "" Then
    Gtitular = grdDest.TextMatrix(1, 3)
   Else
     MsgBox "Registrar Destinatario"
     Exit Sub
   End If
        
        
   If KeyAscii = 13 And Trim(txtIdAut.Text) <> "" And Len(txtCuenta.NroCuenta) = 18 Then
      Dim oCap  As COMDCaptaGenerales.COMDCaptAutorizacion
      Set oCap = New COMDCaptaGenerales.COMDCaptAutorizacion
            Set rs = oCap.SAA(Left(CStr(Nsoperacion), 4) & "00", Vusuario, txtCuenta.NroCuenta, Gtitular, CInt(nmoneda), CLng(Val(txtIdAut.Text)))
      Set oCap = Nothing
     If rs.State = 1 Then
       If rs.RecordCount > 0 Then
        lblMonto.Caption = rs!nMontoAprobado
      Else
          MsgBox "No Existe este Id de Autorización para esta cuenta." & vbCrLf & "Consulte las Operaciones Pendientes.", vbOKOnly + vbInformation, "Atención"
          txtIdAut.Text = ""
       End If
       
     End If
   End If
 If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 13 Or KeyAscii = 8) Then
      KeyAscii = 0
   End If
End Sub

Private Function Cargousu(ByVal NomUser As String) As String
 Dim rs As New ADODB.Recordset
 Dim oCons As COMDConstSistema.DCOMUAcceso
 Set oCons = New COMDConstSistema.DCOMUAcceso
 
 Set rs = oCons.Cargousu(NomUser)
  If Not (rs.EOF And rs.BOF) Then
    Cargousu = rs(0)
  End If
 Set rs = Nothing
 'rs.Close
 Set oCons = Nothing
End Function
