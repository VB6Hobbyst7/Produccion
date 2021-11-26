VERSION 5.00
Begin VB.Form frmViaticoSolReimpr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reimpresión de Solicitud de Viáticos"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8010
   Icon            =   "frmViaticoSolReimpr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   7815
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Elija una solicitud y presione en el boton Reimprimir para Imprimir la solicitud"
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
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   7575
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdReimprimir 
      Caption         =   "Reimprimir"
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Solicitudes de Viáticos"
      Height          =   3015
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   7815
      Begin Sicmact.FlexEdit fgSolViaticos 
         Height          =   2685
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   4736
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Nº de Documento-Fecha Solicitud-Monto-cMovNro"
         EncabezadosAnchos=   "350-2500-2500-1200-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         ColumnasAEditar =   "X-1-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-3-0-0-0"
         EncabezadosAlineacion=   "C-L-C-R-L"
         FormatosEdit    =   "0-0-5-2-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Colaborador"
      ClipControls    =   0   'False
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin Sicmact.TxtBuscar txtBuscaPers 
         Height          =   345
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
         TipoBusPers     =   1
      End
      Begin VB.Label lblAreaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1440
         TabIndex        =   9
         Top             =   1200
         Width           =   6240
      End
      Begin VB.Label lblAreaCod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   840
         TabIndex        =   8
         Top             =   1200
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Area:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblNrodoc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6120
         TabIndex        =   6
         Top             =   840
         Width           =   1560
      End
      Begin VB.Label lblpersNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   840
         TabIndex        =   5
         Top             =   840
         Width           =   4755
      End
      Begin VB.Label lblDni 
         Caption         =   "D.N.I:"
         Height          =   255
         Left            =   5640
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblNombre 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblCodigo 
         Caption         =   "Código:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
   Begin Sicmact.Usuario user 
      Left            =   240
      Top             =   5520
      _ExtentX        =   820
      _ExtentY        =   820
   End
End
Attribute VB_Name = "frmViaticoSolReimpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oArendir As NARendir

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Public Sub CargaDatosPers(psPersCod As String)
    If psPersCod = "" Then
         Exit Sub
    End If
    cmdReimprimir.Enabled = False
    LimpiaFlex fgSolViaticos
        user.DatosPers psPersCod
        txtBuscaPers.Text = user.PersCod
        lblPersNombre = PstaNombre(user.UserNom)
        lblNrodoc = user.NroDNIUser
        lblAreaCod = user.AreaCod
        lblAreaDesc = user.AreaNom
        
        Dim RsSolicitudes  As ADODB.Recordset
        Set RsSolicitudes = New ADODB.Recordset
        Set oArendir = New NARendir
        Set RsSolicitudes = oArendir.ObtenerSolicitudViaticosXPersona(user.PersCod)
        
        Dim lnFila As Long
        If Not RsSolicitudes.EOF And Not RsSolicitudes.BOF Then
            cmdReimprimir.Enabled = True
            Do While Not RsSolicitudes.EOF
                fgSolViaticos.AdicionaFila
                lnFila = fgSolViaticos.row
                fgSolViaticos.TextMatrix(lnFila, 1) = RsSolicitudes!cDocNro
                fgSolViaticos.TextMatrix(lnFila, 2) = RsSolicitudes!dDocFecha
                fgSolViaticos.TextMatrix(lnFila, 3) = Format(RsSolicitudes!nMovMonto, "#,#0.00")
                fgSolViaticos.TextMatrix(lnFila, 4) = RsSolicitudes!cMovNro
                RsSolicitudes.MoveNext
            Loop
            fgSolViaticos.SetFocus
        Else
            MsgBox "No existen Solicitudes de viaticos a nombre de " & PstaNombre(user.UserNom), vbInformation, "Aviso"
            cmdReimprimir.Enabled = False
        End If
End Sub

Private Sub cmdCancelar_Click()
    LimpiaFlex fgSolViaticos
    txtBuscaPers.Text = ""
    lblPersNombre = ""
    lblNrodoc = ""
    lblAreaCod = ""
    lblAreaDesc = ""
    cmdReimprimir.Enabled = False
End Sub

Private Sub cmdReimprimir_Click()
    Dim oImp As NContImprimir
    Dim lsImpresion As String
    Set oImp = New NContImprimir
    
    lsImpresion = ImprimirReciboViaticoData(gnColPage, gdFecSis, gsOpeCod, _
                                            gsInstCmac, gsSimbolo, _
                                            gsNomCmac, gsNomCmacRUC, fgSolViaticos.TextMatrix(fgSolViaticos.row, 4))
        
    EnviaPrevio lsImpresion, Me.Caption, gnLinPage
    
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Reimprimio la Solicitud de Viaticos  del Colaborador : " & lblPersNombre
            Set objPista = Nothing
            '*******
End Sub

Private Sub Form_Load()
    CentraForm Me
    cmdReimprimir.Enabled = False
End Sub

Private Sub txtBuscaPers_EmiteDatos()
    Call CargaDatosPers(txtBuscaPers)
End Sub
