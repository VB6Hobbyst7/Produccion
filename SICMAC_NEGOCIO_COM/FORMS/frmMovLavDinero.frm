VERSION 5.00
Begin VB.Form frmMovLavDinero 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9300
   ClientLeft      =   4470
   ClientTop       =   1380
   ClientWidth     =   6600
   ForeColor       =   &H8000000D&
   Icon            =   "frmMovLavDinero.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   Begin SICMACT.Usuario ctlUsuario 
      Left            =   240
      Top             =   8760
      _extentx        =   820
      _extenty        =   820
   End
   Begin VB.TextBox txtOrigen 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      MaxLength       =   90
      TabIndex        =   20
      Top             =   8400
      Width           =   4575
   End
   Begin VB.Frame fraVisto 
      Caption         =   " Visto Electrónico "
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
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   7200
      Width           =   6435
      Begin VB.TextBox TxtClave 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C16A0B&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   17
         ToolTipText     =   "Ingrese su Clave Secreta"
         Top             =   615
         Width           =   2430
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   16
         Top             =   240
         Width           =   1000
      End
      Begin VB.TextBox txtUsuario 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C16A0B&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   15
         TabIndex        =   15
         ToolTipText     =   "Ingrese su Usuario"
         Top             =   240
         Width           =   2430
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario :"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   255
         Width           =   630
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clave     :"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   630
         Width           =   675
      End
   End
   Begin VB.Frame fraPersona 
      Caption         =   "Persona Realiza Transaccion"
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
      Height          =   2385
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   6435
      Begin VB.CheckBox chkRealiza 
         Caption         =   "No Presencial"
         Height          =   195
         Left            =   3200
         TabIndex        =   21
         Top             =   1973
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddPersRealiza 
         Caption         =   "&Agregar Persona"
         Height          =   300
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelPersRealiza 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1920
         TabIndex        =   11
         Top             =   1920
         Width           =   1140
      End
      Begin SICMACT.FlexEdit grdRealiza 
         Height          =   1605
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   6210
         _extentx        =   10954
         _extenty        =   2831
         cols0           =   9
         highlight       =   1
         allowuserresizing=   3
         visiblepopmenu  =   -1
         encabezadosnombres=   "#-Codigo-Doc Id-Nombre-Direccion-Ocupacion-Nacionalidad-Residente-PEPS"
         encabezadosanchos=   "250-1300-1300-2500-2000-3000-1200-1000-600"
         font            =   "frmMovLavDinero.frx":030A
         font            =   "frmMovLavDinero.frx":0332
         font            =   "frmMovLavDinero.frx":035A
         font            =   "frmMovLavDinero.frx":0382
         font            =   "frmMovLavDinero.frx":03AA
         fontfixed       =   "frmMovLavDinero.frx":03D2
         lbultimainstancia=   -1
         tipobusqueda    =   3
         columnasaeditar =   "X-1-X-X-X-5-X-X-X"
         textstylefixed  =   4
         listacontroles  =   "0-1-0-0-0-3-0-0-0"
         encabezadosalineacion=   "C-L-L-L-L-L-L-C-C"
         formatosedit    =   "0-0-0-0-0-0-0-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbflexduplicados=   0
         colwidth0       =   255
         rowheight0      =   300
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Persona Ordena Transaccion"
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
      Height          =   2385
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   6435
      Begin VB.CheckBox chkOrdena 
         Caption         =   "No Presencial"
         Height          =   195
         Left            =   3200
         TabIndex        =   22
         Top             =   1973
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddPersOrdena 
         Caption         =   "&Agregar Persona"
         Height          =   300
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelPersOrdena 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1920
         TabIndex        =   7
         Top             =   1920
         Width           =   1140
      End
      Begin SICMACT.FlexEdit grdOrdena 
         Height          =   1605
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   6210
         _extentx        =   10954
         _extenty        =   2831
         cols0           =   10
         highlight       =   1
         allowuserresizing=   3
         visiblepopmenu  =   -1
         encabezadosnombres=   "#-Codigo-Doc Id-Nombre-Direccion-Ocupacion-Nacionalidad-Residente-PEPS-Pers"
         encabezadosanchos=   "250-1300-1300-2500-2000-3000-1200-1000-600-1000"
         font            =   "frmMovLavDinero.frx":03F8
         font            =   "frmMovLavDinero.frx":0420
         font            =   "frmMovLavDinero.frx":0448
         font            =   "frmMovLavDinero.frx":0470
         font            =   "frmMovLavDinero.frx":0498
         fontfixed       =   "frmMovLavDinero.frx":04C0
         lbultimainstancia=   -1
         tipobusqueda    =   3
         columnasaeditar =   "X-1-X-X-X-5-X-X-X-X"
         textstylefixed  =   4
         listacontroles  =   "0-1-0-0-0-3-0-0-0-0"
         encabezadosalineacion=   "C-L-L-L-L-L-L-C-C-C"
         formatosedit    =   "0-0-0-0-0-0-0-0-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbflexduplicados=   0
         colwidth0       =   255
         rowheight0      =   300
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Persona Beneficia Transaccion"
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
      Height          =   2385
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   6435
      Begin VB.CheckBox chkBeneficia 
         Caption         =   "No Presencial"
         Height          =   195
         Left            =   3200
         TabIndex        =   24
         Top             =   2000
         Width           =   1335
      End
      Begin VB.CommandButton cmdVisitasEliminar 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1890
         TabIndex        =   23
         Top             =   1920
         Width           =   1140
      End
      Begin VB.CommandButton cmdAgregarPers 
         Caption         =   "&Agregar Persona"
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   1575
      End
      Begin SICMACT.FlexEdit grdBeneficiario 
         Height          =   1605
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   6210
         _extentx        =   10954
         _extenty        =   2831
         cols0           =   10
         highlight       =   1
         allowuserresizing=   3
         visiblepopmenu  =   -1
         encabezadosnombres=   "#-Codigo-Doc Id-Nombre-Direccion-Ocupacion-Nacionalidad-Residente-PEPS-Pers"
         encabezadosanchos=   "250-1300-1300-2500-2000-3000-1200-1000-600-1000"
         font            =   "frmMovLavDinero.frx":04E6
         font            =   "frmMovLavDinero.frx":050E
         font            =   "frmMovLavDinero.frx":0536
         font            =   "frmMovLavDinero.frx":055E
         font            =   "frmMovLavDinero.frx":0586
         fontfixed       =   "frmMovLavDinero.frx":05AE
         lbultimainstancia=   -1
         tipobusqueda    =   3
         columnasaeditar =   "X-1-X-X-X-5-X-X-X-X"
         textstylefixed  =   4
         listacontroles  =   "0-1-0-0-0-3-0-0-0-0"
         encabezadosalineacion=   "C-L-L-L-L-L-L-C-C-C"
         formatosedit    =   "0-0-0-0-0-0-0-0-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbflexduplicados=   0
         colwidth0       =   255
         rowheight0      =   300
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   8880
      Width           =   1000
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   8880
      Width           =   1000
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Origen del Efectivo: "
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   8400
      Width           =   1440
   End
End
Attribute VB_Name = "frmMovLavDinero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCodPersona As String
Dim bAbonaCuenta As Boolean
Dim sTitNombre As String, sTitDocId As String, sTitDireccion As String
Dim nMontoTransaccion As Double
Dim lnTC As Double
Dim lnTipoREU As Integer 'JACA 20110225
Dim sCuenta As String, sOperacion As String
Dim sTipoCuenta As String
Dim bImprimeBoleta As Boolean
Dim nMoneda As Integer
Dim sMoneda As String * 20
Dim sVisPersCod As String   'DAOR 20070511, Codigo de persona que da el visto
Dim lrs As ADODB.Recordset  'By Capi 28022008
Dim foPersona As COMDPersona.UCOMPersona   'WIOR 20130301
Private fTitPersLavDinero As String
Private fTitPersLavDineroNom As String
Private fTitPersLavDineroDir As String
Private fTitPersLavDineroDoc As String
Private fTitPersLavDineroOcu As String 'madm 20100722

Private fOrdPersLavDinero As String
Private fOrdPersLavDineroNom As String
Private fOrdPersLavDineroDir As String
Private fOrdPersLavDineroDoc As String
Private fOrdPersLavDineroOcu As String 'madm 20100722
Private lsOrdPersLavDineroOcu As String 'JACA 20110225
Private fOrdPersLavDineroNac As String 'WIOR 20130301
Private fOrdPersLavDineroRes As String 'WIOR 20130301
Private fOrdPersLavDineroPeps As String 'WIOR 20130301
Private fOrdPersLavDineroPers As String 'NDXERS0062020

'EJVG20120327 *****************************************
Private fOrdPersLavDinero2 As String
Private fOrdPersLavDineroNom2 As String
Private fOrdPersLavDineroDir2 As String
Private fOrdPersLavDineroDoc2 As String
Private fOrdPersLavDineroOcu2 As String
Private lsOrdPersLavDineroOcu2 As String
Private fOrdPersLavDineroNac2 As String 'WIOR 20130301
Private fOrdPersLavDineroRes2 As String 'WIOR 20130301
Private fOrdPersLavDineroPeps2 As String 'WIOR 20130301
Private fOrdPersLavDineroPers2 As String 'NDXERS0062020

Private fOrdPersLavDinero3 As String
Private fOrdPersLavDineroNom3 As String
Private fOrdPersLavDineroDir3 As String
Private fOrdPersLavDineroDoc3 As String
Private fOrdPersLavDineroOcu3 As String
Private lsOrdPersLavDineroOcu3 As String
Private fOrdPersLavDineroNac3 As String 'WIOR 20130301
Private fOrdPersLavDineroRes3 As String 'WIOR 20130301
Private fOrdPersLavDineroPeps3 As String 'WIOR 20130301
Private fOrdPersLavDineroPers3 As String 'NDXERS0062020

Private fOrdPersLavDinero4 As String
Private fOrdPersLavDineroNom4 As String
Private fOrdPersLavDineroDir4 As String
Private fOrdPersLavDineroDoc4 As String
Private fOrdPersLavDineroOcu4 As String
Private lsOrdPersLavDineroOcu4 As String
Private fOrdPersLavDineroNac4 As String 'WIOR 20130301
Private fOrdPersLavDineroRes4 As String 'WIOR 20130301
Private fOrdPersLavDineroPeps4 As String 'WIOR 20130301
Private fOrdPersLavDineroPers4 As String 'NDXERS0062020
'END EJVG**********************************************

Private fReaPersLavDinero As String
Private fReaPersLavDineroNom As String
Private fReaPersLavDineroDir As String
Private fReaPersLavDineroDoc As String
Private fReaPersLavDineroOcu As String 'madm 20100722
Private lsReaPersLavDineroOcu As String 'JACA 20110225
Private fReaPersLavDineroNac As String 'WIOR 20130301
Private fReaPersLavDineroRes As String 'WIOR 20130301
Private fReaPersLavDineroPeps As String 'WIOR 20130301

'EJVG20120327 *****************************************
Private fReaPersLavDinero2 As String
Private fReaPersLavDineroNom2 As String
Private fReaPersLavDineroDir2 As String
Private fReaPersLavDineroDoc2 As String
Private fReaPersLavDineroOcu2 As String
Private lsReaPersLavDineroOcu2 As String
Private fReaPersLavDineroNac2 As String 'WIOR 20130301
Private fReaPersLavDineroRes2 As String 'WIOR 20130301
Private fReaPersLavDineroPeps2 As String 'WIOR 20130301

Private fReaPersLavDinero3 As String
Private fReaPersLavDineroNom3 As String
Private fReaPersLavDineroDir3 As String
Private fReaPersLavDineroDoc3 As String
Private fReaPersLavDineroOcu3 As String
Private lsReaPersLavDineroOcu3 As String
Private fReaPersLavDineroNac3 As String 'WIOR 20130301
Private fReaPersLavDineroRes3 As String 'WIOR 20130301
Private fReaPersLavDineroPeps3 As String 'WIOR 20130301

Private fReaPersLavDinero4 As String
Private fReaPersLavDineroNom4 As String
Private fReaPersLavDineroDir4 As String
Private fReaPersLavDineroDoc4 As String
Private fReaPersLavDineroOcu4 As String
Private lsReaPersLavDineroOcu4 As String
Private fReaPersLavDineroNac4 As String 'WIOR 20130301
Private fReaPersLavDineroRes4 As String 'WIOR 20130301
Private fReaPersLavDineroPeps4 As String 'WIOR 20130301
'END EJVG**********************************************

Private fBenPersLavDinero As String
Private fBenPersLavDineroNom As String
Private fBenPersLavDineroDir As String
Private fBenPersLavDineroDoc As String
Private fBenPersLavDineroOcu As String 'madm 20100722
Private lsBenPersLavDineroOcu As String 'JACA 20110225
Private fBenPersLavDineroNac As String 'WIOR 20130301
Private fBenPersLavDineroRes As String 'WIOR 20130301
Private fBenPersLavDineroPeps As String 'WIOR 20130301
Private fBenPersLavDineroPers As String 'NDXERS0062020

'JACA 20110223
Private fBenPersLavDinero2  As String
Private fBenPersLavDineroNom2 As String
Private fBenPersLavDineroDir2 As String
Private fBenPersLavDineroDoc2 As String
Private fBenPersLavDineroOcu2 As String
Private lsBenPersLavDineroOcu2 As String
Private fBenPersLavDineroNac2 As String 'WIOR 20130301
Private fBenPersLavDineroRes2 As String 'WIOR 20130301
Private fBenPersLavDineroPeps2 As String 'WIOR 20130301
Private fBenPersLavDineroPers2 As String 'NDXERS0062020

Private fBenPersLavDinero3 As String
Private fBenPersLavDineroNom3 As String
Private fBenPersLavDineroDir3 As String
Private fBenPersLavDineroDoc3 As String
Private fBenPersLavDineroOcu3 As String
Private lsBenPersLavDineroOcu3 As String
Private fBenPersLavDineroNac3 As String 'WIOR 20130301
Private fBenPersLavDineroRes3 As String 'WIOR 20130301
Private fBenPersLavDineroPeps3 As String 'WIOR 20130301
Private fBenPersLavDineroPers3 As String 'NDXERS0062020

Private fBenPersLavDinero4 As String
Private fBenPersLavDineroNom4 As String
Private fBenPersLavDineroDir4 As String
Private fBenPersLavDineroDoc4 As String
Private fBenPersLavDineroOcu4 As String
Private lsBenPersLavDineroOcu4 As String
Private fBenPersLavDineroNac4 As String 'WIOR 20130301
Private fBenPersLavDineroRes4 As String 'WIOR 20130301
Private fBenPersLavDineroPeps4 As String 'WIOR 20130301
Private fBenPersLavDineroPers4 As String 'NDXERS0062020

Private fOrigenPersLavDinero As String
Private fnNroREU As String
Private fbImprimir As Boolean 'JACA 20110930
Private fsCtaImprimir As String 'JACA 20110930

'JACA END
Private fVisPersLavDinero As String
Private lnTipoSalida  As Integer
'madm 20100722
Property Let TitPersLavDineroOcu(pOcupa As String)
   fTitPersLavDineroOcu = pOcupa
End Property
Property Get TitPersLavDineroOcu() As String
    TitPersLavDineroDoc = fTitPersLavDineroOcu
End Property

Property Let OrdPersLavDineroOcu(pOcupa As String)
   fOrdPersLavDineroOcu = pOcupa
End Property
Property Get OrdPersLavDineroOcu() As String
    TitPersLavDineroDoc = fOrdPersLavDineroOcu
End Property

Property Let ReaPersLavDineroOcu(pOcupa As String)
   fReaPersLavDineroOcu = pOcupa
End Property
Property Get ReaPersLavDineroOcu() As String
    TitPersLavDineroDoc = fReaPersLavDineroOcu
End Property

Property Let BenPersLavDineroOcu(pOcupa As String)
   fBenPersLavDineroOcu = pOcupa
End Property
Property Get BenPersLavDineroOcu() As String
    TitPersLavDineroDoc = fBenPersLavDineroOcu
End Property

'end madm

'JACA 20110223
Property Let BenPersLavDineroOcu2(pOcupa As String)
   fBenPersLavDineroOcu2 = pOcupa
End Property
Property Get BenPersLavDineroOcu2() As String
    BenPersLavDineroOcu2 = fBenPersLavDineroOcu2
End Property
Property Let BenPersLavDineroOcu3(pOcupa As String)
   fBenPersLavDineroOcu3 = pOcupa
End Property
Property Get BenPersLavDineroOcu3() As String
    BenPersLavDineroOcu3 = fBenPersLavDineroOcu3
End Property
Property Let BenPersLavDineroOcu4(pOcupa As String)
   fBenPersLavDineroOcu4 = pOcupa
End Property
Property Get BenPersLavDineroOcu4() As String
    BenPersLavDineroOcu4 = fBenPersLavDineroOcu4
End Property
Property Let OrigenPersLavDinero(pOcupa As String)
   fOrigenPersLavDinero = pOcupa
End Property

Property Get OrigenPersLavDinero() As String
    OrigenPersLavDinero = fOrigenPersLavDinero
End Property

Property Let NroREU(pNroReu As String)
   fnNroREU = pNroReu
End Property

Property Get NroREU() As String
    NroREU = fnNroREU
End Property

'JACA END

'JACA 20110930**********************************
Property Let Imprimir(pbImprimir As Boolean)
   fbImprimir = pbImprimir
End Property

Property Get Imprimir() As Boolean
    Imprimir = fbImprimir
End Property
Property Let CtaImprimir(psCtaImprimir As String)
   fsCtaImprimir = psCtaImprimir
End Property

Property Get CtaImprimir() As String
    CtaImprimir = fsCtaImprimir
End Property
'JACA END***************************************


Property Let TitPersLavDinero(pPersCod As String)
   fTitPersLavDinero = pPersCod
End Property
Property Get TitPersLavDinero() As String
    TitPersLavDinero = fTitPersLavDinero
End Property

Property Let TitPersLavDineroNom(pNombre As String)
   fTitPersLavDineroNom = pNombre
End Property
Property Get TitPersLavDineroNom() As String
    TitPersLavDineroNom = fTitPersLavDineroNom
End Property

Property Let TitPersLavDineroDir(pDireccion As String)
   fTitPersLavDineroDir = pDireccion
End Property
Property Get TitPersLavDineroDir() As String
    TitPersLavDineroDir = fTitPersLavDineroDir
End Property

Property Let TitPersLavDineroDoc(pDocumento As String)
   fTitPersLavDineroDoc = pDocumento
End Property
Property Get TitPersLavDineroDoc() As String
    TitPersLavDineroDoc = fTitPersLavDineroDoc
End Property


Property Let OrdPersLavDinero(pPersCod As String)
   fOrdPersLavDinero = pPersCod
   'TxtCodigo(0).Text = pPersCod'EJVG20120327
End Property
Property Get OrdPersLavDinero() As String
    OrdPersLavDinero = fOrdPersLavDinero
End Property

Property Let OrdPersLavDineroNom(pNombre As String)
   fOrdPersLavDineroNom = pNombre
End Property
Property Get OrdPersLavDineroNom() As String
    OrdPersLavDineroNom = fOrdPersLavDineroNom
End Property

Property Let OrdPersLavDineroDir(pDireccion As String)
   fOrdPersLavDineroDir = pDireccion
End Property
Property Get OrdPersLavDineroDir() As String
    OrdPersLavDineroDir = fOrdPersLavDineroDir
End Property

Property Let OrdPersLavDineroDoc(pDocumento As String)
   fOrdPersLavDineroDoc = pDocumento
End Property
Property Get OrdPersLavDineroDoc() As String
    OrdPersLavDineroDoc = fOrdPersLavDineroDoc
End Property

'WIOR 20130301 - PERSONA ORDENA 1****************************************
'Nacionalidad
Property Let OrdPersLavDineroNac(pcNac As String)
   fOrdPersLavDineroNac = pcNac
End Property
Property Get OrdPersLavDineroNac() As String
    OrdPersLavDineroNac = fOrdPersLavDineroNac
End Property
'Residente
Property Let OrdPersLavDineroRes(pcRes As String)
   fOrdPersLavDineroRes = pcRes
End Property
Property Get OrdPersLavDineroRes() As String
    OrdPersLavDineroRes = fOrdPersLavDineroRes
End Property
'PEPS
Property Let OrdPersLavDineroPeps(pcPeps As String)
   fOrdPersLavDineroPeps = pcPeps
End Property
Property Get OrdPersLavDineroPeps() As String
    OrdPersLavDineroPeps = fOrdPersLavDineroPeps
End Property
'WIOR FIN ************************************************************

'EJVG20120327
'ORDENA 2
Property Let OrdPersLavDinero2(pPersCod As String)
   fOrdPersLavDinero2 = pPersCod
End Property
Property Get OrdPersLavDinero2() As String
    OrdPersLavDinero2 = fOrdPersLavDinero2
End Property

Property Let OrdPersLavDineroNom2(pNombre As String)
   fOrdPersLavDineroNom2 = pNombre
End Property
Property Get OrdPersLavDineroNom2() As String
    OrdPersLavDineroNom2 = fOrdPersLavDineroNom2
End Property

Property Let OrdPersLavDineroDir2(pDireccion As String)
   fOrdPersLavDineroDir2 = pDireccion
End Property
Property Get OrdPersLavDineroDir2() As String
    OrdPersLavDineroDir2 = fOrdPersLavDineroDir2
End Property

Property Let OrdPersLavDineroDoc2(pDocumento As String)
   fOrdPersLavDineroDoc2 = pDocumento
End Property
Property Get OrdPersLavDineroDoc2() As String
    OrdPersLavDineroDoc2 = fOrdPersLavDineroDoc2
End Property

Property Let OrdPersLavDineroOcu2(pDocumento As String)
    fOrdPersLavDineroOcu2 = pDocumento
End Property
Property Get OrdPersLavDineroOcu2() As String
    OrdPersLavDineroOcu2 = fOrdPersLavDineroOcu2
End Property
'WIOR 20130301 - PERSONA ORDENA 2****************************************
'Nacionalidad
Property Let OrdPersLavDineroNac2(pcNac As String)
   fOrdPersLavDineroNac2 = pcNac
End Property
Property Get OrdPersLavDineroNac2() As String
    OrdPersLavDineroNac2 = fOrdPersLavDineroNac2
End Property
'Residente
Property Let OrdPersLavDineroRes2(pcRes As String)
   fOrdPersLavDineroRes2 = pcRes
End Property
Property Get OrdPersLavDineroRes2() As String
    OrdPersLavDineroRes2 = fOrdPersLavDineroRes2
End Property
'PEPS
Property Let OrdPersLavDineroPeps2(pcPeps As String)
   fOrdPersLavDineroPeps2 = pcPeps
End Property
Property Get OrdPersLavDineroPeps2() As String
    OrdPersLavDineroPeps2 = fOrdPersLavDineroPeps2
End Property
'WIOR FIN ************************************************************

'ORDENA 3
Property Let OrdPersLavDinero3(pPersCod As String)
   fOrdPersLavDinero3 = pPersCod
End Property
Property Get OrdPersLavDinero3() As String
    OrdPersLavDinero3 = fOrdPersLavDinero3
End Property

Property Let OrdPersLavDineroNom3(pNombre As String)
   fOrdPersLavDineroNom3 = pNombre
End Property
Property Get OrdPersLavDineroNom3() As String
    OrdPersLavDineroNom3 = fOrdPersLavDineroNom3
End Property

Property Let OrdPersLavDineroDir3(pDireccion As String)
   fOrdPersLavDineroDir3 = pDireccion
End Property
Property Get OrdPersLavDineroDir3() As String
    OrdPersLavDineroDir3 = fOrdPersLavDineroDir3
End Property

Property Let OrdPersLavDineroDoc3(pDocumento As String)
   fOrdPersLavDineroDoc3 = pDocumento
End Property
Property Get OrdPersLavDineroDoc3() As String
    OrdPersLavDineroDoc3 = fOrdPersLavDineroDoc3
End Property

Property Let OrdPersLavDineroOcu3(pDocumento As String)
    fOrdPersLavDineroOcu3 = pDocumento
End Property
Property Get OrdPersLavDineroOcu3() As String
    OrdPersLavDineroOcu3 = fOrdPersLavDineroOcu3
End Property
'WIOR 20130301 - PERSONA ORDENA 3*************************************
'Nacionalidad
Property Let OrdPersLavDineroNac3(pcNac As String)
   fOrdPersLavDineroNac3 = pcNac
End Property
Property Get OrdPersLavDineroNac3() As String
    OrdPersLavDineroNac3 = fOrdPersLavDineroNac3
End Property
'Residente
Property Let OrdPersLavDineroRes3(pcRes As String)
   fOrdPersLavDineroRes3 = pcRes
End Property
Property Get OrdPersLavDineroRes3() As String
    OrdPersLavDineroRes3 = fOrdPersLavDineroRes3
End Property
'PEPS
Property Let OrdPersLavDineroPeps3(pcPeps As String)
   fOrdPersLavDineroPeps3 = pcPeps
End Property
Property Get OrdPersLavDineroPeps3() As String
    OrdPersLavDineroPeps3 = fOrdPersLavDineroPeps3
End Property
'WIOR FIN ************************************************************
'ORDENA 4
Property Let OrdPersLavDinero4(pPersCod As String)
   fOrdPersLavDinero4 = pPersCod
End Property
Property Get OrdPersLavDinero4() As String
    OrdPersLavDinero4 = fOrdPersLavDinero4
End Property

Property Let OrdPersLavDineroNom4(pNombre As String)
   fOrdPersLavDineroNom4 = pNombre
End Property
Property Get OrdPersLavDineroNom4() As String
    OrdPersLavDineroNom4 = fOrdPersLavDineroNom4
End Property

Property Let OrdPersLavDineroDir4(pDireccion As String)
   fOrdPersLavDineroDir4 = pDireccion
End Property
Property Get OrdPersLavDineroDir4() As String
    OrdPersLavDineroDir4 = fOrdPersLavDineroDir4
End Property

Property Let OrdPersLavDineroDoc4(pDocumento As String)
   fOrdPersLavDineroDoc4 = pDocumento
End Property
Property Get OrdPersLavDineroDoc4() As String
    OrdPersLavDineroDoc4 = fOrdPersLavDineroDoc4
End Property

Property Let OrdPersLavDineroOcu4(pDocumento As String)
    fOrdPersLavDineroOcu4 = pDocumento
End Property
Property Get OrdPersLavDineroOcu4() As String
    OrdPersLavDineroOcu4 = fOrdPersLavDineroOcu4
End Property
'END EJVG
'WIOR 20130301 - PERSONA ORDENA 4*************************************
'Nacionalidad
Property Let OrdPersLavDineroNac4(pcNac As String)
   fOrdPersLavDineroNac4 = pcNac
End Property
Property Get OrdPersLavDineroNac4() As String
    OrdPersLavDineroNac4 = fOrdPersLavDineroNac4
End Property
'Residente
Property Let OrdPersLavDineroRes4(pcRes As String)
   fOrdPersLavDineroRes4 = pcRes
End Property
Property Get OrdPersLavDineroRes4() As String
    OrdPersLavDineroRes4 = fOrdPersLavDineroRes4
End Property
'PEPS
Property Let OrdPersLavDineroPeps4(pcPeps As String)
   fOrdPersLavDineroPeps4 = pcPeps
End Property
Property Get OrdPersLavDineroPeps4() As String
    OrdPersLavDineroPeps4 = fOrdPersLavDineroPeps4
End Property
'WIOR FIN ************************************************************

Property Let ReaPersLavDinero(pPersCod As String)
   fReaPersLavDinero = pPersCod
   'TxtCodigo(1).Text = pPersCod'EJVG20120327
End Property
Property Get ReaPersLavDinero() As String
    ReaPersLavDinero = fReaPersLavDinero
End Property

Property Let ReaPersLavDineroNom(pNombre As String)
   fReaPersLavDineroNom = pNombre
End Property
Property Get ReaPersLavDineroNom() As String
    ReaPersLavDineroNom = fReaPersLavDineroNom
End Property

Property Let ReaPersLavDineroDir(pDireccion As String)
   fReaPersLavDineroDir = pDireccion
End Property
Property Get ReaPersLavDineroDir() As String
    ReaPersLavDineroDir = fReaPersLavDineroDir
End Property

Property Let ReaPersLavDineroDoc(pDocumento As String)
   fReaPersLavDineroDoc = pDocumento
End Property
Property Get ReaPersLavDineroDoc() As String
    ReaPersLavDineroDoc = fReaPersLavDineroDoc
End Property
'WIOR 20130301 - PERSONA REALIZA 1 ***********************************
'Nacionalidad
Property Let ReaPersLavDineroNac(pcNac As String)
   fReaPersLavDineroNac = pcNac
End Property
Property Get ReaPersLavDineroNac() As String
    ReaPersLavDineroNac = fReaPersLavDineroNac
End Property
'Residente
Property Let ReaPersLavDineroRes(pcRes As String)
   fReaPersLavDineroRes = pcRes
End Property
Property Get ReaPersLavDineroRes() As String
    ReaPersLavDineroRes = fReaPersLavDineroRes
End Property
'PEPS
Property Let ReaPersLavDineroPeps(pcPeps As String)
   fReaPersLavDineroPeps = pcPeps
End Property
Property Get ReaPersLavDineroPeps() As String
    ReaPersLavDineroPeps = fReaPersLavDineroPeps
End Property
'WIOR FIN ************************************************************

'EJVG20120327
'REALIZA 2
Property Let ReaPersLavDinero2(pPersCod As String)
   fReaPersLavDinero2 = pPersCod
End Property
Property Get ReaPersLavDinero2() As String
    ReaPersLavDinero2 = fReaPersLavDinero2
End Property

Property Let ReaPersLavDineroNom2(pNombre As String)
   fReaPersLavDineroNom2 = pNombre
End Property
Property Get ReaPersLavDineroNom2() As String
    ReaPersLavDineroNom2 = fReaPersLavDineroNom2
End Property

Property Let ReaPersLavDineroDir2(pDireccion As String)
   fReaPersLavDineroDir2 = pDireccion
End Property
Property Get ReaPersLavDineroDir2() As String
    ReaPersLavDineroDir2 = fReaPersLavDineroDir2
End Property

Property Let ReaPersLavDineroDoc2(pDocumento As String)
   fReaPersLavDineroDoc2 = pDocumento
End Property
Property Get ReaPersLavDineroDoc2() As String
    ReaPersLavDineroDoc2 = fReaPersLavDineroDoc2
End Property

Property Let ReaPersLavDineroOcu2(pDocumento As String)
   fReaPersLavDineroOcu2 = pDocumento
End Property
Property Get ReaPersLavDineroOcu2() As String
    ReaPersLavDineroOcu2 = fReaPersLavDineroOcu2
End Property
'WIOR 20130301 - PERSONA REALIZA 2 ***********************************
'Nacionalidad
Property Let ReaPersLavDineroNac2(pcNac As String)
   fReaPersLavDineroNac2 = pcNac
End Property
Property Get ReaPersLavDineroNac2() As String
    ReaPersLavDineroNac2 = fReaPersLavDineroNac2
End Property
'Residente
Property Let ReaPersLavDineroRes2(pcRes As String)
   fReaPersLavDineroRes2 = pcRes
End Property
Property Get ReaPersLavDineroRes2() As String
    ReaPersLavDineroRes2 = fReaPersLavDineroRes2
End Property
'PEPS
Property Let ReaPersLavDineroPeps2(pcPeps As String)
   fReaPersLavDineroPeps2 = pcPeps
End Property
Property Get ReaPersLavDineroPeps2() As String
    ReaPersLavDineroPeps2 = fReaPersLavDineroPeps2
End Property
'WIOR FIN ************************************************************
'REALIZA 3
Property Let ReaPersLavDinero3(pPersCod As String)
   fReaPersLavDinero3 = pPersCod
End Property
Property Get ReaPersLavDinero3() As String
    ReaPersLavDinero3 = fReaPersLavDinero3
End Property

Property Let ReaPersLavDineroNom3(pNombre As String)
   fReaPersLavDineroNom3 = pNombre
End Property
Property Get ReaPersLavDineroNom3() As String
    ReaPersLavDineroNom3 = fReaPersLavDineroNom3
End Property

Property Let ReaPersLavDineroDir3(pDireccion As String)
    fReaPersLavDineroDir3 = pDireccion
End Property
Property Get ReaPersLavDineroDir3() As String
    ReaPersLavDineroDir3 = fReaPersLavDineroDir3
End Property

Property Let ReaPersLavDineroDoc3(pDocumento As String)
    fReaPersLavDineroDoc3 = pDocumento
End Property
Property Get ReaPersLavDineroDoc3() As String
    ReaPersLavDineroDoc3 = fReaPersLavDineroDoc3
End Property

Property Let ReaPersLavDineroOcu3(pDocumento As String)
    fReaPersLavDineroOcu3 = pDocumento
End Property
Property Get ReaPersLavDineroOcu3() As String
    ReaPersLavDineroOcu3 = fReaPersLavDineroOcu3
End Property
'WIOR 20130301 - PERSONA REALIZA 3 ***********************************
'Nacionalidad
Property Let ReaPersLavDineroNac3(pcNac As String)
   fReaPersLavDineroNac3 = pcNac
End Property
Property Get ReaPersLavDineroNac3() As String
    ReaPersLavDineroNac3 = fReaPersLavDineroNac3
End Property
'Residente
Property Let ReaPersLavDineroRes3(pcRes As String)
   fReaPersLavDineroRes3 = pcRes
End Property
Property Get ReaPersLavDineroRes3() As String
    ReaPersLavDineroRes3 = fReaPersLavDineroRes3
End Property
'PEPS
Property Let ReaPersLavDineroPeps3(pcPeps As String)
   fReaPersLavDineroPeps3 = pcPeps
End Property
Property Get ReaPersLavDineroPeps3() As String
    ReaPersLavDineroPeps3 = fReaPersLavDineroPeps3
End Property
'WIOR FIN ************************************************************
'REALIZA 4
Property Let ReaPersLavDinero4(pPersCod As String)
   fReaPersLavDinero4 = pPersCod
End Property
Property Get ReaPersLavDinero4() As String
    ReaPersLavDinero4 = fReaPersLavDinero4
End Property

Property Let ReaPersLavDineroNom4(pNombre As String)
   fReaPersLavDineroNom4 = pNombre
End Property
Property Get ReaPersLavDineroNom4() As String
    ReaPersLavDineroNom4 = fReaPersLavDineroNom4
End Property

Property Let ReaPersLavDineroDir4(pDireccion As String)
   fReaPersLavDineroDir4 = pDireccion
End Property
Property Get ReaPersLavDineroDir4() As String
    ReaPersLavDineroDir4 = fReaPersLavDineroDir4
End Property

Property Let ReaPersLavDineroDoc4(pDocumento As String)
   fReaPersLavDineroDoc4 = pDocumento
End Property
Property Get ReaPersLavDineroDoc4() As String
    ReaPersLavDineroDoc4 = fReaPersLavDineroDoc4
End Property

Property Let ReaPersLavDineroOcu4(pDocumento As String)
    fReaPersLavDineroOcu4 = pDocumento
End Property
Property Get ReaPersLavDineroOcu4() As String
    ReaPersLavDineroOcu4 = fReaPersLavDineroOcu4
End Property
'END EJVG
'WIOR 20130301 - PERSONA REALIZA 4 ***********************************
'Nacionalidad
Property Let ReaPersLavDineroNac4(pcNac As String)
   fReaPersLavDineroNac4 = pcNac
End Property
Property Get ReaPersLavDineroNac4() As String
    ReaPersLavDineroNac4 = fReaPersLavDineroNac4
End Property
'Residente
Property Let ReaPersLavDineroRes4(pcRes As String)
   fReaPersLavDineroRes4 = pcRes
End Property
Property Get ReaPersLavDineroRes4() As String
    ReaPersLavDineroRes4 = fReaPersLavDineroRes4
End Property
'PEPS
Property Let ReaPersLavDineroPeps4(pcPeps As String)
   fReaPersLavDineroPeps4 = pcPeps
End Property
Property Get ReaPersLavDineroPeps4() As String
    ReaPersLavDineroPeps4 = fReaPersLavDineroPeps4
End Property
'WIOR FIN ************************************************************

Property Let BenPersLavDinero(pPersCod As String)
   fBenPersLavDinero = pPersCod
   'TxtCodigo(2).Text = pPersCod'EJVG20120327
End Property
Property Get BenPersLavDinero() As String
    BenPersLavDinero = fBenPersLavDinero
End Property

Property Let BenPersLavDineroNom(pNombre As String)
   fBenPersLavDineroNom = pNombre
End Property
Property Get BenPersLavDineroNom() As String
    BenPersLavDineroNom = fBenPersLavDineroNom
End Property

Public Property Let BenPersLavDineroDir(ByVal pDireccion As String)
   fBenPersLavDineroDir = pDireccion
End Property
Public Property Get BenPersLavDineroDir() As String
    BenPersLavDineroDir = fBenPersLavDineroDir
End Property

Property Let BenPersLavDineroDoc(pDocumento As String)
   fBenPersLavDineroDoc = pDocumento
End Property
Property Get BenPersLavDineroDoc() As String
    BenPersLavDineroDoc = fBenPersLavDineroDoc
End Property
'WIOR 20130301 - PERSONA BENEFICIA 1 *********************************
'Nacionalidad
Property Let BenPersLavDineroNac(pcNac As String)
   fBenPersLavDineroNac = pcNac
End Property
Property Get BenPersLavDineroNac() As String
    BenPersLavDineroNac = fBenPersLavDineroNac
End Property
'Residente
Property Let BenPersLavDineroRes(pcRes As String)
   fBenPersLavDineroRes = pcRes
End Property
Property Get BenPersLavDineroRes() As String
    BenPersLavDineroRes = fBenPersLavDineroRes
End Property
'PEPS
Property Let BenPersLavDineroPeps(pcPeps As String)
   fBenPersLavDineroPeps = pcPeps
End Property
Property Get BenPersLavDineroPeps() As String
    BenPersLavDineroPeps = fBenPersLavDineroPeps
End Property
'WIOR FIN ************************************************************
'JACA 20110223
'BENEFICIARIO 2
Property Let BenPersLavDinero2(pPersCod As String)
   fBenPersLavDinero2 = pPersCod
   'txtCodigo(2).Text = pPersCod
End Property
Property Get BenPersLavDinero2() As String
    BenPersLavDinero2 = fBenPersLavDinero2
End Property

Property Let BenPersLavDineroNom2(pNombre As String)
   fBenPersLavDineroNom2 = pNombre
End Property
Property Get BenPersLavDineroNom2() As String
    BenPersLavDineroNom2 = fBenPersLavDineroNom2
End Property

Public Property Let BenPersLavDineroDir2(ByVal pDireccion As String)
   fBenPersLavDineroDir2 = pDireccion
End Property
Public Property Get BenPersLavDineroDir2() As String
    BenPersLavDineroDir2 = fBenPersLavDineroDir2
End Property

Property Let BenPersLavDineroDoc2(pDocumento As String)
   fBenPersLavDineroDoc2 = pDocumento
End Property
Property Get BenPersLavDineroDoc2() As String
    BenPersLavDineroDoc2 = fBenPersLavDineroDoc2
End Property
'WIOR 20130301 - PERSONA BENEFICIA 2 *********************************
'Nacionalidad
Property Let BenPersLavDineroNac2(pcNac As String)
   fBenPersLavDineroNac2 = pcNac
End Property
Property Get BenPersLavDineroNac2() As String
    BenPersLavDineroNac2 = fBenPersLavDineroNac2
End Property
'Residente
Property Let BenPersLavDineroRes2(pcRes As String)
   fBenPersLavDineroRes2 = pcRes
End Property
Property Get BenPersLavDineroRes2() As String
    BenPersLavDineroRes2 = fBenPersLavDineroRes2
End Property
'PEPS
Property Let BenPersLavDineroPeps2(pcPeps As String)
   fBenPersLavDineroPeps2 = pcPeps
End Property
Property Get BenPersLavDineroPeps2() As String
    BenPersLavDineroPeps2 = fBenPersLavDineroPeps2
End Property
'WIOR FIN ************************************************************
'BENEFICIARIO 3
Property Let BenPersLavDinero3(pPersCod As String)
   fBenPersLavDinero3 = pPersCod
   'txtCodigo(2).Text = pPersCod
End Property
Property Get BenPersLavDinero3() As String
    BenPersLavDinero3 = fBenPersLavDinero3
End Property

Property Let BenPersLavDineroNom3(pNombre As String)
   fBenPersLavDineroNom3 = pNombre
End Property
Property Get BenPersLavDineroNom3() As String
    BenPersLavDineroNom3 = fBenPersLavDineroNom3
End Property

Public Property Let BenPersLavDineroDir3(ByVal pDireccion As String)
   fBenPersLavDineroDir3 = pDireccion
End Property
Public Property Get BenPersLavDineroDir3() As String
    BenPersLavDineroDir3 = fBenPersLavDineroDir3
End Property

Property Let BenPersLavDineroDoc3(pDocumento As String)
   fBenPersLavDineroDoc3 = pDocumento
End Property
Property Get BenPersLavDineroDoc3() As String
    BenPersLavDineroDoc3 = fBenPersLavDineroDoc3
End Property
'WIOR 20130301 - PERSONA BENEFICIA 3 *********************************
'Nacionalidad
Property Let BenPersLavDineroNac3(pcNac As String)
   fBenPersLavDineroNac3 = pcNac
End Property
Property Get BenPersLavDineroNac3() As String
    BenPersLavDineroNac3 = fBenPersLavDineroNac3
End Property
'Residente
Property Let BenPersLavDineroRes3(pcRes As String)
   fBenPersLavDineroRes3 = pcRes
End Property
Property Get BenPersLavDineroRes3() As String
    BenPersLavDineroRes3 = fBenPersLavDineroRes3
End Property
'PEPS
Property Let BenPersLavDineroPeps3(pcPeps As String)
   fBenPersLavDineroPeps3 = pcPeps
End Property
Property Get BenPersLavDineroPeps3() As String
    BenPersLavDineroPeps3 = fBenPersLavDineroPeps3
End Property
'WIOR FIN ************************************************************
'BENEFICIARIO 4
Property Let BenPersLavDinero4(pPersCod As String)
   fBenPersLavDinero4 = pPersCod
   'txtCodigo(2).Text = pPersCod
End Property
Property Get BenPersLavDinero4() As String
    BenPersLavDinero4 = fBenPersLavDinero4
End Property

Property Let BenPersLavDineroNom4(pNombre As String)
   fBenPersLavDineroNom4 = pNombre
End Property
Property Get BenPersLavDineroNom4() As String
    BenPersLavDineroNom4 = fBenPersLavDineroNom4
End Property

Public Property Let BenPersLavDineroDir4(ByVal pDireccion As String)
   fBenPersLavDineroDir4 = pDireccion
End Property
Public Property Get BenPersLavDineroDir4() As String
    BenPersLavDineroDir4 = fBenPersLavDineroDir4
End Property

Property Let BenPersLavDineroDoc4(pDocumento As String)
   fBenPersLavDineroDoc4 = pDocumento
End Property
Property Get BenPersLavDineroDoc4() As String
    BenPersLavDineroDoc4 = fBenPersLavDineroDoc4
End Property
'JACA END
'WIOR 20130301 - PERSONA BENEFICIA 4 *********************************
'Nacionalidad
Property Let BenPersLavDineroNac4(pcNac As String)
   fBenPersLavDineroNac4 = pcNac
End Property
Property Get BenPersLavDineroNac4() As String
    BenPersLavDineroNac4 = fBenPersLavDineroNac4
End Property
'Residente
Property Let BenPersLavDineroRes4(pcRes As String)
   fBenPersLavDineroRes4 = pcRes
End Property
Property Get BenPersLavDineroRes4() As String
    BenPersLavDineroRes4 = fBenPersLavDineroRes4
End Property
'PEPS
Property Let BenPersLavDineroPeps4(pcPeps As String)
   fBenPersLavDineroPeps4 = pcPeps
End Property
Property Get BenPersLavDineroPeps4() As String
    BenPersLavDineroPeps4 = fBenPersLavDineroPeps4
End Property
'WIOR FIN ************************************************************

Property Get VisPersLavDinero() As String
    VisPersLavDinero = fVisPersLavDinero
End Property

Public Function Inicia(Optional sPersCod As String = "", Optional sNombre As String = "", _
            Optional sDireccion As String, Optional sDocId As String, Optional bMovimiento As Boolean = False, _
            Optional bAbono As Boolean = False, Optional nMonto As Double = 0, Optional sCta As String = "", _
            Optional sOpe As String = "", Optional bImpBol As Boolean = True, Optional sTipo As String = "", _
            Optional sPersCodRea As String = "", Optional sNombreRea As String = "", Optional sDireccionRea As String = "", _
            Optional sDocIdRea As String = "", Optional ByVal pnMoneda As Integer = 0, _
            Optional bOtrEgreWest As Boolean = False, Optional nTipoREU As Integer = 1, Optional gnMontoAcumulado As Double = 0, Optional gsOrigen As String = "", Optional ByVal bNoOperacionPorLote As Boolean = True, _
            Optional ByVal psOpeCod As String = "000000", Optional ByVal pnPrograma As Integer = 0) As String
            'WIOR 20131106 AGREGO Optional ByVal psOpeCod As String = "000000", Optional ByVal pnPrograma As Integer = 0
    
    
    'By Capi 28022008
    Dim lsCuentaWestern As String
    Dim lnMontoOpeCompara As Double
    Dim lnMontoMulCompara As Double
    Dim lnMontoMultipleCliente As Double
    Dim lsPeriodo As String
    
    Dim nTC As Double
    Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim loAcumLavDinero As COMDPersona.DCOMPersonas
    'marg ers073***
    Dim clsLav2 As clases.DParametro
     Set clsLav2 = New clases.DParametro
    'end marg******
    
    Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Set loAcumLavDinero = New COMDPersona.DCOMPersonas
    
    'By Capi 20042008
    'Comentado Porque no permitia hacer retiros O/T GITU
'    If sPersCod <> "" Then
'        fTitPersLavDineroNom = sNombre
'        fTitPersLavDineroDoc = sDocId
'        fTitPersLavDineroDir = sDireccion
'    End If
'
'    Call obtenerDatosPersonasLavDinero
'    'If fTitPersLavDinero = ""  Then
'    'By Capi 05072008
'    If fTitPersLavDinero = "" And fReaPersLavDinero = "" Then
'        Exit Function
'    End If
    
    
    Dim clsTC As COMDConstSistema.NCOMTipoCambio 'WIOR20121107
    Set clsTC = New COMDConstSistema.NCOMTipoCambio 'WIOR20121107
    Dim nTCReu As Double 'WIOR20121107
    nTCReu = clsTC.EmiteTipoCambio(gdFecSis, TCPondREU) 'WIOR20121107
    lsCuentaWestern = ""
    If pnMoneda = gMonedaNacional Then
        'Dim clsTC As COMDConstSistema.NCOMTipoCambio'WIOR20121107 COMENTO ESTA PARTE
        'Set clsTC = New COMDConstSistema.NCOMTipoCambio
        'nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
        nTC = clsTC.EmiteTipoCambio(gdFecSis, TCPondREU)
        'Set clsTC = Nothing
    Else
        nTC = 1
    End If
    Set clsTC = Nothing 'WIOR20121107
    
    If nTC = 0 Then
        MsgBox "Tipo de cambio del REU no esta ingresado"
        fOrdPersLavDinero = ""
        fTitPersLavDinero = ""
        Inicia = ""
        Exit Function
    End If
    'By Capi 10032008
    Dim loConstSis As COMDConstSistema.NCOMConstSistema
    Set loConstSis = New COMDConstSistema.NCOMConstSistema
    
    If Mid(sCta, 9, 1) = "1" Then
        'lsCuentaWestern = loConstSis.LeeConstSistema(gCapCtaWesternLavDineroME) despues cambiar ojo
        lsCuentaWestern = loConstSis.LeeConstSistema(334)
    Else
        'lsCuentaWestern = loConstSis.LeeConstSistema(gCapCtaWesternLavDineroMN) ojo cambiar
        lsCuentaWestern = loConstSis.LeeConstSistema(333)
    End If
    Set loConstSis = Nothing
                
    If sCta = lsCuentaWestern Or bOtrEgreWest Then
        lnMontoOpeCompara = clsLav.GetCapParametro(gMonOpeWestLavDineroME)
        lnMontoMulCompara = clsLav.GetCapParametro(gMonMulWestLavDineroME)
    Else
        'WIOR 20131105 MEMORÁNDUM N° 169-2013-LV-DI-CMAC-M******************
        Dim EsSonada As Boolean
        Dim dCaptaciones As New COMDCaptaGenerales.DCOMCaptaMovimiento
        EsSonada = False
        
        If Len(Trim(sCta)) = 18 Then
            If Trim(Mid(sCta, 6, 3)) = "232" Then
                If dCaptaciones.GetTipoProgramaCaptacion(sCta) = 5 Then 'SOLO PARA CUENTAS SOÑADAS
                    EsSonada = True
                End If
            End If
        Else
            If Mid(Trim(psOpeCod), 1, 2) = "20" Then 'Solo entra ahorros
                If pnPrograma = 5 Then 'SOLO PARA CUENTAS SOÑADAS
                    EsSonada = True
                End If
            End If
        End If
        
        Set dCaptaciones = Nothing
        
'          'SOLO ENTRAN GIROS, CAMBIO DE MONEDA Y OPERACIONES DE CUENTAS SOÑADAS
'          If Mid(Trim(psOpecod), 1, 2) = "31" Or Trim(psOpecod) = "900022" Or Trim(psOpecod) = "900023" Or EsSonada_
'                Or Left(Trim(gsOpeCod), 2) = "93" Then 'PASI20160718 incluyo operaciones de la CCE.
        'If Mid(Trim(psOpeCod), 1, 2) = "31" Or Trim(psOpeCod) = "900022" Or Trim(psOpeCod) = "900023" Or EsSonada Then 'SOLO ENTRAN GIROS, CAMBIO DE MONEDA Y OPERACIONES DE CUENTAS SOÑADAS 'COMMENT BY MARG ERS073 ANEXO 02
        
'***COMMENT BY MARG TIC1711060007********************************************************
'''        If Mid(Trim(psOpeCod), 1, 2) = "31" Or Trim(psOpeCod) = "900022" Or Trim(psOpeCod) = "900023" Or (EsSonada And Left(Trim(psOpeCod), 4) <> "2009") Then   'SOLO ENTRAN GIROS, CAMBIO DE MONEDA Y OPERACIONES DE CUENTAS SOÑADAS 'COMMENT BY MARG ERS073 ANEXO 02
'''
'''             'WIOR 20131104 **********************************************
'''            If gsCodAge = "01" Then 'Oficina Principal
'''                lnMontoOpeCompara = clsLav.GetCapParametro(2133)
'''            ElseIf gsCodAge = "02" Then 'Agencia Huánuco
'''                lnMontoOpeCompara = clsLav.GetCapParametro(2134)
'''            ElseIf gsCodAge = "03" Then 'Agencia Pucallpa
'''                lnMontoOpeCompara = clsLav.GetCapParametro(2135)
'''            ElseIf gsCodAge = "04" Then 'Agencia Calle Arequipa
'''                lnMontoOpeCompara = clsLav.GetCapParametro(2136)
'''            ElseIf gsCodAge = "06" Then 'Agencia Yurimaguas
'''                lnMontoOpeCompara = clsLav.GetCapParametro(2137)
'''            ElseIf gsCodAge = "07" Then 'Agencia Tingo María
'''                lnMontoOpeCompara = clsLav.GetCapParametro(2138)
'''            ElseIf gsCodAge = "09" Then 'Agencia Belén
'''                lnMontoOpeCompara = clsLav.GetCapParametro(2139)
'''            ElseIf gsCodAge = "10" Then 'Agencia Tarapoto
'''                lnMontoOpeCompara = clsLav.GetCapParametro(2140)
'''            ElseIf gsCodAge = "12" Then 'Agencia Aguaytía
'''                lnMontoOpeCompara = clsLav.GetCapParametro(2141)
'''            ElseIf gsCodAge = "13" Then 'Agencia Requena
'''                lnMontoOpeCompara = clsLav.GetCapParametro(2142)
'''            ElseIf gsCodAge = "24" Then 'Agencia Cajamarca
'''                lnMontoOpeCompara = clsLav.GetCapParametro(2143)
'''            ElseIf gsCodAge = "25" Then 'Agencia Cerro de Pasco
'''                lnMontoOpeCompara = clsLav.GetCapParametro(2144)
'''            ElseIf gsCodAge = "31" Then 'Agencia Punchana
'''                lnMontoOpeCompara = clsLav.GetCapParametro(2145)
'''            ElseIf gsCodAge = "33" Then 'Agencia Minka
'''                lnMontoOpeCompara = clsLav.GetCapParametro(2146)
'''            Else 'Otros
'''                lnMontoOpeCompara = clsLav.GetCapParametro(gMonOpeNoWestLavDineroME)
'''            End If
'''            'WIOR FIN *******************************************************
'''        'MARG ERS073*****************************
'''        ElseIf Trim(psOpeCod) = "100225" Then ' LIQUIDACIÓN DE CREDITO CON EL SEG. DESGRAVAMEN
'''            lnMontoOpeCompara = clsLav2.GetParametro(1201, gPrdParamColoc)
'''        ElseIf Mid(Trim(psOpeCod), 1, 4) = "1202" Then ' Desembolso Pignoraticio
'''                lnMontoOpeCompara = clsLav2.GetParametro(1202, gPrdParamColoc)
'''        ElseIf Left(Trim(psOpeCod), 2) = "93" Then ' Transferencia Interbancaria Ordinaria - Originante
'''                lnMontoOpeCompara = clsLav2.GetParametro(1203, gPrdParamCaptac)
'''        'END MARG********************************
'''        'MARG ERS073-ANEXO02*******************
'''        ElseIf Left(Trim(psOpeCod), 4) = "2009" Then ' Transferencias entre cuentas
'''                lnMontoOpeCompara = clsLav2.GetParametro(1204, gPrdParamCaptac)
'''        'END MARG********************************
'''        Else
'''            ''MADM 20120206'WIOR 20131104 COMENTO TODO ESTA PARTE
'''            If gsCodAge = "02" Then
'''                lnMontoOpeCompara = clsLav.GetCapParametro(gMonOpeNoWestLavDineroME02)
'''            ElseIf gsCodAge = "07" Or gsCodAge = "25" Then
'''                lnMontoOpeCompara = clsLav.GetCapParametro(gMonOpeNoWestLavDineroME0725)
'''            ElseIf gsCodAge = "01" Or gsCodAge = "04" Or gsCodAge = "09" Or gsCodAge = "31" Then
'''                lnMontoOpeCompara = clsLav.GetCapParametro(2121)
'''            ElseIf gsCodAge = "13" Or gsCodAge = "03" Then
'''                lnMontoOpeCompara = clsLav.GetCapParametro(2122)
'''            ElseIf gsCodAge = "12" Then
'''                lnMontoOpeCompara = clsLav.GetCapParametro(2123)
'''            Else
'''                lnMontoOpeCompara = clsLav.GetCapParametro(gMonOpeNoWestLavDineroME)
'''            End If
'''            ''END MADM
'''        End If
'''        'WIOR ***************************************************************
'***END MARG********************************************************
            '***MARG TIC1711060007***************************************************
            lnMontoOpeCompara = clsLav2.GetUmbral(gsCodAge, psOpeCod, EsSonada)
            '***END MARG**************************************************************
            lnMontoMulCompara = clsLav.GetCapParametro(gMonMulNoWestLavDineroME)
    End If
    lnMontoMultipleCliente = 0
    lsPeriodo = Mid(gdFecSis, 7, 4) & Mid(gdFecSis, 4, 2)
    Dim nMontTCambMul As Double
    Dim nMontoTem As Double
    Dim nMontoAcTem As Double
    Dim nSalir As Double
    'Set lrs = loAcumLavDinero.ObtenerAcumuladoLavDinero(nTC, lsPeriodo, fTitPersLavDinero)'WIOR20121107 Comentado
    Set lrs = loAcumLavDinero.ObtenerAcumuladoLavDinero(nTCReu, lsPeriodo, fTitPersLavDinero) 'WIOR20121107
    If lrs.RecordCount > 0 Then
        If pnMoneda = gMonedaNacional Then
            nMontTCambMul = lrs!Monto + Round(nMonto / nTC, 2)
            nMontoTem = Round(nMonto / nTC, 2)
            nMontoAcTem = lrs!Monto
        Else
            nMontTCambMul = lrs!Monto + nMonto
            nMontoTem = nMonto
            nMontoAcTem = lrs!Monto
        End If
'ALPA 20081009*****************************************************************
        'If nMontoTem < Round(lnMontoOpeCompara * nTC, 2) Then
        If nMontoTem >= lnMontoOpeCompara Then
'''            If pnMoneda = gMonedaNacional Then
'''                'lnMontoMultipleCliente = lrs!Monto + Round(nMonto / nTC, 2)
'''                lnMontoMultipleCliente = Round(nMonto / nTC, 2)
'''            Else
'''                'lnMontoMultipleCliente = lrs!Monto + nMonto
'''                lnMontoMultipleCliente = nMonto
'''            End If
'            If lnMontoMultipleCliente < lnMontoMulCompara Then
'                fOrdPersLavDinero = "Exit"
'                fTitPersLavDinero = ""
'                Inicia = ""
'                Exit Function
'            End If
        'ElseIf lnMontoMultipleCliente < lnMontoMulCompara Then
        nSalir = 1
'****************************************************************************************
        'ElseIf nMontTCambMul >= lnMontoMulCompara And nMontoAcTem < lnMontoMulCompara Then
        ElseIf nMontTCambMul >= lnMontoMulCompara And nMonto > 0 Then 'WIOR20121107'WIOR 20130124 AGREGO nMonto > 0
'                fOrdPersLavDinero = "Exit"
'                fTitPersLavDinero = ""
'                Inicia = ""
'                Exit Function
                    lnMontoMultipleCliente = nMontTCambMul
            nSalir = 2
        End If
       ' End If
    Else
        If nMonto < Round(lnMontoOpeCompara * nTC, 2) Then
            If lnMontoMultipleCliente < lnMontoMulCompara Then
                fOrdPersLavDinero = "Exit"
                fTitPersLavDinero = ""
                Inicia = ""
                Exit Function
            End If
        End If
    
    End If
    
    'Actualizando en Persona
    Set loAcumLavDinero = Nothing
    Set clsLav = Nothing
    'End By
    If nSalir = 0 Then
        fOrdPersLavDinero = "Exit" ' Cuando la operacion no esta afecta al REU
        fTitPersLavDinero = ""
        Inicia = ""
        Exit Function
    End If
    
    
    
    nMoneda = pnMoneda
    If lnMontoMultipleCliente = 0 Then
        Me.Caption = "Lav de Dinero Oper. Individual - Valor Comparacion " & lnMontoOpeCompara
        nTipoREU = 1
        lnTipoREU = nTipoREU  'JACA 20110225
        
        'ALPA 20081003*******************************************
        gnMontoAcumulado = nMontoTem 'nMonto
        '**********************************************************
    Else
        Me.Caption = "Lav de Dinero Oper. Multiple - Valor Comparacion " & lnMontoMulCompara
        nTipoREU = 2
        lnTipoREU = nTipoREU 'JACA 20110225
        'ALPA 20081003*******************************************
        gnMontoAcumulado = lnMontoMultipleCliente
        '********************************************************
    End If
    sCodPersona = ""
    bAbonaCuenta = bAbono
    nMontoTransaccion = nMonto
    lnTC = nTC
    sCuenta = sCta
    If nMoneda = 0 Then
        sMoneda = "No Determinado"
    Else
        '''sMoneda = IIf(nmoneda = 1, "NUEVOS SOLES", "US DOLARES") 'marg ers044-2016
        sMoneda = IIf(nMoneda = 1, StrConv(gcPEN_PLURAL, vbUpperCase), "US DOLARES") 'marg ers044-2016
    End If
    sOperacion = sOpe
    cmdGrabar.Enabled = False
    bImprimeBoleta = bImpBol
    fbImprimir = bImpBol 'JACA 20110930 agregado en la variable formulario para saber si emvia imprimir
    fsCtaImprimir = sCta 'JACA 20110930 agregado por variable de formulario obtener el nro de cta
    sTipoCuenta = Trim(sTipo)
    
    'Descomentado x JACA 20110303
    If sPersCod <> "" Then
        fTitPersLavDineroNom = sNombre
        fTitPersLavDineroDoc = sDocId
        fTitPersLavDineroDir = sDireccion
    End If
    'END JACA
'    Call obtenerDatosPersonasLavDinero

    '***Modificado por ELRO el 20110915, según Acta 252-2011/TI-D
    'Me.Show 1  'comentado por ELRO el 20110915
    'Inicia = sCodPersona   'comentado por ELRO el 20110915
    If bNoOperacionPorLote = True Then
        Me.Show 1
        Inicia = sCodPersona
    Else
        fOrdPersLavDinero = ""
        fTitPersLavDinero = ""
        Inicia = "Aviso"
        Exit Function
    End If
        
End Function



'By Capi 20012008, Validar que el usuario que  da el visto electrónico tenga los permisos*

Private Sub CmdAceptar_Click()
Dim oAcceso As COMDPersona.UCOMAcceso
    Set oAcceso = New COMDPersona.UCOMAcceso
        '***Modificado por ELRO el 20120403
        'If Not oAcceso.VistoLavDineroEsCorrecto(txtUsuario.Text, TxtClave.Text, sVisPersCod) Then
        '    MsgBox ("Usuario y clave incorrecta o usted no tiene permisos para dar el visto bueno")
        'Else
        '    MsgBox ("Visto satisfactorio, proceda a registrar ")
        '    cmdGrabar.Enabled = True
        '    fVisPersLavDinero = sVisPersCod
        '    txtOrigen.SetFocus
        'End If
        If Len(Trim(txtUsuario)) > 0 And Len(Trim(TxtClave)) > 0 Then
            If Not ValidacionRFIII Then Exit Sub ' RIRO SEGUN TI-ERS108-2013
            If Not oAcceso.VistoLavDineroEsCorrecto(txtUsuario.Text, TxtClave.Text, sVisPersCod) Then
                MsgBox ("Usuario y clave incorrecta o usted no tiene permisos para dar el visto bueno")
            Else
                MsgBox ("Visto satisfactorio, proceda a registrar ")
                cmdGrabar.Enabled = True
                fVisPersLavDinero = sVisPersCod
                txtOrigen.SetFocus
            End If
        ElseIf Len(Trim(txtUsuario)) = 0 Then
            Call MsgBox("Falta ingresar su usuario", vbInformation, "Aviso")
            txtUsuario.SetFocus
            Set oAcceso = Nothing
            Exit Sub
        ElseIf Len(Trim(TxtClave)) = 0 Then
            Call MsgBox("Falta ingresar su clave", vbInformation, "Aviso")
            TxtClave.SetFocus
            Set oAcceso = Nothing
            Exit Sub
        End If
        '***Fin Modificado por ELRO********
    Set oAcceso = Nothing
End Sub
'EJVG20120327
Private Sub cmdAddPersOrdena_Click()
    If grdOrdena.Rows <= 4 Then
        grdOrdena.AdicionaFila
        grdOrdena.SetFocus
        cmdDelPersOrdena.Enabled = True
        SendKeys "{Enter}"
    Else
        MsgBox ("Solo se permite cuatro Personas que Ordenan como máximo")
    End If
End Sub
Private Sub cmdAddPersRealiza_Click()
    If grdRealiza.Rows <= 4 Then
        grdRealiza.AdicionaFila
        grdRealiza.SetFocus
        cmdDelPersRealiza.Enabled = True
        SendKeys "{Enter}"
    Else
        MsgBox ("Solo se permite cuatro Personas que Realizan como máximo")
    End If
End Sub
'END EJVG
'JACA 20110223
Private Sub cmdAgregarPers_Click()
    If grdBeneficiario.Rows <= 4 Then
        grdBeneficiario.AdicionaFila
        grdBeneficiario.SetFocus
        cmdVisitasEliminar.Enabled = True
        SendKeys "{Enter}"
    Else
        MsgBox ("Solo se permite cuatro Beneficiarios como máximo")
    End If

End Sub
'JACA END

Private Sub cmdCancelar_Click()
    lnTipoSalida = 0 'ALPA 20120305
    sCodPersona = ""
    fTitPersLavDinero = ""
    fOrdPersLavDinero = ""
    fReaPersLavDinero = ""
    fBenPersLavDinero = ""
    fVisPersLavDinero = ""
    Call Form_Unload(lnTipoSalida)
    'Unload Me
End Sub


Private Sub cmdDelPersOrdena_Click()
    If MsgBox("¿¿Está seguro de eliminar la selección actual??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        If grdOrdena.Rows = 2 Then
           cmdDelPersOrdena.Enabled = False
        End If
         grdOrdena.EliminaFila grdOrdena.row
    End If
End Sub

Private Sub cmdDelPersRealiza_Click()
    If MsgBox("¿¿Está seguro de eliminar la selección actual??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        If grdRealiza.Rows = 2 Then
           cmdDelPersRealiza.Enabled = False
        End If
         grdRealiza.EliminaFila grdRealiza.row
    End If
End Sub

Private Sub cmdGrabar_Click()
Dim oBoleta As COMNCaptaGenerales.NCOMCaptaImpresion 'NCapImpBoleta
Dim objREU As COMDPersona.DCOMPersonas
Dim lsBoleta As String
Dim N_Ocupa1 As Integer
Dim N_Ocupa2 As Integer

Dim N_Ocupa3 As Integer
Dim i As Integer
'JACA 20110223 AUMENTADO SOLO HASTA 4 BENEF.
Dim N_Ocupa3_2 As Integer
Dim N_Ocupa3_3 As Integer
Dim N_Ocupa3_4 As Integer
'JACA END
lnTipoSalida = 1
Set objREU = New COMDPersona.DCOMPersonas

'Comentado x EJVG20120327
'    fOrdPersLavDinero = Trim(TxtCodigo(0).Text)
'    fOrdPersLavDineroNom = lblNombre(0).Caption
'    fOrdPersLavDineroDir = lblDireccion(0).Caption
'    fOrdPersLavDineroDoc = lblDocID(0).Caption
'    fOrdPersLavDineroOcu = Trim(Right(Me.cboocupa1(0).Text, 4))
'    lsOrdPersLavDineroOcu = Trim(Left(Me.cboocupa1(0).Text, 75))
'
'    fReaPersLavDinero = Trim(TxtCodigo(1).Text)
'    fReaPersLavDineroNom = lblNombre(1).Caption
'    fReaPersLavDineroDir = lblDireccion(1).Caption
'    fReaPersLavDineroDoc = lblDocID(1).Caption
'    fReaPersLavDineroOcu = Trim(Right(Me.cboocupa1(1).Text, 4))
'    lsReaPersLavDineroOcu = Trim(Left(Me.cboocupa1(1).Text, 75))

    'MADM 20100722
    'If Me.cboocupa1(0).ListIndex = -1 Or Me.cboocupa1(1).ListIndex = -1 Or Me.cboocupa1(2).ListIndex = -1 Then 'JACA 20110223
'    If Me.cboocupa1(0).ListIndex = -1 Or Me.cboocupa1(1).ListIndex = -1 Then  ' JACA 20110223
'        MsgBox "Debe registrar la Ocupacion de la Persona", vbInformation, "Aviso"
'        cboocupa1(0).SetFocus
'        Exit Sub
'    End If


    'JACA 20110223
'    fBenPersLavDinero = Trim(txtCodigo(2).Text)
'    fBenPersLavDineroNom = lblNombre(2).Caption
'    fBenPersLavDineroDir = LblDireccion(2).Caption
'    fBenPersLavDineroDoc = lblDocID(2).Caption
'    fBenPersLavDineroOcu = Trim(Right(Me.cboocupa1(2).Text, 4))
        'CAMBIADO X ESTE PROCESO
        'EJVG20120327
        'Validar ingreso de Personas
        If Me.grdRealiza.TextMatrix(1, 1) = "" Then
             MsgBox "Debe Ingresar al menos una Persona que realiza la Transacción", vbInformation, "Aviso"
            Exit Sub
        End If
        If Me.grdOrdena.TextMatrix(1, 1) = "" Then
            MsgBox "Debe Ingresar al menos una Persona que Ordena la Transacción", vbInformation, "Aviso"
            Exit Sub
        End If
        If Me.grdBeneficiario.TextMatrix(1, 1) = "" Then
            'MsgBox "Debe Ingresar al menos un Beneficiario.", vbInformation, "Aviso"
            MsgBox "Debe Ingresar al menos un Beneficiario de la Transacción", vbInformation, "Aviso"
            Exit Sub
        End If
        'Validar las ocupaciones de las Personas
        For i = 1 To grdRealiza.Rows - 1
            If Trim(Right(grdRealiza.TextMatrix(i, 5), 4)) = "" Then
               MsgBox "Debe registrar la Ocupacion de la Persona que realiza la Transacción: " + grdRealiza.TextMatrix(i, 3), vbInformation, "Aviso"
               Exit Sub
            End If
        Next
        For i = 1 To grdOrdena.Rows - 1
            If CInt(grdOrdena.TextMatrix(i, 9)) = gPersonaNat Then    'NDX ERS0062020
                If Trim(Right(grdOrdena.TextMatrix(i, 5), 4)) = "" Then
                   MsgBox "Debe registrar la Ocupacion de la Persona que Ordena: " + grdOrdena.TextMatrix(i, 3), vbInformation, "Aviso"
                   Exit Sub
                End If
            End If  'NDX ERS0062020
        Next
        For i = 1 To grdBeneficiario.Rows - 1
            If Trim(Right(grdBeneficiario.TextMatrix(i, 9), 1)) = gPersonaNat Then    'NDX ERS0062020
                If Trim(Right(grdBeneficiario.TextMatrix(i, 5), 4)) = "" Then
                   MsgBox "Debe registrar la Ocupacion de la Persona que se Beneficia: " + grdBeneficiario.TextMatrix(i, 3), vbInformation, "Aviso"
                   Exit Sub
                End If
            End If  'NDX ERS0062020
        Next
        'NDX ERS0062020
        'Validar las ocupaciones de las Personas si figura como NO DECLARA
        For i = 1 To grdRealiza.Rows - 1
            If Trim(Right(grdRealiza.TextMatrix(i, 5), 4)) = 1999 Then 'NDX ERS0062020
               MsgBox "Debe ACTUALIZAR la Ocupacion de la Persona que realiza la Transacción: " + grdRealiza.TextMatrix(i, 3), vbInformation, "Aviso"
               Exit Sub
            End If
        Next
        For i = 1 To grdOrdena.Rows - 1
            If Trim(Right(grdOrdena.TextMatrix(i, 9), 1)) = gPersonaNat Then    'NDX ERS0062020
                If Trim(Right(grdOrdena.TextMatrix(i, 5), 4)) = 1999 Then 'NDX ERS0062020
                   MsgBox "Debe ACTUALIZAR la Ocupacion de la Persona que Ordena: " + grdOrdena.TextMatrix(i, 3), vbInformation, "Aviso"
                   Exit Sub
                End If
            End If  'NDX ERS0062020
        Next
        For i = 1 To grdBeneficiario.Rows - 1
            If Trim(Right(grdBeneficiario.TextMatrix(i, 9), 1)) = gPersonaNat Then    'NDX ERS0062020
                If Trim(Right(grdBeneficiario.TextMatrix(i, 5), 4)) = 1999 Then 'NDX ERS0062020
                   MsgBox "Debe ACTUALIZAR la Ocupacion de la Persona que se Beneficia: " + grdBeneficiario.TextMatrix(i, 3), vbInformation, "Aviso"
                   Exit Sub
                End If
            End If  'NDX ERS0062020
        Next
        'NDX ERS0062020 END
        
    'JACA END
    'Comentado x EJVG20120728
'    gVarPublicas.gOrdPersLavDinero = fOrdPersLavDinero 'By Capi 20012008
'    gVarPublicas.gReaPersLavDinero = fReaPersLavDinero
'    gVarPublicas.gBenPersLavDinero = fBenPersLavDinero
'    gVarPublicas.gBenPersLavDinero2 = fBenPersLavDinero2 'JACA 20110223
'    gVarPublicas.gBenPersLavDinero3 = fBenPersLavDinero3 'JACA 20110223
'    gVarPublicas.gBenPersLavDinero4 = fBenPersLavDinero4 'JACA 20110223
'    gVarPublicas.gVisPersLavDinero = fVisPersLavDinero 'DAOR 20070511
    gsOrigen = IIf(txtOrigen.Text = "", "", txtOrigen.Text)

    'If fOrdPersLavDinero = "" Or fReaPersLavDinero = "" Or fBenPersLavDinero = "" Then
        'MsgBox "Debe registrar una persona válida.", vbInformation, "Aviso"
        'TxtCodigo(0).SetFocus'EJVG20120327
        'Exit Sub
    'End If

     'MADM 20100308
    If gsOrigen = "" Then
        MsgBox "Debe registrar origen del efectivo.", vbInformation, "Aviso"
        txtOrigen.SetFocus
        Exit Sub
    End If

    'JACA 20110223
    Dim nPersIndex As Integer
    'nPersIndex = BuscarPersonaFlex(fOrdPersLavDinero)
    'EJVG20120327 Se ha cambiado la validación si se repitieran las Personas
    Dim iRea As Integer, iOrd As Integer
    Dim cPersCodRealiza As String, cPersCodOrdena As String

    For iRea = 1 To grdRealiza.Rows - 1
        cPersCodRealiza = grdRealiza.TextMatrix(iRea, 1)
        For iOrd = 1 To grdOrdena.Rows - 1
            cPersCodOrdena = grdOrdena.TextMatrix(iOrd, 1)
            nPersIndex = BuscarPersonaFlex(cPersCodOrdena)
            
            N_Ocupa1 = CInt(Trim(Right(Me.grdRealiza.TextMatrix(iRea, 5), 4)))
            If Me.grdOrdena.TextMatrix(iOrd, 9) = gPersonaNat Then 'NDX ERS0062020
                N_Ocupa2 = CInt(Trim(Right(Me.grdOrdena.TextMatrix(iOrd, 5), 4)))
            Else
                N_Ocupa2 = 0 'NDX ERS0062020
            End If
            
            If (cPersCodRealiza = cPersCodOrdena And nPersIndex <> 0) Then
                N_Ocupa3 = CInt(Trim(Right(Me.grdBeneficiario.TextMatrix(nPersIndex, 5), 4)))
                If ((N_Ocupa1 <> N_Ocupa2) Or (N_Ocupa2 <> N_Ocupa3)) Then
                    MsgBox "Las Ocupaciones de la Persona que Odena,Realiza y se Beneficia:" & Chr(13) & Me.grdRealiza.TextMatrix(iRea, 3) & " no Coinciden", vbInformation, "Aviso"
                    Exit Sub
                End If
            ElseIf (cPersCodOrdena = cPersCodRealiza) Then
                If N_Ocupa1 <> N_Ocupa2 Then
                    MsgBox "Las Ocupaciones de la Persona que Ordena y Realiza:" & Chr(13) & Me.grdRealiza.TextMatrix(iRea, 3) & " no Coinciden", vbInformation, "Aviso"
                    Exit Sub
                End If
            Else
                nPersIndex = BuscarPersonaFlex(cPersCodRealiza)
                If nPersIndex <> 0 Then
                    N_Ocupa3 = CInt(Trim(Right(Me.grdBeneficiario.TextMatrix(nPersIndex, 5), 4)))
                    If N_Ocupa1 <> N_Ocupa3 Then
                        MsgBox "Las Ocupaciones de la Persona que Realiza y se Beneficia: " & Chr(13) & Me.grdRealiza.TextMatrix(iRea, 3) & " no Coinciden", vbInformation, "Aviso"
                        Exit Sub
                    End If
                Else
                    nPersIndex = BuscarPersonaFlex(cPersCodOrdena)
                    If nPersIndex <> 0 Then
                        N_Ocupa3 = CInt(Trim(Right(Me.grdBeneficiario.TextMatrix(nPersIndex, 5), 4)))
                        If N_Ocupa2 <> N_Ocupa3 Then
                            MsgBox "Las Ocupaciones de la Persona que Ordena y se Beneficia:" & Chr(13) & Me.grdOrdena.TextMatrix(iOrd, 3) & " deben coincidir", vbInformation, "Aviso"
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Next
    Next
    'Comentado x EJVG20120327
'    'If fOrdPersLavDinero = fReaPersLavDinero And fOrdPersLavDinero= fBenPersLavDinero Then JACA 20110223
'    If fOrdPersLavDinero = fReaPersLavDinero And nPersIndex <> 0 Then
'
'         'JACA 20110223
'         'N_Ocupa1 = Me.cboocupa1(0).ListIndex
'         'N_Ocupa2 = Me.cboocupa1(1).ListIndex
'         'N_Ocupa3 = Me.cboocupa1(2).ListIndex
'         N_Ocupa1 = Trim(Right(Me.cboocupa1(0).Text, 4))
'         N_Ocupa2 = Trim(Right(Me.cboocupa1(1).Text, 4))
'         N_Ocupa3 = Trim(Right(Me.grdBeneficiario.TextMatrix(nPersIndex, 5), 4))
'
'        If ((N_Ocupa1 <> N_Ocupa2) Or (N_Ocupa2 <> N_Ocupa3)) Then
'            MsgBox "Las Ocupaciones de la Persona que Odena,Realiza y se Beneficia no Coinciden", vbInformation, "Aviso"
'            cboocupa1(0).SetFocus
'            Exit Sub
'        End If
'
'    ElseIf (fOrdPersLavDinero = fReaPersLavDinero) Then
'         N_Ocupa1 = Me.cboocupa1(0).ListIndex
'         N_Ocupa2 = Me.cboocupa1(1).ListIndex
'        If N_Ocupa1 <> N_Ocupa2 Then
'            MsgBox "Las Ocupaciones de la Persona que Ordena y Realiza deben coincidir", vbInformation, "Aviso"
'            cboocupa1(0).SetFocus
'        Exit Sub
'        End If
'    'ElseIf (fBenPersLavDinero = fReaPersLavDinero) Then JACA 20110224
'    Else
'        'JACA 20110224
'        nPersIndex = BuscarPersonaFlex(fReaPersLavDinero)
'        If nPersIndex <> 0 Then
''            N_Ocupa2 = Me.cboocupa1(1).ListIndex
''            N_Ocupa3 = Me.cboocupa1(2).ListIndex
'             N_Ocupa2 = Trim(Right(Me.cboocupa1(1).Text, 4))
'             N_Ocupa3 = Trim(Right(Me.grdBeneficiario.TextMatrix(nPersIndex, 5), 4))
'        'JACA END
'            If N_Ocupa2 <> N_Ocupa3 Then
'                MsgBox "Las Ocupaciones de la Persona que se Realiza y se Benefica deben coincidir", vbInformation, "Aviso"
'                cboocupa1(1).SetFocus
'            Exit Sub
'            End If
'    'ElseIf (fBenPersLavDinero = fOrdPersLavDinero) Then JACA 20110224
'        Else
'            'JACA 20110224
'            nPersIndex = BuscarPersonaFlex(fOrdPersLavDinero)
'            If nPersIndex <> 0 Then
'                'N_Ocupa2 = Me.cboocupa1(0).ListIndex
'                'N_Ocupa3 = Me.cboocupa1(2).ListIndex
'                N_Ocupa2 = Trim(Right(Me.cboocupa1(0).Text, 4))
'                N_Ocupa3 = Trim(Right(Me.grdBeneficiario.TextMatrix(nPersIndex, 5), 4))
'            'JACA END
'
'                If N_Ocupa2 <> N_Ocupa3 Then
'                    MsgBox "Las Ocupaciones de la Persona que Ordena y se Beneficia deben coincidir", vbInformation, "Aviso"
'                    cboocupa1(0).SetFocus
'                Exit Sub
'                End If
'
'            End If
'        End If
'    End If

    'Call objREU.ModificaOcupaPers_REU(fReaPersLavDinero, CInt(Right(Me.cboocupa1(1).Text, 4)))
    'EJVG20120327
    For i = 1 To Me.grdRealiza.Rows - 1
            If Me.grdRealiza.TextMatrix(i, 1) <> "" And i = 1 Then
                fReaPersLavDinero = Me.grdRealiza.TextMatrix(i, 1)
                fReaPersLavDineroDoc = Me.grdRealiza.TextMatrix(i, 2)
                fReaPersLavDineroNom = Me.grdRealiza.TextMatrix(i, 3)
                fReaPersLavDineroDir = Me.grdRealiza.TextMatrix(i, 4)
                fReaPersLavDineroOcu = Trim(Right(Me.grdRealiza.TextMatrix(i, 5), 4))
                lsReaPersLavDineroOcu = Trim(Left(Me.grdRealiza.TextMatrix(i, 5), 75))
                'WIOR 20130301 ********************************************
                fReaPersLavDineroNac = grdRealiza.TextMatrix(i, 6)
                fReaPersLavDineroRes = grdRealiza.TextMatrix(i, 7)
                fReaPersLavDineroPeps = grdRealiza.TextMatrix(i, 8)
                'WIOR FIN *************************************************
            ElseIf Me.grdRealiza.TextMatrix(i, 1) <> "" And i = 2 Then
                fReaPersLavDinero2 = Me.grdRealiza.TextMatrix(i, 1)
                fReaPersLavDineroDoc2 = Me.grdRealiza.TextMatrix(i, 2)
                fReaPersLavDineroNom2 = Me.grdRealiza.TextMatrix(i, 3)
                fReaPersLavDineroDir2 = Me.grdRealiza.TextMatrix(i, 4)
                fReaPersLavDineroOcu2 = Trim(Right(Me.grdRealiza.TextMatrix(i, 5), 4))
                lsReaPersLavDineroOcu2 = Trim(Left(Me.grdRealiza.TextMatrix(i, 5), 75))
                'WIOR 20130301 ********************************************
                fReaPersLavDineroNac2 = grdRealiza.TextMatrix(i, 6)
                fReaPersLavDineroRes2 = grdRealiza.TextMatrix(i, 7)
                fReaPersLavDineroPeps2 = grdRealiza.TextMatrix(i, 8)
                'WIOR FIN *************************************************
            ElseIf Me.grdRealiza.TextMatrix(i, 1) <> "" And i = 3 Then
                fReaPersLavDinero3 = Me.grdRealiza.TextMatrix(i, 1)
                fReaPersLavDineroDoc3 = Me.grdRealiza.TextMatrix(i, 2)
                fReaPersLavDineroNom3 = Me.grdRealiza.TextMatrix(i, 3)
                fReaPersLavDineroDir3 = Me.grdRealiza.TextMatrix(i, 4)
                fReaPersLavDineroOcu3 = Trim(Right(Me.grdRealiza.TextMatrix(i, 5), 4))
                lsReaPersLavDineroOcu3 = Trim(Left(Me.grdRealiza.TextMatrix(i, 5), 75))
                'WIOR 20130301 ********************************************
                fReaPersLavDineroNac3 = grdRealiza.TextMatrix(i, 6)
                fReaPersLavDineroRes3 = grdRealiza.TextMatrix(i, 7)
                fReaPersLavDineroPeps3 = grdRealiza.TextMatrix(i, 8)
                'WIOR FIN *************************************************
            ElseIf Me.grdRealiza.TextMatrix(i, 1) <> "" And i = 4 Then
                fReaPersLavDinero4 = Me.grdRealiza.TextMatrix(i, 1)
                fReaPersLavDineroDoc4 = Me.grdRealiza.TextMatrix(i, 2)
                fReaPersLavDineroNom4 = Me.grdRealiza.TextMatrix(i, 3)
                fReaPersLavDineroDir4 = Me.grdRealiza.TextMatrix(i, 4)
                fReaPersLavDineroOcu4 = Trim(Right(Me.grdRealiza.TextMatrix(i, 5), 4))
                lsReaPersLavDineroOcu4 = Trim(Left(Me.grdRealiza.TextMatrix(i, 5), 75))
                'WIOR 20130301 ********************************************
                fReaPersLavDineroNac4 = grdRealiza.TextMatrix(i, 6)
                fReaPersLavDineroRes4 = grdRealiza.TextMatrix(i, 7)
                fReaPersLavDineroPeps4 = grdRealiza.TextMatrix(i, 8)
                'WIOR FIN *************************************************
            End If
        Next
        For i = 1 To Me.grdOrdena.Rows - 1
            If Me.grdOrdena.TextMatrix(i, 1) <> "" And i = 1 Then
                fOrdPersLavDinero = Me.grdOrdena.TextMatrix(i, 1)
                fOrdPersLavDineroDoc = Me.grdOrdena.TextMatrix(i, 2)
                fOrdPersLavDineroNom = Me.grdOrdena.TextMatrix(i, 3)
                fOrdPersLavDineroDir = Me.grdOrdena.TextMatrix(i, 4)
                fOrdPersLavDineroOcu = Trim(Right(Me.grdOrdena.TextMatrix(i, 5), 4))
                lsOrdPersLavDineroOcu = Trim(Left(Me.grdOrdena.TextMatrix(i, 5), 75))
                'WIOR 20130301 ********************************************
                fOrdPersLavDineroNac = grdOrdena.TextMatrix(i, 6)
                fOrdPersLavDineroRes = grdOrdena.TextMatrix(i, 7)
                fOrdPersLavDineroPeps = grdOrdena.TextMatrix(i, 8)
                fOrdPersLavDineroPers = grdOrdena.TextMatrix(i, 9) 'NDX ERS0062020
                'WIOR FIN *************************************************
            ElseIf Me.grdOrdena.TextMatrix(i, 1) <> "" And i = 2 Then
                fOrdPersLavDinero2 = Me.grdOrdena.TextMatrix(i, 1)
                fOrdPersLavDineroDoc2 = Me.grdOrdena.TextMatrix(i, 2)
                fOrdPersLavDineroNom2 = Me.grdOrdena.TextMatrix(i, 3)
                fOrdPersLavDineroDir2 = Me.grdOrdena.TextMatrix(i, 4)
                fOrdPersLavDineroOcu2 = Trim(Right(Me.grdOrdena.TextMatrix(i, 5), 4))
                lsOrdPersLavDineroOcu2 = Trim(Left(Me.grdOrdena.TextMatrix(i, 5), 75))
                'WIOR 20130301 ********************************************
                fOrdPersLavDineroNac2 = grdOrdena.TextMatrix(i, 6)
                fOrdPersLavDineroRes2 = grdOrdena.TextMatrix(i, 7)
                fOrdPersLavDineroPeps2 = grdOrdena.TextMatrix(i, 8)
                fOrdPersLavDineroPers2 = grdOrdena.TextMatrix(i, 9) 'NDX ERS0062020
                'WIOR FIN *************************************************
            ElseIf Me.grdOrdena.TextMatrix(i, 1) <> "" And i = 3 Then
                fOrdPersLavDinero3 = Me.grdOrdena.TextMatrix(i, 1)
                fOrdPersLavDineroDoc3 = Me.grdOrdena.TextMatrix(i, 2)
                fOrdPersLavDineroNom3 = Me.grdOrdena.TextMatrix(i, 3)
                fOrdPersLavDineroDir3 = Me.grdOrdena.TextMatrix(i, 4)
                fOrdPersLavDineroOcu3 = Trim(Right(Me.grdOrdena.TextMatrix(i, 5), 4))
                lsOrdPersLavDineroOcu3 = Trim(Left(Me.grdOrdena.TextMatrix(i, 5), 75))
                'WIOR 20130301 ********************************************
                fOrdPersLavDineroNac3 = grdOrdena.TextMatrix(i, 6)
                fOrdPersLavDineroRes3 = grdOrdena.TextMatrix(i, 7)
                fOrdPersLavDineroPeps3 = grdOrdena.TextMatrix(i, 8)
                fOrdPersLavDineroPers3 = grdOrdena.TextMatrix(i, 9) 'NDX ERS0062020
                'WIOR FIN *************************************************
            ElseIf Me.grdOrdena.TextMatrix(i, 1) <> "" And i = 4 Then
                fOrdPersLavDinero4 = Me.grdOrdena.TextMatrix(i, 1)
                fOrdPersLavDineroDoc4 = Me.grdOrdena.TextMatrix(i, 2)
                fOrdPersLavDineroNom4 = Me.grdOrdena.TextMatrix(i, 3)
                fOrdPersLavDineroDir4 = Me.grdOrdena.TextMatrix(i, 4)
                fOrdPersLavDineroOcu4 = Trim(Right(Me.grdOrdena.TextMatrix(i, 5), 4))
                lsOrdPersLavDineroOcu4 = Trim(Left(Me.grdOrdena.TextMatrix(i, 5), 75))
                'WIOR 20130301 ********************************************
                fOrdPersLavDineroNac4 = grdOrdena.TextMatrix(i, 6)
                fOrdPersLavDineroRes4 = grdOrdena.TextMatrix(i, 7)
                fOrdPersLavDineroPeps4 = grdOrdena.TextMatrix(i, 8)
                fOrdPersLavDineroPers4 = grdOrdena.TextMatrix(i, 9) 'NDX ERS0062020
                'WIOR FIN *************************************************
            End If
        Next
        For i = 1 To Me.grdBeneficiario.Rows - 1

            If Me.grdBeneficiario.TextMatrix(i, 1) <> "" And i = 1 Then
                fBenPersLavDinero = Me.grdBeneficiario.TextMatrix(i, 1)
                fBenPersLavDineroDoc = Me.grdBeneficiario.TextMatrix(i, 2)
                fBenPersLavDineroNom = Me.grdBeneficiario.TextMatrix(i, 3)
                fBenPersLavDineroDir = Me.grdBeneficiario.TextMatrix(i, 4)
                fBenPersLavDineroOcu = Trim(Right(Me.grdBeneficiario.TextMatrix(i, 5), 4))
                lsBenPersLavDineroOcu = Trim(Left(Me.grdBeneficiario.TextMatrix(i, 5), 75))
                'Comentado x EJVG20120327
'                If fBenPersLavDineroOcu = "" Then
'                   MsgBox "Debe registrar la Ocupacion de la Persona que se Beneficia:" + fBenPersLavDineroNom, vbInformation, "Aviso"
'                   Exit Sub
'                End If
                'WIOR 20130301 ********************************************
                fBenPersLavDineroNac = grdBeneficiario.TextMatrix(i, 6)
                fBenPersLavDineroRes = grdBeneficiario.TextMatrix(i, 7)
                fBenPersLavDineroPeps = grdBeneficiario.TextMatrix(i, 8)
                fBenPersLavDineroPers = grdBeneficiario.TextMatrix(i, 9) 'NDX ERS0062020
                'WIOR FIN *************************************************
            ElseIf Me.grdBeneficiario.TextMatrix(i, 1) <> "" And i = 2 Then
                fBenPersLavDinero2 = Me.grdBeneficiario.TextMatrix(i, 1)
                fBenPersLavDineroDoc2 = Me.grdBeneficiario.TextMatrix(i, 2)
                fBenPersLavDineroNom2 = Me.grdBeneficiario.TextMatrix(i, 3)
                fBenPersLavDineroDir2 = Me.grdBeneficiario.TextMatrix(i, 4)
                fBenPersLavDineroOcu2 = Trim(Right(Me.grdBeneficiario.TextMatrix(i, 5), 4))
                lsBenPersLavDineroOcu2 = Trim(Left(Me.grdBeneficiario.TextMatrix(i, 5), 75))
                'Comentado x EJVG20120327
'                If fBenPersLavDineroOcu2 = "" Then
'                   MsgBox "Debe registrar la Ocupacion de la Persona que se Beneficia:" + fBenPersLavDineroNom2, vbInformation, "Aviso"
'                   Exit Sub
'                End If
                'WIOR 20130301 ********************************************
                fBenPersLavDineroNac2 = grdBeneficiario.TextMatrix(i, 6)
                fBenPersLavDineroRes2 = grdBeneficiario.TextMatrix(i, 7)
                fBenPersLavDineroPeps2 = grdBeneficiario.TextMatrix(i, 8)
                fBenPersLavDineroPers2 = grdBeneficiario.TextMatrix(i, 9) 'NDX ERS0062020
                'WIOR FIN *************************************************
            ElseIf Me.grdBeneficiario.TextMatrix(i, 1) <> "" And i = 3 Then
                fBenPersLavDinero3 = Me.grdBeneficiario.TextMatrix(i, 1)
                fBenPersLavDineroDoc3 = Me.grdBeneficiario.TextMatrix(i, 2)
                fBenPersLavDineroNom3 = Me.grdBeneficiario.TextMatrix(i, 3)
                fBenPersLavDineroDir3 = Me.grdBeneficiario.TextMatrix(i, 4)
                fBenPersLavDineroOcu3 = Trim(Right(Me.grdBeneficiario.TextMatrix(i, 5), 4))
                lsBenPersLavDineroOcu3 = Trim(Left(Me.grdBeneficiario.TextMatrix(i, 5), 75))
                'Comentado x EJVG20120327
'                If fBenPersLavDineroOcu3 = "" Then
'                   MsgBox "Debe registrar la Ocupacion de la Persona que se Beneficia:" + fBenPersLavDineroNom3, vbInformation, "Aviso"
'                   Exit Sub
'                End If
                'WIOR 20130301 ********************************************
                fBenPersLavDineroNac3 = grdBeneficiario.TextMatrix(i, 6)
                fBenPersLavDineroRes3 = grdBeneficiario.TextMatrix(i, 7)
                fBenPersLavDineroPeps3 = grdBeneficiario.TextMatrix(i, 8)
                fBenPersLavDineroPers3 = grdBeneficiario.TextMatrix(i, 9) 'NDX ERS0062020
                'WIOR FIN *************************************************
            ElseIf Me.grdBeneficiario.TextMatrix(i, 1) <> "" And i = 4 Then
                fBenPersLavDinero4 = Me.grdBeneficiario.TextMatrix(i, 1)
                fBenPersLavDineroDoc4 = Me.grdBeneficiario.TextMatrix(i, 2)
                fBenPersLavDineroNom4 = Me.grdBeneficiario.TextMatrix(i, 3)
                fBenPersLavDineroDir4 = Me.grdBeneficiario.TextMatrix(i, 4)
                fBenPersLavDineroOcu4 = Trim(Right(Me.grdBeneficiario.TextMatrix(i, 5), 4))
                lsBenPersLavDineroOcu4 = Trim(Left(Me.grdBeneficiario.TextMatrix(i, 5), 75))
                'Comentado x EJVG20120327
'                If fBenPersLavDineroOcu4 = "" Then
'                   MsgBox "Debe registrar la Ocupacion de la Persona que se Beneficia:" + fBenPersLavDineroNom4, vbInformation, "Aviso"
'                   Exit Sub
'                End If
                'WIOR 20130301 ********************************************
                fBenPersLavDineroNac4 = grdBeneficiario.TextMatrix(i, 6)
                fBenPersLavDineroRes4 = grdBeneficiario.TextMatrix(i, 7)
                fBenPersLavDineroPeps4 = grdBeneficiario.TextMatrix(i, 8)
                fBenPersLavDineroPers4 = grdBeneficiario.TextMatrix(i, 9) 'NDX ERS0062020
                'WIOR FIN *************************************************
            End If
        Next
    'Actualizar los datos de los que realizan la Transacción si es que es su ultima actualizacion fue mayor a 365 dias
    For i = 1 To Me.grdRealiza.Rows - 1
        Dim oPersona As New COMNPersona.NCOMPersona
        If oPersona.NecesitaActualizarDatos(grdRealiza.TextMatrix(i, 1), gdFecSis) Then
            MsgBox "Para continuar con la Operación Ud. debe actualizar los datos de:" & Chr(13) & grdRealiza.TextMatrix(i, 3), vbInformation, "Aviso"
            Dim foPersona As New frmPersona
            Dim DioMantenimiento As Boolean
            DioMantenimiento = False
            Do While Not DioMantenimiento
               DioMantenimiento = foPersona.realizarMantenimiento(grdRealiza.TextMatrix(i, 1))
               If Not DioMantenimiento Then
                   MsgBox "Ud. Necesariamente debe actualizar los datos de: " & grdRealiza.TextMatrix(i, 3), vbInformation, "Aviso"
               End If
            Loop
        End If
    Next
        
    Call objREU.ModificaOcupaPers_REU(fReaPersLavDinero, CInt(fReaPersLavDineroOcu))
    If fReaPersLavDinero2 <> "" Then Call objREU.ModificaOcupaPers_REU(fReaPersLavDinero2, CInt(fReaPersLavDineroOcu2))
    If fReaPersLavDinero3 <> "" Then Call objREU.ModificaOcupaPers_REU(fReaPersLavDinero3, CInt(fReaPersLavDineroOcu3))
    If fReaPersLavDinero4 <> "" Then Call objREU.ModificaOcupaPers_REU(fReaPersLavDinero4, CInt(fReaPersLavDineroOcu4))
    'Call objREU.ModificaOcupaPers_REU(fOrdPersLavDinero, CInt(Right(Me.cboocupa1(0).Text, 4)))
    'Call objREU.ModificaOcupaPers_REU(fOrdPersLavDinero, CInt(fOrdPersLavDineroOcu)) 'NDX ERS0062020 COMENTÓ
    If fOrdPersLavDineroPers = "1" Then Call objREU.ModificaOcupaPers_REU(fOrdPersLavDinero, CInt(fOrdPersLavDineroOcu)) 'NDX ERS0062020
    'If fOrdPersLavDinero2 <> "" Then Call objREU.ModificaOcupaPers_REU(fOrdPersLavDinero2, CInt(fOrdPersLavDineroOcu2))'NDX ERS0062020 COMENTÓ
    If fOrdPersLavDinero2 <> "" And fOrdPersLavDineroPers2 = "1" Then Call objREU.ModificaOcupaPers_REU(fOrdPersLavDinero2, CInt(fOrdPersLavDineroOcu2)) 'NDX ERS0062020
    'If fOrdPersLavDinero3 <> "" Then Call objREU.ModificaOcupaPers_REU(fOrdPersLavDinero3, CInt(fOrdPersLavDineroOcu3))'NDX ERS0062020 COMENTÓ
    If fOrdPersLavDinero3 <> "" And fOrdPersLavDineroPers3 = "1" Then Call objREU.ModificaOcupaPers_REU(fOrdPersLavDinero3, CInt(fOrdPersLavDineroOcu3)) 'NDX ERS0062020
    'If fOrdPersLavDinero4 <> "" Then Call objREU.ModificaOcupaPers_REU(fOrdPersLavDinero4, CInt(fOrdPersLavDineroOcu4))'NDX ERS0062020 COMENTÓ
    If fOrdPersLavDinero4 <> "" And fOrdPersLavDineroPers4 = "1" Then Call objREU.ModificaOcupaPers_REU(fOrdPersLavDinero4, CInt(fOrdPersLavDineroOcu4)) 'NDX ERS0062020
    'Call objREU.ModificaOcupaPers_REU(fBenPersLavDinero, CInt(Trim(Right(Me.grdBeneficiario.TextMatrix(1, 5), 4)))) 'NDX ERS0062020 COMENTÓ
    If fBenPersLavDineroPers = "1" Then Call objREU.ModificaOcupaPers_REU(fBenPersLavDinero, CInt(Trim(Right(Me.grdBeneficiario.TextMatrix(1, 5), 4)))) 'NDX ERS0062020
    'JACA 20110224
    'If fBenPersLavDinero2 <> "" Then Call objREU.ModificaOcupaPers_REU(fBenPersLavDinero2, CInt(Trim(Right(Me.grdBeneficiario.TextMatrix(2, 5), 4))))'NDX ERS0062020 COMENTÓ
    If fBenPersLavDinero2 <> "" And fBenPersLavDineroPers2 = "1" Then Call objREU.ModificaOcupaPers_REU(fBenPersLavDinero2, CInt(Trim(Right(Me.grdBeneficiario.TextMatrix(2, 5), 4)))) 'NDX ERS0062020
    'If fBenPersLavDinero3 <> "" Then Call objREU.ModificaOcupaPers_REU(fBenPersLavDinero3, CInt(Trim(Right(Me.grdBeneficiario.TextMatrix(3, 5), 4))))'NDX ERS0062020 COMENTÓ
    If fBenPersLavDinero3 <> "" And fBenPersLavDineroPers3 = "1" Then Call objREU.ModificaOcupaPers_REU(fBenPersLavDinero3, CInt(Trim(Right(Me.grdBeneficiario.TextMatrix(3, 5), 4)))) 'NDX ERS0062020
    'If fBenPersLavDinero4 <> "" Then Call objREU.ModificaOcupaPers_REU(fBenPersLavDinero4, CInt(Trim(Right(Me.grdBeneficiario.TextMatrix(4, 5), 4))))'NDX ERS0062020 COMENTÓ
    If fBenPersLavDinero4 <> "" And fBenPersLavDineroPers4 = "1" Then Call objREU.ModificaOcupaPers_REU(fBenPersLavDinero4, CInt(Trim(Right(Me.grdBeneficiario.TextMatrix(4, 5), 4)))) 'NDX ERS0062020

    fOrigenPersLavDinero = txtOrigen.Text
    'end madm

     'JACA 20110224
     'Comentado by JACA 20110929*********************************************
'             Dim rsCorr As Recordset
'             Set rsCorr = objREU.ObtenerAgenCorrelativoLavDinero(gsCodAge)
'             If Not (rsCorr.EOF And rsCorr.BOF) Then
'                fnNroREU = rsCorr!nNroREU
'             Else
              'fnNroREU = 0
'             End If
'
'            If bImprimeBoleta Then
'                'By capi 19022009 para que envie la moneda
'                'Call imprimirBoletaREU(, , txtOrigen.Text)
'                Call imprimirBoletaREU(, nmoneda, txtOrigen.Text, fnNroREU) 'JACA 20110325
'
'            End If
    'JACA END***************************************************************
    Set objREU = Nothing
    'Unload Me
    Call Form_Unload(lnTipoSalida)
    sCodPersona = fOrdPersLavDinero
End Sub

'JACA 20110223
Private Function BuscarPersonaFlex(ByVal pcPersCod As String) As Integer
    Dim i As Integer
    BuscarPersonaFlex = 0
    For i = 1 To Me.grdBeneficiario.Rows - 1
        If Me.grdBeneficiario.TextMatrix(i, 1) = pcPersCod Then
            BuscarPersonaFlex = i
            Exit For
        End If
        
    Next
End Function
'JACA END
Private Sub cmdVisitasEliminar_Click()
If MsgBox("¿¿Está seguro de eliminar la selección actual??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    If grdBeneficiario.Rows = 2 Then
       cmdVisitasEliminar.Enabled = False
    End If
     grdBeneficiario.EliminaFila grdBeneficiario.row
End If
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
lnTipoSalida = 0
End Sub
'**CAPI 20080213
'Comentado x EJVG20120327
'Sub obtenerDatosPersonasLavDinero()
'Dim loPersona As COMDPersona.DCOMPersonas
'Dim lnIndexPer As Integer, i As Integer
'Dim lrs As ADODB.Recordset
'Dim lsPersCod As String, lsNombre As String, lsDireccion As String, lsdocumento As String, lsOcupa As String
'    lsPersCod = ""
'    Set loPersona = New COMDPersona.DCOMPersonas
'    Set lrs = loPersona.obtenerDatosPersonasLavDinero(fOrdPersLavDinero, fReaPersLavDinero, fBenPersLavDinero, fTitPersLavDinero)
'    While Not lrs.EOF
'        Select Case lrs!cPersCod
'            Case fOrdPersLavDinero
'                fOrdPersLavDineroNom = PstaNombre(lrs!Nombre)
'                fOrdPersLavDineroDir = lrs!Direccion
'                fOrdPersLavDineroDoc = lrs!IDNum
'                fOrdPersLavDineroDoc = lrs!cActiGiro1
'                lnIndexPer = 0
'            Case fReaPersLavDinero
'                fReaPersLavDineroNom = PstaNombre(lrs!Nombre)
'                fReaPersLavDineroDir = lrs!Direccion
'                fReaPersLavDineroDoc = lrs!IDNum
'                fReaPersLavDineroDoc = lrs!cActiGiro1
'                lnIndexPer = 1
'            Case fBenPersLavDinero
'                fBenPersLavDineroNom = PstaNombre(lrs!Nombre)
'                fBenPersLavDineroDir = lrs!Direccion
'                fBenPersLavDineroDoc = lrs!IDNum
'                fBenPersLavDineroDoc = lrs!cActiGiro1
'                lnIndexPer = 2
'            Case fTitPersLavDinero
'                'By Capi 20042008
'                If lrs!nEstado = 2 Then
'                    MsgBox "!OJO! Persona Registrada como Fraudelento...Coordine con Supervision...Operacion Cancelada", vbInformation, "Aviso"
'                    fTitPersLavDinero = ""
'                    fReaPersLavDinero = ""
'                    Exit Sub
'                ElseIf lrs!nEstado = 1 Then
'                    MsgBox "!OJO! Persona Registrada como Dudoso", vbInformation, "Aviso"
'                ElseIf lrs!nEstado = 3 Then
'                    MsgBox "!OJO! Persona Registrada como PEPS", vbInformation, "Aviso"
'                End If
'                '
'                fTitPersLavDineroNom = PstaNombre(lrs!Nombre)
'                fTitPersLavDineroDir = lrs!Direccion
'                fTitPersLavDineroDoc = lrs!IDNum
'                fTitPersLavDineroDoc = lrs!cActiGiro1
'                lnIndexPer = 4
'        End Select
'        If lsPersCod = "" Then
'            lsPersCod = lrs!cPersCod
'            lsNombre = lrs!Nombre
'            lsDireccion = lrs!Direccion
'            lsdocumento = lrs!IDNum
'        End If
'        If lnIndexPer <> 4 Then
'            lblNombre(lnIndexPer).Caption = lrs!Nombre
'            lblDireccion(lnIndexPer).Caption = lrs!Direccion
'            lblDocID(lnIndexPer).Caption = lrs!IDNum
'        End If
'        lrs.MoveNext
'    Wend
'    For i = 0 To 2
'        If TxtCodigo(i).Text = "" Then
'            TxtCodigo(i).Text = lsPersCod
'            lblNombre(i).Caption = lsNombre
'            lblDireccion(i).Caption = lsDireccion
'            lblDocID(i).Caption = lsdocumento
'            cboocupa1(i).ListIndex = lsOcupa
'        End If
'    Next i
'    Set lrs = Nothing
'    Set loPersona = Nothing
'End Sub

'MAVM 20120228 ***
Private Sub Form_Unload(Cancel As Integer)
'ALPA 20120302************
    If lnTipoSalida = 0 Then
        sCodPersona = ""
        fTitPersLavDinero = ""
        fOrdPersLavDinero = ""
        fReaPersLavDinero = ""
        fBenPersLavDinero = ""
        fVisPersLavDinero = ""
    End If
    Unload Me
End Sub
'***

'JACA 20110223
Private Sub grdBeneficiario_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)

    Dim nPersoneria As PersPersoneria
    Dim nCondicion As Integer 'WIOR 20130301

    If pbEsDuplicado Then
            MsgBox "Persona ya esta registrada en la relación.", vbInformation, "Aviso"
            grdBeneficiario.EliminaFila grdBeneficiario.row
    ElseIf psDataCod = "" Then
                grdBeneficiario.TextMatrix(grdBeneficiario.row, 2) = ""
                grdBeneficiario.TextMatrix(grdBeneficiario.row, 3) = ""
                grdBeneficiario.TextMatrix(grdBeneficiario.row, 4) = ""
                grdBeneficiario.TextMatrix(grdBeneficiario.row, 5) = ""
            Exit Sub
    Else
            'nPersoneria = CLng(Trim(grdBeneficiario.PersPersoneria)) Comentado x JACA20110905
            
   
            'If nPersoneria = gPersonaNat Then Comentado x JACA20110905
                'Call llenar_cboFlex_ocupa
                Call llenar_cboFlex_ocupa(grdBeneficiario)
                Dim ClsPersona As COMDPersona.DCOMPersonas
                Dim R As New ADODB.Recordset
                Dim obj As COMDPersona.DCOMPersonas
                Dim lrsOcupa As ADODB.Recordset 'madm 20100723
                Dim sOcupacion As String
                 
                Set obj = New COMDPersona.DCOMPersonas
                Set lrsOcupa = obj.CargarOcupaciones()
               
                Set ClsPersona = New COMDPersona.DCOMPersonas
                
                Set R = ClsPersona.BuscaCliente(grdBeneficiario.TextMatrix(grdBeneficiario.row, 1), 2)
                
                If Not (R.EOF And R.BOF) Then
                    
                    grdBeneficiario.TextMatrix(grdBeneficiario.row, 2) = IIf(IsNull(R!cPersIDnroDNI), "", IIf(R!cPersIDnroDNI = "", R!cPersIDnroRUC, R!cPersIDnroDNI))
                    grdBeneficiario.TextMatrix(grdBeneficiario.row, 3) = PstaNombre(R!cPersNombre)
                    grdBeneficiario.TextMatrix(grdBeneficiario.row, 4) = R!cPersDireccDomicilio
                    sOcupacion = IIf(IsNull(R!cActiGiro1), "", IIf(R!cActiGiro1 = "", R!cActiGiro1, R!cActiGiro1))
                    'WIOR 20130301 ********************************************
                    grdBeneficiario.TextMatrix(grdBeneficiario.row, 6) = PstaNombre(R!cNacionalidad)
                    grdBeneficiario.TextMatrix(grdBeneficiario.row, 7) = PstaNombre(R!cResidente)
                    Set foPersona = New COMDPersona.UCOMPersona
                    Call foPersona.ValidaEnListaNegativaCondicion(IIf(IsNull(R!cPersIDnroDNI), "", R!cPersIDnroDNI), IIf(IsNull(R!cPersIDnroRUC), "", R!cPersIDnroRUC), nCondicion, R!cPersNombre)
                    grdBeneficiario.TextMatrix(grdBeneficiario.row, 8) = IIf(nCondicion = 3, "SI", "NO")
                    nPersoneria = IIf(IsNull(R!nPersPersoneria), 0, R!nPersPersoneria)  'NDX ERS0062020
                    grdBeneficiario.TextMatrix(grdBeneficiario.row, 9) = nPersoneria                'NDX ERS0062020
                    Set foPersona = Nothing
                    'WIOR FIN *************************************************
                    If sOcupacion <> "" Then
                        Do While Trim(Str(lrsOcupa!nConsValor)) <> sOcupacion
                             lrsOcupa.MoveNext
                        Loop
                            grdBeneficiario.TextMatrix(grdBeneficiario.row, 5) = Trim(lrsOcupa!cConsDescripcion) & space(100) & Trim(Str(lrsOcupa!nConsValor))
                    Else
                        grdBeneficiario.TextMatrix(grdBeneficiario.row, 5) = sOcupacion
                    End If
                    lrsOcupa.Close
                    
                   
                End If
                Set ClsPersona = Nothing
                Set obj = Nothing
             'Comentado x JACA 20110905************************************************
'            Else
'                MsgBox "Persona a registrar debe ser Persona Natural", vbInformation, "Aviso"
'                grdBeneficiario.TextMatrix(grdBeneficiario.Row, 2) = ""
'                grdBeneficiario.TextMatrix(grdBeneficiario.Row, 3) = ""
'                grdBeneficiario.TextMatrix(grdBeneficiario.Row, 4) = ""
'                grdBeneficiario.TextMatrix(grdBeneficiario.Row, 5) = ""
'            End If
            'JACA END***************************************************************
    End If
End Sub
'JACA END

'EJVG20120327
Private Sub grdOrdena_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim nPersoneria As PersPersoneria
    Dim nCondicion As Integer 'WIOR 20130301
    If pbEsDuplicado Then
        MsgBox "Persona ya esta registrada en la relación.", vbInformation, "Aviso"
        grdOrdena.EliminaFila grdOrdena.row
    ElseIf psDataCod = "" Then
                grdOrdena.TextMatrix(grdOrdena.row, 2) = ""
                grdOrdena.TextMatrix(grdOrdena.row, 3) = ""
                grdOrdena.TextMatrix(grdOrdena.row, 4) = ""
                grdOrdena.TextMatrix(grdOrdena.row, 5) = ""
            Exit Sub
    Else
        Call llenar_cboFlex_ocupa(grdOrdena)
        Dim ClsPersona As COMDPersona.DCOMPersonas
        Dim R As New ADODB.Recordset
        Dim obj As COMDPersona.DCOMPersonas
        Dim lrsOcupa As ADODB.Recordset
        Dim sOcupacion As String
          
        Set obj = New COMDPersona.DCOMPersonas
        Set lrsOcupa = obj.CargarOcupaciones()
        Set ClsPersona = New COMDPersona.DCOMPersonas
        Set R = ClsPersona.BuscaCliente(grdOrdena.TextMatrix(grdOrdena.row, 1), 2)
         
        If Not (R.EOF And R.BOF) Then
            grdOrdena.TextMatrix(grdOrdena.row, 2) = IIf(IsNull(R!cPersIDnroDNI), "", IIf(R!cPersIDnroDNI = "", R!cPersIDnroRUC, R!cPersIDnroDNI))
            grdOrdena.TextMatrix(grdOrdena.row, 3) = PstaNombre(R!cPersNombre)
            grdOrdena.TextMatrix(grdOrdena.row, 4) = R!cPersDireccDomicilio
            sOcupacion = IIf(IsNull(R!cActiGiro1), "", IIf(R!cActiGiro1 = "", R!cActiGiro1, R!cActiGiro1))
            'WIOR 20130301 ********************************************
            grdOrdena.TextMatrix(grdOrdena.row, 6) = PstaNombre(R!cNacionalidad)
            grdOrdena.TextMatrix(grdOrdena.row, 7) = PstaNombre(R!cResidente)
            Set foPersona = New COMDPersona.UCOMPersona
            Call foPersona.ValidaEnListaNegativaCondicion(IIf(IsNull(R!cPersIDnroDNI), "", R!cPersIDnroDNI), IIf(IsNull(R!cPersIDnroRUC), "", R!cPersIDnroRUC), nCondicion, R!cPersNombre)
            grdOrdena.TextMatrix(grdOrdena.row, 8) = IIf(nCondicion = 3, "SI", "NO")
            nPersoneria = IIf(IsNull(R!nPersPersoneria), 0, R!nPersPersoneria)  'NDX ERS0062020
            grdOrdena.TextMatrix(grdOrdena.row, 9) = nPersoneria                'NDX ERS0062020
            Set foPersona = Nothing
            'WIOR FIN *************************************************
            If sOcupacion <> "" Then
                Do While Trim(Str(lrsOcupa!nConsValor)) <> sOcupacion
                    lrsOcupa.MoveNext
                Loop
                grdOrdena.TextMatrix(grdOrdena.row, 5) = Trim(lrsOcupa!cConsDescripcion) & space(100) & Trim(Str(lrsOcupa!nConsValor))
            Else
                grdOrdena.TextMatrix(grdOrdena.row, 5) = sOcupacion
            End If
            lrsOcupa.Close
        End If
        Set ClsPersona = Nothing
        Set obj = Nothing
    End If
End Sub
Private Sub grdRealiza_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim nPersoneria As PersPersoneria
    Dim nCondicion As Integer 'WIOR 20130301
    
    If pbEsDuplicado Then
        MsgBox "Persona ya esta registrada en la relación.", vbInformation, "Aviso"
        grdRealiza.EliminaFila grdRealiza.row
    ElseIf psDataCod = "" Then
        grdRealiza.TextMatrix(grdRealiza.row, 2) = ""
        grdRealiza.TextMatrix(grdRealiza.row, 3) = ""
        grdRealiza.TextMatrix(grdRealiza.row, 4) = ""
        grdRealiza.TextMatrix(grdRealiza.row, 5) = ""
        Exit Sub
    Else
        Call llenar_cboFlex_ocupa(grdRealiza)
        Dim ClsPersona As COMDPersona.DCOMPersonas
        Dim R As New ADODB.Recordset
        Dim obj As COMDPersona.DCOMPersonas
        Dim lrsOcupa As ADODB.Recordset
        Dim sOcupacion As String
          
        Set obj = New COMDPersona.DCOMPersonas
        Set lrsOcupa = obj.CargarOcupaciones()
        Set ClsPersona = New COMDPersona.DCOMPersonas
        Set R = ClsPersona.BuscaCliente(grdRealiza.TextMatrix(grdRealiza.row, 1), 2)
         
        If Not (R.EOF And R.BOF) Then
            grdRealiza.TextMatrix(grdRealiza.row, 2) = IIf(IsNull(R!cPersIDnroDNI), "", IIf(R!cPersIDnroDNI = "", R!cPersIDnroRUC, R!cPersIDnroDNI))
            grdRealiza.TextMatrix(grdRealiza.row, 3) = PstaNombre(R!cPersNombre)
            grdRealiza.TextMatrix(grdRealiza.row, 4) = R!cPersDireccDomicilio
            sOcupacion = IIf(IsNull(R!cActiGiro1), "", IIf(R!cActiGiro1 = "", R!cActiGiro1, R!cActiGiro1))
            nPersoneria = IIf(IsNull(R!nPersPersoneria), 0, R!nPersPersoneria)
            'WIOR 20130301 ********************************************
            grdRealiza.TextMatrix(grdRealiza.row, 6) = PstaNombre(R!cNacionalidad)
            grdRealiza.TextMatrix(grdRealiza.row, 7) = PstaNombre(R!cResidente)
            Set foPersona = New COMDPersona.UCOMPersona
            Call foPersona.ValidaEnListaNegativaCondicion(IIf(IsNull(R!cPersIDnroDNI), "", R!cPersIDnroDNI), IIf(IsNull(R!cPersIDnroRUC), "", R!cPersIDnroRUC), nCondicion, R!cPersNombre)
            grdRealiza.TextMatrix(grdRealiza.row, 8) = IIf(nCondicion = 3, "SI", "NO")
            Set foPersona = Nothing
            'WIOR FIN *************************************************
            If nPersoneria = gPersonaNat Then
                If sOcupacion <> "" Then
                    Do While Trim(Str(lrsOcupa!nConsValor)) <> sOcupacion
                         lrsOcupa.MoveNext
                    Loop
                    grdRealiza.TextMatrix(grdRealiza.row, 5) = Trim(lrsOcupa!cConsDescripcion) & space(100) & Trim(Str(lrsOcupa!nConsValor))
                Else
                    grdRealiza.TextMatrix(grdRealiza.row, 5) = sOcupacion
                End If
                lrsOcupa.Close
            Else
                MsgBox "Las Personas que realizan la Transacción deben tener Personería Natural", vbInformation, "Aviso"
                grdRealiza.EliminaFila grdRealiza.row
                Set ClsPersona = Nothing
                Set obj = Nothing
                Exit Sub
            End If
        End If
        Set ClsPersona = Nothing
        Set obj = Nothing
    End If
End Sub
'END EJVG

'LLamar al evento Aceptar_click
Private Sub TxtClave_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call CmdAceptar_Click
End If
End Sub
'MADM 20100723
'Comentado x EJVG20120327
'Sub llenar_cbo_ocupa(Index As Integer)
'Dim obj As COMDPersona.DCOMPersonas
'Dim lrsOcupa As ADODB.Recordset 'madm 20100723
'
'Set obj = New COMDPersona.DCOMPersonas
'
'Set lrsOcupa = obj.CargarOcupaciones()
'Call Llenar_Combo_con_Recordset(lrsOcupa, cboocupa1(Index)) 'madm 20100723
'End Sub
'END MADM

'JACA 20110223
'Sub llenar_cboFlex_ocupa()
Sub llenar_cboFlex_ocupa(ByVal grilla As FlexEdit) 'EJVG20120327
Dim obj As COMDPersona.DCOMPersonas
Dim lrsOcupa As ADODB.Recordset

Set obj = New COMDPersona.DCOMPersonas
Set lrsOcupa = obj.CargarOcupaciones()

'grdBeneficiario.CargaCombo lrsOcupa
grilla.CargaCombo lrsOcupa 'EJVG201200327

End Sub

'JACA END
'Comentado x EJVG20120327
'Private Sub txtCodigo_EmiteDatos(Index As Integer)
'If TxtCodigo(Index).Text <> "" Then
'    Dim nPersoneria As PersPersoneria
'    nPersoneria = CLng(Trim(TxtCodigo(Index).PersPersoneria))
'    Call llenar_cbo_ocupa(Index)
'    'If nPersoneria = gPersonaNat Then Comentado x JACA 20110905
'        lblDocID(Index).Caption = TxtCodigo(Index).sPersNroDoc
'        lblNombre(Index).Caption = TxtCodigo(Index).psDescripcion
'        lblDireccion(Index).Caption = TxtCodigo(Index).sPersDireccion
'        cboocupa1(Index).ListIndex = IndiceListaCombo(cboocupa1(Index), TxtCodigo(Index).sPersOcupa) 'madm 20100723
'        CmdAceptar.SetFocus
''Comentado x JACA 20110905**************************************************
''   Else
''        MsgBox "Persona a registrar debe ser Persona Natural", vbInformation, "Aviso"
''        lblDocID(Index).Caption = ""
''        lblNombre(Index).Caption = ""
''        lblDireccion(Index).Caption = ""
''        TxtCodigo(Index).Text = ""
''        cboocupa1(Index).ListIndex = -1
''        TxtCodigo(Index).SetFocus
''    End If
''JACA END*****************************************************************
'End If
'End Sub

'Private Sub txtCodigo_EmiteDatos(Index As Integer)
'    If TxtCodigo(Index).Text <> "" Then
'        Dim nPersoneria As PersPersoneria
'        nPersoneria = CLng(Trim(TxtCodigo(Index).PersPersoneria))
'        Call llenar_cbo_ocupa(Index)
'        'If nPersoneria = gPersonaNat Then Comentado x JACA 20110905
'        lblDocID(Index).Caption = TxtCodigo(Index).sPersNroDoc
'        lblNombre(Index).Caption = TxtCodigo(Index).psDescripcion
'        lblDireccion(Index).Caption = TxtCodigo(Index).sPersDireccion
'        cboocupa1(Index).ListIndex = IndiceListaCombo(cboocupa1(Index), TxtCodigo(Index).sPersOcupa) 'madm 20100723
'        CmdAceptar.SetFocus
'    End If
'End Sub

Public Sub imprimirBoletaREU(Optional psCta As String = "", Optional ByVal pnMoneda As Integer = 1, Optional ByVal psOrigenEfectivo As String = "", Optional pnNroREU As String = "0")
Dim loBoleta As COMNCaptaGenerales.NCOMCaptaImpresion
Dim lsBoleta As String
Dim lnFicSal As Integer
Dim lbOk As Boolean

   
            If psCta <> "" Then
                sCuenta = psCta
                '''sMoneda = IIf(pnMoneda = 1, "NUEVOS SOLES", "US DOLARES") 'marg ers044-2016
                sMoneda = IIf(pnMoneda = 1, StrConv(gcPEN_PLURAL, vbUpperCase), "US DOLARES") 'marg ers044-2016
                nMoneda = pnMoneda
            End If
        
            Set loBoleta = New COMNCaptaGenerales.NCOMCaptaImpresion
                loBoleta.IniciaImpresora gImpresora
            'ALPA 20081013**************************************************
            
            'By capi 19022009 se incluyo al parametro nMoneda
            
        '    lsBoleta = loBoleta.ImprimeBoletaLavadoDinero(gsNomCmac, gsNomAge, gdFecSis, sCuenta, fTitPersLavDineroNom, fTitPersLavDineroDoc, fTitPersLavDineroDir, _
        '               fReaPersLavDineroNom, fReaPersLavDineroDoc, fReaPersLavDineroDir, _
        '               fBenPersLavDineroNom, fBenPersLavDineroDoc, fBenPersLavDineroDir, _
        '               sOperacion, nMontoTransaccion, sLpt, _
        '               fOrdPersLavDineroNom, fOrdPersLavDineroDoc, fOrdPersLavDineroDir, , True, sTipoCuenta, , gbImpTMU, gsCodAge, psOrigenEfectivo)
        '
            'JACA 20110224
            
        '    lsBoleta = loBoleta.ImprimeBoletaLavadoDinero(gsNomCmac, gsNomAge, gdFecSis, sCuenta, fTitPersLavDineroNom, fTitPersLavDineroDoc, fTitPersLavDineroDir, _
        '               fReaPersLavDineroNom, fReaPersLavDineroDoc, fReaPersLavDineroDir, _
        '               fBenPersLavDineroNom, fBenPersLavDineroDoc, fBenPersLavDineroDir, _
        '               sOperacion, nMontoTransaccion, sLpt, _
        '              fOrdPersLavDineroNom, fOrdPersLavDineroDoc, fOrdPersLavDineroDir, , True, sTipoCuenta, pnMoneda, gbImpTMU, gsCodAge, psOrigenEfectivo)
            
            'EJVG20120329
            Dim ListaPersonasRealizan() As PersonaLavado
            Dim ListaPersonasOrdenan() As PersonaLavado
            Dim ListaPersonasBenefician() As PersonaLavado 'WIOR 20130301
            'Realizan
            'WIOR 20130301 ***********************************************
            If fReaPersLavDinero <> "" Then
                ReDim Preserve ListaPersonasRealizan(0)
                ListaPersonasRealizan(0).PersCod = fReaPersLavDinero
                ListaPersonasRealizan(0).Nombre = fReaPersLavDineroNom
                ListaPersonasRealizan(0).DocumentoId = fReaPersLavDineroDoc
                ListaPersonasRealizan(0).Direccion = fReaPersLavDineroDir
                ListaPersonasRealizan(0).Ocupacion = lsReaPersLavDineroOcu
                ListaPersonasRealizan(0).Nacionalidad = fReaPersLavDineroNac
                ListaPersonasRealizan(0).Residente = fReaPersLavDineroRes
                ListaPersonasRealizan(0).Peps = fReaPersLavDineroPeps
            End If
            'WIOR 20130301 ***********************************************
            If fReaPersLavDinero2 <> "" And fReaPersLavDinero <> "" Then 'WIOR 20130301 adelanto uno los datos y agrego (And fReaPersLavDinero <> "")
                ReDim Preserve ListaPersonasRealizan(1)
                ListaPersonasRealizan(1).PersCod = fReaPersLavDinero2
                ListaPersonasRealizan(1).Nombre = fReaPersLavDineroNom2
                ListaPersonasRealizan(1).DocumentoId = fReaPersLavDineroDoc2
                ListaPersonasRealizan(1).Direccion = fReaPersLavDineroDir2
                ListaPersonasRealizan(1).Ocupacion = lsReaPersLavDineroOcu2
                'WIOR 20130301 ***********************************************
                ListaPersonasRealizan(1).Nacionalidad = fReaPersLavDineroNac2
                ListaPersonasRealizan(1).Residente = fReaPersLavDineroRes2
                ListaPersonasRealizan(1).Peps = fReaPersLavDineroPeps2
                'WIOR FIN ****************************************************
            End If
            If fReaPersLavDinero3 <> "" And fReaPersLavDinero2 <> "" Then 'WIOR 20130301 adelanto uno los datos
                ReDim Preserve ListaPersonasRealizan(2)
                ListaPersonasRealizan(2).PersCod = fReaPersLavDinero3
                ListaPersonasRealizan(2).Nombre = fReaPersLavDineroNom3
                ListaPersonasRealizan(2).DocumentoId = fReaPersLavDineroDoc3
                ListaPersonasRealizan(2).Direccion = fReaPersLavDineroDir3
                ListaPersonasRealizan(2).Ocupacion = lsReaPersLavDineroOcu3
                'WIOR 20130301 ***********************************************
                ListaPersonasRealizan(2).Nacionalidad = fReaPersLavDineroNac3
                ListaPersonasRealizan(2).Residente = fReaPersLavDineroRes3
                ListaPersonasRealizan(2).Peps = fReaPersLavDineroPeps3
                'WIOR FIN ****************************************************
            End If
            If fReaPersLavDinero4 <> "" And fReaPersLavDinero3 <> "" Then 'WIOR 20130301 adelanto uno los datos
                ReDim Preserve ListaPersonasRealizan(3)
                ListaPersonasRealizan(3).PersCod = fReaPersLavDinero4 'WIOR 20130301 fReaPersLavDinero3
                ListaPersonasRealizan(3).Nombre = fReaPersLavDineroNom4 'WIOR 20130301 fReaPersLavDineroNom3
                ListaPersonasRealizan(3).DocumentoId = fReaPersLavDineroDoc4 'WIOR 20130301 fReaPersLavDineroDoc3
                ListaPersonasRealizan(3).Direccion = fReaPersLavDineroDir4 'WIOR 20130301 fReaPersLavDineroDir3
                ListaPersonasRealizan(3).Ocupacion = lsReaPersLavDineroOcu4 'WIOR 20130301 lsReaPersLavDineroOcu3
                'WIOR 20130301 ***********************************************
                ListaPersonasRealizan(3).Nacionalidad = fReaPersLavDineroNac4
                ListaPersonasRealizan(3).Residente = fReaPersLavDineroRes4
                ListaPersonasRealizan(3).Peps = fReaPersLavDineroPeps4
                'WIOR FIN ****************************************************
            End If
            'Ordenan
            'WIOR 20130301 ***********************************************
            If fOrdPersLavDinero <> "" Then
                ReDim Preserve ListaPersonasOrdenan(0)
                ListaPersonasOrdenan(0).PersCod = fOrdPersLavDinero
                ListaPersonasOrdenan(0).Nombre = fOrdPersLavDineroNom
                ListaPersonasOrdenan(0).DocumentoId = fOrdPersLavDineroDoc
                ListaPersonasOrdenan(0).Direccion = fOrdPersLavDineroDir
                ListaPersonasOrdenan(0).Ocupacion = lsOrdPersLavDineroOcu
                ListaPersonasOrdenan(0).Nacionalidad = fOrdPersLavDineroNac
                ListaPersonasOrdenan(0).Residente = fOrdPersLavDineroRes
                ListaPersonasOrdenan(0).Peps = fOrdPersLavDineroPeps
            End If
            'WIOR FIN ****************************************************
            If fOrdPersLavDinero2 <> "" And fOrdPersLavDinero <> "" Then 'WIOR 20130301 adelanto uno los datos y agrego (And fReaPersLavDinero <> "")
                ReDim Preserve ListaPersonasOrdenan(1)
                ListaPersonasOrdenan(1).PersCod = fOrdPersLavDinero2
                ListaPersonasOrdenan(1).Nombre = fOrdPersLavDineroNom2
                ListaPersonasOrdenan(1).DocumentoId = fOrdPersLavDineroDoc2
                ListaPersonasOrdenan(1).Direccion = fOrdPersLavDineroDir2
                ListaPersonasOrdenan(1).Ocupacion = lsOrdPersLavDineroOcu2
                'WIOR 20130301 ***********************************************
                ListaPersonasOrdenan(1).Nacionalidad = fOrdPersLavDineroNac2
                ListaPersonasOrdenan(1).Residente = fOrdPersLavDineroRes2
                ListaPersonasOrdenan(1).Peps = fOrdPersLavDineroPeps2
                'WIOR FIN ****************************************************
            End If
            If fOrdPersLavDinero3 <> "" And fOrdPersLavDinero2 <> "" Then 'WIOR 20130301 adelanto uno los datos
                ReDim Preserve ListaPersonasOrdenan(2)
                ListaPersonasOrdenan(2).PersCod = fOrdPersLavDinero3
                ListaPersonasOrdenan(2).Nombre = fOrdPersLavDineroNom3
                ListaPersonasOrdenan(2).DocumentoId = fOrdPersLavDineroDoc3
                ListaPersonasOrdenan(2).Direccion = fOrdPersLavDineroDir3
                ListaPersonasOrdenan(2).Ocupacion = lsOrdPersLavDineroOcu3
                'WIOR 20130301 ***********************************************
                ListaPersonasOrdenan(2).Nacionalidad = fOrdPersLavDineroNac3
                ListaPersonasOrdenan(2).Residente = fOrdPersLavDineroRes3
                ListaPersonasOrdenan(2).Peps = fOrdPersLavDineroPeps3
                'WIOR FIN ****************************************************
            End If
            If fOrdPersLavDinero4 <> "" And fOrdPersLavDinero3 <> "" Then 'WIOR 20130301 adelanto uno los datos
                ReDim Preserve ListaPersonasOrdenan(3)
                ListaPersonasOrdenan(3).PersCod = fOrdPersLavDinero4
                ListaPersonasOrdenan(3).Nombre = fOrdPersLavDineroNom4
                ListaPersonasOrdenan(3).DocumentoId = fOrdPersLavDineroDoc4
                ListaPersonasOrdenan(3).Direccion = fOrdPersLavDineroDir4
                ListaPersonasOrdenan(3).Ocupacion = lsOrdPersLavDineroOcu4
                'WIOR 20130301 ***********************************************
                ListaPersonasOrdenan(3).Nacionalidad = fOrdPersLavDineroNac4
                ListaPersonasOrdenan(3).Residente = fOrdPersLavDineroRes4
                ListaPersonasOrdenan(3).Peps = fOrdPersLavDineroPeps4
                'WIOR FIN ****************************************************
            End If
            'WIOR 20130301 ***********************************************
            'Beneficia
            If fBenPersLavDinero <> "" Then
                ReDim Preserve ListaPersonasBenefician(0)
                ListaPersonasBenefician(0).PersCod = fBenPersLavDinero
                ListaPersonasBenefician(0).Nombre = fBenPersLavDineroNom
                ListaPersonasBenefician(0).DocumentoId = fBenPersLavDineroDoc
                ListaPersonasBenefician(0).Direccion = fBenPersLavDineroDir
                ListaPersonasBenefician(0).Ocupacion = lsBenPersLavDineroOcu
                ListaPersonasBenefician(0).Nacionalidad = fBenPersLavDineroNac
                ListaPersonasBenefician(0).Residente = fBenPersLavDineroRes
                ListaPersonasBenefician(0).Peps = fBenPersLavDineroPeps
            End If
            If fBenPersLavDinero2 <> "" And fBenPersLavDinero <> "" Then
                ReDim Preserve ListaPersonasBenefician(1)
                ListaPersonasBenefician(1).PersCod = fBenPersLavDinero2
                ListaPersonasBenefician(1).Nombre = fBenPersLavDineroNom2
                ListaPersonasBenefician(1).DocumentoId = fBenPersLavDineroDoc2
                ListaPersonasBenefician(1).Direccion = fBenPersLavDineroDir2
                ListaPersonasBenefician(1).Ocupacion = lsBenPersLavDineroOcu2
                ListaPersonasBenefician(1).Nacionalidad = fBenPersLavDineroNac2
                ListaPersonasBenefician(1).Residente = fBenPersLavDineroRes2
                ListaPersonasBenefician(1).Peps = fBenPersLavDineroPeps2
            End If
            If fBenPersLavDinero3 <> "" And fBenPersLavDinero2 <> "" Then
                ReDim Preserve ListaPersonasBenefician(2)
                ListaPersonasBenefician(2).PersCod = fBenPersLavDinero3
                ListaPersonasBenefician(2).Nombre = fBenPersLavDineroNom3
                ListaPersonasBenefician(2).DocumentoId = fBenPersLavDineroDoc3
                ListaPersonasBenefician(2).Direccion = fBenPersLavDineroDir3
                ListaPersonasBenefician(2).Ocupacion = lsBenPersLavDineroOcu3
                ListaPersonasBenefician(2).Nacionalidad = fBenPersLavDineroNac3
                ListaPersonasBenefician(2).Residente = fBenPersLavDineroRes3
                ListaPersonasBenefician(2).Peps = fBenPersLavDineroPeps3
            End If
            If fBenPersLavDinero4 <> "" And fBenPersLavDinero3 <> "" Then
                ReDim Preserve ListaPersonasBenefician(3)
                ListaPersonasBenefician(3).PersCod = fBenPersLavDinero4
                ListaPersonasBenefician(3).Nombre = fBenPersLavDineroNom4
                ListaPersonasBenefician(3).DocumentoId = fBenPersLavDineroDoc4
                ListaPersonasBenefician(3).Direccion = fBenPersLavDineroDir4
                ListaPersonasBenefician(3).Ocupacion = lsBenPersLavDineroOcu4
                ListaPersonasBenefician(3).Nacionalidad = fBenPersLavDineroNac4
                ListaPersonasBenefician(3).Residente = fBenPersLavDineroRes4
                ListaPersonasBenefician(3).Peps = fBenPersLavDineroPeps4
            End If
            'WIOR FIN ****************************************************
            
            lsBoleta = loBoleta.ImprimeBoletaLavadoDinero(gsNomCmac, gsNomAge, gdFecSis, sCuenta, fTitPersLavDineroNom, fTitPersLavDineroDoc, fTitPersLavDineroDir, fTitPersLavDineroOcu, _
                       fReaPersLavDineroNom, fReaPersLavDineroDoc, fReaPersLavDineroDir, lsReaPersLavDineroOcu, _
                       fBenPersLavDineroNom, fBenPersLavDineroDoc, fBenPersLavDineroDir, lsBenPersLavDineroOcu, _
                       sOperacion, nMontoTransaccion, sLpt, _
                       fOrdPersLavDineroNom, fOrdPersLavDineroDoc, fOrdPersLavDineroDir, lsOrdPersLavDineroOcu, , True, sTipoCuenta, pnMoneda, gbImpTMU, gsCodAge, psOrigenEfectivo, _
                       fBenPersLavDineroNom2, fBenPersLavDineroDoc2, fBenPersLavDineroDir2, lsBenPersLavDineroOcu2, _
                       fBenPersLavDineroNom3, fBenPersLavDineroDoc3, fBenPersLavDineroDir3, lsBenPersLavDineroOcu3, _
                       fBenPersLavDineroNom4, fBenPersLavDineroDoc4, fBenPersLavDineroDir4, lsBenPersLavDineroOcu4, lnTipoREU, pnNroREU, ListaPersonasRealizan, ListaPersonasOrdenan, _
                       ListaPersonasBenefician) 'EJVG20120327 Se agrego las Listas Tipificadas
                       'WIOR 20130301 agrego ListaPersonasBenefician
              'JACA END
            'Se agrego el parametro psOrigenEfectivo
            Set loBoleta = Nothing
                
            lbOk = True
            Do While lbOk
                 lnFicSal = FreeFile
                 Open sLpt For Output As lnFicSal
                     Print #lnFicSal, lsBoleta
                     Print #lnFicSal, ""
                 Close #lnFicSal
                 If MsgBox("Desea Reimprimir Boleta REU ??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                     lbOk = False
                 End If
            Loop
    
End Sub

'JACA 20110224
'Public Sub InsertarLavDinero(Optional sPersLavDinero As String = "", Optional bTransaccion As Boolean = False, Optional pCon As ADODB.Connection = Nothing, Optional nmovnro As Long = 0, Optional sBenPersLavDinero As String, _
'            Optional psTitPersLavDinero As String = "", Optional psOrdPersLavDinero As String = "", Optional psReaPersLavDinero As String = "", _
'            Optional psBenPersLavDinero As String = "", Optional psVisPersLavDinero As String = "", Optional nTipoREU As Integer = 1, Optional nMontoAcu As Double = 0, Optional sOrigen As String = "")
Public Sub InsertarLavDinero(Optional sPersLavDinero As String = "", Optional bTransaccion As Boolean = False, Optional pCon As ADODB.Connection = Nothing, Optional nMovNro As Long = 0, Optional sBenPersLavDinero As String, _
            Optional psTitPersLavDinero As String = "", Optional psOrdPersLavDinero As String = "", Optional psReaPersLavDinero As String = "", _
            Optional psBenPersLavDinero As String = "", Optional psVisPersLavDinero As String = "", Optional nTipoREU As Integer = 1, Optional nMontoAcu As Double = 0, Optional sOrigen As String = "", _
            Optional psBenPersLavDinero2 As String = "", Optional psBenPersLavDinero3 As String = "", Optional psBenPersLavDinero4 As String = "")
            
            Dim clsCap As COMDCaptaGenerales.DCOMCaptaMovimiento
            Set clsCap = New COMDCaptaGenerales.DCOMCaptaMovimiento
            
            'EJVG20120327
            Dim clsGen As COMDCaptaGenerales.DCOMCaptaGenerales
            Set clsGen = New COMDCaptaGenerales.DCOMCaptaGenerales
            
            If Not pCon Is Nothing Then
                    clsCap.SetConexion pCon
                    clsCap.bTransaccion = True
                    bTransaccion = True
            End If
            If psOrdPersLavDinero <> "" And psOrdPersLavDinero <> "Exit" Then
                'clsCap.AgregaMovLavDinero nmovnro, sPersLavDinero, psTitPersLavDinero, psOrdPersLavDinero, psReaPersLavDinero, psBenPersLavDinero, psVisPersLavDinero, nTipoREU, nMontoAcu, sOrigen
                 'clsCap.AgregaMovLavDinero nmovnro, sPersLavDinero, psTitPersLavDinero, psOrdPersLavDinero, psReaPersLavDinero, psBenPersLavDinero, psVisPersLavDinero, nTipoREU, nMontoAcu, sOrigen, psBenPersLavDinero2, psBenPersLavDinero3, psBenPersLavDinero4 ' JACA 20110224
                 'EJVG20120327 Relacion Persona con Lavado Dinero
                 clsCap.AgregaMovLavDinero nMovNro, psTitPersLavDinero, psVisPersLavDinero, nTipoREU, nMontoAcu, sOrigen
                 
                 'MARG ERS052-2017----
                 Dim oVisto As frmVistoElectronico
                 Set oVisto = New frmVistoElectronico
                 oVisto.RegistrarVistoElectronicoLavDinero "910000", psVisPersLavDinero, sOrigen, psTitPersLavDinero, gsCodUser, nMovNro
                'END MARG- -------------
                
                 Call clsGen.AgregaMovLavDineroPersona(nMovNro, psReaPersLavDinero, RealizaTransaccion, chkRealiza.value)                                       'NDX ERS0062020 SE AGREGÓ chkRealiza.value
                 If fReaPersLavDinero2 <> "" Then Call clsGen.AgregaMovLavDineroPersona(nMovNro, fReaPersLavDinero2, RealizaTransaccion, chkRealiza.value)      'NDX ERS0062020 SE AGREGÓ chkRealiza.value
                 If fReaPersLavDinero3 <> "" Then Call clsGen.AgregaMovLavDineroPersona(nMovNro, fReaPersLavDinero3, RealizaTransaccion, chkRealiza.value)      'NDX ERS0062020 SE AGREGÓ chkRealiza.value
                 If fReaPersLavDinero4 <> "" Then Call clsGen.AgregaMovLavDineroPersona(nMovNro, fReaPersLavDinero4, RealizaTransaccion, chkRealiza.value)      'NDX ERS0062020 SE AGREGÓ chkRealiza.value
                 Call clsGen.AgregaMovLavDineroPersona(nMovNro, psOrdPersLavDinero, OrdenaTransaccion, chkOrdena.value)                                         'NDX ERS0062020 SE AGREGÓ chkOrdena.value
                 If fOrdPersLavDinero2 <> "" Then Call clsGen.AgregaMovLavDineroPersona(nMovNro, fOrdPersLavDinero2, OrdenaTransaccion, chkOrdena.value)        'NDX ERS0062020 SE AGREGÓ chkOrdena.value
                 If fOrdPersLavDinero3 <> "" Then Call clsGen.AgregaMovLavDineroPersona(nMovNro, fOrdPersLavDinero3, OrdenaTransaccion, chkOrdena.value)        'NDX ERS0062020 SE AGREGÓ chkOrdena.value
                 If fOrdPersLavDinero4 <> "" Then Call clsGen.AgregaMovLavDineroPersona(nMovNro, fOrdPersLavDinero4, OrdenaTransaccion, chkOrdena.value)        'NDX ERS0062020 SE AGREGÓ chkOrdena.value
                 Call clsGen.AgregaMovLavDineroPersona(nMovNro, psBenPersLavDinero, BeneficiaTransaccion, chkBeneficia.value)                                   'NDX ERS0062020 SE AGREGÓ chkBeneficia.value
                 If fBenPersLavDinero2 <> "" Then Call clsGen.AgregaMovLavDineroPersona(nMovNro, fBenPersLavDinero2, BeneficiaTransaccion, chkBeneficia.value)  'NDX ERS0062020 SE AGREGÓ chkBeneficia.value
                 If fBenPersLavDinero3 <> "" Then Call clsGen.AgregaMovLavDineroPersona(nMovNro, fBenPersLavDinero3, BeneficiaTransaccion, chkBeneficia.value)  'NDX ERS0062020 SE AGREGÓ chkBeneficia.value
                 If fBenPersLavDinero4 <> "" Then Call clsGen.AgregaMovLavDineroPersona(nMovNro, fBenPersLavDinero4, BeneficiaTransaccion, chkBeneficia.value)  'NDX ERS0062020 SE AGREGÓ chkBeneficia.value
                 
                 
                 'JACA 20110929 para imprimir con el correlativo de nMovNro****
                    Dim objREU As New COMDPersona.DCOMPersonas
                    Dim rsCorr As New Recordset
                        Set rsCorr = objREU.ObtenerAgenCorrelativoLavDinero(nMovNro)
                        If Not (rsCorr.EOF And rsCorr.BOF) Then
                           fnNroREU = rsCorr!nNroREU
                        Else
                           fnNroREU = 0
                        End If
                       
                       If Imprimir Then
                            imprimirBoletaREU CtaImprimir, , sOrigen, fnNroREU
                        End If
                        
                 'JACA END*****************************************************
                 
            End If
            'WIOR 20121023 ******************************************************
            If nTipoREU = 2 And Trim(fOrdPersLavDinero) <> "Exit" And Trim(fOrdPersLavDinero) <> "" Then
                Dim Cont As Integer
                Dim nTC As Double
                Dim oDPersona As COMDPersona.DCOMPersona
                Dim oDPersonaMov As COMDPersona.DCOMPersona
                Dim rsDPersona As ADODB.Recordset
                Dim rsDPersonaMov As ADODB.Recordset
            
                Set oDPersona = New COMDPersona.DCOMPersona
                Set oDPersonaMov = New COMDPersona.DCOMPersona
            
                Dim clsTC As COMDConstSistema.NCOMTipoCambio
                Set clsTC = New COMDConstSistema.NCOMTipoCambio
                nTC = clsTC.EmiteTipoCambio(gdFecSis, TCPondREU)
                Set clsTC = Nothing
            
                Dim lnPerido As Integer
                lnPerido = 0

                Set rsDPersona = oDPersona.ObtenerPersOperaciones(psTitPersLavDinero)
                If rsDPersona.RecordCount > 0 Then
                    If Trim(rsDPersona!cPeriodo) = Trim(Mid(Format(gdFecSis, "yyyymmdd"), 1, 6)) Then
                        lnPerido = CInt(rsDPersona!nPeriodo)
                        lnPerido = lnPerido + 1
                        Call oDPersona.InsertaActualizaPersOperaciones(2, psTitPersLavDinero, Trim(rsDPersona!cPeriodo), lnPerido)
                    Else
                        Call oDPersona.InsertaActualizaPersOperaciones(2, psTitPersLavDinero, Trim(Mid(Format(gdFecSis, "yyyymmdd"), 1, 6)), 1)
                    End If
                Else
                    Call oDPersona.InsertaActualizaPersOperaciones(1, psTitPersLavDinero, Trim(Mid(Format(gdFecSis, "yyyymmdd"), 1, 6)), 1)
                End If
                
                If lnPerido = 0 Then
                    lnPerido = 1
                End If
            
                Set rsDPersonaMov = oDPersonaMov.ObtenerPersOperacionesDet(psTitPersLavDinero, nTC, Trim(Mid(Format(gdFecSis, "yyyymmdd"), 1, 6)))
                If rsDPersonaMov.RecordCount > 0 Then
                    If Not (rsDPersonaMov.EOF And rsDPersonaMov.BOF) Then
                        For Cont = 0 To rsDPersonaMov.RecordCount - 1
                            Call oDPersonaMov.InsertaPersOperacionesDet(psTitPersLavDinero, rsDPersonaMov!nMovNro, fnNroREU, Trim(Mid(Format(gdFecSis, "yyyymmdd"), 1, 6)), lnPerido, CInt(rsDPersonaMov!Moneda), CDbl(rsDPersonaMov!Monto))
                            rsDPersonaMov.MoveNext
                        Next Cont
                    End If
                End If
            End If
            'WIOR FIN ***********************************************************
            Set clsCap = Nothing
End Sub
'JACA END


Private Sub txtOrigen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
           If Me.cmdGrabar.Enabled = True Then 'JACA 20111119
                Me.cmdGrabar.SetFocus
           Else
                Me.CmdAceptar.SetFocus
           End If
    End If
End Sub
Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            Me.TxtClave.SetFocus
    End If
    
End Sub

' *** RIRO SEGUN TI-ERS108-2013 ***
Public Function ValidacionRFIII() As Boolean
    Dim rsRF3 As New ADODB.Recordset
    Dim oDCOMMov As COMDMov.DCOMMov
    Dim sMensaje As String
    Dim sCodCargo As String
    Dim sMovNro As String
        
    ctlUsuario.Inicio (Trim(txtUsuario.Text))
    sCodCargo = ctlUsuario.PersCargoCod
    Set rsRF3 = ValidarRFIII
    sMensaje = ""
    
    '*** VERIFICANDO SI RFFIII ESTA EN MODO SUPERVISOR
        Dim oAcceso As UCOMAcceso
        Dim clsGen As COMDConstSistema.DCOMGeneral
        Dim sGrupos, sTemporal, sGrupoRF3 As String
        Dim bModoSupervisor As Boolean
        
        Set oAcceso = New UCOMAcceso
        Set clsGen = New COMDConstSistema.DCOMGeneral
        sGrupoRF3 = clsGen.GetConstante(10027, , "100", "1")!cDescripcion
        sTemporal = ""
        sGrupos = ""
        
        If oAcceso.VerificarUsuarioExistaEnRRHH(Trim(txtUsuario.Text)) Then
            Call oAcceso.CargaGruposUsuario(Trim(txtUsuario.Text), gsDominio)
            sTemporal = oAcceso.DameGrupoUsuario
            Do While Len(sTemporal) > 0
                sGrupos = sGrupos & sTemporal & ","
                sTemporal = oAcceso.DameGrupoUsuario
            Loop
        End If

        sGrupos = Mid(sGrupos, 1, Len(sGrupos) - IIf(Len(sGrupos) > 0, 1, 0))
        Set oAcceso = Nothing
        Set clsGen = Nothing
        If InStr(1, sGrupos, sGrupoRF3) > 0 Then
        
            Dim rsRF3User As New ADODB.Recordset
            Dim oPersona As New COMDPersona.DCOMPersonas
            Set rsRF3User = oPersona.RecuperarGruposRF3(Trim(txtUsuario.Text))
            
            If Not rsRF3User Is Nothing Then
                If Not rsRF3User.BOF And Not rsRF3User.EOF Then
                    If rsRF3User!nEstado = 1 Then
                        bModoSupervisor = True
                    Else
                        bModoSupervisor = False
                    End If
                Else
                    bModoSupervisor = False
                End If
            Else
                bModoSupervisor = False
            End If
        
        Else
            bModoSupervisor = False
        End If
    '*** FIN VERIFICACION
    
    If Not (rsRF3.EOF Or rsRF3.BOF) And rsRF3.RecordCount > 0 Then
        If sCodCargo = "006005" Then     ' *** SI ES "SUPERVISOR"
            If Not rsRF3!bOpcionesSimultaneas And rsRF3!bModoSupervisor Then
                sMensaje = "No es posible emitir el VB virtual porque el RFIII se encuentra activo en modo supervisor si desea dar VB usted debe " & vbNewLine & _
                "desactivar al RFIII en este perfil, de lo contrario solicite a él que emita el VB correspondiente"
            End If
        ElseIf sCodCargo = "007026" Then ' *** SI ES "RFIII"
            If Not bModoSupervisor Then
                sMensaje = "Actualmente no cuenta con permisos para dar VB por lo que el VB virtual debe darlo el supervisor " & _
                           "de operaciones."
            End If
        End If
    End If
    If Len(sMovNro) > 0 Then
        If UCase(Right(sMovNro, 4)) = UCase(Trim(txtUsuario.Text)) Then
            sMensaje = "No puede dar su VB a operaciones realizadas por usted mismo"
        End If
    End If
    If Len(sMensaje) > 0 Then
        MsgBox sMensaje, vbExclamation, "Aviso"
        ValidacionRFIII = False
        Exit Function
    End If
    ValidacionRFIII = True
End Function
' *** FIN RIRO ***
'***JGPA20190610-----------------------------------------
Public Function ObtenerDatosReimpresionReu(ByVal psCodPersona As String) As Variant
    Dim rsDG As ADODB.Recordset
    Dim ClsPersona As COMDPersona.DCOMPersonas
    Dim lrsOcupa As ADODB.Recordset
    Dim sOcupacion As String
    Dim obj As COMDPersona.DCOMPersonas
    Dim MatDatosInterv As Variant
    Dim i As Integer
    Dim foPersona As COMDPersona.UCOMPersona
    Dim nCondicion As Integer
    
    On Error GoTo ErrObtenerDatosReimpresionReu
    
    Set obj = New COMDPersona.DCOMPersonas
    Set lrsOcupa = obj.CargarOcupaciones()
    Set ClsPersona = New COMDPersona.DCOMPersonas
    Set foPersona = New COMDPersona.UCOMPersona
    
    Set rsDG = ClsPersona.BuscaCliente(psCodPersona, BusquedaCodigo)
    
    If Not (rsDG.BOF And rsDG.EOF) Then
        ReDim MatDatosInterv(rsDG.RecordCount, 8)
        For i = 1 To rsDG.RecordCount
            MatDatosInterv(i, 1) = rsDG!cperscod
            MatDatosInterv(i, 2) = PstaNombre(rsDG!cPersNombre)
            MatDatosInterv(i, 3) = IIf(IsNull(rsDG!cPersIDnroDNI), "", IIf(rsDG!cPersIDnroDNI = "", rsDG!cPersIDnroRUC, rsDG!cPersIDnroDNI))
            MatDatosInterv(i, 4) = rsDG!cPersDireccDomicilio
            sOcupacion = IIf(IsNull(rsDG!cActiGiro1), "", IIf(rsDG!cActiGiro1 = "", rsDG!cActiGiro1, rsDG!cActiGiro1))
            If sOcupacion <> "" Then
                Do While Trim(Str(lrsOcupa!nConsValor)) <> sOcupacion
                    lrsOcupa.MoveNext
                Loop
                MatDatosInterv(i, 5) = Trim(lrsOcupa!cConsDescripcion) & space(100) & Trim(Str(lrsOcupa!nConsValor))
            Else
                MatDatosInterv(i, 5) = sOcupacion
            End If
            
            MatDatosInterv(i, 6) = PstaNombre(rsDG!cNacionalidad)
            MatDatosInterv(i, 7) = PstaNombre(rsDG!cResidente)
            Call foPersona.ValidaEnListaNegativaCondicion(IIf(IsNull(rsDG!cPersIDnroDNI), "", rsDG!cPersIDnroDNI), IIf(IsNull(rsDG!cPersIDnroRUC), "", rsDG!cPersIDnroRUC), nCondicion, rsDG!cPersNombre)
            MatDatosInterv(i, 8) = IIf(nCondicion = 3, "SI", "NO")
        Next i
    End If
    
    ObtenerDatosReimpresionReu = MatDatosInterv
    
    RSClose lrsOcupa
    
    Set foPersona = Nothing
    Set ClsPersona = Nothing
    Set obj = Nothing
    Exit Function
ErrObtenerDatosReimpresionReu:
    err.Raise err.Number, "Obtener Datos Reimprime Reu", err.Description
End Function
'***End JGPA---------------------------------------------

'---- NDX ERS0062020



'---- NDX ERS0062020 END
