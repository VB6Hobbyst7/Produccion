VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredMicroseguro 
   Caption         =   "Registro de Microseguro/Multiriesgo"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Height          =   780
      Left            =   120
      TabIndex        =   27
      Top             =   4440
      Width           =   9735
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   32
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   240
         Width           =   1125
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2280
         TabIndex        =   31
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   240
         Width           =   1125
      End
      Begin VB.CommandButton CmdLimpiar 
         Caption         =   "&Limpiar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7200
         TabIndex        =   30
         Top             =   240
         Width           =   1125
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8400
         TabIndex        =   29
         Top             =   240
         Width           =   1125
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
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
         Height          =   390
         Left            =   1200
         TabIndex        =   28
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdExaminar 
      Caption         =   "E&xaminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6240
      TabIndex        =   1
      Top             =   240
      Width           =   1650
   End
   Begin SICMACT.ActXCodCta ActXCodCta 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3645
      _extentx        =   6429
      _extenty        =   688
      texto           =   "Credito"
      enabledcmac     =   -1  'True
      enabledcta      =   -1  'True
      enabledprod     =   -1  'True
      enabledage      =   -1  'True
   End
   Begin TabDlg.SSTab sstRegistro 
      Height          =   3525
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   6218
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Microseguros"
      TabPicture(0)   =   "frmCredMicroseguro.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtNumCertificadoMicro"
      Tab(0).Control(1)=   "cmdEditarPersona"
      Tab(0).Control(2)=   "framePrima"
      Tab(0).Control(3)=   "cmdCancelarPersona"
      Tab(0).Control(4)=   "cmdAgregarPersona"
      Tab(0).Control(5)=   "cmdAceptarPersona"
      Tab(0).Control(6)=   "cmdEliminarPersona"
      Tab(0).Control(7)=   "FERelPers"
      Tab(0).Control(8)=   "Label4"
      Tab(0).Control(9)=   "Label3"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Póliza Multiriesgo"
      TabPicture(1)   =   "frmCredMicroseguro.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblTotalValor"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "feMuebles"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdEliminarMueble"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdAceptarMueble"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdAgregarMueble"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdCancelarMueble"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdEditarMueble"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtNumCertificadoMult"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      Begin VB.TextBox txtNumCertificadoMult 
         Height          =   345
         Left            =   3600
         TabIndex        =   40
         Top             =   400
         Width           =   2055
      End
      Begin VB.TextBox txtNumCertificadoMicro 
         Height          =   345
         Left            =   -70560
         TabIndex        =   39
         Top             =   640
         Width           =   2535
      End
      Begin VB.CommandButton cmdEditarPersona 
         Caption         =   "&Editar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -71280
         TabIndex        =   38
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   2880
         Width           =   1605
      End
      Begin VB.CommandButton cmdEditarMueble 
         Caption         =   "&Editar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3480
         TabIndex        =   37
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   3000
         Width           =   1605
      End
      Begin VB.Frame framePrima 
         Caption         =   "Prima Total Mensual"
         Height          =   615
         Left            =   -74760
         TabIndex        =   34
         Top             =   360
         Width           =   2415
         Begin VB.Label lblPrima 
            AutoSize        =   -1  'True
            Caption         =   "S/. 0.00"
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
            Left            =   720
            TabIndex        =   35
            Top             =   240
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdCancelarMueble 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         TabIndex        =   26
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.CommandButton cmdCancelarPersona 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -72960
         TabIndex        =   25
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.CommandButton cmdAgregarMueble 
         Caption         =   "&Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   23
         Top             =   3000
         Width           =   1605
      End
      Begin VB.CommandButton cmdAceptarMueble 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.CommandButton cmdEliminarMueble 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         TabIndex        =   21
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   3000
         Width           =   1605
      End
      Begin VB.CommandButton cmdAgregarPersona 
         Caption         =   "&Agregar Persona"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -74640
         TabIndex        =   20
         Top             =   2880
         Width           =   1605
      End
      Begin VB.CommandButton cmdAceptarPersona 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -74640
         TabIndex        =   19
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.CommandButton cmdEliminarPersona 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -72960
         TabIndex        =   18
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   2880
         Width           =   1605
      End
      Begin VB.Frame frameExoneraciones 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   3
         Top             =   360
         Width           =   11175
         Begin VB.TextBox txtExoneraDet 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1920
            TabIndex        =   12
            Top             =   1200
            Visible         =   0   'False
            Width           =   6735
         End
         Begin VB.ComboBox cboExoneraDet 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1200
            Visible         =   0   'False
            Width           =   6735
         End
         Begin VB.ComboBox cboUsuExo3 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmCredMicroseguro.frx":0038
            Left            =   6360
            List            =   "frmCredMicroseguro.frx":003A
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   480
            Width           =   2160
         End
         Begin VB.ComboBox cboUsuExo2 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmCredMicroseguro.frx":003C
            Left            =   4200
            List            =   "frmCredMicroseguro.frx":003E
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   480
            Width           =   2160
         End
         Begin VB.CommandButton cmdBorraExo 
            Caption         =   "Borrar"
            Height          =   375
            Left            =   9720
            TabIndex        =   8
            Top             =   1080
            Width           =   1020
         End
         Begin VB.CommandButton cmdAgregaExo 
            Caption         =   "Agrega"
            Height          =   375
            Left            =   9720
            TabIndex        =   7
            Top             =   600
            Width           =   1020
         End
         Begin VB.ComboBox cboUsuarioExone 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmCredMicroseguro.frx":0040
            Left            =   1920
            List            =   "frmCredMicroseguro.frx":0042
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   480
            Width           =   2160
         End
         Begin VB.ComboBox cboQuienExone 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   120
            Width           =   5295
         End
         Begin VB.ComboBox cboExoneraciones 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   840
            Width           =   6735
         End
         Begin SICMACT.FlexEdit FlexExo 
            Height          =   1125
            Left            =   960
            TabIndex        =   13
            Top             =   1560
            Width           =   8250
            _extentx        =   14552
            _extenty        =   1984
            cols0           =   11
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "#-Exoneracion-QuienExonera-Apoder1-Apoder2-Apoder3-cCtaCod-nCodExonera-cCodQuienExo-nDesExoneraCAR-cDesExoneraOtros"
            encabezadosanchos=   "400-3900-1500-700-700-700-0-0-0-0-0"
            font            =   "frmCredMicroseguro.frx":0044
            fontfixed       =   "frmCredMicroseguro.frx":0070
            columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X"
            listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0"
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            encabezadosalineacion=   "L-L-L-C-C-C-C-C-C-C-C"
            formatosedit    =   "0-0-0-0-0-0-0-0-0-0-0"
            textarray0      =   "#"
            lbeditarflex    =   -1  'True
            lbultimainstancia=   -1  'True
            colwidth0       =   405
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Apoderado :"
            Height          =   390
            Left            =   840
            TabIndex        =   16
            ToolTipText     =   "(días de atraso promedio últimas 6 cuotas)"
            Top             =   600
            Width           =   945
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Quien Exonera :"
            Height          =   195
            Left            =   600
            TabIndex        =   15
            ToolTipText     =   "(días de atraso promedio últimas 6 cuotas)"
            Top             =   240
            Width           =   1215
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Exoneración :"
            Height          =   195
            Left            =   840
            TabIndex        =   14
            ToolTipText     =   "(días de atraso promedio últimas 6 cuotas)"
            Top             =   960
            Width           =   975
            WordWrap        =   -1  'True
         End
      End
      Begin SICMACT.FlexEdit feMuebles 
         Height          =   1740
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   8250
         _extentx        =   14552
         _extenty        =   3069
         cols0           =   4
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-Descripcion-Año-Valor"
         encabezadosanchos=   "400-5000-1400-1200"
         font            =   "frmCredMicroseguro.frx":009E
         fontfixed       =   "frmCredMicroseguro.frx":00CA
         columnasaeditar =   "X-1-2-3"
         listacontroles  =   "0-0-0-0"
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         encabezadosalineacion=   "C-L-L-R"
         formatosedit    =   "0-0-0-4"
         cantentero      =   10
         textarray0      =   "#"
         lbultimainstancia=   -1  'True
         lbbuscaduplicadotext=   -1  'True
         colwidth0       =   405
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit FERelPers 
         Height          =   1605
         Left            =   -74760
         TabIndex        =   36
         Top             =   1200
         Width           =   9210
         _extentx        =   16245
         _extenty        =   2831
         cols0           =   7
         highlight       =   1
         allowuserresizing=   3
         visiblepopmenu  =   -1  'True
         encabezadosnombres=   "#-Codigo-Doc Id-Nombre-Direccion-Telefono-Fecha Nacimiento"
         encabezadosanchos=   "250-1300-1300-2500-3000-1000-1600"
         font            =   "frmCredMicroseguro.frx":00F8
         fontfixed       =   "frmCredMicroseguro.frx":0120
         columnasaeditar =   "X-1-X-X-X-X-X"
         textstylefixed  =   4
         listacontroles  =   "0-1-0-0-0-0-0"
         encabezadosalineacion=   "C-L-L-L-L-L-C"
         formatosedit    =   "0-0-0-0-0-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1  'True
         lbflexduplicados=   0   'False
         lbultimainstancia=   -1  'True
         tipobusqueda    =   3
         colwidth0       =   255
         rowheight0      =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Muebles:"
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
         Left            =   240
         TabIndex        =   44
         Top             =   480
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Beneficiarios:"
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
         Left            =   -74760
         TabIndex        =   43
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nº Certificado:"
         Height          =   195
         Left            =   -71640
         TabIndex        =   42
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nº Certificado:"
         Height          =   195
         Left            =   2520
         TabIndex        =   41
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label lblTotalValor 
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00;(0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   33
         Top             =   2640
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total: S/."
         Height          =   195
         Left            =   6480
         TabIndex        =   24
         Top             =   2760
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmCredMicroseguro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmCredMicroseguro
'***     Descripcion:      Registro de Microseguro
'***     Creado por:       WIOR
'***     Maquina:          TIF-1-19
'***     Fecha-Tiempo:     16/05/2012 01:00:00 PM
'***     Ultima Modificacion: Creacion de la Opcion
'*****************************************************************************************
Option Explicit
Dim fnHayMicroseguro As Integer
Dim fnHayMultiriesgo As Integer
Dim fnEditarMueble As Integer
Dim fnEditarPersona As Integer
Dim fnMontoMueble As Double

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim oCredito As COMDCredito.DCOMCredito
        Dim rsCredito As ADODB.Recordset
        Dim nEstadoCred As Integer
        Set oCredito = New COMDCredito.DCOMCredito
        Set rsCredito = oCredito.RecuperaColocacCred(ActXCodCta.NroCuenta)
        If Not (rsCredito.EOF And rsCredito.BOF) Then
            nEstadoCred = oCredito.RecuperaEstadoCredito(ActXCodCta.NroCuenta)
            If nEstadoCred = 2001 Then
                Call CargaDatos(ActXCodCta.NroCuenta)
            Else
                MsgBox "Credito No se encuentra sugerido.", vbInformation, "Aviso"
                Call CargaDatos(ActXCodCta.NroCuenta)
                sstRegistro.Enabled = True
                cmdAgregarMueble.Visible = False
                cmdEditarMueble.Visible = False
                cmdEliminarMueble.Visible = False
                cmdAceptarMueble.Visible = False
                cmdCancelarMueble.Visible = False
                CmdAceptar.Enabled = False
                cmdcancelar.Enabled = False
                
                CmdAceptar.Enabled = False
                cmdcancelar.Enabled = False
                cmdEditar.Enabled = False
                
                Me.cmdAceptarMueble.Visible = False
                Me.cmdCancelarMueble.Visible = False
                Me.cmdAgregarMueble.Visible = False
                Me.cmdEliminarMueble.Visible = False
                Me.cmdEditarMueble.Visible = False
                
                Me.cmdAceptarPersona.Visible = False
                Me.cmdCancelarPersona.Visible = False
                Me.cmdAgregarPersona.Visible = False
                Me.cmdEliminarPersona.Visible = False
                Me.cmdEditarPersona.Visible = False
                txtNumCertificadoMult.Enabled = False
                txtNumCertificadoMicro.Enabled = False
                fnEditarMueble = 0
                fnEditarPersona = 0
            End If
        Else
            MsgBox "Credito No existe.", vbInformation, "Aviso"
        End If
        
        Set oCredito = Nothing
        Set rsCredito = Nothing
    End If
End Sub
Private Sub cmdAceptar_Click()
    Dim oCreditoBD As COMDCredito.DCOMCredActBD
    Dim oConecta As COMConecta.DCOMConecta
    Dim i As Integer
    Dim J As Integer
    
    If MsgBox("Desea Grabar los datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set oCreditoBD = New COMDCredito.DCOMCredActBD
        Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        oConecta.BeginTrans

        If fnHayMicroseguro = 1 Then
            If Trim(Me.txtNumCertificadoMicro.Text) = "" Then
                MsgBox "Ingrese el Nº de Certificado de Microseguros.", vbInformation, "Aviso"
                Exit Sub
            Else
                Call oCreditoBD.ActualizarNumCertificado(ActXCodCta.NroCuenta, Trim(Me.txtNumCertificadoMicro.Text), "Microseguro")
            End If
            oCreditoBD.QuitarBeneficiariosMicroseguro (ActXCodCta.NroCuenta)
            If Me.FERelPers.TextMatrix(1, 1) <> "" Then
                For i = 1 To FERelPers.Rows - 1
                   Call oCreditoBD.RegistrarBeneficiariosMicroseguro(ActXCodCta.NroCuenta, Trim(FERelPers.TextMatrix(i, 1)), gdFecSis, 1)
                Next i
            End If
        End If
        If fnHayMultiriesgo = 1 Then
            If Trim(txtNumCertificadoMult.Text) = "" Then
                MsgBox "Ingrese el Nº de Certificado del Seguro Multiriesgo.", vbInformation, "Aviso"
                Exit Sub
            Else
                Call oCreditoBD.ActualizarNumCertificado(ActXCodCta.NroCuenta, Trim(Me.txtNumCertificadoMult.Text), "Multiriesgo")
            End If
            oCreditoBD.QuitarMueblesMultiriesgo (ActXCodCta.NroCuenta)
            Call oCreditoBD.ActualizarValorMultiriesgo(ActXCodCta.NroCuenta, CDbl(Me.lblTotalValor.Caption))
            If Me.feMuebles.TextMatrix(1, 1) <> "" Then
                For J = 1 To feMuebles.Rows - 1
                    Call oCreditoBD.RegistrarMueblesMultiriesgo(ActXCodCta.NroCuenta, Trim(feMuebles.TextMatrix(J, 1)), Int(feMuebles.TextMatrix(J, 2)), CDbl(feMuebles.TextMatrix(J, 3)), gdFecSis, 1)
                Next J
            End If
        End If
        oConecta.CommitTrans
        oConecta.CierraConexion
        Set oCreditoBD = Nothing
        Set oConecta = Nothing
        MsgBox "Datos registrados correctamente.", vbInformation, "Aviso"
        Call cmdCancelar_Click
    End If
End Sub

Private Sub cmdAceptarMueble_Click()
    Dim i As Long
    Dim J As Long
    For i = 1 To feMuebles.Rows - 1
        For J = 1 To feMuebles.Rows - 1
            If i <> J Then
                If Trim(feMuebles.TextMatrix(i, 1)) = Trim(feMuebles.TextMatrix(J, 1)) And Trim(feMuebles.TextMatrix(i, 2)) = Trim(feMuebles.TextMatrix(J, 2)) And Trim(feMuebles.TextMatrix(i, 3)) = Trim(feMuebles.TextMatrix(J, 3)) Then
                    MsgBox "Mueble ya existe.", vbInformation, "Aviso"
                    feMuebles.Row = J
                    feMuebles.Col = 1
                    feMuebles.SetFocus
                    Exit Sub
                End If
            End If
        Next J
    Next i
    For i = 1 To feMuebles.Rows - 1
        If Trim(feMuebles.TextMatrix(i, 1)) = "" Or Trim(feMuebles.TextMatrix(i, 2)) = "" Or Trim(feMuebles.TextMatrix(i, 3)) = "" Then
            MsgBox "Debe llenar correctamente todos los campos.", vbInformation, "Aviso"
            feMuebles.Row = i
            feMuebles.Col = 1
            feMuebles.SetFocus
            Exit Sub
        End If
    Next i
    feMuebles.lbEditarFlex = False
    cmdAgregarMueble.Visible = True
    cmdEditarMueble.Visible = True
    cmdEliminarMueble.Visible = True
    cmdAceptarMueble.Visible = False
    cmdCancelarMueble.Visible = False
    CmdAceptar.Enabled = True
    cmdcancelar.Enabled = True
    fnEditarMueble = 0
End Sub

Private Sub cmdAceptarPersona_Click()
    Dim i As Long
    Dim J As Long
    For i = 1 To FERelPers.Rows - 1
        For J = 1 To FERelPers.Rows - 1
            If i <> J Then
                If Trim(FERelPers.TextMatrix(i, 1)) = Trim(FERelPers.TextMatrix(J, 1)) Then
                    MsgBox "Persona ya es beneficiaria.", vbInformation, "Aviso"
                    FERelPers.Row = J
                    FERelPers.Col = 1
                    FERelPers.SetFocus
                    Exit Sub
                End If
            End If
        Next J
    Next i
    For i = 1 To FERelPers.Rows - 1
        If Len(Trim(FERelPers.TextMatrix(i, 1))) < 13 Then
            MsgBox "Codigo de Persona Incorrecto", vbInformation, "Aviso"
            FERelPers.Row = i
            FERelPers.Col = 1
            FERelPers.SetFocus
            Exit Sub
        End If
    Next i
    FERelPers.lbEditarFlex = False
    cmdAgregarPersona.Visible = True
    cmdEditarPersona.Visible = True
    cmdEliminarPersona.Visible = True
    cmdAceptarPersona.Visible = False
    cmdCancelarPersona.Visible = False
    CmdAceptar.Enabled = True
    cmdcancelar.Enabled = True
    fnEditarPersona = 0
End Sub

Private Sub cmdAgregarMueble_Click()
    feMuebles.lbEditarFlex = True
    feMuebles.AdicionaFila
    cmdAgregarMueble.Visible = False
    cmdEditarMueble.Visible = False
    cmdEliminarMueble.Visible = False
    cmdAceptarMueble.Visible = True
    Me.cmdCancelarMueble.Visible = True
    CmdAceptar.Enabled = False
    cmdcancelar.Enabled = False
    feMuebles.SetFocus
    fnEditarMueble = 0
End Sub

Private Sub cmdAgregarPersona_Click()
    FERelPers.lbEditarFlex = True
    FERelPers.AdicionaFila
    cmdAgregarPersona.Visible = False
    cmdEditarPersona.Visible = False
    cmdEliminarPersona.Visible = False
    cmdAceptarPersona.Visible = True
    Me.cmdCancelarPersona.Visible = True
    CmdAceptar.Enabled = False
    cmdcancelar.Enabled = False
    FERelPers.SetFocus
    fnEditarPersona = 0
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiaFlex(feMuebles)
    Call LimpiaFlex(FERelPers)
    Me.lblPrima.Caption = "S/. 0.00"
    Me.lblTotalValor.Caption = ""
    Me.sstRegistro.TabVisible(0) = True
    Me.sstRegistro.TabVisible(1) = True
    Me.sstRegistro.Enabled = False
    CmdAceptar.Enabled = False
    cmdcancelar.Enabled = False
    cmdEditar.Enabled = False
    Me.ActXCodCta.Enabled = True
    Call CargaDatos(Me.ActXCodCta.NroCuenta)
    
    Me.cmdAceptarMueble.Visible = False
    Me.cmdCancelarMueble.Visible = False
    Me.cmdAgregarMueble.Visible = True
    Me.cmdEliminarMueble.Visible = True
    Me.cmdEditarMueble.Visible = True
    
    Me.cmdAceptarPersona.Visible = False
    Me.cmdCancelarPersona.Visible = False
    Me.cmdAgregarPersona.Visible = True
    Me.cmdEliminarPersona.Visible = True
    Me.cmdEditarPersona.Visible = True
    
    fnEditarMueble = 0
    fnEditarPersona = 0
End Sub

Private Sub cmdCancelarMueble_Click()
    If fnEditarMueble = 0 Then
        Call feMuebles.EliminaFila(feMuebles.Row)
    End If
    feMuebles.lbEditarFlex = False
    cmdAgregarMueble.Visible = True
    cmdEditarMueble.Visible = True
    cmdEliminarMueble.Visible = True
    cmdAceptarMueble.Visible = False
    cmdCancelarMueble.Visible = False
    CmdAceptar.Enabled = True
    cmdcancelar.Enabled = True
End Sub

Private Sub cmdCancelarPersona_Click()
    If fnEditarPersona = 0 Then
        Call FERelPers.EliminaFila(FERelPers.Row)
    End If
    FERelPers.lbEditarFlex = False
    cmdAgregarPersona.Visible = True
    cmdEditarPersona.Visible = True
    cmdEliminarPersona.Visible = True
    cmdAceptarPersona.Visible = False
    cmdCancelarPersona.Visible = False
    CmdAceptar.Enabled = True
    cmdcancelar.Enabled = True
End Sub

Private Sub cmdEditar_Click()
    Me.sstRegistro.Enabled = True
    Me.CmdAceptar.Enabled = True
    Me.cmdEditar.Enabled = False
    Me.cmdcancelar.Enabled = True
End Sub

Private Sub cmdEditarMueble_Click()
    fnEditarMueble = 1
    feMuebles.lbEditarFlex = True
    cmdAgregarMueble.Visible = False
    cmdEditarMueble.Visible = False
    cmdEliminarMueble.Visible = False
    cmdAceptarMueble.Visible = True
    Me.cmdCancelarMueble.Visible = True
    CmdAceptar.Enabled = False
    cmdcancelar.Enabled = False
    feMuebles.SetFocus
End Sub

Private Sub cmdEditarPersona_Click()
    fnEditarPersona = 1
    FERelPers.lbEditarFlex = True
    cmdAgregarPersona.Visible = False
    cmdEditarPersona.Visible = False
    cmdEliminarPersona.Visible = False
    cmdAceptarPersona.Visible = True
    Me.cmdCancelarPersona.Visible = True
    CmdAceptar.Enabled = False
    cmdcancelar.Enabled = False
    FERelPers.SetFocus
End Sub

Private Sub cmdEliminarMueble_Click()
    If feMuebles.Row < 1 Then
        Exit Sub
    End If
    If MsgBox("Se va a Eliminar el Mueble ''" & feMuebles.TextMatrix(feMuebles.Row, 1) & "'', Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        If feMuebles.Row = 1 And feMuebles.Rows = 2 Then
            feMuebles.TextMatrix(1, 0) = ""
            feMuebles.TextMatrix(1, 1) = ""
            feMuebles.TextMatrix(1, 2) = ""
            feMuebles.TextMatrix(1, 3) = ""
        Else
            Call feMuebles.EliminaFila(feMuebles.Row)
        End If
    End If
    Call SumarPrima(feMuebles)
End Sub

Private Sub cmdEliminarPersona_Click()
    If FERelPers.Row < 1 Then
        Exit Sub
    End If
    If MsgBox("Se va a Eliminar a la Persona " & FERelPers.TextMatrix(FERelPers.Row, 3) & ", Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        If FERelPers.Row = 1 And FERelPers.Rows = 2 Then
            FERelPers.TextMatrix(1, 0) = ""
            FERelPers.TextMatrix(1, 1) = ""
            FERelPers.TextMatrix(1, 2) = ""
            FERelPers.TextMatrix(1, 3) = ""
            FERelPers.TextMatrix(1, 4) = ""
            FERelPers.TextMatrix(1, 5) = ""
            FERelPers.TextMatrix(1, 6) = ""
        Else
            Call FERelPers.EliminaFila(FERelPers.Row)
        End If
    End If
End Sub

Private Sub cmdExaminar_Click()
    Call CmdLimpiar_Click
    ActXCodCta.NroCuenta = frmCredPersEstado.Inicio(Array(gColocEstSolic, gColocEstSug), "Creditos Sugeridos (Microseguros/Multiriesgo)", , , , gsCodAge, , True)
   If ActXCodCta.NroCuenta <> "" Then
         Call CargaDatos(ActXCodCta.NroCuenta)
    Else
        ActXCodCta.CMAC = gsCodCMAC
        ActXCodCta.Age = gsCodAge
        ActXCodCta.SetFocusProd
        ActXCodCta.Enabled = True
    End If
End Sub

Private Sub CmdLimpiar_Click()
    Call LimpiaFlex(feMuebles)
    Call LimpiaFlex(FERelPers)
    Me.lblPrima.Caption = "S/. 0.00"
    Me.lblTotalValor.Caption = ""
    Me.sstRegistro.TabVisible(0) = True
    Me.sstRegistro.TabVisible(1) = True
    Me.sstRegistro.Enabled = False
    CmdAceptar.Enabled = False
    cmdcancelar.Enabled = False
    cmdEditar.Enabled = False
    ActXCodCta.NroCuenta = ""
    ActXCodCta.CMAC = gsCodCMAC
    ActXCodCta.Age = gsCodAge
    Me.ActXCodCta.Enabled = True
    txtNumCertificadoMult.Text = ""
    txtNumCertificadoMicro.Text = ""
    
    Me.cmdAceptarMueble.Visible = False
    Me.cmdCancelarMueble.Visible = False
    Me.cmdAgregarMueble.Visible = True
    Me.cmdEliminarMueble.Visible = True
    Me.cmdEditarMueble.Visible = True
    
    Me.cmdAceptarPersona.Visible = False
    Me.cmdCancelarPersona.Visible = False
    Me.cmdAgregarPersona.Visible = True
    Me.cmdEliminarPersona.Visible = True
    Me.cmdEditarPersona.Visible = True
    
    fnEditarMueble = 0
    fnEditarPersona = 0
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub feMuebles_OnCellChange(pnRow As Long, pnCol As Long)
'Dim nMontoMueble As Double
'nMontoMueble = CDbl(IIf(feMuebles.TextMatrix(feMuebles.Row, 3) = "", 0, feMuebles.TextMatrix(feMuebles.Row, 3)))
    If SumarPrima(feMuebles) Then
        feMuebles.TextMatrix(feMuebles.Row, 1) = UCase(feMuebles.TextMatrix(feMuebles.Row, 1))
    Else
        If fnMontoMueble > 0 Then
            feMuebles.TextMatrix(feMuebles.Row, 3) = fnMontoMueble
        Else
            feMuebles.TextMatrix(feMuebles.Row, 3) = 0
        End If
        If SumarPrima(feMuebles) Then
        End If
    End If
End Sub

Private Function SumarPrima(ByVal pFe As FlexEdit) As Boolean
    Dim nTotal As Double
    Dim nConteo As Integer
    nTotal = 0
    If pFe.Rows - 1 > 0 Then
        For nConteo = 1 To pFe.Rows - 1
            If pFe.TextMatrix(nConteo, 3) <> "0.00" Then
                nTotal = nTotal + CDbl(IIf(pFe.TextMatrix(nConteo, 3) = "", 0, pFe.TextMatrix(nConteo, 3)))
            End If
        Next
    End If
    Me.lblTotalValor.Caption = nTotal
    If nTotal > 30000 Then
        MsgBox "Monto Total no puede sobre pasar los S/.30000.00.", vbInformation, "Aviso"
        SumarPrima = False
        Exit Function
    Else
        fnMontoMueble = CDbl(IIf(feMuebles.TextMatrix(feMuebles.Row, 3) = "", 0, feMuebles.TextMatrix(feMuebles.Row, 3)))
    End If
    SumarPrima = True
End Function

Private Sub FERelPers_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim nPersoneria As PersPersoneria
    If pbEsDuplicado Then
        MsgBox "Persona ya esta registrada en la relación.", vbInformation, "Aviso"
        FERelPers.EliminaFila FERelPers.Row
    ElseIf psDataCod = "" Then
        FERelPers.TextMatrix(FERelPers.Row, 2) = ""
        FERelPers.TextMatrix(FERelPers.Row, 3) = ""
        FERelPers.TextMatrix(FERelPers.Row, 4) = ""
        FERelPers.TextMatrix(FERelPers.Row, 5) = ""
        FERelPers.TextMatrix(FERelPers.Row, 6) = ""
        Exit Sub
    Else
        Dim ClsPersona As COMDpersona.DCOMPersonas
        Dim R As New ADODB.Recordset
          
        Set ClsPersona = New COMDpersona.DCOMPersonas
        Set R = ClsPersona.BuscaCliente(FERelPers.TextMatrix(FERelPers.Row, 1), 2)
         
        If Not (R.EOF And R.BOF) Then
            FERelPers.TextMatrix(FERelPers.Row, 2) = IIf(IsNull(R!cPersIDnroDNI), "", IIf(R!cPersIDnroDNI = "", R!cPersIDnroRUC, R!cPersIDnroDNI))
            FERelPers.TextMatrix(FERelPers.Row, 3) = PstaNombre(R!cPersNombre)
            FERelPers.TextMatrix(FERelPers.Row, 4) = R!cPersDireccDomicilio
            FERelPers.TextMatrix(FERelPers.Row, 5) = IIf(IsNull(R!cPersTelefono) Or R!cPersTelefono = "" Or R!cPersTelefono = "0", "", R!cPersTelefono)
            FERelPers.TextMatrix(FERelPers.Row, 6) = Format(R!dPersNacCreac, "dd/mm/yyyy")
        End If
        Set ClsPersona = Nothing
    End If
End Sub
Private Sub Form_Load()
    Call CentraForm(Me)
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    ActXCodCta.CMAC = gsCodCMAC
    ActXCodCta.Age = gsCodAge
    Me.sstRegistro.Enabled = False
    Me.CmdAceptar.Visible = True
    Me.CmdAceptar.Enabled = False
    Me.cmdEditar.Enabled = False
    Me.cmdcancelar.Enabled = False
End Sub
Private Sub CargaDatos(ByVal psCtaCod As String)
    Dim oCreditoBD As COMDCredito.DCOMCredActBD
    Dim oCredito As COMDCredito.DCOMCredito
    Dim ClsPersona As COMDpersona.DCOMPersonas
    Dim rsBeneficiarios As ADODB.Recordset
    Dim rsPersona As ADODB.Recordset
    Dim rsMuebles As ADODB.Recordset
    Dim rsMicroseguro As ADODB.Recordset
    Dim rsMultiriesgo As ADODB.Recordset
    Dim a As Integer
    Dim b As Integer
    
    Set oCreditoBD = New COMDCredito.DCOMCredActBD
    Set oCredito = New COMDCredito.DCOMCredito
    Set ClsPersona = New COMDpersona.DCOMPersonas
    fnHayMicroseguro = 0
    fnHayMultiriesgo = 0
    fnEditarMueble = 0
    fnEditarPersona = 0

    'Mostrar Microseguro
    Set rsMicroseguro = oCredito.ObtenerMicroseguro(psCtaCod)
    If Not (rsMicroseguro.EOF And rsMicroseguro.BOF) Then
        fnHayMicroseguro = 1
        Me.txtNumCertificadoMicro.Text = rsMicroseguro!cNumCert
        If rsMicroseguro!nTipo = 1 Then
            Me.lblPrima.Caption = "S/. 2.50"
        ElseIf rsMicroseguro!nTipo = 2 Then
            Me.lblPrima.Caption = "S/. 1.50"
        End If
        
        Set rsBeneficiarios = oCreditoBD.ObtenerBeneficiariosMicroseguro(psCtaCod)
        If Not (rsBeneficiarios.EOF And rsBeneficiarios.BOF) Then
            For a = 0 To rsBeneficiarios.RecordCount - 1
                Set rsPersona = ClsPersona.BuscaCliente(rsBeneficiarios!cPersCod, 2)
                If Not (rsPersona.EOF And rsPersona.BOF) Then
                    FERelPers.AdicionaFila
                    FERelPers.TextMatrix(a + 1, 0) = a + 1
                    FERelPers.TextMatrix(a + 1, 1) = rsPersona!cPersCod
                    FERelPers.TextMatrix(a + 1, 2) = IIf(IsNull(rsPersona!cPersIDnroDNI), "", IIf(rsPersona!cPersIDnroDNI = "", rsPersona!cPersIDnroRUC, rsPersona!cPersIDnroDNI))
                    FERelPers.TextMatrix(a + 1, 3) = PstaNombre(rsPersona!cPersNombre)
                    FERelPers.TextMatrix(a + 1, 4) = rsPersona!cPersDireccDomicilio
                    FERelPers.TextMatrix(a + 1, 5) = IIf(IsNull(rsPersona!cPersTelefono) Or rsPersona!cPersTelefono = "" Or rsPersona!cPersTelefono = "0", "", rsPersona!cPersTelefono)
                    FERelPers.TextMatrix(a + 1, 6) = Format(rsPersona!dPersNacCreac, "dd/mm/yyyy")
                End If
                rsBeneficiarios.MoveNext
            Next a
        End If
        Me.cmdEditar.Enabled = True
        Me.ActXCodCta.Enabled = False
    End If

    'Mostrar Multiriesgo
    Set rsMultiriesgo = oCredito.ObtenerMultiriesgo(psCtaCod)
    If Not (rsMultiriesgo.EOF And rsMultiriesgo.BOF) Then
        fnHayMultiriesgo = 1
        Me.txtNumCertificadoMult.Text = rsMultiriesgo!cNumCert
        Me.lblTotalValor.Caption = rsMultiriesgo!nValorTotal
        Set rsMuebles = oCreditoBD.ObtenerMueblesMultiriesgo(psCtaCod)
        If Not (rsMuebles.EOF And rsMuebles.BOF) Then
            For b = 0 To rsMuebles.RecordCount - 1
                    feMuebles.AdicionaFila
                    feMuebles.TextMatrix(b + 1, 0) = b + 1
                    feMuebles.TextMatrix(b + 1, 1) = rsMuebles!cdescripcion
                    feMuebles.TextMatrix(b + 1, 2) = rsMuebles!nAno
                    feMuebles.TextMatrix(b + 1, 3) = rsMuebles!nValor
                rsMuebles.MoveNext
            Next b
        End If
        Me.cmdEditar.Enabled = True
        Me.ActXCodCta.Enabled = False
    End If

    'Mostrar correctamente los tabs
    If fnHayMicroseguro = 0 And fnHayMultiriesgo = 1 Then
        Me.sstRegistro.TabVisible(0) = True
        Me.sstRegistro.TabVisible(1) = True
        Me.sstRegistro.TabVisible(0) = False
        Me.sstRegistro.TabVisible(1) = False
        Me.sstRegistro.TabVisible(1) = True
    ElseIf fnHayMicroseguro = 1 And fnHayMultiriesgo = 0 Then
        Me.sstRegistro.TabVisible(0) = True
        Me.sstRegistro.TabVisible(1) = True
        Me.sstRegistro.TabVisible(1) = False
        Me.sstRegistro.TabVisible(0) = False
        Me.sstRegistro.TabVisible(0) = True
    ElseIf fnHayMicroseguro = 0 And fnHayMultiriesgo = 0 Then
        Call CmdLimpiar_Click
        MsgBox "No cuenta con ningun registro de los 2 tipos de Seguro en la sugerencia.", vbInformation, "Aviso"
    ElseIf fnHayMicroseguro = 1 And fnHayMultiriesgo = 1 Then
        Me.sstRegistro.TabVisible(1) = True
        Me.sstRegistro.TabVisible(0) = True
    End If
End Sub

