VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredFormEvalFormato2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cr�ditos - Evaluaci�n - Formato 2"
   ClientHeight    =   10260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10350
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000000&
   Icon            =   "frmCredFormEvalFormato2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10260
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMNME 
      Caption         =   "MN - ME"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4600
      TabIndex        =   106
      Top             =   9915
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdVerCar 
      Caption         =   "&Ver CAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1770
      TabIndex        =   37
      Top             =   9915
      Width           =   1170
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Hoja Evaluaci�n"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2940
      TabIndex        =   38
      Top             =   9915
      Width           =   1650
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9120
      TabIndex        =   35
      Top             =   9915
      Width           =   1170
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7950
      TabIndex        =   34
      Top             =   9915
      Width           =   1170
   End
   Begin VB.CommandButton cmdInformeVisita 
      Caption         =   "Infor&me de Visita"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   36
      Top             =   9915
      Width           =   1650
   End
   Begin TabDlg.SSTab SSTabInfoNego 
      Height          =   2190
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   3863
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Informaci�n del Negocio"
      TabPicture(0)   =   "frmCredFormEvalFormato2.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtFechaEvaluacion"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ActXCodCta"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "frameLinea"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtGiroNeg"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.TextBox txtGiroNeg 
         Height          =   300
         Left            =   6000
         TabIndex        =   3
         Top             =   360
         Width           =   4155
      End
      Begin VB.Frame frameLinea 
         Height          =   255
         Left            =   6240
         TabIndex        =   91
         Top             =   2040
         Visible         =   0   'False
         Width           =   3855
         Begin VB.TextBox txtNumLinea 
            Height          =   300
            Left            =   1800
            TabIndex        =   93
            Top             =   120
            Width           =   1995
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Linea Autom�tica :"
            Height          =   195
            Left            =   120
            TabIndex        =   92
            Top             =   165
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1405
         Left            =   120
         TabIndex        =   9
         Top             =   700
         Width           =   10095
         Begin VB.TextBox txtCondLocalOtros 
            Height          =   285
            Left            =   6960
            MaxLength       =   250
            TabIndex        =   16
            Top             =   1100
            Visible         =   0   'False
            Width           =   3075
         End
         Begin VB.OptionButton OptCondLocal 
            Caption         =   "Otros"
            Height          =   255
            Index           =   4
            Left            =   6000
            TabIndex        =   15
            Top             =   1110
            Width           =   855
         End
         Begin VB.OptionButton OptCondLocal 
            Caption         =   "Ambulante"
            Height          =   255
            Index           =   3
            Left            =   4680
            TabIndex        =   14
            Top             =   1110
            Width           =   1095
         End
         Begin VB.OptionButton OptCondLocal 
            Caption         =   "Alquilada"
            Height          =   255
            Index           =   2
            Left            =   3480
            TabIndex        =   13
            Top             =   1110
            Width           =   1095
         End
         Begin VB.OptionButton OptCondLocal 
            Caption         =   "Propia"
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   12
            Top             =   1110
            Width           =   855
         End
         Begin VB.TextBox txtNombreCliente 
            Height          =   300
            Left            =   2400
            TabIndex        =   4
            Top             =   120
            Width           =   4035
         End
         Begin MSMask.MaskEdBox txtFecUltEndeuda 
            Height          =   300
            Left            =   9000
            TabIndex        =   8
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin Spinner.uSpinner spnTiempoLocalAnio 
            Height          =   315
            Left            =   2400
            TabIndex        =   0
            Top             =   780
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            Max             =   99
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
         End
         Begin Spinner.uSpinner spnTiempoLocalMes 
            Height          =   315
            Left            =   3720
            TabIndex        =   10
            Top             =   780
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            Max             =   12
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
         End
         Begin SICMACT.EditMoney txtExposicionCredito 
            Height          =   300
            Left            =   8640
            TabIndex        =   11
            Top             =   795
            Width           =   1335
            _extentx        =   2355
            _extenty        =   529
            font            =   "frmCredFormEvalFormato2.frx":0326
            backcolor       =   -2147483643
            forecolor       =   8421504
            text            =   "0"
            enabled         =   -1
         End
         Begin Spinner.uSpinner spnExpEmpAnio 
            Height          =   315
            Left            =   2400
            TabIndex        =   6
            Top             =   450
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            Max             =   99
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   8421504
         End
         Begin Spinner.uSpinner spnExpEmpMes 
            Height          =   315
            Left            =   3720
            TabIndex        =   7
            Top             =   450
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            Max             =   12
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   8421504
         End
         Begin SICMACT.EditMoney txtUltEndeuda 
            Height          =   300
            Left            =   8760
            TabIndex        =   5
            Top             =   150
            Width           =   1215
            _extentx        =   2143
            _extenty        =   529
            font            =   "frmCredFormEvalFormato2.frx":034E
            backcolor       =   -2147483643
            forecolor       =   8421504
            text            =   "0"
            enabled         =   -1
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha �ltimo endeudamiento RCC :"
            Height          =   195
            Left            =   6040
            TabIndex        =   51
            Top             =   480
            Width           =   2510
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ltimo endeudamiento RCC :"
            Height          =   195
            Left            =   6500
            TabIndex        =   50
            Top             =   160
            Width           =   2055
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "meses"
            Height          =   255
            Left            =   4515
            TabIndex        =   49
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "meses"
            Height          =   255
            Left            =   4515
            TabIndex        =   48
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "a�os"
            Height          =   255
            Left            =   3195
            TabIndex        =   47
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "a�os"
            Height          =   255
            Left            =   3195
            TabIndex        =   46
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exposici�n con este cr�dito :"
            Height          =   195
            Left            =   6495
            TabIndex        =   45
            Top             =   795
            Width           =   2055
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Condici�n local :"
            Height          =   210
            Left            =   1170
            TabIndex        =   44
            Top             =   1100
            Width           =   1215
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tiempo en el mismo local :"
            Height          =   255
            Left            =   480
            TabIndex        =   43
            Top             =   800
            Width           =   1935
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Experiencia como empresario :"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   460
            Width           =   2295
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente :"
            Height          =   195
            Left            =   1695
            TabIndex        =   41
            Top             =   120
            Width           =   600
         End
      End
      Begin SICMACT.ActXCodCta ActXCodCta 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3735
         _extentx        =   6588
         _extenty        =   661
         texto           =   "Cr�dito"
      End
      Begin MSMask.MaskEdBox txtFechaEvaluacion 
         Height          =   300
         Left            =   9000
         TabIndex        =   104
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   16777215
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de evaluaci�n al :"
         Height          =   195
         Left            =   7200
         TabIndex        =   105
         Top             =   50
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Giro del Negocio :"
         Height          =   255
         Left            =   4680
         TabIndex        =   52
         Top             =   390
         Width           =   1335
      End
   End
   Begin TabDlg.SSTab SSTabIngresos 
      Height          =   7020
      Left            =   0
      TabIndex        =   39
      Top             =   2190
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   12383
      _Version        =   393216
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Ingresos y Egresos"
      TabPicture(0)   =   "frmCredFormEvalFormato2.frx":0376
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Propuesta del Cr�dito"
      TabPicture(1)   =   "frmCredFormEvalFormato2.frx":0392
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "framePropuesta"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Comentarios y Referidos"
      TabPicture(2)   =   "frmCredFormEvalFormato2.frx":03AE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdQuitar2"
      Tab(2).Control(1)=   "frameReferido"
      Tab(2).Control(2)=   "frameComentario"
      Tab(2).ControlCount=   3
      Begin VB.CommandButton cmdQuitar2 
         Caption         =   "&Quitar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -66240
         TabIndex        =   33
         Top             =   5880
         Width           =   1170
      End
      Begin VB.Frame frameReferido 
         Caption         =   "Referidos :"
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
         Height          =   3135
         Left            =   -74760
         TabIndex        =   89
         Top             =   3240
         Width           =   9855
         Begin VB.CommandButton cmdAgregarRef 
            Caption         =   "&Agregar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   7200
            TabIndex        =   32
            Top             =   2640
            Width           =   1170
         End
         Begin SICMACT.FlexEdit feReferidos 
            Height          =   2295
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   9675
            _extentx        =   17066
            _extenty        =   4048
            cols0           =   7
            highlight       =   1
            encabezadosnombres=   "N-Nombres-DNI-Tel�fono-Comentario-NroDNI-Aux"
            encabezadosanchos=   "350-3200-960-1260-3750-0-0"
            font            =   "frmCredFormEvalFormato2.frx":03CA
            font            =   "frmCredFormEvalFormato2.frx":03F2
            font            =   "frmCredFormEvalFormato2.frx":041A
            font            =   "frmCredFormEvalFormato2.frx":0442
            font            =   "frmCredFormEvalFormato2.frx":046A
            fontfixed       =   "frmCredFormEvalFormato2.frx":0492
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   3
            columnasaeditar =   "X-1-2-3-4-X-X"
            listacontroles  =   "0-0-0-0-0-0-0"
            encabezadosalineacion=   "L-L-L-L-L-L-C"
            formatosedit    =   "0-0-0-0-0-0-0"
            textarray0      =   "N"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            colwidth0       =   345
            rowheight0      =   300
         End
      End
      Begin VB.Frame frameComentario 
         Caption         =   "Comentarios :"
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
         Height          =   2655
         Left            =   -74760
         TabIndex        =   88
         Top             =   360
         Width           =   9855
         Begin VB.TextBox txtComentario 
            Height          =   2250
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   3000
            MultiLine       =   -1  'True
            TabIndex        =   30
            Top             =   240
            Width           =   9615
         End
      End
      Begin VB.Frame framePropuesta 
         Caption         =   "Propuesta del Credito:"
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
         Height          =   6255
         Left            =   -74760
         TabIndex        =   81
         Top             =   360
         Width           =   9855
         Begin VB.TextBox txtEntornoFamiliar2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   24
            Top             =   720
            Width           =   9615
         End
         Begin VB.TextBox txtExperiencia2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   26
            Top             =   2640
            Width           =   9615
         End
         Begin VB.TextBox txtGiroUbicacion2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   1680
            Width           =   9615
         End
         Begin VB.TextBox txtFormalidadNegocio2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   3600
            Width           =   9615
         End
         Begin VB.TextBox txtColaterales2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   28
            Top             =   4560
            Width           =   9615
         End
         Begin VB.TextBox txtDestino2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   29
            Top             =   5520
            Width           =   9615
         End
         Begin MSMask.MaskEdBox txtFechaVisita 
            Height          =   300
            Left            =   8520
            TabIndex        =   23
            Top             =   240
            Width           =   1090
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Visita:"
            Height          =   195
            Left            =   7320
            TabIndex        =   90
            Top             =   300
            Width           =   1140
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el entorno familiar del cliente o representante:"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   87
            Top             =   480
            Width           =   3795
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el giro y la ubicaci�n del negocio:"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   86
            Top             =   1440
            Width           =   2820
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre la experiencia Crediticia:"
            Height          =   195
            Left            =   240
            TabIndex        =   85
            Top             =   2400
            Width           =   2220
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre la consistencia de la informaci�n y la formalidad del negocio:"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   84
            Top             =   3360
            Width           =   4770
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre los colaterales y garant�as:"
            Height          =   195
            Left            =   240
            TabIndex        =   83
            Top             =   4320
            Width           =   2400
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el destino y el impacto del mismo:"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   82
            Top             =   5280
            Width           =   2850
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Ventas y Costos "
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
         Height          =   540
         Left            =   360
         TabIndex        =   76
         Top             =   320
         Width           =   9255
         Begin VB.TextBox txtEgresoNegocio 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4800
            TabIndex        =   18
            Text            =   "0"
            Top             =   200
            Width           =   1515
         End
         Begin VB.TextBox txtIngresoNegocio 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            TabIndex        =   17
            Text            =   "0"
            Top             =   200
            Width           =   1515
         End
         Begin SICMACT.EditMoney txtMargenBruto 
            Height          =   300
            Left            =   7800
            TabIndex        =   77
            Top             =   200
            Width           =   1335
            _extentx        =   2355
            _extenty        =   529
            font            =   "frmCredFormEvalFormato2.frx":04B8
            backcolor       =   -2147483643
            forecolor       =   8421504
            text            =   "0"
            enabled         =   -1
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ingresos del Negocio :"
            Height          =   195
            Left            =   120
            TabIndex        =   80
            Top             =   240
            Width           =   1605
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Egreso por Venta:"
            Height          =   195
            Left            =   3480
            TabIndex        =   79
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Margen Bruto :"
            Height          =   195
            Left            =   6720
            TabIndex        =   78
            Top             =   240
            Width           =   1080
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Gastos del Negocio :"
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
         Height          =   4100
         Left            =   5500
         TabIndex        =   75
         Top             =   840
         Width           =   4695
         Begin SICMACT.FlexEdit feGastosNegocio 
            Height          =   3850
            Left            =   120
            TabIndex        =   20
            Top             =   195
            Width           =   4515
            _extentx        =   7964
            _extenty        =   6800
            cols0           =   5
            highlight       =   1
            encabezadosnombres=   "-N-Concepto-Monto-Aux"
            encabezadosanchos=   "0-300-2800-1300-0"
            font            =   "frmCredFormEvalFormato2.frx":04E0
            font            =   "frmCredFormEvalFormato2.frx":0508
            font            =   "frmCredFormEvalFormato2.frx":0530
            font            =   "frmCredFormEvalFormato2.frx":0558
            font            =   "frmCredFormEvalFormato2.frx":0580
            fontfixed       =   "frmCredFormEvalFormato2.frx":05A8
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   6
            columnasaeditar =   "X-X-X-3-X"
            listacontroles  =   "0-0-0-0-0"
            encabezadosalineacion=   "C-C-L-R-C"
            formatosedit    =   "0-0-0-2-0"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Gastos Familiares : "
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
         Height          =   2660
         Left            =   120
         TabIndex        =   74
         Top             =   4320
         Width           =   5175
         Begin SICMACT.FlexEdit feGastosFamiliares 
            Height          =   2415
            Left            =   120
            TabIndex        =   21
            Top             =   195
            Width           =   4905
            _extentx        =   8652
            _extenty        =   4260
            cols0           =   5
            highlight       =   1
            encabezadosnombres=   "-N-Concepto-Monto-Aux"
            encabezadosanchos=   "0-300-3200-1300-0"
            font            =   "frmCredFormEvalFormato2.frx":05CE
            font            =   "frmCredFormEvalFormato2.frx":05F6
            font            =   "frmCredFormEvalFormato2.frx":061E
            font            =   "frmCredFormEvalFormato2.frx":0646
            font            =   "frmCredFormEvalFormato2.frx":066E
            fontfixed       =   "frmCredFormEvalFormato2.frx":0696
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   6
            columnasaeditar =   "X-X-X-3-X"
            listacontroles  =   "0-0-0-0-0"
            encabezadosalineacion=   "C-C-L-R-C"
            formatosedit    =   "0-0-0-2-0"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Balance General :"
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
         Height          =   3450
         Left            =   100
         TabIndex        =   73
         Top             =   880
         Width           =   5175
         Begin SICMACT.FlexEdit feBalanceGeneral 
            Height          =   3210
            Left            =   120
            TabIndex        =   19
            Top             =   200
            Width           =   4905
            _extentx        =   8652
            _extenty        =   5662
            cols0           =   7
            highlight       =   1
            encabezadosnombres=   "-nConsCod-nConsValor-N-Descripcion-Monto-Aux"
            encabezadosanchos=   "0-0-0-0-3400-1400-0"
            font            =   "frmCredFormEvalFormato2.frx":06BC
            font            =   "frmCredFormEvalFormato2.frx":06E4
            font            =   "frmCredFormEvalFormato2.frx":070C
            font            =   "frmCredFormEvalFormato2.frx":0734
            font            =   "frmCredFormEvalFormato2.frx":075C
            fontfixed       =   "frmCredFormEvalFormato2.frx":0784
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   6
            columnasaeditar =   "X-X-X-X-X-5-X"
            listacontroles  =   "0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-C-L-C-L-R-C"
            formatosedit    =   "0-0-0-0-0-2-0"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Otros Ingresos :"
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
         Height          =   2060
         Left            =   5520
         TabIndex        =   72
         Top             =   4920
         Width           =   4695
         Begin SICMACT.FlexEdit feOtrosIngresos 
            Height          =   1815
            Left            =   120
            TabIndex        =   22
            Top             =   200
            Width           =   4515
            _extentx        =   7964
            _extenty        =   3201
            cols0           =   5
            highlight       =   1
            encabezadosnombres=   "-N-Concepto-Monto-Aux"
            encabezadosanchos=   "0-300-2800-1300-0"
            font            =   "frmCredFormEvalFormato2.frx":07AA
            font            =   "frmCredFormEvalFormato2.frx":07D2
            font            =   "frmCredFormEvalFormato2.frx":07FA
            font            =   "frmCredFormEvalFormato2.frx":0822
            font            =   "frmCredFormEvalFormato2.frx":084A
            fontfixed       =   "frmCredFormEvalFormato2.frx":0872
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   3
            columnasaeditar =   "X-X-X-3-X"
            listacontroles  =   "0-0-0-0-0"
            encabezadosalineacion=   "C-C-L-R-C"
            formatosedit    =   "0-0-0-2-0"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
         End
      End
      Begin VB.Frame frmCredEvalFormato1 
         Caption         =   " Gastos del Negocio :"
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
         Height          =   6015
         Left            =   -74880
         TabIndex        =   59
         Top             =   360
         Width           =   9975
         Begin VB.TextBox txtEntornoFamiliar 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   65
            Top             =   480
            Width           =   9735
         End
         Begin VB.TextBox txtGiroUbicacion 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   64
            Top             =   1440
            Width           =   9735
         End
         Begin VB.TextBox txtExperiencia 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   63
            Top             =   2400
            Width           =   9735
         End
         Begin VB.TextBox txtFormalidadNegocio 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   62
            Top             =   3360
            Width           =   9735
         End
         Begin VB.TextBox txtColaterales 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   61
            Top             =   4320
            Width           =   9735
         End
         Begin VB.TextBox txtDestino 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   60
            Top             =   5280
            Width           =   9735
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el entorno familiar del cliente o representante:"
            Height          =   195
            Left            =   120
            TabIndex        =   71
            Top             =   240
            Width           =   3795
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el giro y la ubicaci�n del negocio:"
            Height          =   195
            Left            =   120
            TabIndex        =   70
            Top             =   1200
            Width           =   2820
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "S�lo la experiencia crediticia:"
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   2160
            Width           =   2070
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre la consistencia de la informaci�n y la formalidad del negocio:"
            Height          =   195
            Left            =   120
            TabIndex        =   68
            Top             =   3120
            Width           =   4770
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre los colaterales y garant�as:"
            Height          =   195
            Left            =   120
            TabIndex        =   67
            Top             =   4080
            Width           =   2400
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre los colaterales y garant�as:"
            Height          =   195
            Left            =   120
            TabIndex        =   66
            Top             =   5040
            Width           =   2400
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Comentarios :"
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
         Height          =   2655
         Left            =   -74760
         TabIndex        =   57
         Top             =   360
         Width           =   9975
         Begin VB.TextBox Text1 
            Height          =   2010
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   58
            Top             =   240
            Width           =   9735
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Comentarios :"
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
         Height          =   2655
         Left            =   -74760
         TabIndex        =   55
         Top             =   3360
         Width           =   9975
         Begin SICMACT.FlexEdit FlexEdit1 
            Height          =   1935
            Left            =   120
            TabIndex        =   56
            Top             =   360
            Width           =   9720
            _extentx        =   17145
            _extenty        =   3413
            cols0           =   6
            highlight       =   1
            encabezadosnombres=   "N�-Nombre-DNI-Telef.-Referido-DNI"
            encabezadosanchos=   "1000-2800-1000-1500-2300-1000"
            font            =   "frmCredFormEvalFormato2.frx":0898
            font            =   "frmCredFormEvalFormato2.frx":08C0
            font            =   "frmCredFormEvalFormato2.frx":08E8
            font            =   "frmCredFormEvalFormato2.frx":0910
            font            =   "frmCredFormEvalFormato2.frx":0938
            fontfixed       =   "frmCredFormEvalFormato2.frx":0960
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   3
            columnasaeditar =   "X-X-X-X-X-X"
            listacontroles  =   "0-0-0-0-0-0"
            encabezadosalineacion=   "L-R-L-C-C-C"
            formatosedit    =   "0-2-0-0-0-0"
            textarray0      =   "N�"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            colwidth0       =   1005
            rowheight0      =   300
         End
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   -74640
         TabIndex        =   54
         Top             =   6120
         Width           =   1170
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73200
         TabIndex        =   53
         Top             =   6120
         Width           =   1170
      End
   End
   Begin TabDlg.SSTab SSTabRatios 
      Height          =   690
      Left            =   0
      TabIndex        =   40
      Top             =   9210
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   1217
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Ratios e Indicadores"
      TabPicture(0)   =   "frmCredFormEvalFormato2.frx":0986
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label22"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label21"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label19(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label13(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblCapaAceptable"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblEndeAceptable"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtExcedenteMensual"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtIngresoNeto"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtEndeudamiento"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtCapacidadNeta"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin SICMACT.EditMoney txtCapacidadNeta 
         Height          =   300
         Left            =   1460
         TabIndex        =   94
         Top             =   330
         Width           =   950
         _extentx        =   1667
         _extenty        =   529
         font            =   "frmCredFormEvalFormato2.frx":09A2
         forecolor       =   8421504
         text            =   "0"
         enabled         =   -1
      End
      Begin SICMACT.EditMoney txtEndeudamiento 
         Height          =   300
         Left            =   4360
         TabIndex        =   95
         Top             =   330
         Width           =   950
         _extentx        =   1667
         _extenty        =   529
         font            =   "frmCredFormEvalFormato2.frx":09CA
         forecolor       =   8421504
         text            =   "0"
         enabled         =   -1
      End
      Begin SICMACT.EditMoney txtIngresoNeto 
         Height          =   300
         Left            =   7160
         TabIndex        =   96
         Top             =   330
         Width           =   1155
         _extentx        =   1931
         _extenty        =   529
         font            =   "frmCredFormEvalFormato2.frx":09F2
         forecolor       =   8421504
         text            =   "0"
         enabled         =   -1
      End
      Begin SICMACT.EditMoney txtExcedenteMensual 
         Height          =   300
         Left            =   9140
         TabIndex        =   97
         Top             =   330
         Width           =   1155
         _extentx        =   1931
         _extenty        =   529
         font            =   "frmCredFormEvalFormato2.frx":0A1A
         forecolor       =   8421504
         text            =   "0"
         enabled         =   -1
      End
      Begin VB.Label lblEndeAceptable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aceptable"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   5340
         TabIndex        =   103
         Top             =   390
         Width           =   750
      End
      Begin VB.Label lblCapaAceptable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aceptable"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   2420
         TabIndex        =   102
         Top             =   390
         Width           =   750
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capacidad de Pago:"
         Height          =   195
         Index           =   0
         Left            =   50
         TabIndex        =   101
         Top             =   380
         Width           =   1440
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endeudamiento:"
         Height          =   195
         Index           =   0
         Left            =   3220
         TabIndex        =   100
         Top             =   380
         Width           =   1170
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso Neto:"
         Height          =   195
         Left            =   6180
         TabIndex        =   99
         Top             =   380
         Width           =   1005
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excedente:"
         Height          =   195
         Left            =   8340
         TabIndex        =   98
         Top             =   380
         Width           =   825
      End
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   10200
      Y1              =   9720
      Y2              =   9720
   End
End
Attribute VB_Name = "frmCredFormEvalFormato2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmCredFormEvalFormato2                                                     *
'** Descripci�n : Formulario para evaluaci�n de Creditos que tienen el tipo de evaluaci�n 2   *
'** Referencia  : ERS004-2016                                                                 *
'** Creaci�n    : LUCV, 20160525 09:00:00 AM                                                  *
'**********************************************************************************************
Option Explicit
    Dim sCtaCod As String
    Dim sPersCod As String
    Dim gsOpeCod As String
    Dim fnTipoRegMant As Integer
    Dim fnTipoPermiso As Integer
    Dim fbPermiteGrabar As Boolean
    Dim fbBloqueaTodo As Boolean
    Dim fnTotalRefGastoNego As Currency
    Dim fnTotalRefGastoFami As Currency
    Dim fsCliente As String
    Dim fsGiroNego As String
    Dim fsAnioExp As Integer
    Dim fsMesExp As Integer
    Dim fsUserAnalista  As String
    Dim fnEstado As Integer
    Dim fnMontoDeudaSbs As Currency
    Dim fnFechaDeudaSbs As Currency
    
    Dim lnCondLocal As Integer
    Dim MatIfiGastoNego As Variant
    Dim MatIfiGastoFami As Variant
    Dim MatReferidos As Variant
    
    Dim rsFeGastoNeg As ADODB.Recordset
    Dim rsFeDatGastoFam As ADODB.Recordset
    Dim rsFeDatOtrosIng As ADODB.Recordset
    Dim rsFeDatBalanGen As ADODB.Recordset
    Dim rsFeDatActivos As ADODB.Recordset
    Dim rsFeDatPasivos As ADODB.Recordset
    Dim rsFeDatPasivosNo As ADODB.Recordset
    Dim rsFeDatPatrimonio As ADODB.Recordset
    Dim rsFeDatRef As ADODB.Recordset
    
    Dim rsCredEval As ADODB.Recordset
    Dim rsDCredito As ADODB.Recordset
    Dim rsAceptableCritico As ADODB.Recordset
    Dim rsCapacPagoNeta As ADODB.Recordset
    Dim rsCuotaIFIs As ADODB.Recordset
    Dim rsPropuesta As ADODB.Recordset
    
    Dim rsDatPasivosNo As ADODB.Recordset
    Dim rsDatActivoPasivo As ADODB.Recordset
    Dim rsDatGastoNeg As ADODB.Recordset
    Dim rsDatGastoFam As ADODB.Recordset
    Dim rsDatOtrosIng As ADODB.Recordset
    Dim rsDatRef As ADODB.Recordset
    Dim rsDatRatioInd As ADODB.Recordset
    Dim rsDatIfiGastoNego As ADODB.Recordset
    Dim rsDatIfiGastoFami As ADODB.Recordset
    Dim rsDatVentaCosto As ADODB.Recordset
    Dim rsDatActivos As ADODB.Recordset
    Dim rsDatPasivos As ADODB.Recordset
    
    Dim nMontoAct As Currency
    Dim nMontoPas As Currency
    Dim nMontoPat As Currency
    Dim nMargenBruto As Currency
    
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim objPista As COMManejador.Pista
    Dim fnFormato As Integer
    Dim fnMontoIni As Double
    Dim lnMin As Double
    Dim lnMax As Double
    Dim lnMinDol As Double
    Dim lnMaxDol As Double
    Dim nTC As Double
    Dim I As Integer
    Dim j As Integer
    Dim K As Integer
    Dim fbGrabar As Boolean
    Dim fnColocCondi As Integer
    Dim fbTieneReferido6Meses As Boolean 'LUCV20171115, Agreg� segun correo: RUSI
    
    'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
    Dim MatBalActCorr As Variant
    Dim MatBalActNoCorr As Variant
    Dim rsDatIfiBalActCorri As ADODB.Recordset
    Dim rsDatIfiBalActNoCorri As ADODB.Recordset
    
    Dim lcMovNro As String 'LUCV20181220 Agreg�, Anexo01 de Acta 199-2018
    
'JOEP20180725 ERS034-2018
Private Sub cmdMNME_Click()
    Call frmCredFormEvalCredCel.Inicio(ActXCodCta.NroCuenta, 11)
End Sub
'JOEP20180725 ERS034-2018

Private Sub feBalanceGeneral_EnterCell()

If fnTipoRegMant = 2 Then
        If feBalanceGeneral.Col = 5 Then
                If CInt(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 0)) = 5 Then
                    feBalanceGeneral.ListaControles = "0-0-0-0-1"
                ElseIf CInt(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 0)) = 6 Then
                    feBalanceGeneral.ListaControles = "0-0-0-0-1"
                Else
                    feBalanceGeneral.ListaControles = "0-0-0-0-0"
                End If
        End If
Else
    If feBalanceGeneral.Col = 5 Then
        If CInt(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 0)) = 5 Then
            feBalanceGeneral.ListaControles = "0-0-0-0-1"
        ElseIf CInt(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 0)) = 6 Then
            feBalanceGeneral.ListaControles = "0-0-0-0-1"
        Else
            feBalanceGeneral.ListaControles = "0-0-0-0-0"
        End If
    End If
End If

End Sub

Private Sub feBalanceGeneral_OnClickTxtBuscar(psCodigo As String, psDescripcion As String)
    Dim fnTotalBalanceActCorriente As Currency
    Dim fnTotalBalanceActNoCorriente As Currency

    psCodigo = 0
    psDescripcion = ""
    psDescripcion = feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 4)
    psCodigo = feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 5)
    
If feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2) = 100 Then
    If psCodigo = 0 Then
         fnTotalBalanceActCorriente = 0
        Set MatBalActCorr = Nothing
        frmCredFormEvalCuotasIfis.Inicio (CLng(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 5))), fnTotalBalanceActCorriente, MatBalActCorr, feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 4)
        psCodigo = Format(fnTotalBalanceActCorriente, "#,##0.00")
    Else
        frmCredFormEvalCuotasIfis.Inicio (CLng(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 5))), fnTotalBalanceActCorriente, MatBalActCorr, feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 4)
        psCodigo = Format(fnTotalBalanceActCorriente, "#,##0.00")
    End If
ElseIf feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2) = 200 Then
    If psCodigo = 0 Then
         fnTotalBalanceActNoCorriente = 0
        Set MatBalActNoCorr = Nothing
        frmCredFormEvalCuotasIfis.Inicio (CLng(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 5))), fnTotalBalanceActNoCorriente, MatBalActNoCorr, feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 4)
        psCodigo = Format(fnTotalBalanceActNoCorriente, "#,##0.00")
    Else
        frmCredFormEvalCuotasIfis.Inicio (CLng(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 5))), fnTotalBalanceActNoCorriente, MatBalActNoCorr, feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 4)
        psCodigo = Format(fnTotalBalanceActNoCorriente, "#,##0.00")
    End If
End If

End Sub
'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
    
'_____________________________________________________________________________________________________________
'******************************************LUCV20160525: EVENTOS Varios***************************************
Private Sub Form_Load()
    fbGrabar = False
    CentraForm Me
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    EnfocaControl spnTiempoLocalAnio
    
'JOEP20180725 ERS034-2018
    If fnTipoRegMant = 3 Then
        If Not ConsultaRiesgoCamCred(sCtaCod) Then
            cmdMNME.Visible = True
        End If
    End If
'JOEP20180725 ERS034-2018
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set MatIfiGastoNego = Nothing 'LUCV20161115
    Set MatIfiGastoFami = Nothing 'LUCV20161115
End Sub

Private Sub Cmdguardar_Click()
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim GrabarDatos As Boolean
    Dim rsGastoNeg As ADODB.Recordset
    Dim rsGastoFam As ADODB.Recordset
    Dim rsOtrosIng As ADODB.Recordset
    Dim rsBalGen As ADODB.Recordset
    Dim MatActiPasivo As Variant
    Dim MatActiPasivoDet As Variant
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim rsRatiosActual As ADODB.Recordset
    Dim rsRatiosAceptableCritico As ADODB.Recordset
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    
    Dim nContadorTotal As Integer
    Dim nContadorDet As Integer
    Dim nContador As Integer
    Set rsGastoNeg = IIf(feGastosNegocio.rows - 1 > 0, feGastosNegocio.GetRsNew(), Nothing)
    Set rsGastoFam = IIf(feGastosFamiliares.rows - 1 > 0, feGastosFamiliares.GetRsNew(), Nothing)
    Set rsOtrosIng = IIf(feOtrosIngresos.rows - 1 > 0, feOtrosIngresos.GetRsNew(), Nothing)
        
    'Contar Totales y Detalles (ActivoPasivo) -> Filas ******
     nContadorTotal = 0
     nContadorDet = 0
     For I = 1 To feBalanceGeneral.rows - 1
        If feBalanceGeneral.TextMatrix(I, 2) = 1000 Or feBalanceGeneral.TextMatrix(I, 2) = 1001 Then
        nContadorTotal = nContadorTotal + 1
        Else
        nContadorDet = nContadorDet + 1
        End If
    Next I
    'Fin Filas <-**********
    
    'LUCV20162606, Carga Matriz Activo, Pasivo, Patrimonio, Totales **********->
    I = 0: j = 0: K = 0: nContador = 0
    ReDim MatActiPasivo(nContadorTotal + 1, 5)
    ReDim MatActiPasivoDet(nContadorDet + 1, 5)
    While feBalanceGeneral.rows - 1 > nContador
        I = I + 1
        'Para Cargar Datos en Matriz-> CredFormEvalActivoPasivo
        If feBalanceGeneral.TextMatrix(I, 3) = "" Then
            j = j + 1
            MatActiPasivo(j, 1) = feBalanceGeneral.TextMatrix(I, 1)
            MatActiPasivo(j, 2) = feBalanceGeneral.TextMatrix(I, 2)
            MatActiPasivo(j, 3) = feBalanceGeneral.TextMatrix(I, 3)
            MatActiPasivo(j, 4) = feBalanceGeneral.TextMatrix(I, 4)
            MatActiPasivo(j, 5) = CDbl(feBalanceGeneral.TextMatrix(I, 5))
         Else 'Para Cargar Datos en Matriz-> CredFormEvalActivoPasivoDet
             K = K + 1
            MatActiPasivoDet(K, 1) = feBalanceGeneral.TextMatrix(I, 1)
            MatActiPasivoDet(K, 2) = feBalanceGeneral.TextMatrix(I, 2)
            MatActiPasivoDet(K, 3) = feBalanceGeneral.TextMatrix(I, 3)
            MatActiPasivoDet(K, 4) = feBalanceGeneral.TextMatrix(I, 4)
            MatActiPasivoDet(K, 5) = CDbl(feBalanceGeneral.TextMatrix(I, 5))
        End If
             nContador = nContador + 1
    Wend
    'Fin LUCV20162606 <-**********
    
    'Flex a Matriz Referidos **********->
        ReDim MatReferidos(feReferidos.rows - 1, 6)
        For I = 1 To feReferidos.rows - 1
            MatReferidos(I, 1) = feReferidos.TextMatrix(I, 0)
            MatReferidos(I, 2) = feReferidos.TextMatrix(I, 1)
            MatReferidos(I, 3) = feReferidos.TextMatrix(I, 2)
            MatReferidos(I, 4) = feReferidos.TextMatrix(I, 3)
            MatReferidos(I, 5) = feReferidos.TextMatrix(I, 4)
            MatReferidos(I, 6) = feReferidos.TextMatrix(I, 5)
         Next I
    'Fin Referidos
    
    If ValidaDatos Then
        If MsgBox("Los Datos seran Grabados, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        If txtUltEndeuda.Text = "__/__/____" Then
            txtUltEndeuda.Text = "01/01/1900"
        End If

        Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
        Set objPista = New COMManejador.Pista
        Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
        If fnTipoPermiso = 3 Then
        GrabarDatos = oNCOMFormatosEval.GrabarCredFormEvalFormato1_5(sCtaCod, fnFormato, fnTipoRegMant, _
                                                                    Trim(txtGiroNeg.Text), CInt(spnExpEmpAnio.valor), CInt(spnExpEmpMes.valor), CInt(spnTiempoLocalAnio.valor), _
                                                                    CInt(spnTiempoLocalMes.valor), CDbl(txtUltEndeuda.Text), Format(txtFecUltEndeuda.Text, "yyyymmdd"), _
                                                                    lnCondLocal, IIf(txtCondLocalOtros.Visible = False, "", txtCondLocalOtros.Text), CDbl(txtExposicionCredito.Text), _
                                                                    Format(txtFechaEvaluacion.Text, "yyyymmdd"), _
                                                                    Format(txtFechaVisita.Text, "yyyymmdd"), _
                                                                    txtEntornoFamiliar2.Text, txtGiroUbicacion2.Text, _
                                                                    txtExperiencia2.Text, txtFormalidadNegocio2.Text, _
                                                                    txtColaterales2, txtDestino2.Text, _
                                                                    txtComentario.Text, MatReferidos, MatIfiGastoNego, MatIfiGastoFami, _
                                                                    rsGastoFam, rsOtrosIng, rsGastoNeg, _
                                                                    CDbl(txtIngresoNegocio.Text), _
                                                                    CDbl(txtEgresoNegocio.Text), _
                                                                    CDbl(txtMargenBruto.Text), _
                                                                    MatActiPasivo, MatActiPasivoDet, , , _
                                                                    gRatioCapacidadPago, _
                                                                    CDbl(Replace(txtCapacidadNeta.Text, "%", "")), _
                                                                    gRatioEndeudamiento, _
                                                                    CDbl(Replace(txtEndeudamiento.Text, "%", "")), _
                                                                    gRatioIngresoNetoNego, _
                                                                    CDbl(txtIngresoNeto.Text), _
                                                                    gRatioExcedenteMensual, _
                                                                    CDbl(txtExcedenteMensual.Text), , , , , , , fnColocCondi, , MatBalActCorr, MatBalActNoCorr)
                                                                    ''Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja - MatBalActCorr, MatBalActNoCorr
                Call oDCOMFormatosEval.RecalculaIndicadoresyRatiosEvaluacion(sCtaCod)
                Set rsRatiosActual = oDCOMFormatosEval.RecuperaDatosRatios(sCtaCod)
                Set rsRatiosAceptableCritico = oDCOMFormatosEval.RecuperaDatosRatiosAceptableCritico(sCtaCod)
        'JOEP20180725 ERS034-2018
            Call EmiteFormRiesgoCamCred(sCtaCod)
        'JOEP20180725 ERS034-2018
        Else
        'GrabarDatos = oNCOMFormatosEval.GrabarCredEvaluacionVerif(sCtaCod, Trim(txtVerif.Text), GeneraMovNro(gdFecSis, gsCodAge, gsCodUser))
        End If
            If GrabarDatos Then
                fbGrabar = True
                    'RECO20161020 ERS060-2016 **********************************************************
                    Dim oNCOMColocEval As New NCOMColocEval
                    'Dim lcMovNro As String 'LUCV20181220 Coment�, Anexo01 de Acta 199-2018
                    lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'LUCV20181220 Agreg�, Anexo01 de Acta 199-2018
                    
                    If Not ValidaExisteRegProceso(sCtaCod, gTpoRegCtrlEvaluacion) Then
                       'lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'LUCV20181220 Coment�, Anexo01 de Acta 199-2018
                       'objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato 2", sCtaCod, gCodigoCuenta 'LUCV20181220 Coment�, Anexo01 de Acta 199-2018
                       Call oNCOMColocEval.insEstadosExpediente(sCtaCod, "Evaluacion de Credito", lcMovNro, "", "", "", 1, 2001, gTpoRegCtrlEvaluacion)
                       Set oNCOMColocEval = Nothing
                    End If
                    'RECO FIN **************************************************************************
                    If fnTipoRegMant = 1 Then
                        objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato 2", sCtaCod, gCodigoCuenta 'LUCV20181220 Agreg�, Anexo01 de Acta 199-2018
                        Set objPista = Nothing 'LUCV20181220 Agreg�, Anexo01 de Acta 199-2018
                        MsgBox "Los datos se grabaron correctamente", vbInformation, "Aviso"
                    Else
                        objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, "Evaluacion Credito Formato 2", sCtaCod, gCodigoCuenta 'LUCV20181220 Agreg�, Anexo01 de Acta 199-2018
                        Set objPista = Nothing 'LUCV20181220 Agreg�, Anexo01 de Acta 199-2018
                        MsgBox "Los datos se actualizaron correctamente", vbInformation, "Aviso"
                    End If
                
                    'Habilita / Deshabilita Botones - Text
                    If fnEstado = 2000 Then             '*****-> Si es Solicitado
                        If fnColocCondi <> 4 Then
                            Me.cmdInformeVisita.Enabled = True
                            Me.cmdVerCar.Enabled = False
                        Else
                            Me.cmdInformeVisita.Enabled = False
                            Me.cmdVerCar.Enabled = False
                        End If
                        Me.cmdImprimir.Enabled = False
                        Me.cmdGuardar.Enabled = False
                    Else                                '*****-> Sugerido +
                        Me.cmdImprimir.Enabled = True
                        Me.cmdGuardar.Enabled = False
                        If fnColocCondi <> 4 Then
                            Me.cmdVerCar.Enabled = True 'No refinanciado
                            Me.cmdInformeVisita.Enabled = True
                        Else
                            Me.cmdVerCar.Enabled = False
                            Me.cmdInformeVisita.Enabled = False
                        End If
                    End If
                    
                    '*****->No Refinanciados (Propuesta Credito)
                    If fnColocCondi <> 4 Then
                        txtFechaVisita.Enabled = True
                        txtEntornoFamiliar2.Enabled = True
                        txtGiroUbicacion2.Enabled = True
                        txtExperiencia2.Enabled = True
                        txtFormalidadNegocio2.Enabled = True
                        txtColaterales2.Enabled = True
                        txtDestino2.Enabled = True
                     Else
                        framePropuesta.Enabled = False
                        txtFechaVisita.Enabled = False
                        txtEntornoFamiliar2.Enabled = False
                        txtGiroUbicacion2.Enabled = False
                        txtExperiencia2.Enabled = False
                        txtFormalidadNegocio2.Enabled = False
                        txtColaterales2.Enabled = False
                        txtDestino2.Enabled = False
                    End If  '*****->Fin No Refinanciados
                    
                    'Actualizaci�n de los Ratios
                        txtCapacidadNeta.Text = CStr(rsRatiosActual!nCapPagNeta * 100) & "%"
                        txtEndeudamiento.Text = CStr(rsRatiosActual!nEndeuPat * 100) & "%"
                        txtIngresoNeto.Text = Format(rsRatiosActual!nIngreNeto, "#,##0.00")
                        txtExcedenteMensual.Text = Format(rsRatiosActual!nExceMensual, "#,##0.00")
                        
                    'Ratios: Aceptable / Critico ->*****
                    If Not (rsRatiosAceptableCritico.EOF Or rsRatiosAceptableCritico.BOF) Then
                    If rsRatiosAceptableCritico!nCapPag = 1 Then 'Capacidad Pago
                        Me.lblCapaAceptable.Caption = "Aceptable"
                        Me.lblCapaAceptable.ForeColor = &H8000&
                    Else
                        Me.lblCapaAceptable.Caption = "Cr�tico"
                        Me.lblCapaAceptable.ForeColor = vbRed
                    End If
                    
                    If rsRatiosAceptableCritico!nEndeud = 1 Then 'Endeudamiento Pat.
                        Me.lblEndeAceptable.Caption = "Aceptable"
                        Me.lblEndeAceptable.ForeColor = &H8000&
                    Else
                        Me.lblEndeAceptable.Caption = "Cr�tico"
                        Me.lblEndeAceptable.ForeColor = vbRed
                    End If
                    Else
                        lblCapaAceptable.Visible = False
                        lblEndeAceptable.Visible = False
                    End If
                'Fin Ratios <-****
                    Set rsRatiosActual = Nothing
                    Set rsRatiosAceptableCritico = Nothing
            Else
                MsgBox "Hubo errores al grabar la informaci�n", vbError, "Error"
            End If
    'Else
    'MsgBox "Ha Ocurrido un Problema o Faltan Ingresar Datos", vbInformation, "Aviso"
    End If
End Sub
Private Sub cmdVerCar_Click()
    Call GeneraVerCar
End Sub
Private Sub cmdInformeVisita_Click()
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim rsInfVisita As ADODB.Recordset
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsInfVisita = New ADODB.Recordset
    Set rsInfVisita = oDCOMFormatosEval.RecuperarDatosInformeVisitaFormato1_6(sCtaCod)
    cmdInformeVisita.Enabled = False
    
    If (rsInfVisita.EOF And rsInfVisita.BOF) Then
        Set oDCOMFormatosEval = Nothing
        MsgBox "No existe datos para este reporte.", vbOKOnly, "Atenci�n"
        Exit Sub
    End If
    Call CargaInformeVisitaPDF(rsInfVisita) 'gCredReportes
    Set rsInfVisita = Nothing
    cmdInformeVisita.Enabled = True
End Sub
Private Sub cmdImprimir_Click()
    Call ImprimirFormatoEvaluacion
End Sub
Private Sub cmdCancelar_Click()
    Unload frmCredFormEvalCuotasIfis
    Unload Me
    Set MatIfiGastoNego = Nothing 'LUCV20161115
    Set MatIfiGastoFami = Nothing 'LUCV20161115
End Sub
Private Sub cmdAgregarRef_Click()
    If feReferidos.rows - 1 < 25 Then
        feReferidos.lbEditarFlex = True
        feReferidos.AdicionaFila
        feReferidos.SetFocus
        feReferidos.AvanceCeldas = Horizontal
        SendKeys "{Enter}"
    Else
        MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
    End If
End Sub
Private Sub cmdQuitar2_Click()
    If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        feReferidos.EliminaFila (feReferidos.row)
    End If
End Sub

'LUCV20160620, KeyPress / GotFocus / LostFocus ->**********
    'TAB0 -> Ingresos/Egresos
Private Sub spnTiempoLocalAnio_KeyPress(KeyAscii As Integer) 'TiempoMismoLocal
    If KeyAscii = 13 Then
        spnTiempoLocalMes.SetFocus
    End If
End Sub
Private Sub spnTiempoLocalMes_KeyPress(KeyAscii As Integer) 'TiempoMismoLocal
    If KeyAscii = 13 Then
        OptCondLocal(1).SetFocus
    End If
End Sub
Private Sub OptCondLocal_KeyPress(index As Integer, KeyAscii As Integer) 'CondicionLocal
    If KeyAscii = 13 Then
        txtIngresoNegocio.SetFocus
    End If
End Sub
Private Sub txtCondLocalOtros_KeyPress(KeyAscii As Integer) 'OtroCondicionLocal
    If KeyAscii = 13 Then
        SSTabIngresos.Tab = 0
        txtIngresoNegocio.SetFocus
    End If
End Sub
Private Sub txtIngresoNegocio_KeyPress(KeyAscii As Integer) 'Ingresos
   KeyAscii = NumerosDecimales(txtIngresoNegocio, KeyAscii, 10, , True)
    If KeyAscii = 45 Then KeyAscii = 0
    If KeyAscii = 13 Then
        SSTabIngresos.Tab = 0
        txtEgresoNegocio.SetFocus
    End If
End Sub
Private Sub txtEgresoNegocio_KeyPress(KeyAscii As Integer) 'EgresoVenta
    KeyAscii = NumerosDecimales(txtEgresoNegocio, KeyAscii, 10, , True)
   If KeyAscii = 45 Then KeyAscii = 0
    If KeyAscii = 13 Then
        Me.feBalanceGeneral.SetFocus
        Me.feBalanceGeneral.row = 1
        Me.feBalanceGeneral.Col = 5
        SendKeys "{F2}"
    End If
End Sub

   'TAB1 ->PropuestaCredito
Private Sub txtFechaVisita_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEntornoFamiliar2.SetFocus
        
        If Not IsDate(txtFechaVisita) Then
            MsgBox "Verifique Dia,Mes,A�o , Fecha Incorrecta", vbInformation, "Aviso"
            txtFechaVisita.SetFocus
        End If
        
    End If
End Sub

Private Sub txtEntornoFamiliar2_KeyPress(KeyAscii As Integer) 'Entornofamiliar
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtGiroUbicacion2.SetFocus
    End If
End Sub
Private Sub txtGiroUbicacion2_KeyPress(KeyAscii As Integer) 'SobreGiro
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtExperiencia2.SetFocus
    End If
End Sub
Private Sub txtExperiencia2_KeyPress(KeyAscii As Integer) 'ExperienciaCrediticia
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtFormalidadNegocio2.SetFocus
    End If
End Sub
Private Sub txtFormalidadNegocio2_KeyPress(KeyAscii As Integer) 'ConsistenciaInformacion
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtColaterales2.SetFocus
    End If
End Sub
Private Sub txtColaterales2_KeyPress(KeyAscii As Integer) 'Colaterales_Garantias
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtDestino2.SetFocus
    End If
End Sub
Private Sub txtDestino2_KeyPress(KeyAscii As Integer) 'Destino del cr�dito
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        SSTabIngresos.Tab = 2
        'If fnColocCondi = 1 Then 'LUCV20171115, Agreg� segun correo: RUSI
        If Not fbTieneReferido6Meses Then
            txtComentario.SetFocus
        Else
            cmdGuardar.SetFocus
        End If
    End If
End Sub
    'TAB1 ->ComentarioReferido
Private Sub txtComentario_KeyPress(KeyAscii As Integer) 'Referidos/ ComentariosReferidos
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        If fnColocCondi = 1 Then
            cmdAgregarRef.SetFocus
        End If
    End If
End Sub

    'GotFocus / LostFocus
Private Sub txtIngresoNegocio_GotFocus()
    fEnfoque txtEgresoNegocio
End Sub
Private Sub txtIngresoNegocio_LostFocus()
    If Trim(txtIngresoNegocio.Text) = "" Then
        txtIngresoNegocio.Text = "0.00"
    Else
        txtIngresoNegocio.Text = Format(txtIngresoNegocio.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(1)
End Sub
Private Sub txtEgresoNegocio_GotFocus()
    fEnfoque txtEgresoNegocio
End Sub
Private Sub txtEgresoNegocio_LostFocus()
    If Trim(txtEgresoNegocio.Text) = "" Then
        txtEgresoNegocio.Text = "0.00"
    Else
        txtEgresoNegocio.Text = Format(txtEgresoNegocio.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(1)
End Sub
'LUCV20160620, KeyPress / GotFocus / LostFocus Fin <-**********

'Para Buscar Cuotas IFIs (GastosNegocio / GastosFamiliares)**********->
Private Sub feGastosNegocio_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    Editar = Split(Me.feGastosNegocio.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        SendKeys "{TAB}", True
        Exit Sub
    End If
End Sub
Private Sub feGastosNegocio_Click() 'GastosNegocio
    If feGastosNegocio.Col = 3 Then
        If CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 0)) = gCodCuotaIfiGastoNego Then
            feGastosNegocio.ListaControles = "0-0-0-1-0"
        Else
            feGastosNegocio.ListaControles = "0-0-0-0-0"
        End If
    End If
     
    Select Case CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 1))
        Case gCodCuotaIfiGastoNego
            'Me.feGastosNegocio.CellBackColor = &HC0FFFF
            Me.feGastosNegocio.BackColorRow &HC0FFFF, True
            Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
            Me.feGastosNegocio.ForeColorRow vbBlack, True
        Case gCodCuotaCmac
            Me.feGastosNegocio.ColumnasAEditar = "X-X-X-X-X"
            Me.feGastosNegocio.ForeColorRow vbBlack, True
        Case Else
            Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
    End Select

End Sub
Private Sub feGastosNegocio_EnterCell() 'LUCV20160525 - Me permite Buscar OtrasCuotasIFIs (GastosNegocio)
    If feGastosNegocio.Col = 3 Then
        If CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 0)) = gCodCuotaIfiGastoNego Then
            feGastosNegocio.ListaControles = "0-0-0-1-0"
        Else
            feGastosNegocio.ListaControles = "0-0-0-0-0"
        End If
    End If
    
    Select Case CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 1))
        Case gCodCuotaIfiGastoNego
            'Me.feGastosNegocio.CellBackColor = &HC0FFFF
            Me.feGastosNegocio.BackColorRow &HC0FFFF, True
            Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
            Me.feGastosNegocio.ForeColorRow vbBlack, True
        Case gCodCuotaCmac
            Me.feGastosNegocio.ColumnasAEditar = "X-X-X-X-X"
            Me.feGastosNegocio.ForeColorRow vbBlack, True
        Case Else
            Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
    End Select
End Sub
Private Sub feGastosNegocio_RowColChange() 'PresionarEnter:Monto
    If feGastosNegocio.Col = 3 Then
        feGastosNegocio.AvanceCeldas = Vertical
    Else
        feGastosNegocio.AvanceCeldas = Horizontal
    End If
    
    If feGastosNegocio.Col = 3 Then
        If CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 0)) = gCodCuotaIfiGastoNego Then
            feGastosNegocio.ListaControles = "0-0-0-1-0"
        Else
            feGastosNegocio.ListaControles = "0-0-0-0-0"
        End If
    End If
    
    Select Case CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 1))
        Case gCodCuotaIfiGastoNego
            'Me.feGastosNegocio.CellBackColor = &HC0FFFF
            Me.feGastosNegocio.BackColorRow &HC0FFFF, True
            Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
            Me.feGastosNegocio.ForeColorRow vbBlack, True
        Case gCodCuotaCmac
            Me.feGastosNegocio.ColumnasAEditar = "X-X-X-X-X"
            Me.feGastosNegocio.ForeColorRow vbBlack, True
        Case Else
            Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
    End Select
End Sub
Private Sub feGastosNegocio_OnClickTxtBuscar(psMontoIfiGastoNego As String, psDescripcion As String) 'GastosNegocio
    psMontoIfiGastoNego = 0
    psDescripcion = ""
    psDescripcion = feGastosNegocio.TextMatrix(feGastosNegocio.row, 2) 'Cuotas Otras IFIs
    psMontoIfiGastoNego = feGastosNegocio.TextMatrix(feGastosNegocio.row, 3) 'Monto
    
    If psMontoIfiGastoNego = 0 Then
         fnTotalRefGastoNego = 0
        Set MatIfiGastoNego = Nothing
        frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosNegocio.TextMatrix(feGastosNegocio.row, 3))), fnTotalRefGastoNego, MatIfiGastoNego, feGastosNegocio.TextMatrix(feGastosNegocio.row, 2)
        psMontoIfiGastoNego = Format(fnTotalRefGastoNego, "#,##0.00")
    Else
        frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosNegocio.TextMatrix(feGastosNegocio.row, 3))), fnTotalRefGastoNego, MatIfiGastoNego, feGastosNegocio.TextMatrix(feGastosNegocio.row, 2)
        psMontoIfiGastoNego = Format(fnTotalRefGastoNego, "#,##0.00")
    End If
    
End Sub
Private Sub feGastosNegocio_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(feGastosNegocio.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feGastosNegocio.TextMatrix(pnRow, pnCol) < 0 Then
            feGastosNegocio.TextMatrix(pnRow, pnCol) = 0
        End If
    Else
        feGastosNegocio.TextMatrix(pnRow, pnCol) = 0
    End If
    
    'If Me.feGastosNegocio.Col = 3 And Me.feGastosNegocio.row = 11 Then 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
    If Me.feGastosNegocio.Col = 3 And Me.feGastosNegocio.row = 12 Then 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        Me.feGastosFamiliares.SetFocus
        feGastosFamiliares.row = 1
        feGastosFamiliares.Col = 3
        SendKeys "{TAB}"
        SendKeys "{F2}"
    End If
    Call CalculoTotal(1)
End Sub
Private Sub feGastosFamiliares_KeyPress(KeyAscii As Integer)
    If (feGastosFamiliares.Col = 1 And feGastosFamiliares.row = 1) Or (feGastosFamiliares.Col = 3 And feGastosFamiliares.row = 7) Then
        If KeyAscii = 13 Then
            feOtrosIngresos.row = 1
            feOtrosIngresos.Col = 3
            EnfocaControl feOtrosIngresos
            SendKeys "{Enter}", True
        End If
    End If
End Sub
Private Sub feGastosFamiliares_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    Editar = Split(Me.feGastosFamiliares.ColumnasAEditar, "-")
    If Me.feGastosFamiliares.row <> 1 Then
        If Editar(pnCol) = "X" Then
            MsgBox "Esta celda no es editable", vbInformation, "Aviso"
            Cancel = False
            SendKeys "{TAB}", True
            Exit Sub
        End If
    End If
End Sub
Private Sub feGastosFamiliares_Click() 'GastosFamiliares
    If feGastosFamiliares.Col = 3 Then
        If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiGastoFami Then
            feGastosFamiliares.ListaControles = "0-0-0-1-0"
        Else
            feGastosFamiliares.ListaControles = "0-0-0-0-0"
        End If
    End If
    Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1))
        Case gCodCuotaIfiGastoFami
           'Me.feGastosFamiliares.CellBackColor = &HC0FFFF
           Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
           Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
           Me.feGastosFamiliares.ForeColorRow vbBlack, True
        Case gCodDeudaLCNUGastoFami
           Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
           Me.feGastosFamiliares.ForeColorRow vbBlack, True
        Case Else
           Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
     End Select
End Sub
Private Sub feGastosFamiliares_EnterCell() 'LUCV20160525 - Me permite Buscar CuotasIFIs(GastosFamiliares)
    If feGastosFamiliares.Col = 3 Then
        If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiGastoFami Then
            feGastosFamiliares.ListaControles = "0-0-0-1-0"
        Else
            feGastosFamiliares.ListaControles = "0-0-0-0-0"
        End If
    End If

    Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1))
        Case gCodCuotaIfiGastoFami
           'Me.feGastosFamiliares.CellBackColor = &HC0FFFF
           Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
           Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
           Me.feGastosFamiliares.ForeColorRow vbBlack, True
        Case gCodDeudaLCNUGastoFami
           Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
           Me.feGastosFamiliares.ForeColorRow vbBlack, True
        Case Else
           Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
     End Select
End Sub
Private Sub feGastosFamiliares_RowColChange() 'PresionarEnter:Monto
    If feGastosFamiliares.Col = 3 Then
        feGastosFamiliares.AvanceCeldas = Vertical
    Else
        feGastosFamiliares.AvanceCeldas = Horizontal
    End If
    
    If feGastosFamiliares.Col = 3 Then
        If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 0)) = gCodCuotaIfiGastoFami Then
            feGastosFamiliares.ListaControles = "0-0-0-1-0"
        Else
            feGastosFamiliares.ListaControles = "0-0-0-0-0"
        End If
    End If

    Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1))
        Case gCodCuotaIfiGastoFami
           'Me.feGastosFamiliares.CellBackColor = &HC0FFFF
           Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
           Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
           Me.feGastosFamiliares.ForeColorRow vbBlack, True
        Case gCodDeudaLCNUGastoFami
           Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
           Me.feGastosFamiliares.ForeColorRow vbBlack, True
        Case Else
           Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
     End Select
End Sub
Private Sub feGastosFamiliares_OnClickTxtBuscar(psMontoIfiGastoFami As String, psDescripcion As String) 'GastosFamiliares
    psMontoIfiGastoFami = 0
    psDescripcion = ""
    psDescripcion = feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2) 'Cuotas Otras IFIs
    psMontoIfiGastoFami = feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3) 'Monto
    
    If psMontoIfiGastoFami = 0 Then
        fnTotalRefGastoFami = 0
        Set MatIfiGastoFami = Nothing
        frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiGastoFami, feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2)
        psMontoIfiGastoFami = Format(fnTotalRefGastoFami, "#,##0.00")
    Else
        frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiGastoFami, feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2)
        psMontoIfiGastoFami = Format(fnTotalRefGastoFami, "#,##0.00")
    End If
End Sub
Private Sub feGastosFamiliares_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(feGastosFamiliares.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feGastosFamiliares.TextMatrix(pnRow, pnCol) < 0 Then
            feGastosFamiliares.TextMatrix(pnRow, pnCol) = 0
        End If
    Else
        feGastosFamiliares.TextMatrix(pnRow, pnCol) = 0
    End If
End Sub

Private Sub OptCondLocal_Click(index As Integer)
    Select Case index
        Case 1, 2, 3
            Me.txtCondLocalOtros.Visible = False
            Me.txtCondLocalOtros.Text = ""
        Case 4
            Me.txtCondLocalOtros.Visible = True
            Me.txtCondLocalOtros.Text = ""
    End Select
    lnCondLocal = index
End Sub

'***** LUCV20160528 - OnCellChange / RowColChange
Private Sub feReferidos_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 1 Or pnCol = 4 Then
        feReferidos.TextMatrix(pnRow, pnCol) = UCase(feReferidos.TextMatrix(pnRow, pnCol))
    End If
    
    Select Case pnCol
    Case 2
        If IsNumeric(feReferidos.TextMatrix(pnRow, pnCol)) Then
            If Len(feReferidos.TextMatrix(pnRow, pnCol)) = 8 Then
                Select Case CCur(feReferidos.TextMatrix(pnRow, pnCol))
                    Case Is > 0
                    Case Else
                        MsgBox "Por favor, verifique el DNI", vbInformation, "Alerta"
                        feReferidos.TextMatrix(pnRow, pnCol) = 0
                    Exit Sub
                End Select
            Else
                MsgBox "El DNI, tiene que ser 8 d�gitos.", vbInformation, "Alerta"
                feReferidos.TextMatrix(pnRow, pnCol) = 0
            End If
            
        Else
            MsgBox "El DNI, tiene que ser num�rico.", vbInformation, "Alerta"
            feReferidos.TextMatrix(pnRow, pnCol) = 0
        End If
    Case 3
        If IsNumeric(feReferidos.TextMatrix(pnRow, pnCol)) Then
            If Len(feReferidos.TextMatrix(pnRow, pnCol)) = 9 Then
                Select Case CCur(feReferidos.TextMatrix(pnRow, pnCol))
                Case Is > 0
                Case Else
                    MsgBox "Tel�fono Mal Ingresado", vbInformation, "Alerta"
                    feReferidos.TextMatrix(pnRow, pnCol) = 0
                    Exit Sub
                End Select
            Else
                MsgBox "Faltan caracteres en el tel�fono / celular.", vbInformation, "Alerta"
                feReferidos.TextMatrix(pnRow, pnCol) = 0
            End If
        Else
            MsgBox "El telefono, solo permite ingreso de datos tipo num�rico." & Chr(10) & "Ejemplo: 065404040, 984047523 ", vbInformation, "Alerta"
            feReferidos.TextMatrix(pnRow, pnCol) = 0
        End If
'    Case 5
'        If IsNumeric(feReferidos.TextMatrix(pnRow, pnCol)) Then
'            If Len(feReferidos.TextMatrix(pnRow, pnCol)) = 8 Then
'                Select Case CCur(feReferidos.TextMatrix(pnRow, pnCol))
'                Case Is > 0
'                Case Else
'                    MsgBox "El DNI del referido, tiene que contener 8 d�gitos", vbInformation, "Alerta"
'                    feReferidos.TextMatrix(pnRow, pnCol) = 0
'                    Exit Sub
'                End Select
'            Else
'                MsgBox "El DNI del referido, tiene que ser 8 d�gitos", vbInformation, "Alerta"
'                feReferidos.TextMatrix(pnRow, pnCol) = 0
'            End If
'        Else
'            MsgBox "El DNI del referido, s�lo permite ingreso de datos tipo num�rico.", vbInformation, "Alerta"
'            feReferidos.TextMatrix(pnRow, pnCol) = 0
'        End If
    End Select
End Sub

Private Sub feReferidos_RowColChange()
    If feReferidos.Col = 1 Then
        feReferidos.MaxLength = "200"
    ElseIf feReferidos.Col = 2 Then
        feReferidos.MaxLength = "8"
    ElseIf feReferidos.Col = 3 Then
        feReferidos.MaxLength = "9"
    ElseIf feReferidos.Col = 4 Then
        feReferidos.MaxLength = "200"
    ElseIf feReferidos.Col = 5 Then
        feReferidos.MaxLength = "8"
    End If
End Sub

Private Sub feBalanceGeneral_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    Editar = Split(Me.feBalanceGeneral.ColumnasAEditar, "-")
    If Me.feBalanceGeneral.row <> 1 Then
        If Editar(pnCol) = "X" Then
            MsgBox "Esta celda no es editable", vbInformation, "Aviso"
            Cancel = False
            SendKeys "{TAB}", True
            Exit Sub
        End If
    End If
End Sub
Private Sub feBalanceGeneral_KeyPress(KeyAscii As Integer)
    'If feBalanceGeneral.Col = 5 And feBalanceGeneral.row = 8 Then 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
    If feBalanceGeneral.Col = 5 And feBalanceGeneral.row = 10 Then 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            If KeyAscii = 13 Then
                EnfocaControl feGastosNegocio
                feGastosNegocio.row = 1
                feGastosNegocio.Col = 3
                SendKeys "{Enter}"
            End If
    End If
End Sub

Private Sub feBalanceGeneral_OnCellChange(pnRow As Long, pnCol As Long)
    'If pnRow = 3 Or pnRow = 6 Or pnRow = 7 Then 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
    If pnRow = 4 Or pnRow = 8 Or pnRow = 9 Then 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        MsgBox "No se puede Editar este Registro", vbInformation, "Aviso"
        feBalanceGeneral.TextMatrix(pnRow, pnCol) = ""
    End If
    
    If IsNumeric(feBalanceGeneral.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feBalanceGeneral.TextMatrix(pnRow, pnCol) < 0 Then
            feBalanceGeneral.TextMatrix(pnRow, pnCol) = 0
        End If
    Else
        feBalanceGeneral.TextMatrix(pnRow, pnCol) = 0
    End If
    
    Select Case feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2)
        Case 1000, 1001
            Me.feBalanceGeneral.BackColorRow (&H80000000)
            Me.feBalanceGeneral.ForeColorRow vbBlack, True
            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-X-X"
        'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        Case IIf((feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2) = 100 And feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 1) = 7026), 100, 0)
            Me.feBalanceGeneral.BackColorRow &HC0FFFF, True
            Me.feBalanceGeneral.ForeColorRow vbBlack, True
            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
        Case IIf((feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2) = 200 And feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 1) = 7026), 200, 0)
            Me.feBalanceGeneral.BackColorRow &HC0FFFF, True
            Me.feBalanceGeneral.ForeColorRow vbBlack, True
            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
        'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        Case 206
            Me.feBalanceGeneral.ForeColorRow vbBlack, True
            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case Else
            Me.feBalanceGeneral.BackColorRow (&HFFFFFF)
            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
    End Select
    Call CalculoTotal(2)
    Call CalculoTotal(1)
End Sub
Private Sub feBalanceGeneral_RowColChange() 'PresionarEnter:Monto
    If feBalanceGeneral.Col = 5 Then
        feBalanceGeneral.AvanceCeldas = Vertical
    Else
        feBalanceGeneral.AvanceCeldas = Horizontal
    End If
    
'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
If fnTipoRegMant = 2 Then
            If feBalanceGeneral.Col = 5 Then
                    If CInt(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 0)) = 5 Then
                        feBalanceGeneral.ListaControles = "0-0-0-0-1"
                    ElseIf CInt(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 0)) = 6 Then
                        feBalanceGeneral.ListaControles = "0-0-0-0-1"
                    Else
                        feBalanceGeneral.ListaControles = "0-0-0-0-0"
                    End If
            End If
Else
    If feBalanceGeneral.Col = 5 Then
        If CInt(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 0)) = 5 Then
            feBalanceGeneral.ListaControles = "0-0-0-0-1"
        ElseIf CInt(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 0)) = 6 Then
            feBalanceGeneral.ListaControles = "0-0-0-0-1"
        Else
            feBalanceGeneral.ListaControles = "0-0-0-0-0"
        End If
    End If
End If
'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
    
    Select Case feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2)
        Case 1000, 1001
            Me.feBalanceGeneral.BackColorRow (&H80000000)
            Me.feBalanceGeneral.ForeColorRow vbBlack, True
            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-X-X"
        'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        Case IIf((feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2) = 100 And feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 1) = 7026), 100, 0)
            Me.feBalanceGeneral.BackColorRow &HC0FFFF, True
            Me.feBalanceGeneral.ForeColorRow vbBlack, True
            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
        Case IIf((feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2) = 200 And feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 1) = 7026), 200, 0)
            Me.feBalanceGeneral.BackColorRow &HC0FFFF, True
            Me.feBalanceGeneral.ForeColorRow vbBlack, True
            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
        'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        Case 206
            Me.feBalanceGeneral.ForeColorRow vbBlack, True
            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case Else
            Me.feBalanceGeneral.BackColorRow (&HFFFFFF)
            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
        End Select
End Sub

Private Sub feOtrosIngresos_RowColChange() 'PresionarEnter:Monto
    If feOtrosIngresos.Col = 3 Then
        feOtrosIngresos.AvanceCeldas = Vertical
    Else
        feOtrosIngresos.AvanceCeldas = Horizontal
    End If
End Sub
Private Sub feOtrosIngresos_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(feOtrosIngresos.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feOtrosIngresos.TextMatrix(pnRow, pnCol) < 0 Then
            feOtrosIngresos.TextMatrix(pnRow, pnCol) = 0
        End If
    Else
        feOtrosIngresos.TextMatrix(pnRow, pnCol) = 0
    End If
    
    If Me.feOtrosIngresos.Col = 3 And Me.feOtrosIngresos.row = 5 Then
        Me.SSTabIngresos.Tab = 1
        SendKeys "{TAB}"
   End If
End Sub
'Fin <- LUCV20160528 - OnCellChange / RowColChange *****

'________________________________________________________________________________________________________________________
'*************************************************LUCV20160525: METODOS Varios **************************************************
Public Function Inicio(ByVal psTipoRegMant As Integer, ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnProducto As Integer, _
                       ByVal pnSubProducto As Integer, ByVal pnMontoExpEsteCred As Double, ByVal pbImprimir As Boolean, ByVal pnEstado As Integer) As Boolean
                       
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim rsDCredEval As ADODB.Recordset
    Dim rsDColCred As ADODB.Recordset
    Dim oTipoCam  As COMDConstSistema.NCOMTipoCambio
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
    Set oTipoCam = New COMDConstSistema.NCOMTipoCambio

    If psCtaCod <> -1 Then '*****-> cCtaCod
        gsOpeCod = ""
        lcMovNro = "" 'LUCV20181220 Agreg�, Anexo01 de Acta 199-2018
        nTC = oTipoCam.EmiteTipoCambio(gdFecSis, TCFijoDia)
        sCtaCod = psCtaCod
        fnTipoRegMant = psTipoRegMant
        ActXCodCta.NroCuenta = sCtaCod
        
        '(3: Analista, 2: Coordinador, 1: JefeAgencia)
        fnTipoPermiso = oNCOMFormatosEval.ObtieneTipoPermisoCredEval(gsCodCargo)  ' Obtener el tipo de Permiso, Segun Cargo
        Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
        Set rsDCredito = oDCOMFormatosEval.RecuperaSolicitudDatoBasicosEval(sCtaCod) ' Datos Basicos del Credito Solicitado
        
        If (rsDCredito!cActiGiro) = "" Then
            MsgBox "Por favor, actualizar los datos del cliente. " & Chr(13) & " (Actividad o Giro del negocio)", vbInformation, "Alerta"
            Exit Function
        End If
        
        '***** Datos b�sicos de cabecera de Formato
        fsGiroNego = IIf((rsDCredito!cActiGiro) = "", "", (rsDCredito!cActiGiro))
        fsCliente = Trim(rsDCredito!cPersNombre)
        fsAnioExp = CInt(rsDCredito!nAnio)
        fsMesExp = CInt(rsDCredito!nMes)
        fnColocCondi = rsDCredito!nColocCondicion
        fbTieneReferido6Meses = rsDCredito!bTieneReferido6Meses   'Si tiene evaluacion registrada 6 meses (LUCV20171115, agreg� seg�n correo: RUSI)
        fnFechaDeudaSbs = IIf(rsDCredito!dFechaUltimaDeudaSBS = "", "__/__/____", rsDCredito!dFechaUltimaDeudaSBS)
        fnMontoDeudaSbs = Format(CCur(rsDCredito!nMontoUltimaDeudaSBS), "#,##0.00")
    
        spnExpEmpAnio.valor = fsAnioExp
        spnExpEmpMes.valor = fsMesExp
        txtUltEndeuda.Text = Format(fnMontoDeudaSbs, "#,##0.00")
        txtFecUltEndeuda.Text = Format(fnFechaDeudaSbs, "dd/mm/yyyy")
        txtExposicionCredito.Text = Format(pnMontoExpEsteCred, "#,##0.00")
        txtFechaEvaluacion.Text = Format(gdFecSis, "dd/mm/yyyy")
        '***** Fin datos de cabecera
         
        Set rsDCredEval = oDCOMFormatosEval.RecuperaColocacCredEval(sCtaCod) 'Ojo: Recuperar Credito Si ha sido Registrado el Form. Eval.
        Set rsAceptableCritico = oDCOMFormatosEval.RecuperaDatosRatiosAceptableCritico(sCtaCod) 'Obtenemos Datos, Aceptable / Critico de los Ratios
            If fnTipoPermiso = 2 Then
                If rsDCredEval.RecordCount = 0 Then ' Si no hay credito registrado
                    MsgBox "El analista no ha registrado la Evaluacion respectiva", vbInformation, "Aviso"
                    fbPermiteGrabar = False
                Else
                    fbPermiteGrabar = True
                End If
            End If
        
        Set rsDCredito = Nothing
        Set rsDCredEval = Nothing
        
        fnEstado = pnEstado
        fnFormato = pnFormato
        SSTabIngresos.Tab = 0
        frameLinea.Visible = False 'Para Ocultar Tab->Linea Credito Aut.
        fbPermiteGrabar = False
        fbBloqueaTodo = False
        Else
            MsgBox "No se ha registrado el n�mero de cuenta del cr�dito a evaluar ", vbInformation, "Aviso"
        End If
        'Fin cCtaCod <-**********
        
        Set oDCOMFormatosEval = Nothing
        Set oTipoCam = Nothing
        Call CargaControlesInicio
        
        If fnTipoRegMant = 3 Then
            fbBloqueaTodo = True
            'LUCV20181220 Agreg�, Anexo01 de Acta 199-2018
            gsOpeCod = gCredConsultarEvaluacionCred
            lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
            Set objPista = New COMManejador.Pista
            objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gConsultar, "Evaluacion Credito Formato 2", sCtaCod, gCodigoCuenta
            Set objPista = Nothing
            'Fin LUCV20181220
        End If
        
        'Carga de Datos Segun Evento: (Registrar / Mantenimiento) *****->
        If CargaDatos Then
            If CargaControlesTipoPermiso(fnTipoPermiso, fbPermiteGrabar, fbBloqueaTodo) Then
                If fnTipoRegMant = 1 Then   'Para el Evento: "Registrar"
                    If Not rsCredEval.EOF Then
                        Call Mantenimiento
                        fnTipoRegMant = 2
                    Else
                        Call Registro
                        fnTipoRegMant = 1
                    End If
                 ElseIf fnTipoRegMant = 2 Then 'Para el Evento. "Mantenimiento"
                    If rsCredEval.EOF Then
                        Call Registro
                        fnTipoRegMant = 1
                    Else
                        Call Mantenimiento
                        fnTipoRegMant = 2
                    End If
                  ElseIf fnTipoRegMant = 3 Then 'Para el Evento. "Consulta"
                  Call Mantenimiento
                  fnTipoRegMant = 3
                End If
            Else
                Unload Me
                Exit Function
            End If
    Else
        If CargaControlesTipoPermiso(1, False) Then
        End If
    End If
    'Fin Crga <-*****
    
    
    
     'Habilita / Deshabilita Botones - Text
        If fnEstado = 2000 Then             '*****-> Si es Solicitado
            'Me.cmdGuardar.Enabled = True
            Me.cmdImprimir.Enabled = False
            Me.cmdInformeVisita.Enabled = False
            If fnColocCondi <> 4 Then
                Me.cmdVerCar.Enabled = False
            Else
                Me.cmdVerCar.Enabled = False
            End If
        Else                                '*****-> Sugerido +
            Me.cmdImprimir.Enabled = True
            'Me.cmdGuardar.Enabled = True
            If fnColocCondi <> 4 Then
                Me.cmdVerCar.Enabled = True 'No refinanciado
                Me.cmdInformeVisita.Enabled = True
            Else
                Me.cmdVerCar.Enabled = False
                Me.cmdInformeVisita.Enabled = False
            End If
        End If
                
        '*****->No Refinanciados (Propuesta Credito)
        If fnColocCondi <> 4 Then
            txtFechaVisita.Enabled = True
            txtEntornoFamiliar2.Enabled = True
            txtGiroUbicacion2.Enabled = True
            txtExperiencia2.Enabled = True
            txtFormalidadNegocio2.Enabled = True
            txtColaterales2.Enabled = True
            txtDestino2.Enabled = True
         Else
            framePropuesta.Enabled = False
            txtFechaVisita.Enabled = False
            txtEntornoFamiliar2.Enabled = False
            txtGiroUbicacion2.Enabled = False
            txtExperiencia2.Enabled = False
            txtFormalidadNegocio2.Enabled = False
            txtColaterales2.Enabled = False
            txtDestino2.Enabled = False
        End If  '*****->Fin No Refinanciados
        
    
    

    Set rsAceptableCritico = Nothing
    fbGrabar = False
    Call CalculoTotal(1)
    If Not pbImprimir Then
        Me.Show 1
    Else
        cmdImprimir_Click
    End If
    Inicio = fbGrabar
End Function

'***** LUCV20160529 / feReferidos2
Public Function ValidaDatosReferencia() As Boolean
    Dim I As Integer, j As Integer
    ValidaDatosReferencia = False
        If feReferidos.rows - 1 < 2 Then
            MsgBox "Debe registrar por lo menos 2 referencias para continuar", vbInformation, "Aviso"
            cmdAgregarRef.SetFocus
            ValidaDatosReferencia = False
            Exit Function
        End If
        For I = 1 To feReferidos.rows - 1  'Verfica Tipo de Valores del DNI
            If Trim(feReferidos.TextMatrix(I, 1)) <> "" Then
                For j = 1 To Len(Trim(feReferidos.TextMatrix(I, 2)))
                    If (Mid(feReferidos.TextMatrix(I, 2), j, 1) < "0" Or Mid(feReferidos.TextMatrix(I, 2), j, 1) > "9") Then
                       MsgBox "Uno de los Digitos del primer DNI de la fila " & I & " no es un Numero", vbInformation, "Aviso"
                       feReferidos.SetFocus
                       ValidaDatosReferencia = False
                       Exit Function
                    End If
                Next j
            End If
        Next I
        For I = 1 To feReferidos.rows - 1  'Verfica Longitud del DNI
            If Trim(feReferidos.TextMatrix(I, 1)) <> "" Then
                If Len(Trim(feReferidos.TextMatrix(I, 2))) <> gnNroDigitosDNI Then
                    MsgBox "El DNI de la fila " & I & " no es de " & gnNroDigitosDNI & " digitos", vbInformation, "Aviso"
                    feReferidos.SetFocus
                    ValidaDatosReferencia = False
                    Exit Function
                End If
            End If
        Next I
        For I = 1 To feReferidos.rows - 1  'Verfica Tipo de Valores del Telefono
            If Trim(feReferidos.TextMatrix(I, 1)) <> "" Then
                For j = 1 To Len(Trim(feReferidos.TextMatrix(I, 3)))
                    If (Mid(feReferidos.TextMatrix(I, 3), j, 1) < "0" Or Mid(feReferidos.TextMatrix(I, 3), j, 1) > "9") Then
                       MsgBox "Uno de los Digitos del tel�fono de la fila " & I & " no es un Numero", vbInformation, "Aviso"
                       feReferidos.SetFocus
                       ValidaDatosReferencia = False
                       Exit Function
                    End If
                Next j
            End If
        Next I

        For I = 1 To feReferidos.rows - 1 'Verfica ambos DNI que no sean iguales
            For j = 1 To feReferidos.rows - 1
                If I <> j Then
                    If feReferidos.TextMatrix(I, 2) = feReferidos.TextMatrix(j, 2) Then
                        MsgBox "No se puede ingresar el mismo DNI mas de una vez en los referidos", vbInformation, "Alerta"
                        ValidaDatosReferencia = False
                        Exit Function
                    End If
                End If
            Next
        Next
    ValidaDatosReferencia = True
End Function

Public Function ValidaGrillas(ByVal Flex As FlexEdit) As Boolean
    Dim I As Integer
    ValidaGrillas = False
        For I = 1 To Flex.rows - 1
            If Flex.TextMatrix(I, 0) <> "" Then
                If Trim(Flex.TextMatrix(I, 1)) = "" Or Trim(Flex.TextMatrix(I, 3)) = "" Then
                    ValidaGrillas = False
                    Exit Function
                End If
            End If
        Next I
    ValidaGrillas = True
End Function

Public Function ValidaDatos() As Boolean
ValidaDatos = False
Dim nIndice As Integer
Dim I As Integer
Dim lsMensajeIfi As String 'LUCV20161115
    If fnTipoPermiso = 3 Then
    '********** Para TAB:0 -> Ingresos y Egresos
        If spnTiempoLocalAnio.valor = "" Then
            MsgBox "Ingrese Tiempo en el mismo local: A�os", vbInformation, "Aviso"
            SSTabIngresos.Tab = 0
            spnTiempoLocalAnio.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If spnTiempoLocalMes.valor = "" Then
            MsgBox "Ingrese Tiempo en el mismo local: Meses", vbInformation, "Aviso"
            SSTabIngresos.Tab = 0
            spnTiempoLocalMes.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If OptCondLocal(1).value = 0 And OptCondLocal(2).value = 0 And OptCondLocal(3).value = 0 And OptCondLocal(4).value = 0 Then
            MsgBox "Falta elegir la Condicion del local", vbInformation, "Aviso"
            ValidaDatos = False
            Exit Function
        End If
        If txtCondLocalOtros.Visible = True Then
            If txtCondLocalOtros.Text = "" Then
                MsgBox "Ingrese la Descripcion de la Opcion: Otro Local", vbInformation, "Aviso"
                SSTabIngresos.Tab = 0
                ValidaDatos = False
                Exit Function
            End If
        End If
        If Trim(txtGiroNeg.Text) = "" Then
            MsgBox "Falta ingresar el Giro del Negocio, Favor Actualizar los Datos del Cliente", vbInformation, "Aviso"
            SSTabIngresos.Tab = 0
            txtGiroNeg.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtFechaEvaluacion.Text) = "__/__/____" Then
            MsgBox "Falta Ingresar la Fecha de Evaluacion", vbInformation, "Aviso"
            SSTabIngresos.Tab = 0
            txtFechaEvaluacion.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If txtIngresoNegocio.Text = "" Then
            MsgBox "Falta Ingresar el Ingreso del Negocio", vbInformation, "Aviso"
            SSTabIngresos.Tab = 0
            txtIngresoNegocio.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If txtEgresoNegocio.Text = "" Then
            MsgBox "Falta Ingresar el Egreso del Negocio", vbInformation, "Aviso"
            SSTabIngresos.Tab = 0
            txtEgresoNegocio.SetFocus
            ValidaDatos = False
            Exit Function
        End If
            
    '********** Para TAB:1 -> Propuesta del Credito
        If fnColocCondi <> 4 Then 'Valida, si el credito no es refinanciado
            If Trim(txtFechaVisita.Text) = "__/__/____" Or Not IsDate(Trim(txtFechaVisita.Text)) Then
                MsgBox "Falta Ingresar la fecha de visita o el formato de fecha no es el correcto." & Chr(10) & " Formato: DD/MM/YYY", vbInformation, "Aviso"
                SSTabIngresos.Tab = 1
                txtFechaVisita.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            If txtEntornoFamiliar2.Text = "" Then
                MsgBox "Por favor Ingrese, El Entorno Familiar del Cliente o Representante", vbInformation, "Aviso"
                SSTabIngresos.Tab = 1
                txtEntornoFamiliar2.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            If txtGiroUbicacion2.Text = "" Then
                MsgBox "Por favor Ingrese, El Giro y la Ubicacion del Negocio", vbInformation, "Aviso"
                SSTabIngresos.Tab = 1
                txtGiroUbicacion2.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            If txtExperiencia2.Text = "" Then
                MsgBox "Por favor Ingrese, Sobre la Experiencia Crediticia", vbInformation, "Aviso"
                SSTabIngresos.Tab = 1
                txtExperiencia2.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            If txtFormalidadNegocio2.Text = "" Then
                MsgBox "Por favor Ingrese, La Formalidad del Negocio", vbInformation, "Aviso"
                SSTabIngresos.Tab = 1
                txtFormalidadNegocio2.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            If txtColaterales2.Text = "" Then
                MsgBox "Por favor Ingrese, Sobre las Garantias y Colaterales", vbInformation, "Aviso"
                SSTabIngresos.Tab = 1
                txtColaterales2.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            If txtDestino2.Text = "" Then
                MsgBox "Por favor Ingrese, El destino del Credito", vbInformation, "Aviso"
                SSTabIngresos.Tab = 1
                txtDestino.SetFocus
                ValidaDatos = False
                Exit Function
            End If
        End If
            
    '********** Para TAB:2 -> Comentarios y referidos
        'LUCV25072016->*****, Si el cliente es Nuevo -> Referente es Obligatorio
            'If fnColocCondi = 1 Then 'LUCV2017115, Seg�n correo: RUSI
            If Not fbTieneReferido6Meses Then
                frameReferido.Enabled = True
                frameComentario.Enabled = True
                    For I = 0 To feReferidos.rows - 1
                        If feReferidos.TextMatrix(I, 0) <> "" Then
                            If Trim(feReferidos.TextMatrix(I, 0)) = "" Or Trim(feReferidos.TextMatrix(I, 1)) = "" Or Trim(feReferidos.TextMatrix(I, 2)) = "" Or Trim(feReferidos.TextMatrix(I, 3)) = "" Or Trim(feReferidos.TextMatrix(I, 4)) = "" Then
                                MsgBox "Faltan datos en la lista de Referencias", vbInformation, "Aviso"
                                SSTabIngresos.Tab = 2
                                ValidaDatos = False
                                Exit Function
                            End If
                        End If
                    Next I
                   
                    If ValidaDatosReferencia = False Then 'Contenido de feReferidos2: Referidos
                        SSTabIngresos.Tab = 2
                        ValidaDatos = False
                        Exit Function
                    End If
                    If txtComentario.Text = "" Then
                        MsgBox "Por favor Ingrese, Comentarios", vbInformation, "Aviso"
                        SSTabIngresos.Tab = 2
                        txtComentario.SetFocus
                        ValidaDatos = False
                        Exit Function
                    End If
            Else
                'si el cliente es nuevo-> referido obligatorio
                    frameReferido.Enabled = False
                    feReferidos.Enabled = False
                    cmdAgregarRef.Enabled = False
                    cmdQuitar2.Enabled = False
                    txtComentario.Enabled = False 'Comentarios
                    frameComentario.Enabled = False
            End If
            'Fin LUCV25072016 <-*****
    
        '********** Para TAB:0 -> Validacion Grillas: GastosNegocio, OtrosIngresos, GastosFamiliares
            If ValidaGrillas(feGastosNegocio) = False Then
                MsgBox "Faltan datos en la lista de Gastos del Negocio", vbInformation, "Aviso"
                SSTabIngresos.Tab = 0
                ValidaDatos = False
                Exit Function
            End If
            If ValidaGrillas(feOtrosIngresos) = False Then
                MsgBox "Faltan datos en la lista de Otros Ingresos", vbInformation, "Aviso"
                SSTabIngresos.Tab = 0
                ValidaDatos = False
                Exit Function
            End If
            If ValidaGrillas(feGastosFamiliares) = False Then
                MsgBox "Faltan datos en la lista de Gastos Familiares", vbInformation, "Aviso"
                SSTabIngresos.Tab = 0
                ValidaDatos = False
                Exit Function
            End If
            
          '********** Para TAB:0 -> Grilla Balance General
          For nIndice = 1 To feBalanceGeneral.rows - 1
            'Activos
                If feBalanceGeneral.TextMatrix(nIndice, 2) = 1000 And feBalanceGeneral.TextMatrix(nIndice, 1) = 7025 Then 'Activo
                    If val(Replace(feBalanceGeneral.TextMatrix(nIndice, 5), ",", "")) <= 0 Then
                        MsgBox "No se ingresaron datos en el Activo", vbInformation, "Alerta"
                        SSTabIngresos.Tab = 0
                        ValidaDatos = False
                        Exit Function
                    End If
                End If
                If feBalanceGeneral.TextMatrix(nIndice, 2) = 100 And feBalanceGeneral.TextMatrix(nIndice, 1) = 7025 Then 'Activo Corriente
                    If val(Replace(feBalanceGeneral.TextMatrix(nIndice, 5), ",", "")) <= 0 Then
                        MsgBox "No se ingresaron datos en el Activo Corriente", vbInformation, "Alerta"
                        SSTabIngresos.Tab = 0
                        ValidaDatos = False
                        Exit Function
                    End If
                End If
            
            'Pasivos
                If feBalanceGeneral.TextMatrix(nIndice, 2) = 1000 And feBalanceGeneral.TextMatrix(nIndice, 1) = 7026 Then 'Pasivo
                    If val(Replace(feBalanceGeneral.TextMatrix(nIndice, 5), ",", "")) < 0 Then
                        MsgBox "No se ingresaron datos en el Pasivo", vbInformation, "Alerta"
                        SSTabIngresos.Tab = 0
                        ValidaDatos = False
                        Exit Function
                    End If
                End If
            'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                'If feBalanceGeneral.TextMatrix(nIndice, 2) = 100 And feBalanceGeneral.TextMatrix(nIndice, 1) = 7026 Then 'Pasivo
                 '   If val(Replace(feBalanceGeneral.TextMatrix(nIndice, 5), ",", "")) <= 0 Then
                  '      MsgBox "No se ingresaron datos en el Pasivo Corriente", vbInformation, "Alerta"
                   '     SSTabIngresos.Tab = 0
                    '    ValidaDatos = False
                     '   Exit Function
                    'End If
                'End If
            'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                'Patrimonio
                If feBalanceGeneral.TextMatrix(nIndice, 2) = 1001 And feBalanceGeneral.TextMatrix(nIndice, 1) = 7026 Then 'Patrimonio
                    If val(Replace(feBalanceGeneral.TextMatrix(nIndice, 5), ",", "")) <= 0 Then
                        MsgBox "Patrimonio = (Total Activo - Total Pasivo) " & Chr(10) & "- No se ingresaron datos en el patrimonio." & Chr(10) & "- El patrimonio no debe ser menor o igual que cero", vbInformation, "Alerta"
                        SSTabIngresos.Tab = 0
                        ValidaDatos = False
                        Exit Function
                    End If
                End If
        Next
      
        '********** Para TAB:0 -> Valida Margen Bruto
            If nMargenBruto <= 0 Then
                 MsgBox "Margen Bruto = (Ingresos del Negocio) - (Egreso por Venta)" & Chr(10) & "El Margen Bruto no debe ser menor o igual que cero.", vbInformation, "Alerta"
                 SSTabIngresos.Tab = 0
                 ValidaDatos = False
                 Exit Function
            End If
            
        'LUCV20161115, Agreg�->Seg�n ERS068-2016
        If Not ValidaIfiExisteCompraDeuda(sCtaCod, MatIfiGastoFami, MatIfiGastoNego, lsMensajeIfi) Or Len(Trim(lsMensajeIfi)) > 0 Then
            MsgBox "Ifi y Cuota registrada en detalle de cambio de estructura de pasivos no coincide:  " & Chr(10) & Chr(10) & " " & lsMensajeIfi & " ", vbInformation, "Aviso"
            SSTabIngresos.Tab = 0
            Exit Function
        End If
   End If
    ValidaDatos = True
End Function

Private Function CargaControlesTipoPermiso(ByVal TipoPermiso As Integer, ByVal pPermiteGrabar As Boolean, Optional ByVal pBloqueaTodo As Boolean = False) As Boolean
    '1: JefeAgencia->
    If TipoPermiso = 1 Then
        Call HabilitaControles(False, False, False)
        CargaControlesTipoPermiso = True
     '2: Coordinador->
    ElseIf TipoPermiso = 2 Then
        Call HabilitaControles(False, False, pPermiteGrabar)
        CargaControlesTipoPermiso = True
     '3: Analista ->
    ElseIf TipoPermiso = 3 Then
        Call HabilitaControles(True, False, True)
        CargaControlesTipoPermiso = True
     'Usuario sin Permisos al formato
    Else
        MsgBox "No tiene Permisos para este m�dulo", vbInformation, "Aviso"
        CargaControlesTipoPermiso = False
    End If
    If pBloqueaTodo Then 'Para el Caso despues de dar Verificacion
        Call HabilitaControles(True, True, False)
        CargaControlesTipoPermiso = True
    End If
End Function

Private Function HabilitaControles(ByVal pbHabilitaA As Boolean, ByVal pbHabilitaRatios As Boolean, ByVal pbHabilitaGuardar As Boolean)
'HabilitacionControlesAnalistas:     pbHabilitaA = True
    'Tab0: Ingresos/Egresos
    spnTiempoLocalAnio.Enabled = pbHabilitaA
    spnTiempoLocalMes.Enabled = pbHabilitaA
    OptCondLocal(1).Enabled = pbHabilitaA
    OptCondLocal(2).Enabled = pbHabilitaA
    OptCondLocal(3).Enabled = pbHabilitaA
    OptCondLocal(4).Enabled = pbHabilitaA
    txtCondLocalOtros.Enabled = pbHabilitaA
    'txtFechaEvaluacion.Enabled = pbHabilitaA
    txtIngresoNegocio.Enabled = pbHabilitaA
    txtEgresoNegocio.Enabled = pbHabilitaA
    feGastosNegocio.Enabled = pbHabilitaA
    feBalanceGeneral.Enabled = pbHabilitaA
    feOtrosIngresos.Enabled = pbHabilitaA
    feGastosFamiliares.Enabled = pbHabilitaA

    'Tab1: Propuesta/Credito
    txtFechaVisita.Enabled = pbHabilitaA
    txtEntornoFamiliar2.Enabled = pbHabilitaA
    txtGiroUbicacion2.Enabled = pbHabilitaA
    txtExperiencia2.Enabled = pbHabilitaA
    txtFormalidadNegocio2.Enabled = pbHabilitaA
    txtColaterales2.Enabled = pbHabilitaA
    txtDestino2.Enabled = pbHabilitaA

    'Tab2: Comentarios/Referidos
    txtComentario.Enabled = pbHabilitaA
    feReferidos.Enabled = pbHabilitaA
    cmdAgregarRef.Enabled = pbHabilitaA
    cmdQuitar2.Enabled = pbHabilitaA
    frameReferido.Enabled = pbHabilitaA

   'txtVerif.Enabled = pbHabilitaB
    If fnEstado = 2000 Then
        SSTabRatios.Visible = False
    Else
        SSTabRatios.Visible = pbHabilitaRatios
    End If

    'cmdInformeVisita.Enabled = pbHabilitaRatios
    'cmdVerCar.Enabled = pbHabilitaRatios
    'cmdImprimir.Enabled = pbHabilitaRatios
    cmdGuardar.Enabled = pbHabilitaGuardar
End Function
Private Sub CargaControlesInicio()
    Call CargarFlexEdit
    'DesHabilita la CargaInicial de Controles
    ActXCodCta.Enabled = False
    txtNombreCliente.Enabled = False
    txtExposicionCredito.Enabled = False
    txtGiroNeg.Enabled = False
    txtUltEndeuda.Enabled = False
    txtFecUltEndeuda.Enabled = False
    spnExpEmpAnio.Enabled = False
    spnExpEmpMes.Enabled = False
    txtMargenBruto.Enabled = False
    
    txtCapacidadNeta.Enabled = False
    txtEndeudamiento.Enabled = False
    txtIngresoNeto.Enabled = False
    txtExcedenteMensual.Enabled = False
    txtIngresoNegocio.Text = "0.00"
    txtEgresoNegocio.Text = "0.00"
    
    SSTabRatios.Visible = False
End Sub
Private Sub CargarFlexEdit() 'Registrar New Formato Evaluacion
    Dim lnFila As Integer
    Dim CargarFlexEdit As Boolean
    Dim nMonto As Double
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim I As Integer
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
    nMonto = Format(0, "00.00")
    
   CargarFlexEdit = oNCOMFormatosEval.CargaDatosFlexEdit(fnFormato, _
                                                        sCtaCod, _
                                                        rsFeGastoNeg, _
                                                        rsFeDatGastoFam, _
                                                        rsFeDatOtrosIng, _
                                                        rsFeDatBalanGen, _
                                                        rsFeDatActivos, _
                                                        rsFeDatPasivos, _
                                                        rsFeDatPasivosNo, _
                                                        rsFeDatPatrimonio, _
                                                        rsFeDatRef)
    'Gastos Negocio
    feGastosNegocio.Clear
    feGastosNegocio.FormaCabecera
    feGastosNegocio.rows = 2
    Call LimpiaFlex(feGastosNegocio)
        Do While Not rsFeGastoNeg.EOF
            feGastosNegocio.AdicionaFila
            lnFila = feGastosNegocio.row
            feGastosNegocio.TextMatrix(lnFila, 1) = rsFeGastoNeg!nConsValor
            feGastosNegocio.TextMatrix(lnFila, 2) = rsFeGastoNeg!cConsDescripcion
            feGastosNegocio.TextMatrix(lnFila, 3) = Format(rsFeGastoNeg!nMonto, "#,##0.00")
                    
            Select Case CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 1))
                Case gCodCuotaIfiGastoNego
                    'Me.feGastosNegocio.CellBackColor = &HC0FFFF
                    Me.feGastosNegocio.BackColorRow &HC0FFFF, True
                    Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
                    Me.feGastosNegocio.ForeColorRow vbBlack, True
                Case gCodCuotaCmac
                    Me.feGastosNegocio.ColumnasAEditar = "X-X-X-X-X"
                    Me.feGastosNegocio.ForeColorRow vbBlack, True
                Case Else
                    Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
            End Select
            rsFeGastoNeg.MoveNext
        Loop
    rsFeGastoNeg.Close
    Set rsFeGastoNeg = Nothing

    'Otros Ingresos
    feOtrosIngresos.Clear
    feOtrosIngresos.FormaCabecera
    feOtrosIngresos.rows = 2
    Call LimpiaFlex(feOtrosIngresos)
        Do While Not rsFeDatOtrosIng.EOF
            feOtrosIngresos.AdicionaFila
            lnFila = feOtrosIngresos.row
            feOtrosIngresos.TextMatrix(lnFila, 1) = rsFeDatOtrosIng!nConsValor
            feOtrosIngresos.TextMatrix(lnFila, 2) = rsFeDatOtrosIng!cConsDescripcion
            feOtrosIngresos.TextMatrix(lnFila, 3) = Format(rsFeDatOtrosIng!nMonto, "#,##0.00")
            rsFeDatOtrosIng.MoveNext
        Loop
    rsFeDatOtrosIng.Close
    Set rsFeDatOtrosIng = Nothing

    'Gastos Familiares
    feGastosFamiliares.Clear
    feGastosFamiliares.FormaCabecera
    feGastosFamiliares.rows = 2
    Call LimpiaFlex(feGastosFamiliares)
        Do While Not rsFeDatGastoFam.EOF
            feGastosFamiliares.AdicionaFila
            lnFila = feGastosFamiliares.row
            feGastosFamiliares.TextMatrix(lnFila, 1) = rsFeDatGastoFam!nConsValor
            feGastosFamiliares.TextMatrix(lnFila, 2) = rsFeDatGastoFam!cConsDescripcion
            feGastosFamiliares.TextMatrix(lnFila, 3) = Format(rsFeDatGastoFam!nMonto, "#,##0.00")
            
            Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1))
                Case gCodCuotaIfiGastoFami
                   'Me.feGastosFamiliares.CellBackColor = &HC0FFFF
                   Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
                   Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
                   Me.feGastosFamiliares.ForeColorRow vbBlack, True
                Case gCodDeudaLCNUGastoFami
                   Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
                   Me.feGastosFamiliares.ForeColorRow vbBlack, True
                Case Else
                   Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
            End Select
            rsFeDatGastoFam.MoveNext
     
        Loop
    rsFeDatGastoFam.Close
    Set rsFeDatGastoFam = Nothing
    
    'Balance General
    feBalanceGeneral.Clear
    feBalanceGeneral.FormaCabecera
    feBalanceGeneral.rows = 2
    Call LimpiaFlex(feBalanceGeneral)
        Do While Not rsFeDatBalanGen.EOF
            feBalanceGeneral.AdicionaFila
            lnFila = feBalanceGeneral.row
            feBalanceGeneral.TextMatrix(lnFila, 1) = rsFeDatBalanGen!nConsCod
            feBalanceGeneral.TextMatrix(lnFila, 2) = rsFeDatBalanGen!nConsValor
            feBalanceGeneral.TextMatrix(lnFila, 3) = rsFeDatBalanGen!nNumAut
            feBalanceGeneral.TextMatrix(lnFila, 4) = rsFeDatBalanGen!cConsDescripcion
            feBalanceGeneral.TextMatrix(lnFila, 5) = Format(rsFeDatBalanGen!nMonto, "#,##0.00")
            
            Select Case feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2)
                Case 1000, 1001
                    Me.feBalanceGeneral.BackColorRow (&H80000000)
                    Me.feBalanceGeneral.ForeColorRow vbBlack, True
                    Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-X-X"
                'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                Case IIf((feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2) = 100 And feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 1) = 7026), 100, 0)
                    Me.feBalanceGeneral.BackColorRow &HC0FFFF, True
                    Me.feBalanceGeneral.ForeColorRow vbBlack, True
                    Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
                Case IIf((feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2) = 200 And feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 1) = 7026), 200, 0)
                    Me.feBalanceGeneral.BackColorRow &HC0FFFF, True
                    Me.feBalanceGeneral.ForeColorRow vbBlack, True
                    Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
                'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                 Case 206
                    Me.feBalanceGeneral.ForeColorRow vbBlack, True
                    Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-X-X"
                 Case Else
                    Me.feBalanceGeneral.BackColorRow (&HFFFFFF)
                    Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
             End Select
            rsFeDatBalanGen.MoveNext
        Loop
    rsFeDatBalanGen.Close
    Set rsFeDatBalanGen = Nothing
End Sub
Private Function CargaDatos() As Boolean 'Mantenimiento Formatos
On Error GoTo ErrorCargaDatos
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim I As Integer
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
 
    CargaDatos = oNCOMFormatosEval.CargaDatosCredEvaluacion2(sCtaCod, _
                                                            fnFormato, _
                                                            rsCredEval, _
                                                            rsDatGastoNeg, _
                                                            rsDatGastoFam, _
                                                            rsDatOtrosIng, _
                                                            rsDatRef, _
                                                            rsDatActivos, _
                                                            rsDatPasivos, _
                                                            rsCuotaIFIs, _
                                                            rsPropuesta, _
                                                            rsCapacPagoNeta, _
                                                            rsDatRatioInd, _
                                                            rsDatActivoPasivo, _
                                                            rsDatIfiGastoNego, _
                                                            rsDatIfiGastoFami, _
                                                            rsDatVentaCosto, , , , , , , rsDatIfiBalActCorri, rsDatIfiBalActNoCorri)
                                            'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja rsDatIfiBalActCorri, rsDatIfiBalActNoCorri)
    Exit Function
ErrorCargaDatos:
    CargaDatos = False
    MsgBox Err.Description + ": Error al carga datos", vbInformation, "Error"
End Function

Private Sub CalculoTotal(ByVal pnTipo As Integer)
On Error GoTo ErrorCalculo
    nMontoAct = 0
    nMontoPas = 0
    nMontoPat = 0
    nMargenBruto = 0
    Select Case pnTipo
        Case 1:
                nMargenBruto = Format(CCur((txtIngresoNegocio.Text)) - CCur(txtEgresoNegocio.Text), "###," & String(15, "#") & "#,##0.00")
                txtMargenBruto.Text = Format(nMargenBruto, "###," & String(15, "#") & "#,##0.00")
        Case 2:
                'Activo Total
                'For I = 1 To 2 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                For I = 1 To 3 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                nMontoAct = nMontoAct + CCur(IIf(Trim(feBalanceGeneral.TextMatrix(I, 5)) = "", 0, Trim(feBalanceGeneral.TextMatrix(I, 5))))
                'feBalanceGeneral.TextMatrix(3, 5) = Format(nMontoAct, "#,##0.00")'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                feBalanceGeneral.TextMatrix(4, 5) = Format(nMontoAct, "#,##0.00") 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                Next I
                
                'Pasivo Total
                'For I = 4 To 6'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                For I = 5 To 8 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                nMontoPas = nMontoPas + CCur(IIf(Trim(feBalanceGeneral.TextMatrix(I, 5)) = "", 0, Trim(feBalanceGeneral.TextMatrix(I, 5))))
                'feBalanceGeneral.TextMatrix(7, 5) = Format(nMontoPas, "#,##0.00")'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                feBalanceGeneral.TextMatrix(9, 5) = Format(nMontoPas, "#,##0.00") 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                Next I
    
                'Patrimonio
                nMontoPat = Format((nMontoAct - nMontoPas), "#,##0.00")
                'feBalanceGeneral.TextMatrix(8, 5) = Format((nMontoPat), "#,##0.00")'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                feBalanceGeneral.TextMatrix(10, 5) = Format((nMontoPat), "#,##0.00") 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
    End Select
    
    Exit Sub
ErrorCalculo:
MsgBox "Informaci�n: Ingrese los datos Correctamente." & Chr(13) & "Detalles: " & Err.Description, vbInformation, "Informacion"
    Select Case pnTipo
        Case 1:
                txtIngresoNegocio.Text = "0.00"
                txtEgresoNegocio.Text = "0.00"
    End Select
 Call CalculoTotal(pnTipo)
End Sub
Private Function Registro()
    gsOpeCod = gCredRegistrarEvaluacionCred
    txtNombreCliente.Text = fsCliente
    txtGiroNeg.Text = fsGiroNego
    cmdInformeVisita.Enabled = False
    cmdVerCar.Enabled = False
    
    txtCapacidadNeta.Enabled = False
    txtEndeudamiento.Enabled = False
    txtIngresoNeto.Enabled = False
    txtExcedenteMensual.Enabled = False
    
        
    'si el cliente es nuevo-> referido obligatorio
    'If fnColocCondi = 1 Then 'LUCV2017115, Seg�n correo: RUSI
    If Not fbTieneReferido6Meses Then
        frameReferido.Enabled = True
        feReferidos.Enabled = True
        cmdAgregarRef.Enabled = True
        cmdQuitar2.Enabled = True
        txtComentario.Enabled = True 'Comentarios
        frameComentario.Enabled = True
    Else
        frameReferido.Enabled = False
        feReferidos.Enabled = False
        cmdAgregarRef.Enabled = False
        cmdQuitar2.Enabled = False
        txtComentario.Enabled = False 'Comentarios
        frameComentario.Enabled = False
    End If
    
    'Ratios: Aceptable / Critico ->*****
    If Not (rsAceptableCritico.BOF Or rsAceptableCritico.EOF) Then
        If rsAceptableCritico!nCapPag = 1 Then 'Capacidad Pago
            Me.lblCapaAceptable.Caption = "Aceptable"
            Me.lblCapaAceptable.ForeColor = &H8000&
        Else
            Me.lblCapaAceptable.Caption = "Cr�tico"
            Me.lblCapaAceptable.ForeColor = vbRed
        End If
        
        If rsAceptableCritico!nEndeud = 1 Then 'Endeudamiento Pat.
            Me.lblEndeAceptable.Caption = "Aceptable"
            Me.lblEndeAceptable.ForeColor = &H8000&
        Else
            Me.lblEndeAceptable.Caption = "Cr�tico"
            Me.lblEndeAceptable.ForeColor = vbRed
        End If
    Else
        lblCapaAceptable.Visible = False
        lblCapaAceptable.Visible = False
    End If
    'Fin Ratios <-****
    
    '*****->No Refinanciados (Propuesta Credito)
    If fnColocCondi <> 4 Then
            txtFechaVisita.Enabled = True
            txtEntornoFamiliar2.Enabled = True
            txtGiroUbicacion2.Enabled = True
            txtExperiencia2.Enabled = True
            txtFormalidadNegocio2.Enabled = True
            txtColaterales2.Enabled = True
            txtDestino2.Enabled = True
    Else
            framePropuesta.Enabled = False
            txtFechaVisita.Enabled = False
            txtEntornoFamiliar2.Enabled = False
            txtGiroUbicacion2.Enabled = False
            txtExperiencia2.Enabled = False
            txtFormalidadNegocio2.Enabled = False
            txtColaterales2.Enabled = False
            txtDestino2.Enabled = False
    End If
    '*****->Fin No Refinanciados
    
End Function
Private Function Mantenimiento()
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Dim lnFila As Integer
        If fnTipoPermiso = 3 Then
            gsOpeCod = gCredMantenimientoEvaluacionCred
        Else
            'gsOpeCod = gCredVerificacionEvaluacionCred
        End If
        
        'Para Botones *****
        If Not fbBloqueaTodo Then
            cmdInformeVisita.Enabled = False
            cmdVerCar.Enabled = False
            cmdImprimir.Enabled = False
        End If
        
        'Ver Ratios *****
        If fnEstado > 2000 Then
            SSTabRatios.Visible = True
        Else
            SSTabRatios.Visible = False
            cmdInformeVisita.Enabled = False
            cmdVerCar.Enabled = False
            cmdImprimir.Enabled = False
        End If
        
        'Ratios/ Indicadores
        txtCapacidadNeta.Enabled = False
        txtEndeudamiento.Enabled = False
        txtIngresoNeto.Enabled = False
        txtExcedenteMensual.Enabled = False
        
        'si el cliente es nuevo-> referido obligatorio
         'If fnColocCondi = 1 Then 'LUCV2017115, Seg�n correo: RUSI
        If Not fbTieneReferido6Meses Then
            frameReferido.Enabled = True
            feReferidos.Enabled = True
            cmdAgregarRef.Enabled = True
            cmdQuitar2.Enabled = True
            txtComentario.Enabled = True 'Comentarios
            frameComentario.Enabled = True
        Else
            frameReferido.Enabled = False
            feReferidos.Enabled = False
            cmdAgregarRef.Enabled = False
            cmdQuitar2.Enabled = False
            txtComentario.Enabled = False 'Comentarios
            frameComentario.Enabled = False
        End If
        
        'Ratios: Aceptable / Critico ->*****
         If Not (rsAceptableCritico.EOF Or rsAceptableCritico.BOF) Then
            If rsAceptableCritico!nCapPag = 1 Then 'Capacidad Pago
                Me.lblCapaAceptable.Caption = "Aceptable"
                Me.lblCapaAceptable.ForeColor = &H8000&
            Else
                Me.lblCapaAceptable.Caption = "Cr�tico"
                Me.lblCapaAceptable.ForeColor = vbRed
            End If
            
            If rsAceptableCritico!nEndeud = 1 Then 'Endeudamiento Pat.
                Me.lblEndeAceptable.Caption = "Aceptable"
                Me.lblEndeAceptable.ForeColor = &H8000&
            Else
                Me.lblEndeAceptable.Caption = "Cr�tico"
                Me.lblEndeAceptable.ForeColor = vbRed
            End If
        Else
            Me.lblCapaAceptable.Visible = False
            Me.lblEndeAceptable.Visible = False
        End If
            'Fin Ratios <-****
            
            '*****->No Refinanciados (Propuesta Credito)
        If fnColocCondi <> 4 Then
                txtFechaVisita.Enabled = True
                txtEntornoFamiliar2.Enabled = True
                txtGiroUbicacion2.Enabled = True
                txtExperiencia2.Enabled = True
                txtFormalidadNegocio2.Enabled = True
                txtColaterales2.Enabled = True
                txtDestino2.Enabled = True
        Else
                framePropuesta.Enabled = False
                txtFechaVisita.Enabled = False
                txtEntornoFamiliar2.Enabled = False
                txtGiroUbicacion2.Enabled = False
                txtExperiencia2.Enabled = False
                txtFormalidadNegocio2.Enabled = False
                txtColaterales2.Enabled = False
                txtDestino2.Enabled = False
        End If
    '*****->Fin No Refinanciados

        'LUCV20160626, Para CARGAR CABECERA->**********
        Set rsDCredito = oDCOMFormatosEval.RecuperaSolicitudDatoBasicosEval(sCtaCod) ' Datos Basicos del Credito Solicitado
        ActXCodCta.NroCuenta = sCtaCod
        txtGiroNeg.Text = rsCredEval!cActividad
        txtNombreCliente.Text = fsCliente
        spnExpEmpAnio.valor = rsCredEval!nExpEmpAnio
        spnExpEmpMes.valor = rsCredEval!nExpEmpMes
        spnTiempoLocalAnio.valor = rsCredEval!nTmpoLocalAnio
        spnTiempoLocalMes.valor = rsCredEval!nTmpoLocalMes
        OptCondLocal(rsCredEval!nCondiLocal).value = 1
        txtCondLocalOtros.Text = rsCredEval!cCondiLocalOtro
        txtExposicionCredito.Text = Format(rsCredEval!nExposiCred, "#,##0.00")
        txtFechaEvaluacion.Text = Format(rsCredEval!dFecEval, "dd/mm/yyyy")
        txtUltEndeuda.Text = Format(rsCredEval!nUltEndeSBS, "#,##0.00")
        txtFecUltEndeuda.Text = Format(rsCredEval!dUltEndeuSBS, "dd/mm/yyyy")
        txtComentario.Text = Trim(rsCredEval!cComentario)
        
        txtIngresoNegocio.Text = Format(rsDatVentaCosto!nIngNegocio, "#,##0.00")
        txtEgresoNegocio.Text = Format(rsDatVentaCosto!nEgrVenta, "#,##0.00")
        txtMargenBruto.Text = Format(rsDatVentaCosto!nMargBruto, "#,##0.00")
        
        'LUCV20160626, Para CARGAR PROPUESTA->**********
        If fnColocCondi <> 4 Then
            txtFechaVisita.Text = Format(rsPropuesta!dFecVisita, "dd/mm/yyyy")
            txtEntornoFamiliar2.Text = Trim(rsPropuesta!cEntornoFami)
            txtGiroUbicacion2.Text = Trim(rsPropuesta!cGiroUbica)
            txtExperiencia2.Text = Trim(rsPropuesta!cExpeCrediticia)
            txtFormalidadNegocio2.Text = Trim(rsPropuesta!cFormalNegocio)
            txtColaterales2.Text = Trim(rsPropuesta!cColateGarantia)
            txtDestino2.Text = Trim(rsPropuesta!cDestino)
        End If
        'LUCV20160626, Para la CARGAR FLEX - Mantenimiento **********->
        
        'Call FormatearGrillas(feGastosNegocio2)
        Call LimpiaFlex(feGastosNegocio)
            Do While Not rsDatGastoNeg.EOF
                feGastosNegocio.AdicionaFila
                lnFila = feGastosNegocio.row
                feGastosNegocio.TextMatrix(lnFila, 1) = rsDatGastoNeg!nConsValor
                feGastosNegocio.TextMatrix(lnFila, 2) = rsDatGastoNeg!cConsDescripcion
                feGastosNegocio.TextMatrix(lnFila, 3) = Format(rsDatGastoNeg!nMonto, "#,##0.00")
                
                Select Case CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 1))
                    Case gCodCuotaIfiGastoNego
                        'Me.feGastosNegocio.CellBackColor = &HC0FFFF
                        Me.feGastosNegocio.BackColorRow &HC0FFFF, True
                        Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
                        Me.feGastosNegocio.ForeColorRow vbBlack, True
                    Case gCodCuotaCmac
                        Me.feGastosNegocio.ColumnasAEditar = "X-X-X-X-X"
                        Me.feGastosNegocio.ForeColorRow vbBlack, True
                    Case Else
                        Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
                End Select
                rsDatGastoNeg.MoveNext
            Loop
        rsDatGastoNeg.Close
        Set rsDatGastoNeg = Nothing
        
        'Call FormatearGrillas(feGastosFamiliares2)
        Call LimpiaFlex(feGastosFamiliares)
            Do While Not rsDatGastoFam.EOF
                feGastosFamiliares.AdicionaFila
                lnFila = feGastosFamiliares.row
                feGastosFamiliares.TextMatrix(lnFila, 1) = rsDatGastoFam!nConsValor
                feGastosFamiliares.TextMatrix(lnFila, 2) = rsDatGastoFam!cConsDescripcion
                feGastosFamiliares.TextMatrix(lnFila, 3) = Format(rsDatGastoFam!nMonto, "#,##0.00")
                     
                Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1))
                    Case gCodCuotaIfiGastoFami
                       'Me.feGastosFamiliares.CellBackColor = &HC0FFFF
                       Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
                       Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
                       Me.feGastosFamiliares.ForeColorRow vbBlack, True
                    Case gCodDeudaLCNUGastoFami
                       Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
                       Me.feGastosFamiliares.ForeColorRow vbBlack, True
                    Case Else
                       Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
                End Select
                rsDatGastoFam.MoveNext
            Loop
        rsDatGastoFam.Close
        Set rsDatGastoFam = Nothing
        
        'Call FormatearGrillas(feOtrosIngresos2)
        Call LimpiaFlex(feOtrosIngresos)
            Do While Not rsDatOtrosIng.EOF
                feOtrosIngresos.AdicionaFila
                lnFila = feOtrosIngresos.row
                feOtrosIngresos.TextMatrix(lnFila, 1) = rsDatOtrosIng!nConsValor
                feOtrosIngresos.TextMatrix(lnFila, 2) = rsDatOtrosIng!cConsDescripcion
                feOtrosIngresos.TextMatrix(lnFila, 3) = Format(rsDatOtrosIng!nMonto, "#,##0.00")
                rsDatOtrosIng.MoveNext
            Loop
        rsDatOtrosIng.Close
        Set rsDatOtrosIng = Nothing
        
        'Call FormatearGrillas(feCuotaIfis)
        Call LimpiaFlex(frmCredFormEvalCuotasIfis.feCuotaIfis)
            Do While Not rsCuotaIFIs.EOF
                frmCredFormEvalCuotasIfis.feCuotaIfis.AdicionaFila
                lnFila = frmCredFormEvalCuotasIfis.feCuotaIfis.row
                frmCredFormEvalCuotasIfis.feCuotaIfis.TextMatrix(lnFila, 1) = rsCuotaIFIs!cDescripcion
                frmCredFormEvalCuotasIfis.feCuotaIfis.TextMatrix(lnFila, 2) = Format(rsCuotaIFIs!nMonto, "#,##0.00")
                rsCuotaIFIs.MoveNext
            Loop
        rsCuotaIFIs.Close
        Set rsCuotaIFIs = Nothing
        
        'Call FormatearGrillas(feReferidos2)
        Call LimpiaFlex(feReferidos)
            Do While Not rsDatRef.EOF
                feReferidos.AdicionaFila
                lnFila = feReferidos.row
                feReferidos.TextMatrix(lnFila, 0) = rsDatRef!nCodRef
                feReferidos.TextMatrix(lnFila, 1) = rsDatRef!cNombre
                feReferidos.TextMatrix(lnFila, 2) = rsDatRef!cDniNom
                feReferidos.TextMatrix(lnFila, 3) = rsDatRef!cTelf
                feReferidos.TextMatrix(lnFila, 4) = rsDatRef!cReferido
                feReferidos.TextMatrix(lnFila, 5) = rsDatRef!cDNIRef
                rsDatRef.MoveNext
            Loop
        rsDatRef.Close
        Set rsDatRef = Nothing
        
        'Call FormatearGrillas(feBalanceGeneral2)
        Call LimpiaFlex(feBalanceGeneral)
            Do While Not rsDatActivoPasivo.EOF
                feBalanceGeneral.AdicionaFila
                lnFila = feBalanceGeneral.row
                feBalanceGeneral.TextMatrix(lnFila, 1) = rsDatActivoPasivo!nConsCod
                'feBalanceGeneral.TextMatrix(lnFila, 2) = rsDatActivoPasivo!nConsValor'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                
            'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                If rsDatActivoPasivo!nConsValor = 201 Then
                    feBalanceGeneral.TextMatrix(lnFila, 2) = 107
                Else
                    feBalanceGeneral.TextMatrix(lnFila, 2) = rsDatActivoPasivo!nConsValor
                End If
            'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                
                feBalanceGeneral.TextMatrix(lnFila, 3) = rsDatActivoPasivo!nNumAut
                feBalanceGeneral.TextMatrix(lnFila, 4) = rsDatActivoPasivo!cConsDescripcion
                feBalanceGeneral.TextMatrix(lnFila, 5) = Format(rsDatActivoPasivo!nTotal, "#,##0.00")
                
                 Select Case feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2)
                    Case 1000, 1001
                        Me.feBalanceGeneral.BackColorRow (&H80000000)
                        Me.feBalanceGeneral.ForeColorRow vbBlack, True
                        Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-X-X"
                'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                    Case IIf((feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2) = 100 And feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 1) = 7026), 100, 0)
                        Me.feBalanceGeneral.BackColorRow &HC0FFFF, True
                        Me.feBalanceGeneral.ForeColorRow vbBlack, True
                        Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
                    Case IIf((feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2) = 200 And feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 1) = 7026), 200, 0)
                        Me.feBalanceGeneral.BackColorRow &HC0FFFF, True
                        Me.feBalanceGeneral.ForeColorRow vbBlack, True
                        Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
                'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                     Case 206
                        Me.feBalanceGeneral.ForeColorRow vbBlack, True
                        Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-X-X"
                     Case Else
                        Me.feBalanceGeneral.BackColorRow (&HFFFFFF)
                        Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
                 End Select
                rsDatActivoPasivo.MoveNext
            Loop
        rsDatActivoPasivo.Close
        Set rsDatActivoPasivo = Nothing
        'LUCV20160626, Fin Carga Flex <-**********
        
    'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        'Carga de rsDatIfiBalActCorri -> Matrix
            ReDim MatBalActCorr(rsDatIfiBalActCorri.RecordCount, 4)
            I = 0
            Do While Not rsDatIfiBalActCorri.EOF
                MatBalActCorr(I, 0) = rsDatIfiBalActCorri!nNroCuota
                MatBalActCorr(I, 1) = rsDatIfiBalActCorri!cDescripcion
                MatBalActCorr(I, 2) = Format(IIf(IsNull(rsDatIfiBalActCorri!nMonto), 0, rsDatIfiBalActCorri!nMonto), "#,##0.00")
                rsDatIfiBalActCorri.MoveNext
                  I = I + 1
            Loop
            rsDatIfiBalActCorri.Close
            Set rsDatIfiBalActCorri = Nothing
            
        'Carga de rsDatIfiBalActNoCorri -> Matrix
            ReDim MatBalActNoCorr(rsDatIfiBalActNoCorri.RecordCount, 4)
            I = 0
            Do While Not rsDatIfiBalActNoCorri.EOF
                MatBalActNoCorr(I, 0) = rsDatIfiBalActNoCorri!nNroCuota
                MatBalActNoCorr(I, 1) = rsDatIfiBalActNoCorri!cDescripcion
                MatBalActNoCorr(I, 2) = Format(IIf(IsNull(rsDatIfiBalActNoCorri!nMonto), 0, rsDatIfiBalActNoCorri!nMonto), "#,##0.00")
                rsDatIfiBalActNoCorri.MoveNext
                  I = I + 1
            Loop
            rsDatIfiBalActNoCorri.Close
            Set rsDatIfiBalActNoCorri = Nothing
        'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        
            'Carga de rsDatIfiGastoNego -> Matrix
            ReDim MatIfiGastoNego(rsDatIfiGastoNego.RecordCount, 4)
            I = 0
            Do While Not rsDatIfiGastoNego.EOF
                MatIfiGastoNego(I, 0) = rsDatIfiGastoNego!nNroCuota
                MatIfiGastoNego(I, 1) = rsDatIfiGastoNego!cDescripcion
                MatIfiGastoNego(I, 2) = Format(IIf(IsNull(rsDatIfiGastoNego!nMonto), 0, rsDatIfiGastoNego!nMonto), "#,##0.00")
                rsDatIfiGastoNego.MoveNext
                  I = I + 1
            Loop
            rsDatIfiGastoNego.Close
            Set rsDatIfiGastoNego = Nothing
    
            'Carga de rsDatIfiGastoFami -> Matrix
            ReDim MatIfiGastoFami(rsDatIfiGastoFami.RecordCount, 4)
            j = 0
            Do While Not rsDatIfiGastoFami.EOF
                MatIfiGastoFami(j, 0) = rsDatIfiGastoFami!nNroCuota
                MatIfiGastoFami(j, 1) = rsDatIfiGastoFami!cDescripcion
                MatIfiGastoFami(j, 2) = Format(IIf(IsNull(rsDatIfiGastoFami!nMonto), 0, rsDatIfiGastoFami!nMonto), "#,##0.00")
                rsDatIfiGastoFami.MoveNext
                j = j + 1
            Loop
            rsDatIfiGastoFami.Close
            Set rsDatIfiGastoFami = Nothing
            
        'LUCV20160628, Para CARGA RATIOS/INDICADORES
        txtCapacidadNeta.Text = CStr(rsDatRatioInd!nCapPagNeta * 100) & "%"
        txtEndeudamiento.Text = CStr(rsDatRatioInd!nEndeuPat * 100) & "%"
        txtIngresoNeto.Text = Format(rsDatRatioInd!nIngreNeto, "#,##0.00")
        txtExcedenteMensual.Text = Format(rsDatRatioInd!nExceMensual, "#,##0.00")
    Set rsDCredito = Nothing
End Function

Private Sub GeneraVerCar()
    Dim oCred As COMNCredito.NCOMFormatosEval
    Dim oDCredSbs As COMDCredito.DCOMFormatosEval
    Dim R As ADODB.Recordset
    Dim lcDNI, lcRUC As String
    Dim RSbs, RDatFin1, RCap As ADODB.Recordset
    
        Set oCred = New COMNCredito.NCOMFormatosEval
        Call oCred.RecuperaDatosInformeComercial(ActXCodCta.NroCuenta, R)
        Set oCred = Nothing
        
        If R.EOF And R.BOF Then
            MsgBox "No existen Datos para el Reporte...", vbInformation, "Aviso"
            Exit Sub
        End If
    
        lcDNI = Trim(R!dni_deudor)
        lcRUC = Trim(R!ruc_deudor)
        
        Set oDCredSbs = New COMDCredito.DCOMFormatosEval
        Set RSbs = oDCredSbs.RecuperaCaliSbs(lcDNI, lcRUC)
        Set RDatFin1 = oDCredSbs.RecuperaDatosFinan(ActXCodCta.NroCuenta, fnFormato)
        Set oDCredSbs = Nothing
        Call ImprimeInformeCriteriosAceptacionRiesgoFormatoEval(ActXCodCta.NroCuenta, gsNomAge, gsCodUser, R, RSbs, RDatFin1)
End Sub

Private Sub ImprimirFormatoEvaluacion()
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim rsInfVisita As ADODB.Recordset
    Dim rsMostrarCuotasIfis As ADODB.Recordset
    Dim rsMostrarCuotasIfisGF As ADODB.Recordset
    Dim rsIngresoEgreso As ADODB.Recordset
    
    Dim oDoc  As cPDF
    Dim psCtaCod As String
    Set oDoc = New cPDF
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsInfVisita = New ADODB.Recordset
    'Set rsInfVisita = oDCOMFormatosEval.RecuperarDatosInformeVisitaFormato1_6(sCtaCod)
    Set rsInfVisita = oDCOMFormatosEval.MostrarFormatoSinConvenioInfVisCabecera(sCtaCod, fnFormato)
    Dim rsRatiosIndicadores As ADODB.Recordset
    
    Set rsMostrarCuotasIfis = oDCOMFormatosEval.MostrarCuotasIfis(sCtaCod, fnFormato, 7022)
    Set rsMostrarCuotasIfisGF = oDCOMFormatosEval.MostrarCuotasIfis(sCtaCod, fnFormato, 7023)
    Set rsRatiosIndicadores = oDCOMFormatosEval.RecuperaDatosRatios(sCtaCod)
    Set rsIngresoEgreso = oDCOMFormatosEval.RecuperaDatosCredEvalVentaCosto(sCtaCod)
    
    Dim A As Currency
    Dim nFila, nFilaFin As Integer
    Dim nFila1 As Integer
    Dim vContrRatios() As Variant
    'A = 50
    
        'Creaci�n del Archivo
        oDoc.Author = gsCodUser
        oDoc.Creator = "SICMACT - Negocio"
        oDoc.Producer = "Caja Municipal de Ahorros y Cr�dito de Maynas S.A."
        oDoc.Subject = "Informe de Visita N� " & sCtaCod
        oDoc.Title = "Informe de Visita N� " & sCtaCod
        
        If Not oDoc.PDFCreate(App.Path & "\Spooler\FormatoEvaluacion_" & sCtaCod & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
            Exit Sub
        End If
    
        'Contenido
        oDoc.Fonts.Add "F1", "Courier New", TrueType, Normal, WinAnsiEncoding
        oDoc.Fonts.Add "F2", "Courier New", TrueType, Bold, WinAnsiEncoding
        oDoc.LoadImageFromFile App.Path & "\logo_cmacmaynas.bmp", "Logo"
        
    If Not (rsInfVisita.BOF Or rsInfVisita.EOF) Then
        'Tama�o de hoja A4
        oDoc.NewPage A4_Vertical
    
        'Call CabeceraImpCuadros(rsInfVisita)
            '---------- cabecera
        oDoc.WImage 45, 45, 45, 113, "Logo"
        oDoc.WTextBox 40, 60, 35, 390, UCase(rsInfVisita!cAgeDescripcion), "F2", 7.5, hLeft
    
        oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
        oDoc.WTextBox 60, 450, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
        oDoc.WTextBox 70, 450, 10, 200, "ANALISTA: " & UCase(rsInfVisita!cUser), "F1", 7.5, hLeft
    
        oDoc.WTextBox 80, 100, 10, 400, "HOJA DE EVALUACION", "F2", 10, hCenter
        oDoc.WTextBox 90, 55, 10, 300, "CODIGO CUENTA: " & Trim(rsInfVisita!cCtaCod), "F1", 7.5, hLeft
        oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(rsInfVisita!cPersCod), "F1", 7.5, hLeft
        oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(rsInfVisita!cPersNombre), "F1", 7.5, hLeft
        oDoc.WTextBox 100, 450, 10, 200, "DNI: " & Trim(rsInfVisita!cPersDni) & "   ", "F1", 7.5, hLeft
        oDoc.WTextBox 110, 450, 10, 200, "RUC: " & Trim(IIf(rsInfVisita!cPersRuc = "-", Space(11), rsInfVisita!cPersRuc)), "F1", 7.5, hLeft
    
        nFila = 110
        nFila = nFila + 10
        
        '*****-> LUCV20160913
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "VENTAS Y COSTOS", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F2", 7.5, hRight
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "Ingresos", "F1", 7.5, hjustify
        oDoc.WTextBox nFila, 140, 1, 160, Format(rsIngresoEgreso!nIngNegocio, "#,##0.00"), "F1", 7.5, hRight
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "Egresos", "F1", 7.5, hjustify
        oDoc.WTextBox nFila, 140, 1, 160, Format(rsIngresoEgreso!nEgrVenta, "#,##0.00"), "F1", 7.5, hRight
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "Margen Bruto", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140, 1, 160, Format(rsIngresoEgreso!nMargBruto, "#,##0.00"), "F2", 7.5, hRight
        nFila = nFila + 10
        '<-***** Fin LUCV20160913
        
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "GASTOS DEL NEGOCIO", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F2", 7.5, hRight
        
        A = 0
            For I = 1 To feGastosNegocio.rows - 1
                nFila = nFila + 10
                oDoc.WTextBox nFila, 55, 15, 250, feGastosNegocio.TextMatrix(I, 2), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150, 15, 150, Format(feGastosNegocio.TextMatrix(I, 3), "#,#0.00"), "F1", 7.5, hRight
                A = A + feGastosNegocio.TextMatrix(I, 3)
            Next I
        nFila = nFila + 10
        oDoc.WTextBox nFila, 140, 1, 160, "TOTAL" & Space(10) & Format(A, "#,##0.00"), "F2", 7.5, hRight
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        
        oDoc.WTextBox nFila, 55, 1, 160, "GASTO DE NEGOCIO - CUOTAS IFIS", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        A = 0
        If Not (rsMostrarCuotasIfis.BOF And rsMostrarCuotasIfis.EOF) Then
            For I = 1 To rsMostrarCuotasIfis.RecordCount
                'oDoc.WTextBox nFila, 55, 1, 160, rsMostrarCuotasIfis!nNroCuota, "F1", 7.5, hjustify
                oDoc.WTextBox nFila, 55, 1, 300, rsMostrarCuotasIfis!cDescripcion, "F1", 7.5, hjustify
                oDoc.WTextBox nFila, 140, 1, 160, Format(rsMostrarCuotasIfis!nMonto, "#,##0.00"), "F1", 7.5, hRight
                A = A + rsMostrarCuotasIfis!nMonto
                rsMostrarCuotasIfis.MoveNext
                nFila = nFila + 10
            Next I
            'nFila = nFila + 10
                oDoc.WTextBox nFila, 140, 1, 160, "TOTAL" & Space(10) & Format(A, "#,##0.00"), "F2", 7.5, hRight
         End If
        
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        '---------------------------------------------------------------------------------------------------------------------------------------------
        
        oDoc.WTextBox nFila, 55, 1, 160, "BALANCE GENERAL", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F2", 7.5, hRight
        A = 0
            For I = 1 To feBalanceGeneral.rows - 1
                nFila = nFila + 10
                oDoc.WTextBox nFila, 55, 15, 250, feBalanceGeneral.TextMatrix(I, 4), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150, 15, 150, Format(feBalanceGeneral.TextMatrix(I, 5), "#,#0.00"), "F1", 7.5, hRight
                A = A + feBalanceGeneral.TextMatrix(I, 5)
            Next I
        nFila = nFila + 10
        oDoc.WTextBox nFila, 140, 1, 160, "TOTAL" & Space(10) & Format(A, "#,##0.00"), "F2", 7.5, hRight
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10

        '---------------------------------------------------------------------------------------------------------------------------------------------
    
        oDoc.WTextBox nFila, 55, 1, 160, "GASTOS FAMILIARES", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F2", 7.5, hRight
        A = 0
            For I = 1 To feGastosFamiliares.rows - 1
                nFila = nFila + 10
                oDoc.WTextBox nFila, 55, 15, 250, feGastosFamiliares.TextMatrix(I, 2), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150, 15, 150, Format(feGastosFamiliares.TextMatrix(I, 3), "#,#0.00"), "F1", 7.5, hRight
                A = A + feGastosFamiliares.TextMatrix(I, 3)
            Next I
        nFila = nFila + 10
        oDoc.WTextBox nFila, 140, 1, 160, "TOTAL" & Space(10) & Format(A, "#,##0.00"), "F2", 7.5, hRight
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        
        oDoc.WTextBox nFila, 55, 1, 160, "GASTOS FAMILIARES  - CUOTAS IFIS", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        A = 0
        If Not (rsMostrarCuotasIfisGF.BOF And rsMostrarCuotasIfisGF.EOF) Then
            For I = 1 To rsMostrarCuotasIfisGF.RecordCount
                'oDoc.WTextBox nFila, 55, 1, 160, rsMostrarCuotasIfisGF!nNroCuota, "F1", 7.5, hjustify
                oDoc.WTextBox nFila, 55, 1, 300, rsMostrarCuotasIfisGF!cDescripcion, "F1", 7.5, hjustify
                oDoc.WTextBox nFila, 140, 1, 160, Format(rsMostrarCuotasIfisGF!nMonto, "#,##0.00"), "F1", 7.5, hRight
                A = A + rsMostrarCuotasIfisGF!nMonto
                nFila = nFila + 10
                rsMostrarCuotasIfisGF.MoveNext
            Next I
            'nFila = nFila + 10
                oDoc.WTextBox nFila, 140, 1, 160, "TOTAL" & Space(10) & Format(A, "#,##0.00"), "F2", 7.5, hRight
         End If
        
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
            
        '---------------------------------------------------------------------------------------------------------------------------------------------
        
        oDoc.WTextBox nFila, 55, 1, 160, "OTROS INGRESOS", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F2", 7.5, hRight
        A = 0
            For I = 1 To feOtrosIngresos.rows - 1
                nFila = nFila + 10
                oDoc.WTextBox nFila, 55, 15, 250, feOtrosIngresos.TextMatrix(I, 2), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150, 15, 150, Format(feOtrosIngresos.TextMatrix(I, 3), "#,#0.00"), "F1", 7.5, hRight
                A = A + feOtrosIngresos.TextMatrix(I, 3)
            Next I
        nFila = nFila + 10
        oDoc.WTextBox nFila, 140, 1, 160, "TOTAL" & Space(10) & Format(A, "#,##0.00"), "F2", 7.5, hRight
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        '---------------------------------------------------------------------------------------------------------------------------------------------
    
        If nFila >= 770 Then
            
            'Tama�o de hoja A4
            oDoc.NewPage A4_Vertical
            
            oDoc.WImage 45, 45, 45, 113, "Logo"
            oDoc.WTextBox 40, 60, 35, 390, UCase(rsInfVisita!cAgeDescripcion), "F2", 7.5, hLeft
        
            oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
            oDoc.WTextBox 60, 450, 10, 200, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
            oDoc.WTextBox 70, 450, 10, 200, "ANALISTA: " & UCase(rsInfVisita!cUser), "F1", 7.5, hLeft
        
            oDoc.WTextBox 80, 100, 10, 400, "HOJA DE EVALUACION", "F2", 10, hCenter
            oDoc.WTextBox 90, 55, 10, 300, "CODIGO CUENTA: " & Trim(rsInfVisita!cCtaCod), "F1", 7.5, hLeft
            oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(rsInfVisita!cPersCod), "F1", 7.5, hLeft
            oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(rsInfVisita!cPersNombre), "F1", 7.5, hLeft
            oDoc.WTextBox 100, 450, 10, 200, "DNI: " & Trim(rsInfVisita!cPersDni) & "   ", "F1", 7.5, hLeft
            oDoc.WTextBox 110, 450, 10, 200, "RUC: " & Trim(IIf(rsInfVisita!cPersRuc = "-", Space(11), rsInfVisita!cPersRuc)), "F1", 7.5, hLeft
                        
            nFila = 110
                                    
            nFila = nFila + 10
            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
            nFila = nFila + 10
            oDoc.WTextBox nFila, 55, 1, 160, "RATIOS E INDICADORES", "F2", 7.5, hjustify
            nFila = nFila + 10
            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
            nFila = nFila + 10
            oDoc.WTextBox nFila, 55, 1, 160, "Capacidad de Pago", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 10, 55, 1, 160, "Endeudamiento", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 20, 55, 1, 160, "Ingreso Neto", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 30, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
        
            'nFila1 = nFila - 20
            oDoc.WTextBox nFila, 150, 15, 150, CStr(rsRatiosIndicadores!nCapPagNeta * 100) & "%", "F1", 7.5, hRight  ''txtEndeudamiento.Text
            oDoc.WTextBox nFila + 10, 150, 15, 150, CStr(rsRatiosIndicadores!nEndeuPat * 100) & "%", "F1", 7.5, hRight ''txtCapacidadNeta.Text
            oDoc.WTextBox nFila + 20, 150, 15, 150, Format(rsRatiosIndicadores!nIngreNeto, "#,#0.00"), "F1", 7.5, hRight ''txtIngresoNeto.Text
            oDoc.WTextBox nFila + 30, 150, 15, 150, Format(rsRatiosIndicadores!nExceMensual, "#,#0.00"), "F1", 7.5, hRight ''txtExcedenteMensual.Text
        
            oDoc.WTextBox nFila, 320, 1, 250, "EN RELACION A SU EXCEDENTE", "F1", 7.5, hLeft
            oDoc.WTextBox nFila + 10, 320, 1, 250, "EN RELACION A SU PATRIMONIO TOTAL", "F1", 7.5, hLeft
            nFila = nFila + 40
            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
            nFila = nFila + 10
        Else
            
            oDoc.WTextBox nFila, 55, 1, 160, "RATIOS E INDICADORES", "F2", 7.5, hjustify
            nFila = nFila + 10
            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
            nFila = nFila + 10
            oDoc.WTextBox nFila, 55, 1, 160, "Capacidad de Pago", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 10, 55, 1, 160, "Endeudamiento", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 20, 55, 1, 160, "Ingreso Neto", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 30, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
        
            'nFila1 = nFila - 20
            oDoc.WTextBox nFila, 150, 15, 150, CStr(rsRatiosIndicadores!nCapPagNeta * 100) & "%", "F1", 7.5, hRight  ''txtEndeudamiento.Text
            oDoc.WTextBox nFila + 10, 150, 15, 150, CStr(rsRatiosIndicadores!nEndeuPat * 100) & "%", "F1", 7.5, hRight ''txtCapacidadNeta.Text
            oDoc.WTextBox nFila + 20, 150, 15, 150, Format(rsRatiosIndicadores!nIngreNeto, "#,#0.00"), "F1", 7.5, hRight ''txtIngresoNeto.Text
            oDoc.WTextBox nFila + 30, 150, 15, 150, Format(rsRatiosIndicadores!nExceMensual, "#,#0.00"), "F1", 7.5, hRight ''txtExcedenteMensual.Text
        
            oDoc.WTextBox nFila, 320, 1, 250, "EN RELACION A SU EXCEDENTE", "F1", 7.5, hLeft
            oDoc.WTextBox nFila + 10, 320, 1, 250, "EN RELACION A SU PATRIMONIO TOTAL", "F1", 7.5, hLeft
            nFila = nFila + 40
            oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
            nFila = nFila + 10
        End If
        
        oDoc.PDFClose
        oDoc.Show
    Else
        MsgBox "Los Datos de la propuesta del Credito no han sido Registrados Correctamente", vbInformation, "Aviso"
    End If
    Set rsInfVisita = Nothing
End Sub

