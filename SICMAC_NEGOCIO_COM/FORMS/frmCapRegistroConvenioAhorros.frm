VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCapRegistroConvenioAhorros 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10575
   Icon            =   "frmCapRegistroConvenioAhorros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraWS 
      Caption         =   "Tiene Servicio Web"
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
      Height          =   1065
      Left            =   8550
      TabIndex        =   81
      Top             =   60
      Width           =   1935
      Begin VB.OptionButton optSi 
         Caption         =   "Si"
         Height          =   285
         Left            =   680
         TabIndex        =   83
         Tag             =   "1"
         Top             =   660
         Width           =   615
      End
      Begin VB.OptionButton optNo 
         Caption         =   "No"
         Height          =   330
         Left            =   680
         TabIndex        =   82
         Tag             =   "0"
         Top             =   270
         Value           =   -1  'True
         Width           =   540
      End
   End
   Begin VB.Frame fraCorreos 
      Caption         =   "Envío Automático"
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
      Height          =   1875
      Left            =   60
      TabIndex        =   62
      Top             =   5970
      Width           =   10440
      Begin SICMACT.EditMoney txtFrecuenciaEnvio 
         Height          =   285
         Left            =   2520
         TabIndex        =   25
         Top             =   300
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtEmail3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5700
         TabIndex        =   32
         Top             =   1440
         Width           =   4335
      End
      Begin VB.TextBox txtEmail2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5700
         TabIndex        =   30
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox txtEmail1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5700
         TabIndex        =   28
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox txtContacto3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1260
         TabIndex        =   31
         Top             =   1440
         Width           =   3195
      End
      Begin VB.TextBox txtContacto2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1260
         TabIndex        =   29
         Top             =   1080
         Width           =   3195
      End
      Begin VB.TextBox txtContacto1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1260
         TabIndex        =   27
         Top             =   720
         Width           =   3195
      End
      Begin MSComCtl2.DTPicker dtProximoEnvio 
         Height          =   285
         Left            =   6200
         TabIndex        =   26
         Top             =   270
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   143261697
         CurrentDate     =   41906
      End
      Begin VB.Label Label18 
         Caption         =   "Email 3:"
         Height          =   285
         Left            =   4850
         TabIndex        =   70
         Top             =   1440
         Width           =   1125
      End
      Begin VB.Label Label17 
         Caption         =   "Email 2:"
         Height          =   285
         Left            =   4850
         TabIndex        =   69
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label Label16 
         Caption         =   "Email 1:"
         Height          =   285
         Left            =   4850
         TabIndex        =   68
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label15 
         Caption         =   "Contacto 3: "
         Height          =   285
         Left            =   120
         TabIndex        =   67
         Top             =   1470
         Width           =   1125
      End
      Begin VB.Label Label14 
         Caption         =   "Contacto 2:"
         Height          =   285
         Left            =   120
         TabIndex        =   66
         Top             =   1110
         Width           =   1125
      End
      Begin VB.Label Label13 
         Caption         =   "Contacto 1: "
         Height          =   285
         Left            =   120
         TabIndex        =   65
         Top             =   750
         Width           =   1125
      End
      Begin VB.Label Label12 
         Caption         =   "Próxima Fecha Envío:"
         Height          =   285
         Left            =   4200
         TabIndex        =   64
         Top             =   300
         Width           =   1995
      End
      Begin VB.Label Label11 
         Caption         =   "Frecuencia Envío (días):"
         Height          =   285
         Left            =   120
         TabIndex        =   63
         Top             =   300
         Width           =   2265
      End
   End
   Begin VB.CommandButton cmdBaja 
      Caption         =   "Dar de &Baja"
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
      Height          =   375
      Left            =   6510
      TabIndex        =   33
      Top             =   7920
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Frame fraCobro 
      Caption         =   "Tipo Cobro"
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
      Height          =   1065
      Left            =   6630
      TabIndex        =   52
      Top             =   60
      Width           =   1845
      Begin VB.OptionButton optConcepto 
         Caption         =   "Por concepto"
         Height          =   330
         Left            =   270
         TabIndex        =   3
         Tag             =   "0"
         Top             =   270
         Value           =   -1  'True
         Width           =   1260
      End
      Begin VB.OptionButton optOperacion 
         Caption         =   "Por operacion"
         Height          =   330
         Left            =   270
         TabIndex        =   4
         Tag             =   "1"
         Top             =   660
         Width           =   1350
      End
   End
   Begin VB.CommandButton cmdCancelar 
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
      Height          =   375
      Left            =   7950
      TabIndex        =   34
      Top             =   7920
      Width           =   1065
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
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
      Height          =   375
      Left            =   60
      TabIndex        =   36
      Top             =   7920
      Width           =   1170
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9060
      TabIndex        =   35
      Top             =   7920
      Width           =   1065
   End
   Begin VB.Frame fraTipoConvenio 
      Caption         =   "Tipo Convenio"
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
      Height          =   765
      Left            =   6300
      TabIndex        =   51
      Top             =   1830
      Width           =   4185
      Begin VB.ComboBox cbTipoConvenio 
         Height          =   315
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   300
         Width           =   3465
      End
   End
   Begin VB.Frame fraDescripcion 
      Caption         =   "Descripción"
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
      Height          =   705
      Left            =   60
      TabIndex        =   50
      Top             =   5220
      Width           =   10440
      Begin VB.TextBox txtServicioVoucher 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6400
         MaxLength       =   7
         TabIndex        =   84
         Top             =   240
         Width           =   3435
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1410
         TabIndex        =   24
         Top             =   240
         Width           =   3825
      End
      Begin VB.Label Label22 
         Caption         =   "Servicio:"
         Height          =   285
         Left            =   5630
         TabIndex        =   77
         Top             =   270
         Width           =   675
      End
      Begin VB.Label Label8 
         Caption         =   "Descripción:"
         Height          =   285
         Left            =   270
         TabIndex        =   44
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.Frame fraDatosConvenio 
      Caption         =   "Datos Generales del Convenio"
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
      Height          =   1005
      Left            =   60
      TabIndex        =   49
      Top             =   4170
      Width           =   10440
      Begin VB.CheckBox ckEsUnicode 
         Height          =   285
         Left            =   8800
         TabIndex        =   79
         Top             =   570
         Width           =   285
      End
      Begin VB.TextBox txtMora 
         Height          =   315
         Left            =   7360
         TabIndex        =   74
         Top             =   560
         Width           =   1035
      End
      Begin VB.ComboBox cbMora 
         Height          =   315
         Left            =   4290
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   560
         Width           =   2685
      End
      Begin VB.TextBox txtConvenioCompartido 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   2420
         TabIndex        =   23
         Top             =   570
         Width           =   1665
      End
      Begin VB.TextBox txtNombreConvenio 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   570
         Width           =   2085
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "Soporta ""Ñ"""
         Height          =   225
         Left            =   8470
         TabIndex        =   78
         Top             =   300
         Width           =   975
      End
      Begin VB.Label lblMoneda 
         Caption         =   "%"
         Height          =   225
         Index           =   6
         Left            =   6980
         TabIndex        =   76
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Label21 
         Caption         =   "Mora"
         Height          =   225
         Left            =   7420
         TabIndex        =   75
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label20 
         Caption         =   "Tipo de Mora"
         Height          =   225
         Left            =   4290
         TabIndex        =   72
         Top             =   330
         Width           =   1305
      End
      Begin VB.Label Label19 
         Caption         =   "Codigo Mype"
         Height          =   255
         Left            =   2420
         TabIndex        =   71
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label Label7 
         Caption         =   "Nombre del Convenio"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   330
         Width           =   1665
      End
   End
   Begin VB.Frame fraCuenta 
      Caption         =   "Cuenta de Recaudo del Convenio"
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
      Height          =   1485
      Left            =   60
      TabIndex        =   48
      Top             =   2640
      Width           =   10440
      Begin VB.CheckBox ckResumen 
         Caption         =   "Imprime Resumen Web"
         Height          =   375
         Left            =   8520
         TabIndex        =   80
         Top             =   360
         Width           =   1575
      End
      Begin VB.Frame fraMontosPago 
         Caption         =   "Montos de Pago"
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
         Left            =   4680
         TabIndex        =   59
         Top             =   120
         Width           =   3525
         Begin VB.TextBox txtPagoMax 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2280
            TabIndex        =   19
            Top             =   300
            Width           =   855
         End
         Begin VB.TextBox txtPagoMin 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   780
            TabIndex        =   18
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "Max"
            Height          =   255
            Left            =   1770
            TabIndex        =   61
            Top             =   330
            Width           =   435
         End
         Begin VB.Label Label9 
            Caption         =   "Min."
            Height          =   255
            Left            =   270
            TabIndex        =   60
            Top             =   330
            Width           =   435
         End
      End
      Begin VB.CommandButton cmdBuscarCuenta 
         Caption         =   "..."
         Height          =   315
         Left            =   3870
         TabIndex        =   17
         Top             =   480
         Width           =   465
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   435
         Left            =   210
         TabIndex        =   16
         Top             =   420
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   767
         Texto           =   "Cuenta Nº"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label txtTitular 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3000
         TabIndex        =   21
         Top             =   1050
         Width           =   5220
      End
      Begin VB.Label txtMoneda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """S/."" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   2
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   990
         TabIndex        =   20
         Top             =   1050
         Width           =   1155
      End
      Begin VB.Label Label6 
         Caption         =   "Titular:"
         Height          =   240
         Left            =   2280
         TabIndex        =   42
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label5 
         Caption         =   "Moneda:"
         Height          =   210
         Left            =   195
         TabIndex        =   41
         Top             =   1080
         Width           =   840
      End
   End
   Begin VB.Frame fraDistribucion 
      Caption         =   "Distribución de la Comisión"
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
      Height          =   765
      Left            =   60
      TabIndex        =   47
      Top             =   1830
      Width           =   6075
      Begin SICMACT.EditMoney txtCliente 
         Height          =   285
         Left            =   4890
         TabIndex        =   14
         Top             =   300
         Width           =   750
         _extentx        =   1323
         _extenty        =   582
         font            =   "frmCapRegistroConvenioAhorros.frx":030A
         text            =   "0"
      End
      Begin SICMACT.EditMoney txtEmpresa 
         Height          =   285
         Left            =   3150
         TabIndex        =   13
         Top             =   300
         Width           =   750
         _extentx        =   1323
         _extenty        =   582
         font            =   "frmCapRegistroConvenioAhorros.frx":0336
         text            =   "0"
      End
      Begin VB.OptionButton optPorcentajeDistribucion 
         Caption         =   "Porcentaje"
         Height          =   315
         Left            =   750
         TabIndex        =   12
         Tag             =   "1"
         Top             =   315
         Width           =   1305
      End
      Begin VB.OptionButton optFijoDistribucion 
         Caption         =   "Fijo"
         Height          =   315
         Left            =   90
         TabIndex        =   11
         Tag             =   "0"
         Top             =   315
         Width           =   645
      End
      Begin VB.Label lblMoneda 
         Caption         =   "S/."
         Height          =   225
         Index           =   5
         Left            =   4620
         TabIndex        =   58
         Top             =   360
         Width           =   330
      End
      Begin VB.Label lblMoneda 
         Caption         =   "S/."
         Height          =   225
         Index           =   4
         Left            =   2910
         TabIndex        =   57
         Top             =   360
         Width           =   330
      End
      Begin VB.Label Label4 
         Caption         =   "Cliente:"
         Height          =   225
         Left            =   3960
         TabIndex        =   40
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label2 
         Caption         =   "Empresa:"
         Height          =   210
         Left            =   2100
         TabIndex        =   39
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.Frame fraComision 
      Caption         =   "Cálculo de la Comisión"
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
      Height          =   615
      Left            =   60
      TabIndex        =   46
      Top             =   1170
      Width           =   10440
      Begin SICMACT.EditMoney txtMaximo 
         Height          =   285
         Left            =   7760
         TabIndex        =   10
         Top             =   210
         Width           =   750
         _extentx        =   1323
         _extenty        =   582
         font            =   "frmCapRegistroConvenioAhorros.frx":0362
         text            =   "0"
      End
      Begin SICMACT.EditMoney txtMinimo 
         Height          =   285
         Left            =   5800
         TabIndex        =   9
         Top             =   210
         Width           =   750
         _extentx        =   1323
         _extenty        =   582
         font            =   "frmCapRegistroConvenioAhorros.frx":038E
         text            =   "0"
      End
      Begin SICMACT.EditMoney txtPorcentaje 
         Height          =   285
         Left            =   3660
         TabIndex        =   8
         Top             =   225
         Width           =   750
         _extentx        =   1323
         _extenty        =   635
         font            =   "frmCapRegistroConvenioAhorros.frx":03BA
         text            =   "0"
      End
      Begin SICMACT.EditMoney txtFijo 
         Height          =   285
         Left            =   1260
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   225
         Width           =   750
         _extentx        =   1323
         _extenty        =   582
         font            =   "frmCapRegistroConvenioAhorros.frx":03E6
         text            =   "0"
      End
      Begin VB.OptionButton optPorcentajeCalculo 
         Caption         =   "Porcentaje"
         Height          =   330
         Left            =   2160
         TabIndex        =   7
         Tag             =   "1"
         Top             =   210
         Width           =   1305
      End
      Begin VB.OptionButton optFijoCalculo 
         Caption         =   "Fijo"
         Height          =   330
         Left            =   210
         TabIndex        =   5
         Tag             =   "0"
         Top             =   210
         Width           =   660
      End
      Begin VB.Label lblMoneda 
         Caption         =   "%"
         Height          =   225
         Index           =   3
         Left            =   3480
         TabIndex        =   56
         Top             =   285
         Width           =   225
      End
      Begin VB.Label lblMoneda 
         Caption         =   "S/."
         Height          =   225
         Index           =   2
         Left            =   7310
         TabIndex        =   55
         Top             =   240
         Width           =   330
      End
      Begin VB.Label lblMoneda 
         Caption         =   "S/."
         Height          =   225
         Index           =   1
         Left            =   930
         TabIndex        =   54
         Top             =   285
         Width           =   330
      End
      Begin VB.Label lblMoneda 
         Caption         =   "S/."
         Height          =   225
         Index           =   0
         Left            =   5440
         TabIndex        =   53
         Top             =   270
         Width           =   330
      End
      Begin VB.Label Label3 
         Caption         =   "Max."
         Height          =   240
         Left            =   6800
         TabIndex        =   38
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Min."
         Height          =   255
         Left            =   4960
         TabIndex        =   37
         Top             =   270
         Width           =   435
      End
   End
   Begin VB.Frame fraConvenio 
      Caption         =   "Institucion Convenio"
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
      Height          =   1065
      Left            =   60
      TabIndex        =   45
      Top             =   60
      Width           =   6510
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2745
         TabIndex        =   1
         Top             =   315
         Width           =   570
      End
      Begin VB.TextBox txtBuscar 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   225
         MaxLength       =   18
         TabIndex        =   0
         Top             =   315
         Width           =   2520
      End
      Begin VB.Label txtConvenio 
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   225
         TabIndex        =   2
         Top             =   660
         Width           =   6090
      End
   End
End
Attribute VB_Name = "frmCapRegistroConvenioAhorros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************************************************************
'* NOMBRE         : "frmCapRegistroConvenioAhorros"
'* DESCRIPCION    : Formulario creado para el registro de convenios segun proyecto: "Mejora del Sistema y Automatizacion de Ahorros y Servicios"
'* CREACION       : RIRO, 20121213 10:00 AM
'************************************************************************************************************************************************

Option Explicit

Private oPersona As COMDPersona.UCOMPersona
Private nTipoFormulario As Integer ' Define el tipo de formulario
Private rsRecaudo As ADODB.Recordset
Private nMoneda As Integer

Private Function validarGrabar() As String

    Dim rsConvenio As ADODB.Recordset
    Dim oRecaudo As COMDCaptaServicios.DCOMServicioRecaudo
    Dim objValidar As COMNCaptaGenerales.NCOMCaptaMovimiento
        
    If txtBuscar.Text = "" Then
        validarGrabar = "Debe seleccionar una institución antes de completar registro"
        Exit Function
    End If
    If optFijoCalculo.value = False And optPorcentajeCalculo.value = False Then
        validarGrabar = "Debe seleccionar el cálculo de comisión antes de completar el registro"
        Exit Function
    End If
    If optPorcentajeCalculo.value Then
        If CDbl(txtMinimo.Text) = CDbl(txtMaximo.Text) Or CDbl(txtMinimo.Text) > CDbl(txtMaximo.Text) Then
            validarGrabar = "El valor mínimo no debe de ser mayor o igual al monto máximo"
            Exit Function
        End If
    End If
    If optFijoDistribucion.value = False And optPorcentajeDistribucion.value = False Then
        validarGrabar = "Debe seleccionar la distribucion de la comisión"
        Exit Function
    End If
    If CDbl(txtEmpresa.Text) < 0 Or CDbl(txtCliente.Text) < 0 Then
        validarGrabar = "Los campos de distribución de comisión no deben contener valores negativos"
        Exit Function
    End If
    If CDbl(txtPorcentaje.Text) > 100 Then
        validarGrabar = "La comisión no debe ser mayor que el 100%"
        Exit Function
    End If
    If txtCuenta.Cuenta = "" Then
        validarGrabar = "Debe seleccionar una cuenta para continuar el proceso."
        Exit Function
    End If
    If Len(Trim(txtNombreConvenio.Text)) = 0 Then
        validarGrabar = "Debe ingresar un nombre para el convenio"
        Exit Function
    End If
    If nTipoFormulario = ConvenioRecaudos.NuevoConvenio Then
    
        Set oRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
        Set rsConvenio = oRecaudo.getBuscarConvenio(, , Trim(txtNombreConvenio.Text))
        If Not rsConvenio.EOF Then
            validarGrabar = "Ya existe un convenio registrado con ese nombre"
            Exit Function
        End If
    ElseIf nTipoFormulario = ConvenioRecaudos.MantenimientoConvenio Then
        Set oRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
        Set rsConvenio = oRecaudo.getBuscarConvenio(, , Trim(txtNombreConvenio.Text))
        If Not rsConvenio.EOF Then
            If rsConvenio!cCodConvenio <> txtBuscar.Text Then
                validarGrabar = "Ya existe un convenio registrado con ese nombre"
                 Exit Function
            End If
        End If
    End If
    Set objValidar = New COMNCaptaGenerales.NCOMCaptaMovimiento
    If Not objValidar.ValidaEstadoCuenta(txtCuenta.NroCuenta, False) Then
        validarGrabar = "Cuenta NO Tiene un estado valido para la Operacion, consulte con el Asistente de Agencia."
        Set objValidar = Nothing
        Exit Function
    End If
    'Correo 01
    Dim nCorreos As Integer
    If Len(Trim(txtContacto1.Text)) > 0 And Len(Trim(txtEmail1.Text)) > 0 Then
        nCorreos = nCorreos + 1
    ElseIf Len(Trim(txtContacto1.Text)) > 0 Or Len(Trim(txtEmail1.Text)) > 0 Then
        validarGrabar = "No es posible registrar un contacto si no se registra un correo electrónico"
        Exit Function
    End If
    'Correo 02
    If Len(Trim(txtContacto2.Text)) > 0 And Len(Trim(TxtEmail2.Text)) > 0 Then
        nCorreos = nCorreos + 1
    ElseIf Len(Trim(txtContacto2.Text)) > 0 Or Len(Trim(TxtEmail2.Text)) > 0 Then
        validarGrabar = "No es posible registrar un contacto si no se registra un correo electrónico"
        Exit Function
    End If
    'Correo 03
    If Len(Trim(txtContacto3.Text)) > 0 And Len(Trim(txtEmail3.Text)) > 0 Then
        nCorreos = nCorreos + 1
    ElseIf Len(Trim(txtContacto3.Text)) > 0 Or Len(Trim(txtEmail3.Text)) > 0 Then
        validarGrabar = "No es posible registrar un contacto si no se registra un correo electrónico"
        Exit Function
    End If
    If nCorreos = 0 Then
        validarGrabar = "Debe registrar al menos un contacto con su correo respectivo"
        Exit Function
    End If
    If val(Right(cbTipoConvenio.Text, 5)) = 5 And Len(Trim(txtConvenioCompartido.Text)) < 11 Then
        validarGrabar = "El código de convenio MYPE debe tener 11 caracteres"
        Exit Function
    End If
    If val(Right(cbTipoConvenio.Text, 5)) = 5 And Len(Trim(txtServicioVoucher.Text)) < 7 Then
        validarGrabar = "La casilla 'Servicio' debe contener un valor de 7 caracteres"
        Exit Function
    End If
    validarGrabar = ""
    Set objValidar = Nothing
End Function

Private Sub MostrarDatos()
    
    On Error GoTo Error
    
    'Muestra datos cuando formulario es de registro de convenio.
    If Not oPersona Is Nothing Then
    
        txtConvenio.Caption = oPersona.sPersNombre
        txtBuscar.Text = oPersona.sPersCod
        txtBuscar.Enabled = False
        cmdBuscar.Enabled = False
        txtTitular.Caption = oPersona.sPersNombre
        optConcepto.SetFocus
    
    'Muesra datos cuando formulario es de Mantenimiento o de consulta.
    ElseIf Not rsRecaudo Is Nothing Then
    
        txtConvenio.Caption = rsRecaudo!cNombreConvenio
        txtBuscar.Text = rsRecaudo!cCodConvenio
        txtTitular.Caption = rsRecaudo!cPersNombre
        txtBuscar.Enabled = False
        cmdBuscar.Enabled = False
        
        Select Case CInt(rsRecaudo!nTipoCobro)
          Case ConvenioRecaudos.TipoCobro_porConcepto
                  optConcepto.value = True
          Case ConvenioRecaudos.TipoCobro_porOperacion
                  optOperacion.value = True
        End Select
        Select Case CBool(rsRecaudo!bTieneWS) 'CTI1 ERS027-2019
          Case ConvenioRecaudos.NoTieneWS
                  optNo.value = True
          Case ConvenioRecaudos.SiTieneWS
                  optSi.value = True
        End Select 'CTI1 ERS027-2019
        Select Case rsRecaudo!nTipoCalculo
          Case ConvenioRecaudos.TipoCalculo_fijo
                  optFijoCalculo.value = True
                  txtFijo.Text = Format$(rsRecaudo!nComision, "#,##0.00")
          Case ConvenioRecaudos.TipoCalculo_porcentaje
                  optPorcentajeCalculo.value = True
                  txtPorcentaje.Text = rsRecaudo!nComision
                  txtMinimo.Text = rsRecaudo!nMinimo
                  txtMaximo.Text = rsRecaudo!nMaximo
        End Select
        cbTipoConvenio.Clear
        iniciarCombo rsRecaudo!nTipoConvenio
        Select Case rsRecaudo!nDistribucion
          Case ConvenioRecaudos.Distribucion_fijo
                  optFijoDistribucion.value = True
          Case ConvenioRecaudos.Distribucion_porcentaje
                  optPorcentajeDistribucion.value = True
        End Select
        
'        iniciarCombo (rsRecaudo!nTipoConvenio)
        txtConvenioCompartido.Text = rsRecaudo!cConvenioCompartido
        txtCuenta.NroCuenta = rsRecaudo!cCtaCod
        nMoneda = val(Mid(txtCuenta.NroCuenta, 9, 1)) ' Obtenemos la moneda: 1 = Soles,   2 = Dolares
        txtMoneda.Caption = IIf(nMoneda = 1, "SOLES", "DOLARES")
        txtTitular.Caption = rsRecaudo!cPersNombre
        txtNombreConvenio.Text = rsRecaudo!cNombreConvenio
        txtDescripcion.Text = rsRecaudo!cDescripcion
        DoEvents
        txtEmpresa.Text = Format(rsRecaudo!nDistEmpresa, "#,##0.00")
        txtCliente.Text = Format(rsRecaudo!nDistCliente, "#,##0.00")
        txtPagoMin.Text = rsRecaudo!nPagoMin
        txtPagoMax.Text = rsRecaudo!nPagoMax
        cbMora.ListIndex = IndiceListaCombo(cbMora, rsRecaudo!nTipoMora)
        TxtMora.Text = Format(CDbl(rsRecaudo!nInteresMoratorio) * 100, "#0.000000")
        txtServicioVoucher.Text = rsRecaudo!cServicio
        ckEsUnicode.value = rsRecaudo!bunicode
        ckResumen.value = IIf(rsRecaudo!bMostrarResumen, 1, 0)
        'Parámetros de envío de correo.
        Dim sTmpC() As String
        Dim sTmpE() As String
        txtFrecuenciaEnvio.value = rsRecaudo!nFrecuenciaEnvio
        dtProximoEnvio.value = rsRecaudo!dProximaFechaEnvio
        sTmpC = Split(rsRecaudo!cContacto, "|")
        sTmpE = Split(rsRecaudo!cCorreo, "|")
        If UBound(sTmpC) >= 0 Then txtContacto1.Text = sTmpC(0)
        If UBound(sTmpC) >= 1 Then txtContacto2.Text = sTmpC(1)
        If UBound(sTmpC) >= 2 Then txtContacto3.Text = sTmpC(2)
        If UBound(sTmpE) >= 0 Then txtEmail1.Text = sTmpE(0)
        If UBound(sTmpE) >= 1 Then TxtEmail2.Text = sTmpE(1)
        If UBound(sTmpE) >= 2 Then txtEmail3.Text = sTmpE(2)
        RefrescarEtiqueta
        
        If cmdGrabar.Enabled Then cmdGrabar.SetFocus
    
    End If
    
    Exit Sub
    
Error:
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub inicia(Optional Formulario As Integer = 1)
    
    On Error GoTo Error
    
    Select Case Formulario
    
        '1
        Case NuevoConvenio
             Me.Caption = "Registro de Convenio de Ahorros"
             FRAConvenio.Caption = "Institucion Convenio"
             iniciarCombo
             iniciarComboMora
             dtProximoEnvio.value = CDate(gdFecSis)
             iniciarComboMora
        '2
        Case MantenimientoConvenio
             Me.Caption = "Mantenimiento de Convenio de Servicios"
             FRAConvenio.Caption = "Convenio"
             cmdBaja.Enabled = True
             cmdBaja.Visible = True
             cbTipoConvenio.Enabled = False
             dtProximoEnvio.value = CDate(gdFecSis)
             txtConvenioCompartido.Enabled = False
             iniciarComboMora
        '3
        Case ConsultaConvenio
             Me.Caption = "Consulta de Convenio"
             FRAConvenio.Caption = "Convenio"
             cmdBaja.Enabled = False
             cmdBaja.Visible = False
             cmdGrabar.Visible = False
             cmdGrabar.Enabled = False
             cbTipoConvenio.Enabled = False
             cmdCancelar.Left = 7875
                          
             cbTipoConvenio.Enabled = False
             DeshabilitarFrames (True)
             dtProximoEnvio.value = CDate(gdFecSis)
             txtConvenioCompartido.Enabled = False
             iniciarComboMora
                
    End Select
 
    nTipoFormulario = Formulario
    Me.Show 1
    
    Exit Sub
    
Error:
    MsgBox Err.Description, vbCritical, "Aviso"

End Sub

Private Sub limpiarForm()

optFijoCalculo.value = False
optPorcentajeCalculo.value = False
optFijoDistribucion.Enabled = True
optFijoDistribucion.value = False
optPorcentajeDistribucion.value = False
optConcepto.value = True
optNo.value = True 'CTI1 ERS027-2019

txtFijo.Text = ""
txtFijo.BackColor = &H8000000F
txtFijo.Enabled = False

txtPorcentaje.Text = ""
txtPorcentaje.BackColor = &H8000000F
txtPorcentaje.Enabled = False

txtMinimo.Text = ""
txtMinimo.BackColor = &H8000000F
txtMinimo.Enabled = False

txtMaximo.Text = ""
txtMaximo.BackColor = &H8000000F
txtMaximo.Enabled = False

txtEmpresa.Text = ""
txtEmpresa.BackColor = &H8000000F
txtEmpresa.Enabled = False

txtCliente.Text = ""
txtCliente.BackColor = &H8000000F
txtCliente.Enabled = False

txtMoneda.Caption = ""
txtTitular.Caption = ""

txtNombreConvenio.Text = ""
txtDescripcion.Text = ""

txtCuenta.Age = ""
txtCuenta.CMAC = ""
txtCuenta.Cuenta = ""
txtCuenta.Prod = ""

txtConvenio.Caption = ""
cbTipoConvenio.Clear


txtFrecuenciaEnvio.value = 0
dtProximoEnvio.value = CDate(gdFecSis)
txtContacto1.Text = ""
txtContacto2.Text = ""
txtContacto3.Text = ""
txtEmail1.Text = ""
TxtEmail2.Text = ""
txtEmail3.Text = ""
TxtMora.Text = "0.00"

Select Case nTipoFormulario
    Case ConvenioRecaudos.NuevoConvenio
         iniciarCombo
         cbMora.ListIndex = 0
    Case ConvenioRecaudos.MantenimientoConvenio
         cbTipoConvenio.ListIndex = -1
         cbMora.ListIndex = 0
End Select
ckEsUnicode.value = 0
ckResumen.value = 0
txtBuscar.Enabled = True
txtBuscar.Text = ""
txtBuscar.Text = ""
cmdBuscar.Enabled = True
txtBuscar.SetFocus
nMoneda = 1
txtConvenioCompartido.Text = ""
txtServicioVoucher.Text = ""
RefrescarEtiqueta

Set oPersona = Nothing
Set rsRecaudo = Nothing

End Sub

' RIRO20150120 ERS127 - Cargando valores al combo Mora *****************
Private Sub iniciarComboMora(Optional sMora As String = "")
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsConst As ADODB.Recordset
    On Error GoTo Error
    
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(10062, , sMora, "")
    Set clsGen = Nothing
    cbMora.Clear
    Do While Not rsConst.EOF
        cbMora.AddItem rsConst("cDescripcion") & Space(100) & rsConst("nConsValor")
        rsConst.MoveNext
    Loop
    If cbMora.ListCount = 0 Then
        cbMora.ListIndex = -1
    Else
        cbMora.ListIndex = 0
    End If
    Exit Sub
Error:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
' END RIRO ************************************************************

Private Sub iniciarCombo(Optional tipoValidacion As String = "")

    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsConst As ADODB.Recordset
    
    On Error GoTo Error
    
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(10009, , tipoValidacion, "")
    Set clsGen = Nothing
    Do While Not rsConst.EOF
        cbTipoConvenio.AddItem rsConst("cDescripcion") & Space(100) & rsConst("nConsValor")
        rsConst.MoveNext
    Loop

    If cbTipoConvenio.ListCount = 0 Then
        cbTipoConvenio.ListIndex = -1
    Else
        cbTipoConvenio.ListIndex = 0
    End If
    
    Exit Sub

Error:
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function ValidarPersona(ByVal cPersCod As String) As Boolean

    Dim oPersona As COMDPersona.DCOMRoles

    On Error GoTo Error

    Set oPersona = New COMDPersona.DCOMRoles
    ValidarPersona = oPersona.ExistePersonaRol(cPersCod, gPersTipoConvenioRecaudo, True)

    Exit Function

Error:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub cbMora_Click()
    If cbMora.ListCount > 0 Then
        If val(Right(cbMora.Text, 5)) = 3 Or val(Right(cbMora.Text, 5)) = 4 Then
            TxtMora.Enabled = True
            TxtMora.SelStart = 0
            TxtMora.SelLength = Len(TxtMora.Text)
            TxtMora.SetFocus
        Else
            TxtMora.Text = "0.00"
            TxtMora.Enabled = False
        End If
    End If
End Sub

Private Sub cbTipoConvenio_Click()
    If val(Right(cbTipoConvenio.Text, 5)) = 3 Then
        txtPagoMin.Enabled = True
        txtPagoMax.Enabled = True
        txtConvenioCompartido.Enabled = False
        txtConvenioCompartido.Text = ""
        txtConvenioCompartido.BackColor = &H8000000F
        txtPagoMin.SetFocus
    Else
        txtPagoMin.Enabled = False
        txtPagoMax.Enabled = False
        txtPagoMin.Text = "0.00"
        txtPagoMax.Text = "999999.00"
        If val(Right(cbTipoConvenio.Text, 5)) = 5 Then 'Validacion MyPe
            If Not nTipoFormulario = 2 And Not nTipoFormulario = 3 Then
                txtConvenioCompartido.Enabled = True
                txtConvenioCompartido.BackColor = &H80000005
            End If
        Else
            txtConvenioCompartido.Enabled = False
            txtConvenioCompartido.Text = ""
            txtConvenioCompartido.BackColor = &H8000000F
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub EditMoney1_Change()

  txtEmpresa.SelStart = 0
  txtEmpresa.SelLength = Len(txtEmpresa.Text)
  txtEmpresa.SetFocus
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
limpiarForm
End Sub

Private Sub optFijoCalculo_Click()

    optFijoDistribucion.Enabled = True
    optPorcentajeDistribucion.Enabled = True

    optPorcentajeDistribucion.value = False
    optFijoDistribucion.value = False

    txtFijo.Enabled = True
    txtFijo.BackColor = &H80000005

    txtPorcentaje.Text = ""
    txtPorcentaje.Enabled = False
    txtPorcentaje.BackColor = &H8000000F

    txtMinimo.Text = ""
    txtMinimo.Enabled = False
    txtMinimo.BackColor = &H8000000F

    txtMaximo.Text = ""
    txtMaximo.Enabled = False
    txtMaximo.BackColor = &H8000000F

    txtEmpresa.Text = ""
    txtEmpresa.Enabled = False
    txtEmpresa.BackColor = &H8000000F

    txtCliente.Text = ""
    txtCliente.Enabled = False
    txtCliente.BackColor = &H8000000F

    txtFijo.SetFocus

End Sub

Private Sub optPorcentajeCalculo_Click()

    optFijoDistribucion.Enabled = False
    optPorcentajeDistribucion.Enabled = True

    optPorcentajeDistribucion.value = False
    optFijoDistribucion.value = False

    txtPorcentaje.Enabled = True
    txtPorcentaje.BackColor = &H80000005

    txtMinimo.Enabled = True
    txtMinimo.BackColor = &H80000005

    txtMaximo.Enabled = True
    txtMaximo.BackColor = &H80000005

    txtFijo.Text = ""
    txtFijo.Enabled = False
    txtFijo.BackColor = &H8000000F

    txtEmpresa.Text = ""
    txtEmpresa.Enabled = False
    txtEmpresa.BackColor = &H8000000F

    txtCliente.Text = ""
    txtCliente.Enabled = False
    txtCliente.BackColor = &H8000000F

    txtPorcentaje.SetFocus

End Sub

Private Sub optFijoDistribucion_Click()

    txtEmpresa.Enabled = True
    txtEmpresa.Text = ""
    txtEmpresa.BackColor = &H80000005

    txtCliente.Enabled = True
    txtCliente.Text = ""
    txtCliente.BackColor = &H80000005
    RefrescarEtiqueta
    txtEmpresa.SetFocus

End Sub

Private Sub optPorcentajeDistribucion_Click()

    txtEmpresa.Enabled = True
    txtEmpresa.Text = ""
    txtEmpresa.BackColor = &H80000005

    txtCliente.Enabled = True
    txtCliente.Text = ""
    txtCliente.BackColor = &H80000005
    RefrescarEtiqueta
    txtEmpresa.SetFocus

End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
        
    On Error GoTo Error
        
    If KeyAscii = 13 Then
    
        If Trim(txtBuscar.Text = "") Then
            MsgBox "Debe ingresar algun valor antes de la búsqueda", vbExclamation, "Aviso"
            Exit Sub
        End If
        
        Dim rsBusqueda As ADODB.Recordset
        Dim oBuscarConvenio As COMDCaptaServicios.DCOMServicioRecaudo
        Dim oBuscaPersonas As COMDPersona.DCOMPersonas
        
        If nTipoFormulario = ConvenioRecaudos.NuevoConvenio Then
            
            Set oBuscaPersonas = New COMDPersona.DCOMPersonas
            Set rsBusqueda = oBuscaPersonas.BuscaCliente(txtBuscar.Text, BusquedaCodigo)
            
            If Not rsBusqueda.EOF Then
                
                Set oPersona = New COMDPersona.UCOMPersona
                oPersona.sPersNombre = rsBusqueda!cPersNombre
                oPersona.sPersCod = rsBusqueda!cPersCod
                
                If ValidarPersona(oPersona.sPersCod) Then
                
                    MostrarDatos
                    
                Else
                    MsgBox "El cliente seleccionado no está definido en el rol de personas como INSTITUCION CONVENIO RECAUDO", vbExclamation, "Aviso"
                
                End If
            
            End If
                    
        ElseIf nTipoFormulario = ConvenioRecaudos.MantenimientoConvenio Then
                
            Set oBuscarConvenio = New COMDCaptaServicios.DCOMServicioRecaudo
            Set rsRecaudo = oBuscarConvenio.getBuscaConvenioXCodigo(Trim(txtBuscar.Text))
            
            If Not rsRecaudo.EOF And Not rsRecaudo.BOF Then
                MostrarDatos
            End If
        
        ElseIf nTipoFormulario = ConvenioRecaudos.ConsultaConvenio Then
        
            Set oBuscarConvenio = New COMDCaptaServicios.DCOMServicioRecaudo
            Set rsRecaudo = oBuscarConvenio.getBuscaConvenioXCodigo(Trim(txtBuscar.Text))
                               
            If Not rsRecaudo.EOF Then
                DeshabilitarFrames (False)
                MostrarDatos
                DeshabilitarFrames (True)
            End If
            
        End If
    Else
    
    KeyAscii = Letras(KeyAscii)
    
    End If
        
    Exit Sub
    
Error:
    MsgBox Err.Description, vbCritical, "Aviso"

End Sub

Private Sub txtCliente_GotFocus()
    txtCliente.SelStart = 0
    txtCliente.SelLength = Len(txtEmpresa.Text)
    txtCliente.SetFocus
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cbTipoConvenio.Enabled Then cbTipoConvenio.SetFocus
    End If
    If txtCliente.Text = "." Then
        txtCliente.Text = 0
    End If
End Sub

Private Sub txtContacto1_KeyPress(KeyAscii As Integer)
KeyAscii = Letras(KeyAscii)
End Sub

Private Sub txtContacto2_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
End Sub

Private Sub txtContacto3_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdGrabar.SetFocus
    Else
        KeyCode = Letras(KeyCode)
    End If
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    Else
        KeyAscii = Letras(KeyAscii)
    End If

End Sub

Private Sub txtEmpresa_Change()

    If txtEmpresa.Text <> "." Then

        If optFijoDistribucion.value Then
            
            If txtFijo.Text - txtEmpresa.Text >= 0 Then
                txtCliente.Text = txtFijo.Text - txtEmpresa.Text
            Else
                txtCliente.Text = 0
                txtEmpresa.Text = txtFijo.Text
            End If
            
        ElseIf optPorcentajeDistribucion.value Then
            
            If 100 - txtEmpresa.Text >= 0 Then
                txtCliente.Text = 100 - txtEmpresa.Text
            Else
                txtCliente.Text = 0
                txtEmpresa.Text = 100
            End If
            
        End If
    Else
        txtEmpresa.Text = 0
    End If

End Sub

Private Sub txtcliente_change()

   If txtCliente.Text <> "." Then

        If optFijoDistribucion.value Then
            If txtFijo.Text - txtCliente.Text >= 0 Then
                txtEmpresa.Text = txtFijo.Text - txtCliente.Text
            Else
                txtEmpresa.Text = 0
                txtCliente.Text = txtFijo.Text
            End If
        ElseIf optPorcentajeDistribucion.value Then
            
            If 100 - txtCliente.Text >= 0 Then
                txtEmpresa.Text = 100 - txtCliente.Text
            Else
                txtCliente.Text = 100
                txtEmpresa.Text = 0
            End If
            
        End If
    Else
        txtCliente.Text = 0

    End If

End Sub

Private Sub txtEmpresa_GotFocus()
    txtEmpresa.SelStart = 0
    txtEmpresa.SelLength = Len(txtEmpresa.Text)
End Sub

Private Sub txtEmpresa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCliente.SetFocus
    End If
End Sub

Private Sub SeleccionarTexto(ByRef txt As TextBox)

    txt.SelStart = 0
    txt.SelLength = Len(txtFijo.Text)
    txt.SetFocus

End Sub

Private Sub txtFijo_Change()
    
    If txtFijo.Text = "." Then
        txtFijo.Text = 0
    End If
    
    If optFijoDistribucion.value Then
        txtEmpresa.Text = Format(txtFijo.Text, "#,##0.00")
    End If
    
End Sub

Private Sub txtFijo_GotFocus()
    txtFijo.SelStart = 0
    txtFijo.SelLength = Len(txtMaximo.Text)
    txtFijo.SetFocus
End Sub

Private Sub txtFijo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            optFijoDistribucion.SetFocus
            Exit Sub
    End If
    
End Sub

Private Sub txtFrecuenciaEnvio_Change()
    If Not IsNumeric(txtFrecuenciaEnvio.Text) Then
        txtMaximo.Text = txtMaximo.value
    End If
End Sub

Private Sub txtFrecuenciaEnvio_LostFocus()
    txtMaximo.Text = Format(txtMaximo.value, "#")
End Sub

Private Sub txtMaximo_Change()
    If txtMaximo.Text = "." Then
        txtMaximo.Text = 0
    End If
End Sub

Private Sub txtMaximo_GotFocus()

    txtMaximo.SelStart = 0
    txtMaximo.SelLength = Len(txtMaximo.Text)
    txtMaximo.SetFocus

End Sub


Private Sub txtMaximo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        optPorcentajeDistribucion.SetFocus
    End If
    
End Sub
Private Sub txtMinimo_Change()
    If txtMinimo.Text = "" Then
        txtMinimo.Text = 0
    End If
End Sub
Private Sub txtMinimo_GotFocus()
    txtMinimo.SelStart = 0
    txtMinimo.SelLength = Len(txtMinimo.Text)
    txtMinimo.SetFocus
End Sub
Private Sub txtMinimo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMaximo.SetFocus
    End If
End Sub
Private Sub txtMora_Change()
    If Not IsNumeric(Trim(TxtMora.Text)) Then
        MsgBox "Debe ingresar valores numéricos", vbInformation, "Aviso"
        TxtMora.Text = "0.000000"
        TxtMora.SelStart = 0
        TxtMora.SelLength = Len(TxtMora.Text)
        TxtMora.SetFocus
        Exit Sub
    End If
    If CDbl(Trim(TxtMora.Text)) < 0 Then
        MsgBox "Debe ingresar valores mayores a ""cero""", vbInformation, "Aviso"
        TxtMora.Text = "0.0000000"
        TxtMora.SelStart = 0
        TxtMora.SelLength = Len(TxtMora.Text)
        TxtMora.SetFocus
    ElseIf CDbl(Trim(TxtMora.Text)) > 200 Then
        MsgBox "Debe ingresar valores menores a 200", vbInformation, "Aviso"
        TxtMora.Text = "0.000000"
        TxtMora.SelStart = 0
        TxtMora.SelLength = Len(TxtMora.Text)
        TxtMora.SetFocus
    Else
        Exit Sub
    End If
End Sub

Private Sub txtMora_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtMora, KeyAscii, , 6)
End Sub

Private Sub txtNombreConvenio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDescripcion.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If

End Sub
Private Sub txtPagoMax_Change()
 
 If Not IsNumeric(txtPagoMax.Text) Then
    
        MsgBox "Debe ingresar valores numericos", vbInformation, "Aviso"
        txtPagoMax.Text = "0.00"
        
    Else
        If txtPagoMax.Text > 999999 Then
            txtPagoMax.Text = "999999.00"
        End If
        
    End If
    
End Sub

Private Sub txtPagoMax_GotFocus()
    With txtPagoMax
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPagoMax_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPagoMax, KeyAscii)
End Sub
 
Private Sub txtPagoMax_LostFocus()
 If txtPagoMax.Text = "" Then
        txtPagoMax.Text = "00.00"
    Else
        txtPagoMax.Text = Format(txtPagoMax.Text, "##0.00")
    End If
End Sub

Private Sub txtPagoMin_Change()
    
    If Not IsNumeric(txtPagoMin.Text) Then
    
        MsgBox "Debe ingresar valores numericos", vbInformation, "Aviso"
        txtPagoMin.Text = "0.00"
        
    Else
        If txtPagoMin.Text > 9999 Then
            txtPagoMin.Text = "9999.00"
        End If
        
    End If

End Sub

Private Sub txtPagoMin_GotFocus()
    With txtPagoMin
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPagoMin_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtPagoMin, KeyAscii)
    If KeyAscii = 13 Then
        txtPagoMax.SetFocus
    End If
End Sub

Private Sub txtPagoMin_LostFocus()
    
    If txtPagoMin.Text = "" Then
        txtPagoMin.Text = "00.00"
    Else
        txtPagoMin.Text = Format(txtPagoMin.Text, "##0.00")
    End If

End Sub

Private Sub txtPorcentaje_Change()
    If txtPorcentaje.Text = "." Then
        txtPorcentaje.Text = Format$(0, "#,##0.00")
    Else
        If txtPorcentaje.Text > 100 Then txtPorcentaje.Text = 100#
    End If
End Sub

Private Sub txtPorcentaje_GotFocus()
    txtPorcentaje.SelStart = 0
    txtPorcentaje.SelLength = Len(txtMinimo.Text)
    txtPorcentaje.SetFocus
End Sub

Private Sub txtPorcentaje_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMinimo.SetFocus
    End If
End Sub

Private Sub RefrescarEtiqueta()

    Dim sMoneda As String
    sMoneda = IIf(nMoneda = 1, "S/. ", IIf(nMoneda = 2, "$ ", ""))

    lblMoneda(0).Caption = sMoneda
    lblMoneda(1).Caption = sMoneda
    lblMoneda(2).Caption = sMoneda
    lblMoneda(4).Caption = ""
    lblMoneda(5).Caption = ""
        
    If optFijoDistribucion.value Then
        lblMoneda(4).Caption = sMoneda
        lblMoneda(5).Caption = sMoneda
    ElseIf optPorcentajeDistribucion.value Then
        lblMoneda(4).Caption = "%"
        lblMoneda(5).Caption = "%"
    End If
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Error

    txtFijo.BackColor = &H8000000F
    txtPorcentaje.BackColor = &H8000000F
    txtMinimo.BackColor = &H8000000F
    txtMaximo.BackColor = &H8000000F
    txtEmpresa.BackColor = &H8000000F
    txtCliente.BackColor = &H8000000F

    txtCuenta.EnabledCMAC = False
    txtCuenta.EnabledAge = False
    txtCuenta.EnabledProd = False
    txtCuenta.EnabledCta = False
    
    
    txtPagoMin.Text = "0.00"
    txtPagoMax.Text = "999999.00"
    txtConvenioCompartido.MaxLength = 11
    nMoneda = 1
    RefrescarEtiqueta
    
    Exit Sub
    
Error:
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub cmdBaja_Click()
        
    Dim oRecaudo As COMDCaptaServicios.DCOMServicioRecaudo
    
    On Error GoTo Error
    
    If rsRecaudo Is Nothing Then
       MsgBox "Debe seleccionar un convenio", vbExclamation + vbDefaultButton1, "Aviso"
       Exit Sub
    End If
    
    If Not MsgBox("Desea dar de baja el Convenio seleccionado?", vbYesNo + vbInformation, "Aviso") = vbYes Then
       Exit Sub
    End If
               
    Set oRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
    oRecaudo.BajaConvenioRecaudo rsRecaudo!cCodConvenio
    limpiarForm
    MsgBox "Se dio de baja el convenio seleccionado", vbInformation, "Aviso"
    
    Exit Sub
    
Error:
    MsgBox Err.Description, vbCritical, "Aviso"

End Sub

Private Sub cmdBuscar_Click()
        
    On Error GoTo Error
    
    Select Case nTipoFormulario
        Case ConvenioRecaudos.NuevoConvenio
            Set oPersona = frmBuscaPersona.Inicio
            If Not oPersona Is Nothing Then
                  If ValidarPersona(oPersona.sPersCod) Then
                      MostrarDatos
                  Else
                      limpiarForm
                      MsgBox "El cliente seleccionado no está definido en el rol de personas como INSTITUCION CONVENIO RECAUDO", vbExclamation, "Aviso"
                  End If
            Else
                    MsgBox "Usted no selecciono ninguna Empresa", vbExclamation, "Aviso"
            End If
        Case ConvenioRecaudos.MantenimientoConvenio
            Set rsRecaudo = frmBuscarConvenio.Inicio
            If Not rsRecaudo Is Nothing Then
                  If Not rsRecaudo.EOF Then
                    MostrarDatos
                  Else
                    limpiarForm
                    MsgBox "Usted no selecciono ningun convenio", vbExclamation, "Aviso"
                  End If
            Else
                  MsgBox "Usted no selecciono ningun convenio", vbExclamation, "Aviso"
            End If
        Case ConvenioRecaudos.ConsultaConvenio ' Implementando para formulario de consulta
            
            Set rsRecaudo = frmBuscarConvenio.Inicio
            If Not rsRecaudo Is Nothing Then
            
                  If Not rsRecaudo.EOF Then
                    DeshabilitarFrames (False)
                    MostrarDatos
                    DeshabilitarFrames (True)
                    
                  Else
                    DeshabilitarFrames (False)
                    limpiarForm
                    DeshabilitarFrames (True)
                    MsgBox "Usted no selecciono ningun convenio", vbExclamation, "Aviso"
                      
                  End If
            Else
                  MsgBox "Usted no selecciono ningun convenio", vbExclamation, "Aviso"
            End If
        
    End Select
    
    Exit Sub
    
Error:
    MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub DeshabilitarFrames(bValor As Boolean)
    fraCobro.Enabled = Not bValor
    FRAComision.Enabled = Not bValor
    fraCuenta.Enabled = Not bValor
    fraDatosConvenio.Enabled = Not bValor
    fraDescripcion.Enabled = Not bValor
    fraDistribucion.Enabled = Not bValor
    fraTipoConvenio.Enabled = Not bValor
    fraCorreos.Enabled = Not bValor
    fraWS.Enabled = Not bValor 'CTI1 TI-ERS027-2019
End Sub

Private Sub cmdCancelar_Click()
    
    On Error GoTo Error
    
    Dim sMensaje As String
    Select Case nTipoFormulario
        Case ConvenioRecaudos.NuevoConvenio
            sMensaje = "Desea cancelar el registro de un nuevo convenio?"
        Case ConvenioRecaudos.MantenimientoConvenio
            sMensaje = "Desea cancelar el mantenimiento del convenio"
        Case ConvenioRecaudos.ConsultaConvenio
            sMensaje = "Desea Consultar Otro Convenio"
            
    End Select
    If MsgBox(sMensaje, vbExclamation + vbYesNo + vbDefaultButton1, "Aviso") = vbYes Then
        Select Case nTipoFormulario
            Case ConvenioRecaudos.NuevoConvenio, ConvenioRecaudos.MantenimientoConvenio
                    limpiarForm
            Case ConvenioRecaudos.ConsultaConvenio
                    DeshabilitarFrames (False)
                    limpiarForm
                    DeshabilitarFrames (True)
        End Select
        Exit Sub
    End If
    Exit Sub
Error:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub CmdGrabar_Click()
    
    Dim sCodConvenio As String
    Dim sPersCod As String
    Dim nTipoCobro As Integer
    Dim nTipoCalculo As Integer
    Dim nComision As Double
    Dim nMinimo As Double
    Dim nMaximo As Double
    Dim nDistribucion As Integer
    Dim nDistEmpresa As Double
    Dim nDistCliente As Double
    Dim nTipoConvenio As Integer
    Dim sCtaCod As String
    Dim sNombreConvenio As String
    Dim sDescripcion As String
    Dim nEstado As Integer
    Dim nPagoMin As Double
    Dim nPagoMax As Double
    Dim sMensaje As String
    Dim sCorreo As String, sContacto As String
    Dim nFrecuenciaEnvio As Integer
    Dim dProximaFechaEnvio As Date
    Dim sConvenioCompartido As String 'Convenio Compartido MYPE
    Dim nTipoMora As Integer
    Dim nMontoMora As Double
    Dim bTieneWS As Boolean 'CTI1 ERS027-2019
        
    Dim oCont As COMNContabilidad.NCOMContFunciones
    Dim oRecaudo As COMDCaptaServicios.DCOMServicioRecaudo
        
    On Error GoTo Error
        
    Set oCont = New COMNContabilidad.NCOMContFunciones
    sMensaje = validarGrabar
    
    If sMensaje = "" Then
            If Len(Trim(txtContacto1.Text)) > 0 And Len(Trim(txtEmail1.Text)) > 0 Then
                sCorreo = Trim(txtEmail1.Text) & "|"
                sContacto = Trim(txtContacto1.Text) & "|"
            End If
            If Len(Trim(txtContacto2.Text)) > 0 And Len(Trim(TxtEmail2.Text)) > 0 Then
                sCorreo = sCorreo & Trim(TxtEmail2.Text) & "|"
                sContacto = sContacto & Trim(txtContacto2.Text) & "|"
            End If
            If Len(Trim(txtContacto3.Text)) > 0 And Len(Trim(txtEmail3.Text)) > 0 Then
                sCorreo = sCorreo & Trim(txtEmail3.Text) & "|"
                sContacto = sContacto & Trim(txtContacto3.Text) & "|"
            End If
            nFrecuenciaEnvio = txtFrecuenciaEnvio.value
            dProximaFechaEnvio = dtProximoEnvio.value
        Select Case nTipoFormulario
            Case ConvenioRecaudos.NuevoConvenio 'Registrando Convenio nuevo
                If Not MsgBox("Esta seguro de registrar un nuevo convenio?", vbInformation + vbYesNo + vbDefaultButton1, "Aviso") = vbYes Then
                    Exit Sub
                End If
                sPersCod = oPersona.sPersCod
                Set oRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
                nTipoCobro = IIf(optConcepto.value, ConvenioRecaudos.TipoCobro_porConcepto, _
                             ConvenioRecaudos.TipoCobro_porOperacion)
                bTieneWS = IIf(optNo.value, ConvenioRecaudos.NoTieneWS, ConvenioRecaudos.SiTieneWS) 'CTI1 ERS027-2019
                nTipoCalculo = IIf(optFijoCalculo.value, ConvenioRecaudos.TipoCalculo_fijo, _
                               ConvenioRecaudos.TipoCalculo_porcentaje)
                nDistribucion = IIf(optFijoDistribucion.value, ConvenioRecaudos.Distribucion_fijo, _
                                ConvenioRecaudos.Distribucion_porcentaje)
                Select Case nTipoCalculo
                    'Se asigno valores a la propiedad Tag de los RadioButon para hacer esta condicion
                    Case CInt(optFijoCalculo.Tag)
                            nComision = txtFijo.Text
                    Case CInt(optPorcentajeCalculo.Tag)
                            nComision = txtPorcentaje.Text
                            nMinimo = txtMinimo.Text
                            nMaximo = txtMaximo.Text
                End Select
                nDistEmpresa = CDbl(txtEmpresa.Text)
                nDistCliente = CDbl(txtCliente.Text)
                nTipoConvenio = CInt(Right(cbTipoConvenio.Text, 5))
                sCtaCod = txtCuenta.NroCuenta
                sNombreConvenio = Trim(txtNombreConvenio.Text)
                sDescripcion = Trim(txtDescripcion.Text)
                nEstado = ConvenioRecaudos.Registrado
                nPagoMin = CDbl(Trim(txtPagoMin.Text))
                nPagoMax = CDbl(Trim(txtPagoMax.Text))
                sConvenioCompartido = Trim(txtConvenioCompartido.Text)
                nTipoMora = val(Right(cbMora.Text, 5))
                nMontoMora = CDbl(Trim(TxtMora.Text)) / 100#
                
                oRecaudo.AgregaConvenioRecaudo _
                                                sPersCod, _
                                                nTipoCobro, _
                                                nTipoCalculo, _
                                                nComision, _
                                                nMinimo, _
                                                nMaximo, _
                                                nDistribucion, _
                                                nDistEmpresa, _
                                                nDistCliente, _
                                                nTipoConvenio, _
                                                sCtaCod, _
                                                sNombreConvenio, _
                                                sDescripcion, _
                                                oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), _
                                                nEstado, _
                                                nPagoMin, _
                                                nPagoMax, sCorreo, sContacto, nFrecuenciaEnvio, _
                                                dProximaFechaEnvio, _
                                                sConvenioCompartido, _
                                                nTipoMora, nMontoMora, _
                                                Trim(txtServicioVoucher.Text), ckEsUnicode.value, ckResumen.value, _
                                                bTieneWS 'CTI1 ERS027-2019
                limpiarForm
                MsgBox "Se ingreso Nuevo convenio de manera satisfactoria", vbInformation, "Aviso"
                txtBuscar.SetFocus
                
            Case ConvenioRecaudos.MantenimientoConvenio 'Mantenimiento de Convenios
            
                If Not MsgBox("Desea guardar los cambios aplicados al convenio seleccionado?", vbYesNo + vbInformation, "Aviso") = vbYes Then
                    
                    Exit Sub
                    
                End If
                
                sCodConvenio = txtBuscar.Text
                sPersCod = rsRecaudo!cPersCod
                
                Set oRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
                nTipoCobro = IIf(optConcepto.value, ConvenioRecaudos.TipoCobro_porConcepto, _
                             ConvenioRecaudos.TipoCobro_porOperacion)
                bTieneWS = IIf(optNo.value, ConvenioRecaudos.NoTieneWS, ConvenioRecaudos.SiTieneWS) 'CTI1 ERS027-2019
                nTipoCalculo = IIf(optFijoCalculo.value, ConvenioRecaudos.TipoCalculo_fijo, _
                               ConvenioRecaudos.TipoCalculo_porcentaje)
                nDistribucion = IIf(optFijoDistribucion.value, ConvenioRecaudos.Distribucion_fijo, _
                                ConvenioRecaudos.Distribucion_porcentaje)
                                
                Select Case nTipoCalculo
        
                    Case CInt(optFijoCalculo.Tag)
                        nComision = txtFijo.Text
                    Case CInt(optPorcentajeCalculo.Tag)
                        nComision = txtPorcentaje.Text
                        nMinimo = txtMinimo.Text
                        nMaximo = txtMaximo.Text
                        
                End Select
        
                nDistEmpresa = CDbl(txtEmpresa.Text)
                nDistCliente = CDbl(txtCliente.Text)
                sCtaCod = txtCuenta.NroCuenta
                sNombreConvenio = Trim(txtNombreConvenio.Text)
                sDescripcion = Trim(txtDescripcion.Text)
                nEstado = ConvenioRecaudos.Registrado
                nPagoMin = CDbl(Trim(txtPagoMin.Text))
                nPagoMax = CDbl(Trim(txtPagoMax.Text))
                sConvenioCompartido = Trim(txtConvenioCompartido.Text)
                nTipoMora = val(Right(cbMora.Text, 5))
                nMontoMora = CDbl(Trim(TxtMora.Text)) / 100#
                
                oRecaudo.EditarConvenioRecaudo _
                                                sCodConvenio, _
                                                nTipoCobro, _
                                                nTipoCalculo, _
                                                nComision, _
                                                nMinimo, _
                                                nMaximo, _
                                                nDistribucion, _
                                                nDistEmpresa, _
                                                nDistCliente, _
                                                sCtaCod, _
                                                sNombreConvenio, _
                                                sDescripcion, _
                                                oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), _
                                                nEstado, nPagoMin, nPagoMax, sCorreo, sContacto, nFrecuenciaEnvio, _
                                                dProximaFechaEnvio, sConvenioCompartido, nTipoMora, _
                                                nMontoMora, Trim(txtServicioVoucher.Text), ckEsUnicode.value, ckResumen.value, _
                                                bTieneWS 'CTI1 ERS027-2019
        
                limpiarForm
                MsgBox "Se actualizó correctamente el convenio", vbInformation, "Aviso"
                txtBuscar.SetFocus
        End Select
    Else
        MsgBox sMensaje, vbExclamation, "Aviso"
    End If
    Exit Sub
Error:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdBuscarCuenta_Click()
    
    Dim sPers As String
    Dim sCta As String
    Dim sRelac As String
    Dim sEstado As String
    Dim clsCuenta As UCapCuenta
    Dim rsProducto As ADODB.Recordset
    Dim rsPers As New ADODB.Recordset
    Dim oRecaudo As COMDCaptaServicios.DCOMServicioRecaudo
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaGenerales
    
    On Error GoTo Error
    
    Select Case nTipoFormulario
    
        Case ConvenioRecaudos.NuevoConvenio
             If oPersona Is Nothing Then
                Exit Sub
             End If
             
             sPers = oPersona.sPersCod
             txtTitular.Caption = oPersona.sPersNombre
             
        Case ConvenioRecaudos.MantenimientoConvenio
             If rsRecaudo Is Nothing Then
                Exit Sub
             End If
             sPers = rsRecaudo!cPersCod
             txtTitular.Caption = rsRecaudo!cPersNombre
             
    End Select
    
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsPers = clsCap.GetCuentasPersona(sPers, 232, True)
    Set clsCap = Nothing
    If Not rsPers.EOF Then
    
        Do While Not rsPers.EOF
            sCta = rsPers("cCtaCod")
            sRelac = rsPers("cRelacion")
            sEstado = Trim(rsPers("cEstado"))
            frmCapMantenimientoCtas.lstCuentas.AddItem sCta & Space(2) & sRelac & Space(2) & sEstado
            rsPers.MoveNext
            
        Loop
        
        Set clsCuenta = New UCapCuenta
        Set clsCuenta = frmCapMantenimientoCtas.inicia
        If clsCuenta Is Nothing Then
            Else
            If clsCuenta.sCtaCod <> "" Then

                Set oRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
                Set rsProducto = oRecaudo.getTipoProductoAhorro(clsCuenta.sCtaCod)

                If Not rsProducto.EOF Then
                    If rsProducto!valor <> 1 Then
                        MsgBox "La cuenta seleccionada debe ser de tipo: Ahorro Corriente", vbExclamation, "Aviso"
                        Exit Sub
                    End If
                End If

                txtCuenta.CMAC = gsCodCMAC
                txtCuenta.Age = Mid(clsCuenta.sCtaCod, 4, 2)
                txtCuenta.Prod = Mid(clsCuenta.sCtaCod, 6, 9)
                txtCuenta.Cuenta = Mid(clsCuenta.sCtaCod, 9, 10)
                txtMoneda.Caption = IIf(Mid(clsCuenta.sCtaCod, 9, 1) = "1", "SOLES", "DOLARES")
                txtCuenta.SetFocusCuenta
                nMoneda = IIf(Mid(clsCuenta.sCtaCod, 9, 1) = "1", 1, 2)
                RefrescarEtiqueta
                
            End If
        End If
        Set clsCuenta = Nothing
        txtNombreConvenio.SetFocus
    Else

        MsgBox "No hay cuentas disponibles", _
            vbInformation, "Aviso"
            
        txtCuenta.NroCuenta = ""

    End If
    rsPers.Close
    Set rsPers = Nothing

    Exit Sub

Error:
    MsgBox Err.Description, vbCritical, "Aviso"

End Sub

Private Sub txtServicioVoucher_KeyPress(KeyAscii As Integer)
KeyAscii = Letras(KeyAscii)
End Sub

'CTI1 ERS027-2019 begin
Private Sub optSi_Click()
    If cbTipoConvenio.ListIndex <> -1 Then
        cbTipoConvenio.ListIndex = 0
        cbTipoConvenio.Enabled = False
    End If
End Sub

Private Sub optNo_Click()
    cbTipoConvenio.Enabled = True
End Sub
'CTI1 ERS027-2019 end
