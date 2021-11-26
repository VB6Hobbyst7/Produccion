VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmAdmCredRegVisitas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Visitas a Clientes"
   ClientHeight    =   8235
   ClientLeft      =   2670
   ClientTop       =   1875
   ClientWidth     =   10455
   Icon            =   "frmAdmCredRegVisitas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   22
      Top             =   4560
      Width           =   10215
      Begin VB.TextBox txtEntrevistado 
         Height          =   285
         Left            =   5880
         MaxLength       =   100
         TabIndex        =   62
         Top             =   360
         Width           =   4095
      End
      Begin VB.TextBox txtDireccion 
         Height          =   285
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   61
         Top             =   360
         Width           =   4335
      End
      Begin VB.TextBox txtVerifGar 
         Height          =   285
         Left            =   2160
         MaxLength       =   150
         TabIndex        =   64
         Top             =   1080
         Width           =   7935
      End
      Begin VB.TextBox txtConclusionAccion 
         Height          =   285
         Left            =   1800
         MaxLength       =   150
         TabIndex        =   68
         Top             =   2760
         Width           =   8295
      End
      Begin VB.ComboBox cboVerifCred 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   720
         Width           =   7335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Opinión del Servicio"
         Height          =   615
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   9975
         Begin VB.ComboBox cboOpiCaja 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   200
            Width           =   3855
         End
         Begin VB.ComboBox cboOpiAna 
            Height          =   315
            Left            =   6000
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   200
            Width           =   3855
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "De la Caja :"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            ToolTipText     =   "(días de atraso promedio últimas 6 cuotas)"
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Del Analista :"
            Height          =   195
            Left            =   5040
            TabIndex        =   24
            ToolTipText     =   "(días de atraso promedio últimas 6 cuotas)"
            Top             =   240
            Width           =   930
         End
      End
      Begin VB.TextBox txtComentarios 
         Height          =   285
         Left            =   1800
         MaxLength       =   150
         TabIndex        =   67
         Top             =   2400
         Width           =   8295
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   300
         Left            =   240
         TabIndex        =   60
         Top             =   360
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Dirección :"
         Height          =   195
         Left            =   1800
         TabIndex        =   31
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Entrevistado :"
         Height          =   195
         Left            =   6000
         TabIndex        =   30
         ToolTipText     =   "(relación)"
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Verificación del destino del crédito :"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         ToolTipText     =   "(días de atraso promedio últimas 6 cuotas)"
         Top             =   720
         Width           =   2535
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Verificación de Garantías :"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         ToolTipText     =   "(días de atraso promedio últimas 6 cuotas)"
         Top             =   1080
         Width           =   1890
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Conclusión y Acción :"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "(días de atraso promedio últimas 6 cuotas)"
         Top             =   2760
         Width           =   1530
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Comentarios, apreciación y otros temas tratados por el entrevistador:"
         Height          =   435
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "(días de atraso promedio últimas 6 cuotas)"
         Top             =   2160
         Width           =   3375
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame3 
      Height          =   495
      Left            =   6960
      TabIndex        =   20
      Top             =   7680
      Width           =   3375
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   160
         Width           =   900
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   255
         Left            =   1440
         TabIndex        =   69
         Top             =   160
         Width           =   900
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   255
         Left            =   2400
         TabIndex        =   21
         Top             =   160
         Width           =   900
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Credito"
      Height          =   4410
      Left            =   75
      TabIndex        =   2
      Top             =   60
      Width           =   10230
      Begin VB.Frame Frame6 
         Caption         =   "Condiciones del crédito"
         ForeColor       =   &H000040C0&
         Height          =   855
         Left            =   120
         TabIndex        =   48
         Top             =   3480
         Width           =   9975
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Cuota :"
            Height          =   195
            Left            =   3720
            TabIndex        =   58
            Top             =   480
            Width           =   510
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   18
            Left            =   4440
            TabIndex        =   57
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Des. :"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   480
            Width           =   915
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   17
            Left            =   1080
            TabIndex        =   55
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Monto :"
            Height          =   195
            Left            =   480
            TabIndex        =   54
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Sald. Cap. :"
            Height          =   195
            Left            =   3360
            TabIndex        =   53
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Plazo :"
            Height          =   195
            Left            =   7080
            TabIndex        =   52
            Top             =   240
            Width           =   480
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   14
            Left            =   1080
            TabIndex        =   51
            Top             =   180
            Width           =   2055
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   15
            Left            =   4440
            TabIndex        =   50
            Top             =   180
            Width           =   2055
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   16
            Left            =   7800
            TabIndex        =   49
            Top             =   180
            Width           =   2055
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Validación de Evaluación"
         ForeColor       =   &H000040C0&
         Height          =   855
         Left            =   120
         TabIndex        =   33
         Top             =   2640
         Width           =   9975
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   13
            Left            =   7800
            TabIndex        =   47
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   10
            Left            =   7800
            TabIndex        =   46
            Top             =   180
            Width           =   2055
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   12
            Left            =   4440
            TabIndex        =   45
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   9
            Left            =   4440
            TabIndex        =   44
            Top             =   180
            Width           =   2055
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   11
            Left            =   1080
            TabIndex        =   43
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Index           =   8
            Left            =   1080
            TabIndex        =   42
            Top             =   180
            Width           =   2055
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Otros Ing. :"
            Height          =   195
            Left            =   3360
            TabIndex        =   39
            Top             =   480
            Width           =   780
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Activo Fijo :"
            Height          =   195
            Left            =   6840
            TabIndex        =   38
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Deudas :"
            Height          =   195
            Left            =   7080
            TabIndex        =   37
            Top             =   480
            Width           =   645
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Inventario :"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   480
            Width           =   795
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Utilidad Neta :"
            Height          =   195
            Left            =   3360
            TabIndex        =   35
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Ventas :"
            Height          =   195
            Left            =   360
            TabIndex        =   34
            Top             =   240
            Width           =   585
         End
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3840
         TabIndex        =   0
         Top             =   315
         Width           =   900
      End
      Begin VB.Frame FraListaCred 
         Caption         =   "&Lista Creditos"
         Height          =   960
         Left            =   4800
         TabIndex        =   3
         Top             =   150
         Width           =   2595
         Begin VB.ListBox LstCred 
            Height          =   645
            ItemData        =   "frmAdmCredRegVisitas.frx":030A
            Left            =   75
            List            =   "frmAdmCredRegVisitas.frx":030C
            TabIndex        =   1
            Top             =   225
            Width           =   2460
         End
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   420
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   741
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblcodigo 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   7
         Left            =   7920
         TabIndex        =   41
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Comportamiento de pago (días de atraso promedio ultimas 6 cuotas) :"
         Height          =   195
         Left            =   3000
         TabIndex        =   40
         Top             =   2445
         Width           =   5010
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   120
         X2              =   10100
         Y1              =   1280
         Y2              =   1280
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Giro del Negocio :"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   2085
         Width           =   1275
      End
      Begin VB.Label lblcodigo 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   6
         Left            =   1440
         TabIndex        =   18
         Top             =   2040
         Width           =   4935
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Analista :"
         Height          =   195
         Left            =   7200
         TabIndex        =   17
         Top             =   2085
         Width           =   645
      End
      Begin VB.Label lblcodigo 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   5
         Left            =   7920
         TabIndex        =   16
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label lblcodigo 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   14
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblcodigo 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   13
         Top             =   1680
         Width           =   4935
      End
      Begin VB.Label lblcodigo 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   2
         Left            =   7920
         TabIndex        =   12
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label lblcodigo 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   3
         Left            =   7920
         TabIndex        =   11
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label lblNat 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Natural :"
         Height          =   195
         Left            =   6870
         TabIndex        =   10
         Top             =   1350
         Width           =   1110
      End
      Begin VB.Label lblTrib 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Juridico :"
         Height          =   195
         Left            =   6870
         TabIndex        =   9
         Top             =   1740
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   735
         TabIndex        =   8
         Top             =   1410
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Deudor :"
         Height          =   195
         Left            =   750
         TabIndex        =   7
         Top             =   1740
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Datos del Crédito"
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
         Left            =   120
         TabIndex        =   6
         Top             =   1065
         Width           =   1485
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Estado Crédito :"
         Height          =   195
         Left            =   3345
         TabIndex        =   5
         Top             =   1410
         Width           =   1125
      End
      Begin VB.Label lblcodigo 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   4
         Left            =   4560
         TabIndex        =   4
         Top             =   1320
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmAdmCredRegVisitas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nMiVivienda As Integer
Dim nPrestamo As Currency
Dim nCalendDinamico As Integer
Dim bCuotaCom As Integer
Dim lcDNI As String, lcRUC As String
Dim lnEstCred As Long

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
Dim oCred As COMDCredito.DCOMCreditos
Dim rsCred As ADODB.Recordset
Dim rsComun As ADODB.Recordset
Dim bCargado As Boolean
Dim sEstado As String

    Set oCred = New COMDCredito.DCOMCreditos
    bCargado = oCred.CargaDatosControlCreditosAdmCred(psCtaCod, rsCred, rsComun, sEstado, lnEstCred)
    Set oCred = Nothing
    If bCargado Then
        lblcodigo(4).Caption = sEstado
        nPrestamo = IIf(IsNull(rsComun!nMontoCol), rsComun!nMontoSol, rsComun!nMontoCol)
        If nPrestamo = 0 Then
            nPrestamo = rsComun!nMontoSol
        End If
               
        lblcodigo(0) = rsComun!cPersCod
        lblcodigo(1) = PstaNombre(rsComun!cTitular)
        lblcodigo(2) = IIf(IsNull(rsComun!Dni), "", rsComun!Dni)
        lblcodigo(3) = IIf(IsNull(rsComun!Ruc), "", rsComun!Ruc)
        lblcodigo(5) = IIf(IsNull(rsComun!cCodAna), "", rsComun!cCodAna)
        lblcodigo(6) = IIf(IsNull(rsComun!CIIU), "", rsComun!CIIU)
        lblcodigo(7) = IIf(IsNull(rsComun!nAtrasoProm), "", rsComun!nAtrasoProm)
        lblcodigo(8) = IIf(IsNull(rsComun!Ventas), "", rsComun!Ventas)
        lblcodigo(9) = IIf(IsNull(rsComun!UtilNeta), "", rsComun!UtilNeta)
        lblcodigo(10) = IIf(IsNull(rsComun!ActivoFijo), "", rsComun!ActivoFijo)
        lblcodigo(11) = IIf(IsNull(rsComun!Inventario), "", rsComun!Inventario)
        lblcodigo(12) = IIf(IsNull(rsComun!OtrosIng), "", rsComun!OtrosIng)
        lblcodigo(13) = IIf(IsNull(rsComun!Deudas), "", rsComun!Deudas)
        lblcodigo(14) = IIf(IsNull(rsComun!nMontoCol), "", rsComun!nMontoCol)
        lblcodigo(15) = IIf(IsNull(rsComun!nSaldo), "", rsComun!nSaldo)
        lblcodigo(16) = IIf(IsNull(rsComun!nPlazo), "", rsComun!nPlazo)
        lblcodigo(17) = IIf(IsNull(rsComun!dVigencia), "", rsComun!dVigencia)
        lblcodigo(18) = IIf(IsNull(rsComun!nCuotas), "", rsComun!nCuotas)

        Me.txtFecha.Text = CDate(gdFecSis)
        Me.txtFecha.SetFocus

    End If

    CargaDatos = bCargado

End Function

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lblcodigo(0).Caption = ""
        lblcodigo(4).Caption = ""
        lblcodigo(1).Caption = ""
        lblcodigo(2).Caption = ""
        lblcodigo(3).Caption = ""
        If CargaDatos(ActxCta.NroCuenta) Then
'            cmdImprimir.Enabled = True
        Else
'            cmdImprimir.Enabled = False
        End If
    End If
End Sub

Private Sub cboOpiAna_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If

End Sub

Private Sub cboOpiCaja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub cboVerifCred_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If

End Sub

Private Sub cmdBuscar_Click()
Dim oCredito As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset
Dim oPers As COMDPersona.UCOMPersona
    
    LstCred.Clear
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        Set oCredito = New COMDCredito.DCOMCreditos
        Set R = oCredito.RecuperaCuentasParaAdmCred(oPers.sPersCod, 2) '2=VISITAS A CLIENTES
        Do While Not R.EOF
            LstCred.AddItem R!cCtaCod
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
        Set oCredito = Nothing
    End If
    If LstCred.ListCount = 0 Then
        MsgBox "El Cliente No Tiene Creditos en estado Vigente Normal", vbInformation, "Aviso"
    End If
    
End Sub

Private Sub cmdCancelar_Click()
    LimpiarControles
End Sub

Private Sub CmdGrabar_Click()
Dim oCred As COMDCredito.DCOMCreditos
Dim vCodCta As String
Dim lsMovNro As String
Dim lnCodVerifi As Integer, lnOpiCaja As Integer, lnOpiAna As Integer
    
    If ValidaDatos = False Then
        Exit Sub
    End If

    vCodCta = ActxCta.NroCuenta

    If MsgBox("Desea Registrar la visita al Cliente.", vbInformation + vbYesNo, "Atención") = vbYes Then
    
        lsMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        
        Set oCred = New COMDCredito.DCOMCreditos

            lnCodVerifi = CInt(Right(Trim(Me.cboVerifCred.Text), 3))
            lnOpiCaja = CInt(Right(Trim(Me.cboOpiCaja.Text), 3))
            lnOpiAna = CInt(Right(Trim(Me.cboOpiAna.Text), 3))

            Call oCred.RegistraVisitaClienteAdmCred(vCodCta, Format(CDate(Me.txtFecha), "yyyymmdd"), Trim(Me.txtDireccion), Trim(Me.txtEntrevistado), lnCodVerifi, Trim(Me.txtVerifGar), lnOpiCaja, lnOpiAna, Trim(Me.txtComentarios), Trim(Me.txtConclusionAccion), lsMovNro)
        
        Set oCred = Nothing

        MsgBox "Los datos se guardaron satisfactoriamente.", vbOKOnly, "Atención"

        LimpiarControles
    End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And ActxCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColPYMEEmp, False)
        If sCuenta <> "" Then
            ActxCta.NroCuenta = sCuenta
            ActxCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
Dim L As ListItem
    CentraForm Me
    
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
   
    Dim oCred As COMDCredito.DCOMCreditos
    Dim oCons As COMDConstantes.DCOMConstantes
    Dim rs As ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    
    Set oCons = New COMDConstantes.DCOMConstantes
        Set rs = oCons.RecuperaConstantes(9006)
    Set oCons = Nothing
    Call Llenar_Combo_con_Recordset(rs, cboVerifCred)

    Set oCons = New COMDConstantes.DCOMConstantes
        Set rs1 = oCons.RecuperaConstantes(9007)
    Set oCons = Nothing
    Call Llenar_Combo_con_Recordset(rs1, cboOpiCaja)
    
    Set oCons = New COMDConstantes.DCOMConstantes
        Set rs2 = oCons.RecuperaConstantes(9007)
    Set oCons = Nothing
    Call Llenar_Combo_con_Recordset(rs2, cboOpiAna)

End Sub

Private Sub LstCred_Click()
        If LstCred.ListCount > 0 And LstCred.ListIndex <> -1 Then
            ActxCta.NroCuenta = LstCred.Text
            ActxCta.SetFocusCuenta
        End If
End Sub

Sub LimpiarControles()
    Dim i As Integer

    ActxCta.Enabled = True
    ActxCta.NroCuenta = fgIniciaAxCuentaCF
    
    For i = 0 To 18
        Me.lblcodigo(i).Caption = ""
    Next
    txtFecha = "__/__/____"
    
    Me.txtDireccion = ""
    Me.txtEntrevistado = ""
    Me.cboVerifCred.ListIndex = -1
    Me.cboOpiCaja.ListIndex = -1
    Me.cboOpiAna.ListIndex = -1
    Me.txtVerifGar = ""
    Me.txtComentarios = ""
    Me.txtConclusionAccion = ""
    
End Sub

Function ValidaDatos() As Boolean

Dim oCred As COMDCredito.DCOMCreditos
        
        
    If Len(ActxCta.NroCuenta) < 18 Then
        MsgBox "Ingrese un crédito.", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
        
    If Len(Trim(lblcodigo(0))) = 0 Then
        MsgBox "Ingrese un crédito.", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
        
    'Verifica que el num de cred no haya sido registrado
    Set oCred = New COMDCredito.DCOMCreditos
    If oCred.BuscaRegistroVisitaClienteAdmCred(ActxCta.NroCuenta, txtFecha) Then
        MsgBox "La visita de este cliente ya fue registrado en esta Fecha.", vbInformation, "Aviso"
        ValidaDatos = False
        LimpiarControles
        Exit Function
    End If
        
    If Len(Trim(Me.txtDireccion)) = 0 Then
        MsgBox "Ingrese una dirección.", vbInformation, "Aviso"
        ValidaDatos = False
        txtDireccion.SetFocus
        Exit Function
    End If
    
    If Len(Trim(Me.txtEntrevistado)) = 0 Then
        MsgBox "Ingrese un entrevistado.", vbInformation, "Aviso"
        ValidaDatos = False
        txtEntrevistado.SetFocus
        Exit Function
    End If
    
    If Me.cboVerifCred.ListIndex = -1 Then
        MsgBox "Seleccione una verificacion de crédito.", vbInformation, "Aviso"
        ValidaDatos = False
        cboVerifCred.SetFocus
        Exit Function
    End If
    
    If Len(Trim(Me.txtVerifGar)) = 0 Then
        MsgBox "Ingrese una verificacion de garantia.", vbInformation, "Aviso"
        ValidaDatos = False
        txtVerifGar.SetFocus
        Exit Function
    End If
    
    If Me.cboOpiCaja.ListIndex = -1 Then
        MsgBox "Seleccione una opinión de la Caja.", vbInformation, "Aviso"
        ValidaDatos = False
        cboOpiCaja.SetFocus
        Exit Function
    End If
    
    If Me.cboOpiAna.ListIndex = -1 Then
        MsgBox "Seleccione una opinión del analista.", vbInformation, "Aviso"
        ValidaDatos = False
        cboOpiAna.SetFocus
        Exit Function
    End If
    
    If Len(Trim(Me.txtComentarios)) = 0 Then
        MsgBox "Ingrese un comentario.", vbInformation, "Aviso"
        ValidaDatos = False
        txtComentarios.SetFocus
        Exit Function
    End If
    
    If Len(Trim(Me.txtConclusionAccion)) = 0 Then
        MsgBox "Ingrese una conclusion.", vbInformation, "Aviso"
        ValidaDatos = False
        txtConclusionAccion.SetFocus
        Exit Function
    End If
        
    ValidaDatos = True
End Function

Private Sub txtcomentarios_KeyPress(KeyAscii As Integer)
     KeyAscii = Letras(KeyAscii)
     If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtConclusionAccion_KeyPress(KeyAscii As Integer)
     KeyAscii = Letras(KeyAscii)
     If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
     KeyAscii = Letras(KeyAscii)
     If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEntrevistado_KeyPress(KeyAscii As Integer)
     KeyAscii = Letras(KeyAscii)
     If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtVerifGar_KeyPress(KeyAscii As Integer)
     KeyAscii = Letras(KeyAscii)
     If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
