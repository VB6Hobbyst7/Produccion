VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmCredLineaCreditoConfiguracion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Tarifarios"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15300
   Icon            =   "FrmCredLineaCreditoConfiguracion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   15300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      Caption         =   "Destino de Crédito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   240
      TabIndex        =   54
      Top             =   11880
      Width           =   4215
      Begin VB.CheckBox chkTDestino 
         Caption         =   "Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   0
         Width           =   2175
      End
      Begin MSComctlLib.ListView lvDestino 
         Height          =   2505
         Left            =   120
         TabIndex        =   56
         Top             =   600
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   4419
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Agencia"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6174
         EndProperty
      End
   End
   Begin VB.CheckBox chkPlazoMes 
      Caption         =   "Plazo (Meses)"
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
      Left            =   3960
      TabIndex        =   53
      Top             =   9840
      Width           =   1575
   End
   Begin VB.Frame Frame7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8400
      TabIndex        =   47
      Top             =   9840
      Width           =   5175
      Begin VB.CheckBox chkA 
         Caption         =   "A"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   360
         Width           =   495
      End
      Begin VB.CheckBox chkB 
         Caption         =   "B"
         Height          =   255
         Left            =   1080
         TabIndex        =   51
         Top             =   360
         Width           =   495
      End
      Begin VB.CheckBox chkC 
         Caption         =   "C"
         Height          =   255
         Left            =   2280
         TabIndex        =   50
         Top             =   360
         Width           =   495
      End
      Begin VB.CheckBox chkD 
         Caption         =   "D"
         Height          =   255
         Left            =   3480
         TabIndex        =   49
         Top             =   360
         Width           =   495
      End
      Begin VB.CheckBox chkE 
         Caption         =   "E"
         Height          =   255
         Left            =   4560
         TabIndex        =   48
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame9 
      Height          =   735
      Left            =   6480
      TabIndex        =   42
      Top             =   10920
      Width           =   4815
      Begin VB.TextBox txtMinEdad 
         Height          =   375
         Left            =   840
         TabIndex        =   44
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtMaxEdad 
         Height          =   375
         Left            =   2880
         TabIndex        =   43
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Max.:"
         Height          =   255
         Left            =   2040
         TabIndex        =   46
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Min.:"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   285
         Width           =   495
      End
   End
   Begin VB.Frame Frame8 
      Height          =   735
      Left            =   3360
      TabIndex        =   38
      Top             =   10920
      Width           =   2775
      Begin VB.TextBox txtCalMinimo 
         Height          =   375
         Left            =   720
         TabIndex        =   39
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Normal"
         Height          =   255
         Left            =   1920
         TabIndex        =   41
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Mínimo"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   240
      TabIndex        =   35
      Top             =   10920
      Width           =   3015
      Begin VB.CheckBox chkFem 
         Caption         =   "Femenino"
         Height          =   375
         Left            =   1680
         TabIndex        =   37
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkMas 
         Caption         =   "Masculino"
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      Height          =   975
      Left            =   3840
      TabIndex        =   30
      Top             =   9840
      Width           =   4335
      Begin VB.TextBox txtPlazoMaximo 
         Height          =   375
         Left            =   2880
         TabIndex        =   32
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtPlazoMinimo 
         Height          =   375
         Left            =   720
         TabIndex        =   31
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Máximo"
         Height          =   255
         Left            =   2040
         TabIndex        =   34
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Mínimo"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Personeria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   27
      Top             =   9840
      Width           =   3495
      Begin VB.CheckBox chkPN 
         Caption         =   "Persona natural"
         Height          =   375
         Left            =   1920
         TabIndex        =   29
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox chkPJ 
         Caption         =   "Persona Juridica"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CheckBox chkEdad 
      Caption         =   "Edad/Experiencia"
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
      Left            =   6600
      TabIndex        =   26
      Top             =   10920
      Width           =   1935
   End
   Begin VB.CheckBox chkCaliSBS 
      Caption         =   "Calificación SBS"
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
      TabIndex        =   25
      Top             =   10905
      Width           =   1935
   End
   Begin VB.CheckBox chkGenero 
      Caption         =   "Género(PN)"
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
      Left            =   360
      TabIndex        =   24
      Top             =   10905
      Width           =   1335
   End
   Begin VB.CheckBox chkCalificacion 
      Caption         =   "Calificación Interna"
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
      Left            =   8520
      TabIndex        =   23
      Top             =   9840
      Width           =   2055
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   15055
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Configuración de Tasas en Productos Crediticios"
      TabPicture(0)   =   "FrmCredLineaCreditoConfiguracion.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCaract"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProd"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "flxCaractProd"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCerrar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdDeshabilidad"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdCancelar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdAceptar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame11"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "FrameDolares"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "FrameSoles"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkSoles"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkDolar"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cboCampana"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cboSubProducto"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdMostrar"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Left            =   7320
         TabIndex        =   20
         Top             =   800
         Width           =   1335
      End
      Begin VB.ComboBox cboSubProducto 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   840
         Width           =   3855
      End
      Begin VB.ComboBox cboCampana 
         Height          =   315
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   840
         Width           =   3015
      End
      Begin VB.CheckBox chkDolar 
         Caption         =   "Dólares"
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
         Left            =   7800
         TabIndex        =   17
         Top             =   1440
         Width           =   975
      End
      Begin VB.CheckBox chkSoles 
         Caption         =   "Soles"
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
         Left            =   480
         TabIndex        =   16
         Top             =   1440
         Width           =   975
      End
      Begin VB.Frame FrameSoles 
         Height          =   2415
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   7215
         Begin VB.CommandButton cmdMasS 
            Caption         =   "+"
            Height          =   375
            Left            =   6600
            TabIndex        =   14
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmdMenosS 
            Caption         =   "-"
            Height          =   375
            Left            =   6600
            TabIndex        =   13
            Top             =   720
            Width           =   495
         End
         Begin SICMACT.FlexEdit FESoles 
            Height          =   1935
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   6375
            _extentx        =   11245
            _extenty        =   3413
            cols0           =   9
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "N°-X-Pref.-Tipo-Desde-Hasta-T.Min-T.Max-Cursor"
            encabezadosanchos=   "350-0-0-2000-900-900-900-900-0"
            font            =   "FrmCredLineaCreditoConfiguracion.frx":0326
            font            =   "FrmCredLineaCreditoConfiguracion.frx":0352
            font            =   "FrmCredLineaCreditoConfiguracion.frx":037E
            font            =   "FrmCredLineaCreditoConfiguracion.frx":03AA
            font            =   "FrmCredLineaCreditoConfiguracion.frx":03D6
            fontfixed       =   "FrmCredLineaCreditoConfiguracion.frx":0402
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            columnasaeditar =   "X-X-X-3-4-5-6-7-X"
            listacontroles  =   "0-0-0-3-0-0-0-0-0"
            encabezadosalineacion=   "R-R-C-L-R-R-R-R-C"
            formatosedit    =   "0-0-0-0-0-0-0-0-0"
            textarray0      =   "N°"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            colwidth0       =   345
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
      End
      Begin VB.Frame FrameDolares 
         Height          =   2415
         Left            =   7560
         TabIndex        =   8
         Top             =   1440
         Width           =   7215
         Begin VB.CommandButton cmdMenosD 
            Caption         =   "-"
            Height          =   375
            Left            =   6600
            TabIndex        =   10
            Top             =   720
            Width           =   495
         End
         Begin VB.CommandButton cmdMasD 
            Caption         =   "+"
            Height          =   375
            Left            =   6600
            TabIndex        =   9
            Top             =   240
            Width           =   495
         End
         Begin SICMACT.FlexEdit FEDolares 
            Height          =   2055
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   6375
            _extentx        =   11245
            _extenty        =   3625
            cols0           =   9
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "N°-X-Pref.-Tipo-Desde-Hasta-T.Min-T.Max-Cursor"
            encabezadosanchos=   "350-0-0-2000-900-900-900-900-0"
            font            =   "FrmCredLineaCreditoConfiguracion.frx":0430
            font            =   "FrmCredLineaCreditoConfiguracion.frx":045C
            font            =   "FrmCredLineaCreditoConfiguracion.frx":0488
            font            =   "FrmCredLineaCreditoConfiguracion.frx":04B4
            font            =   "FrmCredLineaCreditoConfiguracion.frx":04E0
            fontfixed       =   "FrmCredLineaCreditoConfiguracion.frx":050C
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            columnasaeditar =   "X-X-X-3-4-5-6-7-X"
            listacontroles  =   "0-0-0-3-0-0-0-0-0"
            encabezadosalineacion=   "R-R-C-L-R-R-R-R-C"
            formatosedit    =   "0-0-0-0-0-0-0-0-0"
            textarray0      =   "N°"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            colwidth0       =   345
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Agencias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4285
         Left            =   5400
         TabIndex        =   5
         Top             =   3985
         Width           =   3975
         Begin VB.CheckBox chkTAgencias 
            Caption         =   "Todos"
            Height          =   255
            Left            =   200
            TabIndex        =   6
            Top             =   240
            Width           =   2175
         End
         Begin MSComctlLib.ListView lvAgencia 
            Height          =   3485
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   6138
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Agencia"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descripción"
               Object.Width           =   6174
            EndProperty
         End
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   9840
         TabIndex        =   4
         Top             =   4920
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Limpiar"
         Height          =   495
         Left            =   9840
         TabIndex        =   3
         Top             =   5520
         Width           =   1455
      End
      Begin VB.CommandButton cmdDeshabilidad 
         Caption         =   "Dehabilitar"
         Height          =   495
         Left            =   9840
         TabIndex        =   2
         Top             =   6120
         Width           =   1455
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   495
         Left            =   9840
         TabIndex        =   1
         Top             =   6720
         Width           =   1455
      End
      Begin SICMACT.FlexEdit flxCaractProd 
         Height          =   4215
         Left            =   240
         TabIndex        =   57
         Top             =   4080
         Width           =   4815
         _extentx        =   8493
         _extenty        =   7435
         cols0           =   3
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "n-N°-DESCRIPCIÓN"
         encabezadosanchos=   "0-450-4000"
         font            =   "FrmCredLineaCreditoConfiguracion.frx":053A
         font            =   "FrmCredLineaCreditoConfiguracion.frx":0566
         font            =   "FrmCredLineaCreditoConfiguracion.frx":0592
         font            =   "FrmCredLineaCreditoConfiguracion.frx":05BE
         font            =   "FrmCredLineaCreditoConfiguracion.frx":05EA
         fontfixed       =   "FrmCredLineaCreditoConfiguracion.frx":0616
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-X"
         listacontroles  =   "0-0-0"
         encabezadosalineacion=   "C-L-L"
         formatosedit    =   "0-0-0"
         textarray0      =   "n"
         lbeditarflex    =   -1
         lbbuscaduplicadotext=   -1
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin VB.Label lblProd 
         Caption         =   "Producto"
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
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblCaract 
         Caption         =   "Carácteristica"
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
         Left            =   4200
         TabIndex        =   21
         Top             =   600
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmCredLineaCreditoConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gbEstado As Boolean
Dim gnNumDec As Integer
Dim nHab As Integer
'*******NAGL 20190118********'
Dim nRangIniSoles As Currency
Dim nRangFinSoles As Currency
Dim nRangIniDolares As Currency
Dim nRangFinDolares As Currency
Dim pbMostrarMsg As Boolean
'*******NAGL 20190118********'

Public Event Change() 'ALPA 20140328
Public Event KeyPress(KeyAscii As Integer) 'ALPA 20140328

Private Sub Form_Load()
    Call CargaControles
    Call ActualizaControles(False)
End Sub

Private Sub CargaControles()
Dim objLinea As COMDCredito.DCOMLineaCredito
Set objLinea = New COMDCredito.DCOMLineaCredito
Dim lvItem As ListItem
Dim oAge As New DActualizaDatosArea
'psPersCodCaja = "1090100012521"
Dim objRS As ADODB.Recordset
Set objRS = New ADODB.Recordset

Set objRS = objLinea.ObtenerConstanteLineaCredito(3033)
Call LlenarCombo2(cboSubProducto, objRS)
Set objRS = Nothing
Set objRS = New ADODB.Recordset

Set objRS = objLinea.ObtenerCampanaConfguracion
Call LlenarCombo(cboCampana, objRS)
Set objRS = Nothing
Set objRS = New ADODB.Recordset

Set objRS = objLinea.ObtenerProductoCreditocioAgencia("", 0)
If objRS.EOF Then
   RSClose objRS
   MsgBox "No se definieron Agencias en el Sistema...Consultar con Sistemas", vbInformation, "Aviso"
   Exit Sub
End If
Do While Not objRS.EOF
   Set lvItem = lvAgencia.ListItems.Add(, , objRS!cCodigo)
   lvItem.SubItems(1) = objRS!cDescri
   lvItem.Checked = False
   objRS.MoveNext
Loop
RSClose objRS
pbMostrarMsg = True 'NAGL 20190128
'Set objRS = objLinea.ObtenerConstanteLineaCredito(3016)
'If objRS.EOF Then
'   RSClose objRS
'   MsgBox "No se definieron los destinos de creditos en la agencia...Consultar con Sistemas", vbInformation, "Aviso"
'   Exit Sub
'End If
'Do While Not objRS.EOF
'   Set lvItem = lvDestino.ListItems.Add(, , objRS!cCodigo)
'   lvItem.SubItems(1) = objRS!cDescri
'   lvItem.Checked = False
'   objRS.MoveNext
'Loop
'RSClose objRS
'*****Comentado by NAGL Según 20181215 Los Destinos son Config, by Créditos
End Sub

Public Sub ObtieneRangosMontos(ByVal psTpoProd As String)
Dim rs As New ADODB.Recordset
Dim oLinea As New COMDCredito.DCOMLineaCredito
Set rs = oLinea.ObtieneRangosMontosxTarifarios(psTpoProd, gdFecSis)
If Not rs.BOF And Not rs.EOF Then
   nRangIniSoles = rs!cRangoInicioSoles
   nRangFinSoles = rs!cRangoFinSoles
   nRangIniDolares = rs!cRangoInicioDolares
   nRangFinDolares = rs!cRangoFinDolares
End If
End Sub '***NAGL 20190118 ERS042-2018

Private Sub ActualizaControles(ByVal pbMostrar As Boolean)
    cmdAceptar.Enabled = pbMostrar
    cmdCancelar.Enabled = pbMostrar
    cmdDeshabilidad.Enabled = pbMostrar
    cmdCerrar.Enabled = True
End Sub

Private Sub LlenarCombo(ByRef pCombo As ComboBox, ByRef pRs As ADODB.Recordset)
'    pRs.MoveFirst
    If (pRs.BOF Or pRs.EOF) Then
    Exit Sub
    End If
    pCombo.Clear
    Do While Not pRs.EOF
        pCombo.AddItem pRs!cDescri & Space(300) & pRs!cCodigo
        pRs.MoveNext
    Loop
End Sub
Private Sub LlenarCombo2(ByRef pCombo As ComboBox, ByRef pRs As ADODB.Recordset)
'    pRs.MoveFirst
    If (pRs.BOF Or pRs.EOF) Then
    Exit Sub
    End If
    pCombo.Clear
    Do While Not pRs.EOF
        pCombo.AddItem pRs!cCodigo & "-" & pRs!cDescri & Space(300) & pRs!cCodigo
        pRs.MoveNext
    Loop
End Sub

Private Sub cmdMostrar_Click()
Dim nCantidad As Integer
Dim objLinea As COMDCredito.DCOMLineaCredito
Set objLinea = New COMDCredito.DCOMLineaCredito
Dim objRS As ADODB.Recordset
Set objRS = New ADODB.Recordset
Dim lvItem As ListItem
Dim psTpoProd As String 'NAGL 20181217

'*******NAGL 20190118*****
nRangIniSoles = 0
nRangFinSoles = 0
nRangIniDolares = 0
nRangFinDolares = 0
pbMostrarMsg = True
'*******END NAGL**********

If Trim(cboSubProducto.Text) = "" Then
    MsgBox "Debe seleccionar un producto crediticio", vbInformation, "Aviso"
    Exit Sub
End If

If Trim(cboCampana.Text) = "" Then
    MsgBox "Debe seleccionar una campaña de crédito", vbInformation, "Aviso"
    Exit Sub
End If

psTpoProd = Trim(Right(cboSubProducto.Text, 10)) 'NAGL 20181217
Call ObtieneRangosMontos(psTpoProd)   'NAGL 20190118

Set objRS = objLinea.ObtenerProductoCrediticio(Trim(Right(cboSubProducto.Text, 10)), Trim(Right(cboCampana.Text, 10)), "HabCamp")
    If Not (objRS.BOF Or objRS.EOF) Then
           pbMostrarMsg = False 'Para no mostrar el mensaje de la Restricción, ya que está apareciendo por primera vez
           chkSoles.value = objRS!bSoles
           pbMostrarMsg = False
           chkDolar.value = objRS!bDolares
'           If chkSoles.value = 1 Then
'                FrameSoles.Enabled = True
'           End If
'           If chkDolar.value = 1 Then
'                FrameDolares.Enabled = True
'           End If
           'chkPN.value = objRS!bPersoneriaN
           'chkPJ.value = objRS!bPersoneriaJ
           'chkPlazoMes.value = objRS!bPlazo
           'txtPlazoMinimo.Text = objRS!nPlazoMin
           'txtPlazoMaximo.Text = objRS!nPlazoMax
           'chkCalificacion.value = objRS!bCalificacion
           'If chkCalificacion.value = 1 Then
            'chkA.value = objRS!bCalA
            'chkB.value = objRS!bCalB
            'chkC.value = objRS!bCalC
            'chkD.value = objRS!bCalD
            'chkE.value = objRS!bCalE
           'Else
            'chkA.value = 0
            'chkB.value = 0
            'chkC.value = 0
            'chkD.value = 0
            'chkE.value = 0
           'End If
           'chkGenero.value = objRS!bGenero
           'chkMas.value = objRS!bGeneroM
           'chkFem.value = objRS!bGeneroF
           'chkCaliSBS.value = objRS!bCalificacionSBS
           'txtCalMinimo.Text = objRS!nCalificacionSBSPor
           'chkEdad.value = objRS!bEdad
           'txtMinEdad.Text = objRS!nEdadMin
           'txtMaxEdad.Text = objRS!nEdadMax
           '**Comentado by NAGL 20181215, Configuración realizada por Adm.Créditos
           pbMostrarMsg = True '20190128
           If objRS!nEstado = 1 Then
                nHab = 1
                cmdDeshabilidad.Caption = "Deshabilitar"
                cmdAceptar.Enabled = True
           Else
                nHab = 0
                cmdDeshabilidad.Caption = "Habilitar"
                cmdAceptar.Enabled = False
           End If
    Else
            nHab = 1
            cmdDeshabilidad.Caption = "Deshabilitar"
            cmdAceptar.Enabled = True
    End If
    
    nCantidad = 0
    'FrameSoles.Enabled = False
    Call LlenarGrillaSoles(nCantidad)
    
    nCantidad = 0
    'FrameDolares.Enabled = True
    Call LlenarGrillaDolares(nCantidad)
    
    Call CargaCaractGeneralesProd(psTpoProd) 'NAGL 20181219
    'Llenar Agencias
    lvAgencia.ListItems.Clear
    Set objRS = objLinea.ObtenerProductoCreditocioAgencia(Trim(Right(cboSubProducto.Text, 10)), Trim(Right(cboCampana.Text, 10)))
    If Not (objRS.BOF Or objRS.EOF) Then
    Do While Not objRS.EOF
       Set lvItem = lvAgencia.ListItems.Add(, , objRS!cCodigo)
       lvItem.SubItems(1) = objRS!cDescri
       'lvItem.SubItems(2) = IIf(objRs!nEstado = 1, True, False)
       If objRS!nEstado Then
            lvItem.Checked = True
       Else
            lvItem.Checked = False
       End If
       objRS.MoveNext
    Loop
    End If
RSClose objRS
''Tipo de Crédito - Destino
'lvDestino.ListItems.Clear
'Set objRS = objLinea.ObtenerProductoCreditocioDestino(Trim(Right(cboSubProducto.Text, 10)), Trim(Right(cboCampana.Text, 10)))
'If Not (objRS.BOF Or objRS.EOF) Then
'    Do While Not objRS.EOF
'       Set lvItem = lvDestino.ListItems.Add(, , objRS!cCodigo)
'       lvItem.SubItems(1) = objRS!cDescri
'       If objRS!nEstado Then
'            lvItem.Checked = True
'       Else
'            lvItem.Checked = False
'       End If
'       objRS.MoveNext
'    Loop
'    End If
'RSClose objRS
cboSubProducto.Enabled = False
cboCampana.Enabled = False
cmdMostrar.Enabled = False
Call ActualizaControles(True)
End Sub

Public Sub CargaCaractGeneralesProd(psTpoProd As String)
Dim rs As New ADODB.Recordset
'JOEP20190215 CPp
Dim oCredCat As COMDCredito.DCOMCatalogoProd
Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP20190215 CP
Dim i As Long
i = 1
LimpiaFlex flxCaractProd
Set rs = oCredCat.ObtieneCaractCondRequixProducto(psTpoProd, gdFecSis)
Do While Not rs.EOF
    flxCaractProd.AdicionaFila
    flxCaractProd.TextMatrix(i, 1) = rs!nOrden
    flxCaractProd.TextMatrix(i, 2) = rs!cDescripcion
    If rs!bEstado = "N" Then
       flxCaractProd.row = i
       flxCaractProd.Col = 1
       flxCaractProd.CellBackColor = &HE0E0E0
       flxCaractProd.Col = 2
       flxCaractProd.CellBackColor = &HE0E0E0
    End If
    i = i + 1
    rs.MoveNext
Loop
Set rs = Nothing
Set oCredCat = Nothing
End Sub 'NAGL Según ERS 042-2018 20181217

Private Sub LlenarGrillaDolares(ByRef nCantidad As Integer)
    Dim objRS As ADODB.Recordset
    Dim objLinea As COMDCredito.DCOMLineaCredito
    Set objLinea = New COMDCredito.DCOMLineaCredito
    FormateaFlex FEDolares
    Set objRS = Nothing
    Set objRS = New ADODB.Recordset
    If chkDolar.value = 1 Then
        Set objRS = objLinea.ObtenerProductoCrediticioTasas(Trim(Right(cboSubProducto.Text, 10)), Trim(Right(cboCampana.Text, 10)), 2)
        Do While Not objRS.EOF
            nCantidad = nCantidad + 1
            FEDolares.AdicionaFila
            FEDolares.TextMatrix(objRS.Bookmark, 0) = objRS!nOrden
            FEDolares.TextMatrix(objRS.Bookmark, 1) = "0"
            'FEDolares.TextMatrix(objRS.Bookmark, 2) = objRS!bPreferencial
            FEDolares.TextMatrix(objRS.Bookmark, 2) = "" 'FRHU 20171107 Acta 189-2017
            FEDolares.TextMatrix(objRS.Bookmark, 3) = objRS!cConsDescripcion & Space(200) & objRS!nTipo
            FEDolares.TextMatrix(objRS.Bookmark, 4) = objRS!nMontoDesde
            FEDolares.TextMatrix(objRS.Bookmark, 5) = objRS!nMontoHasta
            FEDolares.TextMatrix(objRS.Bookmark, 6) = objRS!nTipoMinino
            FEDolares.TextMatrix(objRS.Bookmark, 7) = objRS!nTipoMaximo
            objRS.MoveNext
        Loop
    FrameDolares.Enabled = True
    End If
End Sub
Private Sub LlenarGrillaSoles(ByRef nCantidad As Integer)
    Dim objRS As ADODB.Recordset
    Dim objLinea As COMDCredito.DCOMLineaCredito
    Set objLinea = New COMDCredito.DCOMLineaCredito
    FormateaFlex FESoles
    Set objRS = Nothing
    Set objRS = New ADODB.Recordset
    If chkSoles.value = 1 Then
        Set objRS = objLinea.ObtenerProductoCrediticioTasas(Trim(Right(cboSubProducto.Text, 10)), Trim(Right(cboCampana.Text, 10)), 1)
        Do While Not objRS.EOF
            nCantidad = nCantidad + 1
            FESoles.AdicionaFila
            'FESoles.TextMatrix(objRS.Bookmark, 0) = objRS!nOrden Comentado by NAGL 20190119
            FESoles.TextMatrix(objRS.Bookmark, 1) = "0"
            'FESoles.TextMatrix(objRS.Bookmark, 2) = objRS!bPreferencial
            FESoles.TextMatrix(objRS.Bookmark, 2) = "" 'FRHU 20171107 Acta 189-2017
            FESoles.TextMatrix(objRS.Bookmark, 3) = objRS!cConsDescripcion & Space(200) & objRS!nTipo
            FESoles.TextMatrix(objRS.Bookmark, 4) = objRS!nMontoDesde
            FESoles.TextMatrix(objRS.Bookmark, 5) = objRS!nMontoHasta
            FESoles.TextMatrix(objRS.Bookmark, 6) = objRS!nTipoMinino
            FESoles.TextMatrix(objRS.Bookmark, 7) = objRS!nTipoMaximo
            objRS.MoveNext
        Loop
    FrameSoles.Enabled = True
    End If
End Sub

Private Sub chkSoles_Click()
Dim nCantidad As Integer
If chkSoles.value = 1 Then
    If nRangIniSoles <> 0 Then 'NAGL 20190123
        FrameSoles.Enabled = True
        nCantidad = 0
        FESoles.Enabled = True
        Call LlenarGrillaSoles(nCantidad)
    Else
        If pbMostrarMsg = True Then
            MsgBox "No se encuentra definido el rango de montos en MN, para este producto", vbInformation, "Aviso"
            chkSoles.value = 0
            Exit Sub
        End If
        pbMostrarMsg = True 'NAGL 20190123
    End If
Else
    FrameSoles.Enabled = False
    FESoles.Enabled = False
    FormateaFlex FESoles
End If
End Sub

Private Sub chkDolar_Click()
Dim nCantidad As Integer
If chkDolar.value = 1 Then
    If nRangIniDolares <> 0 Then  'NAGL 20190123
        FrameDolares.Enabled = True
        FEDolares.Enabled = True
        nCantidad = 0
        Call LlenarGrillaDolares(nCantidad)
    Else
        If pbMostrarMsg = True Then
            MsgBox "No se encuentra definido el rango de montos en ME, para este producto", vbInformation, "Aviso"
            chkDolar.value = 0
            Exit Sub
        End If
        pbMostrarMsg = True 'NAGL 20190123
    End If
Else
    FEDolares.Enabled = False
    FormateaFlex FEDolares
    FEDolares.Enabled = False
End If
End Sub

Private Sub chkTAgencias_Click()
    Dim n As Integer
    If chkTAgencias.value = 1 Then
        For n = 1 To lvAgencia.ListItems.count
            lvAgencia.ListItems(n).Checked = True
        Next n
    End If
    If chkTAgencias.value = 0 Then
        For n = 1 To lvAgencia.ListItems.count
            lvAgencia.ListItems(n).Checked = False
        Next n
    End If
End Sub

Private Sub cmdMasD_Click()
  FEDolares.AdicionaFila
  FEDolares.TextMatrix(FEDolares.row, 1) = 1
  FEDolares.Col = 4
End Sub

Private Sub cmdMasS_Click()
    FESoles.AdicionaFila
    FESoles.TextMatrix(FESoles.row, 1) = 1
    FESoles.Col = 4
End Sub

Private Sub cmdMenosD_Click()
    FEDolares.EliminaFila FEDolares.row
End Sub

Private Sub cmdMenosS_Click()
    FESoles.EliminaFila FESoles.row
End Sub


'*****************************************
'Private Sub chkCalificacion_Click()
'If chkCalificacion.value = 1 Then
'    Frame7.Enabled = True
'Else
'    Frame7.Enabled = False
'End If
'    chkA.value = 0
'    chkB.value = 0
'    chkC.value = 0
'    chkD.value = 0
'    chkE.value = 0
'End Sub
'Private Sub chkCaliSBS_Click()
'If chkCaliSBS.value = 1 Then
'    Frame8.Enabled = True
'Else
'    Frame8.Enabled = False
'End If
'
'End Sub
'Private Sub ChkEdad_Click()
'    If chkEdad.value = 1 Then
'        Frame9.Enabled = True
'    Else
'        Frame9.Enabled = False
'    End If
'    txtMinEdad.Text = ""
'    txtMaxEdad.Text = ""
'End Sub
'Private Sub chkGenero_Click()
'If chkGenero.value = 1 Then
'    Frame6.Enabled = True
'Else
'    Frame6.Enabled = False
'End If
'    chkFem.value = 0
'    chkMas.value = 0
'End Sub
'Private Sub chkPlazoMes_Click()
'If chkPlazoMes.value = 1 Then
'    Frame5.Enabled = True
'Else
'    Frame5.Enabled = False
'End If
'txtPlazoMinimo.Text = ""
'txtPlazoMaximo.Text = ""
'End Sub
'Private Sub chkTDestino_Click()
'    Dim n As Integer
'    If chkTDestino.value = 1 Then
'        For n = 1 To lvDestino.ListItems.count
'            lvDestino.ListItems(n).Checked = True
'        Next n
'    End If
'    If chkTDestino.value = 0 Then
'        For n = 1 To lvDestino.ListItems.count
'            lvDestino.ListItems(n).Checked = False
'        Next n
'    End If
'End Sub

'Function ValidarDestino() As String
'Dim lsTexto As String
'Dim nContDes As Integer
'Dim i As Integer
'
'nContDes = 0
'ValidarDestino = True
'
'
'For i = 1 To lvDestino.ListItems.count
'    If lvDestino.ListItems(i).Checked = True Then
'        nContDes = nContDes + 1
'    End If
'Next i
'ValidarDestino = ""
'If nContDes = 0 Then
'    ValidarDestino = "Debe seleccionar por lo menos un Destino del credito"
'End If
'End Function

'*********Comentado by NAGL 20181215*************

Private Sub CmdAceptar_Click()
Dim objLinea As COMDCredito.DCOMLineaCredito
Set objLinea = New COMDCredito.DCOMLineaCredito
Dim i As Integer
'If Len(Trim(validar)) > 0 Then
If Len(Trim(Validar)) > 0 Then
    MsgBox Validar, vbInformation, "Aviso"
    Exit Sub
End If
'End If
Call objLinea.InsertaProductoCrediticioNew(Trim(Right(cboSubProducto.Text, 10)), Trim(Right(cboCampana.Text, 10)), IIf(chkSoles.value = 1, 1, 0), IIf(chkDolar.value = 1, 1, 0))
Call objLinea.EliminarProductoCreditocioTasas(Trim(Right(cboSubProducto.Text, 10)), Trim(Right(cboCampana.Text, 10)), 1)
Call objLinea.EliminarProductoCreditocioTasas(Trim(Right(cboSubProducto.Text, 10)), Trim(Right(cboCampana.Text, 10)), 2)
If chkSoles.value = 1 Then
    For i = 1 To FESoles.rows - 1
        'Call objLinea.InsertaProductoCrediticioTasas(Trim(Right(cboSubProducto.Text, 10)), Trim(Right(cboCampana.Text, 10)), Right(FESoles.TextMatrix(i, 3), 5), 1, FESoles.TextMatrix(i, 0), FESoles.TextMatrix(i, 4), FESoles.TextMatrix(i, 5), FESoles.TextMatrix(i, 6), FESoles.TextMatrix(i, 7), IIf(FESoles.TextMatrix(i, 2) = "", 0, 0))'Comento caso urgente JOEP20210909
        Call InserTasaSoles(Trim(Right(cboSubProducto.Text, 10)), Trim(Right(cboCampana.Text, 10)), Right(FESoles.TextMatrix(i, 3), 5), 1, FESoles.TextMatrix(i, 0), FESoles.TextMatrix(i, 4), FESoles.TextMatrix(i, 5), FESoles.TextMatrix(i, 6), FESoles.TextMatrix(i, 7), IIf(FESoles.TextMatrix(i, 2) = "", 0, 0)) 'caso urgente JOEP20210909
    Next i
End If
If chkDolar.value = 1 Then
    For i = 1 To FEDolares.rows - 1
        'Call objLinea.InsertaProductoCrediticioTasas(Trim(Right(cboSubProducto.Text, 10)), Trim(Right(cboCampana.Text, 10)), Right(FEDolares.TextMatrix(i, 3), 5), 2, FEDolares.TextMatrix(i, 0), FEDolares.TextMatrix(i, 4), FEDolares.TextMatrix(i, 5), FEDolares.TextMatrix(i, 6), FEDolares.TextMatrix(i, 7), IIf(FEDolares.TextMatrix(i, 2) = "", 0, 0))'Comento caso urgente JOEP20210909
        Call InserTasaDolares(Trim(Right(cboSubProducto.Text, 10)), Trim(Right(cboCampana.Text, 10)), Right(FEDolares.TextMatrix(i, 3), 5), 2, FEDolares.TextMatrix(i, 0), FEDolares.TextMatrix(i, 4), FEDolares.TextMatrix(i, 5), FEDolares.TextMatrix(i, 6), FEDolares.TextMatrix(i, 7), IIf(FEDolares.TextMatrix(i, 2) = "", 0, 0)) 'caso urgente JOEP20210909
    Next i
End If
For i = 1 To lvAgencia.ListItems.count
    Call objLinea.InsertaProductoCreditocioAgencia(Trim(Right(cboSubProducto.Text, 10)), Trim(Right(cboCampana.Text, 10)), Right(lvAgencia.ListItems(i).Text, 2), IIf(lvAgencia.ListItems(i).Checked, 1, 0))
Next
'For i = 1 To lvDestino.ListItems.count
'    Call objLinea.InsertaProductoCreditocioDestino(Trim(Right(cboSubProducto.Text, 10)), Trim(Right(cboCampana.Text, 10)), Right(lvDestino.ListItems(i).Text, 2), IIf(lvDestino.ListItems(i).Checked, 1, 0))
'Next Comentado by NAGL 20181219
MsgBox "Los datos se guardaron satisfactoriamente", vbInformation, "Aviso"
'Call cmdCancelar_Click
End Sub

Private Sub InserTasaSoles(ByVal psTpoProdCod As String, ByVal pnCampanaId As Integer, ByVal pnTipoTasa As Integer, ByVal pnMoneda As Integer, ByVal pnOrden As Integer, ByVal pnMontoDesde As Currency, ByVal pnMontoHasta As Currency, ByVal pnTipoMinino As Double, ByVal pnTipoMaximo As Double, ByVal pbPreferencial As Integer)
    Dim SQL As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    oCon.AbreConexion
    SQL = "exec stp_ins_ProductoCrediticioTasas '" & psTpoProdCod & "'," & pnCampanaId & ",' " & pnTipoTasa & "'," & pnMoneda & "," & pnOrden & "," & pnMontoDesde & "," & pnMontoHasta & "," & pnTipoMinino & "," & pnTipoMaximo & "," & pbPreferencial
    oCon.Ejecutar SQL
    oCon.CierraConexion
End Sub

Private Sub InserTasaDolares(ByVal psTpoProdCod As String, ByVal pnCampanaId As Integer, ByVal pnTipoTasa As Integer, ByVal pnMoneda As Integer, ByVal pnOrden As Integer, ByVal pnMontoDesde As Currency, ByVal pnMontoHasta As Currency, ByVal pnTipoMinino As Double, ByVal pnTipoMaximo As Double, ByVal pbPreferencial As Integer)
    Dim SQL As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    oCon.AbreConexion
    SQL = "exec stp_ins_ProductoCrediticioTasas '" & psTpoProdCod & "'," & pnCampanaId & ",' " & pnTipoTasa & "'," & pnMoneda & "," & pnOrden & "," & pnMontoDesde & "," & pnMontoHasta & "," & pnTipoMinino & "," & pnTipoMaximo & "," & pbPreferencial
    oCon.Ejecutar SQL
    oCon.CierraConexion
End Sub

Private Function Validar() As String
Dim i As Integer

If Trim(cboSubProducto.Text) = "" Then
    Validar = "Debe seleccionar un producto crediticio"
    Exit Function
End If

If Trim(cboCampana.Text) = "" Then
    Validar = "Debe seleccionar una campaña de crédito"
    Exit Function
End If

If chkSoles.value = 1 Then
    For i = 1 To FESoles.rows - 1
        If FESoles.TextMatrix(i, 7) = "" Or FESoles.TextMatrix(i, 3) = "" Or FESoles.TextMatrix(i, 4) = "" Or FESoles.TextMatrix(i, 5) = "" Or FESoles.TextMatrix(i, 6) = "" Then
        Validar = "Debe completar todos los datos de todas las filas de las tasas en Soles"
        Exit Function
    End If
    Next i
End If
If chkDolar.value = 1 Then
    For i = 1 To FEDolares.rows - 1
        If FEDolares.TextMatrix(i, 7) = "" Or FEDolares.TextMatrix(i, 3) = "" Or FEDolares.TextMatrix(i, 4) = "" Or FEDolares.TextMatrix(i, 5) = "" Or FEDolares.TextMatrix(i, 6) = "" Then
            Validar = "Debe completar todos los datos de todas las filas de las tasas en Dolares"
            Exit Function
        End If
    Next i
End If
If chkSoles.value = 1 Then
    If FESoles.rows = 0 Then
        Validar = "Ingresar al menos una fila en la grilla de tasas en soles (S/.)"
        Exit Function
    End If
End If
If chkDolar.value = 1 Then
    If FEDolares.rows = 0 Then
        Validar = "Ingresar al menos una fila en la grilla de tasas en dolares ($)"
        Exit Function
    End If
End If
'**************Comentado by NAGL 20181219*******************
'If chkPJ.value = 0 And chkPN.value = 0 Then
'    Validar = "Selecionar al menos una Personeria "
'    Exit Function
'End If
'If chkGenero.value = 1 Then
'    If chkFem.value = 0 And chkMas.value = 0 Then
'        Validar = "Selecionar al menos un Genero"
'        Exit Function
'    End If
'End If
'
'If chkPlazoMes.value = 1 Then
'    If val(txtPlazoMinimo.Text) = 0 Then
'        Validar = "El monto minimo del plazo debe ser mayor que cero (0)"
'        Exit Function
'    End If
'    If val(txtPlazoMaximo.Text) = 0 Then
'        Validar = "El monto maximo del plazo debe ser mayor que cero (0)"
'        Exit Function
'    End If
'    If val(txtPlazoMinimo.Text) > val(txtPlazoMaximo.Text) Then
'        Validar = "El monto minimo del plazo no debe ser mayor que el monto maximo"
'        Exit Function
'    End If
'End If
'If chkCalificacion.value = 1 Then
'    If chkA.value = 0 And chkB.value = 0 And chkC.value = 0 And chkD.value = 0 And chkE.value = 0 Then
'        Validar = "Selecionar al menos una Calificacion Interna "
'        Exit Function
'    End If
'End If
'If chkEdad.value = 1 Then
'    If val(txtMinEdad.Text) = 0 Then
'        Validar = "La edad mínima debe ser mayor que cero (0)"
'        Exit Function
'    End If
'    If val(txtMaxEdad.Text) = 0 Then
'        Validar = "La edad máxima debe ser mayor que cero (0)"
'        Exit Function
'    End If
'    If val(txtMinEdad.Text) > val(txtMaxEdad.Text) Then
'        Validar = "El edad mínima no debe ser mayor que la edad máxima"
'        Exit Function
'    End If
'End If
'Validar = ValidarDestino
'If Len(Trim(Validar)) > 0 Then Exit Function
'******************END NAGL 20181219*********************
    Validar = ValidarAgencia
    If Len(Trim(Validar)) > 0 Then Exit Function
End Function
Function ValidarAgencia() As String
Dim lsTexto As String
Dim nContAge As Integer
Dim i As Integer

nContAge = 0
ValidarAgencia = True

For i = 1 To lvAgencia.ListItems.count
    If lvAgencia.ListItems(i).Checked = True Then
        nContAge = nContAge + 1
    End If
Next i
ValidarAgencia = ""
If nContAge = 0 Then
    ValidarAgencia = "Debe seleccionar por lo menos una Agencia"
End If
End Function

Private Sub cmdCancelar_Click()
 cboSubProducto.Enabled = True
 cboCampana.Enabled = True
 cmdMostrar.Enabled = True
 FormateaFlex FESoles
 FormateaFlex FEDolares
 FormateaFlex flxCaractProd 'NAGL 20181218
 Call LimpiaControles(Me, True)
 Call LimpiaControlesFaltantes
 Call CargaControles
 Call ActualizaControles(False)
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdDeshabilidad_Click()
Dim objLinea As COMDCredito.DCOMLineaCredito
Set objLinea = New COMDCredito.DCOMLineaCredito

If Trim(cboSubProducto.Text) = "" Then
    MsgBox "Debe seleccionar un producto crediticio", vbInformation, "Aviso"
    Exit Sub
End If

If Trim(cboCampana.Text) = "" Then
    MsgBox "Debe seleccionar una campaña de crédito", vbInformation, "Aviso"
    Exit Sub
End If

If nHab = 1 Then
    Call objLinea.DeshabilitarProductoCrediticio(Trim(Right(cboSubProducto.Text, 10)), 0, Trim(Right(cboCampana.Text, 10)))
    MsgBox "El producto se ha Deshabilitado", vbInformation, "Aviso"
    Call cmdCancelar_Click
    cmdAceptar.Enabled = False
    nHab = 0
Else
    Call objLinea.DeshabilitarProductoCrediticio(Trim(Right(cboSubProducto.Text, 10)), 1, Trim(Right(cboCampana.Text, 10)))
    MsgBox "El producto se ha habilitado", vbInformation, "Aviso" 'NAGL Agregó vbInformation
    cmdDeshabilidad.Caption = "Deshabilitar"
    cmdAceptar.Enabled = True
    nHab = 1
End If
End Sub

Private Sub FEDolares_RowColChange()
If FEDolares.Col = 3 Then
    FEDolares.CargaCombo CargarTipoCredito
End If
If FEDolares.Col = 2 Then
    FEDolares.TextMatrix(FEDolares.row, 2) = ""
End If
End Sub

Private Sub FESoles_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim mat() As String
    Dim valorNewDesde As Currency 'NAGL 20190118
    Dim valorNewHasta As Currency 'NAGL 20190118
    mat = Split(FESoles.ColumnasAEditar, "-")
    If mat(pnCol) = "X" Then
        MsgBox "Esta columna es no editable", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
    If pnCol = 4 Or pnCol = 5 Or pnCol = 6 Or pnCol = 7 Then 'pnCol = 3
        If Not IsNumeric(FESoles.TextMatrix(pnRow, pnCol)) Then
            MsgBox "Ud. debe ingresar un monto mayor a cero", vbInformation, "Aviso"
            Cancel = False
            Exit Sub
        Else
            If CCur(FESoles.TextMatrix(pnRow, pnCol)) <= 0 Then
                MsgBox "Ud. debe ingresar un monto mayor a cero", vbInformation, "Aviso"
                Cancel = False
                Exit Sub
            End If
        End If
    End If
    Dim n As Integer
    For n = 1 To FESoles.rows - 1
    If Trim(Right(FESoles.TextMatrix(n, 3), 5)) = Trim(Right(FESoles.TextMatrix(pnRow, 3), 5)) And pnRow <> n Then
        If (Trim(Right(FESoles.TextMatrix(n, 2), 5)) = Trim(Right(FESoles.TextMatrix(pnRow, 2), 5)) And FESoles.Col <> 2) _
            Or (Trim(Right(FESoles.TextMatrix(n, 2), 5)) <> Trim(Right(FESoles.TextMatrix(pnRow, 2), 5)) And FESoles.Col = 2) Then
           If val(FESoles.TextMatrix(pnRow, 4)) >= val(FESoles.TextMatrix(n, 4)) And val(FESoles.TextMatrix(pnRow, 4)) <= val(FESoles.TextMatrix(n, 5)) Then
                MsgBox "Ud. debe ingresar un monto que no este en otros rangos", vbInformation, "Aviso"
                'FESoles.TextMatrix(pnRow, 3) = 0
                Cancel = False
                Exit Sub
           End If
           If val(FESoles.TextMatrix(pnRow, 5)) >= val(FESoles.TextMatrix(n, 4)) And val(FESoles.TextMatrix(pnRow, 5)) <= val(FESoles.TextMatrix(n, 5)) Then
'            If Trim(Right(FESoles.TextMatrix(n, 2), 5)) = Trim(Right(FESoles.TextMatrix(pnRow, 2), 5)) Then
                 MsgBox "Ud. debe ingresar un monto que no este en otros rangos", vbInformation, "Aviso"
                 Cancel = False
                 Exit Sub
'            End If
           End If
           If val(FESoles.TextMatrix(pnRow, 4)) <= val(FESoles.TextMatrix(n, 4)) And val(FESoles.TextMatrix(pnRow, 5)) >= val(FESoles.TextMatrix(n, 5)) Then
'            If Trim(Right(FESoles.TextMatrix(n, 2), 5)) = Trim(Right(FESoles.TextMatrix(pnRow, 2), 5)) Then
                 MsgBox "El rango del monto no debe cubrir a otros rangos (S/.)", vbInformation, "Aviso"
                 Cancel = False
                 Exit Sub
'            End If
           End If
        End If
    End If
    Next n
    'FRHU 20171107 Acta
    If FESoles.TextMatrix(pnRow, 4) <> "" And FESoles.TextMatrix(pnRow, 5) <> "" Then
        If val(FESoles.TextMatrix(pnRow, 4)) > val(FESoles.TextMatrix(pnRow, 5)) Then
            MsgBox "El valor de la columna [Hasta] no puede ser menor al valor de la columna [Desde]", vbInformation, "Aviso"
            Cancel = False
            Exit Sub
        End If
    End If
    If FESoles.TextMatrix(pnRow, 6) <> "" And FESoles.TextMatrix(pnRow, 7) <> "" Then
        If val(FESoles.TextMatrix(pnRow, 6)) > val(FESoles.TextMatrix(pnRow, 7)) Then
            MsgBox "El valor de la columna [T.Max] no puede ser menor al valor de la columna [T.Min]", vbInformation, "Aviso"
            Cancel = False
            Exit Sub
        End If
    End If
     'FIN FRHU 20171107
'**************NAGL 20190118 Según ERS042-2018**********************************************
If pnCol = 4 Then
    If (FESoles.TextMatrix(pnRow, 4) <> "" And IsNumeric(FESoles.TextMatrix(pnRow, 4)) And Len(FESoles.TextMatrix(pnRow, 4)) < 15) Then
        valorNewDesde = CDbl(FESoles.TextMatrix(pnRow, 4))
        If (valorNewDesde < 0) Then
            MsgBox "No se puede asignar un valor Negativo", vbInformation, "Aviso"
            Cancel = False
            Exit Sub
        ElseIf (valorNewDesde > CDbl(FESoles.TextMatrix(pnRow, 4))) Then
            MsgBox "El valor de la columna [Desde] no puede ser mayor al valor de la columna [Hasta]", vbInformation, "Aviso"
            Cancel = False
            Exit Sub
        ElseIf (nRangIniSoles <> 0 And valorNewDesde < nRangIniSoles) Then
            MsgBox "El Monto inicial no se encuentra en el rango propuesto por créditos ", vbInformation, "Aviso"
            Cancel = False
            Exit Sub
        End If
    End If
ElseIf pnCol = 5 Then
    If (FESoles.TextMatrix(pnRow, 5) <> "" And IsNumeric(FESoles.TextMatrix(pnRow, 5)) And Len(FESoles.TextMatrix(pnRow, 5)) < 15) Then
        valorNewHasta = CDbl(FESoles.TextMatrix(pnRow, 5))
        If (valorNewHasta < 0) Then
            MsgBox "No se puede asignar un valor Negativo", vbInformation, "Aviso"
            Cancel = False
            Exit Sub
        ElseIf (nRangFinSoles <> 0 And valorNewHasta > nRangFinSoles) Then
            MsgBox "El Monto Final no se encuentra en el rango propuesto por créditos ", vbInformation, "Aviso"
             Cancel = False
            Exit Sub
        End If
    End If
End If
'************************END 20190119*********************************************************
End Sub

Private Sub FEDolares_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim mat() As String
    Dim valorNewDesde As Currency 'NAGL 20190118
    Dim valorNewHasta As Currency 'NAGL 20190118
    
    mat = Split(FEDolares.ColumnasAEditar, "-")
    If mat(pnCol) = "X" Then
        MsgBox "Esta columna es no editable", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
    If pnCol = 4 Or pnCol = 5 Or pnCol = 6 Or pnCol = 7 Then 'pnCol = 3
        If Not IsNumeric(FEDolares.TextMatrix(pnRow, pnCol)) Then
            MsgBox "Ud. debe ingresar un monto mayor a cero", vbInformation, "Aviso"
            Cancel = False
            Exit Sub
        Else
            If CCur(FEDolares.TextMatrix(pnRow, pnCol)) <= 0 Then
                MsgBox "Ud. debe ingresar un monto mayor a cero", vbInformation, "Aviso"
                Cancel = False
                Exit Sub
            End If
        End If
    End If
    Dim n As Integer
    For n = 1 To FEDolares.rows - 1
    If Trim(Right(FEDolares.TextMatrix(n, 3), 5)) = Trim(Right(FEDolares.TextMatrix(pnRow, 3), 5)) And pnRow <> n Then
           If (Trim(Right(FEDolares.TextMatrix(n, 2), 5)) = Trim(Right(FEDolares.TextMatrix(pnRow, 2), 5)) And FEDolares.Col <> 2) _
            Or (Trim(Right(FEDolares.TextMatrix(n, 2), 5)) <> Trim(Right(FEDolares.TextMatrix(pnRow, 2), 5)) And FEDolares.Col = 2) Then
           If val(FEDolares.TextMatrix(pnRow, 4)) >= val(FEDolares.TextMatrix(n, 4)) And val(FEDolares.TextMatrix(pnRow, 4)) <= val(FEDolares.TextMatrix(n, 5)) Then
                MsgBox "Ud. debe ingresar un monto que no este en otros rangos(S/.)", vbInformation, "Aviso"
                Cancel = False
                Exit Sub
           End If
           If val(FEDolares.TextMatrix(pnRow, 5)) >= val(FEDolares.TextMatrix(n, 4)) And val(FEDolares.TextMatrix(pnRow, 5)) <= val(FEDolares.TextMatrix(n, 5)) Then
                MsgBox "Ud. debe ingresar un monto que no este en otros rangos ($)", vbInformation, "Aviso"
                Cancel = False
                Exit Sub
           End If
           If val(FEDolares.TextMatrix(pnRow, 4)) <= val(FEDolares.TextMatrix(n, 4)) And val(FEDolares.TextMatrix(pnRow, 5)) >= val(FEDolares.TextMatrix(n, 5)) Then
                MsgBox "El rango del monto no debe cubrir a otros rangos($)", vbInformation, "Aviso"
                Cancel = False
                Exit Sub
           End If
        End If
    End If
    Next n
    
'**************NAGL 20190118 Según ERS042-2018**********************************************
If FEDolares.TextMatrix(pnRow, 4) <> "" And FEDolares.TextMatrix(pnRow, 5) <> "" Then
    If val(FEDolares.TextMatrix(pnRow, 4)) > val(FEDolares.TextMatrix(pnRow, 5)) Then
        MsgBox "El valor de la columna [Hasta] no puede ser menor al valor de la columna [Desde]", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
End If
If FEDolares.TextMatrix(pnRow, 6) <> "" And FEDolares.TextMatrix(pnRow, 7) <> "" Then
    If val(FEDolares.TextMatrix(pnRow, 6)) > val(FEDolares.TextMatrix(pnRow, 7)) Then
        MsgBox "El valor de la columna [T.Max] no puede ser menor al valor de la columna [T.Min]", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
End If

If pnCol = 4 Then
    If (FEDolares.TextMatrix(pnRow, 4) <> "" And IsNumeric(FEDolares.TextMatrix(pnRow, 4)) And Len(FEDolares.TextMatrix(pnRow, 4)) < 15) Then
        valorNewDesde = CDbl(FEDolares.TextMatrix(pnRow, 4))
        If (valorNewDesde < 0) Then
            MsgBox "No se puede asignar un valor Negativo", vbInformation, "Aviso"
            Cancel = False
            Exit Sub
        ElseIf (valorNewDesde > CDbl(FEDolares.TextMatrix(pnRow, 4))) Then
            MsgBox "El valor de la columna [Desde] no puede ser mayor al valor de la columna [Hasta]", vbInformation, "Aviso"
            Cancel = False
            Exit Sub
        ElseIf (nRangIniDolares <> 0 And valorNewDesde < nRangIniDolares) Then
            MsgBox "El Monto inicial no se encuentra en el rango propuesto por créditos ", vbInformation, "Aviso"
            Cancel = False
            Exit Sub
        End If
    End If
ElseIf pnCol = 5 Then
    If (FEDolares.TextMatrix(pnRow, 5) <> "" And IsNumeric(FEDolares.TextMatrix(pnRow, 5)) And Len(FEDolares.TextMatrix(pnRow, 5)) < 15) Then
        valorNewHasta = CDbl(FEDolares.TextMatrix(pnRow, 5))
        If (valorNewHasta < 0) Then
            MsgBox "No se puede asignar un valor Negativo", vbInformation, "Aviso"
            Cancel = False
            Exit Sub
        ElseIf (nRangFinDolares <> 0 And valorNewHasta > nRangFinDolares) Then
            MsgBox "El Monto Final no se encuentra en el rango propuesto por créditos.", vbInformation, "Aviso"
            Cancel = False
            Exit Sub
        End If
    End If
End If
'************************END 20190119*********************************************************
End Sub

Private Sub FESoles_RowColChange()
If FESoles.Col = 3 Then
    FESoles.CargaCombo CargarTipoCredito
End If
If FESoles.Col = 2 Then
   FESoles.TextMatrix(FESoles.row, 2) = ""
End If
End Sub

Private Function CargarTipoCredito() As ADODB.Recordset
Dim rsTipoTasas As ADODB.Recordset
Set rsTipoTasas = New ADODB.Recordset

Dim objLinea As COMDCredito.DCOMLineaCredito
Set objLinea = New COMDCredito.DCOMLineaCredito
Dim objRS As ADODB.Recordset
Set objRS = New ADODB.Recordset
Set objRS = objLinea.ObtenerConstanteLineaCredito(3012)
With rsTipoTasas
    'Crear RecordSet
     .Fields.Append "desc", adVarChar, 50
     .Fields.Append "value", adVarChar, 3
    
    .Open
    'Llenar Recordset
    Do While Not objRS.EOF
        .AddNew
        .Fields("desc") = objRS!cDescri
        .Fields("value") = objRS!cCodigo
        objRS.MoveNext
    Loop
End With
rsTipoTasas.MoveFirst
Set CargarTipoCredito = rsTipoTasas
Set rsTipoTasas = Nothing
End Function

Private Function TienePunto(psCadena As String) As Boolean
If InStr(1, psCadena, ".", vbTextCompare) > 0 Then
    TienePunto = True
Else
    TienePunto = False
End If
End Function

Private Function NumDecimal(psCadena As String) As Integer
Dim lnPos As Integer
lnPos = InStr(1, psCadena, ".", vbTextCompare)
If lnPos > 0 Then
    NumDecimal = Len(psCadena) - lnPos
Else
    NumDecimal = 0
End If
End Function


Public Sub LimpiaControlesFaltantes()
    chkCalificacion.value = 0
    Frame7.Enabled = False
    chkCaliSBS.value = 0
    Frame8.Enabled = False
    chkDolar.value = 0
    FEDolares.Enabled = False
    chkEdad.value = 0
    Frame9.Enabled = False
    chkGenero.value = 0
    Frame6.Enabled = False
    chkPlazoMes.value = 0
    Frame5.Enabled = False
    chkSoles.value = 0
    FESoles.Enabled = False
    chkPJ.value = 0
    chkPN.value = 0
    chkA.value = 0
    chkB.value = 0
    chkC.value = 0
    chkD.value = 0
    chkE.value = 0
    chkSoles.value = 0
    chkDolar.value = 0
'    FrameSoles.Enabled = False
'    FrameDolares.Enabled = False
    chkGenero.value = 0
    chkFem.value = 0
    chkMas.value = 0
    chkTAgencias.value = 0
    chkTDestino.value = 0
End Sub

'******Comentado by NAGL 20181219
'Private Sub txtPlazoMinimo_Change()
'txtPlazoMinimo.SelStart = Len(txtPlazoMinimo)
'gnNumDec = NumDecimal(txtPlazoMinimo)
'If gbEstado And txtPlazoMinimo <> "" Then
'    Select Case gnNumDec
'        Case 0
'                txtPlazoMinimo = Format(txtPlazoMinimo, "#,##0")
'        Case 1
'                txtPlazoMinimo = Format(txtPlazoMinimo, "#,##0.0")
'        Case 2
'                txtPlazoMinimo = Format(txtPlazoMinimo, "#,##0.00")
'        Case 3
'                txtPlazoMinimo = Format(txtPlazoMinimo, "#,##0.000")
'        Case Else
'                txtPlazoMinimo = Format(txtPlazoMinimo, "#,##0.0000")
'    End Select
'End If
'gbEstado = False
'RaiseEvent Change
'End Sub
'
'Private Sub txtPlazoMinimo_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(txtPlazoMinimo, KeyAscii, , 4)
'    If KeyAscii = 13 Then
'
'        txtPlazoMaximo.SetFocus
'    End If
'End Sub
'
'Private Sub txtPlazoMinimo_GotFocus()
'With txtPlazoMinimo
'    .SelStart = 0#
'    .SelLength = Len(.Text)
'End With
'End Sub
'
'Private Sub txtPlazoMaximo_Change()
'txtPlazoMaximo.SelStart = Len(txtPlazoMaximo)
'gnNumDec = NumDecimal(txtPlazoMaximo)
'If gbEstado And txtPlazoMaximo <> "" Then
'    Select Case gnNumDec
'        Case 0
'                txtPlazoMaximo = Format(txtPlazoMaximo, "#,##0")
'        Case 1
'                txtPlazoMaximo = Format(txtPlazoMaximo, "#,##0.0")
'        Case 2
'                txtPlazoMaximo = Format(txtPlazoMaximo, "#,##0.00")
'        Case 3
'                txtPlazoMaximo = Format(txtPlazoMaximo, "#,##0.000")
'        Case Else
'                txtPlazoMaximo = Format(txtPlazoMaximo, "#,##0.0000")
'    End Select
'End If
'gbEstado = False
'RaiseEvent Change
'End Sub
'Private Sub txtPlazoMaximo_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(txtPlazoMaximo, KeyAscii, , 4)
'    If KeyAscii = 13 Then
'        txtPlazoMaximo.SetFocus
'    End If
'End Sub
'Private Sub txtPlazoMaximo_GotFocus()
'With txtPlazoMaximo
'    .SelStart = 0#
'    .SelLength = Len(.Text)
'End With
'End Sub
'Private Sub txtCalMinimo_Change()
'txtCalMinimo.SelStart = Len(txtCalMinimo)
'gnNumDec = NumDecimal(txtCalMinimo)
'If gbEstado And txtCalMinimo <> "" Then
'    Select Case gnNumDec
'        Case 0
'                txtCalMinimo = Format(txtCalMinimo, "#,##0")
'        Case 1
'                txtCalMinimo = Format(txtCalMinimo, "#,##0.0")
'        Case 2
'                txtCalMinimo = Format(txtCalMinimo, "#,##0.00")
'        Case 3
'                txtCalMinimo = Format(txtCalMinimo, "#,##0.000")
'        Case Else
'                txtCalMinimo = Format(txtCalMinimo, "#,##0.0000")
'    End Select
'End If
'gbEstado = False
'RaiseEvent Change
'End Sub
'Private Sub txtCalMinimo_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(txtCalMinimo, KeyAscii, , 4)
'    If KeyAscii = 13 Then
'        txtCalMinimo.SetFocus
'    End If
'End Sub
'
'Private Sub txtCalMinimo_GotFocus()
'With txtCalMinimo
'    .SelStart = 0#
'    .SelLength = Len(.Text)
'End With
'End Sub
'Private Sub txtMinEdad_Change()
'txtMinEdad.SelStart = Len(txtMinEdad)
'gnNumDec = NumDecimal(txtMinEdad)
'If gbEstado And txtMinEdad <> "" Then
'    Select Case gnNumDec
'        Case 0
'                txtMinEdad = Format(txtMinEdad, "#,##0")
'        Case 1
'                txtMinEdad = Format(txtMinEdad, "#,##0.0")
'        Case 2
'                txtMinEdad = Format(txtMinEdad, "#,##0.00")
'        Case 3
'                txtMinEdad = Format(txtMinEdad, "#,##0.000")
'        Case Else
'                txtMinEdad = Format(txtMinEdad, "#,##0.0000")
'    End Select
'End If
'gbEstado = False
'RaiseEvent Change
'End Sub
'Private Sub txtMinEdad_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(txtMinEdad, KeyAscii, , 4)
'    If KeyAscii = 13 Then
'        txtMaxEdad.SetFocus
'    End If
'End Sub
'
'Private Sub txtMinEdad_GotFocus()
'With txtMinEdad
'    .SelStart = 0#
'    .SelLength = Len(.Text)
'End With
'End Sub
'Private Sub txtMaxEdad_Change()
'txtMaxEdad.SelStart = Len(txtMaxEdad)
'gnNumDec = NumDecimal(txtMaxEdad)
'If gbEstado And txtMaxEdad <> "" Then
'    Select Case gnNumDec
'        Case 0
'                txtMaxEdad = Format(txtMaxEdad, "#,##0")
'        Case 1
'                txtMaxEdad = Format(txtMaxEdad, "#,##0.0")
'        Case 2
'                txtMaxEdad = Format(txtMaxEdad, "#,##0.00")
'        Case 3
'                txtMaxEdad = Format(txtMaxEdad, "#,##0.000")
'        Case Else
'                txtMaxEdad = Format(txtMaxEdad, "#,##0.0000")
'    End Select
'End If
'gbEstado = False
'RaiseEvent Change
'End Sub
'Private Sub txtMaxEdad_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(txtMaxEdad, KeyAscii, , 4)
'    If KeyAscii = 13 Then
'        txtMaxEdad.SetFocus
'    End If
'End Sub
'
'Private Sub txtMaxEdad_GotFocus()
'With txtMaxEdad
'    .SelStart = 0#
'    .SelLength = Len(.Text)
'End With
'End Sub
'******END NAGL 20181219


