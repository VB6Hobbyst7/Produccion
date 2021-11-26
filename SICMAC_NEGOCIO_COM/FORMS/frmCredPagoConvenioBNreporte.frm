VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{2E37A9B8-9906-4590-9F3A-E67DB58122F5}#7.0#0"; "OcxLabelX.ocx"
Begin VB.Form frmCredPagoConvenioBNreporte 
   Caption         =   "Reporte de Pagos por Convenio / Corresponsalia BN"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   Icon            =   "frmCredPagoConvenioBNreporte.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   5535
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   10815
      Begin TabDlg.SSTab SSTAB 
         Height          =   5175
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   9128
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         TabCaption(0)   =   "Datos Pagados"
         TabPicture(0)   =   "frmCredPagoConvenioBNreporte.frx":030A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lbl1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lbl4"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lbl3"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblreg"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lbldolar"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lblsoles"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Mshbco"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "Datos No Pagados"
         TabPicture(1)   =   "frmCredPagoConvenioBNreporte.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label1"
         Tab(1).Control(1)=   "Label4"
         Tab(1).Control(2)=   "Label5"
         Tab(1).Control(3)=   "lblimpdx"
         Tab(1).Control(4)=   "lblimpsx"
         Tab(1).Control(5)=   "lblnumx"
         Tab(1).Control(6)=   "Mshbco1"
         Tab(1).ControlCount=   7
         TabCaption(2)   =   "Datos Sobrantes"
         TabPicture(2)   =   "frmCredPagoConvenioBNreporte.frx":0342
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label10"
         Tab(2).Control(1)=   "Label11"
         Tab(2).Control(2)=   "Label12"
         Tab(2).Control(3)=   "lblnumx3"
         Tab(2).Control(4)=   "lblsoles3"
         Tab(2).Control(5)=   "lbldol3"
         Tab(2).Control(6)=   "Mshbco2"
         Tab(2).ControlCount=   7
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Mshbco 
            Height          =   3975
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   7011
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Mshbco1 
            Height          =   3735
            Left            =   -74760
            TabIndex        =   11
            Top             =   600
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   6588
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Mshbco2 
            Height          =   3735
            Left            =   -74880
            TabIndex        =   24
            Top             =   600
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   6588
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin OcxLabelX.LabelX lblsoles 
            Height          =   495
            Left            =   6240
            TabIndex        =   29
            Top             =   4560
            Visible         =   0   'False
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   873
            FondoBlanco     =   0   'False
            Resalte         =   0
            Bold            =   0   'False
            Alignment       =   0
         End
         Begin OcxLabelX.LabelX lbldolar 
            Height          =   495
            Left            =   8880
            TabIndex        =   30
            Top             =   4560
            Visible         =   0   'False
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   873
            FondoBlanco     =   0   'False
            Resalte         =   0
            Bold            =   0   'False
            Alignment       =   0
         End
         Begin OcxLabelX.LabelX lblreg 
            Height          =   495
            Left            =   3720
            TabIndex        =   31
            Top             =   4560
            Visible         =   0   'False
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   873
            FondoBlanco     =   0   'False
            Resalte         =   0
            Bold            =   0   'False
            Alignment       =   0
         End
         Begin OcxLabelX.LabelX lblnumx 
            Height          =   495
            Left            =   -71040
            TabIndex        =   32
            Top             =   4560
            Visible         =   0   'False
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   873
            FondoBlanco     =   0   'False
            Resalte         =   0
            Bold            =   0   'False
            Alignment       =   0
         End
         Begin OcxLabelX.LabelX lblimpsx 
            Height          =   495
            Left            =   -68520
            TabIndex        =   33
            Top             =   4560
            Visible         =   0   'False
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   873
            FondoBlanco     =   0   'False
            Resalte         =   0
            Bold            =   0   'False
            Alignment       =   0
         End
         Begin OcxLabelX.LabelX lblimpdx 
            Height          =   495
            Left            =   -66000
            TabIndex        =   34
            Top             =   4560
            Visible         =   0   'False
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   873
            FondoBlanco     =   0   'False
            Resalte         =   0
            Bold            =   0   'False
            Alignment       =   0
         End
         Begin OcxLabelX.LabelX lbldol3 
            Height          =   495
            Left            =   -66000
            TabIndex        =   35
            Top             =   4560
            Visible         =   0   'False
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   873
            FondoBlanco     =   0   'False
            Resalte         =   0
            Bold            =   0   'False
            Alignment       =   0
         End
         Begin OcxLabelX.LabelX lblsoles3 
            Height          =   495
            Left            =   -68520
            TabIndex        =   36
            Top             =   4560
            Visible         =   0   'False
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   873
            FondoBlanco     =   0   'False
            Resalte         =   0
            Bold            =   0   'False
            Alignment       =   0
         End
         Begin OcxLabelX.LabelX lblnumx3 
            Height          =   495
            Left            =   -70920
            TabIndex        =   37
            Top             =   4560
            Visible         =   0   'False
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   873
            FondoBlanco     =   0   'False
            Resalte         =   0
            Bold            =   0   'False
            Alignment       =   0
         End
         Begin VB.Label lbl3 
            AutoSize        =   -1  'True
            Caption         =   "Imp Soles :"
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
            Height          =   195
            Left            =   5280
            TabIndex        =   8
            Top             =   4680
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label lbl4 
            AutoSize        =   -1  'True
            Caption         =   "Imp Dolares :"
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
            Height          =   195
            Left            =   7680
            TabIndex        =   9
            Top             =   4680
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label lbl1 
            AutoSize        =   -1  'True
            Caption         =   "Procesados :"
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
            Height          =   195
            Left            =   2520
            TabIndex        =   15
            Top             =   4680
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Imp Soles :"
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
            Height          =   195
            Left            =   -69600
            TabIndex        =   18
            Top             =   4680
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Numero :"
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
            Height          =   195
            Left            =   -71880
            TabIndex        =   19
            Top             =   4680
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Imp Dolares :"
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
            Height          =   195
            Left            =   -67320
            TabIndex        =   20
            Top             =   4680
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Imp Dolares :"
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
            Height          =   195
            Left            =   -67440
            TabIndex        =   17
            Top             =   4680
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Numero :"
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
            Height          =   195
            Left            =   -72000
            TabIndex        =   16
            Top             =   4680
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "No procesados :"
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
            Height          =   195
            Left            =   -74760
            TabIndex        =   14
            Top             =   4680
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Dolares :"
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
            Height          =   195
            Left            =   -67200
            TabIndex        =   13
            Top             =   4680
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Imp Soles :"
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
            Height          =   195
            Left            =   -69720
            TabIndex        =   12
            Top             =   4680
            Visible         =   0   'False
            Width           =   960
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      Begin VB.ComboBox cbotipo 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Text            =   "cbotipo"
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmd_buscar 
         Caption         =   "&Buscar"
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
         Left            =   8280
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton btn_generar 
         Caption         =   "&Generar"
         Enabled         =   0   'False
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
         Left            =   9600
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txNotaIni 
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin OcxLabelX.LabelX lblcab1 
         Height          =   495
         Left            =   3360
         TabIndex        =   26
         Top             =   360
         Visible         =   0   'False
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   873
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin OcxLabelX.LabelX lblcab2 
         Height          =   495
         Left            =   5160
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   873
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin OcxLabelX.LabelX lblcab3 
         Height          =   495
         Left            =   6840
         TabIndex        =   28
         Top             =   360
         Visible         =   0   'False
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   873
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Tipo :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   1920
         TabIndex        =   25
         Top             =   120
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Imp Soles :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   5160
         TabIndex        =   23
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Imp Dolares :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   6720
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Procesados :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   3360
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha :"
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
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmCredPagoConvenioBNreporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim est As Integer
Dim sCadImpre As String
 Sub Cargar_grilla(ByVal cod_det As Double)
    Dim rsx As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Dim i As Integer
    Set rsx = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos

    Set rsx = oGen.DevolverDatos_Detalle_108121(cod_det)
    Set oGen = Nothing
    i = 0

    ConfigurarMShComite
      Do Until rsx.EOF
        With Me.Mshbco
             i = i + 1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 2, 0) = i 'rsX!id
            .TextMatrix(.Rows - 2, 1) = rsx!cCodCtafull
            .TextMatrix(.Rows - 2, 2) = rsx!cNombreCliente
            .TextMatrix(.Rows - 2, 3) = rsx!nNroCuota
            .TextMatrix(.Rows - 2, 4) = rsx!nImporteCuota
            .TextMatrix(.Rows - 2, 5) = rsx!nImporteCobrado
        End With
        rsx.MoveNext
    Loop
'  Me.cmdImprimir.Enabled = True
'  Me.cmdReporte.Enabled = True
  rsx.Close
Set rsx = Nothing
End Sub

Sub Cargar_grillaCorresponsalia(ByVal cod_det As Double)
    Dim rsx1 As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Dim i As Integer
    Set rsx1 = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos

    Set rsx1 = oGen.DevolverDatos_Detalle_108121_Corresponsalia(cod_det)
    Set oGen = Nothing
    i = 0

    ConfigurarMShComite
      Do Until rsx1.EOF
        With Me.Mshbco
             i = i + 1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 2, 0) = i 'rsX!id
            .TextMatrix(.Rows - 2, 1) = rsx1!cCodCtafull
            .TextMatrix(.Rows - 2, 2) = rsx1!cNombreCliente
            .TextMatrix(.Rows - 2, 3) = rsx1!nNroCuota
            .TextMatrix(.Rows - 2, 4) = IIf(rsx1!nMoneda = 1, "Soles", "Dolares")
            .TextMatrix(.Rows - 2, 5) = rsx1!nImporteCobrado
        End With
        rsx1.MoveNext
    Loop
'  Me.cmdImprimir.Enabled = True
'  Me.cmdReporte.Enabled = True
  rsx1.Close
Set rsx1 = Nothing
End Sub

Sub Cargar_grilla1(ByVal cod_det As Double)
    Dim rsy1 As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Dim i As Integer
    Set rsy1 = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos

    Set rsy1 = oGen.DevolverDatos_Detalle_108121_NoPagados(cod_det)
    Set oGen = Nothing
    i = 0

    ConfigurarMShComite1
      Do Until rsy1.EOF
        With Me.Mshbco1
             i = i + 1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 2, 0) = i 'rsX!id
            .TextMatrix(.Rows - 2, 1) = rsy1!cCodCtafull
            .TextMatrix(.Rows - 2, 2) = rsy1!cNombreCliente
            .TextMatrix(.Rows - 2, 3) = rsy1!nNroCuota
            .TextMatrix(.Rows - 2, 4) = rsy1!nImporteCuota
            .TextMatrix(.Rows - 2, 5) = rsy1!nImporteCobrado
        End With
        rsy1.MoveNext
    Loop
'  Me.cmdImprimir.Enabled = True
'  Me.cmdReporte.Enabled = True
  rsy1.Close
Set rsy1 = Nothing
End Sub

Sub Cargar_grilla2(ByVal dFecha As String)
    Dim rsz1 As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Dim i As Integer
    Set rsz1 = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos

    Set rsz1 = oGen.DevolverDatos_Detalle_Sobrante(CDate(dFecha))
    Set oGen = Nothing
    i = 0

    ConfigurarMShComite2
      Do Until rsz1.EOF
        With Me.Mshbco2
             i = i + 1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 2, 0) = i 'rsX!id
            .TextMatrix(.Rows - 2, 1) = rsz1!cCodCtafull
            .TextMatrix(.Rows - 2, 2) = rsz1!cNombreCliente
            .TextMatrix(.Rows - 2, 3) = rsz1!nNroCuota
            .TextMatrix(.Rows - 2, 4) = rsz1!nImporteCobrado
            .TextMatrix(.Rows - 2, 5) = rsz1!nMontoDIF
        End With
        rsz1.MoveNext
    Loop
'  Me.cmdImprimir.Enabled = True
'  Me.cmdReporte.Enabled = True
  rsz1.Close
  Set rsz1 = Nothing
End Sub

Sub Cargar_grilla1_Corresponsalia(ByVal cod_det As Double)
    Dim rsy As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Dim i As Integer
    Set rsy = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos

    Set rsy = oGen.DevolverDatos_Detalle_108121_NoPagados_Corresponsalia(cod_det)
    Set oGen = Nothing
    i = 0

    ConfigurarMShComite1
      Do Until rsy.EOF
        With Me.Mshbco1
             i = i + 1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 2, 0) = i 'rsX!id
            .TextMatrix(.Rows - 2, 1) = rsy!cCodCtafull
            .TextMatrix(.Rows - 2, 2) = rsy!cNombreCliente
            .TextMatrix(.Rows - 2, 3) = rsy!nNroCuota
            .TextMatrix(.Rows - 2, 4) = IIf(rsy!nMoneda = 1, "Soles", "Dolares")
            .TextMatrix(.Rows - 2, 5) = rsy!nImporteCobrado
        End With
        rsy.MoveNext
    Loop
'  Me.cmdImprimir.Enabled = True
'  Me.cmdReporte.Enabled = True
  rsy.Close
Set rsy = Nothing
End Sub
Sub Cargar_grilla2_Corresponsalia(ByVal dFecha As String)
    Dim rsz As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Dim i As Integer
    Set rsz = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos

    Set rsz = oGen.DevolverDatos_Detalle_Sobrante_Corresponsalia(CDate(dFecha))
    Set oGen = Nothing
    i = 0

    ConfigurarMShComite2
      Do Until rsz.EOF
        With Me.Mshbco2
             i = i + 1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 2, 0) = i 'rsX!id
            .TextMatrix(.Rows - 2, 1) = rsz!cCodCtafull
            .TextMatrix(.Rows - 2, 2) = rsz!cNombreCliente
            .TextMatrix(.Rows - 2, 3) = rsz!nNroCuota
            .TextMatrix(.Rows - 2, 4) = rsz!nImporteCobrado
            .TextMatrix(.Rows - 2, 5) = rsz!nMontoDIF
        End With
        rsz.MoveNext
    Loop
'  Me.cmdImprimir.Enabled = True
'  Me.cmdReporte.Enabled = True
  rsz.Close
  Set rsz = Nothing
End Sub


Sub ConfigurarMShComite()
 Mshbco.Clear
    Mshbco.Cols = 6
    Mshbco.Rows = 2

    With Mshbco
        .TextMatrix(0, 0) = "Id"
        .TextMatrix(0, 1) = "Cuenta"
        .TextMatrix(0, 2) = "Nombre Cliente"
        .TextMatrix(0, 3) = "Cuota"
        .TextMatrix(0, 4) = "Imp.Cuota"
        .TextMatrix(0, 5) = "Imp.Cobrado"

        .ColWidth(0) = 400
        .ColWidth(1) = 1800
        .ColWidth(2) = 4500
        .ColWidth(3) = 700
        .ColWidth(4) = 1100
        .ColWidth(5) = 1100

    End With
End Sub
Sub ConfigurarMShComite1()
 Mshbco1.Clear
    Mshbco1.Cols = 6
    Mshbco1.Rows = 2

    With Mshbco1
        .TextMatrix(0, 0) = "Id"
        .TextMatrix(0, 1) = "Cuenta"
        .TextMatrix(0, 2) = "Nombre Cliente"
        .TextMatrix(0, 3) = "Cuota"
        .TextMatrix(0, 4) = "Imp.Cuota"
        .TextMatrix(0, 5) = "Imp.Cobrado"

        .ColWidth(0) = 400
        .ColWidth(1) = 1800
        .ColWidth(2) = 4500
        .ColWidth(3) = 700
        .ColWidth(4) = 1100
        .ColWidth(5) = 1100

    End With
End Sub

Sub ConfigurarMShComite2()
 Mshbco2.Clear
    Mshbco2.Cols = 6
    Mshbco2.Rows = 2

    With Mshbco2
        .TextMatrix(0, 0) = "Id"
        .TextMatrix(0, 1) = "Cuenta"
        .TextMatrix(0, 2) = "Nombre Cliente"
        .TextMatrix(0, 3) = "Cuota"
        .TextMatrix(0, 4) = "Imp.Cobrado"
        .TextMatrix(0, 5) = "Diferencia"

        .ColWidth(0) = 400
        .ColWidth(1) = 1800
        .ColWidth(2) = 4500
        .ColWidth(3) = 700
        .ColWidth(4) = 1100
        .ColWidth(5) = 1000

    End With
End Sub
Sub activar_lbl(ByVal xb As Boolean)
lblcab1.Visible = xb
lblcab2.Visible = xb
lblcab3.Visible = xb
Me.lbl1.Visible = xb
Me.lbl3.Visible = xb
Me.Label5.Visible = xb
Me.Label6.Visible = xb
Me.Label7.Visible = xb
Me.Label9.Visible = xb
Me.lbl4.Visible = xb
Me.lbldolar.Visible = xb
Me.lblsoles.Visible = xb
Me.lblimpdx.Visible = xb
Me.lblimpsx.Visible = xb
Me.lblreg.Visible = xb
End Sub

Sub activar_lbl1(ByVal xb As Boolean)
Me.SSTAB.TabVisible(1) = True
Me.Label1.Visible = xb
Me.Label2.Visible = xb
Me.Label3.Visible = xb
Me.Label4.Visible = xb
Me.lblimpdx.Visible = xb
Me.lblimpsx.Visible = xb
Me.lblnumx.Visible = xb
End Sub

Sub activar_lbl2(ByVal xb As Boolean)
Me.Label10.Visible = xb
Me.Label12.Visible = xb
Me.Label11.Visible = xb
Me.lblnumx3.Visible = xb
Me.lblsoles3.Visible = xb
Me.lbldol3.Visible = xb
Me.SSTAB.TabVisible(2) = xb
End Sub
Private Sub btn_generar_Click()
 Dim sCadImp As String
    Dim oPrev As previo.clsprevio
    Dim oGen As COMNCredito.NCOMCredito

    Set oPrev = New previo.clsprevio

    Set oGen = New COMNCredito.NCOMCredito

    If (Me.txNotaIni) <> "" Then
        If Me.cbotipo.ListIndex = 1 Then
            sCadImp = oGen.ImprimeClientesCorresponsaliaBN(gsCodUser, gdFecSis, CDate(Me.txNotaIni), gsNomCmac)
        Else
            sCadImp = oGen.ImprimeClientesRecuperacionesBN(gsCodUser, gdFecSis, CDate(Me.txNotaIni), gsNomCmac)
        End If
    Else
        MsgBox "No se puede presentar el reporte de Pagos"
        Exit Sub
    End If
    
    'Set oNPers = Nothing

    previo.Show sCadImp, "Registro de Archivo de Recuperaciones Cobradas", False
    Set oPrev = Nothing
End Sub

Private Sub cmd_buscar_Click()
If Not IsDate(txNotaIni) Then
    MsgBox "Los Valores indicados en los textos no son Correctos"
    Exit Sub
End If

activar_lbl2 False

If Me.cbotipo.ListIndex = 0 Then
    cargar_datos
Else
    cargar_datosCorres
End If

If est = 1 Then
    btn_generar.Enabled = True
Else
    MsgBox "No hay datos para Mostrar en la Fecha Indicada, Verífique!!"
End If
End Sub

Sub cargar_datos()

    Dim rs1 As ADODB.Recordset
    Dim oGenx As COMDCredito.DCOMCreditos
    Set rs1 = New ADODB.Recordset
    Set oGenx = New COMDCredito.DCOMCreditos

    Set rs1 = oGenx.DevolverDatos_Cabecera_reporte(CDate(txNotaIni.Text))
    Set oGenx = Nothing

     If Not rs1.EOF And Not rs1.BOF Then
          est = 1
        activar_lbl (True)
        
        Me.lblcab1.Caption = rs1("Num_Pro")
        Me.lblcab2.Caption = rs1("nImpToTS")
        Me.lblcab3.Caption = rs1("nImpToTD")
        Me.lblsoles.Caption = rs1("nImpToTS1")
        Me.lbldolar.Caption = rs1("nImpToTD1")
        Me.lblreg.Caption = rs1("Num_Pro")
        
        Cargar_grilla (CDbl(rs1("id")))
           If rs1("Num_NoPro") > 0 Then
                cargar_datos1 (CDbl(rs1("id")))
                Cargar_grilla1 (CDbl(rs1("id")))
            Else
                 SSTAB.Tab = 0
                 SSTAB.TabVisible(1) = False
                 Me.Mshbco1.Clear
            End If
            
            If rs1("Num_Sobra") > 0 Then
                cargar_datos2 (txNotaIni.Text)
                Cargar_grilla2 (txNotaIni.Text)
            Else
                SSTAB.Tab = 0
                SSTAB.TabVisible(2) = False
                Me.Mshbco2.Clear
            End If
    Else
      est = 0
    End If
    rs1.Close
    Set rs1 = Nothing
End Sub

Sub cargar_datos1(ByVal cod_det As Integer)

    Dim rs2 As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Set rs2 = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos

    Set rs2 = oGen.DevolverDatos_Cabecera_108121_NoProcesados(cod_det)
    Set oGen = Nothing

     If Not rs2.EOF And Not rs2.BOF Then
        activar_lbl1 (True)
        Me.lblimpdx.Caption = rs2("nImpToTD")
        Me.lblimpsx.Caption = rs2("nImpToTS")
        Me.lblnumx.Caption = rs2("nNumReg")
     End If

    rs2.Close
    Set rs2 = Nothing
End Sub

Sub cargar_datos1_Corresponsalia(ByVal cod_det As Double)

    Dim rs2c As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Set rs2c = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos

    Set rs2c = oGen.DevolverDatos_Cabecera_108121_NoProcesados_Corresponsalia(cod_det)
    Set oGen = Nothing

     If Not rs2c.EOF And Not rs2c.BOF Then
        activar_lbl1 (True)
        Me.lblimpdx.Caption = rs2c("nImpToTD")
        Me.lblimpsx.Caption = rs2c("nImpToTS")
        Me.lblnumx.Caption = rs2c("nNumReg")
     End If

    rs2c.Close
    Set rs2c = Nothing
End Sub

Sub cargar_datos2(ByVal dfpago As String)

    Dim Rs3 As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Set Rs3 = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos

    Set Rs3 = oGen.DevolverDatos_Cabecera_Sobrante(CDate(dfpago))
    Set oGen = Nothing

     If Not Rs3.EOF And Not Rs3.BOF Then
        activar_lbl2 (True)
        Me.lbldol3.Caption = Rs3("nImpToTD")
        Me.lblsoles3.Caption = Rs3("nImpToTS")
        Me.lblnumx3.Caption = Rs3("nNumReg")
     End If

    Rs3.Close
    Set Rs3 = Nothing
End Sub

Sub cargar_datos2_Corresponsalia(ByVal dfpago As String)

    Dim rs3c As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Set rs3c = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos

    Set rs3c = oGen.DevolverDatos_Cabecera_Sobrante_Corresponsalia(CDate(dfpago))
    Set oGen = Nothing

     If Not rs3c.EOF And Not rs3c.BOF Then
        activar_lbl2 (True)
        Me.lbldol3.Caption = rs3c("nImpToTD")
        Me.lblsoles3.Caption = rs3c("nImpToTS")
        Me.lblnumx3.Caption = rs3c("nNumReg")
     End If

    rs3c.Close
    Set rs3c = Nothing
End Sub

Private Sub Form_Load()
llenar_cbo
ConfigurarMShComite
ConfigurarMShComite1
ConfigurarMShComite2
sCadImpre = ""
Me.SSTAB.TabVisible(1) = False
Me.SSTAB.TabVisible(2) = False
est = 0
End Sub

Sub llenar_cbo()
cbotipo.Clear
cbotipo.AddItem "Cobros", 0
cbotipo.AddItem "Pagos", 1
cbotipo.ListIndex = 0

End Sub

Sub cargar_datosCorres()

    Dim rs1d As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Set rs1d = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos

    Set rs1d = oGen.DevolverDatos_Cabecera_reporte_Corresponsalia(CDate(txNotaIni.Text))
    Set oGen = Nothing

    
    If Not rs1d.EOF And Not rs1d.BOF Then
        est = 1
        activar_lbl (True)
        Me.lblcab1.Caption = rs1d("Num_Pro")
        Me.lblcab2.Caption = rs1d("nImpToTS")
        Me.lblcab3.Caption = rs1d("nImpToTD")
        Me.lblsoles.Caption = rs1d("nImpToTS1")
        Me.lbldolar.Caption = rs1d("nImpToTD1")
        Me.lblreg.Caption = rs1d("Num_Pro")
        
        Cargar_grillaCorresponsalia (CDbl(rs1d("id")))
           If rs1d("Num_NoPro") > 0 Then
                cargar_datos1_Corresponsalia (CDbl(rs1d("id")))
                Cargar_grilla1_Corresponsalia (CDbl(rs1d("id")))
            End If
            
            If rs1d("Num_Sobra") > 0 Then
                cargar_datos2_Corresponsalia (txNotaIni.Text)
                Cargar_grilla2_Corresponsalia (txNotaIni.Text)
            End If
    Else
    est = 0
    End If
    rs1d.Close
    Set rs1d = Nothing
End Sub

Private Sub txNotaIni_KeyPress(KeyAscii As Integer)
  KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cmd_buscar.SetFocus
    End If
End Sub
