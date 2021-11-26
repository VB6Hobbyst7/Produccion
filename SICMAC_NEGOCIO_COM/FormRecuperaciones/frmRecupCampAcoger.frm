VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRecupCampAcoger 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acoger Crédito a Campaña de Recuperaciones"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9645
   Icon            =   "frmRecupCampAcoger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAcoger 
      Caption         =   "Acoger"
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
      Left            =   5760
      TabIndex        =   7
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
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
      Left            =   8400
      TabIndex        =   6
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
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
      Left            =   7080
      TabIndex        =   5
      Top             =   8040
      Width           =   1215
   End
   Begin TabDlg.SSTab sstAcoger 
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Campaña"
      TabPicture(0)   =   "frmRecupCampAcoger.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblMontoPerdonar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblMontoPagar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ActXCtaCred"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdBuscar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraDatosCred"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraCampana"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraDescuento"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.Frame fraDescuento 
         Caption         =   "Descuento en cuotas vencidas"
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
         Height          =   3375
         Left            =   120
         TabIndex        =   8
         Top             =   3960
         Width           =   9255
         Begin SICMACT.FlexEdit feCuotasVenc 
            Height          =   1395
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   9075
            _ExtentX        =   16007
            _ExtentY        =   2461
            Cols0           =   13
            HighLight       =   1
            EncabezadosNombres=   "-Nro--#-Fecha-Atraso-Capital-Interés-Mora-Gastos-Int Gracia-Int. Icv-Total"
            EncabezadosAnchos=   "0-0-300-400-1000-700-900-900-900-900-900-900-1100"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-2-X-X-X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-4-0-0-0-0-0-0-0-0-0-0"
            BackColor       =   16777215
            EncabezadosAlineacion=   "C-L-L-C-C-C-R-R-R-R-R-R-C"
            FormatosEdit    =   "0-0-0-2-5-3-2-2-2-2-2-2-2"
            SelectionMode   =   1
            lbEditarFlex    =   -1  'True
            TipoBusqueda    =   6
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
            CellBackColor   =   16777215
         End
         Begin SICMACT.FlexEdit feCuotasVencAct 
            Height          =   1395
            Left            =   120
            TabIndex        =   35
            Top             =   1920
            Width           =   9075
            _ExtentX        =   16007
            _ExtentY        =   2461
            Cols0           =   11
            HighLight       =   1
            EncabezadosNombres=   "-#-Fecha-Atraso-Capital-Interés-Mora-Gastos-Int. Gracia-Int. Icv-Total"
            EncabezadosAnchos=   "0-700-1000-700-900-900-900-900-900-900-1100"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
            BackColor       =   16777215
            EncabezadosAlineacion=   "C-C-C-C-R-R-R-R-R-R-C"
            FormatosEdit    =   "0-2-5-3-2-2-2-2-2-2-2"
            lbEditarFlex    =   -1  'True
            TipoBusqueda    =   6
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
            CellBackColor   =   16777215
         End
         Begin VB.Label Label12 
            Caption         =   "Cuota a Perdonar:"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1680
            Width           =   1575
         End
      End
      Begin VB.Frame fraCampana 
         Caption         =   "Campaña"
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
         Height          =   1815
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   9255
         Begin VB.ComboBox cmbSubCampana 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   720
            Width           =   6375
         End
         Begin VB.ComboBox cmbCampana 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   360
            Width           =   4455
         End
         Begin MSComCtl2.UpDown udPorcCapital 
            Height          =   300
            Left            =   2640
            TabIndex        =   42
            Top             =   1400
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin SICMACT.EditMoney txtCapital 
            Height          =   300
            Left            =   2040
            TabIndex        =   43
            Top             =   1400
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0.00"
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udPorcInt 
            Height          =   300
            Left            =   3960
            TabIndex        =   44
            Top             =   1400
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin SICMACT.EditMoney txtInt 
            Height          =   300
            Left            =   3360
            TabIndex        =   45
            Top             =   1400
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0.00"
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udPorcMora 
            Height          =   300
            Left            =   5400
            TabIndex        =   46
            Top             =   1400
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin SICMACT.EditMoney txtMora 
            Height          =   300
            Left            =   4800
            TabIndex        =   47
            Top             =   1400
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0.00"
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udPorcGasto 
            Height          =   300
            Left            =   6840
            TabIndex        =   48
            Top             =   1400
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin SICMACT.EditMoney txtGasto 
            Height          =   300
            Left            =   6240
            TabIndex        =   49
            Top             =   1400
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0.00"
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udPorcICV 
            Height          =   300
            Left            =   8280
            TabIndex        =   50
            Top             =   1400
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin SICMACT.EditMoney txtICV 
            Height          =   300
            Left            =   7680
            TabIndex        =   51
            Top             =   1400
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0.00"
            Enabled         =   -1  'True
         End
         Begin VB.Label Label18 
            Caption         =   "ICV:"
            Height          =   255
            Left            =   7200
            TabIndex        =   52
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label17 
            Caption         =   "% Descuentos:"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "Capital:"
            Height          =   255
            Left            =   1440
            TabIndex        =   40
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "Int.:"
            Height          =   255
            Left            =   3000
            TabIndex        =   39
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label13 
            Caption         =   "Mora:"
            Height          =   255
            Left            =   4320
            TabIndex        =   38
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label14 
            Caption         =   "Gasto:"
            Height          =   255
            Left            =   5760
            TabIndex        =   37
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label lblMonedaHasta 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3840
            TabIndex        =   29
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label lblMontoHasta 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4440
            TabIndex        =   28
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblMonedaDesde 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1440
            TabIndex        =   27
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label lblMontoDesde 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2040
            TabIndex        =   26
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "-"
            Height          =   255
            Left            =   3600
            TabIndex        =   23
            Top             =   1080
            Width           =   135
         End
         Begin VB.Label Label7 
            Caption         =   "Rango de Pago:"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Campaña:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Sub Campaña:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame fraDatosCred 
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
         ForeColor       =   &H8000000D&
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   9255
         Begin VB.Label lblMoneda 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1080
            TabIndex        =   19
            Top             =   600
            Width           =   495
         End
         Begin VB.Label lblSaldoCap 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1680
            TabIndex        =   18
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblDiasAtraso 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5280
            TabIndex        =   17
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblDOI 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6720
            TabIndex        =   16
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblCuotaVenc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7680
            TabIndex        =   15
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblCliente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1080
            TabIndex        =   14
            Top             =   240
            Width           =   4815
         End
         Begin VB.Label Label4 
            Caption         =   "Cuotas Vencidas:"
            Height          =   255
            Left            =   6360
            TabIndex        =   13
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Días atraso:"
            Height          =   255
            Left            =   4320
            TabIndex        =   12
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "DOI:"
            Height          =   255
            Left            =   6360
            TabIndex        =   11
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Saldo Cap.:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Cliente:"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   375
         Left            =   3840
         TabIndex        =   2
         Top             =   480
         Width           =   495
      End
      Begin SICMACT.ActXCodCta ActXCtaCred 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   661
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
         CMAC            =   "109"
      End
      Begin VB.Label lblMontoPagar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7920
         TabIndex        =   34
         Top             =   7440
         Width           =   1455
      End
      Begin VB.Label lblMontoPerdonar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   33
         Top             =   7440
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Total a pagar:"
         Height          =   255
         Left            =   6840
         TabIndex        =   31
         Top             =   7440
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Monto perdón:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   7440
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmRecupCampAcoger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmRecupCampAcoger
'** Descripción : Formulario para Acoger a creditos a la campaña de recuperaciones
'**               Creado segun TI-ERS035-2015
'** Creación    : WIOR, 20150522 09:00:00 AM
'**********************************************************************************************

Option Explicit
Private i As Integer
Private MatCamp As Variant
Private MatSubCamp As Variant
Private MatCuotaVenc As Variant
Private MatCuotasPerd As Variant
Private ArrPorcentaje As Variant
Private ArrPorcFijo As Variant
Private nIndSubCamp As Integer
Private fnNroCalen As Integer
Private sPersCod As String 'CROB20180813 ERS055-2018
Private bDesctoAdicional As Boolean 'CROB20180813 ERS055-2018

Private Sub ActXCtaCred_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call CargarDatos(Trim(ActXCtaCred.NroCuenta))
End If
End Sub


Private Sub CargarDatos(ByVal pcCtaCod As String)
Dim oDCredito As COMDCredito.DCOMCredito
Dim oRs As ADODB.Recordset

Set oDCredito = New COMDCredito.DCOMCredito
Set oRs = oDCredito.RecuperarDatosCredCampanaRecup(pcCtaCod)
fnNroCalen = 0
If Not (oRs.EOF And oRs.BOF) Then
    lblCliente.Caption = Trim(oRs!cPersNombre)
    sPersCod = Trim(oRs!cPersCod) 'CROB20180813 ERS055-2018
    lblSaldoCap.Caption = Format(Trim(oRs!nSaldo), "###," & String(15, "#") & "#0.00")
    lblDOI.Caption = Trim(oRs!DOI)
    lblCuotaVenc = Trim(oRs!CantCuoVenc)
    lblDiasAtraso.Caption = Trim(oRs!nDiasAtraso)
    lblMoneda.Caption = IIf(Mid(pcCtaCod, 9, 1) = "1", "S/.", "$")
    cmdAcoger.Enabled = True
    fnNroCalen = CInt(oRs!nNroCalen)
    
    Set oDCredito = New COMDCredito.DCOMCredito
    Set oRs = oDCredito.RecuperarCampanaRecupCuotasVencCred(pcCtaCod)
    
    Call CargaCampanas(pcCtaCod)
    
    ReDim MatCuotaVenc(0, 10)
    LimpiaFlex feCuotasVenc
    If Not (oRs.EOF And oRs.BOF) Then
        'ReDim MatCuotaVenc(oRs.RecordCount - 1, 12)'JOEP
        ReDim MatCuotaVenc(oRs.RecordCount - 1, 13) 'JOEP
        
        For i = 0 To oRs.RecordCount - 1
            MatCuotaVenc(i, 0) = CInt(oRs!nCuota)
            MatCuotaVenc(i, 1) = CDate(oRs!dVenc)
            MatCuotaVenc(i, 2) = CInt(oRs!nDiasMora)
            MatCuotaVenc(i, 3) = CCur(oRs!nCapital)
            MatCuotaVenc(i, 4) = CCur(oRs!nCapitaPag)
            MatCuotaVenc(i, 5) = CCur(oRs!nInt)
            MatCuotaVenc(i, 6) = CCur(oRs!nIntPag)
            MatCuotaVenc(i, 7) = CCur(oRs!nMora)
            MatCuotaVenc(i, 8) = CCur(oRs!nMoraPag)
            MatCuotaVenc(i, 9) = CCur(oRs!nGasto)
            MatCuotaVenc(i, 10) = CCur(oRs!nGastoPag)
            MatCuotaVenc(i, 11) = CCur(oRs!nIntGraPend)
            
            MatCuotaVenc(i, 12) = CCur(oRs!nIntICV) 'JOEP
            MatCuotaVenc(i, 13) = CCur(oRs!nIntICVPag) 'JOEP
            
            feCuotasVenc.AdicionaFila
            feCuotasVenc.TextMatrix(i + 1, 1) = CInt(oRs!nCuota)
            feCuotasVenc.TextMatrix(i + 1, 2) = ""
            feCuotasVenc.TextMatrix(i + 1, 3) = CInt(oRs!nCuota)
            feCuotasVenc.TextMatrix(i + 1, 4) = Format(CDate(oRs!dVenc), "dd/mm/yyyy")
            feCuotasVenc.TextMatrix(i + 1, 5) = CInt(oRs!nDiasMora)
            feCuotasVenc.TextMatrix(i + 1, 6) = Format(CCur(oRs!nCapital) - CCur(oRs!nCapitaPag), "#0.00")
            feCuotasVenc.TextMatrix(i + 1, 7) = Format(CCur(oRs!nInt) - CCur(oRs!nIntPag), "#0.00")
            feCuotasVenc.TextMatrix(i + 1, 8) = Format(CCur(oRs!nMora) - CCur(oRs!nMoraPag), "#0.00")
            feCuotasVenc.TextMatrix(i + 1, 9) = Format(CCur(oRs!nGasto) - CCur(oRs!nGastoPag), "#0.00")
            feCuotasVenc.TextMatrix(i + 1, 10) = Format(CCur(oRs!nIntGraPend), "#0.00")
            
            feCuotasVenc.TextMatrix(i + 1, 11) = Format(CCur(oRs!nIntICV) - CCur(oRs!nIntICVPag), "#0.00") 'JOEP
            
            'feCuotasVenc.TextMatrix(i + 1, 11) = Format(CCur(feCuotasVenc.TextMatrix(i + 1, 6)) + CCur(feCuotasVenc.TextMatrix(i + 1, 7)) + CCur(feCuotasVenc.TextMatrix(i + 1, 8)) + CCur(feCuotasVenc.TextMatrix(i + 1, 9) + CCur(feCuotasVenc.TextMatrix(i + 1, 10))), "#0.00")
            feCuotasVenc.TextMatrix(i + 1, 12) = Format(CCur(feCuotasVenc.TextMatrix(i + 1, 6)) + CCur(feCuotasVenc.TextMatrix(i + 1, 7)) + CCur(feCuotasVenc.TextMatrix(i + 1, 8)) + CCur(feCuotasVenc.TextMatrix(i + 1, 9) + CCur(feCuotasVenc.TextMatrix(i + 1, 10)) + CCur(feCuotasVenc.TextMatrix(i + 1, 11))), "#0.00")
            oRs.MoveNext
        Next i
    End If
    
    Call HabilitarControles(True) 'JOEP
Else
    cmdAcoger.Enabled = False
    MsgBox "El crédito no encontrado o no calififica para las campañas de Recuperaciones", vbInformation, "Aviso"
    Call HabilitarControles(False) 'JOEP
End If

Set oDCredito = Nothing
End Sub

Private Sub cmbCampana_Click()
Dim oDCredito As COMDCredito.DCOMCredito
Dim oRs As ADODB.Recordset
Dim nIndice As Integer
Dim nNumMax As Integer
Dim nValorMaxCampAnual As Integer 'CROB20180813 ERS055-2018
Dim nValorDesctoAdicionalCamp As Byte 'CROB20180813 ERS055-2018
Dim nValorDesctoTotal As Byte 'CROB20180813 ERS055-2018
nIndice = -1

If Trim(cmbCampana.Text) <> "" And Len(Trim(ActXCtaCred.NroCuenta)) = 18 Then
    Set oDCredito = New COMDCredito.DCOMCredito

    nIndice = ObtenerIndice(CInt(Trim(Right(cmbCampana.Text, 6))), MatCamp, 0)
    
    If nIndice = -1 Then Exit Sub
    
    'Obtener Cantidad de veces que acogio campaña un cliente durante el año CROB20180813 ERS055-2018
    Set oRs = oDCredito.RecuperarCantVecesAcogidasCampanasXcliente(sPersCod)
    nNumMax = 0
    If Not (oRs.EOF And oRs.BOF) Then
        nNumMax = CInt(oRs!nNum)
    End If
    'CROB20180813 ERS055-2018
    
'Comentado por CROB20180814
'   Set oRs = oDCredito.RecuperarCampanaRecupCantVeces(CLng(MatCamp(nIndice, 0)), Trim(Me.ActXCtaCred.NroCuenta))
'    nNumMax = 0
'    If Not (oRs.EOF And oRs.BOF) Then
'        nNumMax = CInt(oRs!nNum)
'    End If
'Comentado por CROB20180814
    
    'ReDim MatSubCamp(0, 13)'joep
    ReDim MatSubCamp(0, 14) 'joep
    cmbSubCampana.Clear
    LimpiaFlex feCuotasVencAct
    lblMontoPagar.Caption = "0.00 "
    lblMontoPerdonar.Caption = "0.00 "
    lblMontoDesde.Caption = "0.00 "
    lblMontoHasta.Caption = "0.00 "
    
    nValorMaxCampAnual = oDCredito.ObtenerValorMaxCampanasAnual(Year(gdFecSis))!nCantMaxAnual 'CROB20180813 ERS055-2018
    bDesctoAdicional = IIf((nNumMax >= nValorMaxCampAnual), True, False) 'CROB20180813 ERS055-2018
    
    If nValorMaxCampAnual > 0 Then 'CROB20180813 ERS055-2018
        nValorDesctoAdicionalCamp = oDCredito.ObtenerValorCantAdicionalDesctosCampana(sPersCod)!nCantAdicional 'CROB20180813 ERS055-2018
        nValorDesctoTotal = (nValorMaxCampAnual + nValorDesctoAdicionalCamp) 'CROB20180813 ERS055-2018
    
        'If nNumMax > CInt(MatCamp(nIndice, 5)) Then 'Comentado por CROB20180813
            'MsgBox "El crédito a superado el número máximo de veces que puedes acorgerse establecidas por la campaña.", vbInformation, "Aviso" 'Comentado por CROB20180813
        If nNumMax >= nValorDesctoTotal Then 'CROB20180813 ERS055-2018
            MsgBox "El cliente a superado el número máximo de veces que puede acorgerse por campaña." & Chr(10) & "El cliente ha acogido " & nNumMax & " veces campaña.", vbInformation, "Aviso" 'CROB20180813 ERS055-2018
        Else
            Set oRs = oDCredito.RecuperarCampanaRecupSubCampCred(CLng(MatCamp(nIndice, 0)), Trim(Me.ActXCtaCred.NroCuenta))
            
            If Not (oRs.EOF And oRs.BOF) Then
                'ReDim MatSubCamp(oRs.RecordCount - 1, 13)'joep
                ReDim MatSubCamp(oRs.RecordCount - 1, 14) 'joep
                For i = 0 To oRs.RecordCount - 1
                    MatSubCamp(i, 0) = oRs!nId
                    MatSubCamp(i, 1) = oRs!nIdSubCamp
                    MatSubCamp(i, 2) = Trim(oRs!Descrip)
                    MatSubCamp(i, 3) = CInt(oRs!nDiasAtrasoIni)
                    MatSubCamp(i, 4) = CInt(oRs!nDiasAtrasoFin)
                    MatSubCamp(i, 5) = CInt(oRs!nConsidera)
                    MatSubCamp(i, 6) = CInt(oRs!nTpoConsidera)
                    MatSubCamp(i, 7) = CCur(oRs!nDescCap)
                    MatSubCamp(i, 8) = CCur(oRs!nDescInt)
                    MatSubCamp(i, 9) = CCur(oRs!nDescMora)
                    MatSubCamp(i, 10) = CCur(oRs!nDescGasto)
                    MatSubCamp(i, 11) = CBool(oRs!bVencidos)
                    MatSubCamp(i, 12) = CBool(oRs!bTransferidos)
                    MatSubCamp(i, 13) = CInt(oRs!nTpoGarantia)
                    
                    MatSubCamp(i, 14) = CInt(oRs!nDescIcv) 'JOEP
                    
                    cmbSubCampana.AddItem Trim(oRs!Descrip) & Space(75) & oRs!nIdSubCamp
                    oRs.MoveNext
                Next i
            End If
        End If
    Else 'CROB20180813 ERS055-2018
        MsgBox "El valor maximo de descuentos por año debe ser mayor a 0 (cero).", vbInformation, "Aviso" 'CROB20180813 ERS055-2018
    End If 'CROB20180813 ERS055-2018
End If
End Sub
Private Sub CargaCampanas(ByVal pcCtaCod As String)
Dim oDCredito As COMDCredito.DCOMCredito
Dim oRs As ADODB.Recordset

Set oDCredito = New COMDCredito.DCOMCredito
Set oRs = oDCredito.RecuperarCampanaRecupActivas(gdFecSis, pcCtaCod)

ReDim MatCamp(0, 5)
cmbCampana.Clear
If Not (oRs.EOF And oRs.BOF) Then
    ReDim MatCamp(oRs.RecordCount - 1, 5)
    For i = 0 To oRs.RecordCount - 1
        MatCamp(i, 0) = oRs!nId
        MatCamp(i, 1) = Trim(oRs!cNombre)
        MatCamp(i, 2) = Trim(oRs!cAprobado)
        MatCamp(i, 3) = CDate(oRs!dfechaini)
        MatCamp(i, 4) = CDate(oRs!dfechafin)
        MatCamp(i, 5) = CInt(oRs!nNumMax)
        cmbCampana.AddItem Trim(oRs!cNombre) & Space(75) & oRs!nId
        oRs.MoveNext
    Next i
End If
End Sub

Private Sub cmbSubCampana_Click()
nIndSubCamp = -1
LimpiaFlex feCuotasVencAct

txtCapital.Text = "0.00"
txtInt.Text = "0.00"
txtMora.Text = "0.00"
txtGasto.Text = "0.00"
txtICV.Text = "0.00" 'joep
    
If Trim(cmbSubCampana.Text) <> "" And Len(Trim(ActXCtaCred.NroCuenta)) = 18 Then
    nIndSubCamp = ObtenerIndice(CInt(Trim(Right(cmbSubCampana.Text, 6))), MatSubCamp, 1)
    If nIndSubCamp = -1 Then Exit Sub
    
    'ReDim ArrPorcentaje(3)'JOEP
    ReDim ArrPorcentaje(4) 'JOEP
    ArrPorcentaje(0) = MatSubCamp(nIndSubCamp, 7)
    ArrPorcentaje(1) = MatSubCamp(nIndSubCamp, 8)
    ArrPorcentaje(2) = MatSubCamp(nIndSubCamp, 9)
    ArrPorcentaje(3) = MatSubCamp(nIndSubCamp, 10)
    ArrPorcentaje(4) = MatSubCamp(nIndSubCamp, 14) 'JOEP
    
    'ReDim ArrPorcFijo(3)'JOEP
    ReDim ArrPorcFijo(4) 'JOEP
    ArrPorcFijo(0) = MatSubCamp(nIndSubCamp, 7)
    ArrPorcFijo(1) = MatSubCamp(nIndSubCamp, 8)
    ArrPorcFijo(2) = MatSubCamp(nIndSubCamp, 9)
    ArrPorcFijo(3) = MatSubCamp(nIndSubCamp, 10)
    ArrPorcFijo(4) = MatSubCamp(nIndSubCamp, 14) 'JOEP
    
    txtCapital.Text = Format(ArrPorcentaje(0), "#0.00")
    txtInt.Text = Format(ArrPorcentaje(1), "#0.00")
    txtMora.Text = Format(ArrPorcentaje(2), "#0.00")
    txtGasto.Text = Format(ArrPorcentaje(3), "#0.00")
    
    txtICV.Text = Format(ArrPorcentaje(4), "#0.00") 'JOEP
    
    Call GenerarCuotasAPerdonar
    
'    nDescCap = 0
'    nDescInt = 0
'    nDescMora = 0
'    nDescGasto = 0
'
'    For i = 0 To UBound(MatCuotaVenc)
'        feCuotasVenc.TextMatrix(i + 1, 2) = ""
'        nDescCap = 0
'        nDescInt = 0
'        nDescMora = 0
'        nDescGasto = 0
'
'        bAgreCuota = False
'        If CInt(MatSubCamp(nIndSubCamp, 6)) = 1 And i < CInt(MatSubCamp(nIndSubCamp, 5)) Then 'Cuotas Pagadas y Consideracion
'            bAgreCuota = True
'        ElseIf CInt(MatSubCamp(nIndSubCamp, 6)) = 2 And (CInt(MatSubCamp(nIndSubCamp, 5)) - 1) < (UBound(MatCuotaVenc) - i) Then 'Cuotas Pendientes y Consideracion
'            bAgreCuota = True
'        Else
'            bAgreCuota = False
'        End If
'
'        If bAgreCuota Then
'            feCuotasVenc.TextMatrix(i + 1, 2) = "1"
'            feCuotasVencAct.AdicionaFila
'            feCuotasVencAct.TextMatrix(i + 1, 1) = CInt(MatCuotaVenc(i, 0))
'            feCuotasVencAct.TextMatrix(i + 1, 2) = Format(CDate(MatCuotaVenc(i, 1)), "dd/mm/yyyy")
'            feCuotasVencAct.TextMatrix(i + 1, 3) = CInt(MatCuotaVenc(i, 2))
'
'            nDescCap = Round((CCur(MatCuotaVenc(i, 3)) - CCur(MatCuotaVenc(i, 4))) * (MatPorcentaje(0, 0) / 100), 2)
'            feCuotasVencAct.TextMatrix(i + 1, 4) = Format(CCur(MatCuotaVenc(i, 3)) - CCur(MatCuotaVenc(i, 4) - nDescCap), "#0.00")
'            If nDescCap > 0 Then
'                feCuotasVencAct.col = 4
'                feCuotasVencAct.CellForeColor = vbRed
'            End If
'
'            nDescInt = Round((CCur(MatCuotaVenc(i, 5)) - CCur(MatCuotaVenc(i, 6))) * (MatPorcentaje(0, 1) / 100), 2)
'            feCuotasVencAct.TextMatrix(i + 1, 5) = Format(CCur(MatCuotaVenc(i, 5)) - CCur(MatCuotaVenc(i, 6)) - nDescInt, "#0.00")
'            If nDescInt > 0 Then
'                feCuotasVencAct.col = 5
'                feCuotasVencAct.CellForeColor = vbRed
'            End If
'
'            nDescMora = Round((CCur(MatCuotaVenc(i, 7)) - CCur(MatCuotaVenc(i, 8))) * (MatPorcentaje(0, 2) / 100), 2)
'            feCuotasVencAct.TextMatrix(i + 1, 6) = Format(CCur(MatCuotaVenc(i, 7)) - CCur(MatCuotaVenc(i, 8)) - nDescMora, "#0.00")
'            If nDescMora > 0 Then
'                feCuotasVencAct.col = 6
'                feCuotasVencAct.CellForeColor = vbRed
'            End If
'
'            nDescGasto = Round((CCur(MatCuotaVenc(i, 9)) - CCur(MatCuotaVenc(i, 10))) * (MatPorcentaje(0, 3) / 100), 2)
'            feCuotasVencAct.TextMatrix(i + 1, 7) = Format(CCur(MatCuotaVenc(i, 9)) - CCur(MatCuotaVenc(i, 10)) - nDescGasto, "#0.00")
'            If nDescGasto > 0 Then
'                feCuotasVencAct.col = 7
'                feCuotasVencAct.CellForeColor = vbRed
'            End If
'
'            feCuotasVencAct.TextMatrix(i + 1, 8) = Format(CCur(MatCuotaVenc(i, 11)), "#0.00")
'            feCuotasVencAct.TextMatrix(i + 1, 9) = Format(CCur(feCuotasVencAct.TextMatrix(i + 1, 4)) + CCur(feCuotasVencAct.TextMatrix(i + 1, 5)) + CCur(feCuotasVencAct.TextMatrix(i + 1, 6)) + CCur(feCuotasVencAct.TextMatrix(i + 1, 7)) + CCur(feCuotasVencAct.TextMatrix(i + 1, 8)), "#0.00")
'            feCuotasVencAct.col = 9
'        End If
'    Next i
'
'    Call CalculaMontos
End If
End Sub

Private Sub GenerarCuotasAPerdonar()
Dim nDescCap As Double
Dim nDescInt As Double
Dim nDescMora As Double
Dim nDescGasto As Double
Dim bAgreCuota As Boolean
Dim nDescIcv As Double 'joep
If Trim(cmbSubCampana.Text) <> "" And Len(Trim(ActXCtaCred.NroCuenta)) = 18 Then
    nIndSubCamp = ObtenerIndice(CInt(Trim(Right(cmbSubCampana.Text, 6))), MatSubCamp, 1)
    If nIndSubCamp = -1 Then Exit Sub
    
    nDescCap = 0
    nDescInt = 0
    nDescMora = 0
    nDescGasto = 0
    LimpiaFlex feCuotasVencAct
    For i = 0 To UBound(MatCuotaVenc)
        feCuotasVenc.TextMatrix(i + 1, 2) = ""
        nDescCap = 0
        nDescInt = 0
        nDescMora = 0
        nDescGasto = 0
        
        bAgreCuota = False
        If CInt(MatSubCamp(nIndSubCamp, 6)) = 1 And i < CInt(MatSubCamp(nIndSubCamp, 5)) Then 'Cuotas Pagadas y Consideracion
            bAgreCuota = True
        ElseIf CInt(MatSubCamp(nIndSubCamp, 6)) = 2 And (CInt(MatSubCamp(nIndSubCamp, 5)) - 1) < (UBound(MatCuotaVenc) - i) Then 'Cuotas Pendientes y Consideracion
            bAgreCuota = True
        Else
            bAgreCuota = False
        End If
        
        If bAgreCuota Then
            feCuotasVenc.TextMatrix(i + 1, 2) = "1"
            feCuotasVencAct.AdicionaFila
            feCuotasVencAct.row = i + 1
            
            feCuotasVencAct.TextMatrix(i + 1, 1) = CInt(MatCuotaVenc(i, 0))
            feCuotasVencAct.TextMatrix(i + 1, 2) = Format(CDate(MatCuotaVenc(i, 1)), "dd/mm/yyyy")
            feCuotasVencAct.TextMatrix(i + 1, 3) = CInt(MatCuotaVenc(i, 2))
            
            nDescCap = Round(Format((CDbl(MatCuotaVenc(i, 3)) - CDbl(MatCuotaVenc(i, 4))) * (ArrPorcentaje(0) / 100), "#0.00"), 2)
            feCuotasVencAct.TextMatrix(i + 1, 4) = Format(CCur(MatCuotaVenc(i, 3)) - CCur(MatCuotaVenc(i, 4) - nDescCap), "#0.00")
             
            feCuotasVencAct.Col = 4
            feCuotasVencAct.CellForeColor = vbBlack
            If nDescCap > 0 Then
                feCuotasVencAct.CellForeColor = vbRed
            End If
            
            nDescInt = Round(Format((CDbl(MatCuotaVenc(i, 5)) - CDbl(MatCuotaVenc(i, 6))) * (ArrPorcentaje(1) / 100), "#0.00"), 2)
            feCuotasVencAct.TextMatrix(i + 1, 5) = Format(CCur(MatCuotaVenc(i, 5)) - CCur(MatCuotaVenc(i, 6)) - nDescInt, "#0.00")
            
            feCuotasVencAct.Col = 5
            feCuotasVencAct.CellForeColor = vbBlack
            If nDescInt > 0 Then
                feCuotasVencAct.CellForeColor = vbRed
            End If
            
            nDescMora = Round(Format((CDbl(MatCuotaVenc(i, 7)) - CDbl(MatCuotaVenc(i, 8))) * (ArrPorcentaje(2) / 100), "#0.00"), 2)
            feCuotasVencAct.TextMatrix(i + 1, 6) = Format(CCur(MatCuotaVenc(i, 7)) - CCur(MatCuotaVenc(i, 8)) - nDescMora, "#0.00")
            
            feCuotasVencAct.Col = 6
            feCuotasVencAct.CellForeColor = vbBlack
            If nDescMora > 0 Then
                feCuotasVencAct.CellForeColor = vbRed
            End If
            
            nDescGasto = Round(Format((CDbl(MatCuotaVenc(i, 9)) - CDbl(MatCuotaVenc(i, 10))) * (ArrPorcentaje(3) / 100), "#0.00"), 2)
            feCuotasVencAct.TextMatrix(i + 1, 7) = Format(CCur(MatCuotaVenc(i, 9)) - CCur(MatCuotaVenc(i, 10)) - nDescGasto, "#0.00")
            
            feCuotasVencAct.Col = 7
            feCuotasVencAct.CellForeColor = vbBlack
            If nDescGasto > 0 Then
                feCuotasVencAct.CellForeColor = vbRed
            End If
            
            'JOEP
            nDescIcv = Round(Format((CDbl(MatCuotaVenc(i, 12)) - CDbl(MatCuotaVenc(i, 13))) * (ArrPorcentaje(4) / 100), "#0.00"), 2)
            feCuotasVencAct.TextMatrix(i + 1, 9) = Format(CCur(MatCuotaVenc(i, 12)) - CCur(MatCuotaVenc(i, 13)) - nDescIcv, "#0.00")
            
            feCuotasVencAct.Col = 9
            feCuotasVencAct.CellForeColor = vbBlack
            If nDescIcv > 0 Then
                feCuotasVencAct.CellForeColor = vbRed
            End If
            'JOEP
            
            feCuotasVencAct.TextMatrix(i + 1, 8) = Format(CCur(MatCuotaVenc(i, 11)), "#0.00")
                        
            'feCuotasVencAct.TextMatrix(i + 1, 9) = Format(CCur(feCuotasVencAct.TextMatrix(i + 1, 4)) + CCur(feCuotasVencAct.TextMatrix(i + 1, 5)) + CCur(feCuotasVencAct.TextMatrix(i + 1, 6)) + CCur(feCuotasVencAct.TextMatrix(i + 1, 7)) + CCur(feCuotasVencAct.TextMatrix(i + 1, 8)), "#0.00")'joep
            feCuotasVencAct.TextMatrix(i + 1, 10) = Format(CCur(feCuotasVencAct.TextMatrix(i + 1, 4)) + CCur(feCuotasVencAct.TextMatrix(i + 1, 5)) + CCur(feCuotasVencAct.TextMatrix(i + 1, 6)) + CCur(feCuotasVencAct.TextMatrix(i + 1, 7)) + CCur(feCuotasVencAct.TextMatrix(i + 1, 8)) + CCur(feCuotasVencAct.TextMatrix(i + 1, 9)), "#0.00") 'joep
            'feCuotasVencAct.Col = 9'joep
            feCuotasVencAct.Col = 10 'joep
        End If
    Next i
    
    Call CalculaMontos
End If
End Sub

Private Sub cmdAcoger_Click()
If ValidaDatos Then
    If MsgBox("Estas seguro de acoger al crédito con los datos respectivos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    GrabarDatos
End If
End Sub


Private Sub cmdBuscar_Click()
Dim oPersona As COMDPersona.UCOMPersona
Dim loPersCreditos As COMDCredito.DCOMCredito
Dim lrCreditos As New ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

LimpiarDatos

On Error GoTo ControlError

Set oPersona = Nothing
Set oPersona = New COMDPersona.UCOMPersona
Set oPersona = frmBuscaPersona.Inicio
If oPersona Is Nothing Then Exit Sub

If Trim(oPersona.sPersCod) <> "" Then
    Set loPersCreditos = New COMDCredito.DCOMCredito
    Set lrCreditos = loPersCreditos.CreditosHabilitadosCampRecup(oPersona.sPersCod)
    Set loPersCreditos = Nothing
    sPersCod = oPersona.sPersCod 'CROB20180813 ERS055-2018
End If

If Not (lrCreditos.EOF And lrCreditos.BOF) Then
    Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(oPersona.sPersNombre, lrCreditos)
    If loCuentas.sCtaCod <> "" Then
        ActXCtaCred.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        ActXCtaCred.Enabled = False
        Call ActXCtaCred_KeyPress(13)
    End If
Else
    MsgBox "Cliente no cuenta con créditos aptos para las Campañas de Recuperaciones.", vbInformation, "Aviso"
End If
Set loCuentas = Nothing
Exit Sub
ControlError:
MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub
Private Sub LimpiarDatos()
ActXCtaCred.Age = ""
ActXCtaCred.Prod = ""
ActXCtaCred.Cuenta = ""
ActXCtaCred.Enabled = True
ActXCtaCred.SetFocus

lblCliente.Caption = ""
lblSaldoCap.Caption = ""
lblDOI.Caption = ""
lblCuotaVenc = ""
lblDiasAtraso.Caption = ""
lblMoneda.Caption = ""
lblMontoPagar.Caption = ""
lblMontoPerdonar.Caption = ""

lblMonedaDesde.Caption = ""
lblMontoDesde.Caption = ""
lblMonedaHasta.Caption = ""
lblMontoHasta.Caption = ""

LimpiaFlex feCuotasVenc
LimpiaFlex feCuotasVencAct

cmbCampana.Clear
cmbSubCampana.Clear

cmdAcoger.Enabled = False
fnNroCalen = 0
ReDim MatCamp(0, 0)
ReDim MatSubCamp(0, 0)
ReDim MatCuotaVenc(0, 0)
ReDim MatCuotasPerd(0, 0)

txtCapital.Text = "0.00"
txtInt.Text = "0.00"
txtMora.Text = "0.00"
txtGasto.Text = "0.00"
txtICV.Text = "0.00" 'JOEP

'ReDim ArrPorcentaje(3)'JOEP
ReDim ArrPorcentaje(4) 'JOEP
ArrPorcentaje(0) = 0
ArrPorcentaje(1) = 0
ArrPorcentaje(2) = 0
ArrPorcentaje(3) = 0
ArrPorcentaje(4) = 0 'joep

'ReDim ArrPorcFijo(3)'JOEP
ReDim ArrPorcFijo(4) 'JOEP
ArrPorcFijo(0) = 0
ArrPorcFijo(1) = 0
ArrPorcFijo(2) = 0
ArrPorcFijo(3) = 0
ArrPorcFijo(4) = 0 'JOEP

feCuotasVencAct.Col = 4
feCuotasVencAct.CellForeColor = vbBlack
feCuotasVencAct.Col = 5
feCuotasVencAct.CellForeColor = vbBlack
feCuotasVencAct.Col = 6
feCuotasVencAct.CellForeColor = vbBlack
feCuotasVencAct.Col = 7
feCuotasVencAct.CellForeColor = vbBlack
End Sub

Private Sub cmdCancelar_Click()
LimpiarDatos
Call HabilitarControles(False) 'JOEP
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub feCuotasVenc_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim sColumnas() As String
sColumnas = Split(feCuotasVenc.ColumnasAEditar, "-")
If sColumnas(pnCol) = "X" Or pnCol = 2 Then
   Cancel = False
   SendKeys "{Tab}", True
   Exit Sub
End If
End Sub

Private Sub feCuotasVencAct_OnCellChange(pnRow As Long, pnCol As Long)
Call CalculaMontos
End Sub

Private Sub feCuotasVencAct_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim sColumnas() As String
Dim nIndexCuota As Integer
Dim nIndexSC As Integer
Dim nMontoMinAPag As Double
Dim nMonto As Double
Dim nMontoMaxPerd As Double
Dim nPorcentaje As Double
Dim nDif As Double

sColumnas = Split(feCuotasVencAct.ColumnasAEditar, "-")
If sColumnas(pnCol) = "X" Then
   Cancel = False
   SendKeys "{Tab}", True
   Exit Sub
End If

nIndexCuota = -1
nIndexSC = -1
nMontoMinAPag = 0
nMonto = 0
nMontoMaxPerd = 0
nDif = 0
nPorcentaje = 0
If pnCol = 4 Or pnCol = 5 Or pnCol = 6 Or pnCol = 7 Then
    nIndexCuota = ObtenerIndice(CInt(feCuotasVencAct.TextMatrix(pnRow, 1)), MatCuotaVenc, 0)
    nIndexSC = ObtenerIndice(CInt(Trim(Right(cmbSubCampana.Text, 6))), MatSubCamp, 1)
    
    Select Case pnCol
        Case 4:
                nMonto = MatCuotaVenc(nIndexCuota, 3) - MatCuotaVenc(nIndexCuota, 4)
                nPorcentaje = MatSubCamp(nIndexSC, 7) / 100
                nMontoMaxPerd = Round(nMonto * nPorcentaje, 2)
        Case 5:
                nMonto = MatCuotaVenc(nIndexCuota, 5) - MatCuotaVenc(nIndexCuota, 6)
                nPorcentaje = MatSubCamp(nIndexSC, 8) / 100
                nMontoMaxPerd = Round(nMonto * (MatSubCamp(nIndexSC, 8) / 100), 2)
        Case 6:
                nMonto = MatCuotaVenc(nIndexCuota, 7) - MatCuotaVenc(nIndexCuota, 8)
                nPorcentaje = MatSubCamp(nIndexSC, 9) / 100
                nMontoMaxPerd = Round(nMonto * nPorcentaje, 2)
        Case 7:
                nMonto = MatCuotaVenc(nIndexCuota, 9) - MatCuotaVenc(nIndexCuota, 10)
                nPorcentaje = MatSubCamp(nIndexSC, 10) / 100
                nMontoMaxPerd = Round(nMonto * nPorcentaje, 2)
    End Select
    
    nMontoMinAPag = nMonto - nMontoMaxPerd
    nDif = Round(nMonto - CDbl(feCuotasVencAct.TextMatrix(pnRow, pnCol)), 2)

    feCuotasVencAct.Col = pnCol
    
    If CDbl(feCuotasVencAct.TextMatrix(pnRow, pnCol)) < nMonto Then
        If nDif <= nMontoMaxPerd Then
            feCuotasVencAct.CellForeColor = vbRed
        Else
            MsgBox "- Monto mínimo a pagar de " & _
                    IIf(pnCol = 4, "''Capital''", IIf(pnCol = 5, "''Interes Comp.''", IIf(pnCol = 6, "''Mora''", "''Gasto''"))) & _
                    " de la cuota " & Trim(feCuotasVencAct.TextMatrix(pnRow, 1)) & " es " & Format(nMontoMinAPag, "#0.00") & Chr(10) & _
                    "- Porcentaje Máximo de Descuento es " & Format(nPorcentaje, "#0.00%") & Chr(10) & _
                    "- Monto Máximo a perdonar es " & Format(nMontoMaxPerd, "#0.00"), vbInformation, "Aviso"
            feCuotasVencAct.TextMatrix(pnRow, pnCol) = Format(nMonto, "#0.00")
            feCuotasVencAct.CellForeColor = vbBlack
        End If
    Else
        If CDbl(feCuotasVencAct.TextMatrix(pnRow, pnCol)) > nMonto Then
            MsgBox "Monto a pagar de " & _
                    IIf(pnCol = 4, "''Capital''", IIf(pnCol = 5, "''Interes Comp.''", IIf(pnCol = 6, "''Mora''", "''Gasto''"))) & _
                    " de la cuota " & Trim(feCuotasVencAct.TextMatrix(pnRow, 1)) & " no puede ser mayor al actual.", vbInformation, "Aviso"
            feCuotasVencAct.TextMatrix(pnRow, pnCol) = Format(nMonto, "#0.00")
        End If
        feCuotasVencAct.CellForeColor = vbBlack

    End If
    
    feCuotasVencAct.TextMatrix(pnRow, 9) = Format(CCur(feCuotasVencAct.TextMatrix(pnRow, 4)) + CCur(feCuotasVencAct.TextMatrix(pnRow, 5)) + CCur(feCuotasVencAct.TextMatrix(pnRow, 6)) + CCur(feCuotasVencAct.TextMatrix(pnRow, 7)) + CCur(feCuotasVencAct.TextMatrix(pnRow, 8)), "#0.00")
End If
End Sub


Private Sub Form_Load()
cmdAcoger.Enabled = False
fnNroCalen = 0
Call HabilitarControles(False) 'JOEP
End Sub

Private Function ObtenerIndice(ByVal pnId As Long, ByVal pMatriz As Variant, ByVal pnCampo As Integer) As Integer
Dim nIndex As Integer
nIndex = -1

If IsArray(pMatriz) Then
    If pMatriz(0, 0) <> "" Then
        For i = 0 To UBound(pMatriz)
            If IsNumeric(pMatriz(i, pnCampo)) Then
                If CLng(pMatriz(i, pnCampo)) = pnId Then
                    nIndex = i
                    Exit For
                End If
            End If
        Next i
    End If
End If

ObtenerIndice = nIndex
End Function

Private Sub CalculaMontos()
Dim nMontoPerd As Double
Dim nPagoMin As Double
Dim nPagoMax As Double
Dim nIndexSC As Integer

'Obtener el total perdonado
nMontoPerd = 0
If Trim(feCuotasVencAct.TextMatrix(1, 0)) <> "" Then
    For i = 1 To feCuotasVencAct.Rows - 1
        nMontoPerd = nMontoPerd + (CCur(feCuotasVenc.TextMatrix(i, 6)) - CCur(feCuotasVencAct.TextMatrix(i, 4)))
        nMontoPerd = nMontoPerd + (CCur(feCuotasVenc.TextMatrix(i, 7)) - CCur(feCuotasVencAct.TextMatrix(i, 5)))
        nMontoPerd = nMontoPerd + (CCur(feCuotasVenc.TextMatrix(i, 8)) - CCur(feCuotasVencAct.TextMatrix(i, 6)))
        nMontoPerd = nMontoPerd + (CCur(feCuotasVenc.TextMatrix(i, 9)) - CCur(feCuotasVencAct.TextMatrix(i, 7)))
        nMontoPerd = nMontoPerd + (CCur(feCuotasVenc.TextMatrix(i, 11)) - CCur(feCuotasVencAct.TextMatrix(i, 9))) 'joep
    Next i
End If
lblMontoPerdonar.Caption = Format(nMontoPerd, "###," & String(15, "#") & "#0.00") & " "

'Obtener Rango Pago
nPagoMin = 0
nPagoMax = 0
nIndexSC = ObtenerIndice(CInt(Trim(Right(cmbSubCampana.Text, 6))), MatSubCamp, 1)
If Trim(feCuotasVenc.TextMatrix(1, 0)) <> "" Then
    For i = 1 To feCuotasVenc.Rows - 1
        If feCuotasVenc.TextMatrix(i, 2) = "." Then
            nPagoMin = nPagoMin + (CCur(feCuotasVenc.TextMatrix(i, 6)) - Round(Format(CCur(feCuotasVenc.TextMatrix(i, 6)) * (ArrPorcentaje(0) / 100), "#0.00"), 2))
            nPagoMin = nPagoMin + (CCur(feCuotasVenc.TextMatrix(i, 7)) - Round(Format(CCur(feCuotasVenc.TextMatrix(i, 7)) * (ArrPorcentaje(1) / 100), "#0.00"), 2))
            nPagoMin = nPagoMin + (CCur(feCuotasVenc.TextMatrix(i, 8)) - Round(Format(CCur(feCuotasVenc.TextMatrix(i, 8)) * (ArrPorcentaje(2) / 100), "#0.00"), 2))
            nPagoMin = nPagoMin + (CCur(feCuotasVenc.TextMatrix(i, 9)) - Round(Format(CCur(feCuotasVenc.TextMatrix(i, 9)) * (ArrPorcentaje(3) / 100), "#0.00"), 2))
            
            nPagoMin = nPagoMin + (CCur(feCuotasVenc.TextMatrix(i, 11)) - Round(Format(CCur(feCuotasVenc.TextMatrix(i, 11)) * (ArrPorcentaje(4) / 100), "#0.00"), 2)) 'JOEP
            nPagoMin = nPagoMin + CCur(feCuotasVenc.TextMatrix(i, 10))
            'nPagoMax = nPagoMax + CCur(feCuotasVenc.TextMatrix(i, 11))'JOEP
            nPagoMax = nPagoMax + CCur(feCuotasVenc.TextMatrix(i, 12)) 'JOEP
        End If
    Next i
End If
lblMonedaDesde.Caption = IIf(Mid(Trim(ActXCtaCred.NroCuenta), 9, 1) = "1", "S/.", "$") & " "
lblMonedaHasta.Caption = IIf(Mid(Trim(ActXCtaCred.NroCuenta), 9, 1) = "1", "S/.", "$") & " "
lblMontoDesde.Caption = Format(nPagoMin, "###," & String(15, "#") & "#0.00") & " "
lblMontoHasta.Caption = Format(nPagoMax, "###," & String(15, "#") & "#0.00") & " "

'Obtener Total a Pagar
'lblMontoPagar.Caption = Format(SumarCampo(feCuotasVencAct, 9), "###," & String(15, "#") & "#0.00") & " "'JOEP
lblMontoPagar.Caption = Format(SumarCampo(feCuotasVencAct, 10), "###," & String(15, "#") & "#0.00") & " " 'JOEP
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = True

If Trim(cmbCampana.Text) = "" Then
    MsgBox "Seleccione la Campaña", vbInformation, "Aviso"
    ValidaDatos = False
    cmbCampana.SetFocus
    Exit Function
End If

If Trim(cmbSubCampana.Text) = "" Then
    MsgBox "Seleccione la Sub Campaña", vbInformation, "Aviso"
    ValidaDatos = False
    cmbSubCampana.SetFocus
    Exit Function
End If

If Not IsNumeric(lblMontoPagar.Caption) Then
    MsgBox "Seleccione la Sub Campaña", vbInformation, "Aviso"
    ValidaDatos = False
    feCuotasVencAct.SetFocus
    Exit Function
End If

If CDbl(lblMontoPagar.Caption) = 0 Then
    MsgBox "Seleccione la Sub Campaña", vbInformation, "Aviso"
    ValidaDatos = True
    feCuotasVencAct.SetFocus
    Exit Function
End If

If Not IsNumeric(lblMontoPerdonar.Caption) Then
    MsgBox "Favor de realizar el perdon correspondiente del crédito", vbInformation, "Aviso"
    ValidaDatos = False
    feCuotasVencAct.SetFocus
    Exit Function
End If

If CDbl(lblMontoPerdonar.Caption) = 0 Then
    MsgBox "Favor de realizar el perdon correspondiente del crédito", vbInformation, "Aviso"
    ValidaDatos = False
    feCuotasVencAct.SetFocus
    Exit Function
End If
End Function

Private Sub GrabarDatos()
Dim bgrabar As Boolean
Dim sMovNro As String
Dim oNCredito As COMNCredito.NCOMCredito
Dim oDCredito As COMDCredito.DCOMCredito
Dim rsDatos As ADODB.Recordset
Dim nNivel As Integer
Dim sNivel As String
Dim nMontoPend As Double
Dim nMontoDeuda As Double

On Error GoTo ErrorGrabarDatos
nMontoPend = 0
nMontoDeuda = 0

Set oDCredito = New COMDCredito.DCOMCredito
    
Set rsDatos = oDCredito.RecuperarCampanaRecupNivelApr(CLng(Trim(Right(cmbCampana.Text, 6))), CDbl(lblMontoPerdonar.Caption))
If Not (rsDatos.EOF And rsDatos.BOF) Then
    If bDesctoAdicional = True Then 'CROB20180814 ERS055-2018
        nNivel = 4
        sNivel = "GERENCIA DE CRÈDITOS"
    Else
        nNivel = CInt(rsDatos!nNivel)
        sNivel = Trim(rsDatos!cNivel)
    End If
    'CROB20180814 ERS055-2018

    MsgBox "Se procedera a solicitar el Nivel de Aprobación ''" & sNivel & "''", vbInformation, "Aviso"
    
    'ReDim MatCuotasPerd(feCuotasVencAct.Rows - 2, 13)'JOEP
    ReDim MatCuotasPerd(feCuotasVencAct.Rows - 2, 16) 'JOEP
    For i = 0 To feCuotasVencAct.Rows - 2
        MatCuotasPerd(i, 0) = MatCuotaVenc(i, 0)
        MatCuotasPerd(i, 1) = MatCuotaVenc(i, 3)
        MatCuotasPerd(i, 2) = MatCuotaVenc(i, 4)
        MatCuotasPerd(i, 3) = (CCur(feCuotasVenc.TextMatrix(i + 1, 6)) - CCur(feCuotasVencAct.TextMatrix(i + 1, 4)))
        MatCuotasPerd(i, 4) = MatCuotaVenc(i, 5)
        MatCuotasPerd(i, 5) = MatCuotaVenc(i, 6)
        MatCuotasPerd(i, 6) = (CCur(feCuotasVenc.TextMatrix(i + 1, 7)) - CCur(feCuotasVencAct.TextMatrix(i + 1, 5)))
        MatCuotasPerd(i, 7) = MatCuotaVenc(i, 7) 'Mora
        MatCuotasPerd(i, 8) = MatCuotaVenc(i, 8)
        MatCuotasPerd(i, 9) = (CCur(feCuotasVenc.TextMatrix(i + 1, 8)) - CCur(feCuotasVencAct.TextMatrix(i + 1, 6)))
        MatCuotasPerd(i, 10) = MatCuotaVenc(i, 9) 'Gasto
        MatCuotasPerd(i, 11) = MatCuotaVenc(i, 10)
        MatCuotasPerd(i, 12) = (CCur(feCuotasVenc.TextMatrix(i + 1, 9)) - CCur(feCuotasVencAct.TextMatrix(i + 1, 7)))
        MatCuotasPerd(i, 13) = MatCuotaVenc(i, 11)
        MatCuotasPerd(i, 14) = MatCuotaVenc(i, 12) 'JOEP
        MatCuotasPerd(i, 15) = MatCuotaVenc(i, 13) 'JOEP
        MatCuotasPerd(i, 16) = (CCur(feCuotasVenc.TextMatrix(i + 1, 11)) - CCur(feCuotasVencAct.TextMatrix(i + 1, 9))) 'JOEP
    Next i
    
    For i = 1 To feCuotasVenc.Rows - 1
        'nMontoDeuda = nMontoDeuda + CCur(feCuotasVenc.TextMatrix(i, 11))'JOEP
        nMontoDeuda = nMontoDeuda + CCur(feCuotasVenc.TextMatrix(i, 12)) 'JOEP
        If Trim(feCuotasVenc.TextMatrix(i, 2)) <> "." Then
            'nMontoPend = nMontoPend + CCur(feCuotasVenc.TextMatrix(i, 11))'JOEP
            nMontoPend = nMontoPend + CCur(feCuotasVenc.TextMatrix(i, 12)) 'JOEP
        End If
    Next i

    sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    Set oNCredito = New COMNCredito.NCOMCredito
    bgrabar = oNCredito.AcogerCreditoCampanaRecup(gdFecSis, Trim(ActXCtaCred.NroCuenta), CDbl(lblSaldoCap.Caption), CInt(lblDiasAtraso.Caption), CInt(lblCuotaVenc.Caption), fnNroCalen, _
                                                CLng(Trim(Right(cmbCampana.Text, 6))), CLng(Trim(Right(cmbSubCampana.Text, 6))), CDbl(lblMontoPerdonar.Caption), CDbl(lblMontoPagar.Caption), _
                                                CDbl(lblMontoDesde.Caption), CDbl(lblMontoHasta.Caption), nNivel, nMontoPend, nMontoDeuda, sMovNro, MatCuotasPerd, ArrPorcentaje, gsCodCargo)
    
    If bgrabar Then
        MsgBox "Los datos se grabaron correctamente.", vbInformation, "Aviso"
        
        Set rsDatos = oDCredito.RecuperarCampanaRecupXCredEstadoFecha(Trim(ActXCtaCred.NroCuenta), 1, gdFecSis)
        If Not (rsDatos.BOF And rsDatos.EOF) Then
            If CInt(rsDatos!nEstado) = 1 Then
                MsgBox "Crédito ya fue autorizado automáticamente." & _
                " Ya que la persona que realizo el proceso pertenece al nivel de aprobación solicitado: ''" & sNivel & "''", vbInformation, "Aviso"
            End If
        End If
        LimpiarDatos
    Else
         MsgBox "Hubo errores al grabar las datos", vbError, "Error"
    End If
Else
    MsgBox "No se encontro un nivel de aprobación para el Monto A Perdonar(" & Format(CDbl(lblMontoPerdonar.Caption), "#0.00") & ") de la Campaña ''" & Trim(Left(cmbCampana.Text, 75)) & "''", vbInformation, "Aviso"
End If

Set rsDatos = Nothing
Set oDCredito = Nothing

Exit Sub
ErrorGrabarDatos:
MsgBox Err.Number & " - " & Err.Description, vbError, "Error En Proceso"
End Sub
Private Sub txtCapital_Change()
If Trim(txtCapital.Text) <> "." Then
    If CDbl(txtCapital.Text) > CCur(ArrPorcFijo(0)) Then
        txtCapital.Text = Replace(Mid(txtCapital.Text, 1, Len(txtCapital.Text) - 1), ",", "")
    End If
Else
    txtCapital.Text = "0.00"
End If
ArrPorcentaje(0) = CCur(txtCapital.Text)
Call GenerarCuotasAPerdonar
End Sub

Private Sub txtCapital_LostFocus()
txtCapital.Text = Format(txtCapital.Text, "#0.00")
End Sub

Private Sub txtGasto_Change()
If Trim(txtGasto.Text) <> "." Then
    If CDbl(txtGasto.Text) > CCur(ArrPorcFijo(3)) Then
        txtGasto.Text = Replace(Mid(txtGasto.Text, 1, Len(txtGasto.Text) - 1), ",", "")
    End If
Else
    txtGasto.Text = "0.00"
End If
ArrPorcentaje(3) = CCur(txtGasto.Text)
Call GenerarCuotasAPerdonar
End Sub

Private Sub txtGasto_LostFocus()
txtGasto.Text = Format(txtGasto.Text, "#0.00")
End Sub

Private Sub txtInt_Change()
If Trim(txtInt.Text) <> "." Then
    If CDbl(txtInt.Text) > CCur(ArrPorcFijo(1)) Then
        txtInt.Text = Replace(Mid(txtInt.Text, 1, Len(txtInt.Text) - 1), ",", "")
    End If
Else
    txtInt.Text = "0.00"
End If
ArrPorcentaje(1) = CCur(txtInt.Text)
Call GenerarCuotasAPerdonar
End Sub

Private Sub txtInt_LostFocus()
txtInt.Text = Format(txtInt.Text, "#0.00")
End Sub

Private Sub txtMora_Change()
If Trim(txtMora.Text) <> "." Then
    If CDbl(txtMora.Text) > CCur(ArrPorcFijo(2)) Then
        txtMora.Text = Replace(Mid(txtMora.Text, 1, Len(txtMora.Text) - 1), ",", "")
    End If
Else
    txtMora.Text = "0.00"
End If
ArrPorcentaje(2) = CCur(txtMora.Text)
Call GenerarCuotasAPerdonar
End Sub

Private Sub txtMora_LostFocus()
txtMora.Text = Format(txtMora.Text, "#0.00")
End Sub

'JOEP
Private Sub txticv_change()
If Trim(txtICV.Text) <> "." Then
    If CDbl(txtICV.Text) > CCur(ArrPorcFijo(4)) Then
        txtICV.Text = Replace(Mid(txtICV.Text, 1, Len(txtICV.Text) - 1), ",", "")
    End If
Else
    txtICV.Text = "0.00"
End If
ArrPorcentaje(4) = CCur(txtICV.Text)
Call GenerarCuotasAPerdonar
End Sub

Private Sub txtIcv_LostFocus()
txtICV.Text = Format(txtICV.Text, "#0.00")
End Sub
'JOEP

Private Sub udPorcCapital_DownClick()
Dim valor As Currency
valor = CCur(txtCapital.Text) - 0.01
If valor < 0 Then
    valor = 0
End If
ArrPorcentaje(0) = valor
txtCapital.Text = Format(valor, "#0.00")
Call GenerarCuotasAPerdonar
End Sub

Private Sub udPorcCapital_UpClick()
Dim valor As Currency
valor = CCur(txtCapital.Text) + 0.01
If valor > CCur(ArrPorcFijo(0)) Then
    valor = CCur(ArrPorcFijo(0))
End If
ArrPorcentaje(0) = valor
txtCapital.Text = Format(valor, "#0.00")
Call GenerarCuotasAPerdonar
End Sub

Private Sub udPorcGasto_DownClick()
Dim valor As Currency
valor = CCur(txtGasto.Text) - 0.01
If valor < 0 Then
    valor = 0
End If
ArrPorcentaje(3) = valor
txtGasto.Text = Format(valor, "#0.00")
Call GenerarCuotasAPerdonar
End Sub

Private Sub udPorcGasto_UpClick()
Dim valor As Currency
valor = CCur(txtGasto.Text) + 0.01
If valor > CCur(ArrPorcFijo(3)) Then
    valor = CCur(ArrPorcFijo(3))
End If
ArrPorcentaje(3) = valor
txtGasto.Text = Format(valor, "#0.00")
Call GenerarCuotasAPerdonar
End Sub

'JOEP
Private Sub udPorcICV_DownClick()
    Dim valor As Currency
    valor = CCur(txtICV.Text) - 0.01
    If valor < 0 Then
        valor = 0
    End If
    ArrPorcentaje(4) = valor
    txtICV.Text = Format(valor, "#0.00")
    Call GenerarCuotasAPerdonar
End Sub

Private Sub udPorcIcv_UpClick()
   Dim valor As Currency
    valor = CCur(txtICV.Text) + 0.01
    If valor > CCur(ArrPorcFijo(4)) Then
        valor = CCur(ArrPorcFijo(4))
    End If
    ArrPorcentaje(4) = valor
    txtICV.Text = Format(valor, "#0.00")
    Call GenerarCuotasAPerdonar
End Sub
'JOEP

Private Sub udPorcInt_DownClick()
Dim valor As Currency
valor = CCur(txtInt.Text) - 0.01
If valor < 0 Then
    valor = 0
End If
ArrPorcentaje(1) = valor
txtInt.Text = Format(valor, "#0.00")
Call GenerarCuotasAPerdonar
End Sub

Private Sub udPorcInt_UpClick()
Dim valor As Currency
valor = CCur(txtInt.Text) + 0.01
If valor > CCur(ArrPorcFijo(1)) Then
    valor = CCur(ArrPorcFijo(1))
End If
ArrPorcentaje(1) = valor
txtInt.Text = Format(valor, "#0.00")
Call GenerarCuotasAPerdonar
End Sub

Private Sub udPorcMora_DownClick()
Dim valor As Currency
valor = CCur(txtMora.Text) - 0.01
If valor < 0 Then
    valor = 0
End If
ArrPorcentaje(2) = valor
txtMora.Text = Format(valor, "#0.00")
Call GenerarCuotasAPerdonar
End Sub

Private Sub udPorcMora_UpClick()
Dim valor As Currency
valor = CCur(txtMora.Text) + 0.01
If valor > CCur(ArrPorcFijo(2)) Then
    valor = CCur(ArrPorcFijo(2))
End If
ArrPorcentaje(2) = valor
txtMora.Text = Format(valor, "#0.00")
Call GenerarCuotasAPerdonar
End Sub

'JOEP
Private Sub HabilitarControles(ByVal pbValor As Boolean)
    cmbCampana.Enabled = pbValor
    cmbSubCampana.Enabled = pbValor
    txtCapital.Enabled = pbValor
    txtInt.Enabled = pbValor
    txtMora.Enabled = pbValor
    txtGasto.Enabled = pbValor
    txtICV.Enabled = pbValor
    
    udPorcCapital.Enabled = pbValor
    udPorcInt.Enabled = pbValor
    udPorcMora.Enabled = pbValor
    udPorcGasto.Enabled = pbValor
    udPorcICV.Enabled = pbValor
End Sub
'JOEP

