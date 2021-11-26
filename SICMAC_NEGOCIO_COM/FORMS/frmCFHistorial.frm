VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCFHistorial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estado Carta Fianza"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdSalir 
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
      Height          =   405
      Left            =   7860
      TabIndex        =   37
      ToolTipText     =   "Grabar todos los Cambios Realizados"
      Top             =   5820
      Width           =   1170
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6480
      TabIndex        =   8
      ToolTipText     =   "Grabar todos los Cambios Realizados"
      Top             =   5820
      Width           =   1170
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
      Height          =   405
      Left            =   5160
      TabIndex        =   7
      ToolTipText     =   "Grabar todos los Cambios Realizados"
      Top             =   5820
      Width           =   1170
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3435
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   6059
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "frmCFHistorial.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Historial"
      TabPicture(1)   =   "frmCFHistorial.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Garantias"
      TabPicture(2)   =   "frmCFHistorial.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Otros Datos"
      TabPicture(3)   =   "frmCFHistorial.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame5 
         Height          =   3000
         Left            =   -74880
         TabIndex        =   30
         Top             =   360
         Width           =   8655
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Cancelacion Credito"
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
            Left            =   585
            TabIndex        =   36
            Top             =   1440
            Width           =   1725
         End
         Begin VB.Label lblFechaC 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2385
            TabIndex        =   35
            Top             =   1680
            Width           =   1500
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Cancelacion"
            Height          =   195
            Left            =   585
            TabIndex        =   34
            Top             =   1680
            Width           =   1605
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Rechazo de Credito"
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
            Left            =   585
            TabIndex        =   33
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label lblMotivoR 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2385
            TabIndex        =   32
            Top             =   960
            Width           =   5220
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Motivo de Rechazo"
            Height          =   195
            Left            =   585
            TabIndex        =   31
            Top             =   960
            Width           =   1395
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3000
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   8655
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHGarantias 
            Height          =   2655
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   4683
            _Version        =   393216
            Cols            =   4
            FixedCols       =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   4
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3000
         Left            =   -74880
         TabIndex        =   21
         Top             =   360
         Width           =   8655
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Vigente :"
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
            TabIndex        =   46
            Top             =   1680
            Width           =   780
         End
         Begin VB.Label lblFecha 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   3
            Left            =   1575
            TabIndex        =   45
            Top             =   1680
            Width           =   1500
         End
         Begin VB.Label lblMonto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   3
            Left            =   3375
            TabIndex        =   44
            Top             =   1680
            Width           =   1500
         End
         Begin VB.Label lblMonto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   3360
            TabIndex        =   43
            Top             =   1320
            Width           =   1500
         End
         Begin VB.Label lblFecha 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   1560
            TabIndex        =   42
            Top             =   1320
            Width           =   1500
         End
         Begin VB.Label lblMonto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   3360
            TabIndex        =   41
            Top             =   960
            Width           =   1500
         End
         Begin VB.Label lblFecha 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   1560
            TabIndex        =   40
            Top             =   960
            Width           =   1500
         End
         Begin VB.Label lblMonto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   3360
            TabIndex        =   39
            Top             =   600
            Width           =   1500
         End
         Begin VB.Label lblFecha 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   1560
            TabIndex        =   38
            Top             =   600
            Width           =   1500
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Monto"
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
            Left            =   3840
            TabIndex        =   28
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
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
            Left            =   2040
            TabIndex        =   27
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Sugerido  :"
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
            Left            =   225
            TabIndex        =   26
            Top             =   960
            Width           =   945
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Solicitado :"
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
            Left            =   225
            TabIndex        =   25
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lblFinalidad 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   1560
            TabIndex        =   24
            Top             =   2040
            Width           =   7020
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Finalidad"
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
            Left            =   225
            TabIndex        =   23
            Top             =   2040
            Width           =   780
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Aprobado :"
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
            Left            =   225
            TabIndex        =   22
            Top             =   1320
            Width           =   945
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3000
         Left            =   -74880
         TabIndex        =   9
         Top             =   360
         Width           =   8655
         Begin VB.Label lblModalidad 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3000
            TabIndex        =   49
            Top             =   1920
            Width           =   3615
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Modalidad :"
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
            Left            =   840
            TabIndex        =   20
            Top             =   1920
            Width           =   1005
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Vigencia :"
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
            Left            =   840
            TabIndex        =   19
            Top             =   2280
            Width           =   1725
         End
         Begin VB.Label lblFechaVigencia 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3000
            TabIndex        =   18
            Top             =   2280
            Width           =   3615
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Apoderado :"
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
            Left            =   840
            TabIndex        =   17
            Top             =   1200
            Width           =   1050
         End
         Begin VB.Label lblApoderado 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3000
            TabIndex        =   16
            Top             =   1200
            Width           =   3615
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Condicion Credito :"
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
            Left            =   840
            TabIndex        =   15
            Top             =   1560
            Width           =   1635
         End
         Begin VB.Label lblCondicionCredito 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3000
            TabIndex        =   14
            Top             =   1560
            Width           =   3615
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fuente de Ingreso :"
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
            Left            =   840
            TabIndex        =   13
            Top             =   480
            Width           =   1680
         End
         Begin VB.Label lblFuenteIngreso 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3000
            TabIndex        =   12
            Top             =   480
            Width           =   3615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Analista  :"
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
            Left            =   840
            TabIndex        =   11
            Top             =   840
            Width           =   870
         End
         Begin VB.Label lblAnalista 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3000
            TabIndex        =   10
            Top             =   840
            Width           =   3615
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8955
      Begin VB.CommandButton cmdBuscar 
         Height          =   345
         Left            =   3840
         Picture         =   "frmCFHistorial.frx":0070
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Buscar ..."
         Top             =   300
         Width           =   420
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHPersonas 
         Height          =   1695
         Left            =   4500
         TabIndex        =   47
         Top             =   240
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   2990
         _Version        =   393216
         FixedCols       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin SICMACT.ActXCodCta CodCta 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "C F. Nro:"
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblTipoCF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estado Actual"
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
         TabIndex        =   3
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Carta Fianza"
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
         TabIndex        =   2
         Top             =   720
         Width           =   1785
      End
   End
End
Attribute VB_Name = "frmCFHistorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cPrinter As String
Dim ban As String

Private Function ValidaProducto() As Boolean
    ValidaProducto = IIf(Mid(CodCta.NroCuenta, 6, 3) = "121" Or Mid(CodCta.NroCuenta, 6, 3) = "221" Or Mid(CodCta.NroCuenta, 6, 3) = "514", True, False)
End Function

Private Sub ClearForm()
Dim i As Integer
CodCta.NroCuenta = fgIniciaAxCuentaCF
lblEstado = ""
lblTipoCF = ""
'---------
lblFuenteIngreso = ""
LblAnalista = ""
lblapoderado = ""
lblCondicionCredito = ""
lblModalidad = ""
lblfechavigencia = ""
'-----------
For i = 0 To 3
    lblFecha(i) = ""
    lblMonto(i) = ""
Next i
lblFinalidad = ""
'-----------
MSHPersonas.ClearStructure
MSHGarantias.ClearStructure
'-----------
lblMotivoR = ""
lblFechaC = ""
MarcoFlexGrid
End Sub
Private Sub CargaProducto(ByVal CTA As String)
Dim lsCarta As COMNCartaFianza.NCOMCartaFianzaValida 'NCartaFianzaValida
Dim lsHistorial As COMNCartaFianza.NCOMCartaFianzaValida 'NCartaFianzaValida

Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
Dim Rs3 As New ADODB.Recordset
Dim Rs4 As New ADODB.Recordset
Dim i As Integer
Dim x As Integer

Set lsCarta = New COMNCartaFianza.NCOMCartaFianzaValida
Set rs = lsCarta.RecuperaDatosGeneralesCF(CTA)
If ValidaProducto And Len(CTA) = 18 Then
    
    Set lsCarta = New COMNCartaFianza.NCOMCartaFianzaValida
    Set lsHistorial = New COMNCartaFianza.NCOMCartaFianzaValida
    
    Set rs = lsCarta.RecuperaDatosGeneralesCF(CTA)
    If rs.BOF And rs.EOF Then
       MsgBox "No existe Carta Fianza" & CTA, vbInformation, "AVISO"
    Else
    
      ban = 1
      lblEstado = rs!Estado
      lblFuenteIngreso = IIf(IsNull(rs!F_Ingreso), "", rs!F_Ingreso)
      lblfechavigencia = rs!Vence
      lblTipoCF = IIf(IsNull(rs!Tipo), "", rs!Tipo)
      lblapoderado = IIf(IsNull(rs!Apoderado), "", rs!Apoderado)
      LblAnalista = rs!Analista
      lblFinalidad = IIf(IsNull(rs!Finalidad), "", rs!Finalidad)
      lblModalidad = IIf(IsNull(rs!Modalidad), "", rs!Modalidad)
      lblCondicionCredito = rs!Condicion
      Set rs = Nothing
      Set rs1 = lsHistorial.RecuperaHistorial(CTA)
      x = 0
     
      'MsgBox rs1!Fecha
     '-------------------
      While Not rs1.EOF And x < 4
         lblFecha(x).Caption = IIf(IsNull(rs1!Fecha), "", Format(rs1!Fecha, "dd/mm/yyyy"))
         lblMonto(x).Caption = IIf(IsNull(rs1!Monto), "", Format(rs1!Monto, "#,##0.00#"))
         x = x + 1
         rs1.MoveNext
      Wend
     '-------------------
     Set Rs2 = lsCarta.RecuperaPersonaRelacion(CTA)
     i = 1
     While Not Rs2.EOF
     With MSHPersonas
        .AddItem i, i
        .TextMatrix(i, 0) = IIf(IsNull(Rs2!Nombre), "", Rs2!Nombre)
        .TextMatrix(i, 1) = IIf(IsNull(Rs2!Relacion), "", Rs2!Relacion)
        Rs2.MoveNext
     End With
     Wend
     '--------------------
    
     Set Rs3 = lsCarta.RecuperaGarantias(CTA)
     i = 1
     While Not Rs3.EOF
     With MSHGarantias
        .AddItem i, i
        .TextMatrix(i, 0) = IIf(IsNull(Rs3!NroGarantia), "", Rs3!NroGarantia)
        .TextMatrix(i, 1) = IIf(IsNull(Rs3!Tipo_Garn), "", Rs3!Tipo_Garn)
        .TextMatrix(i, 2) = IIf(IsNull(Rs3!Descripcion), "", Rs3!Descripcion)
        .TextMatrix(i, 3) = IIf(IsNull(Rs3!Tipo_Doc), "", Rs3!Tipo_Doc)
        Rs3.MoveNext
    End With
    Wend
    MSHGarantias.Rows = MSHGarantias.Rows - 1
    MSHPersonas.Rows = MSHPersonas.Rows - 1
    '----------------------
    MarcoFlexGrid
    Set Rs4 = lsCarta.RepcuperaOtrosDt(CTA)
   ' MsgBox rs4.RecordCount
    If Rs4.EOF And Rs4.BOF Then
    Else
        lblMotivoR = IIf(IsNull(Rs4!Motivo), "", Rs4!Motivo)
        lblFechaC = IIf(IsNull(Rs4!Fecha), "", Rs4!Fecha)
    End If
    
    End If
        
    Set lsCarta = Nothing
 Else
    MsgBox "Error en Carta Fianza", vbInformation, "CMACT"
 End If

End Sub

Private Sub cmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersCF As COMNCartaFianza.NCOMCartaFianzaValida
Dim lrCF  As New ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona 'UProdPersona

On Error GoTo ControlError

cmdCancelar_Click

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

' Selecciona Estados
lsEstados = gColocEstSolic & "," & gColocEstSug & "," & gColocEstAprob & "," & gColocEstRetirado & "," & _
            gColocEstRetirado & "," & gColocEstVigNorm & "," & gColocEstCancelado & "," & _
            gColocEstHonrada & "," & gColocEstDevuelta

If Trim(lsPersCod) <> "" Then
    Set loPersCF = New COMNCartaFianza.NCOMCartaFianzaValida
        Set lrCF = loPersCF.dObtieneCFianzaDePersona(lsPersCod, lsEstados)
    Set loPersCF = Nothing
End If

Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrCF)
    If loCuentas.sCtaCod <> "" Then
        CodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        CodCta.SetFocusCuenta
    End If
Set loCuentas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdCancelar_Click()
ClearForm
MSHPersonas.ClearStructure
MSHGarantias.ClearStructure
MSHGarantias.Rows = 2
MSHPersonas.Rows = 2
MarcoFlexGrid
CodCta.Enabled = True
End Sub

Private Sub cmdImprimir_Click()

Dim loRep As COMNCartaFianza.NCOMCartaFianzaReporte 'NCartaFianzaReporte
Dim lsCadImp As String
Dim loPrevio As previo.clsprevio
Dim lsDestino As String  ' P= Previo // I = Impresora // A = Archivo // E = Excel

If Len(CodCta.NroCuenta) = 18 Then
 Set loRep = New COMNCartaFianza.NCOMCartaFianzaReporte
 loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
 
 lsCadImp = loRep.nRepoCartaFianza_Estados(CodCta.NroCuenta, gImpresora)
 lsDestino = "P"
 If lsDestino = "P" Then
    If Len(Trim(lsCadImp)) > 0 Then
        Set loPrevio = New previo.clsprevio
            loPrevio.Show Chr$(27) & Chr$(77) & lsCadImp, "Reporte Historial Carta Fianza", True, , gImpresora
        Set loPrevio = Nothing
    Else
        MsgBox "No Existen Datos para el reporte ", vbInformation, "Aviso"
    End If
 End If
End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub MarcoFlexGrid()
With MSHPersonas
    .TextMatrix(0, 0) = "Nombre de Persona"
    .TextMatrix(0, 1) = "Relacion"
    .ColWidth(0) = 3000
    .ColWidth(1) = 1000
End With
'--------------------
With MSHGarantias
    .TextMatrix(0, 0) = "Nro Garantia"
    .TextMatrix(0, 1) = "Tipo Garantia"
    .TextMatrix(0, 2) = "Descripcion"
    .TextMatrix(0, 3) = "Documento"

    .ColWidth(0) = 1000
    .ColWidth(1) = 2500
    .ColWidth(2) = 2800
    .ColWidth(3) = 2800
    
End With
End Sub

Private Sub CodCta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Len(CodCta.NroCuenta) = 18 Then
    'MSHPersonas.ClearStructure
    'MSHGarantias.ClearStructure
    MSHPersonas.Rows = 2
    MSHGarantias.Rows = 2
    CargaProducto (CodCta.NroCuenta)
    CodCta.Enabled = False
Else
    MsgBox "Nro de Cta Imcompleto", vbInformation, "Aviso"
    MSHPersonas.ClearStructure
    MSHGarantias.ClearStructure
    ClearForm
    MarcoFlexGrid
End If


End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
ban = 0
ClearForm
MarcoFlexGrid
'****************
End Sub

Public Sub CargaCFHistorial(ByVal psCodCta As String)

    MSHPersonas.ClearStructure
    MSHGarantias.ClearStructure
    CodCta.NroCuenta = psCodCta
    CargaProducto (psCodCta)
    CodCta.Enabled = False
    cmdBuscar.Enabled = False
    cmdCancelar.Enabled = False
    Me.Show 1
    
End Sub

'WIOR 20130420 ************************************
Private Sub MSHGarantias_DblClick()
If MSHGarantias.row > 0 Then
    'frmPersGarantias.Inicio ConsultaGarant, Trim(MSHGarantias.TextMatrix(MSHGarantias.row, 0))
    frmGarantia.Consultar Trim(MSHGarantias.TextMatrix(MSHGarantias.row, 0)) 'EJVG20150725
End If
End Sub
'WIOR FIN *****************************************
