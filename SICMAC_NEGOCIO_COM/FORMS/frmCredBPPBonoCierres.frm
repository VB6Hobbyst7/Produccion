VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredBPPBonoCierres 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Realizar Cierres de Bonos Generados"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17730
   Icon            =   "frmCredBPPBonoCierres.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   17730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrarBPP 
      BackColor       =   &H00000000&
      Caption         =   "Cerrar BPP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   8400
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16440
      TabIndex        =   12
      Top             =   8400
      Width           =   1170
   End
   Begin TabDlg.SSTab SSTBusqueda 
      Height          =   7770
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   13705
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Bono Analista"
      TabPicture(0)   =   "frmCredBPPBonoCierres.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraBusqueda"
      Tab(0).Control(1)=   "cmdExportarA"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Bono Coordinador y Jefe de Agencia"
      TabPicture(1)   =   "frmCredBPPBonoCierres.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdExportarCJA"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Bono Jefe Territorial"
      TabPicture(2)   =   "frmCredBPPBonoCierres.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "cmdExportarJT"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton cmdExportarCJA 
         Caption         =   "Exportar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   7200
         Width           =   1170
      End
      Begin VB.CommandButton cmdExportarJT 
         Caption         =   "Exportar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   27
         Top             =   7320
         Width           =   1170
      End
      Begin VB.Frame Frame4 
         Caption         =   "Resultado"
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
         Height          =   6735
         Left            =   -74880
         TabIndex        =   21
         Top             =   480
         Width           =   16695
         Begin SICMACT.FlexEdit feJefeTerritorial 
            Height          =   5775
            Left            =   240
            TabIndex        =   22
            Top             =   675
            Width           =   16335
            _extentx        =   28813
            _extenty        =   10186
            cols0           =   13
            highlight       =   1
            encabezadosnombres=   "#-Zona-Usuario-Meta-Cierre-Meta-Cierre-Caja-Agencia-Agencia-Ag. Bonif.-Porc.-Total"
            encabezadosanchos=   "0-2000-1000-1500-1500-1500-1500-1200-1200-1000-1000-1000-1200"
            font            =   "frmCredBPPBonoCierres.frx":035E
            font            =   "frmCredBPPBonoCierres.frx":0386
            font            =   "frmCredBPPBonoCierres.frx":03AE
            font            =   "frmCredBPPBonoCierres.frx":03D6
            font            =   "frmCredBPPBonoCierres.frx":03FE
            fontfixed       =   "frmCredBPPBonoCierres.frx":0426
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   3
            columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X"
            listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-C-C-R-R-R-R-R-R-R-R-R-R"
            formatosedit    =   "0-0-0-2-2-2-2-2-2-3-3-2-2"
            cantentero      =   15
            textarray0      =   "#"
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Saldo de Cartera"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3255
            TabIndex        =   26
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Saldo Cartera Vencida y Judicial"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   6255
            TabIndex        =   25
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Rendimiento de Cartera"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   9255
            TabIndex        =   24
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Rendimiento de Cartera"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   11655
            TabIndex        =   23
            Top             =   360
            Width           =   3030
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exportar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Top             =   8400
         Width           =   1170
      End
      Begin VB.Frame Frame3 
         Caption         =   "Resultado"
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
         Height          =   6615
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   16815
         Begin SICMACT.FlexEdit feCoordJA 
            Height          =   5775
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   16575
            _extentx        =   29236
            _extenty        =   10186
            cols0           =   12
            highlight       =   1
            encabezadosnombres=   "#-Agencia-Usuario-Comite-% Bonif.-Meta-Cierre-Meta-Cierre-Caja-Agencia-Total"
            encabezadosanchos=   "0-2000-1000-1500-1200-1500-1500-1500-1500-1500-1500-1200"
            font            =   "frmCredBPPBonoCierres.frx":044C
            font            =   "frmCredBPPBonoCierres.frx":0474
            font            =   "frmCredBPPBonoCierres.frx":049C
            font            =   "frmCredBPPBonoCierres.frx":04C4
            font            =   "frmCredBPPBonoCierres.frx":04EC
            fontfixed       =   "frmCredBPPBonoCierres.frx":0514
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   3
            columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
            listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-C-C-L-R-R-R-R-R-R-R-R"
            formatosedit    =   "0-0-0-0-2-2-2-2-2-2-2-2"
            cantentero      =   15
            textarray0      =   "#"
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Saldo Cartera Vencida y Judicial"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   8830
            TabIndex        =   20
            Top             =   285
            Width           =   3015
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Saldo Cartera Vencida y Judicial"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   8840
            TabIndex        =   18
            Top             =   1485
            Width           =   3010
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Saldo de Cartera"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   5835
            TabIndex        =   17
            Top             =   285
            Width           =   3015
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Rendimiento de Cartera"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   11835
            TabIndex        =   16
            Top             =   285
            Width           =   3015
         End
      End
      Begin VB.Frame fraBusqueda 
         Caption         =   "Resultado"
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
         Height          =   6495
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   15375
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   730
            Left            =   2110
            TabIndex        =   5
            Top             =   315
            Width           =   1030
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Analista"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   200
               TabIndex        =   6
               Top             =   315
               Width           =   690
            End
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   730
            Left            =   130
            TabIndex        =   3
            Top             =   315
            Width           =   1995
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Agencia"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   720
               TabIndex        =   4
               Top             =   360
               Width           =   675
            End
         End
         Begin SICMACT.FlexEdit feAnalistas 
            Height          =   5535
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   15135
            _extentx        =   26696
            _extenty        =   9763
            cols0           =   13
            highlight       =   1
            encabezadosnombres=   "#-Agencia-Analista-Saldo-Clientes-Operac.-Mora(8 a 30)-Meta-Plus-Rendimiento-Penalidad-Total-PersCod"
            encabezadosanchos=   "0-2000-1000-1500-1000-1000-1200-1000-1500-1500-1500-1500-0"
            font            =   "frmCredBPPBonoCierres.frx":053A
            font            =   "frmCredBPPBonoCierres.frx":0562
            font            =   "frmCredBPPBonoCierres.frx":058A
            font            =   "frmCredBPPBonoCierres.frx":05B2
            font            =   "frmCredBPPBonoCierres.frx":05DA
            fontfixed       =   "frmCredBPPBonoCierres.frx":0602
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   3
            columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X"
            listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-C-C-R-R-R-R-R-R-R-R-R-L"
            formatosedit    =   "0-0-0-2-3-3-2-2-2-2-2-2-0"
            cantentero      =   15
            textarray0      =   "#"
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "% Cumplimiento Analista"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3135
            TabIndex        =   9
            Top             =   405
            Width           =   4725
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Bono S/."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   7845
            TabIndex        =   8
            Top             =   405
            Width           =   7020
         End
      End
      Begin VB.CommandButton cmdExportarA 
         Caption         =   "Exportar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   1
         Top             =   7200
         Width           =   1170
      End
   End
   Begin VB.Label lblMes 
      AutoSize        =   -1  'True
      Caption         =   "@Mes"
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
      Left            =   1560
      TabIndex        =   11
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Mes generado:"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   1065
   End
End
Attribute VB_Name = "frmCredBPPBonoCierres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private fMatAnalistas() As AnalistaBPP
'Private fMatCoordJA() As CoordinadorJABPP
'Private fMatJefesTerritorial() As JefeTerritorialBPP
'Private i As Integer
'Private nIndex As Integer
'Private fgFecActual As Date
'
'
'Public Sub Inicio(ByVal pnMes As Integer, ByVal pnAnio As Integer, Optional psDesc As String = "", Optional ByVal pnCierre As Boolean = False)
'Me.Caption = Trim(psDesc)
'cmdCerrarBPP.Visible = pnCierre
'fgFecActual = CDate(CStr(pnAnio) & "-" & CStr(IIf(pnMes < 10, "0" & pnMes, pnMes)) & "-01")
'fgFecActual = DateAdd("M", 1, fgFecActual)
'fgFecActual = DateAdd("D", -1, fgFecActual)
'lblMes.Caption = MesAnio(fgFecActual)
'cmdExportarA.Enabled = False
'cmdExportarCJA.Enabled = False
'cmdExportarJT.Enabled = False
'
''ANALISTAS
'Call LlenaDatosAnalista(pnMes, pnAnio)
''COORDINADORES Y JEFES DE AGENCIA
'Call LlenaDatosCoordJA(pnMes, pnAnio)
''JEFES TERRITORIALES
'Call LlenaDatosJT(pnMes, pnAnio)
'Me.Show 1
'End Sub
'
'Private Sub cmdCerrar_Click()
'Unload Me
'End Sub
'
'Private Sub cmdCerrarBPP_Click()
'If MsgBox("Estas seguro de cerrar el Mes Generado del BPP?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'    Dim dFechaNew As Date
'    Dim oBPP As COMNCredito.NCOMBPPR
'    Dim oConsSist As COMDConstSistema.NCOMConstSistema
'    Dim sMovNro As String
'    Set oConsSist = New COMDConstSistema.NCOMConstSistema
'    Set oBPP = New COMNCredito.NCOMBPPR
'    sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'    dFechaNew = DateAdd("D", 1, fgFecActual)
'    dFechaNew = DateAdd("M", 1, dFechaNew)
'    dFechaNew = DateAdd("D", -1, dFechaNew)
'
'    Call oBPP.RegistraCierreBonosGenerados(fgFecActual, sMovNro)
'    Call oConsSist.ActualizaConstSistemas(gConstSistFechaBPP, sMovNro & "'", "'" & Format(dFechaNew, "DD/MM/YYYY") & "'", "", False)
'
'    MsgBox "Mes Cerrado Satisfactoriamente", vbInformation, "Aviso"
'    Set oConsSist = Nothing
'    Set oBPP = Nothing
'    cmdCerrarBPP.Enabled = False
'
'End If
'End Sub
'
'Private Function MesAnio(ByVal dFecha As Date) As String
'Dim sFechaDesc As String
'sFechaDesc = ""
'
'Select Case Month(dFecha)
'    Case 1: sFechaDesc = "Enero"
'    Case 2: sFechaDesc = "Febrero"
'    Case 3: sFechaDesc = "Marzo"
'    Case 4: sFechaDesc = "Abril"
'    Case 5: sFechaDesc = "Mayo"
'    Case 6: sFechaDesc = "Junio"
'    Case 7: sFechaDesc = "Julio"
'    Case 8: sFechaDesc = "Agosto"
'    Case 9: sFechaDesc = "Septiembre"
'    Case 10: sFechaDesc = "Octubre"
'    Case 11: sFechaDesc = "Noviembre"
'    Case 12: sFechaDesc = "Diciembre"
'End Select
'
'sFechaDesc = sFechaDesc & " " & CStr(Year(dFecha))
'MesAnio = UCase(sFechaDesc)
'End Function
'
'Private Sub CargaDatosAnalista(ByVal pnMes As Integer, ByVal pnAnio As Integer)
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim rsBPP As ADODB.Recordset
' On Error GoTo ErrorCargaDatos
'Set oBPP = New COMNCredito.NCOMBPPR
'
'Set rsBPP = oBPP.BonoGenerados(pnMes, pnAnio, 1)
'
'If Not (rsBPP.EOF And rsBPP.BOF) Then
'    ReDim fMatAnalistas(rsBPP.RecordCount - 1)
'    For i = 0 To rsBPP.RecordCount - 1
'        fMatAnalistas(i).cPersCod = rsBPP!cPersCod
'        fMatAnalistas(i).Agencia = rsBPP!Agencia
'        fMatAnalistas(i).Usuario = rsBPP!Usuario
'        fMatAnalistas(i).NombreAnalista = rsBPP!NombreAnalista
'        fMatAnalistas(i).comite = rsBPP!nComite
'        fMatAnalistas(i).Nivel = rsBPP!Nivel
'        fMatAnalistas(i).Categoria = rsBPP!Categoria
'        fMatAnalistas(i).MoraBase = rsBPP!MoraBase
'        fMatAnalistas(i).MetaSaldo = rsBPP!MetaSaldo
'        fMatAnalistas(i).MetaSaldoAG = rsBPP!MetaSaldoAG
'        fMatAnalistas(i).MetaCliente = rsBPP!MetaCliente
'        fMatAnalistas(i).MetaClienteAG = rsBPP!MetaClienteAG
'        fMatAnalistas(i).MetaOperaciones = rsBPP!MetaOperaciones
'        fMatAnalistas(i).MetaOperacionesAG = rsBPP!MetaOperacionesAG
'        fMatAnalistas(i).MetaMora = rsBPP!MetaMora
'        fMatAnalistas(i).MetaMoraAG = rsBPP!MetaMoraAG
'        fMatAnalistas(i).MetaRendimiento = rsBPP!MetaRendimiento
'        fMatAnalistas(i).MetaRendimientoAG = rsBPP!MetaRendimientoAG
'        fMatAnalistas(i).SaldoCapital = rsBPP!SaldoCapital
'        fMatAnalistas(i).SaldoInicial = rsBPP!SaldoInicial
'        fMatAnalistas(i).SaldoEntrante = rsBPP!SaldoEntrante
'        fMatAnalistas(i).SaldoSaliente = rsBPP!SaldoSaliente
'        fMatAnalistas(i).SIA = rsBPP!SIA
'        fMatAnalistas(i).SCE = rsBPP!SCE
'        fMatAnalistas(i).PPOSaldo = rsBPP!PPOSaldo
'        fMatAnalistas(i).PTFSaldo = rsBPP!PTFSaldo
'        fMatAnalistas(i).IXSaldo = rsBPP!IXSaldo
'        fMatAnalistas(i).PXSaldo = rsBPP!PXSaldo
'        fMatAnalistas(i).CantClientes = rsBPP!CantClientes
'        fMatAnalistas(i).ClientesInicial = rsBPP!ClientesInicial
'        fMatAnalistas(i).ClientesEntrantes = rsBPP!ClientesEntrantes
'        fMatAnalistas(i).ClientesSalientes = rsBPP!ClientesSalientes
'        fMatAnalistas(i).NIC = rsBPP!NIC
'        fMatAnalistas(i).NCE = rsBPP!NCE
'        fMatAnalistas(i).PPOCliente = rsBPP!PPOCliente
'        fMatAnalistas(i).PTFCliente = rsBPP!PTFCliente
'        fMatAnalistas(i).IXCliente = rsBPP!IXCliente
'        fMatAnalistas(i).PXCliente = rsBPP!PXCliente
'        fMatAnalistas(i).NFO1 = rsBPP!NFO1
'        fMatAnalistas(i).NOE1 = rsBPP!NOE1
'        fMatAnalistas(i).PPOOpe1 = rsBPP!PPOOpe1
'        fMatAnalistas(i).NFO2 = rsBPP!NFO2
'        fMatAnalistas(i).NOE2 = rsBPP!NOE2
'        fMatAnalistas(i).PPOOpe2 = rsBPP!PPOOpe2
'        fMatAnalistas(i).PTFO = rsBPP!PTFO
'        fMatAnalistas(i).IXOperaciones = rsBPP!IXOperaciones
'        fMatAnalistas(i).PXOperaciones = rsBPP!PXOperaciones
'        fMatAnalistas(i).MF830 = rsBPP!MF830
'        fMatAnalistas(i).MI830 = rsBPP!MI830
'        fMatAnalistas(i).ME830 = rsBPP!ME830
'        fMatAnalistas(i).PP830 = rsBPP!PP830
'        fMatAnalistas(i).PTFMora = rsBPP!PTFMora
'        fMatAnalistas(i).IXM830 = rsBPP!IXM830
'        fMatAnalistas(i).PXMora = rsBPP!PXMora
'        fMatAnalistas(i).ICOB = rsBPP!ICOB
'        fMatAnalistas(i).PESP = rsBPP!PESP
'        fMatAnalistas(i).CCC = rsBPP!CCC
'        fMatAnalistas(i).RCA = rsBPP!RCA
'        fMatAnalistas(i).IXRendimiento = rsBPP!IXRendimiento
'        fMatAnalistas(i).MIMayor30 = rsBPP!MIMayor30
'        fMatAnalistas(i).CJI = rsBPP!CJI
'        fMatAnalistas(i).TMI = rsBPP!TMI
'        fMatAnalistas(i).MFMayor30 = rsBPP!MFMayor30
'        fMatAnalistas(i).CJF = rsBPP!CJF
'        fMatAnalistas(i).TMF = rsBPP!TMF
'        fMatAnalistas(i).BonoMeta = rsBPP!BonoMeta
'        fMatAnalistas(i).BonoPlus = rsBPP!BonoPlus
'        fMatAnalistas(i).BonoRendimiento = rsBPP!BonoRendimiento
'        fMatAnalistas(i).Penalidad = rsBPP!Penalidad
'        fMatAnalistas(i).BonoTotal = rsBPP!BonoTotal
'        fMatAnalistas(i).TopeSaldo = rsBPP!TopeSaldo
'        fMatAnalistas(i).PorcMinSaldo = rsBPP!PorcMinSaldo
'        fMatAnalistas(i).TopeCliente = rsBPP!TopeCliente
'        fMatAnalistas(i).PorcMinCliente = rsBPP!PorcMinCliente
'        fMatAnalistas(i).TopeOperaciones = rsBPP!TopeOperaciones
'        fMatAnalistas(i).PorcMinOperaciones = rsBPP!PorcMinOperaciones
'        fMatAnalistas(i).RangoPerMora = rsBPP!RangoPerMora
'        fMatAnalistas(i).TopeMora = rsBPP!TopeMora
'        fMatAnalistas(i).PorcMinMora = rsBPP!PorcMinMora
'        fMatAnalistas(i).MoraAcepMayor30 = rsBPP!MoraAcepMayor30
'        fMatAnalistas(i).IntCobCMACM = rsBPP!IntCobCMACM
'        fMatAnalistas(i).SaldoCMACM = rsBPP!SaldoCMACM
'        fMatAnalistas(i).RendCMACM = rsBPP!RendCMACM
'        fMatAnalistas(i).MinRendCartera = rsBPP!MinRendCartera
'        fMatAnalistas(i).FactorRend = rsBPP!FactorRend
'        fMatAnalistas(i).PrimQuincena = rsBPP!PrimQuincena
'        fMatAnalistas(i).SegunQuincena = rsBPP!SegunQuincena
'        fMatAnalistas(i).SaldoPlus = rsBPP!SaldoPlus
'        fMatAnalistas(i).ClientesPlus = rsBPP!ClientesPlus
'        fMatAnalistas(i).OperacionesPlus = rsBPP!OperacionesPlus
'        fMatAnalistas(i).MoraPlus = rsBPP!MoraPlus
'        fMatAnalistas(i).Mora830IncialCierre = rsBPP!Mora830IncialCierre
'        fMatAnalistas(i).Mora830Entrante = rsBPP!Mora830Entrante
'        fMatAnalistas(i).Mora830Saliente = rsBPP!Mora830Saliente
'        rsBPP.MoveNext
'    Next i
'Else
'    ReDim fMatAnalistas(0)
'End If
'
'Set rsBPP = Nothing
'Set oBPP = Nothing
'
'Exit Sub
'ErrorCargaDatos:
'ReDim fMatAnalistas(0)
'MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub CargaDatosCoordJA(ByVal pnMes As Integer, ByVal pnAnio As Integer)
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim rsBPP As ADODB.Recordset
'On Error GoTo ErrorCargaDatos
'Set oBPP = New COMNCredito.NCOMBPPR
'Set rsBPP = oBPP.BonoGenerados(pnMes, pnAnio, 2)
'
'If Not (rsBPP.EOF And rsBPP.BOF) Then
'    ReDim fMatCoordJA(rsBPP.RecordCount - 1)
'    For i = 0 To rsBPP.RecordCount - 1
'        fMatCoordJA(i).cCodAge = rsBPP!cCodAge
'        fMatCoordJA(i).Agencia = rsBPP!Agencia
'        fMatCoordJA(i).nCargo = rsBPP!nCargo
'        fMatCoordJA(i).Cargo = rsBPP!Cargo
'        fMatCoordJA(i).cPersCod = rsBPP!cPersCod
'        fMatCoordJA(i).comite = rsBPP!comite
'        fMatCoordJA(i).Usuario = rsBPP!Usuario
'        fMatCoordJA(i).Nombre = rsBPP!Nombre
'        fMatCoordJA(i).CantAnalista = rsBPP!CantAnalista
'        fMatCoordJA(i).PorBonificacion = rsBPP!PorBonificacion
'        fMatCoordJA(i).SaldoCartera = rsBPP!SaldoCartera
'        fMatCoordJA(i).SaldoCarteraCierre = rsBPP!SaldoCarteraCierre
'        fMatCoordJA(i).PorSalCartVencJud = rsBPP!PorSalCartVencJud
'        fMatCoordJA(i).PorcSalVenJud = rsBPP!PorcSalVenJud
'        fMatCoordJA(i).RendCaja = rsBPP!RendCaja
'        fMatCoordJA(i).RendAG = rsBPP!RendAG
'        fMatCoordJA(i).Tope = rsBPP!nTope
'        fMatCoordJA(i).BonoTotal = rsBPP!BonoTotal
'        fMatCoordJA(i).PorBonificacionConf = rsBPP!PorBonificacionConf
'        fMatCoordJA(i).AnalistaBoni = rsBPP!AnalistaBoni
'        fMatCoordJA(i).SaldoCapital = rsBPP!SaldoCapital
'        fMatCoordJA(i).SaldoVencJud = rsBPP!SaldoVencJud
'        fMatCoordJA(i).IntCob = rsBPP!IntCob
'        rsBPP.MoveNext
'    Next i
'Else
'    ReDim fMatCoordJA(0)
'End If
'
'Set rsBPP = Nothing
'Set oBPP = Nothing
'
'Exit Sub
'ErrorCargaDatos:
'ReDim fMatCoordJA(0)
'MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub CargaDatosJT(ByVal pnMes As Integer, ByVal pnAnio As Integer)
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim rsBPP As ADODB.Recordset
' On Error GoTo ErrorCargaDatos
'Set oBPP = New COMNCredito.NCOMBPPR
'Set rsBPP = oBPP.BonoGenerados(pnMes, pnAnio, 3)
'
'
'If Not (rsBPP.EOF And rsBPP.BOF) Then
'    ReDim fMatJefesTerritorial(rsBPP.RecordCount - 1)
'    For i = 0 To rsBPP.RecordCount - 1
'        fMatJefesTerritorial(i).nZona = rsBPP!nZona
'        fMatJefesTerritorial(i).Zona = rsBPP!Zona
'        fMatJefesTerritorial(i).cPersCod = rsBPP!cPersCod
'        fMatJefesTerritorial(i).Usuario = rsBPP!Usuario
'        fMatJefesTerritorial(i).Nombre = rsBPP!Nombre
'        fMatJefesTerritorial(i).SaldoCartera = rsBPP!SaldoCartera
'        fMatJefesTerritorial(i).SaldoCarteraCierre = rsBPP!SaldoCarteraCierre
'        fMatJefesTerritorial(i).PorSalCartVencJud = rsBPP!PorSalCartVencJud
'        fMatJefesTerritorial(i).PorcSalVenJud = rsBPP!PorcSalVenJud
'        fMatJefesTerritorial(i).RendCaja = rsBPP!RendCaja
'        fMatJefesTerritorial(i).Rend = rsBPP!Rend
'        fMatJefesTerritorial(i).CantAge = rsBPP!CantAge
'        fMatJefesTerritorial(i).AgeBoni = rsBPP!AgeBoni
'        fMatJefesTerritorial(i).PorBoni = rsBPP!PorBoni
'        fMatJefesTerritorial(i).PorConfJF = rsBPP!PorConfJF
'        fMatJefesTerritorial(i).Tope = rsBPP!nTope
'        fMatJefesTerritorial(i).BonoTotal = rsBPP!BonoTotal
'        rsBPP.MoveNext
'    Next i
'Else
'    ReDim fMatJefesTerritorial(0)
'End If
'
'Set rsBPP = Nothing
'Set oBPP = Nothing
'
'Exit Sub
'ErrorCargaDatos:
'ReDim fMatJefesTerritorial(0)
'MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub LlenaDatosAnalista(ByVal pnMes As Integer, ByVal pnAnio As Integer)
'LimpiaFlex feAnalistas
'Call CargaDatosAnalista(pnMes, pnAnio)
'
'If fMatAnalistas(0).Usuario <> "" Then
'    For i = 0 To UBound(fMatAnalistas)
'        feAnalistas.AdicionaFila
'        feAnalistas.TextMatrix(i + 1, 1) = fMatAnalistas(i).Agencia
'        feAnalistas.TextMatrix(i + 1, 2) = fMatAnalistas(i).Usuario
'
'        feAnalistas.TextMatrix(i + 1, 3) = Format(fMatAnalistas(i).MetaSaldoAG, "###," & String(15, "#") & "#0.00")
'        feAnalistas.TextMatrix(i + 1, 4) = Format(fMatAnalistas(i).MetaClienteAG, "###," & String(15, "#") & "#0.00")
'        feAnalistas.TextMatrix(i + 1, 5) = Format(fMatAnalistas(i).MetaOperacionesAG, "###," & String(15, "#") & "#0.00")
'        feAnalistas.TextMatrix(i + 1, 6) = Format(fMatAnalistas(i).MetaMoraAG, "###," & String(15, "#") & "#0.00")
'
'
'        feAnalistas.TextMatrix(i + 1, 7) = Format(fMatAnalistas(i).BonoMeta, "###," & String(15, "#") & "#0.00")
'        feAnalistas.TextMatrix(i + 1, 8) = Format(fMatAnalistas(i).BonoPlus, "###," & String(15, "#") & "#0.00")
'        feAnalistas.TextMatrix(i + 1, 9) = Format(fMatAnalistas(i).BonoRendimiento, "###," & String(15, "#") & "#0.00")
'        feAnalistas.TextMatrix(i + 1, 10) = Format(fMatAnalistas(i).Penalidad * 100, "###," & String(15, "#") & "#0.00") & " %"
'        feAnalistas.TextMatrix(i + 1, 11) = Format(fMatAnalistas(i).BonoTotal, "###," & String(15, "#") & "#0.00")
'
'        feAnalistas.TextMatrix(i + 1, 12) = fMatAnalistas(i).cPersCod
'    Next i
'    feAnalistas.TopRow = 1
'    cmdExportarA.Enabled = True
'    cmdCerrarBPP.Enabled = True
'Else
'    cmdCerrarBPP.Enabled = False
'    MsgBox "No hay Datos - Analistas", vbInformation, "Aviso"
'End If
'End Sub
'Private Sub LlenaDatosCoordJA(ByVal pnMes As Integer, ByVal pnAnio As Integer)
'LimpiaFlex feCoordJA
'Call CargaDatosCoordJA(pnMes, pnAnio)
'
'If fMatCoordJA(0).Usuario <> "" Then
'    For i = 0 To UBound(fMatCoordJA)
'        feCoordJA.AdicionaFila
'        feCoordJA.TextMatrix(i + 1, 1) = fMatCoordJA(i).Agencia
'        feCoordJA.TextMatrix(i + 1, 2) = fMatCoordJA(i).Usuario
'        feCoordJA.TextMatrix(i + 1, 3) = fMatCoordJA(i).comite
'        feCoordJA.TextMatrix(i + 1, 4) = Format(fMatCoordJA(i).PorBonificacion * 100, "###," & String(15, "#") & "#0.00")
'        feCoordJA.TextMatrix(i + 1, 5) = Format(fMatCoordJA(i).SaldoCartera, "###," & String(15, "#") & "#0.00")
'        feCoordJA.TextMatrix(i + 1, 6) = Format(fMatCoordJA(i).SaldoCarteraCierre, "###," & String(15, "#") & "#0.00")
'        feCoordJA.TextMatrix(i + 1, 7) = Format(fMatCoordJA(i).PorSalCartVencJud * 100, "###," & String(15, "#") & "#0.00")
'        feCoordJA.TextMatrix(i + 1, 8) = Format(fMatCoordJA(i).PorcSalVenJud * 100, "###," & String(15, "#") & "#0.00")
'        feCoordJA.TextMatrix(i + 1, 9) = Format(fMatCoordJA(i).RendCaja * 100, "###," & String(15, "#") & "#0.00")
'        feCoordJA.TextMatrix(i + 1, 10) = Format(fMatCoordJA(i).RendAG * 100, "###," & String(15, "#") & "#0.00")
'        feCoordJA.TextMatrix(i + 1, 11) = Format(fMatCoordJA(i).BonoTotal, "###," & String(15, "#") & "#0.00")
'    Next i
'    feCoordJA.TopRow = 1
'    cmdExportarCJA.Enabled = True
'    cmdCerrarBPP.Enabled = True
'Else
'    cmdCerrarBPP.Enabled = False
'    MsgBox "No hay Datos - Coordinadores y Jefes de Agencia", vbInformation, "Aviso"
'End If
'
'End Sub
'
'Private Sub LlenaDatosJT(ByVal pnMes As Integer, ByVal pnAnio As Integer)
'LimpiaFlex feJefeTerritorial
'Call CargaDatosJT(pnMes, pnAnio)
'
'If fMatJefesTerritorial(0).Usuario <> "" Then
'    For i = 0 To UBound(fMatJefesTerritorial)
'        feJefeTerritorial.AdicionaFila
'        feJefeTerritorial.TextMatrix(i + 1, 1) = fMatJefesTerritorial(i).Zona
'        feJefeTerritorial.TextMatrix(i + 1, 2) = fMatJefesTerritorial(i).Usuario
'        feJefeTerritorial.TextMatrix(i + 1, 3) = Format(fMatJefesTerritorial(i).SaldoCartera, "###," & String(15, "#") & "#0.00")
'        feJefeTerritorial.TextMatrix(i + 1, 4) = Format(fMatJefesTerritorial(i).SaldoCarteraCierre, "###," & String(15, "#") & "#0.00")
'        feJefeTerritorial.TextMatrix(i + 1, 5) = Format(fMatJefesTerritorial(i).PorSalCartVencJud * 100, "###," & String(15, "#") & "#0.00")
'        feJefeTerritorial.TextMatrix(i + 1, 6) = Format(fMatJefesTerritorial(i).PorcSalVenJud * 100, "###," & String(15, "#") & "#0.00")
'        feJefeTerritorial.TextMatrix(i + 1, 7) = Format(fMatJefesTerritorial(i).RendCaja * 100, "###," & String(15, "#") & "#0.00")
'        feJefeTerritorial.TextMatrix(i + 1, 8) = Format(fMatJefesTerritorial(i).Rend * 100, "###," & String(15, "#") & "#0.00")
'        feJefeTerritorial.TextMatrix(i + 1, 9) = fMatJefesTerritorial(i).CantAge
'        feJefeTerritorial.TextMatrix(i + 1, 10) = fMatJefesTerritorial(i).AgeBoni
'        feJefeTerritorial.TextMatrix(i + 1, 11) = Format(fMatJefesTerritorial(i).PorBoni * 100, "###," & String(15, "#") & "#0.00")
'        feJefeTerritorial.TextMatrix(i + 1, 12) = Format(fMatJefesTerritorial(i).BonoTotal, "###," & String(15, "#") & "#0.00")
'    Next i
'    feJefeTerritorial.TopRow = 1
'    cmdExportarJT.Enabled = True
'    cmdCerrarBPP.Enabled = True
'Else
'    cmdCerrarBPP.Enabled = False
'    MsgBox "No hay Datos - Jefes Territoriales", vbInformation, "Aviso"
'End If
'End Sub
'
'Private Sub GenerarExcelAnalista()
'    Dim fs As Scripting.FileSystemObject
'    Dim xlsAplicacion As Excel.Application
'    Dim lsArchivo As String
'    Dim lsFile As String
'    Dim lsNomHoja As String
'    Dim xlsLibro As Excel.Workbook
'    Dim xlHoja1 As Excel.Worksheet
'    Dim lbExisteHoja As Boolean
'    Dim psArchivoAGrabarC As String
'    Dim lnExcel As Long
'
'    On Error GoTo ErrorGeneraExcelFormato
'
'    Set fs = New Scripting.FileSystemObject
'    Set xlsAplicacion = New Excel.Application
'
'    lsNomHoja = "BPP"
'    lsFile = "FormatoBPPAnalista"
'
'    lsArchivo = "\spooler\" & "BPPAnalistaGeneradoAlCierre" & Format(fgFecActual, "yyyymmdd") & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"
'    If fs.FileExists(App.path & "\FormatoCarta\" & lsFile & ".xls") Then
'        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsFile & ".xls")
'    Else
'        MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia"
'        Exit Sub
'    End If
'
'    'Activar Hoja
'    For Each xlHoja1 In xlsLibro.Worksheets
'       If xlHoja1.Name = lsNomHoja Then
'            xlHoja1.Activate
'         lbExisteHoja = True
'        Exit For
'       End If
'    Next
'
'    If lbExisteHoja = False Then
'        Set xlHoja1 = xlsLibro.Worksheets
'        xlHoja1.Name = lsNomHoja
'    End If
'
'    lnExcel = 5
'    Dim sFormatoConta As String
'    sFormatoConta = "_ * #,##0.00_ ;_ * -#,##0.00_ ;_ *  - ??_ ;_ @_ "
'
'    For i = 0 To UBound(fMatAnalistas)
'        xlHoja1.Cells(lnExcel + i, 1) = i + 1
'        xlHoja1.Cells(lnExcel + i, 2) = fMatAnalistas(i).Agencia
'        xlHoja1.Cells(lnExcel + i, 3).NumberFormat = "@"
'        xlHoja1.Cells(lnExcel + i, 3) = fMatAnalistas(i).cPersCod
'        xlHoja1.Cells(lnExcel + i, 4) = fMatAnalistas(i).comite
'        xlHoja1.Cells(lnExcel + i, 5) = fMatAnalistas(i).Usuario
'        xlHoja1.Cells(lnExcel + i, 6) = fMatAnalistas(i).NombreAnalista
'        xlHoja1.Cells(lnExcel + i, 7) = fMatAnalistas(i).Nivel
'        xlHoja1.Cells(lnExcel + i, 8) = fMatAnalistas(i).Categoria
'        xlHoja1.Cells(lnExcel + i, 9).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 9) = fMatAnalistas(i).MoraBase
'        xlHoja1.Cells(lnExcel + i, 10).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 10) = fMatAnalistas(i).MetaSaldo
'        xlHoja1.Cells(lnExcel + i, 11).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 11) = fMatAnalistas(i).MetaSaldoAG
'        xlHoja1.Cells(lnExcel + i, 12) = fMatAnalistas(i).MetaCliente
'        xlHoja1.Cells(lnExcel + i, 13) = fMatAnalistas(i).MetaClienteAG
'        xlHoja1.Cells(lnExcel + i, 14) = fMatAnalistas(i).MetaOperaciones
'        xlHoja1.Cells(lnExcel + i, 15) = fMatAnalistas(i).MetaOperacionesAG
'        xlHoja1.Cells(lnExcel + i, 16).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 16) = fMatAnalistas(i).MetaMora
'        xlHoja1.Cells(lnExcel + i, 17).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 17) = fMatAnalistas(i).MetaMoraAG
'        xlHoja1.Cells(lnExcel + i, 18).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 18) = fMatAnalistas(i).MetaRendimiento
'        xlHoja1.Cells(lnExcel + i, 19).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 19) = fMatAnalistas(i).MetaRendimientoAG
'        xlHoja1.Cells(lnExcel + i, 20).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 20) = fMatAnalistas(i).SaldoCapital
'        xlHoja1.Cells(lnExcel + i, 21).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 21) = fMatAnalistas(i).SaldoInicial
'        xlHoja1.Cells(lnExcel + i, 22).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 22) = fMatAnalistas(i).SaldoEntrante
'        xlHoja1.Cells(lnExcel + i, 23).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 23) = fMatAnalistas(i).SaldoSaliente
'        xlHoja1.Cells(lnExcel + i, 24).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 24) = fMatAnalistas(i).SIA
'        xlHoja1.Cells(lnExcel + i, 25).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 25) = fMatAnalistas(i).SCE
'        xlHoja1.Cells(lnExcel + i, 26).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 26) = fMatAnalistas(i).PPOSaldo
'        xlHoja1.Cells(lnExcel + i, 27).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 27) = fMatAnalistas(i).PTFSaldo
'        xlHoja1.Cells(lnExcel + i, 28).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 28) = fMatAnalistas(i).IXSaldo
'        xlHoja1.Cells(lnExcel + i, 29).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 29) = fMatAnalistas(i).PXSaldo
'        xlHoja1.Cells(lnExcel + i, 30) = fMatAnalistas(i).CantClientes
'        xlHoja1.Cells(lnExcel + i, 31) = fMatAnalistas(i).ClientesInicial
'        xlHoja1.Cells(lnExcel + i, 32) = fMatAnalistas(i).ClientesEntrantes
'        xlHoja1.Cells(lnExcel + i, 33) = fMatAnalistas(i).ClientesSalientes
'        xlHoja1.Cells(lnExcel + i, 34) = fMatAnalistas(i).NIC
'        xlHoja1.Cells(lnExcel + i, 35) = fMatAnalistas(i).NCE
'        xlHoja1.Cells(lnExcel + i, 36) = fMatAnalistas(i).PPOCliente
'        xlHoja1.Cells(lnExcel + i, 37).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 37) = fMatAnalistas(i).PTFCliente
'        xlHoja1.Cells(lnExcel + i, 38).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 38) = fMatAnalistas(i).IXCliente
'        xlHoja1.Cells(lnExcel + i, 39).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 39) = fMatAnalistas(i).PXCliente
'        xlHoja1.Cells(lnExcel + i, 40) = fMatAnalistas(i).NFO1
'        xlHoja1.Cells(lnExcel + i, 41) = fMatAnalistas(i).NOE1
'        xlHoja1.Cells(lnExcel + i, 42).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 42) = fMatAnalistas(i).PPOOpe1
'        xlHoja1.Cells(lnExcel + i, 43) = fMatAnalistas(i).NFO2
'        xlHoja1.Cells(lnExcel + i, 44) = fMatAnalistas(i).NOE2
'        xlHoja1.Cells(lnExcel + i, 45).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 45) = fMatAnalistas(i).PPOOpe2
'        xlHoja1.Cells(lnExcel + i, 46).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 46) = fMatAnalistas(i).PTFO
'        xlHoja1.Cells(lnExcel + i, 47).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 47) = fMatAnalistas(i).IXOperaciones
'        xlHoja1.Cells(lnExcel + i, 48).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 48) = fMatAnalistas(i).PXOperaciones
'        xlHoja1.Cells(lnExcel + i, 49).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 49) = fMatAnalistas(i).MF830
'        xlHoja1.Cells(lnExcel + i, 50).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 50) = fMatAnalistas(i).MI830
'        xlHoja1.Cells(lnExcel + i, 51).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 51) = fMatAnalistas(i).ME830
'        xlHoja1.Cells(lnExcel + i, 52).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 52) = fMatAnalistas(i).PP830
'        xlHoja1.Cells(lnExcel + i, 53).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 53) = fMatAnalistas(i).PTFMora
'        xlHoja1.Cells(lnExcel + i, 54).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 54) = fMatAnalistas(i).IXM830
'        xlHoja1.Cells(lnExcel + i, 55).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 55) = fMatAnalistas(i).PXMora
'        xlHoja1.Cells(lnExcel + i, 56).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 56) = fMatAnalistas(i).ICOB
'        xlHoja1.Cells(lnExcel + i, 57).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 57) = fMatAnalistas(i).PESP
'        xlHoja1.Cells(lnExcel + i, 58).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 58) = fMatAnalistas(i).CCC
'        xlHoja1.Cells(lnExcel + i, 59).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 59) = fMatAnalistas(i).RCA
'        xlHoja1.Cells(lnExcel + i, 60).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 60) = fMatAnalistas(i).IXRendimiento
'        xlHoja1.Cells(lnExcel + i, 61).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 61) = fMatAnalistas(i).MIMayor30
'        xlHoja1.Cells(lnExcel + i, 62).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 62) = fMatAnalistas(i).CJI
'        xlHoja1.Cells(lnExcel + i, 63).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 63) = fMatAnalistas(i).TMI
'        xlHoja1.Cells(lnExcel + i, 64).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 64) = fMatAnalistas(i).MFMayor30
'        xlHoja1.Cells(lnExcel + i, 65).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 65) = fMatAnalistas(i).CJF
'        xlHoja1.Cells(lnExcel + i, 66).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 66) = fMatAnalistas(i).TMF
'        xlHoja1.Cells(lnExcel + i, 67).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 67) = fMatAnalistas(i).BonoMeta
'        xlHoja1.Cells(lnExcel + i, 68).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 68) = fMatAnalistas(i).BonoPlus
'        xlHoja1.Cells(lnExcel + i, 69).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 69) = fMatAnalistas(i).BonoRendimiento
'        xlHoja1.Cells(lnExcel + i, 70).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 70) = fMatAnalistas(i).Penalidad
'        xlHoja1.Cells(lnExcel + i, 71).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 71) = fMatAnalistas(i).BonoTotal
'        xlHoja1.Cells(lnExcel + i, 72).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 72) = fMatAnalistas(i).TopeSaldo
'        xlHoja1.Cells(lnExcel + i, 73).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 73) = fMatAnalistas(i).PorcMinSaldo
'        xlHoja1.Cells(lnExcel + i, 74).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 74) = fMatAnalistas(i).TopeCliente
'        xlHoja1.Cells(lnExcel + i, 75).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 75) = fMatAnalistas(i).PorcMinCliente
'        xlHoja1.Cells(lnExcel + i, 76).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 76) = fMatAnalistas(i).TopeOperaciones
'        xlHoja1.Cells(lnExcel + i, 77).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 77) = fMatAnalistas(i).PorcMinOperaciones
'        xlHoja1.Cells(lnExcel + i, 78).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 78) = fMatAnalistas(i).RangoPerMora
'        xlHoja1.Cells(lnExcel + i, 79).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 79) = fMatAnalistas(i).TopeMora
'        xlHoja1.Cells(lnExcel + i, 80).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 80) = fMatAnalistas(i).PorcMinMora
'        xlHoja1.Cells(lnExcel + i, 81).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 81) = fMatAnalistas(i).MoraAcepMayor30
'        xlHoja1.Cells(lnExcel + i, 82).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 82) = fMatAnalistas(i).IntCobCMACM
'        xlHoja1.Cells(lnExcel + i, 83).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 83) = fMatAnalistas(i).SaldoCMACM
'        xlHoja1.Cells(lnExcel + i, 84).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 84) = fMatAnalistas(i).RendCMACM
'        xlHoja1.Cells(lnExcel + i, 85).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 85) = fMatAnalistas(i).MinRendCartera
'        xlHoja1.Cells(lnExcel + i, 86).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 86) = fMatAnalistas(i).FactorRend
'        xlHoja1.Cells(lnExcel + i, 87).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 87) = fMatAnalistas(i).PrimQuincena
'        xlHoja1.Cells(lnExcel + i, 88).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 88) = fMatAnalistas(i).SegunQuincena
'        xlHoja1.Cells(lnExcel + i, 89).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 89) = fMatAnalistas(i).SaldoPlus
'        xlHoja1.Cells(lnExcel + i, 90).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 90) = fMatAnalistas(i).ClientesPlus
'        xlHoja1.Cells(lnExcel + i, 91).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 91) = fMatAnalistas(i).OperacionesPlus
'        xlHoja1.Cells(lnExcel + i, 92).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 92) = fMatAnalistas(i).MoraPlus
'        xlHoja1.Cells(lnExcel + i, 93).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 93) = fMatAnalistas(i).Mora830IncialCierre
'        xlHoja1.Cells(lnExcel + i, 94).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 94) = fMatAnalistas(i).Mora830Entrante
'        xlHoja1.Cells(lnExcel + i, 95).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 95) = fMatAnalistas(i).Mora830Saliente
'        xlHoja1.Range(xlHoja1.Cells(lnExcel + i, 1), xlHoja1.Cells(lnExcel + i, 95)).Borders.LineStyle = 1
'    Next i
'    Call xlHoja1.Range(xlHoja1.Cells(3, 71), xlHoja1.Cells(lnExcel + i - 1, 71)).BorderAround(1, xlMedium)
'    Call xlHoja1.Range(xlHoja1.Cells(3, 67), xlHoja1.Cells(lnExcel + i - 1, 71)).BorderAround(1, xlMedium)
'    xlHoja1.Range(xlHoja1.Cells(5, 71), xlHoja1.Cells(lnExcel + i - 1, 71)).Interior.Color = RGB(255, 255, 0)
'    xlHoja1.Range(xlHoja1.Cells(lnExcel + i, 1), xlHoja1.Cells(lnExcel + i, 95)).EntireColumn.AutoFit
'
'
'
'    xlHoja1.SaveAs App.path & lsArchivo
'    psArchivoAGrabarC = App.path & lsArchivo
'    xlsAplicacion.Visible = True
'    xlsAplicacion.Windows(1).Visible = True
'    Set xlsAplicacion = Nothing
'    Set xlsLibro = Nothing
'    Set xlHoja1 = Nothing
'
'    MsgBox "Fromato Generado Satisfactoriamente en la ruta: " & psArchivoAGrabarC, vbInformation, "Aviso"
'
'    Exit Sub
'ErrorGeneraExcelFormato:
'    MsgBox err.Description, vbCritical, "Error a Generar El Formato Excel"
'End Sub
'
'Private Sub GenerarExcelCoordJA()
'    Dim fs As Scripting.FileSystemObject
'    Dim xlsAplicacion As Excel.Application
'    Dim lsArchivo As String
'    Dim lsFile As String
'    Dim lsNomHoja As String
'    Dim xlsLibro As Excel.Workbook
'    Dim xlHoja1 As Excel.Worksheet
'    Dim lbExisteHoja As Boolean
'    Dim psArchivoAGrabarC As String
'    Dim lnExcel As Long
'
'    On Error GoTo ErrorGeneraExcelFormato
'
'    Set fs = New Scripting.FileSystemObject
'    Set xlsAplicacion = New Excel.Application
'
'    lsNomHoja = "BPP"
'    lsFile = "FormatoBPPCoordinadorJefeAgencia"
'
'    lsArchivo = "\spooler\" & "BPPCoordinadorJefeAgenciaGeneradoAlCierre" & Format(fgFecActual, "yyyymmdd") & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"
'    If fs.FileExists(App.path & "\FormatoCarta\" & lsFile & ".xls") Then
'        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsFile & ".xls")
'    Else
'        MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia"
'        Exit Sub
'    End If
'
'    'Activar Hoja
'    For Each xlHoja1 In xlsLibro.Worksheets
'       If xlHoja1.Name = lsNomHoja Then
'            xlHoja1.Activate
'         lbExisteHoja = True
'        Exit For
'       End If
'    Next
'
'    If lbExisteHoja = False Then
'        Set xlHoja1 = xlsLibro.Worksheets
'        xlHoja1.Name = lsNomHoja
'    End If
'
'    lnExcel = 4
'    Dim sFormatoConta As String
'    sFormatoConta = "_ * #,##0.00_ ;_ * -#,##0.00_ ;_ *  - ??_ ;_ @_ "
'
'    For i = 0 To UBound(fMatCoordJA)
'        xlHoja1.Cells(lnExcel + i, 1) = i + 1
'        xlHoja1.Cells(lnExcel + i, 2) = fMatCoordJA(i).Agencia
'        xlHoja1.Cells(lnExcel + i, 3).NumberFormat = "@"
'        xlHoja1.Cells(lnExcel + i, 3) = fMatCoordJA(i).cPersCod
'        xlHoja1.Cells(lnExcel + i, 4) = fMatCoordJA(i).comite
'        xlHoja1.Cells(lnExcel + i, 5) = fMatCoordJA(i).Usuario
'        xlHoja1.Cells(lnExcel + i, 6) = fMatCoordJA(i).Nombre
'        xlHoja1.Cells(lnExcel + i, 7) = fMatCoordJA(i).Cargo
'        xlHoja1.Cells(lnExcel + i, 8) = fMatCoordJA(i).CantAnalista
'        xlHoja1.Cells(lnExcel + i, 9) = fMatCoordJA(i).AnalistaBoni
'        xlHoja1.Cells(lnExcel + i, 10).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 10) = fMatCoordJA(i).PorBonificacion
'        xlHoja1.Cells(lnExcel + i, 11).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 11) = fMatCoordJA(i).SaldoCartera
'        xlHoja1.Cells(lnExcel + i, 12).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 12) = fMatCoordJA(i).SaldoCarteraCierre
'        xlHoja1.Cells(lnExcel + i, 13).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 13) = fMatCoordJA(i).PorSalCartVencJud
'        xlHoja1.Cells(lnExcel + i, 14).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 14) = fMatCoordJA(i).PorcSalVenJud
'        xlHoja1.Cells(lnExcel + i, 15).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 15) = fMatCoordJA(i).RendCaja
'        xlHoja1.Cells(lnExcel + i, 16).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 16) = fMatCoordJA(i).RendAG
'        xlHoja1.Cells(lnExcel + i, 17).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 17) = fMatCoordJA(i).Tope
'        xlHoja1.Cells(lnExcel + i, 18).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 18) = fMatCoordJA(i).PorBonificacionConf
'        xlHoja1.Cells(lnExcel + i, 19).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 19) = fMatCoordJA(i).BonoTotal
'
'        xlHoja1.Range(xlHoja1.Cells(lnExcel + i, 1), xlHoja1.Cells(lnExcel + i, 19)).Borders.LineStyle = 1
'    Next i
'    xlHoja1.Range(xlHoja1.Cells(4, 19), xlHoja1.Cells(lnExcel + i - 1, 19)).Interior.Color = RGB(255, 255, 0)
'    xlHoja1.Range(xlHoja1.Cells(lnExcel + i, 1), xlHoja1.Cells(lnExcel + i, 19)).EntireColumn.AutoFit
'
'    xlHoja1.SaveAs App.path & lsArchivo
'    psArchivoAGrabarC = App.path & lsArchivo
'    xlsAplicacion.Visible = True
'    xlsAplicacion.Windows(1).Visible = True
'    Set xlsAplicacion = Nothing
'    Set xlsLibro = Nothing
'    Set xlHoja1 = Nothing
'
'    MsgBox "Fromato Generado Satisfactoriamente en la ruta: " & psArchivoAGrabarC, vbInformation, "Aviso"
'
'    Exit Sub
'ErrorGeneraExcelFormato:
'    MsgBox err.Description, vbCritical, "Error a Generar El Formato Excel"
'End Sub
'
'Private Sub GenerarExcelJT()
'    Dim fs As Scripting.FileSystemObject
'    Dim xlsAplicacion As Excel.Application
'    Dim lsArchivo As String
'    Dim lsFile As String
'    Dim lsNomHoja As String
'    Dim xlsLibro As Excel.Workbook
'    Dim xlHoja1 As Excel.Worksheet
'    Dim lbExisteHoja As Boolean
'    Dim psArchivoAGrabarC As String
'    Dim lnExcel As Long
'
'    On Error GoTo ErrorGeneraExcelFormato
'
'    Set fs = New Scripting.FileSystemObject
'    Set xlsAplicacion = New Excel.Application
'
'    lsNomHoja = "BPP"
'    lsFile = "FormatoBPPJefeTerritorial"
'
'    lsArchivo = "\spooler\" & "BPPJefeTerritorialGeneradoAlCierre" & Format(fgFecActual, "yyyymmdd") & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"
'    If fs.FileExists(App.path & "\FormatoCarta\" & lsFile & ".xls") Then
'        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsFile & ".xls")
'    Else
'        MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia"
'        Exit Sub
'    End If
'
'    'Activar Hoja
'    For Each xlHoja1 In xlsLibro.Worksheets
'       If xlHoja1.Name = lsNomHoja Then
'            xlHoja1.Activate
'         lbExisteHoja = True
'        Exit For
'       End If
'    Next
'
'    If lbExisteHoja = False Then
'        Set xlHoja1 = xlsLibro.Worksheets
'        xlHoja1.Name = lsNomHoja
'    End If
'
'    lnExcel = 4
'    Dim sFormatoConta As String
'    sFormatoConta = "_ * #,##0.00_ ;_ * -#,##0.00_ ;_ *  - ??_ ;_ @_ "
'
'    For i = 0 To UBound(fMatJefesTerritorial)
'        xlHoja1.Cells(lnExcel + i, 1) = i + 1
'        xlHoja1.Cells(lnExcel + i, 2).NumberFormat = "@"
'        xlHoja1.Cells(lnExcel + i, 2) = fMatJefesTerritorial(i).cPersCod
'        xlHoja1.Cells(lnExcel + i, 3) = fMatJefesTerritorial(i).Zona
'        xlHoja1.Cells(lnExcel + i, 4) = fMatJefesTerritorial(i).Usuario
'        xlHoja1.Cells(lnExcel + i, 5) = fMatJefesTerritorial(i).Nombre
'        xlHoja1.Cells(lnExcel + i, 6).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 6) = fMatJefesTerritorial(i).SaldoCartera
'        xlHoja1.Cells(lnExcel + i, 7).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 7) = fMatJefesTerritorial(i).SaldoCarteraCierre
'        xlHoja1.Cells(lnExcel + i, 8).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 8) = fMatJefesTerritorial(i).PorSalCartVencJud
'        xlHoja1.Cells(lnExcel + i, 9).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 9) = fMatJefesTerritorial(i).PorcSalVenJud
'        xlHoja1.Cells(lnExcel + i, 10).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 10) = fMatJefesTerritorial(i).RendCaja
'        xlHoja1.Cells(lnExcel + i, 11).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 11) = fMatJefesTerritorial(i).Rend
'        xlHoja1.Cells(lnExcel + i, 12) = fMatJefesTerritorial(i).CantAge
'        xlHoja1.Cells(lnExcel + i, 13) = fMatJefesTerritorial(i).AgeBoni
'        xlHoja1.Cells(lnExcel + i, 14).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 14) = fMatJefesTerritorial(i).PorBoni
'        xlHoja1.Cells(lnExcel + i, 15).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 15) = fMatJefesTerritorial(i).Tope
'        xlHoja1.Cells(lnExcel + i, 16).NumberFormat = "0.00%"
'        xlHoja1.Cells(lnExcel + i, 16) = fMatJefesTerritorial(i).PorConfJF
'        xlHoja1.Cells(lnExcel + i, 17).NumberFormat = sFormatoConta
'        xlHoja1.Cells(lnExcel + i, 17) = fMatJefesTerritorial(i).BonoTotal
'
'        xlHoja1.Range(xlHoja1.Cells(lnExcel + i, 1), xlHoja1.Cells(lnExcel + i, 17)).Borders.LineStyle = 1
'    Next i
'    xlHoja1.Range(xlHoja1.Cells(4, 17), xlHoja1.Cells(lnExcel + i - 1, 17)).Interior.Color = RGB(255, 255, 0)
'    xlHoja1.Range(xlHoja1.Cells(lnExcel + i, 1), xlHoja1.Cells(lnExcel + i, 17)).EntireColumn.AutoFit
'
'
'
'    xlHoja1.SaveAs App.path & lsArchivo
'    psArchivoAGrabarC = App.path & lsArchivo
'    xlsAplicacion.Visible = True
'    xlsAplicacion.Windows(1).Visible = True
'    Set xlsAplicacion = Nothing
'    Set xlsLibro = Nothing
'    Set xlHoja1 = Nothing
'
'    MsgBox "Fromato Generado Satisfactoriamente en la ruta: " & psArchivoAGrabarC, vbInformation, "Aviso"
'
'    Exit Sub
'ErrorGeneraExcelFormato:
'    MsgBox err.Description, vbCritical, "Error a Generar El Formato Excel"
'End Sub
'Private Sub cmdExportarA_Click()
'GenerarExcelAnalista
'End Sub
'
'Private Sub cmdExportarCJA_Click()
'GenerarExcelCoordJA
'End Sub
'
'Private Sub cmdExportarJT_Click()
'GenerarExcelJT
'End Sub
