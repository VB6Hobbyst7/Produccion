VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredBPPDetalleBPP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Detalle BPP"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10575
   Icon            =   "frmCredBPPDetalleBPP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Detalle del Bono"
      TabPicture(0)   =   "frmCredBPPDetalleBPP.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraBonoPlus"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fraBonoPlus 
         Caption         =   "Datos Generales"
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
         Height          =   5295
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   9855
         Begin TabDlg.SSTab sstDetalle 
            Height          =   3615
            Left            =   120
            TabIndex        =   2
            Top             =   1440
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   6376
            _Version        =   393216
            TabsPerRow      =   5
            TabHeight       =   520
            TabCaption(0)   =   "Bono Meta"
            TabPicture(0)   =   "frmCredBPPDetalleBPP.frx":0326
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lblBonoMeta"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label4"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "lblBonoPlus"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label8"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "feFactorBono"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "cmdConsideracion1"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "chkPagaMeta"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "chkPagaPlus"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).ControlCount=   8
            TabCaption(1)   =   "Bono Rendimiento"
            TabPicture(1)   =   "frmCredBPPDetalleBPP.frx":0342
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "feVariablesRend"
            Tab(1).Control(1)=   "cmdConsideracion2"
            Tab(1).Control(2)=   "chkPagaRendimiento"
            Tab(1).ControlCount=   3
            TabCaption(2)   =   "Penalidad"
            TabPicture(2)   =   "frmCredBPPDetalleBPP.frx":035E
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "feVariablesPen"
            Tab(2).ControlCount=   1
            Begin VB.CheckBox chkPagaRendimiento 
               Caption         =   "Paga por Rendimiento de Cartera"
               Height          =   195
               Left            =   -74640
               TabIndex        =   18
               Top             =   2760
               Width           =   3375
            End
            Begin VB.CheckBox chkPagaPlus 
               Caption         =   "Paga Plus"
               Height          =   195
               Left            =   3000
               TabIndex        =   17
               Top             =   3120
               Width           =   1335
            End
            Begin VB.CheckBox chkPagaMeta 
               Caption         =   "Paga Meta"
               Height          =   195
               Left            =   3000
               TabIndex        =   16
               Top             =   2760
               Width           =   1335
            End
            Begin VB.CommandButton cmdConsideracion1 
               Caption         =   "Ver Consideraciones "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   7560
               TabIndex        =   4
               Top             =   2760
               Width           =   1770
            End
            Begin VB.CommandButton cmdConsideracion2 
               Caption         =   "Ver Consideraciones "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   -67680
               TabIndex        =   3
               Top             =   2760
               Width           =   1770
            End
            Begin SICMACT.FlexEdit feFactorBono 
               Height          =   2175
               Left            =   240
               TabIndex        =   5
               Top             =   480
               Width           =   9120
               _ExtentX        =   16087
               _ExtentY        =   3836
               Cols0           =   7
               HighLight       =   1
               EncabezadosNombres=   "#-Factor-Meta-% Alcanz.-Inc. Meta-Inc. Plus-Aux"
               EncabezadosAnchos=   "0-3000-1200-1500-1500-1500-0"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnasAEditar =   "X-X-X-X-X-X-X"
               ListaControles  =   "0-0-0-0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-L-R-R-R-R-C"
               FormatosEdit    =   "0-0-2-2-2-2-0"
               TextArray0      =   "#"
               lbUltimaInstancia=   -1  'True
               TipoBusqueda    =   3
               lbBuscaDuplicadoText=   -1  'True
               RowHeight0      =   300
            End
            Begin SICMACT.FlexEdit feVariablesRend 
               Height          =   2175
               Left            =   -74640
               TabIndex        =   6
               Top             =   480
               Width           =   8760
               _ExtentX        =   15452
               _ExtentY        =   3836
               Cols0           =   4
               HighLight       =   1
               EncabezadosNombres=   "#-Variables-Valores-Aux"
               EncabezadosAnchos=   "0-6000-2300-0"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnasAEditar =   "X-X-X-X"
               ListaControles  =   "0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-L-R-C"
               FormatosEdit    =   "0-0-2-0"
               TextArray0      =   "#"
               lbUltimaInstancia=   -1  'True
               TipoBusqueda    =   3
               lbBuscaDuplicadoText=   -1  'True
               RowHeight0      =   300
            End
            Begin SICMACT.FlexEdit feVariablesPen 
               Height          =   2775
               Left            =   -74640
               TabIndex        =   23
               Top             =   600
               Width           =   8760
               _ExtentX        =   15452
               _ExtentY        =   4895
               Cols0           =   4
               HighLight       =   1
               EncabezadosNombres=   "#-Variables-Valores-Aux"
               EncabezadosAnchos=   "0-6000-2300-0"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnasAEditar =   "X-X-X-X"
               ListaControles  =   "0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-L-R-C"
               FormatosEdit    =   "0-0-2-0"
               TextArray0      =   "#"
               lbUltimaInstancia=   -1  'True
               TipoBusqueda    =   3
               lbBuscaDuplicadoText=   -1  'True
               RowHeight0      =   300
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Bono Plus:"
               Height          =   195
               Left            =   360
               TabIndex        =   22
               Top             =   3200
               Width           =   765
            End
            Begin VB.Label lblBonoPlus 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   1320
               TabIndex        =   21
               Top             =   3120
               Width           =   1455
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Bono Meta:"
               Height          =   195
               Left            =   360
               TabIndex        =   20
               Top             =   2800
               Width           =   825
            End
            Begin VB.Label lblBonoMeta 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   1320
               TabIndex        =   19
               Top             =   2760
               Width           =   1455
            End
         End
         Begin VB.Label lblMoraBase 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   4680
            TabIndex        =   15
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblCategoria 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2880
            TabIndex        =   14
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lblNivel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   13
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lblNombre 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2040
            TabIndex        =   12
            Top             =   360
            Width           =   4935
         End
         Begin VB.Label lblUsuario 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   11
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Mora Base:"
            Height          =   195
            Left            =   3840
            TabIndex        =   10
            Top             =   880
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Categoría:"
            Height          =   195
            Left            =   2040
            TabIndex        =   9
            Top             =   880
            Width           =   750
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Usuario:"
            Height          =   195
            Left            =   360
            TabIndex        =   8
            Top             =   420
            Width           =   585
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Nivel:"
            Height          =   195
            Left            =   480
            TabIndex        =   7
            Top             =   880
            Width           =   405
         End
      End
   End
End
Attribute VB_Name = "frmCredBPPDetalleBPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private fAnalista As AnalistaBPP
'Private i As Integer
'
'Private Sub chkPagaMeta_Click()
'chkPagaMeta.value = IIf(fAnalista.BonoMeta > 0, 1, 0)
'End Sub
'
'Private Sub chkPagaPlus_Click()
'chkPagaPlus.value = IIf(fAnalista.BonoPlus > 0, 1, 0)
'End Sub
'
'Private Sub chkPagaRendimiento_Click()
'chkPagaRendimiento.value = IIf(fAnalista.BonoRendimiento > 0, 1, 0)
'End Sub
'
'Private Sub cmdConsideracion1_Click()
'    frmCredBPPRequisitosPago.Inicio 1, fAnalista
'End Sub
'
'Private Sub cmdConsideracion2_Click()
'    frmCredBPPRequisitosPago.Inicio 2, fAnalista
'End Sub
'Public Sub Inicio(ByRef pAnalista As AnalistaBPP)
'fAnalista = pAnalista
'lblUsuario.Caption = fAnalista.Usuario
'lblNombre.Caption = fAnalista.NombreAnalista
'lblCategoria.Caption = fAnalista.Categoria
'lblNivel.Caption = fAnalista.Nivel
'lblMoraBase.Caption = Format(Round(fAnalista.MoraBase * 100, 2), "##0.00") & " %"
'
'lblBonoMeta.Caption = Format(fAnalista.BonoMeta, "###," & String(15, "#") & "#0.00")
'lblBonoPlus.Caption = Format(fAnalista.BonoPlus, "###," & String(15, "#") & "#0.00")
'
'chkPagaMeta.value = IIf(fAnalista.BonoMeta > 0, 1, 0)
'chkPagaPlus.value = IIf(fAnalista.BonoPlus > 0, 1, 0)
'chkPagaRendimiento.value = IIf(fAnalista.BonoRendimiento > 0, 1, 0)
'CargaConstantes
'sstDetalle.TabVisible(1) = True
'Me.Show 1
'End Sub
'
'
'Private Sub CargaConstantes()
'Dim oConst As COMDConstantes.DCOMConstantes
'Dim rsConst As ADODB.Recordset
'
'Set oConst = New COMDConstantes.DCOMConstantes
'
''Factores de Bono
'LimpiaFlex feFactorBono
'Set rsConst = oConst.RecuperaConstantes(7068)
'If Not (rsConst.EOF And rsConst.BOF) Then
'    For i = 0 To rsConst.RecordCount - 1
'        feFactorBono.AdicionaFila
'        feFactorBono.TextMatrix(i + 1, 1) = Trim(rsConst!cConsDescripcion)
'        Select Case CInt(Trim(rsConst!nConsValor))
'            Case 1:
'                feFactorBono.TextMatrix(i + 1, 2) = Format(fAnalista.MetaSaldoAG, "###," & String(15, "#") & "#0.00")
'                feFactorBono.TextMatrix(i + 1, 3) = Format(fAnalista.PorcMinSaldo * 100, "###," & String(15, "#") & "#0.00")
'                feFactorBono.TextMatrix(i + 1, 4) = Format(fAnalista.IXSaldo, "###," & String(15, "#") & "#0.00")
'                feFactorBono.TextMatrix(i + 1, 5) = Format(fAnalista.PXSaldo, "###," & String(15, "#") & "#0.00")
'            Case 2:
'                feFactorBono.TextMatrix(i + 1, 2) = Format(fAnalista.MetaClienteAG, "###," & String(15, "#") & "#0.00")
'                feFactorBono.TextMatrix(i + 1, 3) = Format(fAnalista.PorcMinCliente * 100, "###," & String(15, "#") & "#0.00")
'                feFactorBono.TextMatrix(i + 1, 4) = Format(fAnalista.IXCliente, "###," & String(15, "#") & "#0.00")
'                feFactorBono.TextMatrix(i + 1, 5) = Format(fAnalista.PXCliente, "###," & String(15, "#") & "#0.00")
'            Case 3:
'                feFactorBono.TextMatrix(i + 1, 2) = Format(fAnalista.MetaOperacionesAG, "###," & String(15, "#") & "#0.00")
'                feFactorBono.TextMatrix(i + 1, 3) = Format(fAnalista.PorcMinOperaciones * 100, "###," & String(15, "#") & "#0.00")
'                feFactorBono.TextMatrix(i + 1, 4) = Format(fAnalista.IXOperaciones, "###," & String(15, "#") & "#0.00")
'                feFactorBono.TextMatrix(i + 1, 5) = Format(fAnalista.PXOperaciones, "###," & String(15, "#") & "#0.00")
'            Case 4:
'                feFactorBono.TextMatrix(i + 1, 2) = Format(fAnalista.MetaMoraAG, "###," & String(15, "#") & "#0.00")
'                feFactorBono.TextMatrix(i + 1, 3) = Format(fAnalista.PorcMinMora * 100, "###," & String(15, "#") & "#0.00")
'                feFactorBono.TextMatrix(i + 1, 4) = Format(fAnalista.IXM830, "###," & String(15, "#") & "#0.00")
'                feFactorBono.TextMatrix(i + 1, 5) = Format(fAnalista.PXMora, "###," & String(15, "#") & "#0.00")
'        End Select
'        rsConst.MoveNext
'    Next i
'End If
'Set rsConst = Nothing
'
''Varibles de Rentabilidad
'LimpiaFlex feVariablesRend
'Set rsConst = oConst.RecuperaConstantes(7069)
'If Not (rsConst.EOF And rsConst.BOF) Then
'     For i = 0 To rsConst.RecordCount - 1
'        feVariablesRend.AdicionaFila
'        feVariablesRend.TextMatrix(i + 1, 1) = Trim(rsConst!cConsDescripcion)
'        Select Case CInt(Trim(rsConst!nConsValor))
'            Case 1:
'                feVariablesRend.TextMatrix(i + 1, 2) = Format(fAnalista.RendCMACM * 100, "##0.00") & "%"
'            Case 2:
'                feVariablesRend.TextMatrix(i + 1, 2) = Format(fAnalista.RCA * 100, "##0.00") & "%"
'            Case 3:
'                feVariablesRend.TextMatrix(i + 1, 2) = Format(fAnalista.ICOB, "###," & String(15, "#") & "#0.00")
'            Case 4:
'                feVariablesRend.TextMatrix(i + 1, 2) = Format(fAnalista.FactorRend * 100, "##0.00") & "%"
'            Case 5:
'                feVariablesRend.TextMatrix(i + 1, 2) = Format(fAnalista.IXRendimiento, "###," & String(15, "#") & "#0.00")
'        End Select
'
'        rsConst.MoveNext
'    Next i
'End If
'Set rsConst = Nothing
'
''Varibales de Penalidad
'LimpiaFlex feVariablesPen
'Set rsConst = oConst.RecuperaConstantes(7075)
'If Not (rsConst.EOF And rsConst.BOF) Then
'    For i = 0 To rsConst.RecordCount - 1
'        feVariablesPen.AdicionaFila
'        feVariablesPen.TextMatrix(i + 1, 1) = Trim(rsConst!cConsDescripcion)
'        Select Case CInt(Trim(rsConst!nConsValor))
'            Case 1:
'                feVariablesPen.TextMatrix(i + 1, 2) = Format(fAnalista.MoraAcepMayor30, "###," & String(15, "#") & "#0.00")
'            Case 2:
'                feVariablesPen.TextMatrix(i + 1, 2) = Format(fAnalista.MFMayor30, "###," & String(15, "#") & "#0.00")
'            Case 3:
'                feVariablesPen.TextMatrix(i + 1, 2) = Format(fAnalista.BonoMeta + fAnalista.BonoPlus + fAnalista.BonoRendimiento, "###," & String(15, "#") & "#0.00")
'            Case 4:
'                feVariablesPen.TextMatrix(i + 1, 2) = Format(fAnalista.Penalidad * 100, "###," & String(15, "#") & "#0.00") & " %"
'            Case 5:
'                feVariablesPen.TextMatrix(i + 1, 2) = Format(fAnalista.BonoTotal, "###," & String(15, "#") & "#0.00")
'        End Select
'        rsConst.MoveNext
'    Next i
'End If
'Set rsConst = Nothing
'Set oConst = Nothing
'End Sub
'
'Private Sub feFactorBono_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
'Cancel = ValidaFlex(feFactorBono, pnCol)
'End Sub
'
'Private Sub feVariablesPen_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
'Cancel = ValidaFlex(feVariablesPen, pnCol)
'End Sub
'
'Private Sub feVariablesRend_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
'Cancel = ValidaFlex(feVariablesRend, pnCol)
'End Sub
