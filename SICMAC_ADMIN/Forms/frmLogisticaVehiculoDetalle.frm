VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3A34C7E1-4D73-49FB-9EA9-C6D17991498B}#10.0#0"; "MiPanel.ocx"
Begin VB.Form frmLogisticaVehiculoDetalle 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5775
   ClientLeft      =   1980
   ClientTop       =   1995
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7725
   Begin VB.Frame fraFecha 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   2580
      TabIndex        =   26
      Top             =   3240
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton cmdCancela 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1440
         TabIndex        =   30
         Top             =   960
         Width           =   1170
      End
      Begin VB.CommandButton cmdGraba 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   180
         TabIndex        =   29
         Top             =   960
         Width           =   1170
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   1200
         TabIndex        =   27
         Top             =   420
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   480
         TabIndex        =   28
         Top             =   480
         Width           =   540
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   1455
         Left            =   0
         Top             =   0
         Width           =   2835
      End
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "&Quitar"
      Height          =   375
      Left            =   1380
      TabIndex        =   25
      Top             =   5220
      Width           =   1110
   End
   Begin VB.Frame Frame2 
      Height          =   795
      Left            =   120
      TabIndex        =   3
      Top             =   -60
      Width           =   7485
      Begin LabelBoxOCX.LabelBox lblBSCod 
         Height          =   315
         Left            =   960
         TabIndex        =   4
         Top             =   300
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         FteColor        =   -2147483630
         BeginProperty Fuente {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin LabelBoxOCX.LabelBox lblBSSerie 
         Height          =   315
         Left            =   4680
         TabIndex        =   5
         Top             =   300
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         FteColor        =   -2147483630
         BeginProperty Fuente {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "BSCod"
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
         Left            =   180
         TabIndex        =   7
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "BSSerie"
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
         TabIndex        =   6
         Top             =   360
         Width           =   690
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   7485
      Begin LabelBoxOCX.LabelBox lblTipo 
         Height          =   315
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         FteColor        =   -2147483630
         BeginProperty Fuente {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin LabelBoxOCX.LabelBox lblModelo 
         Height          =   315
         Left            =   960
         TabIndex        =   12
         Top             =   960
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         FteColor        =   -2147483630
         BeginProperty Fuente {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin LabelBoxOCX.LabelBox lblMarca 
         Height          =   315
         Left            =   960
         TabIndex        =   13
         Top             =   600
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         FteColor        =   -2147483630
         BeginProperty Fuente {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin LabelBoxOCX.LabelBox lblPlaca 
         Height          =   315
         Left            =   4680
         TabIndex        =   14
         Top             =   240
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         FteColor        =   -2147483630
         BeginProperty Fuente {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin LabelBoxOCX.LabelBox lblEstado 
         Height          =   315
         Left            =   4680
         TabIndex        =   15
         Top             =   960
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         FteColor        =   -2147483630
         BeginProperty Fuente {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin LabelBoxOCX.LabelBox lblColor 
         Height          =   315
         Left            =   4680
         TabIndex        =   22
         Top             =   600
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         FteColor        =   -2147483630
         BeginProperty Fuente {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Color"
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
         TabIndex        =   23
         Top             =   660
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
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
         TabIndex        =   20
         Top             =   1035
         Width           =   600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo "
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
         Left            =   180
         TabIndex        =   19
         Top             =   285
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Placa"
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
         TabIndex        =   18
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Modelo"
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
         Left            =   180
         TabIndex        =   17
         Top             =   1020
         Width           =   630
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
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
         Left            =   180
         TabIndex        =   16
         Top             =   675
         Width           =   540
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Top             =   5220
      Width           =   1230
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   5220
      Width           =   1110
   End
   Begin TabDlg.SSTab sstReg 
      Height          =   3585
      Left            =   120
      TabIndex        =   21
      Top             =   2100
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   6324
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   6
      TabHeight       =   670
      TabCaption(0)   =   " Asignación       "
      TabPicture(0)   =   "frmLogisticaVehiculoDetalle.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSHAsigna"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdLibera"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Kilometraje      "
      TabPicture(1)   =   "frmLogisticaVehiculoDetalle.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSHKm"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Carga             "
      TabPicture(2)   =   "frmLogisticaVehiculoDetalle.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSHCarga"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Incidencias    "
      TabPicture(3)   =   "frmLogisticaVehiculoDetalle.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdImprimir"
      Tab(3).Control(1)=   "cmdNewInc"
      Tab(3).Control(2)=   "MSHIncidencia"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "SOAT             "
      TabPicture(4)   =   "frmLogisticaVehiculoDetalle.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "MSHSOAT"
      Tab(4).ControlCount=   1
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   -70140
         TabIndex        =   34
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton cmdNewInc 
         Caption         =   "Registrar nueva Incidencia"
         Height          =   375
         Left            =   -72600
         TabIndex        =   33
         Top             =   3120
         Width           =   2415
      End
      Begin VB.CommandButton cmdLibera 
         Caption         =   "Liberar vehículo / conductor"
         Height          =   375
         Left            =   2400
         TabIndex        =   24
         Top             =   3120
         Width           =   2415
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHAsigna 
         Height          =   2505
         Left            =   120
         TabIndex        =   0
         Top             =   540
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4419
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHKm 
         Height          =   2505
         Left            =   -74880
         TabIndex        =   1
         Top             =   540
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4419
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHCarga 
         Height          =   2505
         Left            =   -74880
         TabIndex        =   2
         Top             =   540
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4419
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         AllowUserResizing=   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHIncidencia 
         Height          =   2505
         Left            =   -74880
         TabIndex        =   31
         Top             =   540
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4419
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         AllowUserResizing=   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHSOAT 
         Height          =   2505
         Left            =   -74880
         TabIndex        =   32
         Top             =   540
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4419
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
   End
End
Attribute VB_Name = "frmLogisticaVehiculoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vpHayCambios As Boolean
Dim cBSCod As String, cBSSerie As String, sSql As String

Public Sub Inicio(vBSCod As String, vBSSerie As String)
cBSCod = vBSCod
cBSSerie = vBSSerie
Me.Show 1
End Sub

Private Sub cmdImprimir_Click()
Dim i As Integer, n As Integer
Dim f As Integer, j As Integer
Dim cArchivo As String, v As Variant

f = FreeFile
cArchivo = "\Rep00.txt"
Open App.path + cArchivo For Output As #f

Print #f, ""
Print #f, Space(15) + "REPORTE DE INCIDENCIAS"
Print #f, ""
'Print #f, "-------------------------------------------------"
Print #f, "UNIDAD    : " + lblTipo.Text + " / " + lblMarca.Text
Print #f, "PLACA     : " + lblPlaca.Text
Print #f, "CONDUCTOR : "
Print #f, "-----------------------------------------------------"
Print #f, "FECHA      INCIDENCIA                     LUGAR"
Print #f, "-----------------------------------------------------"
n = MSHIncidencia.Rows - 1
For i = 1 To n
    Print #f, MSHIncidencia.TextMatrix(i, 1) + " " + JIZQ(MSHIncidencia.TextMatrix(i, 2), 30) + " " + JIZQ(MSHIncidencia.TextMatrix(i, 3), 20)
Next
Print #f, "-----------------------------------------------------"
Close #f
v = Shell("notepad " & App.path + cArchivo & "", vbNormalFocus)

End Sub

Private Sub cmdNewInc_Click()
FrmLogisticaVehiculoTipoIncidencia.Show 1
End Sub

Private Sub Form_Load()
vpHayCambios = False
lblBSCod = cBSCod
lblBSSerie = cBSSerie
DatosVehiculo
sstReg.Tab = 0
CargaLista sstReg.Tab
End Sub

Sub DatosVehiculo()
Dim rs As New ADODB.Recordset
Dim oConn As DConecta

Set oConn = New DConecta

If oConn.AbreConexion Then
   sSql = "select LV.cBSCod, LV.cBSSerie, LV.cModelo, LV.nAñoFab, LV.nEstado, " & _
   " LV.cNroMotor , LV.cPlaca, LT.cDescripcion, MV.cMarca, LC.cColor " & _
   " from LogisticaVehiculo LV inner join LogisticaTipoVehiculo LT on LT.nTipoVehiculo = LV.nTipoVehiculo " & _
   "                           inner join logisticaMarcav MV on MV.nMarca = LV.nMarca " & _
   "                           inner join logisticaColor LC on LC.nColor = LV.nColor " & _
   " Where LV.cBSCod = '" & cBSCod & "'  AND LV.cBSSerie = '" & cBSSerie & "' AND " & _
   "       LV.cFlag Is Null  " & _
   "  order by LT.cDescripcion "
   Set rs = oConn.CargaRecordSet(sSql)
   oConn.CierraConexion
   If Not rs.EOF Then
      lblTipo.Text = rs!cDescripcion
      lblModelo.Text = rs!cModelo
      lblMarca.Text = rs!cMarca
      lblPlaca.Text = rs!cPlaca
      lblColor.Text = rs!cColor
      
      lblEstado.BackColor = "&H8000000F"
      lblEstado.FontColor = "&H80C00012"
      Select Case rs!nEstado
          Case 1
               lblEstado.BackColor = "&H80000005"
               lblEstado.FontColor = "&H00C00000"
               lblEstado.Text = "LIBRE"
          Case 2
               lblEstado.Text = "ASIGNADO"
          Case 3
               lblEstado.Text = "REPARACION"
          Case Else
               lblEstado.Text = "LIBRE"
          
      End Select
   End If
End If

End Sub

Private Sub sstReg_Click(PreviousTab As Integer)
CargaLista sstReg.Tab
End Sub

Sub CargaLista(nTab As Integer)
Dim LV As DLogVehiculo
Dim rs As ADODB.Recordset
Dim i As Integer, k As Integer

Set LV = New DLogVehiculo


         'VERIFICACION DE OPERACIONES ---------------------------------
         If lblEstado.Text = "LIBRE" Then
            cmdAgregar.Enabled = False
            cmdQuitar.Enabled = False
            cmdLibera.Enabled = False
         Else
            cmdAgregar.Enabled = True
            cmdQuitar.Enabled = True
         End If
         '-------------------------------------------------------------

FormaMSHFlex nTab
Select Case nTab
    Case 0
         Set rs = LV.GetVehiculoAsignacion(Me.lblBSCod, Me.lblBSSerie)
         i = 0
         If Not rs.EOF Then
            Do While Not rs.EOF
               i = i + 1
               InsRow MSHAsigna, i
               MSHAsigna.TextMatrix(i, 0) = rs!cPersCod
               MSHAsigna.TextMatrix(i, 1) = Format(rs!dFecha, "DD/MM/YYYY")
               MSHAsigna.TextMatrix(i, 2) = Format(rs!dFechaFin, "DD/MM/YYYY")
               MSHAsigna.TextMatrix(i, 3) = rs!cPersNombre
               Select Case rs!cEstado
                   Case 0
                        MSHAsigna.TextMatrix(i, 4) = "  -"
                   Case 1
                        MSHAsigna.TextMatrix(i, 4) = "  -"
                   Case 2
                        MSHAsigna.row = i
                        For k = 1 To 4
                            MSHAsigna.Col = k
                            MSHAsigna.CellBackColor = "&H00DDFFFE"
                            MSHAsigna.CellForeColor = "&H000000C0"
                        Next
                        MSHAsigna.TextMatrix(i, 4) = "ASIGNADO"
                        lblEstado.Text = "ASIGNADO"
                        cmdAgregar.Enabled = False
                        cmdQuitar.Enabled = False
                        cmdLibera.Enabled = True
               End Select
               rs.MoveNext
            Loop
         End If
         If lblEstado.Text = "LIBRE" Then
            cmdAgregar.Enabled = True
            cmdQuitar.Enabled = True
            cmdLibera.Enabled = False
         Else
            cmdAgregar.Enabled = False
            cmdQuitar.Enabled = False
            cmdLibera.Enabled = True
         End If
    Case 1
         Set rs = LV.GetVehiculoKm(Me.lblBSCod, Me.lblBSSerie)
         i = 0
         If Not rs.EOF Then
            Do While Not rs.EOF
               i = i + 1
               InsRow MSHKm, i
               MSHKm.TextMatrix(i, 1) = Format(rs!dFecha, "DD/MM/YYYY")
               MSHKm.TextMatrix(i, 2) = rs!nKm
               MSHKm.TextMatrix(i, 3) = rs!nKmF
               rs.MoveNext
            Loop
         End If
    Case 2
         Set rs = LV.GetVehiculoCarga(cBSCod, cBSSerie)
         i = 0
         If Not rs.EOF Then
            Do While Not rs.EOF
               i = i + 1
               InsRow MSHCarga, i
               MSHCarga.TextMatrix(i, 1) = Format(rs!dFecha, "DD/MM/YYYY")
               MSHCarga.TextMatrix(i, 2) = rs!cDescripcion
               MSHCarga.TextMatrix(i, 3) = rs!NombreAgencia
               MSHCarga.TextMatrix(i, 4) = rs!cDestinoDesc
               rs.MoveNext
            Loop
         End If
    Case 3
         Set rs = LV.GetVehiculoIncidencia(Me.lblBSCod, Me.lblBSSerie)
         i = 0
         If Not rs.EOF Then
            Do While Not rs.EOF
               i = i + 1
               InsRow MSHIncidencia, i
               MSHIncidencia.TextMatrix(i, 1) = Format(rs!dFecha, "DD/MM/YYYY")
               MSHIncidencia.TextMatrix(i, 2) = rs!TipoIncidencia
               MSHIncidencia.TextMatrix(i, 3) = rs!cLugar
               MSHIncidencia.TextMatrix(i, 4) = rs!cDescripcion
               rs.MoveNext
            Loop
         End If
    Case 4
         Set rs = LV.GetVehiculoSoat(Me.lblBSCod, Me.lblBSSerie)
         i = 0
         If Not rs.EOF Then
            Do While Not rs.EOF
               i = i + 1
               InsRow MSHSOAT, i
               Me.MSHSOAT.TextMatrix(i, 1) = Format(rs!dInicio, "DD/MM/YYYY")
               Me.MSHSOAT.TextMatrix(i, 2) = Format(rs!dVencimiento, "DD/MM/YYYY")
               Me.MSHSOAT.TextMatrix(i, 3) = Format(rs!nMonto, "#0.00")
               rs.MoveNext
            Loop
         End If
         cmdAgregar.Enabled = True
         cmdQuitar.Enabled = True
End Select
Set LV = Nothing
Set rs = Nothing
End Sub


Sub FormaMSHFlex(nTab As Integer)
Select Case nTab
    Case 0
        With Me.MSHAsigna
            .Rows = 2:              .RowHeight(0) = 320: .Clear
            .RowHeight(1) = 8
            .ColWidth(0) = 0
            .ColWidth(1) = 950:    .TextMatrix(0, 1) = "  Desde":      .ColAlignment(1) = 4
            .ColWidth(2) = 950:    .TextMatrix(0, 2) = "  Hasta":      .ColAlignment(2) = 4
            .ColWidth(3) = 4100:   .TextMatrix(0, 3) = "Conductor"
            .ColWidth(4) = 950:    .TextMatrix(0, 4) = "Estado":       .ColAlignment(4) = 4
        End With
    Case 1
        With Me.MSHKm
            .Rows = 2:              .RowHeight(0) = 320: .Clear
            .RowHeight(1) = 8
            .ColWidth(0) = 0
            .ColWidth(1) = 1100:    .TextMatrix(0, 1) = "   Fecha":       .ColAlignment(1) = 4
            .ColWidth(2) = 2800:    .TextMatrix(0, 2) = "     Lectura Inicial Tacometro (KM)"
            .ColWidth(3) = 2800:    .TextMatrix(0, 3) = "     Lectura Final Tacometro (KM)"
        End With
    Case 2
        With Me.MSHCarga
            .Rows = 2:              .RowHeight(0) = 320: .Clear
            .RowHeight(1) = 8
            .ColWidth(0) = 0
            .ColWidth(1) = 1100:    .TextMatrix(0, 1) = "   Fecha":       .ColAlignment(1) = 4
            .ColWidth(2) = 2000:    .TextMatrix(0, 2) = "Descripcion"
            .ColWidth(3) = 1800:    .TextMatrix(0, 3) = "Agencia Destino"
            .ColWidth(4) = 2000:    .TextMatrix(0, 4) = "Destino"
        End With
'    Case 3
'        With Me.MSHPapeleta
'            .Rows = 2:              .RowHeight(0) = 320: .Clear
'            .RowHeight(1) = 8
'            .ColWidth(0) = 0
'            .ColWidth(1) = 1100:    .TextMatrix(0, 1) = "   Fecha":      .ColAlignment(1) = 4
'            .ColWidth(2) = 1200:    .TextMatrix(0, 2) = "Monto"
'            .ColWidth(3) = 3000:    .TextMatrix(0, 3) = "Descripcion"
'        End With
    Case 3
        With Me.MSHIncidencia
            .Rows = 2:              .RowHeight(0) = 320: .Clear
            .RowHeight(1) = 8
            .ColWidth(0) = 0
            .ColWidth(1) = 1100:    .TextMatrix(0, 1) = "   Fecha":      .ColAlignment(1) = 4
            .ColWidth(2) = 3000:    .TextMatrix(0, 2) = "Tipo de Incidencia"
            .ColWidth(3) = 1800:    .TextMatrix(0, 3) = "Lugar de Incidencia"
            .ColWidth(4) = 4000:    .TextMatrix(0, 4) = "Descripcion"
        End With
    Case 4
        With Me.MSHSOAT
            .Rows = 2:              .RowHeight(0) = 320: .Clear
            .RowHeight(1) = 8
            .ColWidth(0) = 0
            .ColWidth(1) = 1500:    .TextMatrix(0, 1) = "  Fecha Inicio":    .ColAlignment(1) = 4
            .ColWidth(2) = 1500:    .TextMatrix(0, 2) = "  Fecha Caduca":     .ColAlignment(2) = 4
            .ColWidth(3) = 2000:    .TextMatrix(0, 3) = "              Monto ($)"
            .ColWidth(0) = 0
        End With
End Select
End Sub

Sub InicioLimpia()
    MarcoSoat
    MarcoKm
    MarcoCarga
    MarcoPapeleta
    MarcoIncidencia
    MarcoAsigna
    lblBSCod = ""
    lblBSSerie = ""
    lblEstado = ""
    lblMarca = ""
    lblModelo = ""
    lblPlaca = ""
    lblTipo = ""
End Sub

Private Sub cmdAgregar_Click()
frmLogisticaDatos.Inicio sstReg.Tab, cBSCod, cBSSerie
If frmLogisticaDatos.vpHaGrabado Then
   CargaLista sstReg.Tab
   DatosVehiculo
   vpHayCambios = True
End If
End Sub

Private Sub CmdBuscarPlaca_Click()
Dim LV As DLogVehiculo
Dim rs As ADODB.Recordset
Set LV = New DLogVehiculo
Set rs = LV.GetDatoVehiculo(Trim(Me.TxtPlaca))
InicioLimpia
If Not (rs.EOF And rs.BOF) Then
    lblBSCod = rs!cBSCod
    lblBSSerie = rs!cBSSerie
    
    Select Case rs!nEstado
        Case "1"
            lblEstado = "LIBRE" & Space(100) & rs!nEstado
        Case "2"
            lblEstado = "ASIGNADO" & Space(100) & rs!nEstado
        Case "3"
            lblEstado = "REPARACION" & Space(100) & rs!nEstado
    End Select
    lblMarca = rs!Marcar
    lblModelo = rs!cModelo
    lblPlaca = rs!cPlaca
    lblTipo = rs!TipoV
    CargaSoatVehiculo
    CargaKm
    CargaCargaVehiculo
    CargaPapeleta
    CargaInsidencia
    CargaAsignacion
Else
    MsgBox "No existe Vehiculo", vbInformation, "AVISO"
    Set rs = Nothing
    Set LV = Nothing
    Exit Sub
End If
Set rs = Nothing
Set LV = Nothing
End Sub

Public Sub fEnfoque(ctrControl As Control)
    ctrControl.SelStart = 0
    ctrControl.SelLength = Len(ctrControl.Text)
End Sub

Sub CargaPapeleta()
Dim LV As DLogVehiculo
Dim rs As ADODB.Recordset

Set LV = New DLogVehiculo
Set rs = LV.GetVehiculoPapeleta(Me.lblBSCod, Me.lblBSSerie)
i = 1
While Not rs.EOF
    Me.MSHPapeleta.TextMatrix(i, 1) = Format(rs!dFecha, "DD/MM/YYYY")
    Me.MSHPapeleta.TextMatrix(i, 2) = Format(rs!nMonto, "#0.00")
    Me.MSHPapeleta.TextMatrix(i, 3) = rs!cDescripcion
    i = i + 1
    MSHPapeleta.Rows = MSHPapeleta.Rows + 1
    rs.MoveNext
Wend
If Not (rs.EOF And rs.BOF) Then Me.MSHPapeleta.Rows = MSHPapeleta.Rows - 1
Set LV = Nothing
Set rs = Nothing
End Sub

Sub CargaKm()
Dim LV As DLogVehiculo
Dim rs As ADODB.Recordset

Set LV = New DLogVehiculo
Set rs = LV.GetVehiculoKm(Me.lblBSCod, Me.lblBSSerie)
i = 1
While Not rs.EOF
    Me.MSHKm.TextMatrix(i, 1) = Format(rs!dFecha, "DD/MM/YYYY")
    Me.MSHKm.TextMatrix(i, 2) = rs!nKm
    i = i + 1
    MSHKm.Rows = MSHKm.Rows + 1
    rs.MoveNext
Wend
If Not (rs.EOF And rs.BOF) Then Me.MSHKm.Rows = MSHKm.Rows - 1
Set LV = Nothing
Set rs = Nothing
End Sub

Sub CargaAsignacion()
Dim LV As DLogVehiculo
Dim rs As ADODB.Recordset

Me.TxtFechaAsigna = gdFecSis
Me.TxtFechaFinAsignacion = gdFecSis
If Right(lblEstado, 1) = 1 Then
    Me.FraAsigna.Visible = True
    Me.FraDesasigna.Visible = False
Else
    Me.FraDesasigna.Visible = True
    Me.FraAsigna.Visible = False
End If


Set LV = New DLogVehiculo
Set rs = LV.GetVehiculoAsignacion(Me.lblBSCod, Me.lblBSSerie)
i = 1
MarcoAsigna
While Not rs.EOF
    Me.MSHAsigna.TextMatrix(i, 1) = Format(rs!dFecha, "DD/MM/YYYY")
    Me.MSHAsigna.TextMatrix(i, 2) = rs!cPersNombre
    i = i + 1
    MSHAsigna.Rows = MSHAsigna.Rows + 1
    rs.MoveNext
Wend
If Not (rs.EOF And rs.BOF) Then Me.MSHAsigna.Rows = MSHAsigna.Rows - 1
Set LV = Nothing
Set rs = Nothing
End Sub

Sub CargaSoatVehiculo()
Dim LV As DLogVehiculo
Dim rs As ADODB.Recordset

Set LV = New DLogVehiculo
Set rs = LV.GetVehiculoSoat(Me.lblBSCod, Me.lblBSSerie)
i = 1
While Not rs.EOF
    Me.MSHSOAT.TextMatrix(i, 1) = Format(rs!dInicio, "DD/MM/YYYY")
    Me.MSHSOAT.TextMatrix(i, 2) = Format(rs!dVencimiento, "DD/MM/YYYY")
    Me.MSHSOAT.TextMatrix(i, 3) = Format(rs!nMonto, "#0.00")
    i = i + 1
    MSHSOAT.Rows = MSHSOAT.Rows + 1
    rs.MoveNext
Wend
If Not (rs.EOF And rs.BOF) Then Me.MSHSOAT.Rows = MSHSOAT.Rows - 1
Set LV = Nothing
Set rs = Nothing
End Sub

Private Sub cmdNuevo_Click()
FrmLogisticaVehiculoTipoIncidencia.Show 1
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdGraba_Click()
Dim LV As DLogVehiculo
Dim opt As Integer

If RTrim(lblEstado.Text) = "LIBRE" Then
    MsgBox "El vehiculo se encuentra libre", vbInformation, "AVISO"
    Exit Sub
End If

'If ValFecha(Me.txtFechaFin) = False Then
'    Exit Sub
'End If

opt = MsgBox("¿ Esta seguro de Liberar el Vehículo ? " + Space(10), vbQuestion + vbYesNo, "AVISO")
Set LV = New DLogVehiculo
If opt = vbNo Then Exit Sub

Call LV.LiberaAsignacionVehiculo(cBSCod, cBSSerie, txtFecha)

'MsgBox "El vehículo y el Conductor han sido liberados" + Space(10), vbQuestion + vbYesNo, "Aviso de confirmación"

lblEstado.BackColor = "&H80000005"
lblEstado.FontColor = "&H00C00000"
lblEstado.Text = "LIBRE"
CargaLista sstReg.Tab
CmdSalir.Enabled = True
sstReg.Enabled = True
fraFecha.Visible = False
Set LV = Nothing
End Sub


Private Sub cmdLibera_Click()
If Len(Trim(MSHAsigna.TextMatrix(MSHAsigna.row, 1))) = 0 Then Exit Sub
cmdLibera.Enabled = False
cmdAgregar.Enabled = False
cmdQuitar.Enabled = False
CmdSalir.Enabled = False
sstReg.Enabled = False
txtFecha.Text = Date
fraFecha.Visible = True
txtFecha.SetFocus
End Sub

Private Sub cmdCancela_Click()
cmdAgregar.Enabled = True
cmdQuitar.Enabled = True
CmdSalir.Enabled = True
sstReg.Enabled = True
fraFecha.Visible = False
cmdLibera.Enabled = True
End Sub

Private Sub txtFecha_GotFocus()
txtFecha.SelStart = 0
txtFecha.SelLength = Len(Trim(txtFecha))
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdGraba.SetFocus
End If
End Sub
