VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMntParaViaticos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros de Viáticos: Mantenimiento"
   ClientHeight    =   4350
   ClientLeft      =   960
   ClientTop       =   2145
   ClientWidth     =   11400
   Icon            =   "frmMntParaViaticos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRutas 
      Caption         =   "Rutas/Destinos"
      Height          =   375
      Left            =   7200
      TabIndex        =   14
      Top             =   3885
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid grdV 
      Height          =   2775
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      ForeColor       =   -2147483641
      HeadLines       =   2
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "cCategCod"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "cDestinoCod"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "cTranspCod"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "cObjetoCod"
         Caption         =   "Concepto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "cObjetoDesc"
         Caption         =   "Descripción"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "cViaticoAfectoA"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "cViaticoAfectoADesc"
         Caption         =   "Afecto A"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "cViaticoAfectoTope"
         Caption         =   "Ver Tope"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "nViaticoImporte"
         Caption         =   "Importe"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column02 
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column03 
            ColumnAllowSizing=   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   3600
         EndProperty
         BeginProperty Column05 
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   1755.213
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   765
      Left            =   120
      TabIndex        =   6
      Top             =   90
      Width           =   11205
      Begin VB.ComboBox cboCategoria 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   270
         Width           =   1725
      End
      Begin VB.ComboBox cboDestino 
         Height          =   315
         Left            =   3630
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   270
         Width           =   4635
      End
      Begin VB.ComboBox cboTransporte 
         Height          =   315
         Left            =   9360
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   270
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Categoría"
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
         Left            =   150
         TabIndex        =   12
         Top             =   330
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Destino"
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
         Left            =   2910
         TabIndex        =   11
         Top             =   330
         Width           =   675
      End
      Begin VB.Label Label3 
         Caption         =   "Transporte"
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
         Left            =   8370
         TabIndex        =   10
         Top             =   330
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   9120
      TabIndex        =   5
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   360
      Left            =   1290
      TabIndex        =   1
      Top             =   3885
      Width           =   1100
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   360
      Left            =   2430
      TabIndex        =   2
      Top             =   3885
      Width           =   1100
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   360
      Left            =   3570
      TabIndex        =   3
      Top             =   3885
      Width           =   1100
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   360
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3885
      Width           =   1100
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   360
      Left            =   4710
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3885
      Width           =   1100
   End
End
Attribute VB_Name = "frmMntParaViaticos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCategoriaCod As String
Dim sDestinoCod   As String
Dim sTranspCod    As String

Dim rs As New ADODB.Recordset
Dim clsViaticos As DParaViaticos

Dim nFilas As Integer
Dim lConsulta As Boolean
'ARLO20170208****
Dim objPista As COMManejador.Pista
Dim lsPalabra, lsDesc As String
'************

Public Sub Inicio(plConsulta As Boolean)
lConsulta = plConsulta
Me.Show 1
End Sub

'Private Sub ManejaBoton(plOpcion As Boolean)
'cmdBuscar.Enabled = plOpcion
'If Not lConsulta Then
'   cmdNuevo.Enabled = plOpcion
'   cmdModificar.Enabled = plOpcion
'   cmdEliminar.Enabled = plOpcion
'End If
'cmdImprimir.Enabled = plOpcion
'fg.Enabled = plOpcion
'End Sub

Private Sub cboCategoria_Click()
MuestraParametros
End Sub

Private Sub cboDestino_Click()
MuestraParametros
End Sub

Private Sub cboTransporte_Click()
MuestraParametros
End Sub

Private Sub cmdbuscar_Click()
If Not rs.EOF Then
Dim clsBuscar As New ClassDescObjeto
   clsBuscar.BuscarDato rs, 0, "Viáticos", 3, 4
   Set clsBuscar = Nothing
   grdV.SetFocus
Else
   MsgBox "No existen datos para Buscar", vbInformation, "¡Aviso!"
End If
End Sub

Private Sub CargaVariables()
sCategoriaCod = Trim(Right(cboCategoria, 100))
sDestinoCod = Trim(Right(cboDestino, 100))
sTranspCod = Trim(Right(cboTransporte, 100))
End Sub

Private Sub cmdEliminar_Click()
On Error GoTo Salida
If Not rs.EOF Then
   If MsgBox("¿ Esta seguro de eliminar la Concepto de Viático ?", vbQuestion + vbYesNo, "Confirmación") = vbNo Then
      Exit Sub
   End If
   CargaVariables
   clsViaticos.EliminaParametros sCategoriaCod, sDestinoCod, sTranspCod, rs!cObjetoCod
            'ARLO20170208
            gsOpeCod = LogPistaParaMantViatico
            Set objPista = New COMManejador.Pista
            lsDesc = rs(4)
            lsPalabra = rs!cObjetoCod
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", "Se Elimino el Parametro de Viatico del Concepto : " & lsPalabra & " | Descripción : " & lsDesc
            Set objPista = Nothing
            '*******
   rs.Delete adAffectCurrent
   grdV.SetFocus
Else
   MsgBox "No existen datos para Eliminar", vbInformation, "¡Aviso!"
End If
Exit Sub
Salida:
   MsgBox TextErr(Err.Description), vbCritical, "Aviso"
End Sub

Private Sub cmdImprimir_Click()
Dim sTexto As String
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset
Dim Concp As String
Dim Desc As String
Dim Abrev As String
Dim VerT As String
Set clsViaticos = New DParaViaticos

sTexto = ""
Set rs1 = clsViaticos.CargaParametros(sCategoriaCod, sDestinoCod, sTranspCod, , adLockOptimistic, True)
'Set rs1 = grdV.DataSource

sTexto = sTexto & "Concepto" & Space(10) & "Descripcion" & Space(30) & "Afecto A." & Space(10) & "Ver Tope" & Space(10) & "Importe" & oImpresora.gPrnSaltoLinea
sTexto = sTexto & String(110, "=") & oImpresora.gPrnSaltoLinea

  
   Do While Not rs1.EOF
      Concp = rs1(3)
      Desc = rs1(4)
      Abrev = rs1(6)
      VerT = rs1(7)
      
      sTexto = sTexto & rs1(3) & Space(15 - Len(Concp)) & rs1(4) & Space(45 - Len(Desc)) & rs1(6) & Space(20 - Len(Abrev)) & rs1(7) & Space(18 - Len(VerT)) & rs1(8) & oImpresora.gPrnSaltoLinea
      rs1.MoveNext
   Loop
RSClose rs1
EnviaPrevio sTexto, "Parámetros de Viaticos", gnLinPage, False

            'ARLO20170208
            gsOpeCod = LogPistaParaMantViatico
            Set objPista = New COMManejador.Pista
            lsPalabra = rs!cObjetoCod
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Imprimio el Parametro de Viatico del Concepto "
            Set objPista = Nothing
            '*******


End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdRutas_Click()
    
    frmRutasViaticos.Show 1
    
    LlenaComboConstante gViaticosDestino, cboDestino
    
End Sub

Private Sub Form_Load()
LlenaComboConstante gViaticosCateg, cboCategoria
LlenaComboConstante gViaticosDestino, cboDestino
LlenaComboConstante gViaticosTransporte, cboTransporte
Set clsViaticos = New DParaViaticos

CentraForm Me
'Me.Icon = LoadPicture(App.path & gsRutaIcono)

If cboDestino.ListCount > 0 Then
   cboDestino.ListIndex = 0
End If
If cboTransporte.ListCount > 0 Then
   cboTransporte.ListIndex = 0
End If
If cboCategoria.ListCount > 0 Then
   cboCategoria.ListIndex = 0
End If

MuestraParametros

If lConsulta Then
   cmdNuevo.Visible = False
   cmdModificar.Visible = False
   cmdEliminar.Visible = False
   cmdImprimir.Left = cmdNuevo.Left
End If
End Sub

Function MuestraParametros() As Integer
CargaVariables
Set rs = clsViaticos.CargaParametros(sCategoriaCod, sDestinoCod, sTranspCod, , adLockOptimistic, True)
Set grdV.DataSource = rs
End Function

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdNuevo_Click()
glAceptar = False
CargaVariables
frmMntParaViaticosDat.Inicio sCategoriaCod, sDestinoCod, sTranspCod, "", True
If glAceptar Then
   MuestraParametros
   rs.Find "cObjetoCod = '" & frmMntParaViaticosDat.psCodigo & "'"
End If
grdV.SetFocus
End Sub

Private Sub cmdModificar_Click()
glAceptar = False
If Not rs.EOF Then
   CargaVariables
   frmMntParaViaticosDat.Inicio sCategoriaCod, sDestinoCod, sTranspCod, rs!cObjetoCod, False
   If glAceptar Then
      MuestraParametros
      rs.Find "cObjetoCod = '" & frmMntParaViaticosDat.psCodigo & "'"
   End If
   grdV.SetFocus
Else
   MsgBox "No existen datos para modificar", vbInformation, "¡Aviso!"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
RSClose rs
Set clsViaticos = Nothing
End Sub

Private Sub grdV_HeadClick(ByVal ColIndex As Integer)
If Not rs Is Nothing Then
   If Not rs.EOF Then
      rs.Sort = grdV.Columns(ColIndex).DataField
   End If
End If
End Sub
