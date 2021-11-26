VERSION 5.00
Begin VB.Form frmBuscarObjetoDB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Objeto"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9165
   Icon            =   "frmBuscarObjetoDB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   9165
   Begin SICMACT.FlexEdit FlexObjeto 
      Height          =   2175
      Left            =   2880
      TabIndex        =   8
      Top             =   1080
      Width           =   6015
      _extentx        =   10610
      _extenty        =   3836
      cols0           =   4
      highlight       =   1
      allowuserresizing=   3
      rowsizingmode   =   1
      encabezadosnombres=   "-Codigo-Direccion-Direccion"
      encabezadosanchos=   "400-1200-2500-3000"
      font            =   "frmBuscarObjetoDB.frx":030A
      fontfixed       =   "frmBuscarObjetoDB.frx":0336
      columnasaeditar =   "X-X-X-X"
      listacontroles  =   "0-0-0-0"
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      encabezadosalineacion=   "C-L-L-C"
      formatosedit    =   "0-0-0-0"
      lbultimainstancia=   -1  'True
      appearance      =   0
      colwidth0       =   405
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
   Begin VB.Frame Frame2 
      Caption         =   "Buscar por..."
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
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2655
      Begin VB.OptionButton Opt1 
         Caption         =   "Option1"
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
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2415
      End
      Begin VB.OptionButton Opt2 
         Caption         =   "Option2"
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
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2415
      End
      Begin VB.OptionButton Opt3 
         Caption         =   "Option3"
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
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2415
      End
   End
   Begin VB.TextBox TxtBuscar 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   600
      Width           =   4575
   End
   Begin VB.CommandButton CmdCancelar 
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
      Left            =   1440
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese Dato a Buscar :"
      Height          =   195
      Left            =   2880
      TabIndex        =   7
      Top             =   240
      Width           =   1680
   End
End
Attribute VB_Name = "frmBuscarObjetoDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'    Dim sBusqueda1 As String
'    Dim sBusqueda2 As String
'    Dim sBusqueda3 As String
'    Dim lsCodigo As String
'    Dim lsDescripcion As String
'    Dim lsSql As String
'
Public Sub Inicio(ByVal psSQL As String, ByVal sOption1 As Variant, ByVal sOption2 As Variant, ByVal sOption3 As Variant, ByRef psCodigo As String, ByRef psDescripcion As String, ByVal psTituloForm As String)
'    lsSql = psSql
'    Opt1.Visible = False
'    Opt2.Visible = False
'    Opt3.Visible = False
'    sBusqueda1 = ""
'    sBusqueda2 = ""
'    sBusqueda3 = ""
'
'    lsCodigo = psCodigo
'    lsDescripcion = psDescripcion
'
'    FlexObjeto.ColWidth(3) = 0
'
'    If sOption1(1) = True Then
'        Opt1.Visible = True
'        Opt1.value = 1
'        sBusqueda1 = sOption1(2)
'        Opt1.Caption = sOption1(3)
'    End If
'    If sOption2(1) = True Then
'        Opt2.Visible = True
'        sBusqueda2 = sOption2(2)
'        Opt2.Caption = sOption2(3)
'    End If
'    If sOption3(1) = True Then
'        'Opt3.Visible = True
'        sBusqueda3 = sOption3(2)
'        Opt3.Caption = sOption3(3)
'        FlexObjeto.ColWidth(3) = 250
'    End If
'    Me.Caption = psTituloForm
'    Me.Show 1
'
'    psCodigo = lsCodigo
'    psDescripcion = lsDescripcion
End Sub
'
'
'Private Sub cmdAceptar_Click()
'    Call FlexObjeto_Click
'End Sub
'
'Private Sub CmdCancelar_Click()
'    Unload Me
'End Sub
'
'Private Sub FlexObjeto_Click()
'Dim lnPosi As Integer
'    If FlexObjeto.TextMatrix(1, 1) = "" Then
'        Exit Sub
'    End If
'    lnPosi = FlexObjeto.row
'    lsCodigo = Trim(FlexObjeto.TextMatrix(lnPosi, 1))
'    lsDescripcion = Trim(FlexObjeto.TextMatrix(lnPosi, 2))
'    Unload Me
'End Sub
'
'Private Sub Form_Load()
'CentraForm Me
'End Sub
'
'Private Sub Opt1_Click()
'TxtBuscar = ""
'End Sub
'
'Private Sub Opt2_Click()
'TxtBuscar = ""
'End Sub
'Private Sub Opt3_Click()
'TxtBuscar = ""
'End Sub
'
'
'Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
'  Dim sSql As String
' If KeyAscii = 13 Then
'    If Opt1.value = True Then
'        sSql = lsSql & "Where " & sBusqueda1 & " like '" & TxtBuscar & "%'"
'    End If
'    If Opt2.value = True Then
'        sSql = lsSql & "Where " & sBusqueda2 & " like '" & TxtBuscar & "%'"
'    End If
'    If Opt3.value = True Then
'        sSql = lsSql & "Where " & sBusqueda3 & " like '" & TxtBuscar & "%'"
'    End If
'    Dim objConex As COMConecta.DCOMConecta
'    Dim oRs As ADODB.Recordset
'    Set objConex = New COMConecta.DCOMConecta
'    Set oRs = New ADODB.Recordset
'    objConex.AbreConexion
'        Set oRs = objConex.CargaRecordSet(sSql)
'    objConex.CierraConexion
'    Call LlenarFlexEdit(oRs)
'    oRs.Close
'    Set oRs = Nothing
'    Set objConex = Nothing
' End If
'
'End Sub
'Private Sub LlenarFlexEdit(ByVal oRs As ADODB.Recordset)
'LimpiaFlex FlexObjeto
'If oRs.BOF Or oRs.EOF Then
'    Exit Sub
'End If
'        FlexObjeto.TextMatrix(0, 1) = Opt1.Caption
'        FlexObjeto.TextMatrix(0, 2) = Opt2.Caption
'        FlexObjeto.ColWidth(3) = 0
'        If Trim(sBusqueda3) <> "" Then
'            FlexObjeto.TextMatrix(0, 3) = Opt3.Caption
'            FlexObjeto.ColWidth(3) = 250
'        End If
'    Do While Not oRs.EOF
'        FlexObjeto.AdicionaFila
'        FlexObjeto.TextMatrix(oRs.Bookmark, 1) = oRs!cCodigo
'        FlexObjeto.TextMatrix(oRs.Bookmark, 2) = oRs!cDescripcion
'        If FlexObjeto.ColWidth(3) > 0 Then
'            FlexObjeto.TextMatrix(oRs.Bookmark, 2) = oRs!CampoComp
'        End If
'        oRs.MoveNext
'    Loop
'End Sub
