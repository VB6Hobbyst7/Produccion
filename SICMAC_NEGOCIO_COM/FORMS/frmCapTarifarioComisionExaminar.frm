VERSION 5.00
Begin VB.Form frmCapTarifarioExaminar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Examinar"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5445
   Icon            =   "frmCapTarifarioComisionExaminar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   300
      Left            =   6300
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2070
      Width           =   645
   End
   Begin SICMACT.FlexEdit grdVersiones 
      Height          =   1815
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   3201
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "nId-nIdVersion-nVersion-Version-Fecha Registro-Glosa-00"
      EncabezadosAnchos=   "0-0-0-900-1500-2500-0"
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-L-C-L-C"
      FormatosEdit    =   "0-0-0-0-0-0-0"
      TextArray0      =   "nId"
      SelectionMode   =   1
      lbUltimaInstancia=   -1  'True
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton btnSeleccionar 
      Caption         =   "Seleccionar"
      Height          =   300
      Left            =   2003
      TabIndex        =   1
      Top             =   2070
      Width           =   1455
   End
End
Attribute VB_Name = "frmCapTarifarioExaminar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************************
'* NOMBRE         : frmCapTarifarioExaminar
'* DESCRIPCION    : Proyecto - Tarifario Versionado - Examinar las versiones de tasas y comisiones
'* CREACION       : RIRO, 20160420 10:00 AM
'************************************************************************************************************

Option Explicit

Private rs As ADODB.Recordset
Private bSeleccion As Boolean
Private nId_ As Integer
Private nTipo_ As Integer '1=Comision, 2=tasa
'

Public Property Get nTipo() As Integer
     nTipo = nTipo_
End Property
Public Property Let nTipo(ByVal vNewValue As Integer)
     nTipo_ = vNewValue
End Property
Public Property Get rsExaminar() As ADODB.Recordset
    Set rsExaminar = rs
End Property
Public Property Let rsExaminar(ByVal vNewValue As ADODB.Recordset)
    Set rs = vNewValue
End Property
Private Sub btnSeleccionar_Click()
    bSeleccion = True
    nId_ = CInt(grdVersiones.TextMatrix(grdVersiones.row, 1))
    Unload Me
End Sub
Public Function bRespuesta() As Boolean
    bRespuesta = bSeleccion
End Function
Public Function Id() As Integer
    Id = nId_
End Function
Private Sub cmdsalir_Click()
Unload Me
End Sub
Private Sub Form_Activate()
grdVersiones.SetFocus
End Sub
Private Sub Form_Initialize()
    bSeleccion = False
    nId_ = -1
End Sub
Private Sub cargarVersiones()
    Dim i As Integer
    If nTipo = 1 Then ' Comisiones
        If Not rsExaminar Is Nothing Then
            If rsExaminar.RecordCount > 0 Then
                If (Not rsExaminar.EOF And Not rsExaminar.BOF) Then
                    Do While (Not rsExaminar.EOF And Not rsExaminar.BOF)
                    i = i + 1
                    grdVersiones.AdicionaFila
                    grdVersiones.TextMatrix(i, 0) = i
                    grdVersiones.TextMatrix(i, 1) = rsExaminar!nIdComision
                    grdVersiones.TextMatrix(i, 2) = rsExaminar!nVersion
                    grdVersiones.TextMatrix(i, 3) = rsExaminar!cVersion
                    grdVersiones.TextMatrix(i, 4) = rsExaminar!dFechaRegistro
                    grdVersiones.TextMatrix(i, 5) = rsExaminar!cGlosa
                    rsExaminar.MoveNext
                    Loop
                End If
            End If
        End If
    ElseIf nTipo = 2 Then 'Tasas
        If Not rsExaminar Is Nothing Then
            If rsExaminar.RecordCount > 0 Then
                If (Not rsExaminar.EOF And Not rsExaminar.BOF) Then
                    Do While (Not rsExaminar.EOF And Not rsExaminar.BOF)
                    i = i + 1
                    grdVersiones.AdicionaFila
                    grdVersiones.TextMatrix(i, 0) = i
                    grdVersiones.TextMatrix(i, 1) = rsExaminar!nIdTarifarioTasaCab
                    grdVersiones.TextMatrix(i, 2) = rsExaminar!nVersion
                    grdVersiones.TextMatrix(i, 3) = rsExaminar!cVersion
                    grdVersiones.TextMatrix(i, 4) = rsExaminar!dFechaRegistro
                    grdVersiones.TextMatrix(i, 5) = rsExaminar!cGlosa
                    rsExaminar.MoveNext
                    Loop
                End If
            End If
        End If
    End If
End Sub
Private Sub Form_Load()
cargarVersiones
End Sub
Private Sub grdVersiones_DblClick()
    bSeleccion = True
    nId_ = CInt(grdVersiones.TextMatrix(grdVersiones.row, 1))
    Unload Me
End Sub

Private Sub grdVersiones_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
grdVersiones.SetFocus
End If
End Sub
