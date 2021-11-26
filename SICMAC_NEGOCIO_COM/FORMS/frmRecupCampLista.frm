VERSION 5.00
Begin VB.Form frmRecupCampLista 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Campañas Registradas"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8085
   Icon            =   "frmRecupCampLista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
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
      Left            =   5520
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
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
      Left            =   6840
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin SICMACT.FlexEdit feCampanas 
      Height          =   3705
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   6535
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "#-Campaña-Desde-Hasta-Cod"
      EncabezadosAnchos=   "400-3000-2000-2000-0"
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
      ColumnasAEditar =   "X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0"
      BackColor       =   16777215
      EncabezadosAlineacion=   "C-L-C-C-C"
      FormatosEdit    =   "0-0-5-5-0"
      CantEntero      =   12
      TextArray0      =   "#"
      SelectionMode   =   1
      TipoBusqueda    =   6
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      CellBackColor   =   16777215
   End
End
Attribute VB_Name = "frmRecupCampLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmRecupCampLista
'** Descripción : Formulario para listar las campaña de recuperaciones
'**               Creado segun TI-ERS035-2015
'** Creación    : WIOR, 20150522 09:00:00 AM
'**********************************************************************************************

Option Explicit
Private fsCod As String
Public Function Inicio() As Long
Call CargarCampanas
Me.Show 1
If Trim(fsCod) = "" Then
    fsCod = "0"
End If
Inicio = fsCod
End Function
Private Sub CmdAceptar_Click()
fsCod = feCampanas.TextMatrix(feCampanas.row, 4)

If Trim(fsCod) = "" Then
    MsgBox "Favor de seleccionar una Campaña", vbInformation, "Aviso"
    Exit Sub
End If

Unload Me
End Sub

Private Sub cmdSalir_Click()
fsCod = ""
Unload Me
End Sub

Private Sub CargarCampanas()
Dim i As Long
Dim RsDatos As ADODB.Recordset
Dim oDCredito As COMDCredito.DCOMCredito


LimpiaFlex feCampanas
Set oDCredito = New COMDCredito.DCOMCredito

Set RsDatos = oDCredito.RecuperarCampanaRecup
If Not RsDatos Is Nothing Then
    If Not (RsDatos.EOF And RsDatos.BOF) Then
        For i = 1 To RsDatos.RecordCount
            feCampanas.AdicionaFila
            feCampanas.TextMatrix(i, 1) = Trim(RsDatos!cNombre)
            feCampanas.TextMatrix(i, 2) = Format(RsDatos!dfechaini, "dd/mm/yyyy")
            feCampanas.TextMatrix(i, 3) = Format(RsDatos!dfechafin, "dd/mm/yyyy")
            feCampanas.TextMatrix(i, 4) = Trim(RsDatos!nId)
            RsDatos.MoveNext
        Next i
    End If
End If

Set oDCredito = Nothing
Set RsDatos = Nothing
End Sub

Private Sub feCampanas_DblClick()
CmdAceptar_Click
End Sub
