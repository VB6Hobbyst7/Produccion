VERSION 5.00
Begin VB.Form frmGiroDetalleMantenimiento 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13005
   Icon            =   "frmGiroDetalleMantenimiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   13005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   4920
      Width           =   1100
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   4920
      Width           =   1100
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Giros Pendientes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12795
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtnumdoc 
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin SICMACT.FlexEdit grdGiro1 
         Height          =   3735
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   6588
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Cuenta-Apertura-Remitente-Monto-Destinatario-Agencia Dest."
         EncabezadosAnchos=   "400-2000-1200-3900-1100-3900-1600"
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
         EncabezadosAlineacion=   "C-C-L-L-L-L-L"
         FormatosEdit    =   "0-0-0-0-1-1-1"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label1 
         Caption         =   "Ingrese DNI:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmGiroDetalleMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ValidaDatosDocumentos As Boolean
Public nDocAnt As String
Dim nFila As Long

Private Sub CargaGirosPendientes(ByVal pDocNum As String)
Dim rsGiro As Recordset
Dim clsGiro As COMNCaptaServicios.NCOMCaptaServicios

Dim i As Integer
 
For i = 1 To nFila
      grdGiro1.EliminaFila (i)
Next i
nFila = 0

Set clsGiro = New COMNCaptaServicios.NCOMCaptaServicios
    Set rsGiro = clsGiro.GetGiroPendientexDoc(gsCodAge, pDocNum)
    If Not (rsGiro.EOF And rsGiro.BOF) Then
    
        Do While Not rsGiro.EOF
            If grdGiro1.TextMatrix(1, 1) <> "" Then grdGiro1.Rows = grdGiro1.Rows + 1
            nFila = grdGiro1.Rows - 1
            grdGiro1.TextMatrix(nFila, 0) = nFila
            grdGiro1.TextMatrix(nFila, 1) = rsGiro("cCtaCod")
            grdGiro1.TextMatrix(nFila, 2) = Format$(rsGiro("dPrdEstado"), "dd/mm/yyyy")
            grdGiro1.TextMatrix(nFila, 3) = PstaNombre(rsGiro("cRemitente"), False)
            grdGiro1.TextMatrix(nFila, 4) = Format$(rsGiro("nSaldo"), "#,##0.00")
            grdGiro1.TextMatrix(nFila, 5) = rsGiro("cDestinatario")
            grdGiro1.TextMatrix(nFila, 6) = Trim(rsGiro("cAgencia"))
            rsGiro.MoveNext
        Loop
    Else
        MsgBox "No hay giros pendientes registrados.", vbInformation, "Aviso"
        cmdAceptar.Enabled = False
        fraDatos.Enabled = False
    End If
    
End Sub

Private Sub cmdAceptar_Click()
Dim nFila As Long
nFila = grdGiro1.Row
With frmGiroMantenimiento
    .txtCuenta.Age = Mid(grdGiro1.TextMatrix(nFila, 1), 4, 2)
    .txtCuenta.Cuenta = Right(grdGiro1.TextMatrix(nFila, 1), 10)
    
End With
Unload Me
End Sub

Private Sub CmdBuscar_Click()
Dim j As Integer

If Len(txtnumdoc.Text) = 0 Then
        MsgBox "Falta Ingresar el Numero de Documento", vbInformation, "Aviso"
        txtnumdoc.SetFocus
        ValidaDatosDocumentos = False
        Exit Sub
End If

If Len(txtnumdoc.Text) <> gnNroDigitosDNI Then
      MsgBox "DNI No es de " & gnNroDigitosDNI & " digitos", vbInformation, "Aviso"
      txtnumdoc.SetFocus
      ValidaDatosDocumentos = False
      Exit Sub
End If

For j = 1 To Len(Trim(txtnumdoc.Text))
    If (Mid(txtnumdoc.Text, j, 1) < "0" Or Mid(txtnumdoc.Text, j, 1) > "9") Then
       MsgBox "Uno de los Digitos del DNI no es un Numero", vbInformation, "Aviso"
       txtnumdoc.SetFocus
       ValidaDatosDocumentos = False
       Exit Sub
    End If
Next j
CargaGirosPendientes Me.txtnumdoc
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
ValidaDatosDocumentos = True
End Sub

Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
Else
    KeyAscii = NumerosEnteros(KeyAscii) 'MADM 20090928
End If
End Sub
