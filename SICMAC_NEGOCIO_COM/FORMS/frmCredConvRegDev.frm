VERSION 5.00
Begin VB.Form frmCredConvRegDev 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Devolución por Creditos con Convenio"
   ClientHeight    =   5865
   ClientLeft      =   1560
   ClientTop       =   2610
   ClientWidth     =   11550
   Icon            =   "frmCredConvRegDev.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   11550
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   2805
      TabIndex        =   10
      Top             =   5430
      Width           =   1515
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "&Adicionar"
      Height          =   375
      Left            =   15
      TabIndex        =   9
      Top             =   5430
      Width           =   1350
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   9765
      TabIndex        =   8
      Top             =   5445
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1395
      TabIndex        =   7
      Top             =   5430
      Width           =   1365
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   360
      Left            =   8385
      TabIndex        =   6
      Top             =   5445
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Frame Frame2 
      Height          =   4665
      Left            =   45
      TabIndex        =   3
      Top             =   720
      Width           =   11445
      Begin VB.CommandButton cmdDel 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   500
         TabIndex        =   5
         Top             =   6210
         Visible         =   0   'False
         Width           =   1350
      End
      Begin SICMACT.FlexEdit fgConv 
         Height          =   4290
         Left            =   75
         TabIndex        =   4
         Top             =   225
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   7567
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "item-Codigo-Nombre Persona-Registro-Monto-N° Cheque-Moneda-Nuevo-NumRegistro-Cuenta"
         EncabezadosAnchos=   "650-1500-3500-0-1200-1500-1200-0-1200-1600"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-3-4-5-6-X-X-9"
         ListaControles  =   "0-1-0-2-0-0-3-0-0-0"
         BackColorControl=   65535
         BackColorControl=   65535
         BackColorControl=   65535
         EncabezadosAlineacion=   "C-L-L-C-R-L-L-L-L-L"
         FormatosEdit    =   "0-0-0-0-2-0-0-0-0-0"
         TextArray0      =   "item"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   645
         RowHeight0      =   300
         TipoBusPersona  =   1
         ForeColorFixed  =   -2147483635
      End
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   11430
      Begin SICMACT.TxtBuscar txtInst 
         Height          =   345
         Left            =   150
         TabIndex        =   1
         Top             =   225
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
      End
      Begin VB.Label lblInstDesc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2250
         TabIndex        =   2
         Top             =   240
         Width           =   8580
      End
   End
End
Attribute VB_Name = "frmCredConvRegDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdADD_Click()
'If ValidaInfo(fgConv.Row) = False And fgConv.Row > 1 Then
'    MsgBox "Existe un dato NO ingresado en el registro actual, por favor ingrese toda la informacion", vbInformation, "aviso"
'    fgConv.SetFocus
'    Exit Sub
'End If
'fgConv.AdicionaFila
'fgConv.SetFocus
'SendKeys "{enter}"
If txtInst = "" Then
    MsgBox "Seleccione la institución por favor", vbInformation, "Aviso"
    Exit Sub
End If
frmIngDevConv.Inicio 300408, Trim(txtInst), lblInstDesc, "Pendientes por devolver creditos Convenio"
CargaDatos Trim(txtInst)
End Sub

Private Sub cmdCancelar_Click()
Me.txtInst.Text = ""
Me.lblInstDesc.Caption = ""
Me.fgConv.Clear
Me.fgConv.Rows = 2
Me.fgConv.FormaCabecera
End Sub
'
'Private Sub cmdDel_Click()
'Dim oCred As COMDCredito.DCOMCredito
'Set oCred = New COMDCredito.DCOMCredito
'If MsgBox("Desea Eliminar la fila seleccionada [" & fgConv.Row & "]??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
'    If fgConv.TextMatrix(fgConv.Row, 7) <> "" Then
'        oCred.EliminaColocacConvenioRegDevolucion txtInst, fgConv.TextMatrix(fgConv.Row, 1), fgConv.TextMatrix(fgConv.Row, 3)
'    End If
'    Me.fgConv.EliminaFila fgConv.Row
'    fgConv.SoloFila = False
'End If
'Set oCred = Nothing
'End Sub

Private Sub CmdEliminar_Click()
Dim oCred As COMDCredito.DCOMCredito

Dim bExito As Boolean 'PASI20141222

'Modificado PASI20141222
'If MsgBox("Desea Eliminar la fila seleccionada [" & fgConv.row & "]??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
'    If fgConv.TextMatrix(fgConv.row, 7) <> "" Then
'        oCred.EliminaColocacConvenioRegDevolucion txtInst, fgConv.TextMatrix(fgConv.row, 1), fgConv.TextMatrix(fgConv.row, 3)
'    End If
'    Me.fgConv.EliminaFila fgConv.row
'    'fgConv.SoloFila = False
'End If

    If fgConv.TextMatrix(fgConv.row, 1) = "" Then
        MsgBox "Asegurese de haber seleccionado correctamente un registro para ser Extornado.", vbExclamation, "Aviso"
        Exit Sub
    End If
    If CDate(fgConv.TextMatrix(fgConv.row, 3)) < gdFecSis Then
        MsgBox "No se puede eliminar con una fecha diferente al del registro.", vbExclamation, "Aviso"
        Exit Sub
    End If
    If MsgBox("Esta Seguro de Eliminar el Registro Seleccionado [" & fgConv.row & "]??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        Set oCred = New COMDCredito.DCOMCredito
        bExito = oCred.EliminaColocacConvenioRegDevolucion(txtInst, fgConv.TextMatrix(fgConv.row, 1), fgConv.TextMatrix(fgConv.row, 3), fgConv.TextMatrix(fgConv.row, 8))
        If bExito Then
            Me.fgConv.EliminaFila fgConv.row
            Set oCred = Nothing
            MsgBox "El Registro ha sido eliminado con éxito.", vbInformation, "Aviso"
        Else
            MsgBox "El Registro no ha podido ser Eliminado.", vbInformation + vbCritical, "Aviso"
        End If
    End If
'end PASI
End Sub
Private Sub CmdGrabar_Click()
Dim i As Long
Dim j As Long
Dim oCred As COMDCredito.DCOMCredito
Set oCred = New COMDCredito.DCOMCredito

If Me.fgConv.TextMatrix(1, 1) = "" Then
    MsgBox "Registro no Ingresados", vbInformation, "aviso"
    Exit Sub
End If
For i = 1 To fgConv.Rows - 1
    If ValidaInfo(i) = False Then
        MsgBox "Existe DATOS NO VALIDOS en los informacion ingresada. POR FAVOR VERIFIQUE", vbInformation, "aviso"
        Exit Sub
    End If
Next
For i = 1 To fgConv.Rows - 1
    For j = 1 To fgConv.Rows - 1
        If i <> j Then
            If fgConv.TextMatrix(i, 1) = fgConv.TextMatrix(j, 1) And fgConv.TextMatrix(i, 3) = fgConv.TextMatrix(j, 3) Then
                MsgBox "Existe un cliente registrado con la misma fecha por favor rectifique", vbInformation, "aviso"
                Exit Sub
            End If
        End If
    Next
Next

If MsgBox("Desea grabar los registros ingresado??", vbQuestion + vbYesNo, "Aviso") = vbNo Then
    Exit Sub
End If
For i = 1 To fgConv.Rows - 1
    If fgConv.TextMatrix(i, 7) = "" Then
        'inserta registro nuevo
        oCred.InsertaColocacConvenioRegDevolucion Trim(txtInst), fgConv.TextMatrix(i, 1), _
                fgConv.TextMatrix(i, 3), gsCodAge, Right(Trim(fgConv.TextMatrix(i, 6)), 1), _
                fgConv.TextMatrix(i, 4), fgConv.TextMatrix(i, 5), , Trim(fgConv.TextMatrix(i, 9))
    Else
        'actualiza
        oCred.ActualizaColocacConvenioRegDevolucion Trim(txtInst), fgConv.TextMatrix(i, 1), _
                fgConv.TextMatrix(i, 3), gsCodAge, Right(Trim(fgConv.TextMatrix(i, 6)), 1), _
                fgConv.TextMatrix(i, 4), fgConv.TextMatrix(i, 5)
    End If
Next
Set oCred = Nothing
MsgBox "Grabacion Realizada con exito", vbInformation, "Aviso"
cmdCancelar_Click
End Sub

Sub CargaDatos(ByVal psPerscodInst As String)
Dim oCred As COMDCredito.DCOMCredito
Set oCred = New COMDCredito.DCOMCredito
Me.fgConv.rsFlex = oCred.GetColocacConvRegDevolucion(Trim(psPerscodInst))
Set oCred = Nothing
End Sub
Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim oPersonas As COMDPersona.DCOMPersonas
Dim oConstante As COMDConstantes.DCOMConstantes

Set oPersonas = New COMDPersona.DCOMPersonas

CentraForm Me
txtInst.RS = oPersonas.RecuperaPersonasTipo_Arbol_Agencia(gPersTipoConvenio, gsCodAge)
Set oConstante = New COMDConstantes.DCOMConstantes
fgConv.CargaCombo oConstante.RecuperaConstantes(gMoneda)

Set oPersonas = Nothing
Set oConstante = Nothing
End Sub

Private Sub txtInst_EmiteDatos()
Me.lblInstDesc = Trim(txtInst.psDescripcion)
CargaDatos Trim(txtInst)
Me.cmdADD.SetFocus
End Sub
Function ValidaInfo(ByVal pnRow As Long) As Boolean
Dim i As Integer
ValidaInfo = True
For i = 1 To fgConv.Cols - 1
    If i <> 5 And i <> 7 Then
        If fgConv.TextMatrix(pnRow, i) = "" Then
            ValidaInfo = False
            Exit For
        End If
    End If
Next
End Function
