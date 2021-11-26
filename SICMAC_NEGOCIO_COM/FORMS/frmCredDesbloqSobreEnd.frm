VERSION 5.00
Begin VB.Form frmCredDesbloqSobreEnd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desbloqueo de Sobre Endeudamiento"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11250
   Icon            =   "frmCredDesbloqSobreEnd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
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
      Left            =   1560
      TabIndex        =   7
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdAutorizar 
      Caption         =   "Autorizar"
      Enabled         =   0   'False
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
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Frame fraEvaluacion 
      Caption         =   "Evaluación de SobreEndeudamiento"
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
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   11055
      Begin SICMACT.FlexEdit feCodigos 
         Height          =   2385
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   10650
         _extentx        =   18785
         _extenty        =   4207
         cols0           =   9
         highlight       =   1
         allowuserresizing=   2
         encabezadosnombres=   "#-Codigos-Resultado-Detalle-Plan de Mitigación-nCodigo-nResultado-bGrabado-MensajeDet"
         encabezadosanchos=   "400-1000-2000-2000-5000-0-0-0-0"
         font            =   "frmCredDesbloqSobreEnd.frx":030A
         font            =   "frmCredDesbloqSobreEnd.frx":0336
         font            =   "frmCredDesbloqSobreEnd.frx":0362
         font            =   "frmCredDesbloqSobreEnd.frx":038E
         font            =   "frmCredDesbloqSobreEnd.frx":03BA
         fontfixed       =   "frmCredDesbloqSobreEnd.frx":03E6
         tipobusqueda    =   6
         columnasaeditar =   "X-X-X-3-4-X-X-X-X"
         listacontroles  =   "0-0-0-1-3-0-0-0-0"
         backcolor       =   16777215
         encabezadosalineacion=   "C-C-L-L-L-C-C-C-C"
         formatosedit    =   "0-0-0-0-0-0-0-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbbuscaduplicadotext=   -1
         colwidth0       =   405
         rowheight0      =   300
         cellbackcolor   =   16777215
      End
   End
   Begin VB.Frame fraDatosCliente 
      Caption         =   "Datos del Cliente"
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
      TabIndex        =   0
      Top             =   120
      Width           =   11055
      Begin VB.TextBox txtEvalRisgSobre 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7200
         TabIndex        =   13
         Top             =   960
         Width           =   3015
      End
      Begin VB.CommandButton cmdExaminar 
         Caption         =   "Examinar"
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
         Left            =   7920
         TabIndex        =   8
         Top             =   200
         Width           =   1335
      End
      Begin VB.Label lblEvalRisSob 
         Caption         =   "Evaluación de Riesgo de Sobreendeudamiento"
         Height          =   255
         Left            =   3720
         TabIndex        =   12
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label lblCredito 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "N° Crédito:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DOI:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   330
      End
      Begin VB.Label txtNroDOI 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblNomCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   6735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmCredDesbloqSobreEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************
'** Nombre      : frmCredDesbloqSobreEnd
'** Descripción : Formulario para realizar el desbloqueo de los créditos que entraron por sobreendeudamiento del cliente
'**               Creado según TI-ERS038-2016
'** Creación    : WIOR, 20160623 09:00:00 AM
'*********************************************************************************************
Option Explicit
Private fnId As Long

Private Sub LimpiarDatos()
lblNomCliente.Caption = ""
txtNroDOI.Caption = ""
lblCredito.Caption = ""
txtEvalRisgSobre.Text = "" 'JOEP 13092016
LimpiaFlex feCodigos
cmdAutorizar.Enabled = False
fnId = 0
End Sub

Private Sub cmdAutorizar_Click()
Dim oNCredito As COMNCredito.NCOMCredito
Dim MatCodigos As Variant
Dim nCantCodigos As Integer
Dim bGrabar As Boolean
Dim i As Integer
If Not ValidaDatos Then Exit Sub
If MsgBox("Estás seguro de autorizar el desbloqueo del crédito?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub

On Error GoTo ErrorGrabarDatos

ReDim MatCodigos(2, 0)
nCantCodigos = 0
For i = 1 To feCodigos.rows - 1
    If CInt(feCodigos.TextMatrix(i, 7)) = 0 Then
        ReDim Preserve MatCodigos(2, 0 To nCantCodigos)
        MatCodigos(0, nCantCodigos) = CInt(feCodigos.TextMatrix(i, 5))
        MatCodigos(1, nCantCodigos) = CInt(feCodigos.TextMatrix(i, 6))
        MatCodigos(2, nCantCodigos) = Trim(feCodigos.TextMatrix(i, 4))
        
        nCantCodigos = nCantCodigos + 1
    End If
Next i

Set oNCredito = New COMNCredito.NCOMCredito
bGrabar = False
bGrabar = oNCredito.SobreEndAutorizarDesbloq(Trim(lblCredito.Caption), fnId, MatCodigos, gdFecSis, gsCodUser, gsCodCargo)

If bGrabar Then
    MsgBox "Los datos se grabaron correctamente.", vbInformation, "Aviso"
    LimpiarDatos
Else
     MsgBox "Hubo errores al grabar las datos", vbError, "Error"
End If

Set oNCredito = Nothing

Exit Sub
ErrorGrabarDatos:
MsgBox Err.Number & " - " & Err.Description, vbError, "Error En Proceso"
End Sub

Private Sub cmdCancelar_Click()
LimpiarDatos
End Sub

Private Sub cmdExaminar_Click()
Dim sCta As String
Dim oDCredito As COMDCredito.DCOMCredito
Dim rsCredito As ADODB.Recordset
Dim sMensaje As String
Dim i As Integer
sMensaje = ""
i = 0
fnId = 0
sCta = frmCredPersEstado.InicioDesBloqSobreEnd("Credito a Desbloquear - Sobreendeudamiento", gsCodCargo)

If Len(Trim(sCta)) > 0 Then
'JOEP20190227 CP
feCodigos.Enabled = True
'JOEP20190227 CP
    Set oDCredito = New COMDCredito.DCOMCredito
    Set rsCredito = oDCredito.SobreEndDatosCreditoADesbloq(sCta, gsCodCargo)
     If Not (rsCredito.EOF And rsCredito.BOF) Then
        If CInt(rsCredito!nEstado) = 3 Then
            MsgBox "Crédito ya fue desbloqueado.", vbInformation, "Aviso"
            Exit Sub
        End If
        
        If CInt(rsCredito!nDesbloqARealizar) = 0 Then
            MsgBox "No cuentas con permisos para desbloquear este crédito.", vbInformation, "Aviso"
            Exit Sub
        End If
        
'        If CInt(rsCredito!nDesbloq) = CInt(rsCredito!nDesbloqTiene) Then
'            sMensaje = ""
'            If CInt(rsCredito!nResultadoPS) > 0 Then
'                sMensaje = "Potencial Sobredendeudamiento"
'            End If
'
'            If CInt(rsCredito!nResultadoSE) > 0 Then
'                sMensaje = IIf(Len(sMensaje) > 0, " y ", "")
'                sMensaje = sMensaje & "Sobredendeudamiento"
'            End If
'
'            MsgBox "Crédito ya fue desbloqueado por  " & sMensaje & ".", vbInformation, "Aviso"
'            LimpiarDatos
'            Exit Sub
'        End If
        
        lblNomCliente.Caption = rsCredito!cPersNombre
        txtNroDOI.Caption = rsCredito!cPersIDnro
        lblCredito.Caption = sCta
        fnId = CLng(rsCredito!nId)
        
        Set rsCredito = oDCredito.SobreEndDatosCreditoADesbloqCodigos(sCta, gsCodCargo, True, True)
        LimpiaFlex feCodigos
    'JOEP 13092016 Inicio
        Dim nContador0 As Integer
        Dim nContador1 As Integer
        Dim nContador2 As Integer
        nContador0 = 0
        nContador1 = 0
        nContador2 = 0
    'JOEP 13092016 Fin
        If Not (rsCredito.EOF And rsCredito.BOF) Then
            For i = 1 To rsCredito.RecordCount
                feCodigos.AdicionaFila
                feCodigos.TextMatrix(i, 1) = rsCredito!cCodigo
                feCodigos.TextMatrix(i, 2) = rsCredito!cResultado
                feCodigos.TextMatrix(i, 3) = Mid(Trim(rsCredito!cDetalle), 1, 24) & "..."
                feCodigos.TextMatrix(i, 4) = rsCredito!cPlanmitigacion
                feCodigos.TextMatrix(i, 5) = CInt(rsCredito!nCodigo)
                feCodigos.TextMatrix(i, 6) = CInt(rsCredito!nResultado)
                feCodigos.TextMatrix(i, 7) = CInt(rsCredito!bGrabado)
                feCodigos.TextMatrix(i, 8) = Trim(rsCredito!cDetalle)
                txtEvalRisgSobre.Text = rsCredito!CodFinal
'            'JOEP 13092016 Inicio
'                If rsCredito!nResultado = 0 Then
'                    nContador0 = nContador0 + 1
'                ElseIf rsCredito!nResultado = 1 Then
'                    nContador1 = nContador1 + 1
'                ElseIf rsCredito!nResultado = 2 Then
'                    nContador2 = nContador2 + 1
'                End If
'             'JOEP 13092016 Fin
                rsCredito.MoveNext
            Next i
'          'JOEP 13092016 Inicio
'            If nContador0 = 5 Then
'               txtEvalRisgSobre.Text = "No Aplica"
'            ElseIf nContador2 >= 2 Then
'               txtEvalRisgSobre.Text = "Sobreendeudado"
'            Else
'               txtEvalRisgSobre.Text = "Potencialmente Sobreendeudado"
'            End If
'          'JOEP 13092016 Fin
        Else
            MsgBox "Crédito no cuenta con codigos a desbloquear.", vbInformation, "Aviso"
            LimpiarDatos
            'JOEP20190227 CP
                feCodigos.Enabled = False
            'JOEP20190227 CP
            Exit Sub
        End If
        cmdAutorizar.Enabled = True
     End If
               
     Set oDCredito = Nothing
     Set rsCredito = Nothing
Else
    LimpiarDatos
    Exit Sub
End If

End Sub

'JOEP CODIGOS
Private Sub feCodigos_Click()
Dim oCredito As COMDCredito.DCOMCredito
Dim rsDatosCodigos As ADODB.Recordset
            
Set oCredito = New COMDCredito.DCOMCredito

If feCodigos.TextMatrix(feCodigos.row, 6) <> 0 Then
    If feCodigos.Col = 4 Then
        Set rsDatosCodigos = oCredito.MostrarDatosPlanMitig(5)
            feCodigos.CargaCombo rsDatosCodigos
    End If
End If

Set oCredito = Nothing
RSClose rsDatosCodigos
End Sub

Private Sub feCodigos_DblClick()

If feCodigos.TextMatrix(feCodigos.row, 6) = 0 Then
    feCodigos.ListaControles = "X-X-X-X-X-X-X-X-X"
Else
    feCodigos.ListaControles = "X-X-X-1-3-X-X-X-X"
End If
End Sub

Private Sub feCodigos_EnterCell()
    If feCodigos.TextMatrix(feCodigos.row, 6) = 0 Then
        feCodigos.ListaControles = "X-X-X-X-X-X-X-X-X"
    Else
        feCodigos.ListaControles = "X-X-X-1-3-X-X-X-X"
    End If
End Sub
'JOEP CODIGOS

Private Sub feCodigos_OnClickTxtBuscar(psCodigo As String, psDescripcion As String)
psCodigo = feCodigos.TextMatrix(feCodigos.row, 3)
psDescripcion = feCodigos.TextMatrix(feCodigos.row, 8)

MsgBox psDescripcion, vbInformation, "Detalle del Código"
End Sub

Private Sub feCodigos_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim sColumnas() As String

sColumnas = Split(feCodigos.ColumnasAEditar, "-")

If sColumnas(pnCol) = "X" Then
    Cancel = False
    SendKeys "{Tab}", True
    Exit Sub
End If

If pnCol = 4 Then
    If CInt(feCodigos.TextMatrix(pnRow, 6)) = 0 Then
        Cancel = False
        MsgBox "No se permite ingresar Plan de Mitigación a codigos que no corresponde.", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
    
    If CInt(feCodigos.TextMatrix(pnRow, 6)) <> 0 Then
        If CInt(feCodigos.TextMatrix(pnRow, 7)) = 1 Then
            Cancel = False
            MsgBox "No se permite ingresar Plan de Mitigación a codigos que ya fueron desbloqueados.", vbInformation, "Aviso"
            SendKeys "{Tab}", True
            Exit Sub
        End If
    End If
End If

If pnCol = 3 Then
    feCodigos.TextMatrix(feCodigos.row, 3) = "..."
End If
End Sub

Private Function ValidaDatos() As Boolean
Dim i As Integer
ValidaDatos = True

For i = 1 To feCodigos.rows - 1
    If CInt(feCodigos.TextMatrix(i, 7)) = 0 And Trim(feCodigos.TextMatrix(i, 4)) = "" Then
        ValidaDatos = False
        MsgBox "Favor de ingresar el Plan de Mitigación del " & Trim(feCodigos.TextMatrix(i, 1)) & ".", vbInformation, "Aviso"
        feCodigos.SetFocus
        feCodigos.Col = 4
        feCodigos.row = i
        SendKeys "{Enter}", True
        Exit Function
    End If
Next i

End Function

Private Sub Form_Load()
'JOEP20190227 CP
feCodigos.Enabled = False
'JOEP20190227 CP
End Sub
