VERSION 5.00
Begin VB.Form frmCredFormEvalDescuento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descuentos"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5745
   Icon            =   "frmCredFormEvalDescuento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
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
      Left            =   1080
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtTotalDescuento 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptarDescuento 
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
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdQuitarDescuento 
      Caption         =   "Quitar"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdAgregarDescuento 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin SICMACT.FlexEdit feDescuento 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   3836
      Cols0           =   4
      HighLight       =   1
      AllowUserResizing=   1
      EncabezadosNombres=   "N-Descripcion-Monto-Aux"
      EncabezadosAnchos=   "400-3500-1400-0"
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-2-X"
      ListaControles  =   "0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "R-L-R-C"
      FormatosEdit    =   "3-0-2-0"
      TextArray0      =   "N"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   6
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
   End
   Begin VB.Label Label1 
      Caption         =   "Total :"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   2520
      Width           =   495
   End
End
Attribute VB_Name = "frmCredFormEvalDescuento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre      : frmCredFormEvalDescuento                        '
'** Descripción : Formulario para evaluación de Creditos            '
'** Referencia  : ERS004-2016                                       '
'** Creación    : JOEP, 20160525 09:00:00 AM                        '
'*******************************************************************'
Option Explicit

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&

Dim Suma As Double
Dim Total As Double
Dim vnTotal As Double
Dim MtrDescuento1 As Variant
Dim MtrDescuento2 As Variant
Dim nNum As Integer
Dim i As Integer
Dim sCtaCod As String

Dim nValorOriginal1 As Currency
Dim nValorOriginal2 As Currency

Public Function DisableCloseButton(frm As Form) As Boolean
'PURPOSE: Removes X button from a form
'EXAMPLE: DisableCloseButton Me
'RETURNS: True if successful, false otherwise
'NOTES:   Also removes Exit Item from
'         Control Box Menu
    Dim lHndSysMenu As Long
    Dim lAns1 As Long, lAns2 As Long
    
    lHndSysMenu = GetSystemMenu(frm.hwnd, 0)

    'remove close button
    lAns1 = RemoveMenu(lHndSysMenu, 6, MF_BYPOSITION)

   'Remove seperator bar
    lAns2 = RemoveMenu(lHndSysMenu, 5, MF_BYPOSITION)
    
    'Return True if both calls were successful
    DisableCloseButton = (lAns1 <> 0 And lAns2 <> 0)

End Function

Private Sub cmdCancelar_Click()
If nNum = 1 Then
    vnTotal = nValorOriginal1
Else
    vnTotal = nValorOriginal2
End If
    Unload Me
End Sub

Private Sub feDescuento_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim Editar() As String
    
    Editar = Split(Me.feDescuento.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        SendKeys "{Tab}"
        Exit Sub
    End If
End Sub

Private Sub feDescuento_RowColChange()
    If feDescuento.Col = 2 Then
            feDescuento.AvanceCeldas = Vertical
        Else
            feDescuento.AvanceCeldas = Horizontal
    End If
End Sub

Private Sub Form_Load()
   DisableCloseButton Me
   CentraForm Me
End Sub

Private Function Validar() As Boolean

Validar = True

    If feDescuento.TextMatrix(1, 2) = "0.00" Then
        MsgBox "Ud. debe Ingresar Descuentos ", vbInformation, "Aviso"
        Validar = False
        Exit Function
    End If
    
    If CCur(txtTotalDescuento.Text) < 0 Then
        MsgBox "Total no puede ser NEGATIVO, Verifique los Datos", vbInformation, "Aviso"
        Validar = False
        Exit Function
    End If
    
End Function

Private Sub cmdAceptarDescuento_Click()

If Validar Then

    If nNum = 1 Then
          ReDim MtrDescuento1(2, 0)
            For i = 1 To feDescuento.Rows - 1
               ReDim Preserve MtrDescuento1(2, i)
               MtrDescuento1(0, i) = feDescuento.TextMatrix(i, 0)
                MtrDescuento1(1, i) = feDescuento.TextMatrix(i, 1)
                MtrDescuento1(2, i) = CDbl(feDescuento.TextMatrix(i, 2))
            Next i
    Else
        ReDim MtrDescuento2(2, 0)
            For i = 1 To feDescuento.Rows - 1
                ReDim Preserve MtrDescuento2(2, i)
                MtrDescuento2(0, i) = feDescuento.TextMatrix(i, 0)
                MtrDescuento2(1, i) = feDescuento.TextMatrix(i, 1)
                MtrDescuento2(2, i) = CDbl(feDescuento.TextMatrix(i, 2))
            Next i
    End If
    Call Calculo
    Unload Me
    
End If

End Sub

Private Sub cmdAgregarDescuento_Click()
    If feDescuento.Rows - 1 < 25 Then
            feDescuento.lbEditarFlex = True
            feDescuento.AdicionaFila
            feDescuento.SetFocus
            SendKeys "{Enter}"
        Else
            MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdQuitarDescuento_Click()
    If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        feDescuento.EliminaFila (feDescuento.row)
        Call Calculo
    End If
End Sub

Private Sub feDescuento_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 2 Then
        If IsNumeric(feDescuento.TextMatrix(pnRow, pnCol)) Then
                If feDescuento.TextMatrix(pnRow, pnCol) >= 0 Then
                        feDescuento.TextMatrix(pnRow, pnCol) = UCase(feDescuento.TextMatrix(pnRow, pnCol))
                        Call Calculo
                    Else
                        feDescuento.TextMatrix(pnRow, pnCol) = "0.00"
                End If
            Else
                feDescuento.TextMatrix(pnRow, pnCol) = "0.00"
        End If
    End If
End Sub

Public Sub Calculo()
    Dim nIndice As Integer
    Suma = 0

    For nIndice = 1 To 1
        Suma = Suma + feDescuento.TextMatrix(nIndice, 2)
    Next nIndice
    
    'Suma = feDescuento.TextMatrix(2, 2)
    vnTotal = Suma
     
    txtTotalDescuento.Text = 0
    txtTotalDescuento.Text = Format(Round(feDescuento.TextMatrix(2, 2) - feDescuento.TextMatrix(1, 2) - feDescuento.TextMatrix(3, 2) - feDescuento.TextMatrix(4, 2) - feDescuento.TextMatrix(5, 2) - feDescuento.TextMatrix(6, 2), 2), "#,##0.00")
         
End Sub

Public Sub Inicio1(ByRef psTotal As Double, ByRef pMtrDescuento1 As Variant, ByVal pN As Integer, ByVal psCtaCod As String)
    nNum = pN
    sCtaCod = psCtaCod
    
    If IsArray(pMtrDescuento1) Then
            If UBound(pMtrDescuento1) > 0 Then
                    MtrDescuento1 = pMtrDescuento1
                    Call CargarGridConArray1
                Else
                    Call CargaDescuentos
            End If
        Else
            vnTotal = 0
            nValorOriginal1 = vnTotal
            Call CargaDescuentos
    End If

    Me.Show 1
       
    If IsArray(MtrDescuento1) Then
       pMtrDescuento1 = MtrDescuento1
       psTotal = vnTotal
    End If
    
End Sub

Public Sub Inicio2(ByRef psTotal As Double, ByRef pMtrDescuento2 As Variant, ByVal pN As Integer, ByVal psCtaCod As String)
    nNum = pN
    sCtaCod = psCtaCod
    
    If IsArray(pMtrDescuento2) Then
       
            If UBound(pMtrDescuento2) > 0 Then
                    MtrDescuento2 = pMtrDescuento2
                    Call CargarGridConArray2
                Else
                    Call CargaDescuentos
            End If
       
        Else
            vnTotal = 0
            nValorOriginal2 = vnTotal
            Call CargaDescuentos
        
    End If

    Me.Show 1
    
    If IsArray(MtrDescuento2) Then
       pMtrDescuento2 = MtrDescuento2
       psTotal = vnTotal
    End If
  
End Sub

Private Sub CargarGridConArray1()
    Dim i As Integer
    Dim nTotal As Double

    feDescuento.lbEditarFlex = True
    
    For i = 1 To UBound(MtrDescuento1, 2)
        feDescuento.AdicionaFila
        feDescuento.TextMatrix(i, 0) = MtrDescuento1(0, i)
        feDescuento.TextMatrix(i, 1) = MtrDescuento1(1, i)
        feDescuento.TextMatrix(i, 2) = Format(MtrDescuento1(2, i), "#,##0.00")
        'nTotal = nTotal + feDescuento.TextMatrix(i, 2)
    Next i
    
    txtTotalDescuento.Text = Format(Round(MtrDescuento1(2, 2) - MtrDescuento1(2, 1) - MtrDescuento1(2, 3) - MtrDescuento1(2, 4) - MtrDescuento1(2, 5) - MtrDescuento1(2, 6), 2), "#,##0.00")
    'txtTotalDescuento.Text = Format(nTotal, "#,##0.00")
    nValorOriginal1 = feDescuento.TextMatrix(1, 2)
End Sub

Private Sub CargarGridConArray2()
    Dim i As Integer
    Dim nTotal As Double

    feDescuento.lbEditarFlex = True
    
    For i = 1 To UBound(MtrDescuento2, 2)
        feDescuento.AdicionaFila
        feDescuento.TextMatrix(i, 0) = MtrDescuento2(0, i)
        feDescuento.TextMatrix(i, 1) = MtrDescuento2(1, i)
        feDescuento.TextMatrix(i, 2) = Format(MtrDescuento2(2, i), "#,##0.00")
        'nTotal = nTotal + feDescuento.TextMatrix(i, 2)
    Next i
    
    txtTotalDescuento.Text = Format(Round(MtrDescuento2(2, 2) - MtrDescuento2(2, 1) - MtrDescuento2(2, 3) - MtrDescuento2(2, 4) - MtrDescuento2(2, 5) - MtrDescuento2(2, 6), 2), "#,##0.00")
    'txtTotalDescuento.Text = Format(nTotal, "#,##0.00")
    nValorOriginal2 = feDescuento.TextMatrix(1, 2)
    
End Sub

Private Sub CargaDescuentos()
    Dim oNCred As New COMNCredito.NCOMFormatosEval
    Dim oRS As New ADODB.Recordset
    Dim i As Integer
    
    Set oRS = oNCred.CredFormEvalObtieneDescuentoConv(sCtaCod)
    feDescuento.Clear
    FormateaFlex feDescuento
    For i = 1 To oRS.RecordCount
        feDescuento.AdicionaFila
        feDescuento.TextMatrix(i, 1) = oRS!cConsDescripcion
        feDescuento.TextMatrix(i, 2) = Format(oRS!nMonto, "0.00")
        oRS.MoveNext
    Next
End Sub




