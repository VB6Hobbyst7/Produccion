VERSION 5.00
Begin VB.Form frmCredReprogPropuestaIFIS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuotas IFIs"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9660
   Icon            =   "frmCredReprogPropuestaIFIS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8640
      TabIndex        =   6
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      Top             =   3600
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin SICMACT.FlexEdit fe_ReprogCuotasIfis 
         Height          =   2535
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   4471
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Nombre Ifis-CodIfis-Moneda-nMonedad-Saldo Deuda-Monto de Cuota-fechaRcc-aux"
         EncabezadosAnchos=   "500-4600-0-900-0-1500-1500-0-0"
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
         ColumnasAEditar =   "X-1-X-3-X-X-6-X-X"
         ListaControles  =   "0-0-0-3-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "R-L-R-L-R-R-R-L-C"
         FormatosEdit    =   "3-1-3-3-3-2-2-0-0"
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.EditMoney EditMoneyTotal 
         Height          =   255
         Left            =   7800
         TabIndex        =   4
         Top             =   3120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Total"
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
         Left            =   7200
         TabIndex        =   3
         Top             =   3120
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmCredReprogPropuestaIFIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************************
'** Nombre : frmCredReprogPropuestaIFIS
'** Descripción : Formulario de Deudas de ifis - Modulo de Reprogramacion Propuesta
'** Creación : JOEP, 20200805
'*********************************************************************************************************

Option Explicit

Dim cCtaCod As String
Dim vMatDatosIfis As Variant
Dim Rcc As String

Public Sub Inicio(ByVal pcCtaCod As String, Optional ByRef pMatIFI As Variant)
Set vMatDatosIfis = Nothing
cCtaCod = pcCtaCod

If IsArray(pMatIFI) Then
    vMatDatosIfis = pMatIFI
    CargaDatosMatrix
Else
    CargaDatos
End If
    
frmCredReprogPropuestaIFIS.Show 1
    
If IsArray(vMatDatosIfis) Then
    If UBound(vMatDatosIfis) > 0 Then
        pMatIFI = vMatDatosIfis
    Else
        Set pMatIFI = Nothing
        Set vMatDatosIfis = Nothing
    End If
Else
    Set pMatIFI = Nothing
    Set vMatDatosIfis = Nothing
    Rcc = ""
End If

End Sub

Private Sub CargaDatos()
    Dim oDCred As New COMDCredito.DCOMCredito
    Dim rsLista As New ADODB.Recordset
    Dim i As Integer
    
    Set rsLista = oDCred.ReprogramacionObtieneDatosIfis(cCtaCod)
    LimpiaFlex fe_ReprogCuotasIfis
    If Not (rsLista.EOF And rsLista.BOF) Then
        For i = 1 To rsLista.RecordCount
            fe_ReprogCuotasIfis.AdicionaFila
            fe_ReprogCuotasIfis.TextMatrix(i, 1) = rsLista!Entidad
            fe_ReprogCuotasIfis.TextMatrix(i, 2) = rsLista!codigo
            fe_ReprogCuotasIfis.TextMatrix(i, 3) = rsLista!cMoneda
            fe_ReprogCuotasIfis.TextMatrix(i, 4) = rsLista!Moneda
            fe_ReprogCuotasIfis.TextMatrix(i, 5) = rsLista!Saldo
            fe_ReprogCuotasIfis.TextMatrix(i, 6) = rsLista!MontoCuota
            fe_ReprogCuotasIfis.TextMatrix(i, 7) = rsLista!Fec_Rep
            Rcc = rsLista!Fec_Rep
            rsLista.MoveNext
        Next
        
        EditMoneyTotal.Text = Format(SumarCampo(fe_ReprogCuotasIfis, 6), "#,##0.00")
    End If
    Set oDCred = Nothing
    RSClose rsLista
End Sub

Private Sub CargaDatosMatrix()
Dim i As Integer
    For i = 1 To UBound(vMatDatosIfis)
        fe_ReprogCuotasIfis.AdicionaFila
        fe_ReprogCuotasIfis.TextMatrix(i, 1) = vMatDatosIfis(i, 1)
        fe_ReprogCuotasIfis.TextMatrix(i, 2) = vMatDatosIfis(i, 2)
        fe_ReprogCuotasIfis.TextMatrix(i, 3) = (vMatDatosIfis(i, 3) & Space(50) & vMatDatosIfis(i, 4))
        'fe_ReprogCuotasIfis.TextMatrix(i, 4) = vMatDatosIfis(i, 4)
        fe_ReprogCuotasIfis.TextMatrix(i, 5) = vMatDatosIfis(i, 5)
        fe_ReprogCuotasIfis.TextMatrix(i, 6) = vMatDatosIfis(i, 6)
        fe_ReprogCuotasIfis.TextMatrix(i, 7) = vMatDatosIfis(i, 7)
    Next i
    EditMoneyTotal.Text = Format(SumarCampo(fe_ReprogCuotasIfis, 6), "#,##0.00")
End Sub

Private Sub CargaMatrix()
Dim nIndice As Integer
Dim i As Integer
    nIndice = IIf(fe_ReprogCuotasIfis.TextMatrix(1, 1) = "", 0, fe_ReprogCuotasIfis.rows - 1)
    ReDim vMatDatosIfis(nIndice, 7)
    If nIndice > 0 Then
        For i = 1 To fe_ReprogCuotasIfis.rows - 1
            vMatDatosIfis(i, 0) = fe_ReprogCuotasIfis.TextMatrix(i, 0)
            vMatDatosIfis(i, 1) = Trim(fe_ReprogCuotasIfis.TextMatrix(i, 1))
            vMatDatosIfis(i, 2) = IIf(fe_ReprogCuotasIfis.TextMatrix(i, 2) = "", 0, fe_ReprogCuotasIfis.TextMatrix(i, 2))
            vMatDatosIfis(i, 3) = Left(fe_ReprogCuotasIfis.TextMatrix(i, 3), 10)
            vMatDatosIfis(i, 4) = Right(fe_ReprogCuotasIfis.TextMatrix(i, 3), 1)
            vMatDatosIfis(i, 5) = fe_ReprogCuotasIfis.TextMatrix(i, 5)
            vMatDatosIfis(i, 6) = fe_ReprogCuotasIfis.TextMatrix(i, 6)
            vMatDatosIfis(i, 7) = Rcc
        Next i
    End If
End Sub

Private Sub CmdAceptar_Click()
Call CargaMatrix
   Unload Me
End Sub

Private Sub cmdAgregar_Click()
     If fe_ReprogCuotasIfis.rows - 1 < 25 Then
        fe_ReprogCuotasIfis.AdicionaFila
        fe_ReprogCuotasIfis.SetFocus
        SendKeys "{Enter}"
               
        fe_ReprogCuotasIfis.ColumnasAEditar = "X-1-X-3-4-5-6"
        
    Else
        MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdCancelar_Click()
    Rcc = ""
    Unload Me
End Sub

Private Sub cmdQuitar_Click()
     If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        fe_ReprogCuotasIfis.EliminaFila (fe_ReprogCuotasIfis.row)
        EditMoneyTotal.Text = Format(SumarCampo(fe_ReprogCuotasIfis, 5), "#,##0.00")
        fe_ReprogCuotasIfis.ColumnasAEditar = "X-X-X-X-X-X-6"
    End If
End Sub

Private Sub fe_ReprogCuotasIfis_Click()

    If fe_ReprogCuotasIfis.TextMatrix(fe_ReprogCuotasIfis.row, 2) = "109" Then
        fe_ReprogCuotasIfis.ColumnasAEditar = "X-X-X-X-X-X-X"
    Else
        fe_ReprogCuotasIfis.ColumnasAEditar = "X-X-X-X-X-X-6"
    End If
    EditMoneyTotal.Text = Format(SumarCampo(fe_ReprogCuotasIfis, 6), "#,##0.00")
End Sub

Private Sub fe_ReprogCuotasIfis_DblClick()
    If fe_ReprogCuotasIfis.TextMatrix(fe_ReprogCuotasIfis.row, fe_ReprogCuotasIfis.Col) = "" And fe_ReprogCuotasIfis.TextMatrix(fe_ReprogCuotasIfis.row, 2) <> "109" Then
        fe_ReprogCuotasIfis.ColumnasAEditar = "X-1-X-3-4-5-6"
    Else
        If fe_ReprogCuotasIfis.TextMatrix(fe_ReprogCuotasIfis.row, 2) = "109" Then
            fe_ReprogCuotasIfis.ColumnasAEditar = "X-X-X-X-X-X-X"
        Else
            fe_ReprogCuotasIfis.ColumnasAEditar = "X-X-X-X-X-X-6"
        End If
    End If
    EditMoneyTotal.Text = Format(SumarCampo(fe_ReprogCuotasIfis, 6), "#,##0.00")
End Sub

Private Sub fe_ReprogCuotasIfis_GotFocus()
    If fe_ReprogCuotasIfis.TextMatrix(fe_ReprogCuotasIfis.row, 2) = "109" Then
        fe_ReprogCuotasIfis.ColumnasAEditar = "X-X-X-X-X-X-X"
    Else
        fe_ReprogCuotasIfis.ColumnasAEditar = "X-X-X-X-X-X-6"
    End If
    EditMoneyTotal.Text = Format(SumarCampo(fe_ReprogCuotasIfis, 6), "#,##0.00")
End Sub

Private Sub fe_ReprogCuotasIfis_LostFocus()
    If fe_ReprogCuotasIfis.TextMatrix(fe_ReprogCuotasIfis.row, 2) = "109" Then
        fe_ReprogCuotasIfis.ColumnasAEditar = "X-X-X-X-X-X-X"
    Else
        fe_ReprogCuotasIfis.ColumnasAEditar = "X-X-X-X-X-X-6"
    End If
    EditMoneyTotal.Text = Format(SumarCampo(fe_ReprogCuotasIfis, 6), "#,##0.00")
End Sub

Private Sub fe_ReprogCuotasIfis_OnCellChange(pnRow As Long, pnCol As Long)
    If fe_ReprogCuotasIfis.TextMatrix(pnRow, pnCol) = "" And fe_ReprogCuotasIfis.TextMatrix(pnRow, 2) = "109" Then
        fe_ReprogCuotasIfis.ColumnasAEditar = "X-1-X-3-4-5-6"
    Else
        If fe_ReprogCuotasIfis.TextMatrix(pnRow, 2) = "109" Then
            fe_ReprogCuotasIfis.ColumnasAEditar = "X-X-X-X-X-X-X"
        Else
            fe_ReprogCuotasIfis.ColumnasAEditar = "X-X-X-X-X-X-6"
        End If
    End If
    EditMoneyTotal.Text = Format(SumarCampo(fe_ReprogCuotasIfis, 6), "#,##0.00")
End Sub

Private Sub Form_Load()
Dim rs As ADODB.Recordset
Dim oDCOMCred As COMDConstantes.DCOMConstantes
Set oDCOMCred = New COMDConstantes.DCOMConstantes
Set rs = oDCOMCred.RecuperaConstantes(1011)

fe_ReprogCuotasIfis.ColumnasAEditar = "X-X-X-X-X-X-6"

fe_ReprogCuotasIfis.CargaCombo rs


Set oDCOMCred = Nothing
RSClose rs
End Sub
