VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAsientoModDatosNoContable 
   Caption         =   "Asientos: Modificación de Datos no Contables"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   Icon            =   "frmAsientoModDatosNoContable.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkDoc 
      Caption         =   "&Documento"
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
      Height          =   210
      Left            =   180
      TabIndex        =   24
      Top             =   750
      Width           =   1365
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   1
      EndProperty
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
      ForeColor       =   &H80000012&
      Height          =   285
      Left            =   7875
      TabIndex        =   16
      Top             =   4545
      Width           =   1185
   End
   Begin VB.TextBox txtSTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   1
      EndProperty
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
      ForeColor       =   &H80000012&
      Height          =   285
      Left            =   7875
      TabIndex        =   15
      Top             =   4215
      Width           =   1185
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   6855
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   13
      Top             =   4980
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   7980
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   12
      Top             =   4980
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Operación"
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
      Height          =   645
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   6855
      Begin VB.ComboBox cboOperacion 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   210
         Width           =   4215
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   5520
         TabIndex        =   9
         Top             =   210
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label8 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   150
         TabIndex        =   11
         Top             =   270
         Width           =   675
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha "
         Height          =   240
         Left            =   4950
         TabIndex        =   10
         Top             =   285
         Width           =   555
      End
   End
   Begin VB.Frame fraDato 
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
      Height          =   1095
      Left            =   60
      TabIndex        =   1
      Top             =   1560
      Width           =   9075
      Begin VB.TextBox txtMovDesc 
         Height          =   345
         Left            =   870
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   630
         Width           =   7935
      End
      Begin VB.TextBox txtPersNombre 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2820
         TabIndex        =   2
         Top             =   240
         Width           =   5985
      End
      Begin Sicmact.TxtBuscar txtPersCod 
         Height          =   330
         Left            =   870
         TabIndex        =   3
         Top             =   240
         Width           =   1935
         _extentx        =   3413
         _extenty        =   582
         appearance      =   1
         appearance      =   1
         font            =   "frmAsientoModDatosNoContable.frx":030A
         appearance      =   1
         tipobusqueda    =   3
      End
      Begin VB.Label Label4 
         Caption         =   "Glosa"
         Height          =   255
         Left            =   210
         TabIndex        =   6
         Top             =   660
         Width           =   645
      End
      Begin VB.Label Label5 
         Caption         =   "Persona"
         Height          =   255
         Left            =   210
         TabIndex        =   5
         Top             =   270
         Width           =   645
      End
   End
   Begin Sicmact.FlexEdit fgDoc 
      Height          =   735
      Left            =   60
      TabIndex        =   0
      Top             =   870
      Width           =   9075
      _extentx        =   16007
      _extenty        =   1296
      cols0           =   7
      highlight       =   1
      allowuserresizing=   3
      rowsizingmode   =   1
      encabezadosnombres=   "#-Tipo-Descripción-Numero-Fecha-nMovNro-cMovNro"
      encabezadosanchos=   "400-1200-3530-2100-1500-0-0"
      font            =   "frmAsientoModDatosNoContable.frx":0336
      font            =   "frmAsientoModDatosNoContable.frx":0362
      font            =   "frmAsientoModDatosNoContable.frx":038E
      font            =   "frmAsientoModDatosNoContable.frx":03BA
      font            =   "frmAsientoModDatosNoContable.frx":03E6
      fontfixed       =   "frmAsientoModDatosNoContable.frx":0412
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      columnasaeditar =   "X-1-X-3-4-X-X"
      textstylefixed  =   3
      listacontroles  =   "0-1-0-0-2-0-0"
      encabezadosalineacion=   "C-R-L-L-C-L-L"
      formatosedit    =   "0-3-0-0-0-0-0"
      textarray0      =   "#"
      colwidth0       =   405
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
   Begin Sicmact.FlexEdit fgImp 
      Height          =   1230
      Left            =   1125
      TabIndex        =   14
      Top             =   4155
      Width           =   3465
      _extentx        =   6165
      _extenty        =   1826
      cols0           =   11
      highlight       =   2
      allowuserresizing=   3
      encabezadosnombres=   "-#1-Ok-Impuesto-Tasa-Monto-CtaCont-CtaContDesc-cDocImpDH-cImpDestino-cDocImpOpc"
      encabezadosanchos=   "0-0-350-1000-600-1200-0-0-0-0-0"
      font            =   "frmAsientoModDatosNoContable.frx":0440
      font            =   "frmAsientoModDatosNoContable.frx":046C
      font            =   "frmAsientoModDatosNoContable.frx":0498
      font            =   "frmAsientoModDatosNoContable.frx":04C4
      font            =   "frmAsientoModDatosNoContable.frx":04F0
      fontfixed       =   "frmAsientoModDatosNoContable.frx":051C
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      columnasaeditar =   "X-X-2-X-X-X-X-X-X-X-X"
      textstylefixed  =   4
      listacontroles  =   "0-0-4-0-0-0-0-0-0-0-0"
      encabezadosalineacion=   "C-C-L-C-R-R-C-C-C-C-L"
      formatosedit    =   "0-0-0-2-2-2-2-2-2-2-0"
      lbeditarflex    =   -1
      lbformatocol    =   -1
      lbbuscaduplicadotext=   -1
      rowheight0      =   285
   End
   Begin Sicmact.FlexEdit fgDetalle 
      Height          =   1470
      Left            =   60
      TabIndex        =   17
      Top             =   2670
      Width           =   9060
      _extentx        =   17806
      _extenty        =   2593
      cols0           =   7
      highlight       =   2
      allowuserresizing=   1
      encabezadosnombres=   "#-Código-Descripción-Monto-DH-Grav.-ItemCtaCont"
      encabezadosanchos=   "350-1200-3500-1200-0-0-0"
      font            =   "frmAsientoModDatosNoContable.frx":054A
      font            =   "frmAsientoModDatosNoContable.frx":0576
      font            =   "frmAsientoModDatosNoContable.frx":05A2
      font            =   "frmAsientoModDatosNoContable.frx":05CE
      font            =   "frmAsientoModDatosNoContable.frx":05FA
      fontfixed       =   "frmAsientoModDatosNoContable.frx":0626
      lbultimainstancia=   -1
      columnasaeditar =   "X-X-X-X-X-5-X"
      textstylefixed  =   3
      listacontroles  =   "0-0-0-0-0-4-0"
      encabezadosalineacion=   "C-L-L-R-C-C-C"
      formatosedit    =   "0-0-0-2-2-0-0"
      textarray0      =   "#"
      lbeditarflex    =   -1
      lbformatocol    =   -1
      lbpuntero       =   -1
      lbbuscaduplicadotext=   -1
      colwidth0       =   345
      rowheight0      =   300
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Impuestos"
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
      Height          =   225
      Left            =   135
      TabIndex        =   22
      Top             =   4320
      Width           =   885
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "y/o"
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
      Height          =   195
      Left            =   450
      TabIndex        =   21
      Top             =   4635
      Width           =   315
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Retenc."
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
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Top             =   4935
      Width           =   705
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Height          =   210
      Left            =   6825
      TabIndex        =   19
      Top             =   4575
      Width           =   915
   End
   Begin VB.Label lblSTot 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "SubTotal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6825
      TabIndex        =   18
      Top             =   4260
      Width           =   885
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Height          =   1230
      Left            =   60
      TabIndex        =   23
      Top             =   4140
      Width           =   1035
   End
   Begin VB.Shape ShapeIGV 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   345
      Left            =   6705
      Top             =   4185
      Width           =   2385
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   345
      Left            =   6705
      Top             =   4515
      Width           =   2385
   End
End
Attribute VB_Name = "frmAsientoModDatosNoContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCta As DCtaCont
Dim rs   As ADODB.Recordset

Dim sMovNroAnt As String
Dim nMovNroAnt As Long
Dim lbTieneIGV As Boolean
Dim nTasaIGV   As Currency

Dim lsPersonaAnt As String
Dim lsMovDescAnt As String
Dim lbEditar As Boolean

'ARLO20170208****
Dim objPista As COMManejador.Pista
Dim lsPalabra, lsAccion As String
'************

Public Sub Inicio(psMov As String, pnMov As Long, pbEditar As Boolean)
sMovNroAnt = psMov
nMovNroAnt = pnMov
lbEditar = pbEditar
Me.Show 1
End Sub

Private Sub chkDoc_Click()
If chkDoc.value = vbChecked Then
    fgDoc.lbEditarFlex = True
    If fgDoc.TextMatrix(1, 1) = "" Then
        fgDoc.AdicionaFila
    End If
Else
    fgDoc.lbEditarFlex = False
End If
End Sub

Private Sub cmdAceptar_Click()
Dim oMov As DMov
    Set oMov = New DMov
    Dim lsMovNroMod As String
    Dim lnI As Integer
    
    If MsgBox("Desea Grabar los Cambios ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    lsMovNroMod = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    oMov.BeginTrans
        If lsMovDescAnt <> Me.txtMovDesc.Text Or lsPersonaAnt <> txtPersCod.psCodigoPersona Then
            oMov.InsertaMovModificaNoCont nMovNroAnt, lsMovNroMod, lsPersonaAnt, lsMovDescAnt
            
            If lsPersonaAnt <> txtPersCod.Text Then
                oMov.ActualizaMovGasto nMovNroAnt, txtPersCod.psCodigoPersona
            End If
            
            If lsMovDescAnt <> Me.txtMovDesc.Text Then
                oMov.ActualizaMovDescripcion nMovNroAnt, Me.txtMovDesc.Text
            End If
        End If
        
        oMov.InsertaMovModificaNoContDoc nMovNroAnt, lsMovNroMod
        
        For lnI = 1 To Me.fgDoc.Rows - 1
            If Me.fgDoc.TextMatrix(lnI, 1) <> "" Then
                oMov.InsertaMovDoc nMovNroAnt, Me.fgDoc.TextMatrix(lnI, 1), Me.fgDoc.TextMatrix(lnI, 3), Format(CDate(Me.fgDoc.TextMatrix(lnI, 4)), gsFormatoFecha)
            End If
        Next lnI
        
        For lnI = 1 To Me.fgDetalle.Rows - 1
            If Me.fgDetalle.TextMatrix(lnI, 1) <> "" Then
                oMov.UpdateDocNoCont nMovNroAnt, Me.fgDetalle.TextMatrix(lnI, 4), Me.fgDetalle.TextMatrix(lnI, 1), fgDetalle.TextMatrix(lnI, 3), fgDetalle.TextMatrix(lnI, 5)
            End If
        Next lnI
        
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            If (lsMovDescAnt = "") Then
            lsPalabra = "Agrego"
            lsAccion = "1"
            Else: lsPalabra = "Modifico"
            lsAccion = "2"
            End If
            gsOpeCod = LogPistaRegistoAsientCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, lsAccion, "Se " & lsPalabra & " los Datos No Contables del Tipo de Operacion " & cboOperacion.Text
            Set objPista = Nothing
            '*******
        
    oMov.CommitTrans
    Unload Me
    glAceptar = True


End Sub

Private Sub cmdCancelar_Click()
Unload Me
glAceptar = False

End Sub

Private Sub Form_Load()
Dim i As Integer
CentraForm Me
Set oCta = New DCtaCont
If sMovNroAnt <> "" Then
   cboOperacion.Enabled = False
   txtFecha.Enabled = False
   
   'txtPersCod.TipoBusPers = BusPersDocumentoRuc
   txtPersCod.TipoBusPers = BusPersCodigo
   txtPersCod.TipoBusqueda = BuscaPersona
   
   txtPersCod.EditFlex = False
  
   Dim oMov As New DMov
   fgDoc.rsFlex = oMov.CargaMovDocAsiento(nMovNroAnt)
   
   Set rs = oMov.CargaMovOpeAsiento(nMovNroAnt)
   If Not rs.EOF Then
      txtFecha = GetFechaMov(sMovNroAnt, True)
      cboOperacion.AddItem rs!cOpeCod & "  " & rs!cOpeDesc
      cboOperacion.ListIndex = cboOperacion.ListCount - 1
      txtPersCod.Text = Trim(rs!cRuc)
      txtPersCod.psCodigoPersona = Trim(rs!cPersCod)
      lsPersonaAnt = Trim(rs!cPersCod)
      txtPersNombre = PstaNombre(rs!cPersNombre, False)
      txtMovDesc = rs!cMovDesc
      lsMovDescAnt = rs!cMovDesc
      Me.chkDoc.value = 1
   End If
   
   fgDetalle.rsFlex = oMov.CargaMovDocNoContable(nMovNroAnt)
        txtSTotal = "0"
        For i = 1 To fgDetalle.Rows - 1
            If IsNumeric(fgDetalle.TextMatrix(i, 3)) Then
                Me.txtSTotal = CDbl(Me.txtSTotal) + CDbl(fgDetalle.TextMatrix(i, 3))
            End If
        Next i
        Me.txtTotal = CDbl(Me.txtSTotal)
   End If
    'fgDet.ListaControles = "0-0-0-0-0-4-0"
    Dim oDoc As New DDocumento
    fgDoc.AutoAdd = True
    fgDoc.AvanceCeldas = Horizontal
    fgDoc.psRaiz = "Documentos"
    fgDoc.rsTextBuscar = oDoc.CargaDocumento(, , , 1)
    Set oDoc = Nothing
    
    If lbEditar Then
        cmdAceptar.Visible = True
    Else
        cmdAceptar.Visible = False
    End If
    
    
End Sub

Private Sub MuestraDatos_Doc()
Dim rs As ADODB.Recordset
Dim oDoc As DDocumento
Dim lsTpoDoc As String
lsTpoDoc = fgDoc.TextMatrix(fgDoc.row, 1)

If lsTpoDoc Then
    Exit Sub
End If
Set oDoc = New DDocumento

Dim nRow As Integer
   lbTieneIGV = False
   fgDetalle.Cols = 7
   fgImp.Clear
   fgImp.FormaCabecera
   fgImp.Rows = 2
   fgDetalle.ColumnasAEditar = Left(fgDetalle.ColumnasAEditar, 13)
   fgDetalle.EncabezadosAlineacion = Left(fgDetalle.EncabezadosAlineacion, 13)
   fgDetalle.FormatosEdit = Left(fgDetalle.FormatosEdit, 13)
   fgDetalle.ListaControles = Left(fgDetalle.ListaControles, 13)
   Set rs = New ADODB.Recordset
   Set rs = oDoc.CargaDocImpuesto(Mid(lsTpoDoc, 1, 2))
   Do While Not rs.EOF
      'Primero adicionamos Columna de Impuesto
        fgDetalle.Cols = fgDetalle.Cols + 1
        fgDetalle.ColWidth(fgDetalle.Cols - 1) = 1200
        fgDetalle.ColumnasAEditar = fgDetalle.ColumnasAEditar & "-" & fgDetalle.Cols - 1
        fgDetalle.EncabezadosAlineacion = fgDetalle.EncabezadosAlineacion & "-R"
        fgDetalle.FormatosEdit = fgDetalle.FormatosEdit & "-2"
        fgDetalle.ListaControles = fgDetalle.ListaControles & "-0"
        fgDetalle.TextMatrix(0, fgDetalle.Cols - 1) = rs!cImpAbrev

       'Adicionamos los impuestos en el grid de impuestos
        fgImp.AdicionaFila
        fgImp.col = 0
        nRow = fgImp.row
        fgImp.TextMatrix(nRow, 1) = "."
        If rs!cDocImpOpc = "1" Then
            'activamos el check de impuesto enviandole el valor "1"
'            If cboDocDestino.ListIndex <> 3 Then fgImp.TextMatrix(nRow, 2) = "1"
        End If
        fgImp.TextMatrix(nRow, 3) = rs!cImpAbrev
        fgImp.TextMatrix(nRow, 4) = Format(rs!nImpTasa, gsFormatoNumeroView)
        fgImp.TextMatrix(nRow, 5) = Format(0, gsFormatoNumeroView)
        fgImp.TextMatrix(nRow, 6) = rs!cCtaContCod
        fgImp.TextMatrix(nRow, 7) = rs!cCtaContDesc
        fgImp.TextMatrix(nRow, 8) = rs!cDocImpDH
        fgImp.TextMatrix(nRow, 9) = rs!cImpDestino
        fgImp.TextMatrix(nRow, 10) = rs!cDocImpOpc
        If rs!cCtaContCod = gcCtaIGV Then
            lbTieneIGV = True
            nTasaIGV = rs!nImpTasa
        End If
        rs.MoveNext
   Loop
   fgImp.col = 1
'   If lbTieneIGV = False Then
'      cboDocDestino.ListIndex = -1
'      cboDocDestino.Enabled = False
'   Else
'      cboDocDestino.Enabled = True
'      cboDocDestino.ListIndex = 2
'   End If
   CalculaTotal

Set oDoc = Nothing
End Sub

Private Sub CalculaTotal(Optional lCalcImpuestos As Boolean = True)
Dim n As Integer, m As Integer
Dim nSTot As Currency
Dim nITot As Currency, nImp As Currency
Dim nTot  As Currency
nSTot = 0: nTot = 0
If fgImp.TextMatrix(1, 1) = "" Then
   lCalcImpuestos = False
End If
For m = 1 To fgImp.Rows - 1
   nITot = 0
   For n = 1 To fgDetalle.Rows - 1
      If fgImp.TextMatrix(m, 2) = "." Then
         If lCalcImpuestos Then
            nImp = Round(Val(Format(fgDetalle.TextMatrix(n, 3), gsFormatoNumeroDato)) * Val(Format(fgImp.TextMatrix(m, 4), gsFormatoNumeroDato)) / 100, 2)
            fgDetalle.TextMatrix(n, m + 6) = Format(nImp, gsFormatoNumeroView)
         Else
            nImp = fgDetalle.TextMatrix(n, m + 6)
         End If
         nITot = nITot + nImp
      Else
         If lCalcImpuestos Then fgDetalle.TextMatrix(n, m + 6) = "0.00"
      End If
   Next
   fgImp.TextMatrix(m, 5) = Format(nITot, gsFormatoNumeroView)
   nTot = nTot + nITot * IIf(fgImp.TextMatrix(m, 8) = "D", 1, -1)
Next
For n = 1 To fgDetalle.Rows - 1
   nSTot = nSTot + Val(Format(fgDetalle.TextMatrix(n, 3), gsFormatoNumeroDato))
Next
txtSTotal = Format(nSTot, gsFormatoNumeroView)
txtTotal = Format(nSTot + nTot, gsFormatoNumeroView)
End Sub

Private Sub fgDetalle_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim i As Integer
txtSTotal = "0"
For i = 1 To fgDetalle.Rows - 1
    Me.txtSTotal = CDbl(Me.txtSTotal) + CDbl(fgDetalle.TextMatrix(i, 3))
Next i
Me.txtSTotal = Format(Me.txtSTotal, "#,###0.00")
Me.txtTotal = Format(CDbl(Me.txtSTotal), "#,###0.00")
End Sub

Private Sub Form_Initialize()
    lsPersonaAnt = ""
    lsMovDescAnt = ""
End Sub

Private Sub txtPersCod_EmiteDatos()
    If txtPersCod.Text <> "" Then
        If txtPersCod.Text = "00000000000" Then
            MsgBox "Persona sin RUC", vbInformation, "Aviso"
            'txtPersCod.Text = ""
            'txtPersNombre.Text = ""
         Else
            txtPersCod.Text = txtPersCod.psCodigoPersona
            txtPersNombre.Text = txtPersCod.psDescripcion
         End If
    End If
End Sub
