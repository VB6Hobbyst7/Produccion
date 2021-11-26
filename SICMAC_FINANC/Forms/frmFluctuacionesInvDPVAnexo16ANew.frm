VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFluctuacionesInvDPVAnexo16ANew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fluctuaciones Valores"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   13470
   LinkTopic       =   "Fluctuaciones Valores"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   13470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtVALIBMNTotal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   320
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "0.00"
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox txtVALIBMETotal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   320
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "0.00"
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdBuscarFluctuacion 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtSolesPatr 
      Height          =   325
      Left            =   9000
      TabIndex        =   10
      Text            =   "0.00"
      Top             =   530
      Width           =   1575
   End
   Begin MSMask.MaskEdBox txtFechaPatrimonio 
      Height          =   330
      Left            =   6600
      TabIndex        =   9
      Top             =   530
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   582
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Frame fraPatr 
      Caption         =   "Patrimonio Efectivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   4680
      TabIndex        =   5
      Top             =   240
      Width           =   8655
      Begin VB.TextBox txtDolaresPatr 
         Height          =   352
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   270
         Width           =   1575
      End
      Begin VB.Label lblSoles 
         Caption         =   "Soles:"
         Height          =   255
         Left            =   3720
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lbDolares 
         Caption         =   "Dólares:"
         Height          =   255
         Left            =   6120
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblFechaFluct 
         Caption         =   "Pat. Efectivo (mes ant):"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame fraFechaFluctAnex16A 
      Caption         =   "Fecha de Fluctuación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   705
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2490
      Begin MSMask.MaskEdBox txtFechaFluctuacion 
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Width           =   1395
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   11935
      TabIndex        =   1
      Top             =   5160
      Width           =   1395
   End
   Begin Sicmact.FlexEdit flxFluctuacion 
      Height          =   3135
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   5530
      Cols0           =   11
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Cta Cont-Tipo Instr-Emisor-Cod-Fecha Vencimiento-Moneda-Saldo-Fluctuación-Valor en Libros Total-cCodPers"
      EncabezadosAnchos=   "500-1000-875-2500-0-1500-1300-1500-1500-1720-0"
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
      ColumnasAEditar =   "X-X-X-3-X-X-X-7-8-X-X"
      ListaControles  =   "0-0-0-1-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-C-L-L-C-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-2-2-2-0"
      CantEntero      =   20
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame fraFluctuacion 
      Caption         =   "Inversión en Instrumentos Representativos de Deuda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   3735
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   13215
   End
   Begin VB.Label Label7 
      Caption         =   "VALOR LIBRO MN :"
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
      Left            =   3000
      TabIndex        =   17
      Top             =   5220
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "VALOR LIBRO ME: "
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
      Left            =   6720
      TabIndex        =   16
      Top             =   5220
      Width           =   1695
   End
End
Attribute VB_Name = "frmFluctuacionesInvDPVAnexo16ANew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
'*** Nombre : frmFluctuacionesInvDPVAnexo16ANew
'*** Descripción : Formulario para registrar las fluctuaciones para el Anexo 16A - Cuadro de Liquidez por plazos de Vencimiento de corto Plazo. .
'*** Creación : NAGL el 20170424
'********************************************************************************
Dim oCambio As New nTipoCambio
Dim pdFechaFluct As Date
Dim pdFecPatr As Date
Dim ValorCelda As String
Dim ix As Integer
Dim oValor As New DAnexoRiesgos
Dim ValorCeldaIng As String

Public Sub Inicio(psOpeCod As String, pdFecha As Date)
     txtFechaFluctuacion.Text = Format(pdFecha, "dd/MM/YYYY")
     txtFechaPatrimonio.Text = DateAdd("d", -Day(pdFecha), pdFecha)
     If ValRegFluctuacionValores(pdFecha, "Load1") Then
       Call CargarFluctuacionesValores(pdFecha, "1")
     Else
       Call CargarFluctuacionesValores(pdFecha, "0")
     End If 'NAGL 20171221
     SumaValorLibro 'NAGL 20171221
     CentraForm Me
     Me.Show 1
End Sub

Private Sub CargarFluctuacionesValores(pdFecha As Date, psOpt As String) 'NAGL 20171221 Agregó psOpt
    Dim rs As New ADODB.Recordset
    Dim X As Integer
    
        Set rsValNew = oValor.CargaPatrimonioEfectivo(pdFecha)
            If Not (rsValNew.BOF And rsValNew.EOF) Then
                txtFechaPatrimonio = Format(rsValNew!dFechaPatrimonio, "dd/mm/yyyy") 'NAGL 20171221
                txtSolesPatr = Format(rsValNew!nPatrimonioMN, "###,##0.00")
                txtDolaresPatr = Format(rsValNew!nPatrimonioME, "###,##0.00")
            Else
                txtFechaPatrimonio = Format(DateAdd("D", -Day(pdFecha), pdFecha), "dd/mm/yyyy")
                txtSolesPatr = Format(0, "###,##0.00")
                txtDolaresPatr = Format(0, "###,##0.00")
            End If
                   
        Set rs = oValor.DevuelveFluctuacionesValoresDet(pdFecha, psOpt) 'NAGL 20171221 Agregó psOpt
        flxFluctuacion.Clear
        FormateaFlex flxFluctuacion
    
        If Not (rs.EOF And rs.BOF) Then
            For X = 1 To rs.RecordCount
                flxFluctuacion.AdicionaFila
                flxFluctuacion.TextMatrix(X, 1) = rs!cCtaCod
                flxFluctuacion.TextMatrix(X, 2) = oValor.CargaAsigTipoInstrumento(CStr(flxFluctuacion.TextMatrix(X, 1))) 'NAGL 20171222
                flxFluctuacion.TextMatrix(X, 3) = rs!cPersNombre
                flxFluctuacion.TextMatrix(X, 5) = Format(rs!dFechaVencimiento, "dd/mm/yyyy")
                flxFluctuacion.TextMatrix(X, 6) = rs!cTpoMoneda
                flxFluctuacion.TextMatrix(X, 7) = Format(rs!nSaldo, "#,##0.00")
                If psOpt = "0" Then
                    flxFluctuacion.TextMatrix(X, 8) = Format(0#, "#,##0.00")
                    flxFluctuacion.TextMatrix(X, 9) = Format(rs!nSaldo, "#,##0.00")
                Else
                    flxFluctuacion.TextMatrix(X, 8) = Format(rs!nFluctuacion, "#,##0.00")
                    flxFluctuacion.TextMatrix(X, 9) = Format(rs!nValorLibroTotal, "#,##0.00")
                End If
                    flxFluctuacion.TextMatrix(X, 10) = rs!cPersCod
                rs.MoveNext
            Next
        End If
End Sub

Private Sub cmdBuscarFluctuacion_Click()
Dim rs As New ADODB.Recordset
Dim X As Integer
Dim pdFecha As Date

If txtFechaFluctuacion.Text = "" Or txtFechaFluctuacion.Text = "__/__/____" Then
      MsgBox "Debe Ingresar la Fecha de Fluctuación", vbOKOnly + vbInformation, "Atención"
      txtFechaFluctuacion.SetFocus
ElseIf ValFecha(txtFechaFluctuacion) = False Then
      txtFechaFluctuacion.SetFocus
Else
      pdFecha = txtFechaFluctuacion.Text
      If (Year(pdFecha) < 2050 Or Month(pdFecha) < 12) Then
              If ValRegFluctuacionValores(pdFecha) Then
                   'Set rsValNew = oValor.CargaPatrimonioEfectivo(pdFecha)
                   'If Not (rsValNew.BOF And rsValNew.EOF) Then
                       'txtFechaPatrimonio = Format(rsValNew!dFechaPatrimonio, "dd/mm/yyyy") 'NAGL 20171221
                       'txtSolesPatr = Format(rsValNew!nPatrimonioMN, "###,##0.00")
                       'txtDolaresPatr = Format(rsValNew!nPatrimonioME, "###,##0.00")
                   'Else
                       'txtFechaPatrimonio = Format(DateAdd("D", -Day(pdFecha), pdFecha), "dd/mm/yyyy")
                       'txtSolesPatr = Format(0, "###,##0.00")
                       'txtDolaresPatr = Format(0, "###,##0.00")
                   'End If
              'End If
                   'Set rs = oValor.DevuelveFluctuacionesValoresDet(pdFecha, "1")
                   'flxFluctuacion.Clear
                   'FormateaFlex flxFluctuacion
            
                    'If Not (rs.EOF And rs.BOF) Then
                        'For X = 1 To rs.RecordCount
                            'flxFluctuacion.AdicionaFila
                            'flxFluctuacion.TextMatrix(X, 1) = rs!cCtaCod
                            'flxFluctuacion.TextMatrix(X, 2) = rs!cCtaNombre
                            'flxFluctuacion.TextMatrix(X, 3) = Format(rs!dFechaVencimiento, "dd/mm/yyyy")
                            'flxFluctuacion.TextMatrix(X, 4) = rs!cTpoMoneda
                            'flxFluctuacion.TextMatrix(X, 5) = Format(rs!nSaldo, "#,##0.00")
                            'flxFluctuacion.TextMatrix(X, 6) = Format(rs!nFluctuacion, "#,##0.00")
                            'flxFluctuacion.TextMatrix(X, 7) = Format(rs!nValorLibroTotal, "#,##0.00")
                            'rs.MoveNext
                        'Next
                    'End If 'Comentado by NAGL 20171221
                    Call CargarFluctuacionesValores(pdFecha, "1")
                    SumaValorLibro 'NAGL 20171221
              End If 'NAGL 20171221
        Else
             MsgBox "Debe Ingresar la Fecha de Fluctuación correcta.", vbOKOnly + vbInformation, "Atención"
        End If
End If
End Sub

Private Function LlenarRsFluctuacionValoresDet(ByVal feControl As FlexEdit) As ADODB.Recordset

 Dim rsVD As New ADODB.Recordset
 Dim nIndex As Integer
 
  If feControl.Rows >= 2 Then
         If feControl.TextMatrix(nIndex, 1) = "" Then
            Exit Function
        End If
        
            rsVD.CursorType = adOpenStatic
            rsVD.Fields.Append "cCtaCod", adVarChar, 22, adFldIsNullable
            rsVD.Fields.Append "cPersCod", adVarChar, 13, adFldIsNullable
            rsVD.Fields.Append "dfechaVencimiento", adDate, adFldIsNullable
            rsVD.Fields.Append "cTpoMoneda", adVarChar, 20, adFldIsNullable
            rsVD.Fields.Append "nSaldo", adDouble, adFldIsNullable
            rsVD.Fields.Append "nFluctuacion", adDouble, adFldIsNullable
            rsVD.Fields.Append "nValorLibroTotal", adDouble, adFldIsNullable
            rsVD.Open

        For nIndex = 1 To feControl.Rows - 1
            rsVD.AddNew
            rsVD.Fields("cCtaCod") = feControl.TextMatrix(nIndex, 1)
            rsVD.Fields("cPersCod") = feControl.TextMatrix(nIndex, 10)
            rsVD.Fields("dfechaVencimiento") = feControl.TextMatrix(nIndex, 5)
            rsVD.Fields("cTpoMoneda") = feControl.TextMatrix(nIndex, 6)
            rsVD.Fields("nSaldo") = feControl.TextMatrix(nIndex, 7)
            rsVD.Fields("nFluctuacion") = feControl.TextMatrix(nIndex, 8)
            rsVD.Fields("nValorLibroTotal") = feControl.TextMatrix(nIndex, 9)
            rsVD.Update
            rsVD.MoveFirst
        Next
    End If
    Set LlenarRsFluctuacionValoresDet = rsVD
End Function

Private Sub cmdGuardar_Click()
    Dim lsItemValor As New ADODB.Recordset
    Dim i As Integer
    Dim Registro As Boolean
    Dim lsMovNro As String
    Dim dFechaValor As Date
    Dim oCont As New NContFunciones
    'Added by NAGL Según RFC1712260004**********
    Dim X As Integer, nFilasTotal As Integer, nCantFal As Integer
    nFilasTotal = flxFluctuacion.Rows - 1
    nCantFal = 0
    For X = 1 To nFilasTotal
        If flxFluctuacion.TextMatrix(X, 3) = "" Then
           nCantFal = nCantFal + 1
        End If
    Next X
'    '****END NAGL *************************
        If (txtSolesPatr.Text = "" Or txtSolesPatr.Text = "0.00") Then
                 MsgBox "Falta ingresar el Patrimonio  Efectivo en MN", vbInformation, "Aviso"
                  txtSolesPatr.SetFocus
        ElseIf Len(txtSolesPatr.Text) > 20 Then
                MsgBox "El valor del Patrimonio Efectivo en MN es incorrecto", vbInformation, "Aviso"
                txtSolesPatr.SetFocus
        ElseIf (txtFechaFluctuacion.Text = "" Or txtFechaFluctuacion.Text = "__/__/____") Then
                MsgBox "Debe Ingresar la Fecha de Fluctuación.", vbOKOnly + vbInformation, "Atención"
                txtFechaFluctuacion.SetFocus
        ElseIf (txtFechaPatrimonio = "" Or txtFechaPatrimonio = "__/__/____") Then
                MsgBox "Debe Ingresar la Fecha de Patrimonio.", vbOKOnly + vbInformation, "Atención"
                txtFechaPatrimonio.SetFocus
        Else
            If ValFecha(txtFechaFluctuacion) = True And ValFecha(txtFechaPatrimonio) = True Then
               pdFechaFluct = Format(txtFechaFluctuacion.Text, "dd/mm/yyyy")
               pdFecPatr = Format(txtFechaPatrimonio.Text, "dd/mm/yyyy")
               If Year(pdFechaFluct) < 2050 And Year(pdFecPatr) < 2050 And Month(pdFechaFluct) < 13 And Month(pdFecPatr) < 13 Then
                    If MsgBox("¿Desea Registrar las Fluctuaciones Ingresadas.", vbInformation + vbYesNo, "Atención") = vbYes Then
                       '***NAGL Según RFC1712260004**'
                       If nCantFal <> 0 Then
                          If MsgBox("No se ha ingresado completamente la información de los Emisores...Desea continuar?", vbInformation + vbYesNo, "Atención") = vbNo Then
                             flxFluctuacion.SetFocus
                             Exit Sub
                          End If
                       End If
                       '***END NAGL 20180116******'
                       lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                       Call oValor.ControlFluctuacionValores(Format(txtFechaFluctuacion.Text, "dd/mm/yyyy"), Format(txtFechaPatrimonio.Text, "dd/mm/yyyy"), CDbl(txtSolesPatr.Text), CDbl(txtDolaresPatr.Text), lsMovNro, LlenarRsFluctuacionValoresDet(flxFluctuacion))
                       MsgBox "Los datos se guardaron satisfactoriamente.", vbOKOnly + vbInformation, "Atención"
        
                       If MsgBox("¿Desea Generar el Anexo 16A: Cuadro de Liquidez por plazos de Vencimiento de corto Plazo..", vbInformation + vbYesNo, "Atención") = vbYes Then
                          Unload Me
                          'frmReportes.ReporteAnexo16A (pdFechaFluct) 'Comentado by NAGL 20201211
                          frmAnx16ALiquidezPlazoVencNew.ReporteAnexo16A (pdFechaFluct) 'NAGL 202012 Según ACTA N°094-2020
                       End If
                    End If
                Else
                     MsgBox "La Fecha Ingresada es incorrecta.", vbInformation, "Aviso"
                End If
            End If
        End If
End Sub

Private Function ValRegFluctuacionValores(pdFecha As Date, Optional psOpt As String) As Boolean 'NAGL 20171221 Agregó variable Optional
Dim DAnxVal As New DAnexoRiesgos
Dim psRegistro As String
 psRegistro = DAnxVal.ObtenerRegFluctuacionValores(pdFecha)
 If psOpt = "Load1" Then
    If psRegistro = "0" Then
       ValRegFluctuacionValores = False
    Else
      ValRegFluctuacionValores = True
    End If 'NAGL 20171221
 Else
    If psRegistro = "0" Then
            If MsgBox("No existen datos con la Fecha Registrada, Desea continuar?", vbInformation + vbYesNo, "Atención") = vbNo Then
                txtFechaFluctuacion.SetFocus
                Exit Function
            End If
    End If
    ValRegFluctuacionValores = True
 End If 'NAGL 20171221
End Function 'NAGL 20170425

Private Sub txtFechaFluctuacion_GotFocus()
    fEnfoque txtFechaFluctuacion
End Sub

Private Sub txtFechaFluctuacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (txtFechaFluctuacion.Text = "" Or txtFechaFluctuacion.Text = "__/__/____") Then
        MsgBox "Debe Ingresar la Fecha de Fluctuación.", vbOKOnly + vbInformation, "Atención"
        txtFechaFluctuacion.SetFocus
    Else
         If ValFecha(txtFechaFluctuacion) = True Then
            pdFechaFluct = txtFechaFluctuacion.Text
            If (Year(pdFechaFluct) < 2050 And Month(pdFechaFluct) < 13) Then
               txtFechaPatrimonio.SetFocus
               txtFechaPatrimonio.Text = DateAdd("d", -Day(pdFechaFluct), pdFechaFluct)
            Else
               MsgBox "Fecha Incorrecta.", vbOKOnly + vbInformation, "Atención"
            End If
         End If
    End If
End If
End Sub

Private Sub txtFechaPatrimonio_GotFocus()
  fEnfoque txtFechaPatrimonio
End Sub

Private Sub txtFechaPatrimonio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If (txtFechaPatrimonio.Text = "" Or txtFechaPatrimonio.Text = "__/__/____") Then
            MsgBox "Debe Ingresar la Fecha de Patrimonio.", vbOKOnly + vbInformation, "Atención"
            txtFechaPatrimonio.SetFocus
        Else
             If ValFecha(txtFechaPatrimonio) = True Then
                pdFecPatr = txtFechaPatrimonio.Text
                If (Year(pdFecPatr) < 2050 And Month(pdFecPatr) < 13) Then
                    txtSolesPatr.SetFocus
                Else
                    MsgBox "Fecha Incorrecta.", vbOKOnly + vbInformation, "Atención"
                End If
                
             End If
        End If
    End If
End Sub

Private Sub txtSolesPatr_GotFocus()
    fEnfoque txtSolesPatr
End Sub

Private Sub txtSolesPatr_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtSolesPatr, KeyAscii)
    If KeyAscii = 13 Then
        txtDolaresPatr.SetFocus
        txtSolesPatr = Format(txtSolesPatr, "#,##0.00")
        CalculaPatrimonioDolares
    End If
End Sub

Private Sub txtDolaresPatr_GotFocus()
    fEnfoque txtSolesPatr
End Sub

Private Sub txtDolaresPatr_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtDolaresPatr, KeyAscii)
    If KeyAscii = 13 Then
        cmdGuardar.SetFocus
    End If
End Sub

Public Sub CalculaPatrimonioDolares()
Dim lnTipoCambioFC As Currency
    If (txtSolesPatr.Text = "" Or txtSolesPatr.Text = "0.00") Then
        MsgBox "Debe Ingresar el Patrimonio en Efectivo MN", vbOKOnly + vbInformation, "Atención"
        txtSolesPatr.SetFocus
        
    ElseIf (txtFechaFluctuacion.Text = "" Or txtFechaFluctuacion.Text = "__/__/____") Then
        MsgBox "Debe Ingresar la Fecha de Fluctuación", vbOKOnly + vbInformation, "Atención"
        txtFechaFluctuacion.SetFocus
        
    ElseIf Val(txtFechaFluctuacion) Then
        pdFechaFluct = txtFechaFluctuacion.Text
        If (Year(pdFechaFluct) < 2050 And Month(pdFechaFluct) < 13) Then
            If Month(pdFechaFluct) = Month(DateAdd("d", 1, pdFechaFluct)) Then
                lnTipoCambioFC = Format(oCambio.EmiteTipoCambio(pdFechaFluct, TCFijoDia), "#,##0.0000")
            Else
                lnTipoCambioFC = Format(oCambio.EmiteTipoCambio(DateAdd("d", 1, pdFechaFluct), TCFijoDia), "#,##0.0000")
            End If
            
            If lnTipoCambioFC = 0 Then
                MsgBox "No existe Tipo de Cambio para la fecha ingresada", vbOKOnly + vbInformation, "Atención"
                txtDolaresPatr = Format(0, "#,##0.00")
            Else
                txtDolaresPatr = Format(CCur(txtSolesPatr.Text) / lnTipoCambioFC, "#,##0.00")
            End If
        Else
           MsgBox "El cálculo no será posible por la fecha de fluctuación Ingresada", vbOKOnly + vbInformation, "Atención"
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub flxFluctuacion_RowColChange()
    If flxFluctuacion.col = 8 Then
            flxFluctuacion.AvanceCeldas = Vertical
        Else
            flxFluctuacion.AvanceCeldas = Horizontal
            'cmdGuardar.SetFocus
    End If
End Sub

Private Sub flxFluctuacion_EnterCell()
  If flxFluctuacion.col = 7 Or flxFluctuacion.col = 8 Then
        ValorCelda = flxFluctuacion.TextMatrix(flxFluctuacion.row, flxFluctuacion.col)
  End If 'Condicional by NAGL 20180116
End Sub

Private Sub flxFluctuacion_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
Dim oPersona As New UPersona
Set oPersona = frmBuscaPersona.Inicio()
If Not oPersona Is Nothing Then
    flxFluctuacion.TextMatrix(pnRow, 10) = oPersona.sPersCod
    flxFluctuacion.TextMatrix(pnRow, 3) = oPersona.sPersNombre
End If
End Sub 'NAGL 20171221

Private Sub flxFluctuacion_OnCellChange(pnRow As Long, pnCol As Long)
    Dim valorNew As Currency
    Dim ValorAnterior As Currency
    
    If (ValorCelda <> "") Then
        ValorAnterior = CDbl(ValorCelda)
    End If
             If (IsNumeric(flxFluctuacion.TextMatrix(pnRow, pnCol)) And Len(flxFluctuacion.TextMatrix(pnRow, pnCol)) < 15) Then
                valorNew = CDbl(flxFluctuacion.TextMatrix(pnRow, pnCol))
                If (pnCol = 8) Then
                    flxFluctuacion.TextMatrix(pnRow, 9) = Format(CDbl(flxFluctuacion.TextMatrix(pnRow, 7)) + valorNew, "#,##0.00")
                Else
                    If (valorNew < 0) Then
                        MsgBox "No se puede asignar un valor Negativo", vbInformation, "Aviso"
                        flxFluctuacion.TextMatrix(pnRow, pnCol) = Format(ValorAnterior, "###,##0.00")
                        Exit Sub
                    Else
                        flxFluctuacion.TextMatrix(pnRow, 9) = Format(valorNew + CDbl(flxFluctuacion.TextMatrix(pnRow, 8)), "#,##0.00")
                    End If
                End If
             Else
                If (pnCol = 8) Then
                    flxFluctuacion.TextMatrix(pnRow, pnCol) = Format(ValorAnterior, "###,##0.00")
                    flxFluctuacion.TextMatrix(pnRow, 9) = Format(CDbl(flxFluctuacion.TextMatrix(pnRow, 7)) + ValorAnterior, "#,##0.00")
                Else
                    flxFluctuacion.TextMatrix(pnRow, pnCol) = Format(ValorAnterior, "###,##0.00")
                    flxFluctuacion.TextMatrix(pnRow, 9) = Format(ValorAnterior + CDbl(flxFluctuacion.TextMatrix(pnRow, 8)), "#,##0.00")
                End If
             End If
            SumaValorLibro 'NAGL 20171221
End Sub

Public Sub SumaValorLibro()
Dim nSaldoValLIBMN As Double, nSaldoValLIBME As Double
Dim nNroRegVal As Integer
nSaldoValLIBMN = 0
nSaldoValLIBME = 0
pnRow = 1

If (flxFluctuacion.TextMatrix(flxFluctuacion.Rows - 1, 0) <> "") Then
    nNroRegVal = CInt(flxFluctuacion.TextMatrix(flxFluctuacion.Rows - 1, 0))
    If nNroRegVal <> 0 Then
         For ix = 1 To nNroRegVal
            If (CStr(flxFluctuacion.TextMatrix(pnRow, 6)) = "PEN") Then
                nSaldoValLIBMN = CDbl(flxFluctuacion.TextMatrix(pnRow, 9)) + nSaldoValLIBMN
                
            ElseIf (CStr(flxFluctuacion.TextMatrix(pnRow, 6)) = "USD") Then
                nSaldoValLIBME = CDbl(flxFluctuacion.TextMatrix(pnRow, 9)) + nSaldoValLIBME
            End If
            pnRow = pnRow + 1
         Next ix
    Else
        nSaldoValLIBMN = Format(0, "#,#0.00")
        nSaldoValLIBME = Format(0, "#,#0.00")
    End If
    txtVALIBMNTotal.Text = Format(nSaldoValLIBMN, "#,#0.00")
    txtVALIBMETotal.Text = Format(nSaldoValLIBME, "#,#0.00")
Else
    txtVALIBMNTotal.Text = Format(0, "#,#0.00")
    txtVALIBMETotal.Text = Format(0, "#,#0.00")
End If
End Sub 'NAGL 20171221








