VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmPolizaSeguPatriDistri 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Distribución de Polizas de Seguros Patrimoniales"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12585
   Icon            =   "frmPolizaSeguPatriDistri.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   12585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExtorna 
      Caption         =   "Extorna Asiento"
      Height          =   345
      Left            =   2520
      TabIndex        =   17
      Top             =   6360
      Width           =   1560
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   325
      Left            =   11400
      TabIndex        =   16
      Text            =   "0"
      Top             =   0
      Width           =   1125
   End
   Begin VB.TextBox txtTipCamCompraSBS 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   325
      Left            =   9360
      TabIndex        =   15
      Text            =   "0"
      Top             =   0
      Width           =   1125
   End
   Begin VB.ComboBox cboTpo 
      Height          =   315
      Left            =   4695
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   0
      Width           =   4020
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Height          =   345
      Left            =   45
      TabIndex        =   8
      Top             =   6375
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar / Distribuir Gstos."
      Height          =   345
      Left            =   4215
      TabIndex        =   7
      Top             =   6360
      Width           =   2160
   End
   Begin VB.ComboBox cmbMes 
      Height          =   315
      ItemData        =   "frmPolizaSeguPatriDistri.frx":030A
      Left            =   1935
      List            =   "frmPolizaSeguPatriDistri.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   30
      Width           =   2220
   End
   Begin MSMask.MaskEdBox mskAnio 
      Height          =   300
      Left            =   600
      TabIndex        =   4
      Top             =   60
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   9375
      TabIndex        =   2
      Top             =   6375
      Width           =   960
   End
   Begin VB.CommandButton cmdDeprecia 
      Caption         =   "&Generar"
      Height          =   345
      Left            =   8340
      TabIndex        =   1
      Top             =   6360
      Width           =   960
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   7290
      TabIndex        =   0
      Top             =   6375
      Width           =   960
   End
   Begin Sicmact.FlexEdit FlexEdit1 
      Height          =   5775
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   10186
      Cols0           =   26
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   $"frmPolizaSeguPatriDistri.frx":030E
      EncabezadosAnchos=   "300-1000-0-3000-1200-1200-1200-1200-1200-1200-1200-1200-1000-1000-1000-1000-1200-1200-1200-1000-0-0-0-0-0-0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L-R-R-R-L-L-R-R-R-C-C-L-L-R-R-R-L-C-C-C-L-C-R"
      FormatosEdit    =   "0-0-0-0-2-2-2-0-0-2-2-2-0-0-0-0-2-2-2-0-0-0-0-0-0-2"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      Appearance      =   0
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label Label3 
      Caption         =   "TC Venta :"
      Height          =   255
      Left            =   10560
      TabIndex        =   14
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "TC Fijo :"
      Height          =   255
      Left            =   8760
      TabIndex        =   13
      Top             =   60
      Width           =   615
   End
   Begin VB.Label lblTipo 
      Caption         =   "Tipo :"
      Height          =   210
      Left            =   4215
      TabIndex        =   11
      Top             =   60
      Width           =   510
   End
   Begin VB.OLE OleExcel 
      Appearance      =   0  'Flat
      AutoActivate    =   3  'Automatic
      Enabled         =   0   'False
      Height          =   240
      Left            =   1215
      SizeMode        =   1  'Stretch
      TabIndex        =   9
      Top             =   6420
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblMes 
      Caption         =   "Mes :"
      Height          =   210
      Left            =   1440
      TabIndex        =   6
      Top             =   75
      Width           =   510
   End
   Begin VB.Label lblAnio 
      Caption         =   "Año :"
      Height          =   210
      Left            =   135
      TabIndex        =   3
      Top             =   90
      Width           =   705
   End
End
Attribute VB_Name = "frmPolizaSeguPatriDistri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsCaption As String
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim nMES As Integer
Dim nAnio As Integer
Dim dFecha As Date

Public Sub Ini(psCaption As String)
    lsCaption = psCaption
    Me.Show 1
End Sub



Private Sub cmdDeprecia_Click()
    
    If Not IsNumeric(Me.mskAnio.Text) Then
        MsgBox "Debe Ingresar un año Valido.", vbInformation, "Aviso"
        Me.mskAnio.SetFocus
        Exit Sub
    ElseIf Me.cmbMes.Text = "" Then
        MsgBox "Debe Ingresar un mes Valido.", vbInformation, "Aviso"
        Me.cmbMes.SetFocus
        Exit Sub
    ElseIf Me.cboTpo.Text = "" Then
        MsgBox "Debe Ingresar un tipo de depreciacion Valido.", vbInformation, "Aviso"
        Me.cboTpo.SetFocus
        Exit Sub
    ElseIf Me.txtTipCamCompraSBS.Text = 0 Or Me.Text1.Text = 0 Then
        MsgBox "Por favor ingrese ambos tipos de cambio.", vbInformation, "Aviso"
        Me.txtTipCamCompraSBS.SetFocus
        Exit Sub
    End If
       
    Call llenagrid

End Sub

'Private Sub cmdEliminar_Click()
'    frmLogAFMant.Ini 1, False, False, True, Me.Flex.TextMatrix(Me.Flex.Row, 1), Me.Flex.TextMatrix(Me.Flex.Row, 3), CDate(Me.Flex.TextMatrix(Me.Flex.Row, 18)), Me.Flex.TextMatrix(Me.Flex.Row, 4), 0, 0, Me.Flex.TextMatrix(Me.Flex.Row, 7), Me.Flex.TextMatrix(Me.Flex.Row, 8), Me.Flex.TextMatrix(Me.Flex.Row, 19) & Me.Flex.TextMatrix(Me.Flex.Row, 20), Me.Flex.TextMatrix(Me.Flex.Row, 21), Me.Flex.TextMatrix(Me.Flex.Row, 17)
'End Sub

Private Sub cmdExtorna_Click()
    Dim oContFunc As NContFunciones

    Dim oAge As DOperacion
    Dim overi As DOperacion
    Dim rs As ADODB.Recordset

    Dim lcFecPeriodo As String

    Dim lcMovNroExtorno As Integer
    Dim lnMovNroExtorno As Integer
    Dim lnMovNroExt As Long
    
    lcFecPeriodo = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    
    If Len(Trim(Right(cmbMes.Text, 2))) = 1 Then
        lcFecPeriodo = mskAnio.Text & "0" & Trim(Right(cmbMes.Text, 2))
    Else
        lcFecPeriodo = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    End If
    
    Set overi = New DOperacion
    Set rs = overi.VerificaAsientoCont(gsOpeCod, lcFecPeriodo)
    Set overi = Nothing

    If (rs.EOF And rs.BOF) Then
        MsgBox "Este periodo no tiene asiento contable.", vbCritical, "Aviso!"
        Exit Sub
    Else
        lnMovNroExt = rs!nMovNro
    End If

    If MsgBox("¿Desea extornar el asiento contable? ", vbQuestion + vbYesNo, "Atención") = vbNo Then Exit Sub

    nMES = Val(Trim(Right(cmbMes.Text, 2)))
    nAnio = Val(mskAnio.Text)
    dFecha = DateAdd("m", 1, "01/" & Format(nMES, "00") & "/" & Format(nAnio, "0000")) - 1
    Set oContFunc = New NContFunciones
    If Not oContFunc.PermiteModificarAsiento(Format(dFecha, gsFormatoMovFecha), False) Then
       Set oContFunc = Nothing
       MsgBox "Imposible realizar este proceso ya que la fecha ingresada pertenece a un mes cerrado.", vbInformation, "Aviso"
       Exit Sub
    End If

    Set oAge = New DOperacion
    Call oAge.EliminaAsientoCont(lnMovNroExt)
    Set oAge = Nothing
    
    MsgBox "El asiento contable fue extornado satisfactoriamente.", vbInformation, "Aviso"
        
        
End Sub

Private Sub CmdGrabar_Click()
    
    Dim oMov As DMov
    Set oMov = New DMov
    
    Dim oDep As DOperacion
    Set oDep = New DOperacion
       
    Dim oConect As DConecta
    Set oConect = New DConecta
    
    Dim lnMovNro As Long
    Dim lsMovNro As String
    
    Dim lsTipo As String
    Dim lsFecha As String
    Dim i As Integer
    Dim lnI As Long
    Dim lnContador As Long
    Dim lsCtaCont As String
    Dim oPrevio As clsPrevioFinan
    Dim oAsiento As NContImprimir
    Dim nConta As Integer, lcCtaDif As String
    Dim overi As DOperacion
    Dim lnDebe As Double, lnHaber As Double, lnTotHaber As Double, lnTotDebe As Double
    Dim lnDebeME As Double, lnTotDebeME As Double
    Dim oContFunc As NContFunciones
    Dim lnMontoPrin As Currency
    Dim rsAgesDistrib As ADODB.Recordset
    Dim lsSql As String, lcCtaCont As String
    Dim rsBuscaCuenta As ADODB.Recordset
    Dim lnItemDistri As Integer
    Dim lnRegImporte As Currency
    Dim lnItemPrin As Integer
    Dim lnImpoPrin As Currency
    Dim lnMontoPrME As Currency
    Dim lnRegImporMETot As Currency
    Dim lnMonto As Currency
    Dim lsCtasInexis As String
    Dim lnRegImporME As Currency
    Dim lnRegImporteTot As Currency
    Dim lnImpoPrinME As Currency
    Dim lnMontoDebePrin As Currency
    Set oPrevio = New clsPrevioFinan
    Set oAsiento = New NContImprimir
   
    Dim rs As ADODB.Recordset, rs1 As ADODB.Recordset
    
    Dim ldFechaDepre As Date
    Dim ldFechaRegistro As Date
    
  Select Case Val(Trim(Right(cboTpo.Text, 2)))
            Case 1
                gsOpeCod = "300430"
            Case 2
                gsOpeCod = "300431"
            Case 3
                gsOpeCod = "300432"
            Case 4
                gsOpeCod = "300433"
            Case 5
                gsOpeCod = "300434"
            Case 6
                gsOpeCod = "300435"
        End Select
    
    lsTipo = Trim(Right(cboTpo.Text, 2))

    lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    
    If Len(Trim(Right(cmbMes.Text, 2))) = 1 Then
        lsFecha = mskAnio.Text & "0" & Trim(Right(cmbMes.Text, 2))
    Else
        lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    End If
    
    Set rs = New ADODB.Recordset
   
    Set overi = New DOperacion
    Set rs = overi.VerificaAsientoCont(gsOpeCod, lsFecha)
    Set rs1 = overi.ObtieneCtasIngresosGastosPoliSeguPatri
    Set overi = Nothing
    
    If Not rs.EOF Then
        MsgBox "Asiento ya fue generado ", vbCritical, "Aviso!"
        Exit Sub
    End If
    
    If MsgBox("¿Desea Procesar? ", vbQuestion + vbYesNo, "Atención") = vbNo Then Exit Sub

    nMES = Val(Trim(Right(cmbMes.Text, 2)))
    nAnio = Val(mskAnio.Text)
    dFecha = DateAdd("m", 1, "01/" & Format(nMES, "00") & "/" & Format(nAnio, "0000")) - 1
    Set oContFunc = New NContFunciones
    If Not oContFunc.PermiteModificarAsiento(Format(dFecha, gsFormatoMovFecha), False) Then
       Set oContFunc = Nothing
       MsgBox "Imposible grabar el asiento en un mes cerrado.", vbInformation, "Aviso"
       Exit Sub
    End If

    ldFechaRegistro = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & Right(Me.cmbMes.Text, 2) & "/" & Me.mskAnio.Text)))
    
    lnMontoPrin = 0
    lsCtasInexis = ""
    
    ' genera los montos previos a distribuir por agencia
    
    oMov.BeginTrans
    
        lsMovNro = oMov.GeneraMovNro(ldFechaRegistro, Right(gsCodAge, 2), gsCodUser)
        oMov.InsertaMov lsMovNro, gsOpeCod, "REG. " & Trim(Mid(Me.cboTpo.Text, 1, Len(Me.cboTpo.Text) - 2))
        lnMovNro = oMov.GetnMovNro(lsMovNro)
        
        For lnI = 1 To Me.FlexEdit1.Rows - 1
            If Len(Trim(Me.FlexEdit1.TextMatrix(lnI, 1))) > 0 And Val(Me.FlexEdit1.TextMatrix(lnI, 4)) > 0 And Trim(Me.FlexEdit1.TextMatrix(lnI, 2)) = "01" Then
                lnMontoPrin = Round(Me.FlexEdit1.TextMatrix(lnI, 17), 2)
                lnMontoDebePrin = Round(Me.FlexEdit1.TextMatrix(lnI, 16), 2)
                lnMontoPrME = Round(Me.FlexEdit1.TextMatrix(lnI, 10), 2)
            End If
        Next lnI
    
        If lnMontoPrin > 0 Then
            Set rsAgesDistrib = New ADODB.Recordset
            lsSql = " exec stp_sel_AgenciaPorcentajeGastos "
            Set rsAgesDistrib = oMov.CargaRecordSet(lsSql)
            
            lsCtaCont = oDep.ObtieneCtasPoliSeguPatri(gsOpeCod, "D", 0)
            lsCtaCont = Replace(lsCtaCont, "M", "2")
            
            lnItemDistri = 0
            
            Do While Not rsAgesDistrib.EOF
                lcCtaCont = Replace(lsCtaCont, "AG", rsAgesDistrib!cAgecod)
    
                Set rsBuscaCuenta = oMov.CargaRecordSet("select * From CtaCont where cCtaContCod like '" & lcCtaCont & "' ")
                If Not (rsBuscaCuenta.EOF And rsBuscaCuenta.BOF) Then
                    lnItemDistri = lnItemDistri + 1
                    
                    lnRegImporte = Round(lnMontoPrin * rsAgesDistrib!nAgePorcentaje / 100, 2)
                    lnRegImporME = Round(lnMontoPrME * rsAgesDistrib!nAgePorcentaje / 100, 2)
                    
                    lnRegImporteTot = lnRegImporteTot + lnRegImporte
                    lnRegImporMETot = lnRegImporMETot + lnRegImporME
                    
                    If rsAgesDistrib!cAgecod = "01" Then
                        lnItemPrin = lnItemDistri
                        lnImpoPrin = lnRegImporte
                        lnImpoPrinME = lnRegImporME
                    End If
                    
                    oMov.InsertaMovCta lnMovNro, lnItemDistri, lcCtaCont, lnRegImporte
                    oMov.InsertaMovMe lnMovNro, lnItemDistri, lnRegImporME

                    lnMonto = lnMonto + lnRegImporte
                Else
                    lsCtasInexis = lsCtasInexis + lcCtaCont + "-"
                End If
                rsAgesDistrib.MoveNext
            Loop
            
            lnItemDistri = lnItemDistri + 1
             
            lsCtaCont = oDep.ObtieneCtasPoliSeguPatri(gsOpeCod, "H", 0)
            lsCtaCont = Replace(lsCtaCont, "M", "2")
    
            oMov.InsertaMovCta lnMovNro, lnItemDistri, lsCtaCont, lnMontoDebePrin * -1
            oMov.InsertaMovMe lnMovNro, lnItemDistri, lnMontoPrME * -1
                       
            If lnRegImporteTot > lnMontoPrin Then
                lcCtaCont = Replace(lsCtaCont, "AG", "01")
                Set rsBuscaCuenta = oMov.CargaRecordSet("select * From CtaCont where cCtaContCod like '" & lcCtaCont & "' ")
                If Not (rsBuscaCuenta.EOF And rsBuscaCuenta.BOF) Then
                    oMov.ActualizaMovCta lnMovNro, lnItemPrin, , lnImpoPrin - (lnMonto - lnMontoPrin) '' la dif manda a la ag princ
                Else
                    oMov.ActualizaMovCta lnMovNro, lnItemDistri, , lnRegImporte - (lnMonto - lnMontoPrin) '' si noexiste la age prin manda al ultimo age registrado
                End If
            ElseIf lnMontoPrin > lnRegImporteTot Then
                lcCtaCont = Replace(lsCtaCont, "AG", "01")
                Set rsBuscaCuenta = oMov.CargaRecordSet("select * From CtaCont where cCtaContCod like '" & lcCtaCont & "' ")
                If Not (rsBuscaCuenta.EOF And rsBuscaCuenta.BOF) Then
                    oMov.ActualizaMovCta lnMovNro, lnItemPrin, , lnImpoPrin + (lnMontoPrin - lnMonto) '' la dif manda a la ag princ
                Else
                    oMov.ActualizaMovCta lnMovNro, lnItemDistri, , lnRegImporte + (lnMontoPrin - lnMonto) '' si noexiste la age prin manda al ultimo age registrado
                End If
            End If
            
            If lnRegImporMETot > lnMontoPrME Then
                oMov.ActualizaMovMe lnMovNro, lnItemPrin, lnImpoPrinME - (lnRegImporMETot - lnMontoPrME)
            ElseIf lnMontoPrME > lnRegImporMETot Then
                oMov.ActualizaMovMe lnMovNro, lnItemPrin, lnImpoPrinME + (lnMontoPrME - lnRegImporMETot)
            End If
          
            If lnMontoDebePrin > lnMontoPrin Then
                lcCtaDif = rs1!cCtaGasto
                oMov.InsertaMovCta lnMovNro, nConta, lcCtaDif, (lnMontoDebePrin - lnMontoPrin)
            ElseIf lnMontoPrin > lnMontoDebePrin Then
                lcCtaDif = rs1!cCtaIngreso
                oMov.InsertaMovCta lnMovNro, nConta, lcCtaDif, (lnMontoPrin - lnMontoDebePrin) * -1
            End If
            
            RSClose rsAgesDistrib
            Set rsAgesDistrib = Nothing
        End If

        '****************************

        lnTotHaber = 0: lnHaber = 0: lnDebe = 0

        If lnMontoPrin > 0 Then
            nConta = lnItemDistri + 1
        Else
            nConta = 0
        End If

        For lnI = 1 To Me.FlexEdit1.Rows - 1
            If Len(Trim(Me.FlexEdit1.TextMatrix(lnI, 1))) > 0 And Val(Me.FlexEdit1.TextMatrix(lnI, 4)) > 0 Then
                    
                If lnMontoPrin = 0 Or (lnMontoPrin > 0 And Right(Me.FlexEdit1.TextMatrix(lnI, 2), 2) <> "01") Then
                    nConta = nConta + 1
                    lsCtaCont = oDep.ObtieneCtasPoliSeguPatri(gsOpeCod, "D", 0)
                    lsCtaCont = Replace(lsCtaCont, "AG", Right(Me.FlexEdit1.TextMatrix(lnI, 2), 2))
                    lsCtaCont = Replace(lsCtaCont, "M", "2")
    
                    lnDebe = Round(Me.FlexEdit1.TextMatrix(lnI, 17), 2)
                    lnHaber = Round(Me.FlexEdit1.TextMatrix(lnI, 16), 2)
    
                    lnTotDebe = lnTotDebe + lnDebe
                    lnTotHaber = lnTotHaber + lnHaber
    
                    lnDebeME = Round(Me.FlexEdit1.TextMatrix(lnI, 10), 2)
                    lnTotDebeME = lnTotDebeME + lnDebeME
                    
                    oMov.InsertaMovCta lnMovNro, nConta, lsCtaCont, lnDebe
                    oMov.InsertaMovMe lnMovNro, nConta, lnDebeME
                End If
                
            End If
        Next lnI

        If Me.FlexEdit1.Rows - 1 > 0 Then

            nConta = nConta + 1
             
            lsCtaCont = oDep.ObtieneCtasPoliSeguPatri(gsOpeCod, "H", 0)
            lsCtaCont = Replace(lsCtaCont, "M", "2")
    
            oMov.InsertaMovCta lnMovNro, nConta, lsCtaCont, lnTotHaber * -1
            oMov.InsertaMovMe lnMovNro, nConta, lnTotDebeME * -1
            
            nConta = nConta + 1

            If lnTotHaber > lnTotDebe Then
                lcCtaDif = rs1!cCtaGasto
                oMov.InsertaMovCta lnMovNro, nConta, lcCtaDif, (lnTotHaber - lnTotDebe)
            ElseIf lnTotHaber < lnTotDebe Then
                lcCtaDif = rs1!cCtaIngreso
                oMov.InsertaMovCta lnMovNro, nConta, lcCtaDif, (lnTotDebe - lnTotHaber) * -1
            End If
        End If

    oMov.CommitTrans

    If lsCtasInexis <> "" Then
        MsgBox "Las siguientes cuentas no están en el Plan contable: " + Chr(10) + "lsCtasInexis", vbOKOnly, "Aviso"
    End If

    oPrevio.Show oAsiento.ImprimeAsientoContable(lsMovNro, 60, 80, "A S I E N T O  C O N T A B L E - " + Trim(Left(Me.cboTpo.Text, 30)) + " - " + CStr(dFecha)), "", True
End Sub

Private Sub cmdImprimir_Click()
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    
    If Me.FlexEdit1.TextMatrix(1, 1) = "" Then
        MsgBox "Debe Depreciar antes de imprimir.", vbInformation, "Aviso"
        Me.cmdDeprecia.SetFocus
        Exit Sub
    End If
    
    lsArchivoN = App.path & "\Spooler\" & Format(gdFecSis, "yyyymmdd") & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
       Call GeneraReporte(Me.FlexEdit1.GetRsNew)
       OleExcel.Class = "ExcelWorkSheet"
       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
    End If
    MousePointer = 0
End Sub

'Private Sub cmdModificar_Click()
'    frmLogAFMant.Ini 1, False, True, False, Me.Flex.TextMatrix(Me.Flex.Row, 1), Me.Flex.TextMatrix(Me.Flex.Row, 3), CDate(Me.Flex.TextMatrix(Me.Flex.Row, 18)), Me.Flex.TextMatrix(Me.Flex.Row, 4), 0, 0, Me.Flex.TextMatrix(Me.Flex.Row, 7), Me.Flex.TextMatrix(Me.Flex.Row, 8), Me.Flex.TextMatrix(Me.Flex.Row, 19) & Me.Flex.TextMatrix(Me.Flex.Row, 20), Me.Flex.TextMatrix(Me.Flex.Row, 21), Me.Flex.TextMatrix(Me.Flex.Row, 17)
'End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim oGen As DGeneral
    Set oGen = New DGeneral
    
    Set rs = oGen.GetConstante(1010)
    Me.cmbMes.Clear
    While Not rs.EOF
        cmbMes.AddItem rs.Fields(0) & Space(50) & rs.Fields(1)
        rs.MoveNext
    Wend
    
    Set rs = oGen.GetConstante(9077)
    Me.cboTpo.Clear
    While Not rs.EOF
        cboTpo.AddItem rs.Fields(0) & Space(50) & rs.Fields(1)
        rs.MoveNext
    Wend
    
    txtTipCamCompraSBS = "0.0000"
    Text1 = "0.0000"
   
    Me.mskAnio.Text = Format(gdFecSis, "yyyy")
     
    Me.Caption = lsCaption
End Sub

Private Sub GeneraReporte(prRs As ADODB.Recordset)
    Dim i As Integer
    Dim K As Integer
    Dim j As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Variant 'Currency
    Dim lnSer As Currency
    
    i = -1
    prRs.MoveFirst
    i = i + 1
    For j = 0 To prRs.Fields.Count - 1
        
        xlHoja1.Cells(i + 1, j + 1) = prRs.Fields(j).Name
    Next j
    
    While Not prRs.EOF
        If Len(prRs.Fields(0)) > 0 Then
            i = i + 1
            For j = 0 To prRs.Fields.Count - 1
                xlHoja1.Cells(i + 1, j + 1) = prRs.Fields(j)
            Next j
            'prRs.MoveNext
        End If
        prRs.MoveNext
    Wend
    
    i = i + 1
    xlHoja1.Range("A1:A" & Trim(Str(i))).Font.Bold = True
    xlHoja1.Range("B1:B" & Trim(Str(i))).Font.Bold = True
    xlHoja1.Range("1:1").Font.Bold = True
    
    xlHoja1.Range("D1:D" & Trim(Str(i))).NumberFormat = "#,##0.00"
    xlHoja1.Range("F1:F" & Trim(Str(i))).NumberFormat = "#,##0.00"
    xlHoja1.Range("I1:I" & Trim(Str(i))).NumberFormat = "#,##0.00"
    xlHoja1.Range("J1:J" & Trim(Str(i))).NumberFormat = "#,##0.00"
    xlHoja1.Range("N1:N" & Trim(Str(i))).NumberFormat = "#,##0.00"
    xlHoja1.Range("O1:O" & Trim(Str(i))).NumberFormat = "#,##0.00"
    xlHoja1.Range("P1:P" & Trim(Str(i))).NumberFormat = "#,##0.00"
    xlHoja1.Range("Z1:Z" & Trim(Str(i))).NumberFormat = "#,##0.00"
    xlHoja1.Range("AA1:AA" & Trim(Str(i))).NumberFormat = "#,##0.00"
    xlHoja1.Range("AB1:AB" & Trim(Str(i))).NumberFormat = "#,##0.00"
    xlHoja1.Range("AC1:AC" & Trim(Str(i))).NumberFormat = "#,##0.00"
    
    
    xlHoja1.Range("S1:S" & Trim(Str(i))).NumberFormat = "dd/mm/yyyy"
    xlHoja1.Range("W1:W" & Trim(Str(i))).NumberFormat = "dd/mm/yyyy"
    xlHoja1.Range("Y1:Y" & Trim(Str(i))).NumberFormat = "dd/mm/yyyy"
    
    xlHoja1.Range("D" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
    xlHoja1.Range("F" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
    xlHoja1.Range("I" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
    xlHoja1.Range("J" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
    xlHoja1.Range("N" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
    xlHoja1.Range("O" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
    xlHoja1.Range("P" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
    xlHoja1.Range("Z" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
    xlHoja1.Range("AA" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
    xlHoja1.Range("AB" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
    xlHoja1.Range("AC" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
    
    
    xlHoja1.Range("1:1").Font.Bold = True
    xlHoja1.Columns.AutoFit

    With xlHoja1.PageSetup
        .LeftHeader = ""
        .CenterHeader = "&""Arial,Negrita""&14Reporte de Activo Fijo Mes :  " & Trim(Left(Me.cmbMes, 15)) & " - " & Me.mskAnio.Text & " - " & Trim(Left(Me.cboTpo.Text, 15))
        .RightHeader = "&P"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 70
    End With

    With xlHoja1.Range("A1:AC" & Trim(Str(i + 1))).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A1:AC" & Trim(Str(i + 1))).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A1:AC" & Trim(Str(i + 1))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A1:AC" & Trim(Str(i + 1))).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A1:AC" & Trim(Str(i + 1))).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A1:AC" & Trim(Str(i + 1))).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End Sub

Private Sub GeneraReporteSaldoHistorico(prRs As ADODB.Recordset)
    Dim i As Integer
    Dim K As Integer
    Dim j As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Currency
    
    i = 1
    
    With xlHoja1.Range("A1:N1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("A1:N1").Merge
    xlHoja1.Range("A1:N1").FormulaR1C1 = " REPORTE DE SALDOS HISTORICOS "
    
    With xlHoja1.Range("A2:G2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("A2:G2").Merge
    xlHoja1.Range("A2:G2").FormulaR1C1 = " COSTO DEL ACTIVO FIJO "
    
    With xlHoja1.Range("H2:N2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("H2:N2").Merge
    xlHoja1.Range("H2:N2").FormulaR1C1 = " DEPRECIACION DE ACTIVOS FIJOS "
    
    prRs.MoveFirst
    While Not prRs.EOF
        i = i + 1
        
        If i = 2 Then
            xlHoja1.Cells(i + 1, 1) = "COD CTIVO"
            xlHoja1.Cells(i + 1, 2) = "COD. PATRIM."
            xlHoja1.Cells(i + 1, 3) = "F. ADQUIS"
            xlHoja1.Cells(i + 1, 4) = "SALDO AÑO ANT."
            xlHoja1.Cells(i + 1, 5) = "COMPRAS AÑO"
            xlHoja1.Cells(i + 1, 6) = "RETIROS AÑO"
            xlHoja1.Cells(i + 1, 7) = "SALDO ACTUAL"
            xlHoja1.Cells(i + 1, 8) = "DEP ACUM EJER ANT"
            xlHoja1.Cells(i + 1, 9) = "DEP AL MES ANT"
            xlHoja1.Cells(i + 1, 10) = "DEP DEP MES"
            xlHoja1.Cells(i + 1, 11) = "TOT DEP DEL EJER"
            xlHoja1.Cells(i + 1, 12) = "DEP ACUM DE RET"
            xlHoja1.Cells(i + 1, 13) = "DEP ACUM TOTAL"
            xlHoja1.Cells(i + 1, 14) = "VALOR EN LIBROS"
            
            i = i + 1
            
            xlHoja1.Cells(i + 1, 1) = prRs!Codigo & "-" & prRs!Serie
            xlHoja1.Cells(i + 1, 2) = Format(i - 2, "00000")
            xlHoja1.Cells(i + 1, 3) = "'" & Format(prRs!F_Compra, gsFormatoFechaView)
            xlHoja1.Cells(i + 1, 4) = Format(IIf(Year(prRs!F_Compra) <> Me.mskAnio.Text, prRs!Valor, 0), "#,##0.00")
            xlHoja1.Cells(i + 1, 5) = Format(IIf(Year(prRs!F_Compra) = Me.mskAnio.Text, prRs!Valor, 0), "#,##0.00")
            
            If prRs!Baja = "False" Then
                xlHoja1.Cells(i + 1, 6) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(i + 1, 6) = Format(prRs!Valor, "#,##0.00")
                Else
                    xlHoja1.Cells(i + 1, 6) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(i + 1, 7) = Format(xlHoja1.Cells(i + 1, 4) + xlHoja1.Cells(i + 1, 5) - xlHoja1.Cells(i + 1, 6), "#,##0.00")
            
            xlHoja1.Cells(i + 1, 8) = Format(IIf(Year(prRs!F_Compra) <> Me.mskAnio.Text, prRs!Dep_H_Ejer_Ant, 0), "#,##0.00")
            xlHoja1.Cells(i + 1, 9) = IIf(prRs!Baja = "False", Format(prRs!Dep_H_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(i + 1, 10) = Format(prRs!Dep_H_Mes, "#,##0.00")
            xlHoja1.Cells(i + 1, 11) = Format(xlHoja1.Cells(i + 1, 9) + xlHoja1.Cells(i + 1, 10), "#,##0.00")
            If prRs!Baja = "False" Then
                xlHoja1.Cells(i + 1, 12) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(i + 1, 12) = Format(prRs!Dep_H_Mes_Ant, "#,##0.00")
                Else
                    xlHoja1.Cells(i + 1, 12) = Format(0, "#,##0.00")
                End If
            End If
            xlHoja1.Cells(i + 1, 13) = Format(xlHoja1.Cells(i + 1, 8) + xlHoja1.Cells(i + 1, 11) - xlHoja1.Cells(i + 1, 12), "#,##0.00")
            xlHoja1.Cells(i + 1, 14) = Format(xlHoja1.Cells(i + 1, 13) - xlHoja1.Cells(i + 1, 7), "#,##0.00")
        Else
            xlHoja1.Cells(i + 1, 1) = prRs!Codigo & "-" & prRs!Serie
            xlHoja1.Cells(i + 1, 2) = Format(i - 2, "00000")
            xlHoja1.Cells(i + 1, 3) = "'" & Format(prRs!F_Compra, gsFormatoFechaView)
            xlHoja1.Cells(i + 1, 4) = Format(IIf(Year(prRs!F_Compra) <> Me.mskAnio.Text, prRs!Valor, 0), "#,##0.00")
            xlHoja1.Cells(i + 1, 5) = Format(IIf(Year(prRs!F_Compra) = Me.mskAnio.Text, prRs!Valor, 0), "#,##0.00")
            
            If prRs!Baja = "False" Then
                xlHoja1.Cells(i + 1, 6) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(i + 1, 6) = Format(prRs!Valor, "#,##0.00")
                Else
                    xlHoja1.Cells(i + 1, 6) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(i + 1, 7) = Format(xlHoja1.Cells(i + 1, 4) + xlHoja1.Cells(i + 1, 5) - xlHoja1.Cells(i + 1, 6), "#,##0.00")
            
            xlHoja1.Cells(i + 1, 8) = Format(IIf(Year(prRs!F_Compra) <> Me.mskAnio.Text, prRs!Dep_H_Ejer_Ant, 0), "#,##0.00")
            xlHoja1.Cells(i + 1, 9) = IIf(prRs!Baja = "False", Format(prRs!Dep_H_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(i + 1, 10) = Format(prRs!Dep_H_Mes, "#,##0.00")
            xlHoja1.Cells(i + 1, 11) = Format(xlHoja1.Cells(i + 1, 9) + xlHoja1.Cells(i + 1, 10), "#,##0.00")
            If prRs!Baja = "False" Then
                xlHoja1.Cells(i + 1, 12) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(i + 1, 12) = Format(prRs!Dep_H_Mes_Ant, "#,##0.00")
                Else
                    xlHoja1.Cells(i + 1, 12) = Format(0, "#,##0.00")
                End If
            End If
            xlHoja1.Cells(i + 1, 13) = Format(xlHoja1.Cells(i + 1, 8) + xlHoja1.Cells(i + 1, 11) - xlHoja1.Cells(i + 1, 12), "#,##0.00")
            xlHoja1.Cells(i + 1, 14) = Format(xlHoja1.Cells(i + 1, 13) - xlHoja1.Cells(i + 1, 7), "#,##0.00")
        End If
        prRs.MoveNext
    Wend
        
    xlHoja1.Cells.Select
    xlHoja1.Columns.AutoFit
    
    xlHoja1.Range("A1:N3").Select
    xlHoja1.Range("A1:N3").Font.Bold = True

    With xlHoja1.Range("A2:G2").Interior
        .ColorIndex = 36
        .Pattern = xlSolid
    End With
    With xlHoja1.Range("H2:N2").Interior
        .ColorIndex = 35
        .Pattern = xlSolid
    End With
    
    xlHoja1.Range("A2:N" & Trim(Str(i + 1))).Select
    xlHoja1.Range("A2:N" & Trim(Str(i + 1))).Borders(xlDiagonalDown).LineStyle = xlNone
    xlHoja1.Range("A2:N" & Trim(Str(i + 1))).Borders(xlDiagonalUp).LineStyle = xlNone
    
    With xlHoja1.Range("A2:N" & Trim(Str(i + 1))).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:N" & Trim(Str(i + 1))).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:N" & Trim(Str(i + 1))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:N" & Trim(Str(i + 1))).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:N" & Trim(Str(i + 1))).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:N" & Trim(Str(i + 1))).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
End Sub

Private Sub GeneraReporteSaldoAjustado(prRs As ADODB.Recordset)
    Dim i As Integer
    Dim K As Integer
    Dim j As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Currency
    
    i = 1
    
    With xlHoja1.Range("A1:S1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("A1:T1").Merge
    xlHoja1.Range("A1:T1").FormulaR1C1 = " REPORTE DE SALDOS AJUSTADOS "
    
    With xlHoja1.Range("A2:L2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("A2:L2").Merge
    xlHoja1.Range("A2:L2").FormulaR1C1 = " COSTO DEL AJUSTADO DE ACTIVO FIJO "
    
    With xlHoja1.Range("M2:T2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("M2:T2").Merge
    xlHoja1.Range("M2:T2").FormulaR1C1 = " DEPRECIACION AJUSTADA DE ACTIVOS FIJOS "
    
    prRs.MoveFirst
    While Not prRs.EOF
        i = i + 1
        
        If i = 2 Then
            xlHoja1.Cells(i + 1, 1) = "COD CTIVO"
            xlHoja1.Cells(i + 1, 2) = "COD. PATRIM."
            xlHoja1.Cells(i + 1, 3) = "F. ADQUIS"
            xlHoja1.Cells(i + 1, 4) = "SALDO ACT HIST"
            xlHoja1.Cells(i + 1, 5) = "SALDO AJUS AÑO ANT"
            xlHoja1.Cells(i + 1, 6) = "COM DEL AÑO"
            xlHoja1.Cells(i + 1, 7) = "RET DEL AÑO"
            xlHoja1.Cells(i + 1, 8) = "FAC DE AJUS"
            xlHoja1.Cells(i + 1, 9) = "REEXP VAL AJUS ANT"
            xlHoja1.Cells(i + 1, 10) = "COM AJUS"
            xlHoja1.Cells(i + 1, 11) = "RET AJUS"
            xlHoja1.Cells(i + 1, 12) = "VAL ACT AJUS"
            xlHoja1.Cells(i + 1, 13) = "DEP ACUM AÑO ANT"
            xlHoja1.Cells(i + 1, 14) = "REEXP DEP AJUS AÑO ANT"
            xlHoja1.Cells(i + 1, 15) = "DEP AJUS EJER MES ANT"
            xlHoja1.Cells(i + 1, 16) = "DEP AJUST DEL MES"
            xlHoja1.Cells(i + 1, 17) = "TOT DEP DEL EJER AJSU"
            xlHoja1.Cells(i + 1, 18) = "DEP AJUS ACUM RET"
            xlHoja1.Cells(i + 1, 19) = "DEP AJUS ACUM TOTAL"
            xlHoja1.Cells(i + 1, 20) = "VALOR EN LIBROS AJUS"
            
            i = i + 1

            xlHoja1.Cells(i + 1, 1) = prRs!Codigo & "-" & prRs!Serie
            xlHoja1.Cells(i + 1, 2) = Format(i - 2, "00000")
            xlHoja1.Cells(i + 1, 3) = "'" & Format(prRs!F_Compra, gsFormatoFechaView)
            xlHoja1.Cells(i + 1, 4) = Format(prRs!Valor, "#,##0.00")
            xlHoja1.Cells(i + 1, 5) = IIf(Year(prRs!F_Compra) = Me.mskAnio, 0, Format(prRs!Valor_Ajustado, "#,##0.00"))
            xlHoja1.Cells(i + 1, 6) = Format(IIf(Year(prRs!F_Compra) = Me.mskAnio, prRs!Valor, 0), "#,##0.00")
            
            If prRs!Baja = "False" Or Not IsNumeric(prRs!F_Baja) Then
                xlHoja1.Cells(i + 1, 7) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(i + 1, 7) = Format(prRs!Valor, "#,##0.00")
                Else
                    xlHoja1.Cells(i + 1, 7) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(i + 1, 8) = Format(prRs!F_Ajuste, "#,##0.000")
            xlHoja1.Cells(i + 1, 9) = Format(IIf(Year(prRs!F_Compra) = Me.mskAnio, 0, prRs!Valor_Ajustado) * prRs!F_Ajuste, "#,##0.00")
            xlHoja1.Cells(i + 1, 10) = Format(IIf(Year(prRs!F_Compra) = Year(gdFecSis), prRs!Valor, 0) * prRs!F_Ajuste, "#,##0.00")
            
            If prRs!Baja = "False" Or Not IsNumeric(prRs!F_Baja) Then
                xlHoja1.Cells(i + 1, 11) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(i + 1, 11) = Format(prRs!Valor * prRs!F_Ajuste, "#,##0.00")
                Else
                    xlHoja1.Cells(i + 1, 11) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(i + 1, 12) = Format(xlHoja1.Cells(i + 1, 9) + xlHoja1.Cells(i + 1, 10) - xlHoja1.Cells(i + 1, 11), "#,##0.00")
            
            xlHoja1.Cells(i + 1, 13) = Format(IIf(Year(prRs!F_Compra) >= Me.mskAnio, 0, prRs!Valor_Ajustado), "#,##0.00")
            xlHoja1.Cells(i + 1, 14) = Format(IIf(Year(prRs!F_Compra) >= Me.mskAnio, 0, prRs!Valor_Ajustado) * prRs!F_Ajuste, "#,##0.00")
            xlHoja1.Cells(i + 1, 15) = IIf(prRs!Baja = "False", Format(prRs!Dep_A_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(i + 1, 16) = Format(prRs!Dep_A_Mes, "#,##0.00")
            xlHoja1.Cells(i + 1, 17) = Format(xlHoja1.Cells(i + 1, 15) + xlHoja1.Cells(i + 1, 16), "#,##0.00")
            xlHoja1.Cells(i + 1, 18) = IIf(prRs!Baja <> "False", Format(prRs!Dep_A_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(i + 1, 19) = Format(xlHoja1.Cells(i + 1, 14) + xlHoja1.Cells(i + 1, 17) - xlHoja1.Cells(i + 1, 18), "#,##0.00")
            xlHoja1.Cells(i + 1, 20) = Format(xlHoja1.Cells(i + 1, 12) - xlHoja1.Cells(i + 1, 19), "#,##0.00")
        Else
            xlHoja1.Cells(i + 1, 1) = prRs!Codigo & "-" & prRs!Serie
            xlHoja1.Cells(i + 1, 2) = Format(i - 2, "00000")
            xlHoja1.Cells(i + 1, 3) = "'" & Format(prRs!F_Compra, gsFormatoFechaView)
            xlHoja1.Cells(i + 1, 4) = Format(prRs!Valor, "#,##0.00")
            xlHoja1.Cells(i + 1, 5) = IIf(Year(prRs!F_Compra) = Me.mskAnio, 0, Format(prRs!Valor_Ajustado, "#,##0.00"))
            xlHoja1.Cells(i + 1, 6) = Format(IIf(Year(prRs!F_Compra) = Year(gdFecSis), prRs!Valor, 0), "#,##0.00")
            
            If prRs!Baja = "False" Or Not IsNumeric(prRs!F_Baja) Then
                xlHoja1.Cells(i + 1, 7) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(i + 1, 7) = Format(prRs!Valor, "#,##0.00")
                Else
                    xlHoja1.Cells(i + 1, 7) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(i + 1, 8) = Format(prRs!F_Ajuste, "#,##0.000")
            xlHoja1.Cells(i + 1, 9) = Format(IIf(Year(prRs!F_Compra) = Me.mskAnio, 0, prRs!Valor_Ajustado) * prRs!F_Ajuste, "#,##0.00")
            xlHoja1.Cells(i + 1, 10) = Format(IIf(Year(prRs!F_Compra) = Year(gdFecSis), prRs!Valor, 0) * prRs!F_Ajuste, "#,##0.00")
            
            If prRs!Baja = "False" Or Not IsNumeric(prRs!F_Baja) Then
                xlHoja1.Cells(i + 1, 11) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(i + 1, 11) = Format(prRs!Valor * prRs!F_Ajuste, "#,##0.00")
                Else
                    xlHoja1.Cells(i + 1, 11) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(i + 1, 12) = Format(xlHoja1.Cells(i + 1, 9) + xlHoja1.Cells(i + 1, 10) - xlHoja1.Cells(i + 1, 11), "#,##0.00")
            
            xlHoja1.Cells(i + 1, 13) = Format(IIf(Year(prRs!F_Compra) >= Me.mskAnio, 0, prRs!Valor_Ajustado), "#,##0.00")
            xlHoja1.Cells(i + 1, 14) = Format(IIf(Year(prRs!F_Compra) >= Me.mskAnio, 0, prRs!Valor_Ajustado) * prRs!F_Ajuste, "#,##0.00")
            xlHoja1.Cells(i + 1, 15) = IIf(prRs!Baja = "False", Format(prRs!Dep_A_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(i + 1, 16) = Format(prRs!Dep_A_Mes, "#,##0.00")
            xlHoja1.Cells(i + 1, 17) = Format(xlHoja1.Cells(i + 1, 15) + xlHoja1.Cells(i + 1, 16), "#,##0.00")
            xlHoja1.Cells(i + 1, 18) = IIf(prRs!Baja <> "False", Format(prRs!Dep_A_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(i + 1, 19) = Format(xlHoja1.Cells(i + 1, 14) + xlHoja1.Cells(i + 1, 17) - xlHoja1.Cells(i + 1, 18), "#,##0.00")
            xlHoja1.Cells(i + 1, 20) = Format(-xlHoja1.Cells(i + 1, 12) + xlHoja1.Cells(i + 1, 19), "#,##0.00")

        End If
        prRs.MoveNext
    Wend
        
    xlHoja1.Select
    xlHoja1.Columns.AutoFit
    
    xlHoja1.Range("A1:T3").Select
    xlHoja1.Range("A1:T3").Font.Bold = True

    With xlHoja1.Range("A2:L2").Interior
        .ColorIndex = 36
        .Pattern = xlSolid
    End With
    With xlHoja1.Range("M2:T2").Interior
        .ColorIndex = 35
        .Pattern = xlSolid
    End With
    
    xlHoja1.Range("A2:T" & Trim(Str(i + 1))).Select
    xlHoja1.Range("A2:T" & Trim(Str(i + 1))).Borders(xlDiagonalDown).LineStyle = xlNone
    xlHoja1.Range("A2:T" & Trim(Str(i + 1))).Borders(xlDiagonalUp).LineStyle = xlNone
    
    With xlHoja1.Range("A2:T" & Trim(Str(i + 1))).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:T" & Trim(Str(i + 1))).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:T" & Trim(Str(i + 1))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:T" & Trim(Str(i + 1))).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:T" & Trim(Str(i + 1))).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:T" & Trim(Str(i + 1))).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End Sub


'Private Sub GetResumen()
'    Dim lnI As Long
'    Dim lnAnioAnt As Integer
'    Dim lnSumaVH As Currency
'    Dim lnSumaVA As Currency
'    Dim lnSumaVHM As Currency
'    Dim lnSumaVAM As Currency
'    Dim lnSumaVHMM As Currency
'    Dim lnSumaVAMM As Currency
'    Dim lnSumaVAjuste As Currency
'    Dim lnSumaVHMA As Currency
'    Dim lnSumaVAMA As Currency
'
'    flexRes.Clear
'    flexRes.Rows = 2
'    flexRes.FormaCabecera
'
'
'    lnSumaVH = 0
'    lnSumaVA = 0
'    lnSumaVHM = 0
'    lnSumaVAM = 0
'    lnSumaVHMM = 0
'    lnSumaVAMM = 0
'    lnSumaVAjuste = 0
'    lnSumaVHMA = 0
'    lnSumaVAMA = 0
'
'
'    For lnI = 1 To Me.Flex.Rows - 1
'        If lnAnioAnt <> 0 And lnAnioAnt <> Year(Flex.TextMatrix(lnI, 23)) And Year(Flex.TextMatrix(lnI, 23)) > 1998 Then
'            flexRes.AdicionaFila
'            flexRes.TextMatrix(flexRes.Rows - 1, 1) = lnAnioAnt
'            flexRes.TextMatrix(flexRes.Rows - 1, 2) = Format(lnSumaVH, "#,##0.00")
'            flexRes.TextMatrix(flexRes.Rows - 1, 3) = Format(lnSumaVA, "#,##0.00")
'            flexRes.TextMatrix(flexRes.Rows - 1, 4) = Format(lnSumaVHM, "#,##0.00")
'            flexRes.TextMatrix(flexRes.Rows - 1, 5) = Format(lnSumaVAM, "#,##0.00")
'            flexRes.TextMatrix(flexRes.Rows - 1, 6) = Format(lnSumaVHMM, "#,##0.00")
'            flexRes.TextMatrix(flexRes.Rows - 1, 7) = Format(lnSumaVAMM, "#,##0.00")
'            flexRes.TextMatrix(flexRes.Rows - 1, 8) = Format(lnSumaVAjuste, "#,##0.00")
'            flexRes.TextMatrix(flexRes.Rows - 1, 9) = Format(lnSumaVHMA, "#,##0.00")
'            flexRes.TextMatrix(flexRes.Rows - 1, 10) = Format(lnSumaVAMA, "#,##0.00")
'
'            lnSumaVH = 0
'            lnSumaVA = 0
'            lnSumaVHM = 0
'            lnSumaVAM = 0
'            lnSumaVHMM = 0
'            lnSumaVAMM = 0
'            lnSumaVAjuste = 0
'            lnSumaVHMA = 0
'            lnSumaVAMA = 0
'        End If
'
'        lnSumaVH = lnSumaVH + Flex.TextMatrix(lnI, 4)
'        If IsNumeric(Flex.TextMatrix(lnI, 6)) Then lnSumaVA = lnSumaVA + Flex.TextMatrix(lnI, 6)
'        lnSumaVHM = lnSumaVHM + Flex.TextMatrix(lnI, 9)
'        If IsNumeric(Flex.TextMatrix(lnI, 10)) Then lnSumaVAM = lnSumaVAM + Flex.TextMatrix(lnI, 10)
'        lnSumaVHMM = lnSumaVHMM + Flex.TextMatrix(lnI, 14)
'        If IsNumeric(Flex.TextMatrix(lnI, 15)) Then lnSumaVAMM = lnSumaVAMM + Flex.TextMatrix(lnI, 15)
'        If IsNumeric(Flex.TextMatrix(lnI, 16)) Then lnSumaVAjuste = lnSumaVAjuste + Flex.TextMatrix(lnI, 16)
'        lnSumaVHMA = lnSumaVHMA + Flex.TextMatrix(lnI, 26)
'        lnSumaVAMA = lnSumaVAMA + Flex.TextMatrix(lnI, 27)
'
'        lnAnioAnt = Year(Flex.TextMatrix(lnI, 23))
'    Next lnI
'
'    flexRes.AdicionaFila
'    flexRes.TextMatrix(flexRes.Rows - 1, 1) = lnAnioAnt
'    flexRes.TextMatrix(flexRes.Rows - 1, 2) = Format(lnSumaVH, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 3) = Format(lnSumaVA, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 4) = Format(lnSumaVHM, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 5) = Format(lnSumaVAM, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 6) = Format(lnSumaVHMM, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 7) = Format(lnSumaVAMM, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 8) = Format(lnSumaVAjuste, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 9) = Format(lnSumaVHMA, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 10) = Format(lnSumaVAMA, "#,##0.00")
'
'    lnSumaVH = 0
'    lnSumaVA = 0
'    lnSumaVHM = 0
'    lnSumaVAM = 0
'    lnSumaVHMM = 0
'    lnSumaVAMM = 0
'    lnSumaVAjuste = 0
'    lnSumaVHMA = 0
'    lnSumaVAMA = 0
'
'    For lnI = 1 To Me.flexRes.Rows - 1
'        lnSumaVH = lnSumaVH + flexRes.TextMatrix(lnI, 2)
'        lnSumaVA = lnSumaVA + flexRes.TextMatrix(lnI, 3)
'        lnSumaVHM = lnSumaVHM + flexRes.TextMatrix(lnI, 4)
'        lnSumaVAM = lnSumaVAM + flexRes.TextMatrix(lnI, 5)
'        lnSumaVHMM = lnSumaVHMM + flexRes.TextMatrix(lnI, 6)
'        lnSumaVAMM = lnSumaVAMM + flexRes.TextMatrix(lnI, 7)
'        lnSumaVAjuste = lnSumaVAjuste + flexRes.TextMatrix(lnI, 8)
'        lnSumaVHMA = lnSumaVHMA + flexRes.TextMatrix(lnI, 9)
'        lnSumaVAMA = lnSumaVAMA + flexRes.TextMatrix(lnI, 10)
'    Next lnI
'
'    flexRes.AdicionaFila
'    flexRes.TextMatrix(flexRes.Rows - 1, 1) = "TOTAL"
'    flexRes.TextMatrix(flexRes.Rows - 1, 2) = Format(lnSumaVH, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 3) = Format(lnSumaVA, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 4) = Format(lnSumaVHM, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 5) = Format(lnSumaVAM, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 6) = Format(lnSumaVHMM, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 7) = Format(lnSumaVAMM, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 8) = Format(lnSumaVAjuste, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 9) = Format(lnSumaVHMA, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 10) = Format(lnSumaVAMA, "#,##0.00")
'End Sub
'
'Private Sub GeneraReporteJoyasAdjud(prRs As ADODB.Recordset)
'    Dim i As Integer
'    Dim K As Integer
'    Dim j As Integer
'    Dim nFila As Integer
'    Dim nIni  As Integer
'    Dim lNegativo As Boolean
'    Dim sConec As String
'    Dim lsSuma As String
'    Dim sTipoGara As String
'    Dim sTipoCred As String
'    Dim lnAcum As Variant 'Currency
'    Dim lnSer As Currency
'
'    i = -1
'    prRs.MoveFirst
'    i = i + 1
'    For j = 0 To prRs.Fields.Count - 1
'        xlHoja1.Cells(i + 1, j + 1) = prRs.Fields(j).Name
'    Next j
'
'    While Not prRs.EOF
'        i = i + 1
'        For j = 0 To prRs.Fields.Count - 1
'            xlHoja1.Cells(i + 1, j + 1) = prRs.Fields(j)
'        Next j
'        prRs.MoveNext
'    Wend
'
'End Sub

Private Sub Text1_GotFocus()
fEnfoque Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(Text1, KeyAscii, 14, 5)
End Sub

Private Sub txtTipCamCompraSBS_GotFocus()
fEnfoque txtTipCamCompraSBS
End Sub

Private Sub txtTipCamCompraSBS_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTipCamCompraSBS, KeyAscii, 14, 5)
End Sub

Private Sub llenagrid()
    Dim oDepo As DAgencia
    Set oDepo = New DAgencia
    Dim ldFecha As Date, i As Integer
    Dim rs1 As ADODB.Recordset
    Dim rs9 As ADODB.Recordset
    
    Set rs9 = New ADODB.Recordset
    
    FlexEdit1.Clear
    ldFecha = CDate("01/" & Format(Trim(Right(Me.cmbMes.Text, 5)), "00") & "/" & Me.mskAnio.Text)
    
    FlexEdit1.rsFlex = oDepo.DistribuyePoliSeguPatri(Format(ldFecha, "yyyymmdd"), Val(Trim(Right(Me.cboTpo.Text, 3))), txtTipCamCompraSBS.Text, Text1.Text)
    
End Sub
