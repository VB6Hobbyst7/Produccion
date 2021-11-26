VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRegularizaSobFal 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   Icon            =   "frmRegularizaSobFal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPendientes 
      Appearance      =   0  'Flat
      Caption         =   "Pendientes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2535
      Left            =   60
      TabIndex        =   12
      Top             =   2115
      Width           =   9150
      Begin MSComctlLib.ListView lvwMov 
         Height          =   2130
         Left            =   135
         TabIndex        =   13
         Top             =   270
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3757
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin RichTextLib.RichTextBox R 
      Height          =   150
      Left            =   5730
      TabIndex        =   7
      Top             =   4875
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   265
      _Version        =   393217
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmRegularizaSobFal.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   6105
      TabIndex        =   6
      Top             =   4725
      Width           =   1020
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   7155
      TabIndex        =   0
      Top             =   4725
      Width           =   1020
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8190
      TabIndex        =   1
      Top             =   4725
      Width           =   1020
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Datos para regularizacón"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2130
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   9150
      Begin VB.CheckBox chkIngreso 
         Appearance      =   0  'Flat
         Caption         =   "Enviar el sobrante al Ingreso"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3360
         TabIndex        =   16
         Top             =   1170
         Visible         =   0   'False
         Width           =   3705
      End
      Begin VB.TextBox txtDoc 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1275
         MaxLength       =   15
         TabIndex        =   14
         Top             =   1140
         Width           =   1875
      End
      Begin VB.TextBox txtDesc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1275
         MaxLength       =   200
         TabIndex        =   9
         Top             =   1545
         Width           =   7740
      End
      Begin MSComctlLib.ListView lvwCtas 
         Height          =   795
         Left            =   90
         TabIndex        =   10
         Top             =   270
         Width           =   8940
         _ExtentX        =   15769
         _ExtentY        =   1402
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label lblDoc 
         Caption         =   "Documento  :"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   1207
         Width           =   855
      End
      Begin VB.Label lblDesc 
         Caption         =   "Descripción  :"
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   1545
         Width           =   1155
      End
   End
   Begin VB.Label lblTotDolG 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   4245
      TabIndex        =   5
      Top             =   4755
      Width           =   1440
   End
   Begin VB.Label lblTotDol 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Dolares : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2910
      TabIndex        =   4
      Top             =   4755
      Width           =   2280
   End
   Begin VB.Label lblTotSolG 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1395
      TabIndex        =   3
      Top             =   4755
      Width           =   1440
   End
   Begin VB.Label lblTotSol 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Soles :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   75
      TabIndex        =   2
      Top             =   4755
      Width           =   2280
   End
End
Attribute VB_Name = "frmRegularizaSobFal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsCtaSol As String
Dim lsCtaDol As String
Dim lsCaption As String
Dim lnOpeCod As COMDConstantes.CaptacOperacion 'MAVM 20120105

Public Sub Ini(psCaption As String, Optional ByVal pnOpeCod As Long) 'MAVM 20120105 Se agrego: pnOpeCod***
    lsCaption = psCaption
    lnOpeCod = pnOpeCod
    Me.Show 1
End Sub

Private Sub cmdGrabar_Click()
    'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
    'fin Comprobacion si es RFIII
    
    'Dim lnI As Integer
    Dim lnI As Long
    Dim lnSaldo As Double
    Dim lsHoy As String
    Dim rs As ADODB.Recordset
    Dim nMonto As Double
    Dim sCajero As String, sOperacion As String
    Dim nmoneda As COMDConstantes.Moneda
    
    Set rs = New ADODB.Recordset
    Dim lnNumMov As Long
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim oMov As COMDMov.DCOMMov  'DMov
    Set oMov = New COMDMov.DCOMMov
    Dim oCaj As COMNCajaGeneral.NCOMCajero 'PASI20150925 ERS0692015
    Set oCaj = New COMNCajaGeneral.NCOMCajero
    
    Dim lsMovNro As String
    Dim lnMovNro As Long, nMovNroRef As Long
    
    Dim lsCadImp As String
    Dim nFicSal As String
    
    '*** PEAC 20120803
    Dim lbResultadoVisto As Boolean
    Dim sPersVistoCod  As String
    Dim sPersVistoCom As String
    Dim pnMovMro As Long
    Dim loVistoElectronico As frmVistoElectronico
    Set loVistoElectronico = New frmVistoElectronico
    
    
    If Me.txtDoc.Text = "" Then
        MsgBox "Debe ingresar un numero de documento de referencia.", vbInformation, "Aviso"
        Me.txtDoc.SetFocus
        Exit Sub
    ElseIf Me.txtDesc.Text = "" Then
        MsgBox "Debe ingresar una descripción.", vbInformation, "Aviso"
        Me.txtDoc.SetFocus
        Exit Sub
    End If
    
    
    '*** PEAC 20120803
    If gsOpeCod = "901015" Or gsOpeCod = "300200" Then
    
        lbResultadoVisto = loVistoElectronico.Inicio(3, gsOpeCod)
        If Not lbResultadoVisto Then
            Exit Sub
        End If
    End If
    '*** FIN PEAC ************************************************************
    
    
    '*** el codigo de operacio falta definir para aprobar credito por mientras se puso 999999
    lsMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    Dim oImp As COMNCaptaGenerales.NCOMCaptaImpresion  'NCapImpBoleta
    Set oImp = New COMNCaptaGenerales.NCOMCaptaImpresion

    For lnI = 1 To Me.lvwMov.ListItems.count
        If Me.lvwMov.ListItems(lnI).Checked Then
            nMonto = CDbl(lvwMov.ListItems(lnI).ListSubItems(3))
            nMovNroRef = CLng(lvwMov.ListItems(lnI).ListSubItems(5))
            nmoneda = CLng(lvwMov.ListItems(lnI).ListSubItems(7))
            sOperacion = Trim(lvwMov.ListItems(lnI).ListSubItems(2))
            sCajero = Trim(lvwMov.ListItems(lnI).ListSubItems(1))
            If nMonto > 0 Then

                If chkIngreso.value = 1 Then
                    oMov.InsertaMov lsMovNro, COMDConstSistema.gOpeHabCajIngRegulaSob, Me.txtDesc.Text
                Else
                    oMov.InsertaMov lsMovNro, COMDConstSistema.gOpeHabCajDevClienteRegulaSob, Me.txtDesc.Text
                End If

                lnMovNro = oMov.GetnMovNro(lsMovNro)
                oMov.InsertaMovOpeVarias lnMovNro, Trim(txtDoc.Text), Trim(txtDesc.Text), Abs(nMonto), nmoneda
                oMov.InsertaMovRef lnMovNro, nMovNroRef
                
                If chkIngreso.value = 1 Then
                    lsCadImp = oImp.ImprimeBoleta("ENV. INGRESO", "Envio Cta. Ing." & Me.txtDoc.Text, "", Trim(Str(Abs(nMonto))), sCajero, Right(gsCodAge, 2) & "000" & nmoneda & "000000", "", 0, "", "", 0, 0, , False, , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, False)
                Else
                    lsCadImp = oImp.ImprimeBoleta("DEV.SOBRANTE", "Egreso Efectivo" & Me.txtDoc.Text, "", Trim(Str(Abs(nMonto))), sCajero, Right(gsCodAge, 2) & "000" & nmoneda & "000000", "", 0, "", "", 0, 0, , False, , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, False)
                End If
                Set oImp = Nothing
                
                Do
                    nFicSal = FreeFile
                     Open sLpt For Output As nFicSal
                        Print #nFicSal, lsCadImp & Chr$(12)
                        Print #nFicSal, ""
                     Close #nFicSal
                Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, Me.Caption) = vbNo
            Else
                'PASI20150925 ERS0692015
                If gsCodUser = (oCaj.ObtieneUserxRegularizaFaltante(nMovNroRef)) Then
                    MsgBox "¡¡¡El usuario que generó el faltante no puede pagar este faltante!!!", vbInformation, "Aviso"
                    Exit Sub
                End If
                'end PASI
                
                'Modificado por GITU 23-07-2008
                'oMov.InsertaMov lsMovNro, gOpeHabCajIngEfectRegulaFalt, Me.TxtDesc.Text
                oMov.InsertaMov lsMovNro, gsOpeCod, Me.txtDesc.Text
                'Fin GITU
                lnMovNro = oMov.GetnMovNro(lsMovNro)
                oMov.InsertaMovOpeVarias lnMovNro, Trim(txtDoc.Text), Trim(txtDesc.Text), Abs(nMonto), nmoneda
                oMov.InsertaMovRef lnMovNro, nMovNroRef
                
                lsCadImp = oImp.ImprimeBoleta("PAGO FALTANTE", "Ingreso Efectivo " & Me.txtDoc.Text, "", Trim(Str(Abs(nMonto))), sCajero, Right(gsCodAge, 2) & "000" & nmoneda & "000000", "", 0, "", "", 0, 0, , False, , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, False)
                Set oImp = Nothing
                
                Do
                    nFicSal = FreeFile
                     Open sLpt For Output As nFicSal
                        Print #nFicSal, lsCadImp & Chr$(12)
                        Print #nFicSal, ""
                     Close #nFicSal
                     
                Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, Me.Caption) = vbNo
            End If
        End If
    Next lnI
    
    '*** PEAC 20120803
    loVistoElectronico.RegistraVistoElectronico (pnMovMro)
    '*** FIN PEAC
    
    Set clsCap = Nothing
    Set oMov = Nothing
    'todocompleta avmm
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
   Dim lsCadImp As String
   Dim oCajImp As COMNCajaGeneral.NCOMCajero
   Dim oPrevio As previo.clsprevio
   
   Set oCajImp = New COMNCajaGeneral.NCOMCajero
        lsCadImp = oCajImp.ImpreRegularizacion(gsNomAge, gcEmpresa, gdFecSis, lnOpeCod)
   Set oCajImp = Nothing
   
   Set oPrevio = New previo.clsprevio
    'ALPA 20100202**************************************
     'previo.Show lscadimp, "REGULARIZACION", True
     previo.Show lsCadImp, "REGULARIZACION", True, , gImpresora
   Set oPrevio = Nothing

End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim lAux As ListItem
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    'Dim lnI As Integer '*** PEAC 20121227 por desbordamiento de items ya que solo mostraba 32767
    Dim lnI As Long '*** PEAC 20121227
    
    Dim oConst As COMDConstSistema.NCOMConstSistema
    Set oConst = New COMDConstSistema.NCOMConstSistema
        
    Dim oCaj As COMNCajaGeneral.NCOMCajero
        
    Me.Caption = lsCaption
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
 
    Me.lvwCtas.ColumnHeaders.Clear
    Me.lvwCtas.ListItems.Clear
    Me.lvwCtas.View = lvwReport
    
    Me.lvwCtas.HideColumnHeaders = False
    Me.lvwCtas.ColumnHeaders.Add , , "Cuenta", 2500
    Me.lvwCtas.ColumnHeaders.Add , , "Descripción ", 3500
    
    Me.lvwMov.ColumnHeaders.Clear
    Me.lvwMov.ListItems.Clear
    Me.lvwMov.View = lvwReport
    
    Me.lvwMov.HideColumnHeaders = False
    Me.lvwMov.ColumnHeaders.Add , , "Usuario", 900
    Me.lvwMov.ColumnHeaders.Add , , "Nombre", 3400
    Me.lvwMov.ColumnHeaders.Add , , "Tipo", 1450
    Me.lvwMov.ColumnHeaders.Add , , "Monto", 1500, 1
    Me.lvwMov.ColumnHeaders.Add , , "Fecha", 1300
    Me.lvwMov.ColumnHeaders.Add , , "nNumTranRef", 0
    Me.lvwMov.ColumnHeaders.Add , , "cCodope", 0
    Me.lvwMov.ColumnHeaders.Add , , "Cuenta", 0
    
    lsCtaSol = oConst.LeeConstSistema(61)
    Set lAux = Me.lvwCtas.ListItems.Add(, , lsCtaSol)
    Me.lvwCtas.ListItems(1).Bold = True
    lAux.SubItems(1) = "Cuenta de Sobrante Nuevos Soles"

    lsCtaDol = oConst.LeeConstSistema(62)
    Set lAux = Me.lvwCtas.ListItems.Add(, , lsCtaDol)
    Me.lvwCtas.ListItems(2).ForeColor = &H8000&
    Me.lvwCtas.ListItems(2).Bold = True
    lAux.SubItems(1) = "Cuenta de Sobrante Dolares"
    Me.lvwCtas.ListItems(2).ListSubItems(1).ForeColor = &H8000&
    
    Dim nMonSobrante As Double, nMonFaltante As Double
    Dim oPar As COMNCaptaGenerales.NCOMCaptaDefinicion
    Set oPar = New COMNCaptaGenerales.NCOMCaptaDefinicion
    nMonSobrante = oPar.GetCapParametro(gMontoSobranteCajero)
    nMonFaltante = oPar.GetCapParametro(gMontoFaltanteCajero)
    Set oPar = Nothing
   
    Set oCaj = New COMNCajaGeneral.NCOMCajero
    Set rs = oCaj.ObtenerRegularizacion(nMonSobrante, nMonFaltante)
    Set oCaj = Nothing
    
    lnI = 1
    While Not rs.EOF
        
    'MAVM 20120105 *** Memo 3257
    If lnOpeCod = gOtrOpePagoFaltante Then
        If rs!nMontran < 0 Then
            Set lAux = Me.lvwMov.ListItems.Add(, , rs!cCodusu)
            Me.lvwMov.ListItems(lnI).Bold = True
                
            lAux.SubItems(1) = Trim(rs!cNomusu & "")
            lAux.SubItems(2) = "FALTANTE"
            lAux.SubItems(3) = Format(rs!nMontran, "#,##0.00")
            lAux.SubItems(4) = Format(rs!dFecTran, gsFormatoFechaView)
            lAux.SubItems(5) = rs!cNrotran
            lAux.SubItems(6) = rs!cCodOpe
            lAux.SubItems(7) = rs!cCodCta
                
            If rs!cCodCta = Moneda.gMonedaExtranjera Then
                Me.lvwMov.ListItems(lnI).ForeColor = &H8000&
                Me.lvwMov.ListItems(lnI).ListSubItems(1).ForeColor = &H8000&
                Me.lvwMov.ListItems(lnI).ListSubItems(2).ForeColor = &H8000&
                Me.lvwMov.ListItems(lnI).ListSubItems(3).ForeColor = &H8000&
                Me.lvwMov.ListItems(lnI).ListSubItems(4).ForeColor = &H8000&
            End If
                
            lnI = lnI + 1
        End If
        
    Else
        
        If rs!nMontran > 0 Then
            Set lAux = Me.lvwMov.ListItems.Add(, , rs!cCodusu)
            Me.lvwMov.ListItems(lnI).Bold = True
                
            lAux.SubItems(1) = Trim(rs!cNomusu & "")
            lAux.SubItems(2) = "SOBRANTE"
            lAux.SubItems(3) = Format(rs!nMontran, "#,##0.00")
            lAux.SubItems(4) = Format(rs!dFecTran, gsFormatoFechaView)
            lAux.SubItems(5) = rs!cNrotran
            lAux.SubItems(6) = rs!cCodOpe
            lAux.SubItems(7) = rs!cCodCta
                
            If rs!cCodCta = Moneda.gMonedaExtranjera Then
                Me.lvwMov.ListItems(lnI).ForeColor = &H8000&
                Me.lvwMov.ListItems(lnI).ListSubItems(1).ForeColor = &H8000&
                Me.lvwMov.ListItems(lnI).ListSubItems(2).ForeColor = &H8000&
                Me.lvwMov.ListItems(lnI).ListSubItems(3).ForeColor = &H8000&
                Me.lvwMov.ListItems(lnI).ListSubItems(4).ForeColor = &H8000&
            End If
                
            lnI = lnI + 1
        End If
            
    End If

    rs.MoveNext
        
    'Set lAux = Me.lvwMov.ListItems.Add(, , rs!cCodUsu)
    'Me.lvwMov.ListItems(lnI).Bold = True
    
    'lAux.SubItems(1) = Trim(rs!cNomusu & "")
    
    'If rs!nMonTran > 0 Then
    '    lAux.SubItems(2) = "SOBRANTE"
    'Else
    '    lAux.SubItems(2) = "FALTANTE"
    'End If
    
    'lAux.SubItems(3) = Format(rs!nMonTran, "#,##0.00")
    'lAux.SubItems(4) = Format(rs!dFecTran, gsFormatoFechaView)
    'lAux.SubItems(5) = rs!cNroTran
    'lAux.SubItems(6) = rs!cCodOpe
    'lAux.SubItems(7) = rs!cCodCta
    
    'If rs!cCodCta = Moneda.gMonedaExtranjera Then
    '    Me.lvwMov.ListItems(lnI).ForeColor = &H8000&
    '    Me.lvwMov.ListItems(lnI).ListSubItems(1).ForeColor = &H8000&
    '    Me.lvwMov.ListItems(lnI).ListSubItems(2).ForeColor = &H8000&
    '    Me.lvwMov.ListItems(lnI).ListSubItems(3).ForeColor = &H8000&
    '    Me.lvwMov.ListItems(lnI).ListSubItems(4).ForeColor = &H8000&
    'End If
    
    'lnI = lnI + 1
    'rs.MoveNext
    '***
    Wend
    Set oConst = Nothing
End Sub

Private Sub lvwMov_ItemCheck(ByVal iTem As MSComctlLib.ListItem)
    'Dim lnI As Integer
    Dim lnI As Long
    Dim lnSol As Currency
    Dim lnDol As Currency
    
    For lnI = 1 To Me.lvwMov.ListItems.count
        If Me.lvwMov.ListItems(lnI).Checked Then
            Me.lvwMov.ListItems(lnI).Checked = False
        End If
    Next lnI
    
    iTem.Checked = True
    
    lnSol = 0
    lnDol = 0
    For lnI = 1 To Me.lvwMov.ListItems.count
        If Me.lvwMov.ListItems(lnI).Checked Then
            If lvwMov.ListItems(lnI).ListSubItems(7) = "1" Then
                lnSol = lnSol + CCur(lvwMov.ListItems(lnI).SubItems(3))
            Else
                lnDol = lnDol + CCur(lvwMov.ListItems(lnI).SubItems(3))
            End If
        End If
    Next lnI
    
    Me.lblTotDolG = Format(lnDol, "#,##0.00")
    Me.lblTotSolG = Format(lnSol, "#,##0.00")
End Sub

Private Sub TxtDesc_GotFocus()
    txtDesc.SelStart = 0
    txtDesc.SelLength = 200
End Sub

Private Sub TxtDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.lvwMov.SetFocus
    End If
End Sub

Private Sub txtDoc_GotFocus()
    txtDoc.SelStart = 0
    txtDoc.SelLength = 50
End Sub

Private Sub txtDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtDesc.SetFocus
    End If
End Sub


'sobrante con abono a cuenta
'Private Sub cmdGrabar_Click()
'    Dim lnI As Integer
'    Dim lnSaldo As Double
'    Dim lsHoy As String
'    Dim rs As ADODB.Recordset
'    Dim nMonto As Double
'    Dim sCajero As String, sOperacion As String
'    Dim nmoneda As COMDConstantes.Moneda
'
'    Set rs = New ADODB.Recordset
'    Dim lnNumMov As Long
'    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
'    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
'    Dim oMov As COMDMov.DCOMMov  'DMov
'    Set oMov = New COMDMov.DCOMMov
'    Dim lsMovNro As String
'    Dim lnMovNro As Long, nMovNroRef As Long
'
'    Dim lsCadImp As String
'    Dim nFicSal As String
'
'    If Me.txtDoc.Text = "" Then
'        MsgBox "Debe ingresar un numero de documento de referencia.", vbInformation, "Aviso"
'        Me.txtDoc.SetFocus
'        Exit Sub
'    ElseIf Me.txtDesc.Text = "" Then
'        MsgBox "Debe ingresar una descripción.", vbInformation, "Aviso"
'        Me.txtDoc.SetFocus
'        Exit Sub
'    End If
'
'    lsMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'
'    For lnI = 1 To Me.lvwMov.ListItems.Count
'        If Me.lvwMov.ListItems(lnI).Checked Then
'            nMonto = CDbl(lvwMov.ListItems(lnI).ListSubItems(3))
'            nMovNroRef = CLng(lvwMov.ListItems(lnI).ListSubItems(5))
'            nmoneda = CLng(lvwMov.ListItems(lnI).ListSubItems(7))
'            sOperacion = Trim(lvwMov.ListItems(lnI).ListSubItems(2))
'            sCajero = Trim(lvwMov.ListItems(lnI).ListSubItems(1))
'            If nMonto > 0 Then
'                If nmoneda = gMonedaNacional Then
'                    lnSaldo = clsCap.CapAbonoCuentaAho(lsCtaSol, nMonto, gAhoDepSobCaja, lsMovNro, Me.txtDesc.Text & " - " & nMonto & " - " & sOperacion, , , , , , , , , gsNomAge, sLpt)
'                    lnMovNro = oMov.GetnMovNro(lsMovNro)
'                    oMov.InsertaMovRef lnMovNro, nMovNroRef
'                Else
'                    lnSaldo = clsCap.CapAbonoCuentaAho(lsCtaDol, nMonto, gAhoDepSobCaja, lsMovNro, Me.txtDesc.Text & " - " & nMonto & " - " & sOperacion, , , , , , , , , gsNomAge, sLpt)
'                    lnMovNro = oMov.GetnMovNro(lsMovNro)
'                    oMov.InsertaMovRef lnMovNro, nMovNroRef
'                End If
'            Else
'                oMov.InsertaMov lsMovNro, gOpeHabCajIngEfectRegulaFalt, Me.txtDesc.Text
'                lnMovNro = oMov.GetnMovNro(lsMovNro)
'                oMov.InsertaMovOpeVarias lnMovNro, Trim(txtDoc.Text), Trim(txtDesc.Text), Abs(nMonto), nmoneda
'                oMov.InsertaMovRef lnMovNro, nMovNroRef
'
'                Dim oImp As COMNCaptaGenerales.NCOMCaptaImpresion  'NCapImpBoleta
'                Set oImp = New COMNCaptaGenerales.NCOMCaptaImpresion
'                    lsCadImp = oImp.ImprimeBoleta("PAGO FALTANTE", "Ingreso Efectivo " & Me.txtDoc.Text, "", Trim(Str(Abs(nMonto))), sCajero, Right(gsCodAge, 2) & "000" & nmoneda & "000000", "", 0, "", "", 0, 0, , False, , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, False)
'                Set oImp = Nothing
'
'                Do
'                    nFicSal = FreeFile
'                     Open sLpt For Output As nFicSal
'                        Print #nFicSal, lsCadImp & Chr$(12)
'                        Print #nFicSal, ""
'                     Close #nFicSal
'
'                Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, Me.Caption) = vbNo
'            End If
'        End If
'    Next lnI
'    Set clsCap = Nothing
'    Set oMov = Nothing
'    'todocompleta avmm
'    'Unload Me
'End Sub

