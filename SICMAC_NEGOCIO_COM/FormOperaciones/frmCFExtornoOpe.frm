VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmCFExtornoOpe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carta Fianza - Extorno de Operaciones"
   ClientHeight    =   5610
   ClientLeft      =   735
   ClientTop       =   1950
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCFExtornoOpe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmMotExtorno 
      Caption         =   "Motivos del Extorno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2700
      Left            =   3360
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   2845
      Begin VB.CommandButton cmdExtContinuar 
         Caption         =   "&Continuar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   14
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtDetExtorno 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   240
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox cmbMotivos 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmCFExtornoOpe.frx":030A
         Left            =   240
         List            =   "frmCFExtornoOpe.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalles del Extorno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblExtCmb 
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CheckBox chkOpeCMACLlamada 
      Caption         =   "Operaciones LLamada CMAC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7760
      TabIndex        =   6
      Top             =   0
      Width           =   1695
      Begin VB.CommandButton cmdExtorno 
         Caption         =   "&Extornar"
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
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1245
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
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
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1245
      End
   End
   Begin VB.Frame fraBuscar 
      Caption         =   "Buscar Por"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   180
      TabIndex        =   3
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton opt 
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   1
         Top             =   720
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.OptionButton opt 
         Caption         =   "Nro Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   0
         Top             =   360
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
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
      Left            =   6360
      TabIndex        =   2
      Top             =   240
      Width           =   1005
   End
   Begin MSComctlLib.ListView lstExtorno 
      Height          =   3540
      Left            =   180
      TabIndex        =   4
      Top             =   1320
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   6244
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   17
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCFExtornoOpe.frx":030E
            Key             =   "Cuenta"
         EndProperty
      EndProperty
   End
   Begin SICMACT.ActXCodCta AXCodCta 
      Height          =   390
      Left            =   2340
      TabIndex        =   10
      Top             =   240
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   688
      Texto           =   "Cta Fianza"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin VB.Label lblMensaje 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Por favor, sea muy cuidadoso(a) al utilizar los EXTORNOS. No hay forma de volver a realizar el proceso del Extorno."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   3090
      TabIndex        =   5
      Top             =   4920
      Width           =   6360
   End
End
Attribute VB_Name = "frmCFExtornoOpe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************
'EXTORNO DE OPERACIONES DE CARTA FIANZA
'Archivo:  frmCFExtornoOpe.frm
'LAYG   :  25/09/2002.
'Resumen:  Extornos de Operaciones de Carta fianza

Option Explicit
Dim vPosExtorno As Integer

Dim fsFechaTransac As String
Dim fnOperacion As ColocPOperaciones
Dim fsListaOpe As String
Dim lsOpeExt As String

Public Sub Inicio(ByVal pnOperacion As ColocCFOperaciones, ByVal PsOperacion As String)
fnOperacion = pnOperacion
Me.Caption = "Carta Fianza - Extornos - " & PsOperacion
lsOpeExt = pnOperacion
Select Case fnOperacion
    Case gColCFOpeExtComis
        'fsListaOpe = "('" & gColCFOpeComisEfe & "') "
        fsListaOpe = "('" & gColCFOpeComisEfe & "','" & gComiCredEmisionRenovCF & "') " 'JUEZ 20160226
End Select
Me.chkOpeCMACLlamada.Enabled = False
Me.Show 1
End Sub


Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdBuscar.SetFocus
End Sub

Private Sub cmdBuscar_Click()

Dim lrBusca As New ADODB.Recordset
Dim loValCta As COMNCartaFianza.NCOMCartaFianzaValida
Dim lsmensaje As String

'On Error GoTo ControlError

    'Valida Contrato
    'Limpiar
    lstExtorno.ListItems.Clear
    Set lrBusca = New ADODB.Recordset
    Set loValCta = New COMNCartaFianza.NCOMCartaFianzaValida
        If Me.chkOpeCMACLlamada.value = 1 Then ' Operaciones LLamada CMAC
            'Set lrBusca = loValContrato.nBuscaOperacionesCredPigParaExtornoLLamadaCMAC(fsFechaTransac)
        Else
            If Me.opt(0).value = True Then ' Busca por Codigo
                Set lrBusca = loValCta.nBuscaOperacionesCFParaExtorno(fsFechaTransac, fsListaOpe, Me.AXCodCta.NroCuenta)
            Else
                Set lrBusca = loValCta.nBuscaOperacionesCFParaExtorno(fsFechaTransac, fsListaOpe)
            End If
        End If
    Set loValCta = Nothing
    
    If lrBusca Is Nothing Then ' Hubo un Error
        Set lrBusca = Nothing
        Exit Sub
    End If
    If lrBusca.BOF And lrBusca.EOF Then
        MsgBox "No Existen Operaciones para EXTORNAR", vbInformation, "Aviso"
        Exit Sub
    End If
    
    lstExtorno.ListItems.Clear
    
    Call LLenaLista(lrBusca)
    
    Set lrBusca = Nothing
        
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & err.Number & " " & err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub
'*****CTI3 (ferimoro) 03102018
Sub OcultarControles()
'******CTI3 (ferimoro) 27092018
frmMotExtorno.Visible = False
Me.cmbMotivos.ListIndex = -1
Me.txtDetExtorno.Text = ""
fraBuscar.Enabled = True
AXCodCta.Enabled = True
cmdBuscar.Enabled = True
cmdExtorno.Enabled = False
'******************************
End Sub
Private Sub cmdExtContinuar_Click()
'On Error GoTo ControlError

Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabarExt As COMNCartaFianza.NCOMCartaFianza
Dim loImprime As COMNCartaFianza.NCOMCartaFianzaImpre
Dim loPrevio As previo.clsprevio
Dim lsCadImprimir As String

Dim lsMovNro As String
Dim lsFechaHoraGrab As String

Dim lsNroCta As String
Dim lsOperacion As String
Dim lnMovNroAExt As Long
Dim lnSaldo As Currency
Dim lnMonto As Currency
Dim Fecha As String

'*** PEAC 20081002
Dim lbResultadoVisto As Boolean
Dim sPersVistoCod  As String
Dim sPersVistoCom As String
Dim loVistoElectronico As frmVistoElectronico
Set loVistoElectronico = New frmVistoElectronico

' *** RIRO SEGUN TI-ERS108-2013 ***
    Dim nMovNroOperacion As Long
    nMovNroOperacion = 0
    If Not lstExtorno.SelectedItem Is Nothing Then
    nMovNroOperacion = CDbl(Val(lstExtorno.SelectedItem.ListSubItems(5)))
    End If
' *** FIN RIRO ***

'***CTI3 (FERIMORO)   02102018
Dim DatosExtorna(1) As String

If cmbMotivos.ListIndex = -1 Or Len(txtDetExtorno.Text) <= 0 Then
    MsgBox "Debe ingresar el motivo y/o detalle del Extorno", vbInformation, "Aviso"
    Exit Sub
End If

'***CTI3 (ferimoro)    02102018
frmMotExtorno.Visible = False
DatosExtorna(0) = cmbMotivos.Text
DatosExtorna(1) = txtDetExtorno.Text


'*** PEAC 20081001 - visto electronico ******************************************************
'*** en estos extornos de operaciones pedirá visto electrónico
'Comentado por JUEZ 20160304
'Select Case lsOpeExt
'     Case "149100"
'
'         lbResultadoVisto = loVistoElectronico.Inicio(3, lsOpeExt, , , nMovNroOperacion) 'RIRO SEGUN TI-ERS108-2013/ Se agrego parametro nMovNroOperacion
'         If Not lbResultadoVisto Then
'             Exit Sub
'         End If
'End Select

'*** FIN PEAC ************************************************************


If lstExtorno.SelectedItem.SubItems(2) <> "1" Then
    MsgBox " Debe Extornar el último movimiento de la Carta Fianza ", vbInformation, " Aviso "
    Call OcultarControles
    Exit Sub
Else
    If MsgBox(" Esta Ud seguro de Extornar dicha Operación ? ", vbQuestion + vbYesNo + vbDefaultButton2, " Aviso ") = vbNo Then
        Call OcultarControles
        Exit Sub
    Else
        'JUEZ 20160304 **********************************
        Select Case lsOpeExt
        Case gColCFOpeExtComis
            lbResultadoVisto = loVistoElectronico.Inicio(3, lsOpeExt, , , nMovNroOperacion)
            If Not lbResultadoVisto Then Call OcultarControles: Exit Sub
        End Select
        'END JUEZ ***************************************
        MsgBox " Prepare la impresora para imprimir " & vbCr & _
        " el recibo del Extorno", vbInformation, " Aviso "
    End If
End If
'*** Obtiene Datos de Operacion
lsNroCta = Trim(lstExtorno.SelectedItem)
lsOperacion = Right(lstExtorno.SelectedItem.ListSubItems(1), 6)
lnMovNroAExt = Val(lstExtorno.SelectedItem.ListSubItems(5))
'lnSaldo = CCur(lstExtorno.SelectedItem.ListSubItems(2))
lnMonto = CCur(lstExtorno.SelectedItem.ListSubItems(3))
Fecha = lstExtorno.SelectedItem.ListSubItems(4)

'*** Genera el Mov Nro
Set loContFunct = New COMNContabilidad.NCOMContFunciones
    lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
Set loContFunct = Nothing

lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
    
Set loGrabarExt = New COMNCartaFianza.NCOMCartaFianza
    
    Select Case lsOperacion
        '** Extornar Comision de CF
        'Case gColCFOpeComisEfe
        Case gColCFOpeComisEfe, gComiCredEmisionRenovCF 'JUEZ 20160304
            Call loGrabarExt.nCFExtornoComision(lsNroCta, lsFechaHoraGrab, _
                lsMovNro, lnMovNroAExt, lnMonto, False, DatosExtorna)
            ' *** Impresion
            Set loImprime = New COMNCartaFianza.NCOMCartaFianzaImpre
                lsCadImprimir = loImprime.nPrintReciboCFComisionExtorno(gsNomAge, lsFechaHoraGrab, lsNroCta, _
                       lnMonto, gsCodUser, "", gsCodCMAC, gImpresora, gbImpTMU)
            Set loImprime = Nothing
            Set loPrevio = New previo.clsprevio
                loPrevio.PrintSpool sLpt, lsCadImprimir, True, 22
                Do While True
                    If MsgBox("Desea ReImprimir Recibo ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                        loPrevio.PrintSpool sLpt, lsCadImprimir, True, 22
                    Else
                        Set loPrevio = Nothing
                        Exit Do
                    End If
                Loop
    End Select

    '*** PEAC 20081001
        loVistoElectronico.RegistraVistoElectronico (lnMovNroAExt)
    '*** FIN PEAC

Set loGrabarExt = Nothing

Call OcultarControles 'CTI3

Me.lstExtorno.ListItems.Clear
If lstExtorno.ListItems.count = 0 Then
    cmdExtorno.Enabled = False
End If
If Me.opt(3).value = True Then
    opt_KeyPress 3, 13
Else
    cmdBuscar_Click
End If

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
        
End Sub

Private Sub cmdExtorno_Click()
If lstExtorno.ListItems.count = 0 Then
    cmdExtorno.Enabled = False
    Exit Sub
End If

'******CTI3 (ferimoro) 27092018
 frmMotExtorno.Visible = True
 fraBuscar.Enabled = False
 AXCodCta.Enabled = False
 cmdBuscar.Enabled = False
 cmdExtorno.Enabled = False
 cmbMotivos.SetFocus
'******************************
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

'******CTI3 (ferimoro) 18102018
Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

Set oCons = New COMDConstantes.DCOMConstantes

Set R = oCons.ObtenerConstanteExtornoMotivo

Set oCons = Nothing
Call Llenar_Combo_MotivoExtorno(R, cmbMotivos)

End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    fsFechaTransac = Mid(Format$(gdFecSis, "dd/mm/yyyy"), 7, 4) & Mid(Format$(gdFecSis, "dd/mm/yyyy"), 4, 2) & Mid(Format$(gdFecSis, "dd/mm/yyyy"), 1, 2)
    Call CargaControles 'CTI3
    lstExtorno.ColumnHeaders.Add , , "NroCuenta", 2000
    lstExtorno.ColumnHeaders.Add , , "Operación", 2200
    lstExtorno.ColumnHeaders.Add , , "OpcExt.", 750, lvwColumnCenter
    lstExtorno.ColumnHeaders.Add , , "Monto", 1100, lvwColumnRight
    lstExtorno.ColumnHeaders.Add , , "Fecha de Movimiento", 1750, lvwColumnCenter
    lstExtorno.ColumnHeaders.Add , , "N°Tran", 800, lvwColumnCenter
    lstExtorno.ColumnHeaders.Add , , "Usuario", 800, lvwColumnCenter
    lstExtorno.View = lvwReport
    Limpiar
End Sub

Private Sub LLenaLista(myRs As Recordset)
Dim litmX As ListItem
Dim lsCtaCodAnterior As String

Do While Not myRs.EOF
    Set litmX = lstExtorno.ListItems.Add(, , myRs!cCtaCod, , "Cuenta")           'Nro de Cred Pig
        litmX.SubItems(1) = Mid(myRs!cOpeDesc, 1, 30) & Space(10) & myRs!cOpecod 'Operacion
        litmX.SubItems(3) = Format(myRs!nMonto, "#0.00")                         'Monto Operacion
        litmX.SubItems(4) = fgFechaHoraGrab(myRs!cMovNro)                        'Fecha/hora Operacion
        litmX.SubItems(5) = Str(myRs!nMovNro)                                    'Nro Movimiento(nMovNro)
        litmX.SubItems(6) = Mid(myRs!cMovNro, 22, 4)                             'Usuario
    
    If myRs!cCtaCod = lsCtaCodAnterior Then
        litmX.SubItems(2) = "0"
    Else
        litmX.SubItems(2) = "1"
    End If
    lsCtaCodAnterior = myRs!cCtaCod
    myRs.MoveNext
Loop

End Sub

'Valida el ListView lstExtorno
Private Sub lstExtorno_GotFocus()
If lstExtorno.ListItems.count >= 0 Then
   cmdExtorno.Enabled = True
End If
End Sub

Private Sub lstExtorno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     If lstExtorno.ListItems.count > 0 Then
        cmdExtorno.Enabled = True
        cmdExtorno.SetFocus
     End If
End If
End Sub

Private Sub opt_Click(index As Integer)
Limpiar

Select Case index
    Case 0
        AXCodCta.Visible = True
       
    Case 3
        AXCodCta.Visible = False
End Select
cmdBuscar.Visible = True
End Sub

Private Sub opt_KeyPress(index As Integer, KeyAscii As Integer)
Select Case index
    Case 0
        If KeyAscii = 13 Then
            AXCodCta.SetFocusCuenta
        End If
    Case 3
        If KeyAscii = 13 Then
            cmdBuscar.SetFocus
        End If
End Select
Me.Caption = "Carta Fianza - Extornos"
End Sub


'Inicializa variables
Private Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaCF
    lstExtorno.ListItems.Clear
End Sub

Private Sub txtDetExtorno_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
'If KeyAscii = 13 Then SendKeys "{TAB}": Exit Sub
End Sub
