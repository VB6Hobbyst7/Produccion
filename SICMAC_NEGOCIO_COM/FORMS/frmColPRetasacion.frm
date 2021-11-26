VERSION 5.00
Begin VB.Form frmColPRetasacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Retasación de Joyas"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   Icon            =   "frmColPRetasacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6255
      TabIndex        =   39
      Top             =   6555
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5055
      TabIndex        =   4
      Top             =   6555
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7470
      TabIndex        =   3
      Top             =   6555
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   6330
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   8430
      Begin VB.Frame fraKilataje 
         Caption         =   "Kilataje"
         Height          =   1440
         Left            =   6735
         TabIndex        =   12
         Top             =   2640
         Width           =   1560
         Begin VB.Label lbl21 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   495
            TabIndex        =   38
            Top             =   1125
            Width           =   945
         End
         Begin VB.Label lbl18 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   495
            TabIndex        =   37
            Top             =   840
            Width           =   945
         End
         Begin VB.Label lbl16 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   495
            TabIndex        =   36
            Top             =   525
            Width           =   945
         End
         Begin VB.Label lbl14 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   495
            TabIndex        =   35
            Top             =   195
            Width           =   945
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "21k"
            Height          =   195
            Left            =   75
            TabIndex        =   34
            Top             =   1185
            Width           =   270
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "18k"
            Height          =   195
            Left            =   75
            TabIndex        =   33
            Top             =   900
            Width           =   270
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "16k"
            Height          =   195
            Left            =   75
            TabIndex        =   32
            Top             =   585
            Width           =   270
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "14k"
            Height          =   195
            Left            =   75
            TabIndex        =   31
            Top             =   270
            Width           =   270
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1410
         Left            =   150
         TabIndex        =   11
         Top             =   2655
         Width           =   6555
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Oro Neto (gr.)"
            Height          =   195
            Left            =   75
            TabIndex        =   30
            Top             =   975
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Oro Bruto (gr.)"
            Height          =   195
            Left            =   75
            TabIndex        =   29
            Top             =   660
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Piezas"
            Height          =   195
            Left            =   75
            TabIndex        =   28
            Top             =   360
            Width           =   465
         End
         Begin VB.Label LblOroNeto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1095
            TabIndex        =   27
            Top             =   945
            Width           =   945
         End
         Begin VB.Label lblOroBru 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1095
            TabIndex        =   26
            Top             =   600
            Width           =   945
         End
         Begin VB.Label lblPiezas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1095
            TabIndex        =   25
            Top             =   270
            Width           =   945
         End
         Begin VB.Label lblSaldo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2865
            TabIndex        =   24
            Top             =   930
            Width           =   1140
         End
         Begin VB.Label lblPrestamo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2865
            TabIndex        =   23
            Top             =   585
            Width           =   1140
         End
         Begin VB.Label lblTasacion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2865
            TabIndex        =   22
            Top             =   255
            Width           =   1140
         End
         Begin VB.Label lblEstado 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   4755
            TabIndex        =   21
            Top             =   960
            Width           =   1725
         End
         Begin VB.Label lblFecVenc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   5340
            TabIndex        =   20
            Top             =   615
            Width           =   1140
         End
         Begin VB.Label lblFecPrestamo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   5340
            TabIndex        =   19
            Top             =   270
            Width           =   1140
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Left            =   4095
            TabIndex        =   18
            Top             =   990
            Width           =   495
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Fech. Venc."
            Height          =   195
            Left            =   4095
            TabIndex        =   17
            Top             =   660
            Width           =   870
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Fech. Préstamo"
            Height          =   195
            Left            =   4095
            TabIndex        =   16
            Top             =   360
            Width           =   1110
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Cap."
            Height          =   195
            Left            =   2070
            TabIndex        =   15
            Top             =   975
            Width           =   780
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Préstamo"
            Height          =   195
            Left            =   2100
            TabIndex        =   14
            Top             =   660
            Width           =   660
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Tasación"
            Height          =   195
            Left            =   2100
            TabIndex        =   13
            Top             =   360
            Width           =   660
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cliente"
         Height          =   1785
         Left            =   135
         TabIndex        =   7
         Top             =   870
         Width           =   8175
         Begin SICMACT.FlexEdit FECliente 
            Height          =   1275
            Left            =   105
            TabIndex        =   8
            Top             =   315
            Width           =   6750
            _extentx        =   11906
            _extenty        =   2249
            cols0           =   5
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "#-Cliente-Dirección-Nro DNI-Nro RUC"
            encabezadosanchos=   "0-3000-2500-1200-1200"
            font            =   "frmColPRetasacion.frx":030A
            font            =   "frmColPRetasacion.frx":0336
            font            =   "frmColPRetasacion.frx":0362
            font            =   "frmColPRetasacion.frx":038E
            font            =   "frmColPRetasacion.frx":03BA
            fontfixed       =   "frmColPRetasacion.frx":03E6
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1  'True
            columnasaeditar =   "X-X-X-X-X"
            listacontroles  =   "0-0-0-0-0"
            encabezadosalineacion=   "C-C-C-C-C"
            formatosedit    =   "0-0-0-0-0"
            textarray0      =   "#"
            lbformatocol    =   -1  'True
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
         Begin VB.Label lblTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   6930
            TabIndex        =   10
            Top             =   780
            Width           =   1140
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Tiipo de Contrato"
            Height          =   420
            Left            =   7095
            TabIndex        =   9
            Top             =   300
            Width           =   705
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fraPiezasDet 
         Caption         =   "Detalle de Piezas"
         Height          =   2055
         Left            =   165
         TabIndex        =   5
         Top             =   4125
         Width           =   8130
         Begin SICMACT.FlexEdit FEJoyas 
            Height          =   1695
            Left            =   105
            TabIndex        =   6
            Top             =   240
            Width           =   7890
            _extentx        =   13917
            _extenty        =   2990
            cols0           =   7
            highlight       =   1
            allowuserresizing=   2
            encabezadosnombres=   "Num-Pzas-Material-PBruto-PNeto-Tasac-Descripcion"
            encabezadosanchos=   "400-450-1030-650-650-700-4000"
            font            =   "frmColPRetasacion.frx":0414
            font            =   "frmColPRetasacion.frx":0440
            font            =   "frmColPRetasacion.frx":046C
            font            =   "frmColPRetasacion.frx":0498
            font            =   "frmColPRetasacion.frx":04C4
            fontfixed       =   "frmColPRetasacion.frx":04F0
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1  'True
            columnasaeditar =   "X-X-2-3-4-X-6"
            listacontroles  =   "0-0-3-0-0-0-0"
            encabezadosalineacion=   "C-R-L-R-R-R-L"
            formatosedit    =   "0-2-1-2-2-2-0"
            textarray0      =   "Num"
            lbeditarflex    =   -1  'True
            lbbuscaduplicadotext=   -1  'True
            colwidth0       =   405
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   345
         Left            =   4020
         Picture         =   "frmColPRetasacion.frx":051E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Buscar ..."
         Top             =   390
         Width           =   420
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   3615
         _extentx        =   6376
         _extenty        =   661
         texto           =   "Crédito"
         enabledcta      =   -1  'True
         enabledage      =   -1  'True
         prod            =   "305"
      End
   End
End
Attribute VB_Name = "frmColPRetasacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fsNroProceso As String
Dim fsRemateCadaAgencia As String
Dim fnDiasAtrasoAvisoRemate As Double

Private Sub CargaParametros()
Dim loParam As COMDColocPig.DCOMColPCalculos
Dim loConstSis As COMDConstSistema.NCOMConstSistema
Dim lnProcesoCadaAgencia As Integer

Set loParam = New COMDColocPig.DCOMColPCalculos
     fnDiasAtrasoAvisoRemate = Int(loParam.dObtieneColocParametro(gConsColPDiasAtrasoParaRemate))
    
Set loParam = Nothing
    
Set loConstSis = New COMDConstSistema.NCOMConstSistema
   
    lnProcesoCadaAgencia = loConstSis.LeeConstSistema(121)  ' gConstSistPigRemateCadaAg
    If lnProcesoCadaAgencia = 1 Then  ' En cada agencia
        fsRemateCadaAgencia = gsCodCMAC & gsCodAge
    Else
        fsRemateCadaAgencia = gsCodCMAC & "00"
    End If
Set loConstSis = Nothing

End Sub

Private Sub Limpiar()
    'AXCodCta.NroCuenta = ""
    Dim i As Integer
    For i = 1 To FECliente.Rows - 1
        FECliente.EliminaFila i
    Next i
    
    lblTipo.Caption = ""
    lblPiezas.Caption = ""
    lblOroBru.Caption = ""
    LblOroNeto.Caption = ""
    lblTasacion.Caption = ""
    lblPrestamo.Caption = ""
    lblSaldo.Caption = ""
    lblFecPrestamo.Caption = ""
    lblFecVenc.Caption = ""
    lblEstado.Caption = ""
    lbl14.Caption = ""
    lbl16.Caption = ""
    lbl18.Caption = ""
    lbl21.Caption = ""
   
    
    For i = FEJoyas.Rows - 1 To 2 Step -1
        FEJoyas.EliminaFila i
    Next i
    
   Set FEJoyas.Recordset = Nothing

 
    AXCodCta.Enabled = True
    AXCodCta.EnabledAge = True
    AXCodCta.EnabledCMAC = False
    AXCodCta.EnabledCta = True
    AXCodCta.EnabledProd = False
    
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Limpiar
    fraPiezasDet.Enabled = True

    If ValidaEntraRemate(AXCodCta.NroCuenta) = False Then
            MsgBox "Este contrato no entrará al Remate: " & fsNroProceso & vbCrLf & "No puede realizar esta operación.", vbOKOnly + vbInformation, App.Title
            Exit Sub
    End If

    Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
    Set loMuestraContrato = New COMDColocPig.DCOMColPContrato
    Dim lrCredPigJoyasDet As New ADODB.Recordset
    Dim i As Long
        i = 1
        
        MuestraDatosContrato AXCodCta.NroCuenta
        Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(AXCodCta.NroCuenta)
        
        While Not lrCredPigJoyasDet.EOF
         
            FEJoyas.TextMatrix(i, 0) = lrCredPigJoyasDet!nItem
            FEJoyas.TextMatrix(i, 1) = lrCredPigJoyasDet!npiezas
            FEJoyas.TextMatrix(i, 2) = lrCredPigJoyasDet!ckilataje
            FEJoyas.TextMatrix(i, 3) = lrCredPigJoyasDet!nPesoBruto
            FEJoyas.TextMatrix(i, 4) = lrCredPigJoyasDet!npesoneto
            FEJoyas.TextMatrix(i, 5) = lrCredPigJoyasDet!nvaltasac
            FEJoyas.TextMatrix(i, 6) = Trim(lrCredPigJoyasDet!cdescrip)
            If i < lrCredPigJoyasDet.RecordCount Then
                i = i + 1
                FEJoyas.AdicionaFila i
            End If
            
            lrCredPigJoyasDet.MoveNext
        Wend
       
    Set lrCredPigJoyasDet = Nothing
    Set loMuestraContrato = Nothing
     cmdGrabar.Enabled = True
       cmdCancelar.Enabled = True
     AXCodCta.Enabled = False
     
    
End If
End Sub

Private Sub cmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstado As String
Dim loPersContrato As COMDColocPig.DCOMColPContrato
Dim lrContratos As ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

Limpiar
AXCodCta.NroCuenta = ""
AXCodCta.CMAC = "108"
AXCodCta.Prod = "305"
Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPerscod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing



' Selecciona Estados
lsEstado = gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov

If Trim(lsPersCod) <> "" Then
    Set loPersContrato = New COMDColocPig.DCOMColPContrato
    Set lrContratos = loPersContrato.dObtieneCredPigDePersona(lsPersCod, lsEstado, Mid(gsCodAge, 4, 2))
    Set loPersContrato = Nothing
End If

Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrContratos)
    If loCuentas.sCtaCod <> "" Then
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        AXCodCta.SetFocusCuenta
    End If
Set loCuentas = Nothing

If ValidaEntraRemate(AXCodCta.NroCuenta) = False Then
            MsgBox "Este contrato no entrará al Remate: " & fsNroProceso & vbCrLf & "No puede realizar esta operación.", vbOKOnly + vbInformation, App.Title
            Exit Sub
End If

    Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
    Set loMuestraContrato = New COMDColocPig.DCOMColPContrato
    Dim lrCredPigJoyasDet As New ADODB.Recordset
    Dim i As Long
        i = 1
        
        MuestraDatosContrato AXCodCta.NroCuenta
        Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(AXCodCta.NroCuenta)
        
        While Not lrCredPigJoyasDet.EOF
         
            FEJoyas.TextMatrix(i, 0) = lrCredPigJoyasDet!nItem
            FEJoyas.TextMatrix(i, 1) = lrCredPigJoyasDet!npiezas
            FEJoyas.TextMatrix(i, 2) = lrCredPigJoyasDet!ckilataje
            FEJoyas.TextMatrix(i, 3) = lrCredPigJoyasDet!nPesoBruto
            FEJoyas.TextMatrix(i, 4) = lrCredPigJoyasDet!npesoneto
            FEJoyas.TextMatrix(i, 5) = lrCredPigJoyasDet!nvaltasac
            FEJoyas.TextMatrix(i, 6) = Trim(lrCredPigJoyasDet!cdescrip)
            If i < lrCredPigJoyasDet.RecordCount Then
                i = i + 1
                FEJoyas.AdicionaFila i
            End If
            
            lrCredPigJoyasDet.MoveNext
        
           Wend
           
    cmdGrabar.Enabled = True
    cmdCancelar.Enabled = True
    AXCodCta.Enabled = False
    
    Set lrCredPigJoyasDet = Nothing
    Set loMuestraContrato = Nothing


Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "


End Sub
Private Function ValidaEntraRemate(ByVal sCodCta As String) As Boolean


Dim psfecremate As String
Dim lrdatosrem As ADODB.Recordset
Dim bcolpRemate As Boolean

Dim nColP As COMNColoCPig.NCOMColPRecGar
Set nColP = New COMNColoCPig.NCOMColPRecGar
Set lrdatosrem = New ADODB.Recordset

Dim loParam As COMDColocPig.DCOMColPCalculos
Dim loConstSis As COMDConstSistema.NCOMConstSistema
Dim lnProcesoCadaAgencia As Integer
Dim lsUltRemate As String
Dim lsmensaje As String

Set loConstSis = New COMDConstSistema.NCOMConstSistema

lnProcesoCadaAgencia = loConstSis.LeeConstSistema(121)  ' gConstSistPigRemateCadaAg
    If lnProcesoCadaAgencia = 1 Then  ' En cada agencia
        fsRemateCadaAgencia = gsCodCMAC & gsCodAge
    Else
        fsRemateCadaAgencia = gsCodCMAC & "00"
    End If
    lsUltRemate = nColP.nObtieneNroUltimoProceso("R", fsRemateCadaAgencia, lsmensaje)
    If Trim(lsmensaje) <> "" Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        Exit Function
    End If
    Set lrdatosrem = nColP.nObtieneDatosProcesoRGCredPig("R", lsUltRemate, fsRemateCadaAgencia, lsmensaje)
    If Trim(lsmensaje) <> "" Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        Exit Function
    End If
If (lrdatosrem Is Nothing) Then
    Exit Function
End If
fsNroProceso = lrdatosrem!cNroProceso
psfecremate = Format(lrdatosrem!dProceso, "dd/mm/yyyy")

Set loParam = New COMDColocPig.DCOMColPCalculos
     fnDiasAtrasoAvisoRemate = Int(loParam.dObtieneColocParametro(gConsColPDiasAtrasoParaRemate))
    
Set loParam = Nothing
    
'Set loConstSis = New COMDConstSistema.NCOMConstSistema
Set loConstSis = Nothing

ValidaEntraRemate = False

'bcolpRemate = nColP.bEsParaRemate(sCodCta, psfecremate, pnDiasVctoParaRemate, psFechaSis, psNroRemate)
bcolpRemate = nColP.bEsParaRemate(sCodCta, psfecremate, fnDiasAtrasoAvisoRemate, psfecremate, lsUltRemate)
  If bcolpRemate Then
    ValidaEntraRemate = True
  Else
    ValidaEntraRemate = False
  End If
 
 Set nColP = Nothing

End Function


Public Sub MuestraDatosContrato(ByVal psNroContrato As String)
'On Error GoTo ControlError
Dim lrCredPig As New ADODB.Recordset
Dim lrCredPigPersonas As New ADODB.Recordset
Dim lrCredPigJoyas As New ADODB.Recordset

Dim loConstSis As COMDConstSistema.NCOMConstSistema
Dim lnJoyasDet As Integer

Dim loMuestraContrato As COMDColocPig.DCOMColPContrato

    Set loMuestraContrato = New COMDColocPig.DCOMColPContrato
    
        Set lrCredPig = loMuestraContrato.dObtieneDatosCreditoPignoraticio(psNroContrato)
        Set lrCredPigPersonas = loMuestraContrato.dObtieneDatosCreditoPignoraticioPersonas(psNroContrato)
        Set lrCredPigJoyas = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyas(psNroContrato)
     
    Set loMuestraContrato = Nothing
        
    If lrCredPig.BOF And lrCredPig.EOF Then
        lrCredPig.Close
        Set lrCredPig = Nothing
        Set lrCredPigJoyas = Nothing
        MsgBox " No se encuentra el Credito Pignoraticio " & psNroContrato, vbInformation, " Aviso "
        Exit Sub
    Else
    Dim i As Long
    i = 1
    
      While Not lrCredPigPersonas.EOF
            
                FECliente.TextMatrix(i, 1) = lrCredPigPersonas!cpersapellido & " , " & lrCredPigPersonas!cPersNombre
                FECliente.TextMatrix(i, 2) = Trim(lrCredPigPersonas!cPersDireccDomicilio)
                FECliente.TextMatrix(i, 3) = Trim(IIf(IsNull(lrCredPigPersonas!NroDNI), "", lrCredPigPersonas!NroDNI))
                FECliente.TextMatrix(i, 4) = Trim(IIf(IsNull(lrCredPigPersonas!NroRuc), "", lrCredPigPersonas!NroRuc))
                
              If i < lrCredPigPersonas.RecordCount Then
                i = i + 1
                FECliente.AdicionaFila
                
              End If
              
                
            lrCredPigPersonas.MoveNext
    Wend
        lrCredPigPersonas.Close
        Set lrCredPigPersonas = Nothing
    
       
        lblOroBru.Caption = lrCredPig!nOroBruto
        LblOroNeto.Caption = lrCredPig!nOroNeto
        lblTipo.Caption = lrCredPig!cTipCta
        lblPiezas.Caption = lrCredPig!npiezas
        lblTasacion.Caption = lrCredPig!nTasacion
        lblPrestamo.Caption = lrCredPig!nMontoCol
        lblSaldo.Caption = lrCredPig!nSaldo
        lblEstado.Tag = lrCredPig!nPrdEstado
        lblEstado.Caption = lrCredPig!cEstado
        lblFecPrestamo.Caption = Format(lrCredPig!dVigencia, "dd/mm/yyyy")
        lblFecVenc.Caption = Format(lrCredPig!dVenc, "dd/mm/yyyy")

        lrCredPig.Close
        Set lrCredPig = Nothing

        ' Kilatajes
        
        lbl14.Caption = IIf(IsNull(lrCredPigJoyas!nK14), "0.00", lrCredPigJoyas!nK14)
        lbl16.Caption = IIf(IsNull(lrCredPigJoyas!nK16), "0.00", lrCredPigJoyas!nK16)
        lbl18.Caption = IIf(IsNull(lrCredPigJoyas!nK18), "0.00", lrCredPigJoyas!nK18)
        lbl21.Caption = IIf(IsNull(lrCredPigJoyas!nK21), "0.00", lrCredPigJoyas!nK21)
        
        lrCredPigJoyas.Close
        Set lrCredPigJoyas = Nothing
        
        
    End If
              
      
'Exit Sub

'ControlError:   ' Rutina de control de errores.
'    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
'        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub



Private Sub cmdCancelar_Click()
        Limpiar
        AXCodCta.NroCuenta = ""
        AXCodCta.CMAC = gsCodCMAC
        AXCodCta.Prod = "305"
        fraPiezasDet.Enabled = True
End Sub

Private Sub cmdGrabar_Click()
Dim locolp As COMNColoCPig.NCOMColPContrato, loContFunct As COMNContabilidad.NCOMContFunciones
Dim lsMovNro As String
Dim lrJoyas  As New ADODB.Recordset, lnTasacT As Double
Dim LNBRUTO As Double, LNNETO As Double

    Set lrJoyas = FEJoyas.GetRsNew
    lnTasacT = FEJoyas.SumaRow(5)
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    Set locolp = New COMNColoCPig.NCOMColPContrato
    
Call locolp.nRetasacionCredPignoraticio(AXCodCta.NroCuenta, lsMovNro, lnTasacT, Val(lblPiezas.Caption), Val(lbl14.Caption), Val(lbl16.Caption), Val(lbl18.Caption), Val(lbl21.Caption), lrJoyas, Val(Me.lblOroBru.Caption), Val(Me.LblOroNeto.Caption))
cmdGrabar.Enabled = False
fraPiezasDet.Enabled = False

Set locolp = Nothing

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub FEJoyas_Click()
Dim loConst As COMDConstantes.DCOMConstantes
Dim lrMaterial As New ADODB.Recordset
Dim LRMATERIAL2 As ADODB.Recordset
Set loConst = New COMDConstantes.DCOMConstantes

    Select Case FEJoyas.Col
    Case 2
        'Set lrMaterial = loConst.RecuperaConstantes(gColocPMaterialJoyas, , "C.cConsDescripcion")
        
        'lrMaterial.MoveFirst
        'lrMaterial.MarshalOptions = adMarshalAll
        Set LRMATERIAL2 = New ADODB.Recordset
         LRMATERIAL2.Fields.Append "cConsDescripcion", adChar, 2
         LRMATERIAL2.Open
         
        
        
           LRMATERIAL2.AddNew "cConsDescripcion", "14"
           LRMATERIAL2.AddNew "cConsDescripcion", "16"
           LRMATERIAL2.AddNew "cConsDescripcion", "18"
           LRMATERIAL2.AddNew "cConsDescripcion", "21"
           
           LRMATERIAL2.MoveFirst
           
            
        FEJoyas.CargaCombo LRMATERIAL2
       ' Set lrMaterial = Nothing
        Set LRMATERIAL2 = Nothing
    End Select
    
    
    
    Set loConst = Nothing

End Sub

Private Sub FEJoyas_OnCellChange(pnRow As Long, pnCol As Long)
Dim loColPCalculos As COMDColocPig.DCOMColPCalculos
Dim lnPOro As Double
    
    
    If FEJoyas.Col = 4 Then     'Peso Neto

        If FEJoyas.TextMatrix(FEJoyas.Row, 4) <> "" Then
            If CCur(FEJoyas.TextMatrix(FEJoyas.Row, 4)) < 0 Then
                MsgBox "Peso Neto no puede ser negativo", vbInformation, "Aviso"
                FEJoyas.TextMatrix(FEJoyas.Row, 4) = 0
            Else
                If CCur(FEJoyas.TextMatrix(FEJoyas.Row, 4)) > CCur(FEJoyas.TextMatrix(FEJoyas.Row, 3)) Then
                    MsgBox "Peso Neto no puede ser mayor que Peso Bruto", vbInformation, "Aviso"
                    FEJoyas.TextMatrix(FEJoyas.Row, 4) = 0
                Else
                    'CalculaTasacion
                        Set loColPCalculos = New COMDColocPig.DCOMColPCalculos
                        lnPOro = loColPCalculos.dObtienePrecioMaterial(1, Val(Right(FEJoyas.TextMatrix(FEJoyas.Row, 2), 3)), 1)
                        If lnPOro <= 0 Then
                            MsgBox "Precio del Material No ha sido ingresado en el Tarifario, actualice el Tarifario", vbInformation, "Aviso"
                            Exit Sub
                        End If
                        Set loColPCalculos = Nothing
                        'Calcula el Valor de Tasacion
                        FEJoyas.TextMatrix(FEJoyas.Row, 5) = Format$(Val(FEJoyas.TextMatrix(FEJoyas.Row, 4) * lnPOro), "#####.00")
                            
                End If
            End If
        End If
        
    End If
    
    Dim i As Integer
    Dim v14 As Double, v16 As Double, v18 As Double, v21 As Double
    
    v14 = 0: v16 = 0: v18 = 0: v21 = 0
    
    
    
    For i = 1 To FEJoyas.Rows - 1
            If Left(Trim(FEJoyas.TextMatrix(i, 2)), 2) = "14" Then
                v14 = v14 + Val(FEJoyas.TextMatrix(i, 4))
            ElseIf Left(Trim(FEJoyas.TextMatrix(i, 2)), 2) = "16" Then
                v16 = v16 + Val(FEJoyas.TextMatrix(i, 4))
            ElseIf Left(Trim(FEJoyas.TextMatrix(i, 2)), 2) = "18" Then
                v18 = v18 + Val(FEJoyas.TextMatrix(i, 4))
            ElseIf Left(Trim(FEJoyas.TextMatrix(i, 2)), 2) = "21" Then
                v21 = v21 + Val(FEJoyas.TextMatrix(i, 4))
            End If
            
    Next i
              
    lbl14.Caption = Format(v14, "#0.00")
    lbl16.Caption = Format(v16, "#0.00")
    lbl18.Caption = Format(v18, "#0.00")
    lbl21.Caption = Format(v21, "#0.00")
            
    SumaColumnas
End Sub

' TODOCOMPLETA FUNCION PIG *********************************************
' **********************************************************************
Private Sub SumaColumnas()
Dim i As Integer
Dim lnPiezasT As Integer, lnPBrutoT As Double, lnPNetoT As Double, lnTasacT As Double
    lnPiezasT = 0: lnPBrutoT = 0:       lnPNetoT = 0:       lnTasacT = 0 ':         lnPrestamoT = 0
    'Total Piezas
    lnPiezasT = FEJoyas.SumaRow(1)
    lblPiezas.Caption = Format$(lnPiezasT, "##")

    'PESO BRUTO
    lnPBrutoT = FEJoyas.SumaRow(3)
    lblOroBru.Caption = Format$(lnPBrutoT, "######.00")

    'PESO NETO
    lnPNetoT = FEJoyas.SumaRow(4)
    LblOroNeto.Caption = Format$(lnPNetoT, "######.00")

'    lnTasacT = FEJoyas.SumaRow(5)
'    lblTasacion.Caption = Format$(lnTasacT, "######.00")
'
    
End Sub

Private Sub FEJoyas_OnChangeCombo()
    Dim loColPCalculos As COMDColocPig.DCOMColPCalculos
    Dim lnPOro As Double

Select Case FEJoyas.Col
    Case 2

        Set loColPCalculos = New COMDColocPig.DCOMColPCalculos
        lnPOro = loColPCalculos.dObtienePrecioMaterial(1, Val(Right(FEJoyas.TextMatrix(FEJoyas.Row, 2), 3)), 1)
        If lnPOro <= 0 Then
            MsgBox "Precio del Material No ha sido ingresado en el Tarifario, actualice el Tarifario", vbInformation, "Aviso"
            Exit Sub
        End If
        Set loColPCalculos = Nothing
        'Calcula el Valor de Tasacion
        FEJoyas.TextMatrix(FEJoyas.Row, 5) = Format$(Val(FEJoyas.TextMatrix(FEJoyas.Row, 4) * lnPOro), "#####.00")
End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim loParam As COMDColocPig.DCOMColPContrato
Set loParam = New COMDColocPig.DCOMColPContrato
Dim rsTemp As New ADODB.Recordset
Dim nEstado As Integer

    If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColConsuPrendario, False)
                           
        
        If sCuenta <> "" Then
            AXCodCta.NroCuenta = sCuenta
            AXCodCta.SetFocusCuenta
                      
            Limpiar
    

            If ValidaEntraRemate(AXCodCta.NroCuenta) = False Then
                    MsgBox "Este contrato no entrará al Remate: " & fsNroProceso & vbCrLf & "No puede realizar esta operación.", vbOKOnly + vbInformation, App.Title
                    Exit Sub
            End If
        
            Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
            Set loMuestraContrato = New COMDColocPig.DCOMColPContrato
            Dim lrCredPigJoyasDet As New ADODB.Recordset
            Dim i As Long
                i = 1
                
                MuestraDatosContrato AXCodCta.NroCuenta
                Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(AXCodCta.NroCuenta)
                
                While Not lrCredPigJoyasDet.EOF
                 
                    FEJoyas.TextMatrix(i, 0) = lrCredPigJoyasDet!nItem
                    FEJoyas.TextMatrix(i, 1) = lrCredPigJoyasDet!npiezas
                    FEJoyas.TextMatrix(i, 2) = lrCredPigJoyasDet!ckilataje
                    FEJoyas.TextMatrix(i, 3) = lrCredPigJoyasDet!nPesoBruto
                    FEJoyas.TextMatrix(i, 4) = lrCredPigJoyasDet!npesoneto
                    FEJoyas.TextMatrix(i, 5) = lrCredPigJoyasDet!nvaltasac
                    FEJoyas.TextMatrix(i, 6) = Trim(lrCredPigJoyasDet!cdescrip)
                    If i < lrCredPigJoyasDet.RecordCount Then
                        i = i + 1
                        FEJoyas.AdicionaFila i
                    End If
                    
                    lrCredPigJoyasDet.MoveNext
                Wend
               
            Set lrCredPigJoyasDet = Nothing
            Set loMuestraContrato = Nothing
            cmdGrabar.Enabled = True
            cmdCancelar.Enabled = True
            AXCodCta.Enabled = False
               
                
         End If
    End If
    Set rsTemp = Nothing
    Set loParam = Nothing

End Sub

Private Sub Form_Load()
    AXCodCta.CMAC = gsCodCMAC
End Sub

