VERSION 5.00
Begin VB.Form frmComisionRepTarj 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comision Reposicion de Tarjeta"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Detalle"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   4800
         TabIndex        =   9
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   3480
         TabIndex        =   8
         Top             =   1560
         Width           =   1215
      End
      Begin VB.ComboBox cmbMone 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1590
         Width           =   1095
      End
      Begin SICMACT.TxtBuscar txtCodPers 
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.Label lblMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   2160
         TabIndex        =   7
         Top             =   1590
         Width           =   1215
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   950
         Width           =   5055
      End
      Begin VB.Label Label3 
         Caption         =   "Monto :"
         Height          =   255
         Left            =   195
         TabIndex        =   4
         Top             =   1620
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   195
         TabIndex        =   3
         Top             =   1010
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo :"
         Height          =   255
         Left            =   195
         TabIndex        =   2
         Top             =   420
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmComisionRepTarj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lnMone As Integer
Dim lnSol As Double
Dim lnDol As Double

Private Sub cmbMone_Click()
    lnMone = Right(cmbMone.Text, 1)
    If lnMone = gMonedaNacional Then
        lblmonto.BackColor = &HC0FFFF
        'lblMonto.Caption = Format(CStr(CargaComXRepoTarjeta(gMonedaNacional)), "##0.00")
        lblmonto.Caption = Format(lnSol, "##0.00")
    Else
        lblmonto.BackColor = &HC0FFC0
        'lblMonto.Caption = Format(CStr(CargaComXRepoTarjeta(gMonedaExtranjera)), "##0.00")
        lblmonto.Caption = Format(lnDol, "##0.00")
    End If
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
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

    Dim clsCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim clsCont As COMNContabilidad.NCOMContFunciones
    Dim clsCapM As COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim clsMov As COMDMov.DCOMMov

    Dim lsMov As String
    Dim lnOpeCod As CaptacOperacion
    Dim lnMonto As Currency
    Dim lsPersCod As String

    On Error GoTo Error

    Set clsCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Set clsCont = New COMNContabilidad.NCOMContFunciones
    Set clsCapM = New COMDCaptaGenerales.DCOMCaptaMovimiento
    Set clsMov = New COMDMov.DCOMMov

    lsMov = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    'lnOpeCod = 300470
    lnOpeCod = gComiAhoReposicionTarjeta 'JUEZ 20150928
    lnMonto = CCur(lblmonto.Caption)
    lsPersCod = txtCodPers.Text

    If lsPersCod = "" Then
        MsgBox "Debe seleccionar un cliente", vbExclamation, "MENSAJE DEL SISTEMA"
        txtCodPers.SetFocus
        Exit Sub
    End If

    If Not ValidaTarjAnt(lsPersCod) Then
        LimpiaControles
        MsgBox "El Cliente no tiene tarjeta anterior ", vbExclamation, "MENSAJE DEL SISTEMA"
        txtCodPers.SetFocus
        Exit Sub
    End If

    If MsgBox("Desea Grabar la Información", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion
        Dim oBol As COMNCaptaGenerales.NCOMCaptaImpresion
        Dim lnMoneda As String
        Dim nTC As Double
        Dim loLavDinero As frmMovLavDinero
        Dim sPersLavDinero As String
        Dim nMontoLavDinero As Double
        
        Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
        Set loLavDinero = New frmMovLavDinero
        Set oBol = New COMNCaptaGenerales.NCOMCaptaImpresion

        lnMovNro = clsCapMov.OtrasOperaciones(lsMov, lnOpeCod, lnMonto, "", "COMISION REPOSICION TARJETA", Right(Me.cmbMone.Text, 1), Me.txtCodPers.Text, , loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro)
        'ALPA20131001*****************************
        If gnMovNro = 0 Then
            MsgBox "La operación no se realizó, favor intente nuevamente", vbInformation, "Aviso"
            Exit Sub
        End If
        '*****************************************
        '***ANPS**************************************
        Dim lnCodMotBlo As Integer
        lnCodMotBlo = clsCapMov.ObtenerMotivoBloqueoTarjetaEfectivo(Trim(Me.txtCodPers.Text))
        Set clsCont = Nothing
        Set clsCapMov = Nothing
        Set clsMov = Nothing
        If lnCodMotBlo <> 0 Then
            If lnCodMotBlo = 1 Or lnCodMotBlo = 13 Then
                lsBoleta = oBol.ImprimeBoleta("OTRAS OPERACIONES", "COMISION REPOSICION TARJETA", CStr(lnOpeCod), Str(lnMonto), lblNombre.Caption, "________" & Trim(Right(Me.cmbMone.Text, 1)), "", 0, "0", IIf(Len(lsDocumento) = 0, "", "Nro Documento"), 0, 0, False, False, , , , False, , "Nro Ope. : " & Str(lnMovNro), , gdFecSis, gsNomAge, gsCodUser, sLpt, , False, 0, , , , 1)
                Do
                    If Trim(lsBoleta) <> "" Then
                        nFicSal1 = FreeFile
                        Open sLpt For Output As nFicSal1
                                        Print #nFicSal1, lsImpBoleta
                                        Print #nFicSal1, ""
                                        Print #nFicSal1, ""
                                    Close #nFicSal1
                                End If
                Loop Until MsgBox("¿Desea Re-Imprimir Boletas ?", vbQuestion + vbYesNo, "Aviso") = vbNo
            Else
                lsBoleta = oBol.ImprimeBoleta("OTRAS OPERACIONES", "COMISION REPOSICION TARJETA", CStr(lnOpeCod), Str(lnMonto), lblNombre.Caption, "________" & Trim(Right(Me.cmbMone.Text, 1)), "", 0, "0", IIf(Len(lsDocumento) = 0, "", "Nro Documento"), 0, 0, False, False, , , , False, , "Nro Ope. : " & Str(lnMovNro), , gdFecSis, gsNomAge, gsCodUser, sLpt, , False, 0, , , , 0)
                Do
                    If Trim(lsBoleta) <> "" Then
                        nFicSal1 = FreeFile
                        Open sLpt For Output As nFicSal1
                                        Print #nFicSal1, lsImpBoleta
                                        Print #nFicSal1, ""
                                        Print #nFicSal1, ""
                                    Close #nFicSal1
                                End If
                Loop Until MsgBox("¿Desea Re-Imprimir Boletas ?", vbQuestion + vbYesNo, "Aviso") = vbNo
            End If

        Else
            MsgBox "No se pudo Generar la Boleta, Comuniquese con Sistema"

            End If
        '***fin ANPS**************************************

        lsBoleta = oBol.ImprimeBoleta("OTRAS OPERACIONES", "COMISION REPOSICION TARJETA", CStr(lnOpeCod), Str(lnMonto), lblNombre.Caption, "________" & Trim(Right(Me.cmbMone.Text, 1)), "", 0, "0", IIf(Len(lsDocumento) = 0, "", "Nro Documento"), 0, 0, False, False, , , , False, , "Nro Ope. : " & Str(lnMovNro), , gdFecSis, gsNomAge, gsCodUser, sLpt, , False, 0)
        
        Set oBol = Nothing

        Do
            If Trim(lsBoleta) <> "" Then
                lsBoleta = lsBoleta & oImpresora.gPrnSaltoLinea
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                    Print #nFicSal, oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                    Print #nFicSal, ""
                Close #nFicSal
          End If

            If Trim(lsBoletaITF) <> "" Then
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                Print #nFicSal, lsBoletaITF
                Print #nFicSal, ""
            Close #nFicSal
          End If

        Loop While MsgBox("Desea Re Imprimir ?", vbQuestion + vbYesNo, "Aviso") = vbYes
        Set oBol = Nothing
        'INICIO JHCU ENCUESTA 16-10-2019
        Encuestas gsCodUser, gsCodAge, "ERS0292019", gsOpeCod
        'FIN
        Unload Me
    End If

    Exit Sub
Error:
      MsgBox Str(err.Number) & err.Description
End Sub

Private Sub Form_Load()

    cmbMone.AddItem ("S/." & space(30) & "1")
    cmbMone.AddItem ("$" & space(30) & "2")
    cmbMone.ListIndex = 0

    lblmonto.BackColor = &HC0FFFF
    lblmonto.Caption = Format(CStr(CargaComXRepoTarjeta(1)), "##0.00")

    lnSol = Format(CStr(CargaComXRepoTarjeta(gMonedaNacional)), "##0.00")
    lnDol = Format(CStr(CargaComXRepoTarjeta(gMonedaExtranjera)), "##0.00")

End Sub

Private Sub txtCodPers_EmiteDatos()
    Dim oPersn As New uPersona
    If txtCodPers.Text <> "" Then 'JUEZ 20150928
        Set oPersn = New uPersona
        
        oPersn.ObtieneClientexCodigo (txtCodPers.Text)

        lblNombre.Caption = oPersn.sPersNombre
    Else 'JUEZ 20150928
        lblNombre.Caption = ""
    End If
End Sub
Private Sub LimpiaControles()
    txtCodPers.Text = ""
    lblNombre.Caption = ""
End Sub
