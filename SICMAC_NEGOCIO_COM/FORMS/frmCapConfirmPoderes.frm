VERSION 5.00
Begin VB.Form frmCapConfirmPoderes 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9165
   Icon            =   "frmCapConfirmPoderes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8040
      TabIndex        =   15
      Top             =   3960
      Width           =   1000
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6960
      TabIndex        =   14
      Top             =   3960
      Width           =   1000
   End
   Begin VB.Frame fraCliente 
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
      Height          =   2475
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   8925
      Begin VB.CommandButton cmdMostrarFirma 
         Caption         =   "Mostrar Firma"
         Height          =   315
         Left            =   7350
         TabIndex        =   10
         Top             =   2025
         Width           =   1440
      End
      Begin VB.CommandButton cmdVerRegla 
         Caption         =   "Ver Regla"
         Height          =   315
         Left            =   5760
         TabIndex        =   9
         Top             =   2025
         Visible         =   0   'False
         Width           =   1440
      End
      Begin SICMACT.FlexEdit grdCliente 
         Height          =   1755
         Left            =   150
         TabIndex        =   11
         Top             =   225
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   3096
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Codigo-Nombre-Relacion-Direccion-ID-Firma Oblig-Grupo-Presente"
         EncabezadosAnchos=   "250-1500-3200-1500-0-0-0-1000-1000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-8"
         ListaControles  =   "0-0-0-0-0-0-0-0-4"
         EncabezadosAlineacion=   "C-L-L-L-C-C-C-L-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   255
         RowHeight0      =   300
         TipoBusPersona  =   1
      End
      Begin VB.Label lblTipoCuenta 
         Alignment       =   2  'Center
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
         ForeColor       =   &H80000001&
         Height          =   300
         Left            =   1620
         TabIndex        =   13
         Top             =   2070
         Width           =   1800
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cuenta :"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   2123
         Width           =   960
      End
   End
   Begin VB.Frame fraCuenta 
      Caption         =   "Cuenta"
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
      Height          =   1275
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8910
      Begin VB.Frame fraDatos 
         Height          =   585
         Left            =   50
         TabIndex        =   2
         Top             =   630
         Width           =   8800
         Begin VB.Label lblUltContacto 
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   4740
            TabIndex        =   6
            Top             =   195
            Width           =   1995
         End
         Begin VB.Label lblEtqUltCnt 
            AutoSize        =   -1  'True
            Caption         =   "Ultimo Contacto :"
            Height          =   195
            Left            =   3435
            TabIndex        =   5
            Top             =   255
            Width           =   1215
         End
         Begin VB.Label lblApertura 
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   900
            TabIndex        =   4
            Top             =   195
            Width           =   1965
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Apertura :"
            Height          =   195
            Left            =   135
            TabIndex        =   3
            Top             =   255
            Width           =   690
         End
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   661
         Texto           =   "Cuenta N°:"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   3840
         TabIndex        =   7
         Top             =   240
         Width           =   3960
      End
   End
End
Attribute VB_Name = "frmCapConfirmPoderes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmCapConfirmPoderes
'** Descripción : Formulario para visualizar y confirmar los poderes para realizar operaciones según TI-ERS138-2013
'** Creación : JUEZ, 20131212 09:00:00 AM
'*****************************************************************************************************

Option Explicit

Dim fbConfirmPoderes As Boolean
Public nProducto As COMDConstantes.Producto
Dim nOperacion As COMDConstantes.CaptacOperacion
Dim sPersCodCMAC As String
Dim sNombreCMAC As String, sTipoCuenta As String
Dim sOperacion As String
Dim sCuenta As String
Dim strReglas As String
Dim bProcesoNuevo As Boolean

Public Function Inicia(ByVal psCtaCod As String, ByVal nProd As Producto, ByVal nOpe As CaptacOperacion, _
        ByVal sDescOperacion As String, Optional sCodCmac As String = "", Optional sNomCmac As String) As Boolean

sCuenta = psCtaCod
nProducto = nProd
nOperacion = nOpe
sPersCodCMAC = sCodCmac
sNombreCMAC = sNomCmac
sOperacion = sDescOperacion
fbConfirmPoderes = False

'CARGA DE DATOS SÓLO PARA AHORROS
lblEtqUltCnt = "Ult. Contacto :"
lblUltContacto.Width = 2000
txtCuenta.Prod = Trim(Str(gCapAhorros))
If sPersCodCMAC = "" Then
    Me.Caption = "Captaciones - Cargo - Ahorros - " & sDescOperacion
Else
    Me.Caption = "Captaciones - Cargo - Ahorros - " & sDescOperacion & " - " & sNombreCMAC
End If

grdCliente.ColWidth(6) = 0

txtCuenta.NroCuenta = sCuenta
Call txtCuenta_KeyPress(13)

Me.Show 1

Inicia = fbConfirmPoderes

End Function

Private Sub cmdGrabar_Click()
    If bProcesoNuevo = True Then
    
        If validarReglasPersonas = False Then
            MsgBox "Las personas seleccionadas no tienen suficientes poderes para realizar el retiro", vbInformation, "Aviso"
            Exit Sub
        End If
        
        Dim oPersonaTemp As COMNPersona.NCOMPersona
        Dim iTemp, nMenorEdad As Integer
        
        Set oPersonaTemp = New COMNPersona.NCOMPersona
        
        For iTemp = 1 To grdCliente.Rows - 1
            If oPersonaTemp.validarPersonaMayorEdad(grdCliente.TextMatrix(iTemp, 1), Format(gdFecSis, "dd/mm/yyyy")) = False _
               And grdCliente.TextMatrix(iTemp, 7) <> "PJ" Then
                nMenorEdad = nMenorEdad + 1
            End If
        Next
        
        If nMenorEdad > 0 Then
            If MsgBox("Uno de los intervinientes en la cuenta es menor de edad, SOLO podrá disponer de los fondos con autorización del Juez " & vbNewLine & "Desea continuar?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                
                Dim VistoElectronico As frmVistoElectronico
                Dim ResultadoVisto As Boolean
                Set VistoElectronico = New frmVistoElectronico
                ResultadoVisto = False
                ResultadoVisto = VistoElectronico.Inicio(3, nOperacion)
                If Not ResultadoVisto Then
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        End If
    End If
    
    fbConfirmPoderes = True
    Unload Me
End Sub

Private Sub cmdMostrarFirma_Click()
With grdCliente
    If .TextMatrix(.row, 1) = "" Then Exit Sub
    Call frmPersonaFirma.Inicio(Trim(.TextMatrix(.row, 1)), Trim(txtCuenta.Age), True)
End With
End Sub

Private Sub cmdSalir_Click()
    fbConfirmPoderes = False
    Unload Me
End Sub

Private Sub cmdVerRegla_Click()
    If strReglas <> "" Then
        Call frmCapVerReglas.Inicia(strReglas)
    Else
        MsgBox "Cuenta no tiene reglas definidas", vbInformation, "Aviso"
    End If
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim clsGen As COMDConstSistema.DCOMGeneral
Dim rsCta As ADODB.Recordset, rsRel As New ADODB.Recordset
Dim nEstado As COMDConstantes.CaptacEstado
Dim nRow As Long
Dim sMsg As String, sMoneda As String, sPersona As String
Dim lafirma As frmPersonaFirma
Dim ClsPersona As COMDPersona.DCOMPersonas
Dim Rf As ADODB.Recordset

Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    sMsg = clsCap.ValidaCuentaOperacion(sCuenta)
Set clsCap = Nothing
If sMsg = "" Then
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCuenta)
    If Not (rsCta.EOF And rsCta.BOF) Then
        strReglas = IIf(IsNull(rsCta!cReglas), "", rsCta!cReglas)
        
        nEstado = rsCta("nPrdEstado")
        lblApertura = Format$(rsCta("dApertura"), "dd mmm yyyy hh:mm:ss")
        
        Select Case nProducto
            Case gCapAhorros
                If rsCta("bOrdPag") Then
                    lblMensaje = "AHORROS CON ORDEN DE PAGO" & Chr$(13) & IIf(Mid(sCuenta, 9, 1) = gMonedaNacional, "MONEDA NACIONAL", "MONEDA EXTRANJERA")
                Else
                    lblMensaje = "AHORROS SIN ORDEN DE PAGO" & Chr$(13) & IIf(Mid(sCuenta, 9, 1) = gMonedaNacional, "MONEDA NACIONAL", "MONEDA EXTRANJERA")
                End If
                lblUltContacto = Format$(rsCta("dUltContacto"), "dd mmm yyyy hh:mm:ss")
                
            Case gCapPlazoFijo
                lblUltContacto = rsCta("nPlazo")
                
        End Select
        
        lblTipoCuenta = UCase(rsCta("cTipoCuenta"))
        sTipoCuenta = lblTipoCuenta
        Set rsRel = clsMant.GetPersonaCuenta(sCuenta)
        
        sPersona = ""
        
        Dim dlsMant As COMDCaptaGenerales.DCOMCaptaGenerales
        Set dlsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
                
        Do While Not rsRel.EOF
            If rsRel("cPersCod") = gsCodPersUser Then
                MsgBox "Ud. No puede hacer operaciones con sus propias cuentas.", vbInformation, "Aviso"
                Unload Me
                Exit Sub
            End If
            If sPersona <> rsRel("cPersCod") Then
                grdCliente.AdicionaFila
                nRow = grdCliente.Rows - 1
                grdCliente.TextMatrix(nRow, 1) = rsRel("cPersCod")
                grdCliente.TextMatrix(nRow, 2) = UCase(PstaNombre(rsRel("Nombre")))
                grdCliente.TextMatrix(nRow, 3) = UCase(rsRel("Relacion")) & Space(50) & Trim(rsRel("nPrdPersRelac"))
                grdCliente.TextMatrix(nRow, 4) = rsRel("Direccion") & ""
                grdCliente.TextMatrix(nRow, 5) = rsRel("ID N°")
                 
                If rsRel("cGrupo") <> "" Then
                    bProcesoNuevo = True
                    grdCliente.TextMatrix(nRow, 7) = rsRel("cGrupo")
                Else
                    bProcesoNuevo = False
                    grdCliente.TextMatrix(nRow, 6) = IIf(IsNull(rsRel("cobligatorio")) Or rsRel("cobligatorio") = "N", "NO", IIf(rsRel("cobligatorio") = "S", "SI", "OPCIONAL"))
                                    
                End If
                
                sPersona = rsRel("cPersCod")
            End If
            rsRel.MoveNext
        Loop
        
        
        If bProcesoNuevo Then
            cmdVerRegla.Visible = True
            grdCliente.ColWidth(1) = 1300
            grdCliente.ColWidth(2) = 3600
            grdCliente.ColWidth(3) = 1200
            grdCliente.ColWidth(6) = 0
            grdCliente.ColWidth(7) = 900
            grdCliente.ColWidth(8) = 1000
                
        Else
            cmdVerRegla.Visible = False
            grdCliente.ColWidth(1) = 1700
            grdCliente.ColWidth(2) = 3800
            grdCliente.ColWidth(3) = 1500
            grdCliente.ColWidth(7) = 0
            grdCliente.ColWidth(6) = 1200
            grdCliente.ColWidth(8) = 0
            
            MsgBox "Se recomienda actualizar los grupos y reglas de la cuenta a debitar", vbInformation, "Aviso"
            Set lafirma = New frmPersonaFirma
            Set ClsPersona = New COMDPersona.DCOMPersonas
            
            Set Rf = ClsPersona.BuscaCliente(grdCliente.TextMatrix(nRow, 1), BusquedaCodigo)
            
            If Not Rf.BOF And Not Rf.EOF Then
              If Rf!nPersPersoneria = 1 Then
              Call frmPersonaFirma.Inicio(Trim(grdCliente.TextMatrix(nRow, 1)), Mid(grdCliente.TextMatrix(nRow, 1), 4, 2), False, True)
              End If
            End If
            Set Rf = Nothing
        End If

        rsRel.Close
        Set rsRel = Nothing
        fraCliente.Enabled = True
        
        fraCuenta.Enabled = False
        
        cmdGrabar.Enabled = True
    End If
    
Else
    MsgBox sMsg, vbInformation, "Operacion"
    txtCuenta.SetFocus
End If
Set clsMant = Nothing
End Sub

Private Function validarReglasPersonas() As Boolean
 Dim sReglas() As String
    Dim sGrupos() As String
    Dim sTemporal As String
    Dim v1, v2 As Variant
    Dim bAprobado As Boolean
    Dim intRegla, i, j As Integer
    
    If Trim(strReglas) = "" Then
        validarReglasPersonas = False
        Exit Function
    End If
    sReglas = Split(strReglas, "-")
    For i = 1 To grdCliente.Rows - 1
        If grdCliente.TextMatrix(i, 8) = "." Then
            If j = 0 Then
               sTemporal = sTemporal & grdCliente.TextMatrix(i, 7)
            Else
               sTemporal = sTemporal & "," & grdCliente.TextMatrix(i, 7)
            End If
            j = j + 1
        End If
    Next
    If Trim(sTemporal) = "" Then
        validarReglasPersonas = False
        Exit Function
    End If
    sGrupos = Split(sTemporal, ",")
    For Each v1 In sReglas
        bAprobado = True
        For Each v2 In sGrupos
            If InStr(CStr(v1), CStr(v2)) = 0 Then
                bAprobado = False
                Exit For
            End If
        Next
        If bAprobado Then
            If UBound(sGrupos) = UBound(Split(CStr(v1), "+")) Then
                Exit For
            Else
                bAprobado = False
            End If
        End If
    Next
    validarReglasPersonas = bAprobado
End Function
