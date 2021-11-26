VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCredGarantiaVerificaLegal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verifica Garantia Legal"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9840
   Icon            =   "frmCredGarantiaVerificaLegal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin VB.ComboBox cboestados 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2280
         Width           =   3735
      End
      Begin VB.CommandButton CmdBuscaPersona 
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
         Height          =   320
         Left            =   3600
         TabIndex        =   4
         ToolTipText     =   "Busca Documentos de Persona"
         Top             =   180
         Width           =   1080
      End
      Begin VB.TextBox txtNumGar 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0.00;(0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   3
         Tag             =   "txtPrincipal"
         Text            =   "0"
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton cmdsalir 
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
         Height          =   320
         Left            =   8400
         TabIndex        =   2
         ToolTipText     =   "Busca Documentos de Persona"
         Top             =   2280
         Width           =   1080
      End
      Begin VB.CommandButton CmdActualizaLegal 
         Caption         =   "&Actualizar"
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
         Height          =   320
         Left            =   4200
         TabIndex        =   1
         ToolTipText     =   "Activa Legal "
         Top             =   2280
         Width           =   1080
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAnalista 
         Height          =   1215
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2143
         _Version        =   393216
         Cols            =   7
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lbltitular 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Num Garantia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label lbltitu 
         BackStyle       =   0  'Transparent
         Caption         =   "Titular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmCredGarantiaVerificaLegal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdActualizaLegal_Click()
    Dim oGarantia As COMDCredito.DCOMGarantia
    Dim objPista As COMManejador.Pista
    Dim lnValor As Integer
    Dim lsValor As String
    
    On Error GoTo ErrActualizar
    
    If cboestados.ListIndex = -1 Then
        MsgBox "Ud. debe de seleccionar la Verificación Legal realizada", vbInformation, "Aviso"
        EnfocaControl cboestados
        Exit Sub
    End If
    lnValor = CInt(Trim(Right(cboestados.Text, 3)))
    lsValor = UCase(Trim(Mid(cboestados.Text, 1, Len(cboestados.Text) - 3)))

    If MsgBox("¿Está seguro de guardar la Verificación Legal?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Set oGarantia = New COMDCredito.DCOMGarantia
    oGarantia.dUpdateGarantiasVerificaLegal Trim(grdAnalista.TextMatrix(1, 1)), lnValor 'EJVG20150705
    
    Set objPista = New COMManejador.Pista
    objPista.InsertarPista "190342", GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Verifica Legal -> " & lsValor, Trim(grdAnalista.TextMatrix(1, 1)), gCodigoGarantia
    
    MsgBox "La Garantía N° " & Trim(grdAnalista.TextMatrix(1, 1)) & " se ha actualizado a " & lsValor, vbInformation, "Aviso"
    
    If lnValor = 2 Or lnValor = 4 Then
        oGarantia.dUpdateGarantiasLegal Trim(grdAnalista.TextMatrix(1, 1)), 1
        
        objPista.InsertarPista "190341", GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Bloqueo Legal", Trim(grdAnalista.TextMatrix(1, 1)), gCodigoGarantia
        
        MsgBox "La Garantía N° " & Trim(grdAnalista.TextMatrix(1, 1)) & " se ha bloqueado con éxito", vbInformation, "Aviso"
    'ElseIf lnValor = 1 Or lnValor = 3 Then
    '    oGarantia.dUpdateGarantiasLegal Trim(grdAnalista.TextMatrix(1, 1)), 0
    End If
    Set objPista = Nothing
    Set oGarantia = Nothing
    
    CmdBuscaPersona_Click
    Exit Sub
ErrActualizar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
'EJVG20150706 ***
'Private Sub CmdBuscaPersona_Click()
'Dim bCargaDatos As Boolean
'Dim oGarantia As COMDCredito.DCOMGarantia
'Dim rsGarantia As ADODB.Recordset
'Dim rsGarantReal As ADODB.Recordset
'Dim pbGarantiaVerificaLegal As Boolean
'Dim i As Integer
'
'bCargaDatos = False
'i = 0
'    If val(Me.txtNumGar) <= 0 Then
'        CmdActualizaLegal.Enabled = False
'        Exit Sub
'    End If
'
'    sNumgarant = Format(Me.txtNumGar, "00000000")
'
'    Set oGarantia = New COMDCredito.DCOMGarantia
'    Call oGarantia.CargarDatosGarantiaLegalVerifica(sNumgarant, rsGarantia, rsGarantReal, pbGarantiaVerificaLegal)
'
'    Set oGarantia = Nothing
'
'    If Not (rsGarantia.EOF And rsGarantia.BOF) Then
'        bCargaDatos = True
'    Else
'        lbltitular = ""
'        lbltitu.Visible = False
'        lbltitular.Visible = False
'        bCargaDatos = False
'        lbltitular = ""
'    End If
'
'    If Not bCargaDatos Then
'        MsgBox "El número de garantía que acaba de ingresar no existe.", vbExclamation, "Atención"
'        Exit Sub
'    End If
'
'    If Not (rsGarantReal.EOF And rsGarantReal.BOF) Then
'        lbltitu.Visible = True
'        lbltitular.Visible = True
'        lbltitular = rsGarantReal!cPersNombre
'        CmdActualizaLegal.Enabled = True
'        cboestados.Enabled = True
'        Call CargarCombosEstadoGarantiaLegal
'
'        Call LimpiaFlex(grdAnalista)
'
'            ConfigurarMShComite
'            For i = 0 To rsGarantReal.RecordCount - 1
'
'                grdAnalista.TextMatrix(i + 1, 0) = i + 1
'                grdAnalista.row = i + 1
'                grdAnalista.Col = 1
'
'                grdAnalista.TextMatrix(i + 1, 1) = rsGarantReal!cNumGarant
'                grdAnalista.TextMatrix(i + 1, 2) = rsGarantReal!cdescripcion
'                grdAnalista.TextMatrix(i + 1, 3) = rsGarantReal!dTasacion
'                grdAnalista.TextMatrix(i + 1, 4) = rsGarantReal!nVRM
'                grdAnalista.TextMatrix(i + 1, 5) = IIf(rsGarantReal!nApruebaLegal = 1, "No Modificable", "Modificable")
'
'                Select Case rsGarantReal!nVerificaLegal
'                    Case 1
'                        grdAnalista.TextMatrix(i + 1, 6) = "Pendiente"
'                    Case 2
'                        grdAnalista.TextMatrix(i + 1, 6) = "Aprobado"
'                    Case 3
'                        grdAnalista.TextMatrix(i + 1, 6) = "Desaprobado"
'                    Case 4
'                        grdAnalista.TextMatrix(i + 1, 6) = "Pendiente por Regularizar"
'                    Case 0
'                        grdAnalista.TextMatrix(i + 1, 6) = "Pendiente"
'                End Select
'                cboestados.ListIndex = IndiceListaCombo(cboestados, IIf(rsGarantReal!nVerificaLegal = 0, 1, rsGarantReal!nVerificaLegal))
'                rsGarantReal.MoveNext
'            Next
'
'    Else
'        MsgBox "El número de garantía que acaba de ingresar no es una Real.", vbExclamation, "Atención"
'        CmdActualizaLegal.Enabled = False
'        Call LimpiaFlex(grdAnalista)
'        lbltitular = ""
'        lbltitu.Visible = False
'        lbltitular.Visible = False
'        lbltitular = ""
'        Exit Sub
'    End If
'
'    rsGarantReal.Close
'    rsGarantia.Close
'    Set rsGarantReal = Nothing
'    Set rsGarantia = Nothing
'End Sub
Private Sub CmdBuscaPersona_Click()
    Dim lsNumGarant As String
    Dim oGarantia As COMDCredito.DCOMGarantia
    Dim rsGarantia As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ErrBuscar
    
    lbltitu.Visible = False
    lbltitular.Visible = False
    lbltitular.Caption = ""
    Call LimpiaFlex(grdAnalista)
    ConfigurarMShComite
    cboestados.ListIndex = -1
    cboestados.Enabled = False
    CmdActualizaLegal.Enabled = False
    
    If val(txtNumGar.Text) = 0 Then
        MsgBox "Ud. debe de especificar el Nro. de Garantía", vbInformation, "Aviso"
        EnfocaControl txtNumGar
        Exit Sub
    End If
    lsNumGarant = txtNumGar.Text
    
    Set oGarantia = New COMDCredito.DCOMGarantia
    Set rsGarantia = oGarantia.RecuperaGarantiaxVerificacionLegal(lsNumGarant)
    Set oGarantia = Nothing
    
    If rsGarantia.EOF Then
        MsgBox "No se ha encontrado la garantía especificada, asegurese" & Chr(13) & "digitarlo correctamente y que la garantía sea Real.", vbInformation, "Aviso"
        RSClose rsGarantia
        Exit Sub
    End If
    
    If Not rsGarantia!bTramiteLegal Then
        MsgBox "El número de garantía que acaba de ingresar no es una Real", vbInformation, "Aviso"
        RSClose rsGarantia
        Exit Sub
    End If
    
    lbltitu.Visible = True
    lbltitular.Visible = True
    lbltitular = rsGarantia!cPersNombre
    cboestados.Enabled = True
    CmdActualizaLegal.Enabled = True
    
    grdAnalista.TextMatrix(i + 1, 0) = i + 1
    grdAnalista.row = i + 1
    grdAnalista.col = 1
                   
    grdAnalista.TextMatrix(i + 1, 1) = rsGarantia!cNumGarant
    grdAnalista.TextMatrix(i + 1, 2) = rsGarantia!cDescripcion
    grdAnalista.TextMatrix(i + 1, 3) = Format(rsGarantia!dTasacion, gsFormatoFechaView)
    grdAnalista.TextMatrix(i + 1, 4) = Format(rsGarantia!nVRM, gsFormatoNumeroView)
    grdAnalista.TextMatrix(i + 1, 5) = IIf(rsGarantia!bBloqueoLegal, "Bloqueado", "Desbloqueado")
    grdAnalista.TextMatrix(i + 1, 6) = rsGarantia!cVerificaLegal
    cboestados.ListIndex = IndiceListaCombo(cboestados, rsGarantia!nVerificaLegal)
       
    RSClose rsGarantia
    Exit Sub
ErrBuscar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
'END EJVG *******
Private Sub cmdSalir_Click()
    If lbltitular.Visible = True Then
        'sNumgarant = ""
        lbltitular = ""
        lbltitu.Visible = False
        lbltitular.Visible = False
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    CentraForm Me
    ConfigurarMShComite
    CargarCombosEstadoGarantiaLegal
End Sub

Sub ConfigurarMShComite()
 grdAnalista.Clear
    grdAnalista.Cols = 7
    grdAnalista.Rows = 2
    
    With grdAnalista
        .TextMatrix(0, 1) = "Garantia"
        .TextMatrix(0, 2) = "Descripcion"
        .TextMatrix(0, 3) = "Fecha"
        .TextMatrix(0, 4) = "Valor"
        .TextMatrix(0, 5) = "Estado"
        .TextMatrix(0, 6) = "Estado Aprob."
        
        .ColWidth(0) = 800
        .ColWidth(1) = 1200
        .ColWidth(2) = 2500
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        .ColWidth(6) = 1500
        
    End With
End Sub

Sub CargarCombosEstadoGarantiaLegal()
   Dim OCon As COMDConstantes.DCOMConstantes
   Dim rsT As ADODB.Recordset
   Dim sDes As String
   Dim nCodigo As Integer
    
   On Error GoTo ErrHandler
   
   Set OCon = New COMDConstantes.DCOMConstantes
   Set rsT = OCon.RecuperaConstantes(9985)
   Set OCon = Nothing
   Call Llenar_Combo_con_Recordset(rsT, cboestados)

    Exit Sub
ErrHandler:
    MsgBox "Error al cargar datos", vbInformation, "AVISO"
End Sub
Private Sub txtNumGar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl CmdBuscaPersona
    End If
End Sub
Private Sub txtNumGar_LostFocus()
    txtNumGar = Format(txtNumGar, "00000000")
End Sub
