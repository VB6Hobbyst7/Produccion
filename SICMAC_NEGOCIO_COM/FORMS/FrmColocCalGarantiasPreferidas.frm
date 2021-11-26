VERSION 5.00
Begin VB.Form FrmColocCalGarantiasPreferidas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calificacion - Registrar Garantias Preferidas"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9105
   Icon            =   "FrmColocCalGarantiasPreferidas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   8850
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1170
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   1500
         TabIndex        =   11
         Top             =   240
         Width           =   1080
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   2700
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   3810
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   8850
      Begin VB.TextBox txtGravamen 
         Height          =   285
         Left            =   1260
         TabIndex        =   21
         Top             =   2961
         Width           =   1305
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   1005
         Left            =   1260
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   20
         Top             =   1443
         Width           =   6240
      End
      Begin VB.TextBox txtTitular 
         Height          =   285
         Left            =   1260
         TabIndex        =   19
         Top             =   2562
         Width           =   6285
      End
      Begin VB.TextBox txtEstado 
         Height          =   285
         Left            =   3435
         TabIndex        =   18
         Top             =   2961
         Width           =   1305
      End
      Begin VB.TextBox txtTasacion 
         Height          =   285
         Left            =   6240
         TabIndex        =   17
         Top             =   2961
         Width           =   1305
      End
      Begin VB.CheckBox ChkGarantiaPreferidas 
         Height          =   315
         Left            =   1800
         TabIndex        =   16
         Top             =   3360
         Width           =   285
      End
      Begin VB.ListBox LstGarantias 
         Height          =   645
         Left            =   1260
         TabIndex        =   15
         Top             =   285
         Width           =   6240
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   345
         Left            =   7680
         TabIndex        =   8
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label lblCredito 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4755
         TabIndex        =   23
         Top             =   1044
         Width           =   2745
      End
      Begin VB.Label lblGarantia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1260
         TabIndex        =   22
         Top             =   1044
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nro Garantia"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1089
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Nro Credito"
         Height          =   195
         Left            =   3600
         TabIndex        =   13
         Top             =   1089
         Width           =   795
      End
      Begin VB.Label Label9 
         Caption         =   "Garantias"
         Height          =   255
         Left            =   375
         TabIndex        =   7
         Top             =   285
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Garantias Preferidas"
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   3420
         Width           =   1425
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Valor tas"
         Height          =   195
         Left            =   5400
         TabIndex        =   5
         Top             =   3006
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   2760
         TabIndex        =   4
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Valor Gravan"
         Height          =   195
         Left            =   210
         TabIndex        =   3
         Top             =   3006
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Titular"
         Height          =   195
         Left            =   210
         TabIndex        =   2
         Top             =   2607
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   1443
         Width           =   840
      End
   End
End
Attribute VB_Name = "FrmColocCalGarantiasPreferidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lcCtaCod As String
Dim cPersCod As String
Dim nEstado As String
Dim nRealizacion As Double


'Estado de garantia :  0=Activa // 1=No Activa

Sub ActivarControles(ByVal bEstado As Boolean)
    CmdAceptar.Enabled = bEstado
    LstGarantias.Enabled = bEstado
End Sub
Sub LimpiarCampos()
    
    txtDescripcion = ""
    'txtTitular = ""
    txtGravamen = ""
    txtEstado = ""
    txtTasacion = ""
    CmdAceptar.Enabled = False
    Me.ChkGarantiaPreferidas.value = 0 'JACA 20110627
    Me.lblCredito = ""   'JACA 20110627
    Me.lblGarantia = ""  'JACA 20110627
End Sub



Private Sub CmdAceptar_Click()
    Dim loContFunct As COMNContabilidad.NCOMContFunciones
    Dim objGarantia As COMDCredito.DCOMGarantia
    Dim bolEstado As Boolean
    Dim lsMovNro As String
    On Error GoTo ErrHandler
    If ChkGarantiaPreferidas.value = 1 Then
        nEstado = 0
    Else
        nEstado = 1
    End If
'    If VerificarGuardar = True Then
      If MsgBox("Desea guardar garantia como preferida ? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
             lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        Set objGarantia = New COMDCredito.DCOMGarantia
              'bolEstado = objGarantia.InsertarGarantiaPreferidas(lsMovNro, txtNroGarantia.Text, nRealizacion, cPersCod, nEstado)COMENTADO X JACA 20110627
              bolEstado = objGarantia.InsertarGarantiaPreferidas(lsMovNro, Me.lblGarantia, nRealizacion, cPersCod, IIf(Me.ChkGarantiaPreferidas.value = 1, 3, 4), Me.lblCredito) 'JACA 20110627
        Set objGarantia = Nothing
        If bolEstado = True Then
             MsgBox "Se ha guardado correctamente", vbInformation, "Sistema"
        Else
             MsgBox "Se ha producido un error ", vbInformation, "Sistema"
        End If
        CmdAceptar.Enabled = False
      End If
'    Else
'        MsgBox "Datos incompletos", vbInformation, "Aviso"
'    End If
    Exit Sub
ErrHandler:
    If Not objGarantia Is Nothing Then Set objGarantia = Nothing
End Sub

Private Sub cmdBuscar_Click()
Dim oPers As COMDpersona.UCOMPersona
Dim strNombre As String
On Error GoTo ErrHandler
        Set oPers = frmBuscaPersona.Inicio
        If oPers Is Nothing Then Exit Sub
        cPersCod = oPers.sPersCod
        strNombre = oPers.sPersNombre
        Set oPers = Nothing
        Call LimpiarCampos
        txtTitular.Text = strNombre
        CargarGarantiasPreferidas
    Exit Sub
ErrHandler:
    MsgBox "Se ha producido un error al buscar una persona", vbInformation, "Error"
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiarCampos
 End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
 ActivarControles False
 LimpiarCampos
End Sub

Sub CargarGarantiasPreferidas()
    Dim objGarantia As COMDCredito.DCOMGarantia
    Dim rs As ADODB.Recordset
    Dim strDescripcion As String
    Dim strCodigo As String
    Dim strCtaCod As String
    On Error GoTo ErrHandler
    If cPersCod <> "" Then
        
        Set objGarantia = New COMDCredito.DCOMGarantia
            Set rs = objGarantia.ListaGarantias(cPersCod)
        Set objGarantia = Nothing
        LstGarantias.Clear
        Do Until rs.EOF
            strDescripcion = IIf(IsNull(rs!cdescripcion), "", rs!cdescripcion)
            strCodigo = IIf(IsNull(rs!cNumGarant), "", rs!cNumGarant)
            strCtaCod = IIf(IsNull(rs!cCtaCod), "", rs!cCtaCod)
            'JACA 20110627*********************************************************
            'LstGarantias.AddItem strDescripcion & "-" & strCodigo
            'LstGarantias.ItemData(LstGarantias.NewIndex) = val(strCodigo)
            LstGarantias.AddItem strCtaCod & "-" & strCodigo & "-" & strDescripcion
            LstGarantias.ItemData(LstGarantias.NewIndex) = val(strCodigo)
            'JACA END*********************************************************
            rs.MoveNext
        Loop
        Set rs = Nothing
        LstGarantias.Enabled = True
    End If
    Exit Sub
ErrHandler:
    MsgBox "Se ha producido un error al cargar las garantias", vbInformation, "Error"
End Sub

Private Sub lstGarantias_Click()
    Dim cNumGarant As String
    Dim sMatrix() As String
    If LstGarantias.ListIndex <> -1 Then
         'cNumGarant = ObtenerCodigo(LstGarantias.ItemData(LstGarantias.ListIndex))
'        txtNroGarantia = cNumGarant
'        txtNroGarantia.SetFocus
         'JACA 20110627*****************************************************
         sMatrix = Split(LstGarantias.Text, "-")
         
         Me.lblGarantia = sMatrix(1)
         Me.lblCredito = sMatrix(0)
         CargarDescripcionGarantia Me.lblGarantia, Me.lblCredito
         'JACA END**********************************************************
         
    End If
End Sub

Sub CargarDescripcionGarantia(ByVal pCnumGarant As String, ByVal pCtaCod As String) 'JACA 20110627 SE AGREGO LA VARIABLE pCtaCod
    Dim objGarantia As COMDCredito.DCOMGarantia
    Dim rs As ADODB.Recordset
    On Error GoTo ErrHandler
        Set objGarantia = New COMDCredito.DCOMGarantia
            'Set rs = objGarantia.DescripGarantia(pCnumGarant) 'COMENTADO BY JACA 20110627
            Set rs = objGarantia.DescripGarantia(pCnumGarant, pCtaCod) 'JACA 20110627
        Set objGarantia = Nothing
        If Not rs.EOF Or Not rs.BOF Then
            'Do Until rs.EOF
                txtGravamen = IIf(IsNull(rs!nGravament), "", rs!nGravament)
                txtEstado = IIf(IsNull(rs!cEstado), "", rs!cEstado)
                txtTasacion = IIf(IsNull(rs!nTasacion), "", rs!nTasacion)
                nRealizacion = IIf(IsNull(rs!nRealizacion), 0, rs!nRealizacion)
                nEstado = IIf(IsNull(rs!nEstado), 0, rs!nEstado)
                'JACA 20110627**************************************************
                    If rs!nEstado = 3 Then
                        ChkGarantiaPreferidas.value = 1
                    Else
                        ChkGarantiaPreferidas.value = 0
                    End If
                'JACA END *****************************************************
                cPersCod = rs!cPersCod
                txtTitular.Text = rs!cPersNombre
                Me.CmdAceptar.Enabled = True
             '   rs.MoveNext
            'Loop
        Else
             'JACA 20110627
             MsgBox "La Garantia no se encuentra en la Lista de Calificacion", vbInformation, "Aviso"
             LimpiarCampos
        End If
        Set rs = Nothing
    Exit Sub
ErrHandler:
    'If Not oconecta Is Nothing Then Set oconecta = Nothing
    'If Not rs Is Nothing Then Set rs = Nothing
    MsgBox "Se ha producido un erro al cargar la data ", vbInformation, "Error"
End Sub
Function ObtenerCodigo(ByVal pintCodigo As Long) As String
    Dim strCodigo As String
    Dim nLongitud As Integer
    
    nLongitud = Len(CStr(pintCodigo))
    strCodigo = Right("00000000", 8 - nLongitud) & CStr(pintCodigo)
    
    ObtenerCodigo = strCodigo
End Function

Function VerificarGuardar() As Boolean
'Comentado x JACA 20110530*****************************
'    If txtNroGarantia.Text <> "" Or txtGravamen.Text <> "" Or txtEstado.Text <> "" Or txtTasacion.Text <> "" Then
'        VerificarGuardar = True
'    Else
'        VerificarGuardar = True
'    End If
'JACA END***********************************************
End Function


Private Sub txtNroGarantia_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If Len(txtNroGarantia.Text) > 0 Then
'        'CargarDescripcionGarantia cNumGarant
'
'        CargarDescripcionGarantia txtNroGarantia.Text
'        CmdAceptar.Visible = True
'        CmdAceptar.Enabled = True
'    End If
'End If

End Sub
