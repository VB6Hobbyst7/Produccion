VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBuscarUsuarioConvenio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar Usuario de Convenio"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10605
   Icon            =   "frmBuscarUsuarioConvenio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid dgUsuarioConvenio 
      Height          =   1815
      Left            =   1920
      TabIndex        =   8
      Top             =   840
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   3201
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "cNomCliente"
         Caption         =   "Usuario"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "cDoi"
         Caption         =   "DOI"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "cConcepto"
         Caption         =   "Concepto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "cCodCliente"
         Caption         =   "Codigo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   4199.811
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar por"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton cboCodigo 
         Caption         =   "Codig&o"
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
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton cboNombre 
         Caption         =   "&Nombre"
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton cboDOI 
         Caption         =   "&DOI"
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
         Left            =   105
         TabIndex        =   4
         Top             =   630
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
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
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
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
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtDatoBUscar 
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
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese dato a buscar: "
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
      Left            =   1920
      TabIndex        =   6
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmBuscarUsuarioConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim R As ADODB.Recordset
Dim strConvenio As String
Dim bResultado As Boolean
Dim ClsServicioRecaudoWS As COMDCaptaServicios.DCOMSrvRecaudoWS 'CTI1 TI-ERS027-2019
Dim verificaWS As Boolean 'CTI1 TI-ERS027-2019

Public Function Inicio(ByVal strConvenioParametro As String) As Recordset
    strConvenio = strConvenioParametro
    'CTI1 TI-ERS027-2019 begin
    Set ClsServicioRecaudoWS = New COMDCaptaServicios.DCOMSrvRecaudoWS
    verificaWS = ClsServicioRecaudoWS.VerificarConvenioRecaudoWebService(strConvenio)
    'CTI1 TI-ERS027-2019 end
    Me.Show 1
    
    Set Inicio = R
    Set R = Nothing
End Function

Private Sub cboCodigo_Click()
txtDatoBUscar.Text = Empty
txtDatoBUscar.SetFocus
End Sub

Private Sub cboDOI_Click()
txtDatoBUscar.Text = Empty
txtDatoBUscar.SetFocus
End Sub

Private Sub cboNombre_Click()
txtDatoBUscar.Text = Empty
txtDatoBUscar.SetFocus
End Sub

Private Sub cmdAceptar_Click()

    Dim codigo As String
    Dim ClsServicioRecaudo As COMDCaptaServicios.DCOMServicioRecaudo

    If R Is Nothing Then
        MsgBox "Seleccione un Usuario", vbInformation, "Aviso"
        Exit Sub
    Else
        If R.RecordCount = 0 Then
            MsgBox "Seleccione un Usuario", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    If verificaWS = False Then 'CTI1 TI-ERS027-2019 begin
        Set ClsServicioRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
        dgUsuarioConvenio.Col = 3
        codigo = dgUsuarioConvenio.Text
        Screen.MousePointer = 0
        Set R = ClsServicioRecaudo.getBuscarUsuarioRecaudo(, , codigo, strConvenio, 1)
    End If 'CTI1 TI-ERS027-2019 end
    bResultado = True
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Set R = Nothing
    'Set Persona = Nothing
    Unload Me
End Sub



Private Sub Form_Load()
     dgUsuarioConvenio.MarqueeStyle = dbgHighlightRow
     bResultado = False
End Sub



Private Sub Form_Unload(Cancel As Integer)

    If Not bResultado Then
        Set R = Nothing
        verificaWS = False 'CTI1 TI-ERS027-2019 end
    End If

End Sub

Private Sub txtDatoBUscar_KeyPress(KeyAscii As Integer)
    
Dim ClsServicioRecaudo As COMDCaptaServicios.DCOMServicioRecaudo

If KeyAscii = 13 Then
    If Len(Trim(txtDatoBUscar.Text)) = 0 Then
         MsgBox "Falta Ingresar los datos para la busqueda", vbInformation, "Aviso"
         Exit Sub
    End If
    Screen.MousePointer = 11
    Set ClsServicioRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
    
    If verificaWS = False Then 'CTI1 TI-ERS027-2019 begin
        If cboNombre.value Then
            Set R = ClsServicioRecaudo.getBuscarUsuarioRecaudo(Trim(txtDatoBUscar.Text), , , strConvenio, 3)
        ElseIf cboDOI.value Then
            Set R = ClsServicioRecaudo.getBuscarUsuarioRecaudo(, Trim(txtDatoBUscar.Text), , strConvenio, 2)
        ElseIf cboCodigo.value Then
            Set R = ClsServicioRecaudo.getBuscarUsuarioRecaudo(, , Trim(txtDatoBUscar.Text), strConvenio, 1)
        End If
        
        'CTI7 INI SE MOVIO AL CONDICIONAL
        Set dgUsuarioConvenio.DataSource = R
            dgUsuarioConvenio.Refresh
            Screen.MousePointer = 0
        'CTI7 FIN
    Else 'CTI1 TI-ERS027-2019 begin
        Set ClsServicioRecaudoWS = New COMDCaptaServicios.DCOMSrvRecaudoWS
        Dim RBase As ADODB.Recordset 'CTI7
        Dim nCodConvenioRecaudoWS As Integer
        Dim urlSimaynas As String
        Set RBase = New ADODB.Recordset 'CTI7
        
        nCodConvenioRecaudoWS = ClsServicioRecaudoWS.ObtenerCodConvenioRecaudoWebService(strConvenio)
        urlSimaynas = Trim(LeeConstanteSist(708))

            Select Case nCodConvenioRecaudoWS
                Case 1 'ELUC
                Set R = ClsServicioRecaudoWS.ConsultarSuministroELUC(Trim(txtDatoBUscar.Text), urlSimaynas)
                Set RBase = ClsServicioRecaudoWS.ConsultarSuministroELUCFirstRecord(R) 'CTI7 ERS008-2020
            Case 2 'ELOR
                Set R = ClsServicioRecaudoWS.ConsultarSuministroELOR(Trim(txtDatoBUscar.Text), urlSimaynas) 'CTI1 ERS027-2019 ELOR
                Set RBase = ClsServicioRecaudoWS.ConsultarSuministroELORFirstRecord(R) 'CTI7 ERS008-2020
            Case Else
                    MsgBox "Codigo de Servicio Web Incorrecto", vbInformation, "Aviso"
        End Select
    'End If 'CTI1 TI-ERS027-2019 begin CTI7 COMENTO
        
        'CTI7 INI
        Set dgUsuarioConvenio.DataSource = RBase 'R
            dgUsuarioConvenio.Refresh
            Screen.MousePointer = 0
        'CTI7 FIN
    End If 'CTI1 TI-ERS027-2019 begin
    
    'CTI7 COMENTO ESTO
    'Set dgUsuarioConvenio.DataSource = R
    '    dgUsuarioConvenio.Refresh
    '    Screen.MousePointer = 0

    If R.RecordCount = 0 Then
         MsgBox "No se Encontraron Datos", vbInformation, "Aviso"
         txtDatoBUscar.SetFocus
         cmdAceptar.Default = False
    Else
         cmdAceptar.Default = True
         dgUsuarioConvenio.SetFocus
         
    End If
    
Else
     KeyAscii = Letras(KeyAscii)
     cmdAceptar.Default = False
End If
End Sub