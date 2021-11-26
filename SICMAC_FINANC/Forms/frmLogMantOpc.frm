VERSION 5.00
Begin VB.Form frmLogMantOpc 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2085
   ClientLeft      =   1725
   ClientTop       =   3645
   ClientWidth     =   6915
   Icon            =   "frmLogMantOpc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboOpcion 
      Height          =   315
      Left            =   1740
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1050
      Width           =   1875
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   1890
      TabIndex        =   3
      Top             =   1545
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   3885
      TabIndex        =   4
      Top             =   1545
      Width           =   1305
   End
   Begin VB.TextBox txtDescripcion 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1740
      TabIndex        =   1
      Top             =   630
      Width           =   4155
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Unidad :"
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
      Height          =   210
      Index           =   2
      Left            =   465
      TabIndex        =   7
      Top             =   1095
      Width           =   1170
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1740
      TabIndex        =   0
      Top             =   240
      Width           =   1860
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Descripción :"
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
      Height          =   210
      Index           =   1
      Left            =   465
      TabIndex        =   6
      Top             =   675
      Width           =   1170
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Código :"
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
      Height          =   210
      Index           =   0
      Left            =   465
      TabIndex        =   5
      Top             =   285
      Width           =   1125
   End
End
Attribute VB_Name = "frmLogMantOpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim psFrmTpo As String
Dim psFrmOpc As String
Dim psCodigo As String
Dim psCodSec As String
Dim nRpta As Variant

Public Function Inicio(ByVal psFrmTipo As String, ByVal psFrmOpcion As String, _
Optional ByVal psFrmCodigo As String = "", Optional ByVal psFrmCodSec As String = "") As Variant
'De que Formulario es llamado
psFrmTpo = psFrmTipo
'Para que opción (Ingreso, Mantenimiento)
psFrmOpc = psFrmOpcion
'Un dato a procesar (Código)
psCodigo = psFrmCodigo
psCodSec = psFrmCodSec
Me.Show 1
Inicio = nRpta
End Function

Private Sub CmdAceptar_Click()
    Dim clsDBS As DLogBieSer
    Dim clsDGnral As DLogGeneral
    
    Dim sCod As String, sDes As String
    Dim nOpc As Integer
    Dim sActualiza As String
    Dim nSto As Currency
    
    sCod = Trim(lblCodigo.Caption)
    sDes = Trim(Replace(txtDescripcion.Text, "'", "", , , vbTextCompare))
    nOpc = Val(Right(cboOpcion.Text, 1))
    nSto = 0
    
    If psFrmTpo = "1" Then
        If nOpc = 0 Then
            MsgBox "Determine la Unidad de Medida", vbInformation, " Aviso "
            Exit Sub
        End If
    End If
    
    If sDes = "" Then
        If psFrmTpo = "0" Then
            If psFrmOpc = "1" Then
                MsgBox "Ingrese el nombre", vbInformation, " Aviso "
            End If
        Else
            MsgBox "Ingrese la descripción ", vbInformation, " Aviso "
        End If
        Exit Sub
    End If
    
    If psFrmTpo = "0" Then
        If psFrmOpc = "1" Then
            'BUSQUEDAS
            nRpta = UCase(txtDescripcion.Text)
        End If
    ElseIf psFrmTpo = "1" Then
        'MANTENIMIENTO DE BIENES/SERVICIOS
        Set clsDBS = New DLogBieSer
        If psFrmOpc = "1" Then
            'Ingreso
            If MsgBox("¿ Estás seguro de agregar " & sDes & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                nRpta = clsDBS.GrabaBS(sCod, sDes, nOpc, nSto, sActualiza)
            End If
        ElseIf psFrmOpc = "2" Then
            'Mantenimiento
            If MsgBox("¿ Estás seguro de Modificar " & sDes & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                nRpta = clsDBS.ModificaBS(sCod, sDes, nOpc, nSto, sActualiza)
            End If
        Else
            cmdAceptar.Enabled = False
            MsgBox "Opción de formulario no reconocida", vbInformation, " Aviso"
            Exit Sub
        End If
        Set clsDBS = Nothing
    ElseIf psFrmTpo = "2" Then
        'MANTENIMIENTO DE PARAMETROS
        Set clsDGnral = New DLogGeneral
        
        If psFrmOpc = "1" Then
            'Ingreso
            If MsgBox("¿ Estás seguro de agregar " & sDes & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                nRpta = clsDGnral.ConsGraba(psCodSec, sCod, sDes)
            End If
        ElseIf psFrmOpc = "2" Then
            'Mantenimiento
            If MsgBox("¿ Estás seguro de Modificar " & sDes & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                nRpta = clsDGnral.ConsModifi(psCodSec, sCod, sDes)
            End If
        Else
            cmdAceptar.Enabled = False
            MsgBox "Opción de formulario no reconocida", vbInformation, " Aviso"
            Exit Sub
        End If
        Set clsDGnral = Nothing
    Else
        MsgBox "Opción no reconocida", vbInformation, " Aviso"
    End If
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    If psFrmTpo = "0" Then
        If psFrmOpc = "1" Then
            'BUSQUEDAS
            nRpta = ""
        End If
    Else
        nRpta = 3
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Dim clsDBS As DLogBieSer
    Dim clsDGnral As DLogGeneral
    Dim rs As ADODB.Recordset
    Dim nCont  As Integer
    
    Call CentraForm(Me)
    If psFrmTpo = "0" Then
        If psFrmOpc = "1" Then
            'BUSQUEDAS
            Me.Caption = "Búsqueda..."
            lblEtiqueta(0).Visible = False
            lblEtiqueta(2).Visible = False
            lblCodigo.Visible = False
            cboOpcion.Visible = False
            lblEtiqueta(1).Caption = "Nombre"
            cmdAceptar.Top = cmdAceptar.Top - 300
            cmdCancelar.Top = cmdCancelar.Top - 300
            Me.Height = Me.Height - 300
        End If
    ElseIf psFrmTpo = "1" Then
        'MANTENIMIENTO DE BIENES/SERVICIOS
        Set clsDBS = New DLogBieSer
        Set rs = New ADODB.Recordset
        Set clsDGnral = New DLogGeneral
        
        Set rs = clsDGnral.CargaConstante(gUnidadMedida)
        For nCont = 1 To rs.RecordCount
            cboOpcion.AddItem rs!cConsDescripcion & Space(40) & rs!cConsValor
            rs.MoveNext
        Next
        
        If psFrmOpc = "1" Then
            'Ingreso
            Me.Caption = "Ingreso de Bien/Servicio"
            lblCodigo.Caption = clsDBS.GeneraBSCodNue(psCodigo)
        ElseIf psFrmOpc = "2" Then
            'Mantenimiento
            Me.Caption = "Mantenimiento de Bien/Servicio"
            Set rs = clsDBS.CargaBS(BsUnRegistro, psCodigo)
            If rs.RecordCount = 1 Then
                With rs
                    lblCodigo.Caption = !cBSCod
                    txtDescripcion.Text = !cBSDescripcion
                    For nCont = 1 To cboOpcion.ListCount
                        cboOpcion.ListIndex = nCont - 1
                        If Right(cboOpcion.Text, 1) = Right(!cConsUnidad, 1) Then
                            Exit For
                        End If
                    Next
                End With
            Else
                Set rs = Nothing
                MsgBox "Problemas al cargar bien/servicio", vbInformation, " Aviso"
                Exit Sub
            End If
            
        Else
            MsgBox "Opción del formulario no reconocida", vbInformation, "Aviso"
            Exit Sub
        End If
        Set clsDGnral = Nothing
        Set clsDBS = Nothing
        Set rs = Nothing
    ElseIf psFrmTpo = "2" Then
        'MANTENIMIENTO DE PARAMETROS
        lblEtiqueta(2).Visible = False
        cboOpcion.Visible = False
        cmdAceptar.Top = cmdAceptar.Top - 300
        cmdCancelar.Top = cmdCancelar.Top - 300
        Me.Height = Me.Height - 300
        Set clsDGnral = New DLogGeneral
        
        If psFrmOpc = "1" Then
            'Ingreso
            Me.Caption = "Ingreso de Parámetros"
            'lblCodigo.Caption = clsDBS.GeneraBSCodNue(psCodigo)
            lblCodigo.Caption = clsDGnral.ConsCodNue(psCodSec)
        ElseIf psFrmOpc = "2" Then
            'Mantenimiento
            Me.Caption = "Mantenimiento Parámetros"
            Set rs = New ADODB.Recordset
            'Set rs = clsDBS.CargaBS(BsUnRegistro, psCodigo)
            Set rs = clsDGnral.ConsCarga(psCodSec, psCodigo)
            If rs.RecordCount = 1 Then
                With rs
                    lblCodigo.Caption = !cConsValor
                    txtDescripcion.Text = !cConsDescripcion
                End With
            Else
                Set rs = Nothing
                MsgBox "Problemas al cargar información", vbInformation, " Aviso"
                Exit Sub
            End If
            Set rs = Nothing
        Else
            MsgBox "Opción del formulario no reconocida", vbInformation, "Aviso"
            Exit Sub
        End If
        Set clsDGnral = Nothing
    Else
        cmdAceptar.Enabled = False
        MsgBox "Opción no reconocida", vbInformation, " Aviso"
    End If
End Sub

