VERSION 5.00
Begin VB.Form frmLogMantOpc 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2550
   ClientLeft      =   1725
   ClientTop       =   3645
   ClientWidth     =   6930
   Icon            =   "frmLogMantOpc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5685
      TabIndex        =   8
      Top             =   2100
      Width           =   1200
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   4440
      TabIndex        =   7
      Top             =   2100
      Width           =   1200
   End
   Begin VB.Frame fraProducto 
      Appearance      =   0  'Flat
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1995
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   6840
      Begin VB.TextBox txtStockMin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5520
         TabIndex        =   18
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cboTipoDep 
         Height          =   315
         Left            =   1065
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1193
         Width           =   2775
      End
      Begin VB.CheckBox chkCorrela 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "&Correlativo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   240
         Left            =   2625
         TabIndex        =   11
         ToolTipText     =   "Indica si el bien bien esta vigente en su uso o no lo esta. Si no lo esta no lo muestra en el regsitro de requerimientos"
         Top             =   1620
         Width           =   1245
      End
      Begin VB.CheckBox chkVigente 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "&Vigente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   255
         TabIndex        =   9
         ToolTipText     =   "Indica si el bien bien esta vigente en su uso o no lo esta. Si no lo esta no lo muestra en el regsitro de requerimientos"
         Top             =   1620
         Width           =   975
      End
      Begin VB.TextBox txtPerDep 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5535
         TabIndex        =   14
         Top             =   1215
         Width           =   1185
      End
      Begin VB.CheckBox chkVerifica 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "&Verifica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   4545
         TabIndex        =   12
         ToolTipText     =   "Indica si el bien debe ser verificado por un area tecnica"
         Top             =   1620
         Width           =   975
      End
      Begin VB.CheckBox chkContiene 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "&Es Parte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   5640
         TabIndex        =   13
         ToolTipText     =   "Indica si el bien forma parte de otro"
         Top             =   1620
         Width           =   1050
      End
      Begin VB.CheckBox chkSerie 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "&Serie"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   255
         Left            =   1770
         TabIndex        =   10
         ToolTipText     =   "Indica si el bien tiene o no Numero de Serie"
         Top             =   1620
         Width           =   765
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1065
         TabIndex        =   2
         Top             =   495
         Width           =   5670
      End
      Begin VB.ComboBox cboOpcion 
         Height          =   315
         Left            =   1065
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   825
         Width           =   2775
      End
      Begin VB.Label lblEtiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Stock Min :"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   5
         Left            =   4620
         TabIndex        =   19
         Top             =   870
         Width           =   810
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Tipo Dep :"
         ForeColor       =   &H8000000D&
         Height          =   210
         Index           =   4
         Left            =   150
         TabIndex        =   17
         Top             =   1245
         Width           =   840
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Periodos Deprecia :"
         ForeColor       =   &H8000000D&
         Height          =   210
         Index           =   3
         Left            =   4050
         TabIndex        =   15
         Top             =   1245
         Width           =   1515
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Código :"
         ForeColor       =   &H8000000D&
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   225
         Width           =   690
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Descripción :"
         ForeColor       =   &H8000000D&
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   540
         Width           =   1170
      End
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1065
         TabIndex        =   4
         Top             =   195
         Width           =   2160
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Unidad :"
         ForeColor       =   &H8000000D&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   870
         Width           =   840
      End
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
'ARLO 20170125******************
Dim objPista As COMManejador.Pista
'*******************************

Public Function Inicio(ByVal psFrmTipo As String, ByVal psFrmOpcion As String, _
        Optional ByVal psFrmCodigo As String = "", Optional ByVal psFrmCodSec As String = "", Optional pbSerie As Boolean = False, Optional pbVerifica As Boolean = False, Optional pbContiene As Boolean = False) As Variant
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


Private Sub cboOpcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Mid(psCodigo, 2, 1) = "2" Then
            Me.txtPerDep.SetFocus
        Else
            Me.chkSerie.SetFocus
        End If
    End If
End Sub

Private Sub CmdAceptar_Click()
    Dim clsDBS As DLogBieSer
    Dim clsDGnral As DLogGeneral
    
    Dim sCod As String, sDes As String
    Dim nOpc As Integer
    Dim sActualiza As String
    Dim nSto As Currency
    
    sCod = Trim(lblCodigo.Caption)
    sDes = Trim(Replace(txtDescripcion.Text, "'", "", , , vbTextCompare))
    nOpc = Val(Right(cboOpcion.Text, 3))
    nSto = 0
    
    If Left(lblCodigo.Caption, 3) = "112" Then
        If MsgBox("Ud. esta guardando un Activo Fijo. Se debe escoger numero de serie y correlativo según sea el caso.", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
        If Len(txtPerDep.Text) = 0 Or txtPerDep.Text = "0" Then
            MsgBox "Debe ingresar un dato valido para la los periodos de depreciación.", vbInformation, "Aviso"
            txtPerDep.SetFocus
            Exit Sub
        End If
        chkSerie.value = 1
    ElseIf Left(lblCodigo.Caption, 3) = "113" Then
        If MsgBox("Ud. esta guardando un Bien No depreciable. Se debe escoger numero de serie y correlativo según sea el caso.", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
        chkSerie.value = 1
    End If
    
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
                nRpta = clsDBS.GrabaBS(sCod, sDes, nOpc, Me.chkSerie.value, Me.chkVerifica.value, Me.chkContiene.value, Me.chkVigente.value, Me.chkCorrela.value, Me.txtPerDep.Text, Right(Me.cboTipoDep.Text, 3), sActualiza, IIf(Me.txtStockMin.Text = "", 0, Me.txtStockMin.Text))
                'ARLO 20170125
                gsOpeCod = LogPistaEntraSalidaBienes
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Agrego el Bien Cod : " & sCod & " | " & sDes
                Set objPista = Nothing
                '***********
            End If
        ElseIf psFrmOpc = "2" Then
            'Mantenimiento
            If MsgBox("¿ Estás seguro de Modificar " & sDes & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                nRpta = clsDBS.ModificaBS(sCod, sDes, nOpc, Me.chkSerie.value, Me.chkVerifica.value, Me.chkContiene.value, Me.chkVigente.value, Me.chkCorrela.value, Me.txtPerDep.Text, Right(Me.cboTipoDep.Text, 3), sActualiza, IIf(Me.txtStockMin.Text = "", 0, Me.txtStockMin.Text))
                'ARLO 20170125
                gsOpeCod = LogPistaEntraSalidaBienes
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "2", "Modifico el Bien Cod : " & sCod & " | " & sDes
                Set objPista = Nothing
                '***********
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
    
    If Mid(psCodigo, 3, 1) = "2" Then
        Me.lblEtiqueta(3).Visible = True
        Me.txtPerDep.Visible = True
        Me.cboTipoDep.Visible = True
        Me.lblEtiqueta(4).Visible = True
    Else
        Me.lblEtiqueta(3).Visible = False
        Me.txtPerDep.Visible = False
        Me.cboTipoDep.Visible = False
        Me.lblEtiqueta(4).Visible = False
    End If
        
    Call CentraForm(Me)
    If psFrmTpo = "0" Then
        If psFrmOpc = "1" Then
            'BUSQUEDAS
            Me.Caption = "Búsqueda..."
            lblEtiqueta(0).Visible = False
            lblEtiqueta(2).Visible = False
            lblEtiqueta(1).Visible = False
            lblEtiqueta(3).Visible = False
            lblCodigo.Visible = False
            cboOpcion.Visible = False
            
            lblEtiqueta(1).Caption = "Nombre"
            'cmdAceptar.Top = cmdAceptar.Top - 300
            'cmdCancelar.Top = cmdCancelar.Top - 300
            'Me.Height = Me.Height - 300
            Me.chkContiene.Visible = False
            Me.chkCorrela.Visible = False
            Me.chkVerifica.Visible = False
            Me.chkSerie.Visible = True
        
        End If
    ElseIf psFrmTpo = "1" Then
        'MANTENIMIENTO DE BIENES/SERVICIOS
        Set clsDBS = New DLogBieSer
        Set rs = New ADODB.Recordset
        Set clsDGnral = New DLogGeneral
        
        Set rs = clsDGnral.CargaConstante(gUnidadMedida)
        For nCont = 1 To rs.RecordCount
            cboOpcion.AddItem rs!cConsDescripcion & Space(80) & rs!nConsValor
            rs.MoveNext
        Next
        
        Set rs = clsDGnral.CargaConstante(5062)
        cboTipoDep.Clear
        For nCont = 1 To rs.RecordCount
            cboTipoDep.AddItem rs!cConsDescripcion & Space(80) & rs!nConsValor
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
                    Me.chkSerie.value = IIf(!Serie, 1, 0)
                    Me.chkVerifica.value = IIf(!Verificable, 1, 0)
                    Me.chkContiene.value = IIf(!Contenido, 1, 0)
                    Me.chkVigente.value = IIf(!Vigente, 1, 0)
                    Me.chkCorrela.value = IIf(!Correlativo, 1, 0)
                    Me.txtStockMin.Text = IIf(!StockMin = "", 0, !StockMin)
                    
                    If Mid(psCodigo, 3, 1) = "2" Then
                        Me.txtPerDep.Text = !Per_Depre
                   End If
                    
                    For nCont = 1 To cboOpcion.ListCount
                        cboOpcion.ListIndex = nCont - 1
                        If Val(Right(cboOpcion.Text, 3)) = Val(Right(!cConsUnidad, 3)) Then
                            Exit For
                        End If
                    Next
                    
                    For nCont = 1 To cboTipoDep.ListCount
                        cboTipoDep.ListIndex = nCont - 1
                        If Val(Right(cboTipoDep.Text, 3)) = Val(Right(!Tipo, 3)) Then
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
        'cmdAceptar.Top = cmdAceptar.Top - 300
        'cmdCancelar.Top = cmdCancelar.Top - 300
        'Me.Height = Me.Height - 300
        Set clsDGnral = New DLogGeneral
        
        Me.chkContiene.Visible = False
        Me.chkCorrela.Visible = False
        Me.chkSerie.Visible = False
        Me.chkVigente.Visible = False
        Me.chkVerifica.Visible = False
        
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
                    lblCodigo.Caption = !nConsValor
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

Private Sub txtDescripcion_GotFocus()
    txtDescripcion.SelStart = 0
    txtDescripcion.SelLength = 50
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboOpcion.Enabled And cboOpcion.Visible Then
            Me.cboOpcion.SetFocus
        Else
            Me.cmdAceptar.SetFocus
        End If
    End If
End Sub

Private Sub txtPerDep_GotFocus()
    txtPerDep.SelStart = 0
    txtPerDep.SelLength = 50
End Sub

Private Sub txtPerDep_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        KeyAscii = NumerosEnteros(KeyAscii)
    Else
        Me.chkSerie.SetFocus
    End If
End Sub
