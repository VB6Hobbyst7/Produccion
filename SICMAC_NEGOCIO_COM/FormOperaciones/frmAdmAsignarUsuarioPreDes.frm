VERSION 5.00
Begin VB.Form frmAdmAsignarUsuarioPreDes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignar Usuarios - Pre Desembolso"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7995
   Icon            =   "frmAdmAsignarUsuarioPreDes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fr_Datos 
      Caption         =   "Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.CommandButton cmd_cancelar 
         Caption         =   "Cancelar"
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
         Left            =   6840
         TabIndex        =   11
         Top             =   3480
         Width           =   975
      End
      Begin VB.Frame fr_Buscar 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   7695
         Begin VB.TextBox txt_AgeActual 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2160
            TabIndex        =   13
            Top             =   360
            Width           =   2415
         End
         Begin VB.CommandButton cmd_Nuevo 
            Caption         =   "Nuevo"
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
            Left            =   6720
            TabIndex        =   10
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmd_Buscar 
            Caption         =   "Buscar"
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
            Left            =   4680
            TabIndex        =   8
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txt_usuario 
            Height          =   285
            Left            =   720
            TabIndex        =   7
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lbl_AgeActual 
            Caption         =   "Agencia Actual"
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
            Left            =   2160
            TabIndex        =   12
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label lbl_usuario 
            Caption         =   "Usuario"
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
            Left            =   720
            TabIndex        =   9
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmd_Guardar 
         Caption         =   "Guardar"
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
         Left            =   5760
         TabIndex        =   5
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton cmd_Quitar 
         Caption         =   "Quitar"
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
         TabIndex        =   4
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton cmd_Asignar 
         Caption         =   "<--"
         Height          =   375
         Left            =   3520
         TabIndex        =   3
         ToolTipText     =   "Asignar Agencias"
         Top             =   2040
         Width           =   495
      End
      Begin SICMACT.FlexEdit fe_AgenciaAsignados 
         Height          =   2385
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   4207
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Agencias Asignadas-nCodAge-New"
         EncabezadosAnchos=   "300-2500-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X"
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-R"
         FormatosEdit    =   "0-0-0-3"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit fe_Agencia 
         Height          =   2385
         Left            =   4080
         TabIndex        =   2
         Top             =   1080
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   4207
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Agencias a Asignar--nCodAge"
         EncabezadosAnchos=   "300-2500-500-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-3"
         ListaControles  =   "0-0-4-4"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-L"
         FormatosEdit    =   "0-0-0-0"
         AvanceCeldas    =   1
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmAdmAsignarUsuarioPreDes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Asignar_Click()
    Dim cAgeCod_Sel As String
    Dim cAgeCod_Asig As String
    Dim cAgeCod_New As String
    Dim i As Integer
    
    cAgeCod_Sel = ""
    cAgeCod_Asig = ""
    
    For i = 1 To fe_AgenciaAsignados.Rows - 1
        If fe_AgenciaAsignados.TextMatrix(i, 3) = "1" Then
            cAgeCod_New = cAgeCod_New + fe_AgenciaAsignados.TextMatrix(i, 2) + ","
        Else
            cAgeCod_Asig = cAgeCod_Asig + fe_AgenciaAsignados.TextMatrix(i, 2) + ","
        End If
    Next i
    
    For i = 1 To fe_Agencia.Rows - 1
        If fe_Agencia.TextMatrix(i, 2) = "." Then
            cAgeCod_Sel = cAgeCod_Sel + fe_Agencia.TextMatrix(i, 3) + "," + cAgeCod_New
        End If
    Next i
    
    If Valida(2) = False Then Exit Sub
    
    Call AsignaAgregar(cAgeCod_Sel, cAgeCod_Asig)
    Call HabilitaControles(3, True)
    fe_AgenciaAsignados.col = 1
    fe_AgenciaAsignados.row = 1
End Sub

Private Sub AsignaAgregar(ByVal pcAgeCod_Agr As String, ByVal pcAgeCod_Asig As String)
    Dim objAsigAgre As COMDCredito.DCOMCredito
    Dim rsAsigAgre As ADODB.Recordset
    Dim i As Integer
    Set objAsigAgre = New COMDCredito.DCOMCredito
    
    Set rsAsigAgre = objAsigAgre.AdmCred_AsignaAgencia(pcAgeCod_Asig, pcAgeCod_Agr)
    If Not (rsAsigAgre.BOF And rsAsigAgre.EOF) Then
        LimpiaFlex fe_AgenciaAsignados
        For i = 1 To rsAsigAgre.RecordCount
            fe_AgenciaAsignados.AdicionaFila
                fe_AgenciaAsignados.TextMatrix(i, 1) = rsAsigAgre!cAgeDescripcion
                fe_AgenciaAsignados.TextMatrix(i, 2) = rsAsigAgre!cAgeCod
                fe_AgenciaAsignados.TextMatrix(i, 3) = rsAsigAgre!nNew
                
                If rsAsigAgre!nNew = 1 Then
                    fe_AgenciaAsignados.BackColorRow vbGreen
                Else
                    fe_AgenciaAsignados.BackColorRow vbWhite
                End If
            rsAsigAgre.MoveNext
        Next i
    End If
    
Set objAsigAgre = Nothing
RSClose rsAsigAgre
End Sub

Private Sub cmd_Buscar_Click()
    If Valida(1) = True Then
        If Valida(3) = False Then Exit Sub
            If CargaGrillaAsignarAgencia(Trim(txt_usuario), False) = True Then
                Call HabilitaControles(1, True)
                Call CargaGrillaAgencia
                Call MostrarControles(2)
            Else
                MsgBox "El usuario " & UCase(Trim(txt_usuario)) & " No tiene asignado ninguna agencia.", vbInformation, "Aviso"
                Call LimpiaControles_AmdCred
                Call HabilitaControles(1, False)
                Call HabilitaControles(2, True)
                Call MostrarControles(1)
                Exit Sub
            End If
    End If
End Sub

Private Sub cmd_Guardar_Click()
Dim bGuardar As Boolean
Dim oCont As COMNContabilidad.NCOMContFunciones
Dim cAgeCod_GAsig As String
Dim cMovNro As String
Dim i As Integer
Dim obj As COMDCredito.DCOMCredito
Set obj = New COMDCredito.DCOMCredito
Set oCont = New COMNContabilidad.NCOMContFunciones

If Valida(4) = False Then Exit Sub

cAgeCod_GAsig = ""
cMovNro = oCont.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)

For i = 1 To fe_AgenciaAsignados.Rows - 1
    cAgeCod_GAsig = cAgeCod_GAsig + fe_AgenciaAsignados.TextMatrix(i, 2) + ","
Next i

If MsgBox("Se registrara las [Agencias Asignadas] al usuario: " + UCase(Trim(txt_usuario)) + ", Desea Continuar??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    
    bGuardar = obj.AdmCred_RegistraUserAsigPreDes(Trim(Left(txt_AgeActual, 3)), cAgeCod_GAsig, cMovNro)
    
    If bGuardar = True Then
        MsgBox "Se Registro correctamente los datos...", vbInformation, "Aviso"
        Call CargaGrillaAgencia
        Call cmd_Buscar_Click
    Else
        MsgBox "Error, al registrar los datos...", vbInformation, "Aviso"
    End If
End If
End Sub

Private Sub cmd_Nuevo_Click()
    If Valida(1) = False Then Exit Sub
    If Valida(3) = False Then Exit Sub
    Call HabilitaControles(1, True)
    Call CargaGrillaAgencia
    Call MostrarControles(2)
    If CargaGrillaAsignarAgencia(Trim(txt_usuario), "", True) = True Then Exit Sub
End Sub

Private Sub cmd_Quitar_Click()
    fe_AgenciaAsignados.EliminaFila (fe_AgenciaAsignados.row)
End Sub

Private Sub cmd_cancelar_Click()
    Call LimpiaControles_AmdCred
    Call HabilitaControles(1, False)
    Call MostrarControles(1)
End Sub

Private Sub MostrarControles(ByVal pnCaso As Integer)
    Select Case pnCaso
        Case 1
            cmd_Buscar.Left = 1920
            cmd_Buscar.Top = 240
            lbl_AgeActual.Visible = False
            txt_AgeActual.Visible = False
        Case 2
            cmd_Buscar.Left = 4680
            cmd_Buscar.Top = 240
            lbl_AgeActual.Visible = True
            txt_AgeActual.Visible = True
    End Select
End Sub

Private Sub Form_Load()
    Call MostrarControles(1)
    Call HabilitaControles(1, False)
End Sub

Private Function Valida(ByVal pnOp As Integer) As Boolean
Dim i As Integer
Dim J As Integer

Dim obj As COMDCredito.DCOMCredito
Dim rs As ADODB.Recordset
Set obj = New COMDCredito.DCOMCredito

Valida = True

Select Case pnOp
    Case 1
        If txt_usuario = "" Then
            MsgBox "Ingrese usuario.", vbInformation, "Aviso"
            txt_usuario.SetFocus
            Call LimpiaControles_AmdCred
            Call HabilitaControles(1, False)
            Valida = False
            'Exit Function
        End If
    Case 2
        Dim nCantAge As Integer
        nCantAge = 0
        For i = 1 To fe_Agencia.Rows - 1
            If fe_Agencia.TextMatrix(i, 2) = "." Then
                nCantAge = nCantAge + 1
            End If
        Next i
        
        If nCantAge = 0 Then
            MsgBox "Seleccione al menos una Agencia", vbInformation, "Aviso"
            fe_Agencia.SetFocus
            fe_Agencia.col = 1
            fe_Agencia.row = 1
            Valida = False
            'Exit Function
        End If
                
        For i = 1 To fe_Agencia.Rows - 1
            If fe_Agencia.TextMatrix(i, 2) = "." Then
                For J = 1 To fe_AgenciaAsignados.Rows - 1
                    If fe_Agencia.TextMatrix(i, 3) = fe_AgenciaAsignados.TextMatrix(J, 2) Then
                        MsgBox "La [" & fe_Agencia.TextMatrix(i, 1) & "] esta asignada, no esta permitido duplicados", vbInformation, "Aviso"
                        Valida = False
                        Exit Function
                    End If
                Next J
            End If
        Next i
    Case 3
        Set rs = obj.AdmCred_ValidaCargo(Trim(txt_usuario))
        If Not (rs.BOF And rs.EOF) Then
            If rs!cMsg <> "0" Then
                MsgBox rs!cMsg, vbInformation, "Aviso"
                Valida = False
                Call LimpiaControles_AmdCred
                Call HabilitaControles(1, False)
                'Exit Function
            End If
        End If
    Case 4
        Set rs = obj.AdmCred_ValidaUserReg(gsCodUser)
        If Not (rs.BOF And rs.EOF) Then
            If rs!cMsg <> "0" Then
                MsgBox rs!cMsg, vbInformation, "Aviso"
                Valida = False
            End If
        End If
End Select

Set obj = Nothing
RSClose rs
End Function

Private Function CargaGrillaAsignarAgencia(ByVal pcUser As String, Optional ByVal cAgeAgrega As String = "", Optional ByVal cNew As Boolean = False) As Boolean
    Dim objAsig As COMDCredito.DCOMCredito
    Dim rsAsig As ADODB.Recordset
    Dim i As Integer
    Set objAsig = New COMDCredito.DCOMCredito

If cNew = True Then
    CargaGrillaAsignarAgencia = True
    Set rsAsig = objAsig.AdmCred_BusAgencia(pcUser, IIf(cNew = True, 1, 0))
    If Not (rsAsig.BOF And rsAsig.EOF) Then
        txt_AgeActual.Text = rsAsig!cAgeUserAct
    End If
Else
    CargaGrillaAsignarAgencia = True
    
    Set rsAsig = objAsig.AdmCred_BusAgencia(pcUser, IIf(cNew = True, 1, 0))
    If Not (rsAsig.BOF And rsAsig.EOF) Then
        LimpiaFlex fe_AgenciaAsignados
        For i = 1 To rsAsig.RecordCount
            fe_AgenciaAsignados.AdicionaFila
                fe_AgenciaAsignados.TextMatrix(i, 1) = rsAsig!cAgeDescripcion
                fe_AgenciaAsignados.TextMatrix(i, 2) = rsAsig!cAgeCod
                fe_AgenciaAsignados.TextMatrix(i, 3) = 0
                txt_AgeActual.Text = rsAsig!cAgeUserAct
                fe_AgenciaAsignados.BackColorRow vbWhite
            rsAsig.MoveNext
        Next i
    Else
        CargaGrillaAsignarAgencia = False
    End If
End If

Set objAsig = Nothing
RSClose rsAsig
End Function
Private Sub CargaGrillaAgencia()
    Dim objGri As COMDCredito.DCOMCredito
    Dim rsGri As ADODB.Recordset
    Dim i As Integer
    Set objGri = New COMDCredito.DCOMCredito
    
    Set rsGri = objGri.AdmCred_CargaAgencia
    LimpiaFlex fe_Agencia
    For i = 1 To rsGri.RecordCount
        fe_Agencia.AdicionaFila
        fe_Agencia.TextMatrix(i, 1) = rsGri!cAgeDescripcion
        fe_Agencia.TextMatrix(i, 2) = 0
        fe_Agencia.TextMatrix(i, 3) = rsGri!AgeCod
        rsGri.MoveNext
    Next i
    
Set objGri = Nothing
RSClose rsGri
End Sub

Private Sub CargaCombos()
'    Dim objCmb As COMDCredito.DCOMCredito
'    Dim rsComb As ADODB.Recordset
'    Dim i As Integer
'    Set objCmb = New COMDCredito.DCOMCredito
'
'    Set rsComb = objCmb.AdmCred_CargaAgencia
'
'    If Not (rsComb.BOF And rsComb.EOF) Then
'        For i = 1 To rsComb.RecordCount
'            cmb_Agencias.AddItem rsComb!cAgeDescripcion & Space(500) & rsComb!AgeCod
'            'cmb_Agencias.ItemData(cmb_Agencias.NewIndex) = "" & rsComb!nIdSeccion
'            rsComb.MoveNext
'        Next
'        cmb_Agencias.ListIndex = 0
'    End If
'Set objCmb = Nothing
'RSClose rsComb
End Sub

Private Sub HabilitaControles(ByVal pnOpc As Integer, ByVal pbCondicion As Boolean)
    Select Case pnOpc
        Case 1
            cmd_Nuevo.Enabled = False
            cmd_Quitar.Enabled = pbCondicion
            cmd_Guardar.Enabled = pbCondicion
            cmd_Asignar.Enabled = pbCondicion
        
            fe_AgenciaAsignados.Enabled = pbCondicion
            fe_Agencia.Enabled = pbCondicion
        Case 2
            cmd_Nuevo.Enabled = pbCondicion
        Case 3
            cmd_Guardar.Enabled = pbCondicion
    End Select
End Sub

Private Sub LimpiaControles_AmdCred()
    txt_usuario.SetFocus
    txt_usuario.Text = ""
    txt_AgeActual.Text = ""
    LimpiaFlex fe_AgenciaAsignados
    LimpiaFlex fe_Agencia
End Sub

Private Sub txt_usuario_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras2(KeyAscii)

If KeyAscii <> 0 Then
    KeyAscii = fgIntfMayusculas(KeyAscii)
    If KeyAscii = 13 Then
        EnfocaControl cmd_Buscar
    End If
End If

End Sub
