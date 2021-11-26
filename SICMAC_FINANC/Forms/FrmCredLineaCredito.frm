VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmCredLineaCredito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Líneas de Créditos - Registro MN"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13335
   Icon            =   "FrmCredLineaCredito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   13335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   9530
      TabIndex        =   38
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   10765
      TabIndex        =   37
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   12000
      TabIndex        =   36
      Top             =   6360
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Líneas de Crédito"
      TabPicture(0)   =   "FrmCredLineaCredito.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cLineaCred"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblAbreviado"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtLineaCredito"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.Frame Frame2 
         Caption         =   "Condiciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   12735
         Begin VB.Frame Frame8 
            Caption         =   "Límite Periodo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   7680
            TabIndex        =   34
            Top             =   2880
            Width           =   4935
            Begin MSMask.MaskEdBox txtFechaMaxCred 
               Height          =   300
               Left            =   2280
               TabIndex        =   39
               Top             =   360
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label7 
               Caption         =   "Fecha máx. Créd.:"
               Height          =   255
               Left            =   240
               TabIndex        =   35
               Top             =   360
               Width           =   1455
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Calificación Interna"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   7680
            TabIndex        =   28
            Top             =   2040
            Width           =   4935
            Begin VB.CheckBox chkE 
               Caption         =   "E"
               Height          =   255
               Left            =   4320
               TabIndex        =   33
               Top             =   360
               Width           =   495
            End
            Begin VB.CheckBox chkD 
               Caption         =   "D"
               Height          =   255
               Left            =   3360
               TabIndex        =   32
               Top             =   360
               Width           =   495
            End
            Begin VB.CheckBox chkC 
               Caption         =   "C"
               Height          =   255
               Left            =   2280
               TabIndex        =   31
               Top             =   360
               Width           =   495
            End
            Begin VB.CheckBox chkB 
               Caption         =   "B"
               Height          =   255
               Left            =   1080
               TabIndex        =   30
               Top             =   360
               Width           =   495
            End
            Begin VB.CheckBox chkA 
               Caption         =   "A"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Rango de Tasas-TEM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   7680
            TabIndex        =   23
            Top             =   1080
            Width           =   4935
            Begin VB.TextBox txtRtHasta 
               Height          =   375
               Left            =   3360
               TabIndex        =   25
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox txtRtDesde 
               Height          =   375
               Left            =   1080
               TabIndex        =   24
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label5 
               Caption         =   "Hasta:"
               Height          =   255
               Left            =   2520
               TabIndex        =   27
               Top             =   405
               Width           =   735
            End
            Begin VB.Label Label6 
               Caption         =   "Desde:"
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   405
               Width           =   855
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Agencias"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   3840
            TabIndex        =   20
            Top             =   240
            Width           =   3735
            Begin VB.CheckBox ChkTodosAgencias 
               Caption         =   "Todos"
               Height          =   255
               Left            =   200
               TabIndex        =   21
               Top             =   240
               Width           =   2175
            End
            Begin MSComctlLib.ListView lvAgencia 
               Height          =   2505
               Left            =   120
               TabIndex        =   22
               Top             =   600
               Width           =   3435
               _ExtentX        =   6059
               _ExtentY        =   4419
               View            =   3
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Agencia"
                  Object.Width           =   1411
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Descripción"
                  Object.Width           =   6174
               EndProperty
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Rangos de Monto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   7680
            TabIndex        =   15
            Top             =   240
            Width           =   4935
            Begin VB.TextBox txtRmHasta 
               Height          =   375
               Left            =   3360
               TabIndex        =   19
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox txtRmDesde 
               Height          =   375
               Left            =   1080
               TabIndex        =   16
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label4 
               Caption         =   "Hasta:"
               Height          =   255
               Left            =   2520
               TabIndex        =   18
               Top             =   405
               Width           =   735
            End
            Begin VB.Label Label3 
               Caption         =   "Desde:"
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   405
               Width           =   855
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Tipo de Crédito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   3615
            Begin VB.CheckBox ChkTodosTCred 
               Caption         =   "Todos"
               Height          =   255
               Left            =   200
               TabIndex        =   14
               Top             =   240
               Width           =   2175
            End
            Begin MSComctlLib.ListView lvTipoCredito 
               Height          =   2505
               Left            =   120
               TabIndex        =   13
               Top             =   600
               Width           =   3315
               _ExtentX        =   5847
               _ExtentY        =   4419
               View            =   3
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Tipo Cred"
                  Object.Width           =   1411
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Descripción"
                  Object.Width           =   6174
               EndProperty
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Origen  de Fondos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   12735
         Begin VB.CheckBox cPlazo 
            Caption         =   "Corto Plazo"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6600
            TabIndex        =   42
            Top             =   720
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CheckBox cMoneda 
            Caption         =   "Soles"
            Height          =   375
            Left            =   4920
            TabIndex        =   41
            Top             =   720
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.ComboBox cboRecursosPropios 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   720
            Width           =   2295
         End
         Begin VB.OptionButton OptRp 
            Caption         =   "Recursos Propios"
            Height          =   495
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   2055
         End
         Begin VB.ComboBox cboPagare 
            Height          =   315
            Left            =   9000
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   240
            Width           =   3615
         End
         Begin VB.TextBox txtAdeudadoDes 
            Height          =   375
            Left            =   4440
            TabIndex        =   6
            Top             =   240
            Width           =   3735
         End
         Begin VB.OptionButton OptAdeudado 
            Caption         =   "Adeudado"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
         Begin Sicmact.TxtBuscar txtAdeudo 
            Height          =   375
            Left            =   2040
            TabIndex        =   5
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
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
            Enabled         =   0   'False
            Enabled         =   0   'False
            EnabledText     =   0   'False
         End
         Begin VB.Label Label2 
            Caption         =   "Pagaré:"
            Height          =   255
            Left            =   8280
            TabIndex        =   8
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.TextBox txtLineaCredito 
         Height          =   375
         Left            =   1920
         MaxLength       =   150
         TabIndex        =   2
         Top             =   480
         Width           =   6135
      End
      Begin VB.Label lblAbreviado 
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
         Left            =   8640
         TabIndex        =   43
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label cLineaCred 
         Caption         =   "Código"
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
         Left            =   8640
         TabIndex        =   40
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre de Linea:"
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
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmCredLineaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnOpcion As Integer
Dim psPersCodCaja As String
Dim lslineaCredito As String
Dim lsPersCod As String
Dim MatAgencias() As String
Dim gbEstado As Boolean
Dim gnNumDec As Integer
Public Event Change() 'ALPA 20140328
Public Event KeyPress(KeyAscii As Integer) 'ALPA 20140328
Dim ldCtaIFVenc As Date
Dim cPersCodMivienda As String
Dim lsPersNumero As String
Dim nLogicoGuardar As Integer
Dim fbOK As Boolean

Private Sub cboPagare_Click()
    If cboPagare.ListIndex = -1 Then
        Exit Sub
    End If
    Dim objLinea As DLineaCreditoV2
    Set objLinea = New DLineaCreditoV2
    Dim objRs As ADODB.Recordset
    Set objRs = New ADODB.Recordset
    cLineaCred.Caption = ""
    Dim oPersona As UPersona
    Dim frmBP As frmBuscaPersona

    Set objRs = objLinea.ObtenerSacarCodigoLinea(txtAdeudo.Text, Left(Trim(Right(cboPagare.Text, 10)), 2), Right(Trim(Right(cboPagare.Text, 10)), 7))
    If Not (objRs.BOF Or objRs.EOF) Then
        ldCtaIFVenc = objRs!dCtaIFVenc
        txtFechaMaxCred.Text = Format(ldCtaIFVenc, "DD/MM/YYYY")
        cLineaCred.Caption = Left(Trim(objRs!cCodLinCred), 4)
        If Mid(Trim(Right(cboPagare.Text, 10)), 5, 1) = "1" Then
            cMoneda.value = 1
            cLineaCred.Caption = cLineaCred.Caption & "1"
        Else
            cMoneda.value = 0
            cLineaCred.Caption = cLineaCred.Caption & "2"
        End If
        If cPlazo.value = 1 Then
            cLineaCred.Caption = Left(cLineaCred.Caption, 5) & "1"
        Else
            cLineaCred.Caption = Left(cLineaCred.Caption, 5) & "2"
        End If
    Else
        cLineaCred.Caption = ""
    End If
    If nLogicoGuardar = 1 Then
        If txtAdeudo.Text = cPersCodMivienda Then
            MsgBox "Las lineas de Crédito MIVIVIENDA requieren que se seleccione al prestatario beneficiado de la misma!", vbInformation, "Aviso!"
            Set oPersona = New UPersona
            Set frmBP = New frmBuscaPersona
            Set oPersona = frmBP.Inicio
            If Not oPersona Is Nothing Then
                If oPersona.sPersCod <> "" Then
                    lsPersNumero = oPersona.sPersCod
                    txtLineaCredito.Text = "MIVIVIENDA :::" & oPersona.sPersNombre
                End If
            Else
                lsPersNumero = ""
                txtLineaCredito.Text = ""
                cboPagare.ListIndex = -1

            End If
        Else
'            lsPersNumero = ""
'            txtLineaCredito.Text = ""
'            cboPagare.ListIndex = -1
        End If
    End If
End Sub

Private Sub cboRecursosPropios_Click()
    Dim psCodPagare As String
    If Trim(cboRecursosPropios.Text) = "" Then
'        MsgBox "Elegir tipo de recurso propio", vbCritical, "Aviso"
        Exit Sub
    End If
    If Trim(Right(cboRecursosPropios.Text, 20)) = 1 Then
        cLineaCred.Caption = "0111"
    ElseIf Trim(Right(cboRecursosPropios.Text, 20)) = 2 Then
        cLineaCred.Caption = "0112"
    Else
        cLineaCred.Caption = "0113"
    End If
    lsPersCod = psPersCodCaja
    If cMoneda.value = 1 Then
        cLineaCred.Caption = Left(cLineaCred.Caption, 2) & "1" & Mid(cLineaCred.Caption, 4, 1) & "1"
    Else
        cLineaCred.Caption = Left(cLineaCred.Caption, 2) & "2" & Mid(cLineaCred.Caption, 4, 1) & "2"
    End If
    If cPlazo.value = 1 Then
        cLineaCred.Caption = Left(cLineaCred.Caption, 5) & "1"
    Else
        cLineaCred.Caption = Left(cLineaCred.Caption, 5) & "2"
    End If
End Sub

Private Sub ChkTodosTCred_Click()
    Dim n As Integer
    If ChkTodosTCred.value = 1 Then
        For n = 1 To lvTipoCredito.ListItems.Count
            lvTipoCredito.ListItems(n).Checked = True
        Next n
    End If
    If ChkTodosTCred.value = 0 Then
        For n = 1 To lvTipoCredito.ListItems.Count
            lvTipoCredito.ListItems(n).Checked = False
        Next n
    End If
End Sub
Private Sub ChkTodosAgencias_Click()
    Dim n As Integer
    If ChkTodosAgencias.value = 1 Then
        For n = 1 To lvAgencia.ListItems.Count
            lvAgencia.ListItems(n).Checked = True
        Next n
    End If
    If ChkTodosAgencias.value = 0 Then
        For n = 1 To lvAgencia.ListItems.Count
            lvAgencia.ListItems(n).Checked = False
        Next n
    End If
End Sub

Private Sub CmdCancelar_Click()
    Call InicializarControles
End Sub

Private Sub cmdGuardar_Click()
    Dim objLinea As DLineaCreditoV2
    Set objLinea = New DLineaCreditoV2
    Dim lsFondosAbrev As String
    Dim lsFondosDesc As String
    Dim lsSubFondosAbrev As String
    Dim lsSubFondosDesc As String
    Dim lsSubProductoAbrev As String
    Dim lsSubProductoDesc As String
    Dim lsLineaCod As String
    Dim cTipo As String
    Dim cCodPagare As String
    Dim lsLineaAbrev1 As String
    Dim lsLineaAbrevFinal As String
    Dim n, m As Integer
    Dim lsAdeudados As String
    Dim lnContadorTiposCreditos As Integer
    Dim lnValidarNuevos As Integer
    If RecuperaAgencias = False Then Exit Sub
    If ValidarCreditos = False Then Exit Sub
    
    On Error GoTo ErrGuardar
    
    cmdGuardar.Enabled = False
    Screen.MousePointer = 11
    If Not Validar Then
        cmdGuardar.Enabled = True
        Screen.MousePointer = 0
        Exit Sub
    End If
	
	
'JOEP20211110 Mejora en el registro de Lineas correlativo
Dim i As Integer
Dim nCant As Integer
nCant = 0

 For i = 1 To lvTipoCredito.ListItems.Count
    If lvTipoCredito.ListItems(i).Checked = True Then
        nCant = nCant + 1
    End If
 Next i
 If nCant > 1 Then
    If MsgBox("Selecciono más de un tipo de Crédito. ¿Desea continuar?", vbYesNo + vbQuestion, "Aviso") = vbNo Then
        cmdGuardar.Enabled = True
        Screen.MousePointer = 0
        Exit Sub
    End If
 End If
 
nCant = 0
 For i = 1 To lvAgencia.ListItems.Count
    If lvAgencia.ListItems(i).Checked = True Then
        nCant = nCant + 1
    End If
 Next i
 If nCant > 1 Then
    If MsgBox("Selecciono más de una Agencia. ¿Desea continuar?", vbYesNo + vbQuestion, "Aviso") = vbNo Then
        cmdGuardar.Enabled = True
        Screen.MousePointer = 0
        Exit Sub
    End If
 End If
'JOEP20211110 Mejora en el registro de Lineas correlativo
    
     lsLineaCod = cLineaCred.Caption
     If Left(cLineaCred.Caption, 2) = "01" Then
        lsFondosAbrev = "RPCMM"
        lsFondosDesc = "CAJA MUNICIPAL DE AHORRO Y CREDITO MAYNAS SA"
        lsPersCod = psPersCodCaja
        If Trim(Right(IIf(cboRecursosPropios.Text = "", 0, cboRecursosPropios.Text), 20)) = 1 Then
            lsSubFondosAbrev = "RPPF"
            lsSubFondosDesc = "Plazo Fijo"
        ElseIf Trim(Right(IIf(cboRecursosPropios.Text = "", 0, cboRecursosPropios.Text), 20)) = 2 Then
            lsSubFondosAbrev = "RCTS"
            lsSubFondosDesc = "CTS"
        Else
            lsSubFondosAbrev = "RPAH"
            lsSubFondosDesc = "Ahorro"
        End If

        
    Else
        lsFondosAbrev = objLinea.RecuperaFondos(lsPersCod, 1)
        lsFondosDesc = objLinea.RecuperaFondos(lsPersCod, 2)
        lsSubFondosAbrev = objLinea.RecuperaSubFondos(Left(cLineaCred.Caption, 4), 1)
        lsSubFondosDesc = objLinea.RecuperaSubFondos(Left(cLineaCred.Caption, 4), 2)
    End If
    lsLineaAbrev1 = lsFondosAbrev & "-" & lsSubFondosAbrev & "-" & IIf(Mid(cLineaCred.Caption, 5, 1) = "1", "MN", "ME") & "-" & "LP"
    If Left(cLineaCred.Caption, 2) = "01" Then
        cTipo = "05"
        If Trim(Right(IIf(cboRecursosPropios.Text = "", 0, cboRecursosPropios.Text), 20)) = 1 Then 'PF
            cCodPagare = "01" & Mid(cLineaCred.Caption, 5, 1) & "1" & Mid(cLineaCred.Caption, 5, 1) & "1"
        ElseIf Trim(Right(IIf(cboRecursosPropios.Text = "", 0, cboRecursosPropios.Text), 20)) = 2 Then 'CTS
            cCodPagare = "01" & Mid(cLineaCred.Caption, 5, 1) & "2" & Mid(cLineaCred.Caption, 5, 1) & "1"
        Else 'Ahorro
            cCodPagare = "01" & Mid(cLineaCred.Caption, 5, 1) & "3" & Mid(cLineaCred.Caption, 5, 1) & "1"
        End If

    Else
        cTipo = Left(Trim(Right(cboPagare.Text, 10)), 2)
        cCodPagare = Right(Trim(Right(cboPagare.Text, 10)), 7)
    End If
        
    Dim psAdeudados  As String
    Dim psCorrelativo  As String
    Dim pbNuevo As Boolean
    
    Dim objRs As ADODB.Recordset
    Set objRs = New ADODB.Recordset
    
    Set objRs = objLinea.BuscarCodigoLinea(lsPersCod, cTipo, cCodPagare)
    'If Not (objRs.BOF Or objRs.BOF) And lnOpcion = 1 Then
    If (Not (objRs.BOF Or objRs.BOF)) Then
        If lnOpcion <> 2 Then
            MsgBox "El pagaré seleccionado ya se encuentra registrado en otra línea, favor seleccionar otro pagaré", vbCritical, "Aviso"
            cmdGuardar.Enabled = True
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    
	'JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito
    Dim cLineaCredCorre As String
    cLineaCredCorre = objLinea.CabCorrelativo(cLineaCred & "___0", lsPersCod)
	'JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito
	
	'Comento JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito	
    'psCorrelativo = objLinea.CorrelativoAdeudo(cLineaCred & "___01", lsPersCod, cTipo, cCodPagare)
	'Comento JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito
	
	'JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito
    psCorrelativo = objLinea.CorrelativoAdeudo(cLineaCredCorre, lsPersCod, cTipo, cCodPagare)
    'JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito
	
    If Trim(psCorrelativo) = "00" Then
        'Comento JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito
		'psCorrelativo = objLinea.Correlativo(cLineaCred & "___01")
		'Comento JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito
		
		'JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito
        psCorrelativo = objLinea.Correlativo(cLineaCredCorre)
        'JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito
		
        pbNuevo = True
    Else
        pbNuevo = False
'        NuevaLinea = ModificarLinea(5, psLineaCred, psDescripcion, _
'                    pbEstado, pnPlazoMax, pnPlazoMin, pnMontoMax, _
'                    pnMontoMin, psPersCod, pbPreferencial, pMatAgencias)
'
    End If
'    psCorrelativo = psAdeudados
'    psLineaCred = psLineaCred & psAdeudados
    lnContadorTiposCreditos = 0
    For n = 1 To lvTipoCredito.ListItems.Count
        lnContadorTiposCreditos = lnContadorTiposCreditos + 1
		
		'Comento JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito
        'cLineaCred.Caption = lsLineaCod & Right(lvTipoCredito.ListItems(n).Text, 3) & "01" & psCorrelativo
		'Comento JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito
		
		'JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito
        cLineaCred.Caption = lsLineaCod & Right(lvTipoCredito.ListItems(N).Text, 3) & Right(cLineaCredCorre, 2) & psCorrelativo
        'JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito
		
        lsSubProductoAbrev = objLinea.RecuperaProductosDeCredito(Right(lvTipoCredito.ListItems(n).Text, 3), 1)
        lsSubProductoDesc = objLinea.RecuperaProductosDeCredito(Right(lvTipoCredito.ListItems(n).Text, 3), 2)
        lsLineaAbrevFinal = lsLineaAbrev1 & "-" & lsSubProductoAbrev
        lblAbreviado = lsLineaAbrevFinal
        lsAdeudados = ""
        lnValidarNuevos = objLinea.ObtenerListaLineaCreditoVer(cLineaCred.Caption)
        If lvTipoCredito.ListItems(n).Checked = True Then
            If lnValidarNuevos = 0 Then
                Call objLinea.NuevaLinea(5, cLineaCred.Caption, lsLineaAbrevFinal, 1, 9999, 30, CDbl(txtRmHasta.Text), CDbl(txtRmDesde.Text), lsPersCod, lsFondosDesc, lsSubFondosDesc, lsSubProductoDesc, lsFondosAbrev, lsSubFondosAbrev, 1, MatAgencias, cTipo, cCodPagare, psCorrelativo)
            Else
                Call objLinea.ModificarLinea(5, cLineaCred.Caption, lsLineaAbrevFinal, 1, 9999, 30, CDbl(txtRmHasta.Text), CDbl(txtRmDesde.Text), lsPersCod, 1, MatAgencias, cTipo, cCodPagare)
            End If
            Call objLinea.InsertarLineaCredito(cLineaCred.Caption, txtLineaCredito.Text, lsPersCod, cTipo & cCodPagare, Trim(Right(IIf(cboRecursosPropios.Text = "", 0, cboRecursosPropios.Text), 20)), CDbl(txtRmDesde.Text), CDbl(txtRmHasta.Text), CDbl(txtRtDesde.Text), CDbl(txtRtHasta.Text), chkA.value, chkB.value, chkC.value, chkD.value, chkE.value, CDate(txtFechaMaxCred.Text), 1, lsPersNumero)
'            For M = 1 To lvAgencia.ListItems.Count
'                Call objLinea.InsertaLineaCreditoAgencia(cLineaCred.Caption & Trim(lsAdeudados), Right(lvAgencia.ListItems(M).Text, 2), IIf(lvAgencia.ListItems(M).Checked, 1, 0))
'            Next
        Else
            Call objLinea.ModificarLinea(5, cLineaCred.Caption, lsLineaAbrevFinal, 0, 9999, 30, CDbl(txtRmHasta.Text), CDbl(txtRmDesde.Text), lsPersCod, 1, MatAgencias, cTipo, cCodPagare)
            Call objLinea.InsertarLineaCredito(cLineaCred.Caption, txtLineaCredito.Text, lsPersCod, cTipo & cCodPagare, Trim(Right(IIf(cboRecursosPropios.Text = "", 0, cboRecursosPropios.Text), 20)), CDbl(txtRmDesde.Text), CDbl(txtRmHasta.Text), CDbl(txtRtDesde.Text), CDbl(txtRtHasta.Text), chkA.value, chkB.value, chkC.value, chkD.value, chkE.value, CDate(txtFechaMaxCred.Text), 0, lsPersNumero)
        End If
        'Call objLinea.InsertaLineaCreditoTipoCredito(Trim(Right(cboPagare.Text, 20)) & Trim(Right(IIf(OptAdeudado = True, txtAdeudo.Text, IIf(cboRecursosPropios.Text = "", 0, cboRecursosPropios.Text)), 20)), Right(lvTipoCredito.ListItems(n).Text, 3), IIf(lvTipoCredito.ListItems(n).Checked, 1, 0))
    Next n
    
    cmdGuardar.Enabled = True
    Screen.MousePointer = 0
'    If lnContadorTiposCreditos > 0 Then
        MsgBox "La línea de crédito se guardó correctamente", vbInformation, "Aviso"
        fbOK = True
        nLogicoGuardar = 0
        cLineaCred.Caption = Left(cLineaCred.Caption, 6)
        lblAbreviado.Caption = ""
        Call InicializarControles
        OptAdeudado.value = True
        Call OptAdeudado_Click
        If lnOpcion = 2 Then
        Unload Me
        End If

    Exit Sub
ErrGuardar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Function RecuperaAgencias() As Boolean

Dim nContAge As Integer
Dim I As Integer
    
ReDim MatAgencias(0)
nContAge = 0
RecuperaAgencias = True


For I = 1 To lvAgencia.ListItems.Count
    If lvAgencia.ListItems(I).Checked = True Then
        nContAge = nContAge + 1
        ReDim Preserve MatAgencias(nContAge)
        MatAgencias(nContAge - 1) = Right(lvAgencia.ListItems(I).Text, 2)
    End If
Next I
    
If nContAge = 0 Then
    MsgBox "Debe seleccionar por lo menos una Agencia", vbInformation, "Mensaje"
    RecuperaAgencias = False
End If
    
End Function

Function ValidarCreditos() As Boolean

Dim nContCre As Integer
Dim I As Integer
    
nContCre = 0
ValidarCreditos = True


For I = 1 To lvTipoCredito.ListItems.Count
    If lvTipoCredito.ListItems(I).Checked = True Then
        nContCre = nContCre + 1
    End If
Next I
    
If nContCre = 0 Then
    MsgBox "Debe seleccionar por lo menos un tipo de credito", vbInformation, "Mensaje"
    ValidarCreditos = False
End If
    
End Function
Private Function Validar() As Boolean
    Dim lsFecha As String
    If Trim(txtLineaCredito.Text) = "" Then
        MsgBox "Favor ingresar descripción de linea de credito", vbInformation, "Aviso"
        EnfocaControl txtLineaCredito
        Exit Function
    End If
    If OptAdeudado.value = True Then
        If Trim(txtAdeudo.Text) = "" Then
            MsgBox "Ud. debe elegir un adeudado", vbInformation, "Aviso"
            EnfocaControl txtAdeudo
            Exit Function
        End If
        If cboPagare.ListIndex = -1 Then
            MsgBox "Ud. debe seleccionar un pagaré", vbInformation, "Aviso"
            EnfocaControl cboPagare
            Exit Function
        End If
    Else
        If cboRecursosPropios.ListIndex = -1 Then
            MsgBox "Ud. debe seleccionar un tipo de RRPP", vbInformation, "Aviso"
            EnfocaControl cboRecursosPropios
            Exit Function
        End If
    End If
    If (Not IsNumeric(txtRmDesde.Text) Or Trim(txtRmDesde.Text) = "") Then
        MsgBox "Ud. debe ingresar Monto Inicial", vbInformation, "Aviso"
        EnfocaControl txtRmDesde
        Exit Function
    End If
    If (Not IsNumeric(txtRmHasta.Text) Or Trim(txtRmHasta.Text) = "") Then
        MsgBox "Ud. debe ingresar Monto Final", vbInformation, "Aviso"
        EnfocaControl txtRmHasta
        Exit Function
    End If
    If (Not IsNumeric(txtRtDesde.Text) Or Trim(txtRtDesde.Text) = "") Then
        MsgBox "Ud. debe ingresar Tasa Inicial", vbInformation, "Aviso"
        EnfocaControl txtRtDesde
        Exit Function
    End If
    If (Not IsNumeric(txtRtHasta.Text) Or Trim(txtRtHasta.Text) = "") Then
        MsgBox "Favor ingresar Tasa Final", vbInformation, "Aviso"
        EnfocaControl txtRtHasta
        Exit Function
    End If
    lsFecha = ValidaFecha(txtFechaMaxCred)
    If Len(lsFecha) > 0 Then
        MsgBox lsFecha, vbInformation, "Aviso"
        EnfocaControl txtFechaMaxCred
        Exit Function
    End If
    If chkA.value = False And chkB.value = False And chkC.value = False And chkD.value = False And chkE.value = False Then
        MsgBox "Ud. debe elegir al menos uno de las calificaciones", vbInformation, "Aviso"
        Exit Function
    End If
    If Len(Trim(cLineaCred.Caption)) <> 6 Then
        MsgBox "El pagaré no esta bien configurado, seleccionar otro y avisar al area de TI para su configuración", vbInformation, "Aviso"
        Exit Function
    End If
    If OptAdeudado.value = True Then
        If ldCtaIFVenc < CDate(txtFechaMaxCred.Text) Then
            MsgBox "La fecha de vencimiento debe ser menor o igual a " & Format(ldCtaIFVenc, "DD/MM/YYYY"), vbInformation, "Aviso"
            Exit Function
        End If
    End If
    If CDbl(txtRtDesde.Text) > CDbl(txtRtHasta.Text) Then
        MsgBox "La tasa inicial no debe ser mayor a la tasa final", vbInformation, "Aviso"
        Exit Function
    End If
    If CDbl(txtRmDesde.Text) > CDbl(txtRmHasta.Text) Then
        MsgBox "El monto inicial no debe ser mayor al monto final", vbInformation, "Aviso"
        Exit Function
    End If
    Validar = True
End Function


Private Sub cmdSalir_Click()
Unload Me
End Sub

Public Function registro(Optional ByVal pnOpcion As Integer = 1) As Boolean
    lnOpcion = pnOpcion
    Me.Caption = "Lineas de Crédito - Registro MN"
    Me.Show 1
    registro = fbOK
End Function

Public Function mantenimiento(Optional ByVal pnOpcion As Integer = 2, Optional ByVal pslineaCredito As String = "") As Boolean
    lnOpcion = pnOpcion
    lslineaCredito = pslineaCredito
    Me.Caption = "Lineas de Crédito - Mantenimiento MN"
    nLogicoGuardar = 0
    Call Mostrar
    cmdCancelar.Enabled = False
    Me.Show 1
    mantenimiento = fbOK
End Function

Private Sub Mostrar()
Dim objLinea As DLineaCreditoV2
Set objLinea = New DLineaCreditoV2
Dim lvItem As ListItem
Dim objRs As ADODB.Recordset
Set objRs = New ADODB.Recordset
Call DeshabilitarEdicion(False)
Set objRs = objLinea.ObtenerLineaCredito(lslineaCredito)
lsPersNumero = ""
If Not (objRs.BOF Or objRs.EOF) Then
    Do While Not objRs.EOF
    txtAdeudo.Text = objRs!cPersCod
    ldCtaIFVenc = Format(objRs!dCtaIFVenc, "DD/MM/YYYY")
    Call txtAdeudo_EmiteDatos
    cboPagare.ListIndex = IndiceListaCombo(cboPagare, Trim(objRs!cCodAdeudado))
    OptRp.value = IIf(objRs!nCodRRPP = 0, False, True)
    OptAdeudado.value = IIf(objRs!nCodRRPP = 0, True, False)
    txtLineaCredito.Text = objRs!cLineaCreditoDes
    cboRecursosPropios.ListIndex = IndiceListaCombo(cboRecursosPropios, objRs!nCodRRPP)
    txtRmDesde.Text = objRs!nRanMonDesde
    txtRmHasta.Text = objRs!nRanMonHasta
    txtRtDesde.Text = objRs!nRanTasDesde
    txtRtHasta.Text = objRs!nRanTasHasta
    chkA.value = objRs!bCalA
    chkB.value = objRs!bCalB
    chkC.value = objRs!bCalC
    chkD.value = objRs!bCalD
    chkE.value = objRs!bCalE
    
    lsPersNumero = objRs!cPersCodMivienda
    txtFechaMaxCred.Text = Format(objRs!dFechaMax, "DD/MM/YYYY")
    cLineaCred.Caption = Trim(Left(objRs!cLineaCreditoCod, 6))
    If Mid(objRs!cLineaCreditoCod, 5, 1) = 1 Then
        cMoneda.value = 1
    Else
        cMoneda.value = False
    End If
    If Mid(objRs!cLineaCreditoCod, 6, 1) = 1 Then
        cPlazo.value = 1
    Else
        cPlazo.value = False
    End If
    objRs.MoveNext
    Loop
End If
'txtAdeudo.rs = objRs


'Llenar Agencias
lvAgencia.ListItems.Clear
ChkTodosAgencias.value = 0
Set objRs = objLinea.ObtenerLineaCreditoAgencia(lslineaCredito)
If Not (objRs.BOF Or objRs.EOF) Then
    Do While Not objRs.EOF
       Set lvItem = lvAgencia.ListItems.Add(, , objRs!cCodigo)
       lvItem.SubItems(1) = objRs!CDESCRI
       'lvItem.SubItems(2) = IIf(objRs!nEstado = 1, True, False)
       If objRs!nEstado Then
            lvItem.Checked = True
       Else
            lvItem.Checked = False
       End If
       objRs.MoveNext
    Loop
    End If
RSClose objRs
'Tipo de Crédito
ChkTodosTCred.value = 0
lvTipoCredito.ListItems.Clear
Set objRs = objLinea.ObtenerLineaCreditoTipoCredito(lslineaCredito)
If Not (objRs.BOF Or objRs.EOF) Then
    Do While Not objRs.EOF
       Set lvItem = lvTipoCredito.ListItems.Add(, , objRs!cCodigo)
       lvItem.SubItems(1) = objRs!CDESCRI
       If objRs!nEstado Then
            lvItem.Checked = True
       Else
            lvItem.Checked = False
       End If
       objRs.MoveNext
    Loop
    End If
RSClose objRs
nLogicoGuardar = 1

End Sub

Private Sub LlenarCombo(ByRef pCombo As ComboBox, ByRef prs As ADODB.Recordset)
'    pRs.MoveFirst
    pCombo.Clear
    If (prs.BOF Or prs.EOF) Then
    Exit Sub
    End If
    Do While Not prs.EOF
        pCombo.AddItem prs!CDESCRI & Space(300) & prs!cCodigo
        prs.MoveNext
    Loop
End Sub

Private Sub cMoneda_Click()
    If cLineaCred.Caption = "" Then
        'MsgBox "Elegir tipo de recurso propio", vbCritical, "Aviso"
        Exit Sub
    End If
    cLineaCred.Caption = Left(cLineaCred.Caption, 4)
    If cMoneda.value = 1 Then
        cLineaCred.Caption = Left(cLineaCred.Caption, 4) & "1"
    Else
        cLineaCred.Caption = Left(cLineaCred.Caption, 4) & "2"
    End If
    If cPlazo.value = 1 Then
        cLineaCred.Caption = Left(cLineaCred.Caption, 5) & "1"
    Else
        cLineaCred.Caption = Left(cLineaCred.Caption, 5) & "2"
    End If
End Sub

Private Sub cPlazo_Click()
'    cPlazo.value = 1
    If cLineaCred.Caption = "" Then
        'MsgBox "Elegir tipo de recurso propio", vbCritical, "Aviso"
        Exit Sub
    End If
    cLineaCred.Caption = Left(cLineaCred.Caption, 4)
    If cMoneda.value = 1 Then
        cLineaCred.Caption = Left(cLineaCred.Caption, 4) & "1"
    Else
        cLineaCred.Caption = Left(cLineaCred.Caption, 4) & "2"
    End If
    If cPlazo.value = 1 Then
        cLineaCred.Caption = Left(cLineaCred.Caption, 5) & "1"
    Else
        cLineaCred.Caption = Left(cLineaCred.Caption, 5) & "2"
    End If
End Sub

Private Sub Form_Load()
Dim objLinea As DLineaCreditoV2
Set objLinea = New DLineaCreditoV2
Dim lvItem As ListItem
Dim oAge As New DActualizaDatosArea
psPersCodCaja = "1090100012521"
Dim objRs As ADODB.Recordset
Set objRs = New ADODB.Recordset

Set objRs = objLinea.ObtenerLineaCreditoAdeudado
txtAdeudo.rs = objRs
Set objRs = Nothing
Set objRs = objLinea.ObtenerLineaCreditoRRPP
Call LlenarCombo(cboRecursosPropios, objRs)

ChkTodosAgencias.value = 1
Set objRs = objLinea.ObtenerLineaCreditoAgencia("")
If objRs.EOF Then
   RSClose objRs
   MsgBox "No se definieron Agencias en el Sistema...Consultar con Sistemas", vbInformation, "Aviso"
   Exit Sub
End If
Do While Not objRs.EOF
   Set lvItem = lvAgencia.ListItems.Add(, , objRs!cCodigo)
   lvItem.SubItems(1) = objRs!CDESCRI
   lvItem.Checked = True
   objRs.MoveNext
Loop
RSClose objRs

ChkTodosTCred.value = 1
Set objRs = objLinea.ObtenerLineaCreditoTipoCredito("")
If objRs.EOF Then
   RSClose objRs
   MsgBox "No se definieron los tipos de creditos en la agencia...Consultar con Sistemas", vbInformation, "Aviso"
   Exit Sub
End If
Do While Not objRs.EOF
   Set lvItem = lvTipoCredito.ListItems.Add(, , objRs!cCodigo)
   lvItem.SubItems(1) = objRs!CDESCRI
   lvItem.Checked = True
   lvItem.Checked = True
   objRs.MoveNext
Loop
RSClose objRs
OptAdeudado.value = True
OptRp.value = True
'OptRp.value = False
Call Habilitar(True)
OptAdeudado.value = True
cPlazo.value = 1
cPlazo.Enabled = True
cMoneda.Enabled = False
nLogicoGuardar = 1
cPersCodMivienda = "1093300012530"
Call CambiaTamañoCombo(cboPagare, 300)
fbOK = False
End Sub

Private Sub Habilitar(ByVal lg1 As Boolean)
    txtAdeudo.Enabled = lg1
    txtAdeudadoDes.Enabled = False
    cboPagare.Enabled = lg1
    cboRecursosPropios.Enabled = Not lg1
    cMoneda.Enabled = lg1
    cPlazo.Enabled = lg1
    If lnOpcion = 1 And OptRp.value = True Then
        cMoneda.Enabled = True
        cPlazo.Enabled = True
    End If
    If lnOpcion = 1 And OptAdeudado.value = True Then
        cMoneda.Enabled = False
        cPlazo.Enabled = True
    End If
End Sub

Private Sub OptAdeudado_Click()
    cboRecursosPropios.ListIndex = -1
    cLineaCred.Caption = ""
    Call Habilitar(True)
End Sub

Private Sub OptRp_Click()
    cboPagare.ListIndex = -1
    txtAdeudadoDes.Text = ""
    txtLineaCredito.Text = ""
    txtAdeudo.Text = ""
    cLineaCred.Caption = ""
    txtLineaCredito.Enabled = True
    Call Habilitar(False)
End Sub

Private Sub txtAdeudo_EmiteDatos()
Dim objLinea As DLineaCreditoV2
Set objLinea = New DLineaCreditoV2

Dim objRs As ADODB.Recordset
Set objRs = New ADODB.Recordset
Set objRs = objLinea.ObtenerLineaCreditoAdeudadoPagare(txtAdeudo.Text)
lsPersCod = txtAdeudo.Text
cLineaCred.Caption = ""
txtAdeudadoDes.Text = txtAdeudo.psDescripcion
Call LlenarCombo(cboPagare, objRs)
If txtAdeudo.Text = cPersCodMivienda Then
    'fraMivienda.Visible = True
    txtLineaCredito.Enabled = False
    lsPersNumero = ""
Else
'    fraMivienda.Visible = False
    txtLineaCredito.Enabled = True
    lsPersNumero = ""
End If
'cPersCodMivienda = "1093300012530"
End Sub
Private Sub DeshabilitarEdicion(ByVal lg1 As Boolean)
    OptAdeudado.Enabled = lg1
    OptRp.Enabled = lg1
    txtAdeudo.Enabled = lg1
    txtAdeudadoDes.Enabled = lg1
    cboPagare.Enabled = lg1
    cboRecursosPropios.Enabled = lg1
    cMoneda.Enabled = lg1
    cPlazo.Enabled = lg1
End Sub
Private Sub InicializarControles()
    txtRmDesde.Text = ""
    txtRmHasta.Text = ""
    txtRtDesde.Text = ""
    txtRtHasta.Text = ""
    chkA.value = 0
    chkB.value = 0
    chkC.value = 0
    chkD.value = 0
    chkE.value = 0
    cboPagare.ListIndex = -1
    txtAdeudadoDes.Text = ""
    txtAdeudo.Text = ""
    cboRecursosPropios.ListIndex = -1
    txtLineaCredito.Text = ""
    txtFechaMaxCred = "__/__/____"
    cLineaCred = ""
    cboPagare.Clear
    Call ChkTodosAgencias_Click
    Call ChkTodosTCred_Click
    txtLineaCredito.Enabled = True
    'cMoneda.value = 0
    'cPlazo.value = 0
End Sub
Private Function TienePunto(psCadena As String) As Boolean
If InStr(1, psCadena, ".", vbTextCompare) > 0 Then
    TienePunto = True
Else
    TienePunto = False
End If
End Function

Private Function NumDecimal(psCadena As String) As Integer
Dim lnPos As Integer
lnPos = InStr(1, psCadena, ".", vbTextCompare)
If lnPos > 0 Then
    NumDecimal = Len(psCadena) - lnPos
Else
    NumDecimal = 0
End If
End Function

Private Sub txtLineaCredito_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
End Sub

Private Sub txtLineaCredito_LostFocus()
    txtLineaCredito.Text = UCase(txtLineaCredito.Text)
End Sub

'Private Sub TxtClienteMivienda_EmiteDatos()
'    txtMIVIVIENDADescripcion.Text = TxtClienteMivienda.psDescripcion
'    txtLineaCredito.Text = "MIVIVIENDA :::" & TxtClienteMivienda.psDescripcion
'End Sub

Private Sub txtRmDesde_Change()
txtRmDesde.SelStart = Len(txtRmDesde)
gnNumDec = NumDecimal(txtRmDesde)
If gbEstado And txtRmDesde <> "" Then
    Select Case gnNumDec
        Case 0
                txtRmDesde = Format(txtRmDesde, "#,##0")
        Case 1
                txtRmDesde = Format(txtRmDesde, "#,##0.0")
        Case 2
                txtRmDesde = Format(txtRmDesde, "#,##0.00")
        Case 3
                txtRmDesde = Format(txtRmDesde, "#,##0.000")
        Case Else
                txtRmDesde = Format(txtRmDesde, "#,##0.0000")
    End Select
End If
'If txtRmDesde = "" Then
'    txtRmDesde = "0.00"
'End If
gbEstado = False
RaiseEvent Change
End Sub
Private Sub txtRmDesde_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtRmDesde, KeyAscii, , 4)
    If KeyAscii = 13 Then
'        If txtRmDesde = "" Then
'            txtRmDesde = "0.00"
'        End If
        txtRmHasta.SetFocus
    End If
End Sub

Private Sub txtRmDesde_GotFocus()
With txtRmDesde
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtRmHasta_Change()
txtRmHasta.SelStart = Len(txtRmHasta)
gnNumDec = NumDecimal(txtRmHasta)
If gbEstado And txtRmHasta <> "" Then
    Select Case gnNumDec
        Case 0
                txtRmHasta = Format(txtRmHasta, "#,##0")
        Case 1
                txtRmHasta = Format(txtRmHasta, "#,##0.0")
        Case 2
                txtRmHasta = Format(txtRmHasta, "#,##0.00")
        Case 3
                txtRmHasta = Format(txtRmHasta, "#,##0.000")
        Case Else
                txtRmHasta = Format(txtRmHasta, "#,##0.0000")
    End Select
End If
'If txtRmHasta = "" Then
'    txtRmHasta = 0#
'End If
gbEstado = False
RaiseEvent Change
End Sub
Private Sub txtRmHasta_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtRmHasta, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtRtDesde.SetFocus
    End If
End Sub

Private Sub txtRmHasta_GotFocus()
With txtRmHasta
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub
Private Sub txtRtDesde_Change()
txtRtDesde.SelStart = Len(txtRtDesde)
gnNumDec = NumDecimal(txtRtDesde)
If gbEstado And txtRtDesde <> "" Then
    Select Case gnNumDec
        Case 0
                txtRtDesde = Format(txtRtDesde, "#,##0")
        Case 1
                txtRtDesde = Format(txtRtDesde, "#,##0.0")
        Case 2
                txtRtDesde = Format(txtRtDesde, "#,##0.00")
        Case 3
                txtRtDesde = Format(txtRtDesde, "#,##0.000")
        Case Else
                txtRtDesde = Format(txtRtDesde, "#,##0.0000")
    End Select
End If
'If txtRtDesde = "" Then
'    txtRtDesde = 0#
'End If
gbEstado = False
RaiseEvent Change
End Sub
Private Sub txtRtDesde_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtRtDesde, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtRtHasta.SetFocus
    End If
End Sub

Private Sub txtRtDesde_GotFocus()
With txtRtDesde
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtRtDesde_LostFocus()
    txtRtDesde.Text = Format(txtRtDesde.Text, "#0.00##")
End Sub

Private Sub txtRtHasta_Change()
txtRtHasta.SelStart = Len(txtRtHasta)
gnNumDec = NumDecimal(txtRtHasta)
If gbEstado And txtRtHasta <> "" Then
    Select Case gnNumDec
        Case 0
                txtRtHasta = Format(txtRtHasta, "#,##0")
        Case 1
                txtRtHasta = Format(txtRtHasta, "#,##0.0")
        Case 2
                txtRtHasta = Format(txtRtHasta, "#,##0.00")
        Case 3
                txtRtHasta = Format(txtRtHasta, "#,##0.000")
        Case Else
                txtRtHasta = Format(txtRtHasta, "#,##0.0000")
    End Select
End If
'If txtRtHasta = "" Then
'    txtRtHasta = 0#
'End If
gbEstado = False
RaiseEvent Change
End Sub
Private Sub txtRtHasta_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtRtHasta, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtRtHasta.SetFocus
    End If
End Sub

Private Sub txtRtHasta_GotFocus()
With txtRtHasta
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtRtHasta_LostFocus()
    txtRtHasta.Text = Format(txtRtHasta.Text, "#0.00##")
End Sub
