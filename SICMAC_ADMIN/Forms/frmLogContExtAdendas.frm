VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogContExtAdendas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contratación: Extorno de Adenda"
   ClientHeight    =   4170
   ClientLeft      =   2220
   ClientTop       =   1200
   ClientWidth     =   8070
   Icon            =   "frmLogContExtAdendas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTContratos 
      Height          =   4020
      Left            =   80
      TabIndex        =   0
      Top             =   80
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7091
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Extorno"
      TabPicture(0)   =   "frmLogContExtAdendas.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCancelar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdExtornar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraProv"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame fraProv 
         Caption         =   "Contrato"
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
         Height          =   2805
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   7680
         Begin VB.TextBox txtNAdenda 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   6000
            MaxLength       =   5
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtGlosa 
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
            Height          =   1170
            Left            =   1080
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   1200
            Width           =   6180
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Glosa:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   480
            TabIndex        =   12
            Top             =   1200
            Width           =   450
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Nº Adenda:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5040
            TabIndex        =   11
            Top             =   240
            Width           =   825
         End
         Begin VB.Label lblNAdenda 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   6000
            TabIndex        =   10
            Tag             =   "txtnombre"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Nº Contrato:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   270
            Width           =   870
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   750
            Width           =   780
         End
         Begin VB.Label lblNContrato 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   7
            Tag             =   "txtnombre"
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label lblProveedor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   6
            Tag             =   "txtnombre"
            Top             =   720
            Width           =   6135
         End
      End
      Begin VB.CommandButton cmdExtornar 
         Caption         =   "&Extornar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5570
         TabIndex        =   4
         Top             =   3360
         Width           =   1110
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
         Height          =   360
         Left            =   6690
         TabIndex        =   3
         Top             =   3360
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "_"
         ForeColor       =   &H80000008&
         Height          =   75
         Left            =   3240
         TabIndex        =   1
         Top             =   2760
         Width           =   90
      End
   End
   Begin VB.PictureBox CdlgFile 
      Height          =   615
      Left            =   7440
      ScaleHeight     =   555
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   800
   End
End
Attribute VB_Name = "frmLogContExtAdendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim fsNContrato As String
Dim fsNAdenda As Integer
Dim fnContRef As Integer 'PASI20140823 ti-ers077-2014
Dim fnTipo As Integer
Dim fnEstado As Integer
Dim fntpodocorigen As Integer
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************

'EJVG20131204 ***
'Solo se podrá extornar la ultima Adenda siempre y cuando no hayan realizado Pagos
Public Sub Inicio(ByVal psNContrato As String, ByVal psNAdenda As Integer, ByVal pnTpoAdenda As Integer, Optional ByVal pnTpoDocOrigen As Integer = 0, Optional ByVal pnContRef As Integer = 0) 'pnContRef agregado PASI20140823 ti-ers077-2014
    fsNContrato = psNContrato
    fsNAdenda = psNAdenda
    fnTipo = pnTpoAdenda
    fntpodocorigen = pnTpoDocOrigen
    fnContRef = pnContRef
    If Not ValidaExtornoAdenda Then Exit Sub
    Call CargaDatos
    Me.Show 1
End Sub
Private Function ValidaExtornoAdenda() As Boolean
    Dim olog As New DLogGeneral
    Dim lnUltAdenda As Integer
    ValidaExtornoAdenda = True
    
    lnUltAdenda = olog.ObtenerUltAdendaContratos(fsNContrato, fnContRef) 'fncontref agregado PASIERS0772014
    If fsNAdenda <> lnUltAdenda Then
        ValidaExtornoAdenda = False
        MsgBox "Para continuar primero se debe extornar la Última Adenda Nro. " & Format(lnUltAdenda, "00"), vbInformation, "Aviso"
        Exit Function
    End If
    If olog.RealizaronPagosdeAdenda(fsNContrato, fsNAdenda, fnContRef) Then 'fncontref agregado PASIERS0772014
        ValidaExtornoAdenda = False
        MsgBox "No se puede continuar porque ya se realizaron pagos luego de realizar la Adenda Nro. " & Format(fsNAdenda, "00"), vbInformation, "Aviso"
        Exit Function
    End If
    If fntpodocorigen = LogTipoContrato.ContratoAdqBienes Or fntpodocorigen = LogTipoContrato.ContratoSuministro Then
        If olog.RealizaronPagosdeAdendaBien(fsNContrato, fsNAdenda, fnContRef) Then 'PASIERS0772014
            ValidaExtornoAdenda = False
            MsgBox "No se puede continuar porque ya se realizaron pagos luego de realizar la Adenda Nro. " & Format(fsNAdenda, "00"), vbInformation, "Aviso"
            Exit Function
        End If
    End If
    If fntpodocorigen = LogTipoContrato.ContratoObra Then
        If olog.ExistenPagosValoriza(fsNContrato, fsNAdenda, fnContRef) Then 'PASIERS0772014
            ValidaExtornoAdenda = False
            MsgBox "No se puede continuar porque ya se realizaron pagos luego de realizar la Adenda Nro. " & Format(fsNAdenda, "00"), vbInformation, "Aviso"
            Exit Function
        End If
    End If
End Function
'EBD EJVG *******
Private Sub CargaDatos()
Dim olog As DLogGeneral
Dim rsLog As ADODB.Recordset

Set olog = New DLogGeneral
Set rsLog = olog.ListarDatosContratos(fsNContrato, fnContRef) 'fncontref agregado pasi20140823 ti-ers077-2014

If rsLog.RecordCount > 0 Then
    Me.lblNContrato.Caption = Space(1) & rsLog!NContrato
    Me.lblProveedor.Caption = Space(1) & rsLog!Proveedor
    Me.lblNAdenda.Caption = fsNAdenda
End If

End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdExtornar_Click()
    Dim olog As DLogGeneral
    Dim bTrans As Boolean
    On Error GoTo ErrorExtorno
    
    If Not ValidaExtorno Then Exit Sub
    If MsgBox("Esta seguro de grabar el Extorno?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    Set olog = New DLogGeneral
    olog.dBeginTrans
    bTrans = True
    
    
    olog.ExtornarAdenda_NEW fsNContrato, fsNAdenda, fnContRef 'fnContRef Agregado PASIERS0772014
    olog.RegistrarExtornoAdenda_NEW fsNContrato, fsNAdenda, fnTipo, Trim(txtGlosa.Text), GeneraMov(gdFecSis, "109", gsCodAge, gsCodUser), fnContRef 'fnContRef Agregado PASI20140823 ti-ers077-2014f
    If fnContRef <> 0 Then 'PASIERS0772014
'        If pnTpoDocOrigen = LogTipoContrato.ContratoAdqBienes Or fntpodocorigen = LogTipoContrato.ContratoSuministro Then
'
'        End If
'        If pnTpoDocOrigen = LogTipoContrato.ContratoObra Then
'
'        End If
'        If pnTpoDocOrigen = LogTipoContrato.ContratoArrendamiento Or fntpodocorigen = LogTipoContrato.ContratoServicio Then
'
'        End If
        olog.ActualizaSaldoxAdenda fsNContrato, fnContRef, fsNAdenda
    End If
    
    olog.dCommitTrans
    bTrans = False
    Set olog = Nothing
    Screen.MousePointer = 0
    
    MsgBox "Extorno de Adenda registrada satisfactoriamente", vbInformation, "Aviso"
    'ARLO 20160126 ***
    gsOpeCod = LogPistaRegistraAdenda
    Set objPista = New COMManejador.Pista
    objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", "Extorno la Adenda N° : " & fsNAdenda & " | Del Contrato N° : " & fsNContrato & " | Por Motivo : " & txtGlosa.Text
    Set objPista = Nothing
    '***
    Unload Me
    Exit Sub
ErrorExtorno:
    Screen.MousePointer = 0
    If bTrans Then
        olog.dRollbackTrans
        Set olog = Nothing
    End If
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Function ValidaExtorno() As Boolean
If Trim(Me.txtGlosa.Text) = "" Then
    MsgBox "Ingrese la Glosa del Extorno", vbInformation, "Aviso"
    ValidaExtorno = False
    Exit Function
End If

ValidaExtorno = True
End Function
Sub LimpiarDatos()
    Me.txtGlosa.Text = ""
End Sub
Private Sub Form_Load()
If fsNAdenda = 0 Then
    Me.lblNAdenda.Visible = False
    Me.txtNAdenda.Visible = True
    Me.cmdExtornar.Enabled = False
End If
End Sub

Private Sub txtNAdenda_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    If CInt(Me.txtNAdenda.Text) > 0 Then
        If CargaAdenda Then
            cmdExtornar.Enabled = True
        End If
    Else
        MsgBox "Ingrese Numero Valido", vbInformation, "Aviso"
    End If
End If
End Sub

Private Function CargaAdenda() As Boolean
Dim olog As DLogGeneral
Dim rsLog As ADODB.Recordset
Dim bAdenda As Boolean
Set olog = New DLogGeneral

    Set rsLog = olog.ListarDatosAdenda(fsNContrato, CInt(Me.txtNAdenda.Text), fnContRef) 'fnContRef Agregado PASI20140823 ti-ers077-2014f
    If rsLog.RecordCount > 0 Then
        If CInt(rsLog!nEstado) = 1 Then
            fnTipo = CInt(rsLog!nTipo)
            fsNAdenda = CInt(Me.txtNAdenda.Text)
            CargaAdenda = True
            fnEstado = CInt(rsLog!nEstado)
        ElseIf CInt(rsLog!nEstado) = 2 Then
            MsgBox "No se puede extornar. Adenda Cancelada.", vbInformation, "Aviso"
            CargaAdenda = False
        ElseIf CInt(rsLog!nEstado) = 3 Then
            MsgBox "No se puede extornar. Adenda En proceso de Pago.", vbInformation, "Aviso"
            CargaAdenda = False
        ElseIf CInt(rsLog!nEstado) = 4 Then
            MsgBox "Adenda ya fue extornada.", vbInformation, "Aviso"
            CargaAdenda = False
        End If
    Else
        MsgBox "No se encuentra la adenda", vbInformation, "Aviso"
        CargaAdenda = False
    End If
End Function

Private Function ExitenAdendas() As Boolean
Dim olog As DLogGeneral
Dim rsLog As ADODB.Recordset

Set olog = New DLogGeneral
Set rsLog = olog.ListarDatosContratos(fsNContrato, fnContRef)

If rsLog.RecordCount > 0 Then
    'ADENDAS
    Set rsLog = olog.ListarDatosAdendasPorContrato(fsNContrato, fnContRef)
    If rsLog.RecordCount > 0 Then
        ExitenAdendas = True
    Else
        MsgBox "Contrato no cuenta con Adendas", vbInformation, "Aviso"
        ExitenAdendas = False
    End If
Else
    MsgBox "Contrato no cuenta con Adendas", vbInformation, "Aviso"
    ExitenAdendas = False
End If
End Function


