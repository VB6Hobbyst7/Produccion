VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConstSistema 
   Caption         =   "ConstSistema : Mantenimiento"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9000
   Icon            =   "frmConstSistema.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Constantes de Sistema"
      TabPicture(0)   =   "frmConstSistema.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtDescripcion"
      Tab(0).Control(1)=   "txtValor"
      Tab(0).Control(2)=   "cmdNuevo"
      Tab(0).Control(3)=   "cmdModificar"
      Tab(0).Control(4)=   "cmdEliminar"
      Tab(0).Control(5)=   "cmdAceptar"
      Tab(0).Control(6)=   "grdConstSistema"
      Tab(0).Control(7)=   "cmdSalir"
      Tab(0).Control(8)=   "cmdCancelar"
      Tab(0).Control(9)=   "lbl"
      Tab(0).Control(10)=   "Label1"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Parámetros de Encaje"
      TabPicture(1)   =   "frmConstSistema.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdCancelarE"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdSalirE"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtFechaE"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "grdParamEncaje"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdAceptarE"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtCodigoE"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdEliminarE"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdModificarE"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdNuevoE"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtDescripcionE"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtValorE"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      Begin VB.TextBox txtValorE 
         Alignment       =   1  'Right Justify
         Height          =   340
         Left            =   6360
         TabIndex        =   21
         Top             =   4320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtDescripcionE 
         Enabled         =   0   'False
         Height          =   340
         Left            =   2520
         TabIndex        =   20
         Top             =   4320
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.CommandButton cmdNuevoE 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton cmdModificarE 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton cmdEliminarE 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         Top             =   5040
         Width           =   1215
      End
      Begin VB.TextBox txtCodigoE 
         Enabled         =   0   'False
         Height          =   340
         Left            =   240
         TabIndex        =   18
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdAceptarE 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   6000
         TabIndex        =   14
         Top             =   5040
         Width           =   1215
      End
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         Height          =   340
         Left            =   -74760
         TabIndex        =   9
         Top             =   4320
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.TextBox txtValor 
         Enabled         =   0   'False
         Height          =   340
         Left            =   -70080
         TabIndex        =   8
         Top             =   4320
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   -74760
         TabIndex        =   7
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   -73440
         TabIndex        =   6
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   -72120
         TabIndex        =   5
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   -69000
         TabIndex        =   4
         Top             =   5040
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid grdConstSistema 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "nConsSisCod"
            Caption         =   "Nº"
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
            DataField       =   "nConsSisDesc"
            Caption         =   "Descripcion"
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
            DataField       =   "nConsSisValor"
            Caption         =   "Valor"
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
               DividerStyle    =   6
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3720.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3404.977
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid grdParamEncaje 
         Height          =   3615
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
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
            DataField       =   "nCodigo"
            Caption         =   "Cod."
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
            DataField       =   "dFecha"
            Caption         =   "Fecha"
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
            DataField       =   "cDescripcion"
            Caption         =   "Descripcion"
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
            DataField       =   "nValor"
            Caption         =   "Valor"
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
               DividerStyle    =   6
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3795.024
            EndProperty
            BeginProperty Column03 
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox txtFechaE 
         Height          =   340
         Left            =   1200
         TabIndex        =   19
         Top             =   4320
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   -67680
         TabIndex        =   2
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   -67680
         TabIndex        =   3
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton cmdSalirE 
         Caption         =   "Salir"
         Height          =   375
         Left            =   7320
         TabIndex        =   12
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelarE 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   7320
         TabIndex        =   13
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         TabIndex        =   23
         Top             =   4200
         Width           =   8535
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   120
         TabIndex        =   22
         Top             =   4920
         Width           =   8535
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   -74880
         TabIndex        =   11
         Top             =   4200
         Width           =   8535
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   -74880
         TabIndex        =   10
         Top             =   4920
         Width           =   8535
      End
   End
End
Attribute VB_Name = "frmConstSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lConsulta As Boolean
Dim sSql As String
Dim rsConst As ADODB.Recordset
Dim rsParamEncaje As ADODB.Recordset '*** PEAC 20100706
Dim rsBuscaParamEncaje As ADODB.Recordset '*** PEAC 20100706
Dim i As Integer '*** PEAC 20100706
Dim lNuevo As Boolean
Dim oConst As NConstSistemas
'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim sqlCod As String
Dim Cod As Integer
On Error GoTo AceptarErr
'If Not ValidaDatos Then
'   Exit Sub
'End If
If MsgBox(" ¿ Seguro de grabar los cambios realizados ? ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then 'NAGL Cambio mensaje 20191223
   gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
   Cod = oConst.LeeConstSistemaCod
   Cod = Cod + 1
   Select Case lNuevo
      Case True
            oConst.InsertaValor Cod, txtDescripcion.Text, txtValor.Text, gsMovNro
         
      Case False
            oConst.ActualizaValor rsConst!nConsSisCod, txtDescripcion.Text, txtValor.Text, gsMovNro
   End Select
   CargaConstSistemaTodos
            
            'ARLO20170208
            Dim lsAccion As String
            Set objPista = New COMManejador.Pista
            If lNuevo Then
            lsAccion = "1"
            Else:
            lsAccion = "2"
            End If
            gsOpeCod = ""
            
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, lsAccion, Me.Caption & " : Constante Sistema | Descripción : " & txtDescripcion.Text & " | Valor : " & txtValor.Text
            Set objPista = Nothing
            '*******
   'rsImp.Find "cCtaContCod = '" & txtCta.Text & "'", , , 1
End If
ActivaBotones True
grdConstSistema.SetFocus
Exit Sub
AceptarErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdAceptarE_Click()
Dim sqlCod As String
Dim Cod As Integer
On Error GoTo AceptarErr
If MsgBox(" ¿ Seguro de grabar datos ingesados ? ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
    If Val(Me.txtCodigoE.Text) = 0 Or Len(Format(Me.txtFechaE.Text, "yyyymmdd")) = 0 Or Len(Trim(Me.txtDescripcionE)) = 0 Or Val(Me.txtValorE.Text) = 0 Then
        MsgBox "Ingrese datos válidos.", vbOKOnly, "Atención"
        Exit Sub
    End If
    
    
   Select Case lNuevo
      Case True
            If BuscaParametroEncaje(Val(Me.txtCodigoE.Text), Format(Me.txtFechaE.Text, "yyyymmdd")) Then
                MsgBox "El Parámetro que intenta grabar ya existe.", vbOKOnly, "Atención"
                Exit Sub
            End If
      
            oConst.InsertaValorEncaje Me.txtCodigoE.Text, Format(Me.txtFechaE.Text, "yyyymmdd"), Me.txtDescripcionE, Val(Me.txtValorE.Text)
      Case False
            oConst.ActualizaValorParametroEncaje Me.txtCodigoE.Text, Format(Me.txtFechaE.Text, "yyyymmdd"), Me.txtValorE.Text
   End Select
   CargaParametrosEncaje
            
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            If lNuevo Then
            lsPalabras = "Agrego"
            lsAccion = "1"
            Else: lsPalabras = "Modifico"
            lsAccion = "2"
            End If
            gsOpeCod = ""
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, lsAccion, Me.Caption & " : Parametro Encaje | Descripción : " & txtDescripcionE.Text & " | Valor : " & txtValorE.Text
            Set objPista = Nothing
            '*******
End If
Me.ActivaBotonesEncaje True
Me.grdParamEncaje.SetFocus
Exit Sub
AceptarErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"

End Sub

Private Sub cmdCancelar_Click()
ActivaBotones True
grdConstSistema.SetFocus
End Sub

Private Sub cmdCancelarE_Click()
    Me.ActivaBotonesEncaje True
    Me.grdParamEncaje.SetFocus
End Sub

Private Sub cmdEliminar_Click()
Dim sELIM As String
Dim lrs   As ADODB.Recordset
If rsConst.EOF Then
   MsgBox "No existen datos para Eliminar", vbInformation, "¡Aviso!"
   Exit Sub
End If
If MsgBox(" ¿ Está seguro de Eliminar el Valor ? ", vbQuestion + vbYesNo, "Mensaje de confirmación") = vbYes Then
   oConst.EliminaValor rsConst!nConsSisCod
   'rsConst.Delete adAffectCurrent
            
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            Dim lsOpeDes, lsOpeVal As String
            lsOpeDes = rsConst!nConsSisDesc
            lsOpeVal = rsConst!nConsSisValor
            gsOpeCod = ""
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", Me.Caption & " Constante del Sistema |Descripción : " & lsOpeDes & " Valor : " & lsOpeVal
            Set objPista = Nothing
            '*******
   CargaConstSistemaTodos
   Set oDoc = Nothing
   RSClose lrs
End If
grdConstSistema.SetFocus
End Sub

Private Sub cmdEliminarE_Click()
If rsParamEncaje.EOF Then
   MsgBox "No existen datos para Eliminar", vbInformation, "¡Aviso!"
   Exit Sub
End If
If MsgBox(" ¿ Está seguro de Eliminar el Valor ? ", vbQuestion + vbYesNo, "Mensaje de confirmación") = vbYes Then
   oConst.EliminaValorParamEncaje rsParamEncaje!nCodigo, Format(rsParamEncaje!dFecha, "yyyymmdd")
            
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            Dim lsOpeDes, lsOpeVal As String
            lsOpeDes = rsParamEncaje!cDescripcion
            lsOpeVal = rsParamEncaje!nValor
            gsOpeCod = ""
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", Me.Caption & " Parametro de Encaje |Descripción : " & lsOpeDes & " Valor : " & lsOpeVal
            Set objPista = Nothing
            '*******
   CargaParametrosEncaje
   Set oDoc = Nothing
End If
Me.grdParamEncaje.SetFocus

End Sub

Private Sub cmdModificar_Click()
If rsConst.EOF Then
   MsgBox "No existen datos para Eliminar", vbInformation, "¡Aviso!"
   Exit Sub
End If
txtDescripcion.Text = rsConst!nConsSisDesc
txtValor.Text = rsConst!nConsSisValor
txtDescripcion.Enabled = True
txtValor.Enabled = True
ActivaBotones False
lNuevo = False
End Sub

Private Sub cmdModificarE_Click()
If rsParamEncaje.EOF Then
   MsgBox "No existen datos para Modificar", vbInformation, "¡Aviso!"
   Exit Sub
End If
Me.txtCodigoE.Text = rsParamEncaje!nCodigo
Me.txtFechaE.Text = rsParamEncaje!dFecha
Me.txtDescripcionE.Text = rsParamEncaje!cDescripcion
Me.txtValorE.Text = rsParamEncaje!nValor

Me.txtCodigoE.Enabled = True
Me.txtFechaE.Enabled = True
Me.txtDescripcionE.Enabled = True
Me.txtValorE.Enabled = True
Me.ActivaBotonesEncaje False
lNuevo = False

End Sub

Private Sub cmdNuevo_Click()
txtDescripcion.Text = ""
txtValor.Text = ""
ActivaBotones False
txtDescripcion.Enabled = True
txtValor.Enabled = True
txtDescripcion.SetFocus
lNuevo = True
End Sub

Private Sub cmdNuevoE_Click()
    Me.txtCodigoE.Text = ""
    Me.txtFechaE.Text = "__/__/____"
    Me.txtDescripcionE.Text = ""
    Me.txtValorE.Text = ""
    Me.ActivaBotonesEncaje False
    Me.txtCodigoE.Enabled = True
    Me.txtFechaE.Enabled = True
    Me.txtDescripcionE.Enabled = True
    Me.txtValorE.Enabled = True
    
    Me.txtCodigoE.SetFocus
    lNuevo = True
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSalirE_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Dim oConst As NConstSistemas
    Set oConst = New NConstSistemas
    CargaConstSistemaTodos
    CargaParametrosEncaje
    If lConsulta Then
        cmdNuevo.Visible = False
        cmdModificar.Visible = False
        cmdEliminar.Visible = False
    End If
    CentraForm Me
End Sub
Private Sub CargaConstSistemaTodos()
    Set rsConst = oConst.LeeConstSistemaTodos()
    Set grdConstSistema.DataSource = rsConst
End Sub
Public Sub Inicio(plConsulta As Boolean)
lConsulta = plConsulta
Me.Show 1
End Sub
Sub ActivaBotones(lActiva As Boolean)
'If lActiva Then
'  grdConstSistema.Height = 3120
'Else
'   grdConstSistema.Height = 2640
'End If
cmdNuevo.Visible = lActiva
cmdModificar.Visible = lActiva
cmdEliminar.Visible = lActiva
cmdSalir.Visible = lActiva
cmdAceptar.Visible = Not lActiva
cmdCancelar.Visible = Not lActiva
txtDescripcion.Visible = Not lActiva
txtValor.Visible = Not lActiva
End Sub

Private Sub Form_Unload(Cancel As Integer)
RSClose rsConst
RSClose rsParamEncaje
Set rsConst = Nothing
Set rsParamEncaje = Nothing
End Sub

Private Sub grdConstSistema_GotFocus()
grdConstSistema.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub grdConstSistema_LostFocus()
grdConstSistema.MarqueeStyle = dbgNoMarquee
End Sub
Private Sub grdConstSistema_HeadClick(ByVal ColIndex As Integer)
If Not rsConst Is Nothing Then
   If Not rsConst.EOF Then
      rsConst.Sort = grdConstSistema.Columns(ColIndex).DataField
   End If
End If
End Sub

'*** PEAC 20100706
Private Sub CargaParametrosEncaje()
    Set rsParamEncaje = oConst.LeeParametrosEncaje()
    Set Me.grdParamEncaje.DataSource = rsParamEncaje
End Sub

'*** PEAC 20100707
Private Function BuscaParametroEncaje(pnCodigo As Integer, psFecha As String) As Boolean
    Set rsBuscaParamEncaje = oConst.BuscaParametrosEncaje(pnCodigo, psFecha)
    If Not (rsBuscaParamEncaje.EOF And rsBuscaParamEncaje.BOF) Then
        BuscaParametroEncaje = True
    Else
        BuscaParametroEncaje = False
    End If
    rsBuscaParamEncaje.Close
    Set rsbuscaparamencje = Nothing
End Function



'*** PEAC 20100706
Sub ActivaBotonesEncaje(lActiva As Boolean)
    Me.cmdNuevoE.Visible = lActiva
    Me.cmdModificarE.Visible = lActiva
    Me.cmdEliminarE.Visible = lActiva
    Me.cmdSalirE.Visible = lActiva
    Me.cmdAceptarE.Visible = Not lActiva
    Me.cmdCancelarE.Visible = Not lActiva
    Me.txtCodigoE.Visible = Not lActiva
    Me.txtFechaE.Visible = Not lActiva
    Me.txtDescripcionE.Visible = Not lActiva
    Me.txtValorE.Visible = Not lActiva
End Sub

Private Sub txtCodigoE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtDescripcionE_Change()
    txtDescripcionE.Text = UCase(txtDescripcionE.Text)
    i = Len(txtDescripcionE.Text)
    txtDescripcionE.SelStart = i
End Sub

Private Sub txtDescripcionE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtFechaE_GotFocus()
fEnfoque Me.txtFechaE
End Sub

Private Sub txtFechaE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtValorE_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtValorE, KeyAscii, 16, 2)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtValorE_LostFocus()
    If Len(Trim(txtValorE.Text)) = 0 Then
        txtValorE.Text = "0.00"
    End If
End Sub
