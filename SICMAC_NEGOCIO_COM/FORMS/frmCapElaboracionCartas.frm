VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCapElaboracionCartas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CARTAS DE CLIENTES CON ACTIVIDADES ILICITAS"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   Icon            =   "frmCapElaboracionCartas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraCliente 
      Caption         =   "Procesado (s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Index           =   1
      Left            =   135
      TabIndex        =   25
      Top             =   5505
      Width           =   9810
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   350
         Left            =   8715
         TabIndex        =   27
         Top             =   1905
         Width           =   1000
      End
      Begin SICMACT.FlexEdit grdProcesado 
         Height          =   1995
         Left            =   135
         TabIndex        =   26
         Top             =   255
         Width           =   8475
         _extentx        =   14949
         _extenty        =   3519
         cols0           =   4
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-Dni-Nombre-Codigo"
         encabezadosanchos=   "600-1800-5000-0"
         font            =   "frmCapElaboracionCartas.frx":030A
         font            =   "frmCapElaboracionCartas.frx":0336
         font            =   "frmCapElaboracionCartas.frx":0362
         font            =   "frmCapElaboracionCartas.frx":038E
         font            =   "frmCapElaboracionCartas.frx":03BA
         fontfixed       =   "frmCapElaboracionCartas.frx":03E6
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         columnasaeditar =   "X-X-X-X"
         listacontroles  =   "0-0-0-0"
         encabezadosalineacion=   "C-C-L-C"
         formatosedit    =   "0-0-0-0"
         textarray0      =   "#"
         lbbuscaduplicadotext=   -1  'True
         colwidth0       =   600
         rowheight0      =   300
         forecolor       =   -2147483627
         forecolorfixed  =   -2147483627
         cellforecolor   =   -2147483627
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8940
      TabIndex        =   13
      Top             =   7965
      Width           =   1000
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   7710
      TabIndex        =   12
      Top             =   7965
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   105
      TabIndex        =   11
      Top             =   7935
      Width           =   1000
   End
   Begin RichTextLib.RichTextBox rtfCartas 
      Height          =   330
      Left            =   6555
      TabIndex        =   9
      Top             =   7950
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   582
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmCapElaboracionCartas.frx":0414
   End
   Begin VB.Frame FraCliente 
      Caption         =   "Datos Persona"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2640
      Index           =   0
      Left            =   150
      TabIndex        =   10
      Top             =   2775
      Width           =   9795
      Begin VB.CommandButton cmdReferencia 
         Caption         =   "&Referencia"
         Height          =   350
         Left            =   8685
         TabIndex        =   40
         Top             =   1440
         Width           =   1000
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Nuev. Busq"
         Height          =   350
         Left            =   8700
         TabIndex        =   8
         Top             =   1845
         Width           =   1000
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Enabled         =   0   'False
         Height          =   350
         Left            =   8700
         TabIndex        =   17
         Top             =   2235
         Width           =   1000
      End
      Begin SICMACT.FlexEdit grdCuentas 
         Height          =   1185
         Left            =   135
         TabIndex        =   28
         Top             =   1380
         Width           =   8445
         _extentx        =   14896
         _extenty        =   2090
         cols0           =   5
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-Cuenta-Producto-Estado-Relac."
         encabezadosanchos=   "600-1800-2200-1500-1500"
         font            =   "frmCapElaboracionCartas.frx":0497
         font            =   "frmCapElaboracionCartas.frx":04C3
         font            =   "frmCapElaboracionCartas.frx":04EF
         font            =   "frmCapElaboracionCartas.frx":051B
         font            =   "frmCapElaboracionCartas.frx":0547
         fontfixed       =   "frmCapElaboracionCartas.frx":0573
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         columnasaeditar =   "X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0"
         encabezadosalineacion=   "C-C-L-C-C"
         formatosedit    =   "0-0-0-0-0"
         textarray0      =   "#"
         lbbuscaduplicadotext=   -1  'True
         colwidth0       =   600
         rowheight0      =   300
         forecolor       =   -2147483630
         forecolorfixed  =   -2147483627
         cellforecolor   =   -2147483630
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   350
         Left            =   8670
         TabIndex        =   29
         Top             =   1065
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.Frame FraCli 
         BorderStyle     =   0  'None
         Height          =   810
         Left            =   225
         TabIndex        =   30
         Top             =   330
         Width           =   8880
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2145
            TabIndex        =   7
            Top             =   15
            Width           =   390
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Código :"
            Height          =   195
            Left            =   0
            TabIndex        =   36
            Top             =   60
            Width           =   585
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   15
            TabIndex        =   35
            Top             =   465
            Width           =   600
         End
         Begin VB.Label lblNomCli 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   735
            TabIndex        =   34
            Top             =   405
            Width           =   7650
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Documento Identidad:"
            Height          =   195
            Left            =   4485
            TabIndex        =   33
            Top             =   75
            Width           =   1575
         End
         Begin VB.Label lbllCodigoCli 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   750
            TabIndex        =   32
            Top             =   0
            Width           =   1380
         End
         Begin VB.Label lblDICli 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   6150
            TabIndex        =   31
            Top             =   30
            Width           =   2230
         End
      End
      Begin VB.Frame FraReferencia 
         BorderStyle     =   0  'None
         Height          =   945
         Left            =   135
         TabIndex        =   37
         Top             =   255
         Width           =   8775
         Begin VB.TextBox txtDI 
            Height          =   315
            Left            =   1710
            MaxLength       =   20
            TabIndex        =   14
            Top             =   105
            Width           =   2415
         End
         Begin VB.TextBox TxtNombre 
            Height          =   315
            Left            =   1695
            MaxLength       =   100
            TabIndex        =   16
            Top             =   495
            Width           =   6330
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Documento Identidad:"
            Height          =   195
            Left            =   105
            TabIndex        =   39
            Top             =   180
            Width           =   1575
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   105
            TabIndex        =   38
            Top             =   555
            Width           =   600
         End
      End
   End
   Begin VB.Frame fraCarta 
      Caption         =   "Información Carta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   90
      TabIndex        =   15
      Top             =   15
      Width           =   9840
      Begin VB.TextBox txtSecretario 
         Height          =   360
         Left            =   1455
         MaxLength       =   100
         TabIndex        =   6
         Top             =   2145
         Width           =   7290
      End
      Begin VB.ComboBox cboTpoDelito 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1680
         Width           =   7305
      End
      Begin VB.TextBox txtNroDoc 
         Height          =   345
         Left            =   5325
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1230
         Width           =   2535
      End
      Begin VB.ComboBox cboTpoDoc 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1230
         Width           =   2760
      End
      Begin VB.ComboBox cboProvincia 
         Height          =   315
         Left            =   5310
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   795
         Width           =   3480
      End
      Begin VB.ComboBox cboDepto 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   765
         Width           =   2760
      End
      Begin VB.TextBox txtEntidad 
         Height          =   360
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   0
         Top             =   330
         Width           =   7290
      End
      Begin VB.Label Label10 
         Caption         =   "Tipo de Delito"
         Height          =   240
         Left            =   225
         TabIndex        =   24
         Top             =   1785
         Width           =   1080
      End
      Begin VB.Label Label9 
         Caption         =   "Secretario"
         Height          =   240
         Left            =   225
         TabIndex        =   23
         Top             =   2280
         Width           =   780
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo Doc."
         Height          =   240
         Left            =   225
         TabIndex        =   22
         Top             =   1350
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Depto."
         Height          =   240
         Left            =   225
         TabIndex        =   21
         Top             =   840
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "Entidad"
         Height          =   240
         Left            =   225
         TabIndex        =   20
         Top             =   375
         Width           =   645
      End
      Begin VB.Label Label8 
         Caption         =   "Nro Doc."
         Height          =   240
         Left            =   4485
         TabIndex        =   19
         Top             =   1365
         Width           =   705
      End
      Begin VB.Label Label5 
         Caption         =   "Provincia"
         Height          =   240
         Left            =   4485
         TabIndex        =   18
         Top             =   840
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmCapElaboracionCartas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub TipoActI()
  Dim clsGen As COMNCaptaGenerales.NCOMCaptaGenerales, i As Integer
    Dim rsPar As ADODB.Recordset
    Set clsGen = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsPar = clsGen.GetTpoActI
    
    i = 0
    While Not rsPar.EOF
        
        cboTpoDelito.AddItem Trim(rsPar!cconsdescripcion)
        cboTpoDelito.ItemData(i) = rsPar!nconsvalor
        
        i = i + 1
        rsPar.MoveNext
    Wend
    
    If rsPar.RecordCount > 0 Then cboTpoDelito.ListIndex = 0
    
    Set clsGen = Nothing
    Set rsPar = Nothing
End Sub

Private Sub TipoDocActI()
    Dim clsGen As COMNCaptaGenerales.NCOMCaptaGenerales, i As Integer
    Dim rsPar As ADODB.Recordset
    Set clsGen = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsPar = clsGen.GetTpoDocActI
    i = 0
    While Not rsPar.EOF
        
        cboTpoDoc.AddItem Trim(rsPar!cconsdescripcion)
        cboTpoDoc.ItemData(i) = rsPar!nconsvalor
        i = i + 1
        rsPar.MoveNext
    Wend
    
    If rsPar.RecordCount > 0 Then cboTpoDoc.ListIndex = 0
    
    Set clsGen = Nothing
    Set rsPar = Nothing
End Sub

Private Sub DPTO()
    Dim clsGen As COMNCaptaGenerales.NCOMCaptaGenerales, i As Integer
    Dim rsPar As ADODB.Recordset
    Set clsGen = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsPar = clsGen.GetDPTO
    
    Me.cboDepto.Clear
    i = 0
    While Not rsPar.EOF
    
    
        cboDepto.AddItem Trim(rsPar!cubigeodes)
        cboDepto.ItemData(i) = Mid(rsPar!cubigeoCOD, 2, 2)
        i = i + 1
        rsPar.MoveNext
    Wend
    
    If rsPar.RecordCount > 0 Then cboDepto.ListIndex = 0
    
    Set clsGen = Nothing
    Set rsPar = Nothing
End Sub

Private Sub PROV()
  Dim clsGen As COMNCaptaGenerales.NCOMCaptaGenerales, i As Integer
    Dim rsPar As ADODB.Recordset
    Set clsGen = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsPar = clsGen.GetPROV(Format(CStr(cboDepto.ItemData(cboDepto.ListIndex)), "00"))
    
    cboProvincia.Clear
    While Not rsPar.EOF
        cboProvincia.AddItem Trim(rsPar!cubigeodes) & Space(100 - Len(Trim(rsPar!cubigeodes))) & Trim(rsPar!cubigeoCOD)
        'cboProvincia.ItemData(i) = Mid(rsPar!cubigeoCOD, 2, 4)
        
        rsPar.MoveNext
    Wend
    
    If rsPar.RecordCount > 0 Then cboProvincia.ListIndex = 0
    
    Set clsGen = Nothing
    Set rsPar = Nothing
End Sub


Private Sub cboDepto_Click()
    If cboDepto.ListCount >= 1 Then
        PROV
    End If
End Sub

Private Sub CmdAgregar_Click()
Dim nFila As Integer
nFila = grdProcesado.Rows - 1
If FraCli.Visible = False Then
    If EstaEnLista(Trim(Me.TxtNombre)) Then

        MsgBox "ESTA PERSONA YA FIGURA EN EL LISTADO DE PROCESADOS.", vbOKOnly + vbInformation, "AVISO"
        cmdAgregar.Enabled = False
    
        Exit Sub
        
    End If

    If nFila = 20 Then

        MsgBox "EL NRO MAXIMO DE PROCESADOS POR CARTA ES DE 20", vbOKOnly + vbInformation, "AVISO"
        Exit Sub
    
    End If

    If grdProcesado.TextMatrix(nFila, 2) <> "" Then
    
            nFila = nFila + 1
            grdProcesado.AdicionaFila , , True
                  
            grdProcesado.TextMatrix(nFila, 0) = nFila
            grdProcesado.TextMatrix(nFila, 1) = Me.txtDI
            grdProcesado.TextMatrix(nFila, 2) = Me.TxtNombre
            grdProcesado.TextMatrix(nFila, 3) = Me.txtDI
    
    Else

            grdProcesado.TextMatrix(nFila, 0) = nFila
            grdProcesado.TextMatrix(nFila, 1) = Me.txtDI
            grdProcesado.TextMatrix(nFila, 2) = Me.TxtNombre
            grdProcesado.TextMatrix(nFila, 3) = Me.txtDI
    
    End If

    Me.TxtNombre.Text = ""
    Me.txtDI.Text = ""
    FraReferencia.Visible = False
    FraCli.Visible = True
Else
    If EstaEnLista(Trim(Me.lblNomCli)) Then

        MsgBox "ESTA PERSONA YA FIGURA EN EL LISTADO DE PROCESADOS.", vbOKOnly + vbInformation, "AVISO"
        cmdAgregar.Enabled = False
    
        Exit Sub
        
    End If

    If nFila = 20 Then

        MsgBox "EL NRO MAXIMO DE PROCESADOS POR CARTA ES DE 20", vbOKOnly + vbInformation, "AVISO"
        Exit Sub
    
    End If

    If grdProcesado.TextMatrix(nFila, 2) <> "" Then
    
            nFila = nFila + 1
            grdProcesado.AdicionaFila , , True
                  
            grdProcesado.TextMatrix(nFila, 0) = nFila
            grdProcesado.TextMatrix(nFila, 1) = Me.lblDICli
            grdProcesado.TextMatrix(nFila, 2) = Me.lblNomCli
            grdProcesado.TextMatrix(nFila, 3) = Me.lblDICli
    
    Else

            grdProcesado.TextMatrix(nFila, 0) = nFila
            grdProcesado.TextMatrix(nFila, 1) = Me.lblDICli
            grdProcesado.TextMatrix(nFila, 2) = Me.lblNomCli
            grdProcesado.TextMatrix(nFila, 3) = Me.lblDICli
    
    End If


End If
    
cmdAgregar.Enabled = False

End Sub

Private Function EstaEnLista(ByVal cPerscod As String) As Boolean
Dim i As Integer
    EstaEnLista = False
    
    For i = 1 To grdProcesado.Rows - 1
        If Trim(grdProcesado.TextMatrix(i, 2)) = cPerscod Then
            EstaEnLista = True
            Exit Function
        End If
    Next i

End Function

Private Sub cmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String, lsDni As String
Dim lsEstados As String


'On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
'    If loPers Is Nothing Then
'          FraCli.Visible = False
'          FraReferencia.Visible = True
'          txtDI.Text = ""
'          TxtNombre.Text = ""
'          txtDI.SetFocus
'
'       Exit Sub
'    End If
    
   If loPers Is Nothing Then
       Exit Sub
    End If
    
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
    lsDni = loPers.sPersIdnroDNI

Dim i As Integer
If lsPersCod <> "" Then
    Me.lbllCodigoCli.Caption = lsPersCod
    Me.lblNomCli.Caption = lsPersNombre
    Me.lblDICli.Caption = lsDni
    
    Dim nMant As COMNCaptaGenerales.NCOMCaptaGenerales, rsTemp As Recordset
    Set nMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        Set rsTemp = New ADODB.Recordset
        Set rsTemp = nMant.GetCuentasPersonaAI(lsPersCod)
        
     grdCuentas.Clear
     
     If grdCuentas.Rows - 1 >= 2 Then
        For i = grdCuentas.Rows - 1 To 2 Step -1
            grdCuentas.EliminaFila i
        Next i
     End If
     
     
     grdCuentas.FormaCabecera
      
      If rsTemp.RecordCount > 0 Then
         Set grdCuentas.Recordset = rsTemp
         cmdAgregar.Enabled = False
         CmdImprimir.Enabled = True
      Else
         cmdAgregar.Enabled = True
         CmdImprimir.Enabled = False
      End If
        cmdExit.Enabled = True
        
        Set rsTemp = Nothing
        Set nMant = Nothing
        
 
        
End If

Set loPers = Nothing
'Exit Sub
'
'ControlError:   ' Rutina de control de errores.
'    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
'        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub LimpiarControles()

Dim i As Integer

    txtEntidad.Text = ""
    txtNroDoc.Text = ""
    txtSecretario.Text = ""
    Me.lbllCodigoCli.Caption = ""
    Me.lblDICli.Caption = ""
    Me.lblNomCli = ""
    grdCuentas.Clear
    grdCuentas.FormaCabecera
    grdProcesado.Clear
    grdProcesado.FormaCabecera
    If grdCuentas.Rows - 1 >= 2 Then
      For i = grdCuentas.Rows - 1 To 2 Step -1
        grdCuentas.EliminaFila i
      Next i
    End If
    If grdProcesado.Rows - 1 >= 2 Then
      For i = grdProcesado.Rows - 1 To 2 Step -1
        grdProcesado.EliminaFila i
      Next i
    End If
End Sub
Private Sub cmdCancelar_Click()
  LimpiarControles
  cmdAgregar.Enabled = False
  cmdEliminar.Enabled = False
  CmdImprimir.Enabled = False
  cmdExit.Enabled = False
  cmdGrabar.Enabled = True
End Sub

Private Sub cmdEliminar_Click()
 If grdProcesado.Rows >= 2 And grdProcesado.Row >= 1 And Trim(grdProcesado.TextMatrix(1, 1)) <> "" Then
    If grdProcesado.Row = 1 Then
       grdProcesado.TextMatrix(1, 0) = ""
       grdProcesado.TextMatrix(1, 1) = ""
       grdProcesado.TextMatrix(1, 2) = ""
       
    Else
       grdProcesado.EliminaFila grdProcesado.Row, True
    End If
 End If
 
End Sub

Private Sub cmdExit_Click()
   Me.lbllCodigoCli.Caption = ""
   Me.lblDICli.Caption = ""
   Me.lblNomCli = ""
  Dim i As Integer
   grdCuentas.Clear
   For i = grdCuentas.Rows - 1 To 2 Step -1
      grdCuentas.EliminaFila i
   Next i
 
   grdCuentas.FormaCabecera
   FraCli.Visible = True
   FraReferencia.Visible = False
   cmdBuscar.SetFocus
   cmdAgregar.Enabled = False
   CmdImprimir.Enabled = False
End Sub

Private Sub ImpCarta(ByVal NroCarta As String)
Dim sNomA As String

'rtfCartas.FileName = App.path & "\FormatoCarta\CartaOP.doc"

 sNomA = App.path & "\FormatoCarta\FORMCARTAI.dot"
 Call CartaWORD(sNomA, NroCarta)

End Sub

Private Sub CartaWORD(ByVal psNomPlantilla As String, ByVal NroCarta As String)
Dim aLista() As String
Dim vFilas As Integer

    'Crea una clase que de Word Object
    Dim wApp As Word.Application
    Dim wAppSource As Word.Application
    'Create a new instance of word
    Set wApp = New Word.Application
    Set wAppSource = New Word.Application
     

    Dim RangeSource As Word.Range
    'Abre Documento Plantilla
    wAppSource.Documents.Open FileName:=psNomPlantilla
    Set RangeSource = wAppSource.ActiveDocument.Content
    
    'Lo carga en Memoria
    wAppSource.ActiveDocument.Content.Copy
                
    'Crea Nuevo Documento
    wApp.Documents.Add
    
    wApp.Application.Selection.TypeParagraph


    With wApp.Selection.PageSetup
        .TopMargin = 120 'CentimetersToPoints(10)
        .BottomMargin = 60 'CentimetersToPoints(3)
        .LeftMargin = 140 'CentimetersToPoints(10)
        .RightMargin = 80 'CentimetersToPoints(5)
    End With

           
        wApp.Application.Selection.Paste
        wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
        wApp.Selection.MoveEnd


        With wApp.Selection.Find
            .Text = "<CNROCARTA>"
            .Replacement.Text = Trim(NroCarta)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
       With wApp.Selection.Find
            .Text = "<ENTIDAD>"
            .Replacement.Text = Trim(txtEntidad.Text)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<DESCDPTO>"
            .Replacement.Text = Trim(Left(cboDepto.Text, 100))
            .Forward = True
            .Wrap = wdFindContinue
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<DESCPROV>"
            .Replacement.Text = Trim(Left(cboProvincia.Text, 100))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<CTPODOC>"
            .Replacement.Text = Trim(cboTpoDoc.Text) & " N° " & Trim(txtNroDoc.Text)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        
        With wApp.Selection.Find
            .Text = "<SECRETARIO>"
            .Replacement.Text = Trim(txtSecretario.Text)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll

        With wApp.Selection.Find
            .Text = "<APROCESADO>"
            .Replacement.Text = Trim(grdProcesado.TextMatrix(1, 2))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
            
       If grdProcesado.Rows - 1 >= 2 Then
         With wApp.Selection.Find
                .Text = "<BPROCESADO>"
                .Replacement.Text = Trim(grdProcesado.TextMatrix(2, 2))
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
         wApp.Selection.Find.Execute Replace:=wdReplaceAll
       Else
           With wApp.Selection.Find
                .Text = "<BPROCESADO>"
                .Replacement.Text = "         "
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
         wApp.Selection.Find.Execute Replace:=wdReplaceAll
       
       End If
       
       If grdProcesado.Rows - 1 >= 3 Then
         With wApp.Selection.Find
                .Text = "<CPROCESADO>"
                .Replacement.Text = Trim(grdProcesado.TextMatrix(3, 2))
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
         wApp.Selection.Find.Execute Replace:=wdReplaceAll
       Else
         With wApp.Selection.Find
                .Text = "<CPROCESADO>"
                .Replacement.Text = "         "
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
         wApp.Selection.Find.Execute Replace:=wdReplaceAll
       
       End If
       
       If grdProcesado.Rows - 1 >= 4 Then
         With wApp.Selection.Find
                .Text = "<DPROCESADO>"
                .Replacement.Text = Trim(grdProcesado.TextMatrix(4, 2))
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
         wApp.Selection.Find.Execute Replace:=wdReplaceAll
       Else
           With wApp.Selection.Find
                .Text = "<DPROCESADO>"
                .Replacement.Text = "         "
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
          wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
       End If
       
        If grdProcesado.Rows - 1 >= 5 Then
         With wApp.Selection.Find
                .Text = "<EPROCESADO>"
                .Replacement.Text = Trim(grdProcesado.TextMatrix(5, 2))
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
         wApp.Selection.Find.Execute Replace:=wdReplaceAll
       Else
          With wApp.Selection.Find
                .Text = "<EPROCESADO>"
                .Replacement.Text = "         "
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
         wApp.Selection.Find.Execute Replace:=wdReplaceAll
       
       End If
       
       wApp.GoBack
          
wAppSource.ActiveDocument.Close
wApp.Visible = True

End Sub

Private Sub cmdGrabar_Click()
Dim clsCap As COMDCaptaGenerales.DCOMCaptaGenerales, sMovNro As String
Dim clsMov As COMNContabilidad.NCOMContFunciones, i As Integer, vNroCarta As String
Set clsCap = New COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMov = New COMNContabilidad.NCOMContFunciones
vNroCarta = ""

For i = 1 To grdProcesado.Rows - 1
 sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

 If i = 1 Then
   Call clsCap.AgregaCartaAIp(sMovNro, CStr(i), Trim(Left(grdProcesado.TextMatrix(i, 2), 99)) + Space(100 - Len(Trim(Left(grdProcesado.TextMatrix(i, 2), 99)))) + grdProcesado.TextMatrix(i, 3), cboTpoDelito.ItemData(cboTpoDelito.ListIndex), Trim(txtEntidad.Text), Trim(txtSecretario.Text), cboTpoDoc.ItemData(cboTpoDoc.ListIndex), Trim(Me.txtNroDoc.Text), Right(cboProvincia.Text, 12), vNroCarta)
 End If
 
 If i > 1 Then
   Call clsCap.AgregaCartaAI(sMovNro, CStr(i), Trim(Left(grdProcesado.TextMatrix(i, 2), 99)) + Space(100 - Len(Trim(Left(grdProcesado.TextMatrix(i, 2), 99)))) + grdProcesado.TextMatrix(i, 3), cboTpoDelito.ItemData(cboTpoDelito.ListIndex), Trim(txtEntidad.Text), Trim(txtSecretario.Text), cboTpoDoc.ItemData(cboTpoDoc.ListIndex), Trim(Me.txtNroDoc.Text), Right(cboProvincia.Text, 12), vNroCarta)
 End If
 
Next i

'Call ImpCarta(vNroCarta)

cmdGrabar.Enabled = False
Set clsCap = Nothing
Set clsMov = Nothing

End Sub

Private Sub cmdReferencia_Click()
    FraCli.Visible = False
    FraReferencia.Visible = True
    grdCuentas.Clear
    grdCuentas.FormaCabecera
End Sub

Private Sub cmdSalir_Click()
 Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        SendKeys "{Tab}"
 End If
End Sub

Private Sub Form_Load()

DPTO
TipoActI
TipoDocActI

End Sub
Private Sub txtEntidad_KeyPress(KeyAscii As Integer)
  Dim prletras As String
      prletras = Chr(KeyAscii)
      prletras = UCase(prletras)
      KeyAscii = Asc(prletras)
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
 Dim prletras As String
      prletras = Chr(KeyAscii)
      prletras = UCase(prletras)
      KeyAscii = Asc(prletras)
      
      If KeyAscii = 13 Then
        cmdAgregar.Enabled = True
      End If
      
End Sub

Private Sub txtNroDoc_KeyPress(KeyAscii As Integer)
  Dim prletras As String
      prletras = Chr(KeyAscii)
      prletras = UCase(prletras)
      KeyAscii = Asc(prletras)
End Sub

Private Sub txtSecretario_KeyPress(KeyAscii As Integer)
Dim prletras As String
      prletras = Chr(KeyAscii)
      prletras = UCase(prletras)
      KeyAscii = Asc(prletras)
 
End Sub

Private Function ValidaControl() As Boolean
    ValidaControl = True
    
    If Trim(txtEntidad.Text) = "" Then
      MsgBox "DEBE INDICAR LA ENTIDAD DE DONDE PROVIENE EL DOCUMENTO ", vbOKOnly + vbInformation, "AVISO"
      ValidaControl = False
      Exit Function
    End If
    
    If Trim(txtNroDoc.Text) = "" Then
      MsgBox "DEBE INDICAR EL NRO DE DOCUMENTO ", vbOKOnly + vbInformation, "AVISO"
      ValidaControl = False
      Exit Function
    End If
    
    If Trim(txtSecretario.Text) = "" Then
      MsgBox "DEBE INDICAR EL NOMBRE DEL SECRETARIO ", vbOKOnly + vbInformation, "AVISO"
      ValidaControl = False
      Exit Function
    End If
    
    If Trim(cboDepto.Text) = "" Then
      MsgBox "DEBE INDICAR EL DEPARTAMENTO DE PROCEDENCIA DEL DOCUMENTO.", vbOKOnly + vbInformation, "AVISO"
      ValidaControl = False
      Exit Function
    End If
    
    If Trim(cboProvincia.Text) = "" Then
      MsgBox "DEBE INDICAR LA PROVINCIA DE PROCEDENCIA DEL DOCUMENTO. ", vbOKOnly + vbInformation, "AVISO"
      ValidaControl = False
      Exit Function
    End If
    
    If (grdProcesado.Rows - 1) >= 1 Then
         If Not (Trim(grdProcesado.TextMatrix(1, 2)) <> "") Then
            MsgBox "DEBE INGRESARSE AL MENOS UN PROCESADO.", vbOKOnly + vbInformation, "AVISO"
            ValidaControl = False
            Exit Function
         End If
    End If

End Function
