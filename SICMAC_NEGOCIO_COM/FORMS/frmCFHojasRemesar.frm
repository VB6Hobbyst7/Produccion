VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCFHojasRemesar 
   Caption         =   "Remesar Folios de Cartas Fianza"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTCartasFolios 
      Height          =   5535
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Cartas Fianza"
      TabPicture(0)   =   "frmCFHojasRemesar.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbRegistro"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblAgencia"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblFdesde"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblFhasta"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblFsin"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblFEnvio"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtFechaEnvio"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "FeRemesas"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmbAgencia"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdAceptar"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdQuitar"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdCerrar"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdActualizar"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtFDesde"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtSinFolio"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtFHasta"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      Begin VB.TextBox txtFHasta 
         Height          =   320
         Left            =   3960
         TabIndex        =   2
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtSinFolio 
         Height          =   320
         Left            =   1440
         TabIndex        =   3
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtFDesde 
         Height          =   320
         Left            =   1440
         TabIndex        =   1
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "A&ctualizar"
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   6120
         TabIndex        =   7
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "&Quitar"
         Height          =   375
         Left            =   4920
         TabIndex        =   6
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5640
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ComboBox cmbAgencia 
         Height          =   315
         ItemData        =   "frmCFHojasRemesar.frx":001C
         Left            =   1200
         List            =   "frmCFHojasRemesar.frx":001E
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   960
         Width           =   2535
      End
      Begin SICMACT.FlexEdit FeRemesas 
         Height          =   2115
         Left            =   240
         TabIndex        =   17
         Top             =   2520
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   3731
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Agencia-Fecha Envío-Desde-Hasta-Sin Folio-CodEnvio-Estado"
         EncabezadosAnchos=   "400-2000-1200-1000-1000-1200-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin MSMask.MaskEdBox txtFechaEnvio 
         Height          =   320
         Left            =   3960
         TabIndex        =   4
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "DD/MM/YYYY"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Remesas pendientes de Confirmación"
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   2280
         Width           =   2670
      End
      Begin VB.Label lblFEnvio 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Envío:"
         Height          =   195
         Left            =   2880
         TabIndex        =   14
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblFsin 
         AutoSize        =   -1  'True
         Caption         =   "CF sin Folio:"
         Height          =   195
         Left            =   480
         TabIndex        =   13
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblFhasta 
         AutoSize        =   -1  'True
         Caption         =   "Folio Hasta:"
         Height          =   195
         Left            =   3000
         TabIndex        =   12
         Top             =   1440
         Width           =   840
      End
      Begin VB.Label lblFdesde 
         AutoSize        =   -1  'True
         Caption         =   "Folio Desde:"
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   1440
         Width           =   885
      End
      Begin VB.Label lblAgencia 
         AutoSize        =   -1  'True
         Caption         =   "Agencia:"
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   960
         Width           =   630
      End
      Begin VB.Label lbRegistro 
         AutoSize        =   -1  'True
         Caption         =   "Registro de Remesas"
         Height          =   195
         Left            =   480
         TabIndex        =   9
         Top             =   600
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmCFHojasRemesar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fsCodAgencias As String
Dim fsCodEnvio As String
Private Sub cmbAgencia_Click()
fsCodAgencias = Trim(Right(Me.cmbAgencia.Text, 4))
fsCodAgencias = IIf(Len(fsCodAgencias) < 2, "0" & fsCodAgencias, fsCodAgencias)
End Sub

Private Sub cmdAceptar_Click()
If ValidaDatos Then
    If MsgBox("Esta seguro de Guardar el Envío?", vbInformation + vbYesNo, "Remesar Folios de CF") = vbYes Then
        fsCodEnvio = gsCodCMAC & fsCodAgencias & Format(gdFecSis, "yyyymmdd") & Format(Time, "hhmmss")
        GrabarDatos
        MsgBox "Envio de remesas registrado satisfactoriamente.", vbInformation, "Aviso"
        LimpiarDatos
        LlenarGridRemesas
    Else
        Exit Sub
    End If
End If
End Sub

Private Sub cmdActualizar_Click()
LlenarGridRemesas
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdQuitar_Click()
Dim oCartaFianza As COMNCartaFianza.NCOMCartaFianzaValida
Set oCartaFianza = New COMNCartaFianza.NCOMCartaFianzaValida
If Trim(Me.FeRemesas.TextMatrix(Me.FeRemesas.Row, 6)) = "" Then
    MsgBox "No a selecionado.", vbInformation, "Aviso"
Else
    If MsgBox("Esta seguro de quitar este registro?", vbInformation + vbYesNo, "Remesar Folios de CF") = vbYes Then
        Call oCartaFianza.QuitarEnvioFolios(Trim(Me.FeRemesas.TextMatrix(Me.FeRemesas.Row, 6)))
        LlenarGridRemesas
    End If
End If
End Sub
Private Sub Form_Load()
Call CentraForm(Me)
Me.Icon = LoadPicture(App.path & gsRutaIcono)
Cargar_Objetos_Controles
End Sub

Private Sub Cargar_Objetos_Controles()
Dim oAgencia  As COMDConstantes.DCOMAgencias
Dim rsAgencias As ADODB.Recordset
Set oAgencia = New COMDConstantes.DCOMAgencias
Set rsAgencias = oAgencia.ObtieneAgencias()
Call Llenar_Combo_con_Recordset(rsAgencias, cmbAgencia)
Me.txtFechaEnvio.Text = gdFecSis
LlenarGridRemesas
End Sub


Private Function ValidaDatos() As Boolean
Dim oCartaFianza As COMNCartaFianza.NCOMCartaFianzaValida
Dim rsCartaFianza As ADODB.Recordset

    If Me.cmbAgencia.Text = "" Then
        MsgBox "Seleccione Agencia", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    If txtFechaEnvio.Text = "__/__/____" Then
        MsgBox "Ingrese Fecha de Envío.", vbInformation, "Aviso"
        ValidaDatos = False
        txtFechaEnvio.SetFocus
        Exit Function
    End If
        
    Dim nPosicion As Integer
    nPosicion = InStr(Me.txtFechaEnvio, "_")
    If nPosicion > 0 Then
        MsgBox "Error en Fecha de Envío.", vbInformation, "Aviso"
        ValidaDatos = False
        txtFechaEnvio.SetFocus
        Exit Function
    End If
    
    If CInt(Mid(Me.txtFechaEnvio, 7, 4)) > CInt(Mid(gdFecSis, 7, 4)) Or CInt(Mid(Me.txtFechaEnvio, 7, 4)) < (CInt(Mid(gdFecSis, 7, 4)) - 100) Then
        MsgBox "Año fuera del rango.", vbInformation, "Aviso"
        ValidaDatos = False
        txtFechaEnvio.SetFocus
        Exit Function
    End If
        
    If CInt(Mid(Me.txtFechaEnvio, 4, 2)) > 12 Or CInt(Mid(Me.txtFechaEnvio, 4, 2)) < 1 Then
        MsgBox "Error en Mes.", vbInformation, "Aviso"
        ValidaDatos = False
        txtFechaEnvio.SetFocus
        Exit Function
    End If
            
    Dim nDiasEnMes As Integer
    nDiasEnMes = CInt(DateDiff("d", CDate("01" & Mid(Me.txtFechaEnvio, 3, 8)), DateAdd("M", 1, CDate("01" & Mid(Me.txtFechaEnvio, 3, 8)))))
        
    If CInt(Mid(Me.txtFechaEnvio, 1, 2)) > nDiasEnMes Or CInt(Mid(Me.txtFechaEnvio, 1, 2)) < 1 Then
        MsgBox "Dia fuera del rango.", vbInformation, "Aviso"
        ValidaDatos = False
        txtFechaEnvio.SetFocus
        Exit Function
    End If
    
    
     If Me.txtFDesde.Text = "" Then
        MsgBox "Ingrese inicio Nº Folios.", vbInformation, "Aviso"
        ValidaDatos = False
        txtFDesde.SetFocus
        Exit Function
    End If
    
    If Me.txtFHasta.Text = "" Then
        MsgBox "Ingrese final Nº Folios.", vbInformation, "Aviso"
        ValidaDatos = False
        txtFHasta.SetFocus
        Exit Function
    End If
    If Me.txtSinFolio.Text = "" Then
        MsgBox "Ingrese Cantidad de CF sin Folio.", vbInformation, "Aviso"
        ValidaDatos = False
        txtSinFolio.SetFocus
        Exit Function
    End If
    
    If val(Me.txtFDesde.Text) > 0 Then
        If (val(Me.txtFHasta.Text) - val(Me.txtFDesde.Text)) < 49 Then
            MsgBox "Se debe enviar un minimo 50 unidades foliadas. ", vbInformation, "Aviso"
            ValidaDatos = False
            txtFHasta.SetFocus
            Exit Function
        End If
    End If
    
     If val(Me.txtSinFolio.Text) > 0 Then
        If val(Me.txtSinFolio.Text) < 50 Then
            MsgBox "Debe enviar un minimo de 50 unidades de CF sin Folio.", vbInformation, "Aviso"
            ValidaDatos = False
            txtSinFolio.SetFocus
            Exit Function
        End If
    End If
    
    If val(Me.txtFDesde.Text) = 0 Or val(Me.txtFHasta.Text) = 0 Then
        If val(Me.txtSinFolio.Text) = 0 Then
            MsgBox "Debe ingresar por lo menos CF sin Folio.", vbInformation, "Aviso"
            ValidaDatos = False
            txtSinFolio.SetFocus
            Exit Function
        End If
    End If
    
    Set oCartaFianza = New COMNCartaFianza.NCOMCartaFianzaValida
    Set rsCartaFianza = oCartaFianza.ObtenerEnvioFolios("0", fsCodAgencias)
    If rsCartaFianza.RecordCount > 0 Then
        MsgBox "Agencia aun no confirma la recepción de su envío pendiente.", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    Set rsCartaFianza = oCartaFianza.ObtenerFolios(CLng(Me.txtFDesde.Text), CLng(Me.txtFHasta.Text))
    Dim i As Integer
    Dim sCadAgencias As String
    If val(Me.txtFDesde.Text) > 0 Then
    If rsCartaFianza.RecordCount > 0 Then
        If rsCartaFianza.RecordCount > 1 Then
            For i = 0 To rsCartaFianza.RecordCount - 1
                If i = rsCartaFianza.RecordCount - 1 Then
                sCadAgencias = sCadAgencias & " y " & rsCartaFianza!cAgeDescripcion & "(" & rsCartaFianza!nDesde & "," & rsCartaFianza!nHasta & ")"
                Else
                sCadAgencias = sCadAgencias & IIf(i > 0, ", ", "") & rsCartaFianza!cAgeDescripcion & "(" & rsCartaFianza!nDesde & "," & rsCartaFianza!nHasta & ")"
                End If
            rsCartaFianza.MoveNext
            Next i
            MsgBox "Los Numeros de folio ya fueron utilizados en la " & sCadAgencias & ".", vbInformation, "Aviso"
        Else
            MsgBox "El numero de folio ya fue utilizado en la " & rsCartaFianza!cAgeDescripcion & "(" & rsCartaFianza!nDesde & "," & rsCartaFianza!nHasta & ")" & ".", vbInformation, "Aviso"
        End If
        ValidaDatos = False
        txtFDesde.SetFocus
        Exit Function
    End If
    End If
    
    Set oCartaFianza = Nothing
    Set rsCartaFianza = Nothing
    
    
ValidaDatos = True
End Function

Private Sub txtFDesde_Change()
If val(Me.txtFDesde.Text) > 0 Then
    Me.txtFHasta.Text = val(Me.txtFDesde.Text) + 49
Else
    Me.txtFHasta.Text = "0"
End If
End Sub

Private Sub txtFDesde_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
End Sub
Private Sub txtFHasta_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
End Sub
Private Sub txtSinFolio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
End Sub
Private Sub GrabarDatos()
Dim oCartaFianza As COMNCartaFianza.NCOMCartaFianzaValida
On Error GoTo ErrorGrabar
Set oCartaFianza = New COMNCartaFianza.NCOMCartaFianzaValida
Call oCartaFianza.GrabarEnvioFolios(fsCodEnvio, fsCodAgencias, Me.txtFechaEnvio.Text, CLng(Me.txtFDesde.Text), CLng(Me.txtFHasta.Text), CLng(Me.txtSinFolio.Text))

Exit Sub
ErrorGrabar:
MsgBox Err.Description, vbInformation, "Aviso"
End Sub
Private Sub LlenarGridRemesas()
Dim oCartaFianza As COMNCartaFianza.NCOMCartaFianzaValida
Dim rsCartaFianza As ADODB.Recordset
Dim i As Integer
Set oCartaFianza = New COMNCartaFianza.NCOMCartaFianzaValida
Set rsCartaFianza = oCartaFianza.ObtenerEnvioFolios("0")

Call LimpiaFlex(FeRemesas)
If rsCartaFianza.RecordCount > 0 Then
    If Not (rsCartaFianza.EOF Or rsCartaFianza.BOF) Then
        For i = 0 To rsCartaFianza.RecordCount - 1
            FeRemesas.AdicionaFila
            Me.FeRemesas.TextMatrix(i + 1, 0) = i + 1
            Me.FeRemesas.TextMatrix(i + 1, 1) = rsCartaFianza!cAgeDescripcion
            Me.FeRemesas.TextMatrix(i + 1, 2) = rsCartaFianza!dFechaEnvio
            Me.FeRemesas.TextMatrix(i + 1, 3) = rsCartaFianza!nDesde
            Me.FeRemesas.TextMatrix(i + 1, 4) = rsCartaFianza!nHasta
            Me.FeRemesas.TextMatrix(i + 1, 5) = rsCartaFianza!nSFolio
            Me.FeRemesas.TextMatrix(i + 1, 6) = rsCartaFianza!nCodEnvio
            Me.FeRemesas.TextMatrix(i + 1, 7) = rsCartaFianza!nEstado
            rsCartaFianza.MoveNext
        Next i
    End If
End If
End Sub

Private Sub LimpiarDatos()
Me.txtFDesde.Text = "0"
Me.txtFHasta.Text = "0"
Me.txtSinFolio.Text = "0"
End Sub
