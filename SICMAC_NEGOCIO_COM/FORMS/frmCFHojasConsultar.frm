VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCFHojasConsultar 
   Caption         =   "Consultar Envios de Folios de Cartas Fianza"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTCartasFolios 
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Cartas Fianza"
      TabPicture(0)   =   "frmCFHojasConsultar.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbRegistro"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblAgencia"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblFdesde"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblFhasta"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblFEnvio"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtFechaEnvioB"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtFechaEnvioA"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "FeRemesas"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmbAgencia"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdAceptar"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "fraEstado"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdCerrar"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   8760
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Frame fraEstado 
         Caption         =   "Estados Envío"
         Height          =   1335
         Left            =   4200
         TabIndex        =   12
         Top             =   960
         Width           =   2775
         Begin VB.CheckBox chkEstados 
            Caption         =   "Terminado"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   15
            Top             =   960
            Width           =   1695
         End
         Begin VB.CheckBox chkEstados 
            Caption         =   "Confirmado"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   14
            Top             =   600
            Width           =   1695
         End
         Begin VB.CheckBox chkEstados 
            Caption         =   "Registrado"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   13
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   7560
         TabIndex        =   9
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ComboBox cmbAgencia 
         Height          =   315
         ItemData        =   "frmCFHojasConsultar.frx":001C
         Left            =   1320
         List            =   "frmCFHojasConsultar.frx":001E
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   960
         Width           =   2535
      End
      Begin SICMACT.FlexEdit FeRemesas 
         Height          =   2715
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   9885
         _extentx        =   17436
         _extenty        =   4789
         cols0           =   9
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-Agencia-Fecha Envío-Desde-Hasta-Sin Folio-Sin Folio Utl.-Estado-CodEnvio"
         encabezadosanchos=   "400-2000-1200-1000-1000-1200-1200-1600-0"
         font            =   "frmCFHojasConsultar.frx":0020
         font            =   "frmCFHojasConsultar.frx":004C
         font            =   "frmCFHojasConsultar.frx":0078
         font            =   "frmCFHojasConsultar.frx":00A4
         fontfixed       =   "frmCFHojasConsultar.frx":00D0
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         columnasaeditar =   "X-X-X-X-X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-C-C-C-C-C-C-C"
         formatosedit    =   "0-0-0-0-0-0-0-0-0"
         textarray0      =   "#"
         colwidth0       =   405
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin MSMask.MaskEdBox txtFechaEnvioA 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "DD/MM/YYYY"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaEnvioB 
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   1920
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
         Caption         =   "Envíos Encontrados:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   1500
      End
      Begin VB.Label lblFEnvio 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Envío:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblFhasta 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   1320
         TabIndex        =   6
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label lblFdesde 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   1320
         TabIndex        =   5
         Top             =   1560
         Width           =   510
      End
      Begin VB.Label lblAgencia 
         AutoSize        =   -1  'True
         Caption         =   "Agencia:"
         Height          =   195
         Left            =   600
         TabIndex        =   4
         Top             =   960
         Width           =   630
      End
      Begin VB.Label lbRegistro 
         AutoSize        =   -1  'True
         Caption         =   "Consulta de Remesas"
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   600
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmCFHojasConsultar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fsEstados As String
Dim fsCodAgencias As String
Private Sub chkEstados_Click(Index As Integer)
Dim nPosicion As Integer
Dim nTamano As Integer
Select Case Index
    Case 0:
            If chkEstados(0).value = 1 Then
                fsEstados = fsEstados & "0,"
            Else
                nPosicion = InStr(fsEstados, "0")
                nTamano = Len(fsEstados)
                fsEstados = Mid(fsEstados, 1, nPosicion - 1) & Mid(fsEstados, nPosicion + 2, nTamano - 2)
            End If
    Case 1:
            If chkEstados(1).value = 1 Then
                fsEstados = fsEstados & "1,"
            Else
                nPosicion = InStr(fsEstados, "1")
                nTamano = Len(fsEstados)
                fsEstados = Mid(fsEstados, 1, nPosicion - 1) & Mid(fsEstados, nPosicion + 2, nTamano - 2)
            End If
    Case 2:
            If chkEstados(2).value = 1 Then
                fsEstados = fsEstados & "2,"
            Else
                nPosicion = InStr(fsEstados, "2")
                nTamano = Len(fsEstados)
                fsEstados = Mid(fsEstados, 1, nPosicion - 1) & Mid(fsEstados, nPosicion + 2, nTamano - 2)
            End If
End Select
End Sub
Private Sub cmbAgencia_Click()
fsCodAgencias = Trim(Right(Me.cmbAgencia.Text, 4))
fsCodAgencias = IIf(Len(fsCodAgencias) < 2, "0" & fsCodAgencias, fsCodAgencias)
End Sub

Private Sub cmdAceptar_Click()
If validaDatos Then
    LlenarGridRemesas
End If
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub FeRemesas_DblClick()
Dim oCartaFianza As COMNCartaFianza.NCOMCartaFianzaValida
Dim rsCartaFianza As ADODB.Recordset
Dim sCodEnvio As String
sCodEnvio = Trim(FeRemesas.TextMatrix(Me.FeRemesas.Row, 8))
If sCodEnvio <> "" Then
    Set oCartaFianza = New COMNCartaFianza.NCOMCartaFianzaValida
    Set rsCartaFianza = oCartaFianza.ConsultarEnviosFolios(2, , , , , sCodEnvio)
    If rsCartaFianza.RecordCount > 0 Then
        frmCFHojasConsultarFolio.Inicio (sCodEnvio)
    Else
        MsgBox "No contiene Folios numerados.", vbInformation, "Aviso"
    End If
Else
    MsgBox "Seleccione Correctamente el Registro.", vbInformation, "Aviso"
End If
End Sub

Private Sub Form_Load()
Call CentraForm(Me)
Me.Icon = LoadPicture(App.path & gsRutaIcono)
fsEstados = ""
Cargar_Objetos_Controles
End Sub
Private Sub LlenarGridRemesas()
'ByVal psAgencia As String, ByVal pdFechaA As Date, ByVal pdFechaB As Date, ByVal psEstado As String
Dim oCartaFianza As COMNCartaFianza.NCOMCartaFianzaValida
Dim rsCartaFianza As ADODB.Recordset
Dim i As Integer
Set oCartaFianza = New COMNCartaFianza.NCOMCartaFianzaValida
Set rsCartaFianza = oCartaFianza.ConsultarEnviosFolios(1, fsCodAgencias, Me.txtFechaEnvioA.Text, Me.txtFechaEnvioB.Text, Mid(fsEstados, 1, Len(fsEstados) - 1))

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
            Me.FeRemesas.TextMatrix(i + 1, 6) = rsCartaFianza!nSFolioUtl
            Me.FeRemesas.TextMatrix(i + 1, 7) = rsCartaFianza!Estado
            Me.FeRemesas.TextMatrix(i + 1, 8) = rsCartaFianza!nCodEnvio
            rsCartaFianza.MoveNext
        Next i
    End If
Else
    MsgBox "No exiten datos a mostrar.", vbInformation, "Aviso"
End If
End Sub


Private Function validaDatos() As Boolean

    If Me.cmbAgencia.Text = "" Then
        MsgBox "Seleccione Agencia", vbInformation, "Aviso"
        validaDatos = False
        Exit Function
    End If
    
    If txtFechaEnvioA.Text = "__/__/____" Then
        MsgBox "Ingrese Fecha de Envío Desde.", vbInformation, "Aviso"
        validaDatos = False
        txtFechaEnvioA.SetFocus
        Exit Function
    End If
        
    Dim nPosicion As Integer
    nPosicion = InStr(Me.txtFechaEnvioA, "_")
    If nPosicion > 0 Then
        MsgBox "Error en Fecha de Envío Desde.", vbInformation, "Aviso"
        validaDatos = False
        txtFechaEnvioA.SetFocus
        Exit Function
    End If
    
    If CInt(Mid(Me.txtFechaEnvioA, 7, 4)) > CInt(Mid(gdFecSis, 7, 4)) Or CInt(Mid(Me.txtFechaEnvioA, 7, 4)) < (CInt(Mid(gdFecSis, 7, 4)) - 100) Then
        MsgBox "Año fuera del rango.", vbInformation, "Aviso"
        validaDatos = False
        txtFechaEnvioA.SetFocus
        Exit Function
    End If
        
    If CInt(Mid(Me.txtFechaEnvioA, 4, 2)) > 12 Or CInt(Mid(Me.txtFechaEnvioA, 4, 2)) < 1 Then
        MsgBox "Error en Mes.", vbInformation, "Aviso"
        validaDatos = False
        txtFechaEnvioA.SetFocus
        Exit Function
    End If
            
    Dim nDiasEnMes As Integer
    nDiasEnMes = CInt(DateDiff("d", CDate("01" & Mid(Me.txtFechaEnvioA, 3, 8)), DateAdd("M", 1, CDate("01" & Mid(Me.txtFechaEnvioA, 3, 8)))))
        
    If CInt(Mid(Me.txtFechaEnvioA, 1, 2)) > nDiasEnMes Or CInt(Mid(Me.txtFechaEnvioA, 1, 2)) < 1 Then
        MsgBox "Dia fuera del rango.", vbInformation, "Aviso"
        validaDatos = False
        txtFechaEnvioA.SetFocus
        Exit Function
    End If


    If txtFechaEnvioB.Text = "__/__/____" Then
        MsgBox "Ingrese Fecha de Envío Hasta.", vbInformation, "Aviso"
        validaDatos = False
        txtFechaEnvioB.SetFocus
        Exit Function
    End If
        

    nPosicion = InStr(Me.txtFechaEnvioB, "_")
    If nPosicion > 0 Then
        MsgBox "Error en Fecha de Envío Hasta.", vbInformation, "Aviso"
        validaDatos = False
        txtFechaEnvioB.SetFocus
        Exit Function
    End If
    
    If CInt(Mid(Me.txtFechaEnvioB, 7, 4)) > CInt(Mid(gdFecSis, 7, 4)) Or CInt(Mid(Me.txtFechaEnvioB, 7, 4)) < (CInt(Mid(gdFecSis, 7, 4)) - 100) Then
        MsgBox "Año fuera del rango.", vbInformation, "Aviso"
        validaDatos = False
        txtFechaEnvioB.SetFocus
        Exit Function
    End If
        
    If CInt(Mid(Me.txtFechaEnvioB, 4, 2)) > 12 Or CInt(Mid(Me.txtFechaEnvioB, 4, 2)) < 1 Then
        MsgBox "Error en Mes.", vbInformation, "Aviso"
        validaDatos = False
        txtFechaEnvioB.SetFocus
        Exit Function
    End If

    nDiasEnMes = CInt(DateDiff("d", CDate("01" & Mid(Me.txtFechaEnvioB, 3, 8)), DateAdd("M", 1, CDate("01" & Mid(Me.txtFechaEnvioB, 3, 8)))))
        
    If CInt(Mid(Me.txtFechaEnvioB, 1, 2)) > nDiasEnMes Or CInt(Mid(Me.txtFechaEnvioB, 1, 2)) < 1 Then
        MsgBox "Dia fuera del rango.", vbInformation, "Aviso"
        validaDatos = False
        txtFechaEnvioB.SetFocus
        Exit Function
    End If
    
    If fsEstados = "" Then
        MsgBox "Debe escoger por lo menos un Estado a buscar.", vbInformation, "Aviso"
        validaDatos = False
        Exit Function
    End If
      
validaDatos = True
End Function

Private Sub Cargar_Objetos_Controles()
Dim oAgencia  As COMDConstantes.DCOMAgencias
Dim rsAgencias As ADODB.Recordset
Set oAgencia = New COMDConstantes.DCOMAgencias
Set rsAgencias = oAgencia.ObtieneAgencias()
Call Llenar_Combo_con_Recordset(rsAgencias, cmbAgencia)
End Sub

