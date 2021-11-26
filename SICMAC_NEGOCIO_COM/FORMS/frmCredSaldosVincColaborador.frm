VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredSaldosVincColaborador 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saldo Disponible por Colaborador"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10185
   Icon            =   "frmCredSaldosVincColaborador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Personal CMACM"
      TabPicture(0)   =   "frmCredSaldosVincColaborador.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblSaldoTotal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblSaldoDisponible"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdCerrar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   8280
         TabIndex        =   8
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "Créditos de Colaborador y Vinculados"
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
         Height          =   3015
         Left            =   120
         TabIndex        =   2
         Top             =   1920
         Width           =   9735
         Begin SICMACT.FlexEdit feCreditos 
            Height          =   2535
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   9495
            _extentx        =   16748
            _extenty        =   4471
            cols0           =   8
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "#-Nª Credito-Titular-Relacion-Moneda-Monto Desemb.-Fecha-Saldo Capital"
            encabezadosanchos=   "500-2000-4000-1200-1200-1500-1200-1200"
            font            =   "frmCredSaldosVincColaborador.frx":0326
            font            =   "frmCredSaldosVincColaborador.frx":0352
            font            =   "frmCredSaldosVincColaborador.frx":037E
            font            =   "frmCredSaldosVincColaborador.frx":03AA
            fontfixed       =   "frmCredSaldosVincColaborador.frx":03D6
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            columnasaeditar =   "X-X-X-X-X-X-X-X"
            listacontroles  =   "0-0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-L-C-C-C-C-C-C"
            formatosedit    =   "0-0-0-0-0-0-0-0"
            textarray0      =   "#"
            colwidth0       =   495
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Buscar Colaborador"
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
         TabIndex        =   1
         Top             =   480
         Width           =   9735
         Begin SICMACT.TxtBuscar txtPersona 
            Height          =   285
            Left            =   1080
            TabIndex        =   9
            Top             =   360
            Width           =   2100
            _extentx        =   3704
            _extenty        =   503
            appearance      =   1
            appearance      =   1
            font            =   "frmCredSaldosVincColaborador.frx":0404
            appearance      =   1
            tipobusqueda    =   7
            stitulo         =   ""
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Persona:"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "DOI:"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   840
            Width           =   330
         End
         Begin VB.Label lblNumDoc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1080
            TabIndex        =   11
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label lblNombrePersona 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3240
            TabIndex        =   10
            Top             =   360
            Width           =   6135
         End
      End
      Begin VB.Label lblSaldoDisponible 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5520
         TabIndex        =   7
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Disponible a Solicitud:"
         Height          =   195
         Left            =   3480
         TabIndex        =   6
         Top             =   5040
         Width           =   2010
      End
      Begin VB.Label lblSaldoTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Total:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   5040
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCredSaldosVincColaborador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private i As Integer
Private fnSaldoTrabVinc As Double
Private fnSaldoAceptado As Double
Private fnPatrimonioEfec As Double
'Private fgFecActual As Date 'ORCR20140314**************************
Private ffFecActual As Date 'ORCR20140314**************************
Private fn7PorcPatriEfec As Double
Private fn5PorcDel7PorcPatriEfec As Double
Private fnMontoMax As Double
Private fnResCredVin As Double 'ORCR20140314**************************
Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    CargaValoresParametros
End Sub
'ORCR20140314**************************
Public Sub Inicio()
    Dim oConsSist As New COMDConstSistema.NCOMConstSistema
    Dim oPatrimonioEfectivo As New COMNCredito.NCOMPatrimonioEfectivo
    
    ffFecActual = oConsSist.LeeConstSistema(gConstSistFechaInicioDia)
    
    fnPatrimonioEfec = oPatrimonioEfectivo.ObtenerPatrimonioEfectivo(Year(ffFecActual), Format(Month(ffFecActual), "00"))
    
    If fnPatrimonioEfec = 0 Then
       ffFecActual = DateAdd("m", -1, ffFecActual)
       fnPatrimonioEfec = oPatrimonioEfectivo.ObtenerPatrimonioEfectivo(Year(ffFecActual), Format(Month(ffFecActual), "00"))
       
       If fnPatrimonioEfec = 0 Then
           MsgBox "favor de definir el Patrimonio Efectivo para continuar", vbInformation, "Aviso"
           Exit Sub
       End If
    End If
    
    Me.Show 1
End Sub
'END ORCR20140314**************************

Private Sub txtPersona_EmiteDatos()
On Error GoTo ErrorPersona
    If Trim(txtPersona.psCodigoPersona) <> "" Then
        'ORCR20140314**************************
        Dim oNCOMPersona As New COMNPersona.NCOMPersona
        
        If Not oNCOMPersona.ObtenerPersonaRHEstado(txtPersona.psCodigoPersona) Then
            MsgBox "Trabajador NO Activo", vbInformation, "Aviso"
            Call LimpiaDatos
            Exit Sub
        End If
        'END ORCR20140314**************************
        lblNombrePersona.Caption = txtPersona.psDescripcion
        lblNumDoc.Caption = txtPersona.sPersNroDoc
        
        Call CargaDatos(txtPersona.psCodigoPersona)
        Call CargaCreditosVinc(txtPersona.psCodigoPersona)
        
        Me.lblSaldoTotal.Caption = Format(fnSaldoTrabVinc, "###," & String(15, "#") & "#0.00") & " "
        'Me.lblSaldoDisponible.Caption = Format(fnMontoMax - fnSaldoAceptado - fnSaldoTrabVinc, "###," & String(15, "#") & "#0.00") & " " 'ORCR20140314**************************
        Me.lblSaldoDisponible.Caption = Format((fnPatrimonioEfec * fn7PorcPatriEfec * fn5PorcDel7PorcPatriEfec) - fnSaldoTrabVinc - fnSaldoAceptado, "###," & String(15, "#") & "#0.00") & " " 'ORCR20140314**************************
    End If
    
    Exit Sub
ErrorPersona:
    MsgBox err.Description, vbInformation, "Error"
End Sub
Private Sub txtPersona_KeyPress(KeyAscii As Integer)
    If Len(Trim(txtPersona)) < 13 Then
        LimpiaDatos
    End If
End Sub

Private Sub LimpiaDatos()
    lblNombrePersona.Caption = ""
    lblNumDoc.Caption = ""
    LimpiaFlex feCreditos
    fnSaldoTrabVinc = 0
    
    'ORCR20140314**************************
    lblSaldoTotal.Caption = ""
    lblSaldoDisponible.Caption = ""
    'END ORCR20140314**************************
End Sub
Private Sub CargaDatos(ByVal psPersCod As String)
    'ORCR20140314**************************
'    Dim oCredito As COMNCredito.NCOMCredito
'    Set oCredito = New COMNCredito.NCOMCredito
    Dim oCredito As New COMNCredito.NCOMCredito
    fnSaldoAceptado = oCredito.ObtenerSaldoAsignaTrabEstado(psPersCod, 2)
    Set oCredito = Nothing
    'END ORCR20140314**************************
End Sub

Private Sub CargaCreditosVinc(ByVal psPersCod As String)
    Dim oCredito As COMNCredito.NCOMCredito
    Dim rsCredito As ADODB.Recordset
    Set oCredito = New COMNCredito.NCOMCredito
    
    Set rsCredito = oCredito.ObtenerCreditosTrabVinc(psPersCod)
    LimpiaFlex feCreditos
    fnSaldoTrabVinc = 0
    If Not (rsCredito.EOF And rsCredito.BOF) Then
        For i = 1 To rsCredito.RecordCount
            feCreditos.AdicionaFila
            feCreditos.TextMatrix(i, 1) = rsCredito!cCtaCod
            feCreditos.TextMatrix(i, 2) = rsCredito!cPersNombre
            feCreditos.TextMatrix(i, 3) = rsCredito!Relacion
            feCreditos.TextMatrix(i, 4) = rsCredito!Moneda
            feCreditos.TextMatrix(i, 5) = Format(rsCredito!nMontoCol, "###," & String(15, "#") & "#0.00")
            feCreditos.TextMatrix(i, 6) = Format(rsCredito!dVigencia, "dd/mm/yyyy")
            feCreditos.TextMatrix(i, 7) = Format(rsCredito!nSaldo, "###," & String(15, "#") & "#0.00")
            fnSaldoTrabVinc = fnSaldoTrabVinc + CDbl(rsCredito!nSaldoMN)
            rsCredito.MoveNext
        Next i
    End If
    Set oCredito = Nothing
End Sub
Private Sub CargaValoresParametros()
    'ORCR20140314**************************
'    Dim sAnio As String
'    Dim sMes As String
'    Dim oPar As COMDCredito.DCOMParametro
'    Dim oConsSist As COMDConstSistema.NCOMConstSistema
'
'    Set oConsSist = New COMDConstSistema.NCOMConstSistema
'
'    fgFecActual = oConsSist.LeeConstSistema(gConstSistCierreMesNegocio)
'    sAnio = Year(fgFecActual)
'    sMes = Format(Month(fgFecActual), "00")
'    fnPatrimonioEfec = ObtenerSaldos(sAnio, sMes)
'
'    Set oPar = New COMDCredito.DCOMParametro
'    fn7PorcPatriEfec = oPar.RecuperaValorParametro(102752) / 100
'    fn5PorcDel7PorcPatriEfec = oPar.RecuperaValorParametro(102753) / 100
'
'    fnMontoMax = fn5PorcDel7PorcPatriEfec * fn7PorcPatriEfec * fnPatrimonioEfec
'
'    Set oConsSist = Nothing
'    Set oPar = Nothing
    
    Dim sAnio As String
    Dim sMes As String
    Dim oPar As New COMDCredito.DCOMParametro
    
    sAnio = Year(ffFecActual)
    sMes = Format(Month(ffFecActual), "00")
    
    fn7PorcPatriEfec = oPar.RecuperaValorParametro(102752) / 100
    fn5PorcDel7PorcPatriEfec = oPar.RecuperaValorParametro(102753) / 100
    
    fnMontoMax = fn5PorcDel7PorcPatriEfec * fn7PorcPatriEfec * fnPatrimonioEfec
    
    Set oPar = Nothing
    'END ORCR20140314**************************
End Sub
'ORCR20140314**************************
'Private Function ObtenerSaldos(ByVal psAnio As String, ByVal psMes As String) As Double
'Dim oNContabilidad As COMNContabilidad.NCOMContFunciones
'Dim nSaldo As Double
'Set oNContabilidad = New COMNContabilidad.NCOMContFunciones
'
'nSaldo = oNContabilidad.PatrimonioEfecAjustInfl(psAnio, psMes)
'ObtenerSaldos = nSaldo
'Set oNContabilidad = Nothing
'End Function
'END ORCR20140314**************************

