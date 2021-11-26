VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredSaldosVincAsignar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignación de Saldos a Colaboradores y Vinculados"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14910
   Icon            =   "frmCredSaldosVincAsignar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   14910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Asignación"
      TabPicture(0)   =   "frmCredSaldosVincAsignar.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblSaldoMN"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblSaldoPorAsignar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblSaldoDA"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraSolicitud"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.Frame fraSolicitud 
         Caption         =   "Solicitudes de Saldo"
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
         Height          =   5775
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   14535
         Begin VB.CommandButton cmdAsignar 
            Caption         =   "Asignar"
            Height          =   375
            Left            =   8520
            TabIndex        =   8
            Top             =   5280
            Width           =   1455
         End
         Begin VB.CommandButton cmdCerrar 
            Caption         =   "Cerrar"
            Height          =   375
            Left            =   10080
            TabIndex        =   7
            Top             =   5280
            Width           =   1455
         End
         Begin SICMACT.FlexEdit feCreditos 
            Height          =   4935
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   14175
            _ExtentX        =   25003
            _ExtentY        =   8705
            Cols0           =   13
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Fecha Solicitud-Nª Credito-Titular-Moneda-Monto-MontoMN-Total Vinculados-Disp. de Solic.-Vinculado A-Relación-Asigna-Aux"
            EncabezadosAnchos=   "500-1200-1800-3500-1200-1200-1500-1500-1500-3500-1200-800-0"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-L-C-C-C-C-C-C-C-C-C"
            FormatosEdit    =   "0-5-0-0-0-0-0-0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   495
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Monto Máximo por Colaborador MN:"
            Height          =   195
            Left            =   240
            TabIndex        =   6
            Top             =   5280
            Width           =   2550
         End
         Begin VB.Label lblMontoMax 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2880
            TabIndex        =   5
            Top             =   5280
            Width           =   1935
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Disponible despues de Asignación"
         Height          =   195
         Left            =   9600
         TabIndex        =   12
         Top             =   600
         Width           =   2880
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Saldo por Asignar MN"
         Height          =   195
         Left            =   5760
         TabIndex        =   11
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label lblSaldoDA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12600
         TabIndex        =   10
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblSaldoPorAsignar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7440
         TabIndex        =   9
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblSaldoMN 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3480
         TabIndex        =   2
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Actual Disponible Para Asignación MN:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   3180
      End
   End
End
Attribute VB_Name = "frmCredSaldosVincAsignar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ORCR20140314**************************
'Private fgFecActual As Date
'Private fn7PorcPatriEfec As Double
'Private fn5PorcDel7PorcPatriEfec As Double
'Private fnReservaCred As Double
'Private fnPatrimonioEfec As Double
'
'Private fnSaldosVinculados As Double
'Private fnSaldosAceptados As Double
'
'Private fnTipoAsig As Double
'Private fnSaldoActual As Double
'Private fnSaldoActualAux As Double
'Private fnMontoMax As Double
'Private i As Integer
Private ffFecActual As Date
Private fnPatrimonioEfec As Double

Private fnSaldosAceptados As Double

Private fnTipoAsig As Double
Private fnSaldoActual As Double
Private fnMontoMax As Double
'END ORCR20140314**************************

Private Sub cmdAsignar_Click()
    Dim i As Integer
    If feCreditos.TextMatrix(1, 1) = "" Then
        MsgBox "No se encontraron datos para realizar el proceso.", vbCritical, "Aviso"
        Exit Sub
    End If
    If MsgBox("Estas seguro de procesar la Asignación de Saldos", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Dim oCredito As COMNCredito.NCOMCredito
        Dim sMovNroAsg As String
        Dim nOK As Integer
        Dim sMgsError As String
        Set oCredito = New COMNCredito.NCOMCredito
        For i = 1 To feCreditos.Rows - 1
            If CInt(feCreditos.TextMatrix(i, 12)) = 1 Then
                sMovNroAsg = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                nOK = oCredito.ActualizaCredVinculados(Trim(feCreditos.TextMatrix(i, 2)), sMovNroAsg, 1, 2)
                
                If nOK = 0 Then
                    sMgsError = "Ha ocurrido un Error al Asignar el Saldo del Cliente: " & feCreditos.TextMatrix(i, 3) & Chr(10) & "¿Desea Continuar?"
                    If MsgBox(sMgsError, vbCritical + vbYesNo, "Error") = vbNo Then
                        MsgBox "El formulario se cerrará", vbInformation, "Aviso"
                        Unload Me
                    End If
                End If
            End If
        Next i
        CargaValoresParametros
        CargaDatos
    End If
End Sub
Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    CargaValoresParametros
    CargaDatos
End Sub
'ORCR20140314**************************
'Private Function ObtenerSaldos(ByVal psAnio As String, ByVal psMes As String) As Double
'    Dim oNContabilidad As COMNContabilidad.NCOMContFunciones
'    Dim nSaldo As Double
'    Set oNContabilidad = New COMNContabilidad.NCOMContFunciones
'
'    nSaldo = oNContabilidad.PatrimonioEfecAjustInfl(psAnio, psMes)
'    ObtenerSaldos = nSaldo
'    Set oNContabilidad = Nothing
'End Function
'END ORCR20140314**************************

Public Sub Inicio(ByVal pnTipo As Integer)
    'ORCR20140314**************************
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
    'END ORCR20140314**************************
    fnTipoAsig = pnTipo
    Select Case fnTipoAsig
        Case 1: Me.Caption = "Asignación de Saldos a Colaboradores y Vinculados[Créditos]"
        Case 2: Me.Caption = "Asignación de Saldos a Colaboradores y Vinculados[Ventanilla]"
    End Select
    
    Me.Show 1
End Sub

Private Sub CargaValoresParametros()
    '    Dim sAnio As String
    '    Dim sMes As String
    '    Dim oPar As COMDCredito.DCOMParametro
    '    Dim oCredito As COMNCredito.NCOMCredito
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
    '    fnReservaCred = oPar.RecuperaValorParametro(102751)
    '    fn7PorcPatriEfec = oPar.RecuperaValorParametro(102752) / 100
    '    fn5PorcDel7PorcPatriEfec = oPar.RecuperaValorParametro(102753) / 100
    '
    '    Set oCredito = New COMNCredito.NCOMCredito
    '    fnSaldosVinculados = oCredito.ObtenerSaldoTrabDirVinc
    '    fnSaldosAceptados = oCredito.ObtenerSaldoAsignaEstado(2)
    '
    '    fnSaldoActual = fnPatrimonioEfec - fnSaldosVinculados - fnReservaCred - fnSaldosAceptados
    '    fnMontoMax = fn5PorcDel7PorcPatriEfec * fn7PorcPatriEfec * fnPatrimonioEfec
    '
    '    fnSaldoActualAux = fnSaldoActual
    '    Me.lblSaldoMN.Caption = Format(fnSaldoActual, "###," & String(15, "#") & "#0.00")
    '    Me.lblMontoMax.Caption = Format(fnMontoMax, "###," & String(15, "#") & "#0.00")
    '
    '    'Set oConsSist = Nothing
    '    Set oPar = Nothing
    '    Set oCredito = Nothing
    'ORCR20140314**************************
    Dim nReservaCred As Double
    Dim n7PorcPatriEfec As Double
    Dim n5PorcDel7PorcPatriEfec As Double
    
    Dim fnSaldosVinculados As Double
    
    Dim sAnio As String
    Dim sMes As String
    
    Dim oPar As New COMDCredito.DCOMParametro
    Dim oCredito As New COMNCredito.NCOMCredito

    sAnio = Year(ffFecActual)
    sMes = Format(Month(ffFecActual), "00")
    
    nReservaCred = oPar.RecuperaValorParametro(102751)
    n7PorcPatriEfec = oPar.RecuperaValorParametro(102752) / 100
    n5PorcDel7PorcPatriEfec = oPar.RecuperaValorParametro(102753) / 100
    
    fnSaldosVinculados = oCredito.ObtenerSaldoTrabDirVinc
    fnSaldosAceptados = oCredito.ObtenerSaldoAsignaEstado(2)
    
    fnSaldoActual = (fnPatrimonioEfec * n7PorcPatriEfec) - fnSaldosVinculados - nReservaCred - fnSaldosAceptados
    fnMontoMax = n5PorcDel7PorcPatriEfec * n7PorcPatriEfec * fnPatrimonioEfec
    
    Me.lblSaldoMN.Caption = Format(fnSaldoActual, "###," & String(15, "#") & "#0.00")
    Me.lblMontoMax.Caption = Format(fnMontoMax, "###," & String(15, "#") & "#0.00")
    
    
    Set oPar = Nothing
    Set oCredito = Nothing
    'END ORCR20140314**************************
End Sub

Private Sub CargaDatos()
    '    Dim rsCredito As ADODB.Recordset
    '    Dim oCredito As COMNCredito.NCOMCredito
    '    Dim nDispSolic As Double
    '    Set oCredito = New COMNCredito.NCOMCredito
    '    Set rsCredito = oCredito.ObtenerCreditosAAsignar(fnTipoAsig)
    'ORCR20140314**************************
    Dim rsCredito As ADODB.Recordset
    Dim oCredito As COMNCredito.NCOMCredito
    Dim nDispSolic As Double
    Dim nSaldoActualAux As Double
    
    Dim i As Integer
    
    Set oCredito = New COMNCredito.NCOMCredito
    Set rsCredito = oCredito.ObtenerCreditosAAsignar(fnTipoAsig)
    'END ORCR20140314**************************

    LimpiaFlex feCreditos
    If Not (rsCredito.EOF And rsCredito.BOF) Then
        For i = 1 To rsCredito.RecordCount
            nDispSolic = 0
            nDispSolic = fnMontoMax - CDbl(rsCredito!TotalVinc)
            feCreditos.AdicionaFila
            feCreditos.TextMatrix(i, 1) = Format(rsCredito!Fecha, "dd/mm/yyyy")
            feCreditos.TextMatrix(i, 2) = rsCredito!cCtaCod
            feCreditos.TextMatrix(i, 3) = rsCredito!cPersNombre
            feCreditos.TextMatrix(i, 4) = rsCredito!Moneda
            feCreditos.TextMatrix(i, 5) = Format(CDbl(rsCredito!nMonto), "###," & String(15, "#") & "#0.00")
            feCreditos.TextMatrix(i, 6) = Format(CDbl(rsCredito!nMontoMN), "###," & String(15, "#") & "#0.00")
            feCreditos.TextMatrix(i, 7) = Format(CDbl(rsCredito!TotalVinc), "###," & String(15, "#") & "#0.00")
            feCreditos.TextMatrix(i, 8) = Format(nDispSolic, "###," & String(15, "#") & "#0.00")
            feCreditos.TextMatrix(i, 9) = rsCredito!Vinculado
            feCreditos.TextMatrix(i, 10) = rsCredito!Relacion
            'ORCR20140314**************************
'           If nDispSolic > CDbl(rsCredito!nMontoMN) And fnSaldoActualAux > CDbl(rsCredito!nMontoMN) Then
            If (nDispSolic >= CDbl(rsCredito!nMontoMN) And (fnSaldoActual - nSaldoActualAux) > CDbl(rsCredito!nMontoMN)) Then
                feCreditos.TextMatrix(i, 11) = "SI"
                feCreditos.TextMatrix(i, 12) = "1"
                
                nSaldoActualAux = nSaldoActualAux + CDbl(rsCredito!nMontoMN)
            Else
                feCreditos.TextMatrix(i, 11) = "NO"
                feCreditos.TextMatrix(i, 12) = "0"
            End If
            
'            If feCreditos.TextMatrix(i, 12) = "1" Then
'                fnSaldoActualAux = fnSaldoActualAux - CDbl(rsCredito!nMontoMN)
'            End If
            'END ORCR20140314**************************
            rsCredito.MoveNext
        Next i
    End If
    'ORCR20140314**************************
'    Me.lblSaldoMN.Caption = Format(fnSaldoActualAux, "###," & String(15, "#") & "#0.00")
    
    Me.lblSaldoPorAsignar.Caption = Format(nSaldoActualAux, "###," & String(15, "#") & "#0.00")
    Me.lblSaldoDA.Caption = Format(fnSaldoActual - nSaldoActualAux, "###," & String(15, "#") & "#0.00")
    'END ORCR20140314**************************

    Set oCredito = Nothing
    Set rsCredito = Nothing
End Sub

