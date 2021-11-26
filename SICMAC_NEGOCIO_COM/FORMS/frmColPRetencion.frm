VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmColPRetencion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Retención de Créditos "
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9105
   Icon            =   "frmColPRetencion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab_Retencion 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Config. de Tasas. "
      TabPicture(0)   =   "frmColPRetencion.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdSalir"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdGuardar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCancelar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Retención de créditos."
      TabPicture(1)   =   "frmColPRetencion.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblCampRetencionLeyenda"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ActXCtaCredRC"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "frConfRC"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdSalirRC"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdGrabarRC"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdCancelarRC"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdBuscarRC"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.CommandButton cmdBuscarRC 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   -71280
         TabIndex        =   19
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelarRC 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   -67080
         TabIndex        =   17
         Top             =   4760
         Width           =   855
      End
      Begin VB.CommandButton cmdGrabarRC 
         Caption         =   "Grabar"
         Height          =   315
         Left            =   -68040
         TabIndex        =   16
         Top             =   4760
         Width           =   855
      End
      Begin VB.CommandButton cmdSalirRC 
         Caption         =   "Salir"
         Height          =   315
         Left            =   -74880
         TabIndex        =   15
         Top             =   4760
         Width           =   855
      End
      Begin VB.Frame frConfRC 
         Caption         =   "Configuración de T.E.A"
         Height          =   3735
         Left            =   -74880
         TabIndex        =   8
         Top             =   960
         Width           =   8655
         Begin VB.Frame frAplicaRC 
            Caption         =   "Si aplica retención"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1935
            Left            =   120
            TabIndex        =   9
            Top             =   1680
            Width           =   8415
            Begin SICMACT.EditMoney emTEAMinRC 
               Height          =   255
               Left            =   840
               TabIndex        =   20
               Top             =   360
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   450
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "0"
            End
            Begin VB.TextBox txtObserRC 
               Height          =   855
               Left            =   120
               TabIndex        =   11
               Top             =   960
               Width           =   8175
            End
            Begin SICMACT.EditMoney emTEAMaxRC 
               Height          =   255
               Left            =   2895
               TabIndex        =   21
               Top             =   360
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   450
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "0"
            End
            Begin SICMACT.EditMoney emTEANewRC 
               Height          =   255
               Left            =   5160
               TabIndex        =   26
               Top             =   360
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   450
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "0"
            End
            Begin SICMACT.EditMoney emTEMNewRC 
               Height          =   255
               Left            =   7200
               TabIndex        =   29
               Top             =   360
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   450
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "0"
            End
            Begin VB.Label Label5 
               Caption         =   "Nueva TEM."
               Height          =   255
               Left            =   6240
               TabIndex        =   28
               Top             =   360
               Width           =   975
            End
            Begin VB.Label lblPorNew 
               Caption         =   "%"
               Enabled         =   0   'False
               Height          =   255
               Left            =   5890
               TabIndex        =   25
               Top             =   390
               Width           =   255
            End
            Begin VB.Label lblPorMax 
               Caption         =   "%"
               Enabled         =   0   'False
               Height          =   255
               Left            =   3610
               TabIndex        =   24
               Top             =   390
               Width           =   255
            End
            Begin VB.Label lblPorMin 
               Caption         =   "%"
               Enabled         =   0   'False
               Height          =   255
               Left            =   1570
               TabIndex        =   23
               Top             =   390
               Width           =   255
            End
            Begin VB.Label Label4 
               Caption         =   "Observaciones"
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label Label3 
               Caption         =   "Nueva TEA."
               Height          =   255
               Left            =   4200
               TabIndex        =   14
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label2 
               Caption         =   "TEA Max."
               Height          =   255
               Left            =   2160
               TabIndex        =   13
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label1 
               Caption         =   "TEA Min."
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   360
               Width           =   735
            End
         End
         Begin SICMACT.FlexEdit feDatosRC 
            Height          =   1260
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   8445
            _ExtentX        =   14896
            _ExtentY        =   2223
            Cols0           =   7
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Cliente-DOI-TEA Actual-Saldo-Monto Min. Renovar-aux"
            EncabezadosAnchos=   "400-5200-1100-1000-1200-0-0"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-R-R-C-C"
            FormatosEdit    =   "0-0-3-2-2-2-2"
            CantEntero      =   10
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   4800
         TabIndex        =   6
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   3720
         TabIndex        =   5
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   4680
         Width           =   975
      End
      Begin VB.Frame frDatos 
         Caption         =   "Rangos y Tasas"
         Height          =   4215
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   8655
         Begin VB.CommandButton cmdQuitar 
            Caption         =   "Quitar"
            Height          =   375
            Left            =   1200
            TabIndex        =   3
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmdAñadir 
            Caption         =   "Añadir"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   975
         End
         Begin SICMACT.FlexEdit feConfigTasas 
            Height          =   3300
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   5565
            _ExtentX        =   9816
            _ExtentY        =   5821
            Cols0           =   6
            ScrollBars      =   2
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Monto Inicio-Monto Fin-TEA Min-TEA Max-nid"
            EncabezadosAnchos=   "400-1200-1200-1200-1200-0"
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
            ColumnasAEditar =   "X-1-2-3-4-X"
            ListaControles  =   "0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-R-R-R-R-R"
            FormatosEdit    =   "0-2-2-2-2-3"
            CantEntero      =   10
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin SICMACT.ActXCodCta ActXCtaCredRC 
         Height          =   390
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   3585
         _ExtentX        =   6535
         _ExtentY        =   688
         Texto           =   "Crédito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
         CMAC            =   "109"
      End
      Begin VB.Label lblCampRetencionLeyenda 
         Caption         =   "Después de registrar la retención, por favor realizar la renovación en el mismo día, caso contrario perderá el beneficio."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   -70200
         TabIndex        =   27
         Top             =   360
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmColPRetencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'Archivo:  frmColPRetencion.frm
'JOEP   :  06/09/2021
'Resumen:  Registro y Configuracion de Pignoraticios
'******************************************************

Option Explicit

Dim nTpConfigTasa As Integer
Dim nTpRenteCred As Integer
Public Enum TpOpRete
    TpConfigTasa = 1
    TpRenteCred = 2
End Enum

'********************* Tab Retencion de creditos ****************************
Public Sub RetencionCred()
    SSTab_Retencion.TabVisible(0) = False
    nTpRenteCred = TpRenteCred
    nTpConfigTasa = 0
    Call HabiDeshaControles(TpRenteCred, False)
    Me.Show 1
End Sub

Private Sub cmdGrabarRC_Click()
Dim objRegR As COMDColocPig.DCOMColPContrato
Dim bRegR As Boolean
Dim sMovNroR As String
Dim cMovNroVistoBueno As String
Dim lbResultadoVisto As String
Dim loVistoElectronico As SICMACT.frmVistoElectronico
Set loVistoElectronico = New SICMACT.frmVistoElectronico

Set objRegR = New COMDColocPig.DCOMColPContrato
sMovNroR = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    If ValidacionRegRetencion = False Then
        Exit Sub
    End If
    
    If ValidadDatos(ActXCtaCredRC.NroCuenta, emTEANewRC) = False Then
        Exit Sub
    End If
    
    If MsgBox("Está seguro de registrar los datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    
    lbResultadoVisto = loVistoElectronico.Inicio(23, "910000", "")
        If lbResultadoVisto = True Then
            cMovNroVistoBueno = loVistoElectronico.cMovNroRetencion
            Call objRegR.CampPrenRegRetencion(ActXCtaCredRC.NroCuenta, emTEAMinRC, emTEAMaxRC, emTEANewRC, txtObserRC.Text, sMovNroR, cMovNroVistoBueno, 0)
            MsgBox "Datos registrados correctamente", vbInformation, "Aviso"
            Call cmdCancelarRC_Click
        End If
    End If
End Sub

Private Function ValidacionRegRetencion() As Boolean
ValidacionRegRetencion = True
    
    If (emTEAMinRC = "" Or emTEANewRC = "" Or emTEAMaxRC = "") Then
        MsgBox "falta ingresar las TEAs", vbInformation, "Aviso"
        ValidacionRegRetencion = False
        Exit Function
    End If

    If Not (CDbl(emTEANewRC) >= CDbl(emTEAMinRC) And CDbl(emTEANewRC) <= CDbl(emTEAMaxRC)) Then
        MsgBox "La nueva TEA esta fuera del rango", vbInformation, "Aviso"
        ValidacionRegRetencion = False
        emTEANewRC.SetFocus
        Exit Function
    End If
    
    If txtObserRC.Text = "" Then
        MsgBox "Ingrese datos en la Observacion", vbInformation, "Aviso"
        ValidacionRegRetencion = False
        txtObserRC.SetFocus
        Exit Function
    End If
End Function

Private Sub ActXCtaCredRC_KeyPress(KeyAscii As Integer)
Dim objBus As COMDColocPig.DCOMColPContrato
Dim rsBus As ADODB.Recordset
Dim i As Integer
Set objBus = New COMDColocPig.DCOMColPContrato

If ActXCtaCredRC.NroCuenta <> "" Then
  Set rsBus = objBus.CampPrendarioObtDatosPers(ActXCtaCredRC.NroCuenta)
  Call HabiDeshaControles(TpRenteCred, True)
    If Not (rsBus.BOF And rsBus.EOF) Then
        If rsBus!nPase <> 0 Then
            LimpiaFlex feDatosRC
            For i = 1 To rsBus.RecordCount
                feDatosRC.AdicionaFila
                feDatosRC.TextMatrix(i, 1) = rsBus!cPersNombre
                feDatosRC.TextMatrix(i, 2) = rsBus!cPersIDnro
                feDatosRC.TextMatrix(i, 3) = Format(rsBus!TEA, "#0.00")
                feDatosRC.TextMatrix(i, 4) = Format(rsBus!nSaldo, "#0.00")
                feDatosRC.TextMatrix(i, 5) = Format(rsBus!nMontoMinRenovar, "#0.00")
                emTEAMinRC.Text = Format(rsBus!nTEAInicio, "#0.00")
                emTEAMaxRC.Text = Format(rsBus!nTEAFin, "#0.00")
                rsBus.MoveNext
            Next i
        Else
            MsgBox rsBus!cObs, vbInformation, "Aviso"
            Call HabiDeshaControles(TpRenteCred, False)
            Call cmdCancelarRC_Click
        End If
    End If
Else
    MsgBox "Ingrese el Nº Credito", vbInformation, "Aviso"
End If

Set objBus = Nothing
RSClose rsBus
End Sub

Private Sub cmdCancelarRC_Click()
   LimpiaFlex feDatosRC
   emTEAMinRC.Text = Format(0, "#0.00")
   emTEAMaxRC.Text = Format(0, "#0.00")
   emTEANewRC.Text = Format(0, "#0.00")
   emTEMNewRC.Text = Format(0, "#0.00")
   txtObserRC.Text = ""
   ActXCtaCredRC.NroCuenta = ""
   ActXCtaCredRC.CMAC = "109"
    
End Sub

Private Sub cmdBuscarRC_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersContrato As COMDColocPig.DCOMColPContrato
Dim lrContratos As New ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

Call cmdCancelarRC_Click
Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

' Selecciona Estados
lsEstados = gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov

If Trim(lsPersCod) <> "" Then
    Set loPersContrato = New COMDColocPig.DCOMColPContrato
        Set lrContratos = loPersContrato.dObtieneCredPigDePersona(lsPersCod, lsEstados, Mid(gsCodAge, 4, 2))
    Set loPersContrato = Nothing
End If

Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmCuentasPersona.Inicio(lsPersNombre, lrContratos)
    If loCuentas.sCtaCod <> "" Then
        ActXCtaCredRC.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        ActXCtaCredRC.SetFocusCuenta
    End If
Set loCuentas = Nothing
End Sub

Private Sub cmdSalirRC_Click()
    Unload Me
End Sub

Private Sub emTEANewRC_KeyPress(KeyAscii As Integer)
    If Len(emTEANewRC) = 2 And Right(emTEANewRC, 1) = "." And Left(emTEANewRC, 1) = "." Then
        emTEANewRC = Format(0, "#0.00")
    End If
    If Len(emTEANewRC) = 1 And KeyAscii = 13 Then
        emTEANewRC = Format(0, "#0.00")
    End If
    If Left(emTEANewRC, 1) = "." Then
        emTEANewRC = Format(0, "#0.00")
    End If
    
    If emTEANewRC > 100 Then
        emTEANewRC = 100
    End If
    
    If KeyAscii = 13 Then
        emTEMNewRC = nTEATEM(emTEANewRC)
        txtObserRC.SetFocus
    End If
End Sub

'********************* Tab Retencion de creditos ****************************

Private Sub HabiDeshaControles(ByVal pnOpion As Integer, ByVal bDtAD As Boolean)
    If pnOpion = TpRenteCred Then
        emTEANewRC.Enabled = bDtAD
        txtObserRC.Enabled = bDtAD
        cmdGrabarRC.Enabled = bDtAD
    End If
End Sub

'********************* Tab Configuracion de tasas ****************************
Public Sub ConfigDatos()
    SSTab_Retencion.TabVisible(1) = False
    nTpRenteCred = 0
    nTpConfigTasa = TpConfigTasa
    frDatos.Width = 5775
    SSTab_Retencion.Width = 6000
    Me.Width = 6300
    Call CargaDatosConfig
    Me.Show 1
End Sub

Private Sub CargaDatosConfig()
Dim obj As COMDColocPig.DCOMColPContrato
Dim rs As ADODB.Recordset
Dim i As Integer
Set obj = New COMDColocPig.DCOMColPContrato

Set rs = obj.CampPrendarioConfigDatos

        feConfigTasas.Clear
        feConfigTasas.FormaCabecera
        Call LimpiaFlex(feConfigTasas)
If Not (rs.BOF And rs.EOF) Then
    For i = 1 To rs.RecordCount
        feConfigTasas.AdicionaFila
        feConfigTasas.TextMatrix(i, 1) = Format(rs!nMontoInicio, "#,#0.00")
        feConfigTasas.TextMatrix(i, 2) = Format(rs!nMontoFin, "#,#0.00")
        feConfigTasas.TextMatrix(i, 3) = Format(rs!nTEAInicio, "#0.00")
        feConfigTasas.TextMatrix(i, 4) = Format(rs!nTEAFin, "#0.00")
        feConfigTasas.TextMatrix(i, 5) = rs!nId
        rs.MoveNext
    Next i
End If
    
Set obj = Nothing
RSClose rs
End Sub

Private Sub cmdAñadir_Click()
feConfigTasas.lbEditarFlex = True
feConfigTasas.AdicionaFila
Call AdicionaFila
End Sub

Private Sub cmdCancelar_Click()
    Call CargaDatosConfig
End Sub

Private Sub cmdGuardar_Click()
Dim objReg As COMNColoCPig.NCOMColPContrato
Dim bReg As Boolean
Dim sMovNro As String
Dim i As Integer
Set objReg = New COMNColoCPig.NCOMColPContrato

If ValidaDatos = False Then
    Exit Sub
End If

'llenamos la matriz con los datos
Dim nMatTasas As Variant
Set nMatTasas = Nothing
ReDim nMatTasas(feConfigTasas.rows - 1, 10)
    For i = 1 To feConfigTasas.rows - 1
        nMatTasas(i, 1) = feConfigTasas.TextMatrix(i, 1)
        nMatTasas(i, 2) = feConfigTasas.TextMatrix(i, 2)
        nMatTasas(i, 3) = feConfigTasas.TextMatrix(i, 3)
        nMatTasas(i, 4) = feConfigTasas.TextMatrix(i, 4)
        nMatTasas(i, 5) = IIf(feConfigTasas.TextMatrix(i, 5) = "", i, feConfigTasas.TextMatrix(i, 5))
    Next i
'llenamos la matriz con los datos

bReg = False
sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

    If MsgBox("Está seguro de registrar los datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    bReg = objReg.CampPrenRegTasas(nMatTasas, sMovNro)

    If bReg = True Then
        MsgBox "Datos registrados correctamente?", vbInformation, "Aviso"
    Else
        MsgBox "Hubo un error al registrar los datos, volver intentar", vbInformation, "Aviso"
    End If
        Call CargaDatosConfig
    End If
Set nMatTasas = Nothing
End Sub

Private Sub cmdQuitar_Click()
    If MsgBox("Está seguro de eliminar el registro de la fila " & feConfigTasas.row & "?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        feConfigTasas.EliminaFila (feConfigTasas.row)
    End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub AdicionaFila()
    If (feConfigTasas.rows - 1) = 1 Then
        feConfigTasas.TextMatrix(feConfigTasas.row, 1) = Format(0, "#0.00")
    Else
        If feConfigTasas.TextMatrix(feConfigTasas.row - 1, 2) <> "" Then
            feConfigTasas.TextMatrix(feConfigTasas.row, 1) = Format(CDbl(feConfigTasas.TextMatrix(feConfigTasas.row - 1, 2)) + 0.01, "#,#0.00")
        Else
            MsgBox "Ingrese el monto en la Columna [Monto Fin], fila " & feConfigTasas.row - 1, vbInformation, "Aviso"
            feConfigTasas.EliminaFila (feConfigTasas.row)
        End If
    End If
End Sub

Private Function ValidaDatos() As Boolean
Dim i As Integer
Dim J As Integer
ValidaDatos = True

For i = 1 To feConfigTasas.rows - 1
    If feConfigTasas.TextMatrix(i, 1) = "" Or feConfigTasas.TextMatrix(i, 2) = "" Or feConfigTasas.TextMatrix(i, 3) = "" Or feConfigTasas.TextMatrix(i, 4) = "" Then
        MsgBox "Ingrese los datos", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
Next i

'Verificar la distribucion de los montos
For i = 1 To feConfigTasas.rows - 2
    If CDbl(feConfigTasas.TextMatrix(i, 2)) <> CDbl(feConfigTasas.TextMatrix(i + 1, 1)) - 0.01 Then
        MsgBox "La distribucion de los monto es incorrecto verificar la fila " & feConfigTasas.row, vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
Next i
'Verificar que el monto anterior no sea mayor
For i = 1 To feConfigTasas.rows - 2
    If CDbl(feConfigTasas.TextMatrix(i + 1, 2)) < CDbl(feConfigTasas.TextMatrix(i, 2)) Then
        MsgBox "La distribucion de los monto es incorrecto verificar la fila " & feConfigTasas.row, vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
Next i

End Function

Private Sub emTEANewRC_Validate(Cancel As Boolean)
    If Len(emTEANewRC) = 1 And emTEANewRC = "." Then
        emTEANewRC = Format(0, "#0.00")
    End If
    If Len(emTEANewRC) = 2 And Right(emTEANewRC, 1) = "." Then
        emTEANewRC = Format(0, "#0.00")
    End If
End Sub

Private Sub feConfigTasas_EnterCell()
    If feConfigTasas.col = 3 Then
        feConfigTasas.TextMatrix(feConfigTasas.row, 3) = feConfigTasas.TextMatrix(feConfigTasas.row, 3)
        feConfigTasas.TextMatrix(feConfigTasas.row, 3) = feConfigTasas.TextMatrix(feConfigTasas.row, 3)
    End If
    If feConfigTasas.col = 4 Then
        feConfigTasas.TextMatrix(feConfigTasas.row, 4) = feConfigTasas.TextMatrix(feConfigTasas.row, 4)
        feConfigTasas.TextMatrix(feConfigTasas.row, 4) = feConfigTasas.TextMatrix(feConfigTasas.row, 4)
    End If
End Sub

Private Sub feConfigTasas_KeyPress(KeyAscii As Integer)
If (feConfigTasas.col = 1 Or feConfigTasas.col = 2 Or feConfigTasas.col = 3 Or feConfigTasas.col = 4) And (KeyAscii = 45 Or KeyAscii = 46) Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub feConfigTasas_OnCellChange(pnRow As Long, pnCol As Long)

    If feConfigTasas.col = 2 Then
        If feConfigTasas.row < (feConfigTasas.rows - 1) Then
            feConfigTasas.TextMatrix(feConfigTasas.row + 1, 1) = Format(Replace(feConfigTasas.TextMatrix(feConfigTasas.row, 2) + 0.01, "%", ""), "#,#0.00")
        End If
    End If
    If feConfigTasas.col = 3 Then
        feConfigTasas.TextMatrix(feConfigTasas.row, 3) = feConfigTasas.TextMatrix(feConfigTasas.row, 3)
        feConfigTasas.TextMatrix(feConfigTasas.row, 3) = feConfigTasas.TextMatrix(feConfigTasas.row, 3)
    End If
    If feConfigTasas.col = 4 Then
        feConfigTasas.TextMatrix(feConfigTasas.row, 4) = feConfigTasas.TextMatrix(feConfigTasas.row, 4)
        feConfigTasas.TextMatrix(feConfigTasas.row, 4) = feConfigTasas.TextMatrix(feConfigTasas.row, 4)
    End If
    
End Sub
'************************************ fin ******************************************

Private Sub feConfigTasas_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
If (feConfigTasas.col = 1 Or feConfigTasas.col = 2 Or feConfigTasas.col = 3 Or feConfigTasas.col = 4) And IsNumeric(feConfigTasas.TextMatrix(feConfigTasas.row, feConfigTasas.col)) = False Then
        Cancel = False
        SendKeys "{TAB}"
    End If
End Sub

'Para evitar copiar y pegar en la grilla
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        SendKeys "{Enter}"
    End If
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
End Sub

Public Function nTEATEM(ByVal pTEA As Double)
    nTEATEM = Round((((pTEA / 100# + 1) ^ (1 / 12#)) - 1) * 100#, 4)
End Function

Public Function ValidadDatos(ByVal pcCtaCod As String, ByVal pnTEANew As Double) As Boolean
ValidadDatos = True
Dim oVal As COMDColocPig.DCOMColPContrato
Dim rs As ADODB.Recordset
Set oVal = New COMDColocPig.DCOMColPContrato

Set rs = oVal.CampPrendarioValidadReferido(pcCtaCod, pnTEANew)
If Not (rs.BOF And rs.EOF) Then
    If rs!cMensaje <> "" Then
        MsgBox rs!cMensaje, vbInformation, "Aviso"
        ValidadDatos = False
        Exit Function
    End If
End If

Set oVal = Nothing
RSClose rs
End Function
