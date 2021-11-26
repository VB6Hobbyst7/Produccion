VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGarantiaValorAutoLiquidable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GARANTÍA AUTOLIQUIDABLE"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4995
   Icon            =   "frmGarantiaValorAutoLiquidable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2460
      TabIndex        =   1
      ToolTipText     =   "Cancelar"
      Top             =   3165
      Width           =   1000
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1395
      TabIndex        =   0
      ToolTipText     =   "Aceptar"
      Top             =   3165
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   75
      TabIndex        =   2
      Top             =   75
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5318
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Plazo Fijo"
      TabPicture(0)   =   "frmGarantiaValorAutoLiquidable.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fraDatos 
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
         Height          =   2535
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4575
         Begin VB.ComboBox cmbCuenta 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Monto Cuenta (Utilizable) :"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   1335
            Width           =   1860
         End
         Begin VB.Label lblMontoCuentaDisp 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2040
            TabIndex        =   14
            Top             =   1335
            Width           =   1380
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Disponible:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   2055
            Width           =   780
         End
         Begin VB.Label lblMontoDisponible 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2040
            TabIndex        =   12
            Top             =   2055
            Width           =   1380
         End
         Begin VB.Label lblFormaRetiro 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2040
            TabIndex        =   11
            Top             =   1695
            Width           =   2340
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Forma retiro interés:"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   1695
            Width           =   1380
         End
         Begin VB.Label lblMontoCuenta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2040
            TabIndex        =   9
            Top             =   975
            Width           =   1380
         End
         Begin VB.Label lblMoneda 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2040
            TabIndex        =   8
            Top             =   615
            Width           =   1380
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta:"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   255
            Width           =   1755
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   615
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Monto Cuenta:"
            Height          =   195
            Left            =   135
            TabIndex        =   5
            Top             =   945
            Width           =   1050
         End
      End
   End
End
Attribute VB_Name = "frmGarantiaValorAutoLiquidable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************************************
'** Nombre : frmGarantiaValorAutoliquidable
'** Descripción : Para registro/edición/consulta de Valorización Autoliquida creado segun TI-ERS063-2014
'** Creación : EJVG, 20150116 05:10:01 PM
'*******************************************************************************************************
Option Explicit

Dim fbRegistrar As Boolean
Dim fbEditar As Boolean
Dim fbConsultar As Boolean

Dim fsPersCod As String
Dim fvValorAutoLiquidable As tValorAutoLiquidable
Dim fvListaAutoLiquidable() As tValorAutoLiquidable
Dim fsNumGarant As String

Dim fbPrimero As Boolean
Dim fbOk As Boolean

Dim fnIndex As Integer
Dim fnMontoCuentaIni As Currency
Dim fnMontoCuentaDispIni As Currency
Dim fnMontoDisponibleIni As Currency

Dim fbCancelado As Boolean

Private Sub cmbCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl CmdAceptar
    End If
End Sub

Private Sub cmdAceptar_Click()
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim lsCtaCod As String
    Dim lsValida As String
    
    On Error GoTo ErrAceptar
    If cmbCuenta.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar la Cuenta, verifique el Titular tenga Cuentas Activas en el Sistema", vbInformation, "Aviso"
        EnfocaControl cmbCuenta
        Exit Sub
    End If
    
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    lsCtaCod = Trim(Right(cmbCuenta.Text, 18))
    
    If fbPrimero Then
        lsValida = clsCap.ValidaCuentaOperacion(lsCtaCod)
        If Len(lsValida) > 0 Then
            MsgBox lsValida, vbInformation, "Aviso"
            Set clsCap = Nothing
            Exit Sub
        End If
    Else
        If fbCancelado Then
            MsgBox "La cuenta se encuentra CANCELADA o ANULADA", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
    If fvListaAutoLiquidable(fnIndex).nformaretiro = 4 Then 'Interes adelantado->Primero debieron realizar el retiro
        If Not clsCap.RealizoInteresCash(lsCtaCod) Then
            MsgBox "No se puede continuar ya que aún no se ha realizado el Retiro de Interes adelantado del DPF", vbInformation, "Aviso"
            Set clsCap = Nothing
            Exit Sub
        End If
    End If
    Set clsCap = Nothing
    
    If fvListaAutoLiquidable(fnIndex).nMontoDisponible <= 0# Then
        MsgBox "No existe disponible en la cuenta seleccionada", vbInformation, "Aviso"
        Exit Sub
    End If

    fvValorAutoLiquidable = fvListaAutoLiquidable(fnIndex)

    fbOk = True
    Unload Me
    Exit Sub
ErrAceptar:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Sub Form_Load()
    fbOk = False
    fbCancelado = False
    Screen.MousePointer = 11
    
    CargarControles
    LimpiarControles

    If fbEditar Or fbConsultar Then
        cmbCuenta.ListIndex = IndiceListaCombo(cmbCuenta, fvListaAutoLiquidable(fnIndex).sCtaCod, , 18)
        If fbConsultar Then 'Mostramos los montos como se guardo en la ultima valorización
            fvListaAutoLiquidable(fnIndex).nMontoCuenta = fnMontoCuentaIni
            fvListaAutoLiquidable(fnIndex).nMontoCuentaDisp = fnMontoCuentaDispIni
            fvListaAutoLiquidable(fnIndex).nMontoDisponible = fnMontoDisponibleIni
        End If
        lblFormaRetiro.Caption = fvListaAutoLiquidable(fnIndex).sFormaRetiro
        lblMontoCuenta.Caption = Format(fvListaAutoLiquidable(fnIndex).nMontoCuenta, "#,##0.00")
        lblMontoCuentaDisp.Caption = Format(fvListaAutoLiquidable(fnIndex).nMontoCuentaDisp, "#,##0.00")
        lblMontoDisponible.Caption = Format(fvListaAutoLiquidable(fnIndex).nMontoDisponible, "#,##0.00")
        
        If fbConsultar Then
            FraDatos.Enabled = False
            CmdAceptar.Enabled = False
        End If
        If fbEditar Then
            If Not fbPrimero Then
                'cmbCuenta.Locked = True
            End If
        End If
    End If
    
    If fbRegistrar Then
        Caption = "GARANTÍA AUTOLIQUIDABLE [ NUEVO ]"
    End If
    If fbConsultar Then
        Caption = "GARANTÍA AUTOLIQUIDABLE [ CONSULTAR ]"
    End If
    If fbEditar Then
        Caption = "GARANTÍA AUTOLIQUIDABLE [ EDITAR ]"
    End If
    
    Screen.MousePointer = 0
End Sub
Public Function Registrar(ByVal pbPrimero As Boolean, ByVal psPersCod As String, ByRef pnMoneda As Moneda, ByRef pvValorAutoLiquidable As tValorAutoLiquidable, ByVal psNumGarant As String) As Boolean
    fbRegistrar = True
    fbPrimero = pbPrimero
    fsPersCod = psPersCod
    fvValorAutoLiquidable = pvValorAutoLiquidable
    fsNumGarant = psNumGarant
    Show 1
    pvValorAutoLiquidable = fvValorAutoLiquidable
    pnMoneda = IIf(pvValorAutoLiquidable.sCtaCod <> "", Mid(pvValorAutoLiquidable.sCtaCod, 9, 1), 0)
    
    Registrar = fbOk
End Function
Public Function Editar(ByVal pbPrimero As Boolean, ByVal psPersCod As String, ByRef pnMoneda As Moneda, ByRef pvValorAutoLiquidable As tValorAutoLiquidable, ByVal psNumGarant As String) As Boolean
    fbEditar = True
    fbPrimero = pbPrimero
    fsPersCod = psPersCod
    fvValorAutoLiquidable = pvValorAutoLiquidable
    fsNumGarant = psNumGarant
    Show 1
    pvValorAutoLiquidable = fvValorAutoLiquidable
    pnMoneda = IIf(pvValorAutoLiquidable.sCtaCod <> "", Mid(pvValorAutoLiquidable.sCtaCod, 9, 1), 0)
    
    Editar = fbOk
End Function
Public Sub Consultar(ByVal psPersCod As String, ByRef pvValorAutoLiquidable As tValorAutoLiquidable, ByVal psNumGarant As String)
    fbConsultar = True
    fsPersCod = psPersCod
    fvValorAutoLiquidable = pvValorAutoLiquidable
    fsNumGarant = psNumGarant
    Show 1
End Sub
Private Sub CargarControles()
    Dim oGarant As New COMDCredito.DCOMGarantia
    Dim rsCuenta As ADODB.Recordset
    
    fnIndex = 0
    ReDim fvListaAutoLiquidable(0)
    cmbCuenta.Clear
    If fvValorAutoLiquidable.sCtaCod = "" Or fbPrimero Then 'Muestra todos los DPF/CTS a la fecha
        Set rsCuenta = oGarant.ListaCuentasxGarantiaAutoLiq(fsPersCod, fsNumGarant)
        If Not rsCuenta.EOF Then
            Do While Not rsCuenta.EOF
                ReDim Preserve fvListaAutoLiquidable(rsCuenta.Bookmark)
                fvListaAutoLiquidable(rsCuenta.Bookmark).sCtaCod = rsCuenta!cCtaCod
                fvListaAutoLiquidable(rsCuenta.Bookmark).nformaretiro = rsCuenta!nformaretiro
                fvListaAutoLiquidable(rsCuenta.Bookmark).sFormaRetiro = rsCuenta!cFormaRetiro
                fvListaAutoLiquidable(rsCuenta.Bookmark).nMontoCuenta = rsCuenta!nSaldoCuenta
                fvListaAutoLiquidable(rsCuenta.Bookmark).nMontoCuentaNEW = fvListaAutoLiquidable(rsCuenta.Bookmark).nMontoCuenta
                fvListaAutoLiquidable(rsCuenta.Bookmark).nMontoCuentaDisp = rsCuenta!nSaldoCuentaDisp
                fvListaAutoLiquidable(rsCuenta.Bookmark).nMontoCuentaDispNEW = fvListaAutoLiquidable(rsCuenta.Bookmark).nMontoCuentaDisp
                fvListaAutoLiquidable(rsCuenta.Bookmark).nMontoDisponible = rsCuenta!nSaldoDisp
                fvListaAutoLiquidable(rsCuenta.Bookmark).nMontoDisponibleNEW = fvListaAutoLiquidable(rsCuenta.Bookmark).nMontoDisponible
                
                If fvValorAutoLiquidable.sCtaCod = fvListaAutoLiquidable(rsCuenta.Bookmark).sCtaCod Then
                    fnIndex = rsCuenta.Bookmark
                    fnMontoCuentaIni = fvListaAutoLiquidable(rsCuenta.Bookmark).nMontoCuenta
                    fnMontoCuentaDispIni = fvListaAutoLiquidable(rsCuenta.Bookmark).nMontoCuentaDisp
                    fnMontoDisponibleIni = fvListaAutoLiquidable(rsCuenta.Bookmark).nMontoDisponible
                End If
                
                cmbCuenta.AddItem fvListaAutoLiquidable(rsCuenta.Bookmark).sCtaCod
                rsCuenta.MoveNext
            Loop
        Else
            MsgBox "No se ha encontrado cuentas con el Titular de la Garantía", vbInformation, "Aviso"
        End If
    Else 'Cuando se va a utilizar la misma garantía para AutoLiquidables, debe ser la misma cuenta DPF/CTS, si en caso no lo es, debe crearse otra garantía
        Set rsCuenta = oGarant.ListaCuentasxGarantiaAutoLiq(fsPersCod, fsNumGarant, fvValorAutoLiquidable.sCtaCod)
        If Not rsCuenta.EOF Then
            fbCancelado = rsCuenta!bCancelado
            
            ReDim Preserve fvListaAutoLiquidable(1)
            fvListaAutoLiquidable(1) = fvValorAutoLiquidable
            fvListaAutoLiquidable(1).nMontoCuentaNEW = rsCuenta!nSaldoCuenta
            fvListaAutoLiquidable(1).nMontoCuentaDispNEW = rsCuenta!nSaldoCuentaDisp
            fvListaAutoLiquidable(1).nMontoDisponibleNEW = rsCuenta!nSaldoDisp
            fnIndex = 1
            fnMontoCuentaIni = fvListaAutoLiquidable(1).nMontoCuenta
            fnMontoCuentaDispIni = fvListaAutoLiquidable(1).nMontoCuentaDisp
            fnMontoDisponibleIni = fvListaAutoLiquidable(1).nMontoDisponible
            
            cmbCuenta.AddItem fvListaAutoLiquidable(1).sCtaCod
        End If
    End If
    RSClose rsCuenta
    Set oGarant = Nothing
End Sub
Private Sub cmbCuenta_Click()
    Dim lsCtaCod As String
    Dim i As Integer
    
    If cmbCuenta.ListIndex = -1 Then Exit Sub
    
    On Error GoTo ErrCargaCuenta
    Screen.MousePointer = 11
    
    lblmoneda.Caption = ""
    lblmoneda.BackColor = &H80000005
    lblFormaRetiro.Caption = ""
    lblMontoCuenta.Caption = "0.00"
    lblMontoDisponible.Caption = "0.00"
    
    lsCtaCod = cmbCuenta.Text
    CmdAceptar.Default = False
    For i = 1 To UBound(fvListaAutoLiquidable)
        If lsCtaCod = fvListaAutoLiquidable(i).sCtaCod Then
            fvListaAutoLiquidable(i).nMontoCuenta = fvListaAutoLiquidable(i).nMontoCuentaNEW
            fvListaAutoLiquidable(i).nMontoCuentaDisp = fvListaAutoLiquidable(i).nMontoCuentaDispNEW
            fvListaAutoLiquidable(i).nMontoDisponible = fvListaAutoLiquidable(i).nMontoDisponibleNEW
            
            lblmoneda.Caption = IIf(Mid(fvListaAutoLiquidable(i).sCtaCod, 9, 1) = "1", "SOLES", "DOLARES")
            lblmoneda.BackColor = IIf(Mid(fvListaAutoLiquidable(i).sCtaCod, 9, 1) = "1", &H80000005, &HC0FFC0)
            lblFormaRetiro.Caption = fvListaAutoLiquidable(i).sFormaRetiro
            lblMontoCuenta.Caption = Format(fvListaAutoLiquidable(i).nMontoCuenta, "#,##0.00")
            lblMontoCuentaDisp.Caption = Format(fvListaAutoLiquidable(i).nMontoCuentaDisp, "#,##0.00")
            lblMontoDisponible.Caption = Format(fvListaAutoLiquidable(i).nMontoDisponible, "#,##0.00")
            fnIndex = i
            CmdAceptar.Default = True
            Exit For
        End If
    Next
    
    Screen.MousePointer = 0
    Exit Sub
ErrCargaCuenta:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdCancelar_Click()
    fbOk = False
    Unload Me
End Sub
Private Sub LimpiarControles()
    cmbCuenta.ListIndex = -1
    lblmoneda.Caption = ""
    lblMontoCuenta.Caption = "0.00"
    lblMontoCuentaDisp.Caption = "0.00"
    lblFormaRetiro.Caption = ""
    lblMontoDisponible.Caption = "0.00"
End Sub
