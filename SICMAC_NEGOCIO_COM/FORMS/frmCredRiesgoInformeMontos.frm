VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredRiesgoInformeMontos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Riesgo"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   Icon            =   "frmCredRiesgoInformeMontos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8160
      TabIndex        =   14
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   7080
      TabIndex        =   13
      Top             =   6240
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Informacion del Cliente"
      TabPicture(0)   =   "frmCredRiesgoInformeMontos.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblAgencia"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ActXCodCta"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtAgencia"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.Frame Frame3 
         Caption         =   "Datos de Credito"
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
         TabIndex        =   17
         Top             =   2520
         Width           =   8895
         Begin VB.TextBox txtCuotProp 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1440
            TabIndex        =   30
            Top             =   780
            Width           =   1815
         End
         Begin VB.TextBox txtnTEM 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6600
            TabIndex        =   26
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtnTenAnt 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5160
            TabIndex        =   25
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtnTEA 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6600
            TabIndex        =   24
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtnCuota 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5160
            TabIndex        =   23
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtMontProp 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2040
            TabIndex        =   22
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox txtTpMon 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1440
            TabIndex        =   21
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "Cuota Propuesta"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label10 
            Caption         =   "TEM"
            Height          =   255
            Left            =   6000
            TabIndex        =   28
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label9 
            Caption         =   "T.E.A"
            Height          =   255
            Left            =   6000
            TabIndex        =   27
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label8 
            Caption         =   "TEM Ant"
            Height          =   255
            Left            =   4440
            TabIndex        =   20
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "N° Cuotas"
            Height          =   255
            Left            =   4320
            TabIndex        =   19
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Monto Propuesto"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.TextBox txtAgencia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6240
         TabIndex        =   16
         Top             =   600
         Width           =   2775
      End
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   240
         TabIndex        =   2
         Top             =   955
         Width           =   8895
         Begin VB.TextBox txtFechaExp 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1440
            TabIndex        =   11
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox txtNombCliente 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1440
            TabIndex        =   5
            Top             =   240
            Width           =   7335
         End
         Begin VB.TextBox txtTpCredito 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5280
            TabIndex        =   4
            Top             =   600
            Width           =   3495
         End
         Begin VB.TextBox txtTpProducto 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1440
            TabIndex        =   3
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha de Exp."
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1005
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Titular:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo de Producto:"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo de Credito:"
            Height          =   255
            Left            =   4080
            TabIndex        =   6
            Top             =   645
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Descripcion"
         Height          =   2175
         Left            =   240
         TabIndex        =   1
         Top             =   3840
         Width           =   8895
         Begin VB.TextBox txtDescripcion 
            Height          =   1815
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   8655
         End
      End
      Begin SICMACT.ActXCodCta ActXCodCta 
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   873
         Texto           =   "Credito:"
      End
      Begin VB.Label lblAgencia 
         Caption         =   "Agencia:"
         Height          =   255
         Left            =   5400
         TabIndex        =   15
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCredRiesgoInformeMontos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cAgeCod As String
Dim cPersCod As String
Dim cTpProducto As String
Dim cTpCrdito As String
Dim nCuotaProp As Double

Public Sub Inicio(ByVal psCtaCod As String)

Dim oCreditos As COMDCredito.DCOMCredito
Dim rsCreditos As ADODB.Recordset

Set oCreditos = New COMDCredito.DCOMCredito
Set rsCreditos = oCreditos.MostrarCreditoInformeRiesgoPend(psCtaCod)
Call Cuadro2Creditos(psCtaCod)
                  
ActXCodCta.NroCuenta = psCtaCod
    cAgeCod = rsCreditos!cAgeCod
txtAgencia.Text = rsCreditos!Agencia
    cPersCod = rsCreditos!cPersCod
txtNombCliente.Text = rsCreditos!cPersNombre
    cTpProducto = rsCreditos!cTpoProdCod
txtTpProducto.Text = rsCreditos!TpoProd
    cTpCrdito = rsCreditos!cTpoCredCod
txtTpCredito.Text = rsCreditos!TpoCred
txtFechaExp.Text = rsCreditos!FechaExp

Show 1

RSClose rsCreditos

End Sub

Private Sub CmdGrabar_Click()

Dim oCredito As COMDCredito.DCOMCredito
Set oCredito = New COMDCredito.DCOMCredito
Dim GrabarDatos As Boolean

If validar Then

    If MsgBox("Los Datos seran Grabados, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub

        GrabarDatos = oCredito.GuardarRiesgoInformeMontos(ActXCodCta.NroCuenta, cAgeCod, cPersCod, cTpProducto, cTpCrdito, _
                                                          txtFechaExp.Text, TxtDescripcion.Text, _
                                                          txtMontProp, txtCuotProp, txtnCuota, txtnTEA, txtnTenAnt, txtnTEM)
        
        Call oCredito.ActualizaRiesgoInformeMontos(ActXCodCta.NroCuenta, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), 2)
        frmCredRiesgos.LLenarGrilla
                
            If GrabarDatos Then
                   
                MsgBox "Informe de Riesgo registrado satisfactoriamente.", vbInformation, "Aviso"
                
            Else
        
                MsgBox "Hubo error al grabar la informacion", vbError, "Error"
            
            End If
    Unload Me
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Function validar() As Boolean
    
    validar = True

    If TxtDescripcion.Text = "" Then
        MsgBox "Ingrese Descripcion del Credito a Autorizar", vbInformation, "Aviso"
        SSTab1.Tab = 0
        TxtDescripcion.SetFocus
        validar = False
        Exit Function
    End If

End Function

Private Sub Cuadro2Creditos(ByVal psCtaCod As String)
    Dim clsTC As COMDConstSistema.NCOMTipoCambio
    Dim nTC As Double
    Set clsTC = New COMDConstSistema.NCOMTipoCambio
    
    
    Dim objDPersona As COMDPersona.DCOMPersona
    Dim oRS As ADODB.Recordset
    Set oRS = New ADODB.Recordset
    Set objDPersona = New COMDPersona.DCOMPersona
    nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
    Set clsTC = Nothing
    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro2Creditos(psCtaCod, nTC)
    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
         txtMontProp = Format(oRS!nMontoPropuesto, "###,###,###,##0.00") 'LUCV20160919
         txtTpMon.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "$")
         txtnCuota.Text = oRS!nNrocuotas
         txtnTEA.Text = oRS!nTEA
         txtnTenAnt.Text = IIf(IsNull(oRS!nTEMA), 0#, oRS!nTEMA)
         txtnTEM.Text = IIf(IsNull(oRS!nTem), 0, oRS!nTem)
         txtCuotProp.Text = Format(IIf(IsNull(oRS!nCuoPropuesta), 0, oRS!nCuoPropuesta), "###,###,###,##0.00")
         
        oRS.MoveNext
    Loop
    End If
    
RSClose oRS

End Sub

