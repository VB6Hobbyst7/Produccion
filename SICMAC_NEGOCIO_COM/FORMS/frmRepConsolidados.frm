VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmRepConsolidados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REPORTES CONSOLIDADOS"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   Icon            =   "frmRepConsolidados.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6825
      Left            =   165
      TabIndex        =   30
      Top             =   825
      Width           =   6705
      Begin MSComctlLib.ImageList imglstFiguras 
         Left            =   4215
         Top             =   4470
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRepConsolidados.frx":030A
               Key             =   "Padre"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRepConsolidados.frx":065C
               Key             =   "Bebe"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRepConsolidados.frx":09AE
               Key             =   "Hijo"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRepConsolidados.frx":0D00
               Key             =   "Hijito"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TreeRep 
         Height          =   6525
         Left            =   90
         TabIndex        =   31
         Top             =   180
         Width           =   6390
         _ExtentX        =   11271
         _ExtentY        =   11509
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imglstFiguras"
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6915
      TabIndex        =   25
      Top             =   7245
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8010
      TabIndex        =   24
      Top             =   7245
      Width           =   1035
   End
   Begin VB.Frame fraFecha 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Fechas"
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
      Height          =   975
      Left            =   7140
      TabIndex        =   19
      Top             =   885
      Width           =   1725
      Begin MSMask.MaskEdBox txtFechaF 
         Height          =   300
         Left            =   480
         TabIndex        =   20
         Top             =   570
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   300
         Left            =   480
         TabIndex        =   21
         Top             =   225
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblDel 
         AutoSize        =   -1  'True
         Caption         =   "Del:"
         Height          =   195
         Left            =   75
         TabIndex        =   23
         Top             =   285
         Width           =   285
      End
      Begin VB.Label lblAl 
         AutoSize        =   -1  'True
         Caption         =   "Al:"
         Height          =   195
         Left            =   225
         TabIndex        =   22
         Top             =   630
         Width           =   180
      End
   End
   Begin VB.Frame fraMonto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Montos"
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
      Height          =   975
      Left            =   7125
      TabIndex        =   14
      Top             =   1905
      Visible         =   0   'False
      Width           =   1725
      Begin SICMACT.EditMoney txtMonto 
         Height          =   285
         Left            =   480
         TabIndex        =   15
         Top             =   240
         Width           =   1185
         _extentx        =   2090
         _extenty        =   503
         font            =   "frmRepConsolidados.frx":1052
         appearance      =   0
         text            =   "0"
         enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtMontoF 
         Height          =   285
         Left            =   480
         TabIndex        =   16
         Top             =   585
         Width           =   1185
         _extentx        =   2090
         _extenty        =   503
         font            =   "frmRepConsolidados.frx":107E
         appearance      =   0
         text            =   "0"
         enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Al:"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   630
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Del:"
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   278
         Width           =   285
      End
   End
   Begin VB.Frame fraTipoCambio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Tipo de Cambio"
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
      Height          =   645
      Left            =   7125
      TabIndex        =   12
      Top             =   2940
      Visible         =   0   'False
      Width           =   1725
      Begin SICMACT.EditMoney EditMoney3 
         Height          =   285
         Left            =   480
         TabIndex        =   13
         Top             =   240
         Width           =   1185
         _extentx        =   2090
         _extenty        =   503
         font            =   "frmRepConsolidados.frx":10AA
         appearance      =   0
         text            =   "0"
         enabled         =   -1  'True
      End
   End
   Begin VB.Frame fraAgencias 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Agencias"
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
      Height          =   660
      Left            =   165
      TabIndex        =   8
      Top             =   135
      Width           =   8715
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Todos"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   165
         TabIndex        =   10
         Top             =   300
         Width           =   930
      End
      Begin SICMACT.TxtBuscar TxtAgencia 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   270
         Width           =   1335
         _extentx        =   2355
         _extenty        =   503
         appearance      =   0
         font            =   "frmRepConsolidados.frx":10D6
         appearance      =   0
         stitulo         =   ""
      End
      Begin VB.Label lblAgencia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2505
         TabIndex        =   11
         Top             =   255
         Width           =   6045
      End
   End
   Begin VB.Frame fraUser 
      Height          =   675
      Left            =   7095
      TabIndex        =   5
      Top             =   3600
      Width           =   1770
      Begin VB.CheckBox Check1 
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   30
         TabIndex        =   6
         Top             =   255
         Width           =   690
      End
      Begin SICMACT.TxtBuscar TxtBuscarUser 
         Height          =   345
         Left            =   720
         TabIndex        =   7
         Top             =   195
         Width           =   945
         _extentx        =   1667
         _extenty        =   609
         appearance      =   1
         font            =   "frmRepConsolidados.frx":1102
         enabled         =   0   'False
         enabled         =   0   'False
         stitulo         =   ""
         enabledtext     =   0   'False
         forecolor       =   12582912
      End
   End
   Begin VB.Frame fraCheque 
      Caption         =   "Estado Actual Cheq"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1965
      Left            =   7035
      TabIndex        =   3
      Top             =   4290
      Visible         =   0   'False
      Width           =   1845
      Begin VB.ListBox lstcheques 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         ItemData        =   "frmRepConsolidados.frx":112E
         Left            =   180
         List            =   "frmRepConsolidados.frx":1144
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   315
         Width           =   1560
      End
   End
   Begin VB.Frame fraOrden 
      Caption         =   "Ordenar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   885
      Left            =   6990
      TabIndex        =   26
      Top             =   6300
      Width           =   1890
      Begin VB.OptionButton Option2 
         Caption         =   "Operación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   45
         TabIndex        =   29
         Top             =   510
         Width           =   1080
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Transacción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   28
         Top             =   270
         Value           =   -1  'True
         Width           =   1320
      End
      Begin VB.CheckBox chkTotal 
         Caption         =   "Total"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   27
         Top             =   495
         Width           =   765
      End
   End
   Begin VB.Frame fracmacs 
      Caption         =   "Incluir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   885
      Left            =   6990
      TabIndex        =   0
      Top             =   6285
      Visible         =   0   'False
      Width           =   1890
      Begin VB.CheckBox chkRecepcion 
         Caption         =   "Recepcion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   225
         TabIndex        =   2
         Top             =   555
         Width           =   1080
      End
      Begin VB.CheckBox chkLlamadas 
         Caption         =   "Llamadas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   1
         Top             =   285
         Width           =   1170
      End
   End
   Begin SICMACT.Usuario Usuario 
      Left            =   300
      Top             =   7170
      _extentx        =   820
      _extenty        =   820
   End
End
Attribute VB_Name = "frmRepConsolidados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim Flag As Boolean
Dim flag1 As Boolean
Dim Char12 As Boolean
Dim lbPrtCom As Boolean
Dim NumA As String
Dim SalA As String
Dim NumC As String
Dim NumP As String
Dim SalP As String
Dim SalC As String
Dim TB As Currency
Dim TBD As Currency
Dim ca As Currency
Dim CAE As Currency
Dim RegCmacS As Integer
Dim vBuffer As String
Dim lsCadena As String

Dim oGen As DGeneral




Private Sub Limpia()
    txtFecha.Text = "__/__/____"
    txtFechaF.Text = "__/__/____"
    txtMonto.Text = ""
    txtMontoF.Text = ""
    chkTodos.value = 0
    TxtAgencia.Text = ""
    lblAgencia.Caption = ""
    TxtBuscarUser.Text = ""
    chkLlamadas.value = 0
    chkRecepcion.value = 0
End Sub
Private Sub LlenaArbol()
Dim sqlv As String
'Dim PObjConec as DConecta
Dim rsUsu As New ADODB.Recordset
Dim sOperacion As String, sOpeCod As String
Dim nodOpe As Node
Dim sOpePadre As String
Dim sOpeHijo As String
Dim sOpeHijito As String

Dim PObjConec As New DConecta

    'Set PObjConec = New DConecta

    sqlv = " Select cOpeCod Codigo, UPPER(cCapRepDescripcion) + ' [' + Rtrim(ltrim(Str(nCapRepDiaCalculo))) + '-' + rtrim(ltrim(str(nCapRepRangoMonto))) + '-' + rtrim(ltrim(Str(bCapRepTipoCambio))) + ']' Descripcion, nOpeNiv Nivel From OpeTpo OP" _
         & " Inner Join CaptaReportes CAPR On CAPR.cCapRepCod = OP.cOpeCod" _
         & " Where copecod like '28%' order by cOpeCod"
    PObjConec.AbreConexion
     
    Set rsUsu = PObjConec.CargaRecordSet(sqlv)
    
    Do While Not rsUsu.EOF
        sOpeCod = rsUsu("Codigo")
        sOperacion = sOpeCod & " " & UCase(rsUsu("Descripcion"))
        Select Case rsUsu("Nivel")
            Case "1"
                sOpePadre = "P" & sOpeCod
                Set nodOpe = TreeRep.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
                nodOpe.Tag = sOpeCod
            Case "2"
                sOpeHijo = "H" & sOpeCod
                Set nodOpe = TreeRep.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
                nodOpe.Tag = sOpeCod
            Case "3"
                sOpeHijito = "J" & sOpeCod
                Set nodOpe = TreeRep.Nodes.Add(sOpeHijo, tvwChild, sOpeHijito, sOperacion, "Hijito")
                nodOpe.Tag = sOpeCod
            Case "4"
                Set nodOpe = TreeRep.Nodes.Add(sOpeHijito, tvwChild, "B" & sOpeCod, sOperacion, "Bebe")
                nodOpe.Tag = sOpeCod
        End Select
        rsUsu.MoveNext
    Loop
    rsUsu.Close
    Set rsUsu = Nothing
    
    PObjConec.CierraConexion
Set PObjConec = Nothing
End Sub


Private Sub HabilitaControles(ByVal pfraAgencia As Boolean, ByVal pUser As Boolean, ByVal pfraFecha As Boolean, pFecha As Boolean, pfechaf As Boolean, _
                              ByVal pfraMonto As Boolean, ByVal ptipocambio As Boolean, ByVal pMonto As Boolean, _
                              ByVal pMontof As Boolean, Optional pFraCheque As Boolean = False, Optional pFraOrden As Boolean = False, Optional pFraCmacs As Boolean = False)
    
    fraCheque.Visible = pFraCheque
    frafechacheques.Visible = pFraCheque
    
    fraOrden.Visible = pFraOrden
    
    fraAgencias.Visible = pfraAgencia
    
    fraUser.Visible = pUser
    fraFecha.Visible = pFecha
    
    fraMonto.Visible = pfraMonto
    txtMonto.Visible = pMonto
    txtMontoF.Visible = pMontof
    Label3.Visible = pMontof
    
    fraFecha.Visible = pfraFecha
    txtFecha.Visible = pFecha
    txtFechaF.Visible = pfechaf
    lblAl.Visible = pfechaf
    
    fraTipoCambio.Visible = ptipocambio
    
    fracmacs.Visible = pFraCmacs
    
End Sub
 Dim rs As ADODB.Recordset
    Dim lsCodCab As String
    Dim lsCodCab1 As String
    Dim oCons As DConstante
    Set oCons = New DConstante
    Set oGen = New DGeneral
    LlenaArbol
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    
    txtFecha = Format(gdFecSis, gsFormatoFechaView)
    txtFechaF = Format(gdFecSis, gsFormatoFechaView)
    
    Me.TxtAgencia.rs = oCons.GetAgencias(, , True)
    Usuario.Inicio gsCodUser

Private Sub cmdImprimir_Click()
  Dim Sql As String, rstemp As ADODB.Recordset
  Sql = "SELECT  "
  sSql = sSql & "  "
  
  While Not rstemp.EOF

  End If
End Sub
