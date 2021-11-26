VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmImpreRRHH 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   Icon            =   "frmImpreRRHH.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOpt 
      Caption         =   "Reportes"
      Height          =   3840
      Left            =   75
      TabIndex        =   7
      Top             =   645
      Width           =   4305
      Begin MSComctlLib.ListView lvwImp 
         Height          =   3510
         Left            =   75
         TabIndex        =   8
         Top             =   225
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   6191
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImaLis"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame fraFecha 
      Caption         =   "Fechas"
      Height          =   600
      Left            =   60
      TabIndex        =   2
      Top             =   15
      Width           =   4320
      Begin MSMask.MaskEdBox mskIni 
         Height          =   300
         Left            =   690
         TabIndex        =   3
         Top             =   225
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFin 
         Height          =   300
         Left            =   2970
         TabIndex        =   4
         Top             =   225
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblIni 
         Caption         =   "Inicio:"
         Height          =   255
         Left            =   90
         TabIndex        =   6
         Top             =   255
         Width           =   495
      End
      Begin VB.Label lblFin 
         Caption         =   "Fin:"
         Height          =   255
         Left            =   2610
         TabIndex        =   5
         Top             =   255
         Width           =   375
      End
   End
   Begin MSComctlLib.ImageList ImaLis 
      Left            =   4755
      Top             =   2265
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpreRRHH.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpreRRHH.frx":079C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpreRRHH.frx":0B62
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3435
      TabIndex        =   1
      Top             =   4530
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2355
      TabIndex        =   0
      Top             =   4530
      Width           =   975
   End
End
Attribute VB_Name = "frmImpreRRHH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsContenido As String
Dim ldFecIni As Date
Dim ldFecFin As Date
Dim lbSalir As Boolean
Dim lb(100) As Boolean
Dim lnCantidad As Integer

Private Sub cmdAceptar_Click()
    Dim i As Integer
    
    If mskIni.Enabled Then
        If Not IsDate(mskIni) Then
            MsgBox "Fecha No Valida.", vbInformation, "Aviso"
            mskIni.SetFocus
            Exit Sub
        ElseIf Not IsDate(mskFin) Then
            MsgBox "Fecha No Valida.", vbInformation, "Aviso"
            mskFin.SetFocus
            Exit Sub
        Else
            ldFecIni = CDate(mskIni)
            ldFecFin = CDate(mskFin)
        End If
    End If
    
    For i = 1 To lnCantidad
        lb(i) = lvwImp.ListItems(i).Checked
    Next i
    
    lbSalir = False
    
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    lsContenido = ""
    lbSalir = True
    Unload Me
End Sub

Public Function Ini(psCadena As String, psCaption As String, PB() As Boolean, pdFecIni As Date, pdFecFin As Date, Optional pbFechas As Boolean = True) As Boolean
    Dim i As Integer
    lsContenido = ""
    
    'EJVG 20110829********
    For i = 1 To UBound(lb)
        lb(i) = False
    Next
    'END******************
    
    IniLVW psCadena
    frmImpreRRHH.Caption = psCaption
    
    mskIni.Enabled = pbFechas
    mskFin.Enabled = pbFechas
    mskIni.Visible = pbFechas
    mskFin.Visible = pbFechas
    lblFin.Visible = pbFechas
    lblIni.Visible = pbFechas
    
    mskIni = Format(pdFecIni, gsFormatoFechaView)
    mskFin = Format(pdFecFin, gsFormatoFechaView)
    
    lnCantidad = UBound(PB, 1)
    
    frmImpreRRHH.Show 1
    
    If Not lbSalir Then
        For i = 1 To lnCantidad
            PB(i) = lb(i)
        Next i
    Else
        For i = 1 To lnCantidad
            PB(i) = False
        Next i
    End If
       
     If Not mskIni.Enabled Then
        pdFecIni = ldFecIni
        pdFecFin = ldFecFin
    End If
    
    Ini = lbSalir
End Function

Private Sub IniLVW(psCadena As String)
    Dim lsCadena  As String
    Dim lsItem1 As String
    Dim lnItem1 As Integer
    Dim llAux As ListItem
    Dim lnContador As Integer
    
    lvwImp.ColumnHeaders.Clear
    lvwImp.ListItems.Clear
    
    lvwImp.HideColumnHeaders = False
    lvwImp.ColumnHeaders.Add , , "Opción de Impresión", 4000
    lvwImp.ColumnHeaders.Add , , "Cod", 1
    
    lvwImp.View = lvwReport
    
    lsCadena = psCadena
    lnContador = 0
    While Not lsCadena = ""
        lnItem1 = InStr(1, lsCadena, ";", vbTextCompare)
        lsItem1 = Left(lsCadena, lnItem1 - 1)
        lsCadena = Mid(lsCadena, lnItem1 + 1, Len(lsCadena))
        
        lnContador = lnContador + 1
        
        Set llAux = lvwImp.ListItems.Add(, , lsItem1, , 2)
        llAux.SubItems(1) = lnContador
    Wend

End Sub

Private Sub lvwImp_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        Item.SmallIcon = 1
    Else
        Item.SmallIcon = 2
    End If
End Sub

Private Sub lvwImp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdAceptar.SetFocus
End Sub

Private Sub mskFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lvwImp.SetFocus
End Sub

Private Sub mskIni_GotFocus()
    mskIni.SelStart = 0
    mskIni.SelLength = 100
End Sub

Private Sub mskFin_GotFocus()
    mskFin.SelStart = 0
    mskFin.SelLength = 100
End Sub

Private Sub mskIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then mskFin.SetFocus
End Sub
