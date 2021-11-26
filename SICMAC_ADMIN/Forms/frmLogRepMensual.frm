VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogRepMensual 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte Mensual de Movientos"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   Icon            =   "frmLogRepMensual.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   330
      Left            =   5220
      TabIndex        =   2
      Top             =   1575
      Width           =   1050
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   330
      Left            =   4110
      TabIndex        =   1
      Top             =   1575
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   6210
      Begin VB.ComboBox cboTpoAlm 
         Height          =   315
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   600
         Width           =   2865
      End
      Begin Sicmact.TxtBuscar txtAlmacen 
         Height          =   345
         Left            =   1110
         TabIndex        =   7
         Top             =   240
         Width           =   780
         _extentx        =   1376
         _extenty        =   609
         appearance      =   0
         appearance      =   0
         font            =   "frmLogRepMensual.frx":08CA
         appearance      =   0
         stitulo         =   ""
      End
      Begin MSMask.MaskEdBox mskAno 
         Height          =   285
         Left            =   1125
         TabIndex        =   4
         Top             =   630
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmbMes 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   975
         Width           =   2460
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Almacén Tipo :"
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
         Height          =   180
         Index           =   0
         Left            =   1980
         TabIndex        =   11
         Top             =   660
         Width           =   1500
      End
      Begin VB.Label lblAlmacen 
         Caption         =   "Almacen"
         Height          =   225
         Left            =   315
         TabIndex        =   9
         Top             =   300
         Width           =   780
      End
      Begin VB.Label lblAlmacenG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1950
         TabIndex        =   8
         Top             =   270
         Width           =   4200
      End
      Begin VB.Label lblMes 
         Caption         =   "Mes :"
         Height          =   195
         Left            =   315
         TabIndex        =   6
         Top             =   1035
         Width           =   750
      End
      Begin VB.Label lblFecha 
         Caption         =   "Año :"
         Height          =   195
         Left            =   315
         TabIndex        =   5
         Top             =   690
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmLogRepMensual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
Dim lsPalabras As String
'*******************************

Private Sub cmdProcesar_Click()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim lsCadenaSal As String
    Dim lsCadenaIng As String
    Dim ldFecMesAnt As Date
    Dim ldFecMesIni As Date
    Dim ldFecMesFin As Date
    Dim lsCadena As String

    
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    
    If Me.txtAlmacen.Text = "" Then
        MsgBox "Debe Elegir un Almacen.", vbInformation, "Aviso"
        txtAlmacen.SetFocus
        Exit Sub
    ElseIf Me.cmbMes.Text = "" Then
        MsgBox "Debe Elegir un mes.", vbInformation, "Aviso"
        cmbMes.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(Me.mskAno.Text) Then
        MsgBox "Debe ingresar un año.", vbInformation, "Aviso"
        mskAno.SetFocus
        Exit Sub
    End If
    
    ldFecMesIni = CDate("01/" & Format(Trim(Right(Me.cmbMes.Text, 3)), "00") & "/" & Me.mskAno.Text)
    ldFecMesAnt = DateAdd("d", -1, ldFecMesIni)
    ldFecMesFin = DateAdd("d", -1, DateAdd("m", 1, ldFecMesIni))
    
    lsCadenaSal = ""
    lsCadenaIng = ""
    
    Set rs = oALmacen.GetSalidasAlmacen(ldFecMesIni, ldFecMesFin, txtAlmacen.Text)
    
    While Not rs.EOF
        If lsCadenaSal = "" Then
            lsCadenaSal = Trim(Str(rs!nMovNro))
        Else
            lsCadenaSal = lsCadenaSal & "','" & Trim(Str(rs!nMovNro))
        End If
        rs.MoveNext
    Wend
    
    Set rs = oALmacen.GetIngresosAlmacen(ldFecMesIni, ldFecMesFin, "", txtAlmacen.Text)
        
    While Not rs.EOF
        If lsCadenaIng = "" Then
            lsCadenaIng = Trim(Str(rs!nMovNro))
        Else
            lsCadenaIng = lsCadenaIng & "','" & Trim(Str(rs!nMovNro))
        End If
        rs.MoveNext
    Wend
        
    lsCadena = oALmacen.GetRepLogMensual(lsCadenaIng, lsCadenaSal, ldFecMesAnt, ldFecMesFin, "SALDOS MENSULES LOGISTICA MES " & Trim(Left(Me.cmbMes.Text, 50)), gsNomAge, gsEmpresa, gdFecSis, Val(Right(Me.cboTpoAlm.Text, 5)), txtAlmacen.Text)
    
    oPrevio.Show lsCadena, Me.Caption, True, , gImpresora
        
        'ARLO 20160126 ***
        If (gsOpeCod = 591506) Then
        lsPalabras = "Mensual de Movimientos"
        ElseIf (gsOpeCod = 591502) Then
        lsPalabras = "Notas de Ingreso"
        ElseIf (gsOpeCod = 591503) Then
        lsPalabras = "Guias de Salida"
        End If
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Imprimio el Reporte " & lsPalabras & " del " & ldFecMesFin & " al " & ldFecMesFin
        Set objPista = Nothing
        '**************
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set rs = oCon.GetConstante(5010, False)
    CargaCombo rs, cboTpoAlm
    cboTpoAlm.ListIndex = 0
    
    Dim oDoc As DOperaciones
    Set oDoc = New DOperaciones
    Dim oGen As DGeneral
    Set oGen = New DGeneral
    
    Me.txtAlmacen.rs = oDoc.GetAlmacenes
    Set rs = oGen.GetConstante(1010)
    
    Me.cmbMes.Clear
    
    While Not rs.EOF
        cmbMes.AddItem rs.Fields(0) & Space(50) & rs.Fields(1)
        rs.MoveNext
    Wend
    
    Me.mskAno.Text = Format(gdFecSis, "yyyy")
End Sub

Private Sub txtAlmacen_EmiteDatos()
   Me.lblAlmacenG.Caption = Me.txtAlmacen.psDescripcion
End Sub
