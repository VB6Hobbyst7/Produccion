VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPreRubros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Presupuesto: Mantenimiento de Rubros"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   Icon            =   "frmPreRubros.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   7500
      TabIndex        =   10
      Top             =   6480
      Width           =   1110
   End
   Begin VB.ComboBox cboFecha 
      Height          =   315
      Left            =   540
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   15
      Width           =   1000
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   390
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6480
      Width           =   1110
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Height          =   390
      Left            =   195
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   1110
   End
   Begin VB.ComboBox cboPresu 
      Height          =   315
      Left            =   2505
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   15
      Width           =   4125
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   390
      Left            =   1350
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6480
      Width           =   1110
   End
   Begin VB.CommandButton cmdExa 
      Caption         =   "..."
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CommandButton cmdImprimir 
      Height          =   390
      Left            =   3660
      Picture         =   "frmPreRubros.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6480
      Width           =   1110
   End
   Begin VB.ComboBox cboTpo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   7260
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   30
      Width           =   1635
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgIngreso 
      Height          =   4200
      Left            =   30
      TabIndex        =   4
      Top             =   480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7408
      _Version        =   393216
      Cols            =   4
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483638
      AllowBigSelection=   0   'False
      TextStyleFixed  =   3
      FocusRect       =   0
      FillStyle       =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgDetalle 
      Height          =   1530
      Left            =   30
      TabIndex        =   5
      Top             =   4845
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   2699
      _Version        =   393216
      Cols            =   4
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483638
      AllowBigSelection=   0   'False
      TextStyleFixed  =   3
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin RichTextLib.RichTextBox rtfImp 
      Height          =   315
      Left            =   6675
      TabIndex        =   11
      Top             =   6540
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   556
      _Version        =   393217
      TextRTF         =   $"frmPreRubros.frx":1010
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Año :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   90
      TabIndex        =   14
      Top             =   75
      Width           =   510
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Nombre :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   1710
      TabIndex        =   13
      Top             =   75
      Width           =   825
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Tipo :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   6750
      TabIndex        =   12
      Top             =   75
      Width           =   510
   End
End
Attribute VB_Name = "frmPreRubros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vRTFImp As String
Dim pbIni As Boolean
Dim pRow As Integer

Private Sub cboFecha_Click()
If pbIni Then
    pRow = 1
    Call CargaPresu
    Call CargaDeta
End If
End Sub

Private Sub cboPresu_Click()
    Dim tmpSql As String
    Dim sPresu As String
    Dim oPP As DPresupuesto
    Set oPP = New DPresupuesto
    sPresu = Trim(Right(Trim(cboPresu.Text), 4))
    If Len(sPresu) > 0 Then
        tmpSql = oPP.GetPresupuestoTpo(nVal(sPresu))
        UbicaCombo cboTpo, tmpSql
    End If
    pRow = 1
    Call CargaPresu
    Call CargaDeta
End Sub

Private Sub cmdAgregar_Click()
Dim oPP As DPresupuesto
Set oPP = New DPresupuesto
pRow = fgIngreso.Row
If Not oPP.IsMonto(Me.cboFecha.Text, Right(cboPresu.Text, 1), fgIngreso.TextMatrix(fgIngreso.Row, 1)) Then
   If frmPlaRubIng.Inicio(Left(Trim(cboFecha.Text), 4), Right(cboPresu, 4), fgIngreso.TextMatrix(fgIngreso.Row, 1), "1") Then
       Call CargaPresu
       Call CargaDeta
   End If
Else
    MsgBox "Ya se ha ingresado un monto a este rubro", vbInformation, "Aviso"
End If
fgIngreso.SetFocus
End Sub
Private Sub cmdEliminar_Click()
Dim tmpSql As String
Dim tmpReg As New ADODB.Recordset
Dim nErr As Currency
Dim X As Integer
Dim oPre As New DPresupuesto
On Error GoTo cmdEliminarErr

Set tmpReg = oPre.GetPresupRubro(cboFecha.Text, Right(Trim(cboPresu.Text), 4), fgIngreso.TextMatrix(fgIngreso.Row, 1) & "__")
If (tmpReg.BOF Or tmpReg.EOF) Then
    If MsgBox(" ¿ Esta seguro de Eliminar este Rubro ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
       oPre.EliminaRubro Right(cboPresu.Text, 4), Left(cboFecha.Text, 4), fgIngreso.TextMatrix(fgIngreso.Row, 1)
       fgIngreso.RemoveItem fgIngreso.Row
       Call CargaDeta
    End If
Else
    MsgBox "Existen rubros en niveles inferiores", vbInformation, " Aviso "
End If
RSClose tmpReg
fgIngreso.SetFocus
Exit Sub
cmdEliminarErr:
   If Err.Number = -2147211505 Then
      MsgBox "No se puede eliminar Rubro porque posee relación con Cuentas y/o con Importes Mensuales", vbInformation, "¡Aviso!"
   Else
      MsgBox Err.Description, vbInformation, "¡Aviso!"
   End If
End Sub

Private Sub cmdExa_Click()
Dim vOpc As Boolean
vOpc = frmPlaRubIng.Inicio(Left(Trim(cboFecha.Text), 4), Right(cboPresu, 4), fgIngreso.TextMatrix(fgIngreso.Row, 1), "3")
End Sub

Private Sub CmdImprimir_Click()
    On Error GoTo ControlError
    Dim tmpSql As String
    Dim tmpReg As ADODB.Recordset
    Set tmpReg = New ADODB.Recordset
    Dim oPrevio As clsPrevio
    Set oPrevio = New clsPrevio
    Dim vPage As Integer, vLineas As Integer, vItem As Integer
    Dim m As Integer
    Dim vSpa As String
    Dim vCtas As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lsTem As String
    Dim lnI As Integer
    
    MousePointer = 11
    vRTFImp = ""
    vSpa = Space(4)
    vPage = 1
      
    Call Cabecera("1", vPage)
    oCon.AbreConexion
For m = 2 To fgIngreso.Rows - 1
    vCtas = ""
    vItem = vItem + 1
    '************************************************************************************
    'Lectura de las ctas cnts.
    'If Right(Trim(cboTpo.Text), 1) <> "3" Then
        tmpSql = " SELECT cPresuRubCod, cCtaContCod FROM PresuRubroCta A " & _
               "  " & _
               " Where nPresuAnio = " & Left(cboFecha, 4) & " AND nPresuCod = " & Right(cboPresu, 4) & " " & _
               " AND cPresuRubCod = '" & Trim(fgIngreso.TextMatrix(m, 1)) & "' Order By cCtaContCod "
        Set tmpReg = oCon.CargaRecordSet(tmpSql)
        If Not (tmpReg.BOF Or tmpReg.EOF) Then
            lnI = 0
            Do While Not tmpReg.EOF
                lnI = lnI + 1
                
                If lnI Mod 3 = 0 Then
                    vCtas = vCtas & Trim(tmpReg!cCtaContCod) & "* "
                Else
                    vCtas = vCtas & Trim(tmpReg!cCtaContCod) & ", "
                End If
                tmpReg.MoveNext
            Loop
            If Len(vCtas) >= 52 Then
                lsTem = vCtas
                vCtas = ""
                While InStr(1, lsTem, "*") <> 0
                    If vCtas = "" Then
                        vCtas = Left(lsTem, InStr(1, lsTem, "*") - 1)
                    Else
                        vCtas = vCtas & oImpresora.gPrnSaltoLinea & Space(70) & Left(lsTem, InStr(1, lsTem, "*") - 1)
                    End If
                    lsTem = Mid(lsTem, InStr(1, lsTem, "*") + 1)
                Wend
                
                If InStr(1, lsTem, ",") = 0 Then
                    vCtas = vCtas & oImpresora.gPrnSaltoLinea & Space(70) & Left(lsTem, Len(vCtas) - 2)
                Else
                    vCtas = vCtas & oImpresora.gPrnSaltoLinea & Space(70) & Left(lsTem, InStr(1, lsTem, ",") - 1)
                End If
            Else
                vCtas = Left(vCtas, Len(vCtas) - 2)
            End If
        End If
        tmpReg.Close
        Set tmpReg = Nothing
    'End If
    '************************************************************************************
    vRTFImp = vRTFImp & vSpa & ImpreFormat(vItem, 5, 0) & ImpreFormat(Mid(fgIngreso.TextMatrix(m, 1), 3), 20) & _
        ImpreFormat(fgIngreso.TextMatrix(m, 2), 36) & Space(2) & vCtas & oImpresora.gPrnSaltoLinea
    vLineas = vLineas + 1
    If vLineas >= 55 Then
        vPage = vPage + 1
        Call Cabecera("1", vPage)
        vLineas = 0
    End If
Next
MousePointer = 0
    oPrevio.Show vRTFImp, Caption, True
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdModificar_Click()
pRow = fgIngreso.Row
If fgIngreso.TextMatrix(pRow, 1) = "" Then
   Exit Sub
End If
If frmPlaRubIng.Inicio(Left(Trim(cboFecha.Text), 4), Right(cboPresu, 4), fgIngreso.TextMatrix(pRow, 1), "2") Then
    Call CargaPresu
    fgIngreso.Row = pRow
    fgIngreso.TopRow = pRow
    Call CargaDeta
End If
fgIngreso.SetFocus
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub fgIngreso_Click()
    If Trim(fgIngreso.TextMatrix(fgIngreso.Row, 1)) <> "" Then
        Call CargaDeta
        If fgIngreso.Col = 3 _
            And IsUltNivel(Left(Trim(cboFecha.Text), 4), Right(cboPresu.Text, 4), fgIngreso.TextMatrix(fgIngreso.Row, 1)) And Trim(fgIngreso.TextMatrix(fgIngreso.Row, 1)) <> "00" Then
            cmdExa.Left = fgIngreso.Left + fgIngreso.CellWidth + fgIngreso.CellLeft - cmdExa.Width
            cmdExa.Top = fgIngreso.Top + fgIngreso.CellTop - 15
            cmdExa.Visible = True
            cmdExa.SetFocus
        Else
            cmdExa.Visible = False
        End If
    End If
End Sub

Private Sub fgIngreso_Scroll()
cmdExa.Visible = False
End Sub

Private Sub Form_Load()
    Dim tmpSql As String
    Dim clsDGnral As DLogGeneral
    Set clsDGnral = New DLogGeneral
    Dim oPP As DPresupuesto
    Set oPP = New DPresupuesto
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    CentraForm Me
    pRow = 1
    pbIni = False
    'Carga los Años
    Set rs = clsDGnral.CargaPeriodo
    Call CargaCombo(rs, cboFecha)
    'Carga el nombre de los Presupuestos
    Call CargaCombo(oPP.GetPresupuesto(True), cboPresu, , 1, 0)
    pbIni = True
    Set rs = oCon.GetConstante(gPPPresupuestoTpo)
    CargaCombo rs, cboTpo
End Sub

Private Sub Limpiar()
cmdExa.Visible = False
Call MSHFlex(fgIngreso, 4, "Item-Código-Descripción-Ctas", "450-1800-5800-500", "R-L-L-L")
Call MSHFlex(fgDetalle, 4, "Item-Código-Descripción-Flag", "450-1800-5800-0", "R-L-L-L")
End Sub

Private Sub CargaPresu()
    Dim tmpReg As ADODB.Recordset
    Set tmpReg = New ADODB.Recordset
    Dim tmpSql As String
    Dim X As Integer, n As Integer
    Dim oPP As DPresupuesto
    Set oPP = New DPresupuesto
    
    Limpiar
    AdicionaRow fgIngreso, 1
    If Len(cboPresu.Text) > 4 Then
        fgIngreso.TextMatrix(1, 1) = "00"
        fgIngreso.TextMatrix(1, 2) = Trim(Left(cboPresu.Text, Len(cboPresu.Text) - 4))
    End If
    fgIngreso.Row = 1
    For n = 1 To fgIngreso.Cols - 1
        fgIngreso.Col = n
        fgIngreso.CellBackColor = &HC0FFFF
        fgIngreso.CellForeColor = vbBlue
        'fgIngreso.CellBackColor = &HC0C000         ' Celeste
    Next
    fgIngreso.Redraw = False
    X = 1
    Set tmpReg = oPP.GetRubrosPresupuesto(Me.cboFecha.Text, nVal(Right(Me.cboPresu.Text, 3)))
    If (tmpReg.BOF Or tmpReg.EOF) Then
    Else
        With tmpReg
            Do While Not .EOF
                X = X + 1
                AdicionaRow fgIngreso, X
                fgIngreso.TextMatrix(X, 0) = X
                fgIngreso.TextMatrix(X, 1) = !cPresuRubCod
                fgIngreso.TextMatrix(X, 2) = !cPresuRubDescripcion
                fgIngreso.Row = X
                .MoveNext
            Loop
        End With
    End If
    tmpReg.Close
    Set tmpReg = Nothing
    fgIngreso.Redraw = True
    fgIngreso.TopRow = pRow
    fgIngreso.Row = pRow
End Sub

Private Sub CargaDeta()
    Dim tmpSql As String
    Dim tmpReg As ADODB.Recordset
    Set tmpReg = New ADODB.Recordset
    Dim oPP As DPresupuesto
    Set oPP = New DPresupuesto
    Dim X As Integer
    fgDetalle.Redraw = False
    Call MSHFlex(fgDetalle, 4, "Item-Código-Descripción-Flag", "450-1800-5500-0", "R-L-L-L-L")
    
    If cboPresu.Text = "" Then Exit Sub
    
    Set tmpReg = oPP.GetRubrosPresupuestoDet(Me.cboFecha.Text, Trim(Right(Me.cboPresu.Text, 3)), Me.fgIngreso.TextMatrix(Me.fgIngreso.Row, 1))
    'UbicaCombo cboTpo, tmpSql
    If Not (tmpReg.BOF Or tmpReg.EOF) Then
        With tmpReg
            Do While Not .EOF
                X = X + 1
                AdicionaRow fgDetalle, X
                fgDetalle.TextMatrix(X, 0) = X
                fgDetalle.TextMatrix(X, 1) = !cPresuRubCod
                fgDetalle.TextMatrix(X, 2) = !cPresuRubDescripcion
                .MoveNext
            Loop
        End With
    End If
    tmpReg.Close
    Set tmpReg = Nothing
    fgDetalle.Redraw = True
End Sub



Private Sub Cabecera(ByVal cTipo As String, ByVal nPage As Integer)
If nPage > 1 Then vRTFImp = vRTFImp & oImpresora.gPrnSaltoPagina
vRTFImp = vRTFImp & "  CMAC - TRUJILLO" & Space(90) & Format(gdFecSis & " " & Time, gsFormatoFechaHoraView) & oImpresora.gPrnSaltoLinea
vRTFImp = vRTFImp & ImpreFormat(UCase(gsNomAge), 25) & Space(85) & " Página :" & ImpreFormat(nPage, 5, 0) & oImpresora.gPrnSaltoLinea
If cTipo = "1" Then
    vRTFImp = vRTFImp & Space(25) & "LISTADO DE RUBROS DEL AÑO " & Left(Trim(cboFecha), 4) & " DE : " & UCase(Left(cboPresu, 45)) & oImpresora.gPrnSaltoLinea
End If
vRTFImp = vRTFImp & Space(2) & String(124, "-") & oImpresora.gPrnSaltoLinea
'If Right(Trim(cboTpo.Text), 1) = "3" Then
'    vRTFImp = vRTFImp & Space(2) & "  ITEM          CODIGO                      DESCRIPCION                         " & oImpresora.gPrnSaltoLinea
'Else
    vRTFImp = vRTFImp & Space(2) & "  ITEM          CODIGO                      DESCRIPCION                         CUENTAS CONTABLES" & oImpresora.gPrnSaltoLinea
'End If
vRTFImp = vRTFImp & Space(2) & String(124, "-") & oImpresora.gPrnSaltoLinea
End Sub

Private Sub Form_Unload(Cancel As Integer)
CierraConexion
End Sub



