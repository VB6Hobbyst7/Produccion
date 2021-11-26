VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReportesAdmCred 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes Administración de Créditos"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   Icon            =   "frmReportesAdmCred.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   4935
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   915
      TabIndex        =   12
      Top             =   4920
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   7935
      Begin VB.CommandButton cmdAgencias 
         Caption         =   "Agencias"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   6840
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.PictureBox CR 
         Height          =   480
         Left            =   6240
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   14
         Top             =   3960
         Width           =   1200
      End
      Begin VB.Frame fraProductos 
         Caption         =   "Tipo de crédito"
         ForeColor       =   &H00800000&
         Height          =   4095
         Left            =   840
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   4050
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   3780
            Left            =   60
            TabIndex        =   11
            Top             =   195
            Width           =   3900
            _ExtentX        =   6879
            _ExtentY        =   6668
            _Version        =   393217
            Style           =   7
            Checkboxes      =   -1  'True
            Appearance      =   1
         End
      End
      Begin VB.CheckBox chkCliPref 
         Caption         =   "Cliente Preferencial"
         Height          =   255
         Left            =   5160
         TabIndex        =   9
         Top             =   2040
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ComboBox cboReportes 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   5895
      End
      Begin MSMask.MaskEdBox txtDel 
         Height          =   300
         Left            =   5040
         TabIndex        =   3
         Top             =   1200
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtAl 
         Height          =   300
         Left            =   5040
         TabIndex        =   4
         Top             =   1560
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList imglstFiguras 
         Left            =   5520
         Top             =   3960
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
               Picture         =   "frmReportesAdmCred.frx":030A
               Key             =   "Padre"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReportesAdmCred.frx":065C
               Key             =   "Bebe"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReportesAdmCred.frx":09AE
               Key             =   "Hijo"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReportesAdmCred.frx":0D00
               Key             =   "Hijito"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Al :"
         Height          =   195
         Left            =   4560
         TabIndex        =   8
         Top             =   1560
         Width           =   225
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Del :"
         Height          =   195
         Left            =   4440
         TabIndex        =   7
         Top             =   1200
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Reporte :"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "Procesar"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6840
      TabIndex        =   0
      Top             =   4920
      Width           =   1215
   End
End
Attribute VB_Name = "frmReportesAdmCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MatProductos() As String
Dim lsTitProductos

'Private Sub cboAgencias_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{Tab}", True
'    End If
'End Sub

Private Sub cboReportes_Click()
    'WIOR 20140205 ***********************************************
    If CInt(Right(Trim(Me.cboReportes.Text), 3)) = 14 Then
        chkCliPref.Visible = False
        fraProductos.Enabled = False
    Else
        fraProductos.Enabled = True
        If CInt(Right(Trim(Me.cboReportes.Text), 3)) = 2 Then
            chkCliPref.Visible = True
        Else
            chkCliPref.Visible = False
        End If
    End If
    'WIOR FIN *****************************************************
End Sub

Private Sub cboReportes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub cboReportes_LostFocus()
'    If CInt(Right(Trim(Me.cboReportes.Text), 3)) = 2 Then
'        chkCliPref.Visible = True
'    Else
'        chkCliPref.Visible = False
'    End If
End Sub


'RECO20141226*************
Private Sub cmdAgencias_Click()
    frmSelectAgencias.Inicio Me
    frmSelectAgencias.Show 1
End Sub
'RECI FIN*****************

'MADM 20110802
Private Sub TreeView1_Click()
     ActivaDes TreeView1.SelectedItem
End Sub

Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
Dim nExpande As Integer '1 Si no deja expandir a todos 0 Deja expandir a todos
Dim nUnico As Integer 'Valor del unico que se puede expander

nExpande = val(Text2.Text)
nUnico = val(Text1.Text)

If nExpande = 1 Then
    If InStr(Text1.Text, Mid(Node.Key, 2, 1)) > 0 Then
        Node.Expanded = True
    End If
End If
End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
Dim nExpande As Integer '1 Si no deja expandir a todos 0 Deja expandir a todos
Dim nUnico As Integer 'Valor del unico que se puede expander

nExpande = val(Text2.Text)
nUnico = val(Text1.Text)

If nExpande = 1 Then
    If InStr(Text1.Text, Mid(Node.Key, 2, 1)) = 0 Then
        Node.Expanded = False
        Node.Checked = False
    End If
End If

End Sub

Private Sub TreeView1_KeyPress(KeyAscii As Integer)
    ActivaDes TreeView1.SelectedItem
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
    Node.Selected = True
    ActivaDes Node
End Sub
 
Private Sub ActivaDes(sNode As Node)

    Dim i As Integer
    Dim nExpande As Integer

    nExpande = val(Text2.Text)
         
    If nExpande = 0 Then
'         If Mid(sNode.Key, 2, 1) <> Val(Text1.Text) Then
'            sNode.Checked = False
'         End If
        For i = 1 To TreeView1.Nodes.Count
            If Mid(TreeView1.Nodes(i).Key, 2, 1) = Mid(sNode.Key, 2, 1) And Mid(sNode.Key, 1, 1) = "P" Then
                TreeView1.Nodes(i).Checked = sNode.Checked
            End If
        Next
        
    ElseIf nExpande = 1 Then
        If InStr(Text1.Text, Mid(sNode.Key, 2, 1)) = 0 Then
            sNode.Checked = False
            sNode.Expanded = False
        Else
            TreeView1.SelectedItem = sNode
        Select Case Mid(sNode.Key, 1, 1)
        Case "P"
            If sNode.Checked = True Then
                 For i = 1 To TreeView1.Nodes.Count
                     If Mid(TreeView1.Nodes(i).Key, 2, 1) = Mid(sNode.Key, 2, 1) Then
                         TreeView1.Nodes(i).Checked = True
                     End If
                 Next
            Else
                 For i = 1 To TreeView1.Nodes.Count
                   If Mid(TreeView1.Nodes(i).Key, 2, 1) = Mid(sNode.Key, 2, 1) Then
                     TreeView1.Nodes(i).Checked = False
                   End If
                 Next
            End If

        End Select
        End If
    End If
End Sub

Private Sub ActFiltra(nFiltra As Boolean, Optional nFiltro As String = "")

    Dim i As Integer
    Dim nTempo As Integer
    
    If nFiltra = True Then
        Text2.Text = 1
        Text1.Text = nFiltro
         
        For i = 1 To TreeView1.Nodes.Count
            If InStr(nFiltro, Mid(TreeView1.Nodes(i).Key, 2, 1)) = 0 Then
                TreeView1.Nodes(i).Expanded = False
                TreeView1.Nodes(i).Checked = False
            Else
                TreeView1.Nodes(i).Expanded = True
                TreeView1.Nodes(i).Checked = False
            End If
        Next
        
    Else
        Text2.Text = ""
        Text1.Text = ""
        For i = 1 To TreeView1.Nodes.Count
            TreeView1.Nodes(i).Expanded = False
            TreeView1.Nodes(i).Checked = False
        Next
    End If
End Sub

Private Sub Limpia()
Dim i As Integer
    
    Text2.Text = ""
    Text1.Text = ""
    For i = 1 To TreeView1.Nodes.Count
        TreeView1.Nodes(i).Checked = False
        'TreeView1.Nodes(i).Expanded = False
    Next
 End Sub

Private Sub LlenaProductos()
Dim rs As ADODB.Recordset
Dim oreg As New DCredReporte
Dim sOpePadre As String
Dim sOpeHijo As String
Dim nodOpe As Node
TreeView1.Nodes.Clear
Set rs = New ADODB.Recordset

Set rs = oreg.GetProductos

Do While Not rs.EOF
          
        Select Case rs!cNivel
            Case "1"
                sOpePadre = "P" & rs!cValor
                Set nodOpe = TreeView1.Nodes.Add(, , sOpePadre, rs!cProducto)
                nodOpe.Tag = rs!cValor
            Case "2"
                sOpeHijo = "H" & rs!cValor
                Set nodOpe = TreeView1.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, rs!cProducto)
                nodOpe.Tag = rs!cValor
        
        End Select
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub
'END MADM

Private Sub cmdProcesar_Click()

Dim oCredDoc As COMNCredito.NCOMCredDoc
Dim sCadImp As String
Dim sCadImp_2 As String
Dim Prev As previo.clsprevio
Dim nTipoFormato As Integer

Dim oDCred As COMDCredito.DCOMCreditos
    
Dim lcAge As String
Dim lcNomAge As String
'MADM 20110802 ***********
Dim sCadProd As String
Dim i As Integer, lnPosI As Integer
Dim nContAge As Integer
sCadProd = ""
lcNomAge = ""
lsTitProductos = ""
'END
    'RECO20141226 *********************
    Dim sAgenciasTemp As String, nContAgencias As Integer, sCadAge As String
    For i = 0 To frmSelectAgencias.List1.ListCount - 1
        If frmSelectAgencias.List1.Selected(i) = True Then
            nContAge = nContAge + 1
            nContAgencias = nContAgencias + 1
            If Len(Trim(sCadAge)) = 0 Then
                sCadAge = Mid(frmSelectAgencias.List1.List(i), 1, 2)
                sAgenciasTemp = "" & Mid(frmSelectAgencias.List1.List(i), 1, 2) & ""
            Else
                sCadAge = sCadAge & "," & Mid(frmSelectAgencias.List1.List(i), 1, 2)
                sAgenciasTemp = sAgenciasTemp & ", " & Mid(frmSelectAgencias.List1.List(i), 1, 2) & ""
            End If
        End If
    Next i
    
    'RECO FIN**************************
    If Me.cboReportes.ListIndex = -1 Then
        MsgBox "Seleccione un reporte, por favor", vbOKOnly, "Atención"
        cboReportes.SetFocus
        Exit Sub
    End If
    
'    If Me.cboAgencias.ListIndex = -1 Then
'        MsgBox "Seleccione una Agencia, por favor", vbOKOnly, "Atención"
'        cboAgencias.SetFocus
'        Exit Sub
'    End If
    
    If Me.txtDel.Text = "__/__/____" Or Me.txtAl.Text = "__/__/____" Then
        MsgBox "Ingrese las fechas, por favor", vbOKOnly, "Atención"
        txtDel.SetFocus
        Exit Sub
    End If
    
    If IsDate(Me.txtAl.Text) Or IsDate(Me.txtAl.Text) Then
        If CDate(Me.txtDel.Text) > CDate(Me.txtAl.Text) Then
            MsgBox "La fecha final no puede ser menor a la fecha Inicial.", vbOKOnly, "Atención"
            txtDel.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Ingrese las fechas en formato Correcto, por favor", vbOKOnly, "Atención"
        txtDel.SetFocus
        Exit Sub
    End If
    
'    lcAge = Right(Trim(Me.cboAgencias.Text), 2)'RECO20141226
    lcAge = sCadAge
    'lcNomAge = Trim(Me.cboAgencias.Text)'RECO20141226
    lcNomAge = Trim(Mid(lcNomAge, 1, Len(Trim(lcNomAge)) - Len(Right(Trim(lcNomAge), 2))))
          
    sCadProd = "0"
    'MADM 20110802
    ReDim MatProductos(0)
    nContAge = 0
    sCadProd = "0"
    For i = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(i).Checked = True Then
            If Mid(TreeView1.Nodes(i).Key, 1, 1) = "H" Then
                nContAge = nContAge + 1
                ReDim Preserve MatProductos(nContAge)
                MatProductos(nContAge - 1) = Trim(Mid(TreeView1.Nodes(i).Key, 2, 3))
                sCadProd = sCadProd & "," & MatProductos(nContAge - 1)
                '**DAOR 20070717****
                lnPosI = 0
                lnPosI = InStr(1, TreeView1.Nodes(i).Text, " ")
                If lnPosI > 1 Then
                    lsTitProductos = lsTitProductos & Trim(Mid(TreeView1.Nodes(i).Text, 1, 3)) & " " & Mid(TreeView1.Nodes(i).Text, lnPosI + 1, 3) & "/"
                Else
                    lsTitProductos = lsTitProductos & Trim(Mid(TreeView1.Nodes(i).Text, 1, 3)) & "/"
                End If
                '*******************
            End If
        End If
    Next
    sCadProd = sCadProd & ""
    
    If Len(lsTitProductos) > 1 Then
        lsTitProductos = Left(lsTitProductos, Len(lsTitProductos) - 1)
    End If
    
    If CInt(Right(Trim(Me.cboReportes.Text), 3)) = 2 Or CInt(Right(Trim(Me.cboReportes.Text), 3)) = 4 Or CInt(Right(Trim(Me.cboReportes.Text), 3)) = 5 Or CInt(Right(Trim(Me.cboReportes.Text), 3)) = 6 Or CInt(Right(Trim(Me.cboReportes.Text), 3)) = 7 Or CInt(Right(Trim(Me.cboReportes.Text), 3)) = 8 Then
        If TreeView1.Visible = True Then
            If UBound(MatProductos) = 0 Then
                    MsgBox "Seleccione por lo menos un producto.", vbInformation, "Aviso"
                    Exit Sub
            End If
        End If
    End If
    'END MADM
    
    Select Case CInt(Right(Trim(Me.cboReportes.Text), 3))
    
        Case 1
            Call ImprimeVisitasClientes(lcAge, Format(Me.txtDel.Text, "yyyymmdd"), Format(Me.txtAl.Text, "yyyymmdd"), lcNomAge)
        Case 2
            Call ImprimeControlCreditos(lcAge, Format(Me.txtDel.Text, "yyyymmdd"), Format(Me.txtAl.Text, "yyyymmdd"), lcNomAge, Me.chkCliPref.value, MatProductos)
        Case 3
            Call ImprimeCreditosDesembolsadosPorDia(lcAge, Format(Me.txtDel.Text, "yyyymmdd"), Format(Me.txtAl.Text, "yyyymmdd"), lcNomAge)
        Case 4
            Call ImprimeCreditosDesembolsadosConObservaciones(lcAge, Format(Me.txtDel.Text, "yyyymmdd"), Format(Me.txtAl.Text, "yyyymmdd"), lcNomAge, MatProductos)
        Case 5
            Call ImprimeCreditosDesembolsadosConExoneraciones(lcAge, Format(Me.txtDel.Text, "yyyymmdd"), Format(Me.txtAl.Text, "yyyymmdd"), lcNomAge, MatProductos)
        Case 6
            Call ImprimeCantidadObservacionesPorAnalista(lcAge, Format(Me.txtDel.Text, "yyyymmdd"), Format(Me.txtAl.Text, "yyyymmdd"), lcNomAge, MatProductos)
        Case 7
            Call ImprimeCantidadExoneracionesPorAnalista(lcAge, Format(Me.txtDel.Text, "yyyymmdd"), Format(Me.txtAl.Text, "yyyymmdd"), lcNomAge, MatProductos)
        Case 8
            Call ImprimeCantidadExoneracionesPorRespApr(lcAge, Format(Me.txtDel.Text, "yyyymmdd"), Format(Me.txtAl.Text, "yyyymmdd"), lcNomAge, MatProductos)
        Case 9
            'Call ImprimeCreditosDesembolsadosConPostDesembolso(lcAge, Format(Me.txtDel.Text, "yyyymmdd"), Format(Me.txtAl.Text, "yyyymmdd"), lcNomAge)
            Call ImprimeCreditosCompraDeuda(lcAge, Format(Me.txtDel.Text, "yyyymmdd"), Format(Me.txtAl.Text, "yyyymmdd"), lcNomAge) 'WIOR 20120508
        Case 10
            Call ImprimeCreditosConstitucion(lcAge, Format(Me.txtDel.Text, "yyyymmdd"), Format(Me.txtAl.Text, "yyyymmdd"), lcNomAge)
        Case 11
            Call ImprimeCreditosDesembolsadosConAutorizaciones(lcAge, Format(Me.txtDel.Text, "yyyymmdd"), Format(Me.txtAl.Text, "yyyymmdd"), lcNomAge, MatProductos)
        Case 12
            Call ImprimeCreditosDesembolsadosConPostDesembolso(lcAge, Format(Me.txtDel.Text, "yyyymmdd"), Format(Me.txtAl.Text, "yyyymmdd"), lcNomAge, MatProductos)
        Case 13
            Call ImprimeCFObservaciones(lcAge, Format(Me.txtDel.Text, "yyyymmdd"), Format(Me.txtAl.Text, "yyyymmdd"), lcNomAge)
        'WIOR 20140205*********************
        Case 14
            Call ImprimeAsignacionSaldos(lcAge, Format(Me.txtDel.Text, "yyyymmdd"), Format(Me.txtAl.Text, "yyyymmdd"), lcNomAge)
        'WIOR FIN *************************
    End Select
    Limpia
End Sub
'MADM 20110802 - Optional pMatProd
Private Sub ImprimeControlCreditos(ByVal psCodAge As String, ByVal psDel As String, ByVal psAl As String, ByVal psNomAge As String, Optional ByVal pnCliPref As Integer = 0, Optional ByVal pMatProd As Variant = Nothing)

    Dim fs As Scripting.FileSystemObject
    Dim pRs As ADODB.Recordset
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
    Dim lsArchivo2 As String
    Dim lbLibroOpen As Boolean
    Dim lsNomHoja  As String
    Dim lsMes As String

    Dim oDControl As COMDCredito.DCOMCreditos

    Dim R As ADODB.Recordset

    Dim sCadImp As String, i As Integer, J As Integer
    Dim lsNombreArchivo As String
    Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String
    
    Screen.MousePointer = 11
    
    Set oDControl = New COMDCredito.DCOMCreditos
        Set R = oDControl.RecuperaDatosAprobacionCreditos(psCodAge, psDel, psAl, pnCliPref, pMatProd)
    Set oDControl = Nothing
    
    If R.EOF And R.BOF Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

    'Determinando que Archivo y hoja Excel se debe abrir de acuerdo a eleccion del usuario

    'lsArchivo1 = "ControlCreditosAdmCred"
    lsArchivo1 = "ControlCreditosAdmCredNuevo" 'WIOR 20120620
    lsNomHoja = "AdmCred"
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application

    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo1 & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo1 & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    lsArchivo2 = lsArchivo1 & "_" & gsCodUser & "_" & Format$(gdFecSis, "yyyymmdd") & "_" & Format$(Time(), "HHMMSS")

    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If

    Call InsertaDatosControlAdmCred(R, psCodAge, psDel, psAl, psNomAge, xlHoja1)

    xlHoja1.SaveAs App.path & "\Spooler\" & lsArchivo2 & ".xls"
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

    Screen.MousePointer = 0

End Sub
'MADM 20110803 - parametro convenio
Public Sub InsertaDatosControlAdmCred(ByRef pR As ADODB.Recordset, ByVal psCodAge As String, ByVal psDel As String, ByVal psAl As String, ByVal psNomAge As String, ByRef xlHoja1 As Excel.Worksheet)

Dim i As Integer
Dim lnliquidez As Double, lnCapacidadPago As Double, lnExcedente As Double
Dim lnPatriEmpre As Double, lnPatrimonio As Double, lnIngresoNeto As Double
Dim lnRentabPatrimonial As Double, lnEndeudamiento As Double
Dim lcTpoGarantia As String
Dim lnPorGravar As Double, lnGravado As Double

Dim lcFecDel As String, lcFecAl As String
Dim lnItem As Integer

'MADM 20110309
Dim cPersNombre As String
Dim nChangeSituacion As Integer
cPersNombre = ""
nChangeSituacion = 0
'END MADM

lcFecDel = Mid(psDel, 7, 2) & "/" & Mid(psDel, 5, 2) & "/" & Mid(psDel, 1, 4)
lcFecAl = Mid(psAl, 7, 2) & "/" & Mid(psAl, 5, 2) & "/" & Mid(psAl, 1, 4)

xlHoja1.Cells(2, 1) = Trim(psNomAge) & " - " & " DEL : " & lcFecDel & " AL : " & lcFecAl

i = 3
lnItem = 0

Do While Not pR.EOF
        i = i + 1
        
'        If pR!cPersNombre <> cPersNombre Then 'madm 20110309
'            nChangeSituacion = 1
'            lnItem = lnItem + 1
'        End If
'
'        If nChangeSituacion = 1 Then 'madm 20110309
            xlHoja1.Cells(i, 1) = lnItem
            xlHoja1.Cells(i, 2) = pR!tipoprod
            xlHoja1.Cells(i, 3) = pR!dVigencia 'WIOR 20120620
            xlHoja1.Cells(i, 4) = pR!dFecRevision
            xlHoja1.Cells(i, 5) = pR!cPersNombre
            xlHoja1.Cells(i, 6) = pR!Analista
            xlHoja1.Cells(i, 7) = pR!actividad
            xlHoja1.Cells(i, 8) = pR!Modalidad
            xlHoja1.Cells(i, 9) = pR!Monto
            xlHoja1.Cells(i, 10) = pR!Moneda
            xlHoja1.Cells(i, 11) = pR!ExposiCred
            xlHoja1.Cells(i, 12) = pR!Plazo
            xlHoja1.Cells(i, 13) = pR!cDestinoDescripcion
            nChangeSituacion = 0
'        End If
        
        xlHoja1.Cells(i, 14) = pR!cRegulariza
        xlHoja1.Cells(i, 15) = pR!cObservacion
        xlHoja1.Cells(i, 16) = pR!exoneracion
        xlHoja1.Cells(i, 17) = pR!quienexonera
        xlHoja1.Cells(i, 18) = pR!cUserExonera
        xlHoja1.Cells(i, 19) = pR!Convenio

        xlHoja1.Rows(i + 1).Select
        xlHoja1.Range("A" + Trim(str(i + 1))).EntireRow.Insert
        cPersNombre = pR!cPersNombre
    pR.MoveNext
Loop

xlHoja1.Cells.Select
xlHoja1.Cells.EntireColumn.AutoFit
xlHoja1.Cells(4, 1).Select

End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    CentraForm Me
    
    Dim oCred As COMDCredito.DCOMCreditos
    Dim oCons As COMDConstantes.DCOMConstantes
    Dim rs As ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Dim L As ListItem
    
    Set oCons = New COMDConstantes.DCOMConstantes
        Set rs = oCons.RecuperaConstantes(9008)
    Set oCons = Nothing
    Call Llenar_Combo_con_Recordset(rs, Me.cboReportes)

    Set oCred = New COMDCredito.DCOMCreditos
        Set rs1 = oCred.ObtieneAgenciasAdmCred
    Set oCred = Nothing
'    Call CargaAgenciasAdmCred(rs1)
    
    'MADM 20110802
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    ReDim MatProductos(0)
    fraProductos.Visible = True
    ActFiltra False
    LlenaProductos
    Limpia
    'END MADM

End Sub

'Sub CargaAgenciasAdmCred(ByVal pRs As ADODB.Recordset)
'
'    On Error GoTo ErrHandler
'
'        Do Until pRs.EOF
'
'            Me.cboAgencias.AddItem Left(pRs!cAgeDescripcion, 40) + Space(50) + pRs!cAgeCod
'
'            pRs.MoveNext
'        Loop
'    Exit Sub
'ErrHandler:
'    MsgBox "Error al cargar CargaAgenciasAdmCred", vbInformation, "AVISO"
'End Sub

Private Sub txtAl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtDel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub ImprimeCreditosDesembolsadosPorDia(ByVal psCodAge As String, _
    ByVal psDel As String, ByVal psAl As String, ByVal psNomAge As String)

Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset

Dim sCadImp As String, i As Integer, J As Integer
Dim lsNombreArchivo As String, lcTipProd As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String
Dim lcValor As Integer
Dim lcFecDel As String, lcFecAl As String

    Screen.MousePointer = 11
    Set oDCred = New COMDCredito.DCOMCreditos
        Set R = oDCred.ObtieneCredDesembolsadosPorDia(psCodAge, psDel, psAl)
    Set oDCred = Nothing
    
    If R.EOF And R.BOF Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

lcFecDel = Mid(psDel, 7, 2) & "/" & Mid(psDel, 5, 2) & "/" & Mid(psDel, 1, 4)
lcFecAl = Mid(psAl, 7, 2) & "/" & Mid(psAl, 5, 2) & "/" & Mid(psAl, 1, 4)

'**************************************************************************************
    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = "CAJA MAYNAS"
    ApExcel.Cells(2, 7).Formula = Date + Time()
    ApExcel.Cells(4, 2).Formula = "CREDITOS DESEMBOLSADOS"
    ApExcel.Cells(5, 2).Formula = Trim(psNomAge) & " - " & " DEL : " & lcFecDel & " AL : " & lcFecAl
    
    ApExcel.Range("B4", "G4").MergeCells = True
    ApExcel.Range("B5", "G5").MergeCells = True
    ApExcel.Range("B4", "G5").HorizontalAlignment = xlCenter
    
    ApExcel.Cells(8, 2).Formula = "CUENTA"
    ApExcel.Cells(8, 3).Formula = "MONEDA"
    ApExcel.Cells(8, 4).Formula = "CLIENTE"
    ApExcel.Cells(8, 5).Formula = "ANALISTA"
    ApExcel.Cells(8, 6).Formula = "MONTO"
    ApExcel.Cells(7, 7).Formula = "REGISTRO"
    ApExcel.Cells(8, 7).Formula = "CONTROL"
    ApExcel.Cells(8, 8).Formula = "SUBPRODUCTO" 'RECO 20130915 ERS133
    '**************RECO 20130915 ERS133*******************
    'ApExcel.Range("B2", "G8").Font.Bold = True
    
    'ApExcel.Range("B7", "G8").Interior.Color = RGB(10, 190, 160)
    'ApExcel.Range("B7", "G8").HorizontalAlignment = 3
    
    ApExcel.Range("B2", "H8").Font.Bold = True
    
    ApExcel.Range("B7", "H8").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B7", "H8").HorizontalAlignment = 3
    '********END RECO

    i = 8
    
    Do While Not R.EOF
    i = i + 1
            
                
                ApExcel.Cells(i, 2).Formula = R!cCtaCod
                ApExcel.Cells(i, 3).Formula = R!Moneda
                ApExcel.Cells(i, 4).Formula = R!cPersNombre
                ApExcel.Cells(i, 5).Formula = R!Analista
                ApExcel.Cells(i, 6).Formula = R!nMonto
                ApExcel.Cells(i, 7).Formula = R!registrocontrol
                ApExcel.Cells(i, 8).Formula = R!cSubProducto  'RECO 20130915 ERS133
            R.MoveNext
            
        Loop

    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing
    
End Sub

Private Sub ImprimeCreditosConstitucion(ByVal psCodAge As String, _
    ByVal psDel As String, ByVal psAl As String, ByVal psNomAge As String)

Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset

Dim sCadImp As String, i As Integer, J As Integer
Dim lsNombreArchivo As String, lcTipProd As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String
Dim lcValor As Integer
Dim lcFecDel As String, lcFecAl As String

    Screen.MousePointer = 11
    Set oDCred = New COMDCredito.DCOMCreditos
        Set R = oDCred.ObtieneCreditosConstitucion(psCodAge, psDel, psAl)
        
    Set oDCred = Nothing
    
    If R.EOF And R.BOF Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

    lcFecDel = Mid(psDel, 7, 2) & "/" & Mid(psDel, 5, 2) & "/" & Mid(psDel, 1, 4)
    lcFecAl = Mid(psAl, 7, 2) & "/" & Mid(psAl, 5, 2) & "/" & Mid(psAl, 1, 4)

'**************************************************************************************
    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = "CAJA MAYNAS"
    ApExcel.Cells(2, 6).Formula = Date + Time()
    ApExcel.Cells(4, 2).Formula = "CREDITOS DE CONSTITUCION DE GARANTIAS"
    ApExcel.Cells(5, 2).Formula = Trim(psNomAge) & " - " & " DEL : " & lcFecDel & " AL : " & lcFecAl
    
    ApExcel.Range("B4", "F4").MergeCells = True
    ApExcel.Range("B5", "F5").MergeCells = True
    ApExcel.Range("B4", "F5").HorizontalAlignment = xlCenter
    
    ApExcel.Cells(8, 2).Formula = "CUENTA"
    ApExcel.Cells(8, 3).Formula = "SUBPRODUCTO" 'RECO 20130915 ERS133
    ApExcel.Cells(8, 4).Formula = "MONEDA"
    ApExcel.Cells(8, 5).Formula = "CLIENTE"
    ApExcel.Cells(8, 6).Formula = "ANALISTA"
    ApExcel.Cells(8, 7).Formula = "MONTO"
       
    ApExcel.Range("B2", "H8").Font.Bold = True
    
    '**************RECO 20130915 ERS133**************************
    'ApExcel.Range("B7", "H8").Interior.Color = RGB(10, 190, 160)
    'ApExcel.Range("B7", "H8").HorizontalAlignment = 3
    ApExcel.Range("B7", "I8").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B7", "I8").HorizontalAlignment = 3
    '********END RECO *******************************************
    i = 8
    Do While Not R.EOF
    i = i + 1
        ApExcel.Cells(i, 2).Formula = R!cCtaCod
        ApExcel.Cells(i, 3).Formula = R!cSubProducto 'RECO 20130915 ERS133
        ApExcel.Cells(i, 4).Formula = R!cmoneda
        ApExcel.Cells(i, 5).Formula = R!cCliente
        ApExcel.Cells(i, 6).Formula = R!cAnalista
        ApExcel.Cells(i, 7).Formula = R!nMonto
        R.MoveNext
        Loop
    R.Close
    Set R = Nothing
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Range("B2").Select
    Screen.MousePointer = 0
    ApExcel.Visible = True
    Set ApExcel = Nothing
End Sub

'MADM 20110802 - pMatProd
Private Sub ImprimeCreditosDesembolsadosConObservaciones(ByVal psCodAge As String, _
    ByVal psDel As String, ByVal psAl As String, ByVal psNomAge As String, Optional ByVal pMatProd As Variant = Nothing)

Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset

Dim sCadImp As String, i As Integer, J As Integer
Dim lsNombreArchivo As String, lcTipProd As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String
Dim lcValor As Integer
Dim lcFecDel As String, lcFecAl As String

'MADM 20110309
Dim vcCtaCod As String
Dim nChangeSituacion As Integer
vcCtaCod = ""
nChangeSituacion = 0
'END MADM

    Screen.MousePointer = 11
    Set oDCred = New COMDCredito.DCOMCreditos
        Set R = oDCred.ObtieneCredDesembolsadosConObservaciones(psCodAge, psDel, psAl, pMatProd)
    Set oDCred = Nothing
    
    If R.EOF And R.BOF Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

    lcFecDel = Mid(psDel, 7, 2) & "/" & Mid(psDel, 5, 2) & "/" & Mid(psDel, 1, 4)
    lcFecAl = Mid(psAl, 7, 2) & "/" & Mid(psAl, 5, 2) & "/" & Mid(psAl, 1, 4)

'**************************************************************************************
    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = "CAJA MAYNAS"
    ApExcel.Cells(2, 7).Formula = Date + Time()
    ApExcel.Cells(4, 2).Formula = "CREDITOS DESEMBOLSADOS CON OBSERVACIONES"
    ApExcel.Cells(5, 2).Formula = Trim(psNomAge) & " - " & " DEL : " & lcFecDel & " AL : " & lcFecAl
    
    ApExcel.Range("B4", "G4").MergeCells = True
    ApExcel.Range("B5", "G5").MergeCells = True
    ApExcel.Range("B4", "G5").HorizontalAlignment = xlCenter
    
    ApExcel.Cells(8, 2).Formula = "FECHA REV"
    'ApExcel.Cells(8, 3).Formula = "CUENTA"
    ApExcel.Cells(8, 3).Formula = "CLIENTE"
    ApExcel.Cells(8, 4).Formula = "ANALISTA"
    ApExcel.Cells(8, 5).Formula = "MONEDA"
    ApExcel.Cells(8, 6).Formula = "MONTO"
    ApExcel.Cells(8, 7).Formula = "OBSERVACIONES"
    ApExcel.Cells(8, 8).Formula = "SUBPRODUCTO" 'RECO 20130915 ERS133
    
    '************RECO 20131015********************
    'ApExcel.Range("B2", "G8").Font.Bold = True
    
    'ApExcel.Range("B7", "G8").Interior.Color = RGB(10, 190, 160)
    'ApExcel.Range("B7", "G8").HorizontalAlignment = 3
    
    
    ApExcel.Range("B2", "H8").Font.Bold = True
    
    ApExcel.Range("B7", "H8").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B7", "H8").HorizontalAlignment = 3
    '************END RECO**************************
    i = 8
       
    Do While Not R.EOF
    i = i + 1
                         
        If R!cCtaCod <> vcCtaCod Then 'madm 20110309
            nChangeSituacion = 1
        End If
        
        If nChangeSituacion = 1 Then
            ApExcel.Cells(i, 2).Formula = R!dFecRevision
            'ApExcel.Cells(i, 3).Formula = R!cCtaCod
            ApExcel.Cells(i, 3).Formula = R!cPersNombre
            ApExcel.Cells(i, 4).Formula = R!Analista
            ApExcel.Cells(i, 5).Formula = R!Moneda
            ApExcel.Cells(i, 6).Formula = R!nMonto
            ApExcel.Cells(i, 6).Formula = R!nMonto
            ApExcel.Cells(i, 8).Formula = R!cSubProducto  'RECO 20130915 ERS133
            nChangeSituacion = 0
        End If
        
        ApExcel.Cells(i, 7).Formula = R!cObservacion
        vcCtaCod = R!cCtaCod
        
        R.MoveNext
            
        Loop
            
    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
'    ApExcel.Columns("B:B").ColumnWidth = 6#
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing
    
End Sub

'MADM 20110802 - ImprimeCreditosCompraDeuda
Private Sub ImprimeCreditosDesembolsadosConExoneraciones(ByVal psCodAge As String, _
    ByVal psDel As String, ByVal psAl As String, ByVal psNomAge As String, Optional ByVal pMatProd As Variant = Nothing)

Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset

Dim sCadImp As String, i As Integer, J As Integer
Dim lsNombreArchivo As String, lcTipProd As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String
Dim lcValor As Integer

'MADM 20110309
Dim vcCtaCod As String
Dim nChangeSituacion As Integer
vcCtaCod = ""
nChangeSituacion = 0
'END MADM

Dim lcFecDel As String, lcFecAl As String

    Screen.MousePointer = 11
    Set oDCred = New COMDCredito.DCOMCreditos
        Set R = oDCred.ObtieneCredDesembolsadosConExoneraciones(psCodAge, psDel, psAl, pMatProd)
    Set oDCred = Nothing
    
    If R.EOF And R.BOF Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

    lcFecDel = Mid(psDel, 7, 2) & "/" & Mid(psDel, 5, 2) & "/" & Mid(psDel, 1, 4)
    lcFecAl = Mid(psAl, 7, 2) & "/" & Mid(psAl, 5, 2) & "/" & Mid(psAl, 1, 4)

'**************************************************************************************
    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = "CAJA MAYNAS"
    ApExcel.Cells(2, 9).Formula = Date + Time()
    ApExcel.Cells(4, 2).Formula = "CREDITOS DESEMBOLSADOS CON EXONERACIONES"
    ApExcel.Cells(5, 2).Formula = Trim(psNomAge) & " - " & " DEL : " & lcFecDel & " AL : " & lcFecAl
    
    ApExcel.Range("B4", "I4").MergeCells = True
    ApExcel.Range("B5", "I5").MergeCells = True
    ApExcel.Range("B4", "I5").HorizontalAlignment = xlCenter
    
    ApExcel.Cells(8, 2).Formula = "AGENCIA"
    ApExcel.Cells(8, 3).Formula = "CUENTA"
    ApExcel.Cells(8, 4).Formula = "MONEDA"
    ApExcel.Cells(8, 5).Formula = "CLIENTE"
    ApExcel.Cells(8, 6).Formula = "ANALISTA"
    ApExcel.Cells(8, 7).Formula = "MONTO"
    ApExcel.Cells(8, 8).Formula = "EXONERACIONES"
    ApExcel.Cells(8, 9).Formula = "QUIEN EXONERA"
    ApExcel.Cells(8, 10).Formula = "USU. EXONERA"
    ApExcel.Cells(8, 11).Formula = "SUBPRODUCTO" ' **************RECO 20131015 ERS133
    'ApExcel.Cells(8, 10).Formula = "ESPECIFICACION"
    
    '**************RECO 20131015 ERS133******
    ApExcel.Range("B2", "K8").Font.Bold = True
    
    ApExcel.Range("B7", "K8").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B7", "K8").HorizontalAlignment = 3
    
    'ApExcel.Range("B2", "I8").Font.Bold = True
    
    'ApExcel.Range("B7", "I8").Interior.Color = RGB(10, 190, 160)
    'ApExcel.Range("B7", "I8").HorizontalAlignment = 3
    '*************END RECO*****************
    i = 8
       
    Do While Not R.EOF
    i = i + 1
                           
        If R!cCtaCod <> vcCtaCod Then 'madm 20110309
            nChangeSituacion = 1
        End If
        
        If nChangeSituacion = 1 Then
            ApExcel.Cells(i, 2).Formula = R!cAgeDescripcion
            ApExcel.Cells(i, 3).Formula = R!cCtaCod
            ApExcel.Cells(i, 4).Formula = R!Moneda
            ApExcel.Cells(i, 5).Formula = R!cPersNombre
            ApExcel.Cells(i, 6).Formula = R!Analista
            ApExcel.Cells(i, 7).Formula = R!nMonto
            nChangeSituacion = 0
        End If
        
        ApExcel.Cells(i, 8).Formula = R!exoneracion '*** PEAC 20101223
        ApExcel.Cells(i, 9).Formula = R!quienexonera
        ApExcel.Cells(i, 10).Formula = R!cUserExonera
        '************RECO 20131015 ERS133***************
        'ApExcel.Cells(i, 10).Formula = R!CDescripcionOtro
        ApExcel.Cells(i, 11).Formula = R!cSubProducto
        ApExcel.Cells(i, 12).Formula = R!CDescripcionOtro
        '**************END RECO************
        vcCtaCod = R!cCtaCod
        R.MoveNext
            
        Loop
            
    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing
    
End Sub
'MADM 20110802
Private Sub ImprimeCantidadObservacionesPorAnalista(ByVal psCodAge As String, _
    ByVal psDel As String, ByVal psAl As String, ByVal psNomAge As String, Optional ByVal pMatProd As Variant = Nothing)

Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset

Dim sCadImp As String, i As Integer, J As Integer
Dim lsNombreArchivo As String, lcTipProd As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String
Dim lcValor As Integer
Dim lcFecDel As String, lcFecAl As String

    Screen.MousePointer = 11
    Set oDCred = New COMDCredito.DCOMCreditos
        Set R = oDCred.ObtieneCantidadObservacionesPorAnalista(psCodAge, psDel, psAl, pMatProd)
    Set oDCred = Nothing
    
    If R.EOF And R.BOF Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

    lcFecDel = Mid(psDel, 7, 2) & "/" & Mid(psDel, 5, 2) & "/" & Mid(psDel, 1, 4)
    lcFecAl = Mid(psAl, 7, 2) & "/" & Mid(psAl, 5, 2) & "/" & Mid(psAl, 1, 4)

'**************************************************************************************
    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = "CAJA MAYNAS"
    ApExcel.Cells(2, 4).Formula = Date + Time()
    ApExcel.Cells(4, 2).Formula = "CANTIDAD DE OBSERVACIONES POR ANALISTA"
    ApExcel.Cells(5, 2).Formula = Trim(psNomAge) & " - " & " DEL : " & lcFecDel & " AL : " & lcFecAl
    
    ApExcel.Range("B4", "E4").MergeCells = True
    ApExcel.Range("B5", "E5").MergeCells = True
    ApExcel.Range("B4", "E5").HorizontalAlignment = xlCenter
    
    ApExcel.Cells(8, 2).Formula = "ANALISTA"
    ApExcel.Cells(8, 4).Formula = "CANTIDAD"
    ApExcel.Cells(8, 5).Formula = "TIPO CREDITO"
    
    ApExcel.Cells(8, 6).Formula = "SUBPRODUCTO" '********RECO 20131015 ERS133
    
    '*****************RECO 20131015 ERS133****************
    'ApExcel.Range("B2", "E8").Font.Bold = True
    
    'ApExcel.Range("B7", "E8").Interior.Color = RGB(10, 190, 160)
    'ApExcel.Range("B7", "E8").HorizontalAlignment = 3
    
    ApExcel.Range("B2", "F8").Font.Bold = True
    
    ApExcel.Range("B7", "F8").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B7", "F8").HorizontalAlignment = 3
    '**********************END RECO**********************
    i = 8
       
    Do While Not R.EOF
    i = i + 1
                            
        ApExcel.Cells(i, 2).Formula = R!Analista
        ApExcel.Cells(i, 4).Formula = R!Cantidad
        ApExcel.Cells(i, 5).Formula = R!cConsDescripcion
        ApExcel.Cells(i, 6).Formula = R!cSubProducto '********RECO 20131015 ERS133
        R.MoveNext
            
        Loop
            
    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
'    ApExcel.Columns("B:B").ColumnWidth = 6#
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing
    
End Sub
'MADM 20110802
Private Sub ImprimeCantidadExoneracionesPorAnalista(ByVal psCodAge As String, _
    ByVal psDel As String, ByVal psAl As String, ByVal psNomAge As String, Optional ByVal pMatProd As Variant = Nothing)

Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset

Dim sCadImp As String, i As Integer, J As Integer
Dim lsNombreArchivo As String, lcTipProd As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String
Dim lcValor As Integer
Dim lcFecDel As String, lcFecAl As String

    Screen.MousePointer = 11
    Set oDCred = New COMDCredito.DCOMCreditos
        Set R = oDCred.ImprimeCantidadExoneracionesPorAnalista(psCodAge, psDel, psAl, pMatProd)
    Set oDCred = Nothing
    
    If R.EOF And R.BOF Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

    lcFecDel = Mid(psDel, 7, 2) & "/" & Mid(psDel, 5, 2) & "/" & Mid(psDel, 1, 4)
    lcFecAl = Mid(psAl, 7, 2) & "/" & Mid(psAl, 5, 2) & "/" & Mid(psAl, 1, 4)

'**************************************************************************************
    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = "CAJA MAYNAS"
    ApExcel.Cells(2, 4).Formula = Date + Time()
    ApExcel.Cells(4, 2).Formula = "CANTIDAD DE EXONERACIONES POR ANALISTA"
    ApExcel.Cells(5, 2).Formula = Trim(psNomAge) & " - " & " DEL : " & lcFecDel & " AL : " & lcFecAl
    
    ApExcel.Range("B4", "E4").MergeCells = True
    ApExcel.Range("B5", "E5").MergeCells = True
    ApExcel.Range("B4", "E5").HorizontalAlignment = xlCenter
    
    ApExcel.Cells(8, 2).Formula = "ANALISTA"
    ApExcel.Cells(8, 4).Formula = "CANTIDAD"
    ApExcel.Cells(8, 5).Formula = "DESCRIPCION"
    ApExcel.Cells(8, 6).Formula = "SUBPRODUCTO" '*******RECO 20131015 ERS133
    
    '*************RECO 20131015 ERS133************
    ApExcel.Range("B2", "F8").Font.Bold = True
    
    ApExcel.Range("B7", "F8").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B7", "F8").HorizontalAlignment = 3
    
    'ApExcel.Range("B2", "E8").Font.Bold = True
    
    'ApExcel.Range("B7", "E8").Interior.Color = RGB(10, 190, 160)
    'ApExcel.Range("B7", "E8").HorizontalAlignment = 3
    '*************END RECO************************
    i = 8
       
    Do While Not R.EOF
    i = i + 1
                            
        ApExcel.Cells(i, 2).Formula = R!Analista
        ApExcel.Cells(i, 4).Formula = R!Cantidad
        ApExcel.Cells(i, 5).Formula = R!cConsDescripcion
        ApExcel.Cells(i, 6).Formula = R!cSubProducto '******RECO 20131015 ERS133
        
        R.MoveNext
            
        Loop
            
    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
'    ApExcel.Columns("B:B").ColumnWidth = 6#
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing
    
End Sub
'MADM 20110802
Private Sub ImprimeCantidadExoneracionesPorRespApr(ByVal psCodAge As String, _
    ByVal psDel As String, ByVal psAl As String, ByVal psNomAge As String, Optional ByVal pMatProd As Variant = Nothing)

Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset

Dim sCadImp As String, i As Integer, J As Integer
Dim lsNombreArchivo As String, lcTipProd As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String
Dim lcValor As Integer
Dim lcFecDel As String, lcFecAl As String

    Screen.MousePointer = 11
    Set oDCred = New COMDCredito.DCOMCreditos
        Set R = oDCred.ImprimeCantidadExoneracionesPorRespApr(psCodAge, psDel, psAl, pMatProd)
    Set oDCred = Nothing
    
    If R.EOF And R.BOF Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

    lcFecDel = Mid(psDel, 7, 2) & "/" & Mid(psDel, 5, 2) & "/" & Mid(psDel, 1, 4)
    lcFecAl = Mid(psAl, 7, 2) & "/" & Mid(psAl, 5, 2) & "/" & Mid(psAl, 1, 4)

'**************************************************************************************
    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = "CAJA MAYNAS"
    ApExcel.Cells(2, 4).Formula = Date + Time()
    ApExcel.Cells(4, 2).Formula = "CANTIDAD DE EXON. POR RESPONSABLE DE APROBACION"
    ApExcel.Cells(5, 2).Formula = Trim(psNomAge) & " - " & " DEL : " & lcFecDel & " AL : " & lcFecAl
    
    ApExcel.Range("B4", "D4").MergeCells = True
    ApExcel.Range("B5", "D5").MergeCells = True
    ApExcel.Range("B4", "D5").HorizontalAlignment = xlCenter
    
    ApExcel.Cells(8, 2).Formula = "TIPO CRED"
    ApExcel.Cells(8, 3).Formula = "SUBPRODUCTO" '*******RECO 20131015 ERS133
    ApExcel.Cells(8, 4).Formula = "RESPONSABLE"
    ApExcel.Cells(8, 5).Formula = "CANTIDAD"
    ApExcel.Cells(8, 6).Formula = "OTRO"
    ApExcel.Cells(8, 7).Formula = "Plazo recorrido"
    ApExcel.Cells(8, 8).Formula = "Gravamen/tasaciones"
    ApExcel.Cells(8, 9).Formula = "Número de Ifis"
    ApExcel.Cells(8, 10).Formula = "CAR"
    ApExcel.Cells(8, 11).Formula = "Garantía favor otra Ifi"
    ApExcel.Cells(8, 12).Formula = "Cobertura de Garantía"
    ApExcel.Cells(8, 13).Formula = "Garantía no inscrita"
    ApExcel.Cells(8, 14).Formula = "Más de 14 veces su sueldo"
    ApExcel.Cells(8, 15).Formula = "Calificación SBS"
    ApExcel.Cells(8, 16).Formula = "Capitalizacion, interes y mora"
    ApExcel.Cells(8, 17).Formula = "Plazos Mayores a lo permitido"
    
    '*******RECO 20131015 ERS133********
    'ApExcel.Range("B2", "P8").Font.Bold = True
    'ApExcel.Range("B7", "P8").Interior.Color = RGB(10, 190, 160)
    'ApExcel.Range("B7", "P8").HorizontalAlignment = 3
    ApExcel.Range("B2", "Q8").Font.Bold = True
    ApExcel.Range("B7", "Q8").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B7", "Q8").HorizontalAlignment = 3
    '*******END RECO ********************

    i = 8
    Do While Not R.EOF
    i = i + 1
        ApExcel.Cells(i, 2).Formula = R!cTpoCred
        ApExcel.Cells(i, 3).Formula = R!cSubProducto '*******RECO 20131015 ERS133
        ApExcel.Cells(i, 4).Formula = R!RespApro
        ApExcel.Cells(i, 5).Formula = R!Cantidad
        ApExcel.Cells(i, 6).Formula = R!Otro
        ApExcel.Cells(i, 7).Formula = R!Plazo
        ApExcel.Cells(i, 8).Formula = R!Gravamen
        ApExcel.Cells(i, 9).Formula = R!NumeroIfis
        ApExcel.Cells(i, 10).Formula = R!Car
        ApExcel.Cells(i, 11).Formula = R!GaranIfi
        ApExcel.Cells(i, 12).Formula = R!Cober
        ApExcel.Cells(i, 13).Formula = R!GaranNoIns
        ApExcel.Cells(i, 14).Formula = R!Mas14sueldo
        ApExcel.Cells(i, 15).Formula = R!sBS
        ApExcel.Cells(i, 16).Formula = R!Capitalizacion
        ApExcel.Cells(i, 17).Formula = R!Plazos

        R.MoveNext
        Loop

    R.Close
    Set R = Nothing

    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing
    
End Sub

Private Sub ImprimeVisitasClientes(ByVal psCodAge As String, _
    ByVal psDel As String, ByVal psAl As String, ByVal psNomAge As String)

Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset

Dim sCadImp As String, i As Integer, J As Integer
Dim lsNombreArchivo As String, lcTipProd As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String
Dim lcValor As Integer
Dim lcFecDel As String, lcFecAl As String

    Screen.MousePointer = 11
    Set oDCred = New COMDCredito.DCOMCreditos
        Set R = oDCred.ObtieneVisitasClientes(psCodAge, psDel, psAl)
    Set oDCred = Nothing
    
    If R.EOF And R.BOF Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

    lcFecDel = Mid(psDel, 7, 2) & "/" & Mid(psDel, 5, 2) & "/" & Mid(psDel, 1, 4)
    lcFecAl = Mid(psAl, 7, 2) & "/" & Mid(psAl, 5, 2) & "/" & Mid(psAl, 1, 4)

'**************************************************************************************
    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = "CAJA MAYNAS"
    ApExcel.Cells(2, 15).Formula = Date + Time()
    ApExcel.Cells(4, 2).Formula = "VISITAS A CLIENTES"
    ApExcel.Cells(5, 2).Formula = Trim(psNomAge) & " - " & " DEL : " & lcFecDel & " AL : " & lcFecAl
    
    ApExcel.Range("B4", "O4").MergeCells = True
    ApExcel.Range("B5", "O5").MergeCells = True
    ApExcel.Range("B4", "O5").HorizontalAlignment = xlCenter
    
    ApExcel.Cells(8, 2).Formula = "CUENTA"
    ApExcel.Cells(8, 3).Formula = "FEC. VISITA"
    ApExcel.Cells(8, 4).Formula = "DIRECCION"
    ApExcel.Cells(8, 5).Formula = "DEUDOR"
    ApExcel.Cells(8, 6).Formula = "ENTREVISTADO"
    ApExcel.Cells(8, 7).Formula = "GIRO DEL NEGOCIO"
    ApExcel.Cells(8, 8).Formula = "ANALISTA"
    ApExcel.Cells(8, 9).Formula = "COMPORT. PAGO"
    
    ApExcel.Cells(8, 10).Formula = "VERIF.DEST.CREDITO"
    ApExcel.Cells(8, 11).Formula = "VERIF.GARANTIA"
    ApExcel.Cells(8, 12).Formula = "OPIN.SERV.CAJA"
    ApExcel.Cells(8, 13).Formula = "OPIN.SERV.ANALISTA"
    ApExcel.Cells(8, 14).Formula = "COMENTARIOS"
    ApExcel.Cells(8, 15).Formula = "CONCLUSION Y ACCION"
    
    ApExcel.Range("B2", "O8").Font.Bold = True
    
    ApExcel.Range("B7", "O8").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B7", "O8").HorizontalAlignment = 3

    i = 8
       
    Do While Not R.EOF
    i = i + 1
                            
        ApExcel.Cells(i, 2).Formula = R!cCtaCod
        ApExcel.Cells(i, 3).Formula = "'" & CStr(R!dFecVisita)
        ApExcel.Cells(i, 4).Formula = R!cDireVisita
        ApExcel.Cells(i, 5).Formula = R!Deudor
        ApExcel.Cells(i, 6).Formula = R!cEntrevistadoVisita
        ApExcel.Cells(i, 7).Formula = R!CIIU
        ApExcel.Cells(i, 8).Formula = R!Analista
        ApExcel.Cells(i, 9).Formula = R!nAtrasoProm
        ApExcel.Cells(i, 10).Formula = R!VerifDestino
        ApExcel.Cells(i, 11).Formula = R!cVerifGarantia
        ApExcel.Cells(i, 12).Formula = R!OpiCaja
        ApExcel.Cells(i, 13).Formula = R!OpiAna
        ApExcel.Cells(i, 14).Formula = R!cComentarios
        ApExcel.Cells(i, 15).Formula = R!cConcluAccion

        R.MoveNext
            
        Loop
            
    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
'    ApExcel.Columns("B:B").ColumnWidth = 6#
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing
    
End Sub

'*** PEAC 20101229
Private Sub ImprimeCreditosCompraDeuda(ByVal psCodAge As String, _
    ByVal psDel As String, ByVal psAl As String, ByVal psNomAge As String)

Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset

Dim sCadImp As String, i As Integer, J As Integer
Dim lsNombreArchivo As String, lcTipProd As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String
Dim lcValor As Integer
Dim lcFecDel As String, lcFecAl As String

    Screen.MousePointer = 11
    Set oDCred = New COMDCredito.DCOMCreditos
        Set R = oDCred.ObtieneCreditosCompraDeuda(psCodAge, psDel, psAl)
        
    Set oDCred = Nothing
    
    If R.EOF And R.BOF Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

    lcFecDel = Mid(psDel, 7, 2) & "/" & Mid(psDel, 5, 2) & "/" & Mid(psDel, 1, 4)
    lcFecAl = Mid(psAl, 7, 2) & "/" & Mid(psAl, 5, 2) & "/" & Mid(psAl, 1, 4)

'**************************************************************************************
    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = "CAJA MAYNAS"
    ApExcel.Cells(2, 6).Formula = Date + Time()
    ApExcel.Cells(4, 2).Formula = "CREDITOS PARA COMPRA DE DEUDA"
    ApExcel.Cells(5, 2).Formula = Trim(psNomAge) & " - " & " DEL : " & lcFecDel & " AL : " & lcFecAl
    
    ApExcel.Range("B4", "G4").MergeCells = True
    ApExcel.Range("B5", "G5").MergeCells = True
    ApExcel.Range("B4", "G5").HorizontalAlignment = xlCenter
    
    ApExcel.Cells(8, 2).Formula = "CUENTA"
    ApExcel.Cells(8, 3).Formula = "MONEDA"
    ApExcel.Cells(8, 4).Formula = "CLIENTE"
    ApExcel.Cells(8, 5).Formula = "ANALISTA"
    ApExcel.Cells(8, 6).Formula = "MONTO"
    ApExcel.Cells(8, 7).Formula = "DESCRIPCION"
    ApExcel.Cells(8, 8).Formula = "SUBPRODUCTO" '*******RECO 20131015 ERS133
       
    ApExcel.Range("B2", "H8").Font.Bold = True
    
    '*******RECO 20131015 ERS133***************
    'ApExcel.Range("B7", "H8").Interior.Color = RGB(10, 190, 160)
    'ApExcel.Range("B7", "H8").HorizontalAlignment = 3
    ApExcel.Range("B7", "H8").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B7", "H8").HorizontalAlignment = 3
    '*******END RECO***************************
    i = 8
    Do While Not R.EOF
    i = i + 1
        ApExcel.Cells(i, 2).Formula = R!cCtaCod
        ApExcel.Cells(i, 3).Formula = R!cmoneda
        ApExcel.Cells(i, 4).Formula = R!cCliente
        ApExcel.Cells(i, 5).Formula = R!cAnalista
        ApExcel.Cells(i, 6).Formula = R!nMonto
        ApExcel.Cells(i, 7).Formula = R!DesCompraDeuda
        ApExcel.Cells(i, 8).Formula = R!cSubProducto '*******RECO 20131015 ERS133
        R.MoveNext
        Loop
    R.Close
    Set R = Nothing
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Range("B2").Select
    Screen.MousePointer = 0
    ApExcel.Visible = True
    Set ApExcel = Nothing
End Sub

'MADM 20120329
Private Sub ImprimeCreditosDesembolsadosConAutorizaciones(ByVal psCodAge As String, _
    ByVal psDel As String, ByVal psAl As String, ByVal psNomAge As String, Optional ByVal pMatProd As Variant = Nothing)

Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset

Dim sCadImp As String, i As Integer, J As Integer
Dim lsNombreArchivo As String, lcTipProd As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String
Dim lcValor As Integer

'MADM 20110309
Dim vcCtaCod As String
Dim nChangeSituacion As Integer
vcCtaCod = ""
nChangeSituacion = 0
'END MADM

Dim lcFecDel As String, lcFecAl As String

    Screen.MousePointer = 11
    Set oDCred = New COMDCredito.DCOMCreditos
        Set R = oDCred.ObtieneCredDesembolsadosConAutorizaciones(psCodAge, psDel, psAl, pMatProd)
    Set oDCred = Nothing
    
    If R.EOF And R.BOF Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

    lcFecDel = Mid(psDel, 7, 2) & "/" & Mid(psDel, 5, 2) & "/" & Mid(psDel, 1, 4)
    lcFecAl = Mid(psAl, 7, 2) & "/" & Mid(psAl, 5, 2) & "/" & Mid(psAl, 1, 4)

'**************************************************************************************
    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = "CAJA MAYNAS"
    ApExcel.Cells(2, 9).Formula = Date + Time()
    ApExcel.Cells(4, 2).Formula = "CREDITOS DESEMBOLSADOS CON AUTORIZACIONES"
    ApExcel.Cells(5, 2).Formula = Trim(psNomAge) & " - " & " DEL : " & lcFecDel & " AL : " & lcFecAl
    
    ApExcel.Range("B4", "J4").MergeCells = True
    ApExcel.Range("B5", "J5").MergeCells = True
    ApExcel.Range("B4", "J5").HorizontalAlignment = xlCenter
    
    ApExcel.Cells(8, 2).Formula = "CUENTA"
    ApExcel.Cells(8, 3).Formula = "MONEDA"
    ApExcel.Cells(8, 4).Formula = "CLIENTE"
    ApExcel.Cells(8, 5).Formula = "ANALISTA"
    ApExcel.Cells(8, 6).Formula = "MONTO"
    'ApExcel.Cells(8, 7).Formula = "EXONERACIONES"
    ApExcel.Cells(8, 7).Formula = "AUTORIZACIONES" 'WIOR 20120508
    'ApExcel.Cells(8, 8).Formula = "QUIEN EXONERA"
    ApExcel.Cells(8, 8).Formula = "QUIEN AUTORIZA" 'WIOR 20120508
    'ApExcel.Cells(8, 9).Formula = "USU. EXONERA"
    ApExcel.Cells(8, 9).Formula = "USU. AUTORIZA" 'WIOR 20120508
    ApExcel.Cells(8, 10).Formula = "ESPECIFICACION"
    ApExcel.Cells(8, 11).Formula = "SUBPRODUCTO" '*******RECO 20131015 ERS133
    
    
    '*******RECO 20131015 ERS133***************************
    'ApExcel.Range("B2", "J8").Font.Bold = True
    
    'ApExcel.Range("B7", "J8").Interior.Color = RGB(10, 190, 160)
    'ApExcel.Range("B7", "J8").HorizontalAlignment = 3
    ApExcel.Range("B2", "K8").Font.Bold = True
    
    ApExcel.Range("B7", "K8").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B7", "K8").HorizontalAlignment = 3
    '************END RECO*********************************

    i = 8
       
    Do While Not R.EOF
    i = i + 1
                           
        If R!cCtaCod <> vcCtaCod Then 'madm 20110309
            nChangeSituacion = 1
        End If
        
        If nChangeSituacion = 1 Then
            ApExcel.Cells(i, 2).Formula = R!cCtaCod
            ApExcel.Cells(i, 3).Formula = R!Moneda
            ApExcel.Cells(i, 4).Formula = R!cPersNombre
            ApExcel.Cells(i, 5).Formula = R!Analista
            ApExcel.Cells(i, 6).Formula = R!nMonto
            nChangeSituacion = 0
        End If
        
        ApExcel.Cells(i, 7).Formula = R!exoneracion '*** PEAC 20101223
        ApExcel.Cells(i, 8).Formula = R!quienexonera
        ApExcel.Cells(i, 9).Formula = R!cUserExonera
        ApExcel.Cells(i, 10).Formula = R!CDescripcionOtro
        ApExcel.Cells(i, 11).Formula = R!cSubProducto '*******RECO 20131015 ERS133
        vcCtaCod = R!cCtaCod
        R.MoveNext
            
        Loop
            
    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing
    
End Sub

'MADM 20120329
Private Sub ImprimeCreditosDesembolsadosConPostDesembolso(ByVal psCodAge As String, _
    ByVal psDel As String, ByVal psAl As String, ByVal psNomAge As String, Optional ByVal pMatProd As Variant = Nothing)

Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset

Dim sCadImp As String, i As Integer, J As Integer
Dim lsNombreArchivo As String, lcTipProd As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String
Dim lcValor As Integer

'MADM 20110309
Dim vcCtaCod As String
Dim nChangeSituacion As Integer
vcCtaCod = ""
nChangeSituacion = 0
'END MADM

Dim lcFecDel As String, lcFecAl As String

    Screen.MousePointer = 11
    Set oDCred = New COMDCredito.DCOMCreditos
        Set R = oDCred.ObtieneCredDesembolsadosConPostDesembolso(psCodAge, psDel, psAl, pMatProd)
    Set oDCred = Nothing
    
    If R.EOF And R.BOF Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

    lcFecDel = Mid(psDel, 7, 2) & "/" & Mid(psDel, 5, 2) & "/" & Mid(psDel, 1, 4)
    lcFecAl = Mid(psAl, 7, 2) & "/" & Mid(psAl, 5, 2) & "/" & Mid(psAl, 1, 4)

'**************************************************************************************
    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = "CAJA MAYNAS"
    ApExcel.Cells(2, 9).Formula = Date + Time()
    ApExcel.Cells(4, 2).Formula = "CREDITOS DESEMBOLSADOS CON POST DESEMBOLSO"
    ApExcel.Cells(5, 2).Formula = Trim(psNomAge) & " - " & " DEL : " & lcFecDel & " AL : " & lcFecAl
    
    ApExcel.Range("B4", "H4").MergeCells = True
    ApExcel.Range("B5", "H5").MergeCells = True
    ApExcel.Range("B4", "H5").HorizontalAlignment = xlCenter
    
    ApExcel.Cells(8, 2).Formula = "CUENTA"
    ApExcel.Cells(8, 3).Formula = "SUBPRODUCTO" '*******RECO 20131015 ERS133
    ApExcel.Cells(8, 4).Formula = "MONEDA"
    ApExcel.Cells(8, 5).Formula = "CLIENTE"
    ApExcel.Cells(8, 6).Formula = "ANALISTA"
    ApExcel.Cells(8, 7).Formula = "MONTO"
    'ApExcel.Cells(8, 7).Formula = "REVISION"
    ApExcel.Cells(8, 8).Formula = "ESPECIFICACION"
    
    '*******RECO 20131015 ERS133*************************
    'ApExcel.Range("B2", "G8").Font.Bold = True
    
    'ApExcel.Range("B7", "G8").Interior.Color = RGB(10, 190, 160)
    'ApExcel.Range("B7", "G8").HorizontalAlignment = 3
    ApExcel.Range("B2", "H8").Font.Bold = True
    
    ApExcel.Range("B7", "H8").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B7", "H8").HorizontalAlignment = 3
    '************END RECO*******************************

   'convert(varchar(12),CC.dFecRevision,103) dFecRevision
    'mc.nmonto,mc.nmovnro,cO.CDescripcionOtro cObservacion
    
    i = 8
       
    Do While Not R.EOF
    i = i + 1
                           
        If R!cCtaCod <> vcCtaCod Then 'madm 20110309
            nChangeSituacion = 1
        End If
        
        If nChangeSituacion = 1 Then
            ApExcel.Cells(i, 2).Formula = R!cCtaCod
            ApExcel.Cells(i, 3).Formula = R!cSubProducto '*******RECO 20131015 ERS133
            ApExcel.Cells(i, 4).Formula = R!Moneda
            'ApExcel.Cells(i, 4).Formula = R!dFecRevision
            ApExcel.Cells(i, 5).Formula = R!cPersNombre
            ApExcel.Cells(i, 6).Formula = R!Analista
            ApExcel.Cells(i, 7).Formula = R!nMonto
            nChangeSituacion = 0
        End If
        
        'ApExcel.Cells(i, 7).Formula = R!dFecRevision
        ApExcel.Cells(i, 8).Formula = R!CDescripcionOtro
        vcCtaCod = R!cCtaCod
        R.MoveNext
            
        Loop
            
    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing
    
End Sub

Private Sub ImprimeCFObservaciones(ByVal psCodAge As String, _
    ByVal psDel As String, ByVal psAl As String, ByVal psNomAge As String, Optional ByVal pMatProd As Variant = Nothing)

Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset

Dim sCadImp As String, i As Integer, J As Integer
Dim lsNombreArchivo As String, lcTipProd As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String
Dim lcValor As Integer

'MADM 20110309
Dim vcCtaCod As String
Dim nChangeSituacion As Integer
vcCtaCod = ""
nChangeSituacion = 0
'END MADM

Dim lcFecDel As String, lcFecAl As String

    Screen.MousePointer = 11
    Set oDCred = New COMDCredito.DCOMCreditos
        Set R = oDCred.ObtieneCredCFObservaciones(psCodAge, psDel, psAl)
    Set oDCred = Nothing
    
    If R.EOF And R.BOF Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

    lcFecDel = Mid(psDel, 7, 2) & "/" & Mid(psDel, 5, 2) & "/" & Mid(psDel, 1, 4)
    lcFecAl = Mid(psAl, 7, 2) & "/" & Mid(psAl, 5, 2) & "/" & Mid(psAl, 1, 4)

'**************************************************************************************
    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = "CAJA MAYNAS"
    ApExcel.Cells(2, 9).Formula = Date + Time()
    ApExcel.Cells(4, 2).Formula = "CARTAS FIANZAS CON OBSERVACIONES"
    ApExcel.Cells(5, 2).Formula = Trim(psNomAge) & " - " & " DEL : " & lcFecDel & " AL : " & lcFecAl
    
    ApExcel.Range("B4", "H4").MergeCells = True
    ApExcel.Range("B5", "H5").MergeCells = True
    ApExcel.Range("B4", "H5").HorizontalAlignment = xlCenter
    
    ApExcel.Cells(8, 2).Formula = "CUENTA"
    ApExcel.Cells(8, 3).Formula = "SUBPRODUCTO" '*******RECO 20131015 ERS133
    ApExcel.Cells(8, 4).Formula = "MONEDA"
    ApExcel.Cells(8, 5).Formula = "CLIENTE"
    ApExcel.Cells(8, 6).Formula = "ANALISTA"
    ApExcel.Cells(8, 7).Formula = "MONTO"
    'ApExcel.Cells(8, 7).Formula = "REVISION"
    ApExcel.Cells(8, 8).Formula = "ESPECIFICACION"
    
    '*******RECO 20131015 ERS133*********************************
    'ApExcel.Range("B2", "G8").Font.Bold = True
    
    'ApExcel.Range("B7", "G8").Interior.Color = RGB(10, 190, 160)
    'ApExcel.Range("B7", "G8").HorizontalAlignment = 3
    
    ApExcel.Range("B2", "H8").Font.Bold = True
    
    ApExcel.Range("B7", "H8").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B7", "H8").HorizontalAlignment = 3
    '*******END RECO*********************************************

   'convert(varchar(12),CC.dFecRevision,103) dFecRevision
    'mc.nmonto,mc.nmovnro,cO.CDescripcionOtro cObservacion
    
    i = 8
       
    Do While Not R.EOF
    i = i + 1
                           
        If R!cCtaCod <> vcCtaCod Then 'madm 20110309
            nChangeSituacion = 1
        End If
        
        If nChangeSituacion = 1 Then
            ApExcel.Cells(i, 2).Formula = R!cCtaCod
            ApExcel.Cells(i, 3).Formula = R!cSubProducto '*******RECO 20131015 ERS133
            ApExcel.Cells(i, 4).Formula = R!Moneda
            'ApExcel.Cells(i, 4).Formula = R!dFecRevision
            ApExcel.Cells(i, 5).Formula = R!cPersNombre
            ApExcel.Cells(i, 6).Formula = R!Analista
            ApExcel.Cells(i, 7).Formula = R!nMontoCol
            nChangeSituacion = 0
        End If
        
        'ApExcel.Cells(i, 7).Formula = R!dFecRevision
        ApExcel.Cells(i, 7).Formula = R!cObservacion
        vcCtaCod = R!cCtaCod
        R.MoveNext
            
        Loop
            
    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing
    
End Sub

'WIOR 20140205 **********************************************
Private Sub ImprimeAsignacionSaldos(ByVal psCodAge As String, ByVal psDel As String, ByVal psAl As String, ByVal psNomAge As String, Optional ByVal pMatProd As Variant = Nothing)

Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset
Dim i As Integer

Dim lcFecDel As String, lcFecAl As String

    Screen.MousePointer = 11
    Set oDCred = New COMDCredito.DCOMCreditos
        Set R = oDCred.ObtieneAsignacionesSaldo(psCodAge, psDel, psAl)
    Set oDCred = Nothing
    
    If R.EOF And R.BOF Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

    lcFecDel = Mid(psDel, 7, 2) & "/" & Mid(psDel, 5, 2) & "/" & Mid(psDel, 1, 4)
    lcFecAl = Mid(psAl, 7, 2) & "/" & Mid(psAl, 5, 2) & "/" & Mid(psAl, 1, 4)

    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = "CAJA MAYNAS"
    ApExcel.Cells(2, 13).Formula = Date + Time()
    ApExcel.Cells(4, 2).Formula = "ASIGNACIONES DE SALDOS"
    ApExcel.Cells(5, 2).Formula = Trim(psNomAge) & " - " & " DEL : " & lcFecDel & " AL : " & lcFecAl
    
    ApExcel.Range("B4", "M4").MergeCells = True
    ApExcel.Range("B5", "M5").MergeCells = True
    ApExcel.Range("B4", "M5").HorizontalAlignment = xlCenter
    
    ApExcel.Cells(8, 2).Formula = "Fecha Solicitud"
    ApExcel.Cells(8, 3).Formula = "Fecha Autorización"
    ApExcel.Cells(8, 4).Formula = "Nº Crédito"
    ApExcel.Cells(8, 5).Formula = "Titular"
    ApExcel.Cells(8, 6).Formula = "Moneda"
    ApExcel.Cells(8, 7).Formula = "Monto"
    ApExcel.Cells(8, 8).Formula = "MontoMN"
    ApExcel.Cells(8, 9).Formula = "Total Vinculados"
    ApExcel.Cells(8, 10).Formula = "Vinculado A"
    ApExcel.Cells(8, 11).Formula = "Relación"
    ApExcel.Cells(8, 12).Formula = "Analista"
    ApExcel.Cells(8, 13).Formula = "Tipo de Producto"
    
    ApExcel.Range("B2", "M8").Font.Bold = True
    ApExcel.Range("B8", "M8").Interior.Color = RGB(200, 200, 200)
    ApExcel.Range("B8", "M8").HorizontalAlignment = 3
    

    i = 8
    ApExcel.Range(ApExcel.Cells(i, 2), ApExcel.Cells(i, 13)).Borders.LineStyle = 1
    
    Do While Not R.EOF
    i = i + 1

        ApExcel.Cells(i, 2).HorizontalAlignment = xlCenter
        ApExcel.Cells(i, 3).HorizontalAlignment = xlCenter
        ApExcel.Cells(i, 2).Formula = Format(R!fechasol, "DD/MM/YYYY")
        ApExcel.Cells(i, 3).Formula = Format(R!FechaApr, "DD/MM/YYYY")
        ApExcel.Cells(i, 4).Formula = R!cCtaCod
        ApExcel.Cells(i, 5).Formula = R!cPersNombre
        ApExcel.Cells(i, 6).Formula = R!Moneda
        ApExcel.Cells(i, 7).Formula = Format(R!nMonto, "###," & String(15, "#") & "#0.00")
        ApExcel.Cells(i, 8).Formula = Format(R!nMontoMN, "###," & String(15, "#") & "#0.00")
        ApExcel.Cells(i, 9).Formula = Format(R!TotalVinc, "###," & String(15, "#") & "#0.00")
        ApExcel.Cells(i, 10).Formula = R!Vinculado
        ApExcel.Cells(i, 11).Formula = R!Relacion
        ApExcel.Cells(i, 12).Formula = R!Analista
        ApExcel.Cells(i, 13).Formula = R!TpoProd
        ApExcel.Range(ApExcel.Cells(i, 2), ApExcel.Cells(i, 13)).Borders.LineStyle = 1

        R.MoveNext
            
        Loop
            
    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing
    
End Sub
'WIOR FIN ***************************************************
