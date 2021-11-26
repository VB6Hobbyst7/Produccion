VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAuditListarAcceso 
   Caption         =   "RELACIÓN DE ACCESO Y PERFILES DE LOS USUARIO DE LA CMAC-MAYNAS S.A."
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11970
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAuditListarAcceso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   6975
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   11655
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgUsuario 
         Height          =   6255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   11033
         _Version        =   393216
         Cols            =   5
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
      Begin VB.Label lblRecord 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8760
         TabIndex        =   9
         Top             =   6600
         Width           =   2775
      End
      Begin VB.Label lblProceso 
         Caption         =   "Procesando ... Por Favor Espere un Momento !!! Gracias ...!!!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   6600
         Visible         =   0   'False
         Width           =   6375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros de Búsqueda"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.CommandButton Command2 
         Caption         =   "Excel"
         Height          =   375
         Left            =   10200
         TabIndex        =   10
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSDataListLib.DataCombo dcArea 
         Height          =   315
         Left            =   6480
         TabIndex        =   7
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcAgencia 
         Height          =   315
         Left            =   2280
         TabIndex        =   6
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   10200
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Área:"
         Height          =   255
         Left            =   5880
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Agencia"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmAuditListarAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmAuditListarAcceso
'** Descripción : Formulario para visualizar la Relacion de Accesos y Perfiles del Sistema .
'** Creación : MAVM, 20081010 8:58:15 AM
'** Modificación:
'********************************************************************

Option Explicit
Dim oAcceso As COMDPersona.UCOMAcceso

Private Sub Form_Load()
    lblProceso.Visible = False
    Set oAcceso = New COMDPersona.UCOMAcceso
    CargarAgencias
    CargarAreas
    CargarCabecera
End Sub

Private Sub CargarAgencias()
    Dim rsAgencia As New ADODB.Recordset
    Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Set rsAgencia.DataSource = objCOMNAuditoria.DarAgencias
    dcAgencia.BoundColumn = "AgeCod"
    dcAgencia.DataField = "AgeCod"
    Set dcAgencia.RowSource = rsAgencia
    dcAgencia.ListField = "cAgeDescripcion"
    dcAgencia.BoundText = 0
End Sub

Private Sub CargarAreas()
    Dim rsArea As New ADODB.Recordset
    Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Set rsArea.DataSource = objCOMNAuditoria.DarAreas
    dcArea.BoundColumn = "cAreaCod"
    dcArea.DataField = "cAreaCod"
    Set dcArea.RowSource = rsArea
    dcArea.ListField = "cAreaDescripcion"
    dcArea.BoundText = 0
End Sub

Private Sub CargarCabecera()
    fgUsuario.FormatString = " | Usuario | Nombre(s)/Apellido(s) | Agencia | Area | Grupo"
    fgUsuario.ColWidth(0) = 350
    fgUsuario.ColWidth(1) = 800
    fgUsuario.ColWidth(2) = 3000
    fgUsuario.ColWidth(3) = 1600
    fgUsuario.ColWidth(4) = 2800
    fgUsuario.ColWidth(5) = 2500
End Sub

Private Sub Command1_Click()
    Command2.Visible = False
    lblProceso.Visible = True
    lblRecord.Visible = False
    Limpiar
    LimpiarControles Me, False, True
    CargaControles
    lblProceso.Visible = False
    lblRecord.Visible = True
    Command2.Visible = True
End Sub

Private Sub Limpiar()
    fgUsuario.TextMatrix(1, 1) = ""
    fgUsuario.TextMatrix(1, 2) = ""
    fgUsuario.TextMatrix(1, 3) = ""
    fgUsuario.TextMatrix(1, 4) = ""
    fgUsuario.TextMatrix(1, 5) = ""
End Sub

Public Sub LimpiarControles(frmForm As Form, Optional cText As Boolean, Optional cFG As Boolean, Optional cCombo As Boolean, Optional cDataCombo As Boolean, Optional cLabel As Boolean)
      Dim ctlControl As Object
      On Error Resume Next
      For Each ctlControl In frmForm.Controls
        If cText = True Then
            If TypeOf ctlControl Is TextBox Then
                ctlControl.Text = ""
            End If
        End If
        If cFG = True Then
            If TypeOf ctlControl Is MSHFlexGrid Then
                ctlControl.TextMatrix = ""
                ctlControl.Rows = 2
            End If
        End If
        If cCombo = True Then
            If TypeOf ctlControl Is ComboBox Then
                ctlControl.ListIndex = 0
            End If
        End If
        If cDataCombo = True Then
            If TypeOf ctlControl Is DataCombo Then
                ctlControl.Text = ""
            End If
        End If
        If cLabel = True Then
            If TypeOf ctlControl Is Label Then
                ctlControl.Caption = ""
                ctlControl.BorderStyle = 0
            End If
        End If
         DoEvents
      Next ctlControl
End Sub

Private Sub CargaControles()
    Dim rsOpe As ADODB.Recordset
    Dim RsMenu As ADODB.Recordset
    Dim MatUser() As String
    Dim MatGrupo() As String
    Call oAcceso.CargaControles(gsDominio, MatUser, MatGrupo, RsMenu, rsOpe, gsCodUser)
    CargarDatos MatUser
End Sub

Private Sub CargarDatos(ByVal MatUser As Variant)
    Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Dim lsmensaje As String
    Dim MatGrilla() As String
    Dim rs As ADODB.Recordset
    Dim i, J As Integer
    Dim lsGrupo As String
    Dim lsOperaciones As String
    Dim lsColocaciones As String
    Dim lsOtros As String
    Dim liPosicion As Integer
    Dim lsCadena As String
    Dim liValor As Integer
    Dim contador As Integer
    Dim liStar As Integer
    Dim liFin As Integer
    Dim liGuarda As Integer
    
    ReDim MatGrilla(UBound(MatUser), 1 To 8)
    
    For i = 0 To UBound(MatUser) - 1
        'lsGrupo = oAcceso.CargaUsuarioGrupoAuditoria(MatUser(i, 0), gsDominio)
        lsGrupo = oAcceso.CargaUsuarioGrupo(MatUser(i, 0), gsDominio)
        
        If Len(lsGrupo) <> 0 Then lsGrupo = "'" & lsGrupo & "'"
        liValor = Len(Replace(lsGrupo, "'", " "))
        
        If liValor > 0 Then
            lsOperaciones = DevolverGrupos(lsGrupo)
            lsColocaciones = DevolverGruposColocaciones(lsGrupo)
            lsOtros = DevolverGruposOtros(lsGrupo)
                                        
            lsmensaje = ""
            Set rs = objCOMNAuditoria.DarDatosUsuarioXUser(MatUser(i, 0), dcAgencia.BoundText, dcArea.BoundText, lsmensaje)
                If lsmensaje = "" Then
                    Do While Not rs.EOF
                        MatGrilla(i, 1) = rs("cPersNombre")
                        MatGrilla(i, 2) = rs("cUser")
                        MatGrilla(i, 3) = rs("cAgeDescripcion")
                        MatGrilla(i, 4) = rs("cAreaDescripcion")
                        MatGrilla(i, 5) = lsGrupo
                        MatGrilla(i, 6) = lsOperaciones
                        MatGrilla(i, 7) = lsColocaciones
                        MatGrilla(i, 8) = lsOtros
                        rs.MoveNext
                        Loop
                Else
                    MatGrilla(i, 1) = ""
                    MatGrilla(i, 2) = MatUser(i, 0)
                    MatGrilla(i, 3) = ""
                    MatGrilla(i, 4) = ""
                    MatGrilla(i, 5) = lsGrupo
                    MatGrilla(i, 6) = IIf(lsmensaje = "", "", lsOperaciones)
                    MatGrilla(i, 7) = IIf(lsmensaje = "", "", lsColocaciones)
                    MatGrilla(i, 8) = IIf(lsmensaje = "", "", lsOtros)
                End If
        Else
            lsmensaje = ""
            Set rs = objCOMNAuditoria.DarDatosUsuarioXUser(MatUser(i, 0), dcAgencia.BoundText, dcArea.BoundText, lsmensaje)
            If lsmensaje = "" Then
                Do While Not rs.EOF
                    MatGrilla(i, 1) = rs("cPersNombre")
                    MatGrilla(i, 2) = rs("cUser")
                    MatGrilla(i, 3) = rs("cAgeDescripcion")
                    MatGrilla(i, 4) = rs("cAreaDescripcion")
                    MatGrilla(i, 5) = lsGrupo
                    MatGrilla(i, 6) = lsOperaciones
                    MatGrilla(i, 7) = lsColocaciones
                    MatGrilla(i, 8) = lsOtros
                    rs.MoveNext
                Loop
            Else
                MatGrilla(i, 1) = ""
                MatGrilla(i, 2) = MatUser(i, 0)
                MatGrilla(i, 3) = ""
                MatGrilla(i, 4) = ""
                MatGrilla(i, 5) = lsGrupo
                MatGrilla(i, 6) = IIf(lsmensaje = "", "", lsOperaciones)
                MatGrilla(i, 7) = IIf(lsmensaje = "", "", lsColocaciones)
                MatGrilla(i, 8) = IIf(lsmensaje = "", "", lsOtros)
            End If
        End If
    Next i
    
    Call CargaGrilla(MatGrilla)
    Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
MostrarRelacionUsuarioCMACMAYNASExcel
End Sub

Public Sub MostrarRelacionUsuarioCMACMAYNASExcel()
    Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
    Dim R As ADODB.Recordset
    Dim lMatCabecera As Variant
    Dim lsNombreArchivo As String
    lsNombreArchivo = "RelacionUsuarioCMACMAYNAS"
    ReDim lMatCabecera(8, 2)
    lMatCabecera(0, 0) = "Nombre(s)/Apellido(s)": lMatCabecera(0, 1) = ""
    lMatCabecera(1, 0) = "Usuario": lMatCabecera(1, 1) = ""
    lMatCabecera(2, 0) = "Agencia": lMatCabecera(2, 1) = ""
    lMatCabecera(3, 0) = "Area": lMatCabecera(3, 1) = ""
    lMatCabecera(4, 0) = "Grupo": lMatCabecera(4, 1) = ""
    lMatCabecera(5, 0) = "Operaciones - Captaciones": lMatCabecera(5, 1) = ""
    lMatCabecera(6, 0) = "Operaciones - Colocaciones": lMatCabecera(6, 1) = ""
    lMatCabecera(7, 0) = "Operaciones - Otros": lMatCabecera(7, 1) = ""
    
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Set R = objCOMNAuditoria.DarUsuarioCMACMAYNASExcel
    Set objCOMNAuditoria = Nothing
    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Relación de Usuarios CMACMAYNAS S.A.", "", lsNombreArchivo, lMatCabecera, R, 2, , , True)
End Sub

Public Sub GeneraReporteEnArchivoExcel(ByVal psNomCmac As String, ByVal psNomAge As String, ByVal psCodUser As String, ByVal pdFecSis As Date, ByVal psTitulo As String, ByVal psSubTitulo As String, _
                                    ByVal psNomArchivo As String, ByVal pMatCabeceras As Variant, ByVal prRegistros As ADODB.Recordset, _
                                    Optional pnNumDecimales As Integer, Optional Visible As Boolean = False, Optional psNomHoja As String = "", _
                                    Optional pbSinFormatDeReg As Boolean = False, _
                                    Optional pbUsarCabecerasDeRS As Boolean = False)
    Dim rs As ADODB.Recordset
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim liLineas As Integer, i As Integer
    Dim fs As Scripting.FileSystemObject
    Dim lnNumColumns As Integer


    If Not (prRegistros.EOF And prRegistros.BOF) Then
        If pbUsarCabecerasDeRS = True Then
            lnNumColumns = prRegistros.Fields.Count
        Else
            lnNumColumns = UBound(pMatCabeceras)
            lnNumColumns = IIf(prRegistros.Fields.Count < lnNumColumns, prRegistros.Fields.Count, prRegistros.Fields.Count)
        End If

        If psNomHoja = "" Then psNomHoja = psNomArchivo
        psNomArchivo = psNomArchivo & "_" & psCodUser & ".xls"

        Set fs = New Scripting.FileSystemObject
        Set xlAplicacion = New Excel.Application
        If fs.FileExists(App.path & "\Spooler\" & psNomArchivo) Then
            fs.DeleteFile (App.path & "\Spooler\" & psNomArchivo)
        End If
        Set xlLibro = xlAplicacion.Workbooks.Add
        Set xlHoja1 = xlLibro.Worksheets.Add

        xlHoja1.Name = psNomHoja
        xlHoja1.Cells.Select
        
        'Cabeceras
        xlHoja1.Cells(1, 1) = psNomCmac
        xlHoja1.Cells(1, lnNumColumns) = Trim(Format(pdFecSis, "dd/mm/yyyy hh:mm:ss"))
        xlHoja1.Cells(2, 1) = psNomAge
        xlHoja1.Cells(2, lnNumColumns) = psCodUser
        xlHoja1.Cells(4, 1) = psTitulo
        xlHoja1.Cells(5, 1) = psSubTitulo
        xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, lnNumColumns)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, lnNumColumns)).Merge True
        xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, lnNumColumns)).Merge True
        xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(5, lnNumColumns)).HorizontalAlignment = xlCenter

        liLineas = 6
        If pbUsarCabecerasDeRS = True Then
            For i = 0 To prRegistros.Fields.Count - 1
                xlHoja1.Cells(liLineas, i + 1) = prRegistros.Fields(i).Name
            Next i
        Else
            For i = 0 To lnNumColumns - 1
                If (i + 1) > UBound(pMatCabeceras) Then
                    xlHoja1.Cells(liLineas, i + 1) = prRegistros.Fields(i).Name
                Else
                    xlHoja1.Cells(liLineas, i + 1) = pMatCabeceras(i, 0)
                End If
            Next i
        End If
        
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, lnNumColumns)).Cells.Interior.Color = RGB(220, 220, 220)
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, lnNumColumns)).HorizontalAlignment = xlCenter

        If pbSinFormatDeReg = False Then
            liLineas = liLineas + 1
            While Not prRegistros.EOF
                For i = 0 To lnNumColumns - 1
                    If pMatCabeceras(i, 1) = "" Then  'Verificamos si tiene tipo
                        xlHoja1.Cells(liLineas, i + 1) = prRegistros(i)
                    Else
                        Select Case pMatCabeceras(i, 1)
                            Case "S"
                                xlHoja1.Cells(liLineas, i + 1) = prRegistros(i)
                            Case "N"
                                xlHoja1.Cells(liLineas, i + 1) = Format(prRegistros(i), "#0.00")
                            Case "D"
                                xlHoja1.Cells(liLineas, i + 1) = IIf(Format(prRegistros(i), "yyyymmdd") = "19000101", "", Format(prRegistros(i), "dd/mm/yyyy"))
                        End Select
                    End If
                Next i
                liLineas = liLineas + 1
                prRegistros.MoveNext
            Wend
        Else
            xlHoja1.Range("A7").CopyFromRecordset prRegistros 'Copia el contenido del recordset a excel
        End If

        xlHoja1.SaveAs App.path & "\Spooler\" & psNomArchivo
        MsgBox "Se ha generado el Archivo en " & App.path & "\Spooler\" & psNomArchivo

        If Visible Then
            xlAplicacion.Visible = True
            xlAplicacion.Windows(1).Visible = True
        End If

            xlLibro.Close
            xlAplicacion.Quit
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing

    End If
End Sub

    Private Function DevolverGrupos(ByVal lsGrupo As String) As String
        Dim oConn As COMConecta.DCOMConecta
        Dim rsOpe1 As New ADODB.Recordset
        Dim sSql As String
        Dim J As Integer
        Dim lsOperaciones As String
                
        If lsGrupo <> "" Then
            Set oConn = New COMConecta.DCOMConecta
            If oConn.AbreConexion() Then
                sSql = "select distinct(P.cName), M.cName, M.cAplicacion, Replace(M.cDescrip, '" & "&" & "', '" & "" & "')as cDescrip, M.nOrden from Menu M inner join Permiso P on M.cName=P.cName where M.cAplicacion='" & "1" & "' and M.cDescrip<>'" & "-" & "' and (M.cName > '" & "M020000000000" & "' and M.cName <'" & "M030000000000" & "') and P.cGrupoUsu in (" & lsGrupo & ") order by nOrden"
                Set rsOpe1 = oConn.CargaRecordSet(sSql)
                oConn.CierraConexion
            End If
            Set oConn = Nothing
            
            If rsOpe1.RecordCount <> "0" Then
                For J = 0 To rsOpe1.RecordCount - 1
                    If Mid(rsOpe1("cName"), 5, 1) = "1" And Mid(rsOpe1("cName"), 6, 1) = "0" And Mid(rsOpe1("cName"), 7, 1) = "0" And Mid(rsOpe1("cName"), 8, 1) = "0" And Mid(rsOpe1("cName"), 9, 1) = "0" And Mid(rsOpe1("cName"), 10, 1) = "0" And Mid(rsOpe1("cName"), 11, 1) = "0" Then
                        If Mid(rsOpe1("cName"), 13, 1) = "0" And Mid(rsOpe1("cName"), 12, 1) = "0" Then
                            lsOperaciones = Mid(rsOpe1("cDescrip"), 1, 23)
                        Else
                            lsOperaciones = lsOperaciones + ";" + Mid(rsOpe1("cDescrip"), 1, 23)
                        End If
                    Else
                        If lsOperaciones = "" Then
                            lsOperaciones = Mid(rsOpe1("cDescrip"), 1, 23)
                        Else
                            lsOperaciones = lsOperaciones + ":" + Mid(rsOpe1("cDescrip"), 1, 23)
                        End If
                    End If
                    rsOpe1.MoveNext
                Next J
            End If
        End If
        Set rsOpe1 = Nothing
        DevolverGrupos = lsOperaciones
    End Function
    
    Private Function DevolverGruposColocaciones(ByVal lsGrupo As String) As String
        Dim oConn As COMConecta.DCOMConecta
        Dim sSql As String
        Dim rsOpe1 As New ADODB.Recordset
        Dim J As Integer
        Dim lsOperaciones As String
                
        If lsGrupo <> "" Then
            Set oConn = New COMConecta.DCOMConecta
            If oConn.AbreConexion() Then
                sSql = "select distinct(P.cName), M.cName, M.cAplicacion, Replace(M.cDescrip, '" & "&" & "', '" & "" & "')as cDescrip, M.nOrden from Menu M inner join Permiso P on M.cName=P.cName where M.cAplicacion='" & "1" & "' and M.cDescrip<>'" & "-" & "' and (nOrden between '" & "73" & "' and '" & "230" & "') and P.cGrupoUsu in (" & lsGrupo & ") order by nOrden"
                Set rsOpe1 = oConn.CargaRecordSet(sSql)
                oConn.CierraConexion
            End If
            Set oConn = Nothing
            
            If rsOpe1.RecordCount <> "0" Then
                For J = 0 To rsOpe1.RecordCount - 1
                    If Mid(rsOpe1("cName"), 5, 1) = "1" And Mid(rsOpe1("cName"), 6, 1) = "0" And Mid(rsOpe1("cName"), 7, 1) = "0" And Mid(rsOpe1("cName"), 8, 1) = "0" And Mid(rsOpe1("cName"), 9, 1) = "0" And Mid(rsOpe1("cName"), 10, 1) = "0" And Mid(rsOpe1("cName"), 11, 1) = "0" Then
                        If Mid(rsOpe1("cName"), 13, 1) = "0" And Mid(rsOpe1("cName"), 12, 1) = "0" Then
                            lsOperaciones = Mid(rsOpe1("cDescrip"), 1, 4)
                        Else
                            lsOperaciones = lsOperaciones + ";" + Mid(rsOpe1("cDescrip"), 1, 4)
                        End If
                    Else
                        If lsOperaciones = "" Then
                            lsOperaciones = Mid(rsOpe1("cDescrip"), 1, 4)
                        Else
                            lsOperaciones = lsOperaciones + ":" + Mid(rsOpe1("cDescrip"), 1, 4)
                        End If
                    End If
                    rsOpe1.MoveNext
                Next J
            End If
        
        End If
        Set rsOpe1 = Nothing
        DevolverGruposColocaciones = lsOperaciones
    End Function
    
    Private Function DevolverGruposOtros(ByVal lsGrupo As String) As String
        Dim oConn As COMConecta.DCOMConecta
        Dim rsOpe1 As New ADODB.Recordset
        Dim sSql As String
        Dim J As Integer
        Dim lsOperaciones As String
       
        If lsGrupo <> "" Then
            Set oConn = New COMConecta.DCOMConecta
            If oConn.AbreConexion() Then
                sSql = "select distinct(P.cName), M.cName, M.cAplicacion, Replace(M.cDescrip, '" & "&" & "', '" & "" & "')as cDescrip, M.nOrden from Menu M inner join Permiso P on M.cName=P.cName where M.cAplicacion='" & "1" & "' and M.cDescrip<>'" & "-" & "' and M.cName >= '" & "M040000000000" & "' and P.cGrupoUsu in (" & lsGrupo & ") order by nOrden"
                Set rsOpe1 = oConn.CargaRecordSet(sSql)
                oConn.CierraConexion
            End If
            Set oConn = Nothing
                If rsOpe1.RecordCount <> "0" Then
                    For J = 0 To rsOpe1.RecordCount - 1
                        If Mid(rsOpe1("cName"), 5, 1) = "1" And Mid(rsOpe1("cName"), 6, 1) = "0" And Mid(rsOpe1("cName"), 7, 1) = "0" And Mid(rsOpe1("cName"), 8, 1) = "0" And Mid(rsOpe1("cName"), 9, 1) = "0" And Mid(rsOpe1("cName"), 10, 1) = "0" And Mid(rsOpe1("cName"), 11, 1) = "0" Then
                            If Mid(rsOpe1("cName"), 13, 1) = "0" And Mid(rsOpe1("cName"), 12, 1) = "0" Then
                                lsOperaciones = Mid(rsOpe1("cDescrip"), 1, 22)
                            Else
                                lsOperaciones = lsOperaciones + "; " + Mid(rsOpe1("cDescrip"), 1, 22)
                            End If
                        Else
                            If lsOperaciones = "" Then
                                lsOperaciones = Mid(rsOpe1("cDescrip"), 1, 22)
                            Else
                                lsOperaciones = lsOperaciones + ":" + Mid(rsOpe1("cDescrip"), 1, 22)
                            End If
                        End If
                    rsOpe1.MoveNext
                    Next J
                    
                End If
                
        End If
        
        Set rsOpe1 = Nothing
        DevolverGruposOtros = lsOperaciones
    End Function

Private Sub CargaGrilla(ByVal MatGrilla As Variant)
    Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Dim i As Integer
    Dim Valor As Integer
    objCOMNAuditoria.BorrarUsuarioCMACMAYNASTem
    Me.fgUsuario.Rows = 2
    For i = 1 To UBound(MatGrilla) - 1
        If MatGrilla(i - 1, 1) <> "" Then
            If fgUsuario.TextMatrix(fgUsuario.Rows - 1, 1) <> "" Then Me.fgUsuario.Rows = Me.fgUsuario.Rows + 1
            fgUsuario.TextMatrix(fgUsuario.Rows - 1, 1) = MatGrilla(i - 1, 2)
            fgUsuario.TextMatrix(fgUsuario.Rows - 1, 2) = MatGrilla(i - 1, 1)
            fgUsuario.TextMatrix(fgUsuario.Rows - 1, 3) = MatGrilla(i - 1, 3)
            fgUsuario.TextMatrix(fgUsuario.Rows - 1, 4) = MatGrilla(i - 1, 4)
            fgUsuario.TextMatrix(fgUsuario.Rows - 1, 5) = Replace(MatGrilla(i - 1, 5), "'", "")
            objCOMNAuditoria.InsertarUsuarioCMACMAYNASTem MatGrilla(i - 1, 1), MatGrilla(i - 1, 2), MatGrilla(i - 1, 3), MatGrilla(i - 1, 4), Replace(Replace(MatGrilla(i - 1, 5), ",", "; "), "'", ""), Replace(MatGrilla(i - 1, 6), ",", ";"), Replace(MatGrilla(i - 1, 7), ",", ";"), Replace(MatGrilla(i - 1, 8), ",", ";")
            Valor = Valor + 1
        End If
    Next i
    lblRecord.Caption = "Total de Records" & " " & Valor
End Sub
