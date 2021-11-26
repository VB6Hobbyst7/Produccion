VERSION 5.00
Begin VB.Form frmCredExonSegDesg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exoneración de Seguro Desgravamen"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11115
   Icon            =   "frmCredExonSegDesg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1200
      TabIndex        =   2
      ToolTipText     =   "Grabar Datos de Sugerencia"
      Top             =   1800
      Width           =   1200
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
      Height          =   345
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Grabar Datos de Sugerencia"
      Top             =   1800
      Width           =   840
   End
   Begin SICMACT.FlexEdit feRelaciones 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   2778
      Cols0           =   11
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "#-cPersCod-Exon-Nombre del Cliente-Relación-Edad-Glosa-Aux-nExoTitular-nOrden-nExoTitular"
      EncabezadosAnchos=   "400-0-550-3500-1500-700-4000-0-0-0-0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-2-X-X-X-6-X-X-X-X"
      ListaControles  =   "0-0-4-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-L-C-C-L-C-C-C-C"
      FormatosEdit    =   "0-0-0-5-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   3
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
   End
End
Attribute VB_Name = "frmCredExonSegDesg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&

Public fpnEstGen As Integer
Public fsCtaCod As String
Public fbGrabado As Boolean

Private Sub cmdAceptar_Click()
Dim i As Integer
Dim oCredito As COMDCredito.DCOMCredito
If ValidaDatos Then
    If MsgBox("Estas seguro de guradar los Datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set oCredito = New COMDCredito.DCOMCredito
        Call oCredito.ExoneraPersSegDesgOpe(2, fsCtaCod)
        
        For i = 1 To feRelaciones.rows - 1
            Call oCredito.ExoneraPersSegDesgOpe(1, fsCtaCod, feRelaciones.TextMatrix(i, 1), _
                IIf(Trim(feRelaciones.TextMatrix(i, 2)) = ".", True, False), Trim(feRelaciones.TextMatrix(i, 6)), _
                IIf(Trim(feRelaciones.TextMatrix(i, 8)) = "1", True, False)) 'WIOR 20130929 CAMBIÓ feRelaciones.TextMatrix(i, 5) POR feRelaciones.TextMatrix(j + 1, 6)
                'WIOR 20140825 AGREGO IIf(Trim(feRelaciones.TextMatrix(i, 8)) = ".", True, False)
        Next i
        
        MsgBox "Exoneración del Seguro Desgravamen se Realizado Correctamente.", vbInformation, "Aviso"
        fbGrabado = True
        Unload Me
        fsCtaCod = ""
    End If
End If
End Sub

Private Sub cmdCancelar_Click()
Dim oCredito As COMDCredito.DCOMCredito
Dim rsSegDes  As ADODB.Recordset
Dim rsListExoSegDesg As ADODB.Recordset 'JOEP ERS070-2016
Dim CantExo As Integer
Dim i As Integer

Set oCredito = New COMDCredito.DCOMCredito
fpnEstGen = 0
CantExo = 0

Set rsSegDes = oCredito.ListaPersonaExonerasSegDesg(fsCtaCod, "%")

''JOEP ERS070-20161130
'If gnAgenciaCredEval = 1 Then
'    Set rsListExoSegDesg = oCredito.ObtenerExoSegDesg(fsCtaCod)
'        If Not (rsListExoSegDesg.EOF And rsListExoSegDesg.BOF) Then
'
'        Else
'            MsgBox "Tiene que llenar la Exoneracion de Seguro Desgravamen", vbInformation, "Aviso"
'            Exit Sub
'        End If
'End If
''JOEP ERS070-20161130

If Not (rsSegDes.EOF And rsSegDes.BOF) Then
    For i = 1 To rsSegDes.RecordCount
        If CBool(rsSegDes!bExonera) Then
            CantExo = CantExo + 1
        End If
        rsSegDes.MoveNext
    Next i
Else
    fpnEstGen = 0
End If

If CantExo > 0 Then
    If rsSegDes.RecordCount = CantExo Then
        fpnEstGen = 1
    ElseIf (rsSegDes.RecordCount - CantExo) > 0 Then
        fpnEstGen = 2
    End If
Else
    fpnEstGen = 0
End If


Set oCredito = Nothing
Set rsSegDes = Nothing
fbGrabado = True
Unload Me
End Sub

Public Sub Inicio(ByVal psCtaCod As String, Optional ByRef pnEstado As Integer = 0)
fbGrabado = False
    fpnEstGen = 0
    fsCtaCod = psCtaCod
    Call LlenarDatos
    Me.Show 1
    pnEstado = fpnEstGen
End Sub

Private Sub LlenarDatos()
Dim oCredito As COMDCredito.DCOMCredito
Dim rsCredito As ADODB.Recordset
Dim rsSegDes As ADODB.Recordset
Dim bExo As Boolean
Dim sGlosa As String
Dim i As Integer
Dim J As Integer
'WIOR 20140825 *****
Dim bExoneraSis As Boolean
Dim nOrden As Long
Dim bExoneraTit As Boolean
'WIOR FIN **********

Set oCredito = New COMDCredito.DCOMCredito
'Set rsCredito = oCredito.RecuperaRelacPers(fsCtaCod)'WIOR 20140825 COMENTO
Set rsCredito = oCredito.RecuperaRelacPersSegDes(fsCtaCod)  'WIOR 20140825


LimpiaFlex feRelaciones

bExoneraTit = False 'WIOR 20140825
If Not (rsCredito.EOF And rsCredito.BOF) Then
    J = 0
    For i = 0 To rsCredito.RecordCount - 1
        
        'If Trim(rsCredito!nConsValor) = gColRelPersTitular Or Trim(rsCredito!nConsValor) = gColRelPersCodeudor Then'WIOR 20140825 COMENTO
            feRelaciones.AdicionaFila
            
            bExo = False
            sGlosa = ""
            'WIOR 20140825 ****
            bExoneraSis = False
            nOrden = 0
            'WIOR FIN ********
            Set rsSegDes = oCredito.ListaPersonaExonerasSegDesg(fsCtaCod, Trim(rsCredito!cPersCod))
            
            If Not (rsSegDes.EOF And rsSegDes.BOF) Then
                bExo = rsSegDes!bExonera
                sGlosa = Trim(rsSegDes!cGlosa)
                'WIOR 20140825 *********************
                bExoneraSis = CBool(rsSegDes!bExcluidoSis)
                nOrden = CLng(rsSegDes!Orden)
                
                If nOrden = 1 And bExo Then
                    bExoneraTit = True
                End If
                'WIOR FIN *************************
            End If
            
            feRelaciones.TextMatrix(J + 1, 0) = J + 1
            feRelaciones.TextMatrix(J + 1, 1) = rsCredito!cPersCod
            feRelaciones.TextMatrix(J + 1, 2) = IIf(bExo, 1, 0)
            feRelaciones.TextMatrix(J + 1, 3) = Trim(rsCredito!cPersNombre)
            feRelaciones.TextMatrix(J + 1, 4) = Trim(rsCredito!Relacion) 'WIOR 20140825 COMBIO cConsDescripcion POR Relacion
            feRelaciones.TextMatrix(J + 1, 5) = EdadPersona(rsCredito!dPersNacCreac, gdFecSis) 'DateDiff("yyyy", rsCredito!dPersNacCreac, gdFecSis) 'WIOR 20130829
            feRelaciones.TextMatrix(J + 1, 6) = Trim(sGlosa) 'WIOR 20130829 CAMBIÓ feRelaciones.TextMatrix(j + 1, 5) POR feRelaciones.TextMatrix(j + 1, 6)
            feRelaciones.TextMatrix(J + 1, 7) = Trim(rsCredito!nPersPersoneria) 'WIOR 20130829
            'WIOR 20130829 **********************
            feRelaciones.TextMatrix(J + 1, 8) = IIf(bExoneraSis, 1, 0)
            feRelaciones.TextMatrix(J + 1, 9) = nOrden
            feRelaciones.TextMatrix(J + 1, 10) = IIf(bExoneraTit, 1, 0)
            'WIOR FIN ***************************
            J = J + 1
        'End If'WIOR 20140825 COMENTO
        rsCredito.MoveNext
    Next
End If

Set rsCredito = Nothing
Set rsSegDes = Nothing
End Sub


Private Function ValidaDatos() As Boolean
Dim i As Integer
Dim J As Integer
Dim CantExo As Integer
Dim Items As String
Items = ""
CantExo = 0
fpnEstGen = 0
J = 0
For i = 1 To feRelaciones.rows - 1
    If Trim(feRelaciones.TextMatrix(i, 2)) = "." Then
         If Trim(feRelaciones.TextMatrix(i, 6)) = "" Then 'WIOR 20130829 CAMBIÓ feRelaciones.TextMatrix(i, 5) POR feRelaciones.TextMatrix(i, 6)
            Items = Items & (i) & ","
         End If
         CantExo = CantExo + 1
'JOEP 20171007
    Else
        J = J + 1
'JOEP 20171007
    End If
Next i

'JOEP 20171007
'If (feRelaciones.rows - 1) = J Then
'    MsgBox "Seleccione un Items!! ", vbInformation, "Aviso"
'    ValidaDatos = False
'    Exit Function
'End If
'JOEP 20171007

If Items <> "" Then
    Items = Mid(Items, 1, Len(Items) - 1)
    
    MsgBox "Ingrese la glosa de los Items Nº " & Items, vbInformation, "Aviso"
    ValidaDatos = False
    Exit Function
End If

If CantExo > 0 Then
    If (feRelaciones.rows - 1) = CantExo Then
        fpnEstGen = 1
    ElseIf ((feRelaciones.rows - 1) - CantExo) > 0 Then
        fpnEstGen = 2
    End If
Else
    fpnEstGen = 0
End If


ValidaDatos = True
End Function

'WIOR 20130829 ************************************************************************************
Private Sub feRelaciones_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
On Error GoTo ErrorEdad
Dim oPar As COMDCredito.DCOMParametro
Dim nEdadMax As Double
Dim i As Integer 'WIOR 20140825
Set oPar = New COMDCredito.DCOMParametro

nEdadMax = oPar.RecuperaValorParametro(3204)
If CLng(feRelaciones.TextMatrix(pnRow, 5)) >= nEdadMax Then
    feRelaciones.TextMatrix(pnRow, 2) = "1"
    Exit Sub 'WIOR 20140825
End If

'WIOR 20140825 *********************************
If Trim(feRelaciones.TextMatrix(pnRow, 8)) = "1" Then
    MsgBox "No se pueder quitar a la persona porque estar exonerada por el sistema.", vbInformation, "Aviso"
    feRelaciones.TextMatrix(pnRow, 2) = "1"
    Exit Sub
End If

If Trim(feRelaciones.TextMatrix(pnRow, 9)) = "1" Then
    If Trim(feRelaciones.TextMatrix(pnRow, 2)) = "." Then
        MsgBox "No se puede exonerar al Titular del Seguro." & vbNewLine & "Por favor Coordinar con el Área de Créditos.", vbInformation, "Aviso"
        'Comento JOEP-20171007
        'If MsgBox("Estas exonerando al Titular del Seguro por tanto se exonerara el seguro por completo." & vbNewLine & "¿Desea continuar?", vbInformation + vbYesNo, "Aviso") = vbYes Then
         '   For i = 1 To feRelaciones.rows - 1
          '      feRelaciones.TextMatrix(i, 10) = "1"
           '     feRelaciones.TextMatrix(i, 2) = "1"
           ' Next i
        'Else
            feRelaciones.TextMatrix(pnRow, 2) = "0"
        'End If
        'Comento JOEP-20171007
    Else
        For i = 1 To feRelaciones.rows - 1
            feRelaciones.TextMatrix(i, 10) = "0"
        Next i
    End If
    Exit Sub
Else
If Trim(feRelaciones.TextMatrix(pnRow, 2)) = "." Then
    If Trim(feRelaciones.TextMatrix(pnRow, 4)) = "TITULAR" Then
        MsgBox "No se puede exonerar al Titular del Seguro." & vbNewLine & "Por favor Coordinar con el Área de Créditos.", vbInformation, "Aviso"
        feRelaciones.TextMatrix(pnRow, 2) = "0"
    End If
Else
        For i = 1 To feRelaciones.rows - 1
            feRelaciones.TextMatrix(i, 10) = "0"
        Next i
End If
    Exit Sub
End If

If Trim(feRelaciones.TextMatrix(pnRow, 10)) = "1" Then
    MsgBox "No se pueder quitar a la persona, ya que el titular del seguro esta exonerado.", vbInformation, "Aviso"
    feRelaciones.TextMatrix(pnRow, 2) = "1"
    Exit Sub
End If
'WIOR FIN **************************************

Set oPar = Nothing
Exit Sub
ErrorEdad:
feRelaciones.TextMatrix(pnRow, 2) = "1"
MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub feRelaciones_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim sColumnas() As String
sColumnas = Split(feRelaciones.ColumnasAEditar, "-")
If sColumnas(pnCol) = "X" Or Trim(feRelaciones.TextMatrix(pnRow, 8)) = "1" Then
   Cancel = False
   SendKeys "{Tab}", True
   Exit Sub
End If
End Sub

'WIOR FIN *****************************************************************************************
Private Sub Form_Unload(Cancel As Integer)
If Not fbGrabado Then
    Call cmdCancelar_Click
End If
End Sub

Private Sub Form_Load()
DisableCloseButton Me
End Sub

Public Function DisableCloseButton(frm As Form) As Boolean
'PURPOSE: Removes X button from a form
'EXAMPLE: DisableCloseButton Me
'RETURNS: True if successful, false otherwise
'NOTES:   Also removes Exit Item from
'         Control Box Menu
    Dim lHndSysMenu As Long
    Dim lAns1 As Long, lAns2 As Long
    
    lHndSysMenu = GetSystemMenu(frm.hwnd, 0)

    'remove close button
    lAns1 = RemoveMenu(lHndSysMenu, 6, MF_BYPOSITION)

   'Remove seperator bar
    lAns2 = RemoveMenu(lHndSysMenu, 5, MF_BYPOSITION)
    
    'Return True if both calls were successful
    DisableCloseButton = (lAns1 <> 0 And lAns2 <> 0)

End Function
