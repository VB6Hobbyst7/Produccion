Attribute VB_Name = "ModDescObjeto"
Option Explicit
Public Sub AsignaImgNodo(nodX As Node, Optional plExpand As Boolean = False)

   nodX.Image = "cerrado"
   nodX.ExpandedImage = "abierto"
   If plExpand Then
      nodX.ForeColor = "&H8000000D"
      nodX.Expanded = True
   End If

End Sub
Public Sub CargaNodo(psRaiz As String, tvw As TreeView, rsVista As ADODB.Recordset, pnObjNiv As Integer, pnColCod As Long, pnColDesc As Long, Optional plExpand As Boolean = False)
Dim sCod As String, siSale As Boolean
Dim SiInstancia As Boolean
Dim nodX As Node
Dim pnOk As Integer
Dim nObjNiv As Integer
Dim nNivel As Integer
Dim siGrabo As Boolean
siGrabo = True
siSale = False
Do While Not rsVista.EOF
   If rsVista(2) > pnObjNiv Then
      nNivel = nNivel + 1
      AdicionaNodo rsVista(pnColCod), rsVista(pnColDesc), rsVista(2), tvw, psRaiz, 4, plExpand
      siGrabo = True
      nObjNiv = rsVista(2)
      sCod = rsVista(0)
      rsVista.MoveNext
      CargaNodo sCod, tvw, rsVista, nObjNiv, pnColCod, pnColDesc, plExpand
      nNivel = nNivel - 1
      If Not siGrabo Then
         If rsVista(2) = pnObjNiv Then
            AdicionaNodo rsVista(pnColCod), rsVista(pnColDesc), rsVista(2), tvw, psRaiz, 1, plExpand
            siGrabo = True
            psRaiz = rsVista(pnColCod)
            rsVista.MoveNext
         End If
      End If
   Else
      If rsVista(2) = pnObjNiv Then
         AdicionaNodo rsVista(pnColCod), rsVista(pnColDesc), rsVista(2), tvw, psRaiz, 1, plExpand
         psRaiz = rsVista(pnColCod)
         siGrabo = True
         rsVista.MoveNext
      Else
         If rsVista(2) < pnObjNiv Then
            siGrabo = False
            Exit Sub
         End If
      End If
   End If
Loop
End Sub
Public Sub AdicionaNodo(sCod As String, sDes As String, pnObjNiv As Integer, tvwObjeto As TreeView, psRaiz As String, nTipo As Integer, Optional plExpand As Boolean = False)
Dim nodX As Node
   Set nodX = tvwObjeto.Nodes.Add("K" & psRaiz, nTipo)
   nodX.Key = "K" & sCod
   nodX.Text = sCod & " - " & sDes
   AsignaImgNodo nodX, plExpand
   nodX.Tag = CStr(pnObjNiv)
End Sub

'Public Function Letras(intTecla As Integer, Optional lbMayusculas As Boolean = True) As Integer
'If lbMayusculas Then
'    Letras = Asc(UCase(Chr(intTecla)))
'Else
'    Letras = Asc(LCase(Chr(intTecla)))
'End If
'End Function
'Public Function NumerosEnteros(intTecla As Integer, Optional pbNegativos As Boolean = False) As Integer
'Dim cValidar As String
'    If pbNegativos = False Then
'        cValidar = "0123456789"
'    Else
'        cValidar = "0123456789-"
'    End If
'    If intTecla > 27 Then
'        If InStr(cValidar, Chr(intTecla)) = 0 Then
'            intTecla = 0
'            Beep
'        End If
'    End If
'    NumerosEnteros = intTecla
'End Function
Public Function BuscaDato(ByVal Criterio As String, rsAdo As ADODB.Recordset, ByVal start As Long, ByVal lMsg As Boolean) As Boolean
Dim Pos As Variant
On Error GoTo Errbusq
   BuscaDato = False
   Pos = rsAdo.Bookmark
   rsAdo.Find Criterio, IIf(start = 1, 0, start + 1), adSearchForward, 1
   If rsAdo.EOF Then
      rsAdo.Bookmark = Pos
      If lMsg Then
         MsgBox " ! Dato no encontrado... ! ", vbExclamation, "Error de Busqueda"
         BuscaDato = False
      End If
   Else
      BuscaDato = True
   End If
Exit Function
Errbusq:
   MsgBox (Err.Description), vbInformation, "Aviso"
End Function
'Public Sub CentraForm(frmCentra As Form)
'    frmCentra.Move (Screen.Width - frmCentra.Width) / 2, (Screen.Height - frmCentra.Height) / 2, frmCentra.Width, frmCentra.Height
'End Sub
