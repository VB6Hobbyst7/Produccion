VERSION 5.00
Begin VB.Form frmCredBPPLista 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4545
   Icon            =   "frmCredBPPLista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstLista 
      Height          =   1635
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin SICMACT.FlexEdit feIndCartAtrasada 
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4305
      _extentx        =   7594
      _extenty        =   2778
      cols0           =   5
      highlight       =   1
      encabezadosnombres=   "-De (dias)-A (dias)-+ Judiciales-Aux"
      encabezadosanchos=   "0-1300-1300-1300-0"
      font            =   "frmCredBPPLista.frx":030A
      font            =   "frmCredBPPLista.frx":0336
      font            =   "frmCredBPPLista.frx":0362
      font            =   "frmCredBPPLista.frx":038E
      font            =   "frmCredBPPLista.frx":03BA
      fontfixed       =   "frmCredBPPLista.frx":03E6
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      tipobusqueda    =   3
      columnasaeditar =   "X-1-2-3-X"
      listacontroles  =   "0-0-0-4-0"
      encabezadosalineacion=   "C-R-R-L-C"
      formatosedit    =   "0-3-3-0-0"
      rowheight0      =   300
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   735
   End
End
Attribute VB_Name = "frmCredBPPLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''**********************************************************************************************
''** Nombre : frmCreBPPLista
''** Descripción : Formulario para mostrar datos extras de los parametros del BPP creado segun RFC099-2012
''** Creación : JUEZ, 20121019 09:00:00 AM
''**********************************************************************************************
'
'Option Explicit
'Dim fbCargado As Boolean
'
'Public Sub Inicio(ByVal psIdParametro As String, ByVal pnTipoLista As Integer)
'    Dim oLista As COMDCredito.DCOMBPPR
'    Dim oSubProd As COMDConstantes.DCOMConstantes
'    Dim oAge As COMDConstantes.DCOMAgencias
'    Dim rs As ADODB.Recordset
'    Dim rsLista As ADODB.Recordset
'    Dim i As Integer, j As Integer
'    Dim lnFila As Integer
'
'    fbCargado = False
'
'    Set oLista = New COMDCredito.DCOMBPPR
'
'    Set rs = oLista.RecuperaParametrosLista(psIdParametro, pnTipoLista)
'    Set oLista = Nothing
'
'    If pnTipoLista = 1 Or pnTipoLista = 2 Then
'        Me.Width = 3600
'        lstLista.Visible = True
'        feIndCartAtrasada.Visible = False
'        If pnTipoLista = 1 Then
'            Me.Caption = "Sub Productos"
'            Set oSubProd = New COMDConstantes.DCOMConstantes
'            Set rsLista = oSubProd.RecuperaConstantes(3033, 1)
'            Set oSubProd = Nothing
'        Else
'            Me.Caption = "Agencias"
'            Set oAge = New COMDConstantes.DCOMAgencias
'            Set rsLista = oAge.ObtieneAgencias()
'            Set oAge = Nothing
'        End If
'        lstLista.Clear
'        For i = 0 To rsLista.RecordCount - 1
'            lstLista.AddItem rsLista!nConsValor & " " & Trim(rsLista!cConsDescripcion)
'            rs.MoveFirst
'            For j = 0 To rs.RecordCount - 1
'                If Trim(rsLista!nConsValor) = Trim(rs!nConsValor) Then
'                    lstLista.Selected(i) = True
'                End If
'                rs.MoveNext
'            Next j
'            rsLista.MoveNext
'        Next i
'    ElseIf pnTipoLista = 3 Then
'        Me.Width = 4635
'        lstLista.Visible = False
'        feIndCartAtrasada.Visible = True
'        Me.Caption = "Cartera Atrasada"
'        Call LimpiaFlex(feIndCartAtrasada)
'        Do While Not rs.EOF
'            feIndCartAtrasada.AdicionaFila
'            lnFila = feIndCartAtrasada.row
'            feIndCartAtrasada.TextMatrix(lnFila, 1) = rs!nDiasDesde
'            feIndCartAtrasada.TextMatrix(lnFila, 2) = rs!nDiasHasta
'            feIndCartAtrasada.TextMatrix(lnFila, 3) = IIf(rs!nJudicial = 1, 1, "")
'            rs.MoveNext
'        Loop
'        rs.Close
'        Set rs = Nothing
'    End If
'
'    fbCargado = True
'    Me.Show 1
'End Sub
'
'Private Sub cmdSalir_Click()
'    Unload Me
'End Sub
'
'Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'
'    End If
'End Sub
'
'Private Sub lstLista_ItemCheck(Item As Integer)
'    If fbCargado Then
'        If lstLista.Selected(Item) Then
'            lstLista.Selected(Item) = False
'        Else
'            lstLista.Selected(Item) = True
'        End If
'    End If
'End Sub
