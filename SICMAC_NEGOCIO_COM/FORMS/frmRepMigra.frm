VERSION 5.00
Begin VB.Form frmRepMigra 
   Caption         =   "Form1"
   ClientHeight    =   1530
   ClientLeft      =   4425
   ClientTop       =   3945
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1530
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "frmRepMigra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim ocon As DConecta
    Set ocon = New DConecta
    
    ocon.AbreConexion
    
    Dim rs As ADODB.Recordset
    Set rs = ocon.CargaRecordSet("Exec sp_MigPersona")
    
    ocon.CierraConexion
    
    Dim cadaux As String
    
    cadaux = ""
    cadaux = PrnSet("C+") & ImpreFormat("Codigo", 14) & ImpreFormat("Nombre", 60) & ImpreFormat("Tipo", 20) & _
                ImpreFormat("Tip.Doc.", 10) & ImpreFormat("Nro.Doc.", 15) & ImpreFormat("Zona", 30) & _
                ImpreFormat("Direcc.", 60) & ImpreFormat("Sexo", 4) & vbCrLf
    
    Dim i As Long
    Dim lntotal As Long
    Dim cont As Long
    lntotal = rs.RecordCount
    i = 0
    Do While Not rs.EOF
        cont = cont + 1
        i = i + 1
        cadaux = cadaux & ImpreFormat(rs("Codigo"), 14) & ImpreFormat(rs("Nombre"), 60) & ImpreFormat(rs("Tipo_Persona"), 20) & _
                ImpreFormat(rs("Documento_Identidad"), 10) & ImpreFormat(rs("Numero_Documento"), 15) & ImpreFormat(IIf(IsNull(rs("Zona")), " ", rs("Zona")), 30) & _
                ImpreFormat(Replace(Replace(IIf(IsNull(rs("Direccion")), " ", Trim(rs("Direccion"))), Chr(13), ""), Chr(10), ""), 60) & ImpreFormat(IIf(IsNull(rs("Sexo")), " ", rs("Sexo")), 1) & vbCrLf
                
        If i > 62 Then
            cadaux = cadaux & Chr(12)
            cadaux = cadaux & ImpreFormat("Codigo", 14) & ImpreFormat("Nombre", 60) & ImpreFormat("Tipo", 20) & _
                ImpreFormat("Tip.Doc.", 10) & ImpreFormat("Nro.Doc.", 15) & ImpreFormat("Zona", 30) & _
                ImpreFormat("Direcc.", 60) & ImpreFormat("Sexo", 4) & vbCrLf
            i = 0
        End If
        Me.Caption = "Registro " & cont & " de " & lntotal
        rs.MoveNext
        DoEvents
    Loop
    rs.Close
    Set rs = Nothing
    
'    Dim oprev As Previo.clsPrevio
'    Set oprev = New clsPrevio
'
    EnviaPrevio cadaux, " ", gnLinPage
    
End Sub

