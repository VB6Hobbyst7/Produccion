VERSION 5.00
Begin VB.Form frmComparaTablas 
   Caption         =   "Form1"
   ClientHeight    =   1365
   ClientLeft      =   4365
   ClientTop       =   3765
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   3300
   Begin VB.CommandButton Command1 
      Caption         =   "Comparar Tablas"
      Height          =   450
      Left            =   435
      TabIndex        =   0
      Top             =   255
      Width           =   2040
   End
End
Attribute VB_Name = "frmComparaTablas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oCon As DConecta
Dim oCon1 As DConecta
Private Sub Command1_Click()
Dim sql As String
Dim rsTablas As ADODB.Recordset
Dim rsProd As ADODB.Recordset
Dim rsMig As ADODB.Recordset
Dim lbDifNombre As Boolean
Dim lsDifTipos As String
Dim lsDifPresicion As String
Dim lsDifLongitud As String
Dim lsDifCampos As String
Dim lsCadenas As String
Dim lnTot As Long
Dim i As Long
Dim lnLinea As Long
lsCadenas = "COMPARACION DE CAMPOS BASE DE DATOS [MIG] Y BASE DE PRODUCCION " & Chr(10) & Chr(10)
lsCadenas = lsCadenas & " NOMBRE DE TABLA     TIPO DE DIFERENCIA    DATOS EN [PROD]      DATOS EN [MIG]"
sql = "select * from sysobjects where Xtype ='U'"
Set rsTablas = oCon1.CargaRecordSet(sql)
lnTot = rsTablas.RecordCount
i = 0
lnLinea = 2
Do While Not rsTablas.EOF
    i = i + 1
    Me.Caption = "Tabla :" & i & " de " & lnTot
    lsCadenas = lsCadenas & "TABLA = " & Trim(UCase(rsTablas!Name)) & Chr(10)
    lnLinea = lnLinea + 1
    sql = "EXEC sp_columns '" & Trim(rsTablas!Name) & "'"
    Set rsProd = oCon.CargaRecordSet(sql)
    Set rsMig = oCon1.CargaRecordSet(sql)
    Do While Not rsMig.EOF
        lbDifNombre = False
        If rsProd.RecordCount > 0 Then
            rsProd.MoveFirst
            Do While Not rsProd.EOF
                'comparamos los nombres de columnas
                If Trim(rsProd!COLUMN_NAME) = Trim(rsMig!COLUMN_NAME) Then
                    lbDifNombre = True
                    Exit Do
                Else
                    lbDifNombre = False
                End If
                rsProd.MoveNext
            Loop
            lsDifTipos = ""
            lsDifPresicion = ""
            lsDifLongitud = ""
            lsDifCampos = ""
            If lbDifNombre = True Then 'nombres iguales
                If Trim(rsProd!TYPE_NAME) <> Trim(rsMig!TYPE_NAME) Then
                    lsDifTipos = Space(2) & "DIF.TIPO.DATOS [" & rsMig!COLUMN_NAME & "] =  [PROD]" & rsProd!TYPE_NAME & Space(5) & "[MIG]" & rsMig!TYPE_NAME & Chr(10)
                End If
                If rsProd!Precision <> rsMig!Precision Then
                    lsDifPresicion = Space(2) & "DIF.PRECISION [" & rsMig!COLUMN_NAME & "] = [PROD]" & rsProd!Precision & Space(5) & "[MIG]" & rsMig!Precision & Chr(10)
                End If
                If rsProd!Length <> rsMig!Length Then
                    lsDifLongitud = Space(2) & "DIF.Longitud [" & rsMig!COLUMN_NAME & "] = [PROD]" & rsProd!Length & Space(5) & "[MIG]" & rsMig!Length & Chr(10)
                End If
            Else
                'nombre de campos diferentes
                lsDifCampos = Space(2) & "NOMBRE CAMPO DIFERENTE / CAMPO NO EXISTE EN [PROD]= " & Trim(rsMig!COLUMN_NAME) & Space(5) & Chr(10) '& Trim(rsMig!COLUMN_NAME)
            End If
            If lsDifCampos & lsDifTipos & lsDifPresicion & lsDifLongitud <> "" Then
                lsCadenas = lsCadenas & lsDifCampos & lsDifTipos & lsDifPresicion & lsDifLongitud
                lnLinea = lnLinea + 1
            End If
        Else
            lsCadenas = lsCadenas & "        NO EXISTE EN BASE [PROD] TABLA = " & Trim(rsTablas!Name) & Chr(10)
            lnLinea = lnLinea + 1
            Exit Do
        End If
        If lnLinea > 60 Then
            lsCadenas = lsCadenas & Chr(12)
            lnLinea = 0
        End If
        rsMig.MoveNext
    Loop
    rsProd.Close
    rsMig.Close
    rsTablas.MoveNext
Loop
rsTablas.Close
Set rsTablas = Nothing

EnviaPrevio lsCadenas, "", 66, True

End Sub

Private Sub Form_Load()
Set oCon = New DConecta
Set oCon1 = New DConecta

oCon.AbreConexion
oCon1.AbreConexion "PROVIDER=SQLOLEDB;User ID=dbaccess;Password=cmacica;INITIAL CATALOG=DBCMACICAMIG;DATA SOURCE=01srvsicmac02"


End Sub

Private Sub Form_Unload(Cancel As Integer)
oCon.CierraConexion
oCon1.CierraConexion
End Sub
