VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGarantMultiriesgoMYPE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Declaración Jurada Multiriesgo MYPE"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   Icon            =   "frmGarantMultiriesgoPYME.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGenerarDJ 
      Caption         =   "Generar DJ"
      Height          =   310
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   310
      Left            =   7680
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5741
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Garantías Multiriesgo"
      TabPicture(0)   =   "frmGarantMultiriesgoPYME.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblMontoPrestamo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblTotalAsegurado"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "feGarantias"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin SICMACT.FlexEdit feGarantias 
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   3625
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-N-Sel-Bien en Garantía-Tipo Documento-Moneda-VRM-Valor Asegura-NumGaran"
         EncabezadosAnchos=   "300-0-300-2500-2000-700-1200-1200-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-X-X-X-X-7-X"
         ListaControles  =   "0-0-4-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-L-L-C-R-R-C"
         FormatosEdit    =   "0-0-0-0-0-0-2-2-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label lblTotalAsegurado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7305
         TabIndex        =   6
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Total Asegurado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   5
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label lblMontoPrestamo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Monto de Crédito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Garantías no preferidas no asociadas al crédito:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmGarantMultiriesgoMYPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmGarantMultiriesgoPYME
'** Descripción : Formulario que permite asignar un seguro multiriesgo a una garantia
'** Creación    : RECO, 20150509 - ERS023-2015
'**********************************************************************************************
Option Explicit

Dim sPersCod As String
Dim nMontoPrestamo As Double
Dim sCtaCod As String
Dim nCuotas As Integer
Dim vMatriz() As Variant

Public Function inicia(ByVal psPersCod As String, ByVal pnMontoPrestamo As Double, ByVal psCtaCod As String, ByVal pnCuotas As Integer) As Variant
    psPersCod = psPersCod
    nMontoPrestamo = pnMontoPrestamo
    sCtaCod = psCtaCod
    nCuotas = pnCuotas
    If (CargarDatos(psPersCod) = True) Then
        Me.Show 1
    Else
        MsgBox "No existen garantìas que puedan ser aseguradas por el producto Seguro Multiriesgo MYPE.", vbInformation, "Alerta"
    End If
    inicia = vMatriz
End Function

Private Function CargarDatos(ByVal psPersCod As String) As Boolean
    Dim obj As New COMNCredito.NCOMGarantia
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    CargarDatos = False
    Set rs = obj.ObtieneGarantMultiRiesgoMYPE(psPersCod, sCtaCod)
    If Not (rs.EOF And rs.BOF) Then
        feGarantias.Clear
        FormateaFlex feGarantias
        For i = 1 To rs.RecordCount
            feGarantias.AdicionaFila
            feGarantias.TextMatrix(i, 1) = "1"
            feGarantias.TextMatrix(i, 2) = 0 'RS!nEstado
            feGarantias.TextMatrix(i, 3) = rs!cObjGarantDesc
            feGarantias.TextMatrix(i, 4) = rs!cTpoDoc
            feGarantias.TextMatrix(i, 5) = rs!cMoneda
            feGarantias.TextMatrix(i, 6) = rs!nRealizacion
            feGarantias.TextMatrix(i, 7) = 0
            feGarantias.TextMatrix(i, 8) = rs!cNumGarant
            rs.MoveNext
        Next
        CargarDatos = True
    End If
    lblMontoPrestamo.Caption = Format(nMontoPrestamo, gsFormatoNumeroView)
End Function

Private Sub cmdGenerarDJ_Click()
    Call ImprimeDJ
End Sub

Private Sub CmdGrabar_Click()
    Dim i As Integer, J As Integer
    If CDbl(Replace(lblTotalAsegurado.Caption, ",", "")) <> CDbl(Replace(lblMontoPrestamo.Caption, ",", "")) Then
        MsgBox "El monto total asegurado debe ser igual al monto del crédito.", vbInformation, "Alerta"
        Exit Sub
    End If
    ReDim vMatriz(2, 0)
    For i = 1 To feGarantias.Rows - 1
        If feGarantias.TextMatrix(i, 2) = "." Then
            J = J + 1
            ReDim Preserve vMatriz(2, 0 To J)
            vMatriz(1, J) = feGarantias.TextMatrix(i, 8)
            vMatriz(2, J) = feGarantias.TextMatrix(i, 7)
        End If
    Next
    Unload Me
End Sub

Private Sub feGarantias_OnCellChange(pnRow As Long, pnCol As Long)
    If Not ValidaMontoTabla(pnRow, 7) Then feGarantias.TextMatrix(pnRow, 7) = 0
    lblTotalAsegurado.Caption = Format(SumaFilaSel, gsFormatoNumeroView)
End Sub

Private Sub FeGarantias_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    If Not ValidaMontoTabla(pnRow, 7) Then
        feGarantias.TextMatrix(pnRow, 7) = 0
    End If
    lblTotalAsegurado.Caption = Format(SumaFilaSel, gsFormatoNumeroView)
End Sub

Private Function Cant() As Integer
    Dim i As Integer
    For i = 1 To feGarantias.Rows - 1
        If feGarantias.TextMatrix(i, 2) = "." Then
            Cant = Cant + 1
        End If
    Next
End Function

Private Function ValidaMontoTabla(ByVal pnFil As Integer, ByVal pnCol As Integer) As Boolean
    On Error GoTo Errorr
    If CDbl(feGarantias.TextMatrix(pnFil, pnCol - 1)) < CDbl(feGarantias.TextMatrix(pnFil, pnCol)) Or feGarantias.TextMatrix(pnFil, pnCol) < 0 Then
        ValidaMontoTabla = False
    Else
        ValidaMontoTabla = True
    End If
    Exit Function
Errorr:
    ValidaMontoTabla = False
    feGarantias.TextMatrix(pnFil, pnCol) = 0
End Function

Private Function SumaFilaSel() As Double
    Dim i As Integer
    SumaFilaSel = 0
    For i = 1 To feGarantias.Rows - 1
        If feGarantias.TextMatrix(i, 2) = "." Then
            SumaFilaSel = SumaFilaSel + feGarantias.TextMatrix(i, 7)
        End If
    Next
End Function

Private Sub Form_Load()
    ReDim vMatriz(2, 0)
End Sub

Private Sub ImprimeDJ()
    On Error GoTo ErrorImprimirPDF
    Dim obj As New COMDConstantes.DCOMConstantes
    Dim oGarant  As New COMNCredito.NCOMGarantia
    Dim rsClient As New ADODB.Recordset
    Dim rsGarant As New ADODB.Recordset
    
    Dim oDoc  As New cPDF
    Dim nTipo As Integer
    
    Dim i As Integer
    Dim a As Integer
    Dim nPos As Integer
    Dim nPosicion As Integer
    Dim lnFontSizeBody As Integer
    Dim sCartaCod As String
    Dim sMesDesc As String
    Dim nPosY As Integer
    Dim nPosY2 As Integer
    Dim x As Integer
    'sCartaCod = txtCartaNum.Text & "-" & txtCartaAnio.Text & "-GS-GA/CMACM"
    'sMesDesc = obj.DameDescripcionConstante(1010, Month(gdFecSis))
    
    If Cant < 1 Then
        MsgBox "Debe seleccionar por lo menos una garantía", vbInformation, "Alerta"
        Exit Sub
    End If
    Set rsClient = oGarant.RecuperaDatosClienteDJMYPE(sCtaCod)
        
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "DECLARACION JURADA Nº " & "NUMERO" 'sCartaCod
    oDoc.Title = "DECLARACION JURADA Nº " & "NUMERO" 'sCartaCod
    If Not oDoc.PDFCreate(App.Path & "\Spooler\" & Replace("NUMERO", "/", "-") & "_" & gsCodUser & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
    oDoc.Fonts.Add "F1", "Book Antiqua", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Book Antiqua", TrueType, Bold, WinAnsiEncoding
    oDoc.Fonts.Add "F3", "Book Antiqua", TrueType, BoldItalic, WinAnsiEncoding
    oDoc.Fonts.Add "F4", "Book Antiqua", TrueType, Italic, WinAnsiEncoding
    
    
    lnFontSizeBody = 6
    nPosY = 120
    nPosY2 = 400
    oDoc.NewPage A4_Vertical
    
    oDoc.WTextBox 80, 50, 10, 300, "Pág 1 de 1, " & Day(gdFecSis) & " de " & Left(sMesDesc, 1) & LCase(Mid(sMesDesc, 2, Len(sMesDesc))) & " del " & Year(gdFecSis), "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 80, 450, 10, 300, "Usuario :" & gsCodUser, "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 90, 50, 10, 300, gsNomAge, "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 90, 450, 10, 300, "Fecha :" & gdFecSis, "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 100, 450, 10, 300, "Hora :" & Time, "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 107, 240, 10, 300, "DECLARACIONJURADA DESEGURO MULTIRIESGO", "F1", lnFontSizeBody, hjustify, , vbBlack
    'oDoc.WTextBox 117, 270, 10, 300, "Nº000001", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 127, 50, 10, 550, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 137, 50, 10, 300, "I. RESUMEN INFORMATIVO COLOR NARANJA  PÓLIZA Nº", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 147, 50, 10, 550, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 157, 50, 10, 200, "SUMA ASEGURADA AL   :", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 157, 150, 10, 200, "100% DEL MONTO APROBADO.", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 167, 50, 10, 200, "COBERTURA DE PÓLIZA :", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 167, 150, 10, 500, "INCENDIO, EXPLOSIÓN, INUNDACIÓN Y TERREMOTO HASTA LA SUMA ASEGURADA QUE CORRESPONDA.", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 177, 50, 10, 550, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 187, 50, 10, 500, "II. DATOS DEL CLIENTE", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 197, 50, 10, 550, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", "F1", lnFontSizeBody, hjustify, , vbBlack
    
    oDoc.WTextBox 207, 50, 10, 500, "TIPO DE PERSONA : ", "F1", lnFontSizeBody, hjustify, , vbBlack:                  oDoc.WTextBox 207, nPosY, 10, 500, rsClient!cTpoPers, "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 207, 280, 10, 500, "R. SOCIAL/NOMBRE    : ", "F1", lnFontSizeBody, hjustify, , vbBlack:             oDoc.WTextBox 207, nPosY2, 10, 175, rsClient!cPersNombre, "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 219, 50, 10, 500, "DOI / RUC       : ", "F1", lnFontSizeBody, hjustify, , vbBlack:                  oDoc.WTextBox 219, nPosY, 10, 500, rsClient!cPersIDnro, "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 219, 280, 10, 500, "FECHA NACIMIENTO/CONSTITUCION: ", "F1", lnFontSizeBody, hjustify, , vbBlack:    oDoc.WTextBox 219, nPosY2, 10, 500, rsClient!dPersNacCreac, "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 229, 50, 10, 500, "TIPO DE NEGOCIO : ", "F1", lnFontSizeBody, hjustify, , vbBlack:                  oDoc.WTextBox 229, nPosY, 10, 500, rsClient!cTpoNeg, "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 229, 280, 10, 500, "GIRO DE NEGOCIO :", "F1", lnFontSizeBody, hjustify, , vbBlack:                  oDoc.WTextBox 229, nPosY2, 10, 175, rsClient!cActiGiro, "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 241, 50, 10, 500, "DIRECCION           : ", "F1", lnFontSizeBody, hjustify, , vbBlack:              oDoc.WTextBox 241, nPosY, 10, 500, rsClient!cPersDireccDomicilio, "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 241, 280, 10, 500, "DEPARTAMENTO:     ", "F1", lnFontSizeBody, hjustify, , vbBlack:                 oDoc.WTextBox 241, nPosY2, 10, 500, rsClient!DEP, "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 249, 50, 10, 500, "PROVINCIA: ", "F1", lnFontSizeBody, hjustify, , vbBlack:                         oDoc.WTextBox 249, nPosY, 10, 500, rsClient!PRO, "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 249, 280, 10, 500, "DISTRITO: ", "F1", lnFontSizeBody, hjustify, , vbBlack:                         oDoc.WTextBox 249, nPosY2, 10, 500, rsClient!DIS, "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 259, 50, 10, 500, "TELEFONO                :", "F1", lnFontSizeBody, hjustify, , vbBlack:           oDoc.WTextBox 259, nPosY, 10, 500, rsClient!cPersTelefono, "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 259, 280, 10, 500, "CELULAR             : ", "F1", lnFontSizeBody, hjustify, , vbBlack:             oDoc.WTextBox 259, nPosY2, 10, 500, rsClient!cPersCelular, "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 269, 50, 10, 500, "E-MAIL    : ", "F1", lnFontSizeBody, hjustify, , vbBlack:                        oDoc.WTextBox 269, nPosY, 10, 500, rsClient!cPersEmail, "F1", lnFontSizeBody, hjustify, , vbBlack
    
    oDoc.WTextBox 277, 50, 10, 550, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 287, 50, 10, 500, "III. DIRECCION DE UBICACIÓN DE MATERIA ASEGURADA", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 297, 50, 10, 550, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", "F1", lnFontSizeBody, hjustify, , vbBlack
    
    oDoc.WTextBox 307, 50, 10, 500, "DECLARO(MOS)POSEER BIENES QUE CUBREN EL IMPORTE DEL CRÉDITO APROBADO POR CMAC MAYNAS S.A. QUE SERÁN CUBIERTOS POR LA PÓLIZA MENCIONADA EN EL PRESENTE" & _
                                     " DOCUMENTO Y QUE SON DE LA SIGUIENTE CLASE: MERCADERIA / MAQUINARIA / EQUIPO / MUEBLES", "F1", lnFontSizeBody, hjustify, , vbBlack
    
    oDoc.WTextBox 327, 50, 10, 550, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 337, 50, 10, 500, "IV. DIRECCION DE UBICACIÓN DE MATERIA ASEGURADA", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 347, 50, 10, 550, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 357, 50, 10, 500, "TIPO DETALLE", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 367, 50, 10, 550, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", "F1", lnFontSizeBody, hjustify, , vbBlack
    
    For x = 1 To Cant
        Set rsGarant = oGarant.RecuperaDatosGarantDJMYPE(feGarantias.TextMatrix(x, 8))
        
        oDoc.WTextBox 377, 50, 10, 500, x & rsGarant!cDireccion, "F1", lnFontSizeBody, hjustify, , vbBlack
        oDoc.WTextBox 387, 50, 10, 500, "AÑO DE CONSTRUCCIÓN: " & rsGarant!nAnioConstruc, "F1", lnFontSizeBody, hjustify, , vbBlack
        oDoc.WTextBox 387, 150, 10, 500, "N° DE PISOS:" & rsGarant!nNumPisos, "F1", lnFontSizeBody, hjustify, , vbBlack
        oDoc.WTextBox 387, 250, 10, 500, "USO:" & rsGarant!cUso, "F1", lnFontSizeBody, hjustify, , vbBlack
        oDoc.WTextBox 387, 300, 10, 500, "VIVIENDA    MATERIAL: " & rsGarant!cTpovivienda, "F1", lnFontSizeBody, hjustify, , vbBlack
        oDoc.WTextBox 397, 50, 10, 500, "MATERIA ASEGURADA: EQUIPO" & rsGarant!cMateriaAseg, "F1", lnFontSizeBody, hjustify, , vbBlack
        nPosicion = nPosicion + 10
    Next
    
    oDoc.WTextBox 407 + nPosicion, 50, 10, 550, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 417 + nPosicion, 50, 10, 500, "V.SUMA ASEGURADA", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 427 + nPosicion, 50, 10, 550, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 437 + nPosicion, 50, 10, 500, "MONTO APROBADO   :", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 437 + nPosicion, 150, 10, 500, "S/." & Format(lblTotalAsegurado.Caption, gsFormatoNumeroView) & " (00/100 NUEVOS SOLES)", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 437 + nPosicion, 300, 10, 500, "MONEDA:", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 437 + nPosicion, 350, 10, 500, IIf(Mid(sCtaCod, 9, 1) = 1, "SOLES", "DOLARES"), "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 437 + nPosicion, 400, 10, 500, "N° CUOTAS: ", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 437 + nPosicion, 450, 10, 500, nCuotas, "F1", lnFontSizeBody, hjustify, , vbBlack

    oDoc.WTextBox 447 + nPosicion, 50, 10, 550, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 457 + nPosicion, 50, 10, 500, "DECLARACIONES Y FIRMAS", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 467 + nPosicion, 50, 10, 550, "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 477 + nPosicion, 50, 10, 500, "DECLARO(MOS) CONOCER QUE LA PRESENTE DECLARACIÓN JURADA SE EXTIENDE A UNA INSTITUCIÓND EL SISTEMA FINANCIERO A EFECTOS DE OBTENER UN CRÉDITO, POR LO QUE DE ACUERDO A LO DISPUESTO POR EL ART." & _
                                    " 179° DE LA LEY GENERAL DEL SISTEMA FINANCIERO Y SISTEMA DE SEGUROS Y ORGÁNICA DE LA SUPERINTENDENCIA DE BANCA Y SEGUROS, LEY N°26702, EL BANCO SE ENCUENTRA FACULTADO PARA RESOLVER EL RESPECTIVO CONTRATO" & _
                                    " , ENCONTRÁNDOME SUJETO A LAS RESPONSABILIDADES PENALES CORRESPONDIENTE EN CASOS SE DETERMINE LA FALSEDAD DE LA INFORMACIÓNCON TENIDA EN EL PRESENTE DOUMENTO, SEGÚN LO ESTABLECIDO EN EL ART." & _
                                    " 247° DEL CÓDIGO PENAL.", "F1", lnFontSizeBody, hjustify, , vbBlack

    oDoc.WTextBox 537 + nPosicion, 50, 10, 500, "ASIMISMO, DECLARO(MOS) QUE LA INFORMACIÓN AQUÍ PROPORCIONADA FORMA PARTE INTEGRANTE DE LA PÓLIZA DE SEGURO DE MULTIRIESGO SEGÚN EL IMPORTE DEL CRÉDITO.", "F1", lnFontSizeBody, hjustify, , vbBlack


    oDoc.WTextBox 707 + nPosicion, 50, 10, 550, "-------------------------------------------", "F1", lnFontSizeBody, hjustify, , vbBlack
    oDoc.WTextBox 717 + nPosicion, 50, 10, 500, "FIRMA CLIENTE", "F1", lnFontSizeBody, hjustify, , vbBlack


oDoc.PDFClose
oDoc.Show
Exit Sub
ErrorImprimirPDF:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub
