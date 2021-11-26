Attribute VB_Name = "gImprimirLogistica"
Option Explicit
Public TextoFinLen As Integer

Public Function JustificaTextoCadenaOrdenCompra(sTemp As String, lnColPage As Integer, Optional lsEspIzq As Integer = 0) As String
Dim vTextFin As String
Dim letra As String * 1, i As Integer, K As Integer, N As Integer
Dim nVeces As Long, m As Integer, Fin As Integer, Ini As Integer
Dim nAncho1 As Integer, nSpa As Integer
i = 0
K = 0
N = Len(sTemp)
nAncho1 = lnColPage
Do While i <= N
   K = K + 1
   i = i + 1
   If i > N Then
      Exit Do
   End If
   letra = Mid(sTemp, i, 1)
   If letra = Chr$(27) Then
      vTextFin = vTextFin & letra & Mid(sTemp, i + 1, 1)
      i = i + 1
      K = K + 1
      nAncho1 = nAncho1 + 2
   Else
      If Asc(letra) <> 13 And Asc(letra) <> 10 Then
         If K > nAncho1 Then
            m = 0
            If Mid(sTemp, i, 1) = Chr(32) Then
               vTextFin = Trim(vTextFin)
            Else
               m = InStrRev(vTextFin, " ", , vbTextCompare)
               If m = 0 Then m = 1
               If InStr(Mid(vTextFin, m, Len(vTextFin)), Chr$(27)) Then
                  nAncho1 = nAncho1 - 2
               End If
               i = i - (nAncho1 + 1 - m)
               vTextFin = Mid(vTextFin, 1, m - 1)
            End If
            nSpa = nAncho1 - Len(Trim(vTextFin))
            vTextFin = Trim(vTextFin)
            If nSpa <> 0 Then
               Fin = 1
               nVeces = 0
               m = 1
               Do While m <= nSpa
                  Ini = InStr(Fin, vTextFin, " ", vbTextCompare)
                  If Ini = 0 Then
                     Fin = 1
                     nVeces = nVeces + 1
                     m = m + 1
                  Else
                      vTextFin = Mid(vTextFin, 1, Ini) & " " & RTrim(Mid(vTextFin, Ini + 1, nAncho1))
                      Fin = Ini + 2 + nVeces
                      m = m + 1
                  End If
               Loop
            End If
            vTextFin = vTextFin & oImpresora.gPrnSaltoLinea
            JustificaTextoCadenaOrdenCompra = JustificaTextoCadenaOrdenCompra & Space(lsEspIzq) & Trim(ImpreCarEsp(vTextFin))
            nAncho1 = lnColPage
            vTextFin = ""
            letra = ""
            K = 0
        Else
            vTextFin = vTextFin & letra
            TextoFinLen = Len(Trim(vTextFin))
         End If
      Else
        If i < N Then
            If Asc(Mid(sTemp, i + 1, 1)) = 13 Or Asc(Mid(sTemp, i + 1, 1)) = 10 Then
                i = i + 1
            End If
        End If
         JustificaTextoCadenaOrdenCompra = JustificaTextoCadenaOrdenCompra & Space(lsEspIzq) & Trim(ImpreCarEsp(vTextFin)) & oImpresora.gPrnSaltoLinea
         nAncho1 = lnColPage
         vTextFin = ""
         letra = ""
         K = 0
      End If
   End If
Loop
JustificaTextoCadenaOrdenCompra = JustificaTextoCadenaOrdenCompra & Space(lsEspIzq) & Trim(ImpreCarEsp(vTextFin))
End Function

Public Function ImpreGlosaTeso(psGlosa As String, pnColPage As Integer, Optional psTitGlosa As String = "  GLOSA      : ", Optional pnCols As Integer = 0, Optional lbEnterFinal As Boolean = True, Optional ByRef nLin As Integer = 0) As String
Dim sImpre As String
Dim sTexto As String, N As Integer
Dim nLen As Integer
  nLen = Len(psTitGlosa)
  sTexto = JustificaTexto(psGlosa, IIf(pnCols = 0, pnColPage, pnCols) - nLen)
  sImpre = psTitGlosa
  N = 0
  Do While True
     N = InStr(sTexto, oImpresora.gPrnSaltoLinea)
     If N > 0 Then
        sImpre = sImpre & Mid(sTexto, 1, N - 1) & oImpresora.gPrnSaltoLinea & Space(nLen)
        sTexto = Mid(sTexto, N + 1, Len(sTexto))
        nLin = nLin + 1
     End If
     If N = 0 Then
        sImpre = sImpre & Justifica(sTexto, IIf(pnCols = 0, pnColPage, pnCols) - nLen) & IIf(lbEnterFinal, oImpresora.gPrnSaltoLinea, "")
        If lbEnterFinal Then
            nLin = nLin + 1
        End If
        Exit Do
     End If
  Loop
  ImpreGlosaTeso = sImpre
End Function
