VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

DefInt A-Z

' String
Public sPath_File As String

Public Sub gsExibeErros(objErro As ErrObject, pMetodo As String, pObjeto As String, Optional pSql As String)

    Dim pMsg As String

    Screen.MousePointer = 0
    If objErro.Number = 0 Then Exit Sub

    pMsg = ""
    pMsg = "Ocorreu um Erro Interno no Sistema, Contate " & _
           "o Administrador do Sistema e Informe-o desta " & _
           "Ocorrêcia." & vbNewLine & vbNewLine & _
           "Objeto....: " & pObjeto & vbNewLine & _
           "Método....: " & pMetodo & vbNewLine & vbNewLine & _
           "Ocorrências: " & vbNewLine & "[" & CStr(objErro.Number) & "] - " & objErro.Description

    MsgBox pMsg, vbCritical + vbOKOnly, "Atenção."

    If Len(pSql) > 0 Then
        Call psLogErro(pMetodo, objErro.Number, CStr(objErro.Description), pSql)
    Else
        Call psLogErro(pMetodo, objErro.Number, CStr(objErro.Description))
    End If

End Sub

Private Sub psLogErro(pMetodo As String, pErro As Variant, pDescricao As String, Optional pSql As String)

    Dim sReg As String
    Dim sArquivo As String

    sArquivo = sPath_File & "\LogErro.TXT"
    Open sArquivo For Append As #99

    sReg = "Método.............: " & pMetodo & vbNewLine
    sReg = sReg & "Data/Hora..........: " & Format(Now, "DD/MM/YYYY HH:MM:SS") & vbNewLine
    sReg = sReg & "Numero do Erro.....: #" & CStr(pErro) & vbNewLine
    sReg = sReg & "Descrição..........: " & LTrim(RTrim(pDescricao)) & vbNewLine

    If Len(pSql) > 0 Then
        sReg = sReg & "Instrução SQL......: " & LTrim(RTrim(pSql)) & vbNewLine
    End If

    Print #99, sReg

    Close #99

End Sub

Public Sub gsKeyAscii(vTexto As Variant, iKeyAscii As Integer, vTipo As Variant)

    If iKeyAscii = 8 Then Exit Sub

    Select Case vTipo

        Case 1
            iKeyAscii = Asc(UCase(Chr(iKeyAscii)))

        Case 2
            iKeyAscii = Asc(UCase(Chr(iKeyAscii)))
            If iKeyAscii >= 60 And iKeyAscii <= 90 Or iKeyAscii = 13 Or iKeyAscii = 8 Then
            Else
                iKeyAscii = 0
            End If

        Case 3
            If iKeyAscii >= 48 And iKeyAscii <= 57 Or iKeyAscii = 13 Or iKeyAscii = 8 Then
            Else
                iKeyAscii = 0
            End If

        Case 4
            If iKeyAscii >= 48 And iKeyAscii <= 57 Or iKeyAscii = 13 Or _
               iKeyAscii = 8 Or iKeyAscii = Asc(",") Or iKeyAscii = Asc(".") Then
            Else
                iKeyAscii = 0
            End If

            If InStr(1, vTexto, ",") = 0 Then
                If iKeyAscii >= 48 And iKeyAscii <= 57 Or iKeyAscii = 13 Or _
                   iKeyAscii = 8 Or iKeyAscii <> Asc(",") Or iKeyAscii <> Asc(".") Then
                Else
                    iKeyAscii = 0
                End If
            Else
                If iKeyAscii >= 48 And iKeyAscii <= 57 Or iKeyAscii = 13 Or _
                   iKeyAscii = 8 Or iKeyAscii <> Asc(".") Then
                Else
                    iKeyAscii = 0
                End If
            End If

        Case 6
            If iKeyAscii >= 48 And iKeyAscii <= 57 Or iKeyAscii = 13 Or iKeyAscii = 8 Or _
                iKeyAscii = Asc(".") Or iKeyAscii = Asc("-") Or iKeyAscii = Asc("/") Then
            Else
                iKeyAscii = 0
            End If

        Case 7
            iKeyAscii = Asc(LCase(Chr(iKeyAscii)))

    End Select

End Sub

Public Sub gsLimparMatriz(vMatriz As Variant)

    Dim iLin As Integer
    Dim iCol As Integer
    Dim iT1 As Integer
    Dim iT2 As Integer

    On Error GoTo err_gsLimparMatriz:

    iT1 = UBound(vMatriz, 1)
    iT2 = UBound(vMatriz, 2)

    iLin = 1
    While iLin <= iT1
        iCol = 1
        While iCol <= iT2
            vMatriz(iLin, iCol) = ""
            iCol = iCol + 1
        Wend
        iLin = iLin + 1
    Wend

    Exit Sub

err_gsLimparMatriz:
    Call gsExibeErros(Err, "gsLimparMatriz()", "Classe Sistema")

End Sub

Public Function gfsParseSql(pSql As Variant) As String

    Dim nLoop As Variant
    Dim cAux As String, cPalavra As String
    Dim nContAux As Integer
    Dim bApostrofe As Boolean
    Dim bAnteriorIsChave As Boolean
    Dim cCaractere As String
    Dim bIsDuploBranco As Boolean

    bAnteriorIsChave = False
    bApostrofe = False
    bIsDuploBranco = False
    nContAux = 0
    cAux = ""
    nLoop = 1

    Do While nLoop <= Len(pSql)

        If Mid(pSql, nLoop, 1) = " " Then
            nLoop = nLoop + 1
            cPalavra = " "

            If nLoop > 1 Then
                'verifica se o caractere anterior é espaco em branco
                If Mid(pSql, nLoop - 1, 1) = " " Then
                    bIsDuploBranco = True
                Else
                    bIsDuploBranco = False
                End If
            End If

        Else
            Do While Mid(pSql, nLoop, 1) <> " " And nLoop <= Len(pSql)
                If Mid(pSql, nLoop, 1) = "'" Then
                    bApostrofe = Not bApostrofe
                End If

                cCaractere = Mid(pSql, nLoop, 1)

                If nLoop > 1 Then


                    'verifica se o caractere anterior é espaco em branco
                    If Mid(pSql, nLoop - 1, 1) = " " Then
                        bIsDuploBranco = True
                    Else
                        bIsDuploBranco = False
                    End If
                End If

                'se o caractere for , ; ( ) etc, joga pra outra linha
                If gfbIsPalavraChave(cCaractere) Then
                    cCaractere = cCaractere & vbCrLf
                End If

                cPalavra = cPalavra & cCaractere

                nLoop = nLoop + 1
            Loop
        End If
        If gfbIsPalavraChave(Trim(cPalavra)) Then
            cAux = cAux & cPalavra & vbCrLf
        Else
            If bIsDuploBranco Then
                cAux = cAux & Trim(cPalavra)
            Else
                cAux = cAux & cPalavra
            End If
        End If

    Loop
    gfsParseSql = cAux

End Function

Public Function gfbIsPalavraChave(pPalavra As String) As Boolean

     Select Case UCase(pPalavra)
        Case ";"
            gfbIsPalavraChave = True
        Case "INTO"
            gfbIsPalavraChave = True
'        Case "("
'            gfbIsPalavraChave = True
'        Case ")"
'            gfbIsPalavraChave = True
        Case "END"
            gfbIsPalavraChave = True
        Case "BEGIN"
            gfbIsPalavraChave = True
        Case "INSERT"
            gfbIsPalavraChave = True
        Case "UPDATE"
            gfbIsPalavraChave = True
        Case "DELETE"
            gfbIsPalavraChave = True
        Case ","
            gfbIsPalavraChave = True
        Case "SELECT"
            gfbIsPalavraChave = True
        Case "FROM"
            gfbIsPalavraChave = True
        Case "INNER"
            gfbIsPalavraChave = True
        Case "LEFT"
            gfbIsPalavraChave = True
        Case "RIGHT"
            gfbIsPalavraChave = True
        Case "JOIN"
            gfbIsPalavraChave = True
        Case "ON"
            gfbIsPalavraChave = True
        Case "AND"
            gfbIsPalavraChave = True
        Case "WHERE"
            gfbIsPalavraChave = True
        Case "ORDER"
            gfbIsPalavraChave = True
        Case "BY"
            gfbIsPalavraChave = True
        Case "DISTINCT"
            gfbIsPalavraChave = True
        Case "GROUP"
            gfbIsPalavraChave = True
        Case Else
            gfbIsPalavraChave = False
    End Select

End Function

Public Sub gsLimparText(pNm_Form As String)

    Dim pObj As Object
    Dim pFrm As Form

    For Each pFrm In Forms

        For Each pObj In pFrm.Controls

            If TypeOf pObj Is TextBox And pFrm.Name = pNm_Form Then

                pObj = ""

            Else

                If TypeOf pObj Is RichTextBox And pFrm.Name = pNm_Form Then

                    pObj = ""

                End If

            End If

        Next

    Next

End Sub

Public Function gfsRemoveStr(sTexto As String, sCaracter As String) As String

    '--- Integer
    Dim iContador As Integer
    Dim iPosicao  As Integer

    '--- String
    Dim sCharOrigem As String
    Dim sCaracterConv As String
    Dim sCaractere As String
    Dim sConvertido As String

    sCaractere = sCaracter
    sConvertido = ""
    For iContador = 1 To Len(sTexto)
        DoEvents
        sCharOrigem = Mid$(sTexto, iContador, 1)
        iPosicao = InStr(1, sCaractere, sCharOrigem)
        If iPosicao > 0 Then sCaracterConv = Mid$(sConvertido, iPosicao, 1)
        If Not iPosicao > 0 Then sCaracterConv = sCharOrigem
        gfsRemoveStr = gfsRemoveStr & sCaracterConv
    Next

End Function

Public Function gfsSubstituir(ByVal pTexto As String, pProcurar As String, pTrocar As String) As String

    Dim pPosicao As Long
    Dim pInicio As Long
    Dim pNovoTexto As String

    pInicio = 1
    pPosicao = InStr(pTexto, pProcurar)

    Do While pPosicao > 0
        pNovoTexto = pNovoTexto & Mid(pTexto, pInicio, pPosicao - pInicio) & pTrocar
        pInicio = pPosicao + Len(pProcurar)
        pPosicao = InStr(pInicio, pTexto, pProcurar)
    Loop

    gfsSubstituir = pNovoTexto & Mid(pTexto, pInicio)

End Function

Public Sub gsLimparVetor(pVetor As Variant)

    Dim pLin    As Long
    Dim pTam    As Long

    On Error GoTo err_gsLimparVetor:

    DoEvents: pTam = UBound(pVetor, 1): pLin = 1
    While pLin <= pTam
        pVetor(pLin) = "": pLin = pLin + 1
    Wend

    Exit Sub

err_gsLimparVetor:
    Call gsExibeErros(Err, "gsLimparVetor()", "Classe Sistema")

End Sub

Public Sub gsBuscaCombo(pCombo As Object, pCodigo As Long)

    Dim pAchou  As Boolean
    Dim pLoop   As Long

    On Error GoTo err_gsBuscaCombo:

    If pCombo.ListCount = 0 Then pCombo.ListIndex = -1: Exit Sub
    Screen.MousePointer = 11
    pAchou = False
    For pLoop = 0 To pCombo.ListCount - 1
        If pCombo.ItemData(pLoop) = pCodigo Then pCombo.ListIndex = pLoop: pAchou = True: Exit For
    Next pLoop
    If Not pAchou Then pCombo.ListIndex = -1
    Screen.MousePointer = 0

    Exit Sub

err_gsBuscaCombo:
    pCombo.ListIndex = -1
    Call objSystem.gsExibeErros(Err, "gsBuscaCombo()", "Classe Sistema")

End Sub

Public Sub gsLimparListView(pListView As ListView)

    pListView.ListItems.Clear

End Sub

Public Sub gsLimparTreeView(pTreeView As TreeView)

    pTreeView.Nodes.Clear

End Sub

Public Function gflUlt_Lin_Grid(pGrid As MSFlexGrid) As Long

    Dim pCont As Long

    For pCont = 0 To pGrid.Rows - 1
        pGrid.Row = pCont: pGrid.Col = 0
        If pGrid.Text = "" Then gflUlt_Lin_Grid = pCont: Exit Function
    Next pCont

    pGrid.AddItem "": gflUlt_Lin_Grid = pGrid.Rows - 1

End Function

Public Sub gsLoadForm(pForm As Form)

    pForm.Left = ((Screen.Width - pForm.Width) / 2) + 1000
    pForm.Top = ((Screen.Height - pForm.Height) / 2)
    pForm.Show vbModal

End Sub

Public Sub gsProcurar_No(pNode As Node, pPar As Boolean)

    Dim pCont       As Long
    Dim pNodeAux    As Node

    pNode.Checked = pPar
    If pNode.Children > 0 Then
        Set pNodeAux = pNode.Child
        For pCont = 1 To pNode.Children
            Call gsProcurar_No(pNodeAux, pPar)
            Set pNodeAux = pNodeAux.Next
        Next pCont
    End If

End Sub

Public Function gfsGetIni(pIniFile As String, pSection As String, pKey As String, pDefault As String) As String

    Dim pTemp As String * 256
    Dim pLength As Integer

    On Error GoTo err_gfsGetIni:

    pTemp = Space$(256)
    pLength = GetPrivateProfileString(pSection, pKey, pDefault, pTemp, 255, pIniFile)
    gfsGetIni = Left$(pTemp, pLength)

    Exit Function

err_gfsGetIni:
    Call objSystem.gsExibeErros(Err, "gfsGetIni()", "Classe Sistema")

End Function

Public Sub gfsWriteIni(pIniFile As String, pSection As String, pKey As String, pValue As String)

    Dim pTemp As String
    Dim pCont As Integer

    On Error GoTo err_gfsWriteIni:

    pTemp = pValue
    For pCont = 1 To Len(pValue)
        If Mid$(pValue, pCont, 1) = vbCr Or Mid$(pValue, pCont, 1) = vbLf Then Mid$(pValue, pCont) = " "
    Next pCont

    pCont = WritePrivateProfileString(pSection, pKey, pTemp, pIniFile)

    Exit Sub

err_gfsWriteIni:
    Call gsExibeErros(Err, "gfsWriteIni()", "Classe Sistema")

End Sub

Public Function gfdDiaUtil(pData As Date) As Date

    Dim pWeekday As Integer

    pWeekday = Weekday(pData, vbSunday)
    Select Case pWeekday
        Case 1
            gfdDiaUtil = DateAdd("d", 1, pData)
        Case 7
            gfdDiaUtil = DateAdd("d", 2, pData)
        Case Else
            gfdDiaUtil = pData
    End Select

End Function

Public Sub gsDesmarcar_ListView(pListView As ListView)

    Dim pCont As Integer

    With pListView
        With .ListItems
            For pCont = 1 To .Count
                .Item(pCont).Checked = False
            Next pCont
        End With
    End With

End Sub

Public Function gfbValidarCPF(pCPF As String) As Boolean

    On Error GoTo err_gfbValidarCPF:

    Dim pSoma As Integer
    Dim pResto As Integer
    Dim pContador As Integer

    ''Valida argumento
    If Len(pCPF) <> 11 Then

        gfbValidarCPF = False
        Exit Function

    End If

    If pCPF = "00000000000" Then

        gfbValidarCPF = False
        Exit Function

    End If

    pSoma = 0

    For pContador = 1 To 9
        pSoma = pSoma + Val(Mid$(pCPF, pContador, 1)) * (11 - pContador)
    Next pContador

    pResto = 11 - (pSoma - (Int(pSoma / 11) * 11))

    If pResto = 10 Or pResto = 11 Then pResto = 0

    If pResto <> Val(Mid$(pCPF, 10, 1)) Then

        gfbValidarCPF = False
        Exit Function

    End If

    pSoma = 0

    For pContador = 1 To 10
        pSoma = pSoma + Val(Mid$(pCPF, pContador, 1)) * (12 - pContador)
    Next pContador
    pResto = 11 - (pSoma - (Int(pSoma / 11) * 11))

    If pResto = 10 Or pResto = 11 Then pResto = 0

    If pResto <> Val(Mid$(pCPF, 11, 1)) Then

        gfbValidarCPF = False
        Exit Function

    End If

    gfbValidarCPF = True

    Exit Function

err_gfbValidarCPF:
    Call gsExibeErros(Err, "gfbValidarCPF()", "Classe Sistema")

End Function

Public Function gfsEncryptString(Text As String, Action As Single) As String

    On Error GoTo err_gfsEncryptString:

    ' Define as variaveis usadas
    Dim UserKeyX As String
    Dim Temp     As Integer
    Dim Times    As Integer
    Dim i        As Integer
    Dim j        As Integer
    Dim n        As Integer
    Dim rtn      As String

    ' Obtem os caracteres da chave do usuário
    ' define o comprimento da chave do usuario usada na criptografia
    n = Len(csChave)

    ' Redimensiona o array para o tamanho definido
    ReDim userKeyASCIIS(1 To n)

    ' Preenche o array com caracteres asc
    Debug.Print csChave; "=> ";
    For i = 1 To n

        userKeyASCIIS(i) = Asc(Mid$(csChave, i, 1))
        Debug.Print userKeyASCIIS(i); " ";

    Next

    ' Redimensiona o array com o tamanho do texto
    ' Obtem o caractere de texto
    ReDim TEXTAsciis(Len(Text)) As Integer

    ' Preenche o array com caracteres asc
    Debug.Print
    Debug.Print Text; " => ";

    For i = 1 To Len(Text)

        TEXTAsciis(i) = Asc(Mid$(Text, i, 1))
        Debug.Print TEXTAsciis(i); " ";

    Next

    ' Cifra / Decifra
    If Action = ciEncrypt Then

       For i = 1 To Len(Text)

           j = IIf(j + 1 >= n, 1, j + 1)

           Temp = TEXTAsciis(i) + userKeyASCIIS(j)

           If Temp > 255 Then Temp = Temp - 255

           Debug.Print Temp; " ";

           rtn = rtn + Chr$(Temp)

           Debug.Print rtn

       Next

    ElseIf Action = ciDecrypt Then

       For i = 1 To Len(Text)

           j = IIf(j + 1 >= n, 1, j + 1)

           Temp = TEXTAsciis(i) - userKeyASCIIS(j)

           If Temp < 0 Then Temp = Temp + 255

           rtn = rtn + Chr$(Temp)

       Next

    End If

    ' Retorna o texto
    gfsEncryptString = rtn

    Exit Function

err_gfsEncryptString:
    Call objSystem.gsExibeErros(Err, "gfsEncryptString()", "Classe Sistema")

End Function
