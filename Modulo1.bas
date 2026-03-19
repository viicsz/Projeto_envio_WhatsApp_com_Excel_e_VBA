Attribute VB_Name = "Módulo1"
' 1. FUNÇÃO COPIA TEXTO PURO (igual)
Sub CopiaTextoPuro(Celula As Range)
    Dim objData As Object
    Set objData = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    objData.SetText Celula.Value
    objData.PutInClipboard
End Sub

' 2. SCRIPT ATUALIZADO (mensagem A2 da Planilha2)
Sub EnviarWhatsApp()
    Dim Lin As Integer
    Dim ws As Worksheet
    Dim ws2 As Worksheet  ' Planilha2
    Dim MinTempo As Integer
    Dim MaxTempo As Integer
    Dim Delay As Integer
    
    Set ws = ActiveSheet        ' Planilha atual (contatos)
    Set ws2 = Worksheets("Mensagem")  ' Mensagem (mensagem)
    Randomize
    
    MinTempo = ws.Cells(2, 5).Value  ' E2 atual
    MaxTempo = ws.Cells(2, 6).Value  ' F2 atual
    Lin = 5
    
    Shell "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://web.whatsapp.com"     ' Utilizando o Edge, caso queira utilizar o Chrome, precisa alterar o caminho
    Application.Wait Now + TimeValue("00:00:30")
    
    Do Until ws.Cells(Lin, 2).Value = ""
        ' COPIA CONTATO B (igual)
        CopiaTextoPuro ws.Cells(Lin, 2)
        Application.Wait Now + TimeValue("00:00:01")
        Application.SendKeys "{TAB}"
        Application.Wait Now + TimeValue("00:00:01")
        Application.SendKeys "{TAB}"
        Application.Wait Now + TimeValue("00:00:01")
        Application.SendKeys "{TAB}"
        Application.Wait Now + TimeValue("00:00:01")
        Application.SendKeys "{TAB}"
        Application.Wait Now + TimeValue("00:00:01")
        Application.SendKeys "^v"
        Application.Wait Now + TimeValue("00:00:02")
        Application.SendKeys "{TAB}"
        Application.Wait Now + TimeValue("00:00:01")
        Application.SendKeys "{TAB}"
        Application.Wait Now + TimeValue("00:00:01")
        Application.SendKeys "~"
        Application.Wait Now + TimeValue("00:00:05")
        
        ' MENSAGEM A2 da Planilha2!
        CopiaTextoPuro ws2.Cells(2, 1)     ' ? Planilha2!A2
        Application.Wait Now + TimeValue("00:00:01")
        Application.SendKeys "^v"
        Application.Wait Now + TimeValue("00:00:01")
        Application.SendKeys "~"
        
        Application.SendKeys "{TAB}"
        Application.Wait Now + TimeValue("00:00:01")
        
        Delay = Int((MaxTempo - MinTempo + 1) * Rnd + MinTempo)
        Application.Wait Now + TimeValue("00:00:" & Format(Delay, "00"))
        
        Lin = Lin + 1
    Loop
    
    MsgBox "Mensagens enviadas! (Delay entre mensagens: " & MinTempo & "-" & MaxTempo & "s)"
End Sub
