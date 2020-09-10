Attribute VB_Name = "M�dulo1"
Sub salvar_pdf()
    
    ''Exporta a planilha OPP em PDF e salva na pasta especificada
    Dim FileName    As String
    Dim data_atual  As String
    Dim Assunto     As String
    Dim Emails      As String
    Dim corpo_email As String
    Dim Assinatura As String
    Dim tempoEmail As String
    
        
    ''Armazena nas vari�veis os valores do dia atual
    data_atual = Format(Now(), "dd-mm-yyyy")
    
    '' Cria o endere�o para salvar o arquivo (nome da pasta + data atual)
    FileName = "G:\Operacoes.CORP\03- Dashboards\OPR - One Page Report\historico\OPR - One Page Report_" & CStr(data_atual)
    
    'Salva o arquivo em PDF
    Sheets("OPR").ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
    FileName _
    , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
    :=False, OpenAfterPublish:=False
    
    '' Definindo os par�metros do e-mail
    Emails = Sheets("controle").Range("B2").Value
    corpo_email = Sheets("controle").Range("B3")
    Assunto = Sheets("controle").Range("B4")
    Assinatura = Sheets("controle").Range("B5")
    tempoEmail = Now()
        
    '' Chama a fun��o de envio de e-mail
    Call sendEmail(Assunto, Emails, corpo_email, FileName, Assinatura)
    
    
    '' Chama o pr�ximo agendamento
    Call Agendamento
    
    '' Registrando a data de envio do e-mail
    Sheets("controle").Cells(7, 2).Value = tempoEmail
    
    '' Salva a pasta de trabalho atual
    ActiveWorkbook.Save
End Sub

Sub sendEmail(ByVal vSubject As String, ByVal vUser As String, ByVal vBody As String, vAnexo As String, vAssinatura As String)
    
    Dim oApp        As Object
    Dim oMail       As Object
    Dim anexo       As String

    
    
    Set oApp = CreateObject("Outlook.Application")
    Set oMail = oApp.CreateItem(0)
    
    '' Difinindo o anexo
    anexo = vAnexo & ".pdf"
    
  
    Set objOlAppAnexo = oMail.Attachments.Add(anexo)
    
    With oMail
        .to = vUser
        .Subject = vSubject
        .HTMLBody = vBody & vAssinatura
        .Send
    End With
    
    Set oApp = Nothing
    Set oMail = Nothing
    
End Sub
Sub Agendamento()

    '' Definindo as vari�veis
    Dim tempoBase As String
    Dim tempoEmail As String
    Dim TempoAgendamento As String
    
    '' Definindo os tempo de agendamento
    tempoBase = TimeValue("05:30:00")
    tempoEmail = TimeValue("07:30:00")
    TempoAgendamento = Now()
    
    '' Agendando a atualiza��o
    Application.OnTime tempoBase, "AtualizarBase"
    Application.OnTime tempoEmail, "salvar_pdf"
    
    '' Registrar data da pr�xima atualiza��o
    Sheets("controle").Cells(7, 2).Value = TempoAgendamento
    
    
End Sub

Sub AtualizarBase()
    '' Definindo as vari�veis
    Dim tempoBase   As String
    
    '' Atualizar todas as consultas
    ActiveWorkbook.RefreshAll
    
    '' Coletando a data de atualiza��o
    tempoBase = Now()
    
    '' Registrar data da �ltima atualiza��o
    Sheets("controle").Cells(8, 2).Value = tempoBase
    
End Sub
