' Codigo ASP para envio de formulário do site da extinta "Expresso Tempo Real Ltda" disponibilizao
' pela "InSite" (atual UOL Host).

<%
        strNome = Request.Form("nome")
        strEmail = Request.Form("email")
        strDdd = Request.Form("ddd")
		strTelefone = Request.Form("telefone")
   		strAssunto = Request.Form("assunto")
		strMensagem1 = Request.Form("obs")
    
        Set objCDOSYSMail = Server.CreateObject("CDO.Message")
        
        Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")
        
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
        
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
        
        objCDOSYSCon.Fields.update
        
        Set objCDOSYSMail.Configuration = objCDOSYSCon
        
        objCDOSYSMail.From = strEmail
        
        objCDOSYSMail.To = "matriz@expressotemporeal.com.br"
        
        objCDOSYSMail.Subject = "Contato via site"
        
        'objCDOSYSMail.TextBody = "Teste do componente CDOSYS"
        
        strMensagem = strMensagem & "<table width='400' border='0' cellpadding='8' cellspacing='1' bgcolor='#CCCCCC'>"
        strMensagem = strMensagem & "<tr>"
        strMensagem = strMensagem & "<td bgcolor='#FFFFFF'><table width='400' border='0' cellspacing='0' cellpadding='2'>"
        strMensagem = strMensagem & "<tr>"
        strMensagem = strMensagem & "<td height='35' bgcolor='EFEFEF'><b><font size='4' face='Verdana, Arial, Helvetica, sans-serif'>Conteudo da mensagem</font></b></td>"
        strMensagem = strMensagem & "</tr>"
        strMensagem = strMensagem & "<tr>"
        strMensagem = strMensagem & "<td height='20'><font size='2' face='Verdana, Arial, Helvetica, sans-serif'><b>Nome</b></font></td>"
        strMensagem = strMensagem & "</tr>"
        strMensagem = strMensagem & "<tr>"
        strMensagem = strMensagem & "<td height='20'><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>" & strNome & "</font></td>"
        strMensagem = strMensagem & "</tr>"
        strMensagem = strMensagem & "<tr>"
        strMensagem = strMensagem & "<td height='20'><b><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>E-mail</font></b></td>"
        strMensagem = strMensagem & "</tr>"
        strMensagem = strMensagem & "<tr>"
        strMensagem = strMensagem & "<td height='20'><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>" & strEmail & "</font></td>"
        strMensagem = strMensagem & "</tr>"
        strMensagem = strMensagem & "<tr>"
        strMensagem = strMensagem & "<td height='20'><b><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>DDD</font></b></td>"
        strMensagem = strMensagem & "</tr>"
        strMensagem = strMensagem & "<tr>"
        strMensagem = strMensagem & "<td height='20'><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>" & strDdd & "</font></td>"
        strMensagem = strMensagem & "</tr>"
		strMensagem = strMensagem & "<tr>"
        strMensagem = strMensagem & "<td height='20'><b><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>Telefone</font></b></td>"
        strMensagem = strMensagem & "</tr>"
        strMensagem = strMensagem & "<tr>"
        strMensagem = strMensagem & "<td height='20'><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>" & strTelefone & "</font></td>"
        strMensagem = strMensagem & "</tr>"
		strMensagem = strMensagem & "<tr>"
        strMensagem = strMensagem & "<td height='20'><b><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>Assunto</font></b></td>"
        strMensagem = strMensagem & "</tr>"
        strMensagem = strMensagem & "<tr>"
        strMensagem = strMensagem & "<td height='20'><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>" & strAssunto & "</font></td>"
        strMensagem = strMensagem & "</tr>"
        strMensagem = strMensagem & "<tr>"
        strMensagem = strMensagem & "<td height='20'><font size='2' face='Verdana, Arial, Helvetica, sans-serif'><b>Mensagem</b></font></td>"
        strMensagem = strMensagem & "</tr>"
        strMensagem = strMensagem & "<tr>"
        strMensagem = strMensagem & "<td height='20'><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>" & strMensagem1 & "</font></td>"
        strMensagem = strMensagem & "</tr>"
        strMensagem = strMensagem & "</table></td>"
        strMensagem = strMensagem & "</tr>"
        strMensagem = strMensagem & "</table>"
  
        objCDOSYSMail.HtmlBody = strMensagem
        
        'objCDOSYSMail.fields.update

        objCDOSYSMail.Send
        
        Set objCDOSYSMail = Nothing
        Set objCDOSYSCon = Nothing
        
        Response.Write("<script language=""JavaScript"">")
        'Response.Write("alert('Sua mensagem foi enviada com sucesso');")
        Response.Write("location.href = ""resposta.htm"" ")
        Response.Write("</script>")
%>
