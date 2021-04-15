<%
'----------------------------------------'
'	LOGIN CREDENTIAL SETTING			 '
'	========================			 '
'										 '
'	Please modify ONLY THIS SECTION	 	 '
'	to set the admin security 			 '
'	credential							 '
'										 '
'----------------------------------------'

' SET USERNAME HERE. Your username has to be within quote. 
Username = "admin"

' SET PASSWORD HERE. Your password has to be within quote.
Password = "admin"

' SET EMAIL HERE. Your email has to be within quote.
' The email address will be use for "Forget Password" feature
AdminEmailAddress = "admin@domain.com"


' EMAIL COMPONENT SETTING
'=========================

' SET EMAIL COMPONENT 
' Please choose the Email Server Comonent that is supported by your webhosting. You will have to contact your webhosting to confirm about this

' Options:
' 1 : CDOSYS Email Components
' 2 : CDONTS Email Components
EmailServerComponent = 2


' ====== CDOSYS CONFIGURATION ====== 
' Modify this section ONLY IF you select 1 (CDOSYS Email Components)
' If you don't have this information, please contact your webhosting in order to set this up properly.
EmailServer = "mail.yourdomain.com"

' Modify this section ONLY IF your webhosting requires SMTP Authentication
' Options:
' 1 : YES
' 2 : NO
RequireSMTPAuthentication = 2

' Modify this section ONLY IF your webhosting requires SMTP Authentication
EmailUsername = "emailUserName" ' your username for SMTP Authentication
EmailPassword = "emailPassword" ' your password for SMTP Authentication

' ====== END OF CDOSYS CONFIGURATION =====

'----------------------------------------'
'										 '
'	END OF CREDENTIAL SETTING			 '
'	=========================			 '
'	DO NOT MODIFY ANYTHING BEYOND THIS!	 '
'										 '
'----------------------------------------'

%> 



<%
Session.Timeout = 360

' =========================================================================================='
' *** Validate request to log in to this site.

MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Request.QueryString
MM_valUsername = CStr(Request.Form("username"))
MM_valPassword = Cstr(Request.Form("password"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization= "1"
  MM_redirectLoginSuccess="index.asp"
  MM_redirectLoginFailed= "?valid=false"
    If MM_valUsername = CStr(username) And MM_valPassword = password Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    ' START - create remember me cookie
IF (Request.Form("rememberme") = "yes") then
          Response.Cookies("remember")("username") =  Request.Form("username")
          Response.Cookies("remember")("password") = Request.Form("password")
		  Response.Cookies("remember")("rememberme") = "yes"
          Response.Cookies("remember").Expires = date + 90
		  Else
		  Response.Cookies("remember").Expires = date -1
End If 
	' END - create remember me cookie 
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_fldUserAuthorization)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And true Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<% If Request.Form("SendPassword") = "SendPassword" Then %>
<%
If Request.Form("EmailAddress") = CStr(adminemailaddress) Then
' *** Redirect If Recordset Is Empty
' *** MagicBeat Server Behavior - 2014 - by Jag S. Sidhu - www.magicbeat.com
	
	Dim mailSubject, mailBody
	mailSubject = "Login Credentials"
	mailBody = "Login Credentials:" & vbCrLf _
		& vbCrLf _
		& vbCrLf _ 
		& "Email Address:  " &  adminemailaddress & vbCrLf _ 
		& "User Name:  " & username & vbCrLf _
		& "Password:  " & password & vbCrLf _ 
		& vbCrLf _
		& "Login @ " & Request.ServerVariables("HTTP_REFERER")&  vbCrLf _
		& vbCrLf _
		& vbCrLf

	If EmailServerComponent = 2 Then ' CDONTS
	
		'Create the mail object and send the mail
		Set objMail = Server.CreateObject("CDONTS.NewMail") 'chenged from CreateObject("CDONTS.NewMail")
		objMail.From = adminemailaddress
		objMail.To = adminemailaddress
		objMail.CC = ""
		objMail.BCC = ""
		objMail.Subject = mailSubject
		objMail.Body = 	mailBody
		objMail.Send()
		Set objMail = Nothing
	End If
	
	If EmailServerComponent = 1 Then 'CDOSYS
		theSchema="http://schemas.microsoft.com/cdo/configuration/" 
		Set cdoConfig=server.CreateObject("CDO.Configuration") 
		cdoConfig.Fields.Item(theSchema & "sendusing")= 2
		cdoConfig.Fields.Item(theSchema & "smtpserver")= EmailServer ' use your smtp server name
		
		' Check if it requires SMTP or not
		If RequireSMTPAuthentication = 1 And EmailUsername <> "" And EmailPassword <> "" Then
		response.Write("ok")
			cdoConfig.Fields.Item(theSchema & "smtpauthenticate")= 1
			cdoConfig.Fields.Item(theSchema & "sendusername")= EmailUsername
			cdoConfig.Fields.Item(theSchema & "sendpassword")= EmailPassword
		End If
		cdoConfig.Fields.Update
		
		Set objMail = Server.CreateObject("CDO.Message")
		objMail.Configuration = cdoConfig 
		objMail.From = adminemailaddress  'The mail is sent from the address declared in the variable above
		objMail.To = adminemailaddress 'The mail is sent from the address declared in the variable above
		objMail.Subject = mailSubject
		objMail.TextBody =  mailBody
		objMail.Send 'Send the email! 
		
		set objMail = Nothing
	End If
	
    'Send them to the page specified if requested
Dim rp_redirectpw
      If Request.QueryString <> "" Then
      rp_redirectpw = Request.ServerVariables("HTTP_REFERER") & "&sent=true"
    Else
      rp_redirectpw = Request.ServerVariables("HTTP_REFERER") & "?sent=true"
    End If
	Response.Redirect rp_redirectpw
	
Else
    'Send them to the page specified if requested
Dim rp_redirectpwn
      If Request.QueryString <> "" Then
      rp_redirectpwn = Request.ServerVariables("HTTP_REFERER") & "&sent=false"
    Else
      rp_redirectpwn = Request.ServerVariables("HTTP_REFERER") & "?sent=false"
    End If
	Response.Redirect rp_redirectpwn
	
	End If 
%>
<% end if%>

     
<form name="login" method="POST" action="<%=MM_LoginAction%>" class="form-signin">
        <% IF (Request.Querystring("valid") <> "false") then %>
  <% IF NOT Request.Querystring("sent") <> "" then %>
        <div class="form-group">  <label>Username</label>
        <input type="text" class="form-control" placeholder="Username" required="" autofocus="" name="username" id="username" value="admin"></div>
         <label>Password</label>
        <div class="form-group"> <input type="password" class="form-control" placeholder="Password" required="" name="password" id="password" value="admin"></div>
        <label class="checkbox">
          <input type="checkbox"  name="rememberme" id="rememberme"  <% IF (Request.Cookies("remember")("rememberme") = "yes") then %> checked <%end if%> value="yes"> Remember me
        </label>
        <button class="btn btn-lg btn-primary btn-block" type="submit">Sign in</button>
        
      
        <% end if%>
<% end if%>
</form>

<% IF (Request.Querystring("valid") = "false") then %>
        <font color="#FF0000"><b>You entered either an incorrect Username
        or Password - please <a href="javascript:history.go(-1)">try again</a></b></font>
        <% End If %>
<% IF NOT Request.Querystring("sent") <> "" then %>
<% IF (Request.Querystring("valid") <> "false") then %>
<br />


 <p><a data-toggle="collapse" data-target="#sendpw" href="#">Forgot your password?</a></p>
<div id="sendpw" class="collapse">
<form action="" method="post" name="email">
 
    <input name="EmailAddress" type="hidden" id="EmailAddress" size="30" value="<%=adminemailaddress%>">
    <input name="Submit2" type="submit" value="Click to send password to <%=adminemailaddress%>" class="btn btn-default btn-sm">
    <input name="SendPassword" type="hidden" id="SendPassword" value="SendPasswordNo">
 
</form></div>
<%end if%>
<%end if%>
        <% IF (Request.Querystring("sent") = "false") then %>
      <h2><font color="#FF0000">OOPS</font></h2>
      <p>Sorry but the email address  you entered         was either spelled
        incorrectly or has never been registered. <br>
        Email not in database please <font color="#FF0000"><b><a href="javascript:history.go(-1)">try
        again</a></b></font> or contact system administrator. 
        <% End If %>
      </p>
      <% IF Request.Querystring("sent") then %>
      <h1>Success!</h1>
      <p>Thank you , your password has been sent....<font color="#FF0000"><b><a href="javascript:history.go(-1)">Login</a></b></font></p>
      <% End If %>
