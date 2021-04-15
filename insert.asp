<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="admin_security.asp" -->
<!--#include file="Connections/linksmanager.asp" -->
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_linksmanager_STRING
    MM_editCmd.CommandText = "INSERT INTO Links (ItemName, ItemUrl, CategoryID, ItemDesc, Activated) VALUES (?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 100, Request.Form("LinkName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 100, Request.Form("LinkUrl")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("CategoryID"), Request.Form("CategoryID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 203, 1, 536870910, Request.Form("LinkDescription")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 5, Request.Form("Activated")) ' adVarWChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "index.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If
End If
%>
<%
set list = Server.CreateObject("ADODB.Recordset")
list.ActiveConnection = MM_linksmanager_STRING
list.Source = "SELECT *  FROM Links"
list.CursorType = 0
list.CursorLocation = 2
list.LockType = 3
list.Open()
list_numRows = 0
%>
<%
set Category = Server.CreateObject("ADODB.Recordset")
Category.ActiveConnection = MM_linksmanager_STRING
Category.Source = "SELECT *  FROM LinksCategory ORDER BY CategoryID"
Category.CursorType = 0
Category.CursorLocation = 2
Category.LockType = 3
Category.Open()
Category_numRows = 0
%>
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="">
    <meta name="author" content="">
    <title>Insert</title>
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}
function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') {
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (val<min || max<val) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<link href="bootstrap-3.0.2/css/bootstrap.min.css" rel="stylesheet" type="text/css">
</head>
  <div class="container">
<!--#include file="header.asp" -->
    

      <form method="POST" action="<%=MM_editAction%>" name="form1" class="form-horizontal">
 <label>Insert New Link</label><hr>
 <div class="form-group"><label class="col-sm-2">Link Name:</label>
           <div class="col-sm-10"><input type="text" class="form-control" name="LinkName" value="" required>
           <span class="help-block">Enter a short name to identify the link i.e <em>City Hall.</em></span>           </div>
 </div>
            <div class="form-group">
           <label class="col-sm-2">Link Url:</label>
             <div class="col-sm-10"> <input type="text" name="LinkUrl" class="form-control" value="" required>
              <span class="help-block">Enter the actual URL link i.e.<em> http://www.cityhall.com.</em></span></div>
            </div>
 
      <div class="form-group">
              <label class="col-sm-2">Category:</label>
            <div class="col-sm-10">  <select name="CategoryID" class="form-control" required="required">
                <%
While (NOT Category.EOF)
%>
                <option value="<%=(Category.Fields.Item("CategoryID").Value)%>"><%=(Category.Fields.Item("CategoryName").Value)%></option>
                <%
  Category.MoveNext()
Wend
If (Category.CursorType > 0) Then
  Category.MoveFirst
Else
  Category.Requery
End If
%>
              </select>
<span class="help-block"><a href="javascript:;" onClick="MM_openBrWindow('add_category.asp','Category','scrollbars=yes,width=400,height=300')">add/edit
      category</a>  </span>     
       </div></div>
       
           <div class="form-group">
            <label class="col-sm-2">Link Description:</label>
            <div class="col-sm-10">  <textarea name="LinkDescription" class="form-control" rows="3" required></textarea>
              <span class="help-block">Enter a description of this link i.e. <em>This is a link to the official website of City Hall.</em></span></div>
           </div>
              
            <div class="form-group">
              <label class="col-sm-2">Activated: </label>
             <div class="col-sm-10">
             <input type="checkbox" name="Activated" value="True">
             <p class="help-block">Check if you want this link to be visible to the public. Uncheck if you wish to hide.</p> 
             </div></div>
  
   <div class="form-group">
             <label class="col-sm-2"></label>
             <div class="col-sm-10">
      <div class="pull-right">
  <button type="submit" class="btn btn-primary"> Publish to Links Listing Page </button>      
  <button type="button" class="btn btn-default" onClick="myFunction()">Cancel</button>
</div>
<script>
function myFunction()
{
window.open("index.asp", "_self");
}
</script>

<input type="hidden" name="MM_insert" value="form1">
   </div></div>   </form>

<!--#include file="footer.asp" --> 
   </div>
<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
    <script src="https://code.jquery.com/jquery.js"></script>
    <!-- Include all compiled plugins (below), or include individual files as needed -->
    <script src="bootstrap-3.0.2/js/bootstrap.min.js"></script>
</body>
</html>
<%
list.Close()
%>
<%
Category.Close()
%>



