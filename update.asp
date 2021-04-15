<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="admin_security.asp" -->
<!--#include file="Connections/linksmanager.asp" -->
<%
Dim List__update
List__update = "1000"
If (Request.QueryString("ItemID")  <> "") Then 
  List__update = Request.QueryString("ItemID") 
End If
%>
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
If (CStr(Request("MM_update")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_linksmanager_STRING
    MM_editCmd.CommandText = "UPDATE Links SET ItemName = ?, ItemUrl = ?, ItemDesc = ?, CategoryID = ?, Activated = ? WHERE ItemID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 100, Request.Form("LinkName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 100, Request.Form("LinkUrl")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 203, 1, 536870910, Request.Form("LinkDescription")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("CategoryID"), Request.Form("CategoryID"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 5, Request.Form("Activated")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
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
set List = Server.CreateObject("ADODB.Recordset")
List.ActiveConnection = MM_linksmanager_STRING
List.Source = "SELECT Links.*, LinksCategory.CategoryName, LinksCategory.ParentCategoryIDkey, LinksCategory.CategoryDesc  FROM LinksCategory RIGHT JOIN Links ON LinksCategory.CategoryID = Links.CategoryID  WHERE ItemID = " + Replace(List__update, "'", "''") + "  ORDER BY Links.DateAdded"
List.CursorType = 0
List.CursorLocation = 2
List.LockType = 3
List.Open()
List_numRows = 0
%>
<%
set CategoryListUpdate = Server.CreateObject("ADODB.Recordset")
CategoryListUpdate.ActiveConnection = MM_linksmanager_STRING
CategoryListUpdate.Source = "SELECT *  FROM LinksCategory ORDER BY CategoryID"
CategoryListUpdate.CursorType = 0
CategoryListUpdate.CursorLocation = 2
CategoryListUpdate.LockType = 3
CategoryListUpdate.Open()
CategoryListUpdate_numRows = 0
%>
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="">
    <meta name="author" content="">
<title>Update</title>

<link href="bootstrap-3.0.2/css/bootstrap.min.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>


</head>
<body>
  <div class="container">
<!--#include file="header.asp" -->
</div>
<div class="container">


            
          

      <form method="POST" action="<%=MM_editAction%>" name="form1" class="form-horizontal">
      <label>Update Existing Link </label><hr>
      <div class="form-group"><label class="col-sm-2">Link Name:</label>
           <div class="col-sm-10"><input type="text" class="form-control" name="LinkName" value="<%=(List.Fields.Item("ItemName").Value)%>">
           <span class="help-block">Enter a short name to identify the link i.e &quot;City Hall&quot;.</span>           </div>
      </div>
            <div class="form-group">
           <label class="col-sm-2">Link Url:</label>
             <div class="col-sm-10"> <input type="text" name="LinkUrl" class="form-control" value="<%=(List.Fields.Item("ItemUrl").Value)%>">
              <span class="help-block">Enter the actual URL link i.e. http://www.cityhall.com" | <a href="<%=(List.Fields.Item("ItemUrl").Value)%>" target="_blank">Test Link</a> </span></div>
        </div>
               <div class="form-group">
            <label class="col-sm-2">Link Description:</label>
            <div class="col-sm-10">  <textarea name="LinkDescription" class="form-control" rows="3"><%=(List.Fields.Item("ItemDesc").Value)%></textarea>
              <span class="help-block">Enter a description of this link i.e. &quot;This is a link to the official website of City Hall"</span></div>
        </div>
               <div class="form-group">
              <label class="col-sm-2">Category:</label>
            <div class="col-sm-10">  <select name="CategoryID" class="form-control">
               <%
While (NOT CategoryListUpdate.EOF)
%>
                <option value="<%=(CategoryListUpdate.Fields.Item("CategoryID").Value)%>" <%if (CStr(CategoryListUpdate.Fields.Item("CategoryID").Value) = CStr(List.Fields.Item("CategoryID").Value)) then Response.Write("SELECTED") : Response.Write("")%> ><%=(CategoryListUpdate.Fields.Item("CategoryName").Value)%></option>
                <%
  CategoryListUpdate.MoveNext()
Wend
If (CategoryListUpdate.CursorType > 0) Then
  CategoryListUpdate.MoveFirst
Else
  CategoryListUpdate.Requery
End If
%>
              </select>
             <span class="help-block"><a href="javascript:;" onClick="MM_openBrWindow('add_category.asp','Category','scrollbars=yes,width=400,height=300')">add/edit
      category</a>  </span>       </div>
        </div>
        <div class="form-group">
              <label class="col-sm-2">Activated: </label>
             <div class="col-sm-10"><input type="checkbox" name="Activated" <%If (CStr(List.Fields.Item("Activated").Value) = CStr("True")) Then Response.Write("checked") : Response.Write("")%> value="True">
             <p class="help-block">Check if you want this link to be visible to the public. Uncheck if you wish to hide</p> </div>
        </div> <div class="form-group">
            <label class="col-sm-2">Date Added:</label>
            <div class="col-sm-10"><p class="help-block"><%=(List.Fields.Item("DateAdded").Value)%> </p>
           </div>
            </div>
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
            
<input type="hidden" name="MM_update" value="form1">
<input type="hidden" name="MM_recordId" value="<%= List.Fields.Item("ItemID").Value %>">
</div>
        </div>
      </form>
<!--#include file="footer.asp" --> 
</div>

<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
    <script src="https://code.jquery.com/jquery.js"></script>
    <!-- Include all compiled plugins (below), or include individual files as needed -->
    <script src="bootstrap-3.0.2/js/bootstrap.min.js"></script>
</body>
</html>
<%
List.Close()
%>
<%
CategoryListUpdate.Close()
%>
