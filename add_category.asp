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
  '  Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_linksmanager_STRING
    MM_editCmd.CommandText = "INSERT INTO LinksCategory (CategoryName) VALUES (?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("CategoryName")) ' adVarWChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    'Dim MM_editRedirectUrl
    MM_editRedirectUrl = "closewindow_redirect.asp"
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
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) = "delete" And CStr(Request("MM_recordId")) <> "") Then

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_linksmanager_STRING
    MM_editCmd.CommandText = "DELETE FROM LinksCategory WHERE CategoryID = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    'Dim MM_editRedirectUrl
    MM_editRedirectUrl = "closewindow_redirect.asp"
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
If (CStr(Request("MM_update")) = "form") Then
  If (Not MM_abortEdit) Then
    ' execute the update
 '   Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_linksmanager_STRING
    MM_editCmd.CommandText = "UPDATE LinksCategory SET CategoryName = ? WHERE CategoryID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("CategoryName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    'Dim MM_editRedirectUrl
    MM_editRedirectUrl = "closewindow_redirect.asp"
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
Dim Category
Dim Category_cmd
Dim Category_numRows

Set Category_cmd = Server.CreateObject ("ADODB.Command")
Category_cmd.ActiveConnection = MM_linksmanager_STRING
Category_cmd.CommandText = "SELECT * FROM LinksCategory ORDER BY CategoryID" 
Category_cmd.Prepared = true

Set Category = Category_cmd.Execute
Category_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
Category_numRows = Category_numRows + Repeat1__numRows
%>
<%
' UltraDeviant - Row Number written by Owen Palmer (http://ultradeviant.co.uk)
Dim OP_RowNum
If MM_offset <> "" Then
	OP_RowNum = MM_offset + 1
Else
	OP_RowNum = 1
End If
%>
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="">
    <meta name="author" content="">
<title>Category</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="bootstrap-3.0.2/css/bootstrap.min.css" rel="stylesheet" type="text/css">
</head>
<body>
<div class="table-responsive">
<table width="100%"border="0" cellpadding="0" cellspacing="0" class="table">
  <tr valign="middle"> 
    <td colspan="4" class="tableheader"><table width="100%" border="0" cellspacing="0" cellpadding="5">
        <tr>
          <td valign="baseline" nowrap><strong>Add New Category</strong>
            <form method="POST" action="<%=MM_editAction%>" name="form1" class="form">
              <input type="text" name="CategoryName" value="">
              <input type="submit" value="Insert New" name="submit" class="btn btn-xs btn-default">
              <input type="hidden" name="MM_insert" value="form1">
            </form></td>
        </tr>
    </table>
    </td>
  </tr>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT Category.EOF)) 
%>
  <tr class="<% 
RecordCounter = RecordCounter + 1
If RecordCounter Mod 2 = 1 Then
Response.Write "row1"
Else
Response.Write "row2"
End If
%>"> 
    <td width="2%" valign="baseline"><b> 
    <%Response.Write(RecordCounter)
RecordCounter = RecordCounter%>.</b>    </td>
    <td width="80%" valign="middle"><form name="form" method="POST" action="<%=MM_editAction%>">      
        <div align="left">
          <input name="CategoryName" type="text" id="CategoryName" value="<%=(Category.Fields.Item("CategoryName").Value)%>">
          <input type="hidden" name="MM_update" value="form">
          <input type="hidden" name="MM_recordId" value="<%= Category.Fields.Item("CategoryID").Value %>">
          <input type="submit" name="Submit" value="Update" class="btn btn-xs btn-default">
        </div>
    </form>
    </td>
    <td width="9%" valign="baseline"><form ACTION="<%=MM_editAction%>" METHOD="POST" name="delete">
        <div align="left">
          <input type="submit" name="Submit" value="Delete" class="btn btn-danger btn-xs">
          <input type="hidden" name="MM_delete" value="delete">
<input type="hidden" name="MM_recordId" value="<%= Category.Fields.Item("CategoryID").Value %>">
      </div>
    </form></td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Category.MoveNext()
Wend
%>
</table>
</div>
<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
    <script src="https://code.jquery.com/jquery.js"></script>
    <!-- Include all compiled plugins (below), or include individual files as needed -->
    <script src="bootstrap-3.0.2/js/bootstrap.min.js"></script>
</body>
</html>
<%
Category.Close()
%>
