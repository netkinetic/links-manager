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
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_linksmanager_STRING
    MM_editCmd.CommandText = "DELETE FROM Links WHERE ItemID = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "index.asp"
   ' If (Request.QueryString <> "") Then
     ' If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
      '  MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    '  Else
     '   MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
     ' End If
   ' End If
    Response.Redirect(MM_editRedirectUrl)
  End If

End If
%>
<%
Dim deletelistitem__MMColParam
deletelistitem__MMColParam = "1000"
If (Request.QueryString("ItemID")      <> "") Then 
  deletelistitem__MMColParam = Request.QueryString("ItemID")     
End If
%>
<%
Dim deletelistitem
Dim deletelistitem_cmd
Dim deletelistitem_numRows

Set deletelistitem_cmd = Server.CreateObject ("ADODB.Command")
deletelistitem_cmd.ActiveConnection = MM_linksmanager_STRING
deletelistitem_cmd.CommandText = "SELECT * FROM Links WHERE ItemID = ?" 
deletelistitem_cmd.Prepared = true
deletelistitem_cmd.Parameters.Append deletelistitem_cmd.CreateParameter("param1", 5, 1, -1, deletelistitem__MMColParam) ' adDouble

Set deletelistitem = deletelistitem_cmd.Execute
deletelistitem_numRows = 0
%>
<body onLoad="document.form1.submit()">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
  <input type="hidden" name="MM_delete" value="form1">
<input type="hidden" name="MM_recordId" value="<%= deletelistitem.Fields.Item("ItemID").Value %>">
</form>
<%
deletelistitem.Close()
%>
