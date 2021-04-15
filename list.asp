<!--#include file="admin_security.asp" -->
<!--#include file="Connections/linksmanager.asp" -->
<%
Dim List_Links__MMColParam1
List_Links__MMColParam1 = "%"
If (Request.QueryString("search")   <> "") Then 
  List_Links__MMColParam1 = Request.QueryString("search")  
End If
%>
<%
Dim List_Links__MMColParam2
List_Links__MMColParam2 = "%"
If (Request.Form("search") <> "") Then 
  List_Links__MMColParam2 = Request.Form("search")
End If
%>
<%
Dim List_Links__MMColParam3
List_Links__MMColParam3 = "%"
If (Request.Form("search") <> "") Then 
  List_Links__MMColParam3 = Request.Form("search")
End If
%>
<%
Dim List_Links__MMColParam4
List_Links__MMColParam4 = "%"
If (Request.Form("search")  <> "") Then 
  List_Links__MMColParam4 = Request.Form("search") 
End If
%>
<%
Dim List_Links
Dim List_Links_cmd
Dim List_Links_numRows

Set List_Links_cmd = Server.CreateObject ("ADODB.Command")
List_Links_cmd.ActiveConnection = MM_linksmanager_STRING
List_Links_cmd.CommandText = "SELECT Links.*, LinksCategory.CategoryName, LinksCategory.ParentCategoryIDkey, LinksCategory.CategoryDesc FROM LinksCategory RIGHT JOIN Links ON LinksCategory.CategoryID = Links.CategoryID WHERE Activated = 'True' AND LinksCategory.CategoryName Like ? AND (Links.ItemDesc Like ? OR Links.ItemName Like ? OR Links.ItemURL Like ?) ORDER BY LinksCategory.CategoryName, Links.ItemName" 
List_Links_cmd.Prepared = true
List_Links_cmd.Parameters.Append List_Links_cmd.CreateParameter("param1", 200, 1, 255, List_Links__MMColParam1) ' adVarChar
List_Links_cmd.Parameters.Append List_Links_cmd.CreateParameter("param2", 200, 1, 255, "%" + List_Links__MMColParam2 + "%") ' adVarChar
List_Links_cmd.Parameters.Append List_Links_cmd.CreateParameter("param3", 200, 1, 255, "%" + List_Links__MMColParam3 + "%") ' adVarChar
List_Links_cmd.Parameters.Append List_Links_cmd.CreateParameter("param4", 200, 1, 255, "%" + List_Links__MMColParam4 + "%") ' adVarChar

Set List_Links = List_Links_cmd.Execute
List_Links_numRows = 0
%>
<%
Dim Category
Dim Category_cmd
Dim Category_numRows

Set Category_cmd = Server.CreateObject ("ADODB.Command")
Category_cmd.ActiveConnection = MM_linksmanager_STRING
Category_cmd.CommandText = "SELECT LinksCategory.CategoryID, LinksCategory.CategoryName, LinksCategory.ParentCategoryIDkey, LinksCategory.CategoryDesc FROM LinksCategory INNER JOIN Links ON LinksCategory.CategoryID = Links.CategoryID  GROUP BY LinksCategory.CategoryID, LinksCategory.CategoryName, LinksCategory.ParentCategoryIDkey, LinksCategory.CategoryDesc ORDER BY LinksCategory.CategoryName" 
Category_cmd.Prepared = true

Set Category = Category_cmd.Execute
Category_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
List_Links_numRows = List_Links_numRows + Repeat1__numRows
%>
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="">
    <meta name="author" content="">
<title>List</title>
<link href="bootstrap-3.0.2/css/bootstrap.min.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript">
function confirmDelete(anchor)
  {
    if (confirm('Are you sure?'))
    {
      anchor.href += '&confirm=1';
      return true;
    }
    return false;
  }
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
</script>
</head>
<body>
<div class="container">
<!--#include file="header.asp" -->
<div class="row">
<div class="col-md-6"><% If Not Category.EOF Or Not Category.BOF Then %>
      <form name="form1" method="post" class="form-inline well well-small" action="">
        <select name="Search" id="Search" class="form-control" onChange="MM_jumpMenu('parent',this,0)">
          <option value="<%=Request.ServerVariables("URL")%>" <%If (Not isNull(Request.QueryString("search"))) Then If ("%" = CStr(Request.QueryString("search"))) Then Response.Write("SELECTED") : Response.Write("")%>>Search by Category</option>
          <%
While (NOT Category.EOF)
%>
          <option value="<%=Request.ServerVariables("URL")%>?search=<%=(Category.Fields.Item("CategoryName").Value)%>" <%If (Not isNull(Request.QueryString("search"))) Then If (CStr(Category.Fields.Item("CategoryName").Value) = CStr(Request.QueryString("search"))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(Category.Fields.Item("CategoryName").Value)%></option>
          <%
  Category.MoveNext()
Wend
If (Category.CursorType > 0) Then
  Category.MoveFirst
Else
  Category.Requery
End If
%>
          <option value="<%=Request.ServerVariables("URL")%>" <%If (Not isNull(Request.QueryString("search"))) Then If ("%" = CStr(Request.QueryString("search"))) Then Response.Write("SELECTED") : Response.Write("")%>>Show All</option>
        </select>
      </form>
          <% End If ' end Not Category.EOF Or NOT Category.BOF %>
</div><div class="col-md-6">
      <form name="form" class="form-inline well well-small" method="post" action="">
        <div class="input-group">
          <input type="text" name="Search" class="form-control" placeholder="Search by keyword">
          <span class="input-group-btn">
          <button type="submit" name="submit"  class="btn btn-default">Search</button></span></div>
          
        
      </form>
      </div>
</div>
<div class="table-responsive">
<table class="table table-hover">
<thead>
         
        
  <tr class="active"> 
    <th colspan="2"> Category</th>
    <th>Name </th>
    <th>Description</th>
    <th>Link URL</th>
    <th><div align="center"> <span class="glyphicon glyphicon-plus-sign"></span> <a href="insert.asp">Insert New</a> </div></th>
  </tr></thead>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT List_Links.EOF)) 
%>
  <tr class="<% 
RecordCounter = RecordCounter + 1
If RecordCounter Mod 2 = 1 Then
Response.Write "row1"
Else
Response.Write "row2"
End If
%>"> 
    <td width="2%">      <strong>
    <%Response.Write(RecordCounter)
RecordCounter = RecordCounter%>.      </strong>   </td>
    <td><%=(List_Links.Fields.Item("CategoryName").Value)%></td>
    <td><%=(List_Links.Fields.Item("ItemName").Value)%> </td>
    <td><%=(List_Links.Fields.Item("ItemDesc").Value)%></td>
    <td><a href="<%=(List_Links.Fields.Item("ItemUrl").Value)%>" target="_blank"><%=(List_Links.Fields.Item("ItemUrl").Value)%></a> </td>
    <td nowrap> 
      <div align="center"><span class="glyphicon glyphicon-pencil"></span> <a href="update.asp?ItemID=<%=(List_Links.Fields.Item("ItemID").Value)%>">Edit</a> 
      <span class="glyphicon glyphicon-remove"></span> <a href="delete.asp?ItemID=<%=(List_Links.Fields.Item("ItemID").Value)%>" onClick="return confirmDelete(this)">Delete</a></div>    </td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  List_Links.MoveNext()
Wend
%>
</table>
</div>
<!--#include file="footer.asp" --> 
</div>
<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
    <script src="https://code.jquery.com/jquery.js"></script>
    <!-- Include all compiled plugins (below), or include individual files as needed -->
    <script src="bootstrap-3.0.2/js/bootstrap.min.js"></script>
</body>
</html>
<%
List_Links.Close()
%>
<%
Category.Close()
%>





