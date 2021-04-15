<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_Username")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "index.asp"
  ' redirect with URL parameters (remove the "MM_Logoutnow" query param).
  if (MM_logoutRedirectPage = "") Then MM_logoutRedirectPage = CStr(Request.ServerVariables("URL"))
  If (InStr(1, UC_redirectPage, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
    MM_newQS = "?"
    For Each Item In Request.QueryString
      If (Item <> "MM_Logoutnow") Then
        If (Len(MM_newQS) > 1) Then MM_newQS = MM_newQS & "&"
        MM_newQS = MM_newQS & Item & "=" & Server.URLencode(Request.QueryString(Item))
      End If
    Next
    if (Len(MM_newQS) > 1) Then MM_logoutRedirectPage = MM_logoutRedirectPage & MM_newQS
  End If
  Response.Redirect(MM_logoutRedirectPage)
End If
%>


<style type="text/css">
<!--    
    body {
  padding-top: 20px;
  padding-bottom: 20px;
}
-->
</style>

<div class="navbar navbar-inverse" role="navigation">
  <!-- Brand and toggle get grouped for better mobile display -->
  <div class="navbar-header">
    <a class="navbar-brand" href="index.asp">Links Publisher  <% if Session("MM_UserAuthorization") <> "" THEN %> - Admin Dashboard <span class="glyphicon glyphicon-dashboard"></span></a><% end if%></a>
    <button type="button" class="navbar-toggle" data-toggle="collapse" data-target="#bs-example-navbar-collapse-1">
      <span class="sr-only">Toggle navigation</span>
      <span class="icon-bar"></span>
      <span class="icon-bar"></span>
      <span class="icon-bar"></span>
    </button>
    
   
  </div>

  <!-- Collect the nav links, forms, and other content for toggling -->
  <div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
    <ul class="nav navbar-nav navbar-right">
                <% if Session("MM_UserAuthorization") <> "" THEN %>
          
          <li><a href="index.asp">Home</a></li>
            <li><a href="insert.asp">Insert New Link</a></li>
           <li><a href="list.asp">Archive</a></li>
           
          <li>
  <a href="<%= MM_Logout %>" >Logout</a>
  
 </li>
 <% else%>
  <% if Request.QueryString("valid") = "false" THEN %>
    <a href="javascript:history.go(-1)" class="navbar-brand">
   <span class="label label-danger"> Incorrect username or password - please try again</span> </a>
    <% else%>
 <li>


       <!-- Button trigger modal -->
  <a href="#" data-toggle="modal" data-target="#myModal">Admin Login</a>
 </li><% end if%> <% end if%>
        </ul>

    
  </div><!-- /.navbar-collapse -->
</div>

 
<!-- Modal -->
<div class="modal fade" id="myModal" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
  <div class="modal-dialog">
    <div class="modal-content">
      <div class="modal-header">
        <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
        <h4 class="modal-title" id="myModalLabel">Admin Login <span class="glyphicon glyphicon-lock"></span> </h4>
      </div>
      <div class="modal-body">
       <!--#include file="admin_login.asp" -->
      </div>
    </div><!-- /.modal-content -->
  </div><!-- /.modal-dialog -->
</div><!-- /.modal -->
