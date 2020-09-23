<div align="center">

## ASP\.NET Database Login


</div>

### Description

This is for all you ASP programmers out there looking for a way to validate users against a SQL database and are used to the old ADO RecordSource.EOF way like myself. You must change to represent your database.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Robert L Long](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/robert-l-long.md)
**Level**          |Intermediate
**User Rating**    |3.9 (47 globes from 12 users)
**Compatibility**  |ASP\.NET
**Category**       |[Validation/ Processing](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/validation-processing__10-16.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/robert-l-long-asp-net-database-login__10-118/archive/master.zip)





### Source Code

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="700" id="AutoNumber1">
 <tr>
  <td width="100%"><font size="2">&lt;%@ Page Language=&quot;vb&quot; AutoEventWireup=&quot;false&quot;
  Codebehind=&quot;index.aspx.vb&quot; Inherits=&quot;VirtualManagement.Login&quot;%&gt;<br>
  &lt;%@ Register TagPrefix=&quot;uc1&quot; TagName=&quot;Menu&quot; Src=&quot;Menu.ascx&quot; %&gt;<br>
  &lt;!DOCTYPE HTML PUBLIC &quot;-//W3C//DTD HTML 4.0 Transitional//EN&quot;&gt;<br>
  &lt;HTML&gt;<br>
  &lt;HEAD&gt;<br>
&nbsp;&nbsp;&nbsp; &lt;title&gt;Login&lt;/title&gt;<br>
&nbsp;&nbsp;&nbsp; &lt;meta content=&quot;False&quot; name=&quot;vs_showGrid&quot;&gt;<br>
&nbsp;&nbsp;&nbsp; &lt;meta content=&quot;Microsoft Visual Studio.NET 7.0&quot;
  name=&quot;GENERATOR&quot;&gt;<br>
&nbsp;&nbsp;&nbsp; &lt;meta content=&quot;Visual Basic 7.0&quot; name=&quot;CODE_LANGUAGE&quot;&gt;<br>
&nbsp;&nbsp;&nbsp; &lt;meta content=&quot;JavaScript&quot; name=&quot;vs_defaultClientScript&quot;&gt;<br>
&nbsp;&nbsp;&nbsp; &lt;meta content=&quot;http://schemas.microsoft.com/intellisense/ie5&quot;
  name=&quot;vs_targetSchema&quot;&gt;<br>
&nbsp;&nbsp;&nbsp; &lt;script language=&quot;VB&quot; runat=&quot;server&quot;&gt;<br>
  <br>
&nbsp;&nbsp;&nbsp; Sub Login_Click(Sender As Object, E As EventArgs)<br>
  <font color="#008000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ' Robert
  Long<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ' GlobalDeveloping.com<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ' This is for all you ASP programmers
  out there looking for <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ' a way to validate users against a
  SQL database and are used to <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ' the old ADO RecordSource.EOF way
  like myself.</font><br>
  <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Dim StrUser As String, StrPass As
  String<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Dim BValid As Boolean<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Dim ConVM As SqlConnection<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Dim CmdCoInfo As SqlCommand<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Dim DrCoInfo As SqlDataReader<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Dim StrSQL As String<br>
  <br>
  <font color="#008000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ' We will
  request all variables from our form with this.<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ' Look Mom &quot;No QueryStrings&quot;.</font><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; StrUser = TxtUser.Text<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; StrPass = TxtPass.Text<br>
  <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <font color="#008000">&nbsp;' This is our
  boolean variable for validation purposes set to true if valid user</font><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; BValid = False<br>
  <br>
  <font color="#008000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; '
  Initialize Database Connection<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ' Use whatever server you have</font><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ConVM = New SqlConnection(&quot;Server=MyServer\NetSDK;UID=sa;PWD=;Database=Pubs&quot;)<br>
  <br>
  <font color="#008000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ' Create
  Select Command<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ' Also, notice that it is only
  selecting a record if the username/password match<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ' Otherwise it selects nothing<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ' Make sure you change to match your
  variables</font><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; StrSQL = &quot;SELECT * FROM Pubs WHERE
  UserName='&quot; &amp; StrUser &amp; &quot;'&quot;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; CmdCoInfo = New SqlCommand(StrSQL,
  ConVM)<br>
  <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <font color="#008000">&nbsp;' Retrieve the
  record that matches the username/password</font><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ConVM.Open()<br>
  <br>
  <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; DrCoInfo = CmdCoInfo.ExecuteReader()<br>
  <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <font color="#008000">' This acts
  like the (Not RecordSource.Eof) in ASP 3.0</font><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; While DrCoInfo.Read()<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; If
  DrCoInfo(&quot;UserName&quot;) &lt;&gt; &quot;&quot; And DrCoInfo(&quot;Password&quot;) &lt;&gt; &quot;&quot; Then<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  '(Optional) Write the username to a cookie to personilize the site<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  Response.Cookies(&quot;ValidUser&quot;).Value = DrCoInfo(&quot;UserName&quot;)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  Response.Cookies(&quot;ValidUser&quot;).Expires = DateTime.Now.AddMonths(1)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  BValid = True<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Else<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; End If<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; End While<br>
  <font color="#008000"><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ' Don't forget this</font><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ConVM.Close()<br>
  <font color="#008000"><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ' This handles all response per
  validation<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ' If validated it goes to a
  welcome.aspx page</font><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; If BValid = True Then<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Response.Redirect(&quot;Welcome.aspx&quot;)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ElseIf BValid = False Then<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; LblError.Text = &quot;Login Error: Please
  try again.&quot;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; End If<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; End Sub<br>
  <br>
&nbsp;&nbsp;&nbsp; &lt;/script&gt; <br>
  &lt;/HEAD&gt;<br>
  &lt;body bottomMargin=&quot;0&quot; bgColor=&quot;gray&quot; leftMargin=&quot;0&quot; topMargin=&quot;0&quot;
  rightMargin=&quot;0&quot; MS_POSITIONING=&quot;GridLayout&quot;&gt;<br>
  &lt;form runat=&quot;server&quot;&gt;<br>
  &lt;asp:panel id=&quot;PnlLogin&quot; style=&quot;Z-INDEX: 101; LEFT: 29px; POSITION:
  absolute; TOP: 72px&quot; runat=&quot;server&quot; HorizontalAlign=&quot;Center&quot; Height=&quot;184px&quot;
  BorderColor=&quot;Transparent&quot; BorderWidth=&quot;2px&quot; BorderStyle=&quot;Outset&quot; BackColor=&quot;LightSteelBlue&quot;
  Width=&quot;326px&quot;&gt;<br>
  &lt;P&gt;&amp;nbsp;&lt;/P&gt;<br>
  &lt;P&gt;<br>
  &lt;asp:Label id=&quot;LblUser&quot; runat=&quot;server&quot; Width=&quot;73px&quot; Font-Names=&quot;Arial&quot;
  Font-Size=&quot;X-Small&quot;&gt;Username&lt;/asp:Label&gt;<br>
  &lt;asp:TextBox id=&quot;TxtUser&quot; runat=&quot;server&quot; Width=&quot;120px&quot; BackColor=&quot;Azure&quot;
  Height=&quot;20px&quot;&gt;&lt;/asp:TextBox&gt;&lt;/P&gt;<br>
  &lt;P&gt;<br>
  &lt;asp:Label id=&quot;LblPass&quot; runat=&quot;server&quot; Width=&quot;73px&quot; Font-Names=&quot;Arial&quot;
  Font-Size=&quot;X-Small&quot;&gt;Password&lt;/asp:Label&gt;<br>
  &lt;asp:TextBox id=&quot;TxtPass&quot; runat=&quot;server&quot; Width=&quot;120px&quot; BackColor=&quot;Azure&quot;
  Height=&quot;20px&quot; TextMode=&quot;Password&quot;&gt;&lt;/asp:TextBox&gt;&lt;/P&gt;<br>
  &lt;P&gt;<br>
  &lt;asp:Button id=&quot;CmdLogin&quot; runat=&quot;server&quot; Width=&quot;99px&quot; BackColor=&quot;Gold&quot;
  Height=&quot;25px&quot; Font-Names=&quot;Arial&quot; Font-Size=&quot;X-Small&quot; OnServerClick=&quot;Login_Click&quot;
  Text=&quot;Log-in Now&quot; ForeColor=&quot;Black&quot;&gt;&lt;/asp:Button&gt;&lt;/P&gt;<br>
  &lt;/asp:panel&gt;&lt;asp:panel id=&quot;PnlTitle&quot; style=&quot;Z-INDEX: 102; LEFT: 32px;
  POSITION: absolute; TOP: 75px&quot; runat=&quot;server&quot; HorizontalAlign=&quot;Left&quot;
  Height=&quot;19px&quot; BorderColor=&quot;Navy&quot; BorderWidth=&quot;1px&quot; BorderStyle=&quot;Solid&quot;
  BackColor=&quot;DarkBlue&quot; Width=&quot;320px&quot; Font-Names=&quot;Arial&quot; Font-Size=&quot;X-Small&quot;
  ForeColor=&quot;White&quot; Font-Bold=&quot;True&quot;&gt;&amp;nbsp;Log-on Window&lt;/asp:panel&gt;&amp;nbsp;<br>
  &lt;asp:Label id=&quot;LblError&quot; style=&quot;Z-INDEX: 103; LEFT: 46px; POSITION:
  absolute; TOP: 46px&quot; runat=&quot;server&quot;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  Width=&quot;288px&quot; Font-Names=&quot;Arial Black&quot; ForeColor=&quot;Yellow&quot;&gt;&lt;/asp:Label&gt;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &lt;/form&gt;<br>
&nbsp;&nbsp;&nbsp; &lt;/body&gt;<br>
  &lt;/HTML&gt;</font></td>
 </tr>
</table>
<p>&nbsp;</p>
<p><font size="2"><br>
&nbsp;</font></p>

