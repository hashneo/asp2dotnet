<%@ Language=VBScript %>
<!-- #INCLUDE virtual="/cBook.asp" -->
<!-- #INCLUDE virtual="/cAuthor.asp" -->
<!-- #INCLUDE virtual="/cLibrary.asp" -->
<!-- #INCLUDE virtual="/i_helper.asp" -->

<%


dim MyBook
Set MyBook = new cBook

if request("bookid")<>"" then
    mybook.loadfromid(request("bookid"))    
end if

if request("cmd")<>"" then    
    Mybook.ISBN = Request.Form("isbn")
    Mybook.Title = Request.Form("Title")
    Mybook.SubTitle = Request.Form("SubTitle")
    Mybook.ISBN = Request.Form("isbn")
    Mybook.BindingType = Request.Form("BindingType")
    Mybook.PublishersName = Request.Form("PublishersName")
    Mybook.PublishedYear = Request.Form("PublishedYear")
    Mybook.PageCount = Request.Form("PageCount")
    mybook.store    
end if

%>






<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML>
<HEAD>
<TITLE></TITLE>
<META NAME="Generator" CONTENT="TextPad 4.0">
<META NAME="Author" CONTENT="?">
<META NAME="Keywords" CONTENT="?">
<META NAME="Description" CONTENT="?">
</HEAD>

<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#FF0000" VLINK="#800000" ALINK="#FF00FF" BACKGROUND="?">
<form name="form1" method="post" action="edit_book.asp">
  <table border="1" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="134">Book ID</td>
      <td width="12"><input type="hidden" name="cmd" value="edit"></td>
      <td width="598"><input type="hidden" name="bookid" value="<%=Mybook.ID%>"><%=Mybook.ID%></td>
    </tr>
    <tr> 
      <td>ISBN</td>
      <td>&nbsp;</td>
      <td><input  size="50" name="isbn" value="<%=htmlencode(Mybook.ISBN)%>"></td>
    </tr>
    <tr> 
      <td>Title</td>
      <td>&nbsp;</td>
      <td><input  size="150" name="Title" value="<%=htmlencode(Mybook.Title)%>"></td>
    </tr>
    <tr> 
      <td>SubTitle</td>
      <td>&nbsp;</td>
      <td><input  size="150" name="SubTitle" value="<%=htmlencode(Mybook.SubTitle)%>"></td>
    </tr>
    <tr> 
      <td>Binding Type</td>
      <td>&nbsp;</td>
      <td><input  size="50" name="BindingType" value="<%=htmlencode(Mybook.BindingType)%>"></td>
    </tr>
    <tr> 
      <td>PublishersName</td>
      <td>&nbsp;</td>
      <td><input  size="50" name="PublishersName" value="<%=htmlencode(Mybook.PublishersName)%>"></td>
    </tr>
    <tr> 
      <td>Published Year</td>
      <td>&nbsp;</td>
      <td><input size="50" name="PublishedYear" value="<%=(Mybook.PublishedYear)%>"></td>
    </tr>
    <tr> 
      <td>´Number of pages</td>
      <td>&nbsp;</td>
      <td><input size="50"  name="PageCount" value="<%=htmlencode(Mybook.PageCount)%>"></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp; </td>
      <td><input type="reset" name="Submit2" value="Reset"> <input type="submit" name="Submit" value="Submit"></td>
    </tr>
  </table>
</form>

<a href="./default.asp">List</a>
</BODY>
</HTML>
