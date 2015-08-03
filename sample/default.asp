<!-- #INCLUDE virtual="/cLibrary.asp" -->
<!-- #INCLUDE virtual="/cBook.asp" -->
<!-- #INCLUDE virtual="/cAuthor.asp" -->
<!-- #INCLUDE virtual="/i_helper.asp" -->

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


<%

Dim Library
Set Library = New cLibrary
Call Library.GetAllBooks

for each MyBook in Library.Books.Items
%><table border=0 bgcolor="lightgrey" width="90%">
    <tr bgcolor="skyblue"><TD>ID: <a href="./edit_book.asp?bookid=<%=MyBook.ID%>"><%=MyBook.ID%></a><br>
    
    <%
	
		Dim test
		
		test = adExecuteNoRecords

        for each Author in MyBook.Authors.Items
        echobr Author.ID & " "& Author.FirstName & " "&   Author.LastName & " "&   Author.Title
        next
    %>
    
    
    </td></tr>
    <tr><TD>Title: <%=MyBook.Title%></td></tr>
    <tr><TD>SubTitle: <%=MyBook.SubTitle%></td></tr>
    <tr><TD>Publisher: <%=MyBook.PublishersName%></td></tr>
    <tr><TD>Year: <%=MyBook.PublishedYear%></td></tr>
    <tr><TD>Format: <%=MyBook.BindingType%></td></tr>
    <tr><TD>ISBN: <%=MyBook.ISBN%></td></tr>
    <tr bgcolor="skyblue"><TD><b>
    </b>
    
    </td></tr></table><br>

<%
    set mybook = nothing
next
echo ""

Set Library = New cLibrary

Call Library.GetAllAuthors
for each Author in Library.Authors.Items
    echobr "<hr>"

    echo "<a href=""edit_author.asp?authorid=" & Author.ID & """>" &  Author.ID & "</a> "
    echobr Author.FirstName & " "&   Author.LastName & " "&   Author.Title
    
        for each MyBook in Author.Books.Items 
            echo"<b>"
            echobr MyBook.Title
            echobr MyBook.SubTitle
            echobr MyBook.PublishersName
            echobr MyBook.ISBN        
            echo"</b>"
        next
   
Next 

%>


<a href="./edit_book.asp">New Book</a>
</BODY>
</HTML>
