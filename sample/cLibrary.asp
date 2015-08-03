<% 
Class cLibrary

'Private, class member variable
Private m_Books
Private m_Authors

Sub Class_Initialize()
    set m_Books = Server.CreateObject ("Scripting.Dictionary")
    set m_Authors = Server.CreateObject ("Scripting.Dictionary")
End Sub
Sub Class_Terminate()
    set m_Books = Nothing
    set m_Authors = Nothing
End Sub

'Read the current Books
Public Property Get Books()
    Set Books = m_Books
End Property

'Read the current Authors
Public Property Get Authors()
    Set Authors = m_Authors
End Property


'#############  Public Functions, accessible to the web pages ##############

    'Loads all books into the library
    Public Function GetAllBooks()
        dim strSQL    
        strSQL = "SELECT lngBookID, strTitle, strSubTitle, strISBN, strBindingType, "
        strSQL = strSQL & " strPublishersName, strPublishedYear, strPageCount FROM Book "
        GetAllBooks = LoadBookData (strSQL)
    End Function

    'Loads this object's values by loading a record based on the given ID
    Public Function GetBooksByPublisher(p_Value)
        dim strSQL    
        strSQL = "SELECT lngBookID, strTitle, strSubTitle, strISBN, strBindingType, "
        strSQL = strSQL & " strPublishersName, strPublishedYear, strPageCount FROM Book "
        strSQL = strSQL & " where strPublishersName = '" & SingleQuotes(p_Value) & "'"
        GetBooksByPublisher = LoadBookData (strSQL) 
    End Function    

    'Loads this object's values by loading a record based on the given ID
    Public Function GetBooksByAuthorID(p_Value)
        dim strSQL    
        strSQL = "SELECT Book.lngBookID, Book.strTitle, strSubTitle, strISBN, strBindingType, "
        strSQL = strSQL & " strPublishersName, strPublishedYear, strPageCount FROM Book " & _
                          " INNER JOIN AuthorToBook ON Book.lngBookID = AuthorToBook.lngBookID "
        strSQL = strSQL &   " WHERE (AuthorToBook.lngAuthorID = " & p_Value & ") "   
        GetBooksByAuthorID = LoadBookData (strSQL) 
    End Function    
    
    'Loads this object's values by loading a record based on the given ID
    Public Function GetAllAuthors()
        dim strSQL    
        strSQL = "SELECT Author.lngAuthorID, Author.strFirstName, Author.strLastName, Author.strTitle "
        strSQL = strSQL & " FROM Author "
        GetAllAuthors = LoadAuthorData (strSQL) 
    End Function
    
    'Loads this object's values by loading a record based on the given ID
    Public Function GetAuthorsByBookID(p_Value)
        dim strSQL    
        strSQL = "SELECT Author.lngAuthorID, Author.strFirstName, Author.strLastName, Author.strTitle "
        strSQL = strSQL & " FROM Author INNER JOIN AuthorToBook ON Author.lngAuthorID = AuthorToBook.lngAuthorID "
        strSQL = strSQL &   " WHERE (AuthorToBook.lngBookID = " & p_Value & ") "    
        GetAuthorsByBookID = LoadAuthorData (strSQL) 
    End Function
    
    Public Sub AssociateAuthorWithBook(p_AuthorID, p_BookID)
        dim strSQL
        strSQL = " insert into AuthorToBook "
        strSQL = strSQL & " values (" & p_AuthorID & ", " & p_BookID & ")"
        RunSQL strSQL
    End Sub
'#############  Private Functions                           ##############

    'Takes a recordset
    'Fills the object's properties using the recordset
    Private Function FillBooksFromRS(p_RS)
        dim myBook
        do while not p_RS.eof
            set myBook = New cBook
            MyBook.ID                 = p_RS.fields("lngBookID").Value
            MyBook.Title              = p_RS.fields("strTitle").Value
            MyBook.SubTitle           = p_RS.fields("strSubTitle").Value
            MyBook.ISBN               = p_RS.fields("strISBN").Value
            MyBook.BindingType        = p_RS.fields("strBindingType").Value
            MyBook.PublishersName     = p_RS.fields("strPublishersName").Value
            MyBook.PublishedYear      = p_RS.fields("strPublishedYear").Value
            MyBook.PageCount          = p_RS.fields("strPageCount").Value
            m_Books.Add myBook.ID, myBook
            p_RS.movenext
        loop
    End Function

    Private Function LoadBookData(p_strSQL)
        dim rs
        set rs = LoadRSFromDB(p_strSQL)
        FillBooksFromRS(rs)
        LoadBookData = rs.recordcount
        rs. close
        set rs = nothing
    End Function

    Private Function FillAuthorsFromRS(p_RS)
        dim myBook
        do while not p_RS.eof
			dim myAuthor
            set myAuthor = New cAuthor
            myAuthor.ID                 = p_RS.fields("lngAuthorID").Value
            myAuthor.FirstName          = p_RS.fields("strFirstName").Value
            myAuthor.LastName           = p_RS.fields("strLastName").Value
            myAuthor.Title              = p_RS.fields("strTitle").Value
            m_Authors.Add myAuthor.ID, myAuthor
            p_RS.movenext
        loop
    End Function


    Private Function LoadAuthorData(p_strSQL)
        dim rs
        set rs = LoadRSFromDB(p_strSQL)
        FillAuthorsFromRS(rs)
        LoadAuthorData = rs.recordcount
        rs. close
        set rs = nothing
    End Function    
End Class

%>