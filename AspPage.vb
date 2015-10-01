Public Class AspPage
	Inherits Page

	Private _HttpApplication As System.Web.HttpApplication

	Public Overrides ReadOnly Property Session As HttpSessionState
		Get
			If (Not _HttpApplication Is Nothing) Then
				Return _HttpApplication.Session
			End If
			Return MyBase.Session
		End Get
	End Property

	Public ReadOnly Property Application As System.Web.HttpApplicationState
		Get
			If (Not _HttpApplication Is Nothing) Then
				Return _HttpApplication.Application
			End If
			Return MyBase.Application
		End Get
	End Property

	Public ReadOnly Property Server As System.Web.HttpServerUtility
		Get
			If (Not _HttpApplication Is Nothing) Then
				Return _HttpApplication.Server
			End If
			Return MyBase.Server
		End Get
	End Property

	Private classCache As Dictionary(Of String, PageClass) = New Dictionary(Of String, PageClass)

	Public Function createInstance(What As System.Type) As PageClass
		Dim classInstance As PageClass = Nothing
		SyncLock classCache
			If Not classCache.TryGetValue(What.ToString(), classInstance) Then
				classInstance = Activator.CreateInstance(What, Me)
				classCache.Add(What.ToString(), classInstance)
			End If
		End SyncLock
		Return classInstance
	End Function

	Function Response_WriteLine(Str As String)
		If Not Response Is Nothing Then
			Response.Write(Str & vbCrLf)
		End If
	End Function

	Public Sub New()
		MyBase.New()
	End Sub

	Public Sub New(HttpApplication As System.Web.HttpApplication)
		Me._HttpApplication = HttpApplication
	End Sub
End Class