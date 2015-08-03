Public Class PageClass

	Private _Page As AspPage
	Public ReadOnly Property Page As AspPage
		Get
			Return _Page
		End Get
	End Property

	Protected ReadOnly Property Server As System.Web.HttpServerUtility
		Get
			Return Me.Page.Server
		End Get
	End Property

	Protected ReadOnly Property Application As System.Web.HttpApplicationState
		Get
			Return Me.Page.Application
		End Get
	End Property

	Protected ReadOnly Property Request As System.Web.HttpRequest
		Get
			Return Me.Page.Request
		End Get
	End Property

	Protected ReadOnly Property Response As System.Web.HttpResponse
		Get
			Return Me.Page.Response
		End Get
	End Property

	Public Function createInstance(What As System.Type) As PageClass
		Return Page.createInstance( What )
	End Function

	Protected Sub New(Page As AspPage)
		Me._Page = Page
	End Sub

End Class