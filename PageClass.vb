Public Class PageClass

	Protected ReadOnly Property Page As System.Web.UI.Page

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

	Protected Sub New(Page As System.Web.UI.Page)
		Me._Page = Page
	End Sub

End Class