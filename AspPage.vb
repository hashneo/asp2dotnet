Public Class AspPage
	Inherits Page

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
		Response.Write(Str & vbCrLf)
	End Function
End Class