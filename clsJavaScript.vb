Public Class clsJavaScript

    Public Function MsgBox(ByVal aMessage As String) As String
        Return "<script type=""text/javascript"">" & vbCrLf & _
               "alert('" & Encode(aMessage) & "')" & vbCrLf & _
               "</script>"
    End Function

    Public Function Encode(ByVal aValue As String) As String
        aValue = Trim(aValue)
        aValue = Replace(aValue, "\", "\\")
        aValue = Replace(aValue, "'", "\'")
        aValue = Replace(aValue, """", "\""")
        aValue = Replace(aValue, vbCrLf, "\n")
        Return aValue
    End Function

    Public Function Back() As String
        Return "<script type=""text/javascript"">" & vbCrLf & _
               "location.href='javascript:history.back()'" & vbCrLf & _
               "</script>"
    End Function

    Public Function Redirect(ByVal page As String) As String
        Return "<script type=""text/javascript"">" & vbCrLf & _
               "location.href='" & page & "'" & vbCrLf & _
               "</script>"
    End Function

    Public Sub HighlightTxt(ByVal Ctrl As Control)
        Dim p As Control = Ctrl.Parent
        Dim csType As Type = p.GetType
        Dim sb As New StringBuilder

        sb.Append("<script language='JavaScript'>" & Chr(13) & Chr(10))
        sb.Append("<!--" & Chr(13) & Chr(10))
        sb.Append("function HighlightTxt()" & Chr(13) & Chr(10))
        sb.Append("{" & Chr(13) & Chr(10))
        sb.Append("document.")

        While Not TypeOf p Is System.Web.UI.HtmlControls.HtmlForm
            p = p.Parent
        End While

        sb.Append(p.ClientID)
        sb.Append("['")
        sb.Append(Ctrl.UniqueID.ToString)
        sb.Append("'].select();" & Chr(13) & Chr(10))
        sb.Append("}" & Chr(13) & Chr(10))
        sb.Append("window.onload = HighlightTxt;" & Chr(13) & Chr(10))
        sb.Append("// -->" & Chr(13) & Chr(10))
        sb.Append("</script>")

        Ctrl.Page.ClientScript.RegisterClientScriptBlock(csType, "HighlightTxt", sb.ToString())
    End Sub

    Public Sub MoveCursor(ByVal Ctrl As Control)
        Dim p As Control = Ctrl.Parent
        Dim csType As Type = p.GetType
        Dim sb As New StringBuilder

        sb.Append("<script language='JavaScript'>" & Chr(13) & Chr(10))
        sb.Append("<!--" & Chr(13) & Chr(10))
        sb.Append("function MoveCursor()" & Chr(13) & Chr(10))
        sb.Append("{" & Chr(13) & Chr(10))
        sb.Append("var a = document.")

        While Not TypeOf p Is System.Web.UI.HtmlControls.HtmlForm
            p = p.Parent
        End While

        sb.Append(p.ClientID)
        sb.Append("['")
        sb.Append(Ctrl.UniqueID.ToString)
        sb.Append("'];" & Chr(13) & Chr(10))
        sb.Append("a.focus();" & Chr(13) & Chr(10))
        sb.Append("a.value = a.value;" & Chr(13) & Chr(10))
        sb.Append("}" & Chr(13) & Chr(10))
        sb.Append("window.onload = MoveCursor;" & Chr(13) & Chr(10))
        sb.Append("// -->" & Chr(13) & Chr(10))
        sb.Append("</script>")

        Ctrl.Page.ClientScript.RegisterClientScriptBlock(csType, "MoveCursor", sb.ToString())
    End Sub

    Public Sub OpenPage(ByVal Ctrl As Control, ByVal page As String)
        Dim csType As Type = Ctrl.GetType
        Dim sb As New StringBuilder

        sb.Append("<script language='JavaScript'>" & Chr(13) & Chr(10))
        sb.Append("<!--" & Chr(13) & Chr(10))
        sb.Append("function OpenPage()" & Chr(13) & Chr(10))
        sb.Append("{" & Chr(13) & Chr(10))
        sb.Append("window.open('" & page & "');" & Chr(13) & Chr(10))
        sb.Append("return false;" & Chr(13) & Chr(10))
        sb.Append("}" & Chr(13) & Chr(10))
        sb.Append("window.onload = OpenPage;" & Chr(13) & Chr(10))
        sb.Append("// -->" & Chr(13) & Chr(10))
        sb.Append("</script>")

        Ctrl.Page.ClientScript.RegisterClientScriptBlock(csType, "OpenPage", sb.ToString())
    End Sub

    Public Sub Open2Page(ByVal Ctrl As Control, ByVal page1 As String, ByVal page2 As String)
        Dim csType As Type = Ctrl.GetType
        Dim sb As New StringBuilder

        sb.Append("<script language='JavaScript'>" & Chr(13) & Chr(10))
        sb.Append("<!--" & Chr(13) & Chr(10))
        sb.Append("function Open2Page()" & Chr(13) & Chr(10))
        sb.Append("{" & Chr(13) & Chr(10))
        sb.Append("window.open('" & page1 & "');" & Chr(13) & Chr(10))
        sb.Append("window.open('" & page2 & "');" & Chr(13) & Chr(10))
        sb.Append("return false;" & Chr(13) & Chr(10))
        sb.Append("}" & Chr(13) & Chr(10))
        sb.Append("window.onload = Open2Page;" & Chr(13) & Chr(10))
        sb.Append("// -->" & Chr(13) & Chr(10))
        sb.Append("</script>")

        Ctrl.Page.ClientScript.RegisterClientScriptBlock(csType, "Open2Page", sb.ToString())
    End Sub

    Public Sub Open3Page(ByVal Ctrl As Control, ByVal page1 As String, ByVal page2 As String, ByVal page3 As String)
        Dim csType As Type = Ctrl.GetType
        Dim sb As New StringBuilder

        sb.Append("<script language='JavaScript'>" & Chr(13) & Chr(10))
        sb.Append("<!--" & Chr(13) & Chr(10))
        sb.Append("function Open2Page()" & Chr(13) & Chr(10))
        sb.Append("{" & Chr(13) & Chr(10))
        sb.Append("window.open('" & page1 & "');" & Chr(13) & Chr(10))
        sb.Append("window.open('" & page2 & "');" & Chr(13) & Chr(10))
        sb.Append("window.open('" & page3 & "');" & Chr(13) & Chr(10))
        sb.Append("return false;" & Chr(13) & Chr(10))
        sb.Append("}" & Chr(13) & Chr(10))
        sb.Append("window.onload = Open2Page;" & Chr(13) & Chr(10))
        sb.Append("// -->" & Chr(13) & Chr(10))
        sb.Append("</script>")

        Ctrl.Page.ClientScript.RegisterClientScriptBlock(csType, "Open2Page", sb.ToString())
    End Sub
End Class
