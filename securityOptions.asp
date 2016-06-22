Function setCookies(vars)
  Dim k
  If isArray(vars) = True Then
    For k= 0 To UBound(vars)
      If vars(k).Exists("name") = True And vars(k).Exists("value") = True Then
        If vars(k).Exists("path") = True Then
          Response.AddHeader "Set-Cookie", vars(k)("name") & "=" & vars(k)("value") & ";secure=true; path=" & vars(k)("path") & ";expires=" & CStr(Now + 1) & ";HttpOnly"
        Else
          Response.AddHeader "Set-Cookie", vars(k)("name") & "=" & vars(k)("value") & ";secure=true; path=/;expires=" & CStr(Now + 1) & ";HttpOnly"
        End If
      End if
    Next
  Else
    If vars.Exists("name") = True And vars.Exists("value") = True Then
      If vars.Exists("path") = True Then
        Response.AddHeader "Set-Cookie", vars("name") & "=" & vars("value") & ";secure=true; path=" & vars("path") & ";expires=" & CStr(Now + 1) & ";HttpOnly"
      Else
        Response.AddHeader "Set-Cookie", vars("name") & "=" & vars("value") & ";secure=true; path=/;expires=" & CStr(Now + 1) & ";HttpOnly"
      End If
    End if
  End if
End Function
