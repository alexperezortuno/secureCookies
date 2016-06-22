# Secure Cookies in ASP Classic

Usage for one cookie
```sh
 Dim cookie
 Set cookie = Server.CreateObject("Scripting.Dictionary")
 cookie "name", "mycookie"
 setCookies(cookie)
```

Usage for multiple cookies
```sh
 Dim newCookies(1)
 Set newCookies(0) = Server.CreateObject("Scripting.Dictionary")
 Set newCookies(1) = Server.CreateObject("Scripting.Dictionary")
 newCookies(0).Add "name", "mycookie"
 newCookies(0).Add "value", "cookie 1"
 newCookies(1).Add "name", "mycookie"
 newCookies(1).Add "value", "cookie 2"
 setCookies(newCookies)
```

This function adds tags for secure your cookies and insert HttpOnly, you can read more information in [SecureFlags](https://www.owasp.org/index.php/SecureFlag), You can additional path variable like  cookie "path", "/routeofmypage"  for change modify the path of cookie

### Version
0.1
