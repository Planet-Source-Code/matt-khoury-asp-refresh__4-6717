<div align="center">

## ASP Refresh


</div>

### Description

Instead of using a META refresh, use an ASP refresh to insure you have the most current copy of the web page.
 
### More Info
 
Refreshes the page in the increments you set.

None that I know of.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matt Khoury](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matt-khoury.md)
**Level**          |Beginner
**User Rating**    |4.8 (38 globes from 8 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matt-khoury-asp-refresh__4-6717/archive/master.zip)





### Source Code

```
<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = 0
	Response.Buffer = True
	Response.Clear
	Response.AddHeader "Refresh", "10"
%>
```

