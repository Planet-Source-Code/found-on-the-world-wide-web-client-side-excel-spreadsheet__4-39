<div align="center">

## Client Side Excel Spreadsheet


</div>

### Description

You can create an Excel Spreadsheet on the client side with a HTML data stream. This only works with IE 4 or above and the client must have Excel installed.

The following example script would create an exel spreadsheet with three columns. The first two contain numeric data in the cells and the third sends a function to sum the first two columns.

http://www.truegeeks.com/asp/mam/osdoc/osframe.asp
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Found on the World Wide Web](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/found-on-the-world-wide-web.md)
**Level**          |Advanced
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Server Side](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/server-side__4-31.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/found-on-the-world-wide-web-client-side-excel-spreadsheet__4-39/archive/master.zip)





### Source Code

```
<%@ Language=VBScript %>
<%Response.ContentType = "application/vnd.ms-excel"%>
<table>
<tr>
<th>Col 1</td>
<th>Col 2</td>
<th>Total</td>
</tr>
<tr>
<td>1</td>
<td>2</td>
<td><font color=#000080>=sum(a2:b2)</font></td>
</tr>
<tr>
<td>4</td>
<td>6</td>
<td><font color=#000080>=sum(a3:b3)</font></td>
</tr>
</table>
```

