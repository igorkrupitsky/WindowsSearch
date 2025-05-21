<%@ Page Language="vb" AutoEventWireup="false"%>
<% 
    Dim sFolder As String = "C:\Igor\ReportPortal"
    Dim sServerName As String = ""
    
    If Left(sFolder, 2) = "\\" Then
        'Search File Sahre
        sServerName = sFolder.Substring(2)
        Dim iPos As Integer = sServerName.IndexOf("\", 3)
        sServerName = """" & sServerName.Substring(0, iPos) & """."
        
        sFolder = Replace(sFolder, "\", "/")
        If Right(sFolder, 1) <> "/" Then
            sFolder += "/"
        End If
    End If
%>
<!DOCTYPE html>
<html>
<head>
    <title>Windows Search</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" integrity="sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u" crossorigin="anonymous">
</head>
<body>
<form id="form1" method="post" >
<div class="container">

    <h1>Windows Search
         <a href="WinSearchHelp.htm" target="_blank" ><span class="glyphicon glyphicon-question-sign"></span></a> 
    </h1>

<textarea class="form-control" rows="10" style="width: 100%; margin-top: 10px;margin-bottom: 5px;" name="txtSql"
    ><%If Request.Form("txtSql") <> "" Then
           Response.Write(Request.Form("txtSql"))  
         Else%>SELECT TOP 5 System.ItemUrl, System.ItemName, System.ItemType, System.KindText, System.Filename, System.Size, System.DateModified, System.Search.Rank
FROM <%=sServerName%>SystemIndex 
WHERE scope ='file:<%=sFolder%>' 
AND System.ItemName LIKE '%Test%' 
and FREETEXT('Test')
and CONTAINS('Test')
and System.DateModified > '2007-07-07'<%End If%></textarea>
<button class="btn btn-default" type="submit">Run</button>

    <div style="float:right">
        <a href="https://msdn.microsoft.com/sv-se/library/bb763036(en-us,VS.85).aspx" target="_blank">System Shell Properties</a>        
    </div>

<%If Request.Form("txtSql") <> "" Then
    
        Dim sConnectionString As String = "Provider=Search.CollatorDSO;Extended Properties=""Application=Windows"""
        Dim cn As New System.Data.OleDb.OleDbConnection(sConnectionString)

        Try
            cn.Open()
        Catch ex As Exception
            Response.Write(ex.Message & "; ConnectionString: " & sConnectionString)
        End Try

        Try
            Dim ad As System.Data.OleDb.OleDbDataAdapter = New System.Data.OleDb.OleDbDataAdapter(Request.Form("txtSql"), cn)
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ad.Fill(ds)
            If ds.Tables.Count > 0 Then
                Dim oTable As System.Data.DataTable = ds.Tables(0)
                    
                Response.Write("<table class='table table-striped'><thead><tr>")
                For iCol As Integer = 0 To oTable.Columns.Count - 1
                    Response.Write("<th>" & oTable.Columns(iCol).Caption & "</th>" & vbCrLf)
                Next
                Response.Write("</tr></thead><tbody>" & vbCrLf)
                    
                For iRow As Integer = 0 To oTable.Rows.Count - 1
                    Response.Write("<tr>")
                    For iCol As Integer = 0 To oTable.Columns.Count - 1
                        Response.Write("<td>" & oTable.Rows(iRow)(iCol) & "</td>" & vbCrLf)
                    Next
                    Response.Write("</tr>")
                Next
                
                Response.Write("</tbody></table>")
            End If
        Catch ex As Exception
            Response.Write("<div class='alert alert-danger' style='margin-top: 10px;'>" & ex.Message & "</div>")
        End Try
        cn.Close()
  
End If
%>

</div>

</form>
</body>
</html>
