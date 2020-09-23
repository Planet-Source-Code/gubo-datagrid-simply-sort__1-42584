<div align="center">

## datagrid simply sort


</div>

### Description

a simply sort on de collumsheaders of a datagrid
 
### More Info
 
a datarid,and a connection with a d-base

sort "asc" & "desc"


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[gubo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/gubo.md)
**Level**          |Beginner
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/gubo-datagrid-simply-sort__1-42584/archive/master.zip)





### Source Code

```
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
  With AdoBoeken.Recordset
    If (.Sort = .Fields(ColIndex).[Name] & " Asc") Then
      .Sort = .Fields(ColIndex).[Name] & " Desc"
    Else
      .Sort = .Fields(ColIndex).[Name] & " Asc"
    End If
  End With
End Sub
```

