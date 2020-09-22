<div align="center">

## A Better Way to Test For Tables \(great for begginners \- well commented\)


</div>

### Description

This is the better way to find out whether your particular table, any table, exists in your database. Sequential is NOT the way. Check this out and let me know if you have any questions.

<P>

<P>Please give me a vote if you like this code :)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Beginner
**User Rating**    |4.4 (111 globes from 25 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/a-better-way-to-test-for-tables-great-for-begginners-well-commented__1-9111/archive/master.zip)





### Source Code

```
IF TableExists(strTableName) then MsgBox strTableName & " found." else MsgBox strTableName & " not found."
Private Function TableExists(TableName) As Boolean
'I ususally use a global Database object, however' you can just as easily pass it into the function if you'd prefer
Dim strTableName$ 'string
On Error GoTo NotFound
If TableName <> "" Then strTableName = dbMyDatabase.TableDefs(strTableName).Name
'If the table exists, the string will be filled, 'otherwise it will err out and TableExists will remain false.
TableExists = True
NotFound:
End Function
'I have VERY often seen people use the standard routine of
'going through EACH and EVERY table comparing each one till
'they get the the end, as in
 'For Each MyTable in DB.TableDefs
 ' if MyTable.Name = strNameImLookingFor then
 'TableExists = true
 'Exit For
 'end if
 'Next
'This is NOT the way to do this. You will unecesesarily use up
'yours as well as your users' very valuable time.
'Use this function. Make it private. When you pass the name
'of the table you need to check for into this routine, the
'recordset will either retrieve it, with a quickness, or it
'will error out, which is even quicker. If you have this in
'a private function, the erroring out will equate to it
'returning a negative response for the table search.
'I might add that this technique works superbly with field searches
'as well (such as Serial No, credit cards, socials, phone numbers, etc).
'And, there you have it.
```

