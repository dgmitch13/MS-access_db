 
Attribute VB_Name = "SMT_Term2"
Option Compare Database
Public Sub Qry(TxtLanID)
'Public Sub Qry()
'Code version 1.1 created on 3/13/13 by Donald Mitchell
'Code version 2.0 created on 3/19/13 change SQL statement by Donald Mitchell
'Code version 2.1 created on 3/20/13 add Form_Form1 for front end input and output by Donald Mitchell
'Code version 2.2 created on 4/11/13 add additional outfield from select statement by Donald Mitchell


    Dim db As DAO.Database
    Dim tbl As DAO.TableDef
    Dim MySQLString As String
    Dim UserStr As String
    Dim SQLselect As String
    
    Dim SQLwhere As String
    Dim rst As DAO.Recordset
    Dim TblNames As String
    Dim TblCount As Integer

    'rst.Close 'close object before creating a new one
    
    Set db = CurrentDb() 'creates a database instance for active database
 'UserStr = "789456" 'input field
 'UserStr = "'dmitch05'" 'input field
UserStr = TxtLanID
' Tblname = tbl.Name
  
'ver. 1.1: this for loop will search each table in current database
For Each tbl In db.TableDefs

   ' If tbl.Attributes > 537 Then 'Ignores System Tables but capture link tables
   If tbl.Attributes = 0 Then

     TblNames = tbl.Name 'move table name from object to string
     
          'Code version 2.0 created on 3/19/13 change SQL statement by Donald Mitchell
          'ver. 1.1: this is the "Select" statement variables. Any "Select" changes should be done at this level
          'to insure query will function properly.
          'ver. 2.2: added additionl fields to capture profile etc. In this case lastname.
          
          
                  
          SQLselect = "Select USUS_ID, lastname from "
          SQLwhere = " where USUS_ID = "
                      
          MySQLString = SQLselect & "[" & TblNames & "]" & SQLwhere & UserStr '
         ' Debug.Print MySQLString
            
            Set rst = db.OpenRecordset(MySQLString, dbOpenDynaset)
           ' rst.MoveLast
                
             Dim strHolder As String
             Dim strFound As String
             
             'ver 2.2 rst.fields hold additional values from select statement
             
                If rst.AbsolutePosition = 0 Then
                    Debug.Print "Found User in " & TblNames & " " & rst.Fields(1)
                    
                   strFound = "Found  User in " & TblNames & " " & rst.Fields(1) & Chr$(13) & Chr$(10)
                    
                    Else
                    Debug.Print "Did not find user " & TblNames
                   strFound = "Did not find user " & TblNames & Chr$(13) & Chr$(10)
                    
                End If
                
'Code version 2.1 add variables to capture string output

            strHolder = strHolder & strFound
            Form_FrmTerm.txtOutput = strHolder
            
                e = rst.RecordCount
       
    TblCount = TblCount + 1
    '  Debug.Print tbl.Name
    End If
'  Debug.Print tbl.Attributes & " " & tbl.Name
Next
   Debug.Print "Search has completed"

MsgBox ("Search has completed")

'Form_Form1.TxtLanID = ""

End Sub

Public Sub testid(TxtLanID)
'Public Sub testid()
'Code version 1.1 created on 3/13/13 by Donald Mitchell
'Code version 2.0 created on 3/19/13 change SQL statement by Donald Mitchell
'Code version 2.1 created on 3/20/13 add Form_Form1 for front end input and output by Donald Mitchell
'Code version 2.2 created on 4/11/13 add additional outfield from select statement by Donald Mitchell


    Dim db As DAO.Database
    Dim tbl As DAO.TableDef
    Dim MySQLString As String
    Dim UserStr As String
    Dim SQLselect As String
    
    Dim SQLwhere As String
    Dim rst As DAO.Recordset
    Dim TblNames As String
    Dim TblCount As Integer

    'rst.Close 'close object before creating a new one
    
    Set db = CurrentDb() 'creates a database instance for active database
 'UserStr = "789456" 'input field
' UserStr = "'kmahav01'" 'input field
UserStr = TxtLanID
' Tblname = tbl.Name
  
'ver. 1.1: this for loop will search each table in current database
For Each tbl In db.TableDefs

   ' If tbl.Attributes > 537 Then 'Ignores System Tables but capture link tables
   If tbl.Attributes = 0 Then

     TblNames = tbl.Name 'move table name from object to string
     
          'Code version 2.0 created on 3/19/13 change SQL statement by Donald Mitchell
          'ver. 1.1: this is the "Select" statement variables. Any "Select" changes should be done at this level
          'to insure query will function properly.
          'ver. 2.2: added additionl fields to capture profile etc. In this case lastname.
          
          'ver 3.0 new code to capture testid
          Dim strlen As Integer
        
          strlen = Len(UserStr) - 2 'minus 2 to take into consideration of the '' used in the UserStr variable
          Debug.Print strlen
          Debug.Print UserStr
          
          SQLselect = "Select USUS_ID, USUS_DESC from "
          SQLwhere = " where "
          SQLfilter = "right(USUS_DESC," & strlen & ") ="
                                          
          MySQLString = SQLselect & "[" & TblNames & "]" & SQLwhere & SQLfilter & UserStr
                  
          Debug.Print MySQLString
                
            
            Set rst = db.OpenRecordset(MySQLString, dbOpenDynaset)
           ' rst.MoveLast
                
             Dim strHolder As String
             Dim strFound As String
             
             'ver 2.2 rst.fields hold additional values from select statement
             
                If rst.AbsolutePosition = 0 Then
                    Debug.Print "Found User in table " & TblNames & " " & rst.Fields(0)
                    
                   strFound = "Found  User in table " & TblNames & " " & rst.Fields(0) & rst.Fields(1) & Chr$(13) & Chr$(10)
                    
                    Else
                    Debug.Print "Did not find user " & TblNames
                   strFound = "Did not find user " & TblNames & Chr$(13) & Chr$(10)
                    
                End If
                
'Code version 2.1 add variables to capture string output

            strHolder = strHolder & strFound
            Form_FrmTerm.txtOutput = strHolder
            
                e = rst.RecordCount
       
    TblCount = TblCount + 1
    '  Debug.Print tbl.Name
    End If
'  Debug.Print tbl.Attributes & " " & tbl.Name
Next
   Debug.Print "Search has completed"

MsgBox ("Search has completed")

'Form_Form1.TxtLanID = ""

End Sub





