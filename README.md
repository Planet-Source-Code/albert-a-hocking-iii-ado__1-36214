<div align="center">

## ADO


</div>

### Description

This is a simple sample in two parts, first project1.vbp is a sample of ADO with text and other types of files. The second part is for reading, writing, adding and editing to a database. This is for people who are looking for simple samples of ADO

The ADO through files has to be observed when the code is running becuase it is only meant for tempdata.

This has been written for windows 95, 98, ME, NT and 2000 but it hasn't been tested in only 2000.

Be sure to follow the instructions from the readme, there are almost no instructions but the one that I give, the code depends on. Also, the database is in Microsoft Access 97.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Albert A Hocking III](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/albert-a-hocking-iii.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/albert-a-hocking-iii-ado__1-36214/archive/master.zip)





### Source Code

```
<pre>
Form1:
Private Sub Command1_Click()
'*****************************************************************************************************************************************************************
'    This is for writing data to a temp file each time overwriting the previous data
'*****************************************************************************************************************************************************************
If Err.Number = 76 Then    'Test for 98
 Open "C:\windows\temp\~00001.tmp" For Output As #1  'This opens the file as an output file
 Print #1, "this is a line"   'This writes a line in the text file
 Print #1, "This is another Line"   'This writes a line in the text file
 Close #1    'This closes the file, otherwise it would remain open until the user restarts their machine
Else
 Open "C:\winnt\temp\~00001.tmp" For Output As #1  'This opens the file as an output file
 Print #1, "this is a line"   'This writes a line in the text file
 Print #1, "This is another Line"   'This writes a line in the text file
 Close #1    'This closes the file, otherwise it would remain open until the user restarts their machine
End If
End Sub
Private Sub Command10_Click()
'*****************************************************************************************************************************************************************
'    This is for writing data to a temp file each time overwriting the previous data in Lock Read write mode
'*****************************************************************************************************************************************************************
If Err.Number = 76 Then    'Test for 98
 Open "C:\windows\temp\~00001.tmp" For Output Lock Read Write As #1 'This opens the file as an output file
 Print #1, "this is a line"   'This writes a line in the text file
 Print #1, "This is another Line"   'This writes a line in the text file
 Close #1    'This closes the file, otherwise it would remain open until the user restarts their machine
Else
 Open "C:\winnt\temp\~00001.tmp" For Output Lock Read Write As #1 'This opens the file as an output file
 Print #1, "this is a line"   'This writes a line in the text file
 Print #1, "This is another Line"   'This writes a line in the text file
 Close #1    'This closes the file, otherwise it would remain open until the user restarts their machine
End If
End Sub
Private Sub Command12_Click()
'*****************************************************************************************************************************************************************
'    This is for writing data to a temp file each time overwriting the previous data in a random
'*****************************************************************************************************************************************************************
Dim FileNumber
FileNumber = 1
If Err.Number = 76 Then    'Test for 98
 Open "C:\windows\temp\~00001.tmp" For Output As #FreeFile 'This opens the file as an output file
 Print #FreeFile, "this is a line"   'This writes a line in the text file
 Print #FreeFile, "This is another Line"  'This writes a line in the text file
 Close #FreeFile    'This closes the file, otherwise it would remain open until the user restarts their machine
Else
 Open "C:\winnt\temp\~00001.tmp" For Output As #FreeFile 'This opens the file as an output file
 Print #FreeFile, "this is a line"   'This writes a line in the text file
 Print #FreeFile, "This is another Line"  'This writes a line in the text file
 Close #FreeFile    'This closes the file, otherwise it would remain open until the user restarts their machine
End If
End Sub
Private Sub Command13_Click()
'*****************************************************************************************************************************************************************
'    This is for writing data to a temp file each time overwriting the previous data beginning and specify record lenth
'*****************************************************************************************************************************************************************
If Err.Number = 76 Then    'Test for 98
 Open "C:\windows\temp\~00001.tmp" For Output As #1 Len = 10000 'This opens the file as an output file
 Print #1, "this is a line"   'This writes a line in the text file
 Print #1, "This is another Line"   'This writes a line in the text file
 Close #1    'This closes the file, otherwise it would remain open until the user restarts their machine
Else
 Open "C:\winnt\temp\~00001.tmp" For Output As #1 Len = 10000 'This opens the file as an output file
 Print #1, "this is a line"   'This writes a line in the text file
 Print #1, "This is another Line"   'This writes a line in the text file
 Close #1    'This closes the file, otherwise it would remain open until the user restarts their machine
End If
End Sub
Private Sub Command2_Click()
'*****************************************************************************************************************************************************************
'    This is for writing data to a temp file each time Appending to the previous data
'*****************************************************************************************************************************************************************
If Err.Number = 76 Then    'Test for 98
 Open "C:\Windows\temp\~00001.tmp" For Append As #1  'This opens the file as an output file
 Print #1, "this is a line"   'This writes a line in the text file
 Print #1, "This is another Line"   'This writes a line in the text file
 Close #1    'This closes the file, otherwise it would remain open until the user restarts their machine
Else
 Open "C:\winnt\temp\~00001.tmp" For Append As #1  'This opens the file as an output file
 Print #1, "This is a line that appended"  'This writes a line in the text file
 Print #1, "This is another Line that appended"  'This writes a line in the text file
 Close #1    'This closes the file, otherwise it would remain open until the user restarts their machine
End If
End Sub
Private Sub Command3_Click()
'*****************************************************************************************************************************************************************
'    This Reads Data from a file
'*****************************************************************************************************************************************************************
Dim retval
 Open "C:\ado\test.ini" For Input As #1  'This opens the file as an output file
 Do Until EOF(1)
 Line Input #1, Data
 retval = MsgBox(Data, vbOKOnly, "Data")
 Loop
 Close #1    'This closes the file, otherwise it would remain open until the user restarts their machine
End Sub
Private Sub Command4_Click()
'*****************************************************************************************************************************************************************
'    This is for writing data to a temp file each time Binary
'*****************************************************************************************************************************************************************
If Err.Number = 76 Then    'Test for 98
 Open "C:\Windows\temp\~00001.tmp" For Binary Access Write As #1 'This opens the file as an output file
 Put #1, 30, "This is a line that appended"  'This writes a line in the text file
 Put #1, 100, " "
 Put #1, 500, "This is another Line that appended"
 Close #1    'This closes the file, otherwise it would remain open until the user restarts their machine
Else
 Open "C:\winnt\temp\~00001.tmp" For Binary Access Write As #1 'This opens the file as an output file
 Put #1, 30, "This is a line that appended"  'This writes a line in the text file
 Put #1, 100, " "
 Put #1, 500, "This is another Line that appended"  'This writes a line in the text file
 Close #1    'This closes the file, otherwise it would remain open until the user restarts their machine
End If
End Sub
Private Sub Command5_Click()
'*****************************************************************************************************************************************************************
'    This is for writing in a random order
'*****************************************************************************************************************************************************************
Open "c:\ado\test.db" For Random As #1
 Put #1, 1, "This is a line that appended"  'This writes a line in the text file
 Put #1, 2, "This is another Line that appended"  'This writes a line in the text file
Close #1    'This closes the file, otherwise it would remain open until the user restarts their machine
End Sub
Private Sub Command7_Click()
'*****************************************************************************************************************************************************************
'    This is for writing data to a temp file each time overwriting the previous data in shared mode
'*****************************************************************************************************************************************************************
If Err.Number = 76 Then    'Test for 98
 Open "C:\windows\temp\~00001.tmp" For Output Shared As #1 'This opens the file as an output file
 Print #1, "this is a line"   'This writes a line in the text file
 Print #1, "This is another Line"   'This writes a line in the text file
 Close #1    'This closes the file, otherwise it would remain open until the user restarts their machine
Else
 Open "C:\winnt\temp\~00001.tmp" For Output Shared As #1 'This opens the file as an output file
 Print #1, "this is a line"   'This writes a line in the text file
 Print #1, "This is another Line"   'This writes a line in the text file
 Close #1    'This closes the file, otherwise it would remain open until the user restarts their machine
End If
End Sub
Private Sub Command8_Click()
'*****************************************************************************************************************************************************************
'    This is for writing data to a temp file each time overwriting the previous data in Lock write mode
'*****************************************************************************************************************************************************************
If Err.Number = 76 Then    'Test for 98
 Open "C:\windows\temp\~00001.tmp" For Output Lock Write As #1 'This opens the file as an output file
 Print #1, "this is a line"   'This writes a line in the text file
 Print #1, "This is another Line"   'This writes a line in the text file
 Close #1    'This closes the file, otherwise it would remain open until the user restarts their machine
Else
 Open "C:\winnt\temp\~00001.tmp" For Output Lock Write As #1 'This opens the file as an output file
 Print #1, "this is a line"   'This writes a line in the text file
 Print #1, "This is another Line"   'This writes a line in the text file
 Close #1    'This closes the file, otherwise it would remain open until the user restarts their machine
End If
End Sub
Private Sub Command9_Click()
'*****************************************************************************************************************************************************************
'    This is for writing data to a temp file each time overwriting the previous data in Lock Read mode
'*****************************************************************************************************************************************************************
Dim retval
 Open "C:\ado\test.ini" For Input Lock Read As #1  'This opens the file as an output file
 Do Until EOF(1)
 Line Input #1, Data
 retval = MsgBox(Data, vbOKOnly, "Data")
 Loop
 Close #1    'This closes the file, otherwise it would remain open until the user restarts their machine
End Sub
Private Sub Form_Load()
On Error Resume Next    'This is for continuing on errors
FileCopy "C:\ADO\test.txt", "C:\windows\temp\~00001.tmp" 'This is for Windows 98
If Err.Number = 76 Then    'This is the error code for path not found(This is one way to error handle)
 FileCopy "C:\ADO\test.txt", "C:\winnt\temp\~00001.tmp" 'This is encase you want to use a temp file to write temporary data in windows NT or 2000
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Kill "C:\windows\temp\~00001.tmp"   'This deletes the temp File for 98
If Err.Number = 53 Then
 Kill "C:\winnt\temp\~00001.tmp"   'This deletes the temp File for NT and 2000
End If
End Sub
'*****************************************************************************************************************************************************************
'error 76 = Path not found
'error 53 = file not found
'*****************************************************************************************************************************************************************
This is part 2, writing to a database.
Form 1:
'**********************************************************************************************************************************************************************************
'     Writing to and from a Database
'     Albert A. Hocking III
'**********************************************************************************************************************************************************************************
'     Getting the code to work
'**********************************************************************************************************************************************************************************
'This is from the data form wizard from the addin menu
'This also, is the sample code that you can use for a template
'there are a few things that you have to do to get this working.
'You have to include the data access components so that the database
'can be recognized. This is done by
'On the menu bar Project->References->Microsoft Data Access Objects 2.1 Library(or 2.5 if you want)
'**********************************************************************************************************************************************************************************
'     Using the Data-form Wizard
'**********************************************************************************************************************************************************************************
'If you want to use the data form wizard then,
'Add-Ins->Addin Manager->VB 6 Dataform Wizard.
'This click on Loaded and unloaded and then click load on startup or you will have to do it each time.
'
'After that your done, it should work after you update your ADO folder
'**********************************************************************************************************************************************************************************
Private Sub Command1_Click()
Me.Hide
frmProducts.Show
End Sub
Private Sub Command2_Click()
End
End Sub
Private Sub Command3_Click()
Me.Hide
frmProducts1.Show
End Sub
Private Sub Command5_Click()
Form2.Show
Me.Hide
End Sub
Private Sub Form_Load()
Form1.Top = (Screen.Height - Form1.Height) / 2
Form1.Left = (Screen.Width - Form1.Width) / 2
End Sub
Form2:
'*************************************************************************************************************************************************************************************
'    This form was created by me. This more willingly gives the code
'    for Adding Editing and Deleting Records in a table in a database
'    Be sure if you want to edit the code
'    The specify "ADO code in the wizard with the proper radio button
'    SEE THE COMMENTS FOR THE EXPINATION OF WHAT I'M DOING
'*************************************************************************************************************************************************************************************
Private Sub Command1_Click()     'This routine if for moving to the previous record
On Error Resume Next      'Error Handeling
 Dim db As Connection      'Sets the variable for the database
 Dim rs As Recordset      'Sets the variable for the record set
 Dim retval As Variant      'A BS variable for for msgbox's
 Set db = New Connection     'Specifies a new connection to the database
 Set rs = New Recordset     'does the same aas the recordset
 db.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;" & "Data Source=C:\ADO\Northwind.mdb;" 'Loads the Microsoft Access 97 Driver
 db.Open      'Loads opens the database that is in the line above
 rs.Open "SELECT ProductID, ProductName FROM products", db, adOpenDynamic, adLockPessimistic 'Opens the table with SQL, specifies the Database through the variable, opens it with dynamic so if other users change it, and has a pecimistic lock(that means that when you bite on a record, its locked until you move to another one, otherwise, its locked to other users for update)
 rs.MoveFirst      'Moves to the first record so that the correct on can be found
 If Me.Text2 = rs!ProductID Then     'Once it has moved to the first record, if the text box that has the key is equals the first record in the table then it will not go back
 retval = MsgBox("You have reached the beginning of the Table", vbCritical, "EOF")  'This is the message for the condition when the user can't go back
 Exit Sub      'Ends the code
 End If      'Ends the condition
 If Err.Number = 3021 Then     'This test for the "EOF or BOF is true or the previous record has been deleted"
 retval = MsgBox("Your are at the beginning of the recordset", vbCritical, "An error has ocurred") 'This is the message for the Error
 End If      'This ends the condition
 Me.Text2.Text = Me.Text2.Text - 1     'This subtracts one value from the Key box for comparision with the table
 Do While Not rs.EOF      'Loop that rolls through the table
 rs.MoveNext      'Moves to the next record
 If rs!ProductID = Me.Text2.Text Then    'Tests for the previous record
 Me.Text1 = rs!ProductName     'Places the value in the text box
 Me.Text2 = rs!ProductID     'Places the value in the text box
 Exit Sub      'Exits the Subroutine
 End If      'Ends the Condition
 Loop      'For the repition structure
 Set db = Nothing      'Kills the connection to the Database(DON'T LEAVE THIS OUT, YOU WILL WAIST MEMORY)
 Set rs = Nothing      'This does the same with the recordset
End Sub
Private Sub Command2_Click()     'This moves to the next record - AT THIS POINT i WILL NOT REEXPLAIN PREVIOUS CODE
On Error Resume Next
 Dim db As Connection
 Dim rs As Recordset
 Dim retval As Variant
 Set db = New Connection
 Set rs = New Recordset
 db.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;" & "Data Source=C:\ADO\Northwind.mdb;"
 db.Open
 rs.Open "SELECT ProductID, ProductName FROM products", db, adOpenDynamic, adLockPessimistic
 rs.MoveLast      'This moves the last record
 If Me.Text2 = rs!ProductID Then     'This tests to see if user is on the last record
 retval = MsgBox("You have reached the end of the Table", vbCritical, "EOF")  'This displays a message box
 Exit Sub      'This exits the subroutine
 End If      'Ends the condition
 rs.MoveFirst
 Me.Text2.Text = Me.Text2.Text + 1     'Adds one to the key value so that it can be compired
 If Err.Number = 3021 Then
 retval = MsgBox("Your are at the beginning of the recordset", vbCritical, "An error has ocurred")
 Else
 Do Until rs.EOF
 rs.MoveNext
 If rs!ProductID = Me.Text2.Text Then
 Me.Text1 = rs!ProductName
 Me.Text2 = rs!ProductID
 Exit Sub
 End If
 Loop
 End If
 Set db = Nothing
 Set rs = Nothing
End Sub
Private Sub Command3_Click()     'This edits a recordset
On Error Resume Next
 Dim db As Connection
 Dim rs As Recordset
 Dim retval As Variant
 Set db = New Connection
 Set rs = New Recordset
 db.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;" & "Data Source=C:\ADO\Northwind.mdb;"
 db.Open
 rs.Open "SELECT ProductID, ProductName FROM products", db, adOpenDynamic, adLockPessimistic
 rs.MoveFirst
 If Err.Number = 3021 Then
 retval = MsgBox("Your are at the beginning of the recordset", vbCritical, "An error has ocurred")
 Else
 Do Until rs.EOF
 rs.MoveNext
 If rs!ProductID = Me.Text2.Text Then
 rs!ProductName = Me.Text1     'This gives the same value to the table record(and its buffered)
 rs.Update      'This Writes the value to the database
 rs!ProductID = Me.Text2     'same Code
 rs.Update      'Same Code
 Exit Sub
 End If
 Loop
 End If
 Set db = Nothing
 Set rs = Nothing
End Sub
Private Sub Command4_Click()
Me.Hide
Form1.Show
End Sub
Private Sub Command5_Click()
Me.Text1.Text = ""      'Clears the Boxes of the values
Me.Text2.Text = ""
End Sub
Private Sub Command7_Click()
 Dim db As Connection
 Dim rs As Recordset
 Dim retval As Variant
 Set db = New Connection
 Set rs = New Recordset
 db.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;" & "Data Source=C:\ADO\Northwind.mdb;"
 db.Open
 rs.Open "SELECT * FROM Products", db, adOpenDynamic, adLockPessimistic
 rs.MoveFirst
 Do Until rs.EOF
 If rs!ProductID = Me.Text2.Text Then
 rs.Delete      'Prepares the record to be deleted
 rs.Update      'Updates the recordset
 rs.MoveNext      'Moves to the next record
 Me.Text1 = rs!ProductName     'Populates the box with the next availble record
 Me.Text2 = rs!ProductID     'same code different field
 Exit Sub
 End If
 rs.MoveNext
 Loop
 Set db = Nothing
 Set rs = Nothing
End Sub
Private Sub Command8_Click()
 Dim db As Connection
 Dim rs As Recordset
 Dim retval As Variant
 Set db = New Connection
 Set rs = New Recordset
 db.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;" & "Data Source=C:\ADO\Northwind.mdb;"
 db.Open
 rs.Open "SELECT ProductID, ProductName FROM products", db, adOpenDynamic, adLockPessimistic
 rs.MoveLast
 rs.AddNew      'Tells the database to go to the end
 rs!ProductName = Me.Text1     'transfurs the data
 rs.Update      'Writes the data
 Set db = Nothing
 Set rs = Nothing
End Sub
Private Sub Form_Load()      'Displays the first record
Form2.Top = (Screen.Height - Form2.Height) / 2
Form2.Left = (Screen.Width - Form2.Width) / 2
 Dim db As Connection
 Dim rs As Recordset
 Set db = New Connection
 Set rs = New Recordset
 db.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;" & "Data Source=C:\ADO\Northwind.mdb;"
 db.Open
 rs.Open "SELECT ProductID, ProductName FROM products", db, adOpenDynamic, adLockPessimistic
 rs.MoveFirst
 Me.Text1 = rs!ProductName
 Me.Text2 = rs!ProductID
 Set db = Nothing
 Set rs = Nothing
 Form1.Show
End Sub
Private Sub Form_Unload(Cancel As Integer)    'Property for the exit
Me.Text1.Text = ""      'Clears the textbox
Me.Text2.Text = ""      'Clears the textbox
Form1.Show      'Shows the main form
End Sub
frmProducts:
'*************************************************************************************************************************************************************************************
'    This form is created with the form wizard
'    Be sure if you want to edit the code
'    The specify "ADO code in the wizard with the proper radio button
'*************************************************************************************************************************************************************************************
Dim WithEvents adoPrimaryRS As Recordset
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Private Sub Form_Load()
frmProducts.Top = (Screen.Height - frmProducts.Height) / 2
frmProducts.Left = (Screen.Width - frmProducts.Width) / 2
 Dim db As Connection
 Set db = New Connection
 db.CursorLocation = adUseClient
 db.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=C:\ADO\Northwind.mdb;"
 Set adoPrimaryRS = New Recordset
 adoPrimaryRS.Open "select Discontinued,ProductName,QuantityPerUnit,ReorderLevel,SupplierID,UnitPrice,UnitsInStock,UnitsOnOrder from Products", db, adOpenStatic, adLockOptimistic
 Set grdDataGrid.DataSource = adoPrimaryRS
 mbDataChanged = False
End Sub
Private Sub Form_Resize()
 On Error Resume Next
 'This will resize the grid when the form is resized
 grdDataGrid.Height = Me.ScaleHeight - 30 - picButtons.Height - picStatBox.Height
 lblStatus.Width = Me.Width - 1500
 cmdNext.Left = lblStatus.Width + 700
 cmdLast.Left = cmdNext.Left + 340
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If mbEditFlag Or mbAddNewFlag Then Exit Sub
 Select Case KeyCode
 Case vbKeyEscape
 cmdClose_Click
 Case vbKeyEnd
 cmdLast_Click
 Case vbKeyHome
 cmdFirst_Click
 Case vbKeyUp, vbKeyPageUp
 If Shift = vbCtrlMask Then
 cmdFirst_Click
 Else
 cmdPrevious_Click
 End If
 Case vbKeyDown, vbKeyPageDown
 If Shift = vbCtrlMask Then
 cmdLast_Click
 Else
 cmdNext_Click
 End If
 End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Screen.MousePointer = vbDefault
 Form1.Show
End Sub
Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
 'This will display the current record position for this recordset
 lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub
Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
 'This is where you put validation code
 'This event gets called when the following actions occur
 Dim bCancel As Boolean
 Select Case adReason
 Case adRsnAddNew
 Case adRsnClose
 Case adRsnDelete
 Case adRsnFirstChange
 Case adRsnMove
 Case adRsnRequery
 Case adRsnResynch
 Case adRsnUndoAddNew
 Case adRsnUndoDelete
 Case adRsnUndoUpdate
 Case adRsnUpdate
 End Select
 If bCancel Then adStatus = adStatusCancel
End Sub
Private Sub cmdAdd_Click()
 On Error GoTo AddErr
 adoPrimaryRS.MoveLast
 adoPrimaryRS.AddNew
 grdDataGrid.SetFocus
 Exit Sub
AddErr:
 MsgBox Err.Description
End Sub
Private Sub cmdDelete_Click()
 On Error GoTo DeleteErr
 With adoPrimaryRS
 .Delete
 .MoveNext
 If .EOF Then .MoveLast
 End With
 Exit Sub
DeleteErr:
 MsgBox Err.Description
End Sub
Private Sub cmdRefresh_Click()
 'This is only needed for multi user apps
 On Error GoTo RefreshErr
 Set grdDataGrid.DataSource = Nothing
 adoPrimaryRS.Requery
 Set grdDataGrid.DataSource = adoPrimaryRS
 Exit Sub
RefreshErr:
 MsgBox Err.Description
End Sub
Private Sub cmdEdit_Click()
 On Error GoTo EditErr
 lblStatus.Caption = "Edit record"
 mbEditFlag = True
 SetButtons False
 Exit Sub
EditErr:
 MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
 On Error Resume Next
 SetButtons True
 mbEditFlag = False
 mbAddNewFlag = False
 adoPrimaryRS.CancelUpdate
 If mvBookMark > 0 Then
 adoPrimaryRS.Bookmark = mvBookMark
 Else
 adoPrimaryRS.MoveFirst
 End If
 mbDataChanged = False
End Sub
Private Sub cmdUpdate_Click()
 On Error GoTo UpdateErr
 adoPrimaryRS.UpdateBatch adAffectAll
 If mbAddNewFlag Then
 adoPrimaryRS.MoveLast 'move to the new record
 End If
 mbEditFlag = False
 mbAddNewFlag = False
 SetButtons True
 mbDataChanged = False
 Exit Sub
UpdateErr:
 MsgBox Err.Description
End Sub
Private Sub cmdClose_Click()
 Unload Me
End Sub
Private Sub cmdFirst_Click()
 On Error GoTo GoFirstError
 adoPrimaryRS.MoveFirst
 mbDataChanged = False
 Exit Sub
GoFirstError:
 MsgBox Err.Description
End Sub
Private Sub cmdLast_Click()
 On Error GoTo GoLastError
 adoPrimaryRS.MoveLast
 mbDataChanged = False
 Exit Sub
GoLastError:
 MsgBox Err.Description
End Sub
Private Sub cmdNext_Click()
 On Error GoTo GoNextError
 If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
 If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
 Beep
 'moved off the end so go back
 adoPrimaryRS.MoveLast
 End If
 'show the current record
 mbDataChanged = False
 Exit Sub
GoNextError:
 MsgBox Err.Description
End Sub
Private Sub cmdPrevious_Click()
 On Error GoTo GoPrevError
 If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
 If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
 Beep
 'moved off the end so go back
 adoPrimaryRS.MoveFirst
 End If
 'show the current record
 mbDataChanged = False
 Exit Sub
GoPrevError:
 MsgBox Err.Description
End Sub
Private Sub SetButtons(bVal As Boolean)
 cmdAdd.Visible = bVal
 cmdEdit.Visible = bVal
 cmdUpdate.Visible = Not bVal
 cmdCancel.Visible = Not bVal
 cmdDelete.Visible = bVal
 cmdClose.Visible = bVal
 cmdRefresh.Visible = bVal
 cmdNext.Enabled = bVal
 cmdFirst.Enabled = bVal
 cmdLast.Enabled = bVal
 cmdPrevious.Enabled = bVal
End Sub
frmProducts1:
'*************************************************************************************************************************************************************************************
'    This is the default form is created with the form wizard
'    Be sure if you want to edit the code
'    The specify "ADO code in the wizard with the proper radio button
'*************************************************************************************************************************************************************************************
Dim WithEvents adoPrimaryRS As Recordset
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Private Sub Form_Load()
 Dim db As Connection
 Set db = New Connection
 db.CursorLocation = adUseClient
 db.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=C:\ADO\Northwind.mdb;"
 Set adoPrimaryRS = New Recordset
 adoPrimaryRS.Open "select ProductID,ProductName from Products", db, adOpenStatic, adLockOptimistic
 Dim oText As TextBox
 'Bind the text boxes to the data provider
 For Each oText In Me.txtFields
 Set oText.DataSource = adoPrimaryRS
 Next
 mbDataChanged = False
End Sub
Private Sub Form_Resize()
 On Error Resume Next
 lblStatus.Width = Me.Width - 1500
 cmdNext.Left = lblStatus.Width + 700
 cmdLast.Left = cmdNext.Left + 340
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
frmProducts1.Top = (Screen.Height - frmProducts1.Height) / 2
frmProducts1.Left = (Screen.Width - frmProducts1.Width) / 2
 If mbEditFlag Or mbAddNewFlag Then Exit Sub
 Select Case KeyCode
 Case vbKeyEscape
 cmdClose_Click
 Case vbKeyEnd
 cmdLast_Click
 Case vbKeyHome
 cmdFirst_Click
 Case vbKeyUp, vbKeyPageUp
 If Shift = vbCtrlMask Then
 cmdFirst_Click
 Else
 cmdPrevious_Click
 End If
 Case vbKeyDown, vbKeyPageDown
 If Shift = vbCtrlMask Then
 cmdLast_Click
 Else
 cmdNext_Click
 End If
 End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Screen.MousePointer = vbDefault
 Form1.Show
End Sub
Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
 'This will display the current record position for this recordset
 lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub
Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
 'This is where you put validation code
 'This event gets called when the following actions occur
 Dim bCancel As Boolean
 Select Case adReason
 Case adRsnAddNew
 Case adRsnClose
 Case adRsnDelete
 Case adRsnFirstChange
 Case adRsnMove
 Case adRsnRequery
 Case adRsnResynch
 Case adRsnUndoAddNew
 Case adRsnUndoDelete
 Case adRsnUndoUpdate
 Case adRsnUpdate
 End Select
 If bCancel Then adStatus = adStatusCancel
End Sub
Private Sub cmdAdd_Click()
 On Error GoTo AddErr
 With adoPrimaryRS
 If Not (.BOF And .EOF) Then
 mvBookMark = .Bookmark
 End If
 .AddNew
 lblStatus.Caption = "Add record"
 mbAddNewFlag = True
 SetButtons False
 End With
 Exit Sub
AddErr:
 MsgBox Err.Description
End Sub
Private Sub cmdDelete_Click()
 On Error GoTo DeleteErr
 With adoPrimaryRS
 .Delete
 .MoveNext
 If .EOF Then .MoveLast
 End With
 Exit Sub
DeleteErr:
 MsgBox Err.Description
End Sub
Private Sub cmdRefresh_Click()
 'This is only needed for multi user apps
 On Error GoTo RefreshErr
 adoPrimaryRS.Requery
 Exit Sub
RefreshErr:
 MsgBox Err.Description
End Sub
Private Sub cmdEdit_Click()
 On Error GoTo EditErr
 lblStatus.Caption = "Edit record"
 mbEditFlag = True
 SetButtons False
 Exit Sub
EditErr:
 MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
 On Error Resume Next
 SetButtons True
 mbEditFlag = False
 mbAddNewFlag = False
 adoPrimaryRS.CancelUpdate
 If mvBookMark > 0 Then
 adoPrimaryRS.Bookmark = mvBookMark
 Else
 adoPrimaryRS.MoveFirst
 End If
 mbDataChanged = False
End Sub
Private Sub cmdUpdate_Click()
 On Error GoTo UpdateErr
 adoPrimaryRS.UpdateBatch adAffectAll
 If mbAddNewFlag Then
 adoPrimaryRS.MoveLast 'move to the new record
 End If
 mbEditFlag = False
 mbAddNewFlag = False
 SetButtons True
 mbDataChanged = False
 Exit Sub
UpdateErr:
 MsgBox Err.Description
End Sub
Private Sub cmdClose_Click()
 Unload Me
End Sub
Private Sub cmdFirst_Click()
 On Error GoTo GoFirstError
 adoPrimaryRS.MoveFirst
 mbDataChanged = False
 Exit Sub
GoFirstError:
 MsgBox Err.Description
End Sub
Private Sub cmdLast_Click()
 On Error GoTo GoLastError
 adoPrimaryRS.MoveLast
 mbDataChanged = False
 Exit Sub
GoLastError:
 MsgBox Err.Description
End Sub
Private Sub cmdNext_Click()
 On Error GoTo GoNextError
 If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
 If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
 Beep
 'moved off the end so go back
 adoPrimaryRS.MoveLast
 End If
 'show the current record
 mbDataChanged = False
 Exit Sub
GoNextError:
 MsgBox Err.Description
End Sub
Private Sub cmdPrevious_Click()
 On Error GoTo GoPrevError
 If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
 If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
 Beep
 'moved off the end so go back
 adoPrimaryRS.MoveFirst
 End If
 'show the current record
 mbDataChanged = False
 Exit Sub
GoPrevError:
 MsgBox Err.Description
End Sub
Private Sub SetButtons(bVal As Boolean)
 cmdAdd.Visible = bVal
 cmdEdit.Visible = bVal
 cmdUpdate.Visible = Not bVal
 cmdCancel.Visible = Not bVal
 cmdDelete.Visible = bVal
 cmdClose.Visible = bVal
 cmdRefresh.Visible = bVal
 cmdNext.Enabled = bVal
 cmdFirst.Enabled = bVal
 cmdLast.Enabled = bVal
 cmdPrevious.Enabled = bVal
End Sub
</pre>
```

