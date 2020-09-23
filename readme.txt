Most of the features of SQL Server Developer are easily discovered by simply playing with
the application.  The purpose of this readme file is to inform you of some things that are
not.  

Clicking on a database will populate the top portion of the screen with the tables, stored
procedures and/or triggers that exist for that database.

Clicking on a table will populate a column list.  Double clicking on the table listbox will
produce a sorted list easier to view.

Clicking on a column name will append it to the SQL text box.

Double clicking on a stored procedure or trigger will render the contents of the stored
procedure or trigger in a large text box which can be edited and altered back to the SQL
Server.

Double clicking on a recordset will render a full screen view of the recordset. The buttons
< and > on top or below a recordset will increase or decrease the column sizes.  Right click
on a recordset and see what happens.

If you plan on using the stored procedure debugger, also make sure you have an empty line
after Inserts, Deletes, Updates and Deletes.  Use the Set command to populate the paramaters
you will be using (for example Set @Name = 'George'). If using a cursor, always use the
Do While @@rowcount<>0.

