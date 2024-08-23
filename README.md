# Delphi_DB_Connect
A program that will create a unit file to work with your database.  

Simply create an MS Access 2002/2003 mdb file.  The tables in your database must have prefix "tbl", for example
tblUsers, tblProducts, etc.

The program will create a unit file which you can add to your project.  Using the object, you are able to work with TADOTables, TADOQuery as well as SQL.  Instructions are included in the generated text file.

EXAMPLE USES:

VAR
  objDB : TMYDB;  //TMYDB class contained in generated file
  tblUsers : TADOTable;
  qrySQL : TADOQuery;

BEGIN
  //instantiate the object using the name of your database.
  objDB := TMYDB.create('MyDBFile.mdb');

  //create an alias for your tables
  tblUsers := objDB.gettblUsers;

  //You can now perform TADOTable operations using the usual commands, e.g.
  tblUsers.first;
  tblUsers.insert;
  while not tblUsers.EOF do
  begin
  end;
  etc.

  //Display on a DBGrid
  //You can view TADOTable data as well as TADOQuery
  dbgridUSers.datasource := objDB.getDStblUsers;
  dbGridSQL.datasource := objDB.getDSSQL;

  //Run SQL queries
  runSQL('SELECT * FROM tblUsers');

  //Modify the DB using SQL
  modSQL('DELETE FROM tblUsers WHERE username = "0001");

  //You can also traverse the SQL Query
  qrySQL := objDB.getSQLqry;

  qrySQL.first;
  //etc.
  
  
  

  
  
