unit main_u;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Data.Win.ADODB,
  Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Imaging.pngimage;

type
  TfrmUnitMaker = class(TForm)
    btnSelectDB: TButton;
    lsbTables: TListBox;
    btnCreate: TButton;
    Image1: TImage;
    procedure btnSelectDBClick(Sender: TObject);
    procedure btnCreateClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmUnitMaker: TfrmUnitMaker;
  f, g: textfile;
  iNumTables: integer = 0;

implementation

uses shellAPI;

{$R *.dfm}

procedure TfrmUnitMaker.btnCreateClick(Sender: TObject);
var
  i : integer;
  sTemp : string;
  iSelection : byte;
  bFlag : boolean;
  sFile : string;
  sFileDestination : string;
begin
  sTemp := '';

  for i:=0 to lsbTables.count - 1 do
  begin
    sTemp := sTemp + lsbTables.Items[i] + #13;
  end;


  iSelection := messageDlg('You are about to create a unit file for ' + IntToStr(iNumTables) + ' tables: ' + #13 + sTemp + #13 + 'Do you wish to continue?', mtConfirmation ,mbOkCancel,0);

  case iSelection of
    mrOK : bFlag := true;
    mrCancel : begin
                  bFlag := false;
               end;
  end;


  if not(bFlag) then
  begin
    showmessage('Operation cancelled.');
    abort;
  end;

  try
    assignfile(f,'mydb_u.pas');
    rewrite(f);

   //outline & uses
   writeln(f,'unit mydb_u;');
   writeln(f,'');
   writeln(f,'interface');
   writeln(f,'uses Data.DB, Data.Win.ADODB;');
   writeln(f,'');

   //DB type declaration
   writeln(f,'type');
   writeln(f,'TMYDB = class');
   writeln(f,'private');
   writeln(f,'fADOConnection : TADOConnection;');

   sTemp := '';
   //create a variable for each ADO Table
   for i:=0 to lsbTables.Count - 1 do
   begin
     sTemp := sTemp + 'f' + lsbTables.Items[i] + ',';
   end;
   delete(sTemp,length(sTemp),1);

   sTemp := sTemp + ': TADOTable;';

   writeln(f,sTemp);


   //create a variable for each ADO Table Datasource
   sTemp := '';
   for i:=0 to lsbTables.Count - 1 do
   begin
     sTemp := sTemp + 'fDS' + lsbTables.Items[i] + ',';
   end;
   delete(sTemp,length(sTemp),1);

   sTemp := sTemp + ': TDATASource;';

   writeln(f,sTemp);


   //setup SQL
   writeln(f,'fSQLQuery : TADOQuery;');
   writeln(f,'fDSSQLQuery : TDATASource;');


   //functions and procedures
   writeln(f,'public');

   //constructor
   writeln(f,'constructor create(sDBName : string);');

   //ADO Table Accessors
   for i:=0 to lsbTables.Count -1 do
   begin
     writeln(f,'function get' + lsbTables.Items[i] + ' : TADOTable;');
   end;

   //ADO Table Datasource Accessors
   for i:=0 to lsbTables.Count - 1 do
   begin
     writeln(f,'function getDS' + lsbTables.Items[i] + ' : TDATASource;');
   end;


   //SQL Query Accessor
   writeln(f,'function getSQLqry : TADOQuery;');

   //SQL Datasource Accessor
   writeln(f,'function getSQLDS : TDATASource;');

   //SQL Mutator (for queries)
   writeln(f,'procedure runSQL(sQuery : string);');

   //SQL Mutator (for SQL modifications)
   writeln(f,'procedure modSQL(sQuery : string);');

   writeln(f,'end;');

   writeln(f,'implementation');
   writeln(f,'');
   writeln(f,'uses sysutils;');
   writeln(f,'');

   { TMYDB }


   //constructor
   writeln(f,'constructor TMYDB.create(sDBName: string);');
   writeln(f,'var');
   writeln(f,'sLocation : string;');
   writeln(f,'begin');

   writeln(f,'sLocation := getcurrentdir + ' + quotedstr('\database\') +  ' + sDBName;');
   writeln(f,'fADOConnection := TADOConnection.create(nil);');
   writeln(f,'fADOConnection.connectionstring := ' + quotedstr('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=') + ' + sLocation + ' + quotedstr(';Persist Security Info=False') + ';');
   writeln(f,'fADOConnection.loginprompt := false;');
   writeln(f,'fADOConnection.connected := true;');


   //ado Tables
   for i:=0 to lsbTables.Count - 1 do
   begin
     writeln(f,'');
     writeln(f,'f' + lsbTables.Items[i] + ' := TADOTable.create(fADOConnection);');
     writeln(f,'f' + lsbTables.Items[i] + '.Connection := fADOConnection;');
     writeln(f,'f' + lsbTables.Items[i] + '.TableName := ' + quotedstr(lsbTables.Items[i]) + ';');
     writeln(f,'f' + lsbTables.Items[i] + '.Active := TRUE;');
   end;

   //ado Table DataSource
   for i:=0 to lsbTables.Count - 1 do
   begin
     writeln(f,'');
     writeln(f,'fDS' + lsbTables.Items[i] + ' := TDATASource.create(fADOConnection);');
     writeln(f,'fDS' + lsbTables.Items[i] + '.Dataset := f' + lsbTables.Items[i]+ ';');
   end;

   //SQL QRY
   writeln(f,'');
   writeln(f,'fSQLQuery := TADOQuery.create(fADOConnection);');
   writeln(f,'fSQLQuery.connection := fADOConnection;');

   //SQL DS
   writeln(f,'');
   writeln(f,'fDSSQLQuery := TDATASource.create(fADOConnection);');
   writeln(f,'fDSSQLQuery.Dataset := fSQLQuery;');

   writeln(f,'end;');

   //ADO Table Accessors
   for i:=0 to lsbTables.count -1 do
   begin
      writeln(f,'');
      writeln(f,'function TMYDB.get' + lsbTables.Items[i] + ' : TADOTable;');
      writeln(f,'begin');
      writeln(f,'result := f' + lsbTables.Items[i] + ';');
      writeln(f,'end;');
   end;

   //ADO Table DataSource Accessors
   for i:=0 to lsbTables.count -1 do
   begin
      writeln(f,'');
      writeln(f,'function TMYDB.getDS' + lsbTables.Items[i] + ' : TDATASource;');
      writeln(f,'begin');
      writeln(f,'result := fDS' + lsbTables.Items[i] + ';');
      writeln(f,'end;');
   end;

   //SQL Query Accessor
   writeln(f,'');
   writeln(f,'function TMYDB.getSQLqry : TADOQuery;');
   writeln(f,'begin');
   writeln(f,'result := fSQLQuery;');
   writeln(f,'end;');

   //SQL Query DS Accessor
   writeln(f,'');
   writeln(f,'function TMYDB.getSQLDS : TDATASource;');
   writeln(f,'begin');
   writeln(f,'result := fDSSQLQuery;');
   writeln(f,'end;');

   //SQL Mutator (for queries)
   writeln(f,'');
   writeln(f,'procedure TMYDB.runSQL(sQuery : string);');
   writeln(f,'begin');
   writeln(f,'fSQLQuery.SQL.clear;');
   writeln(f,'fSQLQuery.SQL.text := sQuery;');
   writeln(f,'fSQLQuery.open;');
   writeln(f,'end;');

   //SQL Mutator (for SQL modifications)
   writeln(f,'');
   writeln(f,'procedure TMYDB.modSQL(sQuery : string);');
   writeln(f,'begin');
   writeln(f,'fSQLQuery.SQL.clear;');
   writeln(f,'fSQLQuery.SQL.text := sQuery;');
   writeln(f,'fSQLQuery.ExecSQL;');
   writeln(f,'end;');


   writeln(f,'');
   writeln(f,'end.');


 finally
  closefile(f);
  sFile := 'mydb_u.pas';
  sFileDestination := getcurrentdir + '\DB Unit File\mydb_u.pas';

  CreateDir('DB Unit File');
  CopyFile(pchar(sFile),pchar(sFileDestination), FALSE);
  //MoveFile(pchar(sFile), pchar(sFileDestination));
  DeleteFile(sFile);

  sFileDestination := getcurrentdir + '\DB Unit File\';


  //README FILE
  assignfile(g,'README.txt');
  rewrite(g);


  writeln(g,'      .-.                  ');
  writeln(g,'     (o.o)                 ');
  writeln(g,'      |=|   < Welcome!     ');
  writeln(g,'     __|__                 ');
  writeln(g,'   //.=|=.\\               ');
  writeln(g,'  // .=|=. \\              ');
  writeln(g,'  \\ .=|=. //              ');
  writeln(g,'   \\(_=_)//               ');
  writeln(g,'    (:| |:)                ');
  writeln(g,'     || ||                 ');
  writeln(g,'     () ()                 ');
  writeln(g,'     || ||                 ');
  writeln(g,'     || ||                 ');
  writeln(g,'    =='' ''==              ');
  writeln(g,'==========================================================================');
  writeln(g,'Before proceeding, you must ensure that you database file is in a folder');
  writeln(g,'as shown in one of the two directory outlines below.');
  writeln(g,'==========================================================================');
  writeln(g,'PROJECT FOLDER                 ');
  writeln(g,'    |                          ');
  writeln(g,'    |-- main_p.dproj           ');
  writeln(g,'    |-- main_u.pas             ');
  writeln(g,'    |-- main_p.exe             ');
  writeln(g,'    |-- DATABASE               ');
  writeln(g,'           |                   ');
  writeln(g,'           |-- myDBFile.mdb    ');
  writeln(g,'==========================================================================');
  writeln(g,'Alternative project FOLDER structure:');
  writeln(g,'==========================================================================');
  writeln(g,'PROJECT FOLDER                     ');
  writeln(g,'    |                              ');
  writeln(g,'    |-- main_p.dproj               ');
  writeln(g,'    |-- main_u.pas                 ');
  writeln(g,'  WIN32                            ');
  writeln(g,'    |-- DEBUG                      ');
  writeln(g,'           |                       ');
  writeln(g,'           |-- main_p.exe          ');
  writeln(g,'           |-- DATABASE            ');
  writeln(g,'                  |                ');
  writeln(g,'                  | - myDBFile.mdb ');
  writeln(g,'==========================================================================');
  writeln(g,'In order to make use of this unit file, follow the instructions carefully!');
  writeln(g,'');
  writeln(g,'');
  writeln(g,'==========================================================================');
  writeln(g,'1. Copy "mydb_u.pas" to the PROJECT FOLDER');
  writeln(g,'==========================================================================');
  writeln(g,'2. In Delphi, add "mydb_u.pas" to your project');
  writeln(g,'==========================================================================');
  writeln(g,'3. Add the line:   mydb_u    to your USES');
  writeln(g,'==========================================================================');
  writeln(g,'4. Declare a constant, using the name of your database file');
  writeln(g,'   MYDBFILE = ' + quotedstr('YourDBFileName.mdb') + ';');
  writeln(g,'==========================================================================');
  writeln(g,'5. Declare the following GLOBAL variables:');
  writeln(g,'==========================================================================');
  writeln(g,'   objDB : TMYDB;');
  for i:=0 to lsbTables.Count - 1 do
  begin
    writeln(g,'   ' + lsbTables.Items[i] + ' : TADOTable;');
  end;
  writeln(g,'==========================================================================');
  writeln(g,'6. On FORM CREATE add the following lines:');
  writeln(g,'==========================================================================');
  writeln(g,'   objDB := TMYDB.create(MYDBFiLE);');

  for i := 0 to lsbTables.count - 1 do
  begin
    writeln(g,'   ' + lsbTables.Items[i] + ':= objDB.get' + lsbTables.Items[i] + ';');
  end;

  writeln(g,'   qrySQL := objDB.getSQLqry;');

  closefile(g);

  sFile := 'README.txt';
  sFileDestination := getcurrentdir + '\DB Unit File\README.txt';
  CopyFile(pchar(sFile),pchar(sFileDestination), FALSE);
  //MoveFile(pchar(sFile), pchar(sFileDestination));
  DeleteFile(sFile);


  sFileDestination := getcurrentdir + '\DB Unit File\';
  ShellExecute(0, 'open', PChar(sFileDestination), nil, nil, SW_SHOWNORMAL);

  close;

end;



end;

procedure TfrmUnitMaker.btnSelectDBClick(Sender: TObject);
var
  openDialog: TOpenDialog;
  adoConn: TADOConnection;
  sList: TStringList;
  i: byte;
  bValid: boolean;
begin

  btnCreate.Enabled := false;

  bValid := true;
  openDialog := TOpenDialog.Create(self);
  openDialog.InitialDir := getcurrentdir;
  openDialog.Options := [ofFileMustExist];

  openDialog.Filter := 'Access 2003 mdb files|*.mdb';


    if not(openDialog.execute) then
    begin
      abort;
    end
    else


    adoConn := TADOConnection.Create(self);
    adoConn.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' +
    openDialog.FileName + ';Persist Security Info=False';
    adoConn.LoginPrompt := false;
    try
      adoConn.Connected := true;

    except
      showmessage('Incorrect database format');
      adoConn.Free;
      abort;
    end;


    sList := TStringList.Create;



    adoConn.GetTableNames(sList, false);
    lsbTables.Clear;

    if sList.Count = 0 then
    begin
      showmessage('No tables found in database!');
      abort;
    end;

    for i := 0 to sList.Count - 1 do
    begin

      if pos('tbl', sList[i]) <> 1 then
      begin
        showmessage
        ('Your table names are not in the correct format.  Table names should start with a prefix "tbl", e.g. tblUsers, tblItems, etc');
        lsbTables.Clear;
        bValid := false;
        iNumTables := 0;
        abort;
      end;

      lsbTables.Items.Add(sList[i]);
    end;

    if bValid then
    begin
      btnCreate.Enabled := true;
      adoConn.Free;
      iNumTables := lsbTables.Count;
    end;

  end;



end.
