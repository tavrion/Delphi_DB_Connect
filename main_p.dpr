program main_p;

uses
  Vcl.Forms,
  main_u in 'main_u.pas' {frmUnitMaker};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmUnitMaker, frmUnitMaker);
  Application.Run;
end.
