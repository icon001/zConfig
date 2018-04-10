program zConfig;

uses
  Forms,
  uMain in 'uMain.pas' {Form_Main},
  uCommon in 'uCommon.pas' {DataModule1: TDataModule},
  DIMime in 'DIMime.pas',
  DIMimeStreams in 'DIMimeStreams.pas',
  uLomosUtil in 'uLomosUtil.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm_Main, Form_Main);
  Application.CreateForm(TDataModule1, DataModule1);
  Application.Run;
end.
