unit uCommon;

interface

uses
  SysUtils, Classes;

type
  TDataModule1 = class(TDataModule)
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DataModule1: TDataModule1;
  //bConnected : Boolean; //접속유무
  nSerialPort : integer;
  LoopCount : integer ; //에러 대기 횟수
  DelayTime : integer;  //대기시간

  bMCUIDCheck : Boolean; //ID Check 수신 유무
  bSysInfoCheck:Boolean; //방범사용유무 수신유무
  bLinkusIDCheck:Boolean; //Linkus ID 수신유무
  bLinksTelNoCheck :Boolean;  //방범 전송 전화번호 체크
  bRingCountCheck : Boolean;  //전화 거는 횟수 체크
  bWiznetCheck : Boolean;     //IP 및 체크
  nClose :integer;
  bNext : Boolean;   //다음 진행 유무

implementation

{$R *.dfm}

end.
