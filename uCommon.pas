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
  //bConnected : Boolean; //��������
  nSerialPort : integer;
  LoopCount : integer ; //���� ��� Ƚ��
  DelayTime : integer;  //���ð�

  bMCUIDCheck : Boolean; //ID Check ���� ����
  bSysInfoCheck:Boolean; //���������� ��������
  bLinkusIDCheck:Boolean; //Linkus ID ��������
  bLinksTelNoCheck :Boolean;  //��� ���� ��ȭ��ȣ üũ
  bRingCountCheck : Boolean;  //��ȭ �Ŵ� Ƚ�� üũ
  bWiznetCheck : Boolean;     //IP �� üũ
  nClose :integer;
  bNext : Boolean;   //���� ���� ����

implementation

{$R *.dfm}

end.
