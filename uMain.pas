{***************************************************************}
{                                                               }
{  uMain.Pas :���α��� ��Ͽ� ���α׷� RS-232 ����              }
{                                                               }
{  Copyright (c) 2005 this70@naver.com                          }
{                                                               }
{  All rights reserved.                                         }
{                                                               }
{***************************************************************}

unit uMain;

interface

uses
  DateUtils,
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, LMDAssist, LMDCustomControl, LMDCustomPanel,
  LMDCustomBevelPanel, LMDCustomParentPanel, LMDCustomPanelFill,
  LMDPanelFill, LMDCustomGroupBox, LMDCustomButtonGroup,
  LMDCustomRadioGroup, LMDTextRadioGroup, jpeg, ExtCtrls, LMDControl,
  LMDBaseControl, LMDBaseGraphicButton, LMDCustomShapeButton,
  LMDShapeButton, StdCtrls, LMDCustomListBox, LMDCustomImageListBox,
  LMDTextListBox, LMDCustomScrollBox, LMDListBox, RzButton, RzRadChk, Mask,
  RzEdit, RzLabel, RzPanel, RzBtnEdt, RzSpnEdt, RzPrgres, RzStatus, OoMisc,
  AdStatLt, AdPort, RzCmboBx, ComCtrls, LMDBaseGraphicControl,
  LMDBaseLabel, LMDCustomSimpleLabel, LMDSimpleLabel, LMDCaptionPanel,
  RzTabs, LMDCustomComponent, LMDIniCtrl, DB, dbisamtb,WinSpool
;

const
  TimeoutTicks = 18 * 2; { about 2 seconds }
  MAX_COMPORT = 100;
  
type
  TForm_Main = class(TForm)
    RzStatusBar1: TRzStatusBar;
    RzClockStatus1: TRzClockStatus;
    RzStatusPane1: TRzStatusPane;
    ComPort: TApdComPort;
    led_TX: TApdStatusLight;
    led_RX: TApdStatusLight;
    Label8: TLabel;
    Label9: TLabel;
    RzFieldStatus1: TRzFieldStatus;
    cbSerialPort: TRzComboBox;
    brnConnect: TRzBitBtn;
    Panel1: TPanel;
    ApdSLController: TApdSLController;
    brndisConnect: TRzBitBtn;
    Off_Timer: TTimer;
    LMDCaptionPanel1: TLMDCaptionPanel;
    LMDSimpleLabel1: TLMDSimpleLabel;
    ProgressBar3: TProgressBar;
    PollingTimer: TTimer;
    OpenDialog1: TOpenDialog;
    cb_Next: TCheckBox;
    LMDIniCtrl1: TLMDIniCtrl;
    Panel2: TPanel;
    Notebook1: TNotebook;
    GroupBox1: TGroupBox;
    Image1: TImage;
    LMDShapeButton5: TLMDShapeButton;
    LMDShapeButton6: TLMDShapeButton;
    LMDShapeButton4: TLMDShapeButton;
    LMDShapeButton3: TLMDShapeButton;
    LMDShapeButton2: TLMDShapeButton;
    LMDShapeButton1: TLMDShapeButton;
    LMDTextListBox1: TLMDTextListBox;
    RzBitBtn1: TRzBitBtn;
    btn_Prv: TRzBitBtn;
    btn_Next: TRzBitBtn;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    RzLabel1: TRzLabel;
    RzLabel2: TRzLabel;
    RzLabel3: TRzLabel;
    Edit_DeviceID1: TRzEdit;
    Edit_DeviceID2: TRzEdit;
    Edit_DeviceID3: TRzEdit;
    Edit_DeviceID4: TRzEdit;
    Edit_DeviceID5: TRzEdit;
    Edit_DeviceID6: TRzEdit;
    Edit_DeviceID7: TRzEdit;
    ComboBox_DoorType1: TRzComboBox;
    Edit_Name: TRzEdit;
    btnIDCheck: TRzBitBtn;
    btnIDSet: TRzBitBtn;
    Group_Linkus: TGroupBox;
    RzLabel9: TRzLabel;
    RzLabel17: TRzLabel;
    Label79: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Edit_LinkusID: TRzEdit;
    Edit_LinkusTel: TRzEdit;
    Spinner_Ring: TRzSpinner;
    btnCheckLinkusID: TRzBitBtn;
    btnRegLinkusID: TRzBitBtn;
    LMDPanelFill1: TLMDPanelFill;
    GroupBox4: TGroupBox;
    LMDPanelFill2: TLMDPanelFill;
    RzGroupBox1: TRzGroupBox;
    Label1: TLabel;
    LMDListBox1: TLMDListBox;
    RzGroupBox8: TRzGroupBox;
    RzLabel8: TRzLabel;
    RzLabel10: TRzLabel;
    Edit_ServerIp: TRzEdit;
    Edit_Serverport: TRzEdit;
    Label2: TLabel;
    RzGroupBox7: TRzGroupBox;
    RzLabel4: TRzLabel;
    RzLabel5: TRzLabel;
    RzLabel6: TRzLabel;
    RzLabel7: TRzLabel;
    Label10: TLabel;
    Edit_LocalIP: TRzEdit;
    Edit_Subnet: TRzEdit;
    Edit_Gateway: TRzEdit;
    Edit_LocalPort: TRzEdit;
    Checkbox_DHCP: TRzCheckBox;
    btnCheckwiznet: TRzBitBtn;
    btnSetwiznet: TRzBitBtn;
    GroupBox5: TGroupBox;
    LMDPanelFill3: TLMDPanelFill;
    GroupBox6: TGroupBox;
    RzLabel21: TRzLabel;
    RzLabel22: TRzLabel;
    RzLabel23: TRzLabel;
    RzLabel24: TRzLabel;
    RzLabel25: TRzLabel;
    RzLabel26: TRzLabel;
    RzLabel27: TRzLabel;
    RzLabel28: TRzLabel;
    RzLabel29: TRzLabel;
    RzLabel39: TRzLabel;
    edRegTelNo1: TRzButtonEdit;
    edRegTelNo2: TRzButtonEdit;
    edRegTelNo3: TRzButtonEdit;
    edRegTelNo4: TRzButtonEdit;
    edRegTelNo5: TRzButtonEdit;
    edRegTelNo6: TRzButtonEdit;
    edRegTelNo7: TRzButtonEdit;
    edRegTelNo8: TRzButtonEdit;
    edRegTelNo9: TRzButtonEdit;
    edRegTelNo0: TRzButtonEdit;
    btnCheckTelNo: TRzBitBtn;
    btnSetTelNo: TRzBitBtn;
    GroupBox7: TGroupBox;
    LMDPanelFill4: TLMDPanelFill;
    GroupBox8: TGroupBox;
    Label57: TLabel;
    Label58: TLabel;
    Label59: TLabel;
    Label62: TLabel;
    Label4: TLabel;
    ComboBox1: TComboBox;
    ComboBox2: TComboBox;
    GroupBox9: TGroupBox;
    Label63: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    RzSpinner2: TRzSpinner;
    Btn_RegCalltime: TRzBitBtn;
    Btn_CheckCalltime: TRzBitBtn;
    Btn_CheckDialInfo: TRzBitBtn;
    Btn_RegDialInfo: TRzBitBtn;
    GroupBox10: TGroupBox;
    LMDPanelFill5: TLMDPanelFill;
    GroupBox11: TGroupBox;
    Memo_CardNo: TRzMemo;
    cbAutoReg: TRzCheckBox;
    RzBitBtn5: TRzBitBtn;
    Btn_FileOpen: TRzBitBtn;
    GroupBox12: TGroupBox;
    LMDListBox3: TLMDListBox;
    Label5: TLabel;
    btnCheckCardNo: TRzBitBtn;
    btnRegCardNo: TRzBitBtn;
    btnDeleteCardNo: TRzBitBtn;
    GroupBox13: TGroupBox;
    LMDPanelFill6: TLMDPanelFill;
    Label13: TLabel;
    Label6: TLabel;
    Label14: TLabel;
    ComboBox_CardModeType: TRzComboBox;
    Label3: TLabel;
    ComboBox_LockType: TRzComboBox;
    btnCheckDoorInfo: TRzBitBtn;
    btnRegDoorInfo: TRzBitBtn;
    GroupBox14: TGroupBox;
    LMDPanelFill8: TLMDPanelFill;
    Btn_SyncTime: TRzBitBtn;
    btnCheckVer: TRzBitBtn;
    btnResetDevice: TRzBitBtn;
    btnDoorOPen: TRzBitBtn;
    btnDoorClose: TRzBitBtn;
    Edit_Reset: TRzEdit;
    Edit_Ver: TRzEdit;
    Edit_TimeSync: TRzEdit;
    GroupBox15: TGroupBox;
    LMDListBox4: TLMDListBox;
    Btn_DelEventLog: TRzBitBtn;
    GroupBox16: TGroupBox;
    Label19: TLabel;
    ed_DB_Deviceid: TEdit;
    Label23: TLabel;
    cb_DB_Center: TComboBox;
    Label20: TLabel;
    ed_DB_Local: TEdit;
    Label21: TLabel;
    ed_DB_sLocal: TEdit;
    Label22: TLabel;
    ed_DB_siteName: TEdit;
    Label7: TLabel;
    ed_DB_pstn: TEdit;
    Label24: TLabel;
    ed_DB_ip: TEdit;
    Label25: TLabel;
    ed_DB_subnet: TEdit;
    Label26: TLabel;
    ed_DB_gateway: TEdit;
    Label30: TLabel;
    ed_DB_memo: TEdit;
    Label27: TLabel;
    ed_DB_InstallCo: TEdit;
    Label28: TLabel;
    ed_DB_installName: TEdit;
    Label29: TLabel;
    ed_DB_installTell: TEdit;
    Label31: TLabel;
    RzBitBtn4: TRzBitBtn;
    LMDPanelFill7: TLMDPanelFill;
    TB_DEVICE: TDBISAMTable;
    Label32: TLabel;
    Label33: TLabel;
    SpinEdit_OpenMoni1: TRzSpinEdit;
    SpinEdit_OpenDoor1: TRzSpinEdit;
    Label34: TLabel;
    ComboBox_Fire: TRzComboBox;
    procedure AssistCompleted(Sender: TObject; var Cancel: Boolean);
    procedure brnConnectClick(Sender: TObject);
    procedure ComPortTriggerAvail(CP: TObject; Count: Word);
    procedure brndisConnectClick(Sender: TObject);
    procedure btnIDCheckClick(Sender: TObject);
    procedure btnIDSetClick(Sender: TObject);
    procedure btnCheckwiznetClick(Sender: TObject);
    procedure Off_TimerTimer(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnSetwiznetClick(Sender: TObject);
    procedure edRegTelNo0ButtonClick(Sender: TObject);
    procedure btnCheckTelNoClick(Sender: TObject);
    procedure btnSetTelNoClick(Sender: TObject);
    procedure LMDAssistPage1NextPage(Sender: TObject; var Cancel: Boolean);
    procedure Btn_CheckDialInfoClick(Sender: TObject);
    procedure Btn_RegDialInfoClick(Sender: TObject);
    procedure Btn_CheckCalltimeClick(Sender: TObject);
    procedure Btn_RegCalltimeClick(Sender: TObject);
    procedure btnRegCardNoClick(Sender: TObject);
    procedure btnDeleteCardNoClick(Sender: TObject);
    procedure LMDAssistPage6Exit(Sender: TObject);
    procedure Btn_SyncTimeClick(Sender: TObject);
    procedure btnCheckVerClick(Sender: TObject);
    procedure btnResetDeviceClick(Sender: TObject);
    procedure btnDoorOPenClick(Sender: TObject);
    procedure btnDoorCloseClick(Sender: TObject);
    procedure btnCheckDoorInfoClick(Sender: TObject);
    procedure Btn_DelEventLogClick(Sender: TObject);
    procedure RzBitBtn5Click(Sender: TObject);
    procedure btnCheckCardNoClick(Sender: TObject);
    procedure btnRegDoorInfoClick(Sender: TObject);
    procedure AssistChanging(Sender: TObject; NewPage: TLMDAssistPage;
      var Cancel: Boolean);
    procedure PollingTimerTimer(Sender: TObject);
    procedure edRegTelNo0Change(Sender: TObject);
    procedure edRegTelNo1Click(Sender: TObject);
    procedure Edit_DeviceID1Click(Sender: TObject);
    procedure Edit_ServerIpClick(Sender: TObject);
    procedure ComboBox1Click(Sender: TObject);
    procedure RzSpinner1Click(Sender: TObject);
    procedure Btn_FileOpenClick(Sender: TObject);
    procedure btnCheckLinkusIDClick(Sender: TObject);
    procedure btnRegLinkusIDClick(Sender: TObject);
    procedure Edit_NameEnter(Sender: TObject);
    procedure ComboBox_DoorType1Change(Sender: TObject);
    procedure ComPortTriggerTimer(CP: TObject; TriggerHandle: Word);
    procedure Edit_LinkusIDEnter(Sender: TObject);
    procedure Spinner_RingChanging(Sender: TObject; NewValue: Integer;
      var AllowChange: Boolean);
    procedure Edit_LinkusIDClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure RzBitBtn3Click(Sender: TObject);
    procedure ed_DB_DeviceidKeyPress(Sender: TObject; var Key: Char);
    procedure cb_NextClick(Sender: TObject);
    procedure btn_NextClick(Sender: TObject);
    procedure RzBitBtn1Click(Sender: TObject);
    procedure btn_PrvClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private

    TimeoutTimer : Word;
    RetryCount:    Integer;
    LastSendData: String;
    FConnected: Boolean;
    aSysInfo2 : string;
    L_bFirmWareUpdate : Boolean;

    function  CheckDataPacket(aData: String; var bData:String):String;
    function  DataPacektProcess( aData: string):Boolean;
    Procedure RegDataProcess(aData: String);
    procedure RemoteDataProcess(aData: String);
    Procedure AccessDataProcess(aData: String);
    Procedure PollingAck(aDeviceID: String);
    function  SendPacket(aDeviceID: String;aCmd:Char; aData: String):Boolean;
    //����ȣ ��ȸ
    function CheckID: Boolean;
    procedure RcvDeviceID(aData: String);
    function RegID(aDeviceID: String):Boolean;

    //�ý�������  ��ȸ
    function CheckSysInfo(aDeviceID: String):Boolean;
    Procedure RcvSysinfo(aData: String);
    function RegSysInfo(aDeviceID: String):Boolean;

    Procedure ClearWiznetInfo;
    //wiznet ��ȸ
    function CheckWiznet(aDeviceID: string): Boolean;
    Procedure RcvWiznetInfo(aData:String);
    Procedure RegTellNo(aDeviceID:String;aNo:Integer;aTellno:String);
    Procedure CheckTellNo(aDeviceID:String;aNo:Integer);
    procedure ClearEditTelNo;

    Procedure RcvTellNo(aData:String);
    Procedure RegDialTime(aDeviceID:String;OnTime: Integer;OffTime:Integer);
    Procedure CheckDialTime(aDeviceID:String);
    Procedure RegCallTime(aDeviceID:String;aTime: Integer);
    Procedure CheckCallTime(aDeviceID:String);
    Procedure RcvDialInfo(aData:String);
    Procedure RcvCallTime(aData:String);
    Procedure CD_DownLoad(aDevice: String;aCardNo:String;func:Char);
    Procedure CardAllDownLoad(aFunc:Char);
    Procedure RcvCardRegAck(aData:String);
    Procedure RcvAccEventData(aData:String);
    Procedure RcvDoorEventData(aData:String);
    procedure ACC_sendData(aDeviceID:String; aData:String);
    Procedure Cnt_TimeSync(aDeviceID: String; aTimeStr:String);
    Procedure Cnt_CheckVer(aDeviceID: String);
    Procedure Cnt_Reset(aDeviceID:String);
    procedure RegSysInfo2(aDeviceID: String);
    procedure CheckSysInfo2(aDeviceID: String);
    Procedure RcvSysinfo2(aData: String);
    Procedure DoorControl(DeviceID:String; aNo:Integer ; aControlType:Char; aControl:Char);
    Procedure ClearInfo(aIndex:Integer; isEmpty:Boolean);

    function RegLinkusID(aDeviceID, aLinkusId: String):Boolean;
    function CheckLinkusID(aDeviceID: String):boolean;
    Procedure RcvLinkusId(aData: String);

    function RegLinksTellNo(aDeviceID:String;aNo:Integer;aTellno:String):boolean;
    function CheckLinkusTellNo(aDeviceID:String;aNo:Integer):boolean;
    Procedure RcvLinkusTelNo(aData:String);

    function RegRingCount(aDeviceID:String;aCount: Integer):Boolean;
    function CheckRingCount(aDeviceID:String):Boolean;
    Procedure RcvRingCount(aData:String);

    procedure RegRemoteCheckTime(aDeviceID:String; aTime: Integer);
    procedure autoSetRemoteCheckTime;
    procedure OpenDoor;

    Function IsRegDB(aDeviceID:String):Boolean;
    Procedure InsertDb;
    Procedure UpdateDb;
    //�˶���� ����
    function ChangeAlarmMode(aDeviceID: string; aMode: Char;aQuick:Boolean): Boolean;
    procedure SetConnected(const Value: Boolean);
  private
    ComPortList : TStringList;
    function GetSerialPortList(List : TStringList; const doOpenTest : Boolean = True) : LongWord;
    function EncodeCommportName(PortNum : WORD) : String;
    function DecodeCommportName(PortName : String) : WORD;
  published
    Property bConnected : Boolean read FConnected write SetConnected;

  public
    Rcv_MsgNo     : Char;
    Send_MsgNo    : Integer;
    WiznetData    : String;
    WizNetRegMode : Boolean;
    WizNetRcvACK  : Boolean;
    xDeviceID     : String[9];
  end;

var
  Form_Main: TForm_Main;
  RcvPollingTime: TDatetime;
  ComBuff: String;
  isRegMode: Boolean;
  oldtime1,oldtime2,oldTime3: String;
  Newtime1,Newtime2,NewTime3: String;

implementation

uses
uCommon,uLomosUtil;
{$R *.dfm}

procedure TForm_Main.AssistCompleted(Sender: TObject;
  var Cancel: Boolean);
begin

end;

{����}
procedure TForm_Main.brnConnectClick(Sender: TObject);
var
  nLoop : integer;
  bResult : Boolean;
begin
  if cbSerialPort.ItemIndex > -1 then
    nSerialPort :=  Integer(ComPortList.Objects[cbSerialPort.ItemIndex]);

  try
    LMDIniCtrl1.WriteInteger('����','ComNo',nSerialPort);
    with ComPort do
    begin
      DeviceLayer:= dlWin32;
      Baud:= 38400;
//      Baud:= 9600; //TEST
      ComNumber := nSerialPort;
      ApdSLController.Monitoring := true;
      OPen:= False;
      Sleep(100);
      bConnected := True;
      brnConnect.Enabled := false;
      brndisConnect.Enabled := true;
      Open := true;
      cbSerialPort.Enabled:= False;
      Panel1.Color:= clRed;
    end;
{    ComPort.ComNumber:= cbSerialPort.ItemIndex + 1;

    if ComPort.Open then ComPort.Open:= False;
    ComPort.Open:= True;

    TimeoutTimer := ComPort.AddTimerTrigger;

    ComPort.FlushInBuffer;
    ComPort.FlushOutBuffer;

    cbSerialPort.Enabled:= False;
    ApdSLController.Monitoring:= True;
    Panel1.Color:= clRed;
    PollingTimer.Enabled:= True;
    RcvPollingTime:= Now; }
  except
    brndisConnectClick(Sender);
    ShowMessage('�����Ʈ �� Ȯ���ϼ���');
  end;

  for nLoop := 0 to LoopCount - 1 do
  begin
    bResult := CheckID;
    if Not bConnected then Exit;
    if bResult then break;
  end;

  //ID�� �켱 Ȯ���Ѵ�.
  if Not bResult then
  begin
    brndisConnectClick(Sender);
    ShowMessage('�������');
    exit;
  end;
  //��⼳���� ���ؼ� �ݵ�� ���� ���
  ChangeAlarmMode('000000000','D',True);


end;
{����}
procedure TForm_Main.brndisConnectClick(Sender: TObject);
begin
  bConnected := False;
  PollingTimer.Enabled:= False;
  Panel1.Color:= clSilver;
  cbSerialPort.Enabled:= True;
  brnConnect.Enabled := True;
  brndisConnect.Enabled := False;
  brndisConnect.Down:= True;
  ComPort.Open:= False;
  ApdSLController.Monitoring:= False;
end;


procedure TForm_Main.ComPortTriggerAvail(CP: TObject; Count: Word);
var
  st: string;
  st2: String;
  st3: String;
  aIndex: Integer;
  I: Integer;
  Lenstr:String;
  DataLength: Integer;
begin
  st:= '';
  for I := 1 to Count do st := st + ComPort.GetChar;
  ComBuff:= ComBuff + st;
  Panel1.Color:= clLime;
  //PollingTimer.Enabled:= True;
  //RzFieldStatus1.Caption:=  ComBuff;

  aIndex:= Pos(STX,ComBuff);  // STX ��ġ�� Ȯ�� �Ѵ�.
  if aIndex = 0 then Exit     // STX�� ������ ��ƾ���� ������.
  else if aIndex > 1 then
  begin
    Delete(ComBuff,1,aIndex-1);  //STX��ġ�� 1�ƴϸ� STX�ձ��� ����
  end;
  if Length(Combuff) < 21 then Exit;
  repeat
    st3:= CheckDataPacket(ComBuff,st2);
    ComBuff:= st2;
    if st <> '' then DataPacektProcess(st3);
  until pos(ETX,comBuff) = 0;

end;

function TForm_Main.CheckDataPacket(aData: String;
  var bData: String): String;
var
  aIndex: Integer;
  Lenstr: String;
  DefinedDataLength: Integer;
  StrBuff: String;
  etxIndex: Integer;
begin

  Result:= '';
  Lenstr:= Copy(aData,2,3);
  //������ ���� ��ġ �����Ͱ� ���ڰ� �ƴϸ�...
  if not isDigit(Lenstr) then
  begin
    Delete(aData,1,1);       //1'st STX ����
    aIndex:= Pos(STX,aData); // ���� STX ã��
    if aIndex = 0 then       //STX�� ������...
    begin
      //��ü ������ ����
      bData:= '';
    end else if aIndex > 1 then // STX�� 1'st�� �ƴϸ�
    begin
      Delete(aData,1,aIndex-1);//STX �� ������ ����
      bData:= aData;
    end else
    begin
      bData:= aData;
    end;
    Exit;
  end;

  //��Ŷ�� ���ǵ� ����
  DefinedDataLength:= StrtoInt(Lenstr);
  //��Ŷ�� ���ǵ� ���̺��� ���� �����Ͱ� ������
  if Length(aData) < DefinedDataLength then
  begin
    //���� �����Ͱ� ���̰� ������(���� �� ������ ����)
    etxIndex:= POS(ETX,aData);
    if etxIndex > 0 then
    begin
     Delete(aData,1,etxIndex);
    end;
    bData:= aData;
    Exit;
  end;

  // ���ǵ� ���� ������ �����Ͱ� ETX�� �´°�?
  if aData[DefinedDataLength] = ETX then
  begin
    StrBuff:= Copy(aData,1,DefinedDataLength);
    Result:=StrBuff;
    Delete(aData, 1, DefinedDataLength);
    bData:= aData;
  end else
  begin
    //������ �����Ͱ� EXT�� �ƴϸ� 1'st STX����� ���� STX�� ã�´�.
    Delete(aData,1,1);
    aIndex:= Pos(STX,aData); // ���� STX ã��
    if aIndex = 0 then       //STX�� ������...
    begin
      //��ü ������ ����
      bData:= '';
    end else if aIndex > 1 then // STX�� 1'st�� �ƴϸ�
    begin
      Delete(aData,1,aIndex-1);//STX �� ������ ����
      bData:= aData;
    end else
    begin
      bData:= aData;
    end;
  end;
end;


function TForm_Main.DataPacektProcess(aData: string): Boolean;
var
  aKey: Byte;
  st: string;
  aCommand: Char;
  aCntId: String;
begin
  Result:= False;
  if aData = '' then Exit;

  //31:Q++()./,-**s*S^**+()./,-()
  aKey:= Ord(aData[5]);
  st:= Copy(aData,1,5) + EncodeData(aKey,Copy(aData,6,Length(aData)-6))+aData[Length(aData)];
  aData:= st;
  aCntId:= Copy(aData,8,9);
  aCommand:= aData[17];
  Rcv_MsgNo:= aData[18];

  RzFieldStatus1.Caption:=  'RX:'+aData;
  {A,I,R Commmand�� ������ �ش�.}
  if  (aCommand = 'e') then
  begin
    RcvPollingTime:= Now;
    PollingAck(aCntID);
  end else if (aCommand <> 'c') then
{  if  (aCommand = 'A') or
      (aCommand = 'i') or
      (aCommand = 'R') or
      (aCommand = 'E') or
      (aCommand = 'r')then  // or (aCommand = 'c')  }
  begin
    SendPacket(aCntID,'a','');
  end else if (aCommand = 'a') then
  begin
    Exit;
  end;

  {���� ������ Ŀ�ǵ庰 ó��}
  { ================================================================================
  "A" = Alarm Data
  "I" = Initial Data
  "R" = Remote Command
  "e" = ENQ
  "E" = ERROR
  "a" = ACK
  "n" = NAK
  "r" = Remote Answer
  "c" = Access Control data
  �� c(�������� ������)�ΰ�쿡�� ACK �� 'c' command�� ����� ������ �ؾ� �Ѵ�.
  �� ACK ������ �ι� �־�� �Ѵ�.(����ü ��Ŷ����,���������� ����)
   ================================================================================ }

  //codesite.SendMsg(st);
  case aCommand of
    'A':{�˶�}          begin  end;
    'i':{Initial}
        begin
          RetryCount:=0;
          //ComPort.SetTimerTrigger( TimeoutTimer, TimeoutTicks, False); { disable }
          RegDataProcess(aData);
        end;
    'R':{Remote}
        begin
          RetryCount:=0;
          //ComPort.SetTimerTrigger( TimeoutTimer, TimeoutTicks, False); { disable }
          RemoteDataProcess(aData);
        end;
    'r':{Remote Answer}
        begin
          RetryCount:=0;
          //ComPort.SetTimerTrigger( TimeoutTimer, TimeoutTicks, False); { disable }
          RemoteDataProcess(aData);
        end;
    'c':{��������}
        begin
          RetryCount:=0;
          //ComPort.SetTimerTrigger( TimeoutTimer, TimeoutTicks, False); { disable }
          AccessDataProcess(aData);
        end;
    //'f':{�߿���}        begin  FirmwareProcess(aData)   end;
    'e':{ERROR}
    else {error �߻�: [E003]���� ���� ���� Ŀ�ǵ�}
  end;
  Result:= True;
end;

//��� ����
procedure TForm_Main.RegDataProcess(aData: String);
var
  aDeviceCode: String;
  I: Integer;
  aRegCode: cString;
begin
  aDeviceCode := Copy(aData, 8,7);

  aRegCode:= Copy(aData,19,2);
  //40 K1123456700i1IF00���нý���      61
  if aRegCode = 'ID' then                  //ID ����
  begin
    RcvDeviceID(aData);
  end else if aRegCode = 'SY' then            //�ý��� ����
  begin
    RcvSysinfo(aData);
  end else if aRegCode = 'NW' then        //��Ʈ��ũ���� ����
  begin
    ClearWiznetInfo;
    RcvWiznetInfo(aData);
  end  else if aRegCode = 'TN' then       //��ȭ ��ȣ ��� ����
  begin
    RcvTellNo(aData);
  end else if aRegCode = 'DI' then       // DTMF ��� ����
  begin
    RcvDialInfo(aData);
  end else if aRegCode = 'CT' then
  begin
    RcvCallTime(aData);
  end else if aRegCode = 'Id' then
  begin
    RcvLinkusId(aData);
  end else if aRegCode = 'Tn' then
  begin
    RcvLinkusTelNo(aData);
  end else if aRegCode = 'Rc' then
  begin
    RcvRingCount(aData);
  end;
end;

//�������� ������
procedure TForm_Main.AccessDataProcess(aData: String);
var
  DeviceID: String;
  st: string;
  msgCode: Char;
  accData: String;
begin

  // STX ~ 21's Byte :Header
  //040 K1123456700i1IFN00���нý���      61


  st:= Copy(aData,8,9);
  xDeviceID:= st;
  accData:= Copy(aData,19,Length(aData)-20); //�������� ���� ���� ������
  msgCode:= accData[1];

  {ACK ����:���԰� DOOR}
  if (msgCode = 'E') or (msgCode = 'D') or (msgCode = 'F') then
  begin
    st:='Y' + Copy(aData,20,2)+'  '+'a';
    ACC_sendData(xDeviceID, st);
  end;

  {�������� ������ ó��}
//0460K1100000400c2a51  005000000010000000009E
  case msgcode of
    //'F': RcvTelEventData(accData);
    'E': RcvAccEventData(accData);
    'D': RcvDoorEventData(accData);
    'a': RcvSysinfo2(accData); //��� ��� ����
    'b': RcvSysinfo2(accData);//��� ��ȸ ����
    //'c': RcvAccControl(accData);// �� ���� ����
'l','n','m': RcvCardRegAck(accData);
  end;
end;


//
procedure TForm_Main.RemoteDataProcess(aData: String);
var
  aCode: String;
  st: string;
  aIndex: Integer;
begin
  //037 K1123456700r1TM00050107180637EF
  aCode:= Copy(aData,19,2);
  if aCode = 'TM' then          //�ð�����
  begin
    Edit_TimeSync.Text:= Copy(aData,23,4)+'-'+  //��
                         Copy(aData,27,2)+'-'+  //��
                         Copy(aData,29,2)+' '+  //��
                         Copy(aData,31,2)+':'+  //��
                         Copy(aData,33,2)+':'+  //��
                         Copy(aData,35,2);      //��
  end else if aCode = 'VR' then //����Ȯ��
  begin
    Edit_Ver.Text:= Copy(aData,22,Length(aData)-23);

  end else if aCode = 'RS' then //Reset
  begin
    Edit_Reset.Text:= Copy(aData,22,Length(aData)-23);
  end else if aCode = 'Pt' then  //����üũ ���۽ð�
  begin
    st:= Copy(aData,23,2);
    Label18.Caption:= st + '�ð��� ����üũ �õ�';
  end;

end;


procedure TForm_Main.PollingAck(aDeviceID: String);
var
  ACKStr: String;
  ACKStr2: String;
  aDataLength: Integer;
  aLengthStr: String;
  aKey:Integer;
  aMsgNo: Integer;
  st: string;
begin
  aDataLength:= 21;
  aLengthStr:= FillZeroNumber(aDataLength,3);
  ACkStr:= STX +aLengthStr+  #$20+'K1'+ aDeviceID+ 'a'+Rcv_MsgNo;
  ACkStr:= ACkStr+ MakeCSData(ACKStr+ETX)+ETX;
  aKey:= $20;
  ACkStr2:= Copy(ACKStr,1,5)+EncodeData(aKey,Copy(ACkStr,6,Length(ACkStr)-6))+ETX;
  ComPort.PutString(ACKStr2);

  //st:=  'TX:' + ACkStr2;
  //RzFieldStatus1.Caption:= st;

end;
function TForm_Main.SendPacket(aDeviceID: String; aCmd: Char;
  aData: String): Boolean;
var
  ErrCode: Integer;
  ACKStr: String;
  ACKStr2: String;
  aDataLength: Integer;
  aLengthStr: String;
  aKey:Integer;
  aMsgNo: Integer;
  st: string;
begin

  if not ComPort.Open then
  begin
    bConnected := False;
    ShowMessage('��� ������ �ȵǾ����ϴ�.');
    Exit;
  end;

  ErrCode:= 0;
  Result:= False;
  aDataLength:= 21 + Length(aData);
  aLengthStr:= FillZeroNumber(aDataLength,3);

  if aCmd = 'a' then {���� ó��}
  begin
    ACkStr:= STX +aLengthStr+  #$20+'K1'+ aDeviceID+ aCmd+Rcv_MsgNo;
    ACkStr:= ACkStr+ MakeCSData(ACKStr+ETX)+ETX;
    aKey:= $20;
    ACkStr2:= Copy(ACKStr,1,5)+EncodeData(aKey,Copy(ACkStr,6,Length(ACkStr)-6))+ETX;
  end else           {���� or ��� }
  begin
    aMsgNo:= Send_MsgNo;
    ACkStr:= STX +aLengthStr+ #$20+'K1'+ aDeviceID+ aCmd+InttoStr(aMsgNo) +aData;
    ACkStr:= ACkStr+ MakeCSData(ACKStr+ETX)+ETX;
    aKey:= Ord(ACkStr[5]);
    ACkStr2:= Copy(ACKStr,1,5)+EncodeData(aKey,Copy(ACkStr,6,Length(ACkStr)-6))+ETX;
    if aMsgNo >= 9 then  Send_MsgNo:= 0
    else                 Send_MsgNo:= aMsgNo + 1;

//    ComPort.SetTimerTrigger( TimeoutTimer, TimeoutTicks, True); { TimeOutTimer enable }
    LastSendData:= ACKStr2;

  end;

  ComPort.PutString(ACKStr2);
  st:= 'TX:'+ ACKStr2;
  RzFieldStatus1.Caption:= st;

  st:=  'Server:'+INttoStr(Errcode) +';'+
        Copy(aDeviceID,1,7)+';'+
        Copy(aDeviceID,8,2)+';'+
        //Copy(ACKStr2,17,2)+';'+
        ACKStr2[17]+';'+
        Copy(ACKStr2,19,Length(ACKStr2)-21)+';'+
        ACkStr2+';'+
        'TX';

  Result:= True;
end;


//id ��ȸ
Function TForm_Main.CheckID:Boolean;
var
  stData : string;
  nTime : integer;
  PastTime : dword;
begin
  Result := false;
  stData := 'ID000000000';

  bMCUIDCheck := False;
  SendPacket('000000000', 'Q', stData);

  nTime := 0;
  PastTime := GetTickCount + DelayTime;
  while Not bMCUIDCheck do
  begin
    if Not bConnected then Exit;
    Application.ProcessMessages;
    if GetTickCount > PastTime then Exit;  //300�и����� ���� ������ ���з� ó����
  end;
  Result := true;
end;


//ID ����
procedure TForm_Main.RcvDeviceID(aData: String);
var
  st: string;
begin
  //40 K1123456700i1IF00���нý���      61
  Delete(aData,1,1); //�����ͱ��� 1Byte�� ���߿� �߰��Ǿ� ���Ƿ� 1Byte ���� ó��
  st:=  Copy(aData,22,7);
  xDeviceID:= st+'00';
  Edit_DeviceID1.Color:= clYellow;
  Edit_DeviceID2.Color:= clYellow;
  Edit_DeviceID3.Color:= clYellow;
  Edit_DeviceID4.Color:= clYellow;
  Edit_DeviceID5.Color:= clYellow;
  Edit_DeviceID6.Color:= clYellow;
  Edit_DeviceID7.Color:= clYellow;

  Edit_DeviceID1.Text:= st[1];
  Edit_DeviceID2.Text:= st[2];
  Edit_DeviceID3.Text:= st[3];
  Edit_DeviceID4.Text:= st[4];
  Edit_DeviceID5.Text:= st[5];
  Edit_DeviceID6.Text:= st[6];
  Edit_DeviceID7.Text:= st[7];
  bMCUIDCheck := True;  //MCUIDüũ ����
{  if isRegMode then RegSysInfo(xDeviceID)
  else CheckSysInfo(xDeviceID); }

  //Edit_Name.Text:= '';
  //SHowMessage('ID��ȸ/����� �Ϸ� �Ǿ����ϴ�.'+#13+st)
end;

//ID���
function TForm_Main.RegID(aDeviceID: String):Boolean;
var
  PastTime : dword;
begin
  Result := false;
  if (not isdigit(aDeviceID)) or (Length(aDeviceID) <> 7) then Exit;

  bMCUIDCheck := False;
  SendPacket('000000000', 'I', 'ID00' + aDeviceID);

  PastTime := GetTickCount + DelayTime;
  while Not bMCUIDCheck do
  begin
    if Not bConnected then Exit;
    Application.ProcessMessages;
    if GetTickCount > PastTime then Exit;  //300�и����� ���� ������ ���з� ó����
  end;
  Result := true;
end;

//�ý��� ���� ��û
function TForm_Main.CheckSysInfo(aDeviceID: String):Boolean;
var
  aData: string;
  nTime: integer;
  PastTime : dword;

begin
  Result := false;
  if (not isdigit(aDeviceID)) or (Length(aDeviceID) <> 9) then Exit;
  aData := 'SY00';
  bSysInfoCheck := False;
  SendPacket(aDeviceID, 'Q', aData);

  nTime := 0;
  PastTime := GetTickCount + DelayTime;
  while Not bSysInfoCheck do
  begin
    if Not bConnected then Exit;
    Application.ProcessMessages;
    if GetTickCount > PastTime then Exit;  //300�и����� ���� ������ ���з� ó����
  end;

  Result := true;

end;

procedure TForm_Main.RcvSysinfo(aData: String);
var
  st: string;
  aLocate: String;
begin

  Delete(aData,1,1); //�����ͱ��� 1Byte�� ���߿� �߰��Ǿ� ���Ƿ� 1Byte ���� ó��
  st:= Copy(aData,22,7);
  aLocate:= Copy(aData,29,16);
  if not isDigit(st) then Exit;

  ComboBox_DoorType1.Color:= clYellow;
  ComboBox_DoorType1.ItemIndex:= StrtoInt(st[6]);
  ComboBox_DoorType1.Text:= ComboBox_DoorType1.Items[ComboBox_DoorType1.ItemIndex];
  Edit_Name.Color:= clYellow;
  Edit_Name.Text:= TrimRight(aLocate);

  if ComboBox_DoorType1.ItemIndex = 1 then Group_Linkus.Visible:= False
  else                                     Group_Linkus.Visible:= True;

  bSysInfoCheck :=True;

  if not isRegMode then OpenDoor;

  //SHowMessage('��� ��ȸ/����� �Ϸ� �Ǿ����ϴ�.');

end;

function TForm_Main.RegSysInfo(aDeviceID: String):Boolean;
var
  aData: string;
  PastTime : dword;

begin
  Result := false;
  if (not isdigit(aDeviceID)) or (Length(aDeviceID) <> 9) then Exit;

  aData:= 'SY00' +                                   // COMMAND
          '1' +                                     // �������ÿ���
          '000'+                                    // ��������ð�
          '2'+                                      // ����
          InttoStr(ComboBox_DoorType1.ItemIndex)+
          InttoStr(ComboBox_DoorType1.ItemIndex)+
          SetlengthStr(Edit_Name.Text,16)+
          '000'+                                    // �Խ������ð�
          '1';                                      // ���� ��� ���


  bSysInfoCheck := False;
  SendPacket(aDeviceID, 'I', aData);

  PastTime := GetTickCount + DelayTime;
  while Not bSysInfoCheck do
  begin
    if Not bConnected then Exit;
    Application.ProcessMessages;
    if GetTickCount > PastTime then Exit;  //300�и����� ���� ������ ���з� ó����
  end;

  Result := true;

end;


//ID��ȸ ��ư
procedure TForm_Main.btnIDCheckClick(Sender: TObject);
var
  nLoop : integer;
  bResult : Boolean;
begin
  ClearInfo(1,True);
  for nLoop := 0 to LoopCount - 1 do
  begin
    bResult := CheckID;
    if Not bConnected then Exit;
    if bResult then break;
  end;

  //ID�� �켱 Ȯ���Ѵ�.
  if Not bResult then
  begin
    brndisConnectClick(Sender);
    showmessage('ID Check ��������� �����ϴ�.���۸�带 Ȯ���Ͽ� �ּ���.');
    exit;
  end;

  ComboBox_DoorType1.ItemIndex:= -1;
  for nLoop := 0 to LoopCount - 1 do
  begin
    bResult := CheckSysInfo(xDeviceID);
    if Not bConnected then Exit;
    if bResult then break;
  end;

  //������������ Ȯ���Ѵ�.
  if Not bResult then
  begin
    brndisConnectClick(Sender);
    showmessage('ID Check ��������� �����ϴ�.���۸�带 Ȯ���Ͽ� �ּ���.');
    exit;
  end;

  btnCheckLinkusIDClick(btnCheckLinkusID);

end;

//ID��� ��ư
procedure TForm_Main.btnIDSetClick(Sender: TObject);
var
  aID: String;
  nLoop : integer;
  bResult : Boolean;
begin
  aID:= Edit_DeviceID1.Text+
        Edit_DeviceID2.Text+
        Edit_DeviceID3.Text+
        Edit_DeviceID4.Text+
        Edit_DeviceID5.Text+
        Edit_DeviceID6.Text+
        Edit_DeviceID7.Text;

  ClearInfo(1,False);
  if (Length(aID) < 7) or (isDigit(aID) = False) or (Edit_Name.Text = '') then
  begin
    ShowMessage('���ID�� �ùٸ��� �ʰų� ���α������ �����ϴ�.');
    Exit;
  end;

  for nLoop := 0 to LoopCount - 1 do
  begin
    bResult := RegID(aID);
    if Not bConnected then Exit;
    if bResult then break;
  end;
  if not bResult then
  begin
    showmessage('ID ��� ��������� �����ϴ�.���۸�带 Ȯ���Ͽ� �ּ���.');
    Exit;
  end;

  //���������� ���
  for nLoop := 0 to LoopCount - 1 do
  begin
    bResult :=  RegSysInfo(xDeviceID); // ��ġ��
    if Not bConnected then Exit;
    if bResult then break;
  end;
  if not bResult then
  begin
    showmessage('���������� ��������� �����ϴ�.���۸�带 Ȯ���Ͽ� �ּ���.');
    Exit;
  end;

end;

//����� ��ȸ
procedure TForm_Main.btnCheckwiznetClick(Sender: TObject);
var
  nLoop : integer;
  bResult : boolean;
begin
  WizNetRegMode := False;
  ClearInfo(2,True);
//  for nLoop := 0 to LoopCount - 1 do
//  begin
//    bResult := CheckWiznet(xDeviceID);
//    if Not bConnected then Exit;
//    if bResult then break;
//  end;
//  if not bResult then
//  begin
//    showmessage('����� ��ȸ ��������� �����ϴ�.���۸�带 Ȯ���Ͽ� �ּ���.');
//    Exit;
//  end;

  WizNetRegMode:= False;
  WizNetRcvACK:= False;
  SendPacket(xDeviceID,'Q','NW00');
  //ClearWiznetInfo;
  Off_Timer.Enabled:= True;
  LMDSimpleLabel1.Caption:= '������ ��ȸ�� �Դϴ�. ��ø� ��ٷ� �ֽʽÿ�';
  LMDCaptionPanel1.Left:= 208;
  LMDCaptionPanel1.Top:= 176;

  LMDCaptionPanel1.visible := True;
  LMDCaptionPanel1.BringToFront;
  LMDSimpleLabel1.Twinkle:= True;
  ProgressBar3.Position:= 0;
end;

//����� ����
procedure TForm_Main.btnSetwiznetClick(Sender: TObject);
var
  I: Integer;
  No: Integer;
  st,st2: String;
  DataStr: String;
  aDeviceID: String;
  FHeader:         String[2];
  FMacAddress:    String[12];
  FMode:          String[2];
  FIPAddress:     String[8];
  FSubnet:        String[8];
  FGateway:       String[8];
  FClientPort:    String[4];
  FServerIP:      String[8];
  FServerPort:    String[4];
  FSerial_Baud:   String[2];
  FSerial_data:   String[2];
  FSerial_Parity: String[2];
  FSerial_stop:   String[2];
  FSerial_flow:   String[2];
  FDelimiterChar: String[2];
  FDelimiterSize: String[4];
  FDelimitertime: String[4];
  FDelimiterIdle: String[4];
  FDebugMode:     String[2];
  FROMVer:        String[4];
  FOnDHCP:        String[2];
  FReserve:       String[6];
begin
  WizNetRegMode:= True;
  WizNetRcvACK:= False;
  //1.Header
  FHeader:= 'AA';
  //2.MAC Address ��������
  FMacAddress:='000000000000';
  //3.Mode (Server mode: 02, Mixed Mode: 01 Client mode: 00)
  FMode:= '02';
  {
  if Checkbox_Client.Checked then FMode:= '00'
  else                            FMode:= '01';
   }

  // 4.IP address
  st2:= '';
  if Edit_LocalIP.Text = '' then Edit_LocalIP.Text:= '192.168.10.10';   //IP����
  //Edit_LocalIP.Text:= '192.168.10.10';
  for I:= 0 to 3 do
  begin
    st:= FindCharCopy(Edit_LocalIP.Text,I,'.');
    No:= StrToInt(st);
    st2:= st2 + Char(No);
  end;
  FIPAddress:= ToHexStrNoSpace(st2);

  // 5.Subnet mask
  st2:= '';
  //Edit_Subnet.Text := '255.255.255.0';
  if Edit_Subnet.Text = '' then Edit_Subnet.Text:= '255.255.255.0';
  for I:= 0 to 3 do
  begin
    st:= FindCharCopy(Edit_Subnet.Text,I,'.');
    No:= StrToInt(st);
    st2:= st2 + Char(No);
  end;
  FSubnet:= ToHexStrNoSpace(st2);

  // 6.Gateway address
  st2:= '';
  //Edit_Gateway.Text:= '192.168.10.1';
  if Edit_Gateway.Text = '' then Edit_Gateway.Text:= '192.168.10.1';
  for I:= 0 to 3 do
  begin
    st:= FindCharCopy(Edit_Gateway.Text,I,'.');
    No:= StrToInt(st);
    st2:= st2 + Char(No);
  end;
  FGateway:= ToHexStrNoSpace(st2);

  //7.Port number (Client)
  st2:= '';
  //Edit_LocalPort.Text := '3000';
  if Edit_LocalPort.Text = '' then Edit_LocalPort.Text:= '3000';
  st:= Dec2Hex(StrtoInt(Edit_LocalPort.Text),2);
  if Length(st) < 4 then st:= '0'+st;
  st2:= Chr(Hex2Dec(Copy(st,1,2)))+ Char(Hex2Dec(Copy(st,3,2)));
  FClientPort:= ToHexStrNoSpace(st2);

  //8. Server IP address
  st2:= '';
  if Edit_ServerIp.Text = '' then Edit_ServerIp.Text:= '0.0.0.0';
  for I:= 0 to 3 do
  begin
    st:= FindCharCopy(Edit_ServerIp.Text,I,'.');
    No:= StrToInt(st);
    st2:= st2 + Char(No);
  end;
  FServerIP:= ToHexStrNoSpace(st2);

  //9.  Port number (Server)
  st2:= '';
  if Edit_Serverport.Text = '' then Edit_Serverport.Text:= '3000';
  st2:='';
  st:= Dec2Hex(StrtoInt(Edit_Serverport.Text),2);
  if Length(st) < 4 then st:= '0'+st;
  st2:= Chr(Hex2Dec(Copy(st,1,2)))+ Char(Hex2Dec(Copy(st,3,2)));
  FServerPort:= ToHexStrNoSpace(st2);

  //10. Serial speed (bps)
  FSerial_Baud:= 'FD';
  {
  case ComboBox_Boad.ItemIndex of
    0: FSerial_Baud:= 'F4'; //9600           F4
    1: FSerial_Baud:= 'FA'; //19200          FA
    2: FSerial_Baud:= 'FD'; //38400 Default  FD
    3: FSerial_Baud:= 'FE'; //57600          FE
    4: FSerial_Baud:= 'FF'; //115200         FF
    else FSerial_Baud:= 'FD';
  end;
   }

  //11. Serial data size (08: 8 bit), (07: 7 bit)
  FSerial_data:= '08';
  {
  case ComboBox_Databit.ItemIndex of
      0: FSerial_data:= '07';
      1: FSerial_data:= '08'; //Default
      else FSerial_data:= '08';
  end;
   }

  //12. Parity (00: No), (01: Odd), (02: Even)
  FSerial_Parity:= '00';
  {
  case ComboBox_Parity.ItemIndex of
    0: FSerial_Parity:= '00'; //None Default
    1: FSerial_Parity:= '01'; //Odd
    2: FSerial_Parity:= '02'; //Even
    else FSerial_Parity:= '00';
  end;
   }
  //13. Stop bit
  FSerial_stop:= '01';

  //14.Flow control (00: None), (01: XON/XOFF), (02: CTS/RTS)
  FSerial_flow:= '00';
  {
  case ComboBox_Flow.ItemIndex  of
    0: FSerial_flow:= '00'; //Default
    1: FSerial_flow:= '01';
    2: FSerial_flow:= '02';
  end;
   }

  //15. Delimiter char
  FDelimiterChar:= '03';
  {
  if Edit_Char.Text ='' then Edit_Char.Text:= '00';
  FDelimiterChar:= Edit_Char.Text;
   }

  //16.Delimiter size
  {
  if Edit_Size.Text ='' then Edit_Size.Text:= '0';
  st:= Dec2Hex(StrtoInt(Edit_Size.Text),2);
  st:=FillZeroStrNum(st,4);
   }
  st:= '0000';
  st2:='';
  st2:= Chr(Hex2Dec(Copy(st,1,2)))+ Char(Hex2Dec(Copy(st,3,2)));
  FDelimiterSize:= ToHexStrNoSpace(st2);

  //17. Delimiter time

  //if Edit_Time.Text = '' then Edit_Time.Text:= '20';

  st:= Dec2Hex(StrtoInt('10'),2);
  st:=FillZeroStrNum(st,4);

  st2:='';
  st2:= Chr(Hex2Dec(Copy(st,1,2)))+ Char(Hex2Dec(Copy(st,3,2)));
  FDelimitertime:= ToHexStrNoSpace(st2);

  //18.Delimiter idle time
  //if Edit_Idle.Text = '' then Edit_Idle.Text:= '0';
  //st:= Dec2Hex(StrtoInt(Edit_Idle.Text),2);
  //st:=FillZeroStrNum(st,4);
  st:= '0000';
  st2:='';
  st2:= Chr(Hex2Dec(Copy(st,1,2)))+ Char(Hex2Dec(Copy(st,3,2)));
  FDelimiterIdle:= ToHexStrNoSpace(st2);

  //19. Debug code (00: ON), (01: OFF)
  FDebugMode:= '01';
  {
  if Checkbox_Debugmode.Checked then FDebugMode:= '00'
  else                               FDebugMode:= '01';
   }
  //20.Software major version
  FROMVer:='0000';

  // 21.DHCP option (00: DHCP OFF, 01:DHCP ON)

  if Checkbox_DHCP.Checked then FOnDHCP:= '01'
  else                          FOnDHCP:= '00';

  //22.Reserved for future use
  FReserve:= '000000';

  DataStr:= FHeader+
            FMacAddress+
            FMode+
            FIPAddress+
            FSubnet+
            FGateway+
            FClientPort+
            FServerIP+
            FServerPort+
            FSerial_Baud+
            FSerial_data+
            FSerial_Parity+
            FSerial_stop+
            FSerial_flow+
            FDelimiterChar+
            FDelimiterSize+
            FDelimitertime+
            FDelimiterIdle+
            FDebugMode+
            FROMVer+
            FOnDHCP+
            FReserve;

  WiznetData:= DataStr;

  {
  SHowMessage(FHeader+#13+
              FMacAddress+#13+
              FMode+#13+
              FIPAddress+#13+
              FSubnet+#13+
              FGateway+#13+
              FClientPort+#13+
              FServerIP+#13+
              FServerPort+#13+
              FSerial_Baud+#13+
              FSerial_data+#13+
              FSerial_Parity+#13+
              FSerial_stop+#13+
              FSerial_flow+#13+
              FDelimiterChar+#13+
              FDelimiterSize+#13+
              FDelimitertime+#13+
              FDelimiterIdle+#13+
              FDebugMode+#13+
              FROMVer+#13+
              FOnDHCP+#13+
              FReserve+#13+
            '����:'+InttoStr(Length(DataStr)));
  }

  SendPacket(xDeviceID,'I','NW00'+DataStr);
  ClearInfo(2,True);
  //ClearWiznetInfo;
  Off_Timer.Enabled:= True;
  LMDSimpleLabel1.Caption:= '������ ������ �Դϴ�. ��ø� ��ٷ� �ֽʽÿ�';
  LMDCaptionPanel1.Left:= 208;
  LMDCaptionPanel1.Top:= 176;
  LMDCaptionPanel1.visible := True;
  LMDCaptionPanel1.BringToFront;
  LMDSimpleLabel1.Twinkle:= True;
  ProgressBar3.Position:= 0;
end;


procedure TForm_Main.ClearWiznetInfo;
begin
  LMDListBox1.Clear;
end;

procedure TForm_Main.RcvWiznetInfo(aData: String);
var
  I: Integer;
  TempStr: String;
  st,st2: String;
  DataStr:String;
  ErrorLog: String;
  FHeader:        String[2];
  FMacAddress:    String[12];
  FMode:          String[2];
  FIPAddress:     String[8];
  FSubnet:        String[8];
  FGateway:       String[8];
  FClientPort:    String[4];
  FServerIP:      String[8];
  FServerPort:    String[4];
  FSerial_Baud:   String[2];
  FSerial_data:   String[2];
  FSerial_Parity: String[2];
  FSerial_stop:   String[2];
  FSerial_flow:   String[2];
  FDelimiterChar: String[2];
  FDelimiterSize: String[4];
  FDelimitertime: String[4];
  FDelimiterIdle: String[4];
  FDebugMode:     String[2];
  FROMVer:        String[4];
  FOnDHCP:        String[2];
  FReserve:       String[4];
begin


  Delete(aData,1,1); //�����ͱ��� 1Byte�� ���߿� �߰��Ǿ� ���Ƿ� 1Byte ���� ó��
  bWiznetCheck := True;
  
  DataStr:= copy(aData,22,94);

  FHeader:=            Copy(DataStr, 1,2);
  FMacAddress:=        Copy(DataStr, 3,12);
  FMode:=              Copy(DataStr,15,2);
  FIPAddress:=         Copy(DataStr,17,8);
  FSubnet:=            Copy(DataStr,25,8);
  FGateway:=           Copy(DataStr,33,8);
  FClientPort:=        Copy(DataStr,41,4);
  FServerIP:=          Copy(DataStr,45,8);
  FServerPort:=        Copy(DataStr,53,4);
  FSerial_Baud:=       Copy(DataStr,57,2);
  FSerial_data:=       Copy(DataStr,59,2);
  FSerial_Parity:=     Copy(DataStr,61,2);
  FSerial_stop:=       Copy(DataStr,63,2);
  FSerial_flow:=       Copy(DataStr,65,2);
  FDelimiterChar:=     Copy(DataStr,67,2);
  FDelimiterSize:=     Copy(DataStr,69,4);
  FDelimitertime:=     Copy(DataStr,73,4);
  FDelimiterIdle:=     Copy(DataStr,77,4);
  FDebugMode:=         Copy(DataStr,81,2);
  FROMVer:=            Copy(DataStr,83,4);
  FOnDHCP:=            Copy(DataStr,87,2);
  FReserve:=           Copy(DataStr,89,6);

  //����� aa�̸� ��������
  if FHeader = 'aa' then
  begin
    WizNetRcvACK:= True;
    //btnCheckwiznetClick(Self);
    //Exit;
  end;

  LMDCaptionPanel1.visible := False;
  LMDSimpleLabel1.Twinkle:= False;
  Off_Timer.Enabled:= False;
  LMDListBox1.Clear;

  //2.MAC Address
  TempStr:= 'MAC Addr;'+ FMacAddress;
  LMDListBox1.Items.Add(TempStr);

  //3.Mode (Server mode: 01, Client mode: 00)
  TempStr:= 'Client Mode;';
  if FMode = '00' then TempStr := TempStr + 'Cleint Mode Only'
  else                 TempStr := TempStr + 'Server Mode';
  LMDListBox1.Items.Add(TempStr);

  // 4.IP address
  st2:= '';
  st:= Hex2Ascii(FIPAddress);
  for I:= 1 to 4 do
  begin
    if I < 4 then st2:= st2 + InttoStr(Ord(st[I]))+'.'
    else          st2:= st2 + InttoStr(Ord(st[I]));
  end;
  Edit_LocalIP.Text:= st2;
  Edit_LocalIP.Color := clYellow;
  TempStr:= 'Local/IP;'+ st2;
  LMDListBox1.Items.Add(TempStr);

  // 5.Subnet mask
  st2:= '';
  st:= Hex2Ascii(FSubnet);
  for I:= 1 to 4 do
  begin
    if I < 4 then st2:= st2 + InttoStr(Ord(st[I]))+'.'
    else          st2:= st2 + InttoStr(Ord(st[I]));
  end;
  Edit_Subnet.Text:= st2;
  Edit_Subnet.Color := clYellow;
  TempStr:= 'Local/SubNet;'+ st2;
  LMDListBox1.Items.Add(TempStr);

  // 6.Gateway address
  st2:= '';
  st:= Hex2Ascii(FGateway);
  for I:= 1 to 4 do
  begin
    if I < 4 then st2:= st2 + InttoStr(Ord(st[I]))+'.'
    else          st2:= st2 + InttoStr(Ord(st[I]));
  end;
  Edit_Gateway.Text:= st2;
  Edit_Gateway.Color := clYellow;
  TempStr:= 'Local/Gateway;'+ st2;
  LMDListBox1.Items.Add(TempStr);


  //7.Port number (Client)
  st2:= Hex2DecStr(FClientPort);
  Edit_LocalPort.Text:= st2;
  Edit_LocalPort.Color := clYellow;
  TempStr:= 'Local/Port;'+ st2;
  LMDListBox1.Items.Add(TempStr);


  //8. Server IP address
  st2:= '';
  st:= Hex2Ascii(FServerIP);
  for I:= 1 to 4 do
  begin
    if I < 4 then st2:= st2 + InttoStr(Ord(st[I]))+'.'
    else          st2:= st2 + InttoStr(Ord(st[I]));
  end;
  Edit_ServerIp.Color:= clYellow;
  Edit_ServerIp.Text:= st2;
  TempStr:= 'Server/IP;'+ st2;
  LMDListBox1.Items.Add(TempStr);


  //9.  Port number (Server)
  st2:= '';
  st2:= Hex2DecStr(FServerPort);
  TempStr:= 'Server/Port;'+ st2;
  Edit_Serverport.Color:= clYellow;
  Edit_Serverport.Text:= st2;
  LMDListBox1.Items.Add(TempStr);

  //10. Serial speed (bps)
  TempStr:= 'Serial/Baud;';

  if FSerial_Baud = 'F4' then TempStr:= TempStr + '9600[F4]'
  else if FSerial_Baud = 'FA' then TempStr:= TempStr + '19200[FA]'
  else if FSerial_Baud = 'FD' then TempStr:= TempStr + '38400[FD]'
  else if FSerial_Baud = 'FE' then TempStr:= TempStr + '57600[FE]'
  else if FSerial_Baud = 'FF' then TempStr:= TempStr + '115200[FF]'
  else TempStr:= TempStr +FSerial_Baud;
  LMDListBox1.Items.Add(TempStr);

  //11. Serial data size (08: 8 bit), (07: 7 bit)
  TempStr:= 'Serial/Data;';
  if FSerial_data = '07' then      TempStr:= TempStr + '7'
  else if FSerial_data = '08' then TempStr:= TempStr + '8'
  else                             TempStr:= TempStr + 'error:'+FSerial_data;
  LMDListBox1.Items.Add(TempStr);


  //12. Parity (00: No), (01: Odd), (02: Even)
  TempStr:= 'Serial/Parity;';
  if FSerial_Parity = '00' then      TempStr:= TempStr + 'No'
  else if FSerial_Parity = '01' then TempStr:= TempStr + 'Odd'
  else if FSerial_Parity = '02' then TempStr:= TempStr + 'even'
  else                               TempStr:= TempStr + 'error:'+ FSerial_Parity;
  LMDListBox1.Items.Add(TempStr);

  //13. Stop bit
  TempStr:= 'Serial/Stop;';
  if FSerial_stop = '01' then TempStr:= TempStr + '1'
  else                        TempStr:= TempStr + 'error:'+FSerial_stop;
  LMDListBox1.Items.Add(TempStr);

  //14.Flow control (00: None), (01: XON/XOFF), (02: CTS/RTS)
  TempStr:= 'Serial/Flow;';
  if FSerial_flow = '00' then TempStr:= TempStr + 'None'
  else if FSerial_flow = '01' then TempStr:= TempStr + 'XON/XOFF'
  else if FSerial_flow = '02' then TempStr:= TempStr + 'CTS/RTS'
  else                             TempStr:= TempStr + 'error:'+FSerial_flow;
  LMDListBox1.Items.Add(TempStr);

  //15. Delimiter char
  TempStr:= 'Delimiter/char;';
  TempStr:=TempStr + FDelimiterChar;
  LMDListBox1.Items.Add(TempStr);

  //16.Delimiter size
  TempStr:= 'Delimiter/Size;';
  st2:= '';
  st2:= Hex2DecStr(FDelimiterSize);
  TempStr:= TempStr + st2;
  LMDListBox1.Items.Add(TempStr);

  //17. Delimiter time
  TempStr:= 'Delimiter/Time;';
  st2:= '';
  st2:= Hex2DecStr(FDelimitertime);
  TempStr:= TempStr + st2;
  LMDListBox1.Items.Add(TempStr);

  //18.Delimiter idle time
  TempStr:= 'Delimiter/idleTime;';
  st2:= '';
  st2:= Hex2DecStr(FDelimiterIdle);
  TempStr:= TempStr + st2;
  LMDListBox1.Items.Add(TempStr);

  //19. Debug code (00: ON), (01: OFF)
  TempStr:= 'Debug Mode;';
  if FDebugMode = '00' then TempStr:= TempStr + 'ON[00]'
  else                      TempStr:= TempStr + 'OFF[01]';
  LMDListBox1.Items.Add(TempStr);

  //20.Software major version
  TempStr:= 'version;';
  st:= Hex2Ascii(FROMVer);
  TempStr:= TempStr + InttoStr(Ord(st[1]))+'.'+InttoStr(Ord(st[2]));
  LMDListBox1.Items.Add(TempStr);

  // 21.DHCP option (00: DHCP OFF, 01:DHCP ON)

  TempStr:= 'DHCP;';
  if FOnDHCP = '01' then
  begin
    Checkbox_DHCP.Checked:= True;
    TempStr:= TempStr + 'ON[01]'
  end else if FOnDHCP = '00' then
  begin
    Checkbox_DHCP.Checked:= False;
    TempStr:= TempStr + 'OFF[00]';
  end;

  ErrorLog:= '';

  if FMode <> Copy(wiznetData,15,2)then
     ErrorLog:= ErrorLog +'Mode:' +Copy(wiznetData,15,2) +'<>'+FMode+#13;
  if FIPAddress <> Copy(wiznetData,17,8) then
     ErrorLog:= ErrorLog +'IPAddress:' +Copy(wiznetData,17,8) +'<>'+FIPAddress+#13;
  if FSubnet <> Copy(wiznetData,25,8) then
     ErrorLog:= ErrorLog +'SubNet:' +Copy(wiznetData,25,8) +'<>'+FSubnet+#13;
  if FGateway <> Copy(wiznetData,33,8) then
     ErrorLog:= ErrorLog +'Gateway:' +Copy(wiznetData,33,8) +'<>'+FGateway+#13;
  if FClientPort <> Copy(wiznetData,41,4) then
     ErrorLog:= ErrorLog +'ClientPort:' +Copy(wiznetData,41,8) +'<>'+FClientPort+#13;
  if FServerIP <> Copy(wiznetData,45,8)then
     ErrorLog:= ErrorLog +'ServerIP:' +Copy(wiznetData,45,8) +'<>'+FServerIP+#13;
  if FServerPort <> Copy(wiznetData,53,4)then
     ErrorLog:= ErrorLog +'ServerPort:' +Copy(wiznetData,53,8) +'<>'+FServerPort+#13;
  if FSerial_Baud <> Copy(wiznetData,57,2)then
     ErrorLog:= ErrorLog +'Serial_Baud:' +Copy(wiznetData,57,2) +'<>'+FSerial_Baud+#13;
  if FSerial_data <> Copy(wiznetData,59,2)then
     ErrorLog:= ErrorLog +'Serial_data:' +Copy(wiznetData,59,2) +'<>'+FSerial_data+#13;
  if FSerial_Parity <> Copy(wiznetData,61,2)then
     ErrorLog:= ErrorLog +'Serial_Parity:' +Copy(wiznetData,61,2) +'<>'+FSerial_Parity+#13;
  if FSerial_stop <> Copy(wiznetData,63,2)then
     Errorlog:= Errorlog +'Serial_stop:' +Copy(wiznetData,63,2) +'<>'+FSerial_stop+#13;
  if FSerial_flow <> Copy(wiznetData,65,2)then
     Errorlog:= Errorlog +'Serial_flow:' +Copy(wiznetData,65,2) +'<>'+FSerial_flow+#13;
  if FDelimiterChar <> Copy(wiznetData,67,2)then
     Errorlog:= Errorlog +'DelimiterChar:' +Copy(wiznetData,67,2) +'<>'+FDelimiterChar+#13;
  if FDelimiterSize <> Copy(wiznetData,69,4)then
     Errorlog:= Errorlog +'DelimiterSize:' +Copy(wiznetData,69,2) +'<>'+FDelimiterSize+#13;
  if FDelimitertime <> Copy(wiznetData,73,4)then
     Errorlog:= Errorlog +'Delimitertime:' +Copy(wiznetData,73,4) +'<>'+FDelimitertime+#13;
  if FDelimiterIdle <> Copy(wiznetData,77,4)then
     Errorlog:= Errorlog +'DelimiterIdle:' +Copy(wiznetData,77,4) +'<>'+FDelimiterIdle+#13;
  if FDebugMode <> Copy(wiznetData,81,2) then
     Errorlog:= Errorlog +'DebugMode:' +Copy(wiznetData,81,4) +'<>'+FDebugMode+#13;
  {
  if FROMVer <> Copy(wiznetData,83,4)then
     Errorlog:= Errorlog +'ROMVer:' +Copy(wiznetData,83,4) +'<>'+FROMVer+#13;
  }
  if FOnDHCP <> Copy(wiznetData,87,2)then
     Errorlog:= Errorlog +'OnDHCP:' +Copy(wiznetData,87,4) +'<>'+FOnDHCP;
  if FReserve <> Copy(wiznetData,89,6) then
  if (Errorlog <> '') and (WizNetRegMode = True) then
  begin
     Errorlog:= '�������� ���䰪�� Ʋ���ϴ�.' +#13+
                 '==========================='+#13+
                 '  ������ < ===== > ���䰪  '+#13+
                 '==========================='+#13+
                  Errorlog;
     SHowMessage(Errorlog);
  end;
  {
  else
  begin
    SHowMessage('����/��ȸ �Ϸ� �Ǿ����ϴ�..');
  end;
  }
end;


procedure TForm_Main.Off_TimerTimer(Sender: TObject);
var
  st: string;
  I: Integer;
begin

  if LMDCaptionPanel1.Visible then
  begin
    LMDCaptionPanel1.Left:= 208;
    LMDCaptionPanel1.Top:= 176;
    PollingTimer.Enabled:= False;
    if ProgressBar3.Position >= 19 then
    begin
       if not WizNetRcvACK then
       begin
         LMDCaptionPanel1.visible := False;
         LMDSimpleLabel1.Twinkle:= False;
         Off_Timer.Enabled:= False;
         SHowMessage('��� ��ȸ/������ ���� �߽��ϴ�.��õ� �ϼ���');
         //PollingTimer.Enabled:= True;
       end else
       begin
         btnCheckwiznetClick(Self);
       end;
       Exit;
    end else
    begin
      I:= ProgressBar3.Position;
      Inc(I);
      ProgressBar3.Position:= I;
    end;
  end;
end;

{��ȭ��ȣ ���� ���}
procedure TForm_Main.edRegTelNo0ButtonClick(Sender: TObject);
var
  M_No: Integer;
  T_No: String;
begin
{}
  M_No:= TRzButtonEdit(Sender).Tag;
  T_No:= TRzButtonEdit(Sender).Text;

  //if T_No = '' then ShowMessage('��ȭ��ȣ�� �Է��ϼ���')
  //else if not isdigit(T_No) then ShowMessage('��ȣ�� �Է��ϼ���')
  RegTellNo(xDeviceID,M_No,T_No);
end;

{��ȭ��ȣ ���}
procedure TForm_Main.RegTellNo(aDeviceID: String; aNo: Integer;
  aTellno: String);
var
  NoStr: String[2];
  st: string;
begin
  NoStr:= InttoStr(aNo);
  if Length(NoStr) < 2 then NoStr:= '0'+NoStr;
  st:= SetlengthStr(aTellNo,20);
  SendPacket(aDeviceID,'I','TN01'+NoStr+st);
end;

{��� ��ȭ��ȣ Ȯ��}
procedure TForm_Main.CheckTellNo(aDeviceID: String; aNo: Integer);
var
  NoStr: String[2];
begin
  NoStr:= InttoStr(aNo);
  if Length(NoStr) < 2 then NoStr:= '0'+NoStr;
  SendPacket(aDeviceID,'Q','TN01'+NoStr);
end;

{��ȭ���� ���/��ȸ ��� ó��}
procedure TForm_Main.RcvTellNo(aData: String);
var
  st: string;
  MNo: Integer;
  T_No: String;
  aEdit: TRzButtonEdit;
begin
  Delete(aData,1,1); //�����ͱ��� 1Byte�� ���߿� �߰��Ǿ� ���Ƿ� 1Byte ���� ó��
  st:= copy(aData,22,Length(aData)-24);
  MNo:= StrtoInt(Copy(st,1,2));
  Delete(st,1,2);

  //LMDListBox2.SetItemPart(MNo,1,st);
  //LMDListBox2.ItemIndex:= MNo;
  aEdit:= TRzButtonEdit(FindComponent('edRegTelNo' + InttoStr(MNo)));
  aEdit.Color:= clYellow;
  aEdit.ButtonKind:= bkAccept;
  aEdit.Text:= TrimRight(st);

  if MNO <> 0 then
  begin
    if isRegMode then
    begin
      (*
      repeat
       Inc(MNO);
       T_No:= TRzButtonEdit(FindComponent('edRegTelNo' + InttoStr(MNo))).Text;

      until (T_No <> '') or (MNO >= 9 );

      if (T_No <> '') and (isdigit(T_No) = True) then
        *)
      Inc(MNO);
      if MNO > 9 then MNO := 0;
      T_No:= TRzButtonEdit(FindComponent('edRegTelNo' + InttoStr(MNo))).Text;
      RegTellNo(xDeviceID,MNO,T_No);
    end else
    begin
      Inc(MNO);
      if MNO > 9 then MNO:= 0;
      CheckTellNo(xDeviceID,MNO);

    end;
  end else
  begin
    {
    if isRegMode then ShowMessage('��ȭ��ȣ ����� �Ϸ� �߽��ϴ�.')
    else              ShowMessage('��ȭ��ȣ ��ȸ�� �Ϸ� �߽��ϴ�.');
    }
    Screen.Cursor:= crDefault;
  end;
end;

procedure TForm_Main.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  ComPortList.Free;
  if Comport.Open then
  begin
    Comport.FlushInBuffer;
    Comport.FlushOutBuffer;
    Comport.OPen:= False;
  end;
end;

{��ȭ��ȣ ��ȸ ��ư}
procedure TForm_Main.btnCheckTelNoClick(Sender: TObject);
var
  I: Integer;
begin
  ClearEditTelNo;
  isRegMode:= False;
  CheckTellNo(xDeviceID,1);
  Screen.Cursor:= crHourGlass;
  (*
  for I:= 0 to 9 do
  begin
    CheckTellNo(xDeviceID,I);
    Delay(200);
  end;
  *)
end;

{��ȭ ��ȣ ��ü ���}
procedure TForm_Main.btnSetTelNoClick(Sender: TObject);
var
  I: Integer;
  T_No: String;
const
  NamePrefix = 'edRegTelNo';
begin
  //ClearTelNoList;
  ClearEditTelNo;
  isRegMode:= True;
  T_No:= TRzButtonEdit(FindComponent(NamePrefix + '1')).Text;
  RegTellNo(xDeviceID,1,T_No);
  Screen.Cursor:= crHourGlass;
  //if (T_No <> '') and (isdigit(T_No) = True) then
  //RegTellNo(xDeviceID,I,T_No);
  (*
  for I:= 0 to 9 do
  begin
    T_No:= TRzButtonEdit(FindComponent(NamePrefix + IntToStr(i))).Text;
    if (T_No <> '') and (isdigit(T_No) = True) then
    begin
      RegTellNo(xDeviceID,I,T_No);
      Delay(500);
    end;
  end;
  *)

end;

procedure TForm_Main.LMDAssistPage1NextPage(Sender: TObject;
  var Cancel: Boolean);
begin
end;


{��ȭ����ð� ��ȸ}
procedure TForm_Main.CheckCallTime(aDeviceID: String);
begin
  SendPacket(xDeviceID,'Q','CT01');
end;

{DTMF ��ȸ}
procedure TForm_Main.CheckDialTime(aDeviceID: String);
begin
  isRegMode:= False;
  SendPacket(xDeviceID,'Q','DI01');
end;

{��ȭ ����ð� ���}
procedure TForm_Main.RegCallTime(aDeviceID: String; aTime: Integer);
var
  Timestr: string;
begin
  TimeStr:= FillZeroNumber(aTime,4);
  SendPacket(aDeviceID,'I','CT01'+TimeStr);
end;

{DTMF ���}
procedure TForm_Main.RegDialTime(aDeviceID: String; OnTime,
  OffTime: Integer);
var
  aTime: Integer;
  bTIme: Integer;
  Timestr: string;
begin
  isRegMode:= True;
  aTime:= onTime div 20;
  bTime:= OffTime div 20;

  Timestr:= FillZeroNumber(aTime,2) +   // On Time
            FillZeroNumber(bTime,2);    // Off Time
  SendPacket(aDeviceID,'I','DI01'+TimeStr);
end;

procedure TForm_Main.Btn_CheckDialInfoClick(Sender: TObject);
begin
  CheckDialTime(xDeviceID);
  ClearInfo(4,False);
end;

procedure TForm_Main.Btn_RegDialInfoClick(Sender: TObject);
begin
  RegDialTime(xDeviceID,StrtoInt(ComboBox1.Text),StrtoInt(ComboBox2.Text));
  ClearInfo(4,False);;
end;

procedure TForm_Main.Btn_CheckCalltimeClick(Sender: TObject);
begin
  CheckCallTime(xDeviceID);
end;

procedure TForm_Main.Btn_RegCalltimeClick(Sender: TObject);
begin
  RegCallTime(xDeviceID,RzSpinner2.Value);
end;

{DTMF ���� ó��}
procedure TForm_Main.RcvDialInfo(aData: String);
var
  st: string;

  aTime: Integer;
  bTime: Integer;
begin

  Delete(aData,1,1); //�����ͱ��� 1Byte�� ���߿� �߰��Ǿ� ���Ƿ� 1Byte ���� ó��
  st:= copy(aData,22,4);
  aTime:= StrtoInt(Copy(st,1,2));
  bTime:= StrtoInt(Copy(st,3,2));

  ComboBox1.Color:= clYellow;
  ComboBox2.Color:= clYellow;
  oldtime1:= ComboBox1.Text;
  oldtime2:= ComboBox2.Text;

  ComboBox1.Text:= InttoStr(aTime * 20);
  ComboBox2.Text:= InttoStr(bTime * 20);
  Newtime1:= ComboBox1.Text;
  Newtime2:= ComboBox2.Text;

  {
  SHowMessage('[DTMF ON  TIME]����:'+oldtime1+', ����:'+ComboBox1.Text+#13
             +'[DTMF OFF TIME]����:'+oldtime2+', ����:'+ComboBox2.Text);
   }
  if isRegMode then RegCallTime(xDeviceID,RzSpinner2.Value)
  else              CheckCallTime(xDeviceID);
end;


procedure TForm_Main.RcvCallTime(aData: String);
var
  st: string;
  aTime: Integer;
begin
  Delete(aData,1,1); //�����ͱ��� 1Byte�� ���߿� �߰��Ǿ� ���Ƿ� 1Byte ���� ó��
  st:= copy(aData,22,4);
  aTime:= StrtoInt(st);

  oldTime3:= InttoStr(RzSpinner2.Value);
  NewTime3:= st;
  RzSpinner2.Color:= clYellow;
  RzSpinner2.Value:= aTime;
  //SHowMessage('��ȭ ���� ��ȸ/������ �Ϸ� �߽��ϴ�.');
 {
  SHowMessage('[DTMF ON  TIME]����:'+oldtime1+', ����:'+Newtime1+#13
             +'[DTMF OFF TIME]����:'+oldtime2+', ����:'+Newtime2+#13
             +'[��ȭ�ð�]      ����:'+oldtime3+', ����:'+Newtime3);
  }
  //SHowMessage('[��ȭ�ð�]����:'+oldtime+', ����:'+Newtime);

end;


procedure TForm_Main.CD_DownLoad(aDevice, aCardNo: String; func: Char);
var
  aData: String;
  I: Integer;
  xCardNo: String;

  RealCardNo: String;
  ValidDay: String;
  aLength: Integer;

begin

  //xCardNo:= '00'+ Dec2Hex64(StrtoInt64(aCardNo),8);

  aLength:= Length(aCardNo);
  if aLength < 10 then
    aCardNo:= FillZeroStrNum(aCardNo,10);

  if aLength < 16 then
  begin
    for I := 1  to 16 - aLength do
    begin
      aCardNo:= aCardNo + '0';
    end;
  end;

  RealCardNo:= Copy(aCardNo,1,10);
  ValidDay:=   Copy(aCardNo,11,6);
  xCardNo:=  '00'+EncodeCardNo(RealCardNo);

  aData:= '';
  aData:= func +
          InttoStr(Send_MsgNo)+     // Message Count
          '1'+                      //����ȣ
          '  '+                     //RecordCount #$20 #$20
          '0'+                      //����ڵ�(0:1,2   1:1����,  2:2����, 3:�������)
          xCardNo+                  //ī���ȣ
          ValidDay+                 //��ȿ�Ⱓ
          '0'+                      //��� ���
          '2'+                      //ī�����(0:��������,1:�������,2:���+����)
          '0';                      //Ÿ������

  SendPacket(aDevice,'c',aData);
end;

{��ü ī�� ��ȣ ���}
procedure TForm_Main.CardAllDownLoad(aFunc: Char);
var
  I: Integer;
  DeviceID: String;
  st: string;
begin
  // aFunc:L���,N����,M��ȸ
  if Memo_CardNo.Lines.Count < 1 then
  begin
    ShowMessage('����� ī���ȣ/���Թ�ȣ �� �����ϴ�.');
    Exit;
  end;

  for I:= 0 to Memo_CardNo.Lines.Count -1 do
  begin
    st:= Memo_CardNo.Lines[I];
    if Length(st) < 10 then
    begin
      st:= FillZeroStrNum(st,10);
      Memo_CardNo.Lines[I]:= st;
    end;
    CD_DownLoad(xDeviceID,st,aFunc);
    Delay(200);
  end;

end;

procedure TForm_Main.btnCheckCardNoClick(Sender: TObject);
var
  I: Integer;
begin
  if Memo_CardNo.Lines.Count < 1 then
  begin
    ShowMessage('��ȸ�� ī���ȣ/���Թ�ȣ �� �����ϴ�.');
    Exit;
  end;
  cbAutoReg.Checked:= False;
  CardAllDownLoad('M');
end;


procedure TForm_Main.btnRegCardNoClick(Sender: TObject);
var
  I: Integer;
begin
  if Memo_CardNo.Lines.Count < 1 then
  begin
    ShowMessage('����� ī���ȣ/���Թ�ȣ �� �����ϴ�.');
    Exit;
  end;
  cbAutoReg.Checked:= False;
  CardAllDownLoad('L');
end;

procedure TForm_Main.btnDeleteCardNoClick(Sender: TObject);
var
  I: Integer;
begin
  if Memo_CardNo.Lines.Count < 1 then
  begin
    ShowMessage('������ ī���ȣ/���Թ�ȣ �� �����ϴ�.');
    Exit;
  end;
  cbAutoReg.Checked:= False;
  CardAllDownLoad('N');
end;

{ī���� ����}
procedure TForm_Main.RcvCardRegAck(aData: String);
var
  st:string;
  aCardNo: String;
  bCardNo: int64;
  aIndex: Integer;
  aResult: String;
begin
  aCardNo:= Copy(aData,9,8);
  //bCardNo:= Hex2Dec64(aCardNo);
  //aCardNo:= FillZeroNumber2(bCardNo,10);
  //aIndex:= Memo_CardNo.Lines.IndexOf(st);
  aCardNo:= DeCodeCardNo(aCardNo);

  case aData[23] of
    '1': aResult:= '�̵��';
    '2': aResult:= '���';
    else aResult:= '['+aData[17]+']';
  end;
  if aData[1] = 'l' then
  begin
    if aData[23] = '2' then     st:= aCardNo+';'+'������'
    else if aData[23] = '1' then st:= aCardNo+';'+'�̵��'
    else                         st:= aCardNo+';���['+aData[23]+']';
   //if aIndex < 0 then Memo_CardNo.Lines.Add(aCardNo);
  end else if aData[1] = 'n' then
  begin
    if aData[23] = '2' then st:= aCardNo+';'+'�������'
    else if aData[23] = '1' then st:= aCardNo+';'+'��������'
    else                         st:= aCardNo+';����['+aData[23]+']';

    //if aIndex < 0 then Memo_CardNo.Lines.Delete(aIndex);
  end else if aData[1] = 'm' then
  begin
    if aData[23] = '2' then st:= aCardNo+';'+'������ȸ'
    else if aData[23] = '1' then st:= aCardNo+';'+'��ȸ����'
    else                         st:= aCardNo+';��ȸ['+aData[23]+']';

  end else if aData[1] = 'o' then
  begin
    if aData[23] = '2' then st:= aCardNo+';'+'��ü�����Ϸ�'
    else if aData[23] = '1' then st:= aCardNo+';'+'��ü��������'
    else                         st:= aCardNo+';��ü����['+aData[23]+']';
     //Memo_CardNo.Clear;
  end else
  begin
    st:= aCardNo+';'+aData[1]+aResult;;
  end;
  LMDListBox3.Items.Insert(0,st);
  LMDListBox3.Selected[0]:= True;
end;

{ī�帮���̺�Ʈ���� ó�� }
procedure TForm_Main.RcvAccEventData(aData: String);
var
  st: String;
  aCardNo: String;
  bCardNo: String;
begin
  {0.�ð�}
  st:= Copy(aData,6,2)+'-'+
       Copy(aData,8,2)+'-'+
       Copy(aData,10,2)+' '+
       Copy(aData,12,2)+':'+
       Copy(aData,14,2)+':'+
       Copy(aData,16,2)+';';
  {1.Posi/Nega}
  case aData[18] of
    '0': st:=st + 'Positive;';
    '1': st:= st + 'Negative;';
    else   st:= st+ aData[18];
  end;
  {2.����}
  case aData[19] of
    '0': st:= st+'����;';
    '1': st:= st+'������;';
    else   st:= st+aData[19];
  end;
  {3.�������}
  case aData[20] of
    'C': st:= st + 'ī��;';
    'P': st:= st + '��ȭ;';
    'R': st:= st + '��������;';
    'B': st:= st + '��ư;';
    'S': st:= st + '������;';
    else st:= st + aData[20];
  end;
  {4.���Խ��ΰ��}
  case aData[21]of
    #$30: st:= st + '0;';
    #$31: st:= st + '���Խ���;';
    #$32: st:= st + '���Խ���(*);';
    #$41: st:= st + '�̵��ī��;';
    #$42: st:= st + '���ԺҰ�;';
    #$43: st:= st + '�̵��ī��(*);';
    #$44: st:= st + '��������ԺҰ�;';
    #$45: st:= st + '�������ѽð�;';
    else st:= st + aData[21];
  end;
  {5.���Թ� ����}
  case aData[22]of
    'C': st:= st + '����;';
    'O': st:= st + '����;';
    else st:= st + aData[22];
  end;
  {6.ī���ȣ/���Թ�ȣ}
  aCardNo:= copy(aData,26,8);
  //bCardNo:= Hex2Dec64(aCardNo);
  //st:= st + FillZeroNumber2(bCardNo,10);
  bCardNo:= DecodeCardNo(aCardNo);
  st:= st+ bCardNo;//+ '000000';
  LMDListBox4.Items.InSert(0,st);
  LMDListBox4.Selected[0]:= True;
  //LMDListBox4.Items.Add(st);
  //LMDListBox4.ItemIndex:= LMDListBox1.Items.Count -1;

  if cbAutoReg.Checked = True then
  begin
    CD_DownLoad(xDeviceID,bCardNo+'000000' ,'L');
  end;

end;


procedure TForm_Main.LMDAssistPage6Exit(Sender: TObject);
begin

end;
{�������� ������ }
procedure TForm_Main.ACC_sendData(aDeviceID, aData: String);
begin
  SendPacket(aDeviceID,'c', aData);
end;

procedure TForm_Main.Cnt_CheckVer(aDeviceID: String);
begin
  SendPacket(aDeviceID,'R','VR00');
end;

procedure TForm_Main.Cnt_Reset(aDeviceID: String);
begin
  SendPacket(aDeviceID,'R','RS00Reset');
end;

procedure TForm_Main.Cnt_TimeSync(aDeviceID, aTimeStr: String);
begin
  SendPacket(aDeviceID,'R','TM00'+aTimeStr);
end;

procedure TForm_Main.Btn_SyncTimeClick(Sender: TObject);
var
  TimeStr: String;
begin
  Edit_TimeSync.Text:= '';
  TimeStr:= FormatDatetime('yyyymmddhhnnss',Now);
  Cnt_TimeSync(xDeviceID,TimeStr);
end;

procedure TForm_Main.btnCheckVerClick(Sender: TObject);
begin
  Edit_Ver.Text:= '';
  Cnt_CheckVer(xDeviceID);
end;

procedure TForm_Main.btnResetDeviceClick(Sender: TObject);
begin
  Edit_Reset.Text:= '';
  Cnt_Reset(xDeviceID);
end;

procedure TForm_Main.CheckSysInfo2(aDeviceID: String);
var
  aData: String;
  DeviceID: String;
begin
  aData:= 'B' +                                      //MSG Code
          IntToStr(Send_MsgNo)+                      //Msg Count
          '1'+                                       //����ȣ
          #$20 + #$20 +
          '0' +                                      //1.ī�����
          '0' +                                      //2.���Թ� ����
          '0' +                                      //3.������ð�
          '0' +                                      //4.���氨�ð�
          '0' +                                      //5.�����켳��
          '0' +                                      //6.���Թ���������
          '0' +                                      //7.��������̻�ñ��
          '0' +                                      //8.��Ƽ�н���
          '0' +                                      //9.���氨�ý� ���� ���
          '0' +                                      //10. ����̻�� �������
          '0' +                                      //11.�����Ÿ��
          '0' +                                      //12.ȭ��߻��� ���Թ�����
          '0' +                                      //13.
          '0' +                                      //14.
          '0' +                                      //15.
          '0' +                                      //16.
          '0' +                                      //17.
          '0' +                                      //18.
          '0' +                                      //19.
          '0';                                       //20.
  ACC_sendData(aDeviceID, aData);
end;


{�������� ���}
procedure TForm_Main.RegSysInfo2(aDeviceID: String);
var
  aData: String;
  DeviceID: String;
  aLockType: String[1];
  aDoorLongOpenTime : string;
  stData : string;
begin

  {
  if ComboBox_LockType.ItemIndex = 0 then aLockType:= '1'
  else                                     aLockType:= '4';
   }
  aDoorLongOpenTime := char($30 + SpinEdit_OpenMoni1.IntValue);

  case ComboBox_LockType.ItemIndex of
    0: aLockType:= '1';
    1: aLockType:= '5';
    2: aLockType:= '3';
    else aLockType:= '1';
  end;
  aSysInfo2[2] := IntToStr(Send_MsgNo)[1];
  aSysInfo2[6] := IntToStr(ComboBox_CardModeType.ItemIndex)[1];
  aSysInfo2[8] := IntToStr(SpinEdit_OpenDoor1.intValue)[1];
  aSysInfo2[9] := aDoorLongOpenTime[1];
  aSysInfo2[11] := '1'; //������ ���Թ� ���� ����
  aSysInfo2[16] := aLockType[1];
  aSysInfo2[17] := IntToStr(ComboBox_Fire.ItemIndex)[1];//'0'; //������ ȭ��� ������ ��� ����

{  aData:= 'A' +                                        //MSG Code
          IntToStr(Send_MsgNo)+                         //Msg Count
          '1'+                                          //����ȣ
          #$20 + #$20 +                                 // Record count
          InttoStr(ComboBox_CardModeType.ItemIndex) +   //1.ī�����
          '0' +                                         //2.���Թ� ����(0:�)
          InttoStr(RzSpinner1.Value)+                   //3.Door���� �ð�
          '0'+                                          //4.��ð� ���� �溸 (������)
          '0'+                                          //5.������ ��������(������)
          '1'+                                          //6.���Թ����� ����(���)
          '0'+                                          //7.����̻�� ���
          '0'+                                          //8.�н�ī�� ���� ���(��ɻ���)
          '0'+                                          //9.��ð� ���� �������
          '0'+                                          //10��� �̻�� ���� ���
          aLockType+                                    //11.������ Ÿ��(�Ϲ��� ������ ����)
          '0'+                                          //12.ȭ�� �߻��� ������
          '00000000';                                   //����
   if aSysInfo2 <> aData then
   begin
    showmessage('������ Ʋ��');
   end;  }

   ACC_sendData(xDeviceID, aSysInfo2);
   (*
   Sleep(100);
   aData:=  #90 +
            IntToStr(Send_MsgNo)+                         //Msg Count
            IntToStr(ComboBox_DoorNo2.ItemIndex+1)+
            '  '+#$80;
   ACC_sendData(DeviceID, aData);
   *)

   Delay(1000);
//   I6LP07S601101111111110011111111100
   stData:= 'LP08' + //8����
          'FI' +                      //�����ڵ�
          '1' + //�������� 0.���,1.ȭ��
          '1' + //�˶��߻� ��� 0.1ȸ,1.�߻��ø���
          '1' +   //������ȣ���� 0.���۾���,1����
          '0' +   //+       //�����ð� �������
          '00' +  //{����}
          '00' +  //{���̷�}
          '00' +  //{������1}
          '00' +  //{������2}
          '00' +  //{���Թ� ����1}
          '00' +  //{���Թ� ����2}
          '00' +  //{���η���}
          '00' +  //{���ν��̷�}
          '00' +  //{���θ�����1}
          '00' +  //{���θ�����2}
          '00' +  //{����}
          FillSpace('',16) + //{��ġ����}
          '0400' +  //{�����ð�}
          '1' ; //{���������}

   SendPacket(xDeviceID,'I',stData);
//    ACC_sendData(xDeviceID, stData);
end;


procedure TForm_Main.RcvSysinfo2(aData: String);
var
  aNo: Integer;
  st: String;
begin
//         1         2
//123456789012345678901234567
//a11  0040000000100000000095
  {��� ����ȣ}
  aSysInfo2 := aData;
  aSysInfo2[1]:= 'A';
  aSysInfo2 := copy(aSysInfo2,1,25);

  st:='';
  if aData[3] >= #$30 then
  begin

  end;
  {ī�����}
  if aData[6] >= #$30 then
  begin
    case aData[6] of
      #$30: st:= st + 'ī��:Positive'+#13;
      #$31: st:= st + 'ī��:Negative'+#13;
    end;
    aNo:= StrtoInt(aData[6]);
    ComboBox_CardModeType.Color:= clYellow;
    ComboBox_CardModeType.Text:= ComboBox_CardModeType.Items[aNo];
    ComboBox_CardModeType.ItemIndex:= aNo;
    //ComboBox_CardModeType.Text:= ComboBox_CardModeType.Items[ComboBox_CardModeType.ItemIndex];
  end;
  {���Թ� ����}
  if aData[7] >= #$30 then
  begin
    case aData[7] of
      #$30: st:= st + '���Թ����:�'+#13;
      #$31: st:= st + '���Թ����:����'+#13;
    end;
    //aNo:= StrtoInt(aData[7]);
    //ComboBox_DoorModeType.ItemIndex:= aNo;
    //ComboBox_DoorModeType.Text:= ComboBox_DoorModeType.Items[ComboBox_DoorModeType.ItemIndex];
  end;
  {DOOR����ð�}
  if aData[8] >= #$30 then
  begin
    SpinEdit_OpenDoor1.Color:= ClYellow;
    SpinEdit_OpenDoor1.intValue:= StrtoInt(aData[8]);
    st:= st + '�������ð�:'+aData[8]+'��'+#13;
  end;

  {��ð� ���� �溸}
  if aData[9] >= #$30 then
  begin
   SpinEdit_OpenMoni1.Color:= ClYellow;
   SpinEdit_OpenMoni1.IntValue := ord(aData[9]) - ord(#$30);
  end;
  {������ ���� ����}
  if aData[10] >= #$30 then
  begin
   //aNo:= StrtoInt(aData[10]);
   //ComboBox_UseSch.ItemIndex:= aNo;
   //ComboBox_UseSch.Text:= ComboBox_UseSch.Items[ComboBox_UseSch.ItemIndex];
  end;
  {���Թ� ���� ����}
  if aData[11] >= #$30 then
  begin
   //aNo:= StrtoInt(aData[11]);
   //ComboBox_SendDoorStatus.ItemIndex:= aNo;
   //ComboBox_SendDoorStatus.Text:= ComboBox_SendDoorStatus.Items[ComboBox_SendDoorStatus.ItemIndex];
  end;
  {��� �̻�� ��� }
  if aData[12] >= #$30 then
  begin
   //aNo:= StrtoInt(aData[12]);
   //ComboBox_UseCommOff.ItemIndex:= aNo;
   //ComboBox_UseCommOff.Text:= ComboBox_UseCommOff.Items[ComboBox_UseCommOff.ItemIndex];
  end;
  {��� �̻�� ���� ���}
  if aData[15] >= #$30 then
  begin
   //aNo:= StrtoInt(aData[13]);
   //ComboBox_AlramCommoff.ItemIndex:= aNo;
   //ComboBox_AlramCommoff.Text:= ComboBox_AlramCommoff.Items[ComboBox_AlramCommoff.ItemIndex];
  end;
  {��ð� ���� ���� ���}
  if aData[14] >= #$30 then
  begin
   //aNo:= StrtoInt(aData[14]);
   //ComboBox_AlarmLongOpen.ItemIndex:= aNo;
   //ComboBox_AlarmLongOpen.Text:= ComboBox_AlarmLongOpen.Items[ComboBox_AlarmLongOpen.ItemIndex];
  end;
  {������ Ÿ��}
  if aData[16] >= #$30 then
  begin
    ComboBox_LockType.Color:= clYellow;

    case aData[16] of
      '1': ComboBox_LockType.ItemIndex:= 0;
      '5': ComboBox_LockType.ItemIndex:= 1; 
      '3': ComboBox_LockType.ItemIndex:= 2;
    end;
    {
    if aData[16] = '1' then ComboBox_LockType.ItemIndex:= 0
    else                    ComboBox_LockType.ItemIndex:= 1;
    }
   //aNo:= StrtoInt(aData[16]);
   //ComboBox_LockType.ItemIndex:= aNo;
   //ComboBox_LockType.Text:= ComboBox_LockType.Items[ComboBox_LockType.ItemIndex];
  end;
  {ȭ�� �߻��� ������}
  if aData[17] >= #$30 then
  begin
   aNo:= StrtoInt(aData[17]);
   ComboBox_Fire.ItemIndex:= aNo;
   ComboBox_Fire.Color:= clYellow;
   //ComboBox_ControlFire.Text:= ComboBox_ControlFire.Items[ComboBox_ControlFire.ItemIndex];
  end;
  //ShowMessage(st);

end;

procedure TForm_Main.btnDoorOPenClick(Sender: TObject);
begin
  DoorControl(xDeviceID,1,'2','1' );
end;
procedure TForm_Main.btnDoorCloseClick(Sender: TObject);
begin
  DoorControl(xDeviceID,1,'2','0' );
end;

procedure TForm_Main.DoorControl(DeviceID: String; aNo: Integer;
  aControlType, aControl: Char);
var
  st: string;
begin

  st:= 'C' +
       IntToStr(Send_MsgNo)+                       //Msg Count
       InttoStr(aNo)+                             //��⳻ ����ȣ
       '0'+
       aControlType+                              //'0':�ش���׾���,'1':ī��,'2':���Կ,'3':��������
       aControl+                                  // ī��:Positive:'0',Negative:'1'
       '0000000000';                            // ���Կ:'0':� ,'1':����
                                                  // ��������:'0':����,'1':����
  ACC_sendData(DeviceID, st);
end;

procedure TForm_Main.btnRegDoorInfoClick(Sender: TObject);
begin
  ClearInfo(6,False);
  RegSysInfo2(xDeviceID);
end;


procedure TForm_Main.btnCheckDoorInfoClick(Sender: TObject);
begin
  aSysInfo2:= 'A' +                                        //MSG Code
          '0'+                         //Msg Count
          '1'+                                          //����ȣ
          #$20 + #$20 +                                 // Record count
          '0' +   //1.ī�����
          '0' +                                         //2.���Թ� ����(0:�)
          '5'+                   //3.Door���� �ð�
          '0'+                                          //4.��ð� ���� �溸 (������)
          '0'+                                          //5.������ ��������(������)
          '1'+                                          //6.���Թ����� ����(���)
          '0'+                                          //7.����̻�� ���
          '0'+                                          //8.�н�ī�� ���� ���(��ɻ���)
          '0'+                                          //9.��ð� ���� �������
          '0'+                                          //10��� �̻�� ���� ���
          '1'+                                    //11.������ Ÿ��(�Ϲ��� ������ ����)
          '0'+                                          //12.ȭ�� �߻��� ������
          '00000000';                                   //����
  CheckSysInfo2(xDeviceID);
  ClearInfo(6,False);
end;

procedure TForm_Main.Btn_DelEventLogClick(Sender: TObject);
begin
  LMDListBox4.Clear;
end;

procedure TForm_Main.RzBitBtn5Click(Sender: TObject);
begin
  if MessageDlg('��⳻ ���Ե�� �����Ͱ� ��ü ���� �˴ϴ�.���� �Ͻðڽ��ϱ�?',
    mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    CD_DownLoad(xDeviceID,'0000000000000000','O');
    //Memo_CardNo.Clear;
  end;
end;


procedure TForm_Main.AssistChanging(Sender: TObject;
  NewPage: TLMDAssistPage; var Cancel: Boolean);
begin
{  //SHowMessage(InttoStr(NewPage.PageIndex));
 Screen.Cursor:= crDefault;
  case NewPage.PageIndex of
    1:begin

        if cb_Next.Checked then
        begin
          Cancel:= False;
          btnIDCheckClick(self);
          Exit;
        end;

        if (ComPort.Open = False) or (Panel1.Color = clRed) then
        begin
          SHowMessage('��ſ����� �ϼ���');
          Cancel:= False;

        end else
        begin
          btnIDCheckClick(self);
        end;
      end;
    2:begin
        if Not bConnected then
        begin
          SHowMessage('��ſ����� �ϼ���');
          Cancel:= False;
          Exit;
        end;
        btnCheckwiznetClick(Self);
      end;
    3:begin
        if Not bConnected then
        begin
          SHowMessage('��ſ����� �ϼ���');
          Cancel:= True;
          Exit;
        end;
        btnCheckTelNoClick(Self);
      end;
    4:begin
        if Not bConnected then
        begin
          SHowMessage('��ſ����� �ϼ���');
          Cancel:= True;
          Exit;
        end;
        Btn_CheckDialInfoClick(Self);
      end;
    5:begin
        if Not bConnected then
        begin
          SHowMessage('��ſ����� �ϼ���');
          Cancel:= True;
          Exit;
        end;

      end;
    6:begin
        if Not bConnected then
        begin
          SHowMessage('��ſ����� �ϼ���');
          Cancel:= True;
          Exit;
        end;
        CheckSysInfo2(xDeviceID);
      end;
    7:begin
        if Not bConnected then
        begin
          SHowMessage('��ſ����� �ϼ���');
          Cancel:= True;
          Exit;
        end;

      end;
    8:begin
        if Not bConnected then
        begin
          SHowMessage('��ſ����� �ϼ���');
          Cancel:= True;
          Exit;
        end;
        ed_DB_Deviceid.Text:= xDeviceID;
        ed_DB_siteName.Text:= Edit_Name.Text;

      end;
  end;
  (*
  if NewPage.PageIndex = 1 then
  begin
    if (ComPort.Open = False) or (Panel1.Color = clRed) then
    begin
      SHowMessage('��ſ����� �ϼ���');
      Cancel:= True;
    end else
    begin
      btnIDCheckClick(self);
    end;
  end else if NewPage.PageIndex = 7 then
  begin
    if MessageDlg('��⸦ ����� �Ͻðڽ��ϱ�?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      Cnt_Reset(xDeviceID);
    end;

  end;
  *)   }
end;

procedure TForm_Main.PollingTimerTimer(Sender: TObject);
var
  aTimeOut: TDatetime;
begin
  aTimeOut:= IncTime(RcvPollingTime,00,00,15,00);
  if Now > aTimeout then
  begin
    PollingTimer.Enabled:= False;
    Panel1.Color:= clRed;
    ComPort.FlushInBuffer;
    ComPort.FlushOutBuffer;
    SHowMessage('��� ���̺��� Ȯ�� �Ͻðų� ��Ϲ�ư�� �ٽ� ����ġ �ϼ���');
  end;

end;

procedure TForm_Main.edRegTelNo0Change(Sender: TObject);
begin
  edRegTelNo0.ButtonKind:= bkReject;
end;

procedure TForm_Main.edRegTelNo1Click(Sender: TObject);
begin
  TrzButtonEdit(Sender).Color:= clBtnFace;
  TrzButtonEdit(Sender).ButtonKind:= bkReject;
end;

procedure TForm_Main.ClearEditTelNo;
var
  I: Integer;
  aButtonEdit: TRzButtonEdit;
begin
  for I:= 0 to 9 do
  begin
    aButtonEdit:= TRzButtonEdit(FindComponent('edRegTelNo' + IntToStr(i)));
    aButtonEdit.Color:= clBtnFace;
    aButtonEdit.ButtonKind:= bkReject;
  end;
  
end;


procedure TForm_Main.RcvDoorEventData(aData: String);
var
  st: String;
begin
  {0.�ð�}
  st:= Copy(aData,6,2)+'-'+
       Copy(aData,8,2)+'-'+
       Copy(aData,10,2)+' '+
       Copy(aData,12,2)+':'+
       Copy(aData,14,2)+':'+
       Copy(aData,16,2)+';';
  {1.Posi/Nega}
  case aData[18] of
    '0': st:=st + 'Positive;';
    '1': st:= st + 'Negative;';
    else   st:= st+ aData[18]+' ;';
  end;
  {2.����}
  case aData[19] of
    '0': st:= st+'����;';
    '1': st:= st+'������;';
    else   st:= st+aData[19]+' ;';
  end;
  {3.�������}
  case aData[20] of
    'C': st:= st + 'ī��;';
    'P': st:= st + '��ȭ;';
    'R': st:= st + '��������;';
    'B': st:= st + '��ư;';
    'S': st:= st + '������;';
    else st:= st + aData[20]+' ;';
  end;
  {4.���Խ��ΰ��}
  case aData[21]of
    #$30: st:= st + '0;';
    #$31: st:= st + '���Խ���;';
    #$32: st:= st + '�������;';
    #$41: st:= st + '�̵��ī��;';
    #$42: st:= st + '���ԺҰ�;';
    #$43: st:= st + '����Ұ�;';
    #$44: st:= st + '��������ԺҰ�;';
    #$45: st:= st + '�������ѽð�;';
    else st:= st + aData[21]+' ;';
  end;
  {5.���Թ� ����}
  case aData[22]of
    'C': st:= st + '����';
    'O': st:= st + '����';
    else st:= st + aData[22];
  end;

  LMDListBox4.Items.InSert(0,st);
  LMDListBox4.Selected[0]:= True;

end;

procedure TForm_Main.ClearInfo(aIndex:Integer; isEmpty:Boolean);
begin
  case aIndex of
    1:begin // ��� IP ���
        Edit_DeviceID1.Color:= clWhite;
        Edit_DeviceID2.Color:= clWhite;
        Edit_DeviceID3.Color:= clWhite;
        Edit_DeviceID4.Color:= clWhite;
        Edit_DeviceID5.Color:= clWhite;
        Edit_DeviceID6.Color:= clWhite;
        Edit_DeviceID7.Color:= clWhite;
        ComboBox_DoorType1.Color:= clWhite;
        Edit_Name.Color:= clWHite;
        if isEmpty then
        begin
          Edit_DeviceID1.Text:= '';
          Edit_DeviceID2.Text:= '';
          Edit_DeviceID3.Text:= '';
          Edit_DeviceID4.Text:= '';
          Edit_DeviceID5.Text:= '';
          Edit_DeviceID6.Text:= '';
          Edit_DeviceID7.Text:= '';
          ComboBox_DoorType1.Text:= '';
          Edit_Name.Text:= '';
        end;
      end;
    2:begin  //LAN ��� ����
        Edit_ServerIp.Color:= clWhite;
        Edit_Serverport.Color:= clWhite;
        Edit_LocalIP.Color := clWhite;
        Edit_Subnet.Color := clWhite;
        Edit_Gateway.Color := clWhite;
        Edit_LocalPort.Color := clWhite;
        if isEmpty then
        begin
          Edit_ServerIp.Text:= '';
          Edit_Serverport.Text:= '';
          Edit_LocalIP.Text:= '';
          Edit_Subnet.Text:= '';
          Edit_Gateway.Text:= '';
          Edit_LocalPort.Text:= '';
        end;
      end;
    3:begin
        Edit_LinkusID.Color:= clWhite;
        Edit_LinkusTel.Color:= clWhite;
        Spinner_Ring.Color:= clWhite;
        if isEmpty then
        begin
          Edit_LinkusID.Text:= '';
          Edit_LinkusTel.Text:= '';
          //Spinner_Ring.Value:= 0;
        end;
      end;
    4:begin //��ȭ����
        ComboBox1.Color:= clWhite;
        ComboBox2.Color:= clWhite;
        RzSpinner2.Color:= clWhite;
        if isEmpty then
        begin
          ComboBox1.Text:= '';
          Combobox2.Text:= '';
        end;
      end;
    5:begin
      end;
    6:begin
        SpinEdit_OpenDoor1.Color:= clWhite;
        SpinEdit_OpenMoni1.Color := clWhite;
        ComboBox_CardModeType.Color:= clWhite;
        ComboBox_LockType.Color:= clWhite;
        ComboBox_Fire.Color := clWhite;
        if isEmpty then
        begin
          ComboBox_CardModeType.Text:= '';
          ComboBox_LockType.Text:= '';
        end;
      end;
  end;
end;

procedure TForm_Main.Edit_DeviceID1Click(Sender: TObject);
begin
  ClearInfo(1,False);
end;

procedure TForm_Main.Edit_ServerIpClick(Sender: TObject);
begin
  ClearInfo(2,False);
end;

procedure TForm_Main.ComboBox1Click(Sender: TObject);
begin
  ClearInfo(4,False);
end;

procedure TForm_Main.RzSpinner1Click(Sender: TObject);
begin
  ClearInfo(6,False);
end;

procedure TForm_Main.Btn_FileOpenClick(Sender: TObject);
begin
  //OpenDialog1.
end;


{��Ŀ�� ID ���}
function TForm_Main.RegLinkusID(aDeviceID, aLinkusId: String):boolean;
var
  aID: Integer;
  bID: string;
  PastTime : dword;

begin
  Result := false;
  if not isDigit(Edit_LinkusID.Text) then
  begin
    ShowMessage('�߸��� ����ID�Դϴ�.');
    exit;
  end;
  
  if (not isdigit(aDeviceID)) or (Length(aDeviceID) <> 9) then Exit;
  if not isdigit(aLinkusId) then Exit;
  aID := StrToInt(aLinkusId);
  bID := Dec2Hex(aID, 4);
  bLinkusIDCheck := False;
  SendPacket(aDeviceID, 'I', 'Id00' + bID);

  PastTime := GetTickCount + DelayTime;
  while Not bLinkusIDCheck do
  begin
    if Not bConnected then Exit;
    Application.ProcessMessages;
    if GetTickCount > PastTime then Exit;  //300�и����� ���� ������ ���з� ó����
  end;

  Result := true;

end;

{��Ŀ�� ID ��ȸ}
function TForm_Main.CheckLinkusID(aDeviceID: String):boolean;
var
  aData: string;
  nTime : integer;
  PastTime : dword;

begin
  Result := false;
  if (not isdigit(aDeviceID)) or (Length(aDeviceID) <> 9) then Exit;
  aData := 'Id00';
  
  bLinkusIDCheck  := False;
  SendPacket(aDeviceID, 'Q', aData);

  nTime := 0;
  PastTime := GetTickCount + DelayTime;
  while Not bLinkusIDCheck do
  begin
    if Not bConnected then Exit;
    Application.ProcessMessages;
    if GetTickCount > PastTime then Exit;  //300�и����� ���� ������ ���з� ó����
  end;

  Result := true;
end;

procedure TForm_Main.RcvLinkusId(aData: String);
var
  st: string;
  st2: String;
  aID: Integer;
begin
  Delete(aData,1,1); //�����ͱ��� 1Byte�� ���߿� �߰��Ǿ� ���Ƿ� 1Byte ���� ó��
  st:=  Copy(aData,22,4);
  aID:= Hex2Dec(st);
  st2:= FillZeroNumber(aID,5);
  if st2[1] = '0' then Delete(st2,1,1);
  Edit_LinkusID.Color:= clYellow;
  Edit_LinkusID.Text:= st2;
  bLinkusIDCheck := True;

//  if isRegMode then RegLinksTellNo(xDeviceID,0,Edit_LinkusTel.Text)
//  else              CheckLinkusTellNo(xDeviceID,0);
  

end;

function TForm_Main.CheckLinkusTellNo(aDeviceID: String; aNo: Integer):boolean;
var
  IndexStr: string;
  aData: string;
  PastTime : dword;

begin
  Result := false;
  if (not isdigit(aDeviceID)) or (Length(aDeviceID) <> 9) then Exit;
  IndexStr := '0' + IntToStr(aNo);
  aData := 'Tn00' +  //COMMAND
    IndexStr;        //��ȭ ��ȣ �ε���
  bLinksTelNoCheck := False;
  SendPacket(aDeviceID, 'Q', aData);

  PastTime := GetTickCount + DelayTime;
  while Not bLinksTelNoCheck do
  begin
    if Not bConnected then Exit;
    Application.ProcessMessages;
    if GetTickCount > PastTime then Exit;  //300�и����� ���� ������ ���з� ó����
  end;

  Result := true;
end;

function TForm_Main.RegLinksTellNo(aDeviceID: String; aNo: Integer;aTellno: String):boolean;
var
  NoStr: string[2];
  st: string;
  PastTime : dword;

begin
  Result := false;
  if (not isdigit(aDeviceID)) or (Length(aDeviceID) <> 9) then Exit;
  //if not isdigit(aTellNo) then Exit;
  NoStr := IntToStr(aNo);
  if Length(NoStr) < 2 then NoStr := '0' + NoStr;
  st := SetlengthStr(aTellNo, 20);

  bLinksTelNoCheck := False;
  SendPacket(aDeviceID, 'I', 'Tn00' + NoStr + st);

  PastTime := GetTickCount + DelayTime;
  while Not bLinksTelNoCheck do
  begin
    if Not bConnected then Exit;
    Application.ProcessMessages;
    if GetTickCount > PastTime then Exit;  //300�и����� ���� ������ ���з� ó����
  end;

  Result := true;

end;

procedure TForm_Main.RcvLinkusTelNo(aData: String);
var
  st: string;
  MNo: Integer;
begin
  Delete(aData,1,1); //�����ͱ��� 1Byte�� ���߿� �߰��Ǿ� ���Ƿ� 1Byte ���� ó��
  st:= copy(aData,22,Length(aData)-24);
  MNo:= StrtoInt(Copy(st,1,2));
  Delete(st,1,2);
  DeleteChar(st,' ');
  Edit_LinkusTel.Color:= clYellow;
  if MNo = 0 then Edit_LinkusTel.Text:= st
  else
  begin
    SHowMessage('������ ���� �Դϴ�.');
    Exit;
  End;
  bLinksTelNoCheck := True;

//  if isRegMode then RegRingCount(xDeviceID,Spinner_Ring.Value)
//  else              CheckRingCount(xDeviceID);

end;

function TForm_Main.CheckRingCount(aDeviceID: String):boolean;
var
  aData: string;
  PastTime : dword;
begin
  Result := false;
  if (not isdigit(aDeviceID)) or (Length(aDeviceID) <> 9) then Exit;
  aData := 'Rc00';
  bRingCountCheck := False;
  SendPacket(aDeviceID, 'Q', aData);

  PastTime := GetTickCount + DelayTime;
  while Not bRingCountCheck do
  begin
    if Not bConnected then Exit;
    Application.ProcessMessages;
    if GetTickCount > PastTime then Exit;  //300�и����� ���� ������ ���з� ó����
  end;

  Result := true;
end;

function TForm_Main.RegRingCount(aDeviceID: String; aCount: Integer):Boolean;
var
  Countstr: string;
  PastTime : dword;

begin
  Result := false;
  if (not isdigit(aDeviceID)) or (Length(aDeviceID) <> 9) then Exit;
  CountStr := FillZeroNumber(aCount, 2);
  bRingCountCheck := False;
  SendPacket(aDeviceID, 'I', 'Rc00' + CountStr);

  PastTime := GetTickCount + DelayTime;
  while Not bRingCountCheck do
  begin
    if Not bConnected then Exit;
    Application.ProcessMessages;
    if GetTickCount > PastTime then Exit;  //300�и����� ���� ������ ���з� ó����
  end;

  Result := true;
end;

procedure TForm_Main.RcvRingCount(aData: String);
var
  st: string;
  aCount: Integer;
begin
  Spinner_Ring.Color:= ClYellow;
  Delete(aData,1,1); //�����ͱ��� 1Byte�� ���߿� �߰��Ǿ� ���Ƿ� 1Byte ���� ó��
  st:= copy(aData,22,2);
  aCount:= StrtoInt(st);
  Spinner_Ring.Value:= aCount;
  bRingCountCheck := True;

//  if isRegMode then autoSetRemoteCheckTime;

end;

procedure TForm_Main.btnCheckLinkusIDClick(Sender: TObject);
var
  bResult : Boolean;
  nLoop : integer;
begin
  Edit_LinkusID.Color:= clWhite;
  Edit_LinkusTel.Color:= clWhite;
  Spinner_Ring.Color:= ClWhite;;

  for nLoop := 0 to LoopCount - 1 do
  begin
    bResult := CheckLinkusID(xDeviceID);
    if Not bConnected then Exit;
    if bResult then break;
  end;
  if not bResult then
  begin
    showmessage('CHKKTTID��������� �����ϴ�.���۸�带 Ȯ���Ͽ� �ּ���.');
    Exit;
  end;

  for nLoop := 0 to LoopCount - 1 do
  begin
    bResult := CheckLinkusTellNo(xDeviceID,0);
    if Not bConnected then Exit;
    if bResult then break;
  end;
  if not bResult then
  begin
    showmessage('CHKDECODERNO��������� �����ϴ�.���۸�带 Ȯ���Ͽ� �ּ���.');
    Exit;
  end;

  for nLoop := 0 to LoopCount - 1 do
  begin
    bResult := CheckRingCount(xDeviceID);
    if Not bConnected then Exit;
    if bResult then break;
  end;
  if not bResult then
  begin
    showmessage('CHKDECODERNO��������� �����ϴ�.���۸�带 Ȯ���Ͽ� �ּ���.');
    Exit;
  end;



end;

procedure TForm_Main.btnRegLinkusIDClick(Sender: TObject);
var
  nLoop :integer;
  bResult : Boolean;
begin

  Edit_LinkusID.Color:= clWhite;
  Edit_LinkusTel.Color:= clWhite;
  Spinner_Ring.Color:= ClWhite;

  if (Edit_LinkusID.Text = '') or (isDigit(Edit_LinkusID.Text) = False) then
  begin
    ShowMessage('�ý���ID �� �ùٸ��� �ʽ��ϴ�.');
    Exit;
  end else if (Edit_LinkusTel.Text = '') or (isDigit(Edit_LinkusTel.Text) = False) then
  begin
    ShowMessage('������ȭ��ȣ�� �ùٸ��� �ʽ��ϴ�.');
    Exit;
  end;

  for nLoop := 0 to LoopCount - 1 do
  begin
    bResult := RegLinkusID(xDeviceID,Edit_LinkusID.Text);
    if Not bConnected then Exit;
    if bResult then break;
  end;
  if not bResult then
  begin
    showmessage('�ý���ID ��� ��������� �����ϴ�.���۸�带 Ȯ���Ͽ� �ּ���.');
    Exit;
  end;

  for nLoop := 0 to LoopCount - 1 do
  begin
    bResult := RegLinksTellNo(xDeviceID,0,Edit_LinkusTel.Text);
    if Not bConnected then Exit;
    if bResult then break;
  end;
  if not bResult then
  begin
    showmessage('���ڴ���ȣ ��� ��������� �����ϴ�.���۸�带 Ȯ���Ͽ� �ּ���.');
    Exit;
  end;

  for nLoop := 0 to LoopCount - 1 do
  begin
    bResult := RegRingCount(xDeviceID,Spinner_Ring.Value);
    if Not bConnected then Exit;
    if bResult then break;
  end;

  if not bResult then
  begin
    showmessage('��ī���� ��� ��������� �����ϴ�.���۸�带 Ȯ���Ͽ� �ּ���.');
    Exit;
  end;

  autoSetRemoteCheckTime;
  //ClearInfo(3,True);
end;

procedure TForm_Main.Edit_NameEnter(Sender: TObject);
begin
  Edit_Name.ImeMode:=  imSHanguel;
end;

procedure TForm_Main.ComboBox_DoorType1Change(Sender: TObject);
begin
  if (ComboBox_DoorType1.ItemIndex = 1) then Group_Linkus.Visible:= False
  else                                       Group_Linkus.Visible:= True;
end;

procedure TForm_Main.RegRemoteCheckTime(aDeviceID:String; aTime: Integer);
var
  TimeStr:String;
begin
  TimeStr:= InttoStr(aTime);
  if Length(TimeStr) < 2 then TimeStr:= '0'+ TimeStr;
  SendPacket(aDeviceID,'R','Pt00'+TimeStr);
end;

procedure TForm_Main.ComPortTriggerTimer(CP: TObject; TriggerHandle: Word);
begin
{  if TriggerHandle = TimeoutTimer then
  begin
    if RetryCount < 4 then
    begin
      ComPort.PutString(LastSendData);
      ComPort.SetTimerTrigger( TimeoutTimer, TimeoutTicks, True); { TimeOutTimer enable }
{      Inc(RetryCount);
    end else
    begin
      Screen.Cursor:= crDefault;
    end;
  end;}
end;
procedure TForm_Main.Edit_LinkusIDEnter(Sender: TObject);
begin
  //TRzEdit(Sender).Color:= clwhite;
end;

procedure TForm_Main.Spinner_RingChanging(Sender: TObject;
  NewValue: Integer; var AllowChange: Boolean);
begin
  TrzSpinner(Sender).Color:= clwhite;
end;



procedure TForm_Main.autoSetRemoteCheckTime;
var
  dTime: TDatetime;
  IntervalTime: TDatetime;
  Hour, Min, Sec, MSec: Word;
  aTime: Integer;
begin
  dTime:= Trunc(tomorrow) + EncodeTime(2,0,0,0);
  IntervalTime:= dTime - Now;
  Decodetime(IntervalTime,Hour, Min, Sec, MSec);
  aTime:= Hour+1;
  RegRemoteCheckTime(xDeviceID,aTime);
end;


procedure TForm_Main.Edit_LinkusIDClick(Sender: TObject);
begin
    ClearInfo(3,False);
end;

procedure TForm_Main.OpenDoor;
var
  st: String;
begin
  st:= 'C'+                              //  Msg Code
       '0'+                              //  Msg Count
       '1'+                              //  ��⳻ Door No
       #$30+                             //  RecordCount(����)
       #$33+                             //  RecordCount(���� #$33)
       #$31+
       '0000000000';
  SendPacket(xDeviceID,'c',st);
end;

procedure TForm_Main.FormShow(Sender: TObject);
var
  nCount : integer;
  i : integer;
begin
  TB_DEVICE.CreateTable;
  nSerialPort:= LMDIniCtrl1.ReadInteger('����','ComNo',0);

  nCount := GetSerialPortList(ComPortList);
  cbSerialPort.Items.Clear;
  if nCount > 0 then
  begin
    for i:= 0 to nCount - 1 do
    begin
      cbSerialPort.items.Add(ComPortList.Strings[i])
    end;
    cbSerialPort.ItemIndex := -1;
    if cbSerialPort.Items.Count > 0 then cbSerialPort.ItemIndex := 0;
    for i:= 0 to cbSerialPort.Items.Count -1 do
    begin
      if Integer(ComPortList.Objects[i]) = nSerialPort then
      begin
        cbSerialPort.ItemIndex := i;
        break;
      end;
    end;
  end;

  cb_DB_Center.ItemIndex:= LMDIniCtrl1.ReadInteger('����','KT����',0);
  ed_DB_Local.Text:= LMDIniCtrl1.ReadString('����','KT�����','');
  ed_DB_InstallCo.Text:=LMDIniCtrl1.ReadString('����','��ġ��ü��','');
  ed_DB_installName.Text:=LMDIniCtrl1.ReadString('����','��ġ����ڸ�','');
  ed_DB_installTell.Text:=LMDIniCtrl1.ReadString('����','����ó','');
  LoopCount := LMDIniCtrl1.ReadInteger('����','LoopCount',10)  ; //���� ��� Ƚ��
  DelayTime := LMDIniCtrl1.ReadInteger('����','DelayTime',600);  //���� �۽��� ��� �ð�
  nClose := 0;
  bConnected := False;
  NoteBook1.PageIndex := 0;

  aSysInfo2:= 'A' +                                        //MSG Code
          '0'+                         //Msg Count
          '1'+                                          //����ȣ
          #$20 + #$20 +                                 // Record count
          '0' +   //1.ī�����
          '0' +                                         //2.���Թ� ����(0:�)
          '5'+                   //3.Door���� �ð�
          '0'+                                          //4.��ð� ���� �溸 (������)
          '0'+                                          //5.������ ��������(������)
          '1'+                                          //6.���Թ����� ����(���)
          '0'+                                          //7.����̻�� ���
          '0'+                                          //8.�н�ī�� ���� ���(��ɻ���)
          '0'+                                          //9.��ð� ���� �������
          '0'+                                          //10��� �̻�� ���� ���
          '1'+                                    //11.������ Ÿ��(�Ϲ��� ������ ����)
          '0'+                                          //12.ȭ�� �߻��� ������
          '00000000';                                   //����
end;

procedure TForm_Main.RzBitBtn3Click(Sender: TObject);
begin
  Screen.Cursor:= crHourGlass;
  {DB ���� Ȯ��}
{  try
    ZConnection1.Connect;
  Except
    on E : EDatabaseError do
    begin
      Application.MessageBox (PChar(E.Message),PChar('Ȯ��'),MB_ICONSTOP or MB_OK);
    end else
    begin
      Application.MessageBox (PChar('DB���� ���� ����!!!'),PChar('Ȯ��'),MB_ICONSTOP or MB_OK);
    end;
    Screen.Cursor:= crDefault;
    Exit;
  end;

  if ed_DB_Local.Text = '' then
  begin
    ShowMessage('������� �Է��ϼ���');
    ed_DB_Local.SetFocus;
    Exit;
  end;

  if ed_DB_sLocal.Text = '' then
  begin
    ShowMessage('�������� �Է��ϼ���');
    ed_DB_sLocal.SetFocus;
    Exit;
  end;

  if ed_DB_siteName.Text = '' then
  begin
    ShowMessage('������� �Է��ϼ���');
    ed_DB_siteName.SetFocus;
    Exit;
  end;

  if ed_DB_pstn.Text = '' then
  begin
    ShowMessage('������ȭ��ȣ�� �Է��ϼ���');
    ed_DB_pstn.SetFocus;
    Exit;
  end;

  if ed_DB_ip.Text = '' then
  begin
    ShowMessage('���� IP�ּҸ� �Է��ϼ���');
    ed_DB_ip.SetFocus;
    Exit;
  end;

  if ed_DB_subnet.Text = '' then
  begin
    ShowMessage('���� ������� �Է��ϼ���');
    ed_DB_subnet.SetFocus;
    Exit;
  end;

  if ed_DB_gateway.Text = '' then
  begin
    ShowMessage('���� ����Ʈ���̸� �Է��ϼ���');
    ed_DB_gateway.SetFocus;
    Exit;
  end;

  if ed_DB_InstallCo.Text = '' then
  begin
    ShowMessage('��ġ��ü���� �Է��ϼ���');
    ed_DB_InstallCo.SetFocus;
    Exit;
  end;

  if ed_DB_installName.Text = '' then
  begin
    ShowMessage('��ġ����ڸ��� �Է��ϼ���');
    ed_DB_installName.SetFocus;
    Exit;
  end;

  if ed_DB_installTell.Text = '' then
  begin
    ShowMessage('��ġ����� ����ó�� �Է��ϼ���');
    ed_DB_installTell.SetFocus;
    Exit;
  end;
}
  {�̹� ����� ������� Ȯ��}

{  if IsRegDB(ed_DB_Deviceid.Text) then
  begin
    if MessageDlg('�̹� ��ϵ� ���� �Դϴ�. ������ ���� �Ͻðڽ��ϱ�?',
    mtConfirmation, [mbYes, mbNo], 0) = mrYes then  UpDateDB;
  end else
  begin
    InsertDB;
  end;

  if ZConnection1.Connected then ZConnection1.Disconnect;
  Screen.Cursor:= crDefault;

  LMDIniCtrl1.WriteInteger('����','ComNo',      cbSerialPort.ItemIndex);
  LMDIniCtrl1.WriteInteger('����','KT����',     cb_DB_Center.ItemIndex);
  LMDIniCtrl1.WriteString('����','KT�����',    ed_DB_Local.Text);
  LMDIniCtrl1.WriteString('����','��ġ��ü��',  ed_DB_InstallCo.Text);
  LMDIniCtrl1.WriteString('����','��ġ����ڸ�',ed_DB_installName.Text);
  LMDIniCtrl1.WriteString('����','����ó',      ed_DB_installTell.Text);
}
end;

Function TForm_Main.IsRegDB(aDeviceID:String):Boolean;
begin
  Result:= False;
{  with ZQuery1 do
  begin
    Close;
    SQL.Clear;
    SQL.Add('SELECT DEVICEID');
    SQL.Add('FROM TB_SETUP_SITE');
    SQL.Add('WHERE DEVICEID = :DEVICEID');
    ParamByName('DEVICEID').asString:= aDeviceID;
    Open;
    if RecordCount > 0 then Result:= True;

  end;  }
end;

Procedure TForm_Main.InsertDb;
begin
{  with ZQuery1 do
  begin
    Close;
    SQL.Clear;
    SQL.Add('INSERT INTO TB_SETUP_SITE');
    SQL.Add('(DEVICEID,DOOR_TYPE,KT_CENTER,');
    SQL.Add('KT_LOCAL,KT_SLOCAL,SITENAME,');
    SQL.Add('SYSTEMID,DECODERNO,SERVER,');
    SQL.Add('DOORLOCK,PSTN,SITE_IP,SITE_SUBNET,');
    SQL.Add('SITE_GW,INSTALL_MEMO,INSTALL_CO,');
    SQL.Add('INSTALL_NAME,INSTALL_TEL,REG_DAY)');
    SQL.Add('VALUES(:DEVICEID,:DOOR_TYPE,:KT_CENTER,');
    SQL.Add(':KT_LOCAL,:KT_SLOCAL,:SITENAME,');
    SQL.Add(':SYSTEMID,:DECODERNO,:SERVER,');
    SQL.Add(':DOORLOCK,:PSTN,:SITE_IP,:SITE_SUBNET,');
    SQL.Add(':SITE_GW,:INSTALL_MEMO,:INSTALL_CO,');
    SQL.Add(':INSTALL_NAME,:INSTALL_TEL,:REG_DAY)');
    ParamByName('DEVICEID').asString:=        ed_DB_Deviceid.Text;
    ParamByName('DOOR_TYPE').asString:=       InttoStr(ComboBox_DoorType1.ItemIndex);
    ParamByName('KT_CENTER').asString:=       cb_DB_Center.Text;
    ParamByName('KT_LOCAL').asString:=        ed_DB_Local.Text;
    ParamByName('KT_SLOCAL').asString:=       ed_DB_sLocal.Text;
    ParamByName('SITENAME').asString:=        ed_DB_siteName.Text;
    ParamByName('SYSTEMID').asString:=        Edit_LinkusID.Text;
    ParamByName('DECODERNO').asString:=       Edit_LinkusTel.Text;
    ParamByName('SERVER').asString:=          Edit_ServerIp.Text+':'+Edit_Serverport.Text;
    ParamByName('DOORLOCK').asString:=        ComboBox_LockType.Text[3];
    ParamByName('PSTN').asString:=            ed_DB_pstn.Text;
    ParamByName('SITE_IP').asString:=         ed_DB_ip.Text;
    ParamByName('SITE_SUBNET').asString:=     ed_DB_subnet.Text;
    ParamByName('SITE_GW').asString:=         ed_DB_gateway.Text;
    ParamByName('INSTALL_MEMO').asString:=    ed_DB_memo.Text;
    ParamByName('INSTALL_CO').asString:=      ed_DB_InstallCo.Text;
    ParamByName('INSTALL_NAME').asString:=    ed_DB_installName.Text;
    ParamByName('INSTALL_TEL').asString:=     ed_DB_installTell.Text;
    ParamByName('REG_DAY').asDatetime:=         Now;
    try
      ExecSQL;
    Except
      on E : EDatabaseError do
      begin
        Application.MessageBox (PChar(E.Message),PChar('Ȯ��'),MB_ICONSTOP or MB_OK);
      end else
      begin
        Application.MessageBox (PChar('DB ���� ����!!!'),PChar('Ȯ��'),MB_ICONSTOP or MB_OK);
      end;
      Screen.Cursor:= crDefault;
      Exit;
    end;
  end;
  SHowMessage('��ġ����� �Ϸ� �Ǿ����ϴ�.'+#13+'���� �ϼ̽��ϴ�.');  }
end;

Procedure TForm_Main.UpDateDb;
begin
{
  with ZQuery1 do
  begin
    Close;
    SQL.Clear;
    SQL.Add('UPDATE TB_SETUP_SITE');
    SQL.Add('SET DOOR_TYPE=:DOOR_TYPE ,KT_CENTER=:KT_CENTER ,');
    SQL.Add('KT_LOCAL=:KT_LOCAL,KT_SLOCAL=:KT_SLOCAL,SITENAME=:SITENAME,');
    SQL.Add('SYSTEMID=:SYSTEMID,DECODERNO=:DECODERNO,SERVER=:SERVER,');
    SQL.Add('DOORLOCK=:DOORLOCK,PSTN=:PSTN,SITE_IP=:SITE_IP,SITE_SUBNET=:SITE_SUBNET,');
    SQL.Add('SITE_GW=:SITE_GW,INSTALL_MEMO=:INSTALL_MEMO,INSTALL_CO=:INSTALL_CO,');
    SQL.Add('INSTALL_NAME=:INSTALL_NAME,INSTALL_TEL=:INSTALL_TEL,UPDATE_DAY=:UPDATE_DAY');
    SQL.Add('WHERE DEVICEID =:DEVICEID');

    ParamByName('DEVICEID').asString:=        ed_DB_Deviceid.Text;
    ParamByName('DOOR_TYPE').asString:=       InttoStr(ComboBox_DoorType1.ItemIndex);
    ParamByName('KT_CENTER').asString:=       cb_DB_Center.Text;
    ParamByName('KT_LOCAL').asString:=        ed_DB_Local.Text;
    ParamByName('KT_SLOCAL').asString:=       ed_DB_sLocal.Text;
    ParamByName('SITENAME').asString:=        ed_DB_siteName.Text;
    ParamByName('SYSTEMID').asString:=        Edit_LinkusID.Text;
    ParamByName('DECODERNO').asString:=       Edit_LinkusTel.Text;
    ParamByName('SERVER').asString:=          Edit_ServerIp.Text+':'+Edit_Serverport.Text;
    ParamByName('DOORLOCK').asString:=        ComboBox_LockType.Text[3];
    ParamByName('PSTN').asString:=            ed_DB_pstn.Text;
    ParamByName('SITE_IP').asString:=         ed_DB_ip.Text;
    ParamByName('SITE_SUBNET').asString:=     ed_DB_subnet.Text;
    ParamByName('SITE_GW').asString:=         ed_DB_gateway.Text;
    ParamByName('INSTALL_MEMO').asString:=    ed_DB_memo.Text;
    ParamByName('INSTALL_CO').asString:=      ed_DB_InstallCo.Text;
    ParamByName('INSTALL_NAME').asString:=    ed_DB_installName.Text;
    ParamByName('INSTALL_TEL').asString:=     ed_DB_installTell.Text;
    ParamByName('UPDATE_DAY').asDatetime:=      Now;
    try
      ExecSQL;
    Except
      on E : EDatabaseError do
      begin
        Application.MessageBox (PChar(E.Message),PChar('Ȯ��'),MB_ICONSTOP or MB_OK);
      end else
      begin
        Application.MessageBox (PChar('DB ���� ����!!!'),PChar('Ȯ��'),MB_ICONSTOP or MB_OK);
      end;
      Screen.Cursor:= crDefault;
      Exit;
    end;
  end;
  SHowMessage('��ġ��� ������ �Ϸ� �Ǿ����ϴ�.'+#13+'���� �ϼ̽��ϴ�.');
}
end;



procedure TForm_Main.ed_DB_DeviceidKeyPress(Sender: TObject;
  var Key: Char);
begin
  If Key = #13 Then Begin
    If HiWord(GetKeyState(VK_SHIFT)) <> 0 then
     SelectNext(Sender as TWinControl,False,True)
    else
     SelectNext(Sender as TWinControl,True,True) ;
     Key := #0
   end;

end;

function TForm_Main.ChangeAlarmMode(aDeviceID: string; aMode: Char;
  aQuick: Boolean): Boolean;
var
  aData: string;
begin
  Result := false;
  aData := 'MC' +         //COMMAND
    '00' +                //ZONE ����
    aMode;                //A:���, D:����, P:��ȸ
  SendPacket(aDeviceID, 'R', aData);
  Result := true;

end;

function TForm_Main.CheckWiznet(aDeviceID: string): Boolean;
var
  aData: string;
  PastTime : dword;

begin
  Result := false;
  aData := 'NW00';
  bWiznetCheck := False;
  SendPacket(aDeviceID, 'Q', aData);

  PastTime := GetTickCount + DelayTime;
  while Not bWiznetCheck do
  begin
    if Not bConnected then Exit;
    Application.ProcessMessages;
    if GetTickCount > PastTime then Exit;  //300�и����� ���� ������ ���з� ó����
  end;

  Result := true;
end;

procedure TForm_Main.cb_NextClick(Sender: TObject);
begin
  if cb_Next.Checked then bNext := True
  else bNext := False;

  if bNext then btn_Next.Enabled := True
  else if Not bConnected then btn_Next.Enabled := False;
end;

procedure TForm_Main.btn_NextClick(Sender: TObject);
begin
  btn_Prv.Enabled := True;
  Screen.Cursor:= crDefault;

  if NoteBook1.PageIndex = 7 then
  begin
    if btn_Next.Caption = '����' then btn_Next.Caption := '����'
    else if  btn_Next.Caption = '����' then Close ;
  end;
  if NoteBook1.PageIndex < 7 then    NoteBook1.PageIndex := NoteBook1.PageIndex + 1;

  case NoteBook1.PageIndex of
    1:begin
//        if L_bFirmWareUpdate then
          btnIDCheckClick(self);
      end;
    2:begin
//        if L_bFirmWareUpdate then
          btnCheckwiznetClick(Self);
      end;
    3:begin
//        if L_bFirmWareUpdate then
          btnCheckTelNoClick(Self);
      end;
    4:begin
//        if L_bFirmWareUpdate then
          Btn_CheckDialInfoClick(Self);
      end;
    5:begin

      end;
    6:begin
//        if L_bFirmWareUpdate then
          CheckSysInfo2(xDeviceID);
      end;
    7:begin

      end;
    8:begin
        ed_DB_Deviceid.Text:= xDeviceID;
        ed_DB_siteName.Text:= Edit_Name.Text;

      end;
  end;

end;

procedure TForm_Main.RzBitBtn1Click(Sender: TObject);
begin
  btn_Next.Caption := '����';
  btn_Prv.Enabled := False;
  NoteBook1.PageIndex := 0;
end;

procedure TForm_Main.btn_PrvClick(Sender: TObject);
begin
  btn_Next.Caption := '����';
  if NoteBook1.PageIndex > 0 then  NoteBook1.PageIndex := NoteBook1.PageIndex - 1;
  if NoteBook1.PageIndex = 0 then btn_Prv.Enabled := False;
end;

procedure TForm_Main.SetConnected(const Value: Boolean);
begin
  FConnected := Value;
  if Value then btn_Next.Enabled := True
  else if Not cb_Next.Checked then
  begin
    if Not Value then btn_Next.Enabled := False;
  end;
end;

procedure TForm_Main.FormCreate(Sender: TObject);
begin
  L_bFirmWareUpdate := False;
  ComPortList := TStringList.Create;
end;



function TForm_Main.DecodeCommportName(PortName: String): WORD;
var
 Pt : Integer;
begin
 PortName := UpperCase(PortName);
 if (Copy(PortName, 1, 3) = 'COM') then begin
    Delete(PortName, 1, 3);
    Pt := Pos(':', PortName);
    if Pt = 0 then Result := 0
       else Result := StrToInt(Copy(PortName, 1, Pt-1));
 end
 else if (Copy(PortName, 1, 7) = '\\.\COM') then begin
    Delete(PortName, 1, 7);
    Result := StrToInt(PortName);
 end
 else Result := 0;

end;

function TForm_Main.EncodeCommportName(PortNum: WORD): String;
begin
 if PortNum < 10
    then Result := 'COM' + IntToStr(PortNum) + ':'
    else Result := '\\.\COM'+IntToStr(PortNum);

end;

function TForm_Main.GetSerialPortList(List: TStringList;
  const doOpenTest: Boolean): LongWord;
type
 TArrayPORT_INFO_1 = array[0..0] Of PORT_INFO_1;
 PArrayPORT_INFO_1 = ^TArrayPORT_INFO_1;
var
{$IF USE_ENUMPORTS_API}
 PL : PArrayPORT_INFO_1;
 TotalSize, ReturnCount : LongWord;
 Buf : String;
 CommNum : WORD;
{$IFEND}
 I : LongWord;
 CHandle : THandle;
begin
 List.Clear;
{$IF USE_ENUMPORTS_API}
 EnumPorts(nil, 1, nil, 0, TotalSize, ReturnCount);
 if TotalSize < 1 then begin
    Result := 0;
    Exit;
    end;
 GetMem(PL, TotalSize);
 EnumPorts(nil, 1, PL, TotalSize, TotalSize, Result);

 if Result < 1 then begin
    FreeMem(PL);
    Exit;
    end;

 for I:=0 to Result-1 do begin
    Buf := UpperCase(PL^[I].pName);
    CommNum := DecodeCommportName(PL^[I].pName);
    if CommNum = 0 then Continue;
    List.AddObject(EncodeCommportName(CommNum), Pointer(CommNum));
    end;
{$ELSE}
 for I:=1 to MAX_COMPORT do List.AddObject(EncodeCommportName(I), Pointer(I));
{$IFEND}
 // Open Test
 if List.Count > 0 then for I := List.Count-1 downto 0 do begin
    CHandle := CreateFile(PChar(List[I]), GENERIC_WRITE or GENERIC_READ,
     0, nil, OPEN_EXISTING,
     FILE_ATTRIBUTE_NORMAL,
     0);
    if CHandle = INVALID_HANDLE_VALUE then begin
if doOpenTest or (GetLastError() <> ERROR_ACCESS_DENIED) then List.Delete(I);
Continue;
end;
    CloseHandle(CHandle);
    end;

 Result := List.Count;
{$IF USE_ENUMPORTS_API}
 if Assigned(PL) then FreeMem(PL);
{$IFEND}

end;

end.

//           1         2         3
//0123456789012345678901234567890123456
//E811006050112342300CCO100>217=:15
//E911006050112342300CAO000>217=:15


3	DEVICEID
0	DOOR_TYPE
0	KT_CENTER
0	KT_LOCAL
0	KT_SLOCAL
0	SITENAME
0	SYSTEMID
0	DECODERNO
0	SERVER
0	DOORLOCK
0	PSTN
0	SITE_IP
0	SITE_SUBNET
0	SITE_GW
0	INSTALL_MEMO
0	INSTALL_CO
0	INSTALL_NAME
0	INSTALL_TEL
0	REG_DAY
0	UPDATE_DAY
