{*******************************************************}
{                                                       }
{         ȭ�ϸ�: uLomosUtil.pas                        }
{         ��  ��: ��ƿ��Ƽ                              }
{         �ۼ���:                                       }
{         �ۼ���: ������                                }
{         Copyright                                     }
{                                                       }
{*******************************************************}
//OpenPictureDialog1.InitialDir:= ExtractFileDir(Application.ExeName);
unit uLomosUtil;

interface
uses
  shellapi,
  ComCtrls,
  winsock,
  DBTables,
  Windows,
  SysUtils,
  Controls,
  Classes,
  Graphics,
  Forms,
  DB,
  DBGrids,
  iniFiles,
  DIMime,
  DateUtils;

const
  STX = #$2;
  ETX = #$3;
  ENQ = #$5;
  ACK = #$6;
  NAK = #$15;
  EOT = #$04;
  TAB = #$09;
  HexString : String = '0123456789ABCDEF';

const
  MB_TIMEDOUT = 32000;


type CString = string[100];

 TIsHungAppWindow = function(Wnd: HWND): Bool; stdcall;
 TIsHungThread = function(lpThreadId: DWORD): Bool; stdcall;

  function Ascii2Hex(aData:string;bReverse:Boolean = False):string;
  function GetWeekCode(aDate:string):integer;
  function Dec2Bin(Value : LongInt) : string;
  function Bin2Dec(BinString: string): LongInt;
  function Dec2Hex(N: LongInt; A: Byte): string;
  function Dec2Hex64(N: int64; A: Byte): string;
  procedure SetBitB(var b:byte; BittoSet: integer);
  procedure ClearBitB(var b:byte; BitToClear: integer);
  function IsBitSetB(const b:byte; BitToCheck:integer):Boolean;
  Function MakeCSData(aData: string):String;
  function MakeSum(st:string):Char;
  Function DataConvert1(aMakeValue:Byte;aData:String):String;
  Function DataConvert2(aMakeValue:Byte;aData:String):String;
  function EncodeData(aKey:Byte; aData: String): String;

  function ToHexStr(st:string):String;
  function ToHexStrNoSpace(st:string):String;
  function Hex2Ascii(St: String): String;
  Function Hex2DecStr(S:String):String;
  function Hex2Dec(const S: string): Longint;
  function Hex2Dec64(const S: string): int64;
  function Isdigit(st: string):Boolean;
  function BinToInt(Value : String) : Integer;
  function SetlengthStr(st : String; aLength : Integer): String;
  function FillZeroStrNum(aNo:String; aLength:Integer;bFront:Boolean = True): string;
  function FillZeroNumber(aNo:LongInt; aLength:Integer): string;
  function FillZeroNumber2(aNo:Int64; aLength:Integer): string;
  function IntToBin(Value: Longint; Digits:Integer): string;
  Function FindCharCopy(SourceStr : String; Index : integer; aChar:Char) : String;
  procedure ErrorLogSave(aFileName: String;aError:String;ast:string);
  Function MakeDatetimeStr(aTime: String):String;
  function IncTime(ATime: TDateTime; Hours, Minutes, Seconds,
    MSecs: Integer): TDateTime;
  function DecTime(ATime: TDateTime; Hours, Minutes, Seconds,
    MSecs: Integer): TDateTime;
  procedure LogSave(aFileName,ast:string);
  procedure FileAppend(aFileName,ast:string);
  Function DeleteChar(st : String; DelChar : Char) : String;
  procedure Delay(MSecs: Longint);
  procedure GridtoFile(adbGrid: TDBGrid; aFileName: string);
  function DelChars(const S: string; Chr: Char): string;
  function ActivateScreenSaver( Activate : Boolean ) : Boolean;
  function DecodeCardNo(aCardNo: string): String;
  function EncodeCardNo(aCardNo: string): String;
  function Get_Local_IPAddr : string;
  function GetNodeByText(ATree : TTreeView; AValue:String; AVisible: Boolean): TTreeNode;
  procedure Snooze(ms: Cardinal);
  function ExecFileAndWait(const aCmdLine: String; Hidden, doWait: Boolean): Boolean;
  Procedure ShellExecute_AndWait(FileName:String;Params:String);
  function IsHungWindow(Wnd: HWND): Boolean;
  function MakeXOR(st:string):Char;

  function  ufFindQuery(sQuery : TQuery; sField, sData : String; iClassify : Integer) : Boolean; OverLoad;
  function  ufFindQuery(sQuery : TQuery; sField1, sField2, sData1, sData2 : String; iClassify1, iClassify2 : Integer) : Boolean; OverLoad;
  function  ufFindQuery(sQuery : TQuery; sField1, sField2, sField3, sData1, sData2, sData3 : String; iClassify1, iClassify2, iClassify3 : Integer) : Boolean; OverLoad;
  procedure TerminateApplication;
  procedure InverseString(var S:string;Count:Integer);
  procedure CDLogSave(aFileName: String;ast:string);
  function IsSunday(const D: TDateTime): Boolean;

  Function DateCheck(aDate:string):Boolean;

  function StrToBin(const S: string): string;
  function RightPos(aDelimiter : Char; S:string):integer;
  function FillSpace(aData:string;aLen:integer;bFront:Boolean = False):string;

  function IsDate(aDate:string):Boolean;
  Function StringToBin(aData:string):string;
  Function BinToHexC(aData:string):string;
  //����Ű�� ��������
  Function GetAuthKey(aGubun:string) : string;
  //����ڵ�Ϲ�ȣ�� ��������
  Function GetSaupId(aGubun:string) : string;
  //����Ű���� ������
  Function CompareKey(aSaupId,aKey:string):Boolean;
  //��ġ���ڸ� ��������
  Function GetInstallDate(aGubun:string):string;

  procedure My_RunDosCommand(Command : string;  nShow : Boolean = False; bWait:Boolean = True);
  function MyF_UsingWinNT: Boolean;
  function CheckSumCheck(aBuff:string):Boolean;

  Function WonStringFormat(aAmt:string):string;
  Function MinuteToString(aMinute:integer):string;



{MessageBoxTimeout:����
procedure TForm1.Button1Click(Sender: TObject) ;
var
  iRet: Integer;
  iFlags: Integer;
begin
  iFlags := MB_OK or MB_SETFOREGROUND or MB_SYSTEMMODAL or MB_ICONINFORMATION;
  MessageBoxTimeout(Application.Handle, 'Test a timeout of 2 seconds.', 'MessageBoxTimeout Test', iFlags, 0, 2000) ;

  iFlags := MB_YESNO or MB_SETFOREGROUND or MB_SYSTEMMODAL or MB_ICONINFORMATION;
  iRet := MessageBoxTimeout(Application.Handle, 'Test a timeout of 5 seconds.', 'MessageBoxTimeout Test', iFlags, 0, 5000) ;
  case iRet of
    IDYES:
      ShowMessage('Yes') ;
    IDNO:
      ShowMessage('No') ;
    MB_TIMEDOUT:
      ShowMessage('TimedOut') ;
  end;
end;
 }

  function MessageBoxTimeOut(hWnd: HWND; lpText: PChar; lpCaption: PChar; uType: UINT; wLanguageId: WORD; dwMilliseconds: DWORD): Integer; stdcall; external user32 name 'MessageBoxTimeoutA';
function GetMacAddress: string;

Implementation


function CoCreateGuid(var guid: TGUID): HResult; stdcall; far external 'ole32.dll';
function GetMACAddress:String;
var
UuidCreateFunc : function (var guid: TGUID):HResult;stdcall;
handle : THandle;
g:TGUID;
WinVer:_OSVersionInfoA;
i:integer;
begin
WinVer.dwOSVersionInfoSize := sizeof(WinVer);
getversionex(WinVer);

handle := LoadLibrary('RPCRT4.DLL');
if WinVer.dwMajorVersion >= 5 then {Windows 2000 }
@UuidCreateFunc := GetProcAddress(Handle, 'UuidCreateSequential')
else
@UuidCreateFunc := GetProcAddress(Handle, 'UuidCreate') ;

UuidCreateFunc(g);
result:='';
for i:=2 to 7 do
result:=result+IntToHex(g.d4[i],2);
end;

{function GetMacAddress: string;
var
g: TGUID;
i: Byte;
begin
Result := '';
CoCreateGUID(g);
for i := 2 to 7 do
Result := Result + IntToHex(g.D4[i], 2);
end;  }

function Ascii2Hex(aData:string;bReverse:Boolean = False):string;
var
  i : integer;
  stHex : string;
begin
  stHex := '';
  for i:= 1 to Length(aData) do
  begin
    if Not bReverse then stHex := stHex + Dec2Hex(Ord(aData[i]),2)
    else stHex := Dec2Hex(Ord(aData[i]),2) + stHex;
  end;
  result := stHex;
end;
Function MinuteToString(aMinute:integer):string;
var
  nDD : integer;
  nHH : integer;
  nMM : integer;
begin
  nHH := aMinute div 60;
  nMM := aMinute mod 60; //��...
  if nHH = 0 then
  begin
    result := inttostr(nMM) + '��';
    Exit;
  end;
  nDD := nHH div 60;
  nHH := nHH mod 60; //�ð�
  if nDD = 0 then
  begin
    result := inttostr(nHH) + '�ð� ' + inttostr(nMM) + '��';
    Exit;
  end;
  result := inttostr(nDD) + '�� ' +inttostr(nHH) + '�ð� ' + inttostr(nMM) + '��';

end;

Function WonStringFormat(aAmt:string):string;
begin
  if aAmt = '' then
    aAmt := '0'
  else begin
    aAmt := formatfloat('###,###,##0',StrToint(StringReplace(aAmt,',','',[rfReplaceAll])));
  end;
  result :=aAmt;
end;


function IsSunday(const D: TDateTime): Boolean;
begin
  Result := DayOfWeek(D) = 1;
end;

function GetWeekCode(aDate:string):integer;
var
  dtPresent: TDateTime;
  wYear    : word;
  wMonth   : word;
  wDay     : word;
begin
  wYear  := StrtoInt(Copy(aDate,1,4));
  wMonth := StrtoInt(Copy(aDate,5,2));
  wDay   := StrtoInt(Copy(aDate,7,2));
  dtPresent:= EncodeDatetime(wYear, wMonth, wDay, 00, 00, 00,00);

  Result := DayOfWeek(dtPresent);
end;


procedure InverseString(var S:string;Count:Integer);
var
   TmpStr:string;
   Ctr:Integer;
   Ch:Char;
begin
   TmpStr:=Copy(S,1,Count);
   Ctr:=0;
   while Count>0 do begin
      Ch:=TmpStr[Count];
      Dec(Count);
      Move(Ch,S[Ctr+1],1);
      Inc(Ctr);
   end;
end;


procedure TerminateApplication;
begin
  with Application do begin
    ShowMainForm := False;
    if Handle <> 0 then ShowOwnedPopups(Handle, False);
    Terminate;
  end;
  CallTerminateProcs;
  Halt(10);
end;


{
 ���뼳�� : DB-Grid�� �ش� �������� ��Ŀ���� �̵��Ѵ�.
          : ã�� �������� �ʵ��� ���� �ַ� PK�� �´�.(PK�� 1���� ���)
 �Ķ���� : sQuery : DataSource�� ����Ǿ� �ִ� ����
            sField : ã�� �ʵ��
            sData : ã�� ������ ��(Value)
            iClassify : �ʵ�Ÿ��(1:����, 2:����)
 ��뿹�� : ufFindQuery(qryList, 'Bank_CD', edtBank_CD.Text, 1)
 �� �� �� : True/False

}
function ufFindQuery(sQuery : TQuery; sField, sData : String; iClassify : Integer) : Boolean; OverLoad;
begin
  with sQuery do
  begin
    case iClassify of
      1 : Filter := sField + ' = ' + #39 + sData + #39;
      2 : Filter := sField + ' = ' + sData;
    end;
    if FindFirst then Result := TRUE else Result := FALSE;
  end;
end;



{
 ���뼳�� : DB-Grid�� �ش� �������� ��Ŀ���� �̵��Ѵ�.                     
          : ã�� �������� �ʵ��� ���� �ַ� PK�� �´�.(PK�� 2���� ���)
 �Ķ���� : sQuery : DataSource�� ����Ǿ� �ִ� ����                       
            sField1, sField2 : ã�� �ʵ��                                 
            sData1, sData2 : ã�� ������ ��(Value)                         
            iClassify1, iClassify2 : �ʵ�Ÿ��(1:����, 2:����)              
 ��뿹�� : ufFindQuery(qryList, 'Reg_Date', 'Seq', datReg_Date.Text,      
                        vSeq, 1, 2)                                        
 �� �� �� : True/False                                                     

}
function ufFindQuery(sQuery : TQuery; sField1, sField2, sData1, sData2 : String; iClassify1, iClassify2 : Integer) : Boolean ;OverLoad;
var
  sTempData : String;
begin
  with sQuery do
  begin
    case iClassify1 of
      1 : sTempData := sField1+' = ' + #39 + sData1 + #39 + ' AND ';
      2 : sTempData := sField1+' = ' + sData1 + ' AND ';
    end;
    case iClassify2 of
      1 : sTempData := sTempData+sField2+' = ' + #39 + sData2 + #39;
      2 : sTempData := sTempData+sField2+' = ' + sData2;
    end;
    Filter := sTempData;
    if FindFirst then Result := TRUE else Result := FALSE;
  end;
end;


{

 ���뼳�� : DB-Grid�� �ش� �������� ��Ŀ���� �̵��Ѵ�.
          : ã�� �������� �ʵ��� ���� �ַ� PK�� �´�.(PK�� 3���� ���)
 �Ķ���� : sQuery : DataSource�� ����Ǿ� �ִ� ����
            sField1, sField2, sField3 : ã�� �ʵ��
            sData1, sData2, sData3 : ã�� ������ ��(Value)
            iClassify1, iClassify2, iClassify3 : �ʵ�Ÿ��(1:����, 2:����)
 ��뿹�� : ufFindQuery(qryList, 'Reg_Date', 'Seq', 'Kind',
                        datReg_Date.Text, vSeq, cmbKind.Text, 1, 2, 1)
 �� �� �� : True/False

}
function ufFindQuery(sQuery : TQuery; sField1, sField2, sField3, sData1, sData2, sData3 : String; iClassify1, iClassify2, iClassify3 : Integer) : Boolean; OverLoad;
var
  sTempData : String;
begin
  with sQuery do
  begin
    case iClassify1 of
      1 : sTempData := sField1+' = ' + #39 + sData1 + #39 + ' AND ';
      2 : sTempData := sField1+' = ' + sData1 + ' AND ';
    end;
    case iClassify2 of
      1 : sTempData := sTempData+sField2 + ' = ' + #39 + sData2 + #39 + ' AND ';
      2 : sTempData := sTempData+sField2 + ' = ' + sData2 + ' AND ';
    end;
    case iClassify3 of
      1 : sTempData := sTempData+sField3+' = ' + #39 + sData3 + #39;
      2 : sTempData := sTempData+sField3+' = ' + sData3;
    end;
    Filter := sTempData;
    if FindFirst then Result := TRUE else Result := FALSE;
  end;
end;

 
  

   



function IsHungWindow(Wnd: HWND): Boolean;
var
 IsHungAppWindow: TIsHungAppWindow;
 IsHungThread: TIsHungThread;
 DllHandle: THandle;
 ThreadID: DWORD;
begin
 Result := False;
 DllHandle := LoadLibrary('user32.dll');
 if DllHandle <> 0 then
 begin
   try
     // ������ NT �迭
     if (Win32Platform = VER_PLATFORM_WIN32_NT) then
     begin
       @IsHungAppWindow := GetProcAddress(DLLHandle, 'IsHungAppWindow');
       if Addr(IsHungAppWindow) <> nil then
         Result := IsHungAppWindow(Wnd);
     end
     // ������ 9x �迭
     else begin
       @IsHungThread := GetProcAddress(DLLHandle, 'IsHungThread');
       if Addr(IsHungThread) <> nil then
       begin
         GetWindowThreadProcessId(Wnd, @ThreadID);
         Result := IsHungThread(ThreadID);
       end;
     end;
   finally
     FreeLibrary(DllHandle);
   end;
 end;
end;

function ExecFileAndWait(const aCmdLine: String; Hidden, doWait: Boolean): Boolean;
var
StartupInfo : TStartupInfo;
ProcessInfo : TProcessInformation;
begin
{setup the startup information for the application }
  FillChar(StartupInfo, SizeOf(TStartupInfo), 0);
  with StartupInfo do
  begin
    cb:= SizeOf(TStartupInfo);
    dwFlags:= STARTF_USESHOWWINDOW or STARTF_FORCEONFEEDBACK;
    if Hidden
    then wShowWindow:= SW_HIDE
    else wShowWindow:= SW_SHOWNORMAL;
  end;

  Result := CreateProcess(nil,PChar(aCmdLine), nil, nil, False,
  NORMAL_PRIORITY_CLASS, nil, nil, StartupInfo, ProcessInfo);
  if doWait then
  if Result then
  begin
  WaitForInputIdle(ProcessInfo.hProcess, INFINITE);
  WaitForSingleObject(ProcessInfo.hProcess, INFINITE);
  end;
end;

Procedure ShellExecute_AndWait(FileName:String;Params:String);
var
  exInfo : TShellExecuteInfo;
  Ph     : DWORD;
  errmsg  : String;
begin
 FillChar( exInfo, Sizeof(exInfo), 0 );
 with exInfo do
 begin
   cbSize:= Sizeof( exInfo );
   fMask := SEE_MASK_NOCLOSEPROCESS or SEE_MASK_FLAG_DDEWAIT;
   Wnd := GetActiveWindow();
   ExInfo.lpVerb := 'open';
   ExInfo.lpParameters := PChar(Params);
   lpFile:= PChar(FileName);
   nShow := SW_SHOWNORMAL;
 end;
 if ShellExecuteEx(@exInfo) then
 begin
   Ph := exInfo.HProcess;
 end
 else
 begin
   errmsg:= SysErrorMessage(GetLastError);
   Application.MessageBox (PChar(errmsg),PChar('error'),MB_ICONSTOP or MB_OK);
   exit;
 end;

 while WaitForSingleObject(ExInfo.hProcess, 50) <> WAIT_OBJECT_0 do
 Application.ProcessMessages;
 CloseHandle(Ph);
end;

procedure Snooze(ms: Cardinal);
var
  Stop: Cardinal;
begin
  SetTimer(Application.Handle, 1235, ms, nil);
  Stop := GetTickCount + ms;
  repeat
    Application.HandleMessage;
  until Application.Terminated or (Integer(GetTickCount - Stop) >= 0);
  KillTimer(Application.Handle, 1235);
end;

function GetNodeByText(ATree : TTreeView; AValue:String; AVisible: Boolean): TTreeNode;
var
    Node: TTreeNode;
begin

  Result := nil;
  if ATree.Items.Count = 0 then Exit;
  Node := ATree.Items[0];
  while Node <> nil do
  begin
    if UpperCase(Node.Text) = UpperCase(AValue) then
    begin
      Result := Node;
      if AVisible then
        Result.MakeVisible;
      Break;
    end;
    Node := Node.GetNext;
  end;
end;

function DecodeCardNo(aCardNo: string): String;
var
  I: Integer;
  st: string;
  bCardNo: int64;
begin

  for I := 1 to 8 do
  begin

    if (I mod 2) <> 0 then
    begin
      aCardNo[I] := Char((Ord(aCardNo[I]) shl 4));
    end else
    begin
      aCardNo[I] := Char(Ord(aCardNo[I]) - $30); //�����Ϻ��� 0���� �����.
      //st:= st + char(ord(aCardNo[I-1]) +ord(aCardNo[I]));
      st:= st + char(ord(aCardNo[I-1]) + ord(aCardNo[I]))
    end;
    //aCardNo[I] := Char(Ord(aCardNo[I]) - $30);
    //st := st + aCardNo[I];
  end;


  st:= tohexstrNospace(st);


  bCardNo:= Hex2Dec64(st);
  st:= FillZeroNumber2(bCardNo,10);
  //SHowMessage(st);
  Result:= st;

end;

function EncodeCardNo(aCardNo: string): String;
var
  I: Integer;
  xCardNo: String;
  st: String;
begin
  aCardNo:= Dec2Hex64(StrtoInt64(aCardNo),8);
  xCardNo:= Hex2Ascii(aCardNo);
  for I:= 1 to 4 do
  begin
    st := st + Char((Ord(xCardNo[I]) shr 4) + $30) + Char((Ord(xCardNo[I]) and $F) + $30);
  end;
  Result:= st;
end;

function ActivateScreenSaver( Activate : Boolean ) : Boolean;
var
  IntActive : Byte;
begin
  if Activate then
     IntActive := 1
  else
     IntActive := 0;

  Result := SystemParametersInfo( SPI_SETSCREENSAVEACTIVE, IntActive, nil, 0 );
end;


function DelChars(const S: string; Chr: Char): string;
var
  I: Integer;
begin
  Result := S;
  for I := Length(Result) downto 1 do begin
    if Result[I] = Chr then Delete(Result, I, 1);
  end;
end;

{DBGrid�� CSVȭ�Ϸ� ����}
procedure GridtoFile(adbGrid: TDBGrid; aFileName: string);
var
  st: string;
  st2:string;
  CurMark: TBookmark;
  CurColumn: Integer;
  aStringList: Tstringlist;
begin

  aStringList := TStringList.Create;
  aStringList.Clear;
  //�׸��� ���� ����

  with aDBGrid.Columns do
  begin
    for CurColumn := 0 to Count - 1 do
    begin
      if (CurColumn > 0) then st := st + ', ';
      st := st + aDBGrid.Columns.Items[CurColumn].Title.Caption;
    end;
    aStringList.Add(st);
  end;
  //Title ����
  with aDBGrid.DataSource.Dataset do
  begin
    DisableControls;
    CurMark := GetBookmark; {���� ���ڵ� ������ ����}
    First;
    while not eof do
    begin
      st := '';
      for CurColumn := 0 to aDBGrid.Columns.Count - 1 do
      begin
        if (CurColumn > 0) then st := st + ', ';
        st2:=aDBGrid.Columns[CurColumn].Field.Text;
        st2:= DelChars(st2,',');
        st := st +st2 ; {�ʵ尪}
      end;
      aStringList.Add(st);
      Next;
    end;
    GotoBookmark(CurMark);
    EnableControls;
  end;

  aStringList.SaveToFile(aFileName);
  aStringList.Free;

end;

procedure Delay(MSecs: Longint);
var
  FirstTickCount, Now: Longint;
begin
  FirstTickCount := GetTickCount;
  repeat
    Application.ProcessMessages;
    { allowing access to other controls, etc. }
    Now := GetTickCount;
  until (Now - FirstTickCount >= MSecs) or (Now < FirstTickCount);
end;

Function DeleteChar(st : String; DelChar : Char) : String;
begin
  While POS(DelChar,st) <> 0 do
    st := Copy(st,1,POS(DelChar,st)-1) + Copy(st,POS(DelChar,st)+1,255);
  Result := st;
end;


procedure LogSave(aFileName,ast:string);
Var
  f: TextFile;
  st: string;
  stDir : string;
begin
  {$I-}
  stDir := ExtractFilePath(aFileName);
  if not DirectoryExists(stDir) then CreateDir(stDir);

  AssignFile(f, aFileName);
  Append(f);
  if IOResult <> 0 then Rewrite(f);
  st := FormatDateTIme('yyyy-mm-dd hh:nn:ss:zzz">"',Now) + ' ' + ast;
  WriteLn(f,st);
  System.Close(f);
  {$I+}
end;
procedure FileAppend(aFileName,ast:string);
Var
  f: TextFile;
  st: string;
  stDir : string;
begin
  {$I-}
  stDir := ExtractFilePath(aFileName);
  if not DirectoryExists(stDir) then CreateDir(stDir);

  AssignFile(f, aFileName);
  Append(f);
  if IOResult <> 0 then Rewrite(f);
  st := ast;
  WriteLn(f,st);
  System.Close(f);
  {$I+}
end;

Function MakeDatetimeStr(aTime: String):String;
begin
  MakeDatetimeStr:= Copy(aTime,1,4)+'-'+Copy(aTime,5,2)+'-'+Copy(aTime,7,2) +' '+
                    Copy(aTime,9,2)+':'+Copy(aTime,11,2)+':'+Copy(aTime,13,2);
end;

procedure ErrorLogSave(aFileName: String;aError:String;ast:string);
Var
  f: TextFile;
  st: string;
  stDir : string;
begin
  {$I-}
  stDir := ExtractFilePath(aFileName);
  if not DirectoryExists(stDir) then CreateDir(stDir);
//  aFileName:= 'c:\lomos\log\log'+ FormatDateTIme('yyyymmdd',Now)+'.log';
  AssignFile(f, aFileName);
  Append(f);
  if IOResult <> 0 then Rewrite(f);
  st := FormatDateTIme('hh:nn:ss:zzz">"',Now) + '['+aError+']' + ast;
  WriteLn(f,st);
  System.Close(f);
  {$I+}
end;

procedure CDLogSave(aFileName: String;ast:string);
Var
  f: TextFile;
  st: string;
//  aFileName: String;
  stDir : string;
begin
  {$I-}
  stDir := ExtractFilePath(aFileName);
  if not DirectoryExists(stDir) then CreateDir(stDir);
  //aFileName:= 'c:\lomos\log\CDlog'+ FormatDateTIme('yyyymmdd',Now)+'.log';
  AssignFile(f, aFileName);
  Append(f);
  if IOResult <> 0 then Rewrite(f);
  st := FormatDateTIme('hh:nn:ss:zzz">"',Now) + ' ' + ast;
  WriteLn(f,st);
  System.Close(f);
  {$I+}
end;


Function FindCharCopy(SourceStr : String; Index : integer; aChar:Char) : String;
Var
  a, b : Integer;
  st   : String;
begin
  a := 0;
  b := 1;
  st := '';
  if (Length(SourceStr) < 1) then begin result:= ''; exit;  end;
  for b:=1 to Length(SourceStr) do
  begin
    if a = index then break;
    if SourceStr[b] = aChar then Inc(a);
  end;
  if (a = Index) then
  begin
    while (b <= Length(SourceStr)) and (SourceStr[b] <> aChar) do
    begin
      st := st + SourceStr[b];
      Inc(b);
    end;
  end;
  Result := st;
end;

function IntToBin(Value: Longint; Digits:Integer): string;
begin
  Result := '';
  if Digits > 32 then Digits := 32;
  while Digits > 0 do begin
    Dec(Digits);
    Result := Result + IntToStr((Value shr Digits) and 1);
  end;
end;

function FillZeroStrNum(aNo:String; aLength:Integer;bFront:Boolean = True): string;
var
  I       : Integer;
  st      : string;
  strNo   : String;
  StrCount: Integer;
begin
  Strno:= aNo;
  StrCount:= Length(Strno);
  St:= '';
  StrCount:=  aLength - StrCount;
  if StrCount > 0 then
  begin
    st:='';
    for I:=1 to StrCount do St:=st+'0';
    if bFront then  St:= St + StrNo
    else St:= StrNo + St;
    
    FillZeroStrNum:= st;
  end else FillZeroStrNum:= copy(Strno,1,aLength);
end;

function FillZeroNumber(aNo:LongInt; aLength:Integer): string;
var
  I       : Integer;
  st      : string;
  strNo   : String;
  StrCount: Integer;
begin
  Strno:= InttoStr(aNo);
  StrCount:= Length(Strno);
  St:= '';
  StrCount:=  aLength - StrCount;
  if StrCount > 0 then
  begin
    st:='';
    for I:=1 to StrCount do St:=st+'0';
    St:= St + StrNo;
    FillZeroNumber:= st;
  end else FillZeroNumber:= copy(Strno,1,aLength);
end;

function FillZeroNumber2(aNo:INt64; aLength:Integer): string;
var
  I       : Integer;
  st      : string;
  strNo   : String;
  StrCount: Integer;
begin
  Strno:= InttoStr(aNo);
  StrCount:= Length(Strno);
  St:= '';
  StrCount:=  aLength - StrCount;
  if StrCount > 0 then
  begin
    st:='';
    for I:=1 to StrCount do St:=st+'0';
    St:= St + StrNo;
    FillZeroNumber2:= st;
  end else FillZeroNumber2:= copy(Strno,1,aLength);
end;






function SetlengthStr(st : String; aLength : Integer) : String;
begin
  result := st;
  while Length(Result) < aLength do
    Result := Result + ' ';
  Result := Copy(Result,1,aLength);
end;

function BinToInt(Value : String) : Integer;
var
  i   : Integer;
begin
  Result := 0;
  for i := 1 to Length(Value) do
  begin
    Result := Result shl 1;
    Result := Result + Integer((Value[i] = '1'));
  end;
end;


function Dec2Bin(Value : LongInt) : string;
var  
  i : integer;   
  s : string;   
begin  
  s := '';   
  
  for i := 31 downto 0 do  
    if (Value and (1 shl i)) <> 0 then s := s + '1'  
                                  else s := s + '0';   
  
  Result := s;   
end;   


function Bin2Dec(BinString: string): LongInt;
var  
  i : Integer;   
  Num : LongInt;   
begin  
  Num := 0;   
  
  for i := 1 to Length(BinString) do  
    if BinString[i] = '1' then Num := (Num shl 1) + 1  
                          else Num := (Num shl 1);   
  
  Result := Num;   
end;

function Dec2Hex(N: LongInt; A: Byte): string;
begin
  Result := IntToHex(N, A);
end;

function Dec2Hex64(N: int64; A: Byte): string;
begin
    Result := IntToHex(N, A);
end;

Function Hex2DecStr(S:String):String;
var
  i: longint;
  L: int64;
begin
  Result := '';
  if Length(s) = 0 then Exit;
  L:=0;
  for i := 1 to length(S) do L:=L*16 + pos(S[i],HexString)-1;
  Result:=intToStr(L);
end;


procedure SetBitB(var b:byte; BittoSet: integer);
{ set a bit in a byte }
begin
  if (BittoSet < 0) or (BittoSet > 7) then exit;
  b:= b or ( 1 SHL BittoSet);
end;

procedure ClearBitB(var b:byte; BitToClear: integer);
{ clear a bit in a byte }
begin
  if (BitToClear < 0) or (BitToClear > 7) then exit;
  b := b and not (1 shl BitToClear);
end;

function IsBitSetB(const b:byte; BitToCheck:integer):Boolean;
{ Test bit in a byte }
begin
  Result := false;
  if (BitToCheck < 0) or (BitToCheck > 7) then exit;
  Result := (b and (1 shl BitToCheck)) <> 0;
end;

function MakeSum(st:string):Char;
var
  i: Integer;
  aBcc: Byte;
  BCC: string;
begin
  Result:= #0;
  if st = '' then Exit;
  aBcc := Ord(st[1]);
  for i := 2 to Length(st) do
    aBcc := aBcc + Ord(st[i]);
  BCC := Chr(aBcc);
  Result := BCC[1];
end;


function MakeXOR(st:string):Char;
var
  i: Integer;
  aBcc: Byte;
  BCC: string;
begin
  Result:= #0;
  if st = '' then Exit;
  aBcc := Ord(st[1]);
  for i := 2 to Length(st) do
    aBcc := aBcc XOR Ord(st[i]);
  BCC := Chr(aBcc);
  Result := BCC[1];
end;

{CheckSum�� �����}
Function MakeCSData(aData: string):String;
var
  aSum: Integer;
  st: string;
begin
  aSum:= Ord(MakeSum(aData));
  aSum:= aSum*(-1);
  st:= Dec2Hex(aSum,2);

  Result:= copy(st,Length(st)-1,2);
end;

{���� ��ȣ��(BIT4,BIT3,BIT2,BIT1,BIT0) �� data�� XOR �Ѵ�.}
Function DataConvert1(aMakeValue:Byte;aData:String):String;
var
  I: Integer;
  bData: String;
begin
  bData:= aData;
  for I:= 1 to Length(bData) do
  begin
    bData[I]:= Char(ord(bData[I]) XOR aMakeValue);
  end;
  Result:= bData;
end;

{ ���� ��ȣ��(BIT4,BIT3,BIT2,BIT1,BIT0) �� data�� XOR �� Message No�� ���� Nibble�� �ٽ� XOR �Ѵ�.}
Function DataConvert2(aMakeValue:Byte;aData:String):String;
var
  I: Integer;
  bMakeValue: Byte;
  bData: String;
  TempByte: Byte;
begin
  bData:= aData;
  {13���� Byte �� MessageNo}
  bMakeValue:= Ord(aData[13]) and $F;
  Result:= '';
  for I:= 1 to Length(bData) do
  begin
    if I <> 13 then
    begin
      TempByte:= ord(bData[I]) XOR aMakeValue;
      bData[I]:= Char(TempByte XOR bMakeValue);
    end;
  end;
  Result:= bData;
end;

function EncodeData(aKey:Byte; aData: String): String;
var
  Encodetype: Integer;
  aMakeValue: Byte;
  I: Integer;
begin
  EncodeType:= aKey SHR 6; //7,6 �� Bit�� ���ڵ� Ÿ��
  aMakeValue:= aKey;
  for I:= 5 to 7 do ClearBitB(aMakeValue,I); //1,2,3,4,5 Bit�� ������ȣ

  case EncodeType of
    0: Result:= DataConvert1(aMakeValue,aData);
    1: Result:= DataConvert2(aMakeValue,aData);
    else Result:= aData;
  end;
end;


function ToHexStr(st:string):String;
var
  I : Integer;
  st2: string;
  st3: string[3];
begin
  for I:= 1 to length(st) do
  begin
    st3:= Dec2Hex(ord(st[I]),1);
    if Length(st3) < 2 then st3:= '0'+ st3;
    st2:=st2 +st3 +' ';
  end;
  ToHexStr:= st2;
end;

function ToHexStrNoSpace(st:string):String;
var
  I : Integer;
  st2: string;
  st3: string[3];
begin
  for I:= 1 to length(st) do
  begin
    st3:= Dec2Hex(ord(st[I]),1);
    if Length(st3) < 2 then st3:= '0'+ st3;
    st2:=st2 +st3;
  end;
  ToHexStrnospace:= st2;
end;


function Hex2Ascii(St: String): String;
var
  st2: string;
  I: Integer;
  aLength: Integer;
  aa: Integer;
begin
  st2:= '';
  for I:= 1 to Length(st) do
  begin
    if st[i] <> #$20 then st2:= st2 + st[I];
  end;
  if Length(st2) MOD 2 <> 0 then
  begin
    aLength:= Length(st2);
    st:= copy(st2,1,aLength-1) + '0'+ st2[aLength];
  end else
  begin
   st:= st2;
  end;

  st2:= '';
  while st <> '' do
  begin
    aa:= Hex2Dec(copy(st,1,2));
    st2:= st2 + Char(aa);
    delete(st,1,2);
  end;
  Hex2Ascii:= st2;
end;


function Hex2Dec(const S: string): Longint;
var
  HexStr: string;
begin
  if Pos('$', S) = 0 then HexStr := '$' + S
  else HexStr := S;
  Result := StrToIntDef(HexStr, 0);
end;

function Hex2Dec64(const S: string): int64;
var
  HexStr: string;
begin
  if Pos('$', S) = 0 then HexStr := '$' + S
  else HexStr := S;
  Result := StrToInt64Def(HexStr, 0);
end;

function Isdigit(st: string):Boolean;
var
  I: Integer;
begin
  result:=True;
  if Length(st) < 1 then
  begin
    result:=False;
    Exit;
  end;
  for I:=1 to Length(st) do
    if (st[I]< '0') or (st[I] > '9')  then result:=False
end;

function GetNibble(aValue: Byte; Var NibbleHi:Byte; Var NibbleLo:Byte):Boolean;
begin
  NibbleHi := aValue shr 4;
  NibbleLo := aValue and $F;
  Result:= True;
end;

function IncTime(ATime: TDateTime; Hours, Minutes, Seconds,
  MSecs: Integer): TDateTime;
begin
  Result := ATime + (Hours div 24) + (((Hours mod 24) * 3600000 +
    Minutes * 60000 + Seconds * 1000 + MSecs) / MSecsPerDay);
  if Result < 0 then Result := Result + 1;
end;

function DecTime(ATime: TDateTime; Hours, Minutes, Seconds,
  MSecs: Integer): TDateTime;
begin
  Result := ATime - (Hours div 24) - (((Hours mod 24) * 3600000 -
    Minutes * 60000 - Seconds * 1000 + MSecs) / MSecsPerDay);
  if Result < 0 then Result := Result + 1;
end;

function Get_Local_IPAddr : string;
 type
   TaPInAddr = array [0..10] of PInAddr;
   PaPInAddr = ^TaPInAddr;
 var
   phe : PHostEnt;
   pptr : PaPInAddr;
   Buffer : array [0..63] of char;
   I : Integer;
   GInitData : TWSADATA;
begin
 try
   WSAStartup($101, GInitData);
   Result := '';
   GetHostName(Buffer, SizeOf(Buffer));
   phe := GetHostByName(buffer);
   if phe = nil then Exit;
   pptr := PaPInAddr(Phe^.h_addr_list);
   i := 0;
   result := '';
   while pptr^[I] <> nil do
   begin
     result:= result + StrPas(inet_ntoa(pptr^[I]^)) + ' ';
     Inc(I);
   end;
 finally WSACleanup; end;
end;


Function DateCheck(aDate:string):Boolean;
var
  CheckDate : TDateTime;
begin

  Result := True;
  try
    CheckDate := StrtoDate(aDate);
    Result := True;
  except
    Result := False;
  end;

end;

function StrToBin(const S: string): string;
const
   BitArray: array[0..15] of string =
       ('0000', '0001', '0010', '0011', 
        '0100', '0101', '0110', '0111',
        '1000', '1001', '1010', '1011', 
        '1100', '1101', '1110', '1111');
var
   Index: Integer;
   LoBits: Byte;
   HiBits: Byte;
begin
   Result := '';
   for Index := 1 to Length(S) do
   begin
       HiBits := (Byte( S[Index]) and $F0) shr 4;
       LoBits := Byte( S[Index]) and $0F;
       Result := Result + BitArray[HiBits];
       Result := Result + BitArray[LoBits];
   end;
end;

function RightPos(aDelimiter : Char; S:string):integer;
var
  nLen:integer;
  nPos : integer;
  i:integer;
begin
  nLen := Length(S);
  nPos := 0;
  for i := nLen downto 0 do
  begin
    if S[i] = aDelimiter then
    begin
      nPos := i;
      Break;
    end;
  end;
  result := nPos;
end;

function FillSpace(aData:string;aLen:integer;bFront:Boolean = False):string;
var
  i:integer;
begin
  if Length(aData)>= aLen then
  begin
    result := copy(aData,1,aLen);
    Exit;
  end;
  for i:= Length(aData) to aLen do
  begin
    if bFront then aData := ' ' + aData
    else  aData := aData + ' ';
  end;
  result := copy(aData,1,aLen);

end;

function IsDate(aDate:string):Boolean;
var
  dtTime : TDateTime;
begin
  result := False;
  Try
    dtTime := strtoDate(aDate);
  Except
    Exit;
  End;
  result := True;
end;

Function StringToBin(aData:string):string;
var
  i:integer;
  stTemp : string;
begin

  stTemp := '';
  for i:=1 to Length(aData) do
  begin
    stTemp := stTemp + IntToBin(strtoint(aData[i]),4);
  end;
  result := stTemp;
end;

Function BinToHexC(aData:string):string;
var
  stTemp : string;
  nTemp : integer;
  i : integer;
  stResult : string;
begin
  if (Length(aData) mod 4) <> 0 then Exit;

  stResult := '';
  for i:= 0 to (Length(aData) div 4) -1 do
  begin
    stTemp := copy(aData,(i * 4) + 1,4);
    nTemp := BinToInt(stTemp);
    stResult := stResult + Dec2Hex(nTemp,1);
  end;
  result := stResult;
end;

//����Ű�� ��������
Function GetAuthKey(aGubun:string) : string;
var
  stExeFolder : string;
  ini_fun : TiniFile;
begin
  stExeFolder := ExtractFileDir(Application.ExeName);
  ini_fun := TiniFile.Create(stExeFolder + '\' + 'Key.ini');
  result := ini_fun.ReadString('���������','����Ű' + aGubun,'');
end;

//����ڵ�Ϲ�ȣ�� ��������
Function GetSaupId(aGubun:string) : string;
var
  stExeFolder : string;
  ini_fun : TiniFile;
begin
  stExeFolder := ExtractFileDir(Application.ExeName);
  ini_fun := TiniFile.Create(stExeFolder + '\' + 'Key.ini');
  result := ini_fun.ReadString('���������','����ڵ�Ϲ�ȣ' + aGubun,'');
end;
//����Ű���� ������
Function CompareKey(aSaupId,aKey:string):Boolean;
var
  nSaupID : integer;
  stMac : string;
  nMac : integer;
  stTemp :string;
  i : integer;
  stKey : string;
begin
  result := False;
  nSaupId := strtoint(aSaupId);
  stMac := GetMacAddress;
  nMac := 0;
  for i:=1 to Length(stMac) do
  begin
    nMac := nMac + Ord(stMac[i]);
  end;
  nSaupId := nSaupId + nMac;
  stTemp := inttostr(nSaupId);
  stTemp := StringToBin(stTemp);
  stTemp := copy(stTemp,21,Length(stTemp)-20) +copy(stTemp,11,10) + copy(stTemp,1,10);
  stTemp := BinToHexC(stTemp);
  stKey := copy(stTemp,1,3) + '-' + copy(stTemp,4,4) + '-' + copy(stTemp,8,3);
  if stKey = aKey then Result := True;
end;
//��ġ���ڸ� ��������
Function GetInstallDate(aGubun:string):string;
var
  stDate : string;
  stExeFolder : string;
  ini_fun : TiniFile;
begin
  result := '';
  stExeFolder := ExtractFileDir(Application.ExeName);
  ini_fun := TiniFile.Create(stExeFolder + '\' + 'Key.ini');
  stDate := ini_fun.ReadString('���������','��ġ��'+ aGubun,'');
  if stDate = '' then exit;
  result := MimeDecodeString(stDate);
end;


// ���� ��� ���� �Լ�/���ν���

function MyF_UsingWinNT: Boolean;
var
  OS: TOSVersionInfo; 
begin 
  OS.dwOSVersionInfoSize := Sizeof(OS); 
  GetVersionEx(OS);
  if OS.dwPlatformId = VER_PLATFORM_WIN32_NT then Result:= True 
  else Result:= False; 
end; 


// ���� ��� ���� ���ν��� 
procedure My_RunDosCommand(Command : string;  nShow : Boolean = False; bWait:Boolean = True);
var 
  hReadPipe : THandle; 
  hWritePipe : THandle; 
  SI : TStartUpInfo; 
  PI : TProcessInformation; 
  SA : TSecurityAttributes; 
  SD : TSecurityDescriptor; 
  BytesRead : DWORD; 
  Dest : array[0..1023] of char; 
  CmdLine : array[0..512] of char;
  TmpList : TStringList; 
  S, Param : string; 
  Avail, ExitCode, wrResult : DWORD; 
begin
  if MyF_UsingWinNT then begin
    InitializeSecurityDescriptor(@SD, SECURITY_DESCRIPTOR_REVISION); 
    SetSecurityDescriptorDacl(@SD, True, nil, False); 
    SA.nLength := SizeOf(SA); 
    SA.lpSecurityDescriptor := @SD; 
    SA.bInheritHandle := True; 
    Createpipe(hReadPipe, hWritePipe, @SA, 1024); 
  end else begin 
    Createpipe(hReadPipe, hWritePipe, nil, 1024); 
  end; 
  try
     //Screen.Cursor := crHourglass; 
     FillChar(SI, SizeOf(SI), 0); 
     SI.cb := SizeOf(TStartUpInfo); 
     if nShow then begin 
       SI.wShowWindow := SW_SHOWNORMAL 
     end else begin 
       SI.wShowWindow := SW_HIDE; 
     end;
     SI.dwFlags := STARTF_USESHOWWINDOW;
     SI.dwFlags := SI.dwFlags or STARTF_USESTDHANDLES; 
     SI.hStdOutput := hWritePipe; 
     SI.hStdError := hWritePipe; 
     StrPCopy(CmdLine, Command); 
     //if CreateProcess(nil,CmdLine , nil, nil, True, NORMAL_PRIORITY_CLASS, nil, nil, SI, PI) then begin
     if CreateProcess(nil,pchar(Command) , nil, nil, True,  DETACHED_PROCESS, nil, nil, SI, PI) then begin
       if bWait then
       begin
         ExitCode := 0;
         while ExitCode = 0 do begin
           wrResult := WaitForSingleObject(PI.hProcess, 50);
           if PeekNamedPipe(hReadPipe, nil, 0, nil, @Avail, nil) then begin 
             if Avail > 0 then begin 
               TmpList := TStringList.Create; 
               try 
                 FillChar(Dest, SizeOf(Dest), 0); 
                 ReadFile(hReadPipe, Dest, Avail, BytesRead, nil); 
               finally
                 TmpList.Free; 
               end; 
             end; 
           end;
           if wrResult <> WAIT_TIMEOUT then begin
             ExitCode := 1;
           end; 
           Application.ProcessMessages;
         end; 
         GetExitCodeProcess(PI.hProcess, ExitCode); 
         CloseHandle(PI.hProcess); 
         CloseHandle(PI.hThread);
       end;
     end; 
  finally 
     CloseHandle(hReadPipe); 
     CloseHandle(hWritePipe); 
     Screen.Cursor := crDefault; 
  end;
end;

function CheckSumCheck(aBuff:string):Boolean;
var
  ACKStr: String;
  stCheckSum : string;
begin
  result := False;
  ACKStr := copy(aBuff,1,Length(aBuff) - 3);
  stCheckSum := copy(aBuff,Length(aBuff) - 2,2);
  if stCheckSum = MakeCSData(ACKStr+ETX) then result := True;
end;
end.




//Application.MessageBox('�ùٸ� ��¥�� �ƴմϴ�.', '�Է� Ȯ��', MB_OK + MB_ICONWARNING);




{Select All DbGrid ==============================>
var
  TempBookmark: TBookmark;
begin
    ...
    with Dataset do
    begin
      if (BOF and EOF) then Exit;
      DisableControls;
      try
        TempBookmark:= GetBookmark;
        try
          First;
          while not EOF do
          begin
            DBGrid1.SelectedRows.CurrentRowSelected := True;
            Next;
          end;
        finally
          try
            GotoBookmark(TempBookmark);
          except
          end;
          FreeBookmark(TempBookmark);
        end;
      finally
        EnableControls;
      end;
    ...
end;

}



{Windwos ���α׷� ������� ã��}

// 1. The Documented way

{
  An application can check if a window is responding to messages by
  sending the WM_NULL message with the SendMessageTimeout function.

  Um zu uberprufen, ob ein anderes Fenster (Anwendung) noch reagiert,
  kann man ihr mit der SendMessageTimeout() API eine WM_NULL Nachricht schicken.
}

function AppIsResponding(ClassName: string): Boolean;
const
  { Specifies the duration, in milliseconds, of the time-out period }
  TIMEOUT = 50;
var
  Res: DWORD;
  h: HWND;
begin
  h := FindWindow(PChar(ClassName), nil);
  if h <> 0 then
    Result := SendMessageTimeOut(H,
      WM_NULL,
      0,
      0,
      SMTO_NORMAL or SMTO_ABORTIFHUNG,
      TIMEOUT,
      Res) <> 0
  else
    ShowMessage(Format('%s not found!', [ClassName]));
end;

procedure TForm1.Button1Click(Sender: TObject);
begin 
  if AppIsResponding('OpusApp') then 
    { OpusApp is the Class Name of WINWORD }
    ShowMessage('App. responding');
end; 

// 2. The Undocumented way 

{ 
  // Translated form C to Delphi by Thomas Stutz 
  // Original Code: 
  // (c)1999 Ashot Oganesyan K, SmartLine, Inc 
  // mailto:ashot@aha.ru, http://www.protect-me.com, http://www.codepile.com 

The code doesn't use the Win32 API SendMessageTimout function to
determine if the target application is responding but calls
undocumented functions from the User32.dll.

--> For NT/2000/XP the IsHungAppWindow() API:

The function IsHungAppWindow retrieves the status (running or not responding)
of the specified application

IsHungAppWindow(Wnd: HWND): // handle to main app's window
BOOL;

--> For Windows 95/98/ME we call the IsHungThread() API

The function IsHungThread retrieves the status (running or not responding) of
the specified thread

IsHungThread(DWORD dwThreadId): // The thread's identifier of the main app's window
BOOL;

Unfortunately, Microsoft doesn't provide us with the exports symbols in the
User32.lib for these functions, so we should load them dynamically using the
GetModuleHandle and GetProcAddress functions:
}

// For Win9X/ME
function IsAppRespondig9X(dwThreadId: DWORD): Boolean;
type
  TIsHungThread = function(dwThreadId: DWORD): BOOL; stdcall;
var
  hUser32: THandle;
  IsHungThread: TIsHungThread;
begin
  Result := True;
  hUser32 := GetModuleHandle('user32.dll');
  if (hUser32 > 0) then
  begin
    @IsHungThread := GetProcAddress(hUser32, 'IsHungThread');
    if Assigned(IsHungThread) then
    begin
      Result := not IsHungThread(dwThreadId);
    end;
  end;
end;

// For Win NT/2000/XP
function IsAppRespondigNT(wnd: HWND): Boolean;
type
  TIsHungAppWindow = function(wnd:hWnd): BOOL; stdcall;
var
  hUser32: THandle;
  IsHungAppWindow: TIsHungAppWindow;
begin
  Result := True;
  hUser32 := GetModuleHandle('user32.dll');
  if (hUser32 > 0) then
  begin
    @IsHungAppWindow := GetProcAddress(hUser32, 'IsHungAppWindow');
    if Assigned(IsHungAppWindow) then
    begin
      Result := not IsHungAppWindow(wnd);
    end;
  end;
end;

function IsAppRespondig(Wnd: HWND): Boolean;
begin
if not IsWindow(Wnd) then
begin
   ShowMessage('Incorrect window handle!');
   Exit;
end;
if Win32Platform = VER_PLATFORM_WIN32_NT then
   Result := IsAppRespondigNT(wnd)
else
   Result := IsAppRespondig9X(GetWindowThreadProcessId(Wnd,nil));
end;

// Example: Check if Word is hung/responding

procedure TForm1.Button3Click(Sender: TObject);
var
  Res: DWORD;
  h: HWND;
begin
  // Find Winword by classname
  h := FindWindow(PChar('OpusApp'), nil);
  if h <> 0 then
  begin
    if IsAppRespondig(h) then
      ShowMessage('Word is responding!')
    else
      ShowMessage('Word is not responding!');
  end
  else
    ShowMessage('Word is not open!');
end;



