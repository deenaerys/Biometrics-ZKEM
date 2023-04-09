unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, OleCtrls, zkemkeeper_TLB, DB, UtBase64,
  zkonline_TLB;
const
  EnrollDataLong: integer=1836;
type
  TForm1 = class(TForm)
    CZKEM1: TCZKEM;
    sbtnRS232: TSpeedButton;
    sbtnTCPIP: TSpeedButton;
    Label5: TLabel;
    btnConnect: TButton;
    edtNetAddress: TLabeledEdit;
    edtNetPort: TLabeledEdit;
    edtBaudRate: TEdit;
    edtMessage: TEdit;
    btnFPTConvertion: TButton;
    AFXOnlineMain1: TAFXOnlineMain;
    btnFPRegister: TButton;
    edtEnrollNum: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    procedure btnConnectClick(Sender: TObject);
    procedure btnFPTConvertionClick(Sender: TObject);
    procedure sbtnTCPIPClick(Sender: TObject);
    procedure sbtnRS232Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnFPRegisterClick(Sender: TObject);
  private
    { Private declarations }
    DevID: integer;
    Commport: integer;
    enrollptr: array of Integer;
    enrollbyte: array of byte;
    EnrollData:string; //Fingerprint template
    EnrollNum:integer; //Enroll number
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.btnConnectClick(Sender: TObject);
var
  ErrorCode: integer;
  c: boolean;
begin
  if btnConnect.Tag = 1 then
  begin
    CZKEM1.RefreshData(DevId);
    CZKEM1.EnableDevice(DevId, TRUE);
    CZKEM1.Disconnect;
    btnConnect.Tag := 0;
    btnConnect.Caption := 'Connect';
  end
  else
  begin
    devid:=strtoint(edtNetPort.Text);
    if pos('COM', uppercase(edtNetAddress.text))=1 then
    begin
      commport:=ord(edtNetAddress.text[4])-ord('0');
      CZKEM1.CommPort := commport;
      c:=CZKEM1.Connect_Com(commport, DevID, strtoint(edtBaudRate.text));
    end
    else
    begin
      c:=CZKEM1.Connect_net(edtNetAddress.text, Devid);
      Devid:=1;
    end;                                                                    
    if c then
    begin
      edtMessage.Text := 'Device Connected.';
      CZKEM1.EnableDevice(DevID, false);
      btnConnect.Tag := 1;
      btnConnect.Caption := 'Disconnect';
    end
    else
    begin
      CZKEM1.GetLastError(ErrorCode);
      edtMessage.Text := format('! ConnectDevice ErrorNo.=%d',[ErrorCode]);
    end;
  end;
end;

procedure TForm1.btnFPTConvertionClick(Sender: TObject);
var
  MachineNum,pinnum,i,emPrivilege,emPassword,errorcode,j:integer;
begin
  if EnrollData='' then
  begin
    edtMessage.Text := '��û�еǼ�ָ��'; //��û�еǼ�ָ��
    exit;
  end;
  DecodeData64(EnrollData); //��Base64��ʽ��ģ���ַ���ת��Ϊԭʼ��ָ��ģ��
  Move(enrolldata[1], enrollbyte[0], Length(EnrollData));
  Sleep(100);
  errorcode := 0;
  inc(j);
  if not CZKEM1.SetUserTmp(1, EnrollNum, 1,enrollbyte[0]) then
  begin
    CZKEM1.GetLastError(errorcode);
    edtMessage.Text := 'SetEnrollData fail'; //����ָ�Ƶ��ѻ�ʧ��
    exit;
  end;
  edtMessage.Text := 'SetEnrollData success'; //����ָ���Ѿ����䵽�ѻ�
end;

procedure TForm1.sbtnTCPIPClick(Sender: TObject);
begin
  edtnetAddress.EditLabel.Caption := 'IP Address:';
  edtNetPort.EditLabel.Caption := 'Port Number:';
  edtnetAddress.Text := '192.168.1.201';
  edtNetPort.Text := '$1112';
end;

procedure TForm1.sbtnRS232Click(Sender: TObject);
begin
  edtnetAddress.EditLabel.Caption := 'Com Port:';
  edtNetPort.EditLabel.Caption := 'Device ID:';
  edtnetAddress.Text := 'COM1';
  edtNetPort.Text := '1';

end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  setlength(EnrollPtr, EnrollDataLong div SizeOf(Integer));
  setlength(enrollbyte,2048);
end;

procedure TForm1.btnFPRegisterClick(Sender: TObject);
begin
  try
    EnrollNum:=strtoint(edtEnrollNum.Text );
  except
    edtMessage.Text :='Enroll number is not a valid number';
    exit;
  end;
  if (EnrollNum < 1) and (EnrollNum>65534) then
  begin
    edtMessage.Text :='Enroll number must be in 1-65534 scope';
    exit;
  end;

  if AFXOnlineMain1.Register then
    EnrollData := AFXOnlineMain1.RegisterTemplate;
end;

end.
