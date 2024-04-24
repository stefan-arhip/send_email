unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  EditBtn, Spin, ComObj;

type

  { TForm1 }

  TForm1 = class(TForm)
    Button1: TButton;
    cbAuthenticate: TCheckBox;
    cbSsl: TCheckBox;
    edBCC: TEdit;
    edCC: TEdit;
    edTo: TEdit;
    edPassword: TEdit;
    edServer: TEdit;
    edSubject: TEdit;
    edUsername: TEdit;
    edSender: TEdit;
    fnAttachment: TFileNameEdit;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Memo1: TMemo;
    sePort: TSpinEdit;
    procedure Button1Click(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

procedure TForm1.Button1Click(Sender: TObject);
Const Cdo= 'http://schemas.microsoft.com/cdo/configuration/';
Var Email: OleVariant;
    _Attachment, _Message: Variant;
begin
  _Attachment:= fnAttachment.FileName;

  Email:= ComObj.CreateOleObject('CDO.Message');
  Email.From:= ShortString(edSender.Text);
  // "TO" is a reverved word, so we could use with prefix "&"...
  Email.&To:= ShortString(edTo.Text);
  Email.CC := ShortString(edCC.Text);
  Email.BCC:= ShortString(edBCC.Text);

  Email.Subject:= ShortString(edSubject.Text);
  // Email.HtmlBody should be an Variant variable to be
  _Message:= '<html><head>'+
                   '<style type="text/css">'+
                   '<!--'+
                   'p.text{font-family:arial;font-size:10pt;font-weight:000;color:#0000CC}'+
                   'p.legend{font-family:arial;font-size:10pt;font-weight:000;font-style:italic;color:#cccccc}'+
                   '-->'+
                   '</style>'+
                   '</head><body>'+
                   '<p class="text">'+ Memo1.Lines.Text+ '</p><br><br>'+
                   '<p class="text">See report in attachment</p><br><br>'+
                   '<p class="legend">Email sent from Free Pascal created by https://github.com/stefan-arhip</p>'+
                   '</body></html>';   ;
  Email.HtmlBody:= ShortString(_Message);
  //Email.TextBody:= ShortString(FormatDateTime('yyyy-mm-dd hh:nn:ss', Now())+ #13+ Memo1.Lines.Text);

  If FileExists(_Attachment) Then
    Email.AddAttachment(_Attachment);
  Email.Configuration.Fields.Item(Cdo+ 'sendusing'):= 2;
  Email.Configuration.Fields.Item(Cdo+ 'smtpserver'):= ShortString(edServer.Text);
  Email.Configuration.Fields.Item(Cdo+ 'smtpserverport'):= sePort.Value;
  Email.Configuration.Fields.Item(Cdo+ 'smtpusessl'):= False;
  Email.Configuration.Fields.Item(Cdo+ 'smtpauthenticate'):= False;
  Email.Configuration.Fields.Item(Cdo+ 'sendusername'):= ShortString(edUsername.Text);
  Email.Configuration.Fields.Item(Cdo+ 'sendpassword'):= ShortString(edPassword.Text);
  Email.Configuration.Fields.Item(Cdo+ 'smtpconnectiontimeout'):= 60;
  Email.Configuration.Fields.Update;
  Try
    Email.Send;
  Except
    ShowMessage('EMAIL NOT SENT!');
  End;
End;

end.

