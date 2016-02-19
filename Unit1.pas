unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs,VBIDE_TLB,Word_TLB,Office_TLB,
  Vcl.StdCtrls, Vcl.ComCtrls, Vcl.Samples.Spin;

type
  TForm1 = class(TForm)
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Edit4: TEdit;
    Label4: TLabel;
    Edit5: TEdit;
    Label5: TLabel;
    Edit6: TEdit;
    Label6: TLabel;
    Edit7: TEdit;
    Label7: TLabel;
    Edit8: TEdit;
    Label8: TLabel;
    Edit9: TEdit;
    Label9: TLabel;
    Edit10: TEdit;
    Label10: TLabel;
    Edit11: TEdit;
    Edit12: TEdit;
    Label11: TLabel;
    Label12: TLabel;
    Edit13: TEdit;
    Label13: TLabel;
    Edit14: TEdit;
    Label14: TLabel;
    Edit15: TEdit;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    SpinEdit1: TSpinEdit;
    Edit16: TEdit;
    Label18: TLabel;
    DateTimePicker1: TDateTimePicker;
    Label19: TLabel;
    Edit17: TEdit;
    Label20: TLabel;
    Edit18: TEdit;
    Label21: TLabel;
    Label22: TLabel;
    Edit19: TEdit;
    Label23: TLabel;
    Edit20: TEdit;
    Label24: TLabel;
    Label25: TLabel;
    DateTimePicker2: TDateTimePicker;
    Edit21: TEdit;
    Label26: TLabel;
    Edit22: TEdit;
    Label27: TLabel;
    Label28: TLabel;
    DateTimePicker3: TDateTimePicker;
    Button1: TButton;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var
  WordApp: WordApplication;
  Docs: Documents;
  Doc: WordDocument;
  Pars: Paragraphs;
  Par: Paragraph;
  D: OleVariant;
  i:integer;
begin
 WordApp := CoWordApplication.Create;
 WordApp.Visible := True;

  Docs := WordApp.Documents;
  Doc := Docs.Add('Normal', False, EmptyParam, True);
  Doc.Paragraphs.Format.LeftIndent:=WordApp.CentimetersToPoints(-1.4) ;
  Doc.Paragraphs.Item(1).Range.Text :=
    '���: '+ Edit1.Text
    +#13+'���: '+Edit2.Text +#09+'���. 001'
    +#13+'�������� � ���������� �������� ������ ������������� ��������,'+' ���������� ������� �� ����� �� �������� � �������� ������������ �������, �� ��������� ���������� �� ������ (�����) ������ ������������� �������� � ����� �����, ���������� ������ � ���� ��������������� ������'
    +#13+'����� �������������:'+Edit3.Text+#09+#09+#09+#09+#09+#09+'��� ��������� ����������:'+ Edit4.Text
    +#13+'�������������� � ��������� �����(���):'+ Edit6.Text+#09+#09+#09+#09+'�� ����� ���������� (�����)(���):'+Edit5.Text
    +#13+'�� ������� 20-�� ����� ������ ���������� �� ��������� ������� ����� �������� ����������'
    +#13+ Edit7.Text
    +#13+ Edit8.Text
    +#13+ Edit9.Text
    +#13+ '(���������� �����)'
    +#13+ '��� ���� ������������� ������������ �� �������������� �����:'+Edit10.Text+'.'+Edit11.Text+'.'+Edit12.Text
    +#13+'����� ����������� ��������: '+Edit13.Text
    +#13+'��' +Edit14.Text+ '���������'+#09+#09+'� ����������� �������������� ���������� �(���) ���� ��'+Edit14.Text+'������'
    +#13+ '������������� � ������� ��������, ��������� � ��������� ���������,�����������'+ '����������� ����������� ��������� ������'
    +#13+ inttostr(Spinedit1.Value)+ '1-��������� �����'+#09+#09+#09+#09+'�������� � ������������� ���������'
    +#13+'2-������������� ����������� �����'
    +#13+ Edit7.Text+#09+#09+#09+#09+'������ �������� �����������(���)'
    +#13+ Edit8.Text+#09+#09+#09+#09+'��'+'���������'
    +#13+ Edit9.Text+#09+#09+#09+#09+'� ����������� �������������� ���������� �(���) ���� ��'+'������'
    +#13+'(������, ���, �������� ��������� )'
    +#13+'���� ������������ ���������'+ DateToStr(DateTimePicker1.DateTime)
    +#13+'��������������� �� �'
    +#13
    +#13
    +#13
    +#13
    +#13
    +#13+'(������������ �����������-������������� ����������� �����)'
    +#13+'�������'+#09+#09+#09+#09+'����'+ DateToStr(DateTimePicker1.DateTime)
    +#13+'������������ ���������,��������������� ���������� �������������'+edit17.Text
    +#13
    +#13+'���'+#09+#09+#09+#09+'�������'
    +#13+'���'+ Edit1.Text
    +#13+'���'+Edit2.Text +#09+#09+#09+#09+'���. 001'
    +#13+'������.�������� � ���������� �������� ������ ������������� ��������,'+'���������� ������� �� ����� �� �������� � �������� ������������ �������,'+'�� ��������� ���������� �� ������ (�����) ������ ������������� ��������.'
    +#13+'����������'+'��� ������'+'�������� �����������'
    +#13+'��� ��������� �������������'+#09+#09+#09+#09+'010'+edit18.Text
    +#13+'��� �� �����'+#09+#09+#09+#09+'020'+edit19.Text
    +#13+'����� ����� ��������������� ������, ���������� ������ � ������'+#09+#09+#09+#09+'030'+edit20.Text
    +#13+'���� ������ ��������������� ������'+#09+#09+#09+#09+'040'+DateToStr(DateTimePicker2.DateTime)
    +#13+'�����'+#09+#09+#09+#09+'050'+edit21.Text
    +#13+'���������� �� ������ (�����) ������ ������������� ��������'+'�����'+#09+#09+#09+#09+'060'+edit22.Text
    +#13+'���� ��������'+#09+#09+#09+#09+'070'+DateToStr(DateTimePicker3.DateTime)
    +#13+'��� ������������ �������� ������ ������������� ��������'+#09+#09+#09+#09+'����������� ����� �������'+#09+#09+#09+#09+'����� ��������������� ������'
    +#13
    +#13
    +#13
    +#13
    +#13
    +#13
    +#13+'������������� � ������� ��������, ��������� �� ������ ��������, �����������'
    +#13+'�������'+#09+#09+#09+#09+'����';

for I := 1 to 3 do
 Doc.Paragraphs.Item(i).Alignment:=wdAlignParagraphCenter;
 for I := 27 to 29 do
 Doc.Paragraphs.Item(i).Alignment:=wdAlignParagraphCenter;
//for I := 3 to 6 do
 Doc.Paragraphs.Item(3).Range.Font.Bold := 1;

 Doc.Paragraphs.Item(6).Range.Font.Size := 9;
 for I := 6 to 10 do
   Doc.Paragraphs.Item(i).Alignment:=wdAlignParagraphCenter;
 Doc.Paragraphs.Item(10).Range.Font.Size := 9;
 Doc.Tables.Add(Doc.Paragraphs.Item(14).Range,1,2,wdWord9TableBehavior,wdAutoFitFixed);
 Doc.Tables.Item(1).Cell(1,1).Range.text:=
 '������������� � ������� ��������, ��������� � ��������� ���������,�����������'
 +#13+inttostr(Spinedit1.Value)+ '1-��������� �����'+'2-������������� ����������� �����'
 +#13+ Edit7.Text
 +#13+ Edit8.Text
 +#13+ Edit9.Text
 +#13+'(������, ���, �������� ��������� )'
 +#13;

  Doc.Tables.Item(1).Range.Font.Bold := 1;
end;
end.