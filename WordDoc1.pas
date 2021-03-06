unit WordDoc1;



interface

  uses Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs,VBIDE_TLB,Word_TLB,Office_TLB,
  Vcl.StdCtrls, Vcl.ComCtrls, Vcl.Samples.Spin;

 type
 Aos = array of string;
implementation

procedure sostav(Mas:Aos;Date:TDate;profile:integer);
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
  Doc.Paragraphs.Format.SpaceAfter:=14;
  Doc.Paragraphs.Format.LeftIndent:=WordApp.CentimetersToPoints(-1.4) ;
  Doc.Paragraphs.Item(1).Range.Text :=
    '���: '+ Mas[1]
    +#13+'���: '+Mas[2] +#09+'���. 001'
    +#13+'�������� � ���������� �������� ������ ������������� ��������,'+' ���������� ������� �� ����� �� �������� � �������� ������������ �������, �� ��������� ���������� �� ������ (�����) ������ ������������� �������� � ����� �����, ���������� ������ � ���� ��������������� ������'
    +#13+'����� �������������:'+Mas[3]+#09+#09+#09+#09+#09+#09+'��� ��������� ����������:'+ Mas[4]
    +#13+'�������������� � ��������� �����(���):'+ Mas[6]+#09+#09+#09+#09+'�� ����� ���������� (�����)(���):'+Mas[5]
    +#13+'�� ������� 20-�� ����� ������ ���������� �� ��������� ������� ����� �������� ����������'
    +#13+ Mas[7]
    +#13+ Mas[8]
    +#13+ Mas[9]
    +#13+ '(���������� �����)'
    +#13+ '��� ���� ������������� ������������ �� �������������� �����:'+Mas[10]+'.'+Mas[11]+'.'+Mas[11]
    +#13+'����� ����������� ��������: '+Mas[12]
    +#13+'��' +Mas[13]+ '���������'+#09+#09+'� ����������� �������������� ���������� �(���) ���� ��'+Mas[14]+'������'
    +#13+ '������������� � ������� ��������, ��������� � ��������� ���������,�����������'+ '����������� ����������� ��������� ������'
    +#13
    +#13+'���'+ Mas[1]
    +#13+'���'+Mas[2] +#09+'���. 002'
    +#13+'������.�������� � ���������� �������� ������ ������������� ��������,'+'���������� ������� �� ����� �� �������� � �������� ������������ �������,'+'�� ��������� ���������� �� ������ (�����) ������ ������������� ��������.'
    +#13+'����������'+#09+#09+#09+#09+'��� ������'+#09+#09+#09+#09+'�������� �����������'
    +#13+'��� ��������� �������������'+#09+#09+'010'+#09+#09+#09+#09+#09+#09+Mas[18]
    +#13+'��� �� �����'+#09+#09+#09+#09+#09+'020'+#09+#09+#09+#09+#09+#09+Mas[19]
    +#13+'����� ����� ��������������� ������,'+#09+'030'+#09+#09+#09+#09+#09+#09+Mas[20]
    +#13+ '���������� ������ � ������'
    +#13+'���� ������ ��������������� ������'+#09+#09+'040'+#09+#09+#09+#09+#09+#09+DateToStr(Date.DateTime)
    +#13+' '+' '+#09+#09+#09+#09+'�����'+#09+#09+'050'+#09+#09+#09+#09+#09+#09+Mas[21]
    +#13+'���������� �� ������'+' '+#09+'�����'+#09+#09+'060'+#09+#09+#09+#09+#09+#09+Mas[22]
    +#13+'(�����)������ '
    +#13+'������������� ��������'+#09+'���� ��������'+#09+'070'+#09+#09+#09+#09+#09+#09+DateToStr(Date.DateTime)
    +#13
    +#13+'������������� � ������� ��������, ��������� �� ������ ��������, �����������'
    +#13+'___________(�������)'+#09+#09+'__________(����)';

for I := 1 to 3 do
 Doc.Paragraphs.Item(i).Alignment:=wdAlignParagraphCenter;
 for I := 27 to 29 do
 Doc.Paragraphs.Item(i).Alignment:=wdAlignParagraphCenter;
 Doc.Paragraphs.Item(3).Range.Font.Bold := 1;
 Doc.Paragraphs.Item(6).Range.Font.Size := 9;
 for I := 6 to 10 do
   Doc.Paragraphs.Item(i).Alignment:=wdAlignParagraphCenter;
 Doc.Paragraphs.Item(10).Range.Font.Size := 9;
 Doc.Tables.Add(Doc.Paragraphs.Item(14).Range,1,2,wdWord9TableBehavior,wdAutoFitFixed);
 Doc.Tables.Item(1).Cell(1,1).Range.text:=
 '������������� � ������� ��������, ��������� � ��������� ���������,�����������'
 +#13+inttostr(profile)
 +#13+ Mas[7]
 +#13+ Mas[8]
 +#13+ Mas[9]
 +#13+'(������, ���, �������� ��������� )'
 +#13+ Mas[16]
 +#13+'(������������ �����������-������������� ����������� �����)'
 +#13+'�������'+#09+#09+'����'+ DateToStr(DateTimePicker1.DateTime)
 +#13+'������������ ���������,��������������� ���������� �������������'
 +#13+Mas[17];
  Doc.Tables.Item(1).Cell(1,2).Range.text:=
  '����������� ����������� ��������� ������'
  +#13+'�������� � ������������� ���������'
  +#13+'������ �������� �����������(���):'
  +#13+#9+'��'+'���������'
  +#13+'� ����������� �������������� ���������� �(���) ���� ��'+'������'
  +#13+'���� ������������ ���������'+ DateToStr(DateTimePicker1.DateTime)
  +#13+'��������������� �� �'
  +#13
  +#13
  +#13+'_____________'+#09+#09'__________'
  +#13+'������� �. �.'+#09+#09+#09+'�������';
 for I := 14 to 24 do
  Doc.Paragraphs.Item(i).Alignment:=wdAlignParagraphCenter;
 for I := 36 to 39 do
  Doc.Paragraphs.Item(i).Alignment:=wdAlignParagraphCenter;
   Doc.Paragraphs.Item(40).Range.Font.Bold := 1;
  Doc.Paragraphs.Item(14).Range.Font.Bold := 1;
  Doc.Paragraphs.Item(25).Range.Font.Bold := 1;
 for I := 48 to 50 do
  Doc.Paragraphs.Item(i).Alignment:=wdAlignParagraphLeft;
  Doc.Tables.Add(Doc.Paragraphs.Item(51).Range,1,3,wdWord9TableBehavior,wdAutoFitFixed);
  Doc.Tables.Item(2).Cell(1,1).Range.text:=
  '��� ������������ �������� ������ ������������� ��������'
  +#13+Mas[23]
  +#13+Mas[24]
  +#13+Mas[25];
  Doc.Tables.Item(2).Cell(1,2).Range.text:=
  '����������� ����� �������'
  +#13
  +#13+Mas[26]
  +#13+Mas[27]
  +#13+Mas[28];
  Doc.Tables.Item(2).Cell(1,3).Range.text:=
  '����� ��������������� ������'
  +#13+Mas[29]
  +#13+Mas[30]
  +#13+Mas[31];
  for I := 64 to 66 do
  Doc.Paragraphs.Item(i).Alignment:=wdAlignParagraphCenter;
end;

end.
