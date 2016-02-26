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
  Doc.Paragraphs.Format.SpaceAfter:=14;
  Doc.Paragraphs.Format.LeftIndent:=WordApp.CentimetersToPoints(-1.4) ;
  Doc.Paragraphs.Item(1).Range.Text :=
    'ИНН: '+ Edit1.Text
    +#13+'КПП: '+Edit2.Text +#09+'Стр. 001'
    +#13+'Сведения о количестве объектов водных биологических ресурсов,'+' подлежащих изьятию из среды их обитания в качестве разрешенного прилова, на основании разрешения на добычу (вылов) водных биологических ресурсов и сумма сбора, подлежащих уплате в виде единовременного взноса'
    +#13+'Номер корректировки:'+Edit3.Text+#09+#09+#09+#09+#09+#09+'Год получения разрешения:'+ Edit4.Text
    +#13+'Представляется в налоговый орган(код):'+ Edit6.Text+#09+#09+#09+#09+'По месту нахождения (учета)(Код):'+Edit5.Text
    +#13+'не позднее 20-го числа месяза следующего за последним месяцем срока действия разрешения'
    +#13+ Edit7.Text
    +#13+ Edit8.Text
    +#13+ Edit9.Text
    +#13+ '(плательшик сбора)'
    +#13+ 'Код вида экономической деятельности по классификатору ОКВЭД:'+Edit10.Text+'.'+Edit11.Text+'.'+Edit12.Text
    +#13+'Номер контактного телефона: '+Edit13.Text
    +#13+'На' +Edit14.Text+ 'страницах'+#09+#09+'с приложением подтверждающих документов и(или) коий на'+Edit14.Text+'листах'
    +#13+ 'достоверность и полноту сведений, угазанных в настоящем документе,подтверждаю'+ 'заполняется сотрудником налогвого органа'
    +#13
    +#13+'ИНН'+ Edit1.Text
    +#13+'КПП'+Edit2.Text +#09+'Стр. 001'
    +#13+'Раздел.Сведения о количестве объектов водных биологических ресурсов,'+'подлежащих изьятию из среды их обитания в качестве разрешенного прилова,'+'на основании разрешения на добычу (вылов) водных биологических ресурсов.'
    +#13+'Показатели'+#09+#09+#09+#09+'Код строки'+#09+#09+#09+#09+'Значения показателей'
    +#13+'Код бюджетной классификации'+#09+#09+'010'+#09+#09+#09+#09+#09+#09+edit18.Text
    +#13+'Код по ОКТМО'+#09+#09+#09+#09+#09+'020'+#09+#09+#09+#09+#09+#09+edit19.Text
    +#13+'Общая сумма единовременного взноса,'+#09+'030'+#09+#09+#09+#09+#09+#09+edit20.Text
    +#13+ 'подлежащая уплате в бюджет'
    +#13+'Дата уплаты единовременного взноса'+#09+#09+#09+#09+'040'+DateToStr(DateTimePicker2.DateTime)
    +#13+'серия'+#09+#09+#09+#09+'050'+edit21.Text
    +#13+'Разрешение на добычу (вылов) водных биологических ресурсов'+'номер'+#09+#09+#09+#09+'060'+edit22.Text
    +#13+'Дата полученя'+#09+#09+#09+#09+'070'+DateToStr(DateTimePicker3.DateTime)
    +#13+'Код наименования объектов водных биологических ресурсов'+#09+#09+#09+#09+'Разрешенный объем прилова'+#09+#09+#09+#09+'Сумма единовременного взноса'
    +#13
    +#13
    +#13
    +#13
    +#13
    +#13
    +#13+'Достоверность и полноту сведений, указанных на данной странице, подтверждаю'
    +#13+'Подпись'+#09+#09+#09+#09+'Дата';

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
 'Достоверность и полноту сведений, угазанных в настоящем документе,подтверждаю'
 +#13+inttostr(Spinedit1.Value)
 +#13+ Edit7.Text
 +#13+ Edit8.Text
 +#13+ Edit9.Text
 +#13+'(фамили, имя, отчество полностью )'
 +#13+'(наименование организации-представителя плательшика сбора)'
 +#13+'Подпись'+#09+#09+'дата'+ DateToStr(DateTimePicker1.DateTime)
 +#13+'Наименование документа,подтверждающего полномочия представителя'+edit17.Text;
  Doc.Tables.Item(1).Cell(1,2).Range.text:=
  'Заполняется сотрудником налогвого органа'
  +#13+'Сведения о представлении документа'
  +#13+'данный документ представлен(код):'
  +#13+#9+'на'+'страницах'
  +#13+'с приложением подтверждающих документов и(или) коий на'+'листах'
  +#13+'дата преставления документа'+ DateToStr(DateTimePicker1.DateTime)
  +#13+'Зарегистрирован за №'
  +#13
  +#13
  +#13
  +#13+'ФИО'+#09+#09+#09+#09+'Подпись';
 for I := 14 to 24 do
  Doc.Paragraphs.Item(i).Alignment:=wdAlignParagraphCenter;
 for I := 36 to 39 do
  Doc.Paragraphs.Item(i).Alignment:=wdAlignParagraphCenter;
   Doc.Paragraphs.Item(38).Range.Font.Bold := 1;
  Doc.Paragraphs.Item(14).Range.Font.Bold := 1;
  Doc.Paragraphs.Item(23).Range.Font.Bold := 1;
end;
end.
