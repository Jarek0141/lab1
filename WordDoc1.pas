unit WordDoc1;



interface

  uses Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs,VBIDE_TLB,Word_TLB,Office_TLB,
  Vcl.StdCtrls, Vcl.ComCtrls, Vcl.Samples.Spin;

 type
 Aos = array of string;
implementation

procedure sostav(Mas:Aos;Date:TDate;profile:integer);
begin

 WordApp := CoWordApplication.Create;
 WordApp.Visible := True;

  Docs := WordApp.Documents;
  Doc := Docs.Add('Normal', False, EmptyParam, True);
  Doc.Paragraphs.Format.SpaceAfter:=14;
  Doc.Paragraphs.Format.LeftIndent:=WordApp.CentimetersToPoints(-1.4) ;
  Doc.Paragraphs.Item(1).Range.Text :=
    'ИНН: '+ Mas[1]
    +#13+'КПП: '+Mas[2] +#09+'Стр. 001'
    +#13+'Сведения о количестве объектов водных биологических ресурсов,'+' подлежащих изьятию из среды их обитания в качестве разрешенного прилова, на основании разрешения на добычу (вылов) водных биологических ресурсов и сумма сбора, подлежащих уплате в виде единовременного взноса'
    +#13+'Номер корректировки:'+Mas[3]+#09+#09+#09+#09+#09+#09+'Год получения разрешения:'+ Mas[4]
    +#13+'Представляется в налоговый орган(код):'+ Mas[6]+#09+#09+#09+#09+'По месту нахождения (учета)(Код):'+Mas[5]
    +#13+'не позднее 20-го числа месяза следующего за последним месяцем срока действия разрешения'
    +#13+ Mas[7]
    +#13+ Mas[8]
    +#13+ Mas[9]
    +#13+ '(плательшик сбора)'
    +#13+ 'Код вида экономической деятельности по классификатору ОКВЭД:'+Mas[10]+'.'+Mas[11]+'.'+Mas[11]
    +#13+'Номер контактного телефона: '+Mas[12]
    +#13+'На' +Mas[13]+ 'страницах'+#09+#09+'с приложением подтверждающих документов и(или) коий на'+Mas[14]+'листах'
    +#13+ 'достоверность и полноту сведений, угазанных в настоящем документе,подтверждаю'+ 'заполняется сотрудником налогвого органа'
    +#13
    +#13+'ИНН'+ Mas[1]
    +#13+'КПП'+Mas[2] +#09+'Стр. 002'
    +#13+'Раздел.Сведения о количестве объектов водных биологических ресурсов,'+'подлежащих изьятию из среды их обитания в качестве разрешенного прилова,'+'на основании разрешения на добычу (вылов) водных биологических ресурсов.'
    +#13+'Показатели'+#09+#09+#09+#09+'Код строки'+#09+#09+#09+#09+'Значения показателей'
    +#13+'Код бюджетной классификации'+#09+#09+'010'+#09+#09+#09+#09+#09+#09+Mas[18]
    +#13+'Код по ОКТМО'+#09+#09+#09+#09+#09+'020'+#09+#09+#09+#09+#09+#09+Mas[19]
    +#13+'Общая сумма единовременного взноса,'+#09+'030'+#09+#09+#09+#09+#09+#09+Mas[20]
    +#13+ 'подлежащая уплате в бюджет'
    +#13+'Дата уплаты единовременного взноса'+#09+#09+'040'+#09+#09+#09+#09+#09+#09+DateToStr(Date.DateTime)
    +#13+' '+' '+#09+#09+#09+#09+'серия'+#09+#09+'050'+#09+#09+#09+#09+#09+#09+Mas[21]
    +#13+'Разрешение на добычу'+' '+#09+'номер'+#09+#09+'060'+#09+#09+#09+#09+#09+#09+Mas[22]
    +#13+'(вылов)водных '
    +#13+'биологических ресурсов'+#09+'дата полученя'+#09+'070'+#09+#09+#09+#09+#09+#09+DateToStr(Date.DateTime)
    +#13
    +#13+'Достоверность и полноту сведений, указанных на данной странице, подтверждаю'
    +#13+'___________(подпись)'+#09+#09+'__________(дата)';

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
 'Достоверность и полноту сведений, угазанных в настоящем документе,подтверждаю'
 +#13+inttostr(profile)
 +#13+ Mas[7]
 +#13+ Mas[8]
 +#13+ Mas[9]
 +#13+'(фамили, имя, отчество полностью )'
 +#13+ Mas[16]
 +#13+'(наименование организации-представителя плательшика сбора)'
 +#13+'Подпись'+#09+#09+'дата'+ DateToStr(DateTimePicker1.DateTime)
 +#13+'Наименование документа,подтверждающего полномочия представителя'
 +#13+Mas[17];
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
  +#13+'_____________'+#09+#09'__________'
  +#13+'Фамилия И. О.'+#09+#09+#09+'Подпись';
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
  'Код наименования объектов водных биологических ресурсов'
  +#13+Mas[23]
  +#13+Mas[24]
  +#13+Mas[25];
  Doc.Tables.Item(2).Cell(1,2).Range.text:=
  'Разрешенный объем прилова'
  +#13
  +#13+Mas[26]
  +#13+Mas[27]
  +#13+Mas[28];
  Doc.Tables.Item(2).Cell(1,3).Range.text:=
  'Сумма единовременного взноса'
  +#13+Mas[29]
  +#13+Mas[30]
  +#13+Mas[31];
  for I := 64 to 66 do
  Doc.Paragraphs.Item(i).Alignment:=wdAlignParagraphCenter;
end;

end.
