unit Word;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Word2000, Vcl.OleServer, Vcl.StdCtrls;

type
  TForm1 = class(TForm)
    WordApplication1: TWordApplication;
    WordDocument1: TWordDocument;
    WordFont1: TWordFont;
    WordParagraphFormat1: TWordParagraphFormat;
    ButtonReplace: TButton;
    procedure ButtonReplaceClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.ButtonReplaceClick(Sender: TObject);   //автозамена в выбранном документе Word
var                                                     //изучаем тут: https://forum.vingrad.ru/forum/topic-84634/view-all.html
  Word: OleVariant;

function FindAndReplace(FindText, NewText : string):boolean; {функция поиска и замены}
begin
  Word.Selection.Find.ClearFormatting;                  //форматирование не учитывать при поиске
  Word.Selection.Find.Replacement.ClearFormatting;      //форматирование не учитывать при замене
  Word.Selection.Find.Text:=FindText;                   //что найти
  Word.Selection.Find.Replacement.Text:=NewText;        //на что заменить
  Word.Selection.Find.Forward:=True;                    //искать от начала к концу документа
  Word.Selection.Find.Wrap:=wdFindContinue;             //искать по всему документу (если поиск был не сначала, то по завершении документа он продолжится с начала)
  Word.Selection.Find.Format:=False;                    //не обращать внимание на форматирование искомого
  Word.Selection.Find.MatchCase:= False;                //не придавать замещающему тексту форматирование заменяемого
  Word.Selection.Find.MatchWholeWord:=False;            //искать не только целые слова
  Word.Selection.Find.MatchWildcards:=False;            //не использовать подстановочные знаки
  Word.Selection.Find.MatchAllWordForms:=False;         //искать все формы слова (только для букв)
  Word.Selection.Find.Execute(ReplaceText:=wdReplaceAll);   //выполнить замену, заменить все (только за один проход).
end;                                                    //подробнее о параметрах поиска: https://docs.microsoft.com/ru-ru/office/vba/api/word.find.execute


begin
  Word := CreateOleObject('Word.Application');
  Word.DisplayAlerts := 0;                       //отключить показ окна сохранения при закрытии измененного документа
  Word.Documents.Open('C:\Users\V\Documents\ex.docx', ReadOnly := false); //см. про Open комментарий ниже
  Word.Visible := False;                         //сделать невидимым

  FindAndReplace('  ', ' ');
  While Word.Selection.Find.Execute() do         //пока результат замены true
     FindAndReplace('  ', ' ');                  //снова запускать функцию замены

  Word.ActiveDocument.Close(wdSaveChanges);      //закрыть с сохранением изменений
  Word.Quit;                                     //завершить работу с документом
  Word := UnAssigned;                            //закрыть документ
  ShowMessage('Автозамена завершена');
end;

end.
