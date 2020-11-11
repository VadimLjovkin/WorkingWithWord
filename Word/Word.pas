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

procedure TForm1.ButtonReplaceClick(Sender: TObject);   //���������� � ��������� ��������� Word
var                                                     //������� ���: https://forum.vingrad.ru/forum/topic-84634/view-all.html
  Word: OleVariant;

function FindAndReplace(FindText, NewText : string):boolean; {������� ������ � ������}
begin
  Word.Selection.Find.ClearFormatting;                  //�������������� �� ��������� ��� ������
  Word.Selection.Find.Replacement.ClearFormatting;      //�������������� �� ��������� ��� ������
  Word.Selection.Find.Text:=FindText;                   //��� �����
  Word.Selection.Find.Replacement.Text:=NewText;        //�� ��� ��������
  Word.Selection.Find.Forward:=True;                    //������ �� ������ � ����� ���������
  Word.Selection.Find.Wrap:=wdFindContinue;             //������ �� ����� ��������� (���� ����� ��� �� �������, �� �� ���������� ��������� �� ����������� � ������)
  Word.Selection.Find.Format:=False;                    //�� �������� �������� �� �������������� ��������
  Word.Selection.Find.MatchCase:= False;                //�� ��������� ����������� ������ �������������� �����������
  Word.Selection.Find.MatchWholeWord:=False;            //������ �� ������ ����� �����
  Word.Selection.Find.MatchWildcards:=False;            //�� ������������ �������������� �����
  Word.Selection.Find.MatchAllWordForms:=False;         //������ ��� ����� ����� (������ ��� ����)
  Word.Selection.Find.Execute(ReplaceText:=wdReplaceAll);   //��������� ������, �������� ��� (������ �� ���� ������).
end;                                                    //��������� � ���������� ������: https://docs.microsoft.com/ru-ru/office/vba/api/word.find.execute


begin
  Word := CreateOleObject('Word.Application');
  Word.DisplayAlerts := 0;                       //��������� ����� ���� ���������� ��� �������� ����������� ���������
  Word.Documents.Open('C:\Users\V\Documents\ex.docx', ReadOnly := false); //��. ��� Open ����������� ����
  Word.Visible := False;                         //������� ���������

  FindAndReplace('  ', ' ');
  While Word.Selection.Find.Execute() do         //���� ��������� ������ true
     FindAndReplace('  ', ' ');                  //����� ��������� ������� ������

  Word.ActiveDocument.Close(wdSaveChanges);      //������� � ����������� ���������
  Word.Quit;                                     //��������� ������ � ����������
  Word := UnAssigned;                            //������� ��������
  ShowMessage('���������� ���������');
end;

end.
