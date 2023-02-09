program test_xlsx;

uses
  Vcl.Forms,
  text_xlsx_unit in 'text_xlsx_unit.pas' {Form1},
  BPasZlib in 'BPasZlib.pas',
  BxlsxReader in 'BxlsxReader.pas',
  BxlsxWriter in 'BxlsxWriter.pas',
  BStrTools in 'BStrTools.pas',
  BTinyXML in 'BTinyXML.pas',
  BDate in 'BDate.pas',
  BUnicode in 'BUnicode.pas',
  BdocxWriter in 'BdocxWriter.pas',
  BrtfWriter in 'BrtfWriter.pas',
  BdocxForms in 'BdocxForms.pas',
  BFileTools in 'BFileTools.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
