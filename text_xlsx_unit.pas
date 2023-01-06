unit text_xlsx_unit;

interface

uses
  BxlsxREader,BxlsxWriter,BDocxWriter,BrtfWriter,
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Grids;

type
  TForm1 = class(TForm)
    StringGrid1: TStringGrid;
    OpenDialog1: TOpenDialog;
    Button1: TButton;
    Button2: TButton;
    SaveDialog1: TSaveDialog;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
  private
    { Private declarations }
  public
     FirstCall:boolean;
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

// callback function
procedure rrr(UserParm:nativeUint; sheet,col,row:longword; value:widestring); stdcall;
var f:TForm1;
begin
   f := TForm1(UserParm);
   if sheet =  1 then
   begin
      if f.FirstCall then
      begin
         // First call receive sheet dimension
         f.StringGrid1.ColCount := col;
         f.StringGrid1.RowCount := row;
         f.FirstCall := false;
      end else begin
         if (row = 0) and (col = 0) then
         begin
            //FINISH
         end else begin
            f.stringGrid1.Cells[col-1,row-1] := value;
         end;
      end;
   end else begin
      MessageBox(0,'Sorry, but demo is designet for workbook with only one sheet','info',mb_ok);
   end;
end;

procedure TForm1.Button1Click(Sender: TObject);
var r:BTxlsxReader;
begin
   // Test with call back fill
   StringGrid1.CleanupInstance;
   OpenDialog1.execute;
   r := BTxlsxReader.create;
   FirstCall := true;
   if OpenDialog1.Files.count = 1 then
   begin
      if not r.OpenFile(OpenDialog1.Files[0],@rrr,nativeUint(self)) then
      begin
         MessageBox(0,'error opening file','error',mb_ok);
      end;
   end;
   r.Free;
end;


procedure TForm1.Button3Click(Sender: TObject);
var r:BTxlsxReader;
    mr,mc,rr,cc:longword;
    v:widestring;
begin
   // Test with sequential read
   StringGrid1.CleanupInstance;
   OpenDialog1.execute;
   r := BTxlsxReader.create;
   FirstCall := true;
   if OpenDialog1.Files.count = 1 then
   begin
      if not r.OpenFile(OpenDialog1.Files[0],nil,0) then
      begin
         MessageBox(0,'error opening file','error',mb_ok);
      end else begin
         r.SelectSheet(1);
         r.GetSheetBounds(mc,mr);
         StringGrid1.RowCount := mr;
         StringGrid1.ColCount := mc;
         for rr := 1 to mr do
         begin
            for cc := 1 to mc do
            begin
               if r.GetCellValue(cc,rr,v) then
               begin
                  stringGrid1.Cells[cc-1,rr-1] := v;
               end else begin
                  //error ???
               end;
            end;
         end;
      end;
   end;
   r.Free;
end;

procedure TForm1.Button4Click(Sender: TObject);
var d:BTDocxWriter;
begin
   d := BTDocxWriter.Create;
   d.AddNewPage();
   d.AddText('Hello   mamam friend Боги');
   d.AddNewLine;
   d.AddText('New line text ');
   d.AddText('is here?');
   d.Generate('d:\t1.docx');
   d.Free;
end;

procedure TForm1.Button5Click(Sender: TObject);
var d:BTrtfWriter;
begin
   d := BTrtfWriter.Create;
   d.AddNewPage;
   d.AddText('Hello dear friend wedqw qwed qwed qwe d q');
   d.AddNewLine;
   d.AddText('Hello dear friend 2 wedewq fqw ef qefqef ');
   d.SetTextAlign(2);
   d.AddNewLine;
   d.AddText('Hello dear friend 3 tyhyt tyhtyhtryh tryhtry ');
//   d.AddPicture('d:\mvr-gerb.png',3,5);
   d.AddNewLine;
   d.addtable;
//   d.AddPicture('d:\xa001.jpg',4,4);

   d.Generate('d:\t1.rtf');
   d.Free;
end;



procedure TForm1.Button2Click(Sender: TObject);
var w:BTxlsxWriter;
    sheet:nativeUint;
    font1,font2:longword;
    cell:nativeUint;
begin
   //Write Test
   SaveDialog1.execute;
   if SaveDialog1.files.count = 1 then
   begin
      w := BTxlsxWriter.create;

      font1 := w.SetUpFont('Wingdings',14,0,$0000FF,2); // Blue
      font2 := w.SetUpFont('Times New Roman',14,1,$ff0000); // red

      sheet := w.AddSheet('My test sheet');

      cell := w.AddCell(sheet,1,1,'He Бо');
      w.AddCellStyle(cell,$FFFF00,15,font2,1);
      cell := w.AddCell(sheet,2,1,'A');
      w.AddCellStyle(cell,$A0A0A0,0,font1,2);

      w.AddCell(sheet,3,1,'123.3');

      w.AddCell(sheet,'C1','11');
//      w.AddCell(sheet,'c2','22');
//      w.AddCell(sheet,'e1','C1+C2',2);
      if not w.Generate(SaveDialog1.files[0]) then
      begin

         //error
      end;
         w.GenerateCSV(SaveDialog1.files[0]+'.csv',sheet) ;
      w.Free;
   end;
end;





end.
