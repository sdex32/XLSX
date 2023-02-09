unit text_xlsx_unit;

interface

uses
  BxlsxREader,BxlsxWriter,BDocxWriter,BrtfWriter,BxlsReader,BFileTools,BStrTools,BtinyXML,
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
    Label1: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
  private
    { Private declarations }
  public
     FirstCall:boolean;
    { Public declarations }
     tm:longword;
     r:BTxlsReader;
     rx:BTxlsxReader;
     w:longword;
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
            f.tm := (GetTickCount - f.tm ) div 1000;
            f.label1.Caption := ToStr(f.tm);
         end else begin
            f.stringGrid1.Cells[col-1,row-1] := value;
            inc(f.w);
            f.Label1.caption:= ToStr(f.w);
            sleep(10);

         end;
      end;
   end else begin
      MessageBox(0,'Sorry, but demo is designet for workbook with only one sheet','info',mb_ok);
   end;
end;

procedure TForm1.Button1Click(Sender: TObject);

begin
   w := 0;
   // Test with call back fill
   StringGrid1.CleanupInstance;
   OpenDialog1.execute;
   FirstCall := true;
   if OpenDialog1.Files.count = 1 then
   begin
      tm := GetTickCount;
      if ExtractFileExt(OpenDialog1.Files[0]) = '.xls' then
      begin
         r := BTxlsReader.create;
         if not r.OpenFile(OpenDialog1.Files[0],@rrr,nativeUint(self)) then
         begin
            MessageBox(0,'error opening file','error',mb_ok);
         end;
      end else begin
         rx := BTxlsxReader.create;
         if not rx.OpenFile(OpenDialog1.Files[0],@rrr,nativeUint(self)) then
         begin
            MessageBox(0,'error opening file','error',mb_ok);
         end;
      end;
   end;
end;


procedure TForm1.Button3Click(Sender: TObject);
var r:BTxlsReader;
    mr,mc,rr,cc:longword;
    v:widestring;
begin
   // Test with sequential read
   StringGrid1.CleanupInstance;
   OpenDialog1.execute;
   r := BTxlsReader.create;
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



procedure TForm1.Button6Click(Sender: TObject);
var sa:ansistring;
    xm:BTTinyXML;
    s:string;
    f:longint;
begin
   FileLoadEx('d:\sharedStrings.xml',sa);
   xm := BTTinyXML.Create;
   xm.LoadXML(sa,0);
   f := 0;
   xm.SelectXPath('/sst.uniqueCount',s,f);
   s:= s +'a';

end;

function FastCPosAEX(const a:ansichar; const s:ansistring):longword;
begin
   asm
      push  edi
      mov   eax, 0
      mov   edi, s
      or    edi, edi
      jz    @@out
      mov   ecx, dword ptr [edi - 4]
      mov   al, a
      mov   edx, ecx
      or    ecx, ecx
      jz    @@out

@@lop:
      repne   scasb
      mov     eax, 0
      or      ecx, ecx
      jz      @@out
      mov     eax, edx
      sub     eax, ecx
@@res:
 //     inc eax
@@out:
      pop   edi
      mov   Result, eax
   end;


(*
   /////// fast ansi str len
   cld
   mov ecx, -1
   mov edi, s
   xor eax, eax
   repnz scanb
   mov eax, -2
   sub eax, ecx


   *)
end;


function FastAnsiStrLen(var s):longword;
begin
   asm
      push edi
      cld
      mov ecx, -1
      mov edi, s
      or  edi, edi
      mov   eax, 0
      jz @@11
      repnz scasb
      mov eax, -2
      sub eax, ecx
@@11:
      pop edi
      mov Result, eax
   end;
end;

function FastAnsiStrPos(c:ansichar; var s; len:longword; stpos:longword=1):longword;
begin
   asm
      push  edi
      mov   edi, s
//      test  eax, eax
      or    edi, edi
      jz    @@out
      mov   eax, stpos
      mov   ecx, dword ptr [edi - 4]
      dec   eax
//      test  ecx, ecx
      or    ecx, ecx
      jz    @@out
      mov   dl, c
@@lop:
      cmp   dl, byte ptr [edi + eax]
      je    @@res
      lea   eax, [eax + 1]
      loop  @@lop
      mov   eax, 0
      jmp   @@out
@@res:
      inc   eax
@@out:
      pop   edi
      mov   Result, eax
   end;
end;



function FastPosA_2(const a,s:ansistring; startpos:longword=1):longword;
begin
   asm
      push  edi
      push  esi
      push  ebx
      mov   eax, startpos
      dec   eax
//      mov   eax, 0
      mov   edi, a
      mov   esi, s
      mov   ecx, dword ptr [esi - 4]
      or    ecx, ecx
      jz    @@out
      mov   edx, dword ptr [edi - 4]
      test  edx, $FFFFFFFC //1100 4..max
      jz    @@less4
      sub   ecx, edx
      inc   ecx
      sub   edx, 4
      mov   ebx, dword ptr [edi]
@@lop4:
      cmp   ebx, dword ptr [esi + eax]
      je    @@res4
      lea   eax, [eax + 1]
      loop  @@lop4
      mov   eax, 0
      jmp   @@out
@@res4:
      or    edx, edx
      jz    @@hres4 // no more
      push  ecx
      push  ebx
      push  edi
      push  eax
      mov   ecx, edx // rest
      add   edi, 4
      add   eax, 4
@@bychar:
      mov   bl, byte ptr [edi]
      cmp   bl, byte ptr [esi + eax]
      jne   @@ops
      lea   eax, [eax + 1]
      inc   edi
      loop  @@bychar
      // super ve ahve winner
      add   esp, 16
      sub   eax, 4
      jmp   @@hres4

@@ops:
      pop   eax
      pop   edi
      pop   ebx
      pop   ecx
      lea   eax, [eax + 1]
      jmp   @@lop4

@@hres4: // have result
      sub   eax, edx
      jmp   @@res


@@less4: //--------------------------------------------
      test  edx, $FFFFFFFE //1110 2-3
      jz    @@less2
      sub   ecx, edx
      inc   ecx
      shl   edx, 16 //put size
      mov   dx, word ptr [edi]
@@lop2:
      cmp   dx, word ptr [esi + eax]
      je    @@res2
      lea   eax, [eax + 1]
      loop  @@lop2
      mov   eax, 0
      jmp   @@out
@@res2:
      test  edx, $00010000
      jz    @@res  // only two chars eax on right char -1
      mov   bl, byte ptr [edi + 2] // test third (3) char
      lea   eax, [eax + 1]
      cmp   bl, byte ptr [esi + eax + 1]
      je    @@out // good result
      jmp   @@lop2 //continue

@@less2: //-------------------------------------------
      test  edx, $1  // 1 char
      jz    @@out
      mov   dl, byte ptr [edi]
@@lop:
      cmp   dl, byte ptr [esi + eax]
      je    @@res
      lea   eax, [eax + 1]
      loop  @@lop
      mov   eax, 0
      jmp   @@out
@@res:
      inc   eax

@@out:
      pop   ebx
      pop   esi
      pop   edi
      mov   Result, eax
   end;
end;


procedure TForm1.Button8Click(Sender: TObject);

var sa:ansistring;
    sw:widestring;
    xm:BTTinyXML;
    s:string;
    w,i,k,sz:longint;
    p:pointer;
    sb:array[1..255] of byte;
begin
   FileLoadEx('d:\sharedStrings.xml',sa);
// sa := 'GOTMABogiManilavanilaKotelandia';
// p:= @sa[1];
// p:= @sb[1];
// k := FastAnsiStrLen(pointer(@sa[1])^);
   k := FastPosA_2('ANTROPOVA',sa);
   sw := widestring(sa);
   sz := length(sa);
   w := GetTickCount;
   for i := 1 to 200000 do
   begin
     k:= Pos('ANTROPOVA',sa);
//     k := FastPosA_2('ANTROPOVA',sa);
//     k:=  FastAnsiStrPos('M',p^,sz);
//      k := FastCPosA('M',sa);
//      k := Pos('M',sa);
//      k := FastCPosAEX('M',sa);
//     k := FastCPosW2('M',sw);
   end;
   w := GetTickCount - w;
   label1.Caption := tostr(w);
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
