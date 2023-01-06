unit BrtfWriter;

interface

// work well
//TODO Table   - under construction


type  BTRtfWriter = class {use cp1251 only work for me}
         private
            rtf :string;
            pg :longword;
            parag :string;
            palign :string;
            pdimen :string;
            ParagTouch:boolean;
            ColorTable :array [1..32] of longword;
            ColorTableSize :longword;
            FontTable :array [1..16] of string;
            FontTableSize :longword;
            cp1251 :boolean;
            A4 :boolean;
         public
            BWmode      :boolean;

            constructor Create;
            destructor  Destroy; override;
            procedure   Reset;
            function    GetRTFdoc:string;
            function    Generate(const FileName:string):boolean;

            function    SetupColorTable(R,G,B:longword):longword;
            function    SetupFontTable(const FntName:string):longword;
            procedure   SetupCodePage(const Page:string);

            procedure   AddNewPage(landscape:boolean = false);
            procedure   AddNewPageA4(landscape:boolean = false);  //????
            procedure   AddText(const Txt:string);
            procedure   AddNewLine; // create new paragraph.
            function    AddPicture(const Filename:string; xlng_cm,ylng_cm:single):boolean;
            procedure   AddTable;

            procedure   SetFont(id:longword);
            procedure   SetColor(id:longword);
            procedure   SetBgColor(id:longword);
            procedure   SetBold(state:boolean);
            procedure   SetItalic(state:boolean);
            procedure   SetUnderline(state:boolean);
            procedure   SetFirstLineIdent(cm:single);
            procedure   SetTextAlign(typ:longword);
            procedure   SetFontSize(sz:longword);

      end;




implementation

uses BStrTools,BFileTools,Windows;

{
  \slmult1  line spacing 0-?      1-single
  \pard - reset paragraf values
  \pard\li2414\sa200\sl276\slmult1 -- must be set after
  \sb space before
  \sa space after
  \sl soace beatween lines

   // ruler ot kude do kude
  \li Left ident 0 -def
  \fi first line
  \ri right  ident
  \qj - justify  aligmnet
}


//------------------------------------------------------------------------------
constructor BTRtfWriter.Create;
begin
   Reset;
end;

//------------------------------------------------------------------------------
destructor  BTRtfWriter.Destroy;
begin
   inherited;
end;

//------------------------------------------------------------------------------
procedure   BTRtfWriter.Reset;
begin
   pg := 0;
   A4 := false;
   cp1251 := false;
   BWmode := false;
   rtf := '';
   parag := '';
   palign := '\ql';
   pdimen := '\fi0';
   ParagTouch := false;
   ColorTableSize := 1;
   ColorTable[1] := 0; // black
   FontTableSize := 0;
end;

//------------------------------------------------------------------------------
function    BTRtfWriter.GetRTFdoc:string;
var i,w:longword;
    ortf,urtf:string;
begin
   if ParagTouch then AddNewLine;
//start trhe doc
   ortf := '{\rtf1\ansi';
   if cp1251 then ortf := ortf + '\ansicpg1251';
   ortf := ortf + '\deff0 {\fonttbl {\f0 Courier;';//}}';  // default font
   for i := 1 to FontTableSize do
   begin
      ortf := ortf + '\f'+toStr(i)+' '+FontTable[i]+';';
   end;
   ortf := ortf + '}}'+#13#10;
// Build color table
   ortf := ortf + '{\colortbl;';
   for i := 1 to ColorTableSize do
   begin
      w := ColorTable[i];
      ortf := ortf + '\red'  +toStr(w and $FF);
      ortf := ortf + '\green'+toStr((w shr 8) and $FF);
      ortf := ortf + '\blue' +toStr((w shr 16) and $FF)+';';
   end;
   ortf := ortf + '}'+#13#10;
// add user generated part
   if cp1251 then urtf := string(WideCp1251(rtf)) else urtf := rtf;
   ortf := ortf + '{\*\generator BRtfWriter 1.5}\viewkind4\f0'+#13#10; //4-Normal View  // '\uc1 unicode bytes ??
   ortf := ortf + urtf;
//Finalize
   Result := ortf + '}';
end;

//------------------------------------------------------------------------------
function    BTRtfWriter.Generate(const FileName:string):boolean;
begin
   Result := FileSave(FileName,ansistring(GetRTFdoc));
end;

//------------------------------------------------------------------------------
function    BTRtfWriter.SetupColorTable(R,G,B:longword):longword;
begin
   Result := ColorTableSize + 1;
   if Result <= 32 then
   begin
      ColorTable[Result] := (R and $FF) or ((G and $FF) shl 8) or ((B and $FF) shl 16);
      ColorTableSize := Result;
   end else Result := 0; //def
end;

//------------------------------------------------------------------------------
function    BTRtfWriter.SetupFontTable(const FntName:string):longword;
begin
   Result := FontTableSize + 1;
   if Result <= 16 then
   begin
      FontTable[Result] := FntName;
      FontTableSize := Result;
   end else Result := 0; //def
end;

//------------------------------------------------------------------------------
procedure   BTRtfWriter.SetFontSize(sz:longword);
begin
   rtf := rtf + '\fs' + toStr(sz*2)+' ';   // \fs - font size in half points  -> *2
end;

//------------------------------------------------------------------------------
procedure   BTRtfWriter.SetFont(id:longword);
begin
   if id = 0 then id := 1;
   if id > ColorTableSize then id := 1;
   rtf := rtf + '\f' + toStr(id)+' ';
end;

//------------------------------------------------------------------------------
procedure   BTRtfWriter.SetColor(id:longword);
begin
   if not BWmode then
   begin
      if id = 0 then id := 1;
      if id > ColorTableSize then id := 0;
      rtf := rtf + '\cf' + toStr(id)+' ';
   end;
end;

//------------------------------------------------------------------------------
procedure   BTRtfWriter.SetBgColor(id:longword);
begin
   if not BWmode then
   begin
      if id > ColorTableSize then id := 0;
      rtf := rtf + '\highlight' + toStr(id)+' ';
   end;
end;

//------------------------------------------------------------------------------
procedure   BTRtfWriter.SetBold(state:boolean);
begin
   if state  then rtf := rtf + '\b '
             else rtf := rtf + '\b0 ';
end;

//------------------------------------------------------------------------------
procedure   BTRtfWriter.SetItalic(state:boolean);
begin
   if state  then rtf := rtf + '\i '
             else rtf := rtf + '\i0 ';
end;

//------------------------------------------------------------------------------
procedure   BTRtfWriter.SetUnderline(state:boolean);
begin
   if state  then rtf := rtf + '\u '
             else rtf := rtf + '\u0 ';
end;


//------------------------------------------------------------------------------
procedure   BTRtfWriter.SetFirstLineIdent(cm:single);
begin
   if cm = 0 then pdimen := '\fi0'          // in twips
             else pdimen := '\fi'+toStr(round( ((72/2.54)*cm)*20 ));
end;

//------------------------------------------------------------------------------
procedure   BTRtfWriter.SetTextAlign(typ:longword);
begin
   if typ = 1 then palign := '\ql';   // left
   if typ = 2 then palign := '\qc';   // center
   if typ = 3 then palign := '\qr';   // right
   if typ = 4 then palign := '\qj';   // both (left right justify)
end;


//------------------------------------------------------------------------------
procedure   BTRtfWriter.AddNewPage(landscape:boolean = false);
begin
   if pg <> 0 then rtf := rtf + '\page ';
   if landscape then
   begin
      rtf := rtf + '\landscape ';
   end;
   inc(pg);
end;

//------------------------------------------------------------------------------
procedure   BTRtfWriter.AddNewPageA4(landscape:boolean = false);
begin
   A4 := true;
   AddNewPage(landscape);
end;


//------------------------------------------------------------------------------
const RaragrBegin:string = '\pard\slmult1\sa200\sl276';
procedure   BTRtfWriter.AddNewLine;
begin
   rtf := rtf + RaragrBegin + pdimen + palign + ' '+ parag + '\par'+#13#10;
   parag := '';
   palign := '\ql';
   pdimen := '\fi0';
   paragtouch := false;
end;


//------------------------------------------------------------------------------
procedure   BTRtfWriter.AddText(const Txt:string);
var s1:string;
    i,j,c:longword;
begin
   s1 := '';
   j := length(Txt);
   for i  := 1 to j do
   begin
      c := longword(Txt[i]);
      if c > 127 then
      begin
         if cp1251 then
         begin
            s1 := s1 + char(c); // will be recoded in generation
         end else begin
            // Unicode dirct
            s1 := s1 + '\u'+toStr(c)+'  '; //2 space ???? :(
         end;
      end else begin
         if c >= 32 then
         begin
            if (c = $5C) or (c = $7D) or (c = $7B) then // \ } {
            begin
               s1 := s1 + '\'
            end;
            s1 := s1 + char(c);
         end;
      end;
   end;
   ParagTouch := true;
   parag := parag + s1;
end;

//------------------------------------------------------------------------------
procedure   BTRtfWriter.SetupCodePage(const Page:string);
begin
   if Page = 'cp1251' then cp1251 := true;
end;

//------------------------------------------------------------------------------
function  unicodetocp1251(const txt:string):ansistring;
var i,j,c,w:longword;
begin
   Result := '';
   j := Length(txt);
   if j > 0 then
   begin
      SetLength(Result,j);
      for i := 1 to j do
      begin
         c := longword(txt[i]);
         if c > 127 then
         begin
            w := c;
            if (c >= $410) and (c <= $44F) then c := 192 + (c - $410)
                                           else c := 63; //'?'
            //todo for other special chars
            if w = $2116 then c:= 185; //nomer
         end;
         Result[i] := ansichar(c);
      end;
   end;
end;

// micro GDI + to load eny format picture jpg bmp png
type
   TGPGdiplusStartupInput = packed record
      GdiplusVersion          : Cardinal;  // Must be 1
      DebugEventCallback      : pointer; //TGPDebugEventProc;
      SuppressBackgroundThread: boolean; //BOOL;
      SuppressExternalCodecs  : boolean; //BOOL;
   end;
   PGdiplusStartupInput = ^TGPGdiplusStartupInput;


   TGPGdiplusStartupOutput = packed record
      NotificationHook  : pointer; //TGPNotificationHookProc;
      NotificationUnhook: pointer; //TGPNotificationUnhookProc;
   end;
   PGdiplusStartupOutput = ^TGPGdiplusStartupOutput;


function GdiplusStartup(out token: longword {ULONG}; input: PGdiplusStartupInput; output: PGdiplusStartupOutput) : longint; stdcall;  external 'gdiplus.dll' name 'GdiplusStartup';
procedure GdiplusShutdown(token: longword{ULONG}); stdcall;   external 'gdiplus.dll' name 'GdiplusShutdown';
function GdipLoadImageFromFileICM(filename: PWideCHAR; out image: nativeUint): longint; stdcall; external 'gdiplus.dll' name 'GdipLoadImageFromFileICM';
function GdipBitmapGetPixel(bitmap: nativeUint; x: Integer; y: Integer;  var color: longword): longint; stdcall;  external 'gdiplus.dll' name 'GdipBitmapGetPixel';
function GdipGetImageWidth(bitmap: nativeUint; var width: longint {UINT}): longint {GPSTATUS}; stdcall;  external 'gdiplus.dll' name 'GdipGetImageWidth';
function GdipGetImageHeight(bitmap: nativeUint; var height: longint): longint; stdcall; external 'gdiplus.dll' name 'GdipGetImageHeight';

const BM_INFO :array[0..12] of longword =
      (  40,  //bmiHeader.biSize cardinal = 40
         0,  //bmiHeader.biWidth integer = Xlng
         0,  //bmiHeader.biHeight integer = -Ylng >> start from 0,0 left,top
         $00180001,  //bmiHeader.biPlanes = 1 word   , biBitCount = 24 word   $18 =24 $20=32
         0,  //bmiHeader.biCompression cardinal = 3  BI_BITFIELDS  // was 3 ????? did not work in rtf ????
         0,  //bmiHeader.biSizeImage  cardinal
         0,  //bmiHeader.biXpelsPerMeter integer
         0,  //bmiHeader.biYpelsPerMeter integer
         0,  //bmiHeader.biClrUsed cardinal
         0,  //bmiHeader.biClrImportant cardinal
         $0000FF,  //bmiColors[0]
         $00FF00,  //bmiColors[1]
         $FF0000   //bmiColors[2]
       );

function    BTRtfWriter.AddPicture(const Filename:string; xlng_cm,ylng_cm:single):boolean;
var BMINFO :array[0..12] of longword;
    w,i,alg,algh:longword;
    x,y,bh,bw:longint;
 //   s:ansistring;
   fn:widestring;
//   wi:longint;
   bbm:nativeUint;
   StartupInput :TGPGdiplusStartupInput;
   gdiplusToken :longword;
   StartupOutput :TGPGdiplusStartupOutput;
begin
   Result := false;
   // Initialize GDI+
   StartupInput.DebugEventCallback := nil;
   StartupInput.SuppressBackgroundThread := True;
   StartupInput.SuppressExternalCodecs   := False;
   StartupInput.GdiplusVersion := 1;
   if GdiplusStartup(gdiplusToken, @StartupInput, @StartupOutput) = 0 then
   begin
      fn := widestring(filename) + #0;
      if  GdipLoadImageFromFileICM(@fn[1],bbm) = 0 then
      begin
         // get picture dimentions
         GdipGetImageWidth(bbm,bw);
         GdipGetImageHeight(bbm,bh);

(*
default dpi for RTF is 72 from apple
twips = 1/20 of pixel 96 dpi = 37,9 pixels in 1 cm 756 twips in cm
                      72 dpi = 28,3 pixels in 1 cm 567 twips in cm
the data is infobitmap header + data
*)
//xlng must be even !!!!!! width must be 4 byte align

         algh := (bw*3) and $FFFFFFFC; // for bye align;
         alg := (bw*3) and 3;
         if alg <> 0 then inc(algh);

         parag  := parag + '{\pict\dibitmap0\wbmbitspixel24\wbmplanes1\wbmwidthbytes'+toStr((algh))

         +'\picw'+toStr(bw)+'\pich'+toStr(bh)+'\picwgoal'+toStr({twips}round(xlng_cm*567))+'\pichgoal'+toStr(round(ylng_cm*567))+' '+#13#10;

         //copy in hex BITAMPINFO structure
         for i := 0 to 12 do BMINFO[i] := BM_INFO[i];
         BMINFO[1]:= bw;
         BMINFO[2]:= bh; //was -   did not work with negative value in rtf ??????
         for i := 0 to 9 do parag := parag + toHex( BMINFO[i]         and $FF,2)
                                           + toHex((BMINFO[i] shr 8 ) and $FF,2)
                                           + toHex((BMINFO[i] shr 16) and $FF,2)
                                           + toHex((BMINFO[i] shr 24) and $FF,2);

         // copy in hex DATA 24 bit
         for y := bh-1 downto 0 do
         begin
            for x := 0 to bw-1 do
            begin
               GdipBitmapGetPixel(bbm,x,y,w);
               parag := parag + toHex( w         and $FF,2)
                              + toHex((w shr 8 ) and $FF,2)
                              + toHex((w shr 16) and $FF,2);
            end;
            if alg <> 0 then for i := 3 downto alg do parag := parag + '00';
         end;

         parag := parag + '}';
         ParagTouch := true;
         GdiplusShutdown(gdiplusToken);
         Result := true;
      end;
   end;
end;


procedure   BTRtfWriter.AddTable;
begin
   ParagTouch := true;
   parag := parag + '\trowd\cellx1000\cellx2000\cellx3000\intbl mama\cell\intbl kaka\cell\intbl baba\cell\row '
end;

end.
