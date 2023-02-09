unit BpdfWriter;

interface
{
//TODO//
   pay matrix
   text cyrilic
   text in box auto calculate space
   zip stream
}

type  BTPDFdoc = class
         private
            pdf :ansistring;
            xref :ansistring;
            pages :ansistring;
            pagesobj :ansistring;
            stream :ansistring;
            curfont :ansistring;
            font :ansistring;
            aFontCount :longword;
            aFontResource :ansistring;
            aCurFontSize :single;
            aObjectsCount :longword;
            aOffset :longword;
            aPagesCount :longword;
            aWidth :longword;
            aHeight :longword;
            aPen :boolean;
            aPenColor :longword;
            aPenSet :ansistring;
            aPenWidth :single;
            aPenDash :longword;
            aBrush :boolean;
            aBrushColor :longword;
            aBMPcount :longword;
            aStrokeEnd :ansistring;
            a1251 :boolean;
            a1251index :longword;
            aIncludeFont :boolean;
            aIMAGEsmooth :boolean;
            aXobject :ansistring;
            aBMPdata:ansistring;
            aZIPit :boolean;
            function    _RegisterFont(const name:string; size:single; bold,italic,std:boolean):longword;
            function    _Add1251table:longword;
            procedure   _Arc(x,y,ray,ang1,ang2:single; flg:boolean);
            function    _NextParm(var main:ansistring):ansistring;
            procedure   _NormalizeCord(var x1,y1,x2,y2:single);
            procedure   _SetStrokeEnd;
            procedure   _EndPage;
            procedure   _Additem(s:ansistring);
            procedure   _AddOffset;
            function    _AddObject:longword;
            procedure   _EndObject;
            function    _ZIPer(const a:ansistring):ansistring;
         public
            constructor Create;
            destructor  Destroy; override;
            procedure   Clear;
            procedure   AddPage(Width,Height:longword);
            procedure   AddPageA4(landscape:boolean);
            procedure   SetPen(enable:boolean; color:longword; width:single; LineDash:longword);
            procedure   SetBrush(enable:boolean; color:longword);
            procedure   EnablePenBrush(p,b:boolean);
            function    RegisterFont(const name:string; size:single; bold,italic:boolean):longword;
            function    RegisterTTFFont(const name:string; size:single; bold,italic:boolean):longword;
            procedure   SetFont(hFont:longword; size:single);
            procedure   Line(x1,y1,x2,y2:single);
            procedure   Rectangle(x1,y1,x2,y2:single);
            procedure   RectangleC(x1,y1,w,h:single);
            procedure   Ellipse(x1,y1,x2,y2:single);
            procedure   EllipseC(x1,y1,rx,ry:single);
            procedure   Polygon(data:pointer; count:longword);
            procedure   Pay(x,y,ray,ang1,ang2:single);
            function    RegisterBitmapRaw(xlng,ylng,pitch,bpp,mask:longword; data:pointer):longword;
            function    RegisterBitmap(hBitmap,mask:longword):longword;
            procedure   DrawBitmap(x,y:single; xlng,ylng,a:longword; Image:longword);
            procedure   Text(x,y:single; const s:ansistring);
            procedure   TextBox(x,y,xl,yl:single; s:ansistring);
            function    Save(filename:string):boolean;
            property    PageWidth :longword read aWidth;
            property    PageHeight :longword read aHeight;
      end;


implementation

uses BFileTools,BPasZlib,Windows;


function ast(w:longint):ansistring; inline;
begin
   str(w,result);
end;

function fst(w:single):ansistring; inline;
var i:longword;
begin
   str(w:0:2,Result);
   i :=Pos(',',string(Result));
   if i > 0  then Result[i] := '.';
   i := length(Result);
   if (Result[i] = '0') and (Result[i-1] = '0') then SetLEngth(Result,i-3);
end;

//------------------------------------------------------------------------------
constructor BTPDFdoc.Create;
begin
   Clear;
end;

//------------------------------------------------------------------------------
destructor  BTPDFdoc.Destroy;
begin

   inherited
end;

//------------------------------------------------------------------------------
procedure   BTPDFdoc.Clear;
begin
   pdf := '';
   xref := '';
   pages := '';
   pagesobj := '';
   stream := '';
   font := '';
   curfont := '';
   aPenSet := '';
   aFontCount := 0;
   aOffset := 0;
   aObjectsCount := 0;
   aPagesCount := 0;
   _Additem('%PDF-1.4');
   _Additem('%'+#183#190#173#170);
   _AddObject; // info object must be 1 in my case
   _Additem('<< /Producer (BPDF generator 2014)');
   _Additem('>>');
   _EndObject;
   aPen := false;
   aBrush := false;
   _SetStrokeEnd;
   aPenColor := $FFFFFFFF;
   aPenWidth := 1000;
   aPenDash := $FFFFFFFF;
   aBrushColor := $FFFFFFFF;
   a1251 := false;
   a1251index := 0;
   aIMAGEsmooth := false;
   aBMPcount := 0;
   aXobject := '';
   aBMPdata := '';
   aFontResource := '';
   aIncludeFont := true;
   aZIPit := true;
   RegisterFont('Times',14,false,false); // set default
end;


//------------------------------------------------------------------------------
function    BTPDFdoc._ZIPer(const a:ansistring):ansistring;
var i:longint;
begin
   i := length(a);
   SetLength(Result,i);
   CompressMem(@a[1],@Result[1],i,i);
end;

//------------------------------------------------------------------------------
procedure   BTPDFdoc._EndPage;
var i:longword;
    s:ansistring;
begin
   if aPagesCount > 0 then
   begin  // finish last page
      i := _AddObject;
      PagesObj := PagesObj + ast(i) + '|';
      // next object is stream len
      s := '<< /Length ' + ast(i+1) + ' 0 R ';
      if aZIPit then s := s + #10 + '/Filter [ /FlateDecode ]' + #10;
      _AddItem(s + '>>');
      _Additem('stream');

      if aZIPit then  stream := _ZIPer(stream);

      pdf := pdf + stream;
      inc(aOffset,length(stream));
      _Additem('endstream');
      _EndObject;
      _AddObject;
      _Additem(ast(length(stream)));
      _EndObject;
   end;
end;

//------------------------------------------------------------------------------
procedure   BTPDFdoc.AddPageA4(landscape:boolean);
begin
   if landscape then AddPage(1123,794)
                else AddPage(794,1123); // 794 x 1123 @ 96 DPI
end;

//------------------------------------------------------------------------------
procedure   BTPDFdoc.AddPage(Width,Height:longword);
begin
   _EndPage;

   aWidth := Width;
   aHeight := Height;
   inc(aPagesCount);
   Pages := Pages + '/MediaBox [ 0 0 '+ast(Width)+' '+ast(Height)+']'+#10+ '|'; // end marker
   //start new page
   stream := '';
end;


//------------------------------------------------------------------------------
procedure   BTPDFdoc._Additem(s:ansistring);
begin
   if s[length(s)] <> #10 then s := s + #10;
   pdf := pdf + s;
   inc(aOffset,length(s));
end;

//------------------------------------------------------------------------------
function    BTPDFdoc._AddObject:longword;
begin
   inc(aObjectsCount);
   Result := aObjectsCount;
   _Addoffset;
   _Additem(ast(Result)+' 0 obj');
end;

//------------------------------------------------------------------------------
procedure   BTPDFdoc._EndObject;
begin
   _Additem('endobj');
end;

//------------------------------------------------------------------------------
procedure   BTPDFdoc._AddOffset;  //todo put in add object
var l,i:LongInt;
    o,s:Ansistring;
begin
   str(aOffset,o);
   l := length(o);
   s := '';
   for i := 1 to 10-l do s := s + '0';
   s := s + o;
   xref := xref + s + ' 00000 n'+#13#10;
end;


//------------------------------------------------------------------------------
function    BTPDFdoc.RegisterBitmap(hBitmap,mask:longword):longword;
var BitmapInfo :TBitmapInfo;
    DeviceContext :hDC;
    bm:BITMAP;
    Width,Height,i:longword;
    Data,rData:pointer;
begin
   Result := 0;
   GetObject(hBitmap,sizeof(BITMAP),@bm);
   Width := bm.bmWidth;
   Height := bm.bmHeight;
  // GetDIBSizes(hBitmap, InfoSize, ImageSize);
   with BitmapInfo do
   begin
      bmiHeader.biSize        := Sizeof(BitmapInfo);
      bmiHeader.biWidth       := Width;
      bmiHeader.biHeight      := Height;
      bmiHeader.biPlanes      := 1;
      bmiHeader.biBitCount    := 24;
      bmiHeader.biCompression := BI_RGB;
   end;
   DeviceContext := GetDC(0);
   Data := nil;
   ReallocMem(Data,Width*Height*4+Width);
   if data <> nil  then
   begin
      rdata := data;
      for i := Height-1 downto 0 do
      begin
         if GetDIBits (DeviceContext, hBitmap, i, 1, Data, BitmapInfo, DIB_RGB_COLORS) = 1 then
         begin
            Data := pointer(longword(Data) + Width*3);
         end;
      end;
      Result := RegisterBitmapRaw(Width,Height,Width*3,24,mask,rData);
      ReallocMem(rData,0);
   end;
   ReleaseDC(0, DeviceContext);
end;

//------------------------------------------------------------------------------
// !warning!  data is in format BGR
function    BTPDFdoc.RegisterBitmapRaw(xlng,ylng,pitch,bpp,mask:longword; data:pointer):longword;
var j,r,g,b,mz,mzx:longword;
    p:pointer;
    c:^byte;
    bit,cbit:byte;
    mimage_indx :longword;
    a,pic:ansistring;
    m:boolean;
    color,puter,mputer:longword;
    la,i :longint;
begin
   inc(aBMPcount);
   Result := aBMPcount;
   SetLength(aBMPdata,aBMPcount*8);
   r := (aBMPcount-1)*8;
   longword(pointer(@aBMPdata[r])^) := xlng;
   longword(pointer(@aBMPdata[r+4])^) := ylng;
   m := false;
   mzx :=(xlng div 8) + 1;
   mz := mzx * ylng;
   SetLength(a,mz);  // fuck compiler inside if (mask -> KRASH :(

   mimage_indx := 0;
   if (mask and $FF000000) <> 0 then
   begin
      // prepare mask
      p := Data;
      mputer := 1;
      for r:= 1 to ylng do
      begin
         la := 7;
         cbit := 0;
         puter := mputer;

         c := p;
         for g:= 1 to xlng do
         begin
            color := c^;  // blue
            c := pointer(longword(c) + 1);
            color := color or (longword(c^) shl 8);  // green
            c := pointer(longword(c) + 1);
            color := color or (longword(c^) shl 16);  // green
            c := pointer(longword(c) + 1);
            if bpp = 32 then c := pointer(longword(c) + 1);

            if (mask and $FFFFFF) = color then bit := 1 // clear
                                          else bit := 0; // draw;
            bit := bit shl la;
            cbit := cbit or bit;
            dec(la);
            a[puter] := ansichar(cbit);
            if la < 0  then
            begin
               la := 7;
               cbit := 0;
               inc(puter);
            end;
         end;
         inc(mputer,mzx);
         p := pointer(longword(P) + pitch);
      end;

      m := true;
      inc(aBMPcount);
      mimage_indx := _AddObject;
      aXobject := aXobject + '/Im'+ast(aBMPcount)+ ' ' + ast(mimage_indx) +' 0 R'+#10; // resource list
      _Additem('<< /Type /XObject');
      _Additem('/Length '+ast(length(a))); //ast(mz));
      _Additem('/Subtype /Image');
      _Additem('/ColorSpace /DeviceGray');
      _Additem('/Width '+ast(xlng));
      _Additem('/Height '+ast(ylng));
      _Additem('/BitsPerComponent 1');
      _Additem('/ImageMask true');
      if aZIPit then _Additem('/Filter [ /FlateDecode ]');
      _Additem('>>');
      _Additem('stream');
      if aZIPit then a := _ZIPer(a);

      pdf := pdf + a;
      inc(aOffset,length(a));
      _Additem('endstream');
      _EndObject;
   end;

   // prepare picture
   p := Data;
   setlength(a,xlng*3);
   pic := '';
   for r:= 1 to ylng do
   begin
      c := p;
      for g:= 1 to xlng do
      begin
         b := (g-1)*3;
         a[b + 3] := ansichar(c^);
         c := pointer(longword(c) + 1);
         a[b + 2] := ansichar(c^);
         c := pointer(longword(c) + 1);
         a[b + 1] := ansichar(c^);
         c := pointer(longword(c) + 1);
         if bpp = 32 then c := pointer(longword(c) + 1);
      end;
      pic := pic + a;
      p := pointer(longword(P) + pitch);
   end;

   j := _AddObject;
   aXobject := aXobject + '/Im'+ast(Result)+ ' ' + ast(j) +' 0 R'+#10; // resource list

   _Additem('<< /Type /XObject');
   _Additem('/Length '+ast(length(pic))); //ast(xlng*ylng*3));
   _Additem('/Subtype /Image');
   _Additem('/ColorSpace /DeviceRGB');
   _Additem('/Width '+ast(xlng));
   _Additem('/Height '+ast(ylng));
   _Additem('/BitsPerComponent 8');
   if aIMAGEsmooth then _Additem('/Interpolate true');
   if m then _Additem('/Mask '+ast(mimage_indx)+' 0 R');
   if aZIPit then _Additem('/Filter [ /FlateDecode ]');
   _Additem('/Name /Im'+ast(Result));
   _Additem('>>');
   _Additem('stream');



   if aZIPit then pic := _ZIPer(pic);

   pdf := pdf + pic;

   inc(aOffset,length(pic));

   _Additem('endstream');
   _EndObject;
end;


//------------------------------------------------------------------------------
function    BTPDFdoc._NextParm(var main:ansistring):ansistring;
var a:ansistring;
    k:longword;
begin
   k := Pos('|',string(main));
   Result := Copy(main, 1, k - 1);
   a := Copy(main, k + 1, longword(length(main)) - k);
   main := a;
end;


//------------------------------------------------------------------------------
function    BTPDFdoc.Save(filename:string):boolean;
var xref_start:longword;
    Pages_indx:longword;
    Catalog_indx:longword;
    Resource_indx:longword;
//    Outln_indx:longword;
    i,j,m :longword;
    p :ansistring;
begin
   Result := false;
   if aPagesCount = 0 then Exit;

   _EndPage; // terminate last page


   // Create object <RESOURCE>
   Resource_indx := _AddObject;
   _Additem('<< /ProcSet [ /PDF /Text /ImageB /ImageC /ImageI ]');
   if aFontCount <> 0  then _additem('/Font <<'+#10+ aFontResource +'>>');
   if aBMPcount <> 0 then _additem('/XObject <<'+#10+ aXObject +'>>');

   _Additem('>>');
   _EndObject;

   p := '';
   m := aObjectsCount + aPagesCount + 1; //this will be pages object
   for i := 1 to aPagesCount do
   begin
      // Create object <PAGES>
      j := _AddObject;
      P := P + ' ' + ast(j) + ' 0 R';
      _Additem('<< /Type /Page');
      _Additem(_NextParm(Pages));
      _Additem('/Contents ' + _NextParm(PagesObj) + ' 0 R');
      _Additem('/Resources ' + ast(Resource_indx) + ' 0 R');
      _Additem('/Parent '+ast(m)+' 0 R');
      _Additem('>>');
      _EndObject;
   end;

   // Create object <PAGES>
   Pages_indx := _AddObject;
   _Additem('<< /Type /Pages');
   _Additem('/Kids [' + P + ' ]');
   _Additem('/Count ' + ast(aPagesCount) );
   _Additem('>>');
   _EndObject;


//   // Create object <PAGES>
//   Outln_indx := _AddObject;
//   _Additem('<< /Type /Outlines');
//   _Additem('/Count 0');
//   _Additem('>>');
//   _EndObject;


   // Create object <CATALOG>
   Catalog_indx := _AddObject;
   _Additem('<< /Type /Catalog');
   _Additem('/Pages ' + ast(Pages_indx) + ' 0 R');
//   _Additem('/Outlines ' + ast(Outln_indx) + ' 0 R');
   _Additem('>>');
   _EndObject;

   // Create object <XREFERENCE>
   inc(aObjectsCount);
   xref_start := aOffset ; // mark the start of xref
   _Additem('xref');  // build xref table
   _Additem('0 ' + ast(aObjectsCount));
   _Additem('0000000000 65535 f'+#13);
   _Additem(xref);
   _Additem('trailer');
   _Additem('<< /Root '+ast(Catalog_indx)+' 0 R');
   _Additem('/Info 1 0 R');
   _Additem('/Size '+ ast(aObjectsCount));
   _Additem('>>');
   _Additem('startxref');
   _Additem( ast(xref_start));
   _Additem('%%EOF'); // end of file
   Result := FileSave(filename,pdf);

end;

//------------------------------------------------------------------------------
procedure   BTPDFdoc._SetStrokeEnd;
begin
   if     aPen and     aBrush then aStrokeEnd := 'b'+#10;
   if not aPen and     aBrush then aStrokeEnd := 'f'+#10;
   if     aPen and not aBrush then aStrokeEnd := 's'+#10;
   if not aPen and not aBrush then aStrokeEnd := 'n'+#10;
end;

//------------------------------------------------------------------------------
procedure   BTPDFdoc.SetPen(enable:boolean; color:longword; width:single; LineDash:longword);
begin
   aPen := enable;
   if aPen then
   begin
      if aPenColor <> color then
      begin
      // Set color state
         aPenSet :=
         fst( (color         and $ff) / 255 ) + ' ' +     //R
         fst(((color shr  8) and $ff) / 255 ) + ' ' +     //G
         fst(((color shr 16) and $ff) / 255 ) + ' RG'+#10;//B
         stream := stream + aPenSet;
         aPenColor := color;
      end;
      if aPenWidth <> width then
      begin
         // set width state
         stream := stream +
         fst(width) + ' w'+#10;
         aPenWidth := width;
      end;

//      stream := stream +
//      '[2 1] 0 d'+#10;
      aPenDash := LineDash;
   end;
   _SetStrokeEnd;
end;

//------------------------------------------------------------------------------
procedure   BTPDFdoc.SetBrush(enable:boolean; color:longword);
begin
    aBrush := enable;
    if aBrush then
    begin
       if aBrushColor <> color then
       begin
          // Set color state
          stream := stream +
          fst( (color         and $ff) / 255 ) + ' ' +     //R
          fst(((color shr  8) and $ff) / 255 ) + ' ' +     //G
          fst(((color shr 16) and $ff) / 255 ) + ' rg'+#10;//B
          aBrushColor := color;
       end;
    end;
    _SetStrokeEnd;
end;

//------------------------------------------------------------------------------
procedure   BTPDFdoc.EnablePenBrush(p,b:boolean);
begin
   aPen := p;
   aBrush := b;
end;

//------------------------------------------------------------------------------
procedure   BTPDFdoc.Line(x1,y1,x2,y2:single);
begin
   y1 := aHeight - y1;
   y2 := aHeight - y2;
   stream := stream +
   fst(x1)+' '+fst(y1)+' m'+ #10 +
   fst(x2)+' '+fst(y2)+' l'+ #10 +
   aStrokeEnd;
end;

//------------------------------------------------------------------------------
procedure   BTPDFdoc._NormalizeCord(var x1,y1,x2,y2:single);
var a:single;
begin
   if y1 > y2 then begin a := y2; y2 := y1; y1 := a; end;
   if x1 > x2 then begin a := x2; x2 := x1; x1 := a; end;
   a := y2 - y1; // Y size
   y1 := aHeight - y1 - a ;
   y2 := aHeight - y2 + a ;
end;

//------------------------------------------------------------------------------
procedure   BTPDFdoc.Ellipse(x1,y1,x2,y2:single);
var xm,ym,xk,yk:single;
const kappa = 0.552;
begin
   _NormalizeCord(x1,y1,x2,y2);
   xm := abs(x2 - x1) / 2;
   ym := abs(y2 - y1) / 2;
   xk := xm * kappa;
   yk := ym * kappa;
   ym := y1 + ym;
   xm := x1 + xm;
   stream := stream +
   fst(x1)+' '+fst(ym)+' m'+#10+
   fst(x1)+' '+fst(ym+yk)+' '+fst(xm-xk)+' '+fst(y2)+' '+fst(xm)+' '+fst(y2)+' c'+#10+
   fst(xm+xk)+' '+fst(y2)+' '+fst(x2)+' '+fst(ym+yk)+' '+fst(x2)+' '+fst(ym)+' c'+#10+
   fst(x2)+' '+fst(ym-yk)+' '+fst(xm+xk)+' '+fst(y1)+' '+fst(xm)+' '+fst(y1)+' c'+#10+
   fst(xm-xk)+' '+fst(y1)+' '+fst(x1)+' '+fst(ym-yk)+' '+fst(x1)+' '+fst(ym)+' c'+#10+
   aStrokeEnd;
end;

//------------------------------------------------------------------------------
procedure   BTPDFdoc.EllipseC(x1,y1,rx,ry:single);
begin
   Ellipse(x1-rx,y1-ry,x1+rx,y1+ry);
end;


//------------------------------------------------------------------------------
procedure   BTPDFdoc._Arc(x,y,ray,ang1,ang2:single; flg:boolean);
var
   rx0, ry0, rx1, ry1, rx2, ry2, rx3, ry3 :single;
   x0, y0, x1, y1, x2, y2, x3, y3 :single;
   delta_angle, new_angle, co, si :single;

begin
    delta_angle := (90 - (ang1 + ang2) / 2) / 180 * PI;
    new_angle := (ang2 - ang1) / 2 / 180 * PI;
    rx0 := ray * cos (new_angle);
    ry0 := ray * sin (new_angle);
    rx2 := (ray * 4.0 - rx0) / 3.0;
    ry2 := ((ray * 1.0 - rx0) * (rx0 - ray * 3.0)) / (3.0 * ry0);
    rx1 := rx2;
    ry1 := -ry2;
    rx3 := rx0;
    ry3 := -ry0;
    co := cos(delta_angle);
    si := sin(delta_angle);
    x0 := rx0 * CO - ry0 * SI + x;
    y0 := rx0 * SI + ry0 * CO + y;
    x1 := rx1 * CO - ry1 * SI + x;
    y1 := rx1 * SI + ry1 * CO + y;
    x2 := rx2 * CO - ry2 * SI + x;
    y2 := rx2 * SI + ry2 * CO + y;
    x3 := rx3 * CO - ry3 * SI + x;
    y3 := rx3 * SI + ry3 * CO + y;
    if not flg then
    begin
       stream := stream + fst(x0) + ' ' + fst(y0) + ' m'+#10;
    end;
    stream := stream +
    fst(x1) + ' ' + fst(y1) + ' '+
    fst(x2) + ' ' + fst(y2) + ' '+
    fst(x3) + ' ' + fst(y3) + ' c' + #10;
end;
//------------------------------------------------------------------------------
procedure   BTPDFdoc.Pay(x,y,ray,ang1,ang2:single);
var flg:boolean;
    tmp_ang:single;
begin
   y := aHeight*0.5 - y;
   flg := false;

   if (ang1 >= ang2) or ((ang2 - ang1) >= 360) then Exit;
   while (ang1 < 0) or (ang2 < 0) do
   begin
      ang1 := ang1 + 360;
      ang2 := ang2 + 360;
   end;
   stream := stream + 'q'+#10+
   '1 0 0 .5 0 0 cm'+#10+
   fst(x) + ' ' + fst(y) + ' m'+#10;
   repeat
      if (ang2 - ang1 <= 90) then
      begin
         _Arc (x, y, ray, ang1, ang2, flg);
         break;
      end else begin
         tmp_ang := ang1 + 90;
         _Arc (x, y, ray, ang1, tmp_ang, flg);
         ang1 := tmp_ang;
      end;
      flg := true;
   until (ang1 >= ang2);
   stream := stream + fst(x) + ' ' + fst(y) + ' l'+#10+
   aStrokeEnd + 'Q'+#10;
end;

//------------------------------------------------------------------------------
procedure   BTPDFdoc.Rectangle(x1,y1,x2,y2:single);
var xm,ym:single;
begin
   _NormalizeCord(x1,y1,x2,y2);
   xm := x2 - x1;
   ym := y2 - y1;
   stream := stream +
   fst(x1)+' '+fst(y1)+' '+fst(xm)+' '+fst(ym)+' re'+#10+
   aStrokeEnd;
end;

//------------------------------------------------------------------------------
procedure   BTPDFdoc.RectangleC(x1,y1,w,h:single);
begin
   Rectangle(x1,y1,x1+w-1,y1+h-1);
end;

//------------------------------------------------------------------------------
type _s_a = array [0..0,0..1] of single;
procedure   BTPDFdoc.Polygon(data:pointer; count:longword);
var x0,y0,x1,y1,xs,ys:single;
    i:longword;
    p:^_s_a;
begin
   p := data;
   i := 0;
   x0 := p[i,0];
   y0 := aHeight - p[i,1];
   xs := x0;
   ys := y0;
   stream := stream +
   fst(x0)+' '+fst(y0)+' m'+ #10;
   for i := 1 to Count - 1 do
   begin
      x1 := p[i,0];
      y1 := aHeight - p[i,1];
      stream := stream +
      fst(x1)+' '+fst(y1)+' l'+ #10;
   end;
   stream := stream +
   fst(xs)+' '+fst(ys)+' l'+ #10 +
   aStrokeEnd;
end;

//------------------------------------------------------------------------------
procedure   BTPDFdoc.DrawBitmap(x,y:single; xlng,ylng,a:longword; Image:longword);
var co,si:single;
    r:longword;
begin
   r := (Image-1)*8;
   if (xlng = 0) then xlng := longword(pointer(@aBMPdata[r])^);
   if (ylng = 0) then ylng := longword(pointer(@aBMPdata[r+4])^);

   y := aHeight - y - ylng ;
   co := cos(a * ( PI / 180));
   si := sin(a * ( PI / 180));

   stream := stream + 'q'+#10 +
   fst(co*xlng) +' '+ fst(si*xlng) +' '+
   fst(-si*ylng) +' '+ fst(co*ylng) +' '+
   fst(x)+ ' ' + fst(y) +
   ' cm'+#10+
   '/Im'+ast(Image)+' Do'+#10+
   'Q'+#10;
end;

//------------------------------------------------------------------------------
function    BTPDFdoc._Add1251table:longword;
var t:ansistring;
begin
   if a1251 then
   begin
      if a1251index = 0 then
      begin
         a1251index := _AddObject;
         _Additem('<< /Type /Encoding');
         _Additem('/BaseEncoding /WinAnsiEncoding');
         t :=  '/Differences [129 /afii10052'
         + '/quotesinglbase/afii10100/quotedblbase/ellipsis/dagger/daggerdbl/Euro'
         + '/perthousand/afii10058/guilsinglleft/afii10059/afii10061/afii10060'
         + '/afii10145/afii10099/quoteleft/quoteright/quotedblleft/quotedblright'
         + '/bullet/endash/emdash/space/trademark/afii10106/guilsinglright'
         + '/afii10107/afii10109/afii10108/afii10193/space/afii10062'
         + '/afii10110/afii10057/currency/afii10050/brokenbar/section/afii10023'
         + '/copyright/afii10053/guillemotleft/logicalnot/hyphen/registered'
         + '/afii10056/degree/plusminus/afii10055/afii10103/afii10098/mu'
         + '/paragraph/periodcentered/afii10071/afii61352/afii10101/guillemotright'
         + '/afii10105/afii10054/afii10102/afii10104/afii10017/afii10018/afii10019'
         + '/afii10020/afii10021/afii10022/afii10024/afii10025/afii10026/afii10027'
         + '/afii10028/afii10029/afii10030/afii10031/afii10032/afii10033/afii10034'
         + '/afii10035/afii10036/afii10037/afii10038/afii10039/afii10040/afii10041'
         + '/afii10042/afii10043/afii10044/afii10045/afii10046/afii10047/afii10048'
         + '/afii10049/afii10065/afii10066/afii10067/afii10068/afii10069/afii10070'
         + '/afii10072/afii10073/afii10074/afii10075/afii10076/afii10077/afii10078'
         + '/afii10079/afii10080/afii10081/afii10082/afii10083/afii10084/afii10085'
         + '/afii10086/afii10087/afii10088/afii10089/afii10090/afii10091/afii10092'
         + '/afii10093/afii10094/afii10095/afii10096/afii10097/space]';
         _Additem(t);
         _Additem('>>');
         _EndObject;
      end;
   end;
   Result := a1251index;
end;

//Warning The font name is case sensitive
function    BTPDFdoc._RegisterFont(const name:string; size:single; bold,italic,std:boolean):longword;

var nf_indx:longword;
    fc_xlng,t,a,f,ff,ffs:ansistring;
    Encoding,i,j :longword;
    pm: ^OUTLINETEXTMETRIC;
    pmsize,firstChar,LastChar:longword;
    h_dc,fnt,fnto:longword;
    good:boolean;
    al,cs,Angle,BLD,FItalic,FUnderline,fSize : dword;
    s:string;
    pwidths,pp: PABC;
    pfont:pointer;
    ppp:pansichar;
    psize:longword;
begin
   Result := 0; //error
   h_dc := 0;
   fnt := 0;
   fnto := 0;
   FirstChar := 0;
   LastChar := 0;

   t := ansistring(name);  //format name
//   for i:= 1 to length(t) do t[i] := Upcase(t[i]);
//   t[1] := ansichar(CharUpper(pansichar(@t[1]))); // make only first car


   if not std then
   begin    // add to true tupe font name
      a1251 := true;
      if bold then t := t + ',Bold';  // add to the name
      if italic then t := t + ',Italic';
   end;

   // test is font allready exist
   i := Pos(t,font);
   if i <> 0 then
   begin
      f := font;
      for i := 1 to aFontCount do
      begin
         a := _NextParm(f);
         j := Pos(t,a);
         if j <> 0 then
         begin
            curfont := '/F'+ast(i);
            aCurFontSize := size;
            Result := i;
            Exit; // done
         end;
      end;
   end;

   pm := nil;
   pwidths := nil;
   pfont := nil;
   psize := 1;

   if not std then
   begin    // for TRUE TYPE fonts - get data for the font

      // get font metrics
      h_dc := GetDC(0);

      // calc font dimentions for 96 dpi
      if aHeight > aWidth then
      begin // portret
         fsize := -MulDiv(aHeight, GetDeviceCaps(h_DC, LOGPIXELSY), 96);
      end else begin
         // landscape
         fsize := -MulDiv(aWidth, GetDeviceCaps(h_DC, LOGPIXELSX), 96);
      end;
      Angle := 0;;
      BLD := 0; //700 - for bold
      FItalic := 0;
      FUnderline := 0;
      al := 0;
      cs := DEFAULT_CHARSET;
      if a1251 then cs := RUSSIAN_CHARSET;

      if Bold  then BLD := 700;
      if Italic then FItalic := 1;
//    if bfsUnderline in aStyle  then FUnderline := 1;
//    if bfsAntialiased in aStyle  then al := ANTIALIASED_QUALITY;
      fnt := CreateFont(fSize,0,Angle,0,BLD,FItalic,FUnderline,0,cs,0,0,al,0,@t[1]);

      fnto := SelectObject(h_dc, fnt);

      good := false;
      pmsize := GetOutlineTextMetrics(h_dc, 0, nil);
      if pmsize <> 0 then
      begin
         ReallocMem(pm,pmsize);
         if pm <> nil then
         begin
            i := GetOutlineTextMetrics(h_dc, pmsize, pm);
            if i <> 0 then
            begin
               good := true;
            end;
         end;
      end;

      firstChar := ord(pm.otmTextMetrics.tmFirstChar);
      lastChar := ord(pm.otmTextMetrics.tmLastChar);

      // alloc place for metrics
      ReallocMem(pwidths,sizeof(TABC)*(lastChar-FirstChar+1));
      if pwidths = nil then good := false;

      if aIncludeFont then
      begin
         // alloc plase for font data
         psize := GetFontData(h_dc, 0, 0, nil, 1);
         ReallocMem(pfont,psize);
         if pfont <> nil then
         begin
            GetFontData(h_dc, 0, 0, pfont, psize);
         end else good := false;
      end;

      if not good then
      begin
         if pfont <> nil then  ReallocMem(pfont,0);
         if pwidths <> nil then ReallocMem(pwidths,0);
         if pm <> nil then ReallocMem(pm,0); //free
         SelectObject(h_dc, fnto);
         DeleteObject(fnt);
         ReleaseDC(0,h_dc);
         Exit;
      end;

   end;  // not std

   SetLength(ff,psize);

   Encoding := _Add1251table;

   inc(aFONTcount);
   nf_indx := _AddObject;

   font := font + t +'|'; // new font add to list
   _Additem('<< /Type /Font');
   _Additem('/Name /F'+ast(aFontcount));
   _Additem('/BaseFont /'+t);
   if std then _Additem('/Subtype /Type1')
          else _Additem('/Subtype /TrueType');
   if not a1251  then _Additem('/Encoding /WinAnsiEncoding')
                 else _Additem('/Encoding '+ast(Encoding) + ' 0 R');

   if not std then
   begin  // TRUE TYPE font need data
      _Additem('/FirstChar '+ast(FirstChar));
      _Additem('/LastChar '+ast(LastChar));
      _Additem('/FontDescriptor '+ast( nf_indx + 1 )+' 0 R'); // next one is descriptor
      fc_xlng := '';
      GetCharABCWidths(h_dc, FirstChar, LastChar, pwidths^);
      pp := pwidths;
      for i:= 0 to (LastChar-FirstChar) do
      begin
         fc_xlng := fc_xlng +ast(pp^.abcA + Integer(pp^.abcB) + pp^.abcC) + ' ';
         pp := pointer(longword(pp) + sizeof(tABC));
      end;
      _Additem('/Widths [ '+fc_xlng+' ]');
   end;

   _Additem('>>');
   _EndObject;    // end of the font object

   if not std then
   begin  // TRUE TYPE font need data
      // Descriptor
      _AddObject; // + 1
      _Additem('<< /Type /FontDescriptor');
      _Additem('/FontName /'+t);
      _Additem('/Flags 32');
      _Additem('/FontBBox [' + ast(pm^.otmrcFontBox.Left) + ' '+
                               ast(pm^.otmrcFontBox.Bottom) + ' '+
                               ast(pm^.otmrcFontBox.Right) + ' '+
                               ast(pm^.otmrcFontBox.Top) + ' ]');
      _Additem('/ItalicAngle ' + ast(pm^.otmItalicAngle));
      _Additem('/Ascent ' + ast(pm^.otmAscent));
      _Additem('/Descent ' + ast(pm^.otmDescent));
      _Additem('/Leading ' + ast(pm^.otmTextMetrics.tmInternalLeading));
      _Additem('/CapHeight ' + ast(pm^.otmTextMetrics.tmHeight));
      _Additem('/StemV ' + ast(50 + Round(sqr(pm^.otmTextMetrics.tmWeight / 65))));
      _Additem('/AvgWidth ' + ast(pm^.otmTextMetrics.tmAveCharWidth));
      _Additem('/MaxWidth ' + ast(pm^.otmTextMetrics.tmMaxCharWidth));
      _Additem('/MissingWidth ' + ast(pm^.otmTextMetrics.tmAveCharWidth));
      if aIncludeFont then
      begin
         _Additem('/FontFile2 '+ast(nf_indx + 2)+' 0 R');
      end;
      _Additem('>>');
      _EndObject;


      if aIncludeFont then
      begin
         // àdd font file to pdf
         ppp := pansichar(pfont);
         SetLength(ff,pSize);
         for i:=1 to pSize do
         begin
            ff[i] := ppp^;
            ppp := pointer(longword(ppp) + 1);
         end;
         if aZIPit then
         begin
            ff := _ZIPer(ff);
            pSize := length(ff);
         end;

         _AddObject; // + 2
         _Additem('<< /Length '+ast(psize));
         _Additem('/Length1 '+ast(psize));
         _Additem('/Length2 0');
         _Additem('/Length3 0');
         if aZIPit then _Additem('/Filter [ /FlateDecode ]');
         _Additem('>>');
         _Additem('stream');

         pdf := pdf + ff;
         inc(aOffset,psize);
         _Additem('endstream');
         _EndObject;
      end;

      if pfont <> nil then  ReallocMem(pfont,0);
      if pwidths <> nil then ReallocMem(pwidths,0);
      if pm <> nil then ReallocMem(pm,0); //free
      SelectObject(h_dc, fnto);
      DeleteObject(fnt);
      ReleaseDC(0,h_dc);
   end;

   curfont := '/F'+ast(aFontCount);
   aFontResource := aFontResource + curfont+' '+ast(nf_indx)+' 0 R'+#10;
   aCurFontSize := size;
   Result := nf_indx;

end;



function    BTPDFdoc.RegisterTTFFont(const name:string; size:single; bold,italic:boolean):longword;
begin
   Result := _RegisterFont(name,size,bold,italic,false);
end;

//------------------------------------------------------------------------------
function    BTPDFdoc.RegisterFont(const name:string; size:single; bold,italic:boolean):longword;
var k:longword;        // 0        1         2
    s:string;
const PDF_fonts:ansistring = 'Times Helvetica Courier Symbol ZapfDingbats';
begin
   Result := 1; //the default
   k := Pos(name,string(PDF_fonts));
   if k <> 0 then
   begin
      s := Name;
      if bold or italic then
      begin
         if k < 20 then // Symbols nad Zap dont have bold
         begin
            s := s + '-';
            if Bold then s:= s + 'Bold';
            if Italic then
            begin
               if k > 1  then s := s + 'Oblique'
                         else s := s + 'Italic'; // Times only is italic
            end;
         end;
      end else begin
         if k = 1 then s := s + '-Roman'; // regular Times is Times_Roman
      end;

      Result := _RegisterFont(s,size,bold,italic,true);
  end;
end;

//------------------------------------------------------------------------------
procedure   BTPDFdoc.SetFont(hFont:longword; size:single);
begin
   curfont := '/F'+ast(hFont);
   aCurFontSize := size;
end;

//------------------------------------------------------------------------------
procedure   BTPDFdoc.Text(x,y:single; const s:ansistring);
begin
   TextBox(x,y,0,0,s);
end;

//------------------------------------------------------------------------------
procedure   BTPDFdoc.TextBox(x,y,xl,yl:single; s:ansistring);
var sa:ansistring;
    i:longword;
    ch:ansichar;
begin
   stream := stream +'BT'+#10;

   stream := stream + CurFont+' '+fst(aCurFontSize)+' Tf'+#10; // font & size
   stream := stream + fst(x) +' ' + fst(Y) +' Td' + #10; // start position

   if     aPen and     aBrush then stream := stream + '2 Tr'+#10;
   if not aPen and     aBrush then stream := stream + '0 Tr'+#10;
   if     aPen and not aBrush then stream := stream + '1 Tr'+#10;
   if aPen then stream := stream + aPenSet;

   Sa := '';
   for i:= 1 to length(s) do
   begin
      Ch := s[i];
      if (Ch = '\') or (Ch = '(') or (Ch=')') then Sa := Sa + '\' + Ch
                                              else Sa := Sa + Ch;
   end;

   stream := stream + '('+ Sa +') Tj'+#10; // set text


   stream := stream + '0 -'+fst(aCurFontSize)+' TD'+#10; // to next line

   //   stream := stream + '('+ Sa +') Tj'+#10; // set text

   stream := stream +'ET'+#10;



end;

end.
