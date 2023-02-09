unit BxlsReader;

interface

// Old XLS reader biff8 office format excel97 to office2003

//under construction search TODO


type  BTXlsReader = class
         private
            aSheetCnt:longword;
            aSheet:pointer;
            aCurrent:pointer;
            aStringPool:widestring;
            procedure   _readPoolstr(indx:longword; var Value:widestring);
         public
            constructor Create;
            destructor  Destroy; override;
            procedure   Reset;
            function    OpenFile(const file_name:string; callback:pointer; userparm:nativeUint):boolean;
            function    GetSheetsCount:longword;
            function    GetSheetName(id:longword):string;
            function    SelectSheet(id:longword):boolean;
            function    GetSheetBounds(var max_col,max_row:longword):boolean;
            function    GetCellValue(col,row:longword; var value:widestring):boolean;   //start from (1,1)
      end;

      BTxlsxReaderCallBack = procedure(userparm:nativeUint; sheet,col,row:longword; value:widestring); stdcall;

implementation

uses BStrTools,BFileTools;

type BTCell = record
       next:pointer;
       Row, Col :longword;
       Value :Widestring;
     end;
     PBTCell = ^BTCell;


     BTXlsReaderSheet = record
        next :pointer;
        name :widestring;
        max_col,max_row:longword;
        dataofs :longword;
        Cells:pointer;
     end;
     PBTXlsReaderSheet = ^BTXlsReaderSheet;


//------------------------------------------------------------------------------
constructor BTXlsReader.Create;
begin
   aSheet := nil;
   aCurrent := nil;
   Reset;
end;

//------------------------------------------------------------------------------
destructor  BTXlsReader.Destroy;
begin
   reset;
   inherited;
end;

//------------------------------------------------------------------------------
procedure   BTXlsReader.Reset;
var o,u:PBTXlsReaderSheet;
    oc,uc:PBTCell;
begin
   try
      if aSheet <> nil then
      begin
         o := aSheet;
         while(o<>nil) do
         begin
            u := o;
            // free cels
            oc := u.Cells;
            while (oc<>nil) do
            begin
               uc := oc;
               oc:=oc.next;
               Dispose(uc);
            end;
            o := o.next;
            Dispose(u);
         end;
         aSheet := nil;
      end;
   except
      aSheet := nil;
   end;
   aCurrent := nil;
   aSheetCnt := 0;
   aStringPool := '';
end;

//------------------------------------------------------------------------------
function    BTXlsReader.OpenFile(const file_name:string; callback:pointer; userparm:nativeUint):boolean;
var sa:ansistring;
    j,i,dpnt,k,p,m:longint;
    w,ln:word;
    lw1,lw2,lw3,p1,sz,pn,cm:longword;
    s1:widestring;
    ca:ansichar;
    cw:widechar;
    b1,b2,b3,b4:byte;
    w1,w2,w3,w4:word;
    dp:nativeUint;
    pt:^longword;
    rw:^word;
    o,u:PBTXlsReaderSheet;
    oc,uc:PBTCell;
    lwa:array[1..32] of longword; //danger
    lwac,lwai:longword;
    bl:boolean;
    i64:int64;
    dbl:double;
    dat:widestring;
    callb :BTxlsxReaderCallBack;
    er:longint;
    spt,dpt:pointer;

    function RK(a:longword):double;
    begin
       if (a and 2) = 2 then
       begin
          Result := a shr 2;
       end else begin
          i64 := a;
          i64 := (i64 and $FFFFFFFC) shl 32; // hi words of IEEE num
          REsult := pdouble(pointer(@i64))^;
       end;
       if (a and 1) = 1 then Result := Result /100; //who want this?
    end;


begin
   Reset;
   Result := false;
   aSheetCnt := 0;
   callb := callback;
   er := -1;

   try
   if not FileLoadEX(file_name,sa) then exit;
   j := length(sa);
   if j > 512 then
   begin
      // begining of biff8 D0 CF 11 E0 A1 B1 1A  CFB header
      // signature 3E 00 03 00 FE FF 09  at offset 24
      // first biff start at 512
      if  (plongword(@sa[1])^ = $E011CFD0) and ((plongword(@sa[5])^ and $FFFFFF) = $001AB1A1) then
      begin
         // 25 minir version 3E00
         // 27 majir version 0300  or 0400  version 3 or 4
         // 29 FEFF indicating little-endian byte
         // 31 0900 (indicating the sector size of 512 bytes used for major version 3) or 0C00 (indicating the sector size of 4096 bytes used for major version 4)
         dpnt := 513; //string start from 1
         if pword(@sa[31])^ = $0009 then dpnt := 512+1;
         if pword(@sa[31])^ = $000C then dpnt := 4096+1;


         k := dpnt;
         // ReadGlobals --------------------------------------------------------
         //              for the workbook
         if (pword(@sa[k])^ = $0809) and   // begining of file
            (pword(@sa[k+4])^ = $0600) then // BIFF8
         begin
            while k < j do
            begin
               w := pword(@sa[k])^;
               ln := pword(@sa[k+2])^;
               inc(k,4); //start of data
               // Important TGS
               if w = $00FC then
               begin // SST strin pool
                  lw1 := plongword(@sa[k])^; //str count
                  lw2 := plongword(@sa[k+4])^; //unique str count  c
                  pn := 0;
                  p1 := 0;
                  p := 0;
                  b3 := 0;
                  for i := 1 to lw2 do
                  begin
                     if (k+8+p) >= j then
                     begin
                        break;
                     end;

                     w1 := pword(@sa[k+8+p])^;
                     b1 := pbyte(@sa[k+10+p])^ and 1;
                     b2 := b1 + 1;

                     if w1 = 0 then
                     begin
                       inc(p,3);
                       continue;
                     end;

                     SetLength(s1,w1);
                     if b1 = 0 then
                     begin
                        spt := @sa[k+11+p];
                        dpt := @s1[1];
                        asm
                           push esi
                           push edi
                           push ecx
                           push eax
                           xor  eax,eax
                           mov  esi, spt
                           mov  edi, dpt
                           movzx ecx,w1
                           cld
                         @@11:
                           lodsb
                           stosw
                           loop @@11;
                           pop eax
                           pop ecx
                           pop edi
                           pop esi
                        end;
                     end else begin
                        spt := @sa[k+11+p];
                        dpt := @s1[1];
                        move(spt^,dpt^,w1*2);
                     end;

//                     s1 := '';
//                     for m := 0 to w1-1 do
//                     begin
//                        if b1 = 0 then cw := widechar(pbyte(@sa[k+11+p+m])^)
//                                  else cw := widechar(pword(@sa[k+11+p+m*2])^);
//                        s1 := s1 + cw;
//                     end;

                     if length(s1) = 0  then s1 := 'nop';

                     sz := length(s1);
                     p1 := p1 + 4{bytes header ofs len} + sz;
                     SetLength(aStringPool,p1+4);
                     dp := nativeUint(@aStringPool[1]);
                     spt := @s1[1];
                     dpt := @aStringPool[pn + 5];
                     move(spt^,dpt^,sz*2); // widestring;
//                     for cm := 1 to sz do
//                     begin
//                        aStringPool[pn + 4 +cm] := s1[cm];
//                     end;
                     pt := pointer(dp + pn*2);
                     rw := pointer(dp + pn*2 + 4);
                     pt^ := p1 * 2;
                     rw^ := word(sz);

                     pn := p1;
                     pt := pointer(dp + pn*2);
                     pt^ := 0;

                     inc(b3);
                     p := P + 3 + w1 * b2;
                  end;

               end;
               if w = $0085 then
               begin //sheet boud

                  lw1 := plongword(@sa[k])^; //start offset
                  b1  := pbyte(@sa[k+4])^;  // sheet type   Worksheet=0  MacroSheet=1 Chart=2 VBModule=6
                  b2  := pbyte(@sa[k+5])^ and $3;  // sheet visability  Visible=0 Hidden=1 VeryHidden=2
                  b4  := pbyte(@sa[k+6])^; // string len
                  b3  := pbyte(@sa[k+7])^ and 1; // encoding  ansi=0 unicode=1
                  s1 := '';
                  for m := 0 to b4-1 do
                  begin
                     if b3 = 0 then cw := widechar(pbyte(@sa[k+8+m])^)
                               else cw := widechar(pword(@sa[k+8+m*2])^);
                     s1 := s1 + cw;
                  end;

                  new(o);
                  o.next := nil;
                  o.name := s1;
                  o.max_col := 0;
                  o.max_row := 0;
                  o.dataofs := lw1;
                  o.Cells := nil;

                  if aSheet = nil then
                  begin
                     aSheet := o;
                  end else begin
                     u := aSheet; // link
                     while (u.next <> nil)  do u:= u.next;
                     u.next := o;
                  end;
               end;
               if w = $000A then
               begin
                  break;
               end;
               inc(k,ln);
            end;
         end;


         // READ Sheet data -----------------------------------------------------
         if aSheet <> nil then
         begin
            lwac := 0;
            o := aSheet;
            while o<>nil do
            begin
               inc(aSheetCnt);



               k := o.dataofs + dpnt;  //start offset for sheet 1
               w1 := pword(@sa[k])^;
               if (pword(@sa[k])^ = $0809) and   // begining of file
                  (pword(@sa[k+4])^ = $0600) and // BIFF8  older not
                  (pword(@sa[k+6])^ = $0010) then // worksheet
               begin
                  while k < j do
                  begin
                     w := pword(@sa[k])^;
                     ln := pword(@sa[k+2])^;
                     inc(k,4); //start of data
                     if w = $0200 then
                     begin // dimenstions     (2)
                        lw1 := plongword(@sa[k])^;   //row star
                        lw2 := plongword(@sa[k+4])^;  //row end +1
                        o.max_row := lw2; //Ylng; 1..y
                        w1 := pword(@sa[k+8])^;  // column start
                        w2 := pword(@sa[k+10])^;  //column end +1
                        o.max_col := w2; //xlng 1..y
                        if callback <> nil then callb(userparm,aSheetCnt,o.max_col,o.max_row,'><SHEETSIZE><');
                     end;

                     if w = $020B then
                     begin //index XiffIndex   (1)
                        lw1 := plongword(@sa[k+4])^;  // zero-based index of first existing row
                        lw2 := plongword(@sa[k+8])^;  // zero-based index of last existing row
                        if ln > 16 then
                        begin
                           m := (ln - 16) div 4; // elements count
                           for i := 0 to m-1 do
                           begin
                              lw3 := plongword(@sa[k+16+i*4])^+dpnt; //offset
                              inc(lwac);
                              lwa[lwac] := lw3;
                           end;
                        end;
                     end;

                     if w = $000A then break;
                     inc(k,ln);
                  end;
               end;

               for lwai := 1 to lwac do
               begin
                  k := lwa[lwai];

                  if (pword(@sa[k])^ = $00d7) then   // BDCELL from INDEX
                  begin
                     ln := pword(@sa[k+2])^;
                     inc(k,4); //data zone
                     lw1 := k - 4 - plongword(@sa[k])^;  // Offset of first row linked with this record
                     k := lw1;  // start of the ROW
                     bl := false;
                     repeat //bypass row
                        if pword(@sa[k])^ = $0208 then inc(k,4 + pword(@sa[k+2])^)
                                                   else bl := true;
                     until bl;

                     bl := false;
                     repeat
                        w := pword(@sa[k])^;
                        ln := pword(@sa[k+2])^;
                        inc(k,4); //start of data
                        if w = $000A then break;
                        if w = $00d7 then break;
                        w1 := pword(@sa[k])^; //row
                        w2 := pword(@sa[k+2])^; //col
                        if w1 > o.max_row then continue;
                        if w2 > o.max_col then continue;

                        case w of
                          $0202,$0002: begin // INTEGER, INTEGER_OLD
                              w4 := pword(@sa[k+6])^;
                              dat := ToStr(w4);
                           end;
                           $0203,$0003: begin // NUMBER, NUMBER_OLD
                              dbl := pdouble(@sa[k+6])^;
                              str(dbl,dat);
                           end;
                           $0204,$0004,$00D6: begin //LABEL, LABEL_OLD, RSTRING
                              m := pbyte(@sa[k+6])^; // length
                              Setlength(s1,m);
                              b3 := 0; //??????????? TODO
                              for i := 0 to m-1 do
                              begin
                                //encoding /??????
                                 //TODO
                                 if b3 = 0 then s1[i+1] := widechar(pbyte(@sa[k+8+i])^)
                                           else s1[i+1] := widechar(pword(@sa[k+8+i*2])^);
                              end;
                              Dat := s1;
                           end;
                           $00FD: begin // LABELSST:
                              lw2 := plongword(@sa[k+6])^;
                              _readPoolstr(lw2,Dat); // get from pool
                           end;
                           $027E: begin //RK:
                              lw2 := plongword(@sa[k+6])^;
                              dbl := RK(lw2);
                              str(dbl,dat);
                           end;
                           $00BD: begin //MULRK:
                              m := pword(@sa[k+ln-2])^; // LastColumnIndex
                              for i:= w2{ColumnIndex} to m do
                              begin
                                 lw2 := plongword(@sa[k+6+(6*i)])^;
                                 dbl := rk(lw2);
//TODO input

                                 //add to DATA[row][i]:=dbl;
                              end;
                           end;
                           $0201,$0001,$00BE: begin // BLANK, BLANK_OLD, MULBLANK
                              Dat := '';
                              //blank
                           end;
                           $0406,$0006: begin // FORMULA, FORMULA_OLD
                           //TODO
//                                ((XlsBiffFormulaCell)cell).UseEncoding = m_encoding;
//                                object val = ((XlsBiffFormulaCell)cell).Value;
//                                if (val == null)
//                                    val = string.Empty;
//                                else if (val is FORMULAERROR)
//                                    val = "#" + ((FORMULAERROR)val).ToString();
//                                else if (val is double)
//                                    val = FormatNumber((double)val);
//                                dt.Rows[cell.RowIndex][cell.ColumnIndex] = val.ToString();
                           end;
                           else begin
                              bl := true; //break;
                           end;
                        end;
                        inc(k,ln);

                        new(oc);
                        oc.Row := w1+1; // start from 1
                        oc.Col := w2+1;
                        oc.next := nil;
                        oc.Value := dat;

                        if o.cells = nil then
                        begin
                           o.cells := oc;
                        end else begin
                           uc := o.Cells; // link
                           while (uc.next <> nil)  do uc:= uc.next;
                           uc.next := oc;
                        end;

                        if callback <> nil then callb(userparm,aSheetCnt,w2+1,w1+1,dat);

                        er := 0; //?????? TODO

                     until bl;
                  end;
               end; // for lwa
               if callback <> nil then callb(userparm,aSheetCnt,0,0,'><FINISH><');
               o := o.next; // Next sheet
            end; // While o <> nil
         end; // aSheet <> nil
      end; //Biff8
   end; //length > 512
   except
      aSheetCnt := 0;
   end;
   if aSheetCnt > 0 then Result := true;
end;

//------------------------------------------------------------------------------
function    BTXlsReader.GetSheetsCount:longword;
begin
   Result := aSheetCnt;
end;

//------------------------------------------------------------------------------
function    BTXlsReader.GetSheetName(id:longword):string;
var o:PBTXlsReaderSheet;
    i :longword;
begin
   Result := '';
   try
   if (id > 0) and (id <= aSheetCnt) then
   begin
      if aSheet <> nil then
      begin
         o := aSheet;
         i := 1;
         repeat
            if i = id then
            begin
               Result := o.name;
               break;
            end;
            o := o.Next;
            inc(i);
         until o = nil;
      end;
   end;
   except
      Result := '';
   end;
end;

//------------------------------------------------------------------------------
function    BTXlsReader.SelectSheet(id:longword):boolean;
var o:PBTXlsReaderSheet;
    i :longword;
begin
   Result := false;
   aCurrent := nil;
   try
   if (id > 0) and (id <= aSheetCnt) then
   begin
      if aSheet <> nil then
      begin
         o := aSheet;
         i := 1;
         repeat
            if i = id then
            begin
               aCurrent := o;
               Result := true;
               break;
            end;
            o := o.Next;
            inc(i);
         until o = nil;
      end;
   end;
   except
      Result := false;
   end;
end;

//------------------------------------------------------------------------------
procedure   BTXlsReader._readPoolstr(indx:longword; var Value:widestring);
var k,j,cm,rr,ln:longword;
    pp:nativeUint;
    pt :^longword;
    pw:^word;
begin
   Value := '';
   try
   cm :=0;
   pp := nativeUint(@aStringPool[1]);
   for j := 0 to indx do
   begin
      pt := pointer(pp+cm);
      pw := pointer(pp+cm+4);
      ln := (cm div 2) + 4; // must be 4 but rr := 1
      cm := pt^;
      k := pw^;
      if j = indx then // last get value
      begin
         SetLength(Value,k);
         for rr := 1 to k do Value[rr] := aStringPool[rr+ln];
      end;
   end;
   except
      Value := '';
   end;
end;

//------------------------------------------------------------------------------
function    BTXlsReader.GetCellValue(col,row:longword; var value:widestring):boolean;
var o:PBTXlsReaderSheet;
    oc : PBTCell;
begin
   Result := false;
   Value := '';
   try
      if aCurrent <> nil then
      begin
         o := aCurrent;
         if col = 0 then Exit;
         if row = 0 then Exit;
         if (col <= o.max_col) and (row <= o.max_row) then
         begin
            if o.Cells <> nil then
            begin
               oc := o.Cells;
               while (oc<>nil) do
               begin
                  if (oc.Row = row) and (oc.Col = col) then Value := oc.Value;
                  oc := oc.next;
               end;
               Result := true;
            end;
         end;
      end;
   except
      Result := false;
   end;
end;

//------------------------------------------------------------------------------
function    BTXlsReader.GetSheetBounds(var max_col,max_row:longword):boolean;
var o:PBTXlsReaderSheet;
begin
   Result := false;
   if aCurrent <> nil then
   begin
      try
         o := aCurrent;
         max_col := o.max_col;
         max_row := o.max_row;
         Result := true;
      except
         max_col := 0;
         max_row := 0;
         Result := false;
      end;
   end;
end;


end.
