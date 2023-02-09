unit BxlsxReader;

interface

// version 0.5   Bogi aka sdex32 5.2022  :)
// TODO import csv
//work well very fast after rewriting
//now after parsing i write offset inseide the xml file much faster thar xml parsing !!!
// code is ugly but work fine on 13mb xlsx tested
//TODO !!!!!! need optimisation for speed !!!!!!!!!!!!!!!!

type  BTXlsxReader = class
         private
            aSheetCnt:longword;
            aSheet:pointer;
            aCurrent:pointer;
            aStringPool:widestring;
            procedure   _readPoolstr(const s:ansistring; var Value:widestring);
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



uses BpasZlib,BStrTools,BUnicode,BTinyXML;

type BTXlsxReaderSheet = record
        next :pointer;
        name :string;
        data :ansistring;
        max_col,max_row:longword;
     end;
     PBTXlsxReaderSheet = ^BTXlsxReaderSheet;

//------------------------------------------------------------------------------
constructor BTXlsxReader.Create;
begin
   aSheet := nil;
   aCurrent := nil;
   Reset;
end;

//------------------------------------------------------------------------------
destructor  BTXlsxReader.Destroy;
begin
   reset;
   inherited;
end;

//------------------------------------------------------------------------------
procedure   BTXlsxReader.Reset;
var o,u:PBTXlsxReaderSheet;
begin
   try
      if aSheet <> nil then
      begin
         o := aSheet;
         while(o<>nil) do
         begin
            u := o;
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
function    BTXlsxReader.OpenFile(const file_name:string; callback:pointer; userparm:nativeUint):boolean;
var z:TZipRead;
    i,j,nj,p,k,cm,sz,pc,oc,p1,p3,pn:longword;
    data,s:ansistring;
    cc:ansichar;
    dat:widestring;
    f,f1,er:longint;
    o,u:PBTXlsxReaderSheet;
    s1,m1,l1,n1:string;
    pt:^longword;
    rw:^word;
    dp:nativeUint;
    st:boolean;
    callb :BTxlsxReaderCallBack;
    value :widestring;
    colid,rowid:longword;
    pin:longword;
    spt,dpt:pointer;
begin
   Reset;
   Result := false;
   aSheetCnt := 0;
   aStringPool := '';
   callb := callback;
   try
   z:= TZipRead.Create(file_name);
   if z.count <> 0 then
   begin
      i := z.NameToIndex('xl/sharedStrings.xml');
      data := z.UnZip(i);
      f := 0;
      s := TinyXML_Parse(data,'/sst.uniqueCount',f);
      if f = 0 then
      begin
         i := toval(string(s)); // count of unique strings
         aStringPool := '';
         if i > 0 then
         begin
            pn := 0;
            p1 := 0;
            pin := 1;

            for j := 1 to i do    //work faster
            begin
               pin := FastPosA('<si><t>',data,pin);
               if Pin <> 0 then
               begin
                  f := pin + 7;
                  pin := FastPosA('</t>',data,f);
                  if pin <> 0 then
                  begin
                     s := Copy(Data,f,pin-f);
                     dat := HTMLDecode(UTF82Unicode(s)); //utf8towidestring(s));
                     sz := length(dat);
                     p1 := p1 + 4{bytes header ofs len} + sz;
                     SetLength(aStringPool,p1+4);
                     dp := nativeUint(@aStringPool[1]);

                     spt := @dat[1];
                     dpt := @aStringPool[pn + 5];
                     move(spt^,dpt^,sz*2); // widestring;

//                     for cm := 1 to sz do
//                     begin
//                        aStringPool[pn + 4 +cm] := dat[cm];
//                     end;

                     pt := pointer(dp + pn*2);
                     rw := pointer(dp + pn*2 + 4);
                     pt^ := p1 * 2;
                     rw^ := word(sz);

                     pn := p1;
                     pt := pointer(dp + pn*2);
                     pt^ := 0;

                  end;
               end;
   (*
               f := 0;
               s := TinyXML_Parse(data,'/sst/si['+ansistring(tostr(j))+']/t',f);
               if f = 0 then
               begin
                  dat := HTMLDecode(utf8towidestring(s));
                  sz := length(dat);
                  p1 := p1 + 4{bytes header ofs len} + sz;
                  SetLength(aStringPool,p1+4);
                  dp := nativeUint(@aStringPool[1]);
                  for cm := 1 to sz do
                  begin
                     aStringPool[pn + 4 +cm] := dat[cm];
                  end;
                  pt := pointer(dp + pn*2);
                  rw := pointer(dp + pn*2 + 4);
                  pt^ := p1 * 2;
                  rw^ := word(sz);

                  pn := p1;
                  pt := pointer(dp + pn*2);
                  pt^ := 0;

               end;
     *)
            end;
         end;
      end;

      i := z.NameToIndex('xl/workbook.xml');
      data := z.UnZip(i);
      i := 1;
      repeat
         f := 0;
         s := TinyXML_Parse(data,'/workbook/sheets/sheet['+ansistring(toStr(i))+'].name',f);
         if f = 0 then
         begin
            New(o);
            o.next := nil;
            o.name := utf82Unicode(s); //towidestring(s);
            f := 0;
            s := TinyXML_Parse(data,'/workbook/sheets/sheet['+ansistring(toStr(i))+'].sheetId',f);
            j := z.NameToIndex('xl/worksheets/sheet'+ansistring(s)+'.xml');
            o.data := z.UnZip(j);
            if length(data) > 0 then
            begin
               if aSheet = nil then
               begin
                  aSheet := o;
               end else begin
                  u := aSheet; // link
                  while (u.next <> nil)  do u:= u.next;
                  u.next := o;
               end;
               // parse o.data sheet
               er := 1;
               // get max col and rows
               s1 := string(TinyXML_Parse(o.data,'/worksheet/dimension.ref',f1));
               if f1 = 0 then
               begin
                  m1 := parsestr(s1,1,':');
                  j := length(m1);
                  if j > 0 then
                  begin
                     l1 := '';
                     n1 := '';
                     p := 0;
                     for k := 1 to j do
                     begin
                        if m1[k] < 'A' then p := 1; // number start
                        if p = 0 then l1 := l1 + m1[k] else n1 := n1 + m1[k];
                     end;
                     o.max_row := ToVal(n1);
                     j := length(l1);
                     o.max_col := 0;
                     for k := 1 to j do
                     begin
                        o.max_col := (o.max_col * 26) + (longword(byte(l1[k])) - 64); //A-65
                     end;
                  end;
                  er := 0;
               end;

               if callback <> nil then callb(userparm,i,o.max_col,o.max_row,'><SHEETSIZE><');

               // Create index inside
               // it was very slow for big files  with thausend of rows
               // so i make index inside the text of the xml (sheet data)
               // tread string as memory with start offset 0
               // Ofs 0 -- start offset of first row
               // row tag <row r="1" spans="1:3" ....
               //         0123456789 bytes
               // 0123 (dword) - next offset of next row if = $00000000 then EOF rows
               // 45   (word)  - row id (start from 1...)
               // 6789 (dword) - offset to first col tag
               //
               // col tag <c r="A1"><v>21</v> ..
               //         012345678 bytes
               // 0123 (dword) - next offet of next col data if zero EOF col
               // 45   (word)  - col id
               // 67   (word)  - size of data  if $8000 is up this is index in stringpool
               // 89....  data from tag <v> is shifted to start from offset 8
               //
               p := 0; // first offset point to first row
               j := 1; // start point
               dp := nativeUint(@o.data[1]);
               if er = 0 then
               begin
                  repeat
                     k := j;
                     j := FastPosA('<row ',o.data,k);
                     nj := FastPosA('<row ',o.data,j+5);
                     if nj = 0 then nj := length(o.data);

                     if j<>0 then
                     begin
                        dec(j); // zero start ptr
                        // get the row id
                        sz := 0;
                        s := '';
                        cm := 0;
                        repeat
                           cc := o.Data[j+sz];
                           if cc <> #32 then
                           begin
                              if cm = 3 then
                              begin
                                 if cc = '"' then break;
                                 s := s + cc;
                              end;
                              if cm = 2 then if cc = '"' then cm := 3 else cm := 0;
                              if cm = 1 then if cc = '=' then cm := 2 else cm := 0;
                              if cm = 0 then if cc = 'r' then cm := 1;
                           end;
                           inc(sz);
                        until sz > 10;
                        if length(s) > 0 then sz := Toval(string(s)) else er := 1;
                        if er = 0 then
                        begin
                           // strore the row index
                           pt := pointer(dp + p);
                           pt^ := j; // save row offet on prev pos
                           p := j;
                           pt := pointer(dp + p);
                           pt^ := 0; // save row next ofs EOF
                           pt := pointer(dp + p + 4);
                           pt^ := sz; // save row id
                           rowid := sz;
                           p1 := p + 8;
                           repeat
                              // now start search for
                              oc := FastPosA('<c ',o.data,p1);
                              if (oc > 0) and (oc < nj) then
                              begin
                                 dec(oc); // zero start ptr
                                 sz := 0;
                                 s := '';
                                 st := false;
                                 cm := 0;
                                 repeat
                                    cc := o.Data[oc+sz];
                                    if cc <> #32 then
                                    begin
                                       if cm = 7 then
                                       begin
                                          if cc = 's' then st := true; // we have string index
                                          break;
                                       end;
                                       if cm = 6 then if cc = '"' then cm := 7 else break;
                                       if cm = 5 then if cc = '=' then cm := 6 else break;
                                       if cm = 4 then
                                       begin
                                          if cc = '>' then break;
                                          if cc = 't' then cm := 5;
                                       end;
                                       if cm = 3 then
                                       begin
                                          if cc = '"' then cm := 4;
                                          if (cc >= '0') and (cc <= '9') then cm := 4;
                                          if cm = 3 then s := s + cc;
                                       end;
                                       if cm = 2 then if cc = '"' then cm := 3 else cm := 0;
                                       if cm = 1 then if cc = '=' then cm := 2 else cm := 0;
                                       if cm = 0 then if cc = 'r' then cm := 1;
                                    end;
                                    inc(sz);
                                 until sz > 32;
                                 if length(s) > 0 then
                                 begin
                                    sz := 0;
                                    for cm := 1 to length(s) do
                                    begin
                                      sz := sz *26 + (longword(byte(s[cm])) - 64);
                                    end;
                                 end else er := 1;

                                 if er = 0 then
                                 begin
                                    // store the col index
                                    pt := pointer(dp + p1);
                                    pt^ := oc; // save col offet on prev pos
                                    p1 := oc;
                                    pt := pointer(dp + p1);
                                    pt^ := 0; // save col next ofs EOF
                                    pt := pointer(dp + p1 + 4);
                                    pt^ := sz; // save col id
                                    colid := sz;
                                    Value := '';

                                    p3 := FastPosA('<v>',o.Data,p1+8);
                                    cm := 0;
                                    s := '';
                                    sz := 0;
                                    if p3> 0 then
                                    begin
                                       p3 := p3 + 3;
                                       pn := p1 + 11; //+1 for
                                       repeat
                                         cc := o.Data[p3];
                                         if cc = '<' then cm := 1;
                                         if cm = 0 then
                                         begin
                                           if callback <> nil then Value := Value + Widechar(cc);
                                           o.data[pn+sz] :=  cc;
                                           inc(sz);
                                         end;
                                         inc(p3);
                                       until cm <> 0;
                                    end;
                                    rw := pointer(dp + p1 + 8);
                                    if st then sz := sz or $8000; // string marker
                                    rw^ := word(sz);
                                    if callback <> nil then
                                    begin
                                       if st then _ReadPoolStr(ansistring(Value),Value);
                                       callb(userparm,i,colid,rowid,value);
                                    end;
                                 end;
                              end;
                           until (oc = 0) or (oc >= nj) or (er <> 0);
                        end;
                     end;
                  until (j = 0) or (er <> 0);
               end;

               if er = 0 then inc(aSheetCnt);
            end else begin
               Dispose(o);
            end;
            if callback <> nil then callb(userparm,i,0,0,'><FINISH><');
            inc(i);
         end;
      until f <> 0; //= 100;
   end;
//debug    FileSave('D:\1',o.data);
   z.Destroy;
   except
      aSheetCnt := 0;
   end;
   if aSheetCnt > 0 then Result := true;
end;

//------------------------------------------------------------------------------
function    BTXlsxReader.GetSheetsCount:longword;
begin
   Result := aSheetCnt;
end;

//------------------------------------------------------------------------------
function    BTXlsxReader.GetSheetName(id:longword):string;
var o:PBTXlsxReaderSheet;
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
function    BTXlsxReader.SelectSheet(id:longword):boolean;
var o:PBTXlsxReaderSheet;
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

(*
const ExcelCol:array[0..25] of ansichar = ('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');

function ExcelAdr(row,col:longword):ansistring;
var s:ansistring;
    i:longword;
begin
   s := '';
   while (col <> 0) do
   begin
      col := col - 1;
      i := (col mod 26);
      col := col div 26;
      s := ExcelCol[i] + s;
   end;
   Result := s + ansistring(tostr(row));
end;
*)

//------------------------------------------------------------------------------
procedure   BTXlsxReader._readPoolstr(const s:ansistring; var Value:widestring);
var i,k,j,cm,rr,ln:longword;
    pp:nativeUint;
    pt :^longword;
    pw:^word;
begin
   try
   val(string(s),i,cm);
   if cm = 0  then
   begin
      pp := nativeUint(@aStringPool[1]);
      for j := 0 to i do
      begin
         pt := pointer(pp+cm);
         pw := pointer(pp+cm+4);
         ln := (cm div 2) + 4; // must be 4 but rr := 1
         cm := pt^;
         k := pw^;
         if j = i then // last get value
         begin
            SetLength(Value,k);
            for rr := 1 to k do Value[rr] := aStringPool[rr+ln];
         end;
      end;
   end else begin
      Value := widestring(s);
   end;
   except
      Value := widestring(s);
   end;
end;

//------------------------------------------------------------------------------
function    BTXlsxReader.GetCellValue(col,row:longword; var value:widestring):boolean;
var s:ansistring;
//    co:ansichar;
//    f:longint;
    o:PBTXlsxReaderSheet;
    i,j,rr,cc,ln,cm:longword;
    pt :^longword;
    pp:nativeUint;
    pw:^word;
    st:boolean;
begin
   Result := false;
   Value := '';
//   sr := ansistring(tostr(row));
//   cid := ExcelAdr(row,col);
   try
      if aCurrent <> nil then
      begin
         o := aCurrent;
         pp := longword(@o.data[1]);
         ln := length(o.data);
         if col = 0 then Exit;
         if row = 0 then Exit;

         if (col <= o.max_col) and (row <= o.max_row) then
         begin
            rr := 0;
            repeat
               pt := pointer(pp+rr+4);
               i := pt^; //row id
               if i = row then
               begin
                  pt := pointer(pp+rr+8);
                  cc := pt^; //first col ofs
                  repeat
                     pt := pointer(pp+cc+4);
                     i := pt^; // col id
                     if i = col then
                     begin
                        pw := pointer(pp+cc+8);
                        j := longword(pw^); // size of string
                        if (j and $8000) <> 0 then st := true else st := false;
                        j := j and $7FFF;
                        SetLength(s,j);
//                        s := '';
                        cc := cc + 10;
                        for cm := 1 to j do s[cm] := o.data[cc+cm];
//                        for cm := 1 to j do s := s + o.data[cc+cm];
                        if st then
                        begin // string index
                           _ReadPoolStr(s,Value);
{
                           val(string(s),i,cm);
                           if cm = 0  then
                           begin
//                           cm := 0;
                              pp := nativeUint(@aStringPool[1]);
                              for j := 0 to i do
                              begin
                                 pt := pointer(pp+cm);
                                 pw := pointer(pp+cm+4);
                                 ln := (cm div 2) + 4; // must be 4 but rr := 1
                                 cm := pt^;
                                 k := pw^;
                                 if j = i then // last get value
                                 begin
                                    SetLength(Value,k);
//                                    Value := '';
                                    for rr := 1 to k do Value[rr] := aStringPool[rr+ln];
//                                    for rr := 1 to k do Value := Value + aStringPool[rr+ln];
                                 end;
                              end;
                           end else begin
                              Value := widestring(s);
                           end;
}
                        end else begin
                           Value := HTMLDecode(utf82unicode(s)); //toWidestring(s)); // clear value
                        end;
                        Result := true;
                        Exit;
                     end;
                     pt := pointer(pp+cc);
                     cc := pt^; // next col
                  until (cc = 0) or (cc > ln);
               end;
               pt := pointer(pp+rr);
               rr := pt^; // get offset of row
            until (rr = 0) or (rr > ln);
         end;
      end;
   except
      Result := false;
   end;

(*
      if GetSheetBounds(mc,mr) then
      begin
            o := aCurrent;
{
         if (col <= mc) and (row <=mr) then
         begin

            cid := '<c r="'+cid+'"';
            f := Pos(cid,o.Data);
            if f > 0 then
            begin




            end;
         end;
}


         for rr := 1 to mr do
         begin
            s := TinyXML_Parse2(o.data,'/worksheet/sheetData/row['+ansistring(tostr(rr))+'].r',f); // get row id
            if s = sr then
            begin
//               s := TinyXML_Parse2(o.data,'/worksheet/sheetData/row['+ansistring(tostr(rr))+']',f); // get row id
               for cc := 1 to mc do
               begin
                  s := TinyXML_Parse2(o.data,'/worksheet/sheetData/row['+ansistring(tostr(rr))+']/c['+ansistring(tostr(cc))+'].r',f); // get cel id
                  if s = cid then
                  begin
                     s := TinyXML_Parse2(o.data,'/worksheet/sheetData/row['+ansistring(tostr(rr))+']/c['+ansistring(tostr(cc))+']/v',f); // get cel id
                     value := utf8towidestring(s);
                     Result := true;
                  end;
                  if Result then break;
               end;
            end;
            if Result then break;
         end;

      end;
*)

end;

//------------------------------------------------------------------------------
function    BTXlsxReader.GetSheetBounds(var max_col,max_row:longword):boolean;
var o:PBTXlsxReaderSheet;
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
(*
      s := TinyXML_Parse2(o.data,'/worksheet/dimension.ref',f);
      if f = 0 then
      begin
         m := parsestr(string(s),1,':');
         j := length(m);
         if j > 0 then
         begin
            l := '';
            n := '';
            p := 0;
            for i := 1 to j do
            begin
               if m[i] < 'A' then p := 1; // number start
               if p = 0 then l := l + m[i] else n := n + m[i];
            end;
            max_row := ToVal(n);
            j := length(l);
            max_col := 0;

            for i := 1 to j do
            begin
               max_col := (max_col * 26) + (longword(byte(l[i])) - 64); //A-65
            end;

            Result := true;
         end;
      end;
*)
end;




end.
