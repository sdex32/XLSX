unit BUnicode;

//todo UTF 16

interface

function   Unicode2UTF8(const d:WideString ) :AnsiString;
function   UTF82Unicode(const d:AnsiString ) :WideString;


function   FileData2String( d:AnsiString ) :WideString;
function   String2FileData( d:WideString; encoding:longword ) :AnsiString;

implementation

//------------------------------------------------------------------------------
function   Unicode2UTF8(const d:WideString ) :AnsiString;
var  i,j:longword;
     c:word;
begin
   Result := '';
   j := length(d);
   i := 0;
   while i < j do
   begin
      inc(i);
      c := word(d[i]);
      if c > 2047 then
      begin
         Result := Result + ansichar($E0 or byte((c shr 12) and $F));
         Result := Result + ansichar($80 or byte((c shr 6) and $3F));
         Result := Result + ansichar($80 or byte(c and $3F));
       end else begin
         if c > 127 then
         begin
            Result := Result + ansichar($C0 or byte((c shr 6) and $1F));
            Result := Result + ansichar($80 or byte(c and $3F));
         end else begin
         Result := Result + ansichar(byte(c));
         end;
      end;
   end;
end;

//------------------------------------------------------------------------------
function   UTF82Unicode(const d:AnsiString ) :WideString;
var  i,j:longword;
     c,b2,b3:byte;
begin
   Result := '';
   j := length(d);
   i := 0;
   while i < j do
   begin
      inc(i);
      c := byte(d[i]);
      if True then
//      if (c and $80 = $80) then // have more
///      begin
//         if ((c and $E0) = $C0) ((i+1)<=j) then  // 2 bytes 7FF  110x xxxx  10xx xxxx
//         begin
//            b2 := byte(d[i+1]);
//            if b2 = 0 then break; //error
//
//            Result := Result + WideChar( (longword(c and $1F) shl 6)
//                                      or (longword(b2 and $3F) );
//            Continue;
//         end;


      if ((c and $E0) = $E0) and ((i+2)<=j) then
      begin
         b2 := byte(d[i+1]);
         if b2 = 0 then break; //error
         b3 := byte(d[i+2]);
         if b3 = 0 then break; //error
         Result := result + WideChar( (longword(c  and $F)  shl 12)
                                   or (longword(b2 and $2F) shl 6)
                                   or  longword(b3 and $3F) );
         inc(i,2);
         Continue;
      end;
      if ((c and $C0) = $C0) and ((i+1)<=j)  then
      begin
         b2 := byte(d[i+1]);
         if b2 = 0 then break; //error
         Result := result + WideChar( (longword(c  and $1F) shl 6)
                                    or longword(b2 and $3F) );
         inc(i);
         Continue;
      end;
//      end;
      Result := result + WideChar(c);
   end;
end;

//------------------------------------------------------------------------------
function   FileData2String( d:ansistring ) :WideString;
var j:longword;
    s:ansistring;
begin
   // BOM header  (Byte order mask)
   // UTF-8           EF BB BF             $1234
   // UTF-16 (BE)     FE FF                be - > mem 1 2 3 4
   // UTF-16 (LE)     FF FE                le - > mem 4 3 2 1
   // UTF-32 (BE)     00 00 FE FF
   // UTF-32 (LE)     FF FE 00 00


   j := length(d);
   // test UTF-8
   if j >= 3 then
   begin
      if (d[1] = #$EF) and (d[2] = #$BB) and (d[3] = #$BF) then
      begin
         if j > 3 then s := Copy(d, 4, j - 3)
                  else s := '';
         Result := UTF82Unicode(s);
         Exit;
      end;
   end;

   Result := WideString(d);
end;

//------------------------------------------------------------------------------
function   String2FileData( d:WideString; encoding:longword ) :AnsiString;
begin
   if encoding = 1 then // utf8
   begin
      Result := #$EF#$BB#$BF + Unicode2UTF8(d);
      Exit;
   end;

   Result := ansistring(d);
end;

end.
