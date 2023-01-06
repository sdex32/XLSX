unit BStrTools;

{$IFDEF FPC}
  {$MODE Delphi}
{$ENDIF}

interface

// Big fock STRING is not thread save :( stupid Delphi


{$IFNDEF FPC }
{$IFDEF RELEASE}
{$WEAKLINKRTTI ON}
{$RTTI EXPLICIT METHODS([]) PROPERTIES([]) FIELDS([]) }
{$ENDIF}
{$ENDIF}


const RS_CaseSense = 1;
      RS_DoOnce = 2;

type TStringDynArray = array of string;


function MidStr (const AText :String; const AStart, ACount :longword) :String;
function RightStr (const AText :String; const ACount :longword) :String;
function LeftStr (const AText :String; const ACount :longword) :String;
function Trim (const S :String) :String;
function ToStr (const v :longword) :String;
function ToStrI (const v :longint) :String;
function ToStrF(F:single; Width,Decimal:longword):string;
function ToStr64(i:int64):string;
function ToStrZlead (v,d:longword) :string;
function PZStrToStr(p:pointer; ansi:boolean=true) :string;
function PZStrLen(p:pointer; ansi:boolean=true) :longword;
procedure CopyStrToPZStr(s:string; p:pointer; ansi:boolean=true); // must be allocated
function ToVal (const v :String) :longint;
function ToValDW (const v :String) :longword;
function ToValF (const v :String) :single;
function HexVal (const s :String) :longword;
function ToHex (value, digits :longword) :String;
function ParseStr (const inStr :String; id :longword; delimiter :Char) :String;
function UpperCase (const Source :String) :String;
function LowerCase (const Source :String) :String;
function ReplaceChar (const Source :String; oldChar, newChar :Char) :String;
function ReplaceString (const Source, OldPattern, NewPattern :String; Flags :longword) :String;
function InsertString (const Source :String; Pos:longword; const InString :String) :String;
function string2RGB_REF (const ColorRef :String) :longword;
function RGB_REF2string (Color :longword) :String;
function ClearQuotes(const Atext:String) :String;
function LexParseStr(const Source, Delimiters :AnsiString; var StartPos :longword; EndPos:longword; Trim :boolean) :AnsiString; //TODO ne copa
function ParseStrSub (const inStr :String; id,subid :longword; delimiter,subdelimiter :Char) :String;
function StrToCase(const TheStr: string; CasesList: array of string): Integer;
function SkipChar (const Source :String; skipaChar :Char) :String;
function SkipString (const Source, SkipPattern :String; Flags :longword) :String;
function DwToStr(D:longword; Reverse:boolean = true):ansistring; // binary
function StrToDw(S:Ansistring; Reverse:boolean = true):longword; // binary
function BoolStr(value:boolean; num0_txt1:longword):string;
function StrBool(value:string):boolean;
function DataToStr(p:pointer; len:longword):ansistring;  // binary
function TrimIn(const S:string):string;
function PosA(const SearchTxt,Txt:ansistring):longword;
function SplitStr(const txt,delimiter:string):TStringDynArray;
//function ReadUntil(const Txt,Delimiter:String):string;
function HTTPDecode(const txt: string): string;
function HTTPEncode(const txt: string): string;
function HTMLDecode(const txt: string): string;
function HTMLEncode(const txt: string): string;
function WideCP1251(const txt :widestring):ansistring;
function CP1251Wide(const txt :ansistring):widestring;
function FastCPosW(a:widechar; const s:widestring):longword;  // if you use string and char slow must be same type
function FastCPosA(a:ansichar; const s:ansistring):longword;


implementation

uses windows;

//------------------------------------------------------------------------------
function BoolStr(value:boolean; num0_txt1:longword):string;
begin
   if num0_txt1 = 0 then
   begin
      if value then Result := '1'
               else Result := '0';
   end else begin
      if value then Result := 'true'
               else Result := 'false';
   end;
end;
//------------------------------------------------------------------------------
function StrBool(value:string):boolean;
begin
   value := UpperCase(Trim(value));
   Result := false;
   if (value = '1') or (value = 'TRUE') then Result := true;
end;


//------------------------------------------------------------------------------
// usage   StrToCase(txt,['xsxs','axsaxa','asdsd']) return 0, 1, 2  or -1 default
function StrToCase(const TheStr: string; CasesList: array of string): Integer;
var Idx: integer;
begin
   Result := -1;
   for Idx := 0 to Length(CasesList) - 1 do
   begin
      if TheStr = CasesList[Idx] then
      begin
         Result := Idx;
         Break;
      end;
   end;
end;

//------------------------------------------------------------------------------
function MidStr (const AText :String; const AStart, ACount :longword) :String;
begin
   Result := Copy(AText, AStart, ACount);
end;

//------------------------------------------------------------------------------
function RightStr (const AText :String; const ACount :longword) :String;
begin
   Result := Copy(AText, longword(Length(AText)) + 1 - ACount, ACount);
end;

//------------------------------------------------------------------------------
function LeftStr (const AText :String; const ACount :longword) :String;
begin
   Result := Copy(AText, 1, ACount);
end;

//------------------------------------------------------------------------------
function Trim (const S :String) :String;
var   I, L: Integer;
begin
   Result := '';
   L := Length(S);
   I := 1;
   while (I <= L) and (S[I] <= ' ') do Inc(I);
   if I <= L then
   begin
      while S[L] <= ' ' do Dec(L);
      if I <= L then Result := Copy(S, I, L - I + 1);
   end;
end;

//------------------------------------------------------------------------------
function ToStr (const v :longword) :String;
var s:ShortString;
begin
   str(v,s);
   Result := string(s);
end;

//------------------------------------------------------------------------------
function ToStrI (const v :longint) :String;
var s:ShortString;
begin
   str(v,s);
   Result := string(s);
end;


//------------------------------------------------------------------------------
function ToVal (const v :String) :longint;
var code:longint;
begin
   val(v,Result,code);
   if code <> 0 then Result := 0;
end;

function ToValDW (const v :String) :longword;
var code:longint;
begin
   val(v,Result,code);
   if code <> 0 then Result := 0;
end;

function ToValF (const v :String) :single;
var code:longint;
begin
   val(v,Result,code);
   if code <> 0 then Result := 0;
end;


//------------------------------------------------------------------------------
const hexdigit: array [0..15] of char = ('0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F');

//------------------------------------------------------------------------------
function HexVal (const s :String) :longword;
var res,i,j,l:longword;
    c:char;
begin
   res := 0;
   j := length(s);
   if j > 0 then
   begin
     for i := 1 to j do
     begin
       res := res shl 4;
//       l:=0;
//       for k := 0 to 15 do if hexdigit[k] = UpCase(s[i]) then l:=k;    // this do test for corect chars //TODO
       c := UpCase(s[i]);
       if c < 'A' then l := byte(C) - 48 {0}                       // danger no  test for err //TODO
                  else l :=(byte(C) - 55){A=65 it is 10 so -65+10 = 55};
       res := res + l;
     end;
   end;
   HexVal := res;
end;

//------------------------------------------------------------------------------
function ToHex (value, digits :longword) :String;
var i:longword;
begin
{ old school al -> al char
                and     al, 0Fh
                add     al, 90h
                daa
                adc     al, 40h
                daa
;;;                or      al, 30h
;;;                cmp     al, 3Ah
;;;                jb      short HexPass
;;;                xor     al, 78h
;;;                dec     al
;;;HexPass:
}
  dec(digits);
  Result := '';
  for i := digits downto 0 do result := Result + hexdigit[(Value shr (I*4)) and $F];
end;

//------------------------------------------------------------------------------
function ParseStr (const inStr :String; id :longword; delimiter :Char) :String;
var  i,j,k : longword;
     c:Char;
begin
   result := '';
   if instr = '' then Exit;
   j := length(instr);
   i := 0;
   k := 0;  // counter start from 0
   while i < j do
   begin
      inc(i);
      c := instr[i];
      if c = delimiter then
      begin
         if k = id then break;
         inc(k);
         continue;
      end;
      if k = id then result := result + c;
   end;
end;

//------------------------------------------------------------------------------
function UpperCase (const Source :String) :String;
var i,j:longint;
begin
   j := length(Source);
   SetLength(Result,j);
   if  j > 0 then for i :=1 to j do Result[i] := Char(CharUpper(PChar(longword(Source[i]))));
end;

//------------------------------------------------------------------------------
function LowerCase (const Source :String) :String;
var i,j:longint;
begin
   j := length(Source);
   SetLength(Result,j);
   if  j > 0 then for i :=1 to j do Result[i] := Char(CharLower(PChar(longword(Source[i]))));
end;

//------------------------------------------------------------------------------
function ReplaceChar (const Source :String; oldChar, newChar :Char) :String;
var i: longint;
begin
  Result := Source;
  for i := 1 to Length(Result) do  if Result[i] = oldChar then Result[i] := newChar;
end;

//------------------------------------------------------------------------------
// todo optimize
function ReplaceString (const Source, OldPattern, NewPattern :String; Flags :longword) :String;
var SearchStr, Patt, NewStr: string;
    Offset: Integer;
begin
   if (RS_CaseSense and Flags) <> 0 then
   begin
      SearchStr := Source;
      Patt := OldPattern;
   end else begin
      SearchStr := UpperCase(Source);
      Patt := UpperCase(OldPattern);
   end;

   NewStr := Source;
   Result := '';

   while SearchStr <> '' do
   begin
      Offset := Pos(Patt, SearchStr);
      if Offset = 0 then
      begin
         Result := Result + NewStr;
         Break;
      end;
      Result := Result + Copy(NewStr, 1, Offset - 1) + NewPattern;
      NewStr := Copy(NewStr, Offset + Length(OldPattern), MaxInt);
      if (RS_DoOnce and Flags) <> 0 then
      begin
         Result := Result + NewStr;
         Break;
      end;
      SearchStr := Copy(SearchStr, Offset + Length(Patt), MaxInt);
   end;
end;

//------------------------------------------------------------------------------
function SkipString (const Source, SkipPattern :String; Flags :longword) :String;
var s:string;
begin
   s := '';
   Result := ReplaceString(Source,SkipPattern,s,0);
end;

//------------------------------------------------------------------------------
function SkipChar (const Source :String; skipaChar :Char) :String;
var s:string;
begin
   s := skipaChar;
   Result := SkipString(Source,s,0);
end;




//------------------------------------------------------------------------------
function LexParseStr(const Source, Delimiters :AnsiString; var StartPos :longword; EndPos:longword; Trim :boolean) :AnsiString;
var c:AnsiChar;
    i:longword;
//    done:boolean;
begin
   Result := '';
   i := 0;
//   done := false;
   if (StartPos > 0) and (StartPos <= EndPos) then
   begin
      repeat
         c := Source[StartPos];
         inc(StartPos);
         if not( Trim and (c in [#0..#32])) then
         begin
            if Pos(c,Delimiters) <> 0 then
            begin
               if i = 0 then Result := c;
               Exit;
            end;
            Result := Result + c;
            inc(i);
         end;
         if StartPos > EndPos then StartPos := 0;
      until StartPos = 0;
   end;
end;



//------------------------------------------------------------------------------
function InsertString (const Source :String; Pos:longword; const InString :String) :String;
begin
   if Pos = 1 then
   begin
      Result := InString + Source;
   end else begin
      Result := Copy(Source,1,Pos -1) + InString + Copy(Source,Pos,longword(length(Source)) - Pos + 1);
   end;
end;


//------------------------------------------------------------------------------
{ only RGB no ARGB     string is RRGGBB in HEX }
function string2RGB_REF (const ColorRef :String) :longword;
var r,g,b: longword;
    c:string;
begin
    c :='FF';
    Result := 0; // black;
    if length(ColorRef) = 6 then
    begin
      c[1]:=ColorRef[1];  c[2]:=ColorRef[2];  r := HEXval(c);
      c[1]:=ColorRef[3];  c[2]:=ColorRef[4];  g := HEXval(c);
      c[1]:=ColorRef[5];  c[2]:=ColorRef[6];  b := HEXval(c);
      Result := RGB(r,g,b);
    end;
end;

//------------------------------------------------------------------------------
function  RGB_REF2string (Color :longword) :String;
begin
   Result := toHEX(GetRValue(Color),2) + toHEX(GetGValue(Color),2) + toHEX(GetBValue(Color),2);
end;

//------------------------------------------------------------------------------
function ClearQuotes(const Atext:String) :String;
var l:longword;
begin
   Result := Trim(AText);
   l := length(Result);
   if l > 0 then
   begin
      if (Result[1] = '"') and (Result[l] = '"') then
      begin
         Result := Copy(AText, 2, l - 2);
      end;
   end;
end;

//------------------------------------------------------------------------------
function ParseStrSub (const inStr :String; id,subid :longword; delimiter,subdelimiter :Char) :String;
var substr :string;
begin                    // like  a=1&b=2&c=3
   substr := ParseStr(instr,id,delimiter);
   Result := ParseStr(substr,subid,subdelimiter);
end;

//------------------------------------------------------------------------------
function DwToStr(D:longword; Reverse:boolean = true):ansistring;
begin
   Result := 'xxxx';
   if Reverse then
   begin
      Result[4] := ansichar((D and $FF000000) shr 24);
      Result[3] := ansichar((D and $00FF0000) shr 16);
      Result[2] := ansichar((D and $0000FF00) shr 8);
      Result[1] := ansichar (D and $000000FF);
   end else begin
      Result[1] := ansichar((D and $FF000000) shr 24);
      Result[2] := ansichar((D and $00FF0000) shr 16);
      Result[3] := ansichar((D and $0000FF00) shr 8);
      Result[4] := ansichar (D and $000000FF);
   end;
end;

//------------------------------------------------------------------------------
function StrToDw(S:Ansistring; Reverse:boolean = true):longword;
begin
   Result := 0;
   if length(S) = 4 then
   begin
      if Reverse then
      begin
         REsult := (longword(S[4]) shl 24)
                or (longword(S[3]) shl 16)
                or (longword(S[2]) shl 8)
                or  longword(S[1]);
      end else begin
         REsult := (longword(S[1]) shl 24)
                or (longword(S[2]) shl 16)
                or (longword(S[3]) shl 8)
               or  longword(S[4]);
      end;
   end;
end;

//------------------------------------------------------------------------------
function ToStrZlead (v,d:longword) :string;
var s:string;
    i,j:longword;
begin
   SetLength(Result,d);
   s := ToStr(v);
   j := length(s);
   if j >= d then Result := s
            else begin
               j := d-j+1;
               for i := 1 to d do
               begin
                  if i >= j then Result[i] := s[i-j+1]
                            else Result[i] := '0';
               end;
            end;
end;

//------------------------------------------------------------------------------
function PZStrToStr(p:pointer; ansi:boolean=true) :string; // from ansi
var ch1:pansichar;
    ch2:pwidechar;
    i,j:longword;
begin
   Result := '';
   if p = nil then Exit;
   try
      j := PzStrLen(p,ansi);
      SetLength(Result,j);
      if ansi then
      begin
         ch1 := p;
         for i :=1 to j do
         begin
//         while ch1^ <> #0 do
//         begin
//            Result := Result + char(ch1^);
            Result[i] := char(ch1^);
            inc(ch1,1);
         end;
      end else begin
         ch2 := p;
         for i :=1 to j do
         begin
//         while ch2^ <> #0 do
//         begin
//            Result := Result + char(ch2^);
            Result[i] := char(ch2^);
            inc(ch2,1);
         end;
      end;
   except
      Result := '';
   end;
end;

//------------------------------------------------------------------------------
function PZStrLen(p:pointer; ansi:boolean=true) :longword; // from ansi
var ch1:pansichar;
    ch2:pwidechar;
begin
   Result := 0;
   if p = nil then Exit;
   try
      if ansi then
      begin
         ch1 := p;
         while ch1^ <> #0 do
         begin
            inc(Result);
            inc(ch1,1);
         end;
      end else begin
         ch2 := p;
         while ch2^ <> #0 do
         begin
            inc(Result);
            inc(ch2,1);
         end;
      end;
   except
      Result := 0;
   end;
end;


//------------------------------------------------------------------------------
procedure  CopyStrToPZStr(s:string; p:pointer; ansi:boolean=true); // must be allocated
var ch1:pansichar;   // must be allocated
    ch2:pwidechar;
    i,j:longword;         //todo
begin
   j := Length(s);
   if p = nil then Exit;
   try
      i := PZStrLen(p,ansi);
      if i < j then j:= i;
      i := 1;
      if ansi then
      begin
         ch1 := p;
         while (ch1^ <> #0)and (i<=j) do
         begin
            ansichar(ch1^) := ansichar(s[i]);
            inc(ch1,1);
            inc(i);
         end;
         ansichar(ch1^) := #0; // put end
      end else begin
         ch2 := p;
         while (ch2^ <> #0)and (i<=j) do
         begin
            widechar(ch2^) := widechar(s[i]);
            inc(ch2,1);
            inc(i);
         end;
         widechar(ch2^) := #0; // put end
      end;
   except
     //
   end;
end;


//------------------------------------------------------------------------------
function DataToStr(p:pointer; len:longword):ansistring;
var d:pointer;
begin
   SetLength(Result,len);
   d := @Result[1];
   Move(p^,d^,len);
end;

//------------------------------------------------------------------------------
function TrimIn(const S:string):string;
var m:string;
    c:char;
    i,j:longword;
begin
   Result := '';
   m := Trim(s);
   j := length(m);
   if j > 0 then
   begin
      i:=1;
      repeat
         c := m[i];
         Result := Result + c;
         inc(i);
         if c = ' ' then
         begin
            while m[i] = ' ' do inc(i); // no problem I have space only inside
         end;
      until i > j;
   end;
end;

//------------------------------------------------------------------------------
function SplitStr(const txt,delimiter:string):TStringDynArray;
var sc,dl,sz,i,t,m:longint;
    s:string;
    c:char;
begin
   Result := nil;
   dl := length(delimiter);
   sz := length(txt);
   if (dl = 0) or (sz = 0) then Exit;
   SetLength(s,sz+2);
   sc := 1;
   i := 0;
   m := 0;
   t := 1;
   repeat
      inc(i);
      inc(m);
      c := txt[i];
      if c = delimiter[t] then
      begin
         inc(t);
         if t > dl then // we have delimiter
         begin
            inc(sc);
            s[m] := #0;
            continue;
         end;
      end else t := 1;
      s[m] := c;
   until (i = sz);
   s[m+1]:= #0;

   SetLength(Result,sc);
   for i := 1 to sc do
   begin
      Result[i] := ParseStr(s,i-1,#0);
   end;
end;

function ToStrF(F:single; Width,Decimal:longword):string;
var s:shortstring;
begin
   str(F:Width:Decimal,s);
   Result := string(s);
end;

function ToStr64(i:int64):string;
var s:shortstring;
begin
   str(i,s);
   Result := string(s);
end;




(*
function urlEncode(s :string) :string;
var i:integer;
begin
   Result := '';
   for i:=1 to length(s) do
   begin
      case s[i] of
         ';': result:=result+'%3B';
         '?': result:=result+'%3F';
         '/': result:=result+'%2F';
         ':': result:=result+'%3A';
         '#': result:=result+'%23';
         '&': result:=result+'%26';
         '=': result:=result+'%3D';
         '+': result:=result+'%2B';
         '$': result:=result+'%24';
         ',': result:=result+'%2C';
         ' ': result:=result+'%20'; // %20 or + {deprecated}
         '%': result:=result+'%25';
         '<': result:=result+'%3C';
         '>': result:=result+'%3E';
         '~': result:=result+'%7E';
         else result:=result+s[i];
      end;
   end;
end;
*)

function HTTPDecode(const txt: string): string;
var c:char;
   h: string;
   i,l,m,r,j: integer;
begin
   l := Length(txt);
   Result := txt;
   m := 0;
   i := 1;
   repeat
      inc(m);
      c := txt[i];
      if c = '+' then
      begin
         Result[m] := ' ';
      end else begin
         if c = '%' then
         begin
            h := '$00';
            h[2] := txt[i+1];
            h[3] := txt[i+2];
            val(h,r,j);
            if j = 0 then Result[m] := char(r)
                     else Result[m] := ' ';
            inc(i,2);
         end else begin
            Result[m] := txt[i];
         end;
      end;
      inc(i);
   until (i>l);
   SetLength(Result, m);
end;

function HTTPEncode(const txt: string): string;
const
   HTTPAllowed =
   ['A'..'Z','a'..'z','*','@','.','_','-','0'..'9','$','!','''','(',')'];
var
   c :char;
   l,i: integer;
begin
   l := Length(txt);
   Result := '';
   for i := 1 to l do
   begin
      c := txt[i];
     if ansichar(c) in HTTPAllowed then
      begin
         Result := Result + c;
      end else begin
         if c = ' ' then
         begin
            Result := Result + '+';
         end else begin
            Result := Result + '%' + ToHex(Ord(c),2);
         end;
      end;
   end;
end;

function HTMLEncode(const txt: string): string;
var i,j:longint;
    c:char;
begin
   Result := '';
   j := length(txt);
   for i := 1 to j do
   begin
      c := txt[i];
      if c = '&' then begin Result := Result +'&amp';   c:=';'; end;
      if c = '<' then begin Result := Result +'&lt';    c:=';'; end;
      if c = '>' then begin Result := Result +'&gt';    c:=';'; end;
      if c = '"' then begin Result := Result +'&quot';  c:=';'; end;
      if c = ''''then begin Result := Result +'&apos';  c:=';'; end;
      if word(c)> 255 then
      begin
         Result := Result + '&#'+toStr(word(c));
         c := ';';
      end;
//      if c < #32 then
//      begin
  //TODO
//      end;
      Result := Result + c;
   end;
end;

function HTMLDecode(const txt: string): string;  // uses as XML_UnDecorate

var i,k,t:longword;
    c:char;
    Temp:String;
    ii,ii2:integer;
    done:boolean;
begin
   Result := '';
   Temp:='';
   i := Length(txt);
   k := 1;
   if i > 0 then
   begin
      repeat
         t := 0;
         c := txt[k];
         if c = '&' then
         begin
            done := false;
            t := 1;
            Temp := '&';
            repeat
               if k+t<= i then
               begin
                  c := txt[k+t];
                  Temp := Temp + c;
                  if Temp = '&amp;'  then begin c := '&'; done:=true; end;
                  if Temp = '&lt;'   then begin c := '<'; done:=true; end;
                  if Temp = '&gt;'   then begin c := '>'; done:=true; end;
                  if Temp = '&quot;' then begin c := '"'; done:=true; end;
                  if Temp = '&apos;' then begin c := ''''; done:=true; end;
                  if Temp = '&#' then
                  begin
                     // colect till ;
                     inc(t);
                     Temp := '';
                     repeat
                        c := txt[k+t];
                        if c = ';' then break;
                        Temp := Temp + c;
                        inc(t);
                     until ((k+t)>i);
                     if length(Temp) > 0  then
                     begin
                        If Temp[1] = 'x' then //HEX
                        begin
                           //Todo

                        end else begin
                           val(string(Temp),ii,ii2);
                           if ii = 0 then ii:=32;
                           c := char(ii);
                        end;
                        done := true;
                     end;
                  end;
               end;
               inc(t);
            until (done) or((k+t)>i) or (t > 8) ;
            if not done then t := 0 else dec(t);
         end;
         Result := Result + c;
         inc(k,1+t);
      until (k > i);
   end;
(*

var i,k,Acumolator:longword;
    c:char;
    Temp,s:String;
    ii,ii2:integer;
begin
   Result := '';
   Temp:='';
   i := Length(txt);
   Acumolator := 0;
   for k:= 1 to i do
   begin
      c := txt[k];
      if Acumolator = 1 then
      begin
         Temp := Temp + c;
         if c = ';' then Acumolator := 0;
         if Temp = '&amp;'  then c := '&';
         if Temp = '&lt;'   then c := '<';
         if Temp = '&gt;'   then c := '>';
         if Temp = '&quot;' then c := '"';
         if Temp = '&apos;' then c := '''';
         if (Acumolator = 0) and (length(temp) > 3) and (Temp[1]='&') and (Temp[2]='#') then // &# Dec val ;
         begin
            s := Copy(Temp,3,length(Temp)-3);
            val(s,ii,ii2);
            if ii = 0 then ii:=32;
            c := char(ii);
         end;
         if (Acumolator = 0) and (length(temp) > 4) and (Temp[1]='&') and (Temp[2]='#') and (Temp[3]='x')then // &#x HEXval ;
         begin
            s := Copy(Temp,4,length(Temp)-4);
            ii := HexVal(s);
            if ii = 0 then ii:=32;
            c := char(ii);
         end;
      end;
      if c = '&' then begin Acumolator := 1; Temp := '&'; end;
      if Acumolator = 1 then continue;
      Result := Result + c;
   end;
*)
end;


function PosA(const SearchTxt,Txt:ansistring):longword;
var i,j,k,m,p:longword;
begin
   p := 0;
   j := length(Txt);
   k := length(SearchTxt);
   if (j>0) and (k>0) then
   begin
      m := 1;
      for i := 1 to j do
      begin
         if Txt[i] = SearchTxt[m] then
         begin
            if m = 1 then p := i;
            if m = k then break;
            inc(m);
         end else begin
            p := 0;
            m := 1;
         end;
      end;
   end;
   Result := p;
end;

function FastCPosW(a:widechar; const s:widestring):longword;
//{$IFDEF CPUX64}
var j:longword;
//{$ENDIF}
begin
//{$IFDEF CPUX64}
   Result := 0;
//870
   j := length(s);
   if j > 0 then
   begin
      Result := 1;
      while a <> s[Result] do begin inc(Result); if Result > j then break; end;
      if Result > j then Result := 0;
   end;
{slower
//1200
   for i:= 1 to j do
   begin
      if a = s[i] then
      begin
         Result := i;
         break
      end;
   end;
}
(*
{$ELSE}
   asm
      push  edi
      mov   eax, s
      test  eax, eax
      jz    @@out
      mov   ecx, dword ptr [eax - 4]
      mov   edi, eax
      mov   eax, 0
      mov   dx, a
@@lop:
      cmp   dx, word ptr [edi + eax]
      je    @@res
      lea   eax, [eax + 2]
      loop  @@lop
      mov   eax, 0
      jmp   @@out
@@res:
      shr   eax,1
      inc   eax
@@out:
      pop   edi
      mov   Result, eax
   end;
{$ENDIF}
*)
end;

function FastCPosA(a:ansichar; const s:ansistring):longword;
{$IFDEF CPUX64}
var j:longword;
{$ENDIF}
begin
{$IFDEF CPUX64}
   Result := 0;
   j := length(s);
   if j > 0 then
   begin
      Result := 1;
      while a <> s[Result] do begin inc(Result); if Result > j then break; end;
      if Result > j then Result := 0;
   end;
{$ELSE}
   asm
      push  edi
      mov   eax, s
      test  eax, eax
      jz    @@out
      mov   ecx, dword ptr [eax - 4]
      mov   edi, eax
      mov   eax, 0
      mov   dl, a
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
{$ENDIF}
end;

//------------------------------------------------------------------------------
function WideCP1251(const txt :widestring):ansistring;
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
         if c > 255 then
         begin
            w := c;
            if (c >= $410) and (c <= $44F) then c := 192 + (c - $410)
                                           else c := 63; //'?'-unknown
            //todo for other special chars
            if w = $2116 then c:= 185; //nomer
         end;
         Result[i] := ansichar(c);
      end;
   end;
end;

//------------------------------------------------------------------------------
function CP1251Wide(const txt :ansistring):widestring;
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
            if (c >= 192) and (c <= 255) then c := $410 + (c - 192)
                                           else c := 63; //'?'
            //todo for other special chars
            if w = 185 then c := $2116 //nomer
         end;
         Result[i] := widechar(c);
      end;
   end;

end;


end.
