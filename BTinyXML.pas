unit BTinyXML;

interface

// NOTE working only with AnsiString for 8but UTF8 format
// Search path get elemet        /root/item
//             get attibute      /root/item.atr
//             get Sequentional  /root/item[3] start from 1   !!!!
//                               /root/item[3]/element
//                               /root/item[3].atr
// Flags return error (100-not found) 0-OK

// totaly rewrited old wersion have an error if in sequence we have differnt elemet
// for example <A><B><B><C><B><B><A>   so /A/B[3] did not work
// now it works but it is slow for huge xmls or intesive work
// so i write a tiny class work well for me , for now


function  TinyXML_Parse(const in_XML:AnsiString; const in_Search:AnsiString; var aFlags:longint ) :AnsiString;

function  XML_Document (const RootName, Value :AnsiString; standalone:boolean = false) :AnsiString;
function  XML_AddNode (const  Name, Attribute, data :AnsiString) :AnsiString;
function  XML_DoAttribute (const  Name, value :AnsiString) :AnsiString;
function  XML_DoCDATA (const  value :AnsiString) :AnsiString;

function  XML_Decorate (const value :AnsiString) :AnsiString;
function  XML_UnDecorate (const value :AnsiString) :AnsiString;




type      PBTTinyXMLAttr = ^BTTinyXMLAttr;
          BTTinyXMLAttr = record
             Name : string;
             Value : string;
             Next : PBTTinyXMLAttr;
          end;

          PBTTinyXMLNode = ^BTTinyXMLNode;
          BTTinyXMLNode = record
             Name : string;
             Attributes : PBTTinyXMLAttr;
             ChildNodes : PBTTinyXMLNode;
             Data : string;
             Next : PBTTinyXMLNode;
             Papa : PBTTinyXMLNode;
             indx : longword;
             rem : string;
             cdata : boolean;
          end;


          BTTinyXML = class
             private
                xmlheader :string;
                function    _setgetXpath(const path:string; var value:string; verb,aFlags:longword):longint;
             public
                Nodes : PBTTinyXMLNode;
                constructor Create;
                destructor  Destroy; override;
                function    LoadXML(const in_xml:string; aFlags:longword = 0):longint;
                function    GetXML(var out_xml:string):longint;
                function    SelectXPath(const path:string; var res:string; aFlags:longword = 0):longint;
                function    UpdateXPath(const path:string; value:string; aFlags:longword = 0):longint;
                function    SelectSingleNode(Node:NativeUInt; const NodeName:string; aFlags:longword = 0):NativeUInt; // 0 = root
                function    GetAttribute(Node:NativeUInt; const AttributeName:string):string;
                function    SetAttribute(Node:NativeUInt; const AttributeName,AttributeValue:string):boolean;
                function    GetText(Node:NativeUInt):string;
                function    SetText(Node:NativeUInt; const Value:string; cdata:boolean = false):boolean;
                function    AddElement(Node:NativeUInt; const NodeName:string):NativeUInt;
                function    AddChild(Node:NativeUInt; const NodeName:string):NativeUInt;
                function    DeleteNode(Node:NativeUInt):boolean;
                function    NodeName(Node:NativeUInt; aFlags:longword = 0):string;
                function    FirstChild(Node:NativeUInt):NativeUInt;
                function    NextSibling(Node:NativeUInt):NativeUInt;
                function    AddComment(Node:NativeUInt; const comment:string):boolean;
                //todo xml header encoding
          end;


implementation

uses BStrTools;

// XML READER //////////////////////////////////////////////////////////////////


const Preserve_namespace = $00000001;
      Return_path        = $00000002;
      Undecorate         = $00000004;
      Retrun_Encoding    = $00000008;

function  TinyXML_Parse(const in_XML:AnsiString; const in_Search:AnsiString; var aFlags:longint ) :AnsiString;
var c,opc:ansichar;
    xSize,ofs,err,mode,i,j,lastpopindx,cut:longint;
    tmp,data,stack,lastpop,tag,thetag,res,path,{enc,}t2:ansistring;
    Search:ansistring;
    Attr_need,series,closetag,delm,gottag,fillRes:boolean;
    Attr:ansistring;



    function  _FindTagIndex(const t:ansistring):longword;
    var _i,_j,_m,_k:longword;
        _c:ansichar;
        s:string;
    begin
       Result := 1;
       lastpopindx := 1;
       lastpop := '';
       _j := length(Path);
       _m := 2;
       _k := 0;
       for _i := 1 to _j do
       begin
          _c := Path[_i];
          if _c = #29 then
          begin
             if Path[_i+1] = 'T' then
             begin
                if length(s) <> 0  then
                begin
                 lastpop := Copy(lastpop,1,length(lastpop) - (length(s)+2));
                end;
                if lastpop = t then
                begin
                   Result := lastpopindx + 1;
                end;
             end;
             if _m = 0 then _m := 2;
          end;
          if _m = 0 then
          begin
             if _k = 1 then
             begin
                if _c = ']' then
                begin
                   val(string(s),lastpopindx,_k);
                   _k := 0;
                end else s := s + char(_c);
             end;
             lastpop := lastpop + _c;
             if _c = '[' then
             begin
                s := '';
                _k := 1;
             end;
          end;
          if _c = #28 then
          begin
             _k := 0;
             _m := 0;
             lastpop := '';
             s := '';
          end;
          if _c = '/' then s := '';
       end; //for
    end;

    procedure _cutter(i:longint);
    begin
       if length(res) >= i then Res := Copy(Res,1,length(Res)-i);
    end;

    function _pusht(t:ansistring):boolean;
    var s:shortstring;
        _ti:longword;
    begin
       t2 := t;
       Result := false;
       Stack := stack + '/' + t;
       _ti := _FindTagIndex(Stack);
       if _ti > 1 then
       begin
          str(_ti,s);
          Stack := Stack + '[' + s + ']';
       end;
       Path := Path + #28 + Stack + #29 + t2 + #29 + 'T'; //ToHex(ofs-cut,8);
       if stack = search then Result := true;
       lastpop := '';
    end;

    procedure _popt(const t:ansistring);
    var _i,_j,_f:longint;
    begin
       if (not Attr_need) and GotTag and (TheTag = Stack) then
       begin
          err :=0; // found
          _cutter(cut);
       end;
       _j := length(Stack);
       _f := 0;
       for _i := 1 to _j do if Stack[_i] = '/' then _f := _i; // get last marker
       if _f > 0 then
       begin
//          lastpop := Copy(Stack,_f + 1, _j - _f);
//          _j := pos('[',string(lastpop));
//          if _j <> 0 then
//          begin
//             s := Copy(lastpop,_j+1,length(lastpop)-_j-1);
//             val(string(s),lastpopindx,_i);
//             lastpop := Copy(lastpop,1,_j-1);
//          end;
          Stack := Copy(Stack,1, _f - 1);
       end;
       if  closetag then Path := Path + #28 + lastpop + #29 + t + #29 + 'E' //ToHex(ofs-cut,8);
                    else Path := Path + #28 + lastpop + #29 + t + #29 + 'F';
    end;

    function _peek(s:AnsiString):boolean;
    var _j,_k,_t:longint;
    begin
       Result := false;
       _j := length(s);
       _t := 0;
       if (ofs + _j) <= xSize then
          for _k := 1 to _j do if in_xml[ofs + _k] = s[_k] then inc(_t);
       if _t = _j then
       begin
          if Fillres then Res := Res + s;
          inc(ofs,_j);
          Result := true;
       end;
    end;

// rules ater < no space just name fo tag
begin

   // Set default values
   err := 100;  //not found
   Result := '';
   Res := '';
   Path := '';
//   Enc := '';

   Stack := '';
//   StackHistory := '';

   // Adjust search pattern
   tmp := '';
   Search := '';
   Attr := '';
   Attr_need := false;
   series := false;
   i := length(in_Search);
   if i < 2 then  // min /a - 2 chars
   begin
      if (aFlags and Return_path) = 0 then Exit;  // in callback mode search emnpty
   end else begin    // Adust search data
      for j:=1 to i do
      begin
         c := in_Search[j];
         // test for good chars
         //TODO
         if c = '\' then c:= '/';
         if (j = 1) and (c <> '/') then Search := '/'; //put first if not exist
         if c = '.' then
         begin
            if Attr_need then err := -1; // morw that one attr is error
            Attr_need := true;
            continue;
         end;
         if Attr_need  then Attr := Attr + c
                       else begin
                          if c = '[' then series := true; // open
                          if series then tmp := tmp + c;
                          Search := Search + c;
                          if series and (c = ']')then
                          begin
                             if tmp = '[1]' then // remove series by 1
                             begin
                                tmp := Copy(Search,1,length(Search)-3);
                                Search := tmp;
                             end;
                             tmp := '';
                             series := false;
                          end;
                       end;
      end;
   end;


   //parser
   xSize := length(in_XML);
   mode := 0; // starting parser mode
   tmp := '';
   ofs := 0;  // offset in_XML
   delm := false;
   opc := #0;
   data := '';
   lastpop := '';
   lastpopindx := 1;
   GotTag := false;
   TheTag := '';
   FillRes := false;  // Fill res flag
   cut := 0;
   closetag := false;
   while  (ofs < xSize) and (err=100) do
   begin
      inc(ofs); // Next input char
      inc(cut);
      c := in_XML[ofs];

    (*
      if (mode=0) and (c = '&') then begin tmp := ''; mode := 1; end; //undecore only starting and ending
      if (mode=1) then
      begin
         if c = ';' then
         begin
            if tmp = '&amp'  then c := '&';
            if tmp = '&lt'   then c := '<';
            if tmp = '&gt'   then c := '>';
            if tmp = '&quot' then c := '"';
            if tmp = '&apos' then c := '''';
            if tmp = '&#39'  then c := '''';
            tmp := '';
            mode := 0;
         end else begin
            tmp := tmp + c;
            continue;
         end;
      end;
    *)
      if FillRes then Res := Res + c;

      if mode = 5 then  // we have begin tag <? or <!--
      begin
         if (c = '?') and _peek('>') then mode := 0;
         if (c = '-') and _peek('->') then mode := 0;
         if mode = 0 then // extract encoding
         begin
            if (aFlags and Retrun_Encoding) <> 0 then
            begin
               Result := tmp;
               aFlags := 0;
               Exit;
            end;
            tmp := '';

         end else tmp := tmp + c; // fill tmp without end
         continue;
      end;

      if mode = 6 then  // cdata mode
      begin
         if (c = ']') and _peek(']>') then
         begin
            mode := 0;
         end else res := res + c;
         continue;
      end;

      if mode = 4 then
      begin
         if opc = #0 then
         begin
            if (c = '''') or (c = '"') then opc := c;
            continue;
         end else begin
            if c <> opc then data := data + c
                        else begin
                           Path := Path + #28 + Stack + '.' + tmp + #29 + tmp + #29 + 'A'; //ToHex(ofs-cut,8);
                           //BBBB = ofs   AAAA = ofs - length(data) + 1
                           if GotTag and Attr_need and (tmp = attr) then
                           begin
                              Res := data;
                              err := 0;   //Exit
                           end else begin
                              data := '';
                              opc := #0;
                              mode := 2; // still in tag this is attribute
                           end;
                        end;
         end;
         continue;
      end;

      if (mode=0) and (c = '<') then  // begining of the parsing
      begin
         cut := 1;
         tmp := '';
         mode := 2; // tag start
         delm := false;
         closetag := false;
         continue; //next char is the tag name no space by rule
      end;

      if (mode=3) then //read tag name
      begin
         if (c = ' ') or (c = '/') or (c = '>') then //got name
         begin
            mode := 2; // go back to inside
            tag := tmp;
            tmp := ''; //preaper to store attrib
            if closetag then
            begin
               _popt(tag);
            end else begin
               if _pusht(tag)then
               begin
                  GotTag := true;
                  TheTag := tag;
               end;
            end;
            if c = ' ' then
            begin
               delm := true;
               continue;
            end;
         end else begin
            tmp := tmp + c;
            if ((aFlags and preserve_namespace) = 0) and (c = ':') then tmp := ''; // ignore name space
            continue;
         end;
      end;

      if (mode=2) then // tag start inside
      begin //TAG name read - open
         if cut = 2 then //tag start with  (seconf char)
         begin
            case c of
               '?' : mode := 5;// encoding <? .... ?>
               '!' : begin
                        if _peek('--')      then mode := 5; //comment  <!-- .... -->
                        if _peek('[CDATA[') then mode := 6; //cdata   <![CDATA[ .... ]]>
                     end;
               '/' : begin closetag := true;
               mode := 3; end; //end tag start name
               else begin tmp := c; mode := 3; end; // name of tag begin
            end;
            continue;
         end;

         if (c = '/') and _peek('>') then // force close of open tag
         begin
            _popt(tag);
            mode := 0;
            continue;
         end;

         if c = '>' then // End of tag
         begin
            if GotTag then
            begin
               if Attr_need then
               begin
                  break;    // tag not found
               end;
               FillRes := true;
               if closetag and (TheTag = tag) then
               begin
                  _Cutter(cut);
                  err := 0;
                  continue;
               end;
            end;
            mode := 0;
            continue;
         end;

         if delm then // I have delimiter possible attribute
         begin
            if c = '=' then
            begin
               mode := 4; // read attribute value
               continue;
            end;
            if c = ' ' then // new delimiter may be new attr
            begin
               tmp := '';
               continue;
            end;
            tmp := tmp + c; //acum attrib
            if ((aFlags and preserve_namespace) = 0) and (c = ':') then tmp := ''; // ignore name space
            if closetag then
            begin
               err := - 12;
               break;
            end;
         end;
      end;
   end;


   if err = 0 then Result := Res;
   if (aFlags and Return_path) <> 0 then
   begin
      Result := Path;
      err := 0;
      if length(stack) <> 0 then err := -3;
   end;
   if (aFlags and Undecorate) <> 0 then Result := XML_UnDecorate(Result);

   aFlags := err;
end;






// XML WRITER //////////////////////////////////////////////////////////////////

//------------------------------------------------------------------------------
function XML_DoAttribute(const Name,value:AnsiString):AnsiString;
begin
   Result := ' ' + Name + '="'+ value +'"';
end;

//------------------------------------------------------------------------------
function XML_AddNode(const Name,Attribute,data:AnsiString):AnsiString;
begin
   if data = '' then
   begin
      Result := '<'+Name+Attribute+'/>';
   end else begin
      Result := '<'+Name+Attribute+'>'+data+'</'+Name+'>';
   end;
end;

//------------------------------------------------------------------------------
function XML_DoCDATA(const value:AnsiString):AnsiString;
begin
   Result := '<![CDATA['+value+']]>';
end;

//------------------------------------------------------------------------------
function XML_Document(const RootName,Value:AnsiString; standalone:boolean = false):AnsiString;
begin
   if standalone then Result := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><'+Rootname+'>'+value+'</'+RootName+'>'
                 else Result := '<?xml version="1.0" encoding="UTF-8"?><'+Rootname+'>'+value+'</'+RootName+'>';
end;

//------------------------------------------------------------------------------
function  XML_Decorate (const value :AnsiString) :AnsiString;
begin
   Result := ansistring(HTMLEncode(string(value)));
end;

//------------------------------------------------------------------------------
function  XML_UnDecorate (const value :AnsiString) :AnsiString;
begin
   Result := ansistring(HTMLDecode(string(value)));
end;




//==============================================================================
{  TTTTTT  II  NN  NN  YY  YY   XX   XX  MMM MMM  LL       by Bogi aka SDEX32
     TT    II  NNN NN   YYYY      XXX    MM M MM  LL       vesrion 2
     TT    II  NN NNN    YY       XXX    MM   MM  LL       11.2022
     TT    II  NN  NN    YY     XX   XX  MM   MM  LLLLLL   TinyXML
}

//------------------------------------------------------------------------------
constructor BTTinyXML.Create;
begin
   Nodes := nil;
   xmlheader := '';
end;

//------------------------------------------------------------------------------
destructor  BTTinyXML.Destroy;
begin
   try
      if Nodes <> nil then  DeleteNode(NativeUInt(Nodes));
   except;
      Nodes := nil;
   end;
   Nodes := nil;
   inherited;
end;

//------------------------------------------------------------------------------
function    BTTinyXML.LoadXML(const in_xml:string; aFlags:longword = 0):longint;
var n,nm,nn,nc,ct:PBTTinyXMLNode;
    a,an:PBTTinyXMLAttr;
    err : longint;
    ofs,xSize,mode,cut:longint;
    c,opc:char;
    tmp,Res,data:string;
    FillRes,delm,closetag,cdata:boolean;

    procedure _AddAttr(const Name,value:string);
    begin
       New(an);
       an.Name := Name;
       an.Value := Value;
       an.Next := nil;
       if n.Attributes <> nil then
       begin
          a := n.Attributes;
          while (a.Next <> nil) do a := a.Next;
          a.Next := an;
       end else n.Attributes := an;
    end;

    procedure _pusht(const t:string);
    var i :longword;
    begin
       New(nn);
       nn.Name := t;
       nn.Attributes := nil;
       nn.ChildNodes := nil;
       nn.Data := '';
       nn.Next := nil;
       nn.Papa := nil;
       nn.indx := 1;
       nn.Rem := '';
       nn.cdata := cdata;
       cdata := false;
       i := 1;
       if Nodes <> nil then
       begin
          nm := n;
          if nm.ChildNodes <> nil then
          begin
             nc := nm.ChildNodes;
             if nc.name = nn.name then inc(i);
             while nc.next <> nil do
             begin
                nc := nc.Next;
                if nc <> nil then if nc.name = nn.name then inc(i);
             end;
             nc.next := nn;
             nn.indx := i;
          end else nm.ChildNodes := nn;
          nn.Papa := n;
       end else Nodes := nn;
       n := nn; // n = current tag
    end;

    procedure _popt;
    begin
       n := n.Papa;
    end;

    function _peek(_s:string):boolean;
    var _j,_k,_t:longint;
    begin
       Result := false;
       _j := length(_s);
       _t := 0;
       if (ofs + _j) <= xSize then
          for _k := 1 to _j do if in_xml[ofs + _k] = _s[_k] then inc(_t);
       if _t = _j then
       begin
          if Fillres then Res := Res + _s;
          inc(ofs,_j);
          Result := true;
       end;
    end;

begin
   opc := #0;
   err := 0;
   mode := 0;
   FillRes := false;
   closetag := false;
   delm := false;
   ofs := 0;
   cut := 0;
   tmp := '';
   res := '';
   data := '';
   xmlheader := '';
   ct := nil;
   cdata := false;
   try
   xSize := length(in_xml);
   if xSize > 0 then
   begin
     while  (ofs < xSize) and (err=0) do
      begin
         inc(ofs); // Next input char
         inc(cut);
         c := in_XML[ofs];

         if FillRes then Res := Res + c;     //TODO ne raboti s undecaret ?!?

         if (mode=0) and (c = '&') then begin tmp := ''; mode := 1; end; //undecore only starting and ending
         if (mode=1) then
         begin
            if c = ';' then
            begin
               if tmp = '&amp'  then c := '&';
               if tmp = '&lt'   then c := '<';
               if tmp = '&gt'   then c := '>';
               if tmp = '&quot' then c := '"';
               if tmp = '&apos' then c := '''';
               if tmp = '&#39'  then c := '''';
               tmp := '';
               mode := 0;
            end else begin
               tmp := tmp + c;
               continue;
            end;
         end;

         if mode = 5 then  // we have begin tag <? or <!--
         begin
            if (c = '?') and _peek('>') then mode := 0;
            if (c = '-') and _peek('->') then mode := 0;
            if mode = 0 then // extract encoding
            begin
               xmlheader := Copy(tmp,4,length(tmp)-3);  //store encoding header
               tmp := '';
            end else tmp := tmp + c; // fill tmp without end
            continue;
         end;

         if mode = 6 then  // cdata mode
         begin
            if (c = ']') and _peek(']>') then
            begin
               cdata := true;
               mode := 0;
            end else res := res + c;
            continue;
         end;

         if mode = 4 then
         begin
            if opc = #0 then
            begin
               if (c = '''') or (c = '"') then opc := c;
               continue;
            end else begin
               if c <> opc then data := data + c
                           else begin
                              _AddAttr(tmp,data);
                              data := '';
                              tmp := '';
                              opc := #0;
                              mode := 2; // still in tag this is attribute
                           end;
            end;
            continue;
         end;

         if (mode=0) and (c = '<') then  // begining of the parsing
         begin
          //  FillRes := false;
            cut := 1;
            tmp := '';
            mode := 2; // tag start
            delm := false;
            closetag := false;
            continue; //next char is the tag name no space by rule
         end;

         if (mode=3) then //read tag name
         begin
            if (c = ' ') or (c = '/') or (c = '>') then //got name
            begin
               mode := 2; // go back to inside
               if closetag then
               begin
                  ct := n; //close tag before pop
                  _popt;
               end else begin
                  _pusht(tmp);
               end;
               tmp := ''; //prepare to store attrib
               if c = ' ' then
               begin
                  delm := true;
                  continue;
               end;
            end else begin
               tmp := tmp + c;
               if ((aFlags and preserve_namespace) = 0) and (c = ':') then tmp := ''; // ignore name space
               continue;
            end;
         end;

         if (mode=2) then // tag start inside
         begin //TAG name read - open
            if cut = 2 then //tag start with  (seconf char)
            begin
               case c of
                  '?' : mode := 5;// encoding <? .... ?>
                  '!' : begin
                           if _peek('--')      then mode := 5; //comment  <!-- .... -->
                           if _peek('[CDATA[') then mode := 6; //cdata   <![CDATA[ .... ]]>
                        end;
                  '/' : begin closetag := true;  mode := 3; end; //end tag start name
                  else begin tmp := c; mode := 3; end; // name of tag begin
               end;
               continue;
            end;

            if (c = '/') and _peek('>') then // force close of open tag
            begin
               _popt; //tag
               mode := 0;
               continue;
            end;

            if c = '>' then // End of tag
            begin
               FillRes := true;
               if closetag then
               begin
                  if length(res) >= cut then Res := Copy(Res,1,length(Res)-cut);
                  if ct <> nil then ct.Data := res;
                  FillRes := false;
                  res := '';
               end;
               res := '';
               mode := 0;
               continue;
            end;

            if delm then // I have delimiter possible attribute
            begin
               if c = '=' then
               begin
                  mode := 4; // read attribute value
                  continue;
               end;
               if c = ' ' then // new delimiter may be new attr
               begin
                  tmp := '';
                  continue;
               end;
               tmp := tmp + c; //acum attrib
               if ((aFlags and preserve_namespace) = 0) and (c = ':') then tmp := ''; // ignore name space
               if closetag then
               begin
                  err := - 12;
                  break;
               end;
            end;
         end;
      end;
   end else err := -2;
   except
      err := -1;
   end;
   Result := err;
end;

//------------------------------------------------------------------------------
procedure   _RecGet(n:PBTTinyXMLNode; var out_xml:string);
var a:PBTTinyXMLAttr;
    c:PBTTinyXMLNode;
begin
   if length(n.rem) > 0  then out_xml := out_xml + '<!--' + n.rem + '-->';

   out_xml:= out_xml + '<' + n.Name;
   a := n.Attributes;
   while a <> nil do
   begin
      out_xml := out_xml + ' '+a.Name+'="'+a.Value+'"';
      a := a.Next;
   end;
   if (length(n.Data) > 0) or (n.ChildNodes <> nil) then
   begin
      out_xml := out_xml + '>';
      if n.ChildNodes <> nil then
      begin
         c := n.ChildNodes;
         while c <> nil do
         begin
            _RecGet(c,out_xml);
            c := c.Next;
         end;
      end else begin
         if n.cdata then out_xml:= out_xml + '<![CDATA['+ n.Data +']]>'
                    else out_xml:= out_xml + n.Data;
      end;
      out_xml:= out_xml + '</' + n.Name+'>';
   end else begin
      out_xml := out_xml + '/>';
   end;
end;

function    BTTinyXML.GetXML(var out_xml:string):longint;
begin
   Result := 0;
   out_xml := '';
   try
      if Nodes <> nil then
      begin
         if length(xmlheader)> 0 then out_xml := out_xml + '<?xml '+xmlheader+'?>'#13#10;
         _RecGet(Nodes,out_xml);
      end;
   except
      Result := -1;
   end;
end;

//------------------------------------------------------------------------------
procedure   _GetNode(n:PBTTinyXMLNode; stack:string; const path:string; var found:PBTTinyXMLNode; aFlags:longword);
var c:PBTTinyXMLNode;
    t:string;
    i:longint;
begin
   if n = nil then Exit;
   if found <> nil then Exit;

   t := n.Name;
   if (aFlags and Preserve_namespace) = 0 then // clear name space
   begin
      i := Pos(':',t);
      if i <> 0 then t := Copy(t,i+1,length(t)-i);
   end;
   stack := stack + '/' + t;

   if n.indx > 1 then stack := stack + '['+ToStr(n.indx)+']';
   if stack = path then
   begin
      found := n;
      Exit;
   end;

   c := n.ChildNodes;
   while c <> nil do
   begin
      _GetNode(c,stack,path,found,aFlags);
      if found <> nil then Exit;
      c := c.Next;
   end;
end;

function    _FindNode(n:PBTTinyXMLNode; const path:string; var Attr_need:boolean; var Attr:string; aFlags:longword; var err:longint):PBTTinyXMLNode;
var i,j:longword;
    tmp,Search:string;
    c:char;
    series:boolean;
begin
   series :=false;
   Attr_need:=false;
   tmp:='';
   i := length(Path);
   for j:=1 to i do
   begin
      c := path[j];
      // test for good chars
      if c = '\' then c:= '/';
      if (j = 1) and (c <> '/') then Search := '/'; //put first if not exist
      if c = '.' then
      begin
         if Attr_need then err := -10; // morw that one attr is error
         Attr_need := true;
         continue;
      end;
      if Attr_need  then Attr := Attr + c
                    else begin
                            if c = '[' then series := true; // open
                            if series then tmp := tmp + c;
                            Search := Search + c;
                            if series and (c = ']')then
                            begin
                               if tmp = '[1]' then // remove series by 1
                               begin
                                  tmp := Copy(Search,1,length(Search)-3);
                                  Search := tmp;
                               end;
                               tmp := '';
                               series := false;
                            end;
                         end;
   end;
   Result := nil;
   _GetNode(n,'',search,Result,aFlags);
end;

function    BTTinyXML._setgetXpath(const path:string; var value:string; verb,aFlags:longword):longint;
var f:PBTTinyXMLNode;
    A,A2:PBTTinyXMLAttr;
    attr_need:boolean;
    attr:string;
    err :longint;
begin
   Result := 100; //not found;
   attr := '';
   attr_need := false;
   if verb = 0 then value := '';
   try
      err := 0;
      f := _FindNode(Nodes,path,attr_need,attr,aFlags,err);
      if err <> 0 then
      begin
         Result := err;
         Exit;
      end;

      if f <> nil then
      begin
         if attr_need then
         begin
            A2:= nil;
            A := f.Attributes;
            while (A<> nil) do
            begin
               if A.Name = Attr then A2 := A;
               A := A.Next;
            end;
            if A2 <> nil then
            begin
               if verb = 0 then value := A2.Value
                           else A2.Value := value;
               Result := 0;
            end;
         end else begin
            if verb = 0 then value := f.Data
                        else f.Data := value;
            Result := 0;
         end;
      end;
   except
      Result := -1;
   end;
end;

//------------------------------------------------------------------------------
function    BTTinyXML.SelectXPath(const path:string; var res:string; aFlags:longword = 0):longint;
begin
   Result := _setgetXpath(path,res,0,aFlags);
end;

//------------------------------------------------------------------------------
function    BTTinyXML.UpdateXPath(const path:string; value:string; aFlags:longword = 0):longint;
begin
   Result := _setgetXpath(path,value,1,aFlags);
end;

//------------------------------------------------------------------------------
procedure   _DelRec(n,root:PBTTinyXMLNode);
var c,a,b:PBTTinyXMLNode;
begin
   //1. Dispose childrens
   c := n.ChildNodes;
   while c <> nil do
   begin
      _DelRec(c,root);
      c := c.Next;
   end;
   c := n.Papa;
   //2. Find node in papa child list and unlink;
   if c = nil then c := root; //root
   begin
      a := c.ChildNodes;
      if a <> nil then
      begin
         if a = n then //FirstChild
         begin
            c.ChildNodes := a.Next;
         end else begin
             b := a.Next;
            while b <>nil do
            begin
               if b = n then
               begin
                  a.Next := b.Next;
                  break;
               end;
               b := b.Next;
               a := a.Next;
            end;
         end;
      end;
   end;
   Dispose(n);
end;

function    BTTinyXML.DeleteNode(Node:NativeUInt):boolean;
begin
   Result := true;
   try
      _DelRec(PBTTinyXMLNode(Node),Nodes);
   except
      Result := false;
   end;
end;

//------------------------------------------------------------------------------
function    BTTinyXML.GetAttribute(Node:NativeUInt; const AttributeName:string):string;
var a:PBTTinyXMLAttr;
begin
   Result := '';
   try
      a := PBTTinyXMLNode(Node).Attributes;
      while a <> nil do
      begin
         if a.Name = AttributeName then
         begin
            Result := a.Value;
            Break;
         end;
         a := a.Next;
      end;
   finally
      Result := '';
   end;
end;

//------------------------------------------------------------------------------
function    BTTinyXML.SetAttribute(Node:NativeUInt; const AttributeName,AttributeValue:string):boolean;
var a,b:PBTTinyXMLAttr;
begin
   Result := false;
   try
      a := PBTTinyXMLNode(Node).Attributes;
      while a <> nil do
      begin
         if a.Name = AttributeName then
         begin
            a.Value := AttributeValue;
            Result := true;
            Break;
         end;
         a := a.Next;
      end;
      if not Result then // new Attr
      begin
         New(b);
         b.Name := AttributeName;
         b.Value := AttributeValue;
         b.Next := nil;

         a := PBTTinyXMLNode(Node).Attributes;
         if a = nil then
         begin
            PBTTinyXMLNode(Node).Attributes := b;
         end else begin
            while a.Next <> nil do a := a.Next;
            a.Next := b;
         end;
      end;
   finally
      Result := False;
   end;
end;

//------------------------------------------------------------------------------
function    BTTinyXML.GetText(Node:NativeUInt):string;
var n:PBTTinyXMLNode;
begin
   Result := '';
   try
      n := PBTTinyXMLNode(Node);
      Result := n.Data;
   except
      Result := '';
   end;
end;

//------------------------------------------------------------------------------
function    BTTinyXML.SetText(Node:NativeUInt; const Value:string; cdata:boolean = false):boolean;
var n:PBTTinyXMLNode;
begin
   Result := False;
   try
      n := PBTTinyXMLNode(Node);
      if n.ChildNodes = nil then
      begin
         if length(Value) > 0 then
         begin
            n.Data := Value;
            n.cdata := cdata;
            Result := true;
         end;
      end;
   except
      Result := False;
   end;
end;

//------------------------------------------------------------------------------
function    BTTinyXML.NodeName(Node:NativeUInt; aFlags:longword = 0):string;
var n:PBTTinyXMLNode;
    i:longword;
begin
   Result := '';
   try
      n := PBTTinyXMLNode(Node);
      Result := n.Name;
      if (aFlags and Preserve_namespace) = 0 then
      begin
         i := Pos(':',Result);
         if i <> 0 then  Result := Copy(Result,i+1,length(Result)-i);
      end;
   except
      Result := '';
   end;
end;

//------------------------------------------------------------------------------
function    BTTinyXML.FirstChild(Node:NativeUInt):NativeUInt;
var n:PBTTinyXMLNode;
begin
   try
      n := PBTTinyXMLNode(Node);
      Result := NativeUInt(n.ChildNodes);
   except
      Result := 0;
   end;
end;

//------------------------------------------------------------------------------
function    BTTinyXML.NextSibling(Node:NativeUInt):NativeUInt;
var n:PBTTinyXMLNode;
begin
   try
      n := PBTTinyXMLNode(Node);
      Result := NativeUInt(n.Next);
   except
      Result := 0;
   end;
end;

//------------------------------------------------------------------------------
procedure   _Addnnew(var n:PBTTinyXMLNode; const Name:string);
begin
   new(n);
   n.Name := Name;
   n.Attributes := nil;
   n.ChildNodes := nil;
   n.Data := '';
   n.Next := nil;
   n.Papa := nil;
   n.indx := 1;
   n.Rem := '';
   n.cdata := false;
end;

function    BTTinyXML.AddElement(Node:NativeUInt; const NodeName:string):NativeUInt;
var n,a:PBTTinyXMLNode;
begin
   try
      _AddnNew(n,NodeName);
      a := PBTTinyXMLNode(Node);
      if a = nil then a:= Nodes;
      if a <> nil then
      begin
         while a.Next <> nil do a := a.Next; //find end;
         a.Next := n;
      end else begin
         Nodes := n;
      end;
      Result := NativeUInt(n);
   except
      Result := 0;
   end;
end;

//------------------------------------------------------------------------------
function    BTTinyXML.AddChild(Node:NativeUInt; const NodeName:string):NativeUInt;
var n,a,b:PBTTinyXMLNode;
begin
   Result := 0;
   try
      _AddnNew(n,NodeName);
      a := PBTTinyXMLNode(Node);
      if a = nil then a:= Nodes;
      if a <> nil then
      begin
         b := a.ChildNodes;
         if b = nil then
         begin
            a.ChildNodes := n;
         end else begin
            while b.Next <> nil do b := b.Next; //find end;
            b.Next := n;
         end;
         Result := NativeUInt(n);
      end;
   except
      Result := 0;
   end;
end;

//------------------------------------------------------------------------------
function    BTTinyXML.SelectSingleNode(Node:NativeUInt; const NodeName:string; aFlags:longword = 0):NativeUInt; // 0 = root
var n,a:PBTTinyXMLNode;
    err,i:longint;
    attr_need:boolean;
    attr:string;
    s:string;
begin
   n := nil;
   if Node = 0 then
   begin
      if (Pos('/',NodeName) <> 0 ) or (Pos('\',NodeName)<> 0 ) then
      begin // possible Xpath entry
         err := 0;
         attr_need := false;
         attr := '';
         n := _FindNode(Nodes,NodeName,attr_need,attr,aFlags,err);
         if (err <> 0) or (attr_need) then n := nil; //error or ask for attr in Xpath
      end else begin
//         Node := NativeUInt(Nodes); //start from root
         a := Nodes; // start from root (next)
         while a <> nil do
         begin
            s := a.Name;
            if (aFlags and Preserve_namespace) = 0 then
            begin
               i := Pos(':',s);
               if i > 0 then s := Copy(s,i+1,length(s) - i);
            end;
            if s = NodeName then
            begin
               n := a;
               break;
            end;
            a := a.Next;
         end;
      end;
   end;
   if (n = nil) and (Node <> 0) then
   begin
      a := PBTTinyXMLNode(Node); // in papa childrens look for
      if a <> nil then
      begin
         a := a.ChildNodes;
         while (a<>nil) do
         begin
            s := a.Name;
            if (aFlags and Preserve_namespace) = 0 then
            begin
               i := Pos(':',s);
               if i > 0 then s := Copy(s,i+1,length(s) - i);
            end;
            if s = NodeName then
            begin
               n := a;
               break;
            end;
            a := a.next;
         end;
        end;
   end;
   Result := NativeUInt(n);
end;

//------------------------------------------------------------------------------
function    BTTinyXML.AddComment(Node:NativeUInt; const comment:string):boolean;
begin
   Result := false;
   try
      if length(comment)> 0 then
      begin
         PBTTinyXMLNode(Node).rem := comment;
         Result := true;
      end;
   except
      Result := false;
   end;

end;

end.

