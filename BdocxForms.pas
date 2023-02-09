unit BdocxForms;

interface

//UNDER DEVELOPMENT


type  BTDocxForms = class
         private
            aDocument:AnsiString;
            aFormTemplateFileName:string;
            aToDo :pointer;
            function    _AddTag(typ:longword; const Tag1,Tag2:string):boolean;
            function    _ReBuild:boolean;
         public
            constructor Create;
            destructor  Destroy; override;
            function    OpenForm(const FileName:string):boolean;
            function    WriteForm(const FileName:string):boolean;

            function    ReplaceTag(const TagName,Data:string):boolean;
            function    DeleteParagraph(const TagName:string):boolean;
            function    AcceptParagraph(const TagName:string):boolean;
            function    ReplaceTableTag(const TagName,Data:string):boolean;
            function    DeleteTableTag(const TagName:string):boolean;
            function    DeleteTableRow(TableID,TableRow:longword):boolean;

//TODO
//ReplacePic
//data in table
//Table delete or add ???? un used
         end;


implementation

uses BpasZlib,BStrTools,BTinyXML,BUnicode,Windows;

type TTagToDo = record
        Next :pointer;
        Tag1 :string;
        Tag2 :string;
        Op   :longword;
        res  :longword;
        match_pos :longword;
        match_len :longword;
        match_tag :NativeUInt;
     end;
     PTTagToDo = ^TTagToDo;


constructor BTDocxForms.Create;
begin
   aDocument := '';
   aToDo := nil;
end;

destructor  BTDocxForms.Destroy;
var ss,s:PTTagToDO;
begin
   if aToDo <> nil then
   begin
      ss := aToDo;
      while(ss<>nil) do
      begin
         s := ss;
         ss := s.next;
         Dispose(s);
      end;
   end;
   inherited;
end;


function    BTDocxForms.OpenForm(const FileName:string):boolean;
var z:TZipRead;
    i:integer;
    buf:ansistring;
begin
   Result := false;
   aFormTemplateFileName := FileName;
   z := TZipRead.Create(FileName);
   if z.Count > 0 then
   begin
      i := z.NameToIndex('word/document.xml');
      if i>= 0 then
      begin
         buf := z.UnZip(i);
         if length(buf) > 0 then
         begin
            aDocument := buf; //UTF82Unicode(buf);
            Result := true;
         end;
      end;
   end;
   z.Destroy;
end;

function    BTDocxForms.WriteForm(const FileName:string):boolean;
var z:TZipRead;
    zw:TZipWrite;
    i,j:integer;
    r,buf:ansistring;
begin
   Result := false;
   if _ReBuild then
   begin
      z := TZipRead.Create(aFormTemplateFileName);
      if z.Count > 0 then
      begin
         zw := TZipWrite.Create(FileName);
         for i :=0 to z.Count - 1 do
         begin
            r := ansistring(z.zEntry[i].Name);
            if r = 'word/document.xml' then
            begin
               buf := aDocument; //Unicode2UTF8(aDocument);
               zw.AddDeflated(r,@buf[1],length(buf));
            end else begin
               //just copy
               j := z.NameToIndex(r);
               if j >= 0 then
               begin
                  buf := z.UnZip(j);
                  zw.AddDeflated(r,@buf[1],length(buf));
               end;
            end;
         end;
         Result := true;
         zw.Destroy;
      end;
      z.Destroy;
   end;
   aDocument := '';
end;


var a_a,a_c,a_i:longword;

function    BTDocxForms._ReBuild:boolean;
var i,j,k,ts,ts2,wt,we:longint;
    ttag,obj,el,robj,rr,cc,stag,dtag:NativeUInt;
    s,st,st2:string;
    o1,o2:PTTagToDo;
    Z:BTTinyXML;
    skip_cnt:longint;
    op:longword;
    done,skip,found,delpar:boolean;

   procedure _ClearMatchPos;
   begin
      if aToDO <> nil then
      begin
         o1 := aToDo;
         while (o1 <> nil) do
         begin
            o1.match_tag := 0;
            o1.match_pos := 1;
            o1 := o1.next;
         end;
      end;
   end;


   procedure _SearcTagAndDo;
   var  _i,_j:longint;
         _c :char;
   begin


      if Done then
      begin
         // do the operation
         Z.SetText(ttag,o2.tag2);

         Skip_cnt := 0;  //?????
         done := false;
      end else begin



      s := Z.GetText(ttag);
      if Z.GetAttribute(ttag,'space')='preserve' then
      begin

      end;
      // input stream letter by letter
      _j := length(s);
      for _i := 1 to _j do
      begin
         _c := s[_i];
         if _c = #32 then
         begin
            _ClearMatchPos;
         end else begin
            if aToDO <> nil then
            begin
               o1 := aToDo;
               while (o1 <> nil) do
               begin
                  if _c = o1.tag1[o1.match_pos] then
                  begin
                     if o1.match_pos = 1 then o1.match_tag := robj; //r tag
                     if o1.match_pos = o1.match_len then
                     begin // I have match
                        o2 := o1;
                        found := true;
                        break;
                     end;
                     inc(o1.match_pos);
                  end else begin
                     o1.match_pos := 1;
                  end;
                  o1 := o1.Next;
               end;
            end;
         end;
      end;


         if found then
         begin
            found := false;
            if o2.op = 3 then //Del parag
            begin
               dtag := el; // this parag tag
               delpar := true;
            end;
            if (o2.op = 1) or (o2.op = 2) then // Replase Tag or Accept Parag (just replace tag)
            begin
               robj := o2.match_tag; // go back to begining of tag   stag - tag of fist found
               done := true;
               Skip_cnt := 0;
            end;
         end else begin
            if Skip_cnt > 0  then
            begin
               _i := length(s);
               if _i < Skip_cnt then
               begin
                  s := '';
                  dec(Skip_cnt,_i);
               end else begin
                  s := Copy(s,Skip_cnt+1,length(s) - Skip_cnt);
                  Skip_cnt := 0;
               end;
               Z.SetText(ttag,s);
            end;
         end;

      end;
   end;


begin
   Result := false;
   s := '';
   done := false;
   found := false;
   _ClearMatchPos;

   if length(aDocument) > 0  then
   begin
      Z:=BTTinyXML.Create;
      Z.LoadXML(string(aDocument),1);
      obj := Z.SelectSingleNode(0,'document');
      obj := Z.SelectSingleNode(obj,'body');
      if obj <> 0 then
      begin
         el := Z.FirstChild(Obj);
         while el <> 0  do
         begin

            delpar := false;
            if Z.NodeName(el) = 'p' then
            begin

               Skip_cnt := 0;
               stag := 0; //start tag
               st := ''; // acumml text

               robj := Z.FirstChild(el);
               while robj <> 0 do
               begin
                  found := false;
                  if Z.NodeName(robj) = 'r' then
                  begin
                     ttag := Z.SelectSingleNode(robj,'t');
                     if ttag <> 0 then _SearcTagAndDo
                  end;
                  robj := Z.NextSibling(robj);
               end;
            end;

            if Z.NodeName(el) = 'tbl' then
            begin
               obj := Z.FirstChild(el);
               while obj <> 0 do
               begin
                  if Z.NodeName(obj) = 'tr' then
                  begin
                     cc := Z.FirstChild(obj);
                     while cc <> 0  do
                     begin
                        if Z.NodeName(cc) = 'tc' then
                        begin
                           rr := Z.SelectSingleNode(cc,'p');
                           robj := Z.FirstChild(rr);
                           while robj <> 0 do
                           begin
                              if Z.NodeName(robj) = 'r' then
                              begin
                                 ttag := Z.SelectSingleNode(obj,'t');
                                 if ttag <> 0 then _SearcTagAndDo
                              end;
                              robj := Z.NextSibling(robj);
                           end;
                        end;
                        cc := Z.NextSibling(cc); //column
                     end;
                  end;
                  obj := Z.NextSibling(obj); //row
               end;
            end;
            el := Z.NextSibling(el);
            if delpar then Z.DeleteNode(dtag);
         end;
      end;




     (*

      i := 0;
      repeat // paragraph
         inc(i);
         Done := true;
         j := 0;
         skip := false;
         repeat // word
            inc(j);
            st := '/document/body/p['+Tostr(i)+']/r['+Tostr(j)+']/t';
            ts := Z.SelectXPath(st,s);
            if ts = 0 then
            begin
               if j = 1 then Done := false;
               if aToDO <> nil then
               begin
                  o1 := aToDo;
                  while (o1 <> nil) do
                  begin
                     if (o1.Op = 1) or (o1.Op = 2) then //Replace or Delete
                     begin
                        if Pos(o1.tag1,s) <> 0 then
                        begin
                           s := ReplaceString(s,o1.tag1,o1.tag2,RS_CaseSense);
                           ts2 := Z.UpdateXPath(st,s);
                           break;
                        end;
                     end;
                     if (o1.Op = 3) or (o1.Op = 4) then
                     begin
                        if Pos(o1.tag1,s) <> 0 then // Begin tag
                        begin
                           if o1.Op= 3 then skip := true;
                           ts2 := Z.UpdateXPath(st,'');
                           break
                           //todo
                        end;
                        if Pos(o1.tag2,s) <> 0 then // End tag
                        begin
                           if o1.Op= 3 then skip := false;
                           ts2 := Z.UpdateXPath(st,'');
                           break
                           //todo
                        end;
                     end;
                     o1 := o1.next;
                  end;
               end;
            end;
            if skip then ts := Z.UpdateXPath(st,'');
         until ts = 100;
      until Done;

      // table
      i := 0;
      repeat // table
         inc(i);
         Done := true;
         j := 0;
         repeat // row
            inc(j);
            st := '/document/body/tbl['+Tostr(i)+']/tr['+Tostr(j)+']';
            ts := Z.SelectXPath(st+'/tc[1]/p/r/t',s);
            if ts = 0 then
            begin
               if j = 1 then Done := false;
               k := 0;
               repeat // column
                  inc(k);
                  st2 := '/tc['+ToStr(k)+']/p/r/t';
                  ts2 := Z.SelectXPath(st+st2,s);
                  if ts2 = 0 then
                  begin
                     if aToDO <> nil then
                     begin
                        o1 := aToDo;
                        while (o1 <> nil) do
                        begin
                           if (o1.Op = 5) or (o1.Op = 6) then //Replace or Delete
                           begin
                              if Pos(o1.tag1,s) <> 0 then
                              begin
                                 s := ReplaceString(s,o1.tag1,o1.tag2,RS_CaseSense);
                                 Z.UpdateXPath(st+st2,s);
                                 break;
                              end;
                           end;
                           o1 := o1.next;
                        end;
                     end;
                  end;
               until ts2 = 100;
            end;
         until ts = 100;
      until Done;

      // Adjust and clear unused tags
      i := 0;
      repeat // paragraph
         inc(i);
         Done := true;
         j := 0;
         skip := false;
         wt := 0;
         we := 0;
         repeat // word
            inc(j);
            st := '/document/body/p['+Tostr(i)+']/r['+Tostr(j)+']';
            ts := Z.SelectXPath(st+'/t',s);
            if ts = 0 then
            begin
               if j = 1 then Done := false;
               inc(wt); // have word tag;
               if length(trim(s)) = 0 then
               begin
                  inc(we); //word empty
                  ttag := Z.SelectSingleNode(0,st);
                  if ttag <> 0 then  Z.DeleteNode(ttag);
               end;
            end;
         until ts = 100;
         if we = wt then // pargraph is empty
         begin
            ttag := Z.SelectSingleNode(0,'/document/body/p['+Tostr(i)+']');
            if ttag <> 0 then  Z.DeleteNode(ttag);
            dec(i);
         end;
      until Done;

      i := 0;
      repeat // table
         inc(i);
         Done := true;
         j := 0;
         repeat // row
            inc(j);
            st := '/document/body/tbl['+Tostr(i)+']/tr['+Tostr(j)+']';
            ts := Z.SelectXPath(st+'/tc[1]/p/r/t',s);
            if ts = 0 then
            begin
               if j = 1 then Done := false;
               k := 0;
               wt := 0;
               we := 0;
               repeat // column
                  inc(k);
                  st2 := '/tc['+ToStr(k)+']/p/r/t';
                  ts2 := Z.SelectXPath(st+st2,s);
                  inc(wt);
                  if length(trim(s)) = 0 then inc(we);
               until ts2 = 100;
               if we = wt then // pargraph is empty
               begin
                  ttag := Z.SelectSingleNode(0,st);
                  if ttag <> 0 then  Z.DeleteNode(ttag);
                  dec(j);
               end;
            end;
          until ts = 100;
      until Done;
      *)

      //Second pass rmove empty paragraphs

      Z.GetXML(s);
      aDocument := ansistring(s);
      Z.Free;
      Result := true;
   end;
end;


function    BTDocxForms._AddTag(typ:longword; const Tag1,Tag2:string):boolean;
var c,s:PTTagToDo;
begin
   Result := false;
   try
      new(s);
      s.next := nil;
      s.Tag1 := string(Unicode2UTF8(widestring(tag1)));
      s.Tag2 := string(Unicode2UTF8(widestring(tag2)));
      s.Op := typ and $FF;
      s.res := typ shr 16;
      s.match_pos := 1;
      s.match_len := length(s.Tag1);
      s.match_tag := 0;

      if aToDo = nil then
      begin
         aToDo := s;
      end else begin
         c := aToDo;
         while c.next <> nil do c := c.next;
         c.next := s;
      end;
      Result := true;
   except
      Result := false;
   end;
end;


function    BTDocxForms.ReplaceTag(const TagName,Data:string):boolean;
begin
   Result := _AddTag(1,TagName,Data);
end;

function    BTDocxForms.DeleteParagraph(const TagName:string):boolean;
begin
   Result := _AddTag(3,TagName,'');
end;

function    BTDocxForms.AcceptParagraph(const TagName:string):boolean;
begin
   Result := _AddTag(2,TagName,'');
end;

function    BTDocxForms.ReplaceTableTag(const TagName,Data:string):boolean;
begin
   Result := _AddTag(5,TagName,Data);
end;

function    BTDocxForms.DeleteTableTag(const TagName:string):boolean;
begin
   Result := _AddTag(6,TagName,'');
end;

function    BTDocxForms.DeleteTableRow(TableID,TableRow:longword):boolean;
begin
   Result := _AddTag( ((TableId and $FF) shl 24) or ((TableRow and $FF) shl 16) or 7, '','');
end;


end.
