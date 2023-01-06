unit BFileTools;

{$IFDEF FPC}
  {$MODE Delphi}
{$ENDIF}

interface


{$IFNDEF FPC }
{$IFDEF RELEASE}
{$WEAKLINKRTTI ON}
{$RTTI EXPLICIT METHODS([]) PROPERTIES([]) FIELDS([]) }
{$ENDIF}
{$ENDIF}



function FileExist (const File_Name :string) :boolean;
function FileInUse (const File_Name :string) :boolean;
function FileDelete (const FileName :string) :boolean;
function FileDeleteByPattern (const DirPath,FilePattern :string) :boolean;
function FileRename (const OldName, NewName :string) :boolean;
function FileSize(const FileName :string) :int64;
function FileCopy(const SourceFile, TargetFile :string) :boolean;
function DirectoryExist (const Directory :string) :boolean;
function GetCurrentDir :string;
function SetCurrentDir (const Dir :string) :boolean;
function CreateDir (const Dir :string) :boolean;
function RemoveDir (const Dir :string) :boolean;
function RelDirToAbsDir(const RelFileDir, BaseDir:string) :string;

function ChangeFileExt (const FileName, Extension :string) :string;
function ExtractFilePath (const FileName :string) :string;
function ExtractFileDrive (const FileName :string) :string;
function ExtractFile (const FileName :string) :string;
function ExtractFileName (const FileName :string) :string;
function ExtractFileExt (const FileName :string) :string;

function GetMyFileName:string;

function FileLoad (const FileName :string; var Data:AnsiString ) :boolean;
function FileLoadEx (const FileName :string; var Data:AnsiString ) :boolean; // no limit
function FileSave (const FileName :string; const data :AnsiString) :boolean;
function FileAdd (const FileName :string; const data :AnsiString) :boolean;

function FileReadBlock(const FileName :string; StartOffset :longint; aData :pointer; aDataLen :longword ) :boolean;
function FileWriteBlock(const FileName :string; StartOffset :longint; aData :pointer; aDataLen :longword ) :boolean;

function FileAddTag(const FileName :string; Tag:longword; aData :pointer; aDataLen :longword ) :boolean;
function FileReadAllTags(const FileName :string; StartOffset:longint; ReaderCallBack :pointer; UserData :longword ) :boolean;

procedure CorrectDirChar(var FileName:string);
procedure CorrectDirCharA(var FileName:ansistring);
function  GoodPath(const dir:string):string;  // put / in end

procedure GetDirList(const Dir,Mask,Skip:string; var DirList:string; Flags:longword=1);

function  GetSpecialDir(id :longword) :string;


implementation

uses windows,BStrTools;


//
{
  Flags $1 = get sub dirs
        $2 = get file name with full path
        $4 = set ';' for delimiter  default is #13#10
        $8 = Dir name in list
  Example
  GetDirList(aCurrentDir,'*.pas','',s,0);
   i := 0;
   if length(s)>0  then
   begin
      repeat
         s1:=ParseStr(s,i,';');
         inc(i);
         if length(s1) <> 0 then
         begin
            // do something
         end else i := 0;
      until i = 0
   end;
}
procedure  GetDirList(const Dir,Mask,Skip:string; var DirList:string; Flags:longword=1);
var RE:string;

   function   _DoTestBreak(SM:longword; const Mask,Data:string):boolean;
   var fn,fe,mfn,mfe:string;
   begin
      Result := false; // not in list
      fn := UpperCase(ExtractFileName(Data));
      fe := UpperCase(ExtractFileExt(Data));
      if Pos('.',Mask) <>0 then
      begin
         mfn := Trim(UpperCase(ParseStr(Mask,0,'.')));
         mfe := Trim(UpperCase(ParseStr(Mask,1,'.')));
         if mfn = '*' then
         begin
            if fe = mfe then Result := true;
         end else begin
            if fn = mfn then Result := true;
            mfn:= SkipChar(mfn,'*');
            if Pos(mfn,fn) <> 0 then Result := true;
         end;
      end else begin
         if fn = UpperCase(Mask) then Result := true;
         if Pos(UpperCase(mask),fn) <> 0 then Result := true;
      end;
      if SM = 1  then Result := not Result;
   end;

   function   _TestBreak(SM:longword; const Mask,Data:string):boolean;
   var i,j:longword;
       s:string;
   begin
      Result := false; // do not break
      if Pos('|',Mask) <> 0 then // more than one mask
      begin
         i := 0;
         repeat
            s := Trim(ParseStr(Mask,i,'|'));
            inc(i);
            j := length(s);
            if j > 0 then Result := Result or _DoTestBreak(SM,s,Data);
         until j = 0;
      end else begin
         s := Trim(Mask);
         Result := Result or _DoTestBreak(SM,s,Data);
      end;
   end;

   procedure  _GetDirList(const Dir,Mask,Skip:string; var DirList,ResEmpty:string; Flags:longword);
   var data :win32_Find_dataa;
       h:longword;
       a:ansistring;
       CDir:string;
       delim:string;
   begin
      delim := #13#10;
      if (Flags and 4) <> 0 then delim := ';';

      if ResEmpty = '' then ResEmpty := '\';
      a := ansistring(Dir)+'\*.*'+#0;
      h := FindFirstFileA(@a[1],data);
      if (h <> invalid_handle_value) then
      begin
         repeat
            a := ansistring(data.cFileName);
            if (a<>'.') then
            if (a<>'..') then
            begin
               Cdir := string(a);
               if Skip <> '' then if _TestBreak(0,Skip,CDir) then continue;
               //procceed
               if (data.dwFileAttributes and 16) <> 0 then // directory
               begin
                  if (Flags and 8) <> 0 then
                  begin
                     if (Flags and 2) <> 0 then DirList := DirList + Dir + ResEmpty + string(a) +'\'+Delim // put dir name in list
                                           else DirList := DirList + ResEmpty + string(a) +'\'+Delim;
                  end;
                  CDir := ResEmpty;
                  ResEmpty := ResEmpty + string(a)+'\';
                  if (Flags and 1) <> 0 then _GetDirList(Dir+'\'+ string(a),Mask,Skip,DirList,ResEmpty,Flags); // get sub dirs
                  ResEmpty := Cdir;
               end else begin
                  if Mask <> '' then if _TestBreak(1,Mask,CDir) then continue;
                  if (Flags and 2) <> 0 then DirList := DirList + Dir + ResEmpty + string(a)+delim
                                        else DirList := DirList + ResEmpty + string(a)+delim;
               end;
            end;
         until (FindNextFileA(h,data) = false );
         windows.FindClose(h);
      end;
   end;

begin
   RE := '';
   _GetDirLIst(Dir,Mask,Skip,DirList,Re,Flags);
end;

//------------------------------------------------------------------------------
procedure CorrectDirChar(var FileName:string);
var i,j:longword;
begin
   i := length(FileName);
   for j := 1 to i do if FileName[j] = '/' then FileName[j] := '\';
end;

//------------------------------------------------------------------------------
procedure CorrectDirCharA(var FileName:ansistring);
var i,j:longword;
begin
   i := length(FileName);
   for j := 1 to i do if FileName[j] = '/' then FileName[j] := '\';
end;

//------------------------------------------------------------------------------
function  GoodPath(const dir:string):string;
var i :longint;
begin
   Result := dir;
   CorrectDirChar(Result);
   i := length(Result);
   if i > 0 then if Result[i] <> '\' then Result := Result + '\'
            else Result := '\';
end;

//------------------------------------------------------------------------------
function    FileExist (const File_Name :string) :boolean;
var fn :string;
//    f:file of byte;
//    a :cardinal;
    Code :longint;
begin
   fn := File_Name + #0;
   Code := GetFileAttributes(@fn[1]);
   Result := (Code <> -1) and ((FILE_ATTRIBUTE_DIRECTORY and Code) = 0);
   if not Result then if GetLastError = ERROR_SHARING_VIOLATION then Result := true;

(*
   a := FileMode;
   Result := false;
   Assign(F,File_Name);
   {$I-}
   FileMode := 0; // 0 =  fmOpenRead;  1 write     2 read write
   reset(F);
   {$I+}
   if IOResult = 0 then
   begin
      Result := true;
      system.Close(F);
   end;
   FileMode := a;
*)
end;

//------------------------------------------------------------------------------
function    FileInUse (const File_Name :string) :boolean;
var f:file of byte;
begin
   Result := false;
   if FileExist(File_Name) then
   begin
      Assign(F,File_Name);
      {$I-}
       reset(F);
      {$I+}
      if IOResult = 0 then
      begin
         Result := true;
         system.Close(F);
      end else Result := true;
   end;
end;

//------------------------------------------------------------------------------
function FileDelete (const FileName :string) :boolean;
var fn:string;
begin
   fn := FileName+#0;
   Result := Windows.DeleteFile(PChar(@fn[1]));
end;

//------------------------------------------------------------------------------
function FileRename (const OldName, NewName :string) :boolean;
var so,sn:string;
begin
   so := OldName + #0;
   sn := NewName + #0;
   Result := MoveFile(PChar(@so[1]), PChar(@sn[1]));
end;

//------------------------------------------------------------------------------
function DirectoryExist (const Directory :string) :boolean;
var
   Code: Integer;
   d:string;
begin
   d := Directory+#0;
   Code := GetFileAttributes(PChar(@d[1]));
   Result := (Code <> -1) and ((FILE_ATTRIBUTE_DIRECTORY and Code) <> 0);
end;

//------------------------------------------------------------------------------
function GetCurrentDir :string;
begin
   GetDir(0, Result);
end;

//------------------------------------------------------------------------------
function SetCurrentDir (const Dir :string) :boolean;
var d:string;
begin
   d := Dir + #0;
   Result := SetCurrentDirectory(PChar(@d[1]));
end;

//------------------------------------------------------------------------------
function CreateDir (const Dir :string) :boolean;
var d:string;
begin
   d := Dir +#0;
   Result := CreateDirectory(PChar(@d[1]), nil);
end;

//------------------------------------------------------------------------------
function RemoveDir(const Dir: string): boolean;
var d:string;
begin
   d := Dir +#0;
   Result := RemoveDirectory(PChar(@d[1]));
end;

//------------------------------------------------------------------------------
function _LastDelimiter(const Delimiters, S: string): Integer;
begin
   Result := Length(S);
   while Result > 0 do
   begin
     if Pos(S[Result],Delimiters) <> 0 then Exit;
     Dec(Result);
   end;
end;

const
   PathDelim  = '\';
   DriveDelim = ':';
   PathSep    = ';';

function ChangeFileExt(const FileName, Extension :string) :string;
var
   I: Integer;
begin
   I := _LastDelimiter('.',Filename);
   if (I = 0) or (FileName[I] <> '.') then I := MaxInt;
   Result := Copy(FileName, 1, I ) + Extension;
end;

//------------------------------------------------------------------------------
function ExtractFilePath(const FileName :string) :string;
var
   I: Integer;
begin
   I := _LastDelimiter('\', FileName);
   Result := Copy(FileName, 1, I);
end;

//------------------------------------------------------------------------------
function ExtractFileDrive(const FileName :string) :string;
begin
   Result := '';
   if (FileName[2] = ':') and (FileName[3] = '\') then Result := FileName[1]+':';
end;

//------------------------------------------------------------------------------
function ExtractFile(const FileName :string) :string;
var
   I: Integer;
begin
   I := _LastDelimiter('\', FileName);
   if (I > 0) and (FileName[I] = '\') then
      Result := Copy(FileName, I + 1, MaxInt) else
      Result := FileName;
end;

//------------------------------------------------------------------------------
function ExtractFileName(const FileName :string) :string;
var
   I: Integer;
begin
   Result := ExtractFile(FileName);
   I := _LastDelimiter('.', Result);
   if (I > 0) and (Result[I] = '.') then   Result := Copy(Result, 1, I - 1);
end;

//------------------------------------------------------------------------------
function ExtractFileExt(const FileName :string) :string;
var
   I: Integer;
begin
   I := _LastDelimiter('.', FileName);// + PathDelim + DriveDelim, FileName);
   if (I > 0) and (FileName[I] = '.') then
      Result := Copy(FileName, I + 1, MaxInt) else
      Result := '';
end;

//------------------------------------------------------------------------------
function FileSize(const FileName :string) :int64;
var
   info :TWin32FileAttributeData;
   fn : string;
begin
   fn := FileName + #0;
   Result := -1;
   if not GetFileAttributesEx(PChar(@fn[1]), GetFileExInfoStandard, @info) then Exit;
   Result := info.nFileSizeLow or ( info.nFileSizeHigh shl 32);
end;


//------------------------------------------------------------------------------
// get running application path and file name
function GetMyFileName:string;
var Buf:array[0..254]of char;
begin
   FillChar(Buf,Sizeof(Buf),#0);
   GetModuleFileName(hInstance,Buf,255);
   Result := string(Buf);
end;

//------------------------------------------------------------------------------
function FileReadBlock(const FileName :string; StartOffset :longint; aData :pointer; aDataLen :longword ) :boolean;
var  i :longword;
//    f  :file of byte;
//    a :cardinal;
    H :longint;
    fn :string;
begin
   fn := FileName + #0;    //thread save
   Result := false;
   H := CreateFile(@FN[1], GENERIC_READ, FILE_SHARE_READ, nil, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);
   if H > 0 then
   begin
      i := 0;
      SetFilePointer(H,StartOffset,@i,FILE_BEGIN); //FILE_CURRENT FILE_END    todo INVALIS_SET_FILE_POINTER
      if ReadFile(H,aData^,aDataLen,i,nil) then if i = aDataLen then Result := true;
      CloseHandle(H);
   end;
(*
   Result := false;
   a := FileMode;
   if StartOffset < 0 then Exit;

   system.Assign(f,Filename);
   {$I-}
   FileMode := 0; // 0 =  fmOpenRead;  1 write     2 read write  !!!! this is not thread save
   System.reset(f);
   if IOResult = 0 then
   begin
      system.Seek(f,StartOffset);
      blockread(f,aData^,aDataLen,i);
      if i <> aDataLen then adataLen := 0;
      if IOResult <> 0 then aDataLen := 0;    //todo for empty file!!!!!:(
      if aDataLen <> 0 then Result := true; //OK
   end;
   system.Close(f);
   {$I+}
   FileMode := a;
*)
end;

//------------------------------------------------------------------------------
function FileWriteBlock(const FileName :string; StartOffset :longint; aData :pointer; aDataLen :longword) :boolean;
var j  :longword;
    f  :file of byte;
    fe :integer;
begin
   j := FileSize(FileName);
   if StartOffset = -1 then StartOffset := j; // add to end

   Result := false; // fail
   system.Assign(f,Filename);
   {$I-}
   system.reset(f);
   fe := IOResult; // after call it clears
   if fe <> 0 then
   begin
      system.rewrite(f); // create if not exist
      fe := IOResult;
      StartOffset := 0;
   end;

   if fe = 0 then
   begin
      system.Seek(f,StartOffset);
//      if IOResult = 0 then
//      begin
         blockwrite(f,aData^,aDataLen);
         if IOResult = 0 then Result := True; //OK
//      end;
         if aDataLen <> 0 then Result := true; //OK
   end;
   {$I+}
   system.Close(f);

end;

//------------------------------------------------------------------------------
function FileAdd(const FileName :string; const data :AnsiString) :boolean;
begin
   Result := FileWriteBlock(FileName,-1,@data[1],length(data));
end;

//------------------------------------------------------------------------------
function FileSave(const FileName :string; const data :AnsiString) :boolean;
begin
   Result := false; // Fail
   if FileExist(FileName) then if not FileDelete(FileName) then Exit;
   Result := FileWriteBlock(FileName,0,@data[1],length(data));
end;

//------------------------------------------------------------------------------
function FileLoad(const FileName :string; var Data:AnsiString ) :boolean;
var i  :longword;
begin
   Result := false; // fault
   Data := '';
   i := FileSize(FileName);
   if (i > 0) and (i < 128000) then // put some limit
   begin
      SetLength(Data,i);
      Result := FileReadBlock(FileName,0,@Data[1],i);
   end else begin
      if i = 0 then Result := true;
   end;
end;

//------------------------------------------------------------------------------
function FileLoadEx(const FileName :string; var Data:AnsiString ) :boolean;
var i  :longword;
begin
   Result := false; // fault
   Data := '';
   i := FileSize(FileName);
   if (i > 0) then
   begin
      SetLength(Data,i);
      Result := FileReadBlock(FileName,0,@Data[1],i);
   end else begin
      if i = 0 then Result := true;
   end;
end;

//------------------------------------------------------------------------------
function FileCopy(const SourceFile, TargetFile :string) :boolean;
var S,D:ansistring;
begin
   Result := false;
   S := ansistring(SourceFile)+#0;
   D := ansistring(TargetFile)+#0;
   if copyfileA(@S[1],@D[1],false) then Result := true; //ok ansi
end;

//------------------------------------------------------------------------------
function FileAddTag(const FileName :string; Tag:longword; aData :pointer; aDataLen :longword ) :boolean;
var size : longword;
begin
   Result := false;
   size := aDataLen + 8; { SIZE + TAG }
   if FileWriteBlock(FileName,-1,@size,4) then
    if FileWriteBlock(FileName,-1,@Tag,4) then
     if FileWriteBlock(FileName,-1,aData,aDataLen) then Result := true;
end;

//------------------------------------------------------------------------------
type // fpc need type
   T_Tcb = function(userd,tag:longword; data:pointer; datalen:longword):boolean; stdcall;

function FileReadAllTags(const FileName :string; StartOffset:longint; ReaderCallBack :pointer; UserData :longword ) :boolean;
var f :file of byte;
    cb : T_Tcb;
    p:pointer;
    sz,tg,dn:longword;
    lim,fs :longword;
begin
   dn := 0;
   lim := 0;
   cb := T_Tcb(readerCallBack);
   fs := FileSize(FileName);
   p := nil;
   Result := false; // error
   if StartOffset < 0  then StartOffset := 0;
   fs := fs - longword(StartOffset); //was + 1;

   system.Assign(f,Filename);
   {$I-}
   reset(f);
   if IOResult = 0 then
   begin
      system.Seek(f,StartOffset);
      repeat
         blockread(f,sz,4);
         if IOResult = 0 then
         begin
            dec(fs,4);
            blockread(f,tg,4);
            if IOResult = 0 then
            begin
               dec(fs,4);
               if (sz > 8) and ( sz <= (fs+8)) then
               begin
                  sz := sz - 8; { minus size and tag }
                  if sz > lim then
                  begin
                     lim := sz;
                     ReallocMem(p,lim);
                     if p = nil then break//error
                  end;
                  blockread(f,p^,sz);
                  if IOResult = 0 then
                  begin
                     dec(fs,sz);
                     if not cb(UserData,Tg,p,sz) then dn := 1;
                  end;
               end else break
            end;
         end;
      until (dn = 1) or (IOResult <> 0) or (fs = 0);
   end;
   {$I+}
   if (dn = 0) and (IOResult = 0) and (fs = 0) then Result := true;
   if p <> nil then ReallocMem(p,0); //free
end;

//------------------------------------------------------------------------------
function  GetSpecialDir(id :longword) :string;
var hk:HKEY;
    tp,ss,dl:longword;
    st:longint;
    kn:string;
    ada,k:ansistring;
begin
   Result := '';

    SetLength(kn,128);
    if id = 0 then GetWindowsDirectory(@kn[1],128);
    if id = 1 then GetSystemDirectory(@kn[1],128);
    if id = 2 then GetTempPath(128,@kn[1]);

    if id < 3 then
    begin
       for tp := 1 to 128 do if kn[tp] <> #0 then Result := Result + kn[tp] else break;
       Exit;
    end;


   RegOpenKeyEx($80000001 {HKCU},'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders',0,KEY_READ,hk);
   if hk = 0 then Exit;
   // key exist so read it
   dl := 0;
   case (id - 3) of
     0: {3} kn := 'Desktop';  //-> Desktop
     1: {4} kn := 'AppData';  //-> AppData\Raoming
     2: {5} kn := 'Font';
     3: {6} kn := 'Personal'; //-> Documents
     4: {7} kn := 'Startup';
     5: {8} kn := 'Start Menu'; //-> AppData\Roaming\Microsoft\Windows\Start Menu
     6: {9} kn := 'Programs';  //-> AppData\Roaming\Microsoft\Windows\Start Menu\Programs
     else begin kn := 'Desktop'; dl := 1; end;
   end;

   tp := REG_SZ;
   ss := 0;
   st := RegQueryValueEx(hk,@kn[1],nil,@tp,nil,@ss);
   if st = 0 then
   begin
      SetLength(ada,ss);
      k :=ansistring(kn)+#0;
      st := RegQueryValueExA(hk,@K[1],nil,@tp,@ada[1],@ss);
      if st = 0 then
      begin
         if ada[ss]=#0 then SetLength(ada,ss-1);
         Result := string(ada);
         if dl = 1 then // replace
         begin
            Result := ReplaceString(Result,'Desktop','Downloads',0);
         end;
      end;
   end;
   RegCloseKey(hk);
end;

//------------------------------------------------------------------------------
function FileDeleteByPattern (const DirPath,FilePattern :string) :boolean;
var  data :win32_Find_data;
     h:longword;
     a:ansistring;
     path:string;
begin
   path:= GoodPath(DirPath);
   Result := false;
   a := ansistring(path + FilePattern)+#0;
   h := FindFirstFile(@a[1],data);
   if (h <> invalid_handle_value) then
   begin
      repeat
         a := ansistring(data.cFileName);
         FileDelete(string(a));
      until (FindNextFile(h,data) = false );
      windows.FindClose(h);
      Result := true;
   end;
end;

//------------------------------------------------------------------------------
function PathIsRelative(pszPath: LPCSTR): BOOL; stdcall; external 'shlwapi.dll' name 'PathIsRelativW';
function PathCanonicalize(pszBuf: LPWSTR; pszPath: LPCWSTR): BOOL; stdcall; external 'shlwapi.dll' name 'PathCanonicalizeW';

function RelDirToAbsDir(const RelFileDir, BaseDir:string) :string;
var Buffer:array[0..MAX_PATH] of widechar;
    s,r:widestring;
begin
   Result := '';
   s :=widestring(RelFileDir);
   r := widestring(GoodPath(BaseDir));
   if PathIsRelative(@(s[1])) then
   begin
      r := r + s;
   end else begin
      if Length(ExtractFilePath(RelFileDir)) = 0 then r := r + s  //RelFileDir is only filename
                                                 else r := s;     //RelFileDir is file plus path ignore BaseDir
   end;
   if PathCanonicalize(@Buffer[1],@r) then Result := string(Buffer);
end;

end.
