unit BxlsxWriter;

interface

// version 0.5   Bogi aka sdex32 5.2022  :)
// TODO test for big data time consume
//work well

const BTXW_Fatr_normal    =    0; //mask
      BTXW_Fatr_bold      =  $01;
      BTXW_Fatr_italic    =  $02;
      BTXW_Fatr_underline =  $04;

      // border                      1 - bottom  (result is bit mask)
      //         _______ (2)         2 - top
      //       | (4)        |        3 - right
      //       | _______ (1)| (3)    4 - left
      //

      // charset
      // 0 - ansi
      // 1 - System default
      // 2 - Symbol char set
      // 204 - cyrilic

      // align
      // 0 - default
      // 1 - left
      // 2 - center
      // 3 - right



type  BTXlsxWriter = class
         private
            aData :pointer;
            aStyle :pointer;
            aFonts :pointer;
            aStyleColors :pointer;
            aSheetsCount :longword;
            aStringPool :ansistring;
            aStringCnt :longword;
            aStringUnq :longword;
            aBorderUse :boolean;
         public
            constructor Create;
            destructor  Destroy; override;
            procedure   Reset;
            function    Generate(const FileName:string):boolean;
            function    GenerateCSV(const FileName:string; sheet:nativeUint; Separator:ansistring=';'; PutOnLast:boolean=false):boolean;
            function    SetUpFont(const name:string; size,atr,color:longword; charset:longword=1):longword;  //atr bits  1=bold 2-italic 4-underline

            function    AddSheet(const name:widestring):nativeUint;
            function    AddCell(sheet:nativeUint; Col,Row:longword; const value:widestring; val_typ:longword = 0):nativeUint; overload;
            function    AddCell(sheet:nativeUint; const Adr,value:widestring; val_typ:longword = 0):nativeUint; overload;
            function    AddCellStyle(cell:nativeUint; color,border,font,align:longword):boolean;

//            function    SetRowSize(sheet:nativeUint; Row,Size:longword):boolean;
//            function    SetColSize(sheet:nativeUint; Col,Size:longword):boolean;
// marge cells
// todo formules
// todo optimeze the size of messages
      end;





implementation


uses BpasZlib,BStrTools,BUnicode,BDate,BFileTools;

type BTXLSW_Sheet = record
        next:pointer;
        cells:pointer;
        total:longword;
        maxr,maxc:longword;
        name:ansistring;
     end;
     PBTXLSW_Sheet = ^BTXLSW_Sheet;

     BTXLSW_Cell = record
        next:pointer;
        row , col :longword;
        value: ansistring;
        typ :longword;
        style :longword;
        touch: boolean; //.todo i dont need this

     end;
     PBTXLSW_Cell = ^BTXLSW_Cell;

     BTXLSW_Style = record
        next:pointer;
        color:longword;
        colorId:longword;
        border:longword;
        font:longword;
        align:longword;
     end;
     PBTXLSW_Style = ^BTXLSW_Style;

     BTXLSW_StyleColor = record
        next:pointer;
        color:longword;
     end;
     PBTXLSW_StyleColor = ^BTXLSW_Style;

     BTXLSW_Font = record
        next:pointer;
        name:ansistring;
        size:longword;
        atr:longword;
        color:longword;
        charset:longword;
     end;
     PBTXLSW_Font = ^BTXLSW_Font;

//------------------------------------------------------------------------------
const

//  _rels/.rels
rels : ansistring = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+#13#10
           + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           + '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
           + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
           + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
           + '</Relationships>';


//------------------------------------------------------------------------------
function GetContentTypes(xls:BTXlsxWriter):ansistring; //  [Control_Types].xml
var i :longword;
begin
   Result := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+#13#10+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
	         + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
	         + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
   for i := 1 to xls.aSheetsCount do
   begin
      Result := Result
           + '<Override PartName="/xl/worksheets/sheet'+ansistring(toStr(i))+'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
   end;
   Result := Result
           + '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
           + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
           + '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
           + '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
           + '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>';
end;


//------------------------------------------------------------------------------
function GetPropsApp(xls:BTXlsxWriter):ansistring; //    docProps/app.xml
var i:longword;
begin
   Result := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+#13#10
	         + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
	         + '<Application>Microsoft Excel</Application>'
	         + '<DocSecurity>0</DocSecurity>'
	         + '<ScaleCrop>false</ScaleCrop>'
	         + '<HeadingPairs>'
	         + '<vt:vector size="2" baseType="variant">'
	         + '<vt:variant>'
	         + '<vt:lpstr>Worksheets</vt:lpstr>'
	         + '</vt:variant>'
	         + '<vt:variant>'

	         + '<vt:i4>'+ansistring(ToStr(xls.aSheetsCount))+'</vt:i4>' // set 2 if two sheets

	         + '</vt:variant>'
	         + '</vt:vector>'
	         + '</HeadingPairs>'
	         + '<TitlesOfParts>';

      Result := Result
	         + '<vt:vector size="'+ansistring(toStr(xls.aSheetsCount))+'" baseType="lpstr">'; // set 2 if two sheets
   for i := 1 to xls.aSheetsCount do
   begin
      Result := Result
	         + '<vt:lpstr>Sheet'+ansistring(toStr(i))+'</vt:lpstr>'; // add new row
   end;
   Result := Result
	         + '</vt:vector>'
	         + '</TitlesOfParts>'
	         + '<Company></Company>'
	         + '<LinksUpToDate>false</LinksUpToDate>'
	         + '<SharedDoc>false</SharedDoc>'
	         + '<HyperlinksChanged>false</HyperlinksChanged>'
	         + '<AppVersion>16.0300</AppVersion>'
	         + '</Properties>';
end;

//------------------------------------------------------------------------------
function GetPropsCore:ansistring; // docProps/core.xml
var s:ansistring;
begin

   s := ansistring(DateConToUTCstr(GetTodaySys,false,true)); // need system time greanwich

   Result := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+#13#10
	         + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" '
	         + 'xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
	         + '<dc:creator>Tiny Xls Writer</dc:creator>'
	         + '<cp:lastModifiedBy>Tiny Xls Writer</cp:lastModifiedBy>'

	         + '<dcterms:created xsi:type="dcterms:W3CDTF">'+s+'</dcterms:created>'
	         + '<dcterms:modified xsi:type="dcterms:W3CDTF">'+s+'</dcterms:modified>'

	         + '</cp:coreProperties>';
end;

//------------------------------------------------------------------------------
function GetXlRelsWorkbook(xls:BTXlsxWriter):ansistring; // xl/_rels/workbook.xml.rels
var i:integer;
Begin
   Result := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+#13#10
	         + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
	         + '<Relationship Id="rId'+ansistring(toStr(xls.aSheetsCount + 2))+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
	         + '<Relationship Id="rId'+ansistring(toStr(xls.aSheetsCount + 1))+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>';
   for i := 1 to xls.aSheetsCount do
   begin
      Result := Result
	         + '<Relationship Id="rId'+ansistring(toStr(i))+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet'+ansistring(toStr(i))+'.xml"/>';
   end;

      Result := Result
	         + '<Relationship Id="rId'+ansistring(toStr(xls.aSheetsCount + 3))+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
	         + '</Relationships>';
end;

//------------------------------------------------------------------------------
function GetXlTheme:ansistring; // xl/theme/theme1.xml
begin
   Result := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+#13#10
	         + '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">'
	         + '<a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
	         + '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
	         + '<a:dk2><a:srgbClr val="44546A"/></a:dk2>'
	         + '<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
	         + '<a:accent1><a:srgbClr val="5B9BD5"/></a:accent1>'
	         + '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'
	         + '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>'
	         + '<a:accent4><a:srgbClr val="FFC000"/></a:accent4>'
	         + '<a:accent5><a:srgbClr val="4472C4"/></a:accent5>'
	         + '<a:accent6><a:srgbClr val="70AD47"/></a:accent6>'
	         + '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>'
	         + '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>'
	         + '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/>'
	         + '<a:ea typeface=""/><a:cs typeface=""/>'
	         + '<a:font script="Hebr" typeface="Times New Roman"/></a:majorFont>'
	         + '<a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/>'
	         + '<a:ea typeface=""/><a:cs typeface=""/>'
	         + '<a:font script="Hebr" typeface="Arial"/></a:minorFont></a:fontScheme>'
	         + '<a:fmtScheme name="Office"><a:fillStyleLst>'
	         + '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
	         + '<a:gradFill rotWithShape="1"><a:gsLst>'
	         + '<a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs>'
	         + '<a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs>'
	         + '<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs>'
	         + '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst>'
	         + '<a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs>'
	         + '<a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs>'
	         + '<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs>'
	         + '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst>'
	         + '<a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
	         + '<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
	         + '<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
	         + '<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
	         + '</a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>'
	         + '<a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">'
	         + '<a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>'
	         + '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
	         + '<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>'
	         + '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/>'
	         + '<a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs>'
	         + '<a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/>'
	         + '<a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/>'
	         + '<a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
	         + '</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst>'
	         + '<a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/>'
	         + '</a:ext></a:extLst></a:theme>';
end;

//------------------------------------------------------------------------------
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

//------------------------------------------------------------------------------
function GetXlSheet(xls:BTXlsxWriter; sheet:longword):ansistring; // xl/worksheets/sheet1.xml
var ss:PBTXLSW_Sheet;
    cc:PBTXLSW_Cell;
    i,row,col:longword;
    mrow,mcol:longword;
    done:boolean;
    txt:ansistring;
    rowln:boolean;
begin
   Result := '';
   try

   mrow := 0;
   mcol := 0;
   done := false;
   ss := xls.aData;
   for i := 1 to xls.aSheetsCount do
   begin
      if i = sheet then
      begin
{         mrow := 0;
         mcol := 0;
         total := 0;
         cc := ss.cells;
         while cc <> nil do
         begin
            if mrow < cc.row then mrow := cc.row;
            if mcol < cc.col then mcol := cc.col;
            inc(total);
            cc.touch := false;
            cc := cc.next;
         end;
}
         mrow := ss.maxr;
         mcol := ss.maxc;

         if ss.total > 0 then
         begin
            for row := 1 to mrow do
            begin
               rowln := false;
               for col := 1 to mcol do
               begin
                  cc := ss.cells;
                  while cc <> nil do
                  begin
                     if (cc.row = row) and (cc.col = col) then
                     begin
                        if not rowln then
                        begin
                           Txt := txt + '<row r="'+ansistring(tostr(row))+'" spans="1:3" x14ac:dyDescent="0.25">';
                           rowln := true;
                        end;

                        Txt := txt + '<c r="'+ExcelAdr(row,col)+'"';
                        // add style
                        if cc.style <> 0 then  Txt := txt + ' s="'+ansistring(toStr(cc.style))+'"';
                        if length(cc.value)> 0 then
                        begin
                           if cc.typ = 1 then //string
                           begin
                              Txt := txt + ' t="s"';
                           end;
                           Txt :=  txt + '>';
                           if cc.typ = 2 then //formula
                           begin
                              Txt :=  txt + '<f>'+cc.value+'</f>';
                              cc.value := '0'; //TODO danger if i have to make two generate !!! :(
                           end;
                           Txt :=  txt + '<v>'+cc.value+'</v></c>';
                        end else begin
                           // add empty
                           Txt := txt + '/>';
                        end;
                     end;
                     cc := cc.next;
                  end;
               end;
               if rowln then  Txt := txt + '</row>';
            end;
         end;

         done := true;
         break;
      end;
      ss := ss.next;
   end;

   if not done then exit;

   Result := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+#13#10
           + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
           + 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
           + '<dimension ref="A1:'+ExcelAdr(mrow,mcol)+'"/>'
           + '<sheetViews>'
           + '<sheetView tabSelected="1" workbookViewId="0">'
           + '<selection activeCell="A1" sqref="A1"/>'  //focus
           + '</sheetView>'
           + '</sheetViews>'
           + '<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>'
           + '<sheetData>' + Txt
           + '</sheetData>'
           + '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
           + '</worksheet>';

   except
      Result := '';
   end;
end;

//------------------------------------------------------------------------------
function GetXlShareStrings(xls:BTXlsxWriter):ansistring; // xl/sharedStrings.xml
var s:ansistring;
    i:longword;
    t,h:string;
begin
   s := '';
   if xls.aStringCnt <> 0 then
   begin
      t := string(xls.aStringPool);
      for i := 1 to xls.aStringUnq do
      begin
         h := ParseStr(t,i,#27);
         s := s + '<si><t>'+ansistring(h)+'</t></si>';
      end;
   end else begin
      xls.aStringCnt := 1;
      xls.aStringUnq := 1;
      s := s + '<si><t>nop</t></si>';
   end;

   Result := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+#13#10
           + '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           + 'count="'+ansistring(tostr(xls.aStringCnt))+'" uniqueCount="'+ansistring(tostr(xls.aStringUnq))+'">'
           + s +'</sst>';
end;


//------------------------------------------------------------------------------
const L0 : ansistring = '<left/>';
      L1 : ansistring = '<left style="thin"><color indexed="64"/></left>';
      R0 : ansistring = '<right/>';
      R1 : ansistring = '<right style="thin"><color indexed="64"/></right>';
      T0 : ansistring = '<top/>';
      T1 : ansistring = '<top style="thin"><color indexed="64"/></top>';
      B0 : ansistring = '<bottom/>';
      B1 : ansistring = '<bottom style="thin"><color indexed="64"/></bottom>';

function GetXlstyle(xls:BTXlsxWriter):ansistring; // xl/styles.xml
var s:ansistring;
    i:longword;
    ff:PBTXLSW_Font;
    ss:PBTXLSW_Style;
    so:PBTXLSW_StyleColor;
begin
   try
   Result := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+#13#10
           + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           + 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
           + 'mc:Ignorable="x14ac x16r2" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" '
           + 'xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main">';
   i := 3;
   s := '';
   if xls.aFonts <> nil then
   begin
      ff := xls.aFonts;
      while (ff <> nil) do
      begin
         inc(i);
         s := s + '<font><sz val="'+ansistring(tostr(ff.size))+'"/>';
         if (ff.atr and 1) <> 0 then s := s + '<b/>'; // bold
         if (ff.atr and 2) <> 0 then s := s + '<i/>'; // italic
         if (ff.atr and 4) <> 0 then s := s + '<u/>'; // underline
         if ff.color = 0 then s := s + '<color theme="1"/>'
                         else s := s + '<color rgb="'+ansistring(ToHex(ff.color,8))+'"/>';
//         s := s + '<name val="'+ff.name+'"/><family val="2"/><charset val="204"/><scheme val="minor"/></font>';
         s := s + '<name val="'+ff.name+'"/><family val="2"/><charset val="'+ansistring(tostr(ff.charset))+'"/></font>';
         ff := ff.next;
      end;
   end;

   Result := Result + '<fonts count="'+ansistring(tostr(i))+'" x14ac:knownFonts="'+ansistring(tostr(i))+'">'
           //  '<font><b/><sz... for bold
           //  '<font><i/><sz... for italic
           + '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><charset val="204"/><scheme val="minor"/></font>'
           + '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><charset val="204"/><scheme val="minor"/></font>'
           + '<font><i/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><charset val="204"/><scheme val="minor"/></font>'
           + s
           + '</fonts>';

   i := 2; // default fill color styles
   s := '';
   if xls.aStyleColors <> nil then
   begin
      so := xls.aStyleColors;
      while (so <> nil) do
      begin
         inc(i);
         s := s + '<fill><patternFill patternType="solid"><fgColor rgb="'+ansistring(ToHex(so.color,8))+'"/><bgColor indexed="64"/></patternFill></fill>';
         so := so.next;
      end;
   end;

   Result := Result + '<fills count="'+ansistring(tostr(i))+'">'
           // fill for rgb color
           //	'<fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/><bgColor indexed="64"/></patternFill></fill>'
           + '<fill><patternFill patternType="none"/></fill>'
           + '<fill><patternFill patternType="gray125"/></fill>'
           + s
           + '</fills>';

   if xls.aBorderuse then
   begin
      Result := Result
           + '<borders count="16">'
           + '<border>'+ L0 + R0 + T0 + B0 +'<diagonal/></border>'
           + '<border>'+ L0 + R0 + T0 + B1 +'<diagonal/></border>'
           + '<border>'+ L0 + R0 + T1 + B0 +'<diagonal/></border>'
           + '<border>'+ L0 + R0 + T1 + B1 +'<diagonal/></border>'
           + '<border>'+ L0 + R1 + T0 + B0 +'<diagonal/></border>'
           + '<border>'+ L0 + R1 + T0 + B1 +'<diagonal/></border>'
           + '<border>'+ L0 + R1 + T1 + B0 +'<diagonal/></border>'
           + '<border>'+ L0 + R1 + T1 + B1 +'<diagonal/></border>'
           + '<border>'+ L1 + R0 + T0 + B0 +'<diagonal/></border>'
           + '<border>'+ L1 + R0 + T0 + B1 +'<diagonal/></border>'
           + '<border>'+ L1 + R0 + T1 + B0 +'<diagonal/></border>'
           + '<border>'+ L1 + R0 + T1 + B1 +'<diagonal/></border>'
           + '<border>'+ L1 + R1 + T0 + B0 +'<diagonal/></border>'
           + '<border>'+ L1 + R1 + T0 + B1 +'<diagonal/></border>'
           + '<border>'+ L1 + R1 + T1 + B0 +'<diagonal/></border>'
           + '<border>'+ L1 + R1 + T1 + B1 +'<diagonal/></border>'
           + '</borders>';
   end else begin
      Result := Result // no border default
           + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>';
   end;

      Result := Result
           + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>';

   i := 1; // default draw styles
   s := '';
   if xls.aStyle <> nil then
   begin
      ss := xls.aStyle;
      while (ss <> nil) do
      begin
         inc(i);
         s := s + '<xf numFmtId="0" fontId="'+ansistring(tostr(ss.font))+'" fillId="'+ansistring(tostr(ss.colorID))+'" borderId="'+ansistring(tostr(ss.border))+'" xfId="0"';
         if ss.font <> 0 then s := s + ' applyFont="1"';
         if ss.colorID <> 0 then s := s + ' applyFill="1"';
         if ss.border <> 0 then s := s + ' applyBorder="1"';
         if ss.Align <> 0 then s := s + ' applyAlignment="1"';
         if ss.Align = 0 then
         begin
            s := s + '/>';  // close tag
         end else begin
            s := s + '>';
            if ss.Align = 1 then s := s + '<alignment horizontal="left"/>';
            if ss.Align = 2 then s := s + '<alignment horizontal="center"/>';
            if ss.Align = 3 then s := s + '<alignment horizontal="right"/>';
            s := s + '</xf>' // close tag
         end;
         ss := ss.next;
      end;
   end;

   Result := Result + '<cellXfs count="'+ansistring(tostr(i))+'">'
           // all indexes start from 0
           // index for s="x" in cell
           // syles combined
           // <xf numFmtId="0" fontId="1" fillId="10" borderId="5" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
           // have applay if in this style have font fill or border
           + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' // default style no s="0"
           + s
           + '</cellXfs>'

           + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
           + '<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>'
           + '<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
           + '<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext>'
           + '<ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">'
           + '<x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext></extLst></styleSheet>';
   except
      Result := '';
   end;
end;

//------------------------------------------------------------------------------
function GetXlWorkbook(xls:BTXlsxWriter):ansistring; // xl/workbook.xml
var i:integer;
    ss:PBTXLSW_Sheet;
begin
   try
   Result := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+#13#10
           + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
           + 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" '
           + 'xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">'
           + '<fileVersion appName="xl" lastEdited="6" lowestEdited="6" rupBuild="14420"/>'
           + '<workbookPr defaultThemeVersion="164011"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
           + '<mc:Choice Requires="x15"><x15ac:absPath url="E:\testZ0002\" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/>'
           + '</mc:Choice></mc:AlternateContent><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12300"/>'
           + '</bookViews><sheets>';
   ss := xls.aData;
   for i := 1 to xls.aSheetsCount do
   begin
      Result := Result
           + '<sheet name="'+ss.name+'" sheetId="'+ansistring(toStr(i))+'" r:id="rId'+ansistring(toStr(i))+'"/>'; //TOTO
      ss := ss.next;
   end;
      Result := Result
           + '</sheets><calcPr calcId="162913"/><extLst>'
           + '<ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">'
           + '<x15:workbookPr chartTrackingRefBase="1"/></ext></extLst></workbook>';
   except
      Result := '';
   end;
end;


//------------------------------------------------------------------------------
constructor BTXlsxWriter.Create;
begin
   aData := nil;
   aStyle := nil;
   aFonts := nil;
   aStyleColors := nil;
   Reset;
end;

//------------------------------------------------------------------------------
procedure   BTXlsxWriter.Reset;
var ss,s :PBTXLSW_Sheet;
    cc,c :PBTXLSW_Cell;
    st,t :PBTXLSW_Style;
    so,o :PBTXLSW_StyleColor;
    ff,f :PBTXLSW_Font;
begin
   try
      if aData <> nil then
      begin
         ss := aData;
         while(ss<>nil) do
         begin
            s := ss;
            if s.cells <> nil then
            begin
               cc := s.cells;
               while(cc<>nil) do
               begin
                  c := cc;
                  cc := c.next;
                  Dispose(c);
               end;
            end;
            ss := s.next;
            Dispose(s);
         end;
      end;
      if aStyle <> nil then
      begin
         st := aStyle;
         while (st<>nil) do
         begin
            t := st;
            st := t.next;
            Dispose(t);
         end;
      end;
      if aStyleColors <> nil then
      begin
         so := aStyleColors;
         while (so<>nil) do
         begin
            o := so;
            so := o.next;
            Dispose(o);
         end;
      end;

      if aFonts <> nil then
      begin
         ff := aFonts;
         while (ff<>nil) do
         begin
            f := ff;
            ff := f.next;
            Dispose(f);
         end;
      end;
   except
      //??
   end;
   aFonts := nil;
   aStyle := nil;
   aData := nil;
   aStyleColors := nil;
   aSheetsCount := 0;
   aStringPool := #27;
   aStringCnt := 0;
   aStringUnq := 0;
   aBorderUse := false;
end;

//------------------------------------------------------------------------------
destructor  BTXlsxWriter.Destroy;
begin
   Reset;
   inherited;
end;

//------------------------------------------------------------------------------
function    BTXlsxWriter.AddSheet(const name:widestring):nativeUint;
var ss,c:PBTXLSW_Sheet;
begin
   Result := 0;
   inc(aSheetsCount);
   try
      new(ss);
      ss.next := nil;
      ss.cells := nil;
      ss.total := 0;
      ss.maxr := 0;
      ss.maxc := 0;
      ss.name := Unicode2Utf8(trim(name));
      if aData = nil then
      begin
         aData := ss;
      end else begin
         c := aData;
         while c.next <> nil do c := c.next;
         c.next := ss;
      end;
      Result := nativeUint(ss);
   except
//      Result := 0;
   end;
end;

//------------------------------------------------------------------------------
function  ISDigit(s:ansistring):boolean;
var i,j,c:longword;
begin
   Result := false;
   j := length(s);
   c := 0;
   if j > 0 then
   begin
      for i := 1 to j do
      begin
         //todo not so good must have logis to exclude only .
         if (s[i] >= '0') and (s[i] <= '9') then inc(c);
         if s[i] = '.' then inc(c);
      end;
   end;

   if j = c then  Result := true;
end;

//------------------------------------------------------------------------------
procedure AdrToCord(Adr:widestring; var cc,rr:longword);
var j,p,k:longword;
    ad,l1,n1:string;
begin
   cc := 0;
   rr := 0;
   Ad := UpperCase(string(Adr));
   j := length(Ad);
   if j > 0 then
   begin
      l1 := '';
      n1 := '';
      p := 0;
      for k := 1 to j do
      begin
         if ad[k] < 'A' then p := 1; // number start
         if p = 0 then l1 := l1 + ad[k] else n1 := n1 + ad[k];
      end;
      rr := ToVal(n1);
      j := length(l1);
      cc := 0;
      for k := 1 to j do
      begin
         cc := (cc * 26) + (longword(byte(l1[k])) - 64); //A-65
      end;
      if longint(cc) < 0 then cc := 1;
   end;
end;


//------------------------------------------------------------------------------
function    BTXlsxWriter.AddCell(sheet:nativeUint; const Adr,value:widestring; val_typ:longword = 0):nativeUint;
var c,r:longword;
begin
   AdrToCord(adr,c,r);
   if (c<>0) and (r<>0) then  Result := AddCell(sheet,c,r,Value,val_typ) else Result := 0;
end;

//------------------------------------------------------------------------------
function    BTXlsxWriter.AddCell(sheet:nativeUint; Col,Row:longword; const value:widestring; val_typ:longword = 0):nativeUint;
var ss :PBTXLSW_Sheet;
    cc,c:PBTXLSW_Cell;
    s,h:ansistring;
    t:string;
    i,j:longword;
    have:boolean;
begin
   Result := 0;
   try
      cc := nil;
      ss := PBTXLSW_Sheet(sheet);

      // find did I not already have addressed this cell
      have := false;
      c := ss.Cells;
      while c <> nil do
      begin
         if (c.row = row) and (c.col = col) then
         begin
            cc := c;
            Have := true;
            break;
         end;
         c := c.next;
      end;

      if not have then
      begin
         new(cc);
         cc.next := nil;
         cc.row := row;
         cc.col := col;
         cc.Style := 0; // default
         inc(ss.total);
         if ss.maxr < row then ss.maxr := row;
         if ss.maxc < col then ss.maxc := col;
      end;

      s := Unicode2Utf8(trim(value)) ;//HTMLEncode(trim(value)));
      if length(s) > 0 then
      begin
         if val_typ = 0 then
         begin  // type is ) default generic test for string and set typ = 1 if true
            if not IsDigit(s) then
            begin
               inc(aStringCnt);
               if Pos(#27+s+#27,aStringPool) = 0 then
               begin
                  aStringPool := aStringPool + s + #27;
                  inc(aStringUnq);
               end;
               // Find the position
               t := string(aStringPool);
               j := 0;
               for i := 1 to aStringUnq do
               begin
                  h := ansistring(ParseStr(t,i,#27));
                  if h = s then begin j:= i-1; Break; end;
               end;
               s := ansistring(tostr(j));   // string index start from zero
               val_typ := 1;
            end;
         end;
      end;

      cc.value := s;
      cc.typ := val_typ;

      if not Have then // it is new link it
      begin
         if ss.cells = nil then
         begin
            ss.Cells := cc;
         end else begin
            c := ss.Cells;
            while c.next <> nil do c := c.next;
            c.next := cc;
         end;
      end;
      Result := nativeUint(cc);
   except
//      Result := 0;
   end;
end;

//------------------------------------------------------------------------------
function    BTXlsxWriter.SetUpFont(const name:string; size,atr,color:longword; charset:longword=1):longword;
var ff,f :PBTXLSW_Font;
    i:longword;
begin
   Result := 0; // default;
   try
      new(ff);
      ff.Next := nil;
      ff.name := ansistring(trim(name));
      ff.size := size;
      ff.atr := atr;
      ff.color := color;
      ff.charset := charset;
      i := 2;  // 0 - default  1-default bold  2-default italic
      if aFonts = nil then
      begin
         aFonts := ff;
      end else begin
         inc(i);
         f := aFonts;
         while (f.next <> nil) do
         begin
            inc(i);
            f := f.next;
         end;
         f.next := ff;
      end;
      Result := i + 1; // start from 3
   except
//      Result := 0;
   end;
end;

//------------------------------------------------------------------------------
function    BTXlsxWriter.AddCellStyle(cell:nativeUint; color,border,font,align:longword):boolean;
var st,t :PBTXLSW_Style;
    cc :PBTXLSW_Cell;
    oo,o :PBTXLSW_StyleColor;
    newstyle:boolean;
    newcolor:boolean;
    i,j:longword;
begin
   if border > 15 then border := 0;
   if align > 3 then align := 0;

   Result := true;
   try
      newstyle := true;
      i := 0;
      if aStyle <> nil then // search for exist style
      begin
         st := aStyle;
         while st<>nil do
         begin
            inc(i);
            if (st.color = color) and (st.border = border) and (st.font = font) and (st.align = align) then
            begin
               newstyle := false;
               Break;  // is is the style 1...
            end;
            st := st.next;
         end;
      end;

      newcolor := true;
      j := 0;
      if color <> 0 then
      begin
         if aStyleColors <> nil then
         begin
            oo := aStyleColors;
            while oo<>nil do
            begin
               inc(j);
               if (oo.color = color) then
               begin
                  newcolor := false;
                  Break;  // is is the style 1...
               end;
               oo := oo.next;
            end;
         end;
         if newcolor then
         begin
            new(oo);
            oo.Next := nil;
            oo.color := color;
            if aStyleColors = nil then
            begin
               aStyleColors := oo;
               j := 1;
            end else begin
               o := aStyleColors;
               j := 2;
               while (o.next <> nil) do
               begin
                  inc(j);
                  o := o.next;
               end;
               o.next := oo;
            end;
         end;
         j := j + 1; // I hev 2 default strat from zero first myne must be 2
      end;


      if newstyle then
      begin
         new(st);
         st.next := nil;
         st.color := color;
         st.colorId := j;
         st.border := border and $F;
         st.align := align;
         if border <> 0 then aBorderUse := true;
         st.font := font;
         if aStyle  = nil then
         begin
            aStyle := st;
            i := 1;
         end else begin
            i := 2;
            t := aStyle;
            while (t.Next <> nil ) do
            begin
               inc(i);
               t := t.next;
            end;
            t.next := st;
         end;
      end;
      cc := PBTXLSW_Cell(Cell);
      cc.style := i;
   except
      Result := false;
   end;
end;



(* xlsx structure
  _rels/.rels
  docProps/app.xml
  docProps/core.xml
  xl/_rels/workbook.xml.rels
  xl/theme/theme1.xml
  xl/worksheets/sheet1.xml
  xl/sharedStrings.xml
  xl/styles.xml
  xl/workbook.xml
  [Control_Types].xml
 *)

//------------------------------------------------------------------------------
function    BTXlsxWriter.Generate(const FileName:string):boolean;
var z:TZipWrite;
    s:ansistring;
    i:longword;
begin
   Result := false;
   try
   if aSheetsCount <> 0 then
   begin
      z := TZipWrite.Create(FileName);
      z.AddDeflated('_rels/.rels',@rels[1],length(rels));
      s:= GetContentTypes(self);
      z.AddDeflated('[Content_Types].xml',@s[1],length(s));
      s:= GetPropsApp(self);
      z.AddDeflated('docProps/app.xml',@s[1],length(s));
      s:= GetPropsCore;
      z.AddDeflated('docProps/core.xml',@s[1],length(s));
      s:= GetXlRelsWorkbook(self);
      z.AddDeflated('xl/_rels/workbook.xml.rels',@s[1],length(s));
      s:= GetXlTheme;
      z.AddDeflated('xl/theme/theme1.xml',@s[1],length(s));

      for i := 1 to aSheetsCount do
      begin
         s:= GetXlSheet(self,i);
//         if length(s) = 0 then

         z.AddDeflated('xl/worksheets/sheet'+ansistring(tostr(i))+'.xml',@s[1],length(s));
      end;

      s:= GetXlShareStrings(self);
      z.AddDeflated('xl/sharedStrings.xml',@s[1],length(s));
      s:= GetXlStyle(self);
//         if length(s) = 0 then
      z.AddDeflated('xl/styles.xml',@s[1],length(s));
      s:= GetXlWorkbook(self);  //todo can return error
//         if length(s) = 0 then
      z.AddDeflated('xl/workbook.xml',@s[1],length(s));
      z.Destroy;

      Result := true;
   end;
   except
      Result := false;
   end;
end;


//------------------------------------------------------------------------------
function    BTXlsxWriter.GenerateCSV(const FileName:string; sheet:nativeUint; Separator:ansistring=';'; PutOnLast:boolean=false):boolean;
var ss:PBTXLSW_Sheet;
    cc:PBTXLSW_Cell;
    i,row,col,c:longword;
    txt:ansistring;
    t:string;
begin
   Result := false; //fail
   Txt := '';
   try
      t := string(aStringPool); // already uft-8
      ss := PBTXLSW_Sheet(sheet);
      if ss <> nil then
      begin
         if ss.total > 0 then
         begin
            for row := 1 to ss.maxr do
            begin
               for col := 1 to ss.maxc do
               begin
                  cc := ss.cells;
                  while cc <> nil do
                  begin
                     if (cc.row = row) and (cc.col = col) then
                     begin
                        if length(cc.value)> 0 then
                        begin
                           if cc.typ = 0 then
                           begin
                              Txt := Txt + cc.value;
                           end;
                           if cc.typ = 1 then //string
                           begin
                              val(string(cc.value),c,i);
                              if i = 0 then
                              begin
                                 // Get txt from pool by value must be in utf8
                                 Txt := txt + ansistring(ParseStr(t,c+1,#27));
                              end;
                           end;
//                           if cc.typ = 2 then //formula
                        end;
                     end; // my row,col found
                     cc := cc.next;
                  end; // while run columns
                  if col <> ss.maxc then Txt := Txt + Separator
                                    else if PutOnLast then Txt := Txt + Separator; // for the last
               end; //next col
               Txt := Txt + #13#10;
            end; // next row
            //Write Txt
            //bom for utf-8  $EF $BB $BF - utf8
            txt := #$EF#$BB#$BF + txt;
            Result := FileSave(FileName,txt);
         end; // sheet hava dat
      end; // have sheet
   except
      Result := false;
   end;
end;


end.
