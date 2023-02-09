unit BdocxWriter;

interface


// under constructions

type  BTDocxWriter = class
         private
            aPagesCount,aLineCount,aParCount,aWordCount,aCharCount:longword;
            aParagraphs:pointer;
            aCurParag:pointer;
            procedure   _NewParag;
         public
            constructor Create;
            destructor  Destroy; override;
            procedure   Reset;
            function    Generate(const FileName:string):boolean;

            procedure   AddNewPage(landscape:boolean = false);
            procedure   AddNewPageA4(landscape:boolean = false);
            procedure   AddText(const Txt:string);
            procedure   AddNewLine; // create ne paragraph.

         end;


implementation

uses BpasZlib,BStrTools,BUnicode,BDate;


type  TParagraph = record
         next :pointer;
         Txt :widestring;
      end;
      PTParagraph = ^TParagraph;


//   _rels/.rels
{
const _rels:ansistring =
'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+#13#10+
'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+
'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'+
'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'+
'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'+
'</Relationships>';
}

const
   Bin_rels_len = 590;
   Bin_rels_len_lzo = 298;
   Bin_rels : array [0..297] of byte = (
      $00,$50,$3C,$3F,$78,$6D,$6C,$20,$76,$65,$72,$73,$69,$6F,$6E,$3D,
      $22,$31,$2E,$30,$22,$20,$65,$6E,$63,$6F,$64,$69,$6E,$67,$3D,$22,
      $55,$54,$46,$2D,$38,$22,$20,$73,$74,$61,$6E,$64,$61,$6C,$6F,$6E,
      $65,$3D,$22,$79,$65,$73,$22,$3F,$3E,$0D,$0A,$3C,$52,$65,$6C,$61,
      $74,$69,$6F,$6E,$73,$68,$69,$70,$73,$20,$78,$6D,$6C,$6E,$73,$3D,
      $22,$68,$74,$74,$70,$3A,$2F,$2F,$73,$63,$68,$65,$6D,$61,$73,$2E,
      $6F,$70,$65,$6E,$44,$03,$00,$08,$66,$6F,$72,$6D,$61,$74,$73,$2E,
      $6F,$72,$67,$2F,$70,$61,$63,$6B,$61,$67,$65,$2F,$32,$30,$30,$36,
      $2F,$72,$2A,$0E,$01,$22,$3E,$2B,$4C,$01,$0C,$20,$49,$64,$3D,$22,
      $72,$49,$64,$33,$22,$20,$54,$79,$70,$65,$20,$03,$6C,$01,$0B,$6F,
      $66,$66,$69,$63,$65,$44,$6F,$63,$75,$6D,$65,$6E,$74,$31,$88,$01,
      $00,$02,$2F,$65,$78,$74,$65,$6E,$64,$65,$64,$2D,$70,$72,$6F,$70,
      $65,$72,$74,$69,$65,$73,$58,$0B,$00,$07,$61,$72,$67,$65,$74,$3D,
      $22,$64,$6F,$63,$50,$72,$6F,$70,$73,$2F,$61,$70,$70,$2E,$78,$6D,
      $6C,$22,$2F,$34,$45,$02,$32,$48,$06,$20,$06,$44,$02,$38,$B4,$03,
      $0B,$2F,$6D,$65,$74,$61,$64,$61,$74,$61,$2F,$63,$6F,$72,$65,$3C,
      $3C,$02,$64,$04,$3A,$41,$02,$31,$20,$09,$40,$02,$20,$01,$88,$04,
      $2C,$84,$00,$44,$0B,$D4,$23,$03,$77,$6F,$72,$64,$2F,$64,$CC,$2B,
      $F5,$11,$2F,$2B,$15,$08,$3E,$11,$00,$00   );





//   docProps/app.xml

function getappxml(doc:BTDocxWriter):ansistring;
begin
   Result :=

'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+#13#10+
'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'+
'<Template>Normal</Template>'+
'<TotalTime>0</TotalTime>'+
'<Pages>'+ansistring(ToStr(doc.aPagesCount))+'</Pages>'+
'<Words>'+ansistring(ToStr(doc.aWordCount))+'</Words>'+
'<Characters>'+ansistring(ToStr(doc.aCharCount))+'</Characters>'+
'<Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity>'+
'<Lines>'+ansistring(ToStr(doc.aLineCount))+'</Lines>'+
'<Paragraphs>'+ansistring(ToStr(doc.aParCount))+'</Paragraphs>'+
'<ScaleCrop>false</ScaleCrop>'+
'<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Title</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>'+
'<TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr></vt:lpstr></vt:vector></TitlesOfParts><Company></Company>'+
'<LinksUpToDate>false</LinksUpToDate><CharactersWithSpaces>4</CharactersWithSpaces><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion></Properties>';
end;


//   docProps/core.xml
function getcorexml(doc:BTDocxWriter):ansistring;
var s:ansistring;
begin
   s := ansistring(DateConToUTCstr(GetTodaySys,false,true)); // need system time greanwich

   Result :=  Unicode2Utf8(

'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+#13#10+
'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" '+
'xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'+
'<dc:title></dc:title><dc:subject></dc:subject><dc:creator>Harry Potter</dc:creator><cp:keywords></cp:keywords><dc:description></dc:description>'+
'<cp:lastModifiedBy>Harry Potter</cp:lastModifiedBy><cp:revision>1</cp:revision>'+
'<dcterms:created xsi:type="dcterms:W3CDTF">'+s+'</dcterms:created>'+
'<dcterms:modified xsi:type="dcterms:W3CDTF">'+s+'</dcterms:modified></cp:coreProperties>');
end;


//   word/_rels/document.xml.rels
const xmlrels:ansistring =

'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+#13#10+
'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+
'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>'+
'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'+
'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'+
'<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'+
'<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>'+
'</Relationships>';

const
   Bin_documentxml_len = 817;
   Bin_documentxml_len_lzo = 303;
var
   Bin_documentxml : array [0..302] of byte = (
      $00,$50,$3C,$3F,$78,$6D,$6C,$20,$76,$65,$72,$73,$69,$6F,$6E,$3D,
      $22,$31,$2E,$30,$22,$20,$65,$6E,$63,$6F,$64,$69,$6E,$67,$3D,$22,
      $55,$54,$46,$2D,$38,$22,$20,$73,$74,$61,$6E,$64,$61,$6C,$6F,$6E,
      $65,$3D,$22,$79,$65,$73,$22,$3F,$3E,$0D,$0A,$3C,$52,$65,$6C,$61,
      $74,$69,$6F,$6E,$73,$68,$69,$70,$73,$20,$78,$6D,$6C,$6E,$73,$3D,
      $22,$68,$74,$74,$70,$3A,$2F,$2F,$73,$63,$68,$65,$6D,$61,$73,$2E,
      $6F,$70,$65,$6E,$44,$03,$00,$08,$66,$6F,$72,$6D,$61,$74,$73,$2E,
      $6F,$72,$67,$2F,$70,$61,$63,$6B,$61,$67,$65,$2F,$32,$30,$30,$36,
      $2F,$72,$2A,$0E,$01,$22,$3E,$2B,$4C,$01,$0C,$20,$49,$64,$3D,$22,
      $72,$49,$64,$33,$22,$20,$54,$79,$70,$65,$20,$03,$6C,$01,$0B,$6F,
      $66,$66,$69,$63,$65,$44,$6F,$63,$75,$6D,$65,$6E,$74,$31,$88,$01,
      $09,$2F,$77,$65,$62,$53,$65,$74,$74,$69,$6E,$67,$73,$58,$0A,$04,
      $61,$72,$67,$65,$74,$3D,$22,$29,$50,$00,$03,$2E,$78,$6D,$6C,$22,
      $2F,$34,$21,$02,$32,$44,$06,$20,$28,$21,$02,$73,$2F,$14,$02,$E4,
      $02,$3A,$09,$02,$31,$20,$2C,$08,$02,$01,$74,$79,$6C,$65,$29,$19,
      $04,$73,$9C,$01,$3A,$F9,$01,$35,$20,$2B,$F8,$01,$02,$74,$68,$65,
      $6D,$65,$40,$0A,$D0,$30,$99,$01,$2F,$95,$00,$31,$3A,$0D,$02,$34,
      $48,$06,$20,$28,$14,$06,$05,$66,$6F,$6E,$74,$54,$61,$62,$6C,$29,
      $1C,$02,$27,$48,$00,$F1,$10,$2F,$2B,$A1,$0B,$3E,$11,$00,$00   );




(*
   word/theme/theme1.xml

'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">'
'<a:themeElements><a:clrScheme name="Office">'
'<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
'<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
'<a:dk2><a:srgbClr val="44546A"/></a:dk2>'
'<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
'<a:accent1><a:srgbClr val="5B9BD5"/></a:accent1>'
'<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'
'<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>'
'<a:accent4><a:srgbClr val="FFC000"/></a:accent4>'
'<a:accent5><a:srgbClr val="4472C4"/></a:accent5>'
'<a:accent6><a:srgbClr val="70AD47"/></a:accent6>'
'<a:hlink><a:srgbClr val="0563C1"/></a:hlink>'
'<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>'
'</a:clrScheme><a:fontScheme name="Office"><a:majorFont>'
'<a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/>'
'<a:font script="Jpan" typeface="游ゴシック Light"/>'
'<a:font script="Hang" typeface="맑은 고딕"/>'
'<a:font script="Hans" typeface="等线 Light"/>'
'<a:font script="Hant" typeface="新細明體"/>'
'<a:font script="Arab" typeface="Times New Roman"/>'
'<a:font script="Hebr" typeface="Times New Roman"/>'
'<a:font script="Thai" typeface="Angsana New"/>'
'<a:font script="Ethi" typeface="Nyala"/>'
'<a:font script="Beng" typeface="Vrinda"/>'
'<a:font script="Gujr" typeface="Shruti"/>'
'<a:font script="Khmr" typeface="MoolBoran"/>'
'<a:font script="Knda" typeface="Tunga"/>'
'<a:font script="Guru" typeface="Raavi"/>'
'<a:font script="Cans" typeface="Euphemia"/>'
'<a:font script="Cher" typeface="Plantagenet Cherokee"/>'
'<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>'
'<a:font script="Tibt" typeface="Microsoft Himalaya"/>'
'<a:font script="Thaa" typeface="MV Boli"/>'
'<a:font script="Deva" typeface="Mangal"/>'
'<a:font script="Telu" typeface="Gautami"/>'
'<a:font script="Taml" typeface="Latha"/>'
'<a:font script="Syrc" typeface="Estrangelo Edessa"/>'
'<a:font script="Orya" typeface="Kalinga"/>'
'<a:font script="Mlym" typeface="Kartika"/>'
'<a:font script="Laoo" typeface="DokChampa"/>'
'<a:font script="Sinh" typeface="Iskoola Pota"/>'
'<a:font script="Mong" typeface="Mongolian Baiti"/>'
'<a:font script="Viet" typeface="Times New Roman"/>'
'<a:font script="Uigh" typeface="Microsoft Uighur"/>'
'<a:font script="Geor" typeface="Sylfaen"/>'
'</a:majorFont><a:minorFont>'
'<a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/>'
'<a:font script="Jpan" typeface="游明朝"/>'
'<a:font script="Hang" typeface="맑은 고딕"/>'
'<a:font script="Hans" typeface="等线"/>'
'<a:font script="Hant" typeface="新細明體"/>'
'<a:font script="Arab" typeface="Arial"/>'
'<a:font script="Hebr" typeface="Arial"/>'
'<a:font script="Thai" typeface="Cordia New"/>'
'<a:font script="Ethi" typeface="Nyala"/>'
'<a:font script="Beng" typeface="Vrinda"/>'
'<a:font script="Gujr" typeface="Shruti"/>'
'<a:font script="Khmr" typeface="DaunPenh"/>'
'<a:font script="Knda" typeface="Tunga"/>'
'<a:font script="Guru" typeface="Raavi"/>'
'<a:font script="Cans" typeface="Euphemia"/>'
'<a:font script="Cher" typeface="Plantagenet Cherokee"/>'
'<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>'
'<a:font script="Tibt" typeface="Microsoft Himalaya"/>'
'<a:font script="Thaa" typeface="MV Boli"/>'
'<a:font script="Deva" typeface="Mangal"/>'
'<a:font script="Telu" typeface="Gautami"/>'
'<a:font script="Taml" typeface="Latha"/>'
'<a:font script="Syrc" typeface="Estrangelo Edessa"/>'
'<a:font script="Orya" typeface="Kalinga"/>'
'<a:font script="Mlym" typeface="Kartika"/>'
'<a:font script="Laoo" typeface="DokChampa"/>'
'<a:font script="Sinh" typeface="Iskoola Pota"/>'
'<a:font script="Mong" typeface="Mongolian Baiti"/>'
'<a:font script="Viet" typeface="Arial"/>'
'<a:font script="Uigh" typeface="Microsoft Uighur"/>'
'<a:font script="Geor" typeface="Sylfaen"/>'
'</a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
'<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr">'
'<a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs>'
'<a:gs pos="50000"><a:schemeClr val="phClr">'
'<a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr">'
'<a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
'<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/>'
'<a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr">'
'<a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs>'
'<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/>'
'</a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst>'
'<a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
'<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">'
'<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
'<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
'<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst>'
'<a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>'
'<a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">'
'<a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>'
'</a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
'<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>'
'<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/>'
'<a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000">'
'<a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/>'
'<a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr">'
'<a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/>'
'</a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/>'
'<a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}">'
'<thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" '
'id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>'

*)
const
   Bin_theme1_len = 6799;
   Bin_theme1_len_lzo = 2217;
   Bin_theme1 : array [0..2216] of byte = (
      $00,$4C,$3C,$3F,$78,$6D,$6C,$20,$76,$65,$72,$73,$69,$6F,$6E,$3D,
      $22,$31,$2E,$30,$22,$20,$65,$6E,$63,$6F,$64,$69,$6E,$67,$3D,$22,
      $55,$54,$46,$2D,$38,$22,$20,$73,$74,$61,$6E,$64,$61,$6C,$6F,$6E,
      $65,$3D,$22,$79,$65,$73,$22,$3F,$3E,$0D,$0A,$3C,$61,$3A,$74,$68,
      $65,$6D,$65,$20,$78,$6D,$6C,$6E,$73,$3A,$61,$3D,$22,$68,$74,$74,
      $70,$3A,$2F,$2F,$73,$63,$68,$65,$6D,$61,$73,$2E,$6F,$70,$65,$6E,
      $4C,$03,$0D,$66,$6F,$72,$6D,$61,$74,$73,$2E,$6F,$72,$67,$2F,$64,
      $72,$61,$77,$5C,$0A,$00,$0A,$6D,$6C,$2F,$32,$30,$30,$36,$2F,$6D,
      $61,$69,$6E,$22,$20,$6E,$61,$6D,$65,$3D,$22,$4F,$66,$66,$69,$63,
      $65,$20,$54,$6A,$0A,$22,$3E,$F0,$0B,$05,$45,$6C,$65,$6D,$65,$6E,
      $74,$73,$60,$02,$01,$63,$6C,$72,$53,$69,$0B,$65,$2B,$C4,$00,$8F,
      $05,$64,$6B,$31,$64,$04,$00,$10,$73,$79,$73,$43,$6C,$72,$20,$76,
      $61,$6C,$3D,$22,$77,$69,$6E,$64,$6F,$77,$54,$65,$78,$74,$22,$20,
      $6C,$61,$73,$74,$43,$6C,$72,$3D,$22,$30,$80,$00,$04,$22,$2F,$3E,
      $3C,$2F,$61,$3A,$D2,$06,$6C,$74,$35,$EC,$00,$29,$DD,$00,$46,$80,
      $00,$DC,$06,$C3,$06,$64,$6B,$32,$8F,$0E,$72,$67,$62,$27,$D0,$01,
      $03,$34,$34,$35,$34,$36,$41,$DC,$04,$C2,$04,$6C,$74,$30,$9C,$00,
      $03,$45,$37,$45,$36,$45,$36,$DC,$04,$C0,$04,$02,$61,$63,$63,$65,
      $6E,$CC,$11,$2A,$4C,$01,$03,$35,$42,$39,$42,$44,$35,$CC,$05,$29,
      $90,$00,$A8,$01,$31,$6C,$01,$02,$44,$37,$44,$33,$31,$2B,$BC,$00,
      $90,$04,$BD,$05,$33,$2F,$CE,$02,$41,$35,$64,$00,$2B,$BC,$00,$90,
      $04,$BD,$05,$34,$2F,$BC,$00,$03,$46,$46,$43,$30,$30,$30,$2B,$BC,
      $00,$90,$04,$BD,$05,$35,$2F,$BC,$00,$03,$34,$34,$37,$32,$43,$34,
      $2B,$BC,$00,$90,$04,$BD,$05,$36,$2F,$BC,$00,$03,$37,$30,$41,$44,
      $34,$37,$2B,$BC,$00,$90,$04,$02,$68,$6C,$69,$6E,$6B,$2F,$B4,$00,
      $02,$30,$35,$36,$33,$43,$F4,$1D,$27,$88,$00,$01,$66,$6F,$6C,$48,
      $33,$B8,$00,$03,$39,$35,$34,$46,$37,$32,$D0,$0B,$28,$97,$00,$2F,
      $61,$3A,$27,$74,$09,$68,$0C,$01,$66,$6F,$6E,$74,$36,$AC,$09,$06,
      $6D,$61,$6A,$6F,$72,$46,$6F,$6E,$74,$60,$05,$00,$1D,$6C,$61,$74,
      $69,$6E,$20,$74,$79,$70,$65,$66,$61,$63,$65,$3D,$22,$43,$61,$6C,
      $69,$62,$72,$69,$20,$4C,$69,$67,$68,$74,$22,$20,$70,$61,$6E,$6F,
      $73,$65,$3D,$22,$30,$32,$30,$46,$30,$33,$30,$32,$85,$00,$34,$84,
      $01,$84,$22,$01,$61,$3A,$65,$61,$29,$F4,$00,$7C,$12,$01,$61,$3A,
      $63,$73,$2F,$48,$00,$7C,$11,$0B,$20,$73,$63,$72,$69,$70,$74,$3D,
      $22,$4A,$70,$61,$6E,$22,$29,$88,$00,$0C,$E6,$B8,$B8,$E3,$82,$B4,
      $E3,$82,$B7,$E3,$83,$83,$E3,$82,$AF,$CD,$0F,$2F,$78,$13,$2B,$DC,
      $00,$01,$48,$61,$6E,$67,$2A,$DC,$00,$0A,$EB,$A7,$91,$EC,$9D,$80,
      $20,$EA,$B3,$A0,$EB,$94,$95,$A8,$0F,$2E,$BD,$00,$73,$2A,$BC,$00,
      $03,$E7,$AD,$89,$E7,$BA,$BF,$3A,$78,$01,$48,$6D,$28,$28,$04,$09,
      $E6,$96,$B0,$E7,$B4,$B0,$E6,$98,$8E,$E9,$AB,$94,$31,$74,$01,$01,
      $41,$72,$61,$62,$2A,$74,$01,$0C,$54,$69,$6D,$65,$73,$20,$4E,$65,
      $77,$20,$52,$6F,$6D,$61,$6E,$31,$C4,$00,$01,$48,$65,$62,$72,$20,
      $0D,$C4,$00,$01,$54,$68,$61,$69,$2A,$C4,$00,$04,$41,$6E,$67,$73,
      $61,$6E,$61,$74,$0C,$31,$7F,$01,$45,$74,$68,$2B,$B4,$00,$02,$4E,
      $79,$61,$6C,$61,$31,$9E,$00,$42,$65,$2C,$1C,$05,$02,$56,$72,$69,
      $6E,$64,$32,$A3,$00,$47,$75,$6A,$2B,$C0,$02,$03,$53,$68,$72,$75,
      $74,$69,$31,$47,$01,$4B,$68,$6D,$2B,$A0,$00,$04,$4D,$6F,$6F,$6C,
      $42,$6F,$72,$33,$15,$04,$4B,$6D,$0D,$20,$28,$98,$05,$01,$54,$75,
      $6E,$67,$34,$F2,$01,$72,$75,$2A,$EC,$03,$01,$52,$61,$61,$76,$32,
      $EF,$01,$43,$61,$6E,$2B,$97,$07,$45,$75,$70,$59,$B7,$69,$32,$4B,
      $01,$43,$68,$65,$2B,$98,$02,$09,$50,$6C,$61,$6E,$74,$61,$67,$65,
      $6E,$65,$74,$20,$6C,$03,$01,$6F,$6B,$65,$65,$31,$77,$03,$59,$69,
      $69,$2B,$5C,$05,$0D,$4D,$69,$63,$72,$6F,$73,$6F,$66,$74,$20,$59,
      $69,$20,$42,$61,$69,$33,$4B,$04,$54,$69,$62,$2B,$34,$09,$28,$D0,
      $00,$04,$48,$69,$6D,$61,$6C,$61,$79,$32,$80,$02,$5D,$3D,$61,$2A,
      $CC,$03,$03,$4D,$56,$20,$42,$6F,$6C,$32,$D7,$03,$44,$65,$76,$2C,
      $A4,$00,$02,$61,$6E,$67,$61,$6C,$31,$F3,$02,$54,$65,$6C,$2B,$18,
      $05,$03,$47,$61,$75,$74,$61,$6D,$32,$48,$01,$01,$54,$61,$6D,$6C,
      $2A,$F0,$01,$01,$4C,$61,$74,$68,$32,$90,$02,$01,$53,$79,$72,$63,
      $2A,$9C,$00,$0D,$45,$73,$74,$72,$61,$6E,$67,$65,$6C,$6F,$20,$45,
      $64,$65,$73,$73,$32,$CF,$00,$4F,$72,$79,$2B,$B8,$02,$01,$4B,$61,
      $6C,$69,$34,$D8,$07,$01,$4D,$6C,$79,$6D,$2A,$74,$01,$03,$4B,$61,
      $72,$74,$69,$6B,$32,$4C,$01,$01,$4C,$61,$6F,$6F,$2A,$A4,$00,$05,
      $44,$6F,$6B,$43,$68,$61,$6D,$70,$32,$AC,$00,$01,$53,$69,$6E,$68,
      $2A,$AC,$00,$08,$49,$73,$6B,$6F,$6F,$6C,$61,$20,$50,$6F,$74,$32,
      $BA,$00,$4D,$6F,$2C,$84,$0C,$7C,$01,$02,$6F,$6C,$69,$61,$6E,$37,
      $BB,$07,$56,$69,$65,$2B,$B8,$07,$20,$01,$37,$10,$55,$69,$67,$2B,
      $48,$02,$28,$80,$08,$66,$03,$75,$72,$31,$2F,$07,$47,$65,$6F,$2B,
      $FC,$0A,$03,$53,$79,$6C,$66,$61,$65,$8B,$8D,$2F,$61,$3A,$2B,$5F,
      $17,$6D,$69,$6E,$3F,$90,$17,$2D,$79,$17,$35,$92,$BB,$32,$30,$20,
      $32,$78,$17,$03,$E6,$98,$8E,$E6,$9C,$9D,$31,$44,$03,$20,$25,$48,
      $17,$34,$60,$01,$2B,$3C,$06,$20,$0E,$33,$17,$41,$72,$69,$33,$34,
      $0D,$2E,$08,$17,$36,$9C,$00,$40,$79,$2B,$C8,$10,$02,$43,$6F,$72,
      $64,$69,$20,$81,$DC,$16,$05,$44,$61,$75,$6E,$50,$65,$6E,$68,$31,
      $40,$05,$20,$00,$00,$C7,$D8,$16,$36,$3C,$0F,$68,$B2,$2A,$FC,$18,
      $20,$21,$B0,$16,$28,$7F,$16,$2F,$61,$3A,$22,$D8,$2A,$24,$B8,$2E,
      $23,$D1,$2B,$6D,$37,$EC,$2E,$08,$66,$69,$6C,$6C,$53,$74,$79,$6C,
      $65,$4C,$73,$23,$F8,$2E,$06,$73,$6F,$6C,$69,$64,$46,$69,$6C,$6C,
      $7C,$06,$23,$DD,$3A,$65,$27,$2C,$37,$01,$70,$68,$43,$6C,$97,$C8,
      $2F,$61,$3A,$2B,$9F,$00,$67,$72,$61,$8C,$06,$0D,$20,$72,$6F,$74,
      $57,$69,$74,$68,$53,$68,$61,$70,$65,$3D,$22,$31,$23,$26,$3A,$67,
      $73,$C8,$0B,$06,$67,$73,$20,$70,$6F,$73,$3D,$22,$30,$98,$02,$33,
      $6C,$01,$70,$0E,$03,$6C,$75,$6D,$4D,$6F,$64,$24,$CA,$3A,$31,$31,
      $22,$78,$3A,$B3,$87,$73,$61,$74,$28,$5E,$00,$30,$35,$25,$C4,$35,
      $03,$61,$3A,$74,$69,$6E,$74,$B6,$05,$36,$37,$D0,$02,$7C,$13,$23,
      $5F,$3C,$43,$6C,$72,$23,$61,$3B,$67,$23,$B8,$3C,$E0,$10,$6C,$07,
      $20,$0B,$10,$02,$29,$B0,$01,$2C,$11,$02,$33,$DC,$0D,$2A,$11,$02,
      $37,$F0,$02,$3D,$10,$02,$84,$1B,$20,$24,$15,$02,$39,$31,$15,$02,
      $38,$78,$0B,$7C,$26,$33,$17,$02,$2F,$61,$3A,$27,$54,$06,$02,$6C,
      $69,$6E,$20,$61,$22,$62,$43,$35,$34,$6C,$2D,$06,$30,$22,$20,$73,
      $63,$61,$6C,$65,$64,$7C,$34,$24,$54,$40,$F4,$3B,$70,$33,$EC,$01,
      $20,$24,$A4,$07,$36,$30,$05,$2B,$06,$08,$30,$32,$31,$79,$03,$39,
      $6C,$15,$37,$78,$03,$74,$1B,$A8,$4D,$20,$02,$A4,$07,$2B,$10,$02,
      $74,$26,$23,$22,$3F,$61,$3A,$2C,$10,$02,$68,$23,$7C,$0D,$04,$61,
      $3A,$73,$68,$61,$64,$65,$BC,$4D,$94,$05,$98,$05,$33,$94,$05,$28,
      $18,$02,$C0,$05,$60,$28,$20,$04,$D9,$0B,$39,$28,$48,$07,$2B,$78,
      $02,$68,$21,$9C,$0D,$2B,$1A,$02,$37,$38,$C4,$24,$33,$14,$02,$20,
      $17,$AC,$07,$70,$88,$2D,$6E,$10,$6C,$6E,$2A,$A4,$10,$06,$6C,$6E,
      $20,$77,$3D,$22,$36,$33,$35,$48,$46,$0F,$63,$61,$70,$3D,$22,$66,
      $6C,$61,$74,$22,$20,$63,$6D,$70,$64,$3D,$22,$73,$22,$B8,$2B,$06,
      $61,$6C,$67,$6E,$3D,$22,$63,$74,$72,$A0,$7E,$20,$13,$64,$11,$05,
      $70,$72,$73,$74,$44,$61,$73,$68,$A0,$2E,$90,$8E,$B8,$30,$08,$6D,
      $69,$74,$65,$72,$20,$6C,$69,$6D,$3D,$22,$78,$1F,$27,$FA,$45,$6C,
      $6E,$74,$2C,$AC,$13,$01,$31,$32,$37,$30,$20,$79,$72,$02,$39,$30,
      $20,$71,$E4,$04,$44,$48,$2C,$90,$07,$03,$65,$66,$66,$65,$63,$74,
      $2A,$D8,$07,$29,$44,$00,$64,$2D,$BB,$01,$4C,$73,$74,$AC,$85,$B8,
      $01,$9C,$05,$28,$74,$00,$2D,$38,$00,$20,$0F,$B0,$00,$60,$09,$00,
      $04,$6F,$75,$74,$65,$72,$53,$68,$64,$77,$20,$62,$6C,$75,$72,$52,
      $61,$64,$3D,$22,$35,$37,$31,$60,$28,$04,$64,$69,$73,$74,$3D,$22,
      $31,$B4,$29,$02,$64,$69,$72,$3D,$22,$27,$20,$13,$28,$FC,$09,$2D,
      $B8,$12,$D8,$AE,$2A,$4C,$51,$61,$7C,$30,$BC,$02,$02,$61,$6C,$70,
      $68,$61,$A1,$4E,$36,$2A,$30,$17,$B8,$05,$9C,$C9,$27,$38,$02,$94,
      $01,$B0,$1E,$94,$DC,$40,$28,$B4,$01,$CC,$1E,$2C,$3C,$00,$88,$04,
      $01,$61,$3A,$62,$67,$60,$EC,$84,$04,$C4,$02,$80,$5B,$70,$02,$6C,
      $1D,$33,$88,$10,$B8,$2B,$2C,$9C,$00,$3F,$40,$0D,$74,$09,$28,$79,
      $1C,$39,$74,$A1,$B4,$67,$2B,$54,$11,$74,$E6,$94,$8A,$50,$15,$27,
      $EC,$01,$80,$1B,$2B,$B8,$01,$20,$2C,$E8,$17,$29,$54,$02,$EA,$2C,
      $61,$3A,$2B,$54,$02,$70,$15,$94,$12,$2B,$AD,$13,$39,$70,$7D,$AC,
      $18,$2C,$80,$16,$68,$A3,$7C,$02,$2F,$0C,$03,$20,$10,$E8,$1F,$29,
      $68,$02,$28,$B0,$01,$2B,$68,$02,$64,$16,$31,$68,$02,$68,$48,$68,
      $10,$2E,$EC,$18,$78,$05,$7C,$02,$3D,$68,$02,$20,$03,$90,$18,$29,
      $F0,$19,$27,$68,$0A,$2D,$DC,$04,$60,$21,$90,$13,$33,$B8,$01,$20,
      $1A,$34,$18,$2E,$D4,$0A,$60,$C4,$26,$5C,$29,$94,$45,$23,$EC,$63,
      $2A,$78,$62,$0B,$6F,$62,$6A,$65,$63,$74,$44,$65,$66,$61,$75,$6C,
      $74,$73,$23,$F8,$55,$03,$65,$78,$74,$72,$61,$43,$26,$D8,$62,$A4,
      $81,$02,$61,$3A,$65,$78,$74,$C0,$61,$00,$26,$65,$78,$74,$20,$75,
      $72,$69,$3D,$22,$7B,$30,$35,$41,$34,$43,$32,$35,$43,$2D,$30,$38,
      $35,$45,$2D,$34,$33,$34,$30,$2D,$38,$35,$41,$33,$2D,$41,$35,$35,
      $33,$31,$45,$35,$31,$30,$44,$42,$32,$7D,$22,$3E,$3C,$74,$68,$6D,
      $31,$35,$3A,$8C,$0F,$03,$46,$61,$6D,$69,$6C,$79,$25,$F4,$65,$9C,
      $02,$2F,$05,$66,$6D,$26,$A4,$4D,$03,$2E,$63,$6F,$6D,$2F,$6F,$23,
      $85,$65,$2F,$90,$07,$23,$06,$66,$31,$32,$38,$06,$66,$20,$69,$5C,
      $DA,$00,$17,$7B,$36,$32,$46,$39,$33,$39,$42,$36,$2D,$39,$33,$41,
      $46,$2D,$34,$44,$42,$38,$2D,$39,$43,$36,$42,$2D,$44,$36,$43,$37,
      $44,$46,$44,$43,$35,$38,$39,$46,$7D,$22,$20,$76,$90,$05,$0E,$34,
      $41,$33,$43,$34,$36,$45,$38,$2D,$36,$31,$43,$43,$2D,$34,$36,$30,
      $78,$17,$0D,$38,$39,$2D,$37,$34,$32,$32,$41,$34,$37,$41,$38,$45,
      $34,$41,$7D,$CB,$43,$65,$78,$74,$94,$28,$EC,$21,$06,$2F,$61,$3A,
      $74,$68,$65,$6D,$65,$3E,$11,$00,$00   );




//   word/document.xml
function getdocumentxml(doc:BTDocxWriter):ansistring;
var txt:ansistring;
    p:PTParagraph;
    s,spc,ww:widestring;
    par:string;
    i,j:longword;
    goout,ops:boolean;
    c:widechar;
begin
   txt := '';
   p := PTParagraph(doc.aParagraphs);
   if p <> nil then
   begin
      repeat
         // start paragraph
         s := '<w:p><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr>';   // lang ???

         par := string(p.Txt);
         goout := false;
         i := 1;
         j := length(p.Txt);
         ww := '';
         repeat
            c := p.Txt[i];
            if c = #27 then
            begin // style

            end;

            if c <> ' ' then
            begin
               ww := ww + c;
            end else begin
               s := s + '<w:t>' + ww + '</w:t>';
               spc := ' ';
               if i < j then
               begin
                  ops := false;
                  repeat
                     if i < j then
                     begin
                        if p.Txt[i+1] = ' ' then
                        begin
                           spc := spc + ' ';
                           inc(i);
                        end else ops := true;
                     end;
                  until (i>=j) or ops;
               end;
               s := s + '<w:t xml:space="preserve">' + spc + '</w:t>';
               ww := '';
            end;
            if i = j then
            begin
               if length(ww) <> 0 then
               begin
                  s := s + '<w:t>' + ww + '</w:t>';
               end;
               goout := true;
            end;
            inc(i);

         until goout;
         txt := txt + Unicode2UTF8(s);
         // end paragraph
         txt := txt + '</w:r></w:p>';

         p := p.next; // get next
      until p = nil;
   end;





//rules
{
	<w:r w:rsidRPr="000922F2">
				<w:rPr>
					<w:b/>
					<w:color w:val="FF0000"/>
					<w:highlight w:val="yellow"/>
					<w:lang w:val="en-US"/>
				</w:rPr>
				<w:t>ce</w:t>
			</w:r>

			<w:r w:rsidR="00BE1B61" w:rsidRPr="00BE1B61">
				<w:rPr>
					<w:rFonts w:ascii="Arial Black" w:hAnsi="Arial Black"/>
					<w:sz w:val="24"/>
					<w:szCs w:val="24"/>
					<w:lang w:val="en-US"/>
				</w:rPr>
				<w:t>ne</w:t>
			</w:r>
		</w:p>  // end paragraph

    <w:t xml:space="preserve">Tata </w:t>   or <w:t xml:space="preserve"> </w:t>


      }


   Result :=

'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+#13#10+
'<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" '+
'xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '+
'xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '+
'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" '+
'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '+
'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '+
'xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '+
'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '+
'xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '+
'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '+
'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se wp14">'+
'<w:body>';
//'<w:p w:rsidR="00881B0E" w:rsidRPr="005E430E" w:rsidRDefault="005E430E">'+
//'<w:pPr><w:rPr><w:lang w:val="en-US"/></w:rPr></w:pPr>'+
//'<w:r><w:rPr><w:lang w:val="en-US"/></w:rPr>';
  Result := Result + txt; //<w:t>Bogi</w:t>

//  Result := Result + '</w:r><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p>'+
  Result := Result +'<w:sectPr w:rsidR="00881B0E" w:rsidRPr="005E430E">'+
'<w:pgSz w:w="11906" w:h="16838"/>'+
'<w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>'+
'<w:cols w:space="708"/><w:docGrid w:linePitch="360"/>'+
'</w:sectPr>';
 Result := Result + '</w:body></w:document>';
end;



//   word/fontTable.xml
function getfonttablexml(doc:BTDocxWriter):ansistring;
begin
   Result :=

'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+#13#10+
'<w:fonts xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '+
'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '+
'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" '+
'mc:Ignorable="w14 w15 w16se">'+
'<w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/><w:charset w:val="CC"/><w:family w:val="swiss"/>'+
'<w:pitch w:val="variable"/><w:sig w:usb0="E4002EFF" w:usb1="C000247B" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/></w:font>'+
'<w:font w:name="Times New Roman"><w:panose1 w:val="02020603050405020304"/><w:charset w:val="CC"/><w:family w:val="roman"/>'+
'<w:pitch w:val="variable"/><w:sig w:usb0="E0002EFF" w:usb1="C000785B" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/></w:font>'+
'<w:font w:name="Calibri Light"><w:panose1 w:val="020F0302020204030204"/><w:charset w:val="CC"/><w:family w:val="swiss"/>'+
'<w:pitch w:val="variable"/><w:sig w:usb0="E4002EFF" w:usb1="C000247B" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/></w:font>'+
'</w:fonts>';
end;


//   word/settings.xml
const xmlsettings:ansistring =

'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+#13#10+
'<w:settings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" '+
'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '+
'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" '+
'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '+
'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" '+
'xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main" mc:Ignorable="w14 w15 w16se">'+
'<w:zoom w:percent="100"/><w:proofState w:spelling="clean" w:grammar="clean"/><w:defaultTabStop w:val="708"/>'+
'<w:hyphenationZone w:val="425"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat>'+
'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'+
'<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'+
'<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'+
'<w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'+
'<w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'+
'</w:compat><w:rsids><w:rsidRoot w:val="005E430E"/><w:rsid w:val="005E430E"/>'+
'<w:rsid w:val="00881B0E"/></w:rsids><m:mathPr><m:mathFont m:val="Cambria Math"/>'+
'<m:brkBin m:val="before"/><m:brkBinSub m:val="--"/><m:smallFrac m:val="0"/>'+
'<m:dispDef/><m:lMargin m:val="0"/><m:rMargin m:val="0"/><m:defJc m:val="centerGroup"/>'+
'<m:wrapIndent m:val="1440"/><m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/>'+
'</m:mathPr><w:themeFontLang w:val="bg-BG"/>'+
'<w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" '+
'w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>'+
'<w:shapeDefaults><o:shapedefaults v:ext="edit" spidmax="1026"/><o:shapelayout v:ext="edit">'+
'<o:idmap v:ext="edit" data="1"/></o:shapelayout></w:shapeDefaults><w:decimalSymbol w:val="."/>'+
'<w:listSeparator w:val=";"/><w14:docId w14:val="45737387"/><w15:chartTrackingRefBased/>'+
'<w15:docId w15:val="{54D9D889-9279-4D2B-AB81-FAC0E4C730D3}"/></w:settings>';

const
   Bin_settings_len = 2619;
   Bin_settings_len_lzo = 1376;
   Bin_settings : array [0..1375] of byte = (
      $00,$50,$3C,$3F,$78,$6D,$6C,$20,$76,$65,$72,$73,$69,$6F,$6E,$3D,
      $22,$31,$2E,$30,$22,$20,$65,$6E,$63,$6F,$64,$69,$6E,$67,$3D,$22,
      $55,$54,$46,$2D,$38,$22,$20,$73,$74,$61,$6E,$64,$61,$6C,$6F,$6E,
      $65,$3D,$22,$79,$65,$73,$22,$3F,$3E,$0D,$0A,$3C,$77,$3A,$73,$65,
      $74,$74,$69,$6E,$67,$73,$20,$78,$6D,$6C,$6E,$73,$3A,$6D,$63,$3D,
      $22,$68,$74,$74,$70,$3A,$2F,$2F,$73,$63,$68,$65,$6D,$61,$73,$2E,
      $6F,$70,$65,$6E,$50,$03,$00,$14,$66,$6F,$72,$6D,$61,$74,$73,$2E,
      $6F,$72,$67,$2F,$6D,$61,$72,$6B,$75,$70,$2D,$63,$6F,$6D,$70,$61,
      $74,$69,$62,$69,$6C,$69,$74,$79,$2F,$32,$30,$30,$36,$22,$D8,$08,
      $04,$6F,$3D,$22,$75,$72,$6E,$3A,$C8,$08,$07,$2D,$6D,$69,$63,$72,
      $6F,$73,$6F,$66,$74,$68,$06,$04,$3A,$6F,$66,$66,$69,$63,$65,$D8,
      $00,$E5,$06,$72,$20,$03,$DC,$01,$A4,$07,$05,$44,$6F,$63,$75,$6D,
      $65,$6E,$74,$84,$0E,$0B,$2F,$72,$65,$6C,$61,$74,$69,$6F,$6E,$73,
      $68,$69,$70,$73,$F5,$09,$6D,$20,$17,$34,$01,$01,$6D,$61,$74,$68,
      $F1,$08,$76,$3A,$13,$03,$76,$6D,$6C,$FF,$04,$77,$31,$30,$3A,$A4,
      $00,$B0,$16,$02,$3A,$77,$6F,$72,$64,$27,$C4,$00,$20,$03,$78,$02,
      $60,$06,$04,$70,$72,$6F,$63,$65,$73,$73,$5E,$38,$6D,$6C,$B8,$1D,
      $01,$6D,$61,$69,$6E,$27,$1A,$01,$31,$34,$2F,$20,$01,$27,$A8,$05,
      $02,$2E,$63,$6F,$6D,$2F,$AD,$0F,$2F,$68,$09,$02,$2F,$32,$30,$31,
      $30,$84,$01,$2A,$E5,$02,$35,$20,$0D,$01,$01,$32,$2F,$03,$01,$36,
      $73,$65,$20,$0D,$09,$01,$35,$C8,$08,$03,$2F,$73,$79,$6D,$65,$78,
      $EA,$19,$73,$6C,$2F,$14,$01,$31,$DC,$09,$BC,$49,$03,$4C,$69,$62,
      $72,$61,$72,$A1,$4E,$2F,$BC,$21,$00,$23,$6D,$63,$3A,$49,$67,$6E,
      $6F,$72,$61,$62,$6C,$65,$3D,$22,$77,$31,$34,$20,$77,$31,$35,$20,
      $77,$31,$36,$73,$65,$22,$3E,$3C,$77,$3A,$7A,$6F,$6F,$6D,$20,$77,
      $3A,$70,$65,$72,$63,$65,$6E,$74,$3D,$22,$31,$30,$30,$22,$2F,$60,
      $03,$0F,$70,$72,$6F,$6F,$66,$53,$74,$61,$74,$65,$20,$77,$3A,$73,
      $70,$65,$6C,$6C,$4C,$2D,$0F,$3D,$22,$63,$6C,$65,$61,$6E,$22,$20,
      $77,$3A,$67,$72,$61,$6D,$6D,$61,$72,$E4,$02,$8C,$06,$00,$07,$64,
      $65,$66,$61,$75,$6C,$74,$54,$61,$62,$53,$74,$6F,$70,$20,$77,$3A,
      $76,$61,$6C,$3D,$22,$37,$30,$38,$A8,$0A,$03,$68,$79,$70,$68,$65,
      $6E,$8F,$52,$5A,$6F,$6E,$7C,$0A,$9F,$03,$34,$32,$35,$BC,$03,$00,
      $05,$63,$68,$61,$72,$61,$63,$74,$65,$72,$53,$70,$61,$63,$69,$6E,
      $67,$43,$6F,$6E,$74,$72,$6F,$6C,$FC,$08,$0A,$64,$6F,$4E,$6F,$74,
      $43,$6F,$6D,$70,$72,$65,$73,$73,$C4,$06,$9C,$6C,$78,$15,$A5,$6E,
      $53,$AC,$76,$44,$06,$03,$6E,$61,$6D,$65,$3D,$22,$B4,$02,$DC,$70,
      $01,$4D,$6F,$64,$65,$67,$17,$75,$72,$69,$2F,$48,$05,$37,$70,$08,
      $60,$06,$95,$14,$31,$F0,$14,$88,$0E,$2E,$A0,$01,$06,$6F,$76,$65,
      $72,$72,$69,$64,$65,$54,$7C,$2C,$00,$06,$53,$74,$79,$6C,$65,$46,
      $6F,$6E,$74,$53,$69,$7A,$65,$41,$6E,$64,$4A,$75,$73,$74,$69,$66,
      $69,$63,$84,$1F,$60,$0A,$20,$16,$04,$02,$2A,$CC,$03,$2E,$02,$02,
      $65,$6E,$64,$0F,$0D,$4F,$70,$65,$6E,$54,$79,$70,$65,$46,$65,$61,
      $74,$75,$72,$65,$73,$20,$36,$B0,$01,$84,$31,$0D,$46,$6C,$69,$70,
      $4D,$69,$72,$72,$6F,$72,$49,$6E,$64,$65,$6E,$74,$20,$38,$B0,$01,
      $00,$02,$69,$66,$66,$65,$72,$65,$6E,$74,$69,$61,$74,$65,$4D,$75,
      $6C,$74,$69,$72,$6F,$77,$9C,$2C,$03,$48,$65,$61,$64,$65,$72,$20,
      $1F,$DF,$01,$2F,$77,$3A,$B8,$44,$78,$48,$02,$72,$73,$69,$64,$73,
      $E0,$01,$01,$52,$6F,$6F,$74,$48,$49,$98,$3F,$05,$30,$30,$35,$45,
      $34,$33,$30,$45,$AC,$30,$78,$04,$20,$03,$64,$00,$01,$38,$38,$31,
      $42,$AF,$06,$2F,$77,$3A,$70,$06,$02,$73,$3E,$3C,$6D,$3A,$6A,$A9,
      $50,$72,$E4,$01,$62,$45,$20,$6D,$A8,$65,$06,$43,$61,$6D,$62,$72,
      $69,$61,$20,$4D,$71,$AD,$2F,$64,$04,$03,$62,$72,$6B,$42,$69,$6E,
      $FE,$03,$62,$65,$52,$CD,$65,$22,$29,$67,$00,$53,$75,$62,$F2,$03,
      $2D,$2D,$A0,$03,$06,$73,$6D,$61,$6C,$6C,$46,$72,$61,$63,$E0,$03,
      $84,$7A,$06,$6D,$3A,$64,$69,$73,$70,$44,$65,$66,$90,$07,$02,$6C,
      $4D,$61,$72,$67,$28,$5C,$01,$C5,$04,$72,$33,$54,$00,$01,$64,$65,
      $66,$4A,$27,$2C,$01,$74,$84,$04,$65,$72,$47,$72,$6F,$75,$70,$B4,
      $0D,$01,$77,$72,$61,$70,$A0,$40,$FB,$0D,$31,$34,$34,$DC,$09,$03,
      $69,$6E,$74,$4C,$69,$6D,$FC,$02,$01,$73,$75,$62,$53,$F4,$06,$01,
      $6E,$61,$72,$79,$29,$68,$00,$03,$75,$6E,$64,$4F,$76,$72,$61,$0A,
      $2F,$28,$78,$04,$03,$77,$3A,$74,$68,$65,$6D,$80,$69,$01,$4C,$61,
      $6E,$67,$FC,$2D,$02,$62,$67,$2D,$42,$47,$68,$05,$03,$77,$3A,$63,
      $6C,$72,$53,$6C,$F1,$03,$65,$4D,$61,$70,$70,$69,$88,$04,$08,$62,
      $67,$31,$3D,$22,$6C,$69,$67,$68,$74,$31,$71,$5F,$74,$54,$01,$01,
      $64,$61,$72,$6B,$93,$01,$62,$67,$32,$CD,$03,$32,$8C,$03,$54,$01,
      $6C,$03,$92,$01,$61,$63,$67,$1B,$31,$3D,$22,$C0,$01,$60,$04,$A9,
      $01,$32,$EC,$02,$29,$9D,$00,$33,$ED,$02,$33,$28,$9D,$00,$34,$ED,
      $02,$34,$28,$4D,$00,$35,$ED,$02,$35,$28,$4D,$00,$36,$ED,$02,$36,
      $6C,$02,$08,$68,$79,$70,$65,$72,$6C,$69,$6E,$6B,$3D,$22,$27,$28,
      $00,$7C,$02,$06,$66,$6F,$6C,$6C,$6F,$77,$65,$64,$48,$28,$7C,$00,
      $2F,$48,$00,$B0,$20,$03,$73,$68,$61,$70,$65,$44,$A4,$AF,$4A,$57,
      $6F,$3A,$80,$02,$C8,$B1,$00,$04,$73,$20,$76,$3A,$65,$78,$74,$3D,
      $22,$65,$64,$69,$74,$22,$20,$73,$70,$69,$64,$6D,$61,$78,$7A,$BC,
      $32,$36,$78,$07,$D4,$05,$03,$6C,$61,$79,$6F,$75,$74,$2B,$AC,$00,
      $01,$3E,$3C,$6F,$3A,$71,$05,$70,$2B,$54,$00,$05,$20,$64,$61,$74,
      $61,$3D,$22,$31,$6D,$07,$2F,$2B,$F2,$00,$3E,$3C,$40,$69,$88,$0F,
      $28,$2C,$02,$0A,$77,$3A,$64,$65,$63,$69,$6D,$61,$6C,$53,$79,$6D,
      $62,$28,$21,$17,$2E,$74,$07,$0C,$77,$3A,$6C,$69,$73,$74,$53,$65,
      $70,$61,$72,$61,$74,$6F,$72,$F5,$3B,$3B,$8C,$03,$09,$31,$34,$3A,
      $64,$6F,$63,$49,$64,$20,$77,$31,$34,$A4,$63,$05,$34,$35,$37,$33,
      $37,$33,$38,$37,$BA,$03,$35,$3A,$6C,$C5,$0E,$74,$54,$72,$61,$63,
      $6B,$69,$6E,$67,$52,$65,$66,$42,$61,$73,$65,$64,$72,$D0,$31,$35,
      $27,$E9,$00,$35,$A8,$07,$00,$14,$7B,$35,$34,$44,$39,$44,$38,$38,
      $39,$2D,$39,$32,$37,$39,$2D,$34,$44,$32,$42,$2D,$41,$42,$38,$31,
      $2D,$46,$41,$43,$30,$45,$34,$43,$37,$33,$30,$44,$33,$7D,$60,$0B,
      $09,$2F,$77,$3A,$73,$65,$74,$74,$69,$6E,$67,$73,$3E,$11,$00,$00
         );




(*
//   word/styles.xml

'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
'<w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"  '
'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
'xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" '
'mc:Ignorable="w14 w15 w16se">'
'<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>'
'<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="bg-BG" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr>'
'</w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
'</w:docDefaults><w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="0" w:defUnhideWhenUsed="0" w:defQFormat="0" w:count="371">'
'<w:lsdException w:name="Normal" w:uiPriority="0" w:qFormat="1"/>'
'<w:lsdException w:name="heading 1" w:uiPriority="9" w:qFormat="1"/>'
'<w:lsdException w:name="heading 2" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>'
'<w:lsdException w:name="heading 3" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>'
'<w:lsdException w:name="heading 4" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>'
'<w:lsdException w:name="heading 5" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>'
'<w:lsdException w:name="heading 6" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>'
'<w:lsdException w:name="heading 7" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>'
'<w:lsdException w:name="heading 8" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>'
'<w:lsdException w:name="heading 9" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>'
'<w:lsdException w:name="index 1" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="index 2" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="index 3" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="index 4" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="index 5" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="index 6" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="index 7" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="index 8" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="index 9" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="toc 1" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="toc 2" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="toc 3" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="toc 4" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="toc 5" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="toc 6" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="toc 7" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="toc 8" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="toc 9" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Normal Indent" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="footnote text" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="annotation text" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="header" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="footer" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="index heading" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="caption" w:semiHidden="1" w:uiPriority="35" w:unhideWhenUsed="1" w:qFormat="1"/>'
'<w:lsdException w:name="table of figures" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="envelope address" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="envelope return" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="footnote reference" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="annotation reference" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="line number" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="page number" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="endnote reference" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="endnote text" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="table of authorities" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="macro" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="toa heading" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="List" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="List Bullet" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="List Number" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="List 2" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="List 3" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="List 4" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="List 5" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="List Bullet 2" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="List Bullet 3" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="List Bullet 4" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="List Bullet 5" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="List Number 2" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="List Number 3" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="List Number 4" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="List Number 5" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Title" w:uiPriority="10" w:qFormat="1"/>'
'<w:lsdException w:name="Closing" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Signature" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Default Paragraph Font" w:semiHidden="1" w:uiPriority="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Body Text" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Body Text Indent" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="List Continue" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="List Continue 2" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="List Continue 3" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="List Continue 4" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="List Continue 5" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Message Header" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Subtitle" w:uiPriority="11" w:qFormat="1"/>'
'<w:lsdException w:name="Salutation" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Date" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Body Text First Indent" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Body Text First Indent 2" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Note Heading" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Body Text 2" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Body Text 3" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Body Text Indent 2" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Body Text Indent 3" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Block Text" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Hyperlink" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="FollowedHyperlink" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Strong" w:uiPriority="22" w:qFormat="1"/>'
'<w:lsdException w:name="Emphasis" w:uiPriority="20" w:qFormat="1"/>'
'<w:lsdException w:name="Document Map" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Plain Text" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="E-mail Signature" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="HTML Top of Form" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="HTML Bottom of Form" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Normal (Web)" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="HTML Acronym" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="HTML Address" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="HTML Cite" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="HTML Code" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="HTML Definition" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="HTML Keyboard" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="HTML Preformatted" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="HTML Sample" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="HTML Typewriter" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="HTML Variable" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Normal Table" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="annotation subject" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="No List" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Outline List 1" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Outline List 2" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Outline List 3" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Simple 1" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Simple 2" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Simple 3" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Classic 1" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Classic 2" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Classic 3" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Classic 4" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Colorful 1" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Colorful 2" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Colorful 3" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Columns 1" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Columns 2" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Columns 3" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Columns 4" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Columns 5" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Grid 1" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Grid 2" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Grid 3" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Grid 4" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Grid 5" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Grid 6" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Grid 7" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Grid 8" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table List 1" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table List 2" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table List 3" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table List 4" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table List 5" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table List 6" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table List 7" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table List 8" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table 3D effects 1" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table 3D effects 2" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table 3D effects 3" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Contemporary" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Elegant" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Professional" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Subtle 1" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Subtle 2" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Web 1" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Web 2" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Web 3" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Balloon Text" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Table Grid" w:uiPriority="39"/>'
'<w:lsdException w:name="Table Theme" w:semiHidden="1" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="Placeholder Text" w:semiHidden="1"/>'
'<w:lsdException w:name="No Spacing" w:uiPriority="1" w:qFormat="1"/>'
'<w:lsdException w:name="Light Shading" w:uiPriority="60"/>'
'<w:lsdException w:name="Light List" w:uiPriority="61"/>'
'<w:lsdException w:name="Light Grid" w:uiPriority="62"/>'
'<w:lsdException w:name="Medium Shading 1" w:uiPriority="63"/>'
'<w:lsdException w:name="Medium Shading 2" w:uiPriority="64"/>'
'<w:lsdException w:name="Medium List 1" w:uiPriority="65"/>'
'<w:lsdException w:name="Medium List 2" w:uiPriority="66"/>'
'<w:lsdException w:name="Medium Grid 1" w:uiPriority="67"/>'
'<w:lsdException w:name="Medium Grid 2" w:uiPriority="68"/>'
'<w:lsdException w:name="Medium Grid 3" w:uiPriority="69"/>'
'<w:lsdException w:name="Dark List" w:uiPriority="70"/>'
'<w:lsdException w:name="Colorful Shading" w:uiPriority="71"/>'
'<w:lsdException w:name="Colorful List" w:uiPriority="72"/>'
'<w:lsdException w:name="Colorful Grid" w:uiPriority="73"/>'
'<w:lsdException w:name="Light Shading Accent 1" w:uiPriority="60"/>'
'<w:lsdException w:name="Light List Accent 1" w:uiPriority="61"/>'
'<w:lsdException w:name="Light Grid Accent 1" w:uiPriority="62"/>'
'<w:lsdException w:name="Medium Shading 1 Accent 1" w:uiPriority="63"/>'
'<w:lsdException w:name="Medium Shading 2 Accent 1" w:uiPriority="64"/>'
'<w:lsdException w:name="Medium List 1 Accent 1" w:uiPriority="65"/>'
'<w:lsdException w:name="Revision" w:semiHidden="1"/>
'<w:lsdException w:name="List Paragraph" w:uiPriority="34" w:qFormat="1"/>'
'<w:lsdException w:name="Quote" w:uiPriority="29" w:qFormat="1"/>'
'<w:lsdException w:name="Intense Quote" w:uiPriority="30" w:qFormat="1"/>'
'<w:lsdException w:name="Medium List 2 Accent 1" w:uiPriority="66"/>'
'<w:lsdException w:name="Medium Grid 1 Accent 1" w:uiPriority="67"/>'
'<w:lsdException w:name="Medium Grid 2 Accent 1" w:uiPriority="68"/>'
'<w:lsdException w:name="Medium Grid 3 Accent 1" w:uiPriority="69"/>'
'<w:lsdException w:name="Dark List Accent 1" w:uiPriority="70"/>'
'<w:lsdException w:name="Colorful Shading Accent 1" w:uiPriority="71"/>'
'<w:lsdException w:name="Colorful List Accent 1" w:uiPriority="72"/>'
'<w:lsdException w:name="Colorful Grid Accent 1" w:uiPriority="73"/>'
'<w:lsdException w:name="Light Shading Accent 2" w:uiPriority="60"/>'
'<w:lsdException w:name="Light List Accent 2" w:uiPriority="61"/>'
'<w:lsdException w:name="Light Grid Accent 2" w:uiPriority="62"/>'
'<w:lsdException w:name="Medium Shading 1 Accent 2" w:uiPriority="63"/>'
'<w:lsdException w:name="Medium Shading 2 Accent 2" w:uiPriority="64"/>'
'<w:lsdException w:name="Medium List 1 Accent 2" w:uiPriority="65"/>'
'<w:lsdException w:name="Medium List 2 Accent 2" w:uiPriority="66"/>'
'<w:lsdException w:name="Medium Grid 1 Accent 2" w:uiPriority="67"/>'
'<w:lsdException w:name="Medium Grid 2 Accent 2" w:uiPriority="68"/>'
'<w:lsdException w:name="Medium Grid 3 Accent 2" w:uiPriority="69"/>'
'<w:lsdException w:name="Dark List Accent 2" w:uiPriority="70"/>'
'<w:lsdException w:name="Colorful Shading Accent 2" w:uiPriority="71"/>'
'<w:lsdException w:name="Colorful List Accent 2" w:uiPriority="72"/>'
'<w:lsdException w:name="Colorful Grid Accent 2" w:uiPriority="73"/>'
'<w:lsdException w:name="Light Shading Accent 3" w:uiPriority="60"/>'
'<w:lsdException w:name="Light List Accent 3" w:uiPriority="61"/>'
'<w:lsdException w:name="Light Grid Accent 3" w:uiPriority="62"/>'
'<w:lsdException w:name="Medium Shading 1 Accent 3" w:uiPriority="63"/>'
'<w:lsdException w:name="Medium Shading 2 Accent 3" w:uiPriority="64"/>'
'<w:lsdException w:name="Medium List 1 Accent 3" w:uiPriority="65"/>'
'<w:lsdException w:name="Medium List 2 Accent 3" w:uiPriority="66"/>'
'<w:lsdException w:name="Medium Grid 1 Accent 3" w:uiPriority="67"/>'
'<w:lsdException w:name="Medium Grid 2 Accent 3" w:uiPriority="68"/>'
'<w:lsdException w:name="Medium Grid 3 Accent 3" w:uiPriority="69"/>'
'<w:lsdException w:name="Dark List Accent 3" w:uiPriority="70"/>'
'<w:lsdException w:name="Colorful Shading Accent 3" w:uiPriority="71"/>'
'<w:lsdException w:name="Colorful List Accent 3" w:uiPriority="72"/>'
'<w:lsdException w:name="Colorful Grid Accent 3" w:uiPriority="73"/>'
'<w:lsdException w:name="Light Shading Accent 4" w:uiPriority="60"/>'
'<w:lsdException w:name="Light List Accent 4" w:uiPriority="61"/>'
'<w:lsdException w:name="Light Grid Accent 4" w:uiPriority="62"/>'
'<w:lsdException w:name="Medium Shading 1 Accent 4" w:uiPriority="63"/>'
'<w:lsdException w:name="Medium Shading 2 Accent 4" w:uiPriority="64"/>'
'<w:lsdException w:name="Medium List 1 Accent 4" w:uiPriority="65"/>'
'<w:lsdException w:name="Medium List 2 Accent 4" w:uiPriority="66"/>'
'<w:lsdException w:name="Medium Grid 1 Accent 4" w:uiPriority="67"/>'
'<w:lsdException w:name="Medium Grid 2 Accent 4" w:uiPriority="68"/>'
'<w:lsdException w:name="Medium Grid 3 Accent 4" w:uiPriority="69"/>'
'<w:lsdException w:name="Dark List Accent 4" w:uiPriority="70"/>'
'<w:lsdException w:name="Colorful Shading Accent 4" w:uiPriority="71"/>'
'<w:lsdException w:name="Colorful List Accent 4" w:uiPriority="72"/>'
'<w:lsdException w:name="Colorful Grid Accent 4" w:uiPriority="73"/>'
'<w:lsdException w:name="Light Shading Accent 5" w:uiPriority="60"/>'
'<w:lsdException w:name="Light List Accent 5" w:uiPriority="61"/>'
'<w:lsdException w:name="Light Grid Accent 5" w:uiPriority="62"/>'
'<w:lsdException w:name="Medium Shading 1 Accent 5" w:uiPriority="63"/>'
'<w:lsdException w:name="Medium Shading 2 Accent 5" w:uiPriority="64"/>'
'<w:lsdException w:name="Medium List 1 Accent 5" w:uiPriority="65"/>'
'<w:lsdException w:name="Medium List 2 Accent 5" w:uiPriority="66"/>'
'<w:lsdException w:name="Medium Grid 1 Accent 5" w:uiPriority="67"/>'
'<w:lsdException w:name="Medium Grid 2 Accent 5" w:uiPriority="68"/>'
'<w:lsdException w:name="Medium Grid 3 Accent 5" w:uiPriority="69"/>'
'<w:lsdException w:name="Dark List Accent 5" w:uiPriority="70"/>'
'<w:lsdException w:name="Colorful Shading Accent 5" w:uiPriority="71"/>'
'<w:lsdException w:name="Colorful List Accent 5" w:uiPriority="72"/>'
'<w:lsdException w:name="Colorful Grid Accent 5" w:uiPriority="73"/>'
'<w:lsdException w:name="Light Shading Accent 6" w:uiPriority="60"/>'
'<w:lsdException w:name="Light List Accent 6" w:uiPriority="61"/>'
'<w:lsdException w:name="Light Grid Accent 6" w:uiPriority="62"/>'
'<w:lsdException w:name="Medium Shading 1 Accent 6" w:uiPriority="63"/>'
'<w:lsdException w:name="Medium Shading 2 Accent 6" w:uiPriority="64"/>'
'<w:lsdException w:name="Medium List 1 Accent 6" w:uiPriority="65"/>'
'<w:lsdException w:name="Medium List 2 Accent 6" w:uiPriority="66"/>'
'<w:lsdException w:name="Medium Grid 1 Accent 6" w:uiPriority="67"/>'
'<w:lsdException w:name="Medium Grid 2 Accent 6" w:uiPriority="68"/>'
'<w:lsdException w:name="Medium Grid 3 Accent 6" w:uiPriority="69"/>'
'<w:lsdException w:name="Dark List Accent 6" w:uiPriority="70"/>'
'<w:lsdException w:name="Colorful Shading Accent 6" w:uiPriority="71"/>'
'<w:lsdException w:name="Colorful List Accent 6" w:uiPriority="72"/>'
'<w:lsdException w:name="Colorful Grid Accent 6" w:uiPriority="73"/>'
'<w:lsdException w:name="Subtle Emphasis" w:uiPriority="19" w:qFormat="1"/>'
'<w:lsdException w:name="Intense Emphasis" w:uiPriority="21" w:qFormat="1"/>'
'<w:lsdException w:name="Subtle Reference" w:uiPriority="31" w:qFormat="1"/>'
'<w:lsdException w:name="Intense Reference" w:uiPriority="32" w:qFormat="1"/>'
'<w:lsdException w:name="Book Title" w:uiPriority="33" w:qFormat="1"/>'
'<w:lsdException w:name="Bibliography" w:semiHidden="1" w:uiPriority="37" w:unhideWhenUsed="1"/>'
'<w:lsdException w:name="TOC Heading" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1" w:qFormat="1"/>'
'<w:lsdException w:name="Plain Table 1" w:uiPriority="41"/>'
'<w:lsdException w:name="Plain Table 2" w:uiPriority="42"/>'
'<w:lsdException w:name="Plain Table 3" w:uiPriority="43"/>'
'<w:lsdException w:name="Plain Table 4" w:uiPriority="44"/>'
'<w:lsdException w:name="Plain Table 5" w:uiPriority="45"/>'
'<w:lsdException w:name="Grid Table Light" w:uiPriority="40"/>'
'<w:lsdException w:name="Grid Table 1 Light" w:uiPriority="46"/>'
'<w:lsdException w:name="Grid Table 2" w:uiPriority="47"/>'
'<w:lsdException w:name="Grid Table 3" w:uiPriority="48"/>'
'<w:lsdException w:name="Grid Table 4" w:uiPriority="49"/>'
'<w:lsdException w:name="Grid Table 5 Dark" w:uiPriority="50"/>'
'<w:lsdException w:name="Grid Table 6 Colorful" w:uiPriority="51"/>'
'<w:lsdException w:name="Grid Table 7 Colorful" w:uiPriority="52"/>'
'<w:lsdException w:name="Grid Table 1 Light Accent 1" w:uiPriority="46"/>'
'<w:lsdException w:name="Grid Table 2 Accent 1" w:uiPriority="47"/>'
'<w:lsdException w:name="Grid Table 3 Accent 1" w:uiPriority="48"/>'
'<w:lsdException w:name="Grid Table 4 Accent 1" w:uiPriority="49"/>'
'<w:lsdException w:name="Grid Table 5 Dark Accent 1" w:uiPriority="50"/>'
'<w:lsdException w:name="Grid Table 6 Colorful Accent 1" w:uiPriority="51"/>'
'<w:lsdException w:name="Grid Table 7 Colorful Accent 1" w:uiPriority="52"/>'
'<w:lsdException w:name="Grid Table 1 Light Accent 2" w:uiPriority="46"/>'
'<w:lsdException w:name="Grid Table 2 Accent 2" w:uiPriority="47"/>'
'<w:lsdException w:name="Grid Table 3 Accent 2" w:uiPriority="48"/>'
'<w:lsdException w:name="Grid Table 4 Accent 2" w:uiPriority="49"/>'
'<w:lsdException w:name="Grid Table 5 Dark Accent 2" w:uiPriority="50"/>'
'<w:lsdException w:name="Grid Table 6 Colorful Accent 2" w:uiPriority="51"/>'
'<w:lsdException w:name="Grid Table 7 Colorful Accent 2" w:uiPriority="52"/>'
'<w:lsdException w:name="Grid Table 1 Light Accent 3" w:uiPriority="46"/>'
'<w:lsdException w:name="Grid Table 2 Accent 3" w:uiPriority="47"/>'
'<w:lsdException w:name="Grid Table 3 Accent 3" w:uiPriority="48"/>'
'<w:lsdException w:name="Grid Table 4 Accent 3" w:uiPriority="49"/>'
'<w:lsdException w:name="Grid Table 5 Dark Accent 3" w:uiPriority="50"/>'
'<w:lsdException w:name="Grid Table 6 Colorful Accent 3" w:uiPriority="51"/>'
'<w:lsdException w:name="Grid Table 7 Colorful Accent 3" w:uiPriority="52"/>'
'<w:lsdException w:name="Grid Table 1 Light Accent 4" w:uiPriority="46"/>'
'<w:lsdException w:name="Grid Table 2 Accent 4" w:uiPriority="47"/>'
'<w:lsdException w:name="Grid Table 3 Accent 4" w:uiPriority="48"/>'
'<w:lsdException w:name="Grid Table 4 Accent 4" w:uiPriority="49"/>'
'<w:lsdException w:name="Grid Table 5 Dark Accent 4" w:uiPriority="50"/>'
'<w:lsdException w:name="Grid Table 6 Colorful Accent 4" w:uiPriority="51"/>'
'<w:lsdException w:name="Grid Table 7 Colorful Accent 4" w:uiPriority="52"/>'
'<w:lsdException w:name="Grid Table 1 Light Accent 5" w:uiPriority="46"/>'
'<w:lsdException w:name="Grid Table 2 Accent 5" w:uiPriority="47"/>'
'<w:lsdException w:name="Grid Table 3 Accent 5" w:uiPriority="48"/>'
'<w:lsdException w:name="Grid Table 4 Accent 5" w:uiPriority="49"/>'
'<w:lsdException w:name="Grid Table 5 Dark Accent 5" w:uiPriority="50"/>'
'<w:lsdException w:name="Grid Table 6 Colorful Accent 5" w:uiPriority="51"/>'
'<w:lsdException w:name="Grid Table 7 Colorful Accent 5" w:uiPriority="52"/>'
'<w:lsdException w:name="Grid Table 1 Light Accent 6" w:uiPriority="46"/>'
'<w:lsdException w:name="Grid Table 2 Accent 6" w:uiPriority="47"/>'
'<w:lsdException w:name="Grid Table 3 Accent 6" w:uiPriority="48"/>'
'<w:lsdException w:name="Grid Table 4 Accent 6" w:uiPriority="49"/>'
'<w:lsdException w:name="Grid Table 5 Dark Accent 6" w:uiPriority="50"/>'
'<w:lsdException w:name="Grid Table 6 Colorful Accent 6" w:uiPriority="51"/>'
'<w:lsdException w:name="Grid Table 7 Colorful Accent 6" w:uiPriority="52"/>'
'<w:lsdException w:name="List Table 1 Light" w:uiPriority="46"/>'
'<w:lsdException w:name="List Table 2" w:uiPriority="47"/>'
'<w:lsdException w:name="List Table 3" w:uiPriority="48"/>'
'<w:lsdException w:name="List Table 4" w:uiPriority="49"/>'
'<w:lsdException w:name="List Table 5 Dark" w:uiPriority="50"/>'
'<w:lsdException w:name="List Table 6 Colorful" w:uiPriority="51"/>'
'<w:lsdException w:name="List Table 7 Colorful" w:uiPriority="52"/>'
'<w:lsdException w:name="List Table 1 Light Accent 1" w:uiPriority="46"/>'
'<w:lsdException w:name="List Table 2 Accent 1" w:uiPriority="47"/>'
'<w:lsdException w:name="List Table 3 Accent 1" w:uiPriority="48"/>'
'<w:lsdException w:name="List Table 4 Accent 1" w:uiPriority="49"/>'
'<w:lsdException w:name="List Table 5 Dark Accent 1" w:uiPriority="50"/>'
'<w:lsdException w:name="List Table 6 Colorful Accent 1" w:uiPriority="51"/>'
'<w:lsdException w:name="List Table 7 Colorful Accent 1" w:uiPriority="52"/>'
'<w:lsdException w:name="List Table 1 Light Accent 2" w:uiPriority="46"/>'
'<w:lsdException w:name="List Table 2 Accent 2" w:uiPriority="47"/>'
'<w:lsdException w:name="List Table 3 Accent 2" w:uiPriority="48"/>'
'<w:lsdException w:name="List Table 4 Accent 2" w:uiPriority="49"/>'
'<w:lsdException w:name="List Table 5 Dark Accent 2" w:uiPriority="50"/>'
'<w:lsdException w:name="List Table 6 Colorful Accent 2" w:uiPriority="51"/>'
'<w:lsdException w:name="List Table 7 Colorful Accent 2" w:uiPriority="52"/>'
'<w:lsdException w:name="List Table 1 Light Accent 3" w:uiPriority="46"/>'
'<w:lsdException w:name="List Table 2 Accent 3" w:uiPriority="47"/>'
'<w:lsdException w:name="List Table 3 Accent 3" w:uiPriority="48"/>'
'<w:lsdException w:name="List Table 4 Accent 3" w:uiPriority="49"/>'
'<w:lsdException w:name="List Table 5 Dark Accent 3" w:uiPriority="50"/>'
'<w:lsdException w:name="List Table 6 Colorful Accent 3" w:uiPriority="51"/>'
'<w:lsdException w:name="List Table 7 Colorful Accent 3" w:uiPriority="52"/>'
'<w:lsdException w:name="List Table 1 Light Accent 4" w:uiPriority="46"/>'
'<w:lsdException w:name="List Table 2 Accent 4" w:uiPriority="47"/>'
'<w:lsdException w:name="List Table 3 Accent 4" w:uiPriority="48"/>'
'<w:lsdException w:name="List Table 4 Accent 4" w:uiPriority="49"/>'
'<w:lsdException w:name="List Table 5 Dark Accent 4" w:uiPriority="50"/>'
'<w:lsdException w:name="List Table 6 Colorful Accent 4" w:uiPriority="51"/>'
'<w:lsdException w:name="List Table 7 Colorful Accent 4" w:uiPriority="52"/>'
'<w:lsdException w:name="List Table 1 Light Accent 5" w:uiPriority="46"/>'
'<w:lsdException w:name="List Table 2 Accent 5" w:uiPriority="47"/>'
'<w:lsdException w:name="List Table 3 Accent 5" w:uiPriority="48"/>'
'<w:lsdException w:name="List Table 4 Accent 5" w:uiPriority="49"/>'
'<w:lsdException w:name="List Table 5 Dark Accent 5" w:uiPriority="50"/>'
'<w:lsdException w:name="List Table 6 Colorful Accent 5" w:uiPriority="51"/>'
'<w:lsdException w:name="List Table 7 Colorful Accent 5" w:uiPriority="52"/>'
'<w:lsdException w:name="List Table 1 Light Accent 6" w:uiPriority="46"/>'
'<w:lsdException w:name="List Table 2 Accent 6" w:uiPriority="47"/>'
'<w:lsdException w:name="List Table 3 Accent 6" w:uiPriority="48"/>'
'<w:lsdException w:name="List Table 4 Accent 6" w:uiPriority="49"/>'
'<w:lsdException w:name="List Table 5 Dark Accent 6" w:uiPriority="50"/>'
'<w:lsdException w:name="List Table 6 Colorful Accent 6" w:uiPriority="51"/>'
'<w:lsdException w:name="List Table 7 Colorful Accent 6" w:uiPriority="52"/>'
'</w:latentStyles>'
'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
'<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
'<w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
'<w:style w:type="table" w:default="1" w:styleId="TableNormal">'
'<w:name w:val="Normal Table"/><w:uiPriority w:val="99"/>'
'<w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>'
'<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
'<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
'<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/>'
'<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style></w:styles>'

*)
const
   Bin_styles_len = 28755;
   Bin_styles_len_lzo = 4436;
   Bin_styles : array [0..4435] of byte = (
      $00,$4E,$3C,$3F,$78,$6D,$6C,$20,$76,$65,$72,$73,$69,$6F,$6E,$3D,
      $22,$31,$2E,$30,$22,$20,$65,$6E,$63,$6F,$64,$69,$6E,$67,$3D,$22,
      $55,$54,$46,$2D,$38,$22,$20,$73,$74,$61,$6E,$64,$61,$6C,$6F,$6E,
      $65,$3D,$22,$79,$65,$73,$22,$3F,$3E,$0D,$0A,$3C,$77,$3A,$73,$74,
      $79,$6C,$65,$73,$20,$78,$6D,$6C,$6E,$73,$3A,$6D,$63,$3D,$22,$68,
      $74,$74,$70,$3A,$2F,$2F,$73,$63,$68,$65,$6D,$61,$73,$2E,$6F,$70,
      $65,$6E,$50,$03,$00,$14,$66,$6F,$72,$6D,$61,$74,$73,$2E,$6F,$72,
      $67,$2F,$6D,$61,$72,$6B,$75,$70,$2D,$63,$6F,$6D,$70,$61,$74,$69,
      $62,$69,$6C,$69,$74,$79,$2F,$32,$30,$30,$36,$22,$D9,$08,$72,$20,
      $03,$14,$01,$0B,$6F,$66,$66,$69,$63,$65,$44,$6F,$63,$75,$6D,$65,
      $6E,$74,$9C,$07,$0B,$2F,$72,$65,$6C,$61,$74,$69,$6F,$6E,$73,$68,
      $69,$70,$73,$F5,$09,$77,$20,$03,$34,$01,$08,$77,$6F,$72,$64,$70,
      $72,$6F,$63,$65,$73,$73,$52,$1E,$6D,$6C,$BC,$09,$01,$6D,$61,$69,
      $6E,$27,$1A,$01,$31,$34,$2F,$20,$01,$0B,$6D,$69,$63,$72,$6F,$73,
      $6F,$66,$74,$2E,$63,$6F,$6D,$2F,$A5,$12,$2F,$68,$09,$02,$2F,$32,
      $30,$31,$30,$86,$01,$6D,$6C,$28,$01,$01,$35,$20,$0D,$01,$01,$32,
      $2F,$03,$01,$36,$73,$65,$20,$0D,$09,$01,$35,$C8,$08,$00,$21,$2F,
      $73,$79,$6D,$65,$78,$22,$20,$6D,$63,$3A,$49,$67,$6E,$6F,$72,$61,
      $62,$6C,$65,$3D,$22,$77,$31,$34,$20,$77,$31,$35,$20,$77,$31,$36,
      $73,$65,$22,$3E,$3C,$77,$3A,$64,$6F,$63,$44,$65,$66,$61,$75,$6C,
      $74,$73,$7B,$01,$72,$50,$72,$D8,$01,$D4,$01,$98,$00,$00,$18,$46,
      $6F,$6E,$74,$73,$20,$77,$3A,$61,$73,$63,$69,$69,$54,$68,$65,$6D,
      $65,$3D,$22,$6D,$69,$6E,$6F,$72,$48,$41,$6E,$73,$69,$22,$20,$77,
      $3A,$65,$61,$73,$74,$41,$73,$69,$61,$33,$71,$00,$68,$74,$04,$33,
      $67,$00,$63,$73,$74,$29,$34,$01,$03,$42,$69,$64,$69,$22,$2F,$64,
      $0E,$09,$73,$7A,$20,$77,$3A,$76,$61,$6C,$3D,$22,$32,$32,$E5,$02,
      $43,$64,$10,$2B,$4C,$00,$01,$6C,$61,$6E,$67,$FC,$04,$02,$62,$67,
      $2D,$42,$47,$2A,$24,$02,$04,$3D,$22,$65,$6E,$2D,$55,$53,$68,$02,
      $08,$62,$69,$64,$69,$3D,$22,$61,$72,$2D,$53,$41,$70,$09,$05,$2F,
      $77,$3A,$72,$50,$72,$3E,$3C,$BC,$00,$29,$B9,$03,$70,$2B,$F1,$03,
      $70,$6C,$04,$04,$77,$3A,$73,$70,$61,$63,$69,$9C,$0C,$02,$61,$66,
      $74,$65,$72,$52,$65,$36,$30,$72,$0A,$6C,$69,$7F,$62,$32,$35,$39,
      $F2,$01,$52,$75,$64,$2C,$01,$61,$75,$74,$6F,$D0,$0C,$83,$08,$2F,
      $77,$3A,$2A,$59,$01,$2F,$2F,$C4,$05,$04,$6C,$61,$74,$65,$6E,$74,
      $53,$B0,$6A,$0C,$77,$3A,$64,$65,$66,$4C,$6F,$63,$6B,$65,$64,$53,
      $74,$61,$74,$4C,$3F,$8C,$0D,$0D,$64,$65,$66,$55,$49,$50,$72,$69,
      $6F,$72,$69,$74,$79,$3D,$22,$39,$8C,$0E,$0C,$64,$65,$66,$53,$65,
      $6D,$69,$48,$69,$64,$64,$65,$6E,$3D,$22,$27,$A0,$00,$0A,$6E,$68,
      $69,$64,$65,$57,$68,$65,$6E,$55,$73,$65,$64,$28,$5E,$00,$51,$46,
      $80,$72,$C0,$02,$07,$63,$6F,$75,$6E,$74,$3D,$22,$33,$37,$31,$94,
      $40,$05,$6C,$73,$64,$45,$78,$63,$65,$70,$68,$69,$46,$2A,$6E,$61,
      $65,$3C,$4E,$61,$06,$6C,$7A,$1B,$75,$69,$28,$F0,$01,$9D,$0C,$71,
      $F1,$09,$31,$6E,$1D,$77,$3A,$33,$FF,$00,$68,$65,$61,$7E,$89,$20,
      $31,$2E,$08,$01,$98,$17,$20,$0C,$09,$01,$32,$69,$08,$73,$29,$B4,
      $03,$34,$4D,$01,$75,$2D,$EC,$03,$94,$04,$20,$0C,$A1,$01,$33,$20,
      $47,$A1,$01,$34,$20,$47,$A1,$01,$35,$20,$47,$A1,$01,$36,$20,$47,
      $A1,$01,$37,$20,$47,$A1,$01,$38,$20,$47,$A0,$01,$80,$62,$20,$3B,
      $78,$0B,$02,$69,$6E,$64,$65,$78,$A0,$71,$30,$98,$01,$2F,$D0,$0C,
      $88,$B0,$33,$48,$0F,$BC,$08,$35,$34,$0E,$20,$10,$1C,$01,$35,$B0,
      $0D,$20,$10,$1C,$01,$35,$2C,$0D,$20,$10,$1C,$01,$35,$A8,$0C,$20,
      $10,$1C,$01,$35,$24,$0C,$20,$10,$1C,$01,$35,$A0,$0B,$20,$10,$1C,
      $01,$35,$1C,$0B,$20,$10,$1C,$01,$35,$98,$0A,$20,$0A,$1F,$01,$74,
      $6F,$63,$36,$14,$0A,$29,$89,$19,$33,$80,$0D,$30,$30,$17,$38,$5C,
      $0A,$7C,$0A,$35,$54,$0A,$20,$20,$5C,$01,$35,$94,$0A,$20,$20,$5C,
      $01,$35,$D4,$0A,$20,$20,$5C,$01,$35,$14,$0B,$20,$20,$5C,$01,$35,
      $54,$0B,$20,$20,$5C,$01,$35,$94,$0B,$20,$20,$5C,$01,$35,$D4,$0B,
      $20,$20,$5C,$01,$90,$53,$30,$14,$15,$20,$1C,$5C,$01,$24,$A8,$25,
      $04,$20,$49,$6E,$64,$65,$6E,$74,$34,$D4,$19,$20,$0A,$94,$0D,$09,
      $66,$6F,$6F,$74,$6E,$6F,$74,$65,$20,$74,$65,$78,$20,$21,$34,$01,
      $02,$61,$6E,$6E,$6F,$74,$23,$90,$35,$20,$25,$3C,$01,$22,$5A,$28,
      $65,$72,$20,$24,$90,$03,$20,$22,$18,$01,$A4,$DA,$6F,$12,$69,$6E,
      $67,$20,$20,$52,$02,$63,$61,$23,$08,$2D,$34,$1C,$01,$2A,$84,$08,
      $90,$74,$30,$87,$13,$20,$77,$3A,$20,$04,$B9,$2A,$74,$22,$38,$37,
      $08,$20,$6F,$66,$20,$66,$69,$67,$75,$72,$65,$73,$34,$C0,$01,$20,
      $0A,$C8,$08,$0C,$65,$6E,$76,$65,$6C,$6F,$70,$65,$20,$61,$64,$64,
      $72,$65,$73,$20,$2A,$40,$01,$03,$72,$65,$74,$75,$72,$6E,$20,$20,
      $80,$02,$27,$4C,$0B,$06,$72,$65,$66,$65,$72,$65,$6E,$63,$65,$20,
      $20,$48,$01,$29,$60,$0B,$20,$29,$50,$01,$22,$58,$38,$02,$20,$6E,
      $75,$6D,$62,$20,$22,$6C,$0A,$01,$70,$61,$67,$65,$20,$27,$2F,$01,
      $65,$6E,$64,$88,$82,$20,$29,$A4,$03,$E4,$0A,$20,$24,$90,$11,$27,
      $40,$0B,$01,$61,$75,$74,$68,$22,$99,$3B,$69,$20,$22,$50,$0B,$02,
      $6D,$61,$63,$72,$6F,$20,$20,$98,$08,$01,$74,$6F,$61,$20,$20,$27,
      $7F,$10,$4C,$69,$73,$20,$21,$40,$16,$70,$08,$03,$20,$42,$75,$6C,
      $6C,$65,$20,$26,$2D,$01,$4E,$20,$25,$B8,$0A,$9C,$12,$35,$14,$26,
      $20,$0A,$28,$12,$98,$08,$35,$D0,$25,$20,$0F,$18,$01,$35,$8C,$25,
      $20,$0F,$18,$01,$94,$B5,$30,$C8,$1F,$20,$0F,$18,$01,$AD,$36,$20,
      $20,$26,$88,$04,$D4,$09,$20,$26,$A4,$04,$D4,$09,$20,$26,$C0,$04,
      $D4,$09,$20,$26,$DC,$04,$BC,$53,$20,$27,$DC,$04,$D4,$09,$20,$26,
      $DC,$04,$D4,$09,$20,$26,$DC,$04,$D4,$09,$20,$21,$DC,$04,$01,$54,
      $69,$74,$6C,$84,$E4,$2A,$E9,$4E,$31,$20,$09,$EF,$4E,$43,$6C,$6F,
      $22,$58,$5C,$20,$20,$F0,$14,$05,$53,$69,$67,$6E,$61,$74,$75,$72,
      $8C,$11,$20,$1C,$04,$0D,$25,$9C,$56,$08,$20,$50,$61,$72,$61,$67,
      $72,$61,$70,$68,$20,$22,$54,$5A,$34,$80,$02,$29,$A4,$25,$23,$08,
      $50,$30,$A0,$25,$38,$28,$39,$03,$42,$6F,$64,$79,$20,$54,$20,$23,
      $E0,$2C,$27,$24,$01,$20,$27,$5C,$2F,$84,$A3,$04,$43,$6F,$6E,$74,
      $69,$6E,$75,$20,$21,$40,$05,$2B,$34,$01,$20,$27,$70,$0D,$F5,$13,
      $20,$20,$26,$78,$0D,$27,$3C,$01,$20,$26,$80,$0D,$27,$3C,$01,$20,
      $21,$88,$0D,$01,$4D,$65,$73,$73,$22,$80,$26,$01,$48,$65,$61,$64,
      $20,$22,$B0,$27,$01,$53,$75,$62,$74,$33,$D0,$0E,$84,$58,$20,$04,
      $54,$30,$01,$53,$61,$6C,$75,$24,$A4,$36,$34,$58,$0C,$20,$0A,$BF,
      $1A,$44,$61,$74,$20,$21,$84,$09,$28,$00,$0C,$02,$46,$69,$72,$73,
      $74,$20,$27,$18,$0C,$34,$58,$01,$20,$22,$05,$0B,$4E,$22,$CC,$3C,
      $78,$39,$20,$23,$28,$38,$28,$94,$02,$20,$21,$BC,$1F,$28,$2C,$01,
      $20,$21,$58,$0D,$28,$2C,$01,$24,$B4,$41,$20,$22,$DC,$04,$2F,$48,
      $01,$20,$22,$94,$02,$01,$6C,$6F,$63,$6B,$20,$25,$18,$16,$06,$48,
      $79,$70,$65,$72,$6C,$69,$6E,$6B,$20,$20,$50,$0C,$05,$46,$6F,$6C,
      $6C,$6F,$77,$65,$64,$20,$29,$44,$01,$01,$53,$74,$72,$6F,$24,$D8,
      $41,$2A,$9D,$1E,$32,$84,$46,$20,$04,$C8,$0F,$04,$45,$6D,$70,$68,
      $61,$73,$69,$23,$BC,$3E,$2B,$08,$01,$20,$09,$A8,$1F,$26,$40,$7D,
      $01,$20,$4D,$61,$70,$20,$20,$88,$04,$02,$50,$6C,$61,$69,$6E,$20,
      $25,$DC,$06,$04,$45,$2D,$6D,$61,$69,$6C,$20,$20,$29,$04,$21,$05,
      $48,$54,$4D,$4C,$20,$54,$6F,$70,$22,$DC,$44,$22,$FC,$72,$20,$20,
      $B0,$03,$80,$0A,$03,$42,$6F,$74,$74,$6F,$6D,$20,$28,$4C,$01,$25,
      $24,$50,$02,$28,$57,$65,$62,$29,$20,$25,$80,$02,$04,$41,$63,$72,
      $6F,$6E,$79,$6D,$20,$26,$30,$01,$20,$26,$78,$48,$8A,$27,$43,$69,
      $20,$22,$90,$19,$A6,$09,$6F,$64,$20,$21,$B8,$1A,$84,$09,$03,$44,
      $65,$66,$69,$6E,$69,$22,$1C,$7C,$20,$25,$C0,$04,$05,$4B,$65,$79,
      $62,$6F,$61,$72,$64,$20,$25,$37,$01,$50,$72,$65,$24,$5E,$8D,$74,
      $65,$20,$26,$44,$01,$01,$53,$61,$6D,$70,$24,$A4,$30,$20,$1C,$74,
      $2E,$8C,$27,$03,$54,$79,$70,$65,$77,$72,$50,$3A,$20,$21,$74,$57,
      $9C,$09,$01,$56,$61,$72,$69,$22,$A0,$53,$20,$20,$EC,$04,$D1,$6A,
      $54,$20,$24,$30,$01,$29,$D8,$50,$03,$73,$75,$62,$6A,$65,$63,$20,
      $21,$17,$46,$4E,$6F,$20,$22,$34,$2E,$20,$20,$9F,$03,$4F,$75,$74,
      $23,$FC,$51,$70,$09,$36,$9C,$6D,$20,$0A,$10,$28,$2B,$38,$01,$8C,
      $D2,$20,$1C,$8C,$08,$2B,$38,$01,$20,$21,$88,$20,$97,$39,$20,$53,
      $69,$7C,$57,$20,$22,$B0,$03,$2B,$38,$01,$20,$21,$B0,$03,$2B,$38,
      $01,$20,$27,$B0,$03,$03,$43,$6C,$61,$73,$73,$69,$37,$08,$75,$20,
      $0A,$68,$07,$B8,$13,$FC,$09,$20,$27,$B8,$03,$FC,$09,$20,$2F,$BC,
      $03,$20,$21,$F0,$37,$DC,$1D,$04,$6F,$6C,$6F,$72,$66,$75,$6C,$20,
      $28,$B9,$08,$43,$E0,$0A,$20,$28,$04,$05,$E0,$0A,$20,$28,$08,$05,
      $03,$6F,$6C,$75,$6D,$6E,$73,$20,$2B,$C4,$03,$9C,$09,$20,$2A,$C0,
      $03,$9C,$09,$20,$2F,$BC,$03,$20,$2A,$C8,$08,$9C,$13,$20,$21,$BC,
      $40,$A8,$50,$01,$47,$72,$69,$64,$20,$28,$30,$06,$90,$09,$20,$27,
      $24,$06,$90,$09,$20,$27,$18,$06,$90,$09,$20,$27,$0C,$06,$90,$09,
      $20,$2C,$00,$06,$35,$2C,$83,$20,$10,$00,$15,$84,$13,$35,$00,$83,
      $20,$15,$30,$01,$35,$D4,$82,$20,$10,$30,$01,$20,$26,$08,$20,$B0,
      $56,$90,$09,$20,$27,$9C,$09,$90,$09,$20,$27,$9C,$09,$90,$09,$20,
      $27,$9C,$09,$90,$09,$20,$27,$9C,$09,$90,$09,$20,$27,$9C,$09,$90,
      $09,$20,$27,$9C,$09,$90,$09,$20,$27,$9C,$09,$06,$33,$44,$20,$65,
      $66,$66,$65,$63,$74,$20,$29,$88,$19,$29,$48,$01,$20,$27,$CC,$09,
      $29,$48,$01,$20,$27,$E4,$09,$22,$24,$59,$05,$65,$6D,$70,$6F,$72,
      $61,$72,$79,$20,$20,$E0,$2E,$B8,$6C,$02,$45,$6C,$65,$67,$61,$20,
      $22,$FC,$8F,$B4,$09,$01,$50,$72,$6F,$66,$22,$73,$C3,$6F,$6E,$61,
      $23,$10,$B7,$20,$1C,$EC,$2E,$A8,$0A,$22,$8E,$5A,$6C,$65,$20,$28,
      $F8,$1A,$D8,$09,$20,$27,$93,$07,$57,$65,$62,$20,$28,$68,$02,$6C,
      $09,$20,$2B,$5C,$02,$20,$21,$D4,$09,$03,$42,$61,$6C,$6C,$6F,$6F,
      $20,$26,$6C,$4D,$B8,$39,$74,$D7,$69,$57,$75,$29,$BE,$6C,$33,$39,
      $39,$44,$BF,$B8,$06,$23,$24,$C7,$7C,$06,$20,$1C,$44,$09,$08,$50,
      $6C,$61,$63,$65,$68,$6F,$6C,$64,$65,$72,$35,$BC,$50,$38,$8C,$6E,
      $01,$4E,$6F,$20,$53,$24,$58,$C6,$78,$10,$2A,$F8,$02,$20,$09,$B0,
      $64,$05,$4C,$69,$67,$68,$74,$20,$53,$68,$23,$8C,$C2,$2E,$1A,$01,
      $36,$30,$39,$14,$04,$A4,$07,$74,$A8,$2F,$D8,$00,$3A,$38,$C4,$B8,
      $06,$32,$CD,$05,$36,$25,$8C,$CB,$33,$C8,$B5,$03,$4D,$65,$64,$69,
      $75,$6D,$E0,$15,$A0,$5C,$2A,$C6,$03,$36,$33,$39,$A8,$02,$2D,$F0,
      $00,$84,$5A,$2B,$F1,$00,$34,$20,$01,$F0,$00,$60,$1D,$31,$D9,$01,
      $35,$20,$06,$E4,$00,$30,$CD,$01,$36,$20,$01,$E4,$00,$74,$24,$31,
      $CD,$01,$37,$20,$06,$E4,$00,$30,$CD,$01,$38,$20,$06,$E4,$00,$9C,
      $74,$2B,$84,$04,$3A,$3C,$0C,$01,$44,$61,$72,$6B,$27,$E4,$46,$2A,
      $D5,$00,$37,$3A,$FC,$08,$27,$D0,$38,$35,$F1,$09,$37,$3A,$14,$09,
      $27,$F0,$00,$78,$32,$2E,$FD,$09,$37,$3A,$20,$09,$27,$E4,$00,$70,
      $2B,$2F,$E4,$00,$3A,$14,$09,$A4,$57,$D4,$15,$04,$20,$41,$63,$63,
      $65,$6E,$74,$31,$74,$06,$3A,$CC,$03,$A8,$08,$6C,$16,$38,$FC,$00,
      $3A,$D8,$03,$BC,$07,$64,$17,$38,$FC,$00,$3A,$F0,$03,$2D,$21,$0C,
      $31,$38,$14,$01,$3A,$20,$04,$2D,$15,$01,$32,$38,$14,$01,$20,$08,
      $5C,$0D,$38,$08,$01,$3A,$80,$0D,$01,$52,$65,$76,$69,$22,$E8,$E6,
      $78,$38,$2C,$14,$16,$38,$D0,$14,$90,$2F,$27,$F4,$84,$64,$07,$2A,
      $C5,$0A,$33,$23,$10,$2D,$20,$04,$CF,$69,$51,$75,$6F,$24,$DC,$5E,
      $2A,$FD,$00,$32,$23,$58,$B4,$20,$04,$FC,$00,$05,$49,$6E,$74,$65,
      $6E,$73,$65,$20,$33,$1D,$01,$33,$20,$09,$E0,$6A,$D4,$31,$9C,$1A,
      $39,$28,$06,$20,$08,$B8,$11,$38,$28,$06,$20,$08,$DC,$11,$38,$08,
      $01,$20,$08,$00,$12,$38,$08,$01,$20,$04,$24,$12,$37,$F8,$00,$20,
      $0C,$48,$12,$38,$14,$01,$3A,$90,$0E,$27,$84,$11,$88,$32,$36,$9C,
      $10,$20,$09,$90,$12,$38,$14,$02,$3A,$90,$0E,$A8,$85,$2D,$B4,$12,
      $30,$44,$18,$20,$0D,$B4,$12,$30,$FC,$00,$3A,$20,$04,$A8,$10,$2A,
      $B4,$12,$30,$FC,$00,$20,$13,$B4,$12,$30,$14,$01,$3A,$20,$04,$DC,
      $63,$E4,$21,$27,$88,$0C,$30,$14,$01,$20,$10,$B4,$12,$30,$08,$01,
      $3A,$B4,$12,$C0,$11,$94,$42,$39,$14,$02,$20,$10,$A0,$0E,$30,$14,
      $02,$20,$10,$A0,$0E,$30,$08,$01,$20,$10,$A0,$0E,$30,$08,$01,$20,
      $0C,$A0,$0E,$2F,$F8,$00,$20,$14,$A0,$0E,$30,$14,$01,$3A,$7C,$0A,
      $33,$A0,$0E,$30,$08,$01,$3A,$88,$0A,$27,$08,$01,$3B,$95,$0B,$37,
      $3A,$7C,$0A,$B4,$64,$F8,$53,$D8,$85,$30,$00,$26,$20,$0D,$A0,$0E,
      $30,$FC,$00,$3A,$20,$04,$A8,$10,$2A,$08,$03,$30,$FC,$00,$3A,$14,
      $04,$C4,$5B,$EC,$18,$27,$58,$21,$30,$14,$01,$3A,$20,$04,$2D,$14,
      $01,$27,$88,$0C,$30,$14,$01,$20,$10,$A0,$0E,$30,$08,$01,$20,$10,
      $A0,$0E,$30,$08,$01,$20,$10,$A0,$0E,$30,$08,$01,$20,$10,$A0,$0E,
      $30,$08,$01,$20,$10,$A0,$0E,$30,$08,$01,$20,$0C,$A0,$0E,$2F,$F8,
      $00,$20,$14,$A0,$0E,$30,$14,$01,$3A,$7C,$0A,$27,$94,$0D,$8C,$A7,
      $36,$88,$0C,$20,$11,$44,$1D,$30,$14,$02,$3A,$7C,$0A,$B4,$64,$F0,
      $5C,$D4,$10,$23,$BC,$28,$2A,$04,$28,$20,$06,$CC,$3C,$F0,$FA,$30,
      $FC,$00,$3A,$20,$04,$A8,$10,$2A,$A0,$0E,$30,$FC,$00,$20,$13,$A0,
      $0E,$30,$14,$01,$3A,$20,$04,$36,$A0,$0E,$30,$14,$01,$20,$10,$A0,
      $0E,$30,$08,$01,$20,$10,$A0,$0E,$30,$08,$01,$20,$10,$A0,$0E,$30,
      $08,$01,$20,$10,$A0,$0E,$30,$08,$01,$20,$10,$A0,$0E,$30,$08,$01,
      $20,$0C,$A0,$0E,$2F,$F8,$00,$20,$14,$A0,$0E,$30,$14,$01,$3A,$7C,
      $0A,$33,$A0,$0E,$30,$08,$01,$3A,$88,$0A,$27,$08,$01,$3B,$94,$0B,
      $20,$11,$44,$1D,$23,$40,$63,$20,$1A,$A0,$0E,$30,$FC,$00,$3A,$20,
      $04,$30,$A0,$0E,$30,$FC,$00,$3A,$14,$04,$2D,$88,$0D,$27,$44,$1D,
      $30,$14,$01,$20,$13,$A0,$0E,$30,$14,$01,$20,$10,$A0,$0E,$30,$08,
      $01,$20,$10,$A0,$0E,$30,$08,$01,$20,$10,$A0,$0E,$30,$08,$01,$20,
      $10,$A0,$0E,$30,$08,$01,$20,$10,$A0,$0E,$30,$08,$01,$20,$0C,$A0,
      $0E,$2F,$F8,$00,$20,$14,$A0,$0E,$30,$14,$01,$3A,$7C,$0A,$27,$94,
      $0D,$2A,$44,$1D,$30,$08,$01,$3A,$88,$0A,$27,$08,$01,$2A,$A0,$0E,
      $30,$08,$01,$3A,$7C,$0A,$B4,$64,$2D,$44,$1D,$23,$B0,$70,$20,$1A,
      $A0,$0E,$30,$FC,$00,$3A,$20,$04,$A8,$10,$2A,$08,$03,$30,$FC,$00,
      $3A,$14,$04,$36,$A0,$0E,$30,$14,$01,$3A,$20,$04,$2D,$14,$01,$27,
      $E8,$2B,$30,$14,$01,$20,$10,$A0,$0E,$30,$08,$01,$20,$10,$A0,$0E,
      $30,$08,$01,$20,$10,$A0,$0E,$30,$08,$01,$20,$10,$A0,$0E,$30,$08,
      $01,$20,$10,$A0,$0E,$30,$08,$01,$20,$0C,$A0,$0E,$2F,$F8,$00,$20,
      $14,$A0,$0E,$30,$14,$01,$3A,$7C,$0A,$27,$94,$0D,$2A,$A0,$0E,$30,
      $08,$01,$3A,$88,$0A,$27,$08,$01,$3B,$94,$0B,$3B,$44,$1D,$25,$88,
      $72,$36,$99,$BD,$31,$20,$11,$D4,$53,$36,$29,$01,$32,$20,$09,$E8,
      $6A,$D1,$12,$52,$2A,$24,$FC,$2A,$DD,$10,$33,$20,$09,$28,$01,$26,
      $2C,$56,$38,$2C,$01,$23,$58,$40,$20,$04,$5C,$57,$02,$42,$6F,$6F,
      $6B,$20,$33,$E1,$E1,$33,$23,$BC,$31,$20,$05,$10,$01,$02,$69,$62,
      $6C,$69,$6F,$23,$80,$DF,$35,$58,$7D,$2A,$B0,$72,$23,$48,$84,$20,
      $0B,$74,$DF,$01,$54,$4F,$43,$20,$3B,$84,$CF,$2A,$74,$01,$90,$3B,
      $30,$74,$01,$10,$1F,$94,$06,$25,$2C,$C4,$22,$D0,$B1,$30,$B5,$64,
      $34,$3A,$EC,$0B,$2A,$E4,$00,$9C,$2F,$2A,$75,$07,$34,$3A,$C8,$0B,
      $2A,$E4,$00,$90,$2E,$2B,$E4,$00,$3A,$24,$16,$2A,$E4,$00,$2F,$E1,
      $2A,$34,$3A,$F4,$15,$2A,$E4,$00,$2F,$19,$1C,$34,$3A,$D0,$15,$9C,
      $72,$24,$50,$79,$98,$D1,$2E,$51,$62,$34,$3A,$EC,$47,$29,$F2,$00,
      $31,$20,$34,$F8,$00,$3A,$B4,$16,$29,$F8,$00,$30,$88,$05,$3A,$8C,
      $16,$29,$E0,$00,$30,$84,$05,$3A,$64,$16,$29,$E0,$00,$30,$80,$05,
      $3A,$3C,$16,$29,$E2,$00,$35,$20,$22,$CC,$71,$2E,$9D,$04,$35,$20,
      $06,$9E,$04,$36,$20,$E8,$A2,$2F,$04,$01,$3A,$38,$0A,$29,$FD,$01,
      $37,$38,$04,$01,$3A,$58,$0A,$29,$04,$01,$CC,$35,$26,$18,$40,$88,
      $95,$2B,$90,$0A,$20,$07,$D0,$06,$38,$04,$01,$20,$07,$F4,$06,$38,
      $04,$01,$20,$07,$18,$07,$38,$04,$01,$20,$0C,$3C,$07,$37,$18,$01,
      $20,$11,$60,$07,$38,$28,$01,$20,$10,$84,$07,$38,$28,$01,$20,$15,
      $A8,$07,$30,$98,$0D,$20,$0F,$A8,$07,$30,$04,$01,$20,$0F,$A8,$07,
      $30,$04,$01,$20,$0F,$A8,$07,$30,$04,$01,$20,$14,$A8,$07,$2F,$18,
      $01,$20,$19,$A8,$07,$30,$28,$01,$20,$18,$A8,$07,$30,$28,$01,$20,
      $15,$A8,$07,$30,$60,$14,$20,$0F,$A8,$07,$30,$04,$01,$20,$0F,$A8,
      $07,$30,$04,$01,$20,$0F,$A8,$07,$30,$04,$01,$20,$14,$A8,$07,$2F,
      $18,$01,$20,$19,$A8,$07,$30,$28,$01,$20,$18,$A8,$07,$30,$28,$01,
      $20,$15,$A8,$07,$30,$28,$1B,$20,$0F,$A8,$07,$30,$04,$01,$20,$0F,
      $A8,$07,$30,$04,$01,$20,$0F,$A8,$07,$30,$04,$01,$20,$14,$A8,$07,
      $2F,$18,$01,$20,$19,$A8,$07,$30,$28,$01,$20,$18,$A8,$07,$30,$28,
      $01,$20,$15,$A8,$07,$30,$70,$27,$20,$0F,$A8,$07,$30,$04,$01,$20,
      $0F,$A8,$07,$30,$04,$01,$20,$0F,$A8,$07,$30,$04,$01,$20,$14,$A8,
      $07,$2F,$18,$01,$20,$19,$A8,$07,$30,$28,$01,$20,$18,$A8,$07,$30,
      $28,$01,$20,$15,$A8,$07,$2F,$A1,$3D,$34,$20,$0F,$A8,$07,$30,$04,
      $01,$20,$0F,$A8,$07,$30,$04,$01,$20,$0F,$A8,$07,$30,$04,$01,$20,
      $14,$A8,$07,$2F,$18,$01,$20,$19,$A8,$07,$30,$28,$01,$20,$18,$A8,
      $07,$30,$28,$01,$3A,$A8,$07,$23,$14,$45,$24,$A8,$35,$25,$04,$2E,
      $2E,$08,$30,$3B,$84,$07,$29,$F8,$00,$2F,$8D,$20,$34,$3A,$60,$07,
      $29,$E0,$00,$2F,$C5,$19,$34,$3A,$3C,$07,$29,$E0,$00,$2F,$FD,$12,
      $34,$3A,$18,$07,$29,$E0,$00,$20,$12,$B4,$34,$29,$F4,$00,$20,$16,
      $B4,$34,$29,$04,$01,$20,$16,$B4,$34,$29,$04,$01,$CC,$35,$37,$28,
      $2E,$20,$08,$D0,$06,$38,$04,$01,$20,$07,$F4,$06,$38,$04,$01,$20,
      $07,$18,$07,$38,$04,$01,$20,$0C,$3C,$07,$37,$18,$01,$3B,$58,$0E,
      $29,$50,$05,$28,$60,$07,$38,$28,$01,$3A,$58,$0E,$29,$28,$01,$28,
      $84,$07,$38,$28,$01,$20,$0D,$58,$0E,$FC,$08,$30,$98,$0D,$3A,$04,
      $16,$29,$48,$02,$27,$0C,$5D,$30,$04,$01,$20,$0F,$A8,$07,$30,$04,
      $01,$20,$0F,$A8,$07,$30,$04,$01,$20,$14,$A8,$07,$2F,$18,$01,$20,
      $19,$A8,$07,$30,$28,$01,$20,$18,$A8,$07,$30,$28,$01,$20,$15,$A8,
      $07,$30,$60,$14,$20,$0F,$A8,$07,$30,$04,$01,$20,$0F,$A8,$07,$30,
      $04,$01,$20,$0F,$A8,$07,$30,$04,$01,$20,$14,$A8,$07,$2F,$18,$01,
      $20,$19,$A8,$07,$30,$28,$01,$20,$18,$A8,$07,$30,$28,$01,$20,$15,
      $A8,$07,$30,$28,$1B,$20,$0F,$A8,$07,$30,$04,$01,$20,$0F,$A8,$07,
      $30,$04,$01,$20,$0F,$A8,$07,$30,$04,$01,$20,$14,$A8,$07,$2F,$18,
      $01,$20,$19,$A8,$07,$30,$28,$01,$20,$18,$A8,$07,$30,$28,$01,$20,
      $15,$A8,$07,$2F,$28,$2E,$20,$10,$AC,$1E,$30,$04,$01,$20,$0F,$A8,
      $07,$30,$04,$01,$20,$0F,$A8,$07,$30,$04,$01,$20,$14,$A8,$07,$2F,
      $18,$01,$20,$19,$A8,$07,$30,$28,$01,$20,$18,$A8,$07,$30,$28,$01,
      $20,$15,$A8,$07,$2F,$28,$2E,$20,$10,$A8,$07,$30,$04,$01,$20,$0F,
      $A8,$07,$30,$04,$01,$20,$0F,$A8,$07,$30,$04,$01,$20,$14,$A8,$07,
      $2F,$18,$01,$20,$19,$A8,$07,$30,$28,$01,$20,$18,$A8,$07,$30,$28,
      $01,$8B,$3D,$2F,$77,$3A,$10,$03,$04,$A6,$13,$F0,$A9,$12,$2C,$A6,
      $07,$20,$77,$3A,$74,$79,$70,$65,$3D,$22,$70,$16,$BC,$51,$22,$B9,
      $34,$64,$14,$AA,$AC,$3D,$22,$23,$D4,$62,$13,$24,$B4,$01,$49,$64,
      $3D,$22,$14,$7C,$21,$13,$24,$A5,$12,$E9,$A4,$20,$15,$EC,$AA,$D8,
      $02,$23,$44,$CE,$25,$EC,$73,$4D,$01,$2F,$15,$08,$B5,$30,$A8,$01,
      $04,$63,$68,$61,$72,$61,$63,$74,$14,$30,$49,$35,$A8,$01,$15,$18,
      $54,$27,$18,$CF,$13,$10,$54,$68,$09,$2A,$E0,$01,$D0,$04,$10,$07,
      $AC,$54,$52,$0F,$77,$3A,$28,$A8,$65,$F8,$14,$C8,$2A,$28,$18,$D1,
      $BC,$04,$10,$04,$C8,$01,$48,$02,$38,$DC,$02,$13,$F0,$6E,$39,$78,
      $04,$23,$D4,$39,$D0,$21,$2E,$A8,$02,$B9,$02,$20,$88,$04,$24,$88,
      $DE,$30,$81,$02,$39,$D8,$50,$3E,$84,$02,$02,$77,$3A,$74,$62,$6C,
      $14,$D4,$AE,$07,$74,$62,$6C,$49,$6E,$64,$20,$77,$3A,$77,$15,$C0,
      $AB,$B3,$39,$64,$78,$61,$B0,$0C,$07,$74,$62,$6C,$43,$65,$6C,$6C,
      $4D,$61,$72,$63,$12,$74,$6F,$70,$38,$A8,$00,$01,$6C,$65,$66,$74,
      $B6,$03,$31,$30,$23,$D0,$FC,$2D,$29,$01,$62,$14,$00,$38,$37,$A9,
      $01,$72,$23,$9C,$E9,$98,$03,$32,$00,$01,$40,$50,$2A,$30,$02,$B8,
      $01,$70,$18,$38,$BC,$05,$14,$78,$79,$15,$90,$4A,$35,$9E,$08,$4E,
      $6F,$22,$44,$30,$31,$48,$0A,$24,$B8,$E2,$5C,$39,$32,$28,$08,$20,
      $07,$A4,$05,$28,$6C,$02,$07,$2F,$77,$3A,$73,$74,$79,$6C,$65,$73,
      $3E,$11,$00,$00);


//   word/webSettings.xml
const xmlwebsettings:ansistring =

'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+#13#10+
'<w:webSettings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '+
'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '+
'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '+
'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '+
'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '+
'xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" '+
'mc:Ignorable="w14 w15 w16se"><w:optimizeForBrowser/><w:allowPNG/></w:webSettings>';

const
   Bin_webSettings_len = 576;
   Bin_webSettings_len_lzo = 357;
   Bin_webSettings : array [0..356] of byte = (
      $00,$53,$3C,$3F,$78,$6D,$6C,$20,$76,$65,$72,$73,$69,$6F,$6E,$3D,
      $22,$31,$2E,$30,$22,$20,$65,$6E,$63,$6F,$64,$69,$6E,$67,$3D,$22,
      $55,$54,$46,$2D,$38,$22,$20,$73,$74,$61,$6E,$64,$61,$6C,$6F,$6E,
      $65,$3D,$22,$79,$65,$73,$22,$3F,$3E,$0D,$0A,$3C,$77,$3A,$77,$65,
      $62,$53,$65,$74,$74,$69,$6E,$67,$73,$20,$78,$6D,$6C,$6E,$73,$3A,
      $6D,$63,$3D,$22,$68,$74,$74,$70,$3A,$2F,$2F,$73,$63,$68,$65,$6D,
      $61,$73,$2E,$6F,$70,$65,$6E,$50,$03,$00,$14,$66,$6F,$72,$6D,$61,
      $74,$73,$2E,$6F,$72,$67,$2F,$6D,$61,$72,$6B,$75,$70,$2D,$63,$6F,
      $6D,$70,$61,$74,$69,$62,$69,$6C,$69,$74,$79,$2F,$32,$30,$30,$36,
      $22,$D9,$08,$72,$20,$03,$14,$01,$0B,$6F,$66,$66,$69,$63,$65,$44,
      $6F,$63,$75,$6D,$65,$6E,$74,$9C,$07,$0B,$2F,$72,$65,$6C,$61,$74,
      $69,$6F,$6E,$73,$68,$69,$70,$73,$F5,$09,$77,$20,$03,$34,$01,$08,
      $77,$6F,$72,$64,$70,$72,$6F,$63,$65,$73,$73,$46,$1F,$6D,$6C,$BC,
      $09,$01,$6D,$61,$69,$6E,$27,$1A,$01,$31,$34,$2F,$20,$01,$0B,$6D,
      $69,$63,$72,$6F,$73,$6F,$66,$74,$2E,$63,$6F,$6D,$2F,$A5,$12,$2F,
      $68,$09,$02,$2F,$32,$30,$31,$30,$86,$01,$6D,$6C,$28,$01,$01,$35,
      $20,$0D,$01,$01,$32,$2F,$03,$01,$36,$73,$65,$20,$0D,$09,$01,$35,
      $C8,$08,$00,$13,$2F,$73,$79,$6D,$65,$78,$22,$20,$6D,$63,$3A,$49,
      $67,$6E,$6F,$72,$61,$62,$6C,$65,$3D,$22,$77,$31,$34,$20,$77,$31,
      $35,$20,$77,$31,$36,$73,$65,$22,$3E,$48,$3A,$00,$01,$6F,$70,$74,
      $69,$6D,$69,$7A,$65,$46,$6F,$72,$42,$72,$6F,$77,$73,$65,$72,$2F,
      $78,$02,$05,$61,$6C,$6C,$6F,$77,$50,$4E,$47,$51,$01,$2F,$2B,$DD,
      $07,$3E,$11,$00,$00   );


//   [Content_Types].xml
{
const xmlcontent:ansistring =

'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+#13#10+
'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'+
'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'+
'<Default Extension="xml" ContentType="application/xml"/>'+
'<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'+
'<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'+
'<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'+
'<Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>'+
'<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>'+
'<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'+
'<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'+
'<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>';
}
const
   Bin_Content_Types_len = 1312;
   Bin_Content_Types_len_lzo = 458;
   Bin_Content_Types : array [0..457] of byte = (
      $00,$48,$3C,$3F,$78,$6D,$6C,$20,$76,$65,$72,$73,$69,$6F,$6E,$3D,
      $22,$31,$2E,$30,$22,$20,$65,$6E,$63,$6F,$64,$69,$6E,$67,$3D,$22,
      $55,$54,$46,$2D,$38,$22,$20,$73,$74,$61,$6E,$64,$61,$6C,$6F,$6E,
      $65,$3D,$22,$79,$65,$73,$22,$3F,$3E,$0D,$0A,$3C,$54,$79,$70,$65,
      $73,$20,$78,$6D,$6C,$6E,$73,$3D,$22,$68,$74,$74,$70,$3A,$2F,$2F,
      $73,$63,$68,$65,$6D,$61,$73,$2E,$6F,$70,$65,$6E,$44,$03,$00,$10,
      $66,$6F,$72,$6D,$61,$74,$73,$2E,$6F,$72,$67,$2F,$70,$61,$63,$6B,
      $61,$67,$65,$2F,$32,$30,$30,$36,$2F,$63,$6F,$6E,$74,$65,$6E,$74,
      $2D,$74,$6C,$08,$0D,$22,$3E,$3C,$44,$65,$66,$61,$75,$6C,$74,$20,
      $45,$78,$74,$65,$6E,$A4,$11,$04,$72,$65,$6C,$73,$22,$20,$43,$A0,
      $05,$6C,$0D,$0E,$3D,$22,$61,$70,$70,$6C,$69,$63,$61,$74,$69,$6F,
      $6E,$2F,$76,$6E,$64,$2D,$95,$01,$2D,$C4,$0C,$01,$2E,$72,$65,$6C,
      $8C,$04,$08,$73,$68,$69,$70,$73,$2B,$78,$6D,$6C,$22,$2F,$33,$84,
      $01,$64,$03,$38,$80,$01,$74,$03,$00,$13,$2F,$3E,$3C,$4F,$76,$65,
      $72,$72,$69,$64,$65,$20,$50,$61,$72,$74,$4E,$61,$6D,$65,$3D,$22,
      $2F,$77,$6F,$72,$64,$2F,$64,$6F,$63,$75,$6D,$65,$6E,$74,$2E,$60,
      $05,$38,$18,$01,$31,$9C,$02,$03,$6F,$66,$66,$69,$63,$65,$27,$FC,
      $00,$74,$09,$04,$70,$72,$6F,$63,$65,$73,$73,$4F,$2E,$6D,$6C,$2E,
      $27,$64,$00,$01,$6D,$61,$69,$6E,$FC,$17,$37,$1C,$02,$03,$73,$74,
      $79,$6C,$65,$73,$20,$31,$14,$02,$BC,$0A,$20,$01,$F8,$01,$03,$65,
      $74,$74,$69,$6E,$67,$20,$33,$00,$02,$C4,$0B,$3F,$08,$02,$01,$77,
      $65,$62,$53,$CC,$05,$20,$31,$18,$04,$29,$70,$01,$3F,$20,$02,$06,
      $66,$6F,$6E,$74,$54,$61,$62,$6C,$65,$20,$31,$18,$02,$27,$68,$01,
      $3F,$10,$02,$03,$74,$68,$65,$6D,$65,$2F,$95,$00,$31,$20,$20,$1C,
      $02,$98,$08,$3A,$C8,$01,$09,$64,$6F,$63,$50,$72,$6F,$70,$73,$2F,
      $63,$6F,$72,$20,$12,$D8,$03,$E8,$76,$74,$07,$07,$2D,$70,$72,$6F,
      $70,$65,$72,$74,$69,$65,$27,$D0,$0E,$32,$D0,$0B,$27,$C7,$01,$61,
      $70,$70,$20,$20,$7D,$03,$65,$7F,$8D,$64,$65,$64,$31,$EC,$01,$04,
      $2F,$54,$79,$70,$65,$73,$3E,$11,$00,$00   );




function lzo_decompress(const CData; CSize: LongInt; var Data; var Size: LongInt): LongInt; cdecl;
asm
  DB $51
  DD $458B5653,$C558B08,$F08BD003,$33FC5589,$144D8BD2,$68A1189,$3C10558B,$331C7611,$83C88AC9
  DD $8346EFC1,$820F04F9,$1C9,$8846068A,$75494202,$3366EBF7,$460E8AC9,$F10F983,$8D83,$75C98500,$8107EB18
  DD $FFC1,$3E804600,$33F47400,$83068AC0,$C8030FC0,$83068B46,$28904C6,$4904C283,$F9832F74,$8B217204,$83028906
  DD $C68304C2,$4E98304,$7304F983,$76C985EE,$46068A14,$49420288,$9EBF775,$8846068A,$75494202,$8AC933F7
  DD $F983460E,$C12B7310,$828D02E9,$FFFFF7FF,$C933C12B,$C1460E8A,$C12B02E1,$8840088A,$88A420A,$420A8840
  DD $288008A,$113E942,$F9830000,$8B207240,$FF428DD9,$8302EBC1,$C32B07E3,$1E8ADB33,$3E3C146,$2B05E9C1
  DD $D9E949C3,$83000000,$2F7220F9,$851FE183,$EB1875C9,$FFC18107,$46000000,$74003E80,$8AC033F4,$1FC08306
  DD $F46C803,$FBC11EB7,$FF428D02,$C683C32B,$8369EB02,$457210F9,$D98BC28B,$C108E383,$C32B0BE3,$8507E183
  DD $EB1875C9,$FFC18107,$46000000,$74003E80,$8ADB33F4,$7C3831E,$F46CB03,$FBC11EB7,$83C32B02,$D03B02C6
  DD $9A840F,$2D0000,$EB000040,$2E9C11F,$2BFF428D,$8AC933C1,$E1C1460E,$8AC12B02,$A884008,$88008A42
  DD $51EB4202,$7206F983,$2BDA8B37,$4FB83D8,$188B2E7C,$8904C083,$4C2831A,$8B02E983,$831A8918,$C08304C2
  DD $4E98304,$7304F983,$76C985EE,$40188A20,$49421A88,$15EBF775,$8840188A,$188A421A,$421A8840,$8840188A
  DD $7549421A,$8AC933F7,$E183FE4E,$FC98503,$FFFE4284,$46068AFF,$49420288,$C933F775,$E9460E8A,$FFFFFECA
  DD $8B10552B,$10891445,$75FC753B,$EBC03304,$FFF8B80D,$753BFFFF,$830372FC,$5B5E04C0,$90C35D59
end;
procedure DecompressData(const InData: Pointer; InSize: LongInt; const OutData: Pointer; var OutSize: LongInt);
begin
  lzo_decompress(InData^, InSize, OutData^, OutSize);
end;




constructor BTDocxWriter.Create;
begin
   aParagraphs := nil;
   aCurParag := nil;
end;

destructor  BTDocxWriter.Destroy;
begin
   Reset;
   inherited;
end;

procedure   BTDocxWriter.Reset;
var p,o:PTParagraph;
begin
   if aParagraphs <> nil then
   begin
      p := PTParagraph( aParagraphs);
      repeat
         o := p;
         p := p.Next;
         Dispose(o);
      until p= nil;
   end;
   aParagraphs := nil;
   aCurParag := nil;
   aPagesCount := 0;
   aLineCount := 0;
   aParCount := 0;
   aWordCount := 0;
   aCharCount := 0;
end;

function    BTDocxWriter.Generate(const FileName:string):boolean;
var z:TZipWrite;
    s,d:ansistring;
//    i:longword;
    p:pointer;
    m:longint;
begin
//1    _rels/.rels
//2    docProps/app.xml
//3    docProps/core.xml
//4    word/_rels/document.xml.rels
//5    word/theme/theme1.xml
//6    word/document.xml
//7    word/fontTable.xml
//8    word/settings.xml
//9    word/styles.xml
//10   word/webSettings.xml
//11   [Content_Types].xml


   Result := True; //ok


   z := TZipWrite.Create(FileName);
   {1}
   SetLength(d,Bin_rels_len);
   p := @Bin_rels[0];
   m := Bin_rels_len;
   DecompressData(p,Bin_rels_len_lzo,@d[1],m);
   z.AddDeflated('_rels/.rels',@d[1],length(d));
   {2}
   s:= getappxml(self);
   z.AddDeflated('docProps/app.xml',@s[1],length(s));
   {3}
   s:= getcorexml(self);
   z.AddDeflated('docProps/core.xml',@s[1],length(s));
   {4}
   z.AddDeflated('word/_rels/document.xml.rels',@xmlrels[1],length(xmlrels));
   {5}
   SetLength(d,Bin_theme1_len);
   p := @Bin_theme1[0];
   m := Bin_theme1_len;
   DecompressData(p,Bin_theme1_len_lzo,@d[1],m);
   z.AddDeflated('word/theme/theme1.xml',@d[1],length(d));
   {6}
   s:= getdocumentxml(self);
   z.AddDeflated('word/document.xml',@s[1],length(s));
   {7}
   s:= getfonttablexml(self);
   z.AddDeflated('word/fontTable.xml',@s[1],length(s));
   {8}
   z.AddDeflated('word/settings.xml',@xmlsettings[1],length(xmlsettings));
   {9}
   SetLength(d,Bin_styles_len);
   p := @Bin_styles[0];
   m := Bin_styles_len;
   DecompressData(p,Bin_styles_len_lzo,@d[1],m);
   z.AddDeflated('word/styles.xml',@d[1],length(d));
   {10}
   z.AddDeflated('word/webSettings.xml',@xmlwebsettings[1],length(xmlwebsettings));
   {11}
   SetLength(d,Bin_content_types_len);
   p := @Bin_content_types[0];
   m := Bin_content_types_len;
   DecompressData(p,Bin_content_types_len_lzo,@d[1],m);
   z.AddDeflated('[Content_Types].xml',@d[1],length(d));
   z.Destroy;
end;


procedure   BTDocxWriter._NewParag;
var p,o:PTParagraph;
begin
   new(o);
   o.next := nil;
   o.Txt := '';
   inc(aParCount);

   p := aParagraphs;
   if p = nil then
   begin
      aParagraphs := o;
   end else begin
      while p.next <> nil do p := p.next;
      p.next := pointer(o);
   end;
   aCurParag := pointer(o);

end;

procedure   BTDocxWriter.AddNewPage(landscape:boolean = false);
begin
   _NewParag;
end;

procedure   BTDocxWriter.AddNewPageA4(landscape:boolean = false);
begin
   _NewParag;
end;

procedure   BTDocxWriter.AddText(const Txt:string);
var p:PTParagraph;
begin
   p := PTParagraph(aCurParag);
   if p <> nil then
   begin
      p.txt := p.txt + widestring(Txt);
   end;
end;

procedure   BTDocxWriter.AddNewLine; // create new paragraph.
begin
  _NewParag;
end;



end.
