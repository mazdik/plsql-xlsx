create or replace package PKG_XLSX_HELPER is

  -- Author  : MAZITOV.RRI
  -- Created : 24.03.2017 16:39:43
  -- Purpose : Работа с XLSX

  /* Печатать коллекции */
  bDEBUG boolean := false;

  /* Колонки */
  type tCOL is record(
    INDX  pls_integer,
    WIDTH number);
  type tbCOLS is table of tCOL;
  /* Объединенные ячейки */
  type tMERGE_CELL is record(
    CELL1_COL pls_integer,
    CELL1_ROW pls_integer,
    CELL2_COL pls_integer,
    CELL2_ROW pls_integer);
  type tbMERGE_CELLS is table of tMERGE_CELL;
  /* Таблица общих строк */
  type tbSTRINGS is table of varchar2(32767) index by pls_integer;
  /* Ячейки */
  type tCELL is record(
    ROW         pls_integer,
    COL         pls_integer,
    CELL_TYPE   varchar2(1),
    STRING_VAL  varchar2(4000),
    NUMBER_VAL  number,
    DATE_VAL    date,
    STYLE_INDEX pls_integer,
    FORMULA     varchar2(4000));
  type tbCELLS is table of tCELL;
  /* Таблица стилей */
  type tCELL_XF is record(
    NUM_FMT_ID pls_integer,
    FONT_ID    pls_integer,
    FILL_ID    pls_integer,
    BORDER_ID  pls_integer,
    HORIZONTAL varchar2(100),
    VERTICAL   varchar2(100),
    WRAPTEXT   boolean);
  type tbCELL_XFS is table of tCELL_XF index by pls_integer;
  /* Границы ячеек */
  type tBORDER is record(
    LEFT   varchar2(100),
    RIGHT  varchar2(100),
    TOP    varchar2(100),
    BOTTOM varchar2(100));
  type tbBORDERS is table of tBORDER index by pls_integer;
  /* Шрифты */
  type tFONT is record(
    NAME      varchar2(100),
    FAMILY    pls_integer default 2,
    FONTSIZE  pls_integer default 11,
    COLOR     varchar2(8),
    UNDERLINE boolean default false,
    ITALIC    boolean default false,
    BOLD      boolean default false);
  type tbFONTS is table of tFONT index by pls_integer;

  /* Загрузка из директории */
  function FILE2BLOB(sDIR in varchar2, sFILE_NAME in varchar2) return blob;
  /* Колонки */
  function GET_COLS(cSHEET_XML in clob) return tbCOLS;
  /* Объединенные ячейки */
  function GET_MERGE_CELLS(cSHEET_XML in clob) return tbMERGE_CELLS;
  /* Таблица общих строк */
  function GET_SHARED_STRINGS(cSHARED_STRINGS_XML in clob) return tbSTRINGS;
  /* Получить стили */
  function GET_CELL_STYLES(cSTYLES_XML in clob) return tbCELL_XFS;
  /* Границы ячеек  */
  function GET_BORDERS(cSTYLES_XML in clob) return tbBORDERS;
  /* Шрифты  */
  function GET_FONTS(cSTYLES_XML in clob) return tbFONTS;
  /* Данные листа */
  function GET_SHEET_DATA(cSHEET_XML in clob, cSHARED_STRINGS_XML in clob) return tbCELLS;
  /* Формат даты */
  function GET_DATE_1904(cWORKBOOK_XML in clob) return boolean;

  /* Сгенерировать код для AS_XLSX */
  procedure GEN_CODE_FOR_AS_XLSX(cSHEET_XML in clob, cSHARED_STRINGS_XML in clob, cSTYLES_XML in clob);
  /* Формирование шаблона отчета */
  procedure CREATE_REPORT_TEMPLATE(nREPORT_ID in number, iVERSION pls_integer default 1);
  /* Данные из файла */
  function GET_DATA(bXLSX in blob) return tbCELLS;

end PKG_XLSX_HELPER;
/
create or replace package body PKG_XLSX_HELPER is

  sXMLNS     varchar2(200) := 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
  bDATE_1904 boolean := true;

  /* Формат ячеек */
  type tNUM_FMT is record(
    NUM_FMT_ID  pls_integer,
    FORMAT_CODE varchar2(100));
  type tbNUM_FMTS is table of tNUM_FMT index by pls_integer;
  /* Строки */
  type tROW is record(
    INDX   pls_integer,
    HEIGHT number);
  type tbROWS is table of tROW;
  tabROWS tbROWS := tbROWS();
  /* Заливка */
  type tFILL is record(
    PATTERN_TYPE varchar2(30),
    FG_RGB       varchar2(8));
  type tbFILLS is table of tFILL index by pls_integer;

  /* Загрузка из clob */
  function CLOB2NODE(cCLOB in clob) return dbms_xmldom.domnode as
    xPARSER DBMS_XMLPARSER.parser;
    xDOC    DBMS_XMLDOM.domdocument;
    xROOT   DBMS_XMLDOM.domnode;
  begin
    if cCLOB is null then
      return null;
    end if;
    xPARSER := DBMS_XMLPARSER.newParser;
    begin
      DBMS_XMLPARSER.parseClob(xPARSER, cCLOB);
      xDOC := DBMS_XMLPARSER.getDocument(xPARSER);
      begin
        xROOT := DBMS_XMLDOM.makeNode(DBMS_XMLDOM.getDocumentElement(xDOC));
      exception
        when others then
          DBMS_XMLDOM.freeDocument(xDOC);
          raise;
      end;
    exception
      when others then
        DBMS_XMLPARSER.freeParser(xPARSER);
        raise;
    end;
    DBMS_XMLPARSER.freeParser(xPARSER);
    return xROOT;
  end CLOB2NODE;

  /* Загрузка из директории */
  function FILE2BLOB(sDIR in varchar2, sFILE_NAME in varchar2) return blob is
    bFILE_LOB  bfile;
    bFILE_BLOB blob;
  begin
    bFILE_LOB := BFILENAME(SDIR, SFILE_NAME);
    DBMS_LOB.OPEN(bFILE_LOB, DBMS_LOB.FILE_READONLY);
    DBMS_LOB.CREATETEMPORARY(bFILE_BLOB, TRUE, DBMS_LOB.CALL);
    DBMS_LOB.LOADFROMFILE(bFILE_BLOB, bFILE_LOB, DBMS_LOB.LOBMAXSIZE);
    DBMS_LOB.CLOSE(bFILE_LOB);
    return bFILE_BLOB;
  exception
    when others then
      if DBMS_LOB.ISOPEN(bFILE_LOB) = 1 then
        DBMS_LOB.CLOSE(bFILE_LOB);
      end if;
      if DBMS_LOB.ISTEMPORARY(bFILE_BLOB) = 1 then
        DBMS_LOB.FREETEMPORARY(bFILE_BLOB);
      end if;
      raise;
  end FILE2BLOB;

  /* Конвертация строки в число */
  function CONVERT_NUMBER(sVALUE in varchar2) return number as
  begin
    return to_number(sVALUE,
                     case when instr(sVALUE, 'E') = 0 then translate(sVALUE, '.012345678,-+', 'D999999999') else
                     translate(substr(sVALUE, 1, instr(sVALUE, 'E') - 1), '.012345678,-+', 'D999999999') || 'EEEE' end,
                     'NLS_NUMERIC_CHARACTERS=.,');
  end CONVERT_NUMBER;

  /* Конвертация строки в дату */
  function CONVERT_DATE(sVALUE in varchar2) return date as
  begin
    if (bDATE_1904) then
      return to_date('01-01-1904', 'DD-MM-YYYY') + to_number(sVALUE);
    else
      return to_date('01-03-1900', 'DD-MM-YYYY') +(to_number(sVALUE) - 61);
    end if;
  end CONVERT_DATE;

  /* Конверитровать boolean в varchar2 */
  function BOOLEAN_TO_CHAR(bVALUE in boolean) return varchar2 is
  begin
    return case bVALUE when true then 'TRUE' when false then 'FALSE' else 'NULL' end;
  end BOOLEAN_TO_CHAR;

  /* Конверитровать названия колонок Excel в цифры */
  function COL_ALFAN(sCOL varchar2) return pls_integer is
  begin
    return ascii(substr(sCOL, -1)) - 64 + nvl((ascii(substr(sCOL, -2, 1)) - 64) * 26, 0) + nvl((ascii(substr(sCOL, -3, 1)) - 64) * 676,
                                                                                               0);
  end COL_ALFAN;

  /* Конвертировать ячейки в формате A1:F1 в числовой формат */
  function PARSE_MERGE_CELL_NAME(sMERGE_CELL in varchar2) return tMERGE_CELL as
    tMC        tMERGE_CELL;
    sCELL1_COL varchar2(100);
    sCELL2_COL varchar2(100);
    iCELL1_ROW pls_integer;
    iCELL2_ROW pls_integer;
  begin
    begin
      select regexp_substr(CELL1, '[^[:digit:]]{1,3}', 1, 1) as CELL1_COL,
             regexp_substr(CELL1, '[[:digit:]]{1,3}', 1, 1) as CELL1_ROW,
             regexp_substr(CELL2, '[^[:digit:]]{1,3}', 1, 1) as CELL2_COL,
             regexp_substr(CELL2, '[[:digit:]]{1,3}', 1, 1) as CELL2_ROW
        into sCELL1_COL, iCELL1_ROW, sCELL2_COL, iCELL2_ROW
        from (select T.MERGE_CELL,
                     regexp_substr(T.MERGE_CELL, '[^:]+', 1, 1) as CELL1,
                     regexp_substr(T.MERGE_CELL, '[^:]+', 1, 2) as CELL2
                from (select sMERGE_CELL as MERGE_CELL from dual) T) T;
    exception
      when others then
        raise;
    end;
    tMC.CELL1_COL := COL_ALFAN(sCELL1_COL);
    tMC.CELL1_ROW := iCELL1_ROW;
    tMC.CELL2_COL := COL_ALFAN(sCELL2_COL);
    tMC.CELL2_ROW := iCELL2_ROW;
    return tMC;
  end PARSE_MERGE_CELL_NAME;

  /* Определить формат даты в ячейке */
  function CELL_IS_DATE_FORMAT(iSTYLE_INDEX in pls_integer) return boolean as
    bRESULT     boolean := false;
    iNUM_FMT_ID pls_integer;
  begin
    /* TODO */
    if (iNUM_FMT_ID >= 14 and iNUM_FMT_ID <= 22) then
      bRESULT := true;
    end if;
    return bRESULT;
  end CELL_IS_DATE_FORMAT;

  /* Колонки */
  function GET_COLS(cSHEET_XML in clob) return tbCOLS as
    tabCOLS        tbCOLS := tbCOLS();
    tDOM_NODE      dbms_xmldom.domnode;
    tDOM_NODE_LIST dbms_xmldom.domnodelist;
    sMIN           varchar2(4000);
    sMAX           varchar2(4000);
    sWIDTH         varchar2(4000);
    iMIN           pls_integer;
    iMAX           pls_integer;
    nWIDTH         number;
  begin
    tDOM_NODE      := CLOB2NODE(cSHEET_XML);
    tDOM_NODE_LIST := dbms_xslprocessor.selectnodes(tDOM_NODE, '/worksheet/cols/col');
  
    for cur in 0 .. dbms_xmldom.getLength(tDOM_NODE_LIST) - 1 loop
      sMIN   := dbms_xslprocessor.valueof(dbms_xmldom.item(tDOM_NODE_LIST, cur), '@min');
      sMAX   := dbms_xslprocessor.valueof(dbms_xmldom.item(tDOM_NODE_LIST, cur), '@max');
      sWIDTH := dbms_xslprocessor.valueof(dbms_xmldom.item(tDOM_NODE_LIST, cur), '@width');
      iMIN   := to_number(sMIN);
      iMAX   := to_number(sMAX);
      nWIDTH := to_number(sWIDTH);
    
      for indx in iMIN .. iMAX loop
        tabCOLS.extend;
        tabCOLS(tabCOLS.last).INDX := indx;
        tabCOLS(tabCOLS.last).WIDTH := nWIDTH;
      end loop;
    end loop;
  
    return tabCOLS;
  end GET_COLS;

  /* Объединенные ячейки */
  function GET_MERGE_CELLS(cSHEET_XML in clob) return tbMERGE_CELLS as
    tabMERGE_CELLS tbMERGE_CELLS := tbMERGE_CELLS();
    tDOM_NODE      dbms_xmldom.domnode;
    tDOM_NODE_LIST dbms_xmldom.domnodelist;
    sMERGE_CELL    varchar2(4000);
  begin
    tDOM_NODE      := CLOB2NODE(cSHEET_XML);
    tDOM_NODE_LIST := dbms_xslprocessor.selectnodes(tDOM_NODE, '/worksheet/mergeCells/mergeCell');
  
    for cur in 0 .. dbms_xmldom.getLength(tDOM_NODE_LIST) - 1 loop
      sMERGE_CELL := dbms_xslprocessor.valueof(dbms_xmldom.item(tDOM_NODE_LIST, cur), '@ref');
      tabMERGE_CELLS.extend;
      tabMERGE_CELLS(tabMERGE_CELLS.last) := PARSE_MERGE_CELL_NAME(sMERGE_CELL);
    end loop;
    return tabMERGE_CELLS;
  end GET_MERGE_CELLS;

  /* Таблица общих строк */
  function GET_SHARED_STRINGS(cSHARED_STRINGS_XML in clob) return tbSTRINGS as
    tabSTRINGS      tbSTRINGS;
    tDOM_NODE       dbms_xmldom.domnode;
    tDOM_NODE_LIST  dbms_xmldom.domnodelist;
    tDOM_NODE_LIST2 dbms_xmldom.domnodelist;
    iCOUNT          pls_integer := 0;
    iNUM            pls_integer := 0;
    iPOSITION       pls_integer := 5000;
  begin
    tDOM_NODE := CLOB2NODE(cSHARED_STRINGS_XML);
  
    loop
      tDOM_NODE_LIST := dbms_xslprocessor.selectnodes(tDOM_NODE,
                                                      '/sst/si[position()>="' || to_char(iNUM * iPOSITION + 1) ||
                                                      '" and position()<=" ' || to_char((iNUM + 1) * iPOSITION) || '"]',
                                                      sXMLNS);
      exit when dbms_xmldom.getlength(tDOM_NODE_LIST) = 0;
      iNUM := iNUM + 1;
      for i in 0 .. dbms_xmldom.getlength(tDOM_NODE_LIST) - 1 loop
        iCOUNT := tabSTRINGS.count;
        tabSTRINGS(iCOUNT) := dbms_xslprocessor.valueof(dbms_xmldom.item(tDOM_NODE_LIST, i), '.');
        if tabSTRINGS(iCOUNT) is null then
          tabSTRINGS(iCOUNT) := dbms_xslprocessor.valueof(dbms_xmldom.item(tDOM_NODE_LIST, i), '*/text()');
          if tabSTRINGS(iCOUNT) is null then
            tDOM_NODE_LIST2 := dbms_xslprocessor.selectnodes(dbms_xmldom.item(tDOM_NODE_LIST, i), 'r/t/text()');
            for j in 0 .. dbms_xmldom.getlength(tDOM_NODE_LIST2) - 1 loop
              tabSTRINGS(iCOUNT) := tabSTRINGS(iCOUNT) || dbms_xmldom.getnodevalue(dbms_xmldom.item(tDOM_NODE_LIST2, j));
            end loop;
          end if;
        end if;
      end loop;
    end loop;
  
    return tabSTRINGS;
  end GET_SHARED_STRINGS;

  /* Получить стили */
  function GET_CELL_STYLES(cSTYLES_XML in clob) return tbCELL_XFS as
    tabCELL_XFS         tbCELL_XFS;
    tDOM_NODE           dbms_xmldom.domnode;
    tDOM_NODE_LIST      dbms_xmldom.domnodelist;
    tDOM_NODE_ALIGNMENT dbms_xmldom.domnode;
    sTEMP               varchar2(32767);
    iCOUNT              pls_integer;
  begin
    tDOM_NODE      := CLOB2NODE(cSTYLES_XML);
    tDOM_NODE_LIST := dbms_xslprocessor.selectnodes(tDOM_NODE, '/styleSheet/cellXfs/xf', sXMLNS);
  
    for cur in 0 .. dbms_xmldom.getLength(tDOM_NODE_LIST) - 1 loop
      iCOUNT := tabCELL_XFS.count;
      sTEMP := dbms_xslprocessor.valueof(dbms_xmldom.item(tDOM_NODE_LIST, cur), '@numFmtId');
      tabCELL_XFS(iCOUNT).NUM_FMT_ID := to_number(sTEMP);
    
      sTEMP := dbms_xslprocessor.valueof(dbms_xmldom.item(tDOM_NODE_LIST, cur), '@fontId');
      tabCELL_XFS(iCOUNT).FONT_ID := to_number(sTEMP);
    
      sTEMP := dbms_xslprocessor.valueof(dbms_xmldom.item(tDOM_NODE_LIST, cur), '@fillId');
      tabCELL_XFS(iCOUNT).FILL_ID := to_number(sTEMP);
    
      sTEMP := dbms_xslprocessor.valueof(dbms_xmldom.item(tDOM_NODE_LIST, cur), '@borderId');
      tabCELL_XFS(iCOUNT).BORDER_ID := to_number(sTEMP);
    
      tDOM_NODE_ALIGNMENT := dbms_xslprocessor.selectSingleNode(dbms_xmldom.item(tDOM_NODE_LIST, cur), 'alignment');
    
      sTEMP := dbms_xslprocessor.valueof(tDOM_NODE_ALIGNMENT, '@horizontal');
      tabCELL_XFS(iCOUNT).HORIZONTAL := sTEMP;
    
      sTEMP := dbms_xslprocessor.valueof(tDOM_NODE_ALIGNMENT, '@vertical');
      tabCELL_XFS(iCOUNT).VERTICAL := sTEMP;
    
      sTEMP := dbms_xslprocessor.valueof(tDOM_NODE_ALIGNMENT, '@wrapText');
      if (sTEMP = '1' or sTEMP = 'true') then
        tabCELL_XFS(iCOUNT).WRAPTEXT := true;
      else
        tabCELL_XFS(iCOUNT).WRAPTEXT := false;
      end if;
    end loop;
  
    if (bDEBUG) then
      for i in 0 .. tabCELL_XFS.count - 1 loop
        dbms_output.put_line('NUM_FMT_ID: ' || tabCELL_XFS(i).NUM_FMT_ID || ' FONT_ID: ' || tabCELL_XFS(i).FONT_ID || ' FILL_ID: ' || tabCELL_XFS(i)
                             .FILL_ID || ' BORDER_ID: ' || tabCELL_XFS(i).BORDER_ID || ' HORIZONTAL: ' || tabCELL_XFS(i).HORIZONTAL ||
                             ' VERTICAL: ' || tabCELL_XFS(i).VERTICAL || ' WRAPTEXT: ' || BOOLEAN_TO_CHAR(tabCELL_XFS(i).WRAPTEXT));
      end loop;
    end if;
  
    return tabCELL_XFS;
  end GET_CELL_STYLES;

  /* Формат ячеек */
  function GET_NUM_FMTS(cSTYLES_XML in clob) return tbNUM_FMTS as
    tabNUM_FMTS    tbNUM_FMTS;
    tDOM_NODE      dbms_xmldom.domnode;
    tDOM_NODE_LIST dbms_xmldom.domnodelist;
    sTEMP          varchar2(4000);
    iCOUNT         pls_integer;
  begin
    tDOM_NODE      := CLOB2NODE(cSTYLES_XML);
    tDOM_NODE_LIST := dbms_xslprocessor.selectnodes(tDOM_NODE, '/styleSheet/numFmts/numFmt');
  
    for cur in 0 .. dbms_xmldom.getLength(tDOM_NODE_LIST) - 1 loop
      sTEMP := dbms_xslprocessor.valueof(dbms_xmldom.item(tDOM_NODE_LIST, cur), '@numFmtId');
      /* numFmtId как индекс массива */
      iCOUNT := to_number(sTEMP);
      tabNUM_FMTS(iCOUNT).NUM_FMT_ID := to_number(sTEMP);
    
      sTEMP := dbms_xslprocessor.valueof(dbms_xmldom.item(tDOM_NODE_LIST, cur), '@formatCode');
      tabNUM_FMTS(iCOUNT).FORMAT_CODE := sTEMP;
    
    end loop;
  
    if (bDEBUG) then
      if (tabNUM_FMTS.count > 0) then
        for i in tabNUM_FMTS.first .. tabNUM_FMTS.last loop
          dbms_output.put_line('NUM_FMT_ID: ' || tabNUM_FMTS(i).NUM_FMT_ID || ' FORMAT_CODE: ' || tabNUM_FMTS(i).FORMAT_CODE);
        end loop;
      end if;
    end if;
  
    return tabNUM_FMTS;
  end GET_NUM_FMTS;

  /* Границы ячеек  */
  function GET_BORDERS(cSTYLES_XML in clob) return tbBORDERS as
    tabBORDERS     tbBORDERS;
    tDOM_NODE      dbms_xmldom.domnode;
    tDOM_NODE_LIST dbms_xmldom.domnodelist;
    tDOM_NODE2     dbms_xmldom.domnode;
    sTEMP          varchar2(4000);
    iCOUNT         pls_integer;
  begin
    tDOM_NODE      := CLOB2NODE(cSTYLES_XML);
    tDOM_NODE_LIST := dbms_xslprocessor.selectnodes(tDOM_NODE, '/styleSheet/borders/border');
  
    for cur in 0 .. dbms_xmldom.getLength(tDOM_NODE_LIST) - 1 loop
      iCOUNT := tabBORDERS.count;
      tDOM_NODE2 := dbms_xslprocessor.selectSingleNode(dbms_xmldom.item(tDOM_NODE_LIST, cur), 'left');
      sTEMP := dbms_xslprocessor.valueof(tDOM_NODE2, '@style');
      tabBORDERS(iCOUNT).LEFT := sTEMP;
    
      tDOM_NODE2 := dbms_xslprocessor.selectSingleNode(dbms_xmldom.item(tDOM_NODE_LIST, cur), 'right');
      sTEMP := dbms_xslprocessor.valueof(tDOM_NODE2, '@style');
      tabBORDERS(iCOUNT).RIGHT := sTEMP;
    
      tDOM_NODE2 := dbms_xslprocessor.selectSingleNode(dbms_xmldom.item(tDOM_NODE_LIST, cur), 'top');
      sTEMP := dbms_xslprocessor.valueof(tDOM_NODE2, '@style');
      tabBORDERS(iCOUNT).TOP := sTEMP;
    
      tDOM_NODE2 := dbms_xslprocessor.selectSingleNode(dbms_xmldom.item(tDOM_NODE_LIST, cur), 'bottom');
      sTEMP := dbms_xslprocessor.valueof(tDOM_NODE2, '@style');
      tabBORDERS(iCOUNT).BOTTOM := sTEMP;
    end loop;
  
    if (bDEBUG) then
      for i in 0 .. tabBORDERS.count - 1 loop
        dbms_output.put_line('LEFT: ' || tabBORDERS(i).LEFT || ' RIGHT: ' || tabBORDERS(i).RIGHT || ' TOP: ' || tabBORDERS(i).TOP ||
                             ' BOTTOM: ' || tabBORDERS(i).BOTTOM);
      end loop;
    end if;
  
    return tabBORDERS;
  end GET_BORDERS;

  /* Заливка  */
  function GET_FILLS(cSTYLES_XML in clob) return tbFILLS as
    tabFILLS       tbFILLS;
    tDOM_NODE      dbms_xmldom.domnode;
    tDOM_NODE_LIST dbms_xmldom.domnodelist;
    tDOM_NODE2     dbms_xmldom.domnode;
    sTEMP          varchar2(4000);
    iCOUNT         pls_integer;
  begin
    tDOM_NODE      := CLOB2NODE(cSTYLES_XML);
    tDOM_NODE_LIST := dbms_xslprocessor.selectnodes(tDOM_NODE, '/styleSheet/fills/fill');
  
    for cur in 0 .. dbms_xmldom.getLength(tDOM_NODE_LIST) - 1 loop
      iCOUNT     := tabFILLS.count;
      tDOM_NODE2 := dbms_xslprocessor.selectSingleNode(dbms_xmldom.item(tDOM_NODE_LIST, cur), 'patternFill');
      if (not dbms_xmldom.isnull(tDOM_NODE2)) then
        sTEMP := dbms_xslprocessor.valueof(tDOM_NODE2, '@patternType');
        tabFILLS(iCOUNT).PATTERN_TYPE := sTEMP;
      end if;
    
      tDOM_NODE2 := dbms_xslprocessor.selectSingleNode(dbms_xmldom.item(tDOM_NODE_LIST, cur), 'patternFill/fgColor');
      if (not dbms_xmldom.isnull(tDOM_NODE2)) then
        sTEMP := dbms_xslprocessor.valueof(tDOM_NODE2, '@rgb');
        tabFILLS(iCOUNT).FG_RGB := sTEMP;
      end if;
    
    end loop;
  
    if (bDEBUG) then
      for i in 0 .. tabFILLS.count - 1 loop
        dbms_output.put_line('PATTERN_TYPE: ' || tabFILLS(i).PATTERN_TYPE || ' FG_RGB: ' || tabFILLS(i).FG_RGB);
      end loop;
    end if;
  
    return tabFILLS;
  end GET_FILLS;

  /* Шрифты  */
  function GET_FONTS(cSTYLES_XML in clob) return tbFONTS as
    tabFONTS       tbFONTS;
    tDOM_NODE      dbms_xmldom.domnode;
    tDOM_NODE_LIST dbms_xmldom.domnodelist;
    tDOM_NODE2     dbms_xmldom.domnode;
    sTEMP          varchar2(4000);
    iCOUNT         pls_integer;
  begin
    tDOM_NODE      := CLOB2NODE(cSTYLES_XML);
    tDOM_NODE_LIST := dbms_xslprocessor.selectnodes(tDOM_NODE, '/styleSheet/fonts/font');
  
    for cur in 0 .. dbms_xmldom.getLength(tDOM_NODE_LIST) - 1 loop
      iCOUNT     := tabFONTS.count;
      tDOM_NODE2 := dbms_xslprocessor.selectSingleNode(dbms_xmldom.item(tDOM_NODE_LIST, cur), 'name');
      if (not dbms_xmldom.isnull(tDOM_NODE2)) then
        sTEMP := dbms_xslprocessor.valueof(tDOM_NODE2, '@val');
        tabFONTS(iCOUNT).NAME := sTEMP;
      end if;
    
      tDOM_NODE2 := dbms_xslprocessor.selectSingleNode(dbms_xmldom.item(tDOM_NODE_LIST, cur), 'family');
      if (not dbms_xmldom.isnull(tDOM_NODE2)) then
        sTEMP := dbms_xslprocessor.valueof(tDOM_NODE2, '@val');
        tabFONTS(iCOUNT).FAMILY := sTEMP;
      end if;
    
      tDOM_NODE2 := dbms_xslprocessor.selectSingleNode(dbms_xmldom.item(tDOM_NODE_LIST, cur), 'sz');
      if (not dbms_xmldom.isnull(tDOM_NODE2)) then
        sTEMP := dbms_xslprocessor.valueof(tDOM_NODE2, '@val');
        tabFONTS(iCOUNT).FONTSIZE := sTEMP;
      end if;
    
      tDOM_NODE2 := dbms_xslprocessor.selectSingleNode(dbms_xmldom.item(tDOM_NODE_LIST, cur), 'color');
      if (not dbms_xmldom.isnull(tDOM_NODE2)) then
        sTEMP := dbms_xslprocessor.valueof(tDOM_NODE2, '@rgb');
        tabFONTS(iCOUNT).COLOR := sTEMP;
      end if;
    
      tDOM_NODE2 := dbms_xslprocessor.selectSingleNode(dbms_xmldom.item(tDOM_NODE_LIST, cur), 'b');
      if (not dbms_xmldom.isnull(tDOM_NODE2)) then
        tabFONTS(iCOUNT).BOLD := true;
      end if;
    
      tDOM_NODE2 := dbms_xslprocessor.selectSingleNode(dbms_xmldom.item(tDOM_NODE_LIST, cur), 'i');
      if (not dbms_xmldom.isnull(tDOM_NODE2)) then
        tabFONTS(iCOUNT).ITALIC := true;
      end if;
    
      tDOM_NODE2 := dbms_xslprocessor.selectSingleNode(dbms_xmldom.item(tDOM_NODE_LIST, cur), 'u');
      if (not dbms_xmldom.isnull(tDOM_NODE2)) then
        tabFONTS(iCOUNT).UNDERLINE := true;
      end if;
    
    end loop;
  
    if (bDEBUG) then
      for i in 0 .. tabFONTS.count - 1 loop
        dbms_output.put_line('NAME: ' || tabFONTS(i).NAME || ' FAMILY: ' || tabFONTS(i).FAMILY || ' FONTSIZE: ' || tabFONTS(i)
                             .FONTSIZE || ' COLOR: ' || tabFONTS(i).COLOR || ' BOLD: ' || BOOLEAN_TO_CHAR(tabFONTS(i).BOLD) ||
                             ' ITALIC: ' || BOOLEAN_TO_CHAR(tabFONTS(i).ITALIC) || ' UNDERLINE: ' ||
                             BOOLEAN_TO_CHAR(tabFONTS(i).UNDERLINE));
      end loop;
    end if;
  
    return tabFONTS;
  end GET_FONTS;

  /* Данные листа */
  function GET_SHEET_DATA(cSHEET_XML in clob, cSHARED_STRINGS_XML in clob) return tbCELLS as
    tabCELLS        tbCELLS := tbCELLS();
    tabSTRINGS      tbSTRINGS;
    tDOM_NODE       dbms_xmldom.domnode;
    tDOM_NODE_LIST  dbms_xmldom.domnodelist;
    tDOM_NODE_LIST2 dbms_xmldom.domnodelist;
    sTEMP           varchar2(32767);
    sCELL_TYPE      varchar2(100);
    sCELL_VALUE     varchar2(32767);
    sSTYLE_INDEX    varchar2(100);
    sFORMULA        varchar2(32767);
    nVALUE          number;
  begin
    tabROWS.delete;
    tabSTRINGS     := GET_SHARED_STRINGS(cSHARED_STRINGS_XML);
    tDOM_NODE      := CLOB2NODE(cSHEET_XML);
    tDOM_NODE_LIST := dbms_xslprocessor.selectnodes(tDOM_NODE, '/worksheet/sheetData/row');
  
    for r in 0 .. dbms_xmldom.getlength(tDOM_NODE_LIST) - 1 loop
      tabROWS.extend;
      sTEMP := dbms_xslprocessor.valueof(dbms_xmldom.item(tDOM_NODE_LIST, r), '@r');
      tabROWS(tabROWS.last).INDX := to_number(sTEMP);
      sTEMP := dbms_xslprocessor.valueof(dbms_xmldom.item(tDOM_NODE_LIST, r), '@ht');
      tabROWS(tabROWS.last).HEIGHT := to_number(sTEMP);
    
      tDOM_NODE_LIST2 := dbms_xslprocessor.selectnodes(dbms_xmldom.item(tDOM_NODE_LIST, r), 'c');
      for j in 0 .. dbms_xmldom.getlength(tDOM_NODE_LIST2) - 1 loop
        tabCELLS.extend;
        sTEMP        := dbms_xslprocessor.valueof(dbms_xmldom.item(tDOM_NODE_LIST2, j), '@r', sXMLNS);
        sCELL_VALUE  := dbms_xslprocessor.valueof(dbms_xmldom.item(tDOM_NODE_LIST2, j), 'v');
        sCELL_TYPE   := dbms_xslprocessor.valueof(dbms_xmldom.item(tDOM_NODE_LIST2, j), '@t');
        sSTYLE_INDEX := dbms_xslprocessor.valueof(dbms_xmldom.item(tDOM_NODE_LIST2, j), '@s');
        sFORMULA     := dbms_xslprocessor.valueof(dbms_xmldom.item(tDOM_NODE_LIST2, j), 'f');
      
        tabCELLS(tabCELLS.last).ROW := ltrim(sTEMP, rtrim(sTEMP, '0123456789'));
        tabCELLS(tabCELLS.last).COL := COL_ALFAN(rtrim(sTEMP, '0123456789'));
        tabCELLS(tabCELLS.last).STYLE_INDEX := to_number(sSTYLE_INDEX);
        tabCELLS(tabCELLS.last).FORMULA := sFORMULA;
      
        if sCELL_TYPE in ('str', 'inlineStr', 'e') then
          tabCELLS(tabCELLS.last).CELL_TYPE := 'S';
          tabCELLS(tabCELLS.last).STRING_VAL := sCELL_VALUE;
        elsif sCELL_TYPE = 's' then
          tabCELLS(tabCELLS.last).CELL_TYPE := 'S';
          if (sCELL_VALUE is not null) then
            tabCELLS(tabCELLS.last).STRING_VAL := tabSTRINGS(to_number(sCELL_VALUE));
          end if;
        else
          nVALUE := CONVERT_NUMBER(sCELL_VALUE);
          if sSTYLE_INDEX is not null and CELL_IS_DATE_FORMAT(to_number(sSTYLE_INDEX)) then
            tabCELLS(tabCELLS.last).CELL_TYPE := 'D';
            tabCELLS(tabCELLS.last).DATE_VAL := CONVERT_DATE(nVALUE);
          else
            tabCELLS(tabCELLS.last).CELL_TYPE := 'N';
            nVALUE := round(nVALUE, 14 - substr(to_char(nVALUE, 'TME'), -3));
            tabCELLS(tabCELLS.last).NUMBER_VAL := nVALUE;
          end if;
        end if;
      end loop;
    end loop;
  
    if (bDEBUG) then
      for i in 1 .. tabCELLS.count loop
        dbms_output.put_line('ROW: ' || tabCELLS(i).ROW || ' COL: ' || tabCELLS(i).COL || ' CELL_TYPE: ' || tabCELLS(i).CELL_TYPE ||
                             ' STRING_VAL: ' || tabCELLS(i).STRING_VAL || ' NUMBER_VAL: ' || tabCELLS(i).NUMBER_VAL || ' DATE_VAL: ' || tabCELLS(i)
                             .DATE_VAL || ' STYLE_INDEX: ' || tabCELLS(i).STYLE_INDEX || ' FORMULA: ' || tabCELLS(i).FORMULA);
      end loop;
      for i in 1 .. tabROWS.count loop
        dbms_output.put_line('INDX: ' || tabROWS(i).INDX || ' HEIGHT: ' || tabROWS(i).HEIGHT);
      end loop;
    end if;
  
    return tabCELLS;
  end GET_SHEET_DATA;

  /* Формат даты */
  function GET_DATE_1904(cWORKBOOK_XML in clob) return boolean as
    tDOM_NODE dbms_xmldom.domnode;
    bRESULT   boolean;
  begin
    tDOM_NODE := CLOB2NODE(cWORKBOOK_XML);
    bRESULT   := lower(dbms_xslprocessor.valueof(tDOM_NODE, '/workbook/workbookPr/@date1904', sXMLNS)) in ('true', '1');
    bRESULT   := nvl(bRESULT, false);
  
    if (bDEBUG) then
      dbms_output.put_line('DATE_1904: ' || BOOLEAN_TO_CHAR(bRESULT));
    end if;
  
    return bRESULT;
  end GET_DATE_1904;

  /* Сгенерировать код для AS_XLSX */
  procedure GEN_CODE_FOR_AS_XLSX(cSHEET_XML in clob, cSHARED_STRINGS_XML in clob, cSTYLES_XML in clob) as
    tabCOLS        tbCOLS := tbCOLS();
    tabMERGE_CELLS tbMERGE_CELLS := tbMERGE_CELLS();
    tabCELLS       tbCELLS := tbCELLS();
    tabCELL_XFS    tbCELL_XFS;
    tabBORDERS     tbBORDERS;
    sALIGNMENT     varchar2(4000);
    iNUM           pls_integer;
    sBORDER        varchar2(4000);
    iBORDER_ID     pls_integer;
  begin
    tabCOLS        := GET_COLS(cSHEET_XML);
    tabMERGE_CELLS := GET_MERGE_CELLS(cSHEET_XML);
    tabCELL_XFS    := GET_CELL_STYLES(cSTYLES_XML);
    tabBORDERS     := GET_BORDERS(cSTYLES_XML);
    tabCELLS       := GET_SHEET_DATA(cSHEET_XML, cSHARED_STRINGS_XML);
  
    for i in 1 .. tabCOLS.count loop
      dbms_output.put_line('AS_XLSX.SET_COLUMN_WIDTH(' || tabCOLS(i).INDX || ', ' || tabCOLS(i).WIDTH || ');');
    end loop;
  
    for i in 1 .. tabCELLS.count loop
      iNUM       := to_number(tabCELLS(i).STYLE_INDEX);
      sALIGNMENT := '';
      sBORDER    := '';
      if (iNUM is not null) then
        if (tabCELL_XFS.exists(iNUM)) then
          if (tabCELL_XFS(iNUM).HORIZONTAL is not null) then
            sALIGNMENT := sALIGNMENT || 'P_HORIZONTAL => ''' || tabCELL_XFS(iNUM).HORIZONTAL || ''', ';
          end if;
          if (tabCELL_XFS(iNUM).VERTICAL is not null) then
            sALIGNMENT := sALIGNMENT || 'P_VERTICAL => ''' || tabCELL_XFS(iNUM).VERTICAL || ''', ';
          end if;
          if (tabCELL_XFS(iNUM).WRAPTEXT is not null) then
            sALIGNMENT := sALIGNMENT || 'P_WRAPTEXT => ' || BOOLEAN_TO_CHAR(tabCELL_XFS(iNUM).WRAPTEXT) || ', ';
          end if;
          sALIGNMENT := rtrim(sALIGNMENT, ', ');
          if (sALIGNMENT is not null) then
            sALIGNMENT := ', P_ALIGNMENT => AS_XLSX.GET_ALIGNMENT(' || sALIGNMENT || ')';
          end if;
        
          iBORDER_ID := to_number(tabCELL_XFS(iNUM).BORDER_ID);
          if (tabBORDERS.exists(iBORDER_ID)) then
            if (tabBORDERS(iBORDER_ID).TOP is not null or tabBORDERS(iBORDER_ID).BOTTOM is not null or tabBORDERS(iBORDER_ID)
               .LEFT is not null or tabBORDERS(iBORDER_ID).RIGHT is not null) then
              sBORDER := ', P_BORDERID => AS_XLSX.GET_BORDER(P_TOP=>''' || tabBORDERS(iBORDER_ID).TOP || ''', P_BOTTOM=>''' || tabBORDERS(iBORDER_ID)
                        .BOTTOM || ''', P_LEFT=>''' || tabBORDERS(iBORDER_ID).LEFT || ''', P_RIGHT=>''' || tabBORDERS(iBORDER_ID)
                        .RIGHT || ''')';
            end if;
          end if;
        end if;
      end if;
    
      if (tabCELLS(i).CELL_TYPE = 'N' and tabCELLS(i).NUMBER_VAL is not null) then
        dbms_output.put_line('AS_XLSX.CELL(' || tabCELLS(i).COL || ', ' || tabCELLS(i).ROW || ', ' || tabCELLS(i).NUMBER_VAL || '' ||
                             sALIGNMENT || sBORDER || ');');
      elsif (tabCELLS(i).CELL_TYPE = 'D' and tabCELLS(i).DATE_VAL is not null) then
        dbms_output.put_line('AS_XLSX.CELL(' || tabCELLS(i).COL || ', ' || tabCELLS(i).ROW || ', ''' ||
                             to_char(tabCELLS(i).DATE_VAL, 'DD.MM.YYYY HH24:MI') || '''' || sALIGNMENT || sBORDER || ');');
      else
        dbms_output.put_line('AS_XLSX.CELL(' || tabCELLS(i).COL || ', ' || tabCELLS(i).ROW || ', ''' || tabCELLS(i).STRING_VAL || '''' ||
                             sALIGNMENT || sBORDER || ');');
      end if;
    end loop;
  
    for i in 1 .. tabMERGE_CELLS.count loop
      dbms_output.put_line('AS_XLSX.MERGECELLS(' || tabMERGE_CELLS(i).CELL1_COL || ', ' || tabMERGE_CELLS(i).CELL1_ROW || ', ' || tabMERGE_CELLS(i)
                           .CELL2_COL || ', ' || tabMERGE_CELLS(i).CELL2_ROW || ');');
    end loop;
  
  end GEN_CODE_FOR_AS_XLSX;

  /* Построить AS_XLSX */
  procedure BUILD_AS_XLSX(cSHEET_XML in clob, cSHARED_STRINGS_XML in clob, cSTYLES_XML in clob) as
    tabCOLS        tbCOLS := tbCOLS();
    tabMERGE_CELLS tbMERGE_CELLS := tbMERGE_CELLS();
    tabCELLS       tbCELLS := tbCELLS();
    tabCELL_XFS    tbCELL_XFS;
    tabBORDERS     tbBORDERS;
    tabFONTS       tbFONTS;
    tabNUM_FMTS    tbNUM_FMTS;
    tabFILLS       tbFILLS;
    iSTYLE_INDEX   pls_integer;
    tALIGNMENT     AS_XLSX.TP_ALIGNMENT;
    iBORDER_INDEX  pls_integer;
    iBORDER_ID     pls_integer;
    iFONT_INDEX    pls_integer;
    iFONT_ID       pls_integer;
    iNUM_FMT_INDEX pls_integer;
    iNUM_FMT_ID    pls_integer;
    iFILL_INDEX    pls_integer;
    iFILL_ID       pls_integer;
  begin
    tabCOLS        := GET_COLS(cSHEET_XML);
    tabMERGE_CELLS := GET_MERGE_CELLS(cSHEET_XML);
    tabCELL_XFS    := GET_CELL_STYLES(cSTYLES_XML);
    tabBORDERS     := GET_BORDERS(cSTYLES_XML);
    tabFONTS       := GET_FONTS(cSTYLES_XML);
    tabNUM_FMTS    := GET_NUM_FMTS(cSTYLES_XML);
    tabFILLS       := GET_FILLS(cSTYLES_XML);
    tabCELLS       := GET_SHEET_DATA(cSHEET_XML, cSHARED_STRINGS_XML);
  
    AS_XLSX.CLEAR_WORKBOOK;
    AS_XLSX.NEW_SHEET('sheet1');
  
    for i in 1 .. tabCOLS.count loop
      AS_XLSX.SET_COLUMN_WIDTH(tabCOLS(i).INDX, tabCOLS(i).WIDTH);
    end loop;
  
    for i in 1 .. tabCELLS.count loop
      iSTYLE_INDEX := to_number(tabCELLS(i).STYLE_INDEX);
      if (iSTYLE_INDEX is not null) then
        if (tabCELL_XFS.exists(iSTYLE_INDEX)) then
          tALIGNMENT := AS_XLSX.GET_ALIGNMENT(P_HORIZONTAL => tabCELL_XFS(iSTYLE_INDEX).HORIZONTAL,
                                              P_VERTICAL   => tabCELL_XFS(iSTYLE_INDEX).VERTICAL,
                                              P_WRAPTEXT   => tabCELL_XFS(iSTYLE_INDEX).WRAPTEXT);
        
          iBORDER_INDEX := tabCELL_XFS(iSTYLE_INDEX).BORDER_ID;
          if (tabBORDERS.exists(iBORDER_INDEX)) then
            iBORDER_ID := AS_XLSX.GET_BORDER(P_TOP    => tabBORDERS(iBORDER_INDEX).TOP,
                                             P_BOTTOM => tabBORDERS(iBORDER_INDEX).BOTTOM,
                                             P_LEFT   => tabBORDERS(iBORDER_INDEX).LEFT,
                                             P_RIGHT  => tabBORDERS(iBORDER_INDEX).RIGHT);
          end if;
        
          iFONT_INDEX := tabCELL_XFS(iSTYLE_INDEX).FONT_ID;
          if (tabFONTS.exists(iFONT_INDEX)) then
            iFONT_ID := AS_XLSX.GET_FONT(P_NAME      => tabFONTS(iFONT_INDEX).NAME,
                                         P_FAMILY    => tabFONTS(iFONT_INDEX).FAMILY,
                                         P_FONTSIZE  => tabFONTS(iFONT_INDEX).FONTSIZE,
                                         P_UNDERLINE => tabFONTS(iFONT_INDEX).UNDERLINE,
                                         P_ITALIC    => tabFONTS(iFONT_INDEX).ITALIC,
                                         P_BOLD      => tabFONTS(iFONT_INDEX).BOLD,
                                         P_RGB       => tabFONTS(iFONT_INDEX).COLOR);
          end if;
        
          iNUM_FMT_INDEX := tabCELL_XFS(iSTYLE_INDEX).NUM_FMT_ID;
          if (iNUM_FMT_INDEX < 164) then
            iNUM_FMT_ID := iNUM_FMT_INDEX;
          else
            if (tabNUM_FMTS.exists(iNUM_FMT_INDEX)) then
              iNUM_FMT_ID := AS_XLSX.GET_NUMFMT(tabNUM_FMTS(iNUM_FMT_INDEX).FORMAT_CODE);
            end if;
          end if;
        
          iFILL_INDEX := tabCELL_XFS(iSTYLE_INDEX).FILL_ID;
          if (tabFILLS.exists(iFILL_INDEX)) then
            iFILL_ID := AS_XLSX.GET_FILL(tabFILLS(iFILL_INDEX).PATTERN_TYPE, tabFILLS(iFILL_INDEX).FG_RGB);
          end if;
        
        end if;
      end if;
    
      if (tabCELLS(i).CELL_TYPE = 'N' and tabCELLS(i).NUMBER_VAL is not null) then
        AS_XLSX.CELL(tabCELLS   (i).COL,
                     tabCELLS   (i).ROW,
                     tabCELLS   (i).NUMBER_VAL,
                     P_ALIGNMENT => tALIGNMENT,
                     P_BORDERID  => iBORDER_ID,
                     P_FONTID    => iFONT_ID,
                     P_NUMFMTID  => iNUM_FMT_ID,
                     P_FILLID    => iFILL_ID);
      elsif (tabCELLS(i).CELL_TYPE = 'D' and tabCELLS(i).DATE_VAL is not null) then
        AS_XLSX.CELL(tabCELLS   (i).COL,
                     tabCELLS   (i).ROW,
                     tabCELLS   (i).DATE_VAL,
                     P_ALIGNMENT => tALIGNMENT,
                     P_BORDERID  => iBORDER_ID,
                     P_FONTID    => iFONT_ID,
                     P_NUMFMTID  => iNUM_FMT_ID,
                     P_FILLID    => iFILL_ID);
      else
        if (tabCELLS(i).STRING_VAL is null) then
          iNUM_FMT_ID := null;
        end if;
        AS_XLSX.CELL(tabCELLS   (i).COL,
                     tabCELLS   (i).ROW,
                     tabCELLS   (i).STRING_VAL,
                     P_ALIGNMENT => tALIGNMENT,
                     P_BORDERID  => iBORDER_ID,
                     P_FONTID    => iFONT_ID,
                     P_NUMFMTID  => iNUM_FMT_ID,
                     P_FILLID    => iFILL_ID);
      end if;
      tALIGNMENT  := null;
      iBORDER_ID  := null;
      iFONT_ID    := null;
      iNUM_FMT_ID := null;
    end loop;
  
    for i in 1 .. tabMERGE_CELLS.count loop
      AS_XLSX.MERGECELLS(tabMERGE_CELLS(i).CELL1_COL,
                         tabMERGE_CELLS(i).CELL1_ROW,
                         tabMERGE_CELLS(i).CELL2_COL,
                         tabMERGE_CELLS(i).CELL2_ROW);
    end loop;
  
    for i in 1 .. tabROWS.count loop
      if (tabROWS(i).HEIGHT is not null) then
        AS_XLSX.SET_ROW_HEIGHT(tabROWS(i).INDX, tabROWS(i).HEIGHT);
      end if;
    end loop;
  
  end BUILD_AS_XLSX;

  /* Шаблон отчета в XLSX */
  function GET_TEMPLATE(nREPORT_ID in number, iVERSION pls_integer) return blob as
    bTEMPLATE blob;
  begin
    begin
      select T.TEMPLATE_DATA
        into bTEMPLATE
        from REPORT_TEMPLATE T
       where T.REPORT_ID = nREPORT_ID
         and T.VERSION = iVERSION;
    exception
      when NO_DATA_FOUND then
        RAISE_APPLICATION_ERROR(-20001, 'Шаблон отчета не найден');
      when others then
        raise;
    end;
    return bTEMPLATE;
  end GET_TEMPLATE;

  /* Формирование шаблона отчета */
  procedure CREATE_REPORT_TEMPLATE(nREPORT_ID in number, iVERSION pls_integer default 1) as
    bXLSX               blob;
    bUNZIP_FILE         blob;
    xFILE               XMLTYPE;
    cWORKBOOK_XML       clob;
    cSHEET_XML          clob;
    cSHARED_STRINGS_XML clob;
    cSTYLES_XML         clob;
  begin
    bXLSX := GET_TEMPLATE(nREPORT_ID, iVERSION);
  
    bUNZIP_FILE := AS_ZIP.GET_FILE(bXLSX, 'xl/workbook.xml');
    xFILE       := XMLTYPE.CREATEXML(bUNZIP_FILE, nls_charset_id('AL32UTF8'), null);
    if DBMS_LOB.ISTEMPORARY(bUNZIP_FILE) = 1 then
      DBMS_LOB.FREETEMPORARY(bUNZIP_FILE);
    end if;
    if (xFILE is not null) then
      cWORKBOOK_XML := xFILE.GETCLOBVAL();
    end if;
  
    bUNZIP_FILE := AS_ZIP.GET_FILE(bXLSX, 'xl/worksheets/sheet1.xml');
    xFILE       := XMLTYPE.CREATEXML(bUNZIP_FILE, nls_charset_id('AL32UTF8'), null);
    if DBMS_LOB.ISTEMPORARY(bUNZIP_FILE) = 1 then
      DBMS_LOB.FREETEMPORARY(bUNZIP_FILE);
    end if;
    if (xFILE is not null) then
      cSHEET_XML := xFILE.GETCLOBVAL();
    end if;
  
    bUNZIP_FILE := AS_ZIP.GET_FILE(bXLSX, 'xl/sharedStrings.xml');
    xFILE       := XMLTYPE.CREATEXML(bUNZIP_FILE, nls_charset_id('AL32UTF8'), null);
    if DBMS_LOB.ISTEMPORARY(bUNZIP_FILE) = 1 then
      DBMS_LOB.FREETEMPORARY(bUNZIP_FILE);
    end if;
    if (xFILE is not null) then
      cSHARED_STRINGS_XML := xFILE.GETCLOBVAL();
    end if;
  
    bUNZIP_FILE := AS_ZIP.GET_FILE(bXLSX, 'xl/styles.xml');
    xFILE       := XMLTYPE.CREATEXML(bUNZIP_FILE, nls_charset_id('AL32UTF8'), null);
    if DBMS_LOB.ISTEMPORARY(bUNZIP_FILE) = 1 then
      DBMS_LOB.FREETEMPORARY(bUNZIP_FILE);
    end if;
    if (xFILE is not null) then
      cSTYLES_XML := xFILE.GETCLOBVAL();
    end if;
  
    bDATE_1904 := GET_DATE_1904(cWORKBOOK_XML);
  
    BUILD_AS_XLSX(cSHEET_XML, cSHARED_STRINGS_XML, cSTYLES_XML);
  end CREATE_REPORT_TEMPLATE;

  /* Данные из файла */
  function GET_DATA(bXLSX in blob) return tbCELLS is
    bUNZIP_FILE         blob;
    xFILE               XMLTYPE;
    cWORKBOOK_XML       clob;
    cSHEET_XML          clob;
    cSHARED_STRINGS_XML clob;
  begin
    bUNZIP_FILE := AS_ZIP.GET_FILE(bXLSX, 'xl/workbook.xml');
    xFILE       := XMLTYPE.CREATEXML(bUNZIP_FILE, nls_charset_id('AL32UTF8'), null);
    if DBMS_LOB.ISTEMPORARY(bUNZIP_FILE) = 1 then
      DBMS_LOB.FREETEMPORARY(bUNZIP_FILE);
    end if;
    if (xFILE is not null) then
      cWORKBOOK_XML := xFILE.GETCLOBVAL();
    end if;
  
    bUNZIP_FILE := AS_ZIP.GET_FILE(bXLSX, 'xl/worksheets/sheet1.xml');
    xFILE       := XMLTYPE.CREATEXML(bUNZIP_FILE, nls_charset_id('AL32UTF8'), null);
    if DBMS_LOB.ISTEMPORARY(bUNZIP_FILE) = 1 then
      DBMS_LOB.FREETEMPORARY(bUNZIP_FILE);
    end if;
    if (xFILE is not null) then
      cSHEET_XML := xFILE.GETCLOBVAL();
    end if;
  
    bUNZIP_FILE := AS_ZIP.GET_FILE(bXLSX, 'xl/sharedStrings.xml');
    xFILE       := XMLTYPE.CREATEXML(bUNZIP_FILE, nls_charset_id('AL32UTF8'), null);
    if DBMS_LOB.ISTEMPORARY(bUNZIP_FILE) = 1 then
      DBMS_LOB.FREETEMPORARY(bUNZIP_FILE);
    end if;
    if (xFILE is not null) then
      cSHARED_STRINGS_XML := xFILE.GETCLOBVAL();
    end if;
  
    bDATE_1904 := GET_DATE_1904(cWORKBOOK_XML);
  
    return GET_SHEET_DATA(cSHEET_XML, cSHARED_STRINGS_XML);
  end GET_DATA;

begin
  /* Обязательно */
  execute immediate 'ALTER SESSION SET NLS_NUMERIC_CHARACTERS = ''.,''';
end PKG_XLSX_HELPER;
/
