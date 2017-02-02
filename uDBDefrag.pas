unit uDBDefrag;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Buttons, NxColumns, NxColumnClasses, NxScrollControl,
  NxCustomGridControl, NxCustomGrid, NxGrid, IBDatabase, DB, IBCustomDataSet,
  IBQuery, XLSFile, XLSWorkbook, XLSFormat, ComObj;

procedure BackupDBFile;
procedure XLS_Init;
procedure XLS_SetStyle(nRow, nCol: integer; bHeader: Boolean = False);
procedure XLS_ShowReport;
procedure PrepareGrid;
procedure DeleteDirectoryByID(table, id: string);
function CheckDirectoryIsInUse(field, match: string): boolean;
procedure ToggleButtons(bEnabled: boolean);
procedure DoDBDefrag(bBackUp: boolean = false);

const
  strHeader0: string = 'Директория';

implementation

uses Main;

var
  XLSDoc: TXLSFile;
  XLS_RowCounter: integer;

  // backup DB file 'usrdt_31.01.2017_20-18-47.msq'

procedure BackupDBFile;
var
  strDateTime, strDBFile, strDBFile_new: string;
begin
  strDBFile := AppPath + 'usrdt.msq';
  if not FileExists(strDBFile) then
  begin
    MessageBox(FormMain.Handle, 'Файл баз данных не найден.', 'Ошибка', MB_OK or MB_ICONERROR);
    exit;
  end;

  strDateTime := DateTimeToStr(now);
  strDateTime := StringReplace(strDateTime, ':', '-', [rfReplaceAll, rfIgnoreCase]);
  strDateTime := StringReplace(strDateTime, ' ', '_', [rfReplaceAll, rfIgnoreCase]);
  strDBFile_new := AppPath + '\usrdt' + '_' + strDateTime + '.msq';

  CopyFile(PChar(strDBFile), PChar(strDBFile_new), false);
end;

procedure XLS_Init;
begin
  XLSDoc := TXLSFile.Create;
  XLSDoc.Workbook.Sheets[0].Name := 'Дефрагментация БД ' + DateToStr(Now());
  XLS_RowCounter := 1; // because we have header row

  // Format Excel Doc
  with XLSDoc.Workbook.Sheets[0] do
  begin
    // Header
    Cells[0, 0].Value := 'Директория';
    XLS_SetStyle(0, 0, True);
    Cells[0, 1].Value := 'ID';
    XLS_SetStyle(0, 1, True);
    Cells[0, 2].Value := 'Название';
    XLS_SetStyle(0, 2, True);
    Freeze(1, 0);

    // Cols width
    Columns[0].WidthPx := 150;
    Columns[1].WidthPx := 50;
    Columns[2].WidthPx := 1500;

    // Print settings
    PageSetup.CenterHorizontally := False;
    PageSetup.CenterVertically := False;
    PageSetup.HeaderMargin := 0;
    PageSetup.FooterMargin := 0;
    PageSetup.TopMargin := 0.4;
    PageSetup.BottomMargin := 0.4;
    PageSetup.LeftMargin := 0.5;
    PageSetup.RightMargin := 0.4;
    PageSetup.PrintRowsOnEachPageFrom := 0;
    PageSetup.PrintRowsOnEachPageTo := 0;
    PageSetup.Orientation := xlLandscape;
    PageSetup.FitPagesWidth := 1;
    PageSetup.FitPagesHeight := 0;
    PageSetup.Zoom := False;
  end;
end;

procedure XLS_SetStyle(nRow, nCol: integer; bHeader: Boolean = False);
begin
  if (nCol < 0) or (nRow < 0) then
    Exit;

  with XLSDoc.Workbook.Sheets[0] do
  begin
    if bHeader then
    begin
      Cells[nRow, nCol].HAlign := xlHAlignCenter;
      Cells[nRow, nCol].VAlign := xlVAlignCenter;
      // Cells[nRow, nCol].FontBold:= True;
      Rows[nRow].HeightPx := 30;
      Cells[nRow, nCol].BorderStyle[xlBorderAll] := bsThin;
      Cells[nRow, nCol].BorderColorRGB[xlBorderAll] := RGB(0, 0, 0);
    end
    else
    begin
      Cells[nRow, nCol].VAlign := xlVAlignTop;
      Cells[nRow, nCol].Wrap := True;
      Rows[nRow].AutoFit;
    end;
    Cells[nRow, nCol].FontName := 'Tahoma';
    Cells[nRow, nCol].FontHeight := 10;
  end;
end;

procedure XLS_ShowReport;
var
  strReportFileName, strError: string;
  bShowReport: boolean;
  ExcelApp: OleVariant;
  i: integer;
begin
  if FormMain.SGDataUtils.RowCount = 0 then
  begin
    XLSDoc.Destroy;
    exit;
  end;

  bShowReport := True;
  strReportFileName := AppPath + 'report.xls';
  strError := 'ExcelApp: no errors';

  try // Closing
    ExcelApp := GetActiveOleObject('Excel.Application');
    ExcelApp.Visible := True;
    for i := ExcelApp.Workbooks.Count downto 1 do
      ExcelApp.Workbooks[i].Save;
    ExcelApp.Workbooks.Close;
    ExcelApp.Quit;
    VarClear(ExcelApp);
  except
    on E: Exception do
      strError := 'ExcelApp: ' + E.Message;
  end;

  try // Saving
    XLSDoc.SaveAs(strReportFileName);
  except
    on E: Exception do
    begin
      MessageBox(FormMain.Handle, PChar('Ошибка при сохранении файла отчета:' + #13 + strError + #13 + E.Message), 'Ошибка',
        MB_OK or MB_ICONERROR);
      bShowReport := False;
    end;
  end;

  if bShowReport then // Showing
  begin
    try
      ExcelApp := CreateOLEObject('Excel.Application');
      ExcelApp.Visible := True;
      ExcelApp.Workbooks.Open(strReportFileName);
      // ShellExecute(0, 'open', PChar(strReportFileName'), nil, nil, SW_SHOW);
    except
      on E: Exception do
      begin
        MessageBox(FormMain.Handle, PChar('Ошибка при отображении файла отчета:' + #13 + E.Message), 'Ошибка', MB_OK or MB_ICONERROR);
      end;
    end;
  end;

  XLSDoc.Destroy;
end;

procedure PrepareGrid;
var
  i: integer;
begin
  with FormMain do
  begin
    SGDataUtils.BeginUpdate;
    SGDataUtils.ClearRows;
    for i := SGDataUtils.Columns.Count - 1 downto 0 do
      SGDataUtils.Columns[i].Free;
    SGDataUtils.Columns.Add(TNxTextColumn, strHeader0);
    SGDataUtils.Columns[0].Position := 0;
    SGDataUtils.Columns[0].Width := 80;
    SGDataUtils.Columns.Add(TNxTextColumn, 'ID');
    SGDataUtils.Columns[1].Position := 1;
    SGDataUtils.Columns[1].Width := 40;
    SGDataUtils.Columns[1].Header.Glyph := NxTextColumn1.Header.Glyph;
    SGDataUtils.Columns[1].Header.DisplayMode := dmTextAndImage;
    SGDataUtils.Columns.Add(TNxTextColumn, 'Название');
    SGDataUtils.Columns[2].Position := 2;
    SGDataUtils.Columns[2].Width := 740;
    SGDataUtils.Columns[2].Header.Glyph := NxTextColumn2.Header.Glyph;
    SGDataUtils.Columns[2].Header.DisplayMode := dmTextAndImage;
    SGDataUtils.EndUpdate;
  end;
end;

procedure DeleteDirectoryByID(table, id: string);
var
  Q: TIBQuery;
begin
  if (Trim(table) = EmptyStr) or (Trim(id) = EmptyStr) then
    exit;
  Q := TIBQuery.Create(nil);
  Q.Database := FormMain.IBDatabase1;
  Q.Transaction := FormMain.IBTransaction1;
  Q.Close;
  Q.SQL.Text := 'delete from ' + table + ' where ID = :ID';
  Q.ParamByName('ID').AsString := id;
  try
    Q.Open;
    FormMain.IBTransaction1.CommitRetaining;
  finally
    Q.Close;
    Q.Free;
  end;
end;

function CheckDirectoryIsInUse(field, match: string): boolean;
var
  Q: TIBQuery;
begin
  result := false;
  if (Trim(field) = EmptyStr) or (Trim(match) = EmptyStr) then
    exit;
  Q := TIBQuery.Create(nil);
  Q.Database := FormMain.IBDatabase1;
  Q.Transaction := FormMain.IBTransaction1;
  Q.Close;
  Q.SQL.Text := 'select ID from BASE where lower(' + field + ') like :STR rows 1';
  Q.ParamByName('STR').AsString := match;
  Q.Open;
  if Q.RecordCount > 0 then
  begin
    result := true;
  end;
  Q.Close;
  Q.Free;
end;

procedure ToggleButtons(bEnabled: boolean);
begin
  with FormMain do
  begin
    sPageControl1.Enabled := bEnabled;
    btnImageExistingCheck.Enabled := bEnabled;
    btnEmailCheck.Enabled := bEnabled;
    btnWEBCheck.Enabled := bEnabled;
    btnDBDefrag.Enabled := bEnabled;
  end;
end;

procedure DoDBDefrag(bBackUp: boolean = false);
var
  Query: TIBQuery;
  CurrentProgress, i: Integer;
  isDeleted: boolean;
  strDirType, strID, strDIR: string;
begin
  ToggleButtons(false);
  PrepareGrid;

  if bBackUp = true then
    BackupDBFile;

  Query := TIBQuery.Create(nil);
  Query.Database := FormMain.IBDatabase1;
  Query.Transaction := FormMain.IBTransaction1;
  Query.SQL.Text := 'SELECT t1.id, t1.name, ''CURATOR'' as DIR FROM CURATOR as t1 UNION ALL ' +
    'SELECT t2.id, t2.name, ''RUBRIKATOR'' as DIR FROM RUBRIKATOR as t2 UNION ALL ' +
    'SELECT t3.id, t3.name, ''TYPE'' as DIR FROM TYPE as t3 UNION ALL ' +
    'SELECT t4.id, t4.name, ''NAPRAVLENIE'' as DIR FROM NAPRAVLENIE as t4 UNION ALL ' +
    'SELECT t5.id, t5.name, ''OFFICETYPE'' as DIR FROM OFFICETYPE as t5 UNION ALL ' +
    'SELECT t6.id, t6.name, ''COUNTRY'' as DIR FROM COUNTRY as t6 UNION ALL ' +
    'SELECT t7.id, t7.name, ''GOROD'' as DIR FROM GOROD as t7';
  try
    Query.Open;
  except
    begin
      MessageBox(FormMain.Handle, 'Ошибка при создании общей таблицы дерикторий.', 'Ошибка', MB_OK or MB_ICONERROR);
      Query.Close;
      Query.Free;
      exit;
    end;
  end;
  Query.FetchAll;

  with FormMain do
  begin
    XLS_Init;
    gaugeProgress.Visible := True;
    CurrentProgress := 1;
    gaugeProgress.MaxValue := Query.RecordCount;
    gaugeProgress.Progress := 1;
    SGDataUtils.OnDblClick := nil;
    SGDataUtils.BeginUpdate;

    for i := 1 to Query.RecordCount do
    begin
      isDeleted := false;
      strDirType := 'Неизвестно';

      strID := Query.FieldByName('ID').AsString;
      strDIR := Trim(Query.FieldByName('DIR').AsString);

      if strDIR = 'CURATOR' then
      begin
        if not CheckDirectoryIsInUse('CURATOR', '%#' + strID + '$%') then
        begin
          DeleteDirectoryByID('CURATOR', strID);
          isDeleted := true;
          strDirType := 'Куратор';
        end;
      end
      else if strDIR = 'RUBRIKATOR' then
      begin
        if not CheckDirectoryIsInUse('RUBR', '%#' + strID + '$%') then
        begin
          DeleteDirectoryByID('RUBRIKATOR', strID);
          isDeleted := true;
          strDirType := 'Рубрика';
        end;
      end
      else if strDIR = 'TYPE' then
      begin
        if not CheckDirectoryIsInUse('TYPE', '%#' + strID + '$%') then
        begin
          DeleteDirectoryByID('TYPE', strID);
          isDeleted := true;
          strDirType := 'Тип фирмы';
        end;
      end
      else if strDIR = 'NAPRAVLENIE' then
      begin
        if not CheckDirectoryIsInUse('NAPRAVLENIE', '%#' + strID + '$%') then
        begin
          DeleteDirectoryByID('NAPRAVLENIE', strID);
          isDeleted := true;
          strDirType := 'Деятельность';
        end;
      end
      else if strDIR = 'OFFICETYPE' then
      begin
        if not CheckDirectoryIsInUse('ADRES', '%#@' + strID + '$%') then
        begin
          DeleteDirectoryByID('OFFICETYPE', strID);
          isDeleted := true;
          strDirType := 'Тип адреса';
        end;
      end
      else if strDIR = 'COUNTRY' then
      begin
        if not CheckDirectoryIsInUse('ADRES', '%#&' + strID + '$%') then
        begin
          DeleteDirectoryByID('COUNTRY', strID);
          isDeleted := true;
          strDirType := 'Страна';
        end;
      end
      else if strDIR = 'GOROD' then
      begin
        if not CheckDirectoryIsInUse('ADRES', '%#^' + strID + '$%') then
        begin
          DeleteDirectoryByID('GOROD', strID);
          isDeleted := true;
          strDirType := 'Город';
        end;
      end;

      if isDeleted = true then
      begin
        SGDataUtils.AddRow;
        SGDataUtils.Cells[0, SGDataUtils.LastAddedRow] := strDirType;
        SGDataUtils.Cells[1, SGDataUtils.LastAddedRow] := Query.FieldByName('ID').AsString;
        SGDataUtils.Cells[2, SGDataUtils.LastAddedRow] := Query.FieldByName('NAME').AsString;

        with XLSDoc.Workbook.Sheets[0] do
        begin
          Cells[XLS_RowCounter, 0].Value := strDirType;
          XLS_SetStyle(XLS_RowCounter, 0);
          Cells[XLS_RowCounter, 1].Value := Query.FieldByName('ID').AsString;
          XLS_SetStyle(XLS_RowCounter, 1);
          Cells[XLS_RowCounter, 2].Value := Query.FieldByName('NAME').AsString;
          XLS_SetStyle(XLS_RowCounter, 2);

          Inc(XLS_RowCounter, 1);
        end;
      end;

      Inc(CurrentProgress, 1);
      gaugeProgress.Progress := CurrentProgress;
      Application.ProcessMessages;
      Query.Next;
    end;

    SGDataUtils.EndUpdate;
    gaugeProgress.Visible := False;
    XLS_ShowReport;
  end;

  Query.Close;
  Query.Free;
  FormMain.IBDatabase1.Connected := False;
  ToggleButtons(true);
end;

end.
