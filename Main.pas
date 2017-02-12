unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Buttons, NxColumns, NxColumnClasses, NxScrollControl, StrUtils,
  NxCustomGridControl, NxCustomGrid, NxGrid, ComCtrls, sPageControl,
  StdCtrls, NxEdit, sSpeedButton, ExtCtrls, sPanel, sSkinManager, sEdit,
  sSkinProvider, sCheckBox, sGauge, DBAccess, IBC, MemDS, DB;

function QueryCreate: TIBCQuery;

type
  TFormMain = class(TForm)
    sPanel1: TsPanel;
    sPanel2: TsPanel;
    btnSettings: TsSpeedButton;
    editSearch: TNxEdit;
    sPageControl1: TsPageControl;
    sTabSheet1: TsTabSheet;
    SGFirm: TNextGrid;
    NxTextColumn1: TNxTextColumn;
    NxTextColumn2: TNxTextColumn;
    NxTextColumn3: TNxTextColumn;
    sSkinProvider1: TsSkinProvider;
    sSkinManager1: TsSkinManager;
    sTabSheet2: TsTabSheet;
    SGRubr: TNextGrid;
    NxTextColumn4: TNxTextColumn;
    NxTextColumn5: TNxTextColumn;
    NxTextColumn6: TNxTextColumn;
    NxTextColumn7: TNxTextColumn;
    NxTextColumn8: TNxTextColumn;
    NxTextColumn9: TNxTextColumn;
    NxTextColumn10: TNxTextColumn;
    sTabSheet3: TsTabSheet;
    btnImageExistingCheck: TsSpeedButton;
    gaugeProgress: TsGauge;
    SGDataUtils: TNextGrid;
    btnEmailCheck: TsSpeedButton;
    btnWEBCheck: TsSpeedButton;
    btnDBDefrag: TsSpeedButton;
    IBQuery1: TIBCQuery;
    IBDatabase1: TIBCConnection;
    IBTransaction1: TIBCTransaction;
    procedure FormCreate(Sender: TObject);
    procedure GetFirmList(Request, ID: string; ClearRows: Boolean);
    procedure GetRubrList(Request, ID: string; ClearRows: Boolean);
    procedure SGAfterSort(Sender: TObject; ACol: Integer);
    procedure editSearchChange(Sender: TObject);
    procedure sPageControl1Change(Sender: TObject);
    procedure btnSettingsClick(Sender: TObject);
    procedure btnImageExistingCheckClick(Sender: TObject);
    function IsValidEmail(const Value: string): Boolean;
    function IsValidWeb(const Value: string): Boolean;
    procedure EmailWebCheck(Sender: TObject);
    procedure btnDBDefragClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

const
  FBUserName: string = 'SYSDBA';
  FBUserPassword: string = 'masterkey';
  MainDB: string = 'usrdt.msq';

var
  FormMain: TFormMain;
  AppPath: string;

implementation

uses Edit, uDBDefrag;

{$R *.dfm}

function QueryCreate: TIBCQuery;
var
  Query: TIBCQuery;
begin
  Query := TIBCQuery.Create(nil);
  Query.Connection := FormMain.IBDatabase1;
  Query.Transaction := FormMain.IBTransaction1;
  Query.AutoCommit := False;
  Query.FetchRows := 1;
  result := Query;
end;

procedure TFormMain.FormCreate(Sender: TObject);
begin
  AppPath := ExtractFilePath(Application.ExeName);
  IBDatabase1.Database := AppPath + MainDB;
  IBDatabase1.Params.Clear;
  IBDatabase1.Params.Add('user_name=' + FBUserName);
  IBDatabase1.Params.Add('password=' + FBUserPassword);
  IBQuery1.Connection := IBDatabase1;
  IBQuery1.Transaction := IBTransaction1;
  IBTransaction1.DefaultConnection := IBDatabase1;
  try
    IBDatabase1.Connected := True;
  except
    begin
      MessageBox(handle, 'Ошибка при подключении файлов баз данных.', 'Ошибка', MB_OK or MB_ICONERROR);
      FormMain.Free;
      Halt;
    end;
  end;
  GetFirmList('select * from BASE', '', True);
  GetRubrList('select * from RUBRIKATOR', '', True);
end;

procedure TFormMain.GetFirmList(Request, ID: string; ClearRows: Boolean);
var
  i: integer;
  ReklamaStr, tmp: string;

  procedure StringParse(ReklamaType: string; Col: integer);
  begin
    if pos(ReklamaType, ReklamaStr) > 0 then
    begin
      tmp := ReklamaStr;
      delete(tmp, 1, pos(ReklamaType, ReklamaStr) + (Length(ReklamaType) - 1));
      if tmp[1] = '0' then
        SGFirm.Cell[Col, SGFirm.LastAddedRow].Color := $008888FF;
      if tmp[1] = '1' then
        SGFirm.Cell[Col, SGFirm.LastAddedRow].Color := $00A6FFA6;
      delete(tmp, 1, 1);
      delete(tmp, pos('$', tmp), length(tmp));
      tmp := AnsiReplaceStr(tmp, '<[>', '');
      tmp := AnsiReplaceStr(tmp, '<]>', '');
      SGFirm.Cells[Col, SGFirm.LastAddedRow] := tmp;
    end;
  end;

begin
  if ClearRows then
    SGFirm.ClearRows;
  IBQuery1.Close;
  IBQuery1.SQL.Text := Request;;
  if IBQuery1.ParamCount > 0 then
    IBQuery1.Params[0].AsString := ID;
  IBQuery1.Open;
  IBQuery1.FetchAll := True;
  if IBQuery1.RecordCount = 0 then
    exit;
  SGFirm.BeginUpdate;
  for i := 1 to IBQuery1.RecordCount do
  begin
    SGFirm.AddRow;
    SGFirm.Cells[0, SGFirm.LastAddedRow] := IBQuery1.FieldByName('ID').AsString;
    if IBQuery1.FieldByName('ACTIVITY').AsInteger = 0 then
      SGFirm.Cell[1, SGFirm.LastAddedRow].Color := $008888FF;
    SGFirm.Cells[1, SGFirm.LastAddedRow] := IBQuery1.FieldByName('NAME').AsString;
    if IBQuery1.FieldValues['REKLAMA'] <> null then
    begin
      ReklamaStr := IBQuery1.FieldByName('REKLAMA').AsString;
      StringParse('bannerright=', 2);
      StringParse('bannermain=', 3);
      StringParse('text=', 4);
      StringParse('site=', 5);
      StringParse('doc=', 6);
    end;
    IBQuery1.Next;
  end;
  SGFirm.Resort;
  SGFirm.EndUpdate;
  IBQuery1.Close;
  IBDatabase1.Close;
end;

procedure TFormMain.GetRubrList(Request, ID: string; ClearRows: Boolean);
var
  i: integer;
  ReklamaStr, tmp: string;

  procedure StringParse(ReklamaType: string; Col: integer);
  begin
    if pos(ReklamaType, ReklamaStr) > 0 then
    begin
      tmp := ReklamaStr;
      delete(tmp, 1, pos(ReklamaType, ReklamaStr) + (Length(ReklamaType) - 1));
      if tmp[1] = '0' then
        SGRubr.Cell[Col, SGRubr.LastAddedRow].Color := $008888FF;
      if tmp[1] = '1' then
        SGRubr.Cell[Col, SGRubr.LastAddedRow].Color := $00A6FFA6;
      delete(tmp, 1, 1);
      delete(tmp, pos('$', tmp), length(tmp));
      tmp := AnsiReplaceStr(tmp, '<[>', '');
      tmp := AnsiReplaceStr(tmp, '<]>', '');
      SGRubr.Cells[Col, SGRubr.LastAddedRow] := tmp;
    end;
  end;

begin
  if ClearRows then
    SGRubr.ClearRows;
  IBQuery1.Close;
  IBQuery1.SQL.Text := Request;
  if IBQuery1.ParamCount > 0 then
    IBQuery1.Params[0].AsString := ID;
  IBQuery1.Open;
  IBQuery1.FetchAll := True;
  if IBQuery1.RecordCount = 0 then
    exit;
  SGRubr.BeginUpdate;
  for i := 1 to IBQuery1.RecordCount do
  begin
    SGRubr.AddRow;
    SGRubr.Cells[0, SGRubr.LastAddedRow] := IBQuery1.FieldByName('ID').AsString;
    SGRubr.Cells[1, SGRubr.LastAddedRow] := IBQuery1.FieldByName('NAME').AsString;
    if IBQuery1.FieldValues['REKLAMA'] <> null then
    begin
      ReklamaStr := IBQuery1.FieldByName('REKLAMA').AsString;
      StringParse('bannermainrubr=', 2);
    end;
    IBQuery1.Next;
  end;
  SGRubr.Resort;
  SGRubr.EndUpdate;
  IBQuery1.Close;
  IBDatabase1.Close;
end;

procedure TFormMain.editSearchChange(Sender: TObject);
var
  s, t: string;
  RowVisible: Boolean;
  i: integer;
  SG: TNextGrid;
begin
  SG := nil;
  if sPageControl1.ActivePageIndex = 0 then
    SG := SGFirm;
  if sPageControl1.ActivePageIndex = 1 then
    SG := SGRubr;
  if sPageControl1.ActivePageIndex = 2 then
    exit;
  SG.BeginUpdate;
  s := AnsiUpperCase(editSearch.Text);
  for i := 0 to SG.RowCount - 1 do
  begin
    t := AnsiUpperCase(SG.Cells[1, i]);
    RowVisible := (s = '') or (pos(s, t) > 0);
    SG.RowVisible[i] := RowVisible;
  end;
  SG.EndUpdate;
  SG.Resort;
  SG.SelectFirstRow;
end;

procedure TFormMain.sPageControl1Change(Sender: TObject);
begin
  editSearch.Clear;
  case sPageControl1.ActivePageIndex of
    0:
      btnSettings.Enabled := True;
    1:
      btnSettings.Enabled := True;
    2:
      btnSettings.Enabled := False;
  end;
end;

procedure TFormMain.SGAfterSort(Sender: TObject; ACol: Integer);
var
  i, n: integer;
  C: TColor;
begin
  with Sender as TNextGrid do
  begin
    BeginUpdate;
    n := 0;
    for i := 1 to RowCount do
    begin
      if RowVisible[i - 1] then
        inc(n, 1);
      if Odd(n) then
        C := $00EDE9EB { clMenuBar }
      else
        C := clWindow;
      if name = 'SGFirm' then
      begin
        Cell[0, i - 1].Color := C;
        if Cell[1, i - 1].Color <> $008888FF then
          Cell[1, i - 1].Color := C;
        if Trim(Cells[2, i - 1]) = '' then
          Cell[2, i - 1].Color := C;
        if Trim(Cells[3, i - 1]) = '' then
          Cell[3, i - 1].Color := C;
        if Trim(Cells[4, i - 1]) = '' then
          Cell[4, i - 1].Color := C;
        if Trim(Cells[5, i - 1]) = '' then
          Cell[5, i - 1].Color := C;
        if Trim(Cells[6, i - 1]) = '' then
          Cell[6, i - 1].Color := C;
      end;
      if name = 'SGRubr' then
      begin
        Cell[0, i - 1].Color := C;
        Cell[1, i - 1].Color := C;
        if Trim(Cells[2, i - 1]) = '' then
          Cell[2, i - 1].Color := C;
      end;
    end;
    EndUpdate;
  end;
end;

procedure TFormMain.btnSettingsClick(Sender: TObject);
var
  ID, ReklamaStr, tmp: string;
  isFirm: Boolean;

  procedure StringParse(ReklamaType: string; Edit: TsEdit; CB: TsCheckBox);
  begin
    if pos(ReklamaType, ReklamaStr) > 0 then
    begin
      tmp := ReklamaStr;
      delete(tmp, 1, pos(ReklamaType, ReklamaStr) + (Length(ReklamaType) - 1));
      if tmp[1] = '0' then
        CB.Checked := False;
      if tmp[1] = '1' then
        CB.Checked := True;
      delete(tmp, 1, 1);
      delete(tmp, pos('$', tmp), length(tmp));
      tmp := AnsiReplaceStr(tmp, '<[>', '');
      tmp := AnsiReplaceStr(tmp, '<]>', '');
      Edit.Text := tmp;
    end;
  end;

begin
  if sPageControl1.ActivePageIndex = 0 then // ФИРМА
  begin
    if SGFirm.SelectedCount = 0 then
      exit;
    ID := SGFirm.Cells[0, SGFirm.SelectedRow];
    if Trim(ID) = '' then
      exit;
    IBQuery1.Close;
    IBQuery1.SQL.Text := 'select * from BASE where ID = :ID';
    IBQuery1.Params[0].AsString := ID;
    IBQuery1.Open;
    IBQuery1.FetchAll := True;
    FormEdit.ClearEdits;
    FormEdit.lblID.Caption := IBQuery1.FieldByName('ID').AsString;
    FormEdit.editName.Text := IBQuery1.FieldByName('NAME').AsString;
    ReklamaStr := '';
    if IBQuery1.FieldValues['REKLAMA'] <> null then
    begin
      ReklamaStr := IBQuery1.FieldByName('REKLAMA').AsString;
      StringParse('bannerright=', FormEdit.editBannerRight, FormEdit.cbBannerRight);
      StringParse('bannermain=', FormEdit.editBannerMain, FormEdit.cbBannerMain);
      StringParse('text=', FormEdit.editReklamaText, FormEdit.cbReklamaText);
      StringParse('site=', FormEdit.editReklamaSite, FormEdit.cbReklamaSite);
      StringParse('doc=', FormEdit.editReklamaDoc, FormEdit.cbReklamaDoc);
    end;
    IBQuery1.Close;
    IBDatabase1.Close;
    FormEdit.gbBannerRight.Enabled := True;
    FormEdit.gbBannerMain.Enabled := True;
    FormEdit.gbReklamaText.Enabled := True;
    FormEdit.gbReklamaSite.Enabled := True;
    FormEdit.gbReklamaDoc.Enabled := True;
    FormEdit.gbBannerMainRubr.Enabled := False;
    FormEdit.btnSave.OnClick := FormEdit.SaveFirmData;
    FormEdit.Show;
  end;
  if sPageControl1.ActivePageIndex = 1 then // РУБРИКА
  begin
    if SGRubr.SelectedCount = 0 then
      exit;
    ID := SGRubr.Cells[0, SGRubr.SelectedRow];
    if Trim(ID) = '' then
      exit;
    IBQuery1.Close;
    IBQuery1.SQL.Text := 'select * from RUBRIKATOR where ID = :ID';
    IBQuery1.Params[0].AsString := ID;
    IBQuery1.Open;
    IBQuery1.FetchAll := True;
    FormEdit.ClearEdits;
    FormEdit.lblID.Caption := IBQuery1.FieldByName('ID').AsString;
    FormEdit.editName.Text := IBQuery1.FieldByName('NAME').AsString;
    ReklamaStr := '';
    if IBQuery1.FieldValues['REKLAMA'] <> null then
    begin
      ReklamaStr := IBQuery1.FieldByName('REKLAMA').AsString;
      StringParse('bannermainrubr=', FormEdit.editBannerMainRubr, FormEdit.cbBannerMainRubr);
    end;
    IBQuery1.Close;
    IBDatabase1.Close;
    FormEdit.gbBannerRight.Enabled := False;
    FormEdit.gbBannerMain.Enabled := False;
    FormEdit.gbReklamaText.Enabled := False;
    FormEdit.gbReklamaSite.Enabled := False;
    FormEdit.gbReklamaDoc.Enabled := False;
    FormEdit.gbBannerMainRubr.Enabled := True;
    FormEdit.btnSave.OnClick := FormEdit.SaveRubrData;
    FormEdit.Show;
  end;
  if sPageControl1.ActivePageIndex = 2 then // РАБОТА С БАЗОЙ
  begin
    if SGDataUtils.SelectedCount = 0 then
      exit;
    if SGDataUtils.Columns[0].Header.Caption = uDBDefrag.strHeader0 then
      exit;
    ID := SGDataUtils.Cells[1, SGDataUtils.SelectedRow];
    if Trim(ID) = '' then
      exit;
    if SGDataUtils.Cells[0, SGDataUtils.SelectedRow] = 'Фирма' then
      isFirm := True
    else
      isFirm := False;
    IBQuery1.Close;
    if isFirm then
      IBQuery1.SQL.Text := 'select * from BASE where ID = :ID'
    else
      IBQuery1.SQL.Text := 'select * from RUBRIKATOR where ID = :ID';
    IBQuery1.Params[0].AsString := ID;
    IBQuery1.Open;
    IBQuery1.FetchAll := True;
    FormEdit.ClearEdits;
    FormEdit.lblID.Caption := IBQuery1.FieldByName('ID').AsString;
    FormEdit.editName.Text := IBQuery1.FieldByName('NAME').AsString;
    ReklamaStr := '';
    if isFirm then
    begin
      if IBQuery1.FieldValues['REKLAMA'] <> null then
      begin
        ReklamaStr := IBQuery1.FieldByName('REKLAMA').AsString;
        StringParse('bannerright=', FormEdit.editBannerRight, FormEdit.cbBannerRight);
        StringParse('bannermain=', FormEdit.editBannerMain, FormEdit.cbBannerMain);
        StringParse('text=', FormEdit.editReklamaText, FormEdit.cbReklamaText);
        StringParse('site=', FormEdit.editReklamaSite, FormEdit.cbReklamaSite);
        StringParse('doc=', FormEdit.editReklamaDoc, FormEdit.cbReklamaDoc);
      end;
      IBQuery1.Close;
      IBDatabase1.Close;
      FormEdit.gbBannerRight.Enabled := True;
      FormEdit.gbBannerMain.Enabled := True;
      FormEdit.gbReklamaText.Enabled := True;
      FormEdit.gbReklamaSite.Enabled := True;
      FormEdit.gbReklamaDoc.Enabled := True;
      FormEdit.gbBannerMainRubr.Enabled := False;
      FormEdit.btnSave.OnClick := FormEdit.SaveFirmData;
      FormEdit.Show;
    end
    else
    begin
      if IBQuery1.FieldValues['REKLAMA'] <> null then
      begin
        ReklamaStr := IBQuery1.FieldByName('REKLAMA').AsString;
        StringParse('bannermainrubr=', FormEdit.editBannerMainRubr, FormEdit.cbBannerMainRubr);
      end;
      IBQuery1.Close;
      IBDatabase1.Close;
      FormEdit.gbBannerRight.Enabled := False;
      FormEdit.gbBannerMain.Enabled := False;
      FormEdit.gbReklamaText.Enabled := False;
      FormEdit.gbReklamaSite.Enabled := False;
      FormEdit.gbReklamaDoc.Enabled := False;
      FormEdit.gbBannerMainRubr.Enabled := True;
      FormEdit.btnSave.OnClick := FormEdit.SaveRubrData;
      FormEdit.Show;
    end;
  end;
end;

procedure TFormMain.btnImageExistingCheckClick(Sender: TObject);
var
  MaxProgress, CurrentProgress, i: Integer;
  banner_Right, banner_Main, reklama_Doc, banner_Rubr: string;
begin
  if not DirectoryExists(AppPath + 'Pic') then
  begin
    MessageBox(handle, 'Не удалось найти директорию "Pic"', 'Ошибка', MB_OK or MB_ICONERROR);
    exit;
  end;
  if not DirectoryExists(AppPath + 'Doc') then
  begin
    MessageBox(handle, 'Не удалось найти директорию "Doc"', 'Ошибка', MB_OK or MB_ICONERROR);
    exit;
  end;
  SGDataUtils.BeginUpdate;
  SGDataUtils.ClearRows;
  for i := SGDataUtils.Columns.Count - 1 downto 0 do
    SGDataUtils.Columns[i].Free;
  SGDataUtils.Columns.Add(TNxTextColumn, 'Раздел');
  SGDataUtils.Columns[0].Position := 0;
  SGDataUtils.Columns[0].Width := 80;
  SGDataUtils.Columns.Add(TNxTextColumn, 'ID');
  SGDataUtils.Columns[1].Position := 1;
  SGDataUtils.Columns[1].Width := 40;
  SGDataUtils.Columns[1].Header.Glyph := NxTextColumn1.Header.Glyph;
  SGDataUtils.Columns[1].Header.DisplayMode := dmTextAndImage;
  SGDataUtils.Columns.Add(TNxTextColumn, 'Название');
  SGDataUtils.Columns[2].Position := 2;
  SGDataUtils.Columns[2].Width := 300;
  SGDataUtils.Columns[2].Header.Glyph := NxTextColumn2.Header.Glyph;
  SGDataUtils.Columns[2].Header.DisplayMode := dmTextAndImage;
  SGDataUtils.Columns.Add(TNxTextColumn, 'Баннер (ротация)');
  SGDataUtils.Columns[3].Position := 3;
  SGDataUtils.Columns[3].Width := 145;
  SGDataUtils.Columns.Add(TNxTextColumn, 'Баннер (основной)');
  SGDataUtils.Columns[4].Position := 4;
  SGDataUtils.Columns[4].Width := 145;
  SGDataUtils.Columns.Add(TNxTextColumn, 'Статья');
  SGDataUtils.Columns[5].Position := 5;
  SGDataUtils.Columns[5].Width := 145;
  gaugeProgress.Visible := True;
  CurrentProgress := 0;
  gaugeProgress.Progress := 0;
  MaxProgress := SGFirm.RowCount + SGRubr.RowCount;
  gaugeProgress.MaxValue := MaxProgress;
  for i := 0 to SGFirm.RowCount - 1 do
  begin
    banner_Right := Trim(SGFirm.Cells[2, i]);
    banner_Main := Trim(SGFirm.Cells[3, i]);
    reklama_Doc := Trim(SGFirm.Cells[6, i]);
    if ((banner_Right <> '') and (not FileExists(AppPath + 'Pic\' + banner_Right))) or
      ((banner_Main <> '') and (not FileExists(AppPath + 'Pic\' + banner_Main))) or
      ((reklama_Doc <> '') and (not FileExists(AppPath + 'Doc\' + reklama_Doc))) then
    begin
      SGDataUtils.AddRow;
      SGDataUtils.Cells[0, SGDataUtils.LastAddedRow] := 'Фирма';
      SGDataUtils.Cells[1, SGDataUtils.LastAddedRow] := SGFirm.Cells[0, i];
      SGDataUtils.Cells[2, SGDataUtils.LastAddedRow] := SGFirm.Cells[1, i];
      if ((banner_Right <> '') and (not FileExists(AppPath + 'Pic\' + banner_Right))) then
        SGDataUtils.Cells[3, SGDataUtils.LastAddedRow] := SGFirm.Cells[2, i];
      if ((banner_Main <> '') and (not FileExists(AppPath + 'Pic\' + banner_Main))) then
        SGDataUtils.Cells[4, SGDataUtils.LastAddedRow] := SGFirm.Cells[3, i];
      if ((reklama_Doc <> '') and (not FileExists(AppPath + 'Doc\' + reklama_Doc))) then
        SGDataUtils.Cells[5, SGDataUtils.LastAddedRow] := SGFirm.Cells[6, i];
    end;
    Inc(CurrentProgress, 1);
    gaugeProgress.Progress := CurrentProgress;
    Application.ProcessMessages;
  end;
  for i := 0 to SGRubr.RowCount - 1 do
  begin
    banner_Rubr := Trim(SGRubr.Cells[2, i]);
    if ((banner_Rubr <> '') and not(FileExists(AppPath + 'Pic\' + banner_Rubr))) then
    begin
      SGDataUtils.AddRow;
      SGDataUtils.Cells[0, SGDataUtils.LastAddedRow] := 'Рубрика';
      SGDataUtils.Cells[1, SGDataUtils.LastAddedRow] := SGRubr.Cells[0, i];
      SGDataUtils.Cells[2, SGDataUtils.LastAddedRow] := SGRubr.Cells[1, i];
      SGDataUtils.Cells[4, SGDataUtils.LastAddedRow] := SGRubr.Cells[2, i];
    end;
    Inc(CurrentProgress, 1);
    gaugeProgress.Progress := CurrentProgress;
    Application.ProcessMessages;
  end;
  gaugeProgress.Visible := False;
  SGDataUtils.EndUpdate;
  SGDataUtils.OnDblClick := btnSettingsClick;
end;

function TFormMain.IsValidEmail(const Value: string): Boolean;

  function CheckAllowed(const s: string): Boolean;
  var
    i: integer;
  begin
    Result := false;
    for i := 1 to Length(s) do
    begin
      if not(s[i] in ['a' .. 'z', 'A' .. 'Z', '0' .. '9', '_', '-', '.', '&']) then
        exit;
    end;
    Result := true;
  end;

var
  i: integer;
  namePart, serverPart: string;
begin
  Result := false;
  i := Pos('@', Value);
  if i = 0 then
    exit;
  namePart := Copy(Value, 1, i - 1);
  serverPart := Copy(Value, i + 1, Length(Value));
  if (Length(namePart) = 0) or ((Length(serverPart) < 4)) then
    exit;
  i := Pos('.', serverPart);
  if (i = 0) or (i > (Length(serverPart) - 2)) then
    exit;
  Result := CheckAllowed(namePart) and CheckAllowed(serverPart);
end;

function TFormMain.IsValidWeb(const Value: string): Boolean;

  function CheckAllowed(const s: string): Boolean;
  var
    i: integer;
  begin
    Result := false;
    for i := 1 to Length(s) do
    begin
      if not(s[i] in ['a' .. 'z', 'A' .. 'Z', '0' .. '9', '_', '-', '.']) then
        exit;
    end;
    Result := true;
  end;

var
  i: integer;
  wwwPart, domenPart: string;
begin
  Result := False;
  i := Pos('.', Value);
  if i = 0 then
    exit;
  wwwPart := Copy(Value, 1, i - 1);
  domenPart := Copy(Value, i + 1, Length(Value));
  if (AnsiLowerCase(wwwPart) <> 'www') or ((Length(domenPart) < 5)) then
    exit;
  i := Pos('.', domenPart);
  if (i = 0) or (i > (Length(domenPart) - 2)) then
    exit;
  Result := CheckAllowed(wwwPart) and CheckAllowed(domenPart);
end;

procedure TFormMain.EmailWebCheck(Sender: TObject);
var
  str, tmp: string;
  i: integer;
begin
  SGDataUtils.BeginUpdate;
  SGDataUtils.ClearRows;
  for i := SGDataUtils.Columns.Count - 1 downto 0 do
    SGDataUtils.Columns[i].Free;
  SGDataUtils.Columns.Add(TNxTextColumn, 'ID');
  SGDataUtils.Columns[0].Position := 0;
  SGDataUtils.Columns[0].Width := 40;
  SGDataUtils.Columns[0].Header.Glyph := NxTextColumn1.Header.Glyph;
  SGDataUtils.Columns[0].Header.DisplayMode := dmTextAndImage;
  SGDataUtils.Columns.Add(TNxTextColumn, 'Название');
  SGDataUtils.Columns[1].Position := 1;
  SGDataUtils.Columns[1].Width := 300;
  SGDataUtils.Columns[1].Header.Glyph := NxTextColumn2.Header.Glyph;
  SGDataUtils.Columns[1].Header.DisplayMode := dmTextAndImage;
  if TsSpeedButton(Sender).Name = 'btnEmailCheck' then
    SGDataUtils.Columns.Add(TNxTextColumn, 'Email');
  if TsSpeedButton(Sender).Name = 'btnWEBCheck' then
    SGDataUtils.Columns.Add(TNxTextColumn, 'Сайт (WEB)');
  SGDataUtils.Columns[2].Position := 2;
  SGDataUtils.Columns[2].Width := 300;
  IBQuery1.Close;
  if TsSpeedButton(Sender).Name = 'btnEmailCheck' then
    IBQuery1.SQL.Text := 'select ID,NAME,EMAIL from BASE';
  if TsSpeedButton(Sender).Name = 'btnWEBCheck' then
    IBQuery1.SQL.Text := 'select ID,NAME,WEB from BASE';
  IBQuery1.Open;
  IBQuery1.FetchAll := True;
  gaugeProgress.MinValue := 0;
  gaugeProgress.MaxValue := IBQuery1.RecordCount;
  gaugeProgress.Visible := IBQuery1.RecordCount > 0;
  for i := 0 to IBQuery1.RecordCount do
  begin
    str := IBQuery1.Fields[2].AsString;
    if Length(Trim(str)) > 0 then
    begin
      if str[Length(str)] <> ',' then
        str := str + ',';
      while pos(',', str) > 0 do
      begin
        tmp := copy(str, 0, pos(',', str));
        delete(str, 1, length(tmp));
        tmp := Trim(tmp);
        delete(tmp, length(tmp), 1);
        if ((TsSpeedButton(Sender).Name = 'btnEmailCheck') and (not IsValidEmail(tmp))) then
        begin
          SGDataUtils.AddRow;
          SGDataUtils.Cells[0, SGDataUtils.LastAddedRow] := IBQuery1.FieldByName('ID').AsString;
          SGDataUtils.Cells[1, SGDataUtils.LastAddedRow] := IBQuery1.FieldByName('NAME').AsString;
          SGDataUtils.Cells[2, SGDataUtils.LastAddedRow] := tmp;
        end;
        if ((TsSpeedButton(Sender).Name = 'btnWEBCheck') and (not IsValidWeb(tmp))) then
        begin
          SGDataUtils.AddRow;
          SGDataUtils.Cells[0, SGDataUtils.LastAddedRow] := IBQuery1.FieldByName('ID').AsString;
          SGDataUtils.Cells[1, SGDataUtils.LastAddedRow] := IBQuery1.FieldByName('NAME').AsString;
          SGDataUtils.Cells[2, SGDataUtils.LastAddedRow] := tmp;
        end;
      end;
    end;
    gaugeProgress.Progress := i;
    Application.ProcessMessages;
    IBQuery1.Next;
  end;
  gaugeProgress.Visible := False;
  SGDataUtils.EndUpdate;
  IBQuery1.Close;
  IBDatabase1.Close;
  SGDataUtils.OnDblClick := nil;
end;

procedure TFormMain.btnDBDefragClick(Sender: TObject);
begin
  DoDBDefrag;
end;

end.
