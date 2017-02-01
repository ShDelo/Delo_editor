unit Edit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, sPanel, StdCtrls, sCheckBox, sEdit, sGroupBox,
  Buttons, sSpeedButton, sLabel, NxScrollControl, NxCustomGridControl,
  NxCustomGrid, NxGrid;

type
  TFormEdit = class(TForm)
    sPanel1: TsPanel;
    gbBannerRight: TsGroupBox;
    editBannerRight: TsEdit;
    cbBannerRight: TsCheckBox;
    bntBannerRightDelete: TsSpeedButton;
    bntBannerRightBrowse: TsSpeedButton;
    gbReklamaText: TsGroupBox;
    bntReklamaTextDelete: TsSpeedButton;
    editReklamaText: TsEdit;
    cbReklamaText: TsCheckBox;
    gbBannerMain: TsGroupBox;
    btnBannerMainDelete: TsSpeedButton;
    btnBannerMainBrowse: TsSpeedButton;
    editBannerMain: TsEdit;
    cbBannerMain: TsCheckBox;
    gbReklamaSite: TsGroupBox;
    btnReklamaSite: TsSpeedButton;
    editReklamaSite: TsEdit;
    cbReklamaSite: TsCheckBox;
    gbBannerMainRubr: TsGroupBox;
    btnBannerMainRubrDelete: TsSpeedButton;
    btnBannerMainRubrBrowse: TsSpeedButton;
    editBannerMainRubr: TsEdit;
    cbBannerMainRubr: TsCheckBox;
    btnSave: TsSpeedButton;
    btnCancel: TsSpeedButton;
    editName: TsEdit;
    lblID: TsLabel;
    OpenDialog1: TOpenDialog;
    gbReklamaDoc: TsGroupBox;
    btnReklamaDocDelete: TsSpeedButton;
    editReklamaDoc: TsEdit;
    cbReklamaDoc: TsCheckBox;
    btnReklamaDocBrowse: TsSpeedButton;
    procedure btnCancelClick(Sender: TObject);
    procedure ClearEdits;
    procedure SaveFirmData(Sender: TObject);
    procedure SaveRubrData(Sender: TObject);
    procedure bntBannerRightDeleteClick(Sender: TObject);
    procedure bntBannerRightBrowseClick(Sender: TObject);
  private
    { Private declarations }
  public
    procedure CreateParams(var Params: TCreateParams); override;
    { Public declarations }
  end;

var
  FormEdit: TFormEdit;

implementation

uses Main, IBQuery;

{$R *.dfm}

procedure TFormEdit.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
  Params.ExStyle := Params.ExStyle or WS_Ex_AppWindow;
end;

procedure TFormEdit.ClearEdits;
begin
  lblID.Caption := '';
  editName.Text := '';
  editBannerRight.Text := '';
  editBannerMain.Text := '';
  editReklamaText.Text := '';
  editReklamaSite.Text := '';
  editReklamaDoc.Text := '';
  editBannerMainRubr.Text := '';
  cbBannerRight.Checked := False;
  cbBannerMain.Checked := False;
  cbReklamaText.Checked := False;
  cbReklamaSite.Checked := False;
  cbReklamaDoc.Checked := False;
  cbBannerMainRubr.Checked := False;
  gbBannerRight.Enabled := True;
  gbBannerMain.Enabled := True;
  gbReklamaText.Enabled := True;
  gbReklamaSite.Enabled := True;
  gbReklamaDoc.Enabled := True;
  gbBannerMainRubr.Enabled := True;
end;

procedure TFormEdit.btnCancelClick(Sender: TObject);
begin
  FormEdit.Close;
end;

procedure TFormEdit.bntBannerRightDeleteClick(Sender: TObject);
begin
  if TsSpeedButton(Sender).Name = 'bntBannerRightDelete' then
    editBannerRight.Text := '';
  if TsSpeedButton(Sender).Name = 'btnBannerMainDelete' then
    editBannerMain.Text := '';
  if TsSpeedButton(Sender).Name = 'bntReklamaTextDelete' then
    editReklamaText.Text := '';
  if TsSpeedButton(Sender).Name = 'btnReklamaSite' then
    editReklamaSite.Text := '';
  if TsSpeedButton(Sender).Name = 'btnReklamaDocDelete' then
    editReklamaDoc.Text := '';
  if TsSpeedButton(Sender).Name = 'btnBannerMainRubrDelete' then
    editBannerMainRubr.Text := '';
end;

procedure TFormEdit.bntBannerRightBrowseClick(Sender: TObject);
var
  Edit: TsEdit;
begin
  if TsSpeedButton(Sender).Name = 'btnReklamaDocBrowse' then
    OpenDialog1.Filter := 'Файл Microsoft Office или изображение|*.doc;*.jpg;*.bmp'
  else
    OpenDialog1.Filter := 'Файл Jpg или Bmp|*.jpg;*.bmp';
  Edit := nil;
  if TsSpeedButton(Sender).Name = 'bntBannerRightBrowse' then
    Edit := editBannerRight;
  if TsSpeedButton(Sender).Name = 'btnBannerMainBrowse' then
    Edit := editBannerMain;
  if TsSpeedButton(Sender).Name = 'btnBannerMainRubrBrowse' then
    Edit := editBannerMainRubr;
  if TsSpeedButton(Sender).Name = 'btnReklamaDocBrowse' then
    Edit := editReklamaDoc;
  if OpenDialog1.Execute then
    if Assigned(Edit) then
      Edit.Text := ExtractFileName(OpenDialog1.FileName);
end;

procedure TFormEdit.SaveFirmData(Sender: TObject);
var
  finalStr, tmpStr: string;
begin
  if Trim(editBannerRight.Text) = '' then
    cbBannerRight.Checked := False;
  if Trim(editBannerMain.Text) = '' then
    cbBannerMain.Checked := False;
  if Trim(editReklamaText.Text) = '' then
    cbReklamaText.Checked := False;
  if Trim(editReklamaSite.Text) = '' then
    cbReklamaSite.Checked := False;
  if Trim(editReklamaDoc.Text) = '' then
    cbReklamaDoc.Checked := False;
  finalStr := '';
  if cbBannerRight.Checked then
    tmpStr := '#bannerright=1<[>' + editBannerRight.Text + '<]>$'
  else
    tmpStr := '#bannerright=0<[>' + editBannerRight.Text + '<]>$';
  finalStr := finalStr + tmpStr;
  if cbBannerMain.Checked then
    tmpStr := '#bannermain=1<[>' + editBannerMain.Text + '<]>$'
  else
    tmpStr := '#bannermain=0<[>' + editBannerMain.Text + '<]>$';
  finalStr := finalStr + tmpStr;
  if cbReklamaText.Checked then
    tmpStr := '#text=1<[>' + editReklamaText.Text + '<]>$'
  else
    tmpStr := '#text=0<[>' + editReklamaText.Text + '<]>$';
  finalStr := finalStr + tmpStr;
  if cbReklamaSite.Checked then
    tmpStr := '#site=1<[>' + editReklamaSite.Text + '<]>$'
  else
    tmpStr := '#site=0<[>' + editReklamaSite.Text + '<]>$';
  finalStr := finalStr + tmpStr;
  if cbReklamaDoc.Checked then
    tmpStr := '#doc=1<[>' + editReklamaDoc.Text + '<]>$'
  else
    tmpStr := '#doc=0<[>' + editReklamaDoc.Text + '<]>$';
  finalStr := finalStr + tmpStr;
  with FormMain.IBQuery1 do
  begin
    Close;
    SQL.Text := 'update BASE set REKLAMA = :REKLAMA where ID = :ID';
    ParamByName('REKLAMA').AsString := finalStr;
    ParamByName('ID').AsString := lblID.Caption;
    ExecSQL;
    FormMain.IBTransaction1.CommitRetaining;
  end;
  if FormMain.SGFirm.FindText(0, lblID.Caption, [soCaseInsensitive, soExactMatch]) then
    FormMain.SGFirm.DeleteRow(FormMain.SGFirm.SelectedRow);
  FormMain.GetFirmList('select * from BASE where ID = :ID', lblID.Caption, False);
  FormMain.SGFirm.FindText(0, lblID.Caption, [soCaseInsensitive, soExactMatch]);
  FormEdit.Close;
end;

procedure TFormEdit.SaveRubrData(Sender: TObject);
var
  finalStr: string;
begin
  if Trim(editBannerMainRubr.Text) = '' then
    cbBannerMainRubr.Checked := False;
  if cbBannerMainRubr.Checked then
    finalStr := '#bannermainrubr=1<[>' + editBannerMainRubr.Text + '<]>$'
  else
    finalStr := '#bannermainrubr=0<[>' + editBannerMainRubr.Text + '<]>$';
  with FormMain.IBQuery1 do
  begin
    Close;
    SQL.Text := 'update RUBRIKATOR set REKLAMA = :REKLAMA where ID = :ID';
    ParamByName('REKLAMA').AsString := finalStr;
    ParamByName('ID').AsString := lblID.Caption;
    ExecSQL;
    FormMain.IBTransaction1.CommitRetaining;
  end;
  if FormMain.SGRubr.FindText(0, lblID.Caption, [soCaseInsensitive, soExactMatch]) then
    FormMain.SGRubr.DeleteRow(FormMain.SGRubr.SelectedRow);
  FormMain.GetRubrList('select * from RUBRIKATOR where ID = :ID', lblID.Caption, False);
  FormMain.SGRubr.FindText(0, lblID.Caption, [soCaseInsensitive, soExactMatch]);
  FormEdit.Close;
end;

end.
