unit dataunit;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, ZConnection, ZDataset, ZStoredProcedure,
  ZSqlProcessor, ZSqlUpdate, db, InstellingenUnit, dialogs;

type

  { TDM }

  TDM = class(TDataModule)
    ZConnection: TZConnection;
    ZDaginfoNieuw: TZQuery;
    ZJaarweekUren: TZQuery;
    ZPMTInfo: TZQuery;
    ZJaarweek: TZQuery;
    ZAfdelinginfo: TZReadOnlyQuery;
    ZAfdelingUren: TZQuery;
    ZUpdateSQL1: TZUpdateSQL;
    ZWeekInfoNieuw: TZReadOnlyQuery;
    ZVoorraadcorrectiesDelete: TZSQLProcessor;
    ZvoorraadcorrectiesQuery: TZQuery;
    ZVoorraadcorrectiesAdd: TZSQLProcessor;
    ZWeekinfo: TZQuery;
    ZOmzetgegevensAdd: TZSQLProcessor;
    ZAfdeling: TZQuery;
    ZDaginfo: TZQuery;
    procedure DataModuleCreate(Sender: TObject);
    procedure ZAfdelingUrenBeforePost(DataSet: TDataSet);
    procedure ZConnectionBeforeConnect(Sender: TObject);
    procedure ZJaarweekUrenAfterScroll(DataSet: TDataSet);
    procedure ZomzetgegevensStoredBeforeOpen(DataSet: TDataSet);
    procedure ZUpdateSQL1BeforeInsertSQL(Sender: TObject);
    procedure ZUpdateSQL1BeforeModifySQL(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  DM: TDM;

implementation

{$R *.lfm}

{ TDM }

procedure TDM.ZomzetgegevensStoredBeforeOpen(DataSet: TDataSet);
begin

end;

procedure TDM.ZUpdateSQL1BeforeInsertSQL(Sender: TObject);
begin

end;

procedure TDM.ZUpdateSQL1BeforeModifySQL(Sender: TObject);
begin
  dm.ZAfdelingUren.ParamByName('jaarweek').AsInteger := dm.ZJaarweekUren.FieldByName('jaarweek').AsInteger;
end;

procedure TDM.DataModuleCreate(Sender: TObject);
begin

end;

procedure TDM.ZAfdelingUrenBeforePost(DataSet: TDataSet);
begin

  dm.ZAfdelingUren.fieldByName('jaarweek').AsInteger := dm.ZJaarweekUren.FieldByName('jaarweek').AsInteger;
  showmessage('before post : '+inttostr(dm.ZAfdelingUren.fieldByName('jaarweek').AsInteger));
end;

procedure TDM.ZConnectionBeforeConnect(Sender: TObject);
begin
  zConnection.Database:= Forminstellingen.EditDatabase.Text;
  zConnection.HostName:= FormInstellingen.EditHost.Text;
  zConnection.Password:= FormInstellingen.EditPassword.Text;
  zConnection.User:= FormInstellingen.EditUser.Text;
end;

procedure TDM.ZJaarweekUrenAfterScroll(DataSet: TDataSet);
begin

end;

end.

