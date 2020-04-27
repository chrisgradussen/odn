unit dataunit;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, ZConnection, ZDataset, ZStoredProcedure,
  ZSqlProcessor, ZSqlUpdate, db, InstellingenUnit, dialogs;

type

  { TDM }

  TDM = class(TDataModule)
    ZAfdelingAFDELING_ID: TLongintField;
    ZAfdelingOMSCHRIJVING: TStringField;
    ZAfdelingUrenAFDELING: TLongintField;
    ZAfdelingUrenJAARWEEK: TLongintField;
    ZAfdelingUrenNORMUREN: TFloatField;
    ZAfdelingUrenUREN: TFloatField;
    ZConnection: TZConnection;
    ZDaginfoNieuw: TZQuery;
    ZOmzetgegevensDelete: TZSQLProcessor;
    ZWagAdd: TZSQLProcessor;
    ZPMTInfo: TZQuery;
    ZJaarweek: TZQuery;
    ZAfdelinginfo: TZReadOnlyQuery;
    ZAfdelingUren: TZQuery;
    ZWeekInfoNieuw: TZReadOnlyQuery;
    ZVoorraadcorrectiesDelete: TZSQLProcessor;
    ZvoorraadcorrectiesQuery: TZQuery;
    ZVoorraadcorrectiesAdd: TZSQLProcessor;
    ZWeekinfo: TZQuery;
    ZOmzetgegevensAdd: TZSQLProcessor;
    ZAfdeling: TZQuery;
    ZDaginfo: TZQuery;
    procedure DataModuleCreate(Sender: TObject);
    procedure ZAfdelingUrenAFDELINGChange(Sender: TField);
    procedure ZAfdelingUrenAFDELINGGetText(Sender: TField; var aText: string;
      DisplayText: Boolean);
    procedure ZAfdelingUrenAFDELINGSetText(Sender: TField; const aText: string);
    procedure ZAfdelingUrenAFDELINGValidate(Sender: TField);
    procedure ZAfdelingUrenBeforePost(DataSet: TDataSet);
    procedure ZConnectionBeforeConnect(Sender: TObject);
    procedure ZJaarweekUrenAfterScroll(DataSet: TDataSet);
    procedure ZomzetgegevensStoredBeforeOpen(DataSet: TDataSet);
    procedure ZUpdateSQL1BeforeInsertSQL(Sender: TObject);
    procedure ZUpdateSQL1BeforeModifySQL(Sender: TObject);
    procedure ZWagAddAfterExecute(Processor: TZSQLProcessor;
      StatementIndex: Integer);
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

end;

procedure TDM.ZWagAddAfterExecute(Processor: TZSQLProcessor;
  StatementIndex: Integer);
begin

end;

procedure TDM.DataModuleCreate(Sender: TObject);
begin

end;

procedure TDM.ZAfdelingUrenAFDELINGChange(Sender: TField);
begin
  showmessage(' onchange : ' +sender.Text);
end;

procedure TDM.ZAfdelingUrenAFDELINGGetText(Sender: TField; var aText: string;
  DisplayText: Boolean);
begin
  aText := copy(atext,1,2);

end;

procedure TDM.ZAfdelingUrenAFDELINGSetText(Sender: TField; const aText: string);
begin
  showmessage (' on settext : ' + aText);
end;

procedure TDM.ZAfdelingUrenAFDELINGValidate(Sender: TField);
begin
  showmessage(' on validate : ' + sender.Text)
end;

procedure TDM.ZAfdelingUrenBeforePost(DataSet: TDataSet);
begin

//  dm.ZAfdelingUren.fieldByName('jaarweek').AsInteger := dm.ZJaarweek.FieldByName('jaarweek').AsInteger;
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

