unit urenunit;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, db, FileUtil, rxdbgrid, rxmemds, Forms, Controls, Graphics,
  Dialogs, ExtCtrls, DBGrids, dataunit;

type

  { TForm2 }

  TForm2 = class(TForm)
    DataSource2: TDataSource;
    DatasourceJaarweek: TDataSource;
    DBGrid1: TDBGrid;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    RxDBGrid1: TRxDBGrid;
    RxDBGrid2: TRxDBGrid;
    procedure DataSource1DataChange(Sender: TObject; Field: TField);
    procedure DataSource1StateChange(Sender: TObject);
    procedure DataSource1UpdateData(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure RxDBGrid2AfterQuickSearch(Sender: TObject; Field: TField;
      var AValue: string);
    procedure TRxColumnEditButtons1Click(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  Form2: TForm2;

implementation

{$R *.lfm}

{ TForm2 }

procedure TForm2.RxDBGrid2AfterQuickSearch(Sender: TObject; Field: TField;
  var AValue: string);
begin

end;

procedure TForm2.TRxColumnEditButtons1Click(Sender: TObject);
begin
  showmessage(' on click' );
end;

procedure TForm2.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin

   dataunit.DM.ZAfdelingUren.ApplyUpdates;
  dataunit.DM.ZConnection.Commit;



end;

procedure TForm2.DataSource1DataChange(Sender: TObject; Field: TField);
begin
  if dm.ZAfdelingUren.Active then
  begin
  //showmessage('afdelinguren : '+ inttostr(dm.ZAfdelingUren.ParamByName('jaarweek').AsInteger));
 // dm.ZAfdelingUren.ParamByName('jaarweek').AsInteger := dm.ZJaarweekUren.FieldByName('jaarweek').AsInteger;
  //dm.ZAfdelingUren.Refresh;
  //showmessage('afdelinguren : '+ inttostr(dm.ZAfdelingUren.ParamByName('jaarweek').AsInteger));
  end;

end;

procedure TForm2.DataSource1StateChange(Sender: TObject);
begin
  showmessage('state change');
end;

procedure TForm2.DataSource1UpdateData(Sender: TObject);
begin
  showmessage('updatedata');
end;

end.

