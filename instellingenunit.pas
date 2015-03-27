unit InstellingenUnit;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, RTTICtrls, Forms, Controls, Graphics, Dialogs,
  ExtCtrls, StdCtrls, IniPropStorage;

type

  { TFormInstellingen }

  TFormInstellingen = class(TForm)
    Button1: TButton;
    EditDatabase: TEdit;
    EditHost: TEdit;
    EditUser: TEdit;
    EditPassword: TEdit;
    EditDagInfoRapport: TEdit;
    EditWeekInfoRapport: TEdit;
    EditPMTRapport: TEdit;
    EditAfdelingRapport: TEdit;
    IniPropStorage1: TIniPropStorage;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Panel1: TPanel;
    Panel2: TPanel;
    procedure Button1Click(Sender: TObject);
    procedure EditWeekInfoRapportChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  FormInstellingen: TFormInstellingen;

implementation

{$R *.lfm}

{ TFormInstellingen }

procedure TFormInstellingen.Button1Click(Sender: TObject);
begin
  close;
end;

procedure TFormInstellingen.EditWeekInfoRapportChange(Sender: TObject);
begin

end;

procedure TFormInstellingen.FormClose(Sender: TObject;
  var CloseAction: TCloseAction);
begin

end;

procedure TFormInstellingen.FormCreate(Sender: TObject);
begin

end;

procedure TFormInstellingen.FormDestroy(Sender: TObject);
begin

end;

procedure TFormInstellingen.FormShow(Sender: TObject);
begin

end;

end.

