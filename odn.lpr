program odn;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, laz_fpspreadsheet_visual, rxnew, tachartlazaruspkg,
  runtimetypeinfocontrols, zcomponent, mainunit, dataunit, InstellingenUnit,
  urenunit;

{$R *.res}

begin
  RequireDerivedFormResource := True;
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TFormInstellingen, FormInstellingen);
  Application.CreateForm(TDM, DM);
  Application.CreateForm(TForm2, Form2);
  Application.Run;
end.

