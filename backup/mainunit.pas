unit mainunit;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, TADbSource, TAGraph, LR_Desgn, LR_Class,
  LR_DBSet, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls, Buttons,
  Menus, fpspreadsheet, xlsbiff8, xlsbiff5, laz_fpspreadsheet, fpspreadsheetgrid,
  fpspreadsheetctrls, rxdbgrid, dataunit, db, urenunit,fpsallformats,fpstypes, vinfo;



type

  { TForm1 }

  TForm1 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    R10Button: TButton;
    DatasourceAfdelingInfo: TDatasource;
    DatasourceDaginfoNieuw: TDataSource;
    DatasourceWeekinfoNieuw: TDatasource;
    frDBDataSetAfdelingInfo: TfrDBDataSet;
    frDBDatasetDaginfoNieuw: TfrDBDataSet;
    frDBDataSetWeekinfoNieuw: TfrDBDataSet;
    HfdMenu: TMainMenu;
    MenuInstellingen: TMenuItem;
    MenuUren: TMenuItem;
    sWorkbook: TsWorkbookSource;
    Daginfo: TfrReport;
    PMTinfo: TfrReport;
    WeekinfoNieuwButton: TButton;
    DatasourcePMTinfo: TDatasource;
    DatasourceWeekinfo: TDatasource;
    DatasourceAfdeling: TDatasource;
    DatasourceJaarweek: TDatasource;
    DatasourceDaginfo: TDatasource;
    frDBDatasetPMTinfo: TfrDBDataSet;
    frDBDataSetWeekinfo: TfrDBDataSet;
    frDBDataSetAfdeling: TfrDBDataSet;
    frDBDataSetJaarweek: TfrDBDataSet;
    frDBDatasetDaginfo: TfrDBDataSet;
    Weekinfo: TfrReport;
    DervingInfoButten: TButton;
    WeekinfoGrid: TRxDBGrid;
    WeekInfoButton: TButton;
    buttonomzet: TButton;
    Memo: TMemo;
    Omzetdialoog: TOpenDialog;
    Panel1: TPanel;
    Panel2: TPanel;
    OmzetDirectoryDialoog: TSelectDirectoryDialog;
    sOmzetGrid : TsWorkSheetgrid;
    DaginfoNieuwButton1: TButton;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MenuInstellingenClick(Sender: TObject);
    procedure DaginfoNieuwButton1Click(Sender: TObject);
    procedure R10ButtonClick(Sender: TObject);
    procedure WeekinfoGridAfterQuickSearch(Sender: TObject; Field: TField;
      var AValue: string);
    procedure WeekinfoNieuwButtonClick(Sender: TObject);
    procedure buttonomzetClick(Sender: TObject);
    procedure DervingInfoButtenClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure MenuUrenClick(Sender: TObject);
    procedure WeekInfoButtonClick(Sender: TObject);
    procedure Panel2Click(Sender: TObject);
    procedure sOmzetGridClick(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
    procedure leesomzetspreadsheet(filename : string);
    procedure leesomzetR10(filename : string);
    procedure leesOPE1010;
    procedure leesomzetopdatum;
  end;

var
  Form1: TForm1;

implementation

uses instellingenunit;

{$R *.lfm}

{ TForm1 }

{ Leesomzetopdatum
R10 OPE1010 omzet op presentatiegroep, drillen op datum, exporteren als flat excel
}

procedure TForm1.leesomzetopdatum;
var
   datumstr        : string;
   mindatum        : tdatetime;
   maxdatum        : tdatetime;
   wagnummer       : string;
   wagomschrijving : string;
   wagomzet        : array[1..7] of string;
   i,y             : integer;
   omzetgroep      : string;
   l_decimalseparator : char;
   datumstring     : string;
   waarde          : string;

begin
  // eerst alle oude datums verwijderen
  i := 3;
  mindatum := 0;
  maxdatum := 0;
  while assigned(somzetgrid.worksheet.findcell(i,0)) do
  begin
    if  (pos('Totaal' ,somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue)= 1) then
    begin
      if ((mindatum <> 0) and (maxdatum <> 0)) then
      begin
        //data verwijderen....
       // showmessage ('mindatum : ' + datetimetostr(mindatum) + ' maxdatum : ' +datetimetostr(maxdatum));
        dm.ZOmzetgegevensDelete.ParamByName('mindatum').AsDate:= mindatum;
        dm.ZOmzetgegevensDelete.ParamByName('maxdatum').AsDate:= maxdatum;
        dm.ZOmzetgegevensDelete.Execute;
        dm.ZOmzetgegevensDelete.Connection.Commit;

      end;
      break;
    end;
    if ((mindatum = 0) and (maxdatum = 0)) then
    begin
      mindatum := somzetgrid.worksheet.findcell(i,2)^.DateTimeValue;
      maxdatum :=somzetgrid.worksheet.findcell(i,2)^.DateTimeValue;
    end
    else
    begin
      if somzetgrid.worksheet.findcell(i,2)^.DateTimeValue  < mindatum then
        mindatum := somzetgrid.worksheet.findcell(i,2)^.DateTimeValue;
      if somzetgrid.worksheet.findcell(i,2)^.DateTimeValue > maxdatum then
        maxdatum := somzetgrid.worksheet.findcell(i,2)^.DateTimeValue;
    end;
    inc(i);
  end;
  //omzetgegevens toevoegen
  i := 3;
  while assigned(somzetgrid.worksheet.findcell(i,0)) do
  begin
    if  (pos('Totaal' ,somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue)= 1) then
    begin
      break;
    end;
   { for y := 0 to 3 do
    begin
      showmessage(inttostr(y) + '  '  + somzetgrid.worksheet.findcell(i,y)^.UTF8StringValue);
      case  somzetgrid.worksheet.findcell(i,y)^.ContentType of
         cctnumber : showmessage(' number' );
         cctUTF8String  : showmessage(' string');
         cctDateTime : showmessage('datetime');
      end;
    end;   }
    dm.OmzetgegevensAdd.ParamByName('datum').asdatetime := somzetgrid.worksheet.findcell(i,2)^.DateTimeValue;
    dm.OmzetgegevensAdd.ParamByName('wag_id').AsInteger:= strtoint(somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue);
    dm.OmzetgegevensAdd.ParamByName('waarde').AsFloat:= somzetgrid.worksheet.findcell(i,3)^.NumberValue;
   // showmessage(dm.zomzetgegevensadd.parambyname('datum').AsString + '   '  +dm.zomzetgegevensadd.parambyname('wag_id').AsString + '  '  + dm.zomzetgegevensadd.parambyname('waarde').asstring);
    dm.OmzetgegevensAdd.Execute;
    dm.OmzetgegevensAdd.Connection.Commit;
    inc(i);
  end;
  {
      mindatum := 0;
      maxdatum := 0;
      for i := 2 to somzetgrid.RowCount-1 do
      //for i := 2 to 10 do
      begin
        if assigned(somzetgrid.worksheet.findcell(i,0)) then
        begin
         // memo.lines.add(somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue);
          if (somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue[1] in ['0'..'9']) then
          begin
            //showmessage('i : '+ inttostr(i) + ' ; y : '+ inttostr(y) + ' : ' +  floattostr(somzetgrid.worksheet.findcell(i,0)^.NumberValue));
            //showmessage('dateseparator is : "' + dateseparator+'" celinhoud is : "'+somzetgrid.cells[0,i]+'"');
            datumstring := somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue;
            datumstring[3] := dateseparator;
            datumstring[6] := dateseparator;
            if ((mindatum = 0) and (maxdatum = 0)) then
            begin
              mindatum := strtodate(datumstring);
              maxdatum := strtodate(datumstring);
            end
            else
            begin
              if strtodate(datumstring) < mindatum then
                mindatum := strtodate(datumstring);
               if strtodate(datumstring) > maxdatum then
                maxdatum := strtodate(datumstring);
            end;
          end;
        end;
      end;
      if mindatum <> 0 then
      begin
        dm.ZVoorraadcorrectiesDelete.ParamByName('mindatum').AsDate:= mindatum;
        dm.ZVoorraadcorrectiesDelete.ParamByName('maxdatum').AsDate:= maxdatum;
        dm.ZVoorraadcorrectiesDelete.Execute;
        dm.ZVoorraadcorrectiesDelete.Connection.Commit;

     //   dm.ZvoorraadcorrectiesQuery.ApplyUpdates;
     //   dm.zvoorraadcorre
      //  dm.ZvoorraadcorrectiesQuery.CommitUpdates;
      end;
      memo.Lines.Add('mindatum is : ' + datetostr(mindatum));
      memo.Lines.Add('maxdatum is : ' + datetostr(maxdatum));

      for i := 2 to somzetgrid.RowCount -1 do
  {for debug enabled*/} //   for i := 2 to 10 do
      begin
        if assigned(somzetgrid.worksheet.findcell(i,0)) then
        begin
          if not (somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue[1] in ['0'..'9']) then
          begin
             memo.lines.add(somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue);
             omzetgroep := somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue;
             omzetgroep := copy(omzetgroep,3,length(omzetgroep));
          // showmessage('"'+omzetgroep+'"');

          end
          else
          begin
            if dm.ZvoorraadcorrectiesQuery.Active then
            dm.ZvoorraadcorrectiesQuery.close;
            dm.ZVoorraadCorrectiesQuery.ParamByName('omzetgroep').AsString:= omzetgroep;
            dm.ZVoorraadCorrectiesQuery.ParamByName('soort').AsString:= somzetgrid.worksheet.findcell(i,5)^.UTF8StringValue;
           // showmessage(dm.zvoorraadcorrectiesquery.parambyname('soort').asstring);
            datumstring := somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue;
            datumstring[3] := dateseparator;
            datumstring[6] := dateseparator;

            dm.ZVoorraadCorrectiesQuery.ParamByName('datum').AsDate:= strtodate(datumstring);
           // showmessage(dm.zvoorraadcorrectiesquery.parambyname('datum').asstring);
            dm.ZVoorraadCorrectiesQuery.ParamByName('artikelnummer').AsInteger:= round(somzetgrid.worksheet.findcell(i,1)^.NumberValue);
           // showmessage(dm.zvoorraadcorrectiesquery.parambyname('artikelnummer').asstring);
            dm.ZVoorraadCorrectiesQuery.ParamByName('omschrijving').AsString:= somzetgrid.worksheet.findcell(i,2)^.UTF8StringValue;
           // showmessage(dm.zvoorraadcorrectiesquery.parambyname('omschrijving').asstring);
    //        dm.ZVoorraadCorrectiesQuery.ParamByName('aantal').AsFloat:= strtofloat(StringReplace(somzetgrid.worksheet.findcell(i,3)^.NumberValue, ' ', '',[rfReplaceAll, rfIgnoreCase]))/1000;
            dm.ZVoorraadCorrectiesQuery.ParamByName('aantal').AsFloat:= somzetgrid.worksheet.findcell(i,3)^.NumberValue;
            //somzetgrid.worksheet.findcell(i,3)^.ContentType:= tcellcontenttype.cctUTF8String;
           // showmessage('aantal is : '+ floattostr(somzetgrid.worksheet.findcell(i,3)^.NumberValue));
          //  floattostr(dm.zvoorraadcorrectiesquery.parambyname('aantal').asfloat));
          //  waarde := somzetgrid.worksheet.findcell(i,6)^.NumberValue;
         //   waarde := StringReplace(waarde, ' ', '',[rfReplaceAll, rfIgnoreCase]);
           // waarde := StringReplace(waarde, '.', '',[rfReplaceAll, rfIgnoreCase]);
            dm.ZVoorraadCorrectiesQuery.ParamByName('waarde').AsFloat:=
              -somzetgrid.worksheet.findcell(i,6)^.NumberValue;
            dm.ZvoorraadcorrectiesQuery.Active:= true;
            dm.ZvoorraadcorrectiesQuery.ApplyUpdates;
            dm.ZvoorraadcorrectiesQuery.CommitUpdates;
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('oomzetgroep').AsString+'"');
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('osoort').AsString+'"');
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('oartikelnummer').AsString+'"');
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('oomschrijving').AsString+'"');
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('oaantal').AsString+'"');
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('owaarde').AsString+'"');
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('omzetgroep_id').AsString+'"');
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('soort_id').AsString+'"');
            {dm.ZVoorraadCorrectiesAdd.ParamByName('omzetgroep').AsString:= omzetgroep;
            dm.ZVoorraadCorrectiesAdd.ParamByName('soort').AsString:= somzetgrid.cells[5,i];
            dm.ZVoorraadCorrectiesAdd.ParamByName('artikelnummer').AsInteger:= strtoint(somzetgrid.cells[1,i]);
            dm.ZVoorraadCorrectiesAdd.ParamByName('omschrijving').AsString:= somzetgrid.cells[2,i];
            dm.ZVoorraadCorrectiesAdd.ParamByName('aantal').AsFloat:= strtofloat(StringReplace(somzetgrid.cells[3,i], ' ', '',[rfReplaceAll, rfIgnoreCase]))/1000;
            dm.ZVoorraadCorrectiesAdd.ParamByName('waarde').AsFloat:= strtofloat(
            StringReplace(somzetgrid.cells[6,i], ' ', '',[rfReplaceAll, rfIgnoreCase]));

            dm.ZVoorraadCorrectiesAdd.Execute;}
          end;
        end;
      end;
      memo.Lines.add('KLAAR');;
      dm.ZConnection.commit;
    end
    else
      showmessage('Onbekend soort spreadsheet');
 //   showmessage(inttostr(pos('Weekoverzicht Van :',somzetgrid.Cells[0,0])));
  end;}
end;



procedure TForm1.leesomzetspreadsheet(filename : string);
var
   datumstr        : array[1..7] of string;
   mindatum        : tdatetime;
   maxdatum        : tdatetime;
   wagnummer       : string;
   wagomschrijving : string;
   wagomzet        : array[1..7] of string;
   i,y             : integer;
   omzetgroep      : string;
   l_decimalseparator : char;
   datumstring     : string;
   waarde          : string;

begin
  for y := 0 to 6 do
  begin
    wagomzet[y+1] := '';
  end;
  try
    sWorkbook.LoadFromSpreadsheetFile(Filename, sfExcel5);
  except
    try
      sWorkbook.LoadFromSpreadsheetFile(Filename, sfExcel8);
    except
    end;
  end;
 // somzetgrid.Worksheet.Cells[1,1] := 'hallo';
  if pos('Weekoverzicht Van :',somzetgrid.Worksheet.FindCell(0,0)^.UTF8StringValue) = 1 then
  begin
    //load datum
    for i := 0 to 6 do
    begin
    //  showmessage(somzetgrid.worksheet.findcell(11,4+2*i)^.UTF8StringValue);
      datumstr[i+1] := copy(somzetgrid.worksheet.findcell(11,4+2*i)^.UTF8StringValue,1,10);
      datumstr[i+1][3] := '.';
      datumstr[i+1][6] := '.';
    end;
    //load groep number
    i := 0;
    while assigned(somzetgrid.worksheet.findcell(16+i,0)) do
   // while somzetgrid.worksheet.findcell(16+i,0)^.UTF8StringValue <> '' do
    begin
      wagnummer := copy(somzetgrid.worksheet.findcell(16+i,0)^.UTF8StringValue,1,4);
      wagomschrijving := copy(somzetgrid.worksheet.findcell(16+i,0)^.UTF8StringValue,6,100);
      for y := 0 to 6 do
      begin
        wagomzet[y+1] := '';
        if somzetgrid.worksheet.findcell(16+i,3+2*y) <> nil then
        begin
        //  showmessage('i : '+ inttostr(i) + ' ; y : '+ inttostr(y) + ' : ' +  floattostr(somzetgrid.worksheet.findcell(16+i,3+2*y)^.NumberValue));
          wagomzet[y+1] := floattostr(somzetgrid.worksheet.findcell(16+i,3+2*y)^.NumberValue);
        end;
        memo.Lines.Add(wagnummer+','''+datumstr[y+1]+''','''+wagomzet[y+1]+'''');
        if wagomzet[y+1] <> '' then
        begin
          datumstr[y+1,3] := dateseparator;
          datumstr[y+1,6] := dateseparator;
          dm.ZOmzetgegevensAdd.ParamByName('datum').asdatetime:= strtodate(datumstr[y+1]);
          dm.ZOmzetgegevensAdd.ParamByName('wag_id').AsInteger:= strtoint(wagnummer);
          dm.ZOmzetgegevensAdd.ParamByName('waarde').AsFloat:= strtofloat(wagomzet[y+1]);
          dm.ZOmzetgegevensAdd.Execute;
          dm.ZOmzetgegevensAdd.Connection.Commit;
        end;
      end;
      //showmessage('voor inc i : '+ inttostr(i) +' wagnummer : '+ wagnummer);
      inc(i);
      //showmessage('na inc i : '+ inttostr(i)  +' wagnummer : '+ wagnummer);
    end;
    //showmessage('na while loop')
  end
  else
  begin
    if (pos('Winkelassortimentsgroep',somzetgrid.worksheet.findcell(0,0)^.UTF8StringValue) = 1) and
        (pos('Totaal',somzetgrid.worksheet.findcell(1,0)^.UTF8StringValue) = 1) then
    begin
      mindatum := 0;
      maxdatum := 0;
      for i := 2 to somzetgrid.RowCount-1 do
      //for i := 2 to 10 do
      begin
        if assigned(somzetgrid.worksheet.findcell(i,0)) then
        begin
         // memo.lines.add(somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue);
          if (somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue[1] in ['0'..'9']) then
          begin
            //showmessage('i : '+ inttostr(i) + ' ; y : '+ inttostr(y) + ' : ' +  floattostr(somzetgrid.worksheet.findcell(i,0)^.NumberValue));
            //showmessage('dateseparator is : "' + dateseparator+'" celinhoud is : "'+somzetgrid.cells[0,i]+'"');
            datumstring := somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue;
            datumstring[3] := dateseparator;
            datumstring[6] := dateseparator;
            if ((mindatum = 0) and (maxdatum = 0)) then
            begin
              mindatum := strtodate(datumstring);
              maxdatum := strtodate(datumstring);
            end
            else
            begin
              if strtodate(datumstring) < mindatum then
                mindatum := strtodate(datumstring);
               if strtodate(datumstring) > maxdatum then
                maxdatum := strtodate(datumstring);
            end;
          end;
        end;
      end;
      if mindatum <> 0 then
      begin
        dm.ZVoorraadcorrectiesDelete.ParamByName('mindatum').AsDate:= mindatum;
        dm.ZVoorraadcorrectiesDelete.ParamByName('maxdatum').AsDate:= maxdatum;
        dm.ZVoorraadcorrectiesDelete.Execute;
        dm.ZVoorraadcorrectiesDelete.Connection.Commit;

     //   dm.ZvoorraadcorrectiesQuery.ApplyUpdates;
     //   dm.zvoorraadcorre
      //  dm.ZvoorraadcorrectiesQuery.CommitUpdates;
      end;
      memo.Lines.Add('mindatum is : ' + datetostr(mindatum));
      memo.Lines.Add('maxdatum is : ' + datetostr(maxdatum));

      for i := 2 to somzetgrid.RowCount -1 do
  {for debug enabled*/} //   for i := 2 to 10 do
      begin
        if assigned(somzetgrid.worksheet.findcell(i,0)) then
        begin
          if not (somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue[1] in ['0'..'9']) then
          begin
             memo.lines.add(somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue);
             omzetgroep := somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue;
             omzetgroep := copy(omzetgroep,3,length(omzetgroep));
          // showmessage('"'+omzetgroep+'"');

          end
          else
          begin
            if dm.ZvoorraadcorrectiesQuery.Active then
            dm.ZvoorraadcorrectiesQuery.close;
            dm.ZVoorraadCorrectiesQuery.ParamByName('omzetgroep').AsString:= omzetgroep;
            dm.ZVoorraadCorrectiesQuery.ParamByName('soort').AsString:= somzetgrid.worksheet.findcell(i,5)^.UTF8StringValue;
           // showmessage(dm.zvoorraadcorrectiesquery.parambyname('soort').asstring);
            datumstring := somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue;
            datumstring[3] := dateseparator;
            datumstring[6] := dateseparator;

            dm.ZVoorraadCorrectiesQuery.ParamByName('datum').AsDate:= strtodate(datumstring);
           // showmessage(dm.zvoorraadcorrectiesquery.parambyname('datum').asstring);
            dm.ZVoorraadCorrectiesQuery.ParamByName('artikelnummer').AsInteger:= round(somzetgrid.worksheet.findcell(i,1)^.NumberValue);
           // showmessage(dm.zvoorraadcorrectiesquery.parambyname('artikelnummer').asstring);
            dm.ZVoorraadCorrectiesQuery.ParamByName('omschrijving').AsString:= somzetgrid.worksheet.findcell(i,2)^.UTF8StringValue;
           // showmessage(dm.zvoorraadcorrectiesquery.parambyname('omschrijving').asstring);
    //        dm.ZVoorraadCorrectiesQuery.ParamByName('aantal').AsFloat:= strtofloat(StringReplace(somzetgrid.worksheet.findcell(i,3)^.NumberValue, ' ', '',[rfReplaceAll, rfIgnoreCase]))/1000;
            dm.ZVoorraadCorrectiesQuery.ParamByName('aantal').AsFloat:= somzetgrid.worksheet.findcell(i,3)^.NumberValue;
            //somzetgrid.worksheet.findcell(i,3)^.ContentType:= tcellcontenttype.cctUTF8String;
           // showmessage('aantal is : '+ floattostr(somzetgrid.worksheet.findcell(i,3)^.NumberValue));
          //  floattostr(dm.zvoorraadcorrectiesquery.parambyname('aantal').asfloat));
          //  waarde := somzetgrid.worksheet.findcell(i,6)^.NumberValue;
         //   waarde := StringReplace(waarde, ' ', '',[rfReplaceAll, rfIgnoreCase]);
           // waarde := StringReplace(waarde, '.', '',[rfReplaceAll, rfIgnoreCase]);
            dm.ZVoorraadCorrectiesQuery.ParamByName('waarde').AsFloat:=
              -somzetgrid.worksheet.findcell(i,6)^.NumberValue;
            dm.ZvoorraadcorrectiesQuery.Active:= true;
            dm.ZvoorraadcorrectiesQuery.ApplyUpdates;
            dm.ZvoorraadcorrectiesQuery.CommitUpdates;
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('oomzetgroep').AsString+'"');
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('osoort').AsString+'"');
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('oartikelnummer').AsString+'"');
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('oomschrijving').AsString+'"');
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('oaantal').AsString+'"');
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('owaarde').AsString+'"');
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('omzetgroep_id').AsString+'"');
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('soort_id').AsString+'"');
            {dm.ZVoorraadCorrectiesAdd.ParamByName('omzetgroep').AsString:= omzetgroep;
            dm.ZVoorraadCorrectiesAdd.ParamByName('soort').AsString:= somzetgrid.cells[5,i];
            dm.ZVoorraadCorrectiesAdd.ParamByName('artikelnummer').AsInteger:= strtoint(somzetgrid.cells[1,i]);
            dm.ZVoorraadCorrectiesAdd.ParamByName('omschrijving').AsString:= somzetgrid.cells[2,i];
            dm.ZVoorraadCorrectiesAdd.ParamByName('aantal').AsFloat:= strtofloat(StringReplace(somzetgrid.cells[3,i], ' ', '',[rfReplaceAll, rfIgnoreCase]))/1000;
            dm.ZVoorraadCorrectiesAdd.ParamByName('waarde').AsFloat:= strtofloat(
            StringReplace(somzetgrid.cells[6,i], ' ', '',[rfReplaceAll, rfIgnoreCase]));

            dm.ZVoorraadCorrectiesAdd.Execute;}
          end;
        end;
      end;
      memo.Lines.add('KLAAR');;
      dm.ZConnection.commit;
    end
    else
      showmessage('Onbekend soort spreadsheet');
 //   showmessage(inttostr(pos('Weekoverzicht Van :',somzetgrid.Cells[0,0])));
  end;
end;


procedure TForm1.leesomzetR10(filename : string);
var
   datumstr        : array[1..7] of string;
   mindatum        : tdatetime;
   maxdatum        : tdatetime;
   wagnummer       : string;
   wagomschrijving : string;
   wagomzet        : array[1..7] of string;
   i,y             : integer;
   omzetgroep      : string;
   l_decimalseparator : char;
   datumstring     : string;
   waarde          : string;
   zoekstring      : string;

begin
  for y := 0 to 6 do
  begin
    wagomzet[y+1] := '';
  end;
  somzetgrid.Clear;
  try
    sWorkbook.LoadFromSpreadsheetFile(Filename, sfExcelXML);
  except
    try
      sWorkbook.LoadFromSpreadsheetFile(Filename, sfOOXML);
    except
      try
         sWorkbook.LoadFromSpreadsheetFile(Filename, sfExcel5);
      except
         try
          sWorkbook.LoadFromSpreadsheetFile(Filename, sfExcel8);

        except
           showmessage('niets kunnen laden');
          exit;
        end;
      end;

    end;
  end;
  //showmessage(somzetgrid.worksheet.findcell(2,13)^.UTF8StringValue);
  //showmessage(somzetgrid.worksheet.findcell(6,2)^.utf8stringvalue);
 // somzetgrid.Worksheet.Cells[1,1] := 'hallo';
 zoekstring := '' ;
  if assigned(somzetgrid.Worksheet.FindCell(0,0)) then
 begin
   zoekstring :=  somzetgrid.Worksheet.FindCell(0,0)^.UTF8StringValue;
 end;
 if (pos('OPE1010 - Vestiging Omzet totaal->Datum',zoekstring)  = 1) then
 begin
   leesomzetopdatum;
   exit;
 end;
 showmessage('na leesomzetopdatum, verder zoeken');
   zoekstring := '' ;
  if assigned(somzetgrid.Worksheet.FindCell(1,13)) then
  begin
    zoekstring :=  somzetgrid.Worksheet.FindCell(1,13)^.UTF8StringValue;
  end;
  if ((pos('OPE1010 - Omzet totaal',zoekstring)  = 1) or (pos('OPE1010 - Omzet Totaal',zoekstring)  = 1)) then
  begin
    leesOPE1010;;
  end
  else
  begin
    if (pos('Winkelassortimentsgroep',somzetgrid.worksheet.findcell(0,0)^.UTF8StringValue) = 1) and
        (pos('Totaal',somzetgrid.worksheet.findcell(1,0)^.UTF8StringValue) = 1) then
    begin
      mindatum := 0;
      maxdatum := 0;
      for i := 2 to somzetgrid.RowCount-1 do
      //for i := 2 to 10 do
      begin
        if assigned(somzetgrid.worksheet.findcell(i,0)) then
        begin
         // memo.lines.add(somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue);
          if (somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue[1] in ['0'..'9']) then
          begin
            //showmessage('i : '+ inttostr(i) + ' ; y : '+ inttostr(y) + ' : ' +  floattostr(somzetgrid.worksheet.findcell(i,0)^.NumberValue));
            //showmessage('dateseparator is : "' + dateseparator+'" celinhoud is : "'+somzetgrid.cells[0,i]+'"');
            datumstring := somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue;
            datumstring[3] := dateseparator;
            datumstring[6] := dateseparator;
            if ((mindatum = 0) and (maxdatum = 0)) then
            begin
              mindatum := strtodate(datumstring);
              maxdatum := strtodate(datumstring);
            end
            else
            begin
              if strtodate(datumstring) < mindatum then
                mindatum := strtodate(datumstring);
               if strtodate(datumstring) > maxdatum then
                maxdatum := strtodate(datumstring);
            end;
          end;
        end;
      end;
      if mindatum <> 0 then
      begin
        dm.ZVoorraadcorrectiesDelete.ParamByName('mindatum').AsDate:= mindatum;
        dm.ZVoorraadcorrectiesDelete.ParamByName('maxdatum').AsDate:= maxdatum;
        dm.ZVoorraadcorrectiesDelete.Execute;
        dm.ZVoorraadcorrectiesDelete.Connection.Commit;

     //   dm.ZvoorraadcorrectiesQuery.ApplyUpdates;
     //   dm.zvoorraadcorre
      //  dm.ZvoorraadcorrectiesQuery.CommitUpdates;
      end;
      memo.Lines.Add('mindatum is : ' + datetostr(mindatum));
      memo.Lines.Add('maxdatum is : ' + datetostr(maxdatum));

      for i := 2 to somzetgrid.RowCount -1 do
  {for debug enabled*/} //   for i := 2 to 10 do
      begin
        if assigned(somzetgrid.worksheet.findcell(i,0)) then
        begin
          if not (somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue[1] in ['0'..'9']) then
          begin
             memo.lines.add(somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue);
             omzetgroep := somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue;
             omzetgroep := copy(omzetgroep,3,length(omzetgroep));
          // showmessage('"'+omzetgroep+'"');

          end
          else
          begin
            if dm.ZvoorraadcorrectiesQuery.Active then
            dm.ZvoorraadcorrectiesQuery.close;
            dm.ZVoorraadCorrectiesQuery.ParamByName('omzetgroep').AsString:= omzetgroep;
            dm.ZVoorraadCorrectiesQuery.ParamByName('soort').AsString:= somzetgrid.worksheet.findcell(i,5)^.UTF8StringValue;
           // showmessage(dm.zvoorraadcorrectiesquery.parambyname('soort').asstring);
            datumstring := somzetgrid.worksheet.findcell(i,0)^.UTF8StringValue;
            datumstring[3] := dateseparator;
            datumstring[6] := dateseparator;

            dm.ZVoorraadCorrectiesQuery.ParamByName('datum').AsDate:= strtodate(datumstring);
           // showmessage(dm.zvoorraadcorrectiesquery.parambyname('datum').asstring);
            dm.ZVoorraadCorrectiesQuery.ParamByName('artikelnummer').AsInteger:= round(somzetgrid.worksheet.findcell(i,1)^.NumberValue);
           // showmessage(dm.zvoorraadcorrectiesquery.parambyname('artikelnummer').asstring);
            dm.ZVoorraadCorrectiesQuery.ParamByName('omschrijving').AsString:= somzetgrid.worksheet.findcell(i,2)^.UTF8StringValue;
           // showmessage(dm.zvoorraadcorrectiesquery.parambyname('omschrijving').asstring);
    //        dm.ZVoorraadCorrectiesQuery.ParamByName('aantal').AsFloat:= strtofloat(StringReplace(somzetgrid.worksheet.findcell(i,3)^.NumberValue, ' ', '',[rfReplaceAll, rfIgnoreCase]))/1000;
            dm.ZVoorraadCorrectiesQuery.ParamByName('aantal').AsFloat:= somzetgrid.worksheet.findcell(i,3)^.NumberValue;
            //somzetgrid.worksheet.findcell(i,3)^.ContentType:= tcellcontenttype.cctUTF8String;
           // showmessage('aantal is : '+ floattostr(somzetgrid.worksheet.findcell(i,3)^.NumberValue));
          //  floattostr(dm.zvoorraadcorrectiesquery.parambyname('aantal').asfloat));
          //  waarde := somzetgrid.worksheet.findcell(i,6)^.NumberValue;
         //   waarde := StringReplace(waarde, ' ', '',[rfReplaceAll, rfIgnoreCase]);
           // waarde := StringReplace(waarde, '.', '',[rfReplaceAll, rfIgnoreCase]);
            dm.ZVoorraadCorrectiesQuery.ParamByName('waarde').AsFloat:=
              -somzetgrid.worksheet.findcell(i,6)^.NumberValue;
            dm.ZvoorraadcorrectiesQuery.Active:= true;
            dm.ZvoorraadcorrectiesQuery.ApplyUpdates;
            dm.ZvoorraadcorrectiesQuery.CommitUpdates;
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('oomzetgroep').AsString+'"');
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('osoort').AsString+'"');
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('oartikelnummer').AsString+'"');
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('oomschrijving').AsString+'"');
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('oaantal').AsString+'"');
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('owaarde').AsString+'"');
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('omzetgroep_id').AsString+'"');
            memo.Lines.Add('"'+dm.ZvoorraadcorrectiesQuery.FieldByName('soort_id').AsString+'"');
            {dm.ZVoorraadCorrectiesAdd.ParamByName('omzetgroep').AsString:= omzetgroep;
            dm.ZVoorraadCorrectiesAdd.ParamByName('soort').AsString:= somzetgrid.cells[5,i];
            dm.ZVoorraadCorrectiesAdd.ParamByName('artikelnummer').AsInteger:= strtoint(somzetgrid.cells[1,i]);
            dm.ZVoorraadCorrectiesAdd.ParamByName('omschrijving').AsString:= somzetgrid.cells[2,i];
            dm.ZVoorraadCorrectiesAdd.ParamByName('aantal').AsFloat:= strtofloat(StringReplace(somzetgrid.cells[3,i], ' ', '',[rfReplaceAll, rfIgnoreCase]))/1000;
            dm.ZVoorraadCorrectiesAdd.ParamByName('waarde').AsFloat:= strtofloat(
            StringReplace(somzetgrid.cells[6,i], ' ', '',[rfReplaceAll, rfIgnoreCase]));

            dm.ZVoorraadCorrectiesAdd.Execute;}
          end;
        end;
      end;
      memo.Lines.add('KLAAR');;
      dm.ZConnection.commit;
    end
    else
      showmessage('Onbekend soort spreadsheet');
 //   showmessage(inttostr(pos('Weekoverzicht Van :',somzetgrid.Cells[0,0])));
  end
end;

procedure Tform1.leesOPE1010();
var
  datumstr        : array[1..7] of string;
  mindatum        : tdatetime;
  maxdatum        : tdatetime;
  wagnummer       : string;
  wagomschrijving : string;
  wagomzet        : array[1..7] of string;
  i,y             : integer;
  omzetgroep      : string;
  l_decimalseparator : char;
  datumstring     : string;
  waarde          : string;
  zoekstring      : string;
begin
  for i := 0 to 0 do
  begin
    //  showmessage(somzetgrid.worksheet.findcell(11,4+2*i)^.UTF8StringValue);
    datumstr[i+1] := copy(somzetgrid.worksheet.findcell(5,2)^.UTF8StringValue,9,10);
    datumstr[i+1][3] := dateseparator;
    if datumstr[i+1][6] = '-' then
    begin
      datumstr[i+1][6] := dateseparator;
    end
    else
    begin
      datumstr[i+1] := copy(somzetgrid.worksheet.findcell(5,2)^.UTF8StringValue,9,9);
      datumstr[i+1][3] := dateseparator;
      datumstr[i+1][5] := dateseparator;
    end  ;
  end;
  //load groep number
  i := 0;
  //while assigned(somzetgrid.worksheet.findcell(16+i,0)) do
  // while somzetgrid.worksheet.findcell(16+i,0)^.UTF8StringValue <> '' do
  while pos('Totaal',somzetgrid.Worksheet.FindCell(14+i,2)^.UTF8StringValue) <> 1 do
  begin
    wagnummer := copy(somzetgrid.worksheet.findcell(14+i,2)^.UTF8StringValue,1,4);
    wagomschrijving := copy(somzetgrid.worksheet.findcell(14+i,5)^.UTF8StringValue,6,100);
    for y := 0 to 0 do
    begin
      wagomzet[y+1] := '';
      if somzetgrid.worksheet.findcell(14+i,9) <> nil then
      begin
      //  showmessage('i : '+ inttostr(i) + ' ; y : '+ inttostr(y) + ' : ' +  floattostr(somzetgrid.worksheet.findcell(16+i,3+2*y)^.NumberValue));
        wagomzet[y+1] := floattostr(somzetgrid.worksheet.findcell(14+i,9)^.NumberValue);
      end;
      memo.Lines.Add(wagnummer+','''+datumstr[y+1]+''','''+wagomzet[y+1]+'''');
      if wagomzet[y+1] <> '' then
      begin
      //  datumstr[y+1,3] := dateseparator;
      //  datumstr[y+1,6] := dateseparator;
        dm.ZOmzetgegevensAdd.ParamByName('datum').asdatetime:= strtodate(datumstr[y+1]);
        dm.ZOmzetgegevensAdd.ParamByName('wag_id').AsInteger:= strtoint(wagnummer);
        dm.ZOmzetgegevensAdd.ParamByName('waarde').AsFloat:= strtofloat(wagomzet[y+1]);
        try
          dm.ZOmzetgegevensAdd.Execute;
        except
          try
            dm.ZWagAdd.ParamByName('wag_id').asinteger := strtoint(wagnummer);
            dm.ZWagAdd.ParamByName('omschrijving').AsString := (wagomschrijving);
            dm.ZWagAdd.Execute;
            dm.ZWagAdd.connection.commit;
            dm.ZOmzetgegevensAdd.Execute;
          except
            showmessage('iets fout met wagnummer');
          end;
        end;
        dm.ZOmzetgegevensAdd.Connection.Commit;
      end;
    end;
      //showmessage('voor inc i : '+ inttostr(i) +' wagnummer : '+ wagnummer);
    inc(i);
     //showmessage('na inc i : '+ inttostr(i)  +' wagnummer : '+ wagnummer);
  end;
    //showmessage('na while loop')
end;





procedure TForm1.buttonomzetClick(Sender: TObject);
begin
   if omzetdialoog.Execute then
   begin
      leesomzetspreadsheet(omzetdialoog.filename);
   end;
   dataunit.DM.ZConnection.Disconnect;
   dataunit.DM.ZConnection.Connect;
   dataunit.DM.ZJaarweek.Open;
   //dm.ZJaarweek.refresh;
end;

procedure TForm1.DervingInfoButtenClick(Sender: TObject);
begin
   dm.ZDaginfo.Close;
   if dm.ZDaginfo.Params.Count = 1 then
   dm.ZDaginfo.ParamByName('JAARWEEK').AsInteger:= dm.ZJAARWEEK.FieldByName('JAARWEEK').AsInteger;
   dm.ZDaginfo.Active:= true;
   weekinfo.LoadFromFile(Instellingenunit.FormInstellingen.EditDagInfoRapport.Text);
  // weekinfo.LoadFromFile('C:\Users\chrgra\Documents\Projecten\ODN\ODN Eiland\reports\daginfo.lrf');
   Weekinfo.ShowReport;
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  Info: TVersionInfo;
  Version: string;
begin
  decimalseparator := '.';
  thousandseparator := ',';
  Info := TVersionInfo.Create;
  Info.Load(HINSTANCE);
  self.Caption := Format('Dagomzet voor Personeelstool Versie %d.%d.%d build %d Hoofdmenu', [Info.FixedInfo.FileVersion[0],Info.FixedInfo.FileVersion[1],Info.FixedInfo.FileVersion[2],Info.FixedInfo.FileVersion[3]]);
  Info.Free;


end;

procedure TForm1.MenuUrenClick(Sender: TObject);
begin
  urenunit.Form2.ShowModal;
end;

procedure TForm1.WeekInfoButtonClick(Sender: TObject);
begin
   dm.ZWeekinfo.Close;
   if dm.ZWeekinfo.Params.Count = 1 then
   dm.ZWeekinfo.ParamByName('JAARWEEK').AsInteger:= dm.ZJAARWEEK.FieldByName('JAARWEEK').AsInteger;
   dm.zweekinfo.Active:= true;
   weekinfo.LoadFromFile(Instellingenunit.FormInstellingen.EditWeekInfoRapport.text);
   weekinfo.LoadFromFile('C:\Users\chrgra\Documents\Projecten\ODN\ODN Eiland\reports\weekinfo3.lrf');
   Weekinfo.ShowReport;
end;

procedure TForm1.Panel2Click(Sender: TObject);
begin

end;

procedure TForm1.sOmzetGridClick(Sender: TObject);
begin
   //showmessage('active col : '+inttostr(somzetgrid.Worksheet.ActiveCellCol) +
     //' active row : ' + inttostr(Somzetgrid.Worksheet.ActiveCellRow));
end;

const wildcardsearch = '\*';
      filefound = 0;

procedure TForm1.Button1Click(Sender: TObject);
Var
   SearchResult : TSearchRec;
   MoreFiles : Integer;
begin
  if omzetdirectorydialoog.Execute then
  begin
    If FindFirst (omzetdirectorydialoog.FileName+wildcardsearch, (faAnyFile And Not faDirectory) , SearchResult) = FileFound Then
    Begin
      memo.Lines.Add(omzetdirectorydialoog.FileName +'\'+ searchresult.Name);
      leesomzetspreadsheet(omzetdirectorydialoog.FileName +'\'+ searchresult.Name);
      While FindNext (SearchResult) = FileFound Do
      Begin
        memo.Lines.Add(omzetdirectorydialoog.FileName +'\'+ searchresult.Name);
        leesomzetspreadsheet(omzetdirectorydialoog.FileName +'\'+ searchresult.Name);
      End;
    End;
    FindClose (SearchResult);
  end;
  dm.ZJaarweek.refresh;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
   dm.ZPMTInfo.Close;
   if dm.ZPMTinfo.Params.Count = 1 then
   dm.ZPMTinfo.ParamByName('JAARWEEK').AsInteger:= dm.ZJAARWEEK.FieldByName('JAARWEEK').AsInteger;
   dm.ZPMTinfo.Active:= true;
  // weekinfo.LoadFromFile(Instellingenunit.FormInstellingen.EditPMTRapport.text);
  // weekinfo.LoadFromFile('C:\Users\chrgra\Documents\Projecten\ODN\ODN Eiland\reports\daginfo_voor_personeelstool.lrf');
   PMTinfo.ShowReport;
end;

procedure TForm1.Button3Click(Sender: TObject);
var
   i : integer;
begin
  for i := 0 to 52 do
  begin

 {
  dm.ZAfdelinginfo.Close;
  dm.ZConnection.Commit;
  if dm.ZAfdelinginfo.params.Count = 1 then
  dm.ZAfdelinginfo.ParamByName('JAARWEEK').AsInteger:= dm.ZJAARWEEK.FieldByName('JAARWEEK').AsInteger;
  dm.ZAfdelinginfo.Active:= true;
  weekinfo.LoadFromFile(Instellingenunit.FormInstellingen.EditAfdelingRapport.text);
  //weekinfo.LoadFromFile('C:\Users\chrgra\Documents\Projecten\ODN\reports\afdelinginfo.lrf');
  Weekinfo.ShowReport;}

  end;
end;

procedure TForm1.Button4Click(Sender: TObject);
begin

end;

procedure TForm1.FormShow(Sender: TObject);
begin
  dm.ZConnection.Connect;
  dm.ZJaarweek.Active:= true;
end;

procedure TForm1.MenuInstellingenClick(Sender: TObject);
begin
  forminstellingen.showmodal;
  dataunit.DM.ZConnection.Disconnect;
  dataunit.DM.ZConnection.Connect;
  dataunit.DM.ZJaarweek.Open;
end;

procedure TForm1.DaginfoNieuwButton1Click(Sender: TObject);
begin
   dm.ZDaginfonieuw.Close;
   dm.ZConnection.Commit;
   if dm.Zdaginfonieuw.Params.Count = 1 then
   dm.Zdaginfonieuw.ParamByName('JAARWEEK').AsInteger:= dm.ZJAARWEEK.FieldByName('JAARWEEK').AsInteger;
   dm.zdaginfonieuw.Active:= true;
   //weekinfo.LoadFromFile(Instellingenunit.FormInstellingen.EditDagInfoRapport.Text);
  // weekinfo.LoadFromFile('C:\Users\chrgra\Documents\Projecten\ODN\reports\daginfo_nieuw.lrf');

   daginfo.ShowReport;

end;

procedure TForm1.R10ButtonClick(Sender: TObject);
begin
   if omzetdialoog.Execute then
   begin
      leesomzetR10(omzetdialoog.filename);
   end;
   dataunit.DM.ZConnection.Disconnect;
   dataunit.DM.ZConnection.Connect;
   dataunit.DM.ZJaarweek.Open;
   //dm.ZJaarweek.refresh;
end;

procedure TForm1.WeekinfoGridAfterQuickSearch(Sender: TObject; Field: TField;
  var AValue: string);
begin

end;

procedure TForm1.WeekinfoNieuwButtonClick(Sender: TObject);
begin
   dm.Zweekinfonieuw.Close;
   dm.ZConnection.Commit;
   if dm.Zweekinfonieuw.Params.Count = 1 then
   dm.Zweekinfonieuw.ParamByName('JAARWEEK').AsInteger:= dm.ZJAARWEEK.FieldByName('JAARWEEK').AsInteger;
   dm.zweekinfonieuw.Active:= true;
   //weekinfo.LoadFromFile(Instellingenunit.FormInstellingen.EditWeekInfoRapport.text);
  // weekinfo.LoadFromFile('C:\Users\chrgra\Documents\Projecten\ODN\reports\weekinfo4.lrf');
   Weekinfo.ShowReport;
end;

end.

