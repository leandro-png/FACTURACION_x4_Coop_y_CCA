unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  ComCtrls, ExtCtrls, db, dbf, BufDataset,
  fpstypes, fpspreadsheet, fpsallformats, fpsexport;

type

  { TForm1 }

  TForm1 = class(TForm)
    Bevel1: TBevel;
    Bevel2: TBevel;
    Bevel3: TBevel;
    BtnCreateDatabase: TButton;
    BtnExport: TButton;
    BtnImport: TButton;
    EdRecordCount: TEdit;
    HeaderLabel3: TLabel;
    InfoLabel2: TLabel;
    HeaderLabel1: TLabel;
    InfoLabel1: TLabel;
    InfoLabel3: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    HeaderLabel2: TLabel;
    FileList: TListBox;
    Label3: TLabel;
    PageControl: TPageControl;
    Panel1: TPanel;
    RgDatabaseType: TRadioGroup;
    RgFileFormat: TRadioGroup;
    RgExportMode: TRadioGroup;
    TabDataGenerator: TTabSheet;
    TabExport: TTabSheet;
    TabImport: TTabSheet;
    procedure BtnCreateDatabaseClick(Sender: TObject);
    procedure BtnExportClick(Sender: TObject);
    procedure BtnImportClick(Sender: TObject);
    procedure FileListClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure PageControlChange(Sender: TObject);
    procedure RgFileFormatSelectionChanged(Sender: TObject);
  private
    { private declarations }
    FExportDataset: TDataset;
    FImportDataset: TDbf;
    FWorkbook: TsWorkbook;
    FHeaderTemplateCell: PCell;
    FDateTemplateCell: PCell;
    FCurrencyTemplatecell: PCell;
    FSocialSecurityTemplateCell: PCell;
    FImportedFieldNames: TStringList;
    FImportedRowCells: Array of TCell;
    // Actual export code when using FPSpreadsheet's fpsexport:
    // reads dbf and writes to spreadsheet
    // Expects FExportDataset to be available
    procedure ExportUsingFPSExport(MultipleSheets: Boolean; var DataFileName: string);
    // Actual export code when using virtual mode:
    // reads dbf and writes to spreadsheet
    // Expects FExportDataset to be available
    procedure ExportUsingVirtualMode(var DataFileName: string);
    // FPSExport: Get sheet name
    procedure ExporterGetSheetNameHandler(Sender: TObject; ASheetIndex: Integer;
      var ASheetName: String);
    // Virtual mode: for reading: all data for the database is generated here out of the spreadsheet file
    procedure ReadCellDataHandler(Sender: TObject; ARow, ACol: Cardinal;
      const ADataCell: PCell);
    // Virtual mode: for writing: all data for the cells is generated here (out of the .dbf file)
    procedure WriteCellDataHandler(Sender: TsWorksheet; ARow, ACol: Cardinal;
      var AValue: variant; var AStyleCell: PCell);
  public
    { public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

uses
  fpsNumFormat;

type
  // Ways to export dbf/dataset. Corresponds to the items
  // of the RgExportMode radiogroup
  TsExportModes=(
    seVirtual,       {manual coding using Virtual Mode}
    seFPSExport,     {uses FpSpreadsheet's fpsexport - takes more memory}
    seFPSExportMulti {uses FpSpreadsheet??s fpsexport to multiple sheets}
  );

const
  // Parameters for generating dbf file contents
  NUM_LAST_NAMES = 8;
  NUM_FIRST_NAMES = 8;
  NUM_CITIES = 10;
  LAST_NAMES: array[0..NUM_LAST_NAMES-1] of string = (
    'Chaplin', 'Washington', 'Dylan', 'Springsteen', 'Brando',
    'Monroe', 'Dean', 'Lincoln');
  FIRST_NAMES: array[0..NUM_FIRST_NAMES-1] of string = (
    'Charley', 'George', 'Bob', 'Bruce', 'Marlon',
    'Marylin', 'James', 'Abraham');
  CITIES: array[0..NUM_CITIES-1] of string = (
    'New York', 'Los Angeles', 'San Francisco', 'Chicago', 'Miami',
    'New Orleans', 'Washington', 'Boston', 'Seattle', 'Las Vegas');

  TABLENAME = 'people'; //name for the database table, extension will be added
  DATADIR = 'data'; //subdirectory where .dbf is stored

  // File formats corresponding to the items of the RgFileFormat radiogroup
  // Items in RadioGroup in Export tab match this order
  FILE_FORMATS: array[0..4] of TsSpreadsheetFormat = (
    sfExcel2, sfExcel5, sfExcel8, sfOOXML, sfOpenDocument
  );
  // Spreadsheet files will get the TABLENAME and have one of these extensions.
  FILE_EXT: array[0..4] of string = (
    '_excel2.xls', '_excel5.xls', '.xls', '.xlsx', '.ods');
  // Extension of database files supported
  DB_EXT: array[0..1] of string = (
    '.dbf', '.db');


{ TForm1 }

function RandomNumberStr(ALength: Integer): String;
var
  i: Integer;
begin
  SetLength(Result, ALength);
  for i := 1 to ALength do
    Result[i] := char(ord('0') + Random(10));
end;

{ This procedure creates a test dbf table with random data for us to play with }
procedure TForm1.BtnCreateDatabaseClick(Sender: TObject);
var
  i: Integer;
  startDate: TDate;
  maxAge: Integer = 80 * 365;
  f: TField;
  fn: String;
begin
  if FExportDataset <> nil then
    FExportDataset.Free;

  ForceDirectories(DATADIR);
  startDate := EncodeDate(2010, 8, 1);

  fn := DATADIR + DirectorySeparator + TABLENAME + DB_EXT[RgDatabaseType.itemIndex];
  case RgDatabaseType.ItemIndex of
    0: begin
         FExportDataset := TDbf.Create(self);
         TDbf(FExportDataset).FilePathFull := ExtractFilePath(fn);
         TDbf(FExportDataset).TableName := ExtractFileName(fn);
//         TDbf(FExportDataset).TableLevel := 4;  // DBase IV: most widely used.
         TDbf(FExportDataset).TableLevel := 25;   // FoxPro supports FieldType nfCurrency
       end;
    1: begin
         FExportDataset := TBufDataset.Create(self);
         TBufDataset(FExportDataset).Filename := fn;
       end;
    2: raise Exception.Create('Database type not supported');
  end;
  FExportDataset.FieldDefs.Add('Last name', ftString);
  FExportDataset.FieldDefs.Add('First name', ftString);
  FExportDataset.FieldDefs.Add('City', ftString);
  FExportDataset.FieldDefs.Add('Social sec', ftString);
  FExportDataset.FieldDefs.Add('Birthday', ftDate);
  FExportDataset.FieldDefs.Add('Salary', ftCurrency);
  FExportDataset.FieldDefs.Add('Work begin', ftDateTime);
  FExportDataset.FieldDefs.Add('Work end', ftDateTime);
  FExportDataset.FieldDefs.Add('Size', ftFloat);
  DeleteFile(fn);
  case RgDatabaseType.ItemIndex of
    0: TDbf(FExportDataset).CreateTable;
    1: TBufDataset(FExportDataset).CreateDataset;
  end;

  FExportDataset.Open;
  // We generate random records by combining first names, last names and cities
  // defined in the FIRST_NAMES, LAST_NAMES and CITIES arrays. We also add a
  // random birthday.

  for i:=1 to StrToInt(EdRecordCount.Text) do begin
    if (i mod 1000 = 0) then
    begin
      if FExportDataset is TBufDataset then
        TBufDataset(FExportDataset).MergeChangeLog;
      InfoLabel1.Caption := Format('Adding record %d...', [i]);
      Application.ProcessMessages;
    end;
    FExportDataset.Insert;
    FExportDataset.FieldByName('Last name').AsString := LAST_NAMES[Random(NUM_LAST_NAMES)];
    FExportDataset.FieldByName('First name').AsString := FIRST_NAMES[Random(NUM_FIRST_NAMES)];
    FExportDataset.FieldByName('City').AsString := CITIES[Random(NUM_CITIES)];
    FExportDataset.FieldByName('Birthday').AsDateTime := startDate - random(maxAge);
    FExportDataset.FieldByName('Salary').AsFloat := 1000+Random(9000);
    FExportDataset.FieldByName('Size').AsFloat := (160 + Random(50)) / 100;
    FExportDataSet.FieldByName('Work begin').AsDateTime := 40000+EncodeTime(6+Random(4), Random(60), Random(60), 0);
    FExportDataSet.FieldByName('Work end').AsDateTime := EncodeTime(15+Random(4), Random(60), Random(60), 0);
    FExportDataSet.FieldByName('Social sec').AsString := RandomNumberStr(16);
    FExportDataset.Post;
  end;

  FExportDataset.Close;

  InfoLabel1.Caption := Format('Done. Created file "%s" in folder "data".', [
    ExtractFileName(fn), ExtractFileDir(fn)
  ]);
  InfoLabel2.Caption := '';
  InfoLabel3.Caption := '';
  Application.ProcessMessages;
end;

{ This procedure exports the data in the dbf file created by BtnCreateDbfClick
  to a spreadsheet file. The workbook operates in virtual mode to minimize
  memory load of this process }
procedure TForm1.BtnExportClick(Sender: TObject);
var
  DataFileName: string; //export file name
begin
  InfoLabel2.Caption := '';
  Application.ProcessMessages;

  DataFileName := DATADIR + DirectorySeparator + TABLENAME + DB_EXT[RgDatabaseType.ItemIndex];;
  if FExportDataset = nil then
  begin
    case RgDatabaseType.ItemIndex of
      0: begin
           FExportDataset := TDbf.Create(self);
           TDbf(FExportDataset).FilePathFull := ExtractFilePath(DatafileName);
           TDbf(FExportDataset).TableName := ExtractFileName(DatafileName);
         end;
      1: begin
           FExportDataset := TBufDataset.Create(self);
           TBufDataset(FExportDataset).FileName := DatafileName;
         end;
      else
        raise Exception.Create('Database type not supported.');
    end;
  end;

  if not FileExists(DataFileName) then
  begin
    MessageDlg(Format('Database file "%s" not found. Please run "Create database" first.',
      [DataFileName]), mtError, [mbOK], 0);
    exit;
  end;

  // Make user aware export may take some time by changing cursor
  Screen.Cursor := crHourGlass;
  try
    InfoLabel1.Caption := ('Starting database export.');
    case TsExportModes(RgExportMode.ItemIndex) of
      seVirtual:
        ExportUsingVirtualMode(DataFileName);
      seFPSExport:
        ExportUsingFPSExport(false, DataFileName);
      seFPSExportMulti:
        ExportUsingFPSExport(true, DataFileName);
      else
      begin
        ShowMessageFmt('Unknown export mode number %d. Please correct source code.',[RgExportMode.ItemIndex]);
        exit;
      end;
    end;
  finally
    Screen.Cursor := crDefault;
  end;

  InfoLabel2.Caption := Format('Done. Database exported to file "%s" in folder "%s"', [
    ChangeFileExt(ExtractFileName(DataFileName), FILE_EXT[RgFileFormat.ItemIndex]),
    DATADIR
  ]);
end;

{ This procedure imports the contents of the selected spreadsheet file into a
  new dbf database file using virtual mode. }
procedure TForm1.BtnImportClick(Sender: TObject);
var
  DataFileName: String;
  fmt: TsSpreadsheetFormat;
  ext: String;
begin
  if FileList.ItemIndex = -1 then begin
    MessageDlg('Please select a file in the listbox.', mtInformation, [mbOK], 0);
    exit;
  end;

  // Determine the file format from the filename - just to avoid the annoying
  // exceptions that occur for Excel2 and Excel5.
  DataFileName := FileList.Items[FileList.ItemIndex];
  ext := lowercase(ExtractFileExt(DataFileName));
  case ext of
    '.xls':
    begin
      if pos(FILE_EXT[0], DataFileName) > 0 then
        fmt := sfExcel2
      else
      if pos(FILE_EXT[1], DataFileName) > 0 then
        fmt := sfExcel5
      else
        fmt := sfExcel8;
    end;
    '.xlsx':
      fmt := sfOOXML;
    '.ods':
      fmt := sfOpenDocument;
    else
    begin
      MessageDlg('Unknown spreadsheet file format.', mtError, [mbOK], 0);
      exit;
    end;
  end;

  // Make user aware import may take some time by changing cursor
  Screen.Cursor := crHourglass;
  try
    DataFileName := DATADIR + DirectorySeparator + DataFileName;

    // Prepare dbf table for the spreadsheet data to be imported
    if FImportDataset <> nil then
      FImportDataset.Free;
    FImportDataset := TDbf.Create(self);
    FImportDataset.FilePathFull := DATADIR + DirectorySeparator;
    FImportDataset.TableName := 'imported_' + TABLENAME + '.dbf';
    FImportDataset.TableLevel := 4; //DBase IV; most widely used.
    DeleteFile(FImportDataset.FilePathFull + FImportDataset.TableName);

    // The stringlist will temporarily store the field names ...
    if FImportedFieldNames = nil then
      FImportedFieldNames := TStringList.Create;
    FImportedFieldNames.Clear;

    // ... and this array will temporarily store the cells of the second row
    // until we have all information to create the dbf table.
    SetLength(FImportedRowCells, 0);

    // Create the workbook and activate virtual mode
    FWorkbook := TsWorkbook.Create;
    try
      FWorkbook.Options := FWorkbook.Options + [boVirtualMode];
      FWorkbook.OnReadCellData := @ReadCellDataHandler;
      // Read the data from the spreadsheet file transparently into the dbf file
      // The data are not permanently available in the worksheet and do not occupy
      // memory there - this is virtual mode.
      FWorkbook.ReadFromFile(DataFilename, fmt);
      // We close the ImportDataset after import process has finished:
      FImportDataset.Close;
      InfoLabel3.Caption := Format('Done. File "%s" imported in database "%s".',
        [ExtractFileName(DataFileName), FImportDataset.TableName]);
    finally
      FWorkbook.Free;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TForm1.ExportUsingFPSExport(MultipleSheets: Boolean;
  var DataFileName: string);
var
  Exporter: TFPSExport;
  ExportSettings: TFPSExportFormatSettings;
begin
  FExportDataset.Open;
  // TCustomDatasetExporter dsecendants like TFPSExport will start to export
  // from current record so make sure we get everything
  FExportDataset.First;

  Exporter := TFPSExport.Create(nil);
  ExportSettings := TFPSExportFormatSettings.Create(true);
  try
    // Write header row with field names
    ExportSettings.HeaderRow := true;
    case FILE_FORMATS[RgFileFormat.ItemIndex] of
      sfExcel2, sfExcel5:
        begin
          ShowMessage('Format not supported using this mode.');
          exit;
        end;
      sfExcel8:
        ExportSettings.ExportFormat := efXLS;
      sfOOXML:
        ExportSettings.ExportFormat := efXLSX;
      sfOpenDocument:
        ExportSettings.ExportFormat := efODS;
      else
        begin
          ShowMessage('Unknown export format. Please correct the source code.');
          exit;
        end;
    end;
    // Actually apply settings
    Exporter.FormatSettings := ExportSettings;

    // Write
    Exporter.Dataset := FExportDataset;
    Exporter.FileName := ChangeFileExt(DataFileName, FILE_EXT[
      RgFileFormat.ItemIndex]);

    // Export to multiple sheets
    if MultipleSheets then
    begin
      Exporter.MultipleSheets := true;
      Exporter.OnGetSheetName := @ExporterGetSheetNameHandler;

      // On the first sheet we want "Last name", "First name" and "City"
      Exporter.ExportFields.AddField('Last name');
      Exporter.ExportFields.AddField('First name');
      Exporter.ExportFields.AddField('City');
      Exporter.Execute;

      // On the second sheet we want "Last name", "First name", "Birthday" and "Social security number"
      Exporter.ExportFields.Clear;
      Exporter.ExportFields.AddField('Last name');
      Exporter.ExportFields.AddField('First name');
      Exporter.ExportFields.AddField('Birthday');
      Exporter.ExportFields.AddField('Social sec');
      Exporter.Execute;

      // On the third sheet we want "Last name", "First name" and "Income"
      Exporter.ExportFields.Clear;
      Exporter.ExportFields.AddField('Last name');
      Exporter.ExportFields.AddField('First name');
      Exporter.ExportFields.AddField('Salary');
      Exporter.Execute;

      // On the 4th sheet we want "Last name", "First name" and "Work begin/end times"
      Exporter.ExportFields.Clear;
      Exporter.ExportFields.AddField('Last name');
      Exporter.ExportFields.AddField('First name');
      Exporter.ExportFields.AddField('Work begin');
      Exporter.ExportFields.AddField('Work end');
      Exporter.Execute;

      // On the 5th sheet we want "Last name", "First name" and "Size"
      Exporter.ExportFields.Clear;
      Exporter.ExportFields.AddField('Last name');
      Exporter.ExportFields.AddField('First name');
      Exporter.ExportFields.AddField('Size');
      Exporter.Execute;

      // Export complete --> we can write to file
      Exporter.WriteExportFile;
    end
    // Export of all records to single sheet
    else
      Exporter.Execute;
  finally
    Exporter.Free;
    ExportSettings.Free;
  end;
end;

procedure TForm1.FileListClick(Sender: TObject);
begin
  BtnImport.Enabled := (FileList.ItemIndex > -1);
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  InfoLabel1.Caption := '';
  InfoLabel2.Caption := '';
  InfoLabel3.Caption := '';
  PageControl.ActivePageIndex := 0;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  FreeAndNil(FImportedFieldNames);
end;

{ When we activate the "Import" page of the pagecontrol we read the data
  folder and collect all spreadsheet files available in a list box. The user
  will have to select the one to be converted to dbf. }
procedure TForm1.PageControlChange(Sender: TObject);
var
  sr: TSearchRec;
  ext: String;
begin
  if PageControl.ActivePage = TabImport then begin
    FileList.Clear;
    if FindFirst(DATADIR + DirectorySeparator + TABLENAME + '*.*', faAnyFile, sr) = 0
    then begin
      repeat
        if (sr.Name = '.') or (sr.Name = '..') then
          Continue;
        ext := lowercase(ExtractFileExt(sr.Name));
        if (ext = '.xls') or (ext = '.xlsx') or (ext = '.ods') then
          FileList.Items.Add(sr.Name);
      until FindNext(sr) <> 0;
      FindClose(sr);
    end;
    BtnImport.Enabled := FileList.ItemIndex > -1;
  end;
end;

procedure TForm1.ExportUsingVirtualMode(var DataFileName: string);
var
  worksheet: TsWorksheet;
  tablename: String;
begin
  {
  if FILE_FORMATS[RgFileFormat.ItemIndex] = sfOpenDocument then
  begin
    MessageDlg('Virtual mode is not yet implemented for .ods files.', mtError, [mbOK], 0);
    exit;
  end;
   }
  FExportDataset.Open;

  case RgDatabaseType.ItemIndex of
    0: tablename := TDbf(FExportDataset).TableName;
    1: tablename := ExtractFilename(TBufDataset(FExportDataset).FileName);
    else raise Exception.Create('Database type not supported.');
  end;

  FWorkbook := TsWorkbook.Create;
  try
    FWorkbook.FormatSettings.ShortDateFormat := 'dd.mm.yyyy "r."';
    worksheet := FWorkbook.AddWorksheet(tableName);

    // Make header line frozen - but not in Excel2 where frozen panes do not yet work properly
    if FILE_FORMATS[RgFileFormat.ItemIndex] <> sfExcel2 then begin
      worksheet.Options := worksheet.Options + [soHasFrozenPanes];
      worksheet.TopPaneHeight := 1;
    end;

    // Use cell A1 as format template of header line
    FHeaderTemplateCell := worksheet.GetCell(0, 0);
    worksheet.WriteFontStyle(FHeaderTemplateCell, [fssBold]);
    worksheet.WriteBackgroundColor(FHeaderTemplateCell, scGray);
    if FILE_FORMATS[RgFileFormat.ItemIndex] <> sfExcel2 then
      worksheet.WriteFontColor(FHeaderTemplateCell, scWhite);  // Does not look nice in the limited Excel2 format

    // Use cell B1 as format template of date column
    FDateTemplateCell := worksheet.GetCell(0, 1);
    worksheet.WriteDateTimeFormat(FDateTemplateCell, nfShortDate);

    // Use cell C1 as format template of currency column
    FCurrencyTemplateCell := worksheet.GetCell(0, 2);
    worksheet.WriteNumberFormat(FCurrencyTemplateCell, nfCurrency);

    // Use cell D1 as format template for social security number column
    FSocialSecurityTemplatecell := worksheet.GetCell(0, 3);
    worksheet.WriteNumberFormat(FSocialSecurityTemplateCell, nfText);

    // Make rows a bit wider
    worksheet.WriteColWidth(0, 20);
    worksheet.WriteColWidth(1, 20);
    worksheet.WriteColWidth(2, 20);
    worksheet.WriteColWidth(3, 15);

    // Setup virtual mode to save memory
//    FWorkbook.Options := FWorkbook.Options + [boVirtualMode, boBufStream];
    FWorkbook.Options := FWorkbook.Options + [boVirtualMode];
    worksheet.OnWriteCellData := @WriteCellDataHandler;
    worksheet.VirtualRowCount := FExportDataset.RecordCount + 1;  // +1 for the header line
    worksheet.VirtualColCount := FExportDataset.FieldCount;

    // Write
    DataFileName := ChangeFileExt(DataFileName, FILE_EXT[
      RgFileFormat.ItemIndex]);
    FWorkbook.WriteToFile(DataFileName, FILE_FORMATS[RgFileFormat.ItemIndex],
      true);
  finally
    FreeAndNil(FWorkbook);
  end;
end;

{ Determines the sheet name of the export using FPExport }
procedure TForm1.ExporterGetSheetNameHandler(Sender: TObject; ASheetIndex: Integer;
  var ASheetName: String);
begin
  case ASheetIndex of
    0: ASheetName := 'City';
    1: ASheetName := 'Birthday';
    2: ASheetName := 'Salary';
    3: ASheetName := 'Work time';
    4: ASheetName := 'Size';
  end;
end;

{ This is the event handler for reading a spreadsheet file in virtual mode.
  ADataCell has just been read from the spreadsheet file, but will not be added
  to the workbook and will be discarded. The event handler, however, can pick
  the data and post them to the database table.
  Note that we do not make too many assumptions on the data structure here.
  Therefore we have to buffer the first two rows of the spreadsheet file until
  the structure of the table is clear. }
procedure TForm1.ReadCellDataHandler(Sender: TObject; ARow, ACol: Cardinal;
  const ADataCell: PCell);
var
  i: Integer;
  fieldType: TFieldType;
  fmt: TsCellFormat;
  nfp: TsNumFormatParams;
begin
  // The first row (index 0) holds the field names. We temporarily store the
  // field names in a stringlist because we don't know the data types of the
  // cell until we have read the second row (index 1).
  if ARow = 0 then begin
    // We know that the first row contains string cells -> no further checks.
    FImportedFieldNames.Add(ADataCell^.UTF8StringValue);
  end
  else
  // We have to buffer the second row (index 1) as well. When it is fully read
  // we can put everything together and create the dbf table.
  if ARow = 1 then begin
    if Length(FImportedRowCells) = 0 then
      SetLength(FImportedRowCells, FImportedFieldNames.Count);
    FImportedRowCells[ACol] := ADataCell^;
    // The row is read completely, all field types are known --> we create the
    // table
    if ACol = High(FImportedRowCells) then begin
      // Add fields - the required information is stored in FImportedFieldNames
      // and FImportedFieldTypes
      for i:=0 to High(FImportedRowCells) do begin
        fmt := TsWorksheet(ADataCell^.Worksheet).ReadCellFormat(ADataCell);
        nfp := FWorkbook.GetNumberFormat(fmt.NumberFormatIndex);
        case FImportedRowCells[i].ContentType of
          cctNumber     : if IsCurrencyFormat(nfp) then fieldType := ftCurrency
                            else fieldType := ftFloat;
          cctDateTime   : fieldType := ftDateTime;
          else            fieldType := ftString;
        end;
        FImportDataset.FieldDefs.Add(FImportedFieldNames[i], fieldType);
      end;
      // Create the table and open it
      DeleteFile(FImportDataset.FilePathFull + FImportDataset.TableName);
      FImportDataset.CreateTable;
      FImportDataset.Open;
      // Now we have to post the cells of the buffered row, otherwise these data
      // will be lost
      FImportDataset.Insert;
      for i:=0 to High(FImportedRowCells) do
        case FImportedRowCells[i].ContentType of
          cctNumber    : FImportDataset.Fields[i].AsFloat := FImportedRowCells[i].NumberValue;
          cctDateTime  : FImportDataset.Fields[i].AsDateTime := FImportedRowCells[i].DateTimeValue;
          else           FImportDataset.Fields[i].AsString := FImportedRowCells[i].UTF8StringValue;
        end;
      FImportDataset.Post;
      // Finally we dispose the buffered cells, we don't need them any more
      SetLength(FImportedRowCells, 0);
    end;
  end
  else
  begin
    // Now that we know everything we can add the data to the table
    if ARow mod 25 = 0 then
    begin
      InfoLabel3.Caption := Format('Writing row %d to database...', [ARow]);
      Application.ProcessMessages;
    end;

    if ACol = 0 then
      FImportDataset.Insert;
    case ADataCell^.ContentType of
      cctNumber    : FImportDataSet.Fields[Acol].AsFloat :=  ADataCell^.NumberValue;
      cctUTF8String: FImportDataset.Fields[Acol].AsString := ADataCell^.UTF8StringValue;
      cctDateTime  : FImportDataset.Fields[ACol].AsDateTime := ADataCell^.DateTimeValue;
    end;
    if ACol = FImportedFieldNames.Count-1 then
      FImportDataset.Post;   // We post the data after the last cell of the row has been received.
  end;
end;

procedure TForm1.RgFileFormatSelectionChanged(Sender: TObject);
begin
  RgExportMode.Controls[2].Enabled := RgFileFormat.ItemIndex <> 0;
end;

{ This is the event handler for exporting a database file to spreadsheet format
  in virtual mode. Data are not written into the worksheet, they exist only
  temporarily. }
procedure TForm1.WriteCellDataHandler(Sender: TsWorksheet; ARow, ACol: Cardinal;
  var AValue: variant; var AStyleCell: PCell);
begin
  // Header line: we want to show the field names here.
  if ARow = 0 then
  begin
    AValue := FExportDataset.Fields[ACol].FieldName;
    AStyleCell := FHeaderTemplateCell;
    FExportDataset.First;
  end
  else
  // After the header line we write the record data. Note that we are
  // responsible for advancing the dataset cursor whenever a row is complete.
  begin
    AValue := FExportDataset.Fields[ACol].Value;
    if FExportDataset.Fields[ACol].DataType = ftDate then
      AStyleCell := FDateTemplateCell
    else if FExportDataset.Fields[ACol].DataType = ftCurrency then
      AStyleCell := FCurrencyTemplateCell
    else if SameText(FExportDataset.Fields[ACol].FieldName, 'Social sec') then
      AStyleCell := FSocialSecurityTemplateCell;
    if ACol = Sender.VirtualColCount-1 then
    begin
      // Move to next record after last field has been written
      FExportDataset.Next;
      // Progress display
      if (ARow-1) mod 1000 = 0 then
      begin
        InfoLabel2.Caption := Format('Writing record %d to spreadsheet...', [ARow-1]);
        Application.ProcessMessages;
      end;
    end;
  end;
end;

end.

