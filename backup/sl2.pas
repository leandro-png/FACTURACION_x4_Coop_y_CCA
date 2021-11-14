unit SL2;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  ExtCtrls, Grids, ComCtrls, fpspreadsheetgrid, fpspreadsheet,
  fpspreadsheetctrls, xlsbiff8, LazUTF8, fpsTypes, dateutils, strutils;

type

  { Txls }

  Txls = class(TForm)
    ApplicationProperties1: TApplicationProperties;
    boton_archivo_1: TButton;
    boton_archivo_2: TButton;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    Button6: TButton;
    Edit1: TEdit;
    grilla2: TStringGrid;
    grilla3: TStringGrid;
    grilla4: TStringGrid;
    grilla5: TStringGrid;
    grilla6: TStringGrid;
    grilla7: TStringGrid;
    grilla8: TStringGrid;
    grilla9: TStringGrid;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label2: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    Label27: TLabel;
    Label28: TLabel;
    Label29: TLabel;
    Label3: TLabel;
    Label30: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    OpenDialog1: TOpenDialog;
    grilla1: TStringGrid;
    StringGrid1: TStringGrid;
    procedure boton_archivo_1Click(Sender: TObject);
    procedure boton_archivo_2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Edit1Change(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Label10Click(Sender: TObject);
    procedure Label1Click(Sender: TObject);
    procedure Label9Click(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  xls: Txls;
  Mitrabajo: TsWorkbook;
  Miplanilla: TsWorksheet;


implementation

{$R *.lfm}

{ Txls }

procedure Txls.boton_archivo_1Click(Sender: TObject);


var
    nombre_de_archivo_1 , tmp1 , tmp2:         string;
    scan_fila , x , y , filas_grilla2 ,lugar:  integer;
    Archivo_1_de_excel: TsWorkbook;
    planilla_dentro_de_Archivo_1_de_excel: TsWorksheet;



begin
  opendialog1.execute;
  if opendialog1.FileName <>'' then begin

                                         nombre_de_archivo_1 :=opendialog1.FileName ;
                                         label2.caption:='Archivo: '+(ExtractFilename(opendialog1.Filename));
                                         Archivo_1_de_excel := TsWorkbook.Create;

                                         Archivo_1_de_excel.Options := Archivo_1_de_excel.Options + [boReadFormulas];

                                         Archivo_1_de_excel.ReadFromFile(nombre_de_archivo_1, sfExcel8);    //sfOOxml
                                         planilla_dentro_de_Archivo_1_de_excel := Archivo_1_de_excel.GetWorksheetByName('Sheet');

                                         scan_fila:=0;


                                          boton_archivo_1.visible:=false;
                                          boton_archivo_2.visible:=true;
                                          button2.visible:=true;


                                         ////////////ENCUENTRA LA CANTIDAD MAXIMA DE FILAS/////////////////////////
                                          repeat
                                          inc(scan_fila);
                                          until (planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(scan_fila+1,0)='')and(scan_fila>30);


                                         //   if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(4,0)='' then lugar:=10 else lugar:=8;
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(7,0)='Fecha Hora' then lugar:=7 else
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(8,0)='Fecha Hora' then lugar:=8 else
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(9,0)='Fecha Hora' then lugar:=9 else
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(10,0)='Fecha Hora' then lugar:=10 else
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(11,0)='Fecha Hora' then lugar:=11 else
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(12,0)='Fecha Hora' then lugar:=12   ;



                                          //    label9.caption:=planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(8,0);
                                              scan_fila:=scan_fila-lugar;
                                         ////////////ENCUENTRA LA CANTIDAD MAXIMA DE FILAS/////////////////////////
                                       //  label1.caption:='Filas: '+inttostr(scan_fila);
                                          ///////////////ACOMODA TAMAÑO DE LA GRILLA/////////////////////////////////////////////
                                          grilla1.rowcount:=scan_fila+1;
                                          grilla1.colcount:=4;
                                          ///////////////ACOMODA TAMAÑO DE LA GRILLA/////////////////////////////////////////////


                                          //=====================PASA LOS DATOS A LA GRILLA1 =====================

                                          for y:=0 to scan_fila do
                                          begin

                                            grilla1.Cells[0,y]:=planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(y+lugar,0);
                                            tmp1:=grilla1.Cells[0,y][1]+grilla1.Cells[0,y][2]+grilla1.Cells[0,y][3]+grilla1.Cells[0,y][4]+grilla1.Cells[0,y][5]+grilla1.Cells[0,y][6]+'20'+grilla1.Cells[0,y][7]+grilla1.Cells[0,y][8];
                                            tmp2:=grilla1.Cells[0,y][10]+grilla1.Cells[0,y][11]+grilla1.Cells[0,y][12]+grilla1.Cells[0,y][13]+grilla1.Cells[0,y][14];
                                            grilla1.Cells[0,y]:=tmp1;
                                            grilla1.Cells[1,y]:=tmp2+':00';
                                            grilla1.Cells[2,y]:=planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(y+lugar,9);
                                            grilla1.Cells[3,y]:=planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(y+lugar,11);
                                           end;



                                             grilla1.Cells[0,0]:='Fecha';
                                             grilla1.Cells[1,0]:='Hora';
                                           //=====================PASA LOS DATOS A LA GRILLA1 =====================

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-     AHORA EMPIEZA LA GRILLA 2       -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
x:=3;

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=- }
                                          filas_grilla2:=1;
                                          grilla2.rowcount:=scan_fila+50000;
                                          grilla2.colcount:=4;
                                          grilla2.Cells[1,0]:=timetostr(INCMINUTE(strtotime(grilla1.Cells[1,1]),-15));
                                          grilla2.Cells[0,0]:=grilla1.Cells[0,1];
                                          y:=1;
                                          repeat
                                            begin                         for x:=0 to 3 do  grilla2.Cells[x,filas_grilla2]:=grilla1.Cells[x,y];

                                                                          if strtodate(grilla2.Cells[0,filas_grilla2])=strtodate(grilla2.Cells[0,filas_grilla2-1])then BEGIN     //si los dias son =

                                                                          if ((strtotime(grilla2.Cells[1,filas_grilla2])-strtotime(grilla2.Cells[1,filas_grilla2-1])=strtotime('00:00:00')))then grilla2.Cells[1,filas_grilla2]:=timetostr((strtotime(grilla2.Cells[1,filas_grilla2-1]))) else

                                                                          if ((strtotime(grilla2.Cells[1,filas_grilla2])-strtotime(grilla2.Cells[1,filas_grilla2-1])<strtotime('00:00:00')))then begin   if (grilla2.Cells[1,filas_grilla2])<>'00:00:00' then grilla2.Cells[1,filas_grilla2]:=grilla2.Cells[1,filas_grilla2-1];
                                                                                                                                                                                                 end else

                                                                          if ((strtotime(grilla2.Cells[1,filas_grilla2])-strtotime(grilla2.Cells[1,filas_grilla2-1])>strtotime('00:15:10')))then  begin
                                                                                                                                                                                                       for x:=0 to 3 do  grilla2.Cells[x,filas_grilla2+1]:=grilla2.Cells[x,filas_grilla2];
                                                                                                                                                                                                       grilla2.Cells[1,filas_grilla2]:=timetostr(INCMINUTE(strtotime(grilla2.Cells[1,filas_grilla2-1]),15));
                                                                                                                                                                                                       for x:=2 to 3 do  grilla2.Cells[x,filas_grilla2+1]:='0';
                                                                                                                                                                                                       for x:=2 to 3 do  grilla2.Cells[x,filas_grilla2]:='0';
                                                                                                                                                                                                       dec(y);

                                                                                                                                                                                                  end;
                                                                                                                                                                         END else      //si los dias son distintos



                                                                                                                                                                         BEGIN                         if (grilla2.Cells[1,filas_grilla2])='00:00:00'then        else
                                                                                                                                                                                                      if (((strtotime(grilla2.Cells[1,filas_grilla2])-strtotime(grilla2.Cells[1,filas_grilla2-1])<>strtotime('00:15:00')))) then  begin


                                                                                                                                                                                                      for x:=0 to 3 do  grilla2.Cells[x,filas_grilla2+1]:=grilla2.Cells[x,filas_grilla2];



                                                                                                                                                                                                      repeat

                                                                                                                                                                                                    //  grilla2.Cells[0,filas_grilla2]:=grilla2.Cells[0,filas_grilla2-1];
                                                                                                                                                                                                      grilla2.Cells[1,filas_grilla2]:=timetostr(INCMINUTE(strtotime(grilla2.Cells[1,filas_grilla2-1]),15));
                                                                                                                                                                                                      grilla2.Cells[0,filas_grilla2]:=grilla2.Cells[0,filas_grilla2-1];
                                                                                                                                                                                                      for x:=2 to 3 do  grilla2.Cells[x,filas_grilla2]:='0';
                                                                                                                                                                                                      inc(filas_grilla2);
                                                                                                                                                                                                      until grilla2.Cells[1,filas_grilla2-1]='23:45:00';


                                                                                                                                                                                                   //   grilla2.Cells[0,filas_grilla2]:=grilla2.Cells[0,filas_grilla2-1];
                                                                                                                                                                                                      grilla2.Cells[1,filas_grilla2]:='00:00:00';
                                                                                                                                                                                                      for x:=2 to 3 do  grilla2.Cells[x,filas_grilla2]:='0';
                                                                                                                                                                                                      grilla2.Cells[0,filas_grilla2]:=datetostr(INCday(strtodate(grilla2.Cells[0,filas_grilla2-1]),1));     //repetido
                                                                                                                                                                                                      inc(filas_grilla2);

                                                                                                                                                                                                      for x:=0 to 3 do  grilla2.Cells[x,filas_grilla2]:=grilla2.Cells[x,filas_grilla2];


                                                                                                                                                                                                      grilla2.Cells[0,filas_grilla2]:=datetostr(INCday(strtodate(grilla2.Cells[0,filas_grilla2-1]),0));
                                                                                                                                                                                                      grilla2.Cells[1,filas_grilla2]:='00:15:00';
                                                                                                                                                                                                      for x:=2 to 3 do  grilla2.Cells[x,filas_grilla2]:='0';
                                                                                                                                                                                                      dec(y);




                                                                                                                                                                                                                                                                                                                                end;

                                                                                                                                                                        END;
                                                                          inc(filas_grilla2);
                                                                          inc(y);
                                            end;
                                       until y=scan_fila+1;



     grilla2.rowcount:=filas_grilla2;




  grilla2.Cells[0,0]:='Fecha';
  grilla2.Cells[1,0]:='Hora';
  grilla2.Cells[2,0]:='Wh';
  grilla2.Cells[3,0]:='VARh';

  label3.caption:='TERMINAL: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(2,1);
  label4.caption:='DOMICILIO: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(3,1);
  label5.caption:='SERVICIO: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(4,1);
  label6.caption:='TITULAR: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(2,10);

  if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(3,10)<>'' then
  label7.caption:='PERIODO: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(3,10)
  else label7.caption:='PERIODO: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(4,10);

  label8.caption:='Antes: '+inttostr(scan_fila)+', después: '+inttostr(filas_grilla2-1)+' ('+floattostr((filas_grilla2-1)/96)+' días)';
  if int((filas_grilla2-1)/96)=((filas_grilla2-1)/96)then
  Edit1.text:=floattostr((filas_grilla2-1)/96)
  else  showmessage('Atención, faltan datos en el archivo XLS, pues al dividir la cantidad de datos totales por 96, no dá un entero. Completalos manualmente editando el XLS');


     end;

end;

procedure Txls.boton_archivo_2Click(Sender: TObject);


var
    nombre_de_archivo_1 , tmp1 , tmp2:         string;
    scan_fila , x , y , filas_grilla4 ,lugar:  integer;
    Archivo_1_de_excel: TsWorkbook;
    planilla_dentro_de_Archivo_1_de_excel: TsWorksheet;



begin
  opendialog1.execute;
  if opendialog1.FileName <>'' then begin

                                         nombre_de_archivo_1 :=opendialog1.FileName ;
                                         label15.caption:='Archivo: '+(ExtractFilename(opendialog1.Filename));
                                         Archivo_1_de_excel := TsWorkbook.Create;

                                         Archivo_1_de_excel.Options := Archivo_1_de_excel.Options + [boReadFormulas];

                                         Archivo_1_de_excel.ReadFromFile(nombre_de_archivo_1, sfExcel8);    //sfOOxml
                                         planilla_dentro_de_Archivo_1_de_excel := Archivo_1_de_excel.GetWorksheetByName('Sheet');

                                         scan_fila:=0;


                                          boton_archivo_2.visible:=false;
                                          button4.visible:=true;


                                         ////////////ENCUENTRA LA CANTIDAD MAXIMA DE FILAS/////////////////////////
                                          repeat
                                          inc(scan_fila);
                                          until (planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(scan_fila+1,0)='')and(scan_fila>30);


                                         //   if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(4,0)='' then lugar:=10 else lugar:=8;
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(7,0)='Fecha Hora' then lugar:=7 else
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(8,0)='Fecha Hora' then lugar:=8 else
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(9,0)='Fecha Hora' then lugar:=9 else
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(10,0)='Fecha Hora' then lugar:=10 else
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(11,0)='Fecha Hora' then lugar:=11 else
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(12,0)='Fecha Hora' then lugar:=12   ;



                                          //    label9.caption:=planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(8,0);
                                              scan_fila:=scan_fila-lugar;
                                         ////////////ENCUENTRA LA CANTIDAD MAXIMA DE FILAS/////////////////////////
                                        // label1.caption:='Filas: '+inttostr(scan_fila);
                                          ///////////////ACOMODA TAMAÑO DE LA GRILLA/////////////////////////////////////////////
                                          grilla3.rowcount:=scan_fila+1;
                                          grilla3.colcount:=4;
                                          ///////////////ACOMODA TAMAÑO DE LA GRILLA/////////////////////////////////////////////


                                          //=====================PASA LOS DATOS A LA GRILLA3 =====================

                                          for y:=0 to scan_fila do
                                          begin

                                            grilla3.Cells[0,y]:=planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(y+lugar,0);
                                            tmp1:=grilla3.Cells[0,y][1]+grilla3.Cells[0,y][2]+grilla3.Cells[0,y][3]+grilla3.Cells[0,y][4]+grilla3.Cells[0,y][5]+grilla3.Cells[0,y][6]+'20'+grilla3.Cells[0,y][7]+grilla3.Cells[0,y][8];
                                            tmp2:=grilla3.Cells[0,y][10]+grilla3.Cells[0,y][11]+grilla3.Cells[0,y][12]+grilla3.Cells[0,y][13]+grilla3.Cells[0,y][14];
                                            grilla3.Cells[0,y]:=tmp1;
                                            grilla3.Cells[1,y]:=tmp2+':00';
                                            grilla3.Cells[2,y]:=planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(y+lugar,9);
                                            grilla3.Cells[3,y]:=planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(y+lugar,11);
                                           end;



                                             grilla3.Cells[0,0]:='Fecha';
                                             grilla3.Cells[1,0]:='Hora';
                                           //=====================PASA LOS DATOS A LA GRILLA3 =====================

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-     AHORA EMPIEZA LA GRILLA 4       -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
x:=3;

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=- }
                                          filas_grilla4:=1;
                                          grilla4.rowcount:=scan_fila+50000;
                                          grilla4.colcount:=4;
                                          grilla4.Cells[1,0]:=timetostr(INCMINUTE(strtotime(grilla3.Cells[1,1]),-15));
                                          grilla4.Cells[0,0]:=grilla3.Cells[0,1];
                                          y:=1;
                                          repeat
                                            begin                         for x:=0 to 3 do  grilla4.Cells[x,filas_grilla4]:=grilla3.Cells[x,y];

                                                                          if strtodate(grilla4.Cells[0,filas_grilla4])=strtodate(grilla4.Cells[0,filas_grilla4-1])then BEGIN     //si los dias son =

                                                                          if ((strtotime(grilla4.Cells[1,filas_grilla4])-strtotime(grilla4.Cells[1,filas_grilla4-1])=strtotime('00:00:00')))then grilla4.Cells[1,filas_grilla4]:=timetostr((strtotime(grilla4.Cells[1,filas_grilla4-1]))) else

                                                                          if ((strtotime(grilla4.Cells[1,filas_grilla4])-strtotime(grilla4.Cells[1,filas_grilla4-1])<strtotime('00:00:00')))then begin   if (grilla4.Cells[1,filas_grilla4])<>'00:00:00' then grilla4.Cells[1,filas_grilla4]:=grilla4.Cells[1,filas_grilla4-1];
                                                                                                                                                                                                 end else

                                                                          if ((strtotime(grilla4.Cells[1,filas_grilla4])-strtotime(grilla4.Cells[1,filas_grilla4-1])>strtotime('00:15:10')))then  begin
                                                                                                                                                                                                       for x:=0 to 3 do  grilla4.Cells[x,filas_grilla4+1]:=grilla4.Cells[x,filas_grilla4];
                                                                                                                                                                                                       grilla4.Cells[1,filas_grilla4]:=timetostr(INCMINUTE(strtotime(grilla4.Cells[1,filas_grilla4-1]),15));
                                                                                                                                                                                                       for x:=2 to 3 do  grilla4.Cells[x,filas_grilla4+1]:='0';
                                                                                                                                                                                                       for x:=2 to 3 do  grilla4.Cells[x,filas_grilla4]:='0';
                                                                                                                                                                                                       dec(y);

                                                                                                                                                                                                  end;
                                                                                                                                                                         END else      //si los dias son distintos



                                                                                                                                                                         BEGIN                         if (grilla4.Cells[1,filas_grilla4])='00:00:00'then        else
                                                                                                                                                                                                      if (((strtotime(grilla4.Cells[1,filas_grilla4])-strtotime(grilla4.Cells[1,filas_grilla4-1])<>strtotime('00:15:00')))) then  begin


                                                                                                                                                                                                      for x:=0 to 3 do  grilla4.Cells[x,filas_grilla4+1]:=grilla4.Cells[x,filas_grilla4];



                                                                                                                                                                                                      repeat

                                                                                                                                                                                                    //  grilla4.Cells[0,filas_grilla4]:=grilla4.Cells[0,filas_grilla4-1];
                                                                                                                                                                                                      grilla4.Cells[1,filas_grilla4]:=timetostr(INCMINUTE(strtotime(grilla4.Cells[1,filas_grilla4-1]),15));
                                                                                                                                                                                                      grilla4.Cells[0,filas_grilla4]:=grilla4.Cells[0,filas_grilla4-1];
                                                                                                                                                                                                      for x:=2 to 3 do  grilla4.Cells[x,filas_grilla4]:='0';
                                                                                                                                                                                                      inc(filas_grilla4);
                                                                                                                                                                                                      until grilla4.Cells[1,filas_grilla4-1]='23:45:00';


                                                                                                                                                                                                   //   grilla4.Cells[0,filas_grilla4]:=grilla4.Cells[0,filas_grilla4-1];
                                                                                                                                                                                                      grilla4.Cells[1,filas_grilla4]:='00:00:00';
                                                                                                                                                                                                      for x:=2 to 3 do  grilla4.Cells[x,filas_grilla4]:='0';
                                                                                                                                                                                                      grilla4.Cells[0,filas_grilla4]:=datetostr(INCday(strtodate(grilla4.Cells[0,filas_grilla4-1]),1));     //repetido
                                                                                                                                                                                                      inc(filas_grilla4);

                                                                                                                                                                                                      for x:=0 to 3 do  grilla4.Cells[x,filas_grilla4]:=grilla4.Cells[x,filas_grilla4];


                                                                                                                                                                                                      grilla4.Cells[0,filas_grilla4]:=datetostr(INCday(strtodate(grilla4.Cells[0,filas_grilla4-1]),0));
                                                                                                                                                                                                      grilla4.Cells[1,filas_grilla4]:='00:15:00';
                                                                                                                                                                                                      for x:=2 to 3 do  grilla4.Cells[x,filas_grilla4]:='0';
                                                                                                                                                                                                      dec(y);




                                                                                                                                                                                                                                                                                                                                end;

                                                                                                                                                                        END;
                                                                          inc(filas_grilla4);
                                                                          inc(y);
                                            end;
                                       until y=scan_fila+1;



     grilla4.rowcount:=filas_grilla4;




  grilla4.Cells[0,0]:='Fecha';
  grilla4.Cells[1,0]:='Hora';
  grilla4.Cells[2,0]:='Wh';
  grilla4.Cells[3,0]:='VARh';

  label9.caption:='TERMINAL: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(2,1);
  label10.caption:='DOMICILIO: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(3,1);
  label11.caption:='SERVICIO: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(4,1);
  label12.caption:='TITULAR: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(2,10);

  if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(3,10)<>'' then
  label13.caption:='PERIODO: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(3,10)
  else label13.caption:='PERIODO: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(4,10);

  label14.caption:='Antes: '+inttostr(scan_fila)+', después: '+inttostr(filas_grilla4-1)+' ('+floattostr((filas_grilla4-1)/96)+' días)';

  if (int((filas_grilla4-1)/96)=((filas_grilla4-1)/96))then  begin
              if edit1.text<>floattostr((filas_grilla4-1)/96) then showmessage('Atención: Parece que los dos archivos tienen cantidad de días distintos, fijate que correspondan al mismo mes.');
                                                             end
  else  showmessage('Atención, faltan datos en el archivo XLS, pues al dividir la cantidad de datos totales por 96, no dá un entero. Completalos manualmente editando el XLS');


     end;

end;














procedure Txls.Button1Click(Sender: TObject);
var a,dias,TOTAL:integer;//  temp:extended;
    suma_valle,suma_pico,suma_resto:extended;
    activa_total,reactiva_total:extended;
    MAX_pico,MAX_resto,MAX_valle:extended;


begin
   stringgrid1.Visible:=true;
  TOTAL:=strtoint(edit1.Text);
  grilla5.Cells[0,0]:='A1*4';
  grilla5.Cells[1,0]:='A2*4';

  grilla5.Cells[10,0]:='A3*4';
  grilla5.Cells[11,0]:='A4*4';

  grilla5.Cells[2,0]:='Demanda';
  grilla5.Cells[3,0]:='Activa';
  grilla5.Cells[4,0]:='Reactiva';
  grilla5.Cells[5,0]:='Pico';
  grilla5.Cells[6,0]:='Resto';
  grilla5.Cells[7,0]:='Valle';
  activa_total:=0;
  reactiva_total:=0;
  suma_valle:=0;
  suma_pico:=0;
  suma_resto:=0;

  MAX_pico:=0;
  MAX_resto:=0;
  MAX_valle:=0;


  button3.visible:=true;

  for a:=1 to TOTAL*96 do begin
     {MEDIDOR1}       grilla5.cells[0,a]:=floattostr(4*strtofloat(grilla2.Cells[2,a]));
     {MEDIDOR2}       grilla5.cells[1,a]:=floattostr(4*strtofloat(grilla4.Cells[2,a]));

     {MEDIDOR3}       grilla5.cells[10,a]:=floattostr(4*strtofloat(grilla7.Cells[2,a]));
     {MEDIDOR4}       grilla5.cells[11,a]:=floattostr(4*strtofloat(grilla9.Cells[2,a]));

     {DEMANDA}        grilla5.cells[2,a]:=floattostr(4*strtofloat(grilla2.Cells[2,a])+4*strtofloat(grilla4.Cells[2,a])+4*strtofloat(grilla7.Cells[2,a])+4*strtofloat(grilla9.Cells[2,a]));        {+grilla7+grilla9}
     {ACTIVA}         grilla5.cells[3,a]:=floattostr(strtofloat(grilla2.Cells[2,a])+strtofloat(grilla4.Cells[2,a])+strtofloat(grilla7.Cells[2,a])+strtofloat(grilla9.Cells[2,a]));            {+grilla7+grilla9}
     {REACTIVA}       grilla5.cells[4,a]:=floattostr(strtofloat(grilla2.Cells[3,a])+strtofloat(grilla4.Cells[3,a])+strtofloat(grilla7.Cells[3,a])+strtofloat(grilla9.Cells[3,a]));            {+grilla7+grilla9}

     {FECHA}          grilla5.cells[8,a]:=grilla2.cells[0,a];
     {HORA}           grilla5.cells[9,a]:=grilla2.cells[1,a];

                      if grilla5.cells[3,a]<>'' then activa_total  :=  activa_total+strtofloat(grilla5.cells[3,a]);
                      if grilla5.cells[4,a]<>'' then reactiva_total:=reactiva_total+strtofloat(grilla5.cells[4,a]);

  end;

// PRIMERA PORCION DEL DIA /////////////////////////////////////////////////////////////////////////////////////////////////////////
for a:=1 to 20 do begin
                        grilla5.cells[7,a]:=grilla5.cells[3,a];//valle incompleto
                        suma_valle:= suma_valle+strtofloat(grilla5.cells[3,a]);
                  end;
// PRIMERA PORCION DEL DIA /////////////////////////////////////////////////////////////////////////////////////////////////////////





// DIAS COMPLETOS /////////////////////////////////////////////////////////////////////////////////////////////////////////
for dias:=0 to TOTAL do begin
  for a:=1 to 52 do begin
                          grilla5.cells[6,a+20+dias*96]:=grilla5.cells[3,a+20+dias*96];             //resto
                          if grilla5.cells[6,a+20+dias*96]<>'' then
                          suma_resto:= suma_resto+strtofloat(grilla5.cells[6,a+20+dias*96]);        //SUMA RESTO
                    end;

  for a:=1 to 20 do begin
                         grilla5.cells[5,a+72+dias*96]:=grilla5.cells[3,a+72+dias*96];              //pico
                         if grilla5.cells[5,a+72+dias*96]<>'' then
                         suma_pico:= suma_pico+strtofloat(grilla5.cells[5,a+72+dias*96]);           //SUMA PICO
                    end;
  for a:=1 to 24 do begin
                         grilla5.cells[7,a+92+dias*96]:=grilla5.cells[3,a+92+dias*96];              //valle
                         if grilla5.cells[7,a+92+dias*96]<>'' then
                         suma_valle:= suma_valle+strtofloat(grilla5.cells[7,a+92+dias*96]);         //SUMA VALLE
                    end;
                    end;
// DIAS COMPLETOS /////////////////////////////////////////////////////////////////////////////////////////////////////////




//ULTIMA PORCION DEL ULTIMO DIA /////////////////////////////////////////////////////////////////////////////////////////////////////////
   for a:=1 to 52 do grilla5.cells[6,a+20+(TOTAL+1)*96]:=grilla5.cells[3,a+20+(TOTAL+1)*96];             //resto ultimo
   if  grilla5.cells[5,a+20+(TOTAL+1)*96]<>'' then                                                       //SUMA RESTO
   suma_resto:= suma_resto+strtofloat(grilla5.cells[5,a+20+(TOTAL+1)*96]);                                 //SUMA RESTO

   for a:=1 to 24 do grilla5.cells[5,a+72+(TOTAL+1)*96]:=grilla5.cells[3,a+72+(TOTAL+1)*96];             //pico incompleto
   if  grilla5.cells[5,a+72+(TOTAL+1)*96]<>'' then                                                       //SUMA PICO
   suma_pico:= suma_pico+strtofloat(grilla5.cells[5,a+72+(TOTAL+1)*96]);                                 //SUMA PICO
//ULTIMA PORCION DEL ULTIMO DIA /////////////////////////////////////////////////////////////////////////////////////////////////////////


///////////////////////////////////////////////////////////////////////////////////////////
//DEMANDA MAXIMA PICO RESTO VALLE
for a:=1 to TOTAL*96 do begin
    if (grilla5.cells[5,a]<>'')and(strtofloat(grilla5.cells[5,a])>MAX_pico)  then MAX_pico := strtofloat(grilla5.cells[5,a]);   //pico
    if (grilla5.cells[6,a]<>'')and(strtofloat(grilla5.cells[6,a])>MAX_resto) then MAX_resto:= strtofloat(grilla5.cells[6,a]);   //resto
    if (grilla5.cells[7,a]<>'')and(strtofloat(grilla5.cells[7,a])>MAX_valle) then MAX_valle:= strtofloat(grilla5.cells[7,a]);   //valle
                        end;
///////////////////////////////////////////////////////////////////////////////////////////



Stringgrid1.cells[1,0]:=floattostr(activa_total);
Stringgrid1.cells[1,1]:=floattostr(suma_valle);
Stringgrid1.cells[1,2]:=floattostr(MAX_valle*4);
Stringgrid1.cells[1,3]:=floattostr(suma_resto);
Stringgrid1.cells[1,4]:=floattostr(MAX_resto*4);
Stringgrid1.cells[1,5]:=floattostr(suma_pico);
Stringgrid1.cells[1,6]:=floattostr(MAX_pico*4);
Stringgrid1.cells[1,7]:=floattostr(reactiva_total);




end;

procedure Txls.Button2Click(Sender: TObject);
begin
Edit1.text:='';
stringgrid1.Visible:=false;
   label2.caption:='';
   label3.caption:='';
   label4.caption:='';
   label5.caption:='';
   label6.caption:='';
   label7.caption:='';
   label8.caption:='';
   label9.caption:='';
   label10.caption:='';
   label11.caption:='';
   label12.caption:='';
   label13.caption:='';
   label14.caption:='';
   label15.caption:='';
       grilla1.Clean;
       grilla2.Clean;
       grilla3.Clean;
       grilla4.Clean;
       grilla5.Clean;
       grilla6.Clean;
       grilla7.Clean;
       grilla8.Clean;
       grilla9.Clean;

       Stringgrid1.Cells[1,0]:='';
       Stringgrid1.Cells[1,1]:='';
       Stringgrid1.Cells[1,2]:='';
       Stringgrid1.Cells[1,3]:='';
       Stringgrid1.Cells[1,4]:='';
       Stringgrid1.Cells[1,5]:='';
       Stringgrid1.Cells[1,6]:='';
       Stringgrid1.Cells[1,7]:='';
       boton_archivo_1.visible:=true;
       boton_archivo_2.visible:=false;
        opendialog1.FileName:='';
        button1.visible:=false;
         button2.visible:=false;
         button3.visible:=false;
         button4.visible:=false;
         button5.visible:=false;
         button6.visible:=false;




            label17.caption:='';
            label18.caption:='';
            label19.caption:='';
            label20.caption:='';
            label21.caption:='';
            label22.caption:='';
            label23.caption:='';
            label24.caption:='';
            label25.caption:='';
            label26.caption:='';
            label27.caption:='';
            label28.caption:='';
            label29.caption:='';
            label30.caption:='';


end;

procedure Txls.Button3Click(Sender: TObject);
var GR: TGridRect;
begin
GR.Left:=0;
GR.Right:=4;
GR.Top:=0;
GR.Bottom:=grilla2.rowcount;
stringgrid1.Selection:=GR;

stringgrid1.CopyToClipboard(True);

end;

procedure Txls.Button4Click(Sender: TObject);
var
    nombre_de_archivo_1 , tmp1 , tmp2:         string;
    scan_fila , x , y , filas_grilla7 ,lugar:  integer;
    Archivo_1_de_excel: TsWorkbook;
    planilla_dentro_de_Archivo_1_de_excel: TsWorksheet;



begin
  opendialog1.execute;
  if opendialog1.FileName <>'' then begin

                                         nombre_de_archivo_1 :=opendialog1.FileName ;
                                         label23.caption:='Archivo: '+(ExtractFilename(opendialog1.Filename));
                                         Archivo_1_de_excel := TsWorkbook.Create;

                                         Archivo_1_de_excel.Options := Archivo_1_de_excel.Options + [boReadFormulas];

                                         Archivo_1_de_excel.ReadFromFile(nombre_de_archivo_1, sfExcel8);    //sfOOxml
                                         planilla_dentro_de_Archivo_1_de_excel := Archivo_1_de_excel.GetWorksheetByName('Sheet');

                                         scan_fila:=0;


                                          button4.visible:=false;
                                          button5.visible:=true;


                                         ////////////ENCUENTRA LA CANTIDAD MAXIMA DE FILAS/////////////////////////
                                          repeat
                                          inc(scan_fila);
                                          until (planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(scan_fila+1,0)='')and(scan_fila>30);


                                         //   if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(4,0)='' then lugar:=10 else lugar:=8;
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(7,0)='Fecha Hora' then lugar:=7 else
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(8,0)='Fecha Hora' then lugar:=8 else
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(9,0)='Fecha Hora' then lugar:=9 else
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(10,0)='Fecha Hora' then lugar:=10 else
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(11,0)='Fecha Hora' then lugar:=11 else
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(12,0)='Fecha Hora' then lugar:=12   ;



                                          //    label9.caption:=planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(8,0);
                                              scan_fila:=scan_fila-lugar;
                                         ////////////ENCUENTRA LA CANTIDAD MAXIMA DE FILAS/////////////////////////
                                        // label1.caption:='Filas: '+inttostr(scan_fila);
                                          ///////////////ACOMODA TAMAÑO DE LA GRILLA/////////////////////////////////////////////
                                          grilla6.rowcount:=scan_fila+1;
                                          grilla6.colcount:=4;
                                          ///////////////ACOMODA TAMAÑO DE LA GRILLA/////////////////////////////////////////////


                                          //=====================PASA LOS DATOS A LA grilla6 =====================

                                          for y:=0 to scan_fila do
                                          begin

                                            grilla6.Cells[0,y]:=planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(y+lugar,0);
                                            tmp1:=grilla6.Cells[0,y][1]+grilla6.Cells[0,y][2]+grilla6.Cells[0,y][3]+grilla6.Cells[0,y][4]+grilla6.Cells[0,y][5]+grilla6.Cells[0,y][6]+'20'+grilla6.Cells[0,y][7]+grilla6.Cells[0,y][8];
                                            tmp2:=grilla6.Cells[0,y][10]+grilla6.Cells[0,y][11]+grilla6.Cells[0,y][12]+grilla6.Cells[0,y][13]+grilla6.Cells[0,y][14];
                                            grilla6.Cells[0,y]:=tmp1;
                                            grilla6.Cells[1,y]:=tmp2+':00';
                                            grilla6.Cells[2,y]:=planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(y+lugar,9);
                                            grilla6.Cells[3,y]:=planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(y+lugar,11);
                                           end;



                                             grilla6.Cells[0,0]:='Fecha';
                                             grilla6.Cells[1,0]:='Hora';
                                           //=====================PASA LOS DATOS A LA grilla6 =====================

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-     AHORA EMPIEZA LA GRILLA 4       -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
x:=3;

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=- }
                                          filas_grilla7:=1;
                                          grilla7.rowcount:=scan_fila+50000;
                                          grilla7.colcount:=4;
                                          grilla7.Cells[1,0]:=timetostr(INCMINUTE(strtotime(grilla6.Cells[1,1]),-15));
                                          grilla7.Cells[0,0]:=grilla6.Cells[0,1];
                                          y:=1;
                                          repeat
                                            begin                         for x:=0 to 3 do  grilla7.Cells[x,filas_grilla7]:=grilla6.Cells[x,y];

                                                                          if strtodate(grilla7.Cells[0,filas_grilla7])=strtodate(grilla7.Cells[0,filas_grilla7-1])then BEGIN     //si los dias son =

                                                                          if ((strtotime(grilla7.Cells[1,filas_grilla7])-strtotime(grilla7.Cells[1,filas_grilla7-1])=strtotime('00:00:00')))then grilla7.Cells[1,filas_grilla7]:=timetostr((strtotime(grilla7.Cells[1,filas_grilla7-1]))) else

                                                                          if ((strtotime(grilla7.Cells[1,filas_grilla7])-strtotime(grilla7.Cells[1,filas_grilla7-1])<strtotime('00:00:00')))then begin   if (grilla7.Cells[1,filas_grilla7])<>'00:00:00' then grilla7.Cells[1,filas_grilla7]:=grilla7.Cells[1,filas_grilla7-1];
                                                                                                                                                                                                 end else

                                                                          if ((strtotime(grilla7.Cells[1,filas_grilla7])-strtotime(grilla7.Cells[1,filas_grilla7-1])>strtotime('00:15:10')))then  begin
                                                                                                                                                                                                       for x:=0 to 3 do  grilla7.Cells[x,filas_grilla7+1]:=grilla7.Cells[x,filas_grilla7];
                                                                                                                                                                                                       grilla7.Cells[1,filas_grilla7]:=timetostr(INCMINUTE(strtotime(grilla7.Cells[1,filas_grilla7-1]),15));
                                                                                                                                                                                                       for x:=2 to 3 do  grilla7.Cells[x,filas_grilla7+1]:='0';
                                                                                                                                                                                                       for x:=2 to 3 do  grilla7.Cells[x,filas_grilla7]:='0';
                                                                                                                                                                                                       dec(y);

                                                                                                                                                                                                  end;
                                                                                                                                                                         END else      //si los dias son distintos



                                                                                                                                                                         BEGIN                         if (grilla7.Cells[1,filas_grilla7])='00:00:00'then        else
                                                                                                                                                                                                      if (((strtotime(grilla7.Cells[1,filas_grilla7])-strtotime(grilla7.Cells[1,filas_grilla7-1])<>strtotime('00:15:00')))) then  begin


                                                                                                                                                                                                      for x:=0 to 3 do  grilla7.Cells[x,filas_grilla7+1]:=grilla7.Cells[x,filas_grilla7];



                                                                                                                                                                                                      repeat

                                                                                                                                                                                                    //  grilla7.Cells[0,filas_grilla7]:=grilla7.Cells[0,filas_grilla7-1];
                                                                                                                                                                                                      grilla7.Cells[1,filas_grilla7]:=timetostr(INCMINUTE(strtotime(grilla7.Cells[1,filas_grilla7-1]),15));
                                                                                                                                                                                                      grilla7.Cells[0,filas_grilla7]:=grilla7.Cells[0,filas_grilla7-1];
                                                                                                                                                                                                      for x:=2 to 3 do  grilla7.Cells[x,filas_grilla7]:='0';
                                                                                                                                                                                                      inc(filas_grilla7);
                                                                                                                                                                                                      until grilla7.Cells[1,filas_grilla7-1]='23:45:00';


                                                                                                                                                                                                   //   grilla7.Cells[0,filas_grilla7]:=grilla7.Cells[0,filas_grilla7-1];
                                                                                                                                                                                                      grilla7.Cells[1,filas_grilla7]:='00:00:00';
                                                                                                                                                                                                      for x:=2 to 3 do  grilla7.Cells[x,filas_grilla7]:='0';
                                                                                                                                                                                                      grilla7.Cells[0,filas_grilla7]:=datetostr(INCday(strtodate(grilla7.Cells[0,filas_grilla7-1]),1));     //repetido
                                                                                                                                                                                                      inc(filas_grilla7);

                                                                                                                                                                                                      for x:=0 to 3 do  grilla7.Cells[x,filas_grilla7]:=grilla7.Cells[x,filas_grilla7];


                                                                                                                                                                                                      grilla7.Cells[0,filas_grilla7]:=datetostr(INCday(strtodate(grilla7.Cells[0,filas_grilla7-1]),0));
                                                                                                                                                                                                      grilla7.Cells[1,filas_grilla7]:='00:15:00';
                                                                                                                                                                                                      for x:=2 to 3 do  grilla7.Cells[x,filas_grilla7]:='0';
                                                                                                                                                                                                      dec(y);




                                                                                                                                                                                                                                                                                                                                end;

                                                                                                                                                                        END;
                                                                          inc(filas_grilla7);
                                                                          inc(y);
                                            end;
                                       until y=scan_fila+1;



     grilla7.rowcount:=filas_grilla7;




  grilla7.Cells[0,0]:='Fecha';
  grilla7.Cells[1,0]:='Hora';
  grilla7.Cells[2,0]:='Wh';
  grilla7.Cells[3,0]:='VARh';

  label17.caption:='TERMINAL: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(2,1);
  label18.caption:='DOMICILIO: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(3,1);
  label19.caption:='SERVICIO: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(4,1);
  label20.caption:='TITULAR: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(2,10);

  if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(3,10)<>'' then
  label21.caption:='PERIODO: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(3,10)
  else label21.caption:='PERIODO: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(4,10);

  label22.caption:='Antes: '+inttostr(scan_fila)+', después: '+inttostr(filas_grilla7-1)+' ('+floattostr((filas_grilla7-1)/96)+' días)';

  if (int((filas_grilla7-1)/96)=((filas_grilla7-1)/96))then  begin
              if edit1.text<>floattostr((filas_grilla7-1)/96) then showmessage('Atención: Parece que los dos archivos tienen cantidad de días distintos, fijate que correspondan al mismo mes.');
                                                             end
  else  showmessage('Atención, faltan datos en el archivo XLS, pues al dividir la cantidad de datos totales por 96, no dá un entero. Completalos manualmente editando el XLS');


     end;


end;

procedure Txls.Button5Click(Sender: TObject);

var
    nombre_de_archivo_1 , tmp1 , tmp2:         string;
    scan_fila , x , y , filas_grilla9 ,lugar:  integer;
    Archivo_1_de_excel: TsWorkbook;
    planilla_dentro_de_Archivo_1_de_excel: TsWorksheet;



begin
  opendialog1.execute;
  if opendialog1.FileName <>'' then begin

                                         nombre_de_archivo_1 :=opendialog1.FileName ;
                                         label30.caption:='Archivo: '+(ExtractFilename(opendialog1.Filename));
                                         Archivo_1_de_excel := TsWorkbook.Create;

                                         Archivo_1_de_excel.Options := Archivo_1_de_excel.Options + [boReadFormulas];

                                         Archivo_1_de_excel.ReadFromFile(nombre_de_archivo_1, sfExcel8);    //sfOOxml
                                         planilla_dentro_de_Archivo_1_de_excel := Archivo_1_de_excel.GetWorksheetByName('Sheet');

                                         scan_fila:=0;


                                          button5.visible:=false;
                                          button1.visible:=true;
                                          button6.visible:=true;

                                         ////////////ENCUENTRA LA CANTIDAD MAXIMA DE FILAS/////////////////////////
                                          repeat
                                          inc(scan_fila);
                                          until (planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(scan_fila+1,0)='')and(scan_fila>30);


                                         //   if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(4,0)='' then lugar:=10 else lugar:=8;
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(7,0)='Fecha Hora' then lugar:=7 else
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(8,0)='Fecha Hora' then lugar:=8 else
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(9,0)='Fecha Hora' then lugar:=9 else
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(10,0)='Fecha Hora' then lugar:=10 else
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(11,0)='Fecha Hora' then lugar:=11 else
                                              if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(12,0)='Fecha Hora' then lugar:=12   ;



                                          //    label9.caption:=planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(8,0);
                                              scan_fila:=scan_fila-lugar;
                                         ////////////ENCUENTRA LA CANTIDAD MAXIMA DE FILAS/////////////////////////
                                        // label1.caption:='Filas: '+inttostr(scan_fila);
                                          ///////////////ACOMODA TAMAÑO DE LA GRILLA/////////////////////////////////////////////
                                          grilla8.rowcount:=scan_fila+1;
                                          grilla8.colcount:=4;
                                          ///////////////ACOMODA TAMAÑO DE LA GRILLA/////////////////////////////////////////////


                                          //=====================PASA LOS DATOS A LA grilla8 =====================

                                          for y:=0 to scan_fila do
                                          begin

                                            grilla8.Cells[0,y]:=planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(y+lugar,0);
                                            tmp1:=grilla8.Cells[0,y][1]+grilla8.Cells[0,y][2]+grilla8.Cells[0,y][3]+grilla8.Cells[0,y][4]+grilla8.Cells[0,y][5]+grilla8.Cells[0,y][6]+'20'+grilla8.Cells[0,y][7]+grilla8.Cells[0,y][8];
                                            tmp2:=grilla8.Cells[0,y][10]+grilla8.Cells[0,y][11]+grilla8.Cells[0,y][12]+grilla8.Cells[0,y][13]+grilla8.Cells[0,y][14];
                                            grilla8.Cells[0,y]:=tmp1;
                                            grilla8.Cells[1,y]:=tmp2+':00';
                                            grilla8.Cells[2,y]:=planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(y+lugar,9);
                                            grilla8.Cells[3,y]:=planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(y+lugar,11);
                                           end;



                                             grilla8.Cells[0,0]:='Fecha';
                                             grilla8.Cells[1,0]:='Hora';
                                           //=====================PASA LOS DATOS A LA grilla8 =====================

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-     AHORA EMPIEZA LA GRILLA 4       -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=}
x:=3;

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=- }
                                          filas_grilla9:=1;
                                          grilla9.rowcount:=scan_fila+50000;
                                          grilla9.colcount:=4;
                                          grilla9.Cells[1,0]:=timetostr(INCMINUTE(strtotime(grilla8.Cells[1,1]),-15));
                                          grilla9.Cells[0,0]:=grilla8.Cells[0,1];
                                          y:=1;
                                          repeat
                                            begin                         for x:=0 to 3 do  grilla9.Cells[x,filas_grilla9]:=grilla8.Cells[x,y];

                                                                          if strtodate(grilla9.Cells[0,filas_grilla9])=strtodate(grilla9.Cells[0,filas_grilla9-1])then BEGIN     //si los dias son =

                                                                          if ((strtotime(grilla9.Cells[1,filas_grilla9])-strtotime(grilla9.Cells[1,filas_grilla9-1])=strtotime('00:00:00')))then grilla9.Cells[1,filas_grilla9]:=timetostr((strtotime(grilla9.Cells[1,filas_grilla9-1]))) else

                                                                          if ((strtotime(grilla9.Cells[1,filas_grilla9])-strtotime(grilla9.Cells[1,filas_grilla9-1])<strtotime('00:00:00')))then begin   if (grilla9.Cells[1,filas_grilla9])<>'00:00:00' then grilla9.Cells[1,filas_grilla9]:=grilla9.Cells[1,filas_grilla9-1];
                                                                                                                                                                                                 end else

                                                                          if ((strtotime(grilla9.Cells[1,filas_grilla9])-strtotime(grilla9.Cells[1,filas_grilla9-1])>strtotime('00:15:10')))then  begin
                                                                                                                                                                                                       for x:=0 to 3 do  grilla9.Cells[x,filas_grilla9+1]:=grilla9.Cells[x,filas_grilla9];
                                                                                                                                                                                                       grilla9.Cells[1,filas_grilla9]:=timetostr(INCMINUTE(strtotime(grilla9.Cells[1,filas_grilla9-1]),15));
                                                                                                                                                                                                       for x:=2 to 3 do  grilla9.Cells[x,filas_grilla9+1]:='0';
                                                                                                                                                                                                       for x:=2 to 3 do  grilla9.Cells[x,filas_grilla9]:='0';
                                                                                                                                                                                                       dec(y);

                                                                                                                                                                                                  end;
                                                                                                                                                                         END else      //si los dias son distintos



                                                                                                                                                                         BEGIN                         if (grilla9.Cells[1,filas_grilla9])='00:00:00'then        else
                                                                                                                                                                                                      if (((strtotime(grilla9.Cells[1,filas_grilla9])-strtotime(grilla9.Cells[1,filas_grilla9-1])<>strtotime('00:15:00')))) then  begin


                                                                                                                                                                                                      for x:=0 to 3 do  grilla9.Cells[x,filas_grilla9+1]:=grilla9.Cells[x,filas_grilla9];



                                                                                                                                                                                                      repeat

                                                                                                                                                                                                    //  grilla9.Cells[0,filas_grilla9]:=grilla9.Cells[0,filas_grilla9-1];
                                                                                                                                                                                                      grilla9.Cells[1,filas_grilla9]:=timetostr(INCMINUTE(strtotime(grilla9.Cells[1,filas_grilla9-1]),15));
                                                                                                                                                                                                      grilla9.Cells[0,filas_grilla9]:=grilla9.Cells[0,filas_grilla9-1];
                                                                                                                                                                                                      for x:=2 to 3 do  grilla9.Cells[x,filas_grilla9]:='0';
                                                                                                                                                                                                      inc(filas_grilla9);
                                                                                                                                                                                                      until grilla9.Cells[1,filas_grilla9-1]='23:45:00';


                                                                                                                                                                                                   //   grilla9.Cells[0,filas_grilla9]:=grilla9.Cells[0,filas_grilla9-1];
                                                                                                                                                                                                      grilla9.Cells[1,filas_grilla9]:='00:00:00';
                                                                                                                                                                                                      for x:=2 to 3 do  grilla9.Cells[x,filas_grilla9]:='0';
                                                                                                                                                                                                      grilla9.Cells[0,filas_grilla9]:=datetostr(INCday(strtodate(grilla9.Cells[0,filas_grilla9-1]),1));     //repetido
                                                                                                                                                                                                      inc(filas_grilla9);

                                                                                                                                                                                                      for x:=0 to 3 do  grilla9.Cells[x,filas_grilla9]:=grilla9.Cells[x,filas_grilla9];


                                                                                                                                                                                                      grilla9.Cells[0,filas_grilla9]:=datetostr(INCday(strtodate(grilla9.Cells[0,filas_grilla9-1]),0));
                                                                                                                                                                                                      grilla9.Cells[1,filas_grilla9]:='00:15:00';
                                                                                                                                                                                                      for x:=2 to 3 do  grilla9.Cells[x,filas_grilla9]:='0';
                                                                                                                                                                                                      dec(y);




                                                                                                                                                                                                                                                                                                                                end;

                                                                                                                                                                        END;
                                                                          inc(filas_grilla9);
                                                                          inc(y);
                                            end;
                                       until y=scan_fila+1;



     grilla9.rowcount:=filas_grilla9;




  grilla9.Cells[0,0]:='Fecha';
  grilla9.Cells[1,0]:='Hora';
  grilla9.Cells[2,0]:='Wh';
  grilla9.Cells[3,0]:='VARh';

  label24.caption:='TERMINAL: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(2,1);
  label25.caption:='DOMICILIO: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(3,1);
  label26.caption:='SERVICIO: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(4,1);
  label27.caption:='TITULAR: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(2,10);

  if planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(3,10)<>'' then
  label28.caption:='PERIODO: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(3,10)
  else label28.caption:='PERIODO: '+planilla_dentro_de_Archivo_1_de_excel.ReadAsUTF8Text(4,10);

  label29.caption:='Antes: '+inttostr(scan_fila)+', después: '+inttostr(filas_grilla9-1)+' ('+floattostr((filas_grilla9-1)/96)+' días)';

  if (int((filas_grilla9-1)/96)=((filas_grilla9-1)/96))then  begin
              if edit1.text<>floattostr((filas_grilla9-1)/96) then showmessage('Atención: Parece que los dos archivos tienen cantidad de días distintos, fijate que correspondan al mismo mes.');
                                                             end
  else  showmessage('Atención, faltan datos en el archivo XLS, pues al dividir la cantidad de datos totales por 96, no dá un entero. Completalos manualmente editando el XLS');


     end;

end;

procedure Txls.Button6Click(Sender: TObject);
  var y:integer;
begin
for y:=1 to 48 do begin
                            grilla2.Cells[2,y]:='0';
                            grilla2.Cells[3,y]:='0';
                            grilla4.Cells[2,y]:='0';
                            grilla4.Cells[3,y]:='0';
                            grilla7.Cells[2,y]:='0';
                            grilla7.Cells[3,y]:='0';
                            grilla9.Cells[2,y]:='0';
                            grilla9.Cells[3,y]:='0';
                  end;


 for y:=grilla2.RowCount-1 downto grilla2.RowCount-48 do begin
                                                               grilla2.Cells[2,y]:='0';
                                                               grilla2.Cells[3,y]:='0';
                                                               grilla4.Cells[2,y]:='0';
                                                               grilla4.Cells[3,y]:='0';
                                                               grilla7.Cells[2,y]:='0';
                                                               grilla7.Cells[3,y]:='0';
                                                               grilla9.Cells[2,y]:='0';
                                                               grilla9.Cells[3,y]:='0';
                                                     end;
end;

procedure Txls.Edit1Change(Sender: TObject);
begin

end;

procedure Txls.FormCreate(Sender: TObject);
begin
 //    label1.caption:='';
   label2.caption:='';
   label3.caption:='';
   label4.caption:='';
   label5.caption:='';
   label6.caption:='';
   label7.caption:='';
   label8.caption:='';
   label9.caption:='';
   label10.caption:='';
   label11.caption:='';
   label12.caption:='';
   label13.caption:='';
   label14.caption:='';
   label15.caption:='';

   label17.caption:='';
   label18.caption:='';
   label19.caption:='';
   label20.caption:='';
   label21.caption:='';
   label22.caption:='';
   label23.caption:='';
   label24.caption:='';
   label25.caption:='';
   label26.caption:='';
   label27.caption:='';
   label28.caption:='';
   label29.caption:='';
   label30.caption:='';


   button1.visible:=false;
     button2.visible:=false;
       button3.visible:=false;
       button4.visible:=false;
       button5.visible:=false;
       button6.visible:=false;
 boton_archivo_2.visible:=false;
 stringgrid1.Visible:=false;


end;

procedure Txls.Label10Click(Sender: TObject);
begin

end;

procedure Txls.Label1Click(Sender: TObject);
begin

end;

procedure Txls.Label9Click(Sender: TObject);
begin

end;





end.

