program ECOLOG;

{$mode objfpc}{$H+}

uses {$IFDEF UNIX} {$IFDEF UseCThreads}
  cthreads, {$ENDIF} {$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms,
  printer4lazarus,
  lazdbexport,
  multiframepackage,
  zmsql,
  FrameViewer09,
  main,
  info,
  filter,
  about,
  cluster,
  nmds,
  cca,
  ca,
  export,
  pco,
  pca,
  rda,
  child,
  splash,
  report,
  useful,
  diversity,
 {$IFDEF WINDOWS}ShpAPI129,
 {$ENDIF} progress,
  eval;

{$R *.res}

var
  SplashScreen: TSplashForm;

begin
  RequireDerivedFormResource := True;
  Application.Scaled := True;
  Application.Initialize;
  SplashScreen := TSplashForm.Create(Application);
  SplashScreen.Show;
  SplashScreen.Update;
  while not SplashScreen.Completed do
    Application.ProcessMessages;
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(TProjectDlg, ProjectDlg);
  Application.CreateForm(TFilterDlg, FilterDlg);
  Application.CreateForm(TAboutBox, AboutBox);
  Application.CreateForm(TClusterDlg, ClusterDlg);
  Application.CreateForm(TNMDSDlg, NMDSDlg);
  Application.CreateForm(TCCADlg, CCADlg);
  Application.CreateForm(TCADlg, CADlg);
  Application.CreateForm(TPCODlg, PCODlg);
  Application.CreateForm(TPCADlg, PCADlg);
  Application.CreateForm(TRDADlg, RDADlg);
  SplashScreen.Hide;
  SplashScreen.Free;
  Application.CreateForm(TProgressDlg, ProgressDlg);
  Application.Run;
end.
