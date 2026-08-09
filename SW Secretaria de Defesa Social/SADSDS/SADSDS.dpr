program SADSDS;
//library SADSDS;


uses
  //ISAPIApp,
  //IWInitISAPI,
  Forms,
  IWMain,
  Unit1 in 'Unit1.pas' {IWForm1: TIWAppForm},
  ServerController in 'ServerController.pas' {IWServerController: TIWServerControllerBase},
  UserSessionUnit in 'UserSessionUnit.pas' {IWUserSession: TIWUserSessionBase},
  uIntercritério in 'uIntercritério.pas' {frmaintercriterio: TIWAppForm},
  uResultados in 'uResultados.pas' {frmResultados: TIWAppForm},
  lpsolve51 in 'lpsolve51.pas',
  uHasseDiagram in 'uHasseDiagram.pas' {frmhasseD: TIWAppForm},
  uCode in 'uCode.pas',
  uAPI_SolveLP in 'uAPI_SolveLP.pas',
  untIntro in 'untIntro.pas' {FrmEntrada: TIWAppForm};

{$R *.res}

 //inicio ISAPI
//exports

  //GetExtensionVersion, HttpExtensionProc, TerminateExtension;
 //fim ISAPI
begin
  //IWRun;
  //ISAPI
  Application.Initialize;
  Application.CreateForm(TformIWMain, formIWMain);
  Application.Run;
end.
