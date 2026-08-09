unit untIntro;

interface

uses
  Classes, SysUtils, IWAppForm, IWApplication, IWColor, IWTypes, IWCompLabel,
  IWCompText, pngimage, IWVCLBaseControl, IWBaseControl, IWBaseHTMLControl,
  IWControl, IWExtCtrls, Controls, Forms, IWVCLBaseContainer, IWContainer,
  IWHTMLContainer, IWHTML40Container, IWRegion, jpeg, IWCompButton,
  IWCompListbox, IWCompEdit, IWGrids, ActnList, Comobj, IWWebGrid, IWAdvWebGrid,
  IWAdvWebGridExcel, TeeGDIPlus, TeEngine, ExtCtrls, TeeProcs, Chart, Series,
  IWDBStdCtrls, WebComp, WebAdapt, Math, IdMessage, IdTCPConnection,
  IdTCPClient, Graphics,
  IdExplicitTLSClientServerBase, IdMessageClient, IdSMTPBase, IdSMTP,
  IdBaseComponent, IdComponent, IdIOHandler, IdIOHandlerSocket,
  IdIOHandlerStack, IdSSL, IdSSLOpenSSL, IdText, IWProgressIndicator,
  IWCompProgressBar, IWCompCheckbox, IWBaseComponent, IWBaseHTMLComponent,
  IWBaseHTML40Component, IWTMSPopup, IWHTMLControls, Menus, IWHelpTip,
  IWCompRadioButton, DateUtils
  { , IdContext, IdBaseComponent, IdComponent, IdCustomTCPServer,
    IdTCPServer, IdCmdTCPServer, IdDICTServer } ;

type
  TFrmEntrada = class(TIWAppForm)
    TIWAdvWebGridExcelIO1: TTIWAdvWebGridExcelIO;
    TIWAdvWebGridExcelIO2: TTIWAdvWebGridExcelIO;
    tmr1: TIWTimer;
    tmrcontinue: TIWTimer;
    tmrcont2: TIWTimer;
    tmrstep: TIWTimer;
    rgnlanguage: TIWRegion;
    IWImage3: TIWImage;
    IWLabel58: TIWLabel;
    btnportuguese: TIWButton;
    btnEnglish: TIWButton;
    btnclosela: TIWButton;
    IWLabel59: TIWLabel;
    IWRegion13: TIWRegion;
    pmnHelp: TPopupMenu;
    About1: TMenuItem;
    UserGuide1: TMenuItem;
    English1: TMenuItem;
    Portuguese1: TMenuItem;
    pmnResetoption: TPopupMenu;
    Restartthecurrentproblem1: TMenuItem;
    Logout1: TMenuItem;
    AlreadyRakend1: TMenuItem;
    rgnintro: TIWRegion;
    lblWarning: TIWLabel;
    btnLogin: TIWButton;
    btnRegister: TIWButton;
    imgSolve: TIWImage;
    lblSolve: TIWLabel;
    txtwarnning: TIWText;
    IWButton3: TIWButton;
    IWLabel46: TIWLabel;
    IWImage9: TIWImage;
    IWLabel52: TIWLabel;
    IWLink3: TIWLink;
    IWLabel70: TIWLabel;
    IWImage23: TIWImage;
    IWImage10: TIWImage;
    IWImage27: TIWImage;
    TIWPopupMenuButton12: TTIWPopupMenuButton;
    Chart1: TChart;
    Series1: TBarSeries;
    IWRegion9: TIWRegion;
    IWRegion16: TIWRegion;
    rgnCadusuario: TIWRegion;
    lblCadusuario: TIWLabel;
    lblOptional: TIWLabel;
    lblCademail: TIWLabel;
    lblCadtitle: TIWLabel;
    lblCadfirstname: TIWLabel;
    lblCadLastname: TIWLabel;
    edtcademail: TIWEdit;
    edtcadtitle: TIWEdit;
    edtcadname: TIWEdit;
    edtcadlast: TIWEdit;
    lblCadcity: TIWLabel;
    lblCadcountry: TIWLabel;
    lblCadinsticity: TIWLabel;
    lblCadinstiname: TIWLabel;
    lblCadaffiliation: TIWLabel;
    edtpassword: TIWEdit;
    edtcadaddcity: TIWEdit;
    edtcadaddresscoun: TIWEdit;
    edtcadaddresscity: TIWEdit;
    edtcadinstname: TIWEdit;
    lblCadwhats: TIWLabel;
    edtcadwhats: TIWEdit;
    btnSend: TIWButton;
    lblPassword: TIWLabel;
    btnVoltarinicio: TIWButton;
    IWComboBox1: TIWComboBox;
    IWLabel1: TIWLabel;
    IWRadioGroup1: TIWRadioGroup;
    edtdia: TIWEdit;
    edtmes: TIWEdit;
    IWLabel5: TIWLabel;
    IWLabel4: TIWLabel;
    IWLabel6: TIWLabel;
    edtano: TIWEdit;
    IWLabel2: TIWLabel;
    IWRadioGroup2: TIWRadioGroup;
    IWLabel3: TIWLabel;
    edtcruso: TIWEdit;
    IWLabel22: TIWLabel;
    IWLabel23: TIWLabel;
    cmbbxAffiliation: TIWComboBox;
    edtconfirmpass: TIWEdit;
    lblconfirmpassword: TIWLabel;
    lblpasswarning: TIWLabel;
    chckbxanalyst: TIWCheckBox;
    IWLabel27: TIWLabel;
    IWLabel28: TIWLabel;
    IWLabel29: TIWLabel;
    IWLabel30: TIWLabel;
    IWLabel31: TIWLabel;
    IWLabel32: TIWLabel;
    IWLabel33: TIWLabel;
    lblAffiliation1: TIWLabel;
    lblAffiliation2: TIWLabel;
    IWLabel36: TIWLabel;
    txtconfirmation: TIWText;
    IWImage11: TIWImage;
    IWLabel37: TIWLabel;
    IWImage6: TIWImage;
    IWLink1: TIWLink;
    IWLabel71: TIWLabel;
    IWImage28: TIWImage;
    IWImage29: TIWImage;
    TIWPopupMenuButton8: TTIWPopupMenuButton;
    IWRegion15: TIWRegion;
    rgnConseqvalues: TIWRegion;
    lblEnterconsequence: TIWLabel;
    btnSaveproblem: TIWButton;
    btnSaveStart: TIWButton;
    btnreset2: TIWButton;
    IWLabel25: TIWLabel;
    btnlogoutrgncons: TIWButton;
    gn: TIWButton;
    IWLabel34: TIWLabel;
    btnAboutDC: TIWButton;
    IWLabel51: TIWLabel;
    IWImage13: TIWImage;
    IWLabel40: TIWLabel;
    IWImage8: TIWImage;
    IWLink2: TIWLink;
    IWLabel35: TIWLabel;
    TIWPopupMenuButton4: TTIWPopupMenuButton;
    IWImage7: TIWImage;
    IWImage30: TIWImage;
    TIWPopupMenuButton5: TTIWPopupMenuButton;
    rgnaboutdiscrete: TIWRegion;
    lbltitleAD: TIWLabel;
    txtimpAD: TIWText;
    IWImage2: TIWImage;
    btncloseIDC: TIWButton;
    rgnconsman: TIWRegion;
    rgntabela: TIWRegion;
    IWLabel48: TIWLabel;
    grdCritInfo: TIWGrid;
    rgnConfirmBackEditar: TIWRegion;
    IWLabel67: TIWLabel;
    btnConfBack: TIWButton;
    btnConfNao: TIWButton;
    IWText1: TIWText;
    IWImage4: TIWImage;
    IWRegion5: TIWRegion;
    IWRegion17: TIWRegion;
    rgntabelaalter: TIWRegion;
    grdinformaalt: TIWGrid;
    IWLabel69: TIWLabel;
    rgnImportExport: TIWRegion;
    grdImportproblema: TTIWAdvWebGrid;
    btnExporttemplate: TIWButton;
    lblPleaseselectfile: TIWLabel;
    btnReset4: TIWButton;
    flImportarproblema: TIWFile;
    lblnameproblema: TIWLabel;
    edtEntername: TIWEdit;
    lblAdvertencia: TIWLabel;
    grdImportacao: TTIWAdvWebGrid;
    lblproblematica: TIWLabel;
    cmbbxproblem: TIWComboBox;
    btnslotherfile: TIWButton;
    btnLogoutRgnImportExport: TIWButton;
    IWLabel13: TIWLabel;
    IWLabel14: TIWLabel;
    IWLabel15: TIWLabel;
    IWLabel16: TIWLabel;
    btnmessageDisc: TIWButton;
    lblsheet: TIWLabel;
    btnguide: TIWButton;
    IWLabel49: TIWLabel;
    btnStartelicitation: TIWButton;
    btnauxstartelicitation: TIWButton;
    IWImage14: TIWImage;
    IWLink4: TIWLink;
    IWLabel11: TIWLabel;
    TIWPopupMenuButton3: TTIWPopupMenuButton;
    IWImage31: TIWImage;
    IWImage32: TIWImage;
    btnImporttabela: TIWButton;
    TIWPopupMenuButton6: TTIWPopupMenuButton;
    rgnWarning: TIWRegion;
    txtwarnings: TIWText;
    IWLabel12: TIWLabel;
    IWLabel19: TIWLabel;
    IWLabel41: TIWLabel;
    lblimpdisc: TIWLabel;
    IWLabel55: TIWLabel;
    IWRegion1: TIWRegion;
    IWRegion18: TIWRegion;
    rgnOpcaoInput: TIWRegion;
    lblpleasechoose: TIWLabel;
    btnImport: TIWButton;
    btnRegisternew: TIWButton;
    btnContinueregist: TIWButton;
    btnLogoutRgnOpcaoInput: TIWButton;
    IWImage37: TIWImage;
    IWImage38: TIWImage;
    IWLabel44: TIWLabel;
    IWImage39: TIWImage;
    IWLink9: TIWLink;
    IWRegion7: TIWRegion;
    IWRegion22: TIWRegion;
    IWButton7: TIWButton;
    lblproblematic1: TIWLabel;
    IWLabel72: TIWLabel;
    lblchosenproblematic: TIWLabel;
    rgnEnterLogin: TIWRegion;
    lbluserid: TIWLabel;
    edtemail: TIWEdit;
    btnEnter: TIWButton;
    lblSenha: TIWLabel;
    edtSenha: TIWEdit;
    IWLabel8: TIWLabel;
    IWLabel9: TIWLabel;
    IWButton1: TIWButton;
    btnrecover: TIWButton;
    txtpassrecover: TIWText;
    txtnewcode: TIWText;
    IWImage16: TIWImage;
    IWLink5: TIWLink;
    IWLabel10: TIWLabel;
    IWImage15: TIWImage;
    IWImage33: TIWImage;
    TIWPopupMenuButton7: TTIWPopupMenuButton;
    IWRegion2: TIWRegion;
    rgnvalidationpassword: TIWRegion;
    IWImage1: TIWImage;
    IWLabel18: TIWLabel;
    rgnvalidation: TIWRegion;
    IWLabel42: TIWLabel;
    edtvalidation: TIWEdit;
    btnvalidaremail: TIWButton;
    lblnewvalidationcode: TIWLabel;
    btnnewcode: TIWButton;
    IWRegion3: TIWRegion;
    rgnnewpassword: TIWRegion;
    IWLabel20: TIWLabel;
    edtnewpassword: TIWEdit;
    edtnewpassconfir: TIWButton;
    IWLabel21: TIWLabel;
    edtnewpasswordconfirm: TIWEdit;
    lbldifcamp: TIWLabel;
    IWRegion19: TIWRegion;
    rgnCofirmReset: TIWRegion;
    IWImage5: TIWImage;
    IWLabel68: TIWLabel;
    IWText2: TIWText;
    IWButton8: TIWButton;
    btnConfReset: TIWButton;
    lbllembretebudget: TIWLabel;
    rgnInputMatrizConseq: TIWRegion;
    btnLogoutRgnInputMatrizConseq: TIWButton;
    btnResetRgnInputMatrizCon: TIWButton;
    IWLabel17: TIWLabel;
    btnauxstep1: TIWButton;
    btnStep1: TIWButton;
    btnauxstep2: TIWButton;
    btnStep2: TIWButton;
    IWImage20: TIWImage;
    IWLabel73: TIWLabel;
    btnmenu: TTIWPopupMenuButton;
    IWImage18: TIWImage;
    IWImage36: TIWImage;
    TIWPopupMenuButton9: TTIWPopupMenuButton;
    rgnInputConsequences: TIWRegion;
    grdMatrizConseq: TIWGrid;
    lblInputValues: TIWLabel;
    rgnEquivalencedistance: TIWRegion;
    lblEquivalencedistance: TIWLabel;
    edtMaxdif: TIWEdit;
    IWLabel74: TIWLabel;
    rgnhide1: TIWRegion;
    IWRegion20: TIWRegion;
    rgnCadprojeto: TIWRegion;
    lblRegisternew: TIWLabel;
    edtProjectname: TIWEdit;
    lblProjectname: TIWLabel;
    btnSavecadastro: TIWButton;
    btnReset: TIWButton;
    IWLabel7: TIWLabel;
    cmproblematica1: TIWComboBox;
    btnLogoutRgnCadProjeto: TIWButton;
    IWLabel24: TIWLabel;
    btnDiscrete0: TIWButton;
    IWLabel38: TIWLabel;
    btnplusalt: TIWButton;
    btnminalt: TIWButton;
    btnpluscrit: TIWButton;
    btnmincrit: TIWButton;
    txthelpdisc: TIWText;
    IWButton2: TIWButton;
    IWButton4: TIWButton;
    IWImage19: TIWImage;
    IWLabel39: TIWLabel;
    IWImage22: TIWImage;
    IWLink7: TIWLink;
    IWLabel53: TIWLabel;
    TIWPopupMenuButton1: TTIWPopupMenuButton;
    IWImage21: TIWImage;
    IWImage35: TIWImage;
    TIWPopupMenuButton10: TTIWPopupMenuButton;
    IWRegion4: TIWRegion;
    IWRegion10: TIWRegion;
    IWRegion11: TIWRegion;
    IWLabel56: TIWLabel;
    lblNamealternat: TIWLabel;
    lblNumberalternat: TIWLabel;
    edtNamealternat: TIWEdit;
    edtNumberalternat: TIWEdit;
    btnAddalternative: TIWButton;
    IWLabel54: TIWLabel;
    lstbxNamealternat: TIWListbox;
    lblselecttoremove: TIWLabel;
    btneditalternative: TIWButton;
    btndeletealternat: TIWButton;
    TIWHelpTip1: TTIWHelpTip;
    IWRegion21: TIWRegion;
    lblproblematic2: TIWLabel;
    rgnLoginproject: TIWRegion;
    lblChooseproj: TIWLabel;
    lstbxProjects: TIWListbox;
    btnChoose: TIWButton;
    btnReset3: TIWButton;
    btnLogoutRgnLogin: TIWButton;
    IWLabel26: TIWLabel;
    lblMexplanation: TIWLabel;
    IWLabel43: TIWLabel;
    btnauxchoose: TIWButton;
    IWImage24: TIWImage;
    IWLabel50: TIWLabel;
    IWLabel76: TIWLabel;
    IWImage26: TIWImage;
    IWLink8: TIWLink;
    TIWPopupMenuButton2: TTIWPopupMenuButton;
    IWImage25: TIWImage;
    IWImage34: TIWImage;
    TIWPopupMenuButton11: TTIWPopupMenuButton;
    IWRegion6: TIWRegion;
    IWRegion8: TIWRegion;
    IWRegion23: TIWRegion;
    rgnproblematica: TIWRegion;
    IWLabel63: TIWLabel;
    IWButton5: TIWButton;
    IWLabel61: TIWLabel;
    IWLabel64: TIWLabel;
    IWLabel65: TIWLabel;
    btnproblematicalogout: TIWButton;
    IWLabel66: TIWLabel;
    IWRegion14: TIWRegion;
    rgnchoice: TIWRegion;
    imagechoice: TIWImage;
    rgnsorting: TIWRegion;
    imagesorting: TIWImage;
    rgnport: TIWRegion;
    rgnrank: TIWRegion;
    imagerank: TIWImage;
    IWImage12: TIWImage;
    IWImage40: TIWImage;
    IWImage41: TIWImage;
    IWLink10: TIWLink;
    IWRegion24: TIWRegion;
    IWLabel62: TIWLabel;
    rgnAbout: TIWRegion;
    IWImage42: TIWImage;
    IWImage58: TIWImage;
    IWImage59: TIWImage;
    IWLink25: TIWLink;
    IWLabel116: TIWLabel;
    IWLabel128: TIWLabel;
    lnkcloseabout: TIWLink;
    IWRegion58: TIWRegion;
    IWRegion59: TIWRegion;
    IWText31: TIWText;
    lblbudget: TIWLabel;
    edtbudget: TIWEdit;
    hptpbud: TTIWHelpTip;
    rgnPortGeneration: TIWRegion;
    IWLabel86: TIWLabel;
    IWLabel77: TIWLabel;
    IWLabel99: TIWLabel;
    IWLink19: TIWLink;
    IWImage50: TIWImage;
    IWImage55: TIWImage;
    IWImage43: TIWImage;
    IWRegion26: TIWRegion;
    txtsesanw: TIWText;
    IWRegion47: TIWRegion;
    PgbarPort: TIWProgressBar;
    rgnportComb: TIWRegion;
    imgPortComb: TIWImage;
    IWLabel78: TIWLabel;
    tmrPortfoliocombination: TIWTimer;
    IWRegion12: TIWRegion;
    lblCinfCA: TIWLabel;
    lblNameattribut: TIWLabel;
    lblNumberattribut: TIWLabel;
    edtNumberattribut: TIWEdit;
    edtNameattribut: TIWEdit;
    IWLabel57: TIWLabel;
    lblCaracteristica: TIWLabel;
    lblDireçãoproblema: TIWLabel;
    cmbxScaletype: TIWComboBox;
    cmbxCriteriondirection: TIWComboBox;
    lblaboutDC0: TIWLabel;
    lblNumberlevels: TIWLabel;
    lblselecttoremoveattribut: TIWLabel;
    edtNumberlevels: TIWEdit;
    btnAddattribut: TIWButton;
    lstbxNamecriteria: TIWListbox;
    btneditcriteria: TIWButton;
    btndeletecriteria: TIWButton;
    hptpcriCA: TTIWHelpTip;
    chckbxinteger: TIWCheckBox;
    IWRegion25: TIWRegion;
    lblacaointra: TIWLabel;
    rdbtndeclaretype: TIWRadioButton;
    rdbtnelictfunction: TIWRadioButton;
    lblTipoproblema: TIWLabel;
    cmbxFO: TIWComboBox;
    lblfunc: TIWLabel;
    txtparametros: TIWText;
    lblparA: TIWLabel;
    edtparA: TIWEdit;
    lblparB: TIWLabel;
    edtparB: TIWEdit;
    hptpintraexp: TTIWHelpTip;
    imageport: TIWImage;
    IWLink11: TIWLink;
    IWLink12: TIWLink;
    IWLink13: TIWLink;
    IWLink6: TIWLink;
    btnovoproblema: TIWButton;
    procedure btnLoginClick(Sender: TObject);
    procedure btnRegisterClick(Sender: TObject);
    procedure btnEnterClick(Sender: TObject);
    procedure btnRegisternewClick(Sender: TObject);
    procedure btnContinueregistClick(Sender: TObject);
    procedure btnSendClick(Sender: TObject);
    procedure btnSavecadastroClick(Sender: TObject);
    procedure btnreset2Click(Sender: TObject);
    procedure btnResetClick(Sender: TObject);
    procedure btnImportClick(Sender: TObject);
    procedure btndeletealternatClick(Sender: TObject);
    procedure btndeletecriteriaClick(Sender: TObject);
    procedure btnReset3Click(Sender: TObject);
    procedure btnReset4Click(Sender: TObject);
    procedure btnImporttabelaClick(Sender: TObject);
    procedure btnAddalternativeClick(Sender: TObject);
    procedure btnAddattributClick(Sender: TObject);
    procedure btnExporttemplateClick(Sender: TObject);
    procedure btneditalternativeClick(Sender: TObject);
    procedure btneditcriteriaClick(Sender: TObject);
    procedure btnSaveproblemClick(Sender: TObject);
    procedure btnSaveStartClick(Sender: TObject);
    procedure btnStartelicitationClick(Sender: TObject);
    procedure StartelicitationChoiRank;
    procedure btnChooseClick(Sender: TObject);
    procedure btnVoltarinicioClick(Sender: TObject);
    procedure btnStep1Click(Sender: TObject);
    procedure btnStep2Click(Sender: TObject);
    Procedure Inicializa;
    procedure CostNorm;
    procedure cmbbxproblemChange(Sender: TObject);
    function TrocaPV(Num: String): Double;
    function TrocaVP(Num: Double): String;
    function TrocaVP1(Num: String): Double;
    function idUnique: String;
    procedure btnslotherfileClick(Sender: TObject);
    procedure SalvarBD;
    procedure cmproblematica1Change(Sender: TObject);
    procedure btnLogoutRgnOpcaoInputClick(Sender: TObject);
    procedure Logout;
    procedure btnLogoutRgnImportExportClick(Sender: TObject);
    procedure btnLogoutRgnInputMatrizConseqClick(Sender: TObject);
    procedure btnResetRgnInputMatrizConClick(Sender: TObject);
    procedure btnLogoutRgnCadProjetoClick(Sender: TObject);
    procedure btnLogoutRgnLoginClick(Sender: TObject);
    procedure rgnintroCreate(Sender: TObject);
    procedure IWDBLabel1AsyncClick(Sender: TObject; EventParams: TStringList);
    procedure AdapterActionAfterExecute(Sender: TObject; Params: TStrings);
    procedure validardiscretos(ValDiscreto: integer);
    function send_email(destinatario, assunto, corpoMensagem: string): boolean;
    procedure btnrecoverClick(Sender: TObject);
    procedure btnvalidaremailClick(Sender: TObject);
    procedure btnnewcodeClick(Sender: TObject);
    procedure edtnewpassconfirClick(Sender: TObject);
    procedure IWButton1Click(Sender: TObject);
    procedure cmbxFOChange(Sender: TObject);
    procedure cmbxScaletypeChange(Sender: TObject);
    procedure btnlogoutrgnconsClick(Sender: TObject);
    procedure rgnInputMatrizConseqCreate(Sender: TObject);
    procedure alinhamento;
    procedure cmbbxAffiliationChange(Sender: TObject);
    procedure gnClick(Sender: TObject);
    procedure btncloseIDCClick(Sender: TObject);
    procedure btnAboutDCClick(Sender: TObject);
    procedure btnDiscrete0Click(Sender: TObject);
    procedure btnmessageDiscClick(Sender: TObject);
    procedure btnConfBackClick(Sender: TObject);
    procedure btnConfNaoClick(Sender: TObject);
    procedure btnminaltClick(Sender: TObject);
    procedure btnpluscritClick(Sender: TObject);
    procedure btnplusaltClick(Sender: TObject);
    procedure btnmincritClick(Sender: TObject);
    procedure btnguideClick(Sender: TObject);
    procedure imgCabecalhoClick(Sender: TObject);
    procedure IWButton2Click(Sender: TObject);
    procedure IWButton3Click(Sender: TObject);
    procedure btnauxchooseClick(Sender: TObject);
    procedure tmrcontinueTimer(Sender: TObject);
    procedure btnauxstartelicitationClick(Sender: TObject);
    procedure tmrcont2Timer(Sender: TObject);
    procedure btnauxstep2Click(Sender: TObject);
    procedure btnauxstep1Click(Sender: TObject);
    procedure tmrstepTimer(Sender: TObject);
    procedure btnportugueseClick(Sender: TObject);
    procedure btnEnglishClick(Sender: TObject);
    procedure btnlcloseClick(Sender: TObject);
    procedure btncloselaClick(Sender: TObject);
    // Procedure Normaliza; // foi para a untintracrit
    procedure ImporttabelaChoiRank;
    procedure IWButton5Click(Sender: TObject);
    procedure imagechoiceClick(Sender: TObject);
    procedure imagerankClick(Sender: TObject);
    procedure imagesortingClick(Sender: TObject);

    // ** procedimentos problematica de portfolio CB**//
    Procedure ImporttabelaPortfolioCB;
    Procedure ImporttabelaClassificacao;
    Procedure ExporttemplatePortfolioCB;
    Procedure StartelicitationPortfolioCB;
    procedure imageportClick(Sender: TObject);
    procedure btnovoproblemaClick(Sender: TObject);
    procedure IWButton7Click(Sender: TObject);
    procedure btnConfResetClick(Sender: TObject);
    procedure IWButton8Click(Sender: TObject);
    procedure lnkcloseaboutClick(Sender: TObject);
    procedure About1Click(Sender: TObject);
    procedure English1Click(Sender: TObject);
    procedure Portuguese1Click(Sender: TObject);

    // procedimentos criados em 20.01.22, devido untIntracriterio//
    procedure rdbtndeclaretypeClick(Sender: TObject);
    procedure rdbtnelictfunctionClick(Sender: TObject);

    // **  Procedimentos para problemática de Portfólio combinatório **//
    Procedure ImporttabelaPortfolioC;
    procedure imgPortCombClick(Sender: TObject);
    procedure tmrPortfoliocombinationTimer(Sender: TObject);
    Procedure StartelicitationPortfolioComb;

    // ** Modificações 24/10 - Sistema de Carla **//
    Procedure Checar_Siglas;
    Procedure Siglas;
    Procedure Definir_Decimais;
    Function trescasas(Num: Double): String;
    Procedure HeuQuest;

    Procedure AreThereDiscCriteria;
    procedure IWLink11Click(Sender: TObject);
    procedure IWLink12Click(Sender: TObject);
    procedure IWLink13Click(Sender: TObject);

  public
  end;

implementation

uses UserSessionUnit, ServerController, lpsolve51, uIntercritério,
  uResultados, uAPI_SolveLP, uCode, Unit1;
{$R *.dfm}

procedure TFrmEntrada.alinhamento;
begin
  rgnintro.BorderOptions.Color := clnone;
  rgnintro.Align := Talign(5);
  rgnCadusuario.Align := Talign(5);
  rgnConseqvalues.Align := Talign(5);
  rgnEnterLogin.Align := Talign(5);
  rgnImportExport.Align := Talign(5);
  rgnLoginproject.Align := Talign(5);
  rgnCadprojeto.Align := Talign(5);
  rgnInputMatrizConseq.Align := Talign(5);
  rgnOpcaoInput.Left := 454;
  rgnOpcaoInput.Top := 100;
  rgnlanguage.Top := 137;
  rgnlanguage.Left := 856;
  rgnproblematica.Left := 300;
  rgnproblematica.Top := 0;
  rgnPortGeneration.Top := 10;
  rgnPortGeneration.Left := 10;
  rgnAbout.Left := 250;
  rgnAbout.Top := 180;

  { if rgnopcaoinput.Visible = true then
    begin
    if TIWUserSession(WebApplication.Data).Problematica = 2 then
    begin
    btncontinueregist.Visible := false;
    IWButton7.Left := 100;
    IWButton7.Top := 303;
    IWButton7.Width := 241;
    end
    else
    begin
    btncontinueregist.Visible := true;
    IWButton7.Left := 100;
    IWButton7.Top := 357;
    IWButton7.Width := 241;
    end;
    end; }

  // ** Se for o primeiro acesso, desabilita o botão retomada do problema **//
  if not TIWUserSession(WebApplication.Data).primeiroacesso then
    btnContinueregist.Enabled := true
  else
    btnContinueregist.Enabled := false;
end;

function TFrmEntrada.send_email(destinatario: string; assunto: string;
  corpoMensagem: string): boolean;
var
  // variáveis e objetos necessários para o envio
  IdSSLIOHandlerSocket: TIdSSLIOHandlerSocketOpenSSL;
  IdSMTP: TIdSMTP;
  IdMessage: TIdMessage;
  IdText: TIdText;
  sAnexo: string;
begin
  // Result:=False;

  // instanciação dos objetos
  IdSSLIOHandlerSocket := TIdSSLIOHandlerSocketOpenSSL.Create(Self);
  IdSMTP := TIdSMTP.Create(Self);
  IdMessage := TIdMessage.Create(Self);

  try
    // Configuração do protocolo SSL (TIdSSLIOHandlerSocketOpenSSL)
    IdSSLIOHandlerSocket.SSLOptions.Method := sslvSSLv23;
    IdSSLIOHandlerSocket.SSLOptions.Mode := sslmClient;
    IdSSLIOHandlerSocket.ReadTimeout := 30000;

    // Configuração do servidor SMTP (TIdSMTP)
    IdSMTP.IOHandler := IdSSLIOHandlerSocket;
    IdSMTP.UseTLS := utUseImplicitTLS;
    IdSMTP.AuthType := satDefault;
    IdSMTP.Port := 465;
    IdSMTP.Host := 'smtp.gmail.com';
    IdSMTP.Username := 'fitradeoff@cdsid.org.br';
    IdSMTP.Password := 'cdsid@ufpe';
    // IdSMTP.ReadTimeout:=30000;

    // Configuração da mensagem (TIdMessage)
    IdMessage.From.Address := 'fitradeoff@cdsid.org.br';
    IdMessage.From.Name := 'FITRADEOFF TECHNICAL SUPPORT';
    IdMessage.ReplyTo.EMailAddresses := IdMessage.From.Address;
    IdMessage.Recipients.EMailAddresses := destinatario;
    IdMessage.Subject := assunto;
    IdMessage.Body.Clear;
    IdMessage.Body.Add(corpoMensagem);

    // Conexão e autenticação
    try
      IdSMTP.Connect;
      IdSMTP.Authenticate;
    except
      on E: Exception do
      begin
        // WebApplication.ShowMessage('Erro na conexão ou autenticação: '+
        // MessageDlg('Erro na conexão ou autenticação: ' +
        // E.Message);
        Result := false;
        Exit;
      end;
    end;

    // Envio da mensagem
    try
      IdSMTP.Send(IdMessage);
      Result := true;
      // MessageDlg('Mensagem enviada com sucesso!', mtInformation, [mbOK], 0);
    except
      On E: Exception do
      begin
        // MessageDlg('Erro ao enviar a mensagem: ' +
        // E.Message, mtWarning, [mbOK], 0);
        Result := false;
      end;
    end;
  finally
    // desconecta do servidor
    IdSMTP.Disconnect;
    // liberação da DLL
    // UnLoadOpenSSLLibrary;
    // liberação dos objetos da memória
    FreeAndNil(IdMessage);
    FreeAndNil(IdSSLIOHandlerSocket);
    FreeAndNil(IdSMTP);
  end;

end;

Function TFrmEntrada.trescasas(Num: Double): String;
var
  S, formato: string;
  i, k: integer;
  inicia: boolean;
begin
  with TIWUserSession(WebApplication.Data) do
  begin
    if StrDInd <> -1 then
      formato := TIWUserSession(WebApplication.Data)
        .StrDecimal[TIWUserSession(WebApplication.Data).StrDInd]
    else
      formato := '0.00';

    StrDInd := -1;
  end;
  TIWUserSession(WebApplication.Data).StrDInd := -1;
  inicia := false;
  k := 0;
  S := FloatToStr(Num);
  // ** Contando o número de casas decimais **//
  for i := 2 to length(S) - 1 do
  begin
    if inicia = true then
      INC(k);
    if (S[i] = ',') or (S[i] = '.') then
      inicia := true;
  end;

  if k + 2 >= length(formato) then
    Result := FormatFloat(formato, Num)
  else
    Result := S;
end;

procedure TFrmEntrada.Siglas;
var
  i, j: integer;
begin
  with TIWUserSession(WebApplication.Data) do
  begin
    // ** Siglas dos critérios **//
    Setlength(CritSiglas, NumCrit);
    for i := 0 to NumCrit - 1 do
    begin
      CritSiglas[i] := '';
      for j := 1 to length(NomeCrit[i]) do
      begin
        CritSiglas[i] := CritSiglas[i] + NomeCrit[i][j];
        if length(CritSiglas[i]) >= 5 then
          Break;
      end;
    end;

    // ** Siglas das alternativas **//
    Setlength(AltSiglas, NumAlt);
    for i := 0 to NumAlt - 1 do
    begin
      AltSiglas[i] := '';
      for j := 1 to length(NomeAlt[i]) do
      begin
        AltSiglas[i] := AltSiglas[i] + NomeAlt[i][j];
        if length(AltSiglas[i]) >= 8 then
          Break;
      end;
    end;
  end;

end;

procedure TFrmEntrada.HeuQuest;
var
  i: integer;
begin
  with TIWUserSession(WebApplication.Data) do
  begin
    for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 2 do
      // ** Se o critério não estiver blindado ou se tiver sido blindado na elicitação **//
      if (TIWUserSession(WebApplication.Data).vtrrelacoesblin[i] = false) or
        (TIWUserSession(WebApplication.Data).Cri_Cong[i] = 1) then
      begin
        TIWUserSession(WebApplication.Data).IdAHeur := TIWUserSession
          (WebApplication.Data).Representantes
          [TIWUserSession(WebApplication.Data).vtrposicoes[i] - 1];
        Indice := i;
        Break;
      end;
    TIWUserSession(WebApplication.Data).IdBHeur := Representantes
      [TIWUserSession(WebApplication.Data).vtrposicoes[NumCrit - 1] - 1];
  end;
end;

procedure TFrmEntrada.Checar_Siglas;
var
  i, j, k, P, T, a, z, Niguais: integer;
  Diferentes, Same, escreve, Siguais: boolean;
  iguais, numIgualdade: array of integer;
  Vchar, Vchar2: set of Char;
  AuxSiglas: Array of string;
  aux, Aux1: string;
begin
  // ** Cria as Siglas **//
  Siglas;
  // ** Fica em looping até que todas as siglas sejam diferentes **//
  with TIWUserSession(WebApplication.Data) do // Comando para chamar as variáveis da userssession//
  begin
    // ** Primeiro checa as siglas dos critérios **//
    T := 0;
    Diferentes := false;
    { while not Diferentes do
      Begin }
    P := 0;
    Setlength(iguais, NumCrit);
    Setlength(numIgualdade, NumCrit);

    // Setlength(AuxSiglas, Numcrit);

    for i := 0 to NumCrit - 1 do
      numIgualdade[i] := 0;

    Vchar := ['A', 'a', 'E', 'e', 'I', 'i', 'O', 'o', 'U', 'u'];
    // Vchar2 := ['\', '/', ')', '(', '*', '#', '*', '-', '+', ',', '.', ';', '|'];

    Diferentes := true;
    Same := false;

    for j := 0 to NumCrit - 1 do
    begin
      iguais[j] := -1;
    end;

    for i := 0 to NumCrit - 1 do
      for j := i + 1 to NumCrit - 1 do
        if CritSiglas[i] = TIWUserSession(WebApplication.Data)
          .CritSiglas[j] then
        begin
          Same := true;
          if iguais[j] = -1 then
          begin
            iguais[j] := i;
            iguais[i] := -5;
          end;
        end;

    if Same then
    begin
      Same := false;
      for j := 0 to NumCrit - 1 do
      begin
        aux := NomeCrit[j][1];
        escreve := false;
        P := 0;

        // ** Tentativa 1 (segunda palavra) **//
        for k := 2 to length(NomeCrit[j]) do
        begin
          if NomeCrit[j][k - 1] = ' ' then
            escreve := true;

          if escreve then
          begin
            aux := aux + NomeCrit[j][k];
            INC(P);
          end;

          if P >= 4 then
          begin
            Break;
          end;
        end;

        TIWUserSession(WebApplication.Data).CritSiglas[j] := aux;
        If length(aux) <= 1 then
          Same := true;
      end; // j

      if not Same then
      begin
        for i := 0 to NumCrit - 1 do
        begin
          for j := i + 1 to NumCrit - 1 do
            if CritSiglas[i] = TIWUserSession(WebApplication.Data)
              .CritSiglas[j] then
            begin
              Same := true;
              Break;
            end;
          if Same then
            Break;
        end; // i
      end; // not same
    end; // same

    if (Same) then
    begin
      Same := false;
      for j := 0 to NumCrit - 1 do
      begin
        aux := '';
        a := 0;

        // ** Tentativa 2 (Consoantes) **//
        for k := 1 to length(NomeCrit[j]) do
          if not(NomeCrit[j][k] in Vchar) then
          begin
            aux := aux + NomeCrit[j][k];
            INC(a);
            if a >= 5 then
            begin
              Break;
            end;
          end;

        TIWUserSession(WebApplication.Data).CritSiglas[j] := aux;
        if length(aux) <= 1 then
          Same := true;
      end; // for j

      if not Same then
      begin
        for i := 0 to NumCrit - 1 do
        begin
          for j := i + 1 to NumCrit - 1 do
            if CritSiglas[i] = TIWUserSession(WebApplication.Data)
              .CritSiglas[j] then
            begin
              Same := true;
              Break;
            end;
          if Same then
            Break;
        end; // i
      end; // not same
    end; // same

    Niguais := 0;
    if (Same) then
    begin

      // ** Tentativa Três (Adicionando número) **//
      for j := 0 to NumCrit - 1 do
      begin
        aux := '';
        Aux1 := '';

        if (iguais[j] <> -1) and (iguais[j] <> -5) then
        begin
          // ** Índice da alternativa com a qual Crit[j] é igual **//
          P := iguais[j];
          // ** Incrementa o número de alternativas iguais à Crit[p] **//
          INC(numIgualdade[P]);
          Aux1 := inttostr(numIgualdade[P] + 1);
        end
        else if (iguais[z] = -5) then
          Aux1 := '1';

        for k := 1 to length(NomeCrit[j]) do
        begin
          aux := aux + NomeCrit[j][k];

          if (k = length(NomeCrit[j])) or (k = 6) then
          begin
            CritSiglas[j] := aux + Aux1;
            Break;
          end;
        end;
      end;

      Diferentes := false;
    end; // same

    // ** Depois checa as siglas das alternativas **//
    T := 0;
    Diferentes := false;
    P := 0;
    Setlength(iguais, NumAlt);
    Setlength(numIgualdade, NumAlt);

    // Setlength(AuxSiglas, Numcrit);

    for i := 0 to NumAlt - 1 do
      numIgualdade[i] := 0;

    Vchar := ['A', 'a', 'E', 'e', 'I', 'i', 'O', 'o', 'U', 'u'];
    // Vchar2 := ['\', '/', ')', '(', '*', '#', '*', '-', '+', ',', '.', ';', '|'];

    Diferentes := true;
    Same := false;

    for j := 0 to NumAlt - 1 do
    begin
      iguais[j] := -1;
    end;

    // ** Verifica e armazena as alternativas iguais, caso não existam igualdades
    // ** Não é preciso modificar a sigla inicialmente proposta **//
    for i := 0 to NumAlt - 1 do
      for j := i + 1 to NumAlt - 1 do
        if AltSiglas[i] = TIWUserSession(WebApplication.Data).AltSiglas[j] then
        begin
          Same := true;
          if iguais[j] = -1 then
          begin
            iguais[j] := i;
            iguais[i] := -5;
          end;
        end;

    if Same then
    begin
      Same := false;
      for z := 0 to NumAlt - 1 do
      begin
        aux := NomeAlt[z][1];
        escreve := false;
        P := 0;
        // ** Tentativa 1 (segunda palavra) **//
        for k := 2 to length(NomeAlt[z]) do
        begin
          if NomeAlt[z][k - 1] = ' ' then
            escreve := true;

          if escreve then
          begin
            aux := aux + NomeAlt[z][k];
            INC(P);
          end;

          if P >= 7 then
          begin
            Break;
          end;
        end;
        TIWUserSession(WebApplication.Data).AltSiglas[z] := aux;
        if length(aux) <= 1 then
          Same := true;
      end; // z

      if not Same then
      begin
        for i := 0 to NumAlt - 1 do
        begin
          for j := i + 1 to NumAlt - 1 do
            if AltSiglas[i] = TIWUserSession(WebApplication.Data)
              .AltSiglas[j] then
            begin
              Same := true;
              Break;
            end;
          if Same then
            Break;
        end; // i
      end; // not same
    end; // same

    if (Same) then
    begin
      Same := false;
      for z := 0 to NumAlt - 1 do
      begin
        aux := '';
        a := 0;
        // ** Tentativa 2 (Consoantes) **//
        for k := 1 to length(NomeAlt[z]) do
          if not(NomeAlt[z][k] in Vchar) then
          begin
            aux := aux + NomeAlt[z][k];
            INC(a);
            if a >= 8 then
            begin
              Break;
            end;
          end;
        TIWUserSession(WebApplication.Data).AltSiglas[z] := aux;
        if length(aux) <= 1 then
          Same := true;
      end; // for z

      if not Same then
      begin
        for i := 0 to NumAlt - 1 do
        begin
          for j := i + 1 to NumAlt - 1 do
            if AltSiglas[i] = TIWUserSession(WebApplication.Data)
              .AltSiglas[j] then
            begin
              Same := true;
              Break;
            end;
          if Same then
            Break;
        end; // i
      end; // not same
    end; // same

    Niguais := 0;
    if (Same) then
    begin
      // ** Tentativa Três (Adicionando número) **//
      for z := 0 to NumAlt - 1 do
      begin
        aux := '';
        Aux1 := '';

        if (iguais[z] <> -1) and (iguais[z] <> -5) then
        begin
          // ** Índice da alternativa com a qual alt[z] é igual **//
          P := iguais[z];
          // ** Incrementa o número de alternativas iguais à alt[p] **//
          INC(numIgualdade[P]);
          Aux1 := inttostr(numIgualdade[P] + 1);
        end
        else if (iguais[z] = -5) then
          Aux1 := '1';

        for k := 1 to length(NomeAlt[z]) do
        begin
          aux := aux + NomeAlt[z][k];

          if (k = length(NomeAlt[z])) or (k = 6) then
          begin
            AltSiglas[z] := aux + Aux1;
            Break;
          end;
        end;
      end;

      Diferentes := false;
    end; // same

  end;
end;

procedure TFrmEntrada.Definir_Decimais;
var
  i: integer;
  j: real;
begin
  with TIWUserSession(WebApplication.Data) do
  begin
    Setlength(StrDecimal, NumCrit);
    for i := 0 to NumCrit - 1 do
    begin
      j := MaxCons[i] - MinCons[i];
      // ** Considerando o módulo da diferença **//
      if j < 0 then
        j := (-1) * j;
      if (j < 100) then
        StrDecimal[i] := '0.00';
      if (j >= 100) then
        TIWUserSession(WebApplication.Data).StrDecimal[i] := '0';

    end;

    StrDInd := -1;
  end;
end;

procedure TFrmEntrada.tmrcont2Timer(Sender: TObject);
begin
  btnStartelicitationClick(Self);
end;

procedure TFrmEntrada.tmrcontinueTimer(Sender: TObject);
begin
  btnChooseClick(Self);
end;

Procedure TFrmEntrada.AreThereDiscCriteria;
var
  i, j, contador: integer;
  StrAux: string;
  contcri: array of integer;
  STR22: string;
begin
  with TIWUserSession(WebApplication.Data) do
  begin
    StrAux := '';
    contcri := 0;
    if Idioma = 0 then
      STR22 := 'and'
    else
      STR22 := 'e';
    for j := 0 to NumCrit - 1 do
      // ** Caso o critério não tenha sido declarado como inteiro nem discreto **//
      if (Typeofcrit[j] < 2) and (not Cri_inteiro[j]) then
      begin
        contador := 0;
        for i := 0 to NumAlt - 1 do
          // ** Consequência inteira **//
          if ConsMatrix[i, j] = Round(ConsMatrix[i, j]) then
            INC(contador);

        if contador = NumAlt then
        begin
          Setlength(contcri, length(contcri) + 1);
          contcri[length(contcri) - 1] := j;
        end;
      end; // if type

    for i := 0 to length(contcri) - 1 do
      if i = 0 then
        StrAux := StrAux + '"' + NomeCrit[contcri[i]] + '"'
      else if i < length(contcri) - 1 then
        StrAux := StrAux + ', "' + NomeCrit[contcri[i]] + '"'
      else
        StrAux := StrAux + ' ' + STR22 + ' "' + NomeCrit[contcri[i]] + '"';
    if Idioma = 0 then
      if StrAux <> '' then
        if length(contcri) = 1 then
          WebApplication.ShowMessage
            ('All consequences in the criterion ' + StrAux +
              ' are integer. We advise you to declare it as:' + #13#10 +
              '> A discrete criterion -' +
              ' in case it is evaluated trough a built scale;' + #13#10 +
              '> An integer criterion -' +
              ' in case it only admits integer consequences (i.e number of people);'
              + #13#10 +
              '> Keep it as a continuous criterion - in case it also admits non-integer consequences.' +
              #13#10 + 'If you have any doubts, take a look at the user guide.')
        else
          WebApplication.ShowMessage
            ('All consequences in the criteria ' + StrAux +
              ' are integer. We advise you to declare them as:' + #13#10 +
              '> Discrete criteria -' +
              ' in case they are evaluated trough built scales;' + #13#10 +
              '> Integer criteria -' +
              ' in case they only admit integer consequences (i.e. number of people); ' +
              #13#10
              + '> Keep them as a continuous criteria - in case they also admit non-integer consequences.' + #13#10 + 'If you have any doubts, take a look at the user guide.');

    if Idioma = 1 then
      if StrAux <> '' then
        if length(contcri) = 1 then
          WebApplication.ShowMessage
            ('Todas as consequências no critério ' + StrAux +
              ' são inteiras. Aconselhamos que o declare como:' + #13#10 +
              '> Critério discreto -' +
              ' caso seja avaliado através de uma escala construída;' +
              #13#10 + '> Critério inteiro -' +
              ' caso apenas admita consequências inteiras(ex: número de pessoas);'
              + #13#10 +
              '> Mantenha-o como contínuo - caso também admita consequências reais não inteiras.' +
              #13#10 + 'Em caso de dúvidas, consulte o guia.')
        else
          WebApplication.ShowMessage
            ('Todas as consequências nos critérios ' + StrAux +
              ' são inteiras. Aconselhamos que os declare como: ' + #13#10 +
              '> Critérios discretos -' +
              ' caso sejam avaliados através de uma escala construída;' +
              #13#10 + '> Critérios inteiros -' +
              ' caso apenas admitam consequências inteiras(ex: número de pessoas);'
              + #13#10 +
              '> Mantenha-os como contínuos - caso também admitam consequências reais não inteiras.' +
              #13#10 + 'Em caso de dúvidas, consulte o guia.');

  end; // with
end;

procedure TFrmEntrada.tmrPortfoliocombinationTimer(Sender: TObject);
var
  percentual: real;
  Str1: string;
begin
  {with UserSession.ZQuery1 do
  begin
    // ** Verifica se já existem portfólios sendo gerados **//
    SQL.Text := 'SELECT percentual FROM line_portfolio WHERE id_problem="' +
      (TIWUserSession(WebApplication.Data).ID_Problema) + '";';
    ExecSQL;
    open;
    Str1 := FieldByName('percentual').AsString;
    close;
  end;}
  percentual := StrToFloat(Str1);
  PgbarPort.Percent := Round(percentual);

  if percentual = 100 then
  begin
    // ** Pega os dados de volta do banco <Portfólios gerados> **//
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 43;
    SalvarBD;
    // ** Desabilita o Timer **//
    tmrPortfoliocombination.Enabled := false;
    // ** Oculta a região de espera **//
    rgnPortGeneration.Hide;
    // ** Indica que a geração dos portfólios já foi concluída **//
    TIWUserSession(WebApplication.Data).PortfolioG := false;
    // ** Segue o fluxo do sistema **//
    btnStartelicitationClick(Self);
  end;
end;

procedure TFrmEntrada.tmrstepTimer(Sender: TObject);
begin
  case TIWUserSession(WebApplication.Data).stage of
    1:
      btnStep1Click(Self);
    2:
      btnStep2Click(Self);
  end;

end;

{ procedure TFrmEntrada.Normaliza;
  var
  i, j: integer;
  Zmax, Zmin: Double;
  matriztestesnormal: array of array of Double;
  auxcost: string;
  begin
  /// / Esse procedimento realiza a normalização das consequências considerando /
  /// a função valor linear para todos os critérios. ////////////////////////////
  with TIWUserSession(WebApplication.Data) do
  begin
  setlength(TIWUserSession(WebApplication.Data).ConsNorm, NumAlt, numcrit);
  for j := 0 to numcrit - 1 do
  for i := 0 to NumAlt - 1 do

  begin
  /// verificar se os criterios DISCRETOS devem seguir uma regra diferente de transformação da escala///
  case FTipo[j] of // função que armazena o vetor dos tipos de FO declaradas na planilha

  1: // procedimento realiza a normalização das consequências considerando a função valor linear ///////////////////////
  begin
  // Normalização para os critérios de maximização //////////////////////
  if (Typeofcrit[j] = 1) or (Typeofcrit[j] = 3) then
  ConsNorm[i, j] := (ConsMatrix[i, j] - Min[j]) /
  (Max[j] - Min[j]);
  /// Normalização para os critérios de minimização //////////////////////
  if (Typeofcrit[j] = 0) or (Typeofcrit[j] = 2) then
  ConsNorm[i, j] := (ConsMatrix[i, j] - Max[j]) /
  (-Max[j] + Min[j]);

  end;

  2: // procedimento realiza a normalização das consequencias considerando função valor exponencial
  begin
  if Min[j] = 0 then
  Min[j] := Power(10, -5); // O QUE É POWER??
  Zmin := Exp(parA[j] * Min[j]); // exponencial do parametro A vezes o valor minimo capturado no criterio
  Zmax := Exp(parA[j] * Max[j]); // exponencial do parametro A vezes o valor máximo capturado no criterio

  // normalização para os criterios de maximização//
  if (Typeofcrit[j] = 1) or (Typeofcrit[j] = 3) then
  ConsNorm[i, j] := (Exp(parA[j] * ConsMatrix[i, j]) - Zmin) /
  (Zmax - Zmin);
  // normalização para os criterios de minimização//
  if (Typeofcrit[j] = 0) or (Typeofcrit[j] = 2) then
  ConsNorm[i, j] := (Exp(parA[j] * ConsMatrix[i, j]) - Zmax) /
  (Zmin - Zmax);
  end;

  3: // procedimento realiza a normalização das consequencias considerando função valor logaritmica
  begin
  if Min[j] = 0 then
  Min[j] := Power(10, -5); // O QUE É POWER??
  Zmin := Ln(parA[j] * Min[j] + 1 - parA[j] * Min[j]);
  Zmax := Ln(parA[j] * Max[j] + 1 - parA[j] * Min[j]);

  // normalização para os criterios de maximização///
  if (Typeofcrit[j] = 1) or (Typeofcrit[j] = 3) then
  ConsNorm[i, j] :=
  (Ln(parA[j] * ConsMatrix[i, j] + 1 - parA[j] * Min[j])
  - Zmin) / (Zmax - Zmin);
  // normalização para os criterios de minimização//
  if (Typeofcrit[j] = 0) or (Typeofcrit[j] = 2) then
  ConsNorm[i, j] :=
  (Ln(parA[j] * ConsMatrix[i, j] + 1 - parA[j] * Min[j])
  - Zmax) / (Zmin - Zmax);
  {else
  ConsNorm[i, j] := 1;
  end; }

{ 4: // procedimento realiza a normalização das consequencias considerando função valor logistica (S-Shape)
  begin
  if Min[j] = 0 then
  Min[j] := Power(10, -5);

  Zmax := Exp(-parA[j] / Max[j]);
  Zmin := Exp(-parA[j] / Min[j]);

  // normalização para os criterios de maximização///
  if (Typeofcrit[j] = 1) or (Typeofcrit[j] = 3) then
  ConsNorm[i, j] := (Exp(-parA[j] / ConsMatrix[i, j]) - Zmin) /
  (Zmax - Zmin);
  // normalização para os criterios de minimização//
  if (Typeofcrit[j] = 0) or (Typeofcrit[j] = 2) then
  ConsNorm[i, j] := (Exp(-parA[j] / ConsMatrix[i, j]) - Zmax) /
  (Zmin - Zmax);
  { else
  ConsNorm[i, j] := 1;
  end;
  end;

  end; // final dos contadores

  //Este trecho foi levado para executar na untIntracrit- 21.01.22 //

  /// / Se a problematica for de portfolio então aplica-se uma transformação linear
  /// ao vetor de custos e se calcula o fatormultiplicativo que é uma espécie
  /// de fator corretivo para os parâmetros utilizados na verificação sas relações
  /// de dominância dado o efeito da transformação no valor da max dif ////
  if Problematica = 2 then
  begin
  // ** Identificando o custo maximo **//
  TIWUserSession(WebApplication.Data).MaxCost := TIWUserSession
  (WebApplication.Data).cost[0];
  TIWUserSession(WebApplication.Data).MinCost := cost[0];
  for i := 0 to NumAlt - 1 do
  begin
  if TIWUserSession(WebApplication.Data).cost[i] > TIWUserSession
  (WebApplication.Data).MaxCost then
  TIWUserSession(WebApplication.Data).MaxCost := TIWUserSession
  (WebApplication.Data).cost[i];
  if TIWUserSession(WebApplication.Data).cost[i] < TIWUserSession
  (WebApplication.Data).MinCost then
  TIWUserSession(WebApplication.Data).MinCost := TIWUserSession
  (WebApplication.Data).cost[i];
  end;
  // ** Verifica a ordem de grandeza do maior custo **//
  auxcost := FloatToStr(MinCost);
  j := 0;
  for i := 0 to length(auxcost) - 1 do
  if auxcost[i] <> '' then
  if (auxcost[i] = ',') or (auxcost[i] = '.') then
  begin
  INC(j, -1);
  break;
  end
  else
  INC(j);
  // ** Constrói um múltiplo de 10 de grandeza semelhante ao custo máximo **//
  auxcost := '1';
  for i := 0 to j - 1 do
  auxcost := auxcost + '0';

  // ** Calcula o fator multiplicativo **//
  fmult := TIWUserSession(WebApplication.Data).MinCost / StrToInt(auxcost);

  // ** Transformação do vetor de custos **//
  setlength(CostNorm, NumAlt);

  for i := 0 to NumAlt - 1 do
  TIWUserSession(WebApplication.Data).CostNorm[i] := cost[i]
  / TIWUserSession(WebApplication.Data).MinCost;

  end; // problematica = 2
  end; // final dos casos
  end; }

procedure TFrmEntrada.Portuguese1Click(Sender: TObject);
begin
  WebApplication.SendFile(
    'C:\inetpub\fitradeoff\Images\Guia rapido Fitradeoff web' + '.pdf', true);
end;

procedure TFrmEntrada.validardiscretos(ValDiscreto: integer);
/// / Procedimento responsável por fazer a validação da escala introduzida para
/// os critérios discretos. //////////////////////////////////////////////////
var
  i, j, S, P, NumCri, NumAlt: integer;
  Value: real;
  NonInteger, OutofScale: array of boolean;
  Outof, NonInt: string;

begin

  /// //////////////////////// Validação dos dados ////////////////////////////////
  NumCri := TIWUserSession(WebApplication.Data).NumCrit;
  NumAlt := TIWUserSession(WebApplication.Data).NumAlt;

  // setando tamanho dos vetores auxiliares
  Setlength(NonInteger, NumCri + 1);
  Setlength(OutofScale, NumCri + 1);
  Outof := '';
  NonInt := '';

  for i := 0 to NumCri - 1 do
  begin

    if (TIWUserSession(WebApplication.Data).Niveis[i] > 2) and
      (TIWUserSession(WebApplication.Data).Niveis[i] < 8) then
    // Para validar somente Cri discretos
    begin

      for j := 0 to NumAlt - 1 do
      begin
        Value := TIWUserSession(WebApplication.Data).ConsMatrix[j, i];
        // Recebendo as consequências
        // Verifica se a consequência é um valor inteiro
        P := Round(Value); // Arredondando a cnsequência para o inteiro mais próximo

        If (Value - P) <> 0 then
        begin
          NonInteger[0] := true; // O primeiro índice sendo verdadeiro indica a existência de ao menos uma consequência não-inteira em pelo menos um critério discreto
          NonInteger[i + 1] := true; // Salva qual critério possui uma ou mais consequências não-inteiras
        end; // If (Value-P) <> 0

        if (Value < 1) or (Value > TIWUserSession(WebApplication.Data)
            .Niveis[i]) then // Verifica se as consequências estão dentro da escala
        begin
          OutofScale[0] := true; // O primeiro índice sendo verdadeiro indica a existência de ao menos uma consequência fora da escala em pelo menos um critério discreto
          OutofScale[i + 1] := true; // Salva qual critério possui uma ou mais consequências fora da escala
        end; // (Value < 1) or (Value > NiveisDiscreto[i])

      end; // For J

    end // If Niveisdiscreto
    Else if TIWUserSession(WebApplication.Data).Niveis[i] = 2 then
    begin

      for j := 0 to NumAlt - 1 do
      begin
        Value := TIWUserSession(WebApplication.Data).ConsMatrix[j, i];
        // Recebendo as consequências
        // Verifica se a consequência é um valor inteiro
        P := Round(Value); // Arredondando a cnsequência para o inteiro mais próximo

        If (Value - P) <> 0 then
        begin
          NonInteger[0] := true; // O primeiro índice sendo verdadeiro indica a existência de ao menos uma consequência não-inteira em pelo menos um critério discreto
          NonInteger[i + 1] := true; // Salva qual critério possui uma ou mais consequências não-inteiras
        end; // If (Value-P) <> 0

        if (Value < 0) or (Value > 1) then // Verifica se as consequências estão dentro da escala
        begin
          OutofScale[0] := true; // O primeiro índice sendo verdadeiro indica a existência de ao menos uma consequência fora da escala em pelo menos um critério discreto
          OutofScale[i + 1] := true; // Salva qual critério possui uma ou mais consequências fora da escala
        end; // (Value < 1) or (Value > NiveisDiscreto[i])

      end; // For J

    end; // If Niveisdiscreto = 2

  end; // For I

  { O código abaixo tem como função apenas escrever quais critérios Precisam
    ser reavaliados pelo usuário identificando o mesmo através de mensagens }
  if (OutofScale[0] = true) then
  Begin
    S := 0;
    P := 0;
    for i := 1 to NumCri do
      if OutofScale[i] = true then
        S := S + 1;

    i := 1;

    while P < S do
    begin
      if OutofScale[i] = true then
      begin
        if P = 0 then
          Outof := TIWUserSession(WebApplication.Data).NomeCrit[i - 1]
        else if P = S - 1 then
          Outof := Outof + ' and ' + TIWUserSession(WebApplication.Data)
            .NomeCrit[i - 1]
        else
          Outof := Outof + ', ' + TIWUserSession(WebApplication.Data)
            .NomeCrit[i - 1];
        P := P + 1;
      end; // if
      i := i + 1;
    end; // While

    if S > 1 then
      WebApplication.ShowMessage(
        'It has been found Out-Scale-Consequences in the criteria ' + Outof +
          '. Please, check the "Important information about discrete criteria".'
        )
    else
      WebApplication.ShowMessage(
        'It has been found Out-Scale-Consequences in the criterion ' + Outof +
          '. Please, check the "Important information about discrete criteria"'
        );
  End; // IF

  if (NonInteger[0] = true) then
  Begin
    S := 0;
    P := 0;
    for i := 1 to NumCri do
      if NonInteger[i] = true then
        S := S + 1;

    i := 1;

    while P < S do
    begin
      if NonInteger[i] = true then
      begin
        if P = 0 then
          NonInt := TIWUserSession(WebApplication.Data).NomeCrit[i - 1]
        else if P = S - 1 then
          NonInt := NonInt + ' and ' + TIWUserSession(WebApplication.Data)
            .NomeCrit[i - 1]
        else
          NonInt := NonInt + ', ' + TIWUserSession(WebApplication.Data)
            .NomeCrit[i - 1];
        P := P + 1;
      end; // if
      i := i + 1;
    end; // While

    if S > 1 then
      WebApplication.ShowMessage(
        'It has been found Non-Integer-Consequences in the criteria ' +
          NonInt + '. Please, check it before continuing.')
    else
      WebApplication.ShowMessage(
        'It has been found Non-Integer-Consequences in the criterion ' +
          NonInt + '. Please, check it before continuing.');

  End; // IF
  if (NonInteger[0] = true) or (OutofScale[0] = true) then
    TIWUserSession(WebApplication.Data).ValDiscreto := -1
  else
    TIWUserSession(WebApplication.Data).ValDiscreto := 1;
  /// ///////////////////////////////////////////////////////////////////////////

end;

procedure TFrmEntrada.Logout;
begin

  rgnEnterLogin.Visible := false;
  rgnOpcaoInput.Visible := false;
  rgnImportExport.Visible := false;
  rgnCadprojeto.Visible := false;
  rgnInputMatrizConseq.Visible := false;
  rgnLoginproject.Visible := false;
  rgnConseqvalues.Visible := false;
  rgnproblematica.Visible := false;
  rgnintro.Visible := true;
  edtemail.Clear;
  edtSenha.Clear;
  lblbudget.Visible := false; // adicionado aqui para garantir que componente não aparecerá qndo sair
  lbllembretebudget.Visible := false; // adicionado aqui para garantir que componente não aparecerá qndo sair
  edtbudget.Visible := false; // adicionado aqui para garantir que componente não aparecerá qndo sair
  hptpbud.Visible := false; // adicionado aqui para garantir que componente não aparecerá qndo sair

  rgnchoice.BorderOptions.NumericWidth := 0;
  rgnrank.BorderOptions.NumericWidth := 0;
  rgnsorting.BorderOptions.NumericWidth := 0;
  rgnport.BorderOptions.NumericWidth := 0;
  TIWUserSession(WebApplication.Data).Problematica := 0;

end;

procedure TFrmEntrada.rdbtndeclaretypeClick(Sender: TObject);
/// neste, o decisor declara, na entrada manual de dados, que deseja fornecer as
/// funções analisticas de cada criterio anlisado, bem como parametros, quando aplicavel

begin
  rdbtnelictfunction.Checked := false; // outro radio buton é desabilitado pra não dar bug
  rdbtndeclaretype.Checked := true;
  cmbxFO.ItemIndex := 1; // default linear
  lblTipoproblema.Visible := true; // libera componentes para declarar tipo FO
  cmbxFO.Visible := true;

end;

procedure TFrmEntrada.rdbtnelictfunctionClick(Sender: TObject);
/// neste, o decisor declara, na entrada manual de dados, que deseja realizar
/// a elicitação intra. Dando a ideia inicial que Ftipo=0

var
  i: integer;

begin
  rdbtndeclaretype.Checked := false; // outro radio buton é desabilitado pra não dar bug
  rdbtnelictfunction.Checked := true;
  // TIWUserSession(WebApplication.Data).AuxFuncTipo[i]:=0;

  lblTipoproblema.Visible := false; // esconde componentes para declarar tipo FO
  cmbxFO.ItemIndex := -1; // garantir que sempre esteja desmarcado
  cmbxFO.Visible := false;
  txtparametros.Visible := false;
  lblparA.Visible := false;
  edtparA.Visible := false;
  lblparB.Visible := false;
  edtparB.Visible := false;
  lblfunc.Visible := false; // só tem que aparecer quando altera cmbbxFO

end;

procedure TFrmEntrada.rgnInputMatrizConseqCreate(Sender: TObject);
begin

  if TIWUserSession(WebApplication.Data).AUXLOGOUT = -363636 then
  begin
    rgnintro.Visible := true;
    rgnInputMatrizConseq.Visible := false;
    TIWUserSession(WebApplication.Data).AUXLOGOUT := 0;
  end;
end;

procedure TFrmEntrada.rgnintroCreate(Sender: TObject);
begin
  if TIWUserSession(WebApplication.Data).VerificaUnit = -50 then
  begin
    // **Altera a label da problematica **//
    case TIWUserSession(WebApplication.Data).Problematica of
      0:
        begin
          lblproblematic1.Text := 'Choice';
          lblchosenproblematic.Text := 'Choice';
        end;
      1:
        begin
          lblproblematic1.Text := 'Ranking';
          lblchosenproblematic.Text := 'Ranking';
        end;
      2:
        begin
          lblproblematic1.Text := 'Portfolio CB';
          lblchosenproblematic.Text := 'Portfolio CB';
        end;
      3:
        begin
          lblproblematic1.Text := 'Sorting';
          lblchosenproblematic.Text := 'Sorting';
        end;
      4:
        begin
          lblproblematic1.Text := 'Combinatorial Portfolio';
          lblchosenproblematic.Text := 'Combinatorial Portfolio';
        end;
    end;
    rgnintro.Visible := false;
    rgnInputMatrizConseq.Visible := true;
    // **Deixa visível a opção para seguir pra elicitação direto **//
    AlreadyRakend1.Visible := true;
    edtEntername.Text := TIWUserSession(WebApplication.Data).probname;
    btnStartelicitationClick(Self); // Constroi o grid com as variáveis armazenada
    TIWUserSession(WebApplication.Data).VerificaUnit := 0;

  end;

  if TIWUserSession(WebApplication.Data).VerificaUnit = -55 then
  begin
    // **Altera a label da problematica **//
    case TIWUserSession(WebApplication.Data).Problematica of
      0:
        begin
          lblproblematic1.Text := 'Choice';
          lblchosenproblematic.Text := 'Choice';
        end;
      1:
        begin
          lblproblematic1.Text := 'Ranking';
          lblchosenproblematic.Text := 'Ranking';
        end;
      2:
        begin
          lblproblematic1.Text := 'Portfolio CB';
          lblchosenproblematic.Text := 'Portfolio CB';
        end;
      3:
        begin
          lblproblematic1.Text := 'Sorting';
          lblchosenproblematic.Text := 'Sorting';
          // **Inicializa o vetor com tamanho 0 ao clicar para inserir um novo prob **//
          Setlength(TIWUserSession(WebApplication.Data).perfis, 0);
          // ** Informa que não se trata de um reinício do problema **//
          TIWUserSession(WebApplication.Data).ClassRestart := false;
        end;
      4:
        begin
          lblproblematic1.Text := 'Combinatorial Portfolio';
          lblchosenproblematic.Text := 'Combinatorial Portfolio';
        end;
    end;
    rgnintro.Visible := false;
    rgnOpcaoInput.Visible := true;
    btnslotherfileClick(Self); // Limpas as variáveis para recomeçar
    Setlength(TIWUserSession(WebApplication.Data).Vetorpergunta, 0);
    TIWUserSession(WebApplication.Data).VerificaUnit := -55;
  end;
  edtemail.Clear;
  edtSenha.Clear;
  alinhamento;
end;

procedure TFrmEntrada.About1Click(Sender: TObject);
begin
  rgnAbout.show;
end;

procedure TFrmEntrada.AdapterActionAfterExecute
  (Sender: TObject; Params: TStrings);
begin

  btnExporttemplateClick(Self);
end;

procedure TFrmEntrada.btnAboutDCClick(Sender: TObject);
begin
  rgnaboutdiscrete.Visible := true;
  rgntabela.Visible := false;
  rgntabelaalter.Visible := false;
end;

procedure TFrmEntrada.btnAddalternativeClick(Sender: TObject);
var
  i, j, a, b, c: integer;
  // apagar depois (variáveis locais pra verificação)
  nome: array of string;
  Num: integer;
begin

  (* Parte do código referente ao Nùmero de Alternativas preenchido pelo usuário que não é mais utilizado *)
  { Try
    TIWUserSession(WebApplication.Data).NumAlt := strtoint
    (edtNumberalternat.text);
    except
    begin
    WebApplication.ShowMessage(
    'Please, fill in the field "Number of Alternatives" only with numbers');
    Exit;
    end;
    End; }
  { if (strtoint(edtNumberalternat.text) < 2) or (edtNumberalternat.text = '')
    then
    begin
    WebApplication.ShowMessage(
    'Please, fill in the number of alternatives. It should be at least 2.');
    Exit;
    end; }

  { Verificações }
  if edtNamealternat.Text = '' then
  begin
    WebApplication.ShowMessage('Please enter a name for the alternative');
    Exit;
  end;

  with TIWUserSession(WebApplication.Data) do
  begin

    for a := 0 to lstbxNamealternat.Items.Count - 1 do
    begin
      if (lstbxNamealternat.Items[a] = edtNamealternat.Text) then
      begin
        if (EditaAlt = 0) OR
          ((EditaAlt = 1) AND (a <> lstbxNamealternat.ItemIndex)) then
        begin
          WebApplication.ShowMessage
            ('There already is an alternative called ' + edtNamealternat.Text +
              '.');
          Exit;
        end; // if
      end; // for
    end;

    // Se nenhuma alternativa tiver sido selecionada, o índice vai ser -1 e sempre diferente de a.
    // Caso uma alternativa tenha sido selecionada, a verificação é feita entre todos os elementtos do listbox excetoa
    // permitindo que o usuário adicione a mesma alternativa quando tiver editando (em outras palavras, que o usuário aperte p editar, não faça nenhuma alteração e salve ovamnte)

    { Passando pelas verificações, uma alternativa foi, de fato, adicionada ou editada
      e o listbox pode ser alterado(uma nova alternativa pode ser selecionada) }
    lstbxNamealternat.Enabled := true;

    Setlength(AuxNomeAlt, 100); // um valor qualquer muito grande

    { if lstbxNamealternat.ItemIndex = -1 then // Adcionando novo
      begin
      EditaAlt := 0;
      end; }
    // Se qualquer nome tiver sido selecionado mas o botão editar não tiver sido pressionado, EditaAlt tem que ser zero do mesmo jeito

    if EditaAlt = 1 then // editando
    begin
      // Alteração no listbox
      b := AuxIndiceAlt; // AuxIndiceAlt definido no botão de editar alternativa (índice no listbox)
      lstbxNamealternat.Items[b] := edtNamealternat.Text;

      // Descobrindo o índce da alternativa a ser editada no vetor com os nomes das alternativas
      // Nesse vetor auxiliar as alternativas apagadas estão com espaço vazio, não sabemos até onde ele está preenchido
      for c := 0 to 99 do
      begin
        if TIWUserSession(WebApplication.Data).AuxNomeAlt[c] = TIWUserSession
          (WebApplication.Data).NomeAlternativaEditada then
        begin
          i := c; // i é o índice da alternativa que será alterada *no vetor*
          Break;
        end;
      end;
    end;

    if EditaAlt = 0 then // criando uma nova
    begin
      i := AuxNumAlt; // inicializado com zero e depois incrementado a medida que se add alt
      AuxNumAlt := AuxNumAlt + 1;
      lstbxNamealternat.Items.Add(edtNamealternat.Text);
    end;

    TIWUserSession(WebApplication.Data).AuxNomeAlt[i] := edtNamealternat.Text;

    edtNamealternat.Clear;
    lstbxNamealternat.ItemIndex := -1;

    /// ////// APAGAR  //////////////
    Setlength(nome, 100);
    Num := AuxNumAlt;
    for j := 0 to AuxNumAlt - 1 do
    begin
      nome[j] := AuxNomeAlt[j];
    end;
    /// /////////////

    if (StrToInt(edtNumberalternat.Text) < (lstbxNamealternat.Items.Count)) then
    begin
      edtNumberalternat.Text := inttostr(lstbxNamealternat.Items.Count);
    end;
    // btnminalt.Enabled := true;  //Esse botão não está mais sendo utilizado
  end;

  // No final do procedimento uma alternativa já foi adicionada ou editada, então
  // a variável edita alt volta a ser zero indicando que uma nova alternativa será adicionada ´na próxima vez, a menos que o botão editar seja pressionado e altere esse parâmetro
  TIWUserSession(WebApplication.Data).EditaAlt := 0;
  btnAddalternative.Caption := 'Add';

end;

procedure TFrmEntrada.btnAddattributClick(Sender: TObject);
var
  i, j, a, b, c: integer;
  verificacrit: integer;
  VerificaParA: Double;

  // apagar depois (variáveis locais pra verificação)
  nome: array of string;
  afunctipo, aescala, adirec, aniveis: array of integer;
  apara, aparb: array of Double;

begin

  { ---- Validando os critérios ------
    Try
    verificacrit := strtoint(edtNumberattribut.text);
    except
    begin
    WebApplication.ShowMessage(
    'Please, fill in the field "Number of Criteria" with numbers only.');
    Exit;
    end;
    End; }

  { if (edtNumberattribut.text = '') then
    begin
    WebApplication.ShowMessage('Please fill in the number of criteria.');
    Exit;
    end; }

  if edtNameattribut.Text = '' then
  begin
    WebApplication.ShowMessage('Please enter the name of criteria.');
    Exit;
  end;

  with TIWUserSession(WebApplication.Data) do
  begin
    for a := 0 to lstbxNamecriteria.Items.Count - 1 do
    begin
      if (lstbxNamecriteria.Items[a] = edtNameattribut.Text) then
      begin
        if (EditaCrit = 0) OR
          ((EditaCrit = 1) AND (a <> lstbxNamecriteria.ItemIndex)) then
        begin
          WebApplication.ShowMessage
            ('There already is a criterion called ' + edtNameattribut.Text +
              '.');
          Exit;
        end;
      end;
    end;

    Setlength(AuxNomeCrit, 100); // um valor qualquer muito grande
    Setlength(AuxFuncTipo, 100);
    Setlength(AuxEscalaTipo, 100);
    Setlength(AuxDirecTipo, 100);
    Setlength(AuxNiveis, 100);
    Setlength(AuxParA, 100);
    Setlength(AuxParB, 100);

    /// //////APAGAR (verificação) ////////////////
    Setlength(nome, 100);
    Setlength(afunctipo, 100);
    Setlength(aescala, 100);
    Setlength(adirec, 100);
    Setlength(aniveis, 100);
    Setlength(apara, 100);
    Setlength(aparb, 100);
    /// ///////////////////////////

    { ------------- LISTBOX CRITÉRIO ---------------- }

    { if lstbxNamecriteria.ItemIndex = -1 then
      // Nenhuma alternativa selecionada
      begin
      EditaCrit := 0;
      end; }

    if EditaCrit = 1 then // Editando um critério
    begin
      b := AuxIndice; // Indice *no listbox* critério selecionado para edição
      // lstbxNamecriteria.Items[I] := edtNameattribut.text;
      i := IndiceCriterioEditado;
      // ** Declaração do critério inteiro **//
      if chckbxinteger.Checked then
        Cri_inteiro[i] := true
      else
        Cri_inteiro[i] := false;
    end;

    if EditaCrit = 0 then // Criando um novo
    begin
      i := AuxNumCrit; // inicializada com zero
      AuxNumCrit := AuxNumCrit + 1;
      // lstbxNamecriteria.Items.Add(edtNameattribut.text);
      // ** Declaração do critério inteiro **//
      Setlength(Cri_inteiro, length(Cri_inteiro) + 1);
      if chckbxinteger.Checked then
        Cri_inteiro[length(Cri_inteiro) - 1] := true
      else
        Cri_inteiro[length(Cri_inteiro) - 1] := false;
    end;

    // O nome do critério vai ser passado para o listbox no final do procedimento, depois de todas as verificações dos parâmetros

    { ---------------- DIREÇÃO DO CRITÉRIO ------------------------- }
    begin
      case cmbxCriteriondirection.ItemIndex of
        - 1:
          begin
            WebApplication.ShowMessage
              ('Please select the criterion direction.');
            if EditaCrit = 0 then
              AuxNumCrit := AuxNumCrit - 1;
            Exit;
          end;
        0: // maximização
          AuxDirecTipo[i] := 0;
        1: // minimização
          AuxDirecTipo[i] := 1;
      end;
    end;

    { ----------------------- ESCALA DO CRITÉRIO -------------------- }
    case cmbxScaletype.ItemIndex of
      - 1:
        begin
          WebApplication.ShowMessage('Please select the scale type.');
          if EditaCrit = 0 then
            AuxNumCrit := AuxNumCrit - 1;
        end;
      0: // contínuo
        begin
          AuxEscalaTipo[i] := 0; // contínuo
          AuxNiveis[i] := 0; // é zero mesmo?
        end;
      1: // discreto
        begin
          AuxEscalaTipo[i] := 1; // discreto
          if (edtNumberlevels.Text = '') OR
            (StrToInt(edtNumberlevels.Text) > 7) OR
            (StrToInt(edtNumberlevels.Text) < 2) then
          begin // verificação do número de níveis
            WebApplication.ShowMessage(
              'Please enter a valid number of levels. It should be between 2 and 7.');
            if EditaCrit = 0 then
              AuxNumCrit := AuxNumCrit - 1;
            Exit;
          end
          else
          begin // número de níveis para critérios N-lineares
            AuxNiveis[i] := StrToInt(edtNumberlevels.Text);
          end;
        end;
    end;

    { -------------------- TIPO DA FUNÇÃO -------------------------- }

    begin
      if cmbxFO.ItemIndex = -1 then // verificação para elicitar todas as funções
      { if (rdbtndeclaretype.Checked = false) AND
        (rdbtnelictfunction.Checked = false) then }
      // NADA FOI SELELCIONADO

      begin { ALTERADO EM 01.03.22- DEVIDO INSERÃO DA UNITINTRA }
        if TIWUserSession(WebApplication.Data).Idioma = 0 then
          WebApplication.ShowMessage(
            'Please select an action to the Intra-criteria evaluation')
        else
          WebApplication.ShowMessage(
            'Por favor, selecione uma opção para a Avaliação Intracritério');
      end;

      if (rdbtndeclaretype.Checked = true) OR
        (rdbtnelictfunction.Checked = true) then
      BEGIN
        case cmbxFO.ItemIndex of
          0: // Logarithmic
            AuxFuncTipo[i] := 3;
          1: // Linear
            begin
              AuxFuncTipo[i] := 1;
              AuxParA[i] := 0;
              AuxParB[i] := 0;
            end;
          2: // Logistic
            AuxFuncTipo[i] := 4;
          3: // Exponential
            AuxFuncTipo[i] := 2;
          -1: // nenhuma seleção
            begin
              if rdbtnelictfunction.Checked = true then
                // mas botao de elicitar estiver marcado
                AuxFuncTipo[i] := 0;
              // TIWUserSession(WebApplication.Data).Verificaintra[i]:= true;

              if AuxEscalaTipo[i] = 1 then // se for crit discreto, mostra advertencia
              begin
                WebApplication.ShowMessage(
                  'Please select the criterion value function');
                if EditaCrit = 0 then
                  AuxNumCrit := AuxNumCrit - 1;
                Exit;
              end;
            end;
        end;
      END;

      if cmbxFO.ItemIndex <> 1 then // caso não linear
      begin
        if AuxEscalaTipo[i] = 1 then // se for crit discreto, mostra advertencia. relaxa trava pra continuo
        begin
          if (edtparA.Text = '') then
          begin
            WebApplication.ShowMessage(
              'Please fill in the function parameter "A".');

            if EditaCrit = 0 then
              AuxNumCrit := AuxNumCrit - 1;
            Exit;
          end;
        end
        else

          if AuxEscalaTipo[i] = 1 then // se for crit discreto, mostra advertencia. relaxa trava pra continuo
        begin
          if edtparB.Text = '' then
          begin
            WebApplication.ShowMessage(
              'Please fill in the function parameter "B".');

            if EditaCrit = 0 then
              AuxNumCrit := AuxNumCrit - 1;
            Exit;
          end;
        end
        // O erro de não aceitar o ParA como não inteiro utilizando ponto ocorre aqui
        // ele não consegue fazer essa comparação

        else
        begin
          if AuxEscalaTipo[i] = 1 then // se for crit discreto,
          begin
            if trocapv(edtparB.Text) = 0 then
            begin
              WebApplication.ShowMessage(
                'The function parameter "B" cannot assume the value 0. Please, introduce a valid value.');

              if EditaCrit = 0 then
                AuxNumCrit := AuxNumCrit - 1;
              Exit;
            end; // AuxEscalaTipo[i] = 1
          end;

          if AuxEscalaTipo[i] = 1 then // se for crit discreto
          begin
            VerificaParA := trocapv(edtparA.Text);
            if VerificaParA < 0 then
            begin
              WebApplication.ShowMessage(
                'The function parameter "A" cannot assume the value 0. Please, introduce a valid value.');

              if EditaCrit = 0 then
                AuxNumCrit := AuxNumCrit - 1;
              Exit;
            end; // AuxEscalaTipo[i] = 1

          end
          else
          begin
            if AuxEscalaTipo[i] = 1 then // se for crit discreto
            begin
              AuxParA[i] := trocapv(edtparA.Text);
              AuxParB[i] := trocapv(edtparB.Text);
            end; // if AuxEscalaTipo[i] = 1
          end;
        end;
      end;

      { ------ Depois de passar por todas as verificações ------ }

      // Nome do critério
      AuxNomeCrit[i] := edtNameattribut.Text;
      btnmincrit.Enabled := true;

      // Adicionar o critério no listbox ou alterar o nome em caso de edição
      if EditaCrit = 0 then
        lstbxNamecriteria.Items.Add(edtNameattribut.Text);
      if EditaCrit = 1 then
        lstbxNamecriteria.Items[AuxIndice] := edtNameattribut.Text;

      if StrToInt(edtNumberattribut.Text) < (lstbxNamecriteria.Items.Count) then
      begin
        edtNumberattribut.Text := inttostr(lstbxNamecriteria.Items.Count);
      end;

      // "Inicializa" EditaCrit com 0 para a adição de um novo critério
      // Caso o usuário aperte o botão de editar esse valor se torna 1
      EditaCrit := 0;
      btnAddattribut.Caption := 'Add'; // Altera o nome do botão para ficar claro para o usuário
      lstbxNamecriteria.Enabled := true;

      { ----- Limpando ---- }
      edtNameattribut.Clear;
      cmbxScaletype.ItemIndex := -1;
      cmbxCriteriondirection.ItemIndex := -1;
      chckbxinteger.Checked := false;
      chckbxinteger.Visible := false;
      // lblNumberlevels.Visible := false;
      // edtNumberlevels.Visible := false;

      { txtparametros.Visible := false;
        edtparA.Visible := false;
        edtparB.Visible := false; }

      if cmbxScaletype.ItemIndex = 1 then
        edtNumberlevels.Clear;

      if cmbxFO.ItemIndex <> 1 then
      begin
        edtparA.Clear;
        edtparB.Clear;
      end;
    end;

    /// // Edits e labels de níveis e parâmetros visiveis ou não
    { if cmbxFO.ItemIndex <> 1 then
      begin
      txtparametros.Visible := true;
      lblparA.Visible := true;
      lblparB.Visible := true;
      edtparA.Visible := true;
      edtparB.Visible := true;
      end
      else } // alterado em 20.01.22
    begin
      txtparametros.Visible := false;
      lblparA.Visible := false;
      lblparB.Visible := false;
      edtparA.Visible := false;
      edtparB.Visible := false;
    end;

    if cmbxScaletype.ItemIndex = 1 then
    begin
      lblNumberlevels.Visible := true;
      // edtNumberlevels.Visible := true;
    end
    else
      lblNumberlevels.Visible := false;

    lstbxNamecriteria.ItemIndex := -1; // tira a seleção

    /// ///// APAGAR  (verificação) /////////////
    for j := 0 to AuxNumCrit - 1 do
    begin
      nome[j] := AuxNomeCrit[j];
      afunctipo[j] := AuxFuncTipo[j];
      aniveis[j] := AuxNiveis[j];
      aescala[j] := AuxEscalaTipo[j];
      adirec[j] := AuxDirecTipo[j];
      apara[j] := AuxParA[j];
      aparb[j] := AuxParB[j];
    end;
    /// /////////////////////
  end; // with
end;

procedure TFrmEntrada.btnauxchooseClick(Sender: TObject);
begin
  // ** Elementos para dar a sensação de carregamento **//
  btnauxchoose.Enabled := false;
  rgnLoginproject.Cursor := crWait;
  tmrcontinue.Enabled := true;
end;

procedure TFrmEntrada.btnauxstartelicitationClick(Sender: TObject);
begin
  // ** Elementos para dar a sensação de carregamento **//
  btnauxstartelicitation.Enabled := false;
  rgnImportExport.Cursor := crWait;
  tmrcont2.Enabled := true;

end;

procedure TFrmEntrada.btnauxstep1Click(Sender: TObject);
var
  i: integer;
  CMinord, Cmaxord: array of real;
begin
  // **Elementos para impressão de carregamento **//

  Setlength(CMinord, TIWUserSession(WebApplication.Data).NumCrit);
  Setlength(Cmaxord, TIWUserSession(WebApplication.Data).NumCrit);
  Setlength(TIWUserSession(WebApplication.Data).Bconseqord, TIWUserSession
      (WebApplication.Data).NumCrit);
  Setlength(TIWUserSession(WebApplication.Data).Wconseqord, TIWUserSession
      (WebApplication.Data).NumCrit);

  // procedimento para identificar o que é criterio de minimização e o que é criterio de maximizãção
  for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
  begin
    Cmaxord[i] := TIWUserSession(WebApplication.Data).MaxCons[i];
    CMinord[i] := TIWUserSession(WebApplication.Data).MinCons[i];
    if (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 1) or
      (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 3) then
    begin
      TIWUserSession(WebApplication.Data).Bconseqord[i] := Cmaxord[i];
      TIWUserSession(WebApplication.Data).Wconseqord[i] := CMinord[i];
    end; // referente a if (Typeofcrit[i] = 0) or (Typeofcrit[i] = 2)
    if (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 0) or
      (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 2) then
    begin
      TIWUserSession(WebApplication.Data).Bconseqord[i] := CMinord[i];
      TIWUserSession(WebApplication.Data).Wconseqord[i] := Cmaxord[i];
    end; // referente a if (Typeofcrit[i] = 0) or (Typeofcrit[i] = 2)
  end; // referente a for

  rgnInputMatrizConseq.Cursor := crWait;
  btnauxstep1.Cursor := crWait;
  btnauxstep2.Cursor := crWait;
  btnauxstep1.Enabled := false;
  btnauxstep2.Enabled := false;
  TIWUserSession(WebApplication.Data).stage := 1;
  TIWUserSession(WebApplication.Data).idioma := 1;
  tmrstep.Enabled := true;
end;

procedure TFrmEntrada.btnauxstep2Click(Sender: TObject);
begin
  // **Elementos para impressão de carregamento **//
  rgnInputMatrizConseq.Cursor := crWait;
  btnauxstep1.Cursor := crWait;
  btnauxstep2.Cursor := crWait;
  btnauxstep1.Enabled := false;
  btnauxstep2.Enabled := false;
  TIWUserSession(WebApplication.Data).stage := 2;
  tmrstep.Enabled := true;
  // atualizando a data de entrada no BD
  TIWUserSession(WebApplication.Data).AuxsalvarBD := 34;
  SalvarBD;
end;

procedure TFrmEntrada.gnClick(Sender: TObject);
begin
  // Abre a região de confirmação
  rgnConfirmBackEditar.Visible := true;

  btnSaveproblem.Enabled := false;
  btnSaveStart.Enabled := false;

end;

// proc
procedure TFrmEntrada.btnChooseClick(Sender: TObject);
var
  i, j: integer;
  auxnaname: string;
  // ArrayAlt, ArrayCrt: array of TIWLabel; // variáveis foram declaradas globalmente
  AuxTipo: array [0 .. 3] of string;
  mensageretorno: integer; // recebe -90998 quando não pode retornar para o local
begin
  tmrcontinue.Enabled := false;
  TIWUserSession(WebApplication.Data).stage := 0;

  // primeira verificação se nome realmente foi selecionado - alteração em 20.02.22
  if (lstbxProjects.ItemIndex = -1) then

  begin
    WebApplication.ShowMessage('Select a Project to continue');
    tmrcontinue.Enabled := false;
    btnauxchoose.Enabled := true;
    Exit;
  end;

  // captura o nome e id do problema selecionado do listbox
  TIWUserSession(WebApplication.Data).auxnome := '';
  auxnaname := lstbxProjects.Items[lstbxProjects.ItemIndex];
  for i := 1 to length(auxnaname) - 15 do
  begin
    TIWUserSession(WebApplication.Data).auxnome := TIWUserSession
      (WebApplication.Data).auxnome + auxnaname[i];
  end;
  TIWUserSession(WebApplication.Data).probname := TIWUserSession
    (WebApplication.Data).auxnome;
  { TIWUserSession(WebApplication.Data).auxnome := lstbxProjects.Items
    [lstbxProjects.ItemIndex]; }
  TIWUserSession(WebApplication.Data).ID_Problema := TIWUserSession
    (WebApplication.Data).idpoblemasuser[lstbxProjects.ItemIndex];

  TIWUserSession(WebApplication.Data).ID_chave := TIWUserSession
    (WebApplication.Data).chavesproblemas[lstbxProjects.ItemIndex];

  TIWUserSession(WebApplication.Data).dtinicretorn := TIWUserSession
    (WebApplication.Data).datainicioproblema[lstbxProjects.ItemIndex];

  TIWUserSession(WebApplication.Data).dtfnretorn := TIWUserSession
    (WebApplication.Data).datafinalproblema[lstbxProjects.ItemIndex];

  // verifica se um nome realmente foi selecionado- alteração 20.02.22
  if TIWUserSession(WebApplication.Data).auxnome = '' then
  begin
    WebApplication.ShowMessage('Select a Project to continue');
    tmrcontinue.Enabled := false;
    btnauxchoose.Enabled := true;
    Exit;
  end;

  // procedimento para com o nome do projeto, e o usuario pegar o id_problema, a problematica e o status
  TIWUserSession(WebApplication.Data).AuxsalvarBD := 14;
  SalvarBD;

  if TIWUserSession(WebApplication.Data).Problematica = 3 then
  begin
    TIWUserSession(WebApplication.Data).elicitarperfis := false;
    // ** Verifica se o problema foi reiniciado antes de ser retomado **//

    TIWUserSession(WebApplication.Data).AuxsalvarBD := 37;
    SalvarBD;
  end;

  // ALTERAÇÃO EM 20.02.22 para retomada manual
  // só entra nos cases referentes à intra, quando não for salvamento parcial da entrada manual - alteração em 21.01.22
  if (TIWUserSession(WebApplication.Data).estagio <> 6) AND
    (TIWUserSession(WebApplication.Data).estagio <> -1) then
  begin
    { após pegar id do problema, verifica round, alternativas e criterios que está, do problema selelcionado }
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 38;
    SalvarBD;

    if TIWUserSession(WebApplication.Data).Problematica = 4 then
    begin
      // ** Busca as informações a respeito dos portfólios **//
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 44;
      SalvarBD;
    end;
    // ** Se os portfólios ainda estiverem sendo gerados, leva para a tela de carregamento **//
    if TIWUserSession(WebApplication.Data).PortfolioG then
    begin
      rgnLoginproject.Hide;
      tmrPortfoliocombination.Enabled := true;
      rgnPortGeneration.show;
      Exit;
    end;
    { else
      begin }
    // ** Se os portfólios já tiverem sido gerados, apenas se resgata seus valores **//
    // ** Busca as informações a respeito dos portfólios **//
    if TIWUserSession(WebApplication.Data).Problematica = 4 then
    begin
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 43;
      SalvarBD;
    end;

    { e depois, se o id_problem retomado é de um problema que passou pela elicitação intra, identifica
      rounds intra que o problema foi interrompido, de acordo com o round do sistema }
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 39;
    SalvarBD;

    { retornando todos oa valores referentes às infos obtidas nos cases 37 e 38. infos que foram elicitadas
      são resgatadas para fazer o preenchimento correto da matriz }
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 40;
    SalvarBD;

    { retornando infos da ultima atualização feita no vtr DinamVerificaIntra. }
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 41;
    SalvarBD;
  end;
  // fim alteração 20.02

  if (StrToDate(TIWUserSession(WebApplication.Data).dtinicretorn) < StrToDate
      ('19/04/2021')) AND (StrToDate(TIWUserSession(WebApplication.Data)
        .dtfnretorn) >= StrToDate('19/04/2021')) then
  begin
    if (TIWUserSession(WebApplication.Data).estagio = 0) or
      (TIWUserSession(WebApplication.Data).estagio = 1) then
    begin
      TIWUserSession(WebApplication.Data).tempinicio := TimeToStr(Now);

      // procedimento para pegar o round, caracteristicas dos criterios e das alternativas
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 15;
      SalvarBD;
      // rgnInputMatrizConseq.Visible := true;
      // rgnLoginproject.Visible := false;
      // **Deixa visível a opção para seguir pra elicitação direto **//
      AlreadyRakend1.Visible := true;

      // procedimento para preencher as informações na região rgnInputMatrizConseq

      Inicializa;
      if TIWUserSession(WebApplication.Data).Flag <> -2 then
      begin
        with TIWUserSession(WebApplication.Data) do
        begin
          /// / Setando o tamanho dos grids que exibem a informação importada na região imput ////
          grdMatrizConseq.RowCount := NumAlt + 7;
          grdMatrizConseq.ColumnCount := NumCrit + 1;
          probname := edtEntername.Text;

          /// / Preenchendo a informação fixa dos grids ////
          grdMatrizConseq.Cell[0, 0].Text := 'Criteria:';
          grdMatrizConseq.Cell[1, 0].Text :=
            '0-Cont Min; 1-Cont Max; 2-Disc Min; 3- Disc Max; 4- Int Min; 5- Int Max:';
          { grdInputvalues.Cell[2, 0].text := 'Weights:'; }
          grdMatrizConseq.Cell[2, 0].Text := '';
          grdMatrizConseq.Cell[3, 0].Text := '';
          grdMatrizConseq.Cell[4, 0].Text := '';
          grdMatrizConseq.Cell[5, 0].Text :=
            'Number of levels of discrete criteria';
          grdMatrizConseq.Cell[6, 0].Text := 'Consequence Matrix:';

          /// / Preenchendo os Grids com a informação Imputada ///////////////////////
          for j := 0 to NumCrit - 1 do
          begin
            grdMatrizConseq.Cell[0, j + 1].Text := NomeCrit[j];
            if Cri_inteiro[j] then
            begin
              if Typeofcrit[j] = 0 then
                grdMatrizConseq.Cell[1, j + 1].Text := '4';
              if Typeofcrit[j] = 1 then
                grdMatrizConseq.Cell[1, j + 1].Text := '5';
            end
            else
              grdMatrizConseq.Cell[1, j + 1].Text := inttostr(Typeofcrit[j]);
            { grdMatrizConseq.Cell[2, j + 1].Text := inttostr(FTipo[j]);
              if FTipo[j] <> 1 then
              begin
              grdMatrizConseq.Cell[3, j + 1].Text := FloatToStr(parA[j]);
              grdMatrizConseq.Cell[4, j + 1].Text := FloatToStr(parB[j]);
              end; }
            grdMatrizConseq.Cell[5, j + 1].Text := inttostr(Niveis[j]);
          end;

          for i := 0 to NumAlt - 1 do
          begin
            grdMatrizConseq.Cell[i + 7, 0].Text := NomeAlt[i];
            for j := 0 to NumCrit - 1 do
            begin
              grdMatrizConseq.Cell[i + 7, j + 1].Text := TrocaVP
                (ConsMatrix[i, j]);
            end;
          end;
          /// ////////////////////////////////////////////////////////////////////////
          // end;
        end; // with

        // após recuperar todas as informações necessárias para retomar um problema, verifica em qual estagiointra
        // a elicitação foi interrompida, redirecionando paraa tela apropriada
        if (TIWUserSession(WebApplication.Data).Estagiointra = 0) { OR
          (TIWUserSession(WebApplication.Data).Estagiointra = 1) } then
        begin
          TIWUserSession(WebApplication.Data).retornointra := 0; // var global que auxilia na determinação de qual tela intra o sis retomará
          //TTFrmIntraCriterio.Create(WebApplication).show;
          rgnLoginproject.Visible := false;
        end; // estagiointra 0

        { ALTERAÇÃO 20.02.22 }
        if (TIWUserSession(WebApplication.Data).Estagiointra = 1) then
        begin
          TIWUserSession(WebApplication.Data).retornointra := 1;
          //TTFrmIntraCriterio.Create(WebApplication).show;
          rgnLoginproject.Visible := false;

          for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
            TIWUserSession(WebApplication.Data).FTipo[i] := TIWUserSession
              (WebApplication.Data).NewFtipo[i];

        end; // estagiointra 1  - FIM ALTERAÇÃO

        if (TIWUserSession(WebApplication.Data).Estagiointra = 2) OR
          (TIWUserSession(WebApplication.Data).Estagiointra = 3) then
        begin
          rgnLoginproject.Visible := false;
          rgnInputMatrizConseq.Visible := true;
        end; // estagiointra 2 ou 3

      end; // flag
    end; // if estagio

    if (TIWUserSession(WebApplication.Data).estagio = 2) or
      (TIWUserSession(WebApplication.Data).estagio = 3) then
    begin
      // procedimento para pegar o round, caracteristicas dos criterios e das alternativas
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 15;
      SalvarBD;
      // procedimento para retornar dados dos criterios
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 16;
      SalvarBD;

      // procedimento para retornar dados da AH
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 19;
      SalvarBD;

      // procedimento para retornar dados da elicitação
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 18;
      SalvarBD;

      // procedimento para retornar dados da elicitação
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 21;
      SalvarBD;

      // procedimento para retornar dados da avaliação holistica
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 29;
      SalvarBD;

      /// ////////////////////////////////////////////////////////////////////////////////////////
      // preenchimento de variaveis que vão depender da resposta da heuristica e de outros fatores
      if (TIWUserSession(WebApplication.Data).HeurQuest = 0) or
        (TIWUserSession(WebApplication.Data).HeurQuest = 2) then
      begin
        // se a resposta da heuristica for A ou indiferente
        TIWUserSession(WebApplication.Data).Etapa := 2;
      end; // referente a  if (TIWUserSession(WebApplication.Data).HeurQuest=0) or (TIWUserSession(WebApplication.Data).HeurQuest=2) then

      if TIWUserSession(WebApplication.Data).HeurQuest = 1 then
      begin
        // se a resposta da heuristica for igual a B

        if TIWUserSession(WebApplication.Data).NumQuest =
          (TIWUserSession(WebApplication.Data).NumCrit - TIWUserSession
            (WebApplication.Data).NumCriBlind) then
        begin
          TIWUserSession(WebApplication.Data).Etapa := 2;
        end; // referente a if TIWUserSession(WebApplication.Data).NumQuest=(TIWUserSession(WebApplication.Data).NumCrit-TIWUserSession(WebApplication.Data).NumCriBlind) then

        if TIWUserSession(WebApplication.Data).NumQuest <
          (Round(0.5 * (TIWUserSession(WebApplication.Data)
                .NumCrit - TIWUserSession(WebApplication.Data).NumCriBlind)))
          then
        begin
          TIWUserSession(WebApplication.Data).Etapa := 1;
          TIWUserSession(WebApplication.Data).Fase := 1;
        end // referente a if TIWUserSession(WebApplication.Data).NumQuest < (round(0.5*(TIWUserSession(WebApplication.Data).NumCrit - TIWUserSession(WebApplication.Data).NumCriBlind)))then
        else
        begin
          TIWUserSession(WebApplication.Data).Etapa := 1;
          TIWUserSession(WebApplication.Data).Fase := 2;
        end;
      end; // referente a if TIWUserSession(WebApplication.Data).HeurQuest=1 then

      // ** devemos retomar as informações dos limites das relações
      // ** entre os crit não adj e quais foram blindadas **//

      TIWUserSession(WebApplication.Data).AuxsalvarBD := 45;
      SalvarBD;



      /// ///////////////////////////////////////////////////////////////////////////////////////////////////

      { if TIWUserSession(WebApplication.Data).auxparacasonsaieretone=-1002525 then
        TIWUserSession(WebApplication.Data).stage:=0; }

      TIWUserSession(WebApplication.Data).RETPROBDEPOIS := -4555;
      Inicializa;
      // TIWUserSession(WebApplication.Data).stage := 3;
      TIWUserSession(WebApplication.Data).epsilon := trocapv(edtMaxdif.Text);
      rgnLoginproject.Visible := false;



      // Verificação para passar ou não direto para a unit de resultados
      if (TIWUserSession(WebApplication.Data).elicitarperfis = true) and
        (TIWUserSession(WebApplication.Data).Problematica = 3) then
      begin
        Tfrmaintercriterio.Create(WebApplication).btnprofileelicitClick(Self);
      end
      else
      begin
        TIWUserSession(WebApplication.Data).elicitarperfis := false;
      Tfrmaintercriterio.Create(WebApplication).show;

      // Verificação para passar ou não direto para a unit de resultados (caso não seja necessário realizar a elicitção flexível)
      { if TIWUserSession(WebApplication.Data).Sol <> -5 then
        begin }
      Tfrmaintercriterio.Create(WebApplication).BtnshowresultsClick(Self);
      // end;

      end;

    end; // referente a if (TIWUserSession(WebApplication.Data).estagio = 2) or (TIWUserSession(WebApplication.Data).estagio = 3) then

    if TIWUserSession(WebApplication.Data).estagio = -1 then
    begin
      rgnLoginproject.Visible := false;
      rgnConseqvalues.Visible := true;
      TIWUserSession(WebApplication.Data).Nomeprojeto := TIWUserSession
        (WebApplication.Data).auxnome;
      // procedimento para retornar dados da elicitação
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 25;
      SalvarBD;

      with TIWUserSession(WebApplication.Data) do
      begin
        Setlength(ArrayAlt, TIWUserSession(WebApplication.Data).NumAlt);
        for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
        begin
          ArrayAlt[i] := TIWLabel.Create(Self);
          ArrayAlt[i].parent := rgnconsman;
          ArrayAlt[i].Text := ('Alt' + inttostr(i + 1));
          ArrayAlt[i].Left := (5);
          ArrayAlt[i].Top := (50 + (i * 25));
          ArrayAlt[i].Font.Style := [fsbold];
          ArrayAlt[i].Font.FontName := 'Verdana';
        end;

        Setlength(ArrayCrt, TIWUserSession(WebApplication.Data).NumCrit);
        for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
        begin
          ArrayCrt[i] := TIWLabel.Create(Self);
          ArrayCrt[i].parent := rgnconsman;
          ArrayCrt[i].Text := ('C' + inttostr(i + 1));
          ArrayCrt[i].Left := (45 + (i * 60));
          ArrayCrt[i].Top := (20);
          ArrayCrt[i].Font.Style := [fsbold];
          ArrayCrt[i].Font.FontName := 'Verdana';
        end;
      end;

      Setlength(TIWUserSession(WebApplication.Data).MatEdit, TIWUserSession
          (WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data)
          .NumCrit);
      // alteração feita em 20.02 para permitir que retomada manual com info parcial ocorra corretamente
      Setlength(TIWUserSession(WebApplication.Data).Verificaintra,
        TIWUserSession(WebApplication.Data).NumCrit);

      for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
      begin
        for j := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do

        begin
          // Criando os edits que recaberão os desempenhos a partir de uma Matriz de TIWEdit's
          TIWUserSession(WebApplication.Data).MatEdit[i][j] := TIWEdit.Create
            (Self);
          TIWUserSession(WebApplication.Data).MatEdit[i][j].parent :=
            rgnconsman;
          TIWUserSession(WebApplication.Data).MatEdit[i][j].Text := ('');
          TIWUserSession(WebApplication.Data).MatEdit[i][j].Left :=
            (45 + (j * 60));
          TIWUserSession(WebApplication.Data).MatEdit[i][j].Top :=
            (25 + (i * 25) + 23);
          TIWUserSession(WebApplication.Data).MatEdit[i][j].Width := 60;
        end;
      end;
      // cmproblematica1.ItemIndex := -1;  // Não limpo o nome caso o usuário volte para essa página

      AuxTipo[0] := 'Linear';
      AuxTipo[1] := 'Exponential';
      AuxTipo[2] := 'Logarithmic';
      AuxTipo[3] := 'Logistic';
      // ** Preenchendo a tabela de informações dos critérios **//

      grdCritInfo.ColumnCount := 7;
      grdCritInfo.RowCount := TIWUserSession(WebApplication.Data).NumCrit + 1;

      grdCritInfo.Cell[0, 0].Text := 'Name';
      grdCritInfo.Cell[0, 1].Text := 'Scale';
      grdCritInfo.Cell[0, 2].Text := 'Number of levels';
      grdCritInfo.Cell[0, 3].Text := 'Value Function';
      grdCritInfo.Cell[0, 4].Text := 'Direction';
      grdCritInfo.Cell[0, 5].Text := 'Par a';
      grdCritInfo.Cell[0, 6].Text := 'Par b';

      with TIWUserSession(WebApplication.Data) do // ALTERAÇÃO EM 20.02.22
        for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
        begin
          grdCritInfo.Cell[i + 1, 0].Text := TIWUserSession(WebApplication.Data)
            .NomeCrit[i];
          if TIWUserSession(WebApplication.Data).Typeofcrit[i] > 1 then
          begin
            grdCritInfo.Cell[i + 1, 1].Text := 'Discrete';
            grdCritInfo.Cell[i + 1, 2].Text := inttostr
              (TIWUserSession(WebApplication.Data).Niveis[i]);
          end // end do   if TIWUserSession(WebApplication.Data).Typeofcrit[I] > 1
          else
          begin
            grdCritInfo.Cell[i + 1, 1].Text := 'Continuous';
            grdCritInfo.Cell[i + 1, 2].Text := '---';
          end; // end do else

          { ALTERAÇÃO EM 20.02.22 }
          // quando Ftipo=0 (elic intra), escreve direto na célula do grid. Sem utilizar AuxTipo
          if FTipo[i] = 0 then
          begin
            grdCritInfo.Cell[i + 1, 4].Text := '---';
            TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
          end
          else
            grdCritInfo.Cell[i + 1, 3].Text := AuxTipo
              [TIWUserSession(WebApplication.Data).FTipo[i] - 1];

          /// //COLOCANDO TYPECRIT SEPARADO, DEVIDO QUESTÃO DE EDTS VAZIOS//////
          if (Typeofcrit[i] = 1) then // CONT MAX
          begin
            grdCritInfo.Cell[i + 1, 5].Text := 'Maximization';

            if (edtparA.Text = '') then
            begin
              grdCritInfo.Cell[i + 1, 6].Text := '---';
              if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
            end;

            if (edtparB.Text = '') then
            begin
              grdCritInfo.Cell[i + 1, 7].Text := '---';
              if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
            end // se vazio

            else
            begin
              grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parA[i]);
              grdCritInfo.Cell[i + 1, 7].Text := FloatToStr(parB[i]);
              if (parA[i] = 0) then
                grdCritInfo.Cell[i + 1, 6].Text := '---';
              if (parB[i] = 0) then
                grdCritInfo.Cell[i + 1, 7].Text := '---';
            end; // se possui valor
          end; // type=1

          if (Typeofcrit[i] = 3) then // DISC MAX
          begin
            grdCritInfo.Cell[i + 1, 5].Text := 'Maximization';
            grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parA[i]);
            grdCritInfo.Cell[i + 1, 7].Text := FloatToStr(parB[i]);
            if (parA[i] = 0) then
              grdCritInfo.Cell[i + 1, 6].Text := '---';
            if (parB[i] = 0) then
              grdCritInfo.Cell[i + 1, 7].Text := '---';
          end
          else

            if (Typeofcrit[i] = 0) then // CONT MIN
          begin
            grdCritInfo.Cell[i + 1, 5].Text := 'Minimization';

            if (edtparA.Text = '') then
            begin
              grdCritInfo.Cell[i + 1, 6].Text := '---';
              if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
            end;

            if (edtparB.Text = '') then
            begin
              grdCritInfo.Cell[i + 1, 7].Text := '---';
              if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
            end // se vazio

            else
            begin
              grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parA[i]);
              grdCritInfo.Cell[i + 1, 7].Text := FloatToStr(parB[i]);
              if (parA[i] = 0) then
                grdCritInfo.Cell[i + 1, 6].Text := '---';
              if (parB[i] = 0) then
                grdCritInfo.Cell[i + 1, 7].Text := '---';
            end; // se possui valor
          end; // type= 0

          if (Typeofcrit[i] = 2) then // DISC MIN
          begin
            grdCritInfo.Cell[i + 1, 5].Text := 'Minimization';
            grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parA[i]);
            grdCritInfo.Cell[i + 1, 7].Text := FloatToStr(parB[i]);
            if (parA[i] = 0) then
              grdCritInfo.Cell[i + 1, 6].Text := '---';
            if (parB[i] = 0) then
              grdCritInfo.Cell[i + 1, 7].Text := '---';
          end;
        end;
      // FINAL ALTERAÇÃO EM 20.02.22

      // criando o grid com as alternativas
      grdinformaalt.ColumnCount := 2;
      grdinformaalt.RowCount := TIWUserSession(WebApplication.Data).NumAlt;
      for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
      begin
        grdinformaalt.Cell[i, 0].Text := 'Alt' + ' ' + inttostr(i + 1);
        grdinformaalt.Cell[i, 1].Text := TIWUserSession(WebApplication.Data)
          .NomeAlt[i];
      end; // referente a for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do

    end; // end do -1

    if TIWUserSession(WebApplication.Data).estagio = 6 then
    begin
      rgnLoginproject.Visible := false;
      rgnConseqvalues.Visible := true;
      TIWUserSession(WebApplication.Data).Nomeprojeto := TIWUserSession
        (WebApplication.Data).auxnome;
      // procedimento para retornar dados da elicitação
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 31;
      SalvarBD;

      with TIWUserSession(WebApplication.Data) do
      begin
        Setlength(ArrayAlt, TIWUserSession(WebApplication.Data).NumAlt);
        for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
        begin
          ArrayAlt[i] := TIWLabel.Create(Self);
          ArrayAlt[i].parent := rgnconsman;
          ArrayAlt[i].Text := ('Alt' + inttostr(i + 1));
          ArrayAlt[i].Left := (5);
          ArrayAlt[i].Top := (50 + (i * 25));
          ArrayAlt[i].Font.Style := [fsbold];
          ArrayAlt[i].Font.FontName := 'Verdana';
        end;

        Setlength(ArrayCrt, TIWUserSession(WebApplication.Data).NumCrit);
        for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
        begin
          ArrayCrt[i] := TIWLabel.Create(Self);
          ArrayCrt[i].parent := rgnconsman;
          ArrayCrt[i].Text := ('C' + inttostr(i + 1));
          ArrayCrt[i].Left := (45 + (i * 60));
          ArrayCrt[i].Top := (20);
          ArrayCrt[i].Font.Style := [fsbold];
          ArrayCrt[i].Font.FontName := 'Verdana';
        end;
      end;

      Setlength(TIWUserSession(WebApplication.Data).MatEdit, TIWUserSession
          (WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data)
          .NumCrit);
      // alteração feita em 21.01.22 para permitir que retomada manual com info parcial ocorra corretamente
      Setlength(TIWUserSession(WebApplication.Data).Verificaintra,
        TIWUserSession(WebApplication.Data).NumCrit);

      for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
      begin
        for j := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do

        begin
          // Criando os edits que recaberão os desempenhos a partir de uma Matriz de TIWEdit's
          TIWUserSession(WebApplication.Data).MatEdit[i][j] := TIWEdit.Create
            (Self);
          TIWUserSession(WebApplication.Data).MatEdit[i][j].parent :=
            rgnconsman;
          if TIWUserSession(WebApplication.Data).ConsMatrix[i, j] <> 0 then
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Text := FloatToStr
              (TIWUserSession(WebApplication.Data).ConsMatrix[i, j]);
          if TIWUserSession(WebApplication.Data).ConsMatrix[i, j] = 0 then
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Text := '';
          TIWUserSession(WebApplication.Data).MatEdit[i][j].Left :=
            (45 + (j * 60));
          TIWUserSession(WebApplication.Data).MatEdit[i][j].Top :=
            (25 + (i * 25) + 23);
          TIWUserSession(WebApplication.Data).MatEdit[i][j].Width := 60;
        end;
      end;
      // cmproblematica1.ItemIndex := -1;  // Não limpo o nome caso o usuário volte para essa página

      AuxTipo[0] := 'Linear';
      AuxTipo[1] := 'Exponential';
      AuxTipo[2] := 'Logarithmic';
      AuxTipo[3] := 'Logistic';


      // Alterado em 21.01.22//
      // ** Preenchendo a tabela de informações dos critérios **//

      grdCritInfo.ColumnCount := 7;
      grdCritInfo.RowCount := TIWUserSession(WebApplication.Data).NumCrit + 1;

      grdCritInfo.Cell[0, 0].Text := 'Name';
      grdCritInfo.Cell[0, 1].Text := 'Scale';
      grdCritInfo.Cell[0, 2].Text := 'Number of levels';
      grdCritInfo.Cell[0, 3].Text := 'Value Function';
      grdCritInfo.Cell[0, 4].Text := 'Direction';
      grdCritInfo.Cell[0, 5].Text := 'Par a';
      grdCritInfo.Cell[0, 6].Text := 'Par b';

      with TIWUserSession(WebApplication.Data) do
        for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
        begin
          grdCritInfo.Cell[i + 1, 0].Text := TIWUserSession(WebApplication.Data)
            .NomeCrit[i];
          if TIWUserSession(WebApplication.Data).Typeofcrit[i] > 1 then
          begin
            grdCritInfo.Cell[i + 1, 1].Text := 'Discrete';
            grdCritInfo.Cell[i + 1, 2].Text := inttostr
              (TIWUserSession(WebApplication.Data).Niveis[i]);
          end // end do   if TIWUserSession(WebApplication.Data).Typeofcrit[I] > 1
          else
          begin
            grdCritInfo.Cell[i + 1, 1].Text := 'Continuous';
            grdCritInfo.Cell[i + 1, 2].Text := '---';
          end; // end do else

          // quando Ftipo=0 (elic intra), escreve direto na célula do grid. Sem utilizar AuxTipo
          if FTipo[i] = 0 then
          begin
            grdCritInfo.Cell[i + 1, 3].Text := '---';
            TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
          end
          else
            grdCritInfo.Cell[i + 1, 3].Text := AuxTipo
              [TIWUserSession(WebApplication.Data).FTipo[i] - 1];

          /// //COLOCANDO TYPECRIT SEPARADO, DEVIDO QUESTÃO DE EDTS VAZIOS//////
          if (Typeofcrit[i] = 1) then // CONT MAX
          begin
            grdCritInfo.Cell[i + 1, 4].Text := 'Maximization';

            if (edtparA.Text = '') then
            begin
              grdCritInfo.Cell[i + 1, 5].Text := '---';
              if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
            end;

            if (edtparB.Text = '') then
            begin
              grdCritInfo.Cell[i + 1, 6].Text := '---';
              if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
            end // se vazio

            else
            begin
              grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
              grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
              if (parA[i] = 0) then
                grdCritInfo.Cell[i + 1, 5].Text := '---';
              if (parB[i] = 0) then
                grdCritInfo.Cell[i + 1, 6].Text := '---';
            end; // se possui valor
          end; // type=1

          if (Typeofcrit[i] = 3) then // DISC MAX
          begin
            grdCritInfo.Cell[i + 1, 4].Text := 'Maximization';
            grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
            grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
            if (parA[i] = 0) then
              grdCritInfo.Cell[i + 1, 5].Text := '---';
            if (parB[i] = 0) then
              grdCritInfo.Cell[i + 1, 6].Text := '---';
          end
          else

            if (Typeofcrit[i] = 0) then // CONT MIN
          begin
            grdCritInfo.Cell[i + 1, 4].Text := 'Minimization';

            if (edtparA.Text = '') then
            begin
              grdCritInfo.Cell[i + 1, 5].Text := '---';
              if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
            end;

            if (edtparB.Text = '') then
            begin
              grdCritInfo.Cell[i + 1, 6].Text := '---';
              if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
            end // se vazio

            else
            begin
              grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
              grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
              if (parA[i] = 0) then
                grdCritInfo.Cell[i + 1, 5].Text := '---';
              if (parB[i] = 0) then
                grdCritInfo.Cell[i + 1, 6].Text := '---';
            end; // se possui valor
          end; // type= 0

          if (Typeofcrit[i] = 2) then // DISC MIN
          begin
            grdCritInfo.Cell[i + 1, 4].Text := 'Minimization';
            grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
            grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
            if (parA[i] = 0) then
              grdCritInfo.Cell[i + 1, 5].Text := '---';
            if (parB[i] = 0) then
              grdCritInfo.Cell[i + 1, 6].Text := '---';
          end;
        end;

      { grdCritInfo.Cell[i + 1, 3].Text := AuxTipo
        [TIWUserSession(WebApplication.Data).FTipo[i] - 1];

        if (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 1) OR
        (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 3) then
        begin
        grdCritInfo.Cell[i + 1, 4].Text := 'Maximization';
        grdCritInfo.Cell[i + 1, 5].Text := FloatToStr
        (TIWUserSession(WebApplication.Data).parA[i]);
        grdCritInfo.Cell[i + 1, 6].Text := FloatToStr
        (TIWUserSession(WebApplication.Data).parB[i]);
        if (TIWUserSession(WebApplication.Data).parA[i] = 0) then
        grdCritInfo.Cell[i + 1, 5].Text := '---';
        if (TIWUserSession(WebApplication.Data).parB[i] = 0) then
        grdCritInfo.Cell[i + 1, 6].Text := '---';
        end // end do  if (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 1)...
        else if (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 0) OR
        (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 2) then
        begin
        grdCritInfo.Cell[i + 1, 4].Text := 'Minimization';
        grdCritInfo.Cell[i + 1, 5].Text := FloatToStr
        (TIWUserSession(WebApplication.Data).parA[i]);
        grdCritInfo.Cell[i + 1, 6].Text := FloatToStr
        (TIWUserSession(WebApplication.Data).parB[i]);
        if (TIWUserSession(WebApplication.Data).parA[i] = 0) then
        grdCritInfo.Cell[i + 1, 5].Text := '---';
        if (TIWUserSession(WebApplication.Data).parB[i] = 0) then
        grdCritInfo.Cell[i + 1, 6].Text := '---';
        end; // end do else
        end; // end do for I := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
        } // fim da alteração 21.01.22

      // criando o grid com as alternativas
      grdinformaalt.ColumnCount := 2;
      grdinformaalt.RowCount := TIWUserSession(WebApplication.Data).NumAlt;
      for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
      begin
        grdinformaalt.Cell[i, 0].Text := 'Alt' + ' ' + inttostr(i + 1);
        grdinformaalt.Cell[i, 1].Text := TIWUserSession(WebApplication.Data)
          .NomeAlt[i];
      end; // referente a for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do

    end; // end do 6

    if (TIWUserSession(WebApplication.Data).estagio = 4) or
      (TIWUserSession(WebApplication.Data).estagio = 5) then
    begin
      // procedimento para pegar o round, caracteristicas dos criterios e das alternativas
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 15;
      SalvarBD;
      // procedimento para retornar dados dos criterios
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 16;
      SalvarBD;

      // procedimento para retornar dados da AH
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 19;
      SalvarBD;

      // procedimento para retornar dados da elicitação
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 18;
      SalvarBD;

      // procedimento para retornar dados da elicitação
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 21;
      SalvarBD;

      // procedimento para retornar dados da avaliação holistica
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 29;
      SalvarBD;

      /// ////////////////////////////////////////////////////////////////////////////////////////
      // preenchimento de variaveis que vão depender da resposta da heuristica e de outros fatores
      if (TIWUserSession(WebApplication.Data).HeurQuest = 0) or
        (TIWUserSession(WebApplication.Data).HeurQuest = 2) then
      begin
        // se a resposta da heuristica for A ou indiferente
        TIWUserSession(WebApplication.Data).Etapa := 2;
      end; // referente a  if (TIWUserSession(WebApplication.Data).HeurQuest=0) or (TIWUserSession(WebApplication.Data).HeurQuest=2) then

      if TIWUserSession(WebApplication.Data).HeurQuest = 1 then
      begin
        // se a resposta da heuristica for igual a B

        if TIWUserSession(WebApplication.Data).NumQuest =
          (TIWUserSession(WebApplication.Data).NumCrit - TIWUserSession
            (WebApplication.Data).NumCriBlind) then
        begin
          TIWUserSession(WebApplication.Data).Etapa := 2;
        end; // referente a if TIWUserSession(WebApplication.Data).NumQuest=(TIWUserSession(WebApplication.Data).NumCrit-TIWUserSession(WebApplication.Data).NumCriBlind) then

        if TIWUserSession(WebApplication.Data).NumQuest <
          (Round(0.5 * (TIWUserSession(WebApplication.Data)
                .NumCrit - TIWUserSession(WebApplication.Data).NumCriBlind)))
          then
        begin
          TIWUserSession(WebApplication.Data).Etapa := 1;
          TIWUserSession(WebApplication.Data).Fase := 1;
        end // referente a if TIWUserSession(WebApplication.Data).NumQuest < (round(0.5*(TIWUserSession(WebApplication.Data).NumCrit - TIWUserSession(WebApplication.Data).NumCriBlind)))then
        else
        begin
          TIWUserSession(WebApplication.Data).Etapa := 1;
          TIWUserSession(WebApplication.Data).Fase := 2;
        end;
      end; // referente a if TIWUserSession(WebApplication.Data).HeurQuest=1 then

      // ** devemos retomar as informações dos limites das relações
      // ** entre os crit não adj e quais foram blindadas **//

      TIWUserSession(WebApplication.Data).AuxsalvarBD := 45;
      SalvarBD;
      /// ///////////////////////////////////////////////////////////////////////////////////////////////////

      { if TIWUserSession(WebApplication.Data).auxparacasonsaieretone=-1002525 then
        TIWUserSession(WebApplication.Data).stage:=0; }

      TIWUserSession(WebApplication.Data).RETPROBDEPOIS := -4555;
      Inicializa;
      // TIWUserSession(WebApplication.Data).stage := 3;
      TIWUserSession(WebApplication.Data).epsilon := trocapv(edtMaxdif.Text);
      rgnLoginproject.Visible := false;



      // Verificação para passar ou não direto para a unit de resultados
      if (TIWUserSession(WebApplication.Data).elicitarperfis = true) and
        (TIWUserSession(WebApplication.Data).Problematica = 3) then
      begin
        Tfrmaintercriterio.Create(WebApplication).btnprofileelicitClick(Self);
      end
      else
      begin
        Tfrmaintercriterio.Create(WebApplication).show;

      // Verificação para passar ou não direto para a unit de resultados
      { if TIWUserSession(WebApplication.Data).Sol <> -5 then
        begin }
      TIWUserSession(WebApplication.Data).RETPROBDEPOIS := -4555;
      Tfrmaintercriterio.Create(WebApplication).BtnshowresultsClick(Self);
      end;
      // end;

      // Verificação para passar ou não direto para a unit de resultados
      if TIWUserSession(WebApplication.Data).Sol = -5 then
      begin
        TIWUserSession(WebApplication.Data).RETPROBDEPOIS := -4555;
        Tfrmaintercriterio.Create(WebApplication).btnstopelicitationClick(Self);
      end;

    end; // referente a if (TIWUserSession(WebApplication.Data).estagio = 5) or (TIWUserSession(WebApplication.Data).estagio = 4) then

  end; // referente a if datafinalproblema='19/04/2021' then

  // Problema iniciado depois de 19/04/2021
  if (StrToDate(TIWUserSession(WebApplication.Data).dtinicretorn) = StrToDate
      ('19/04/2021')) or (StrToDate(TIWUserSession(WebApplication.Data)
        .dtinicretorn) > StrToDate('19/04/2021')) then
  begin
    if (TIWUserSession(WebApplication.Data).estagio = 0) or
      (TIWUserSession(WebApplication.Data).estagio = 1) then
    begin
      TIWUserSession(WebApplication.Data).tempinicio := TimeToStr(Now);

      // procedimento para pegar o round, caracteristicas dos criterios e das alternativas
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 15;
      SalvarBD;

      // ** Busca as informações a respeito dos portfólios **//
      if TIWUserSession(WebApplication.Data).Problematica = 4 then
      begin
        TIWUserSession(WebApplication.Data).AuxsalvarBD := 44;
        SalvarBD;
      end;

      // ** Se os portfólios ainda estiverem sendo gerados, leva para a tela de carregamento **//
      if TIWUserSession(WebApplication.Data).PortfolioG then
      begin
        rgnLoginproject.Hide;
        tmrPortfoliocombination.Enabled := true;
        rgnPortGeneration.show;
        Exit;
      end;
      { else
        begin }
      // ** Se os portfólios já tiverem sido gerados, apenas se resgata seus valores **//
      // ** Busca as informações a respeito dos portfólios **//
      if TIWUserSession(WebApplication.Data).Problematica = 4 then
      begin
        TIWUserSession(WebApplication.Data).AuxsalvarBD := 43;
        SalvarBD;
      end;
      // rgnInputMatrizConseq.Visible := true;
      // rgnLoginproject.Visible := false;
      // **Deixa visível a opção para seguir pra elicitação direto **//
      AlreadyRakend1.Visible := true;


      // procedimento para preencher as informações na região rgnInputMatrizConseq

      Inicializa;
      if TIWUserSession(WebApplication.Data).Flag <> -2 then
      begin
        with TIWUserSession(WebApplication.Data) do
        begin
          /// / Setando o tamanho dos grids que exibem a informação importada na região imput ////
          grdMatrizConseq.RowCount := NumAlt + 7;
          grdMatrizConseq.ColumnCount := NumCrit + 1;
          if Problematica = 4 then
            grdMatrizConseq.ColumnCount := NumCrit + 2;
          probname := edtEntername.Text;

          /// / Preenchendo a informação fixa dos grids ////
          grdMatrizConseq.Cell[0, 0].Text := 'Criteria:';
          grdMatrizConseq.Cell[1, 0].Text :=
            '0-Cont Min; 1-Cont Max; 2-Disc Min; 3- Disc Max; 4- Int Min; 5- Int Max:';
          { grdInputvalues.Cell[2, 0].text := 'Weights:'; }
          grdMatrizConseq.Cell[2, 0].Text := '';
          grdMatrizConseq.Cell[3, 0].Text := '';
          grdMatrizConseq.Cell[4, 0].Text := '';
          grdMatrizConseq.Cell[5, 0].Text :=
            'Number of levels of discrete criteria';
          grdMatrizConseq.Cell[6, 0].Text := 'Consequence Matrix:';

          /// / Preenchendo os Grids com a informação Imputada ///////////////////////
          for j := 0 to NumCrit - 1 do
          begin
            grdMatrizConseq.Cell[0, j + 1].Text := NomeCrit[j];
            if Cri_inteiro[j] then
            begin
              if Typeofcrit[j] = 0 then
                grdMatrizConseq.Cell[1, j + 1].Text := '4';
              if Typeofcrit[j] = 1 then
                grdMatrizConseq.Cell[1, j + 1].Text := '5';
            end
            else
              grdMatrizConseq.Cell[1, j + 1].Text := inttostr(Typeofcrit[j]);
            { grdMatrizConseq.Cell[2, j + 1].Text := inttostr(FTipo[j]);
              if FTipo[j] <> 1 then
              begin
              grdMatrizConseq.Cell[3, j + 1].Text := FloatToStr(parA[j]);
              grdMatrizConseq.Cell[4, j + 1].Text := FloatToStr(parB[j]);
              end; }
            grdMatrizConseq.Cell[5, j + 1].Text := inttostr(Niveis[j]);
          end;

          for i := 0 to NumAlt - 1 do
          begin
            grdMatrizConseq.Cell[i + 7, 0].Text := TIWUserSession
              (WebApplication.Data).NomeAlt[i];
            if Problematica = 4 then
              grdMatrizConseq.Cell[i + 7, NumCrit].Text := TrocaVP
                (TIWUserSession(WebApplication.Data).Cost[i]);
            for j := 0 to NumCrit - 1 do
            begin
              grdMatrizConseq.Cell[i + 7, j + 1].Text := TrocaVP
                (TIWUserSession(WebApplication.Data).ConsMatrix[i, j]);
            end;
          end;
          /// ////////////////////////////////////////////////////////////////////////
          // end;
        end;

        // após recuperar todas as informações necessárias para retomar um problema, verifica em qual estagiointra
        // a elicitação foi interrompida, redirecionando paraa tela apropriada
        if (TIWUserSession(WebApplication.Data).Estagiointra = 0) { OR
          (TIWUserSession(WebApplication.Data).Estagiointra = 1) } then
        begin
          TIWUserSession(WebApplication.Data).retornointra := 0; // var global que auxilia na determinação de qual tela intra o sis retomará
          //TTFrmIntraCriterio.Create(WebApplication).show;
          rgnLoginproject.Visible := false;
        end; // estagiointra 0

        { ALTERAÇÃO 20.02.22 }
        if (TIWUserSession(WebApplication.Data).Estagiointra = 1) then
        begin
          TIWUserSession(WebApplication.Data).retornointra := 1;
          //TTFrmIntraCriterio.Create(WebApplication).show;
          rgnLoginproject.Visible := false;

          for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
            TIWUserSession(WebApplication.Data).FTipo[i] := TIWUserSession
              (WebApplication.Data).NewFtipo[i];

        end; // estagiointra 1  - FIM ALTERAÇÃO

        if (TIWUserSession(WebApplication.Data).Estagiointra = 2) OR
          (TIWUserSession(WebApplication.Data).Estagiointra = 3) then
        begin
          rgnLoginproject.Visible := false;
          rgnInputMatrizConseq.Visible := true;
        end; // estagiointra 2 ou 3

      end;
    end;

    if (TIWUserSession(WebApplication.Data).estagio = 2) or
      (TIWUserSession(WebApplication.Data).estagio = 3) then
    begin
      // procedimento para pegar o round, caracteristicas dos criterios e das alternativas
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 15;
      SalvarBD;

      // ** Recupera a informação dos portfólios ** //
      if TIWUserSession(WebApplication.Data).Problematica = 4 then
      begin
        TIWUserSession(WebApplication.Data).AuxsalvarBD := 43;
        SalvarBD;
      end;

      // procedimento para retornar dados dos criterios
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 16;
      SalvarBD;

      // procedimento para retornar dados da AH
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 19;
      SalvarBD;

      // procedimento para retornar dados da elicitação
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 18;
      SalvarBD;

      // procedimento para retornar dados da elicitação
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 21;
      SalvarBD;

      // procedimento para retornar dados da avaliação holistica
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 29;
      SalvarBD;

      /// ////////////////////////////////////////////////////////////////////////////////////////
      // preenchimento de variaveis que vão depender da resposta da heuristica e de outros fatores
      if (TIWUserSession(WebApplication.Data).HeurQuest = 0) or
        (TIWUserSession(WebApplication.Data).HeurQuest = 2) then
      begin
        // se a resposta da heuristica for A ou indiferente
        TIWUserSession(WebApplication.Data).Etapa := 2;
      end; // referente a  if (TIWUserSession(WebApplication.Data).HeurQuest=0) or (TIWUserSession(WebApplication.Data).HeurQuest=2) then

      if TIWUserSession(WebApplication.Data).HeurQuest = 1 then
      begin
        // se a resposta da heuristica for igual a B

        if TIWUserSession(WebApplication.Data).NumQuest =
          (TIWUserSession(WebApplication.Data).NumCrit - TIWUserSession
            (WebApplication.Data).NumCriBlind) then
        begin
          TIWUserSession(WebApplication.Data).Etapa := 2;
        end; // referente a if TIWUserSession(WebApplication.Data).NumQuest=(TIWUserSession(WebApplication.Data).NumCrit-TIWUserSession(WebApplication.Data).NumCriBlind) then

        if TIWUserSession(WebApplication.Data).NumQuest <
          (Round(0.5 * (TIWUserSession(WebApplication.Data)
                .NumCrit - TIWUserSession(WebApplication.Data).NumCriBlind)))
          then
        begin
          TIWUserSession(WebApplication.Data).Etapa := 1;
          TIWUserSession(WebApplication.Data).Fase := 1;
        end // referente a if TIWUserSession(WebApplication.Data).NumQuest < (round(0.5*(TIWUserSession(WebApplication.Data).NumCrit - TIWUserSession(WebApplication.Data).NumCriBlind)))then
        else
        begin
          TIWUserSession(WebApplication.Data).Etapa := 1;
          TIWUserSession(WebApplication.Data).Fase := 2;
        end;
      end; // referente a if TIWUserSession(WebApplication.Data).HeurQuest=1 then

      // devemos retomar as informações dos limites das relações
      // ** entre os crit não adj e quais foram blindadas **//

      TIWUserSession(WebApplication.Data).AuxsalvarBD := 45;
      SalvarBD;
      /// ///////////////////////////////////////////////////////////////////////////////////////////////////

      { if TIWUserSession(WebApplication.Data).auxparacasonsaieretone=-1002525 then
        TIWUserSession(WebApplication.Data).stage:=0; }

      TIWUserSession(WebApplication.Data).RETPROBDEPOIS := -4555;
      Inicializa;
      // TIWUserSession(WebApplication.Data).stage := 3;
      TIWUserSession(WebApplication.Data).epsilon := trocapv(edtMaxdif.Text);
      rgnLoginproject.Visible := false;





      // Verificação para passar ou não direto para a unit de resultados
      if (TIWUserSession(WebApplication.Data).elicitarperfis = true) and
        (TIWUserSession(WebApplication.Data).Problematica = 3) then
      begin
        Tfrmaintercriterio.Create(WebApplication).btnprofileelicitClick(Self);
      end
      else
        begin
        Tfrmaintercriterio.Create(WebApplication).show;

      // Verificação para passar ou não direto para a unit de resultados (caso não seja necessário realizar a elicitção flexível)
      { if TIWUserSession(WebApplication.Data).Sol <> -5 then
        begin }
      Tfrmaintercriterio.Create(WebApplication).BtnshowresultsClick(Self);
      // end;

        end;

    end; // referente a if (TIWUserSession(WebApplication.Data).estagio = 2) or (TIWUserSession(WebApplication.Data).estagio = 3) then

    if TIWUserSession(WebApplication.Data).estagio = -1 then
    begin
      rgnLoginproject.Visible := false;
      rgnConseqvalues.Visible := true;
      TIWUserSession(WebApplication.Data).Nomeprojeto := TIWUserSession
        (WebApplication.Data).auxnome;
      // procedimento para retornar dados da elicitação
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 25;
      SalvarBD;

      with TIWUserSession(WebApplication.Data) do
      begin
        Setlength(ArrayAlt, TIWUserSession(WebApplication.Data).NumAlt);

        if (Problematica = 2) or (Problematica = 4) then
        begin
          Setlength(ArrayCrt, NumCrit + 1);
          Setlength(MatEdit, NumAlt, NumCrit + 1);

          lbllembretebudget.Visible := true;
          lbllembretebudget.Caption := 'Budget value: $' + FloatToStr
            (TIWUserSession(WebApplication.Data).budget);

          for i := 0 to NumCrit do

          Begin
            if i < NumCrit then
            begin
              ArrayCrt[i] := TIWLabel.Create(Self);
              ArrayCrt[i].parent := rgnconsman;
              ArrayCrt[i].Text := ('C' + inttostr(i + 1));
              ArrayCrt[i].Left := (45 + (i * 60));
              ArrayCrt[i].Top := (20);
              TIWUserSession(WebApplication.Data).ArrayCrt[i].Font.Style :=
                [fsbold];
            end
            else
            begin
              ArrayCrt[i] := TIWLabel.Create(Self);
              ArrayCrt[i].parent := rgnconsman;
              TIWUserSession(WebApplication.Data).ArrayCrt[i].Text :=
                ('Cost ($)');
              ArrayCrt[i].Left := (45 + (i * 60));
              ArrayCrt[i].Top := (20);
              ArrayCrt[i].Font.Style := [fsbold];
            end;
          end;
        end
        else
          for i := 0 to NumCrit - 1 do
          begin
            ArrayCrt[i] := TIWLabel.Create(Self);
            ArrayCrt[i].parent := rgnconsman;
            ArrayCrt[i].Text := ('C' + inttostr(i + 1));
            ArrayCrt[i].Left := (45 + (i * 60));
            ArrayCrt[i].Top := (20);
            ArrayCrt[i].Font.Style := [fsbold];
          end;
      end;

      for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
      begin
        TIWUserSession(WebApplication.Data).ArrayAlt[i] := TIWLabel.Create
          (Self);
        TIWUserSession(WebApplication.Data).ArrayAlt[i].parent := rgnconsman;
        TIWUserSession(WebApplication.Data).ArrayAlt[i].Text :=
          ('Alt' + inttostr(i + 1));
        TIWUserSession(WebApplication.Data).ArrayAlt[i].Left := (5);
        TIWUserSession(WebApplication.Data).ArrayAlt[i].Top := (50 + (i * 25));
        TIWUserSession(WebApplication.Data).ArrayAlt[i].Font.Style := [fsbold];
        TIWUserSession(WebApplication.Data).ArrayAlt[i].Font.FontName :=
          'Verdana';
      end;

      { setlength(ArrayCrt, TIWUserSession(WebApplication.Data).numcrit);
        for i := 0 to TIWUserSession(WebApplication.Data).numcrit - 1 do
        begin
        ArrayCrt[i] := TIWLabel.Create(Self);
        ArrayCrt[i].parent := rgnconsman;
        ArrayCrt[i].Text := ('C' + inttostr(i + 1));
        ArrayCrt[i].Left := (45 + (i * 60));
        ArrayCrt[i].Top := (20);
        ArrayCrt[i].Font.Style := [fsbold];
        end;
        end; }

      if (TIWUserSession(WebApplication.Data).Problematica = 2) or
        (TIWUserSession(WebApplication.Data).Problematica = 4) then
      begin
        Setlength(TIWUserSession(WebApplication.Data).MatEdit, TIWUserSession
            (WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data)
            .NumCrit + 1);
      end
      else
      begin
        Setlength(TIWUserSession(WebApplication.Data).MatEdit, TIWUserSession
            (WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data)
            .NumCrit);
      end;
      { for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
        begin
        for j := 0 to TIWUserSession(WebApplication.Data).numcrit - 1 do

        begin
        // Criando os edits que recaberão os desempenhos a partir de uma Matriz de TIWEdit's
        TIWUserSession(WebApplication.Data).MatEdit[i][j] := TIWEdit.Create
        (Self);
        TIWUserSession(WebApplication.Data).MatEdit[i][j].parent :=
        rgnconsman;
        TIWUserSession(WebApplication.Data).MatEdit[i][j].Text := ('');
        TIWUserSession(WebApplication.Data).MatEdit[i][j].Left :=
        (45 + (j * 60));
        TIWUserSession(WebApplication.Data).MatEdit[i][j].Top :=
        (25 + (i * 25) + 23);
        TIWUserSession(WebApplication.Data).MatEdit[i][j].Width := 60;
        end;
        end; }

      for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
      begin

        if (TIWUserSession(WebApplication.Data).Problematica = 2) or
          (TIWUserSession(WebApplication.Data).Problematica = 4) then
        begin
          for j := 0 to TIWUserSession(WebApplication.Data).NumCrit do
          begin
            // Criando os edits que recaberão consequencias e valor dos custos na Matriz de TIWEdit's
            TIWUserSession(WebApplication.Data).MatEdit[i][j] := TIWEdit.Create
              (Self);
            TIWUserSession(WebApplication.Data).MatEdit[i][j].parent :=
              rgnconsman;
            if j = TIWUserSession(WebApplication.Data).NumCrit then
              TIWUserSession(WebApplication.Data).MatEdit[i][j].BGColor :=
                clMenuBar;
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Text := ('');
            TIWUserSession(WebApplication.Data)
              .MatEdit[i][TIWUserSession(WebApplication.Data).NumCrit].Text :=
              ('$');
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Left :=
              (45 + (j * 60));
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Top :=
              (25 + (i * 25) + 23);
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Width := 60;
          end;
        end
        else
          for j := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
          begin
            // Criando os edits que recaberão os desempenhos a partir de uma Matriz de TIWEdit's
            TIWUserSession(WebApplication.Data).MatEdit[i][j] := TIWEdit.Create
              (Self);
            TIWUserSession(WebApplication.Data).MatEdit[i][j].parent :=
              rgnconsman;
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Text := ('');
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Left :=
              (45 + (j * 60));
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Top :=
              (25 + (i * 25) + 23);
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Width := 60;
          end; // if prob=2

      end; // for i

      // alteração feita em 20.02.22 para permitir que retomada manual com info parcial ocorra corretamente
      Setlength(TIWUserSession(WebApplication.Data).Verificaintra,
        TIWUserSession(WebApplication.Data).NumCrit);

      // cmproblematica1.ItemIndex := -1;  // Não limpo o nome caso o usuário volte para essa página

      AuxTipo[0] := 'Linear';
      AuxTipo[1] := 'Exponential';
      AuxTipo[2] := 'Logarithmic';
      AuxTipo[3] := 'Logistic';
      // ** Preenchendo a tabela de informações dos critérios **//

      grdCritInfo.ColumnCount := 7;
      grdCritInfo.RowCount := TIWUserSession(WebApplication.Data).NumCrit + 1;

      grdCritInfo.Cell[0, 0].Text := 'Name';
      grdCritInfo.Cell[0, 1].Text := 'Scale';
      grdCritInfo.Cell[0, 2].Text := 'Number of levels';
      grdCritInfo.Cell[0, 3].Text := 'Value Function';
      grdCritInfo.Cell[0, 4].Text := 'Direction';
      grdCritInfo.Cell[0, 5].Text := 'Par a';
      grdCritInfo.Cell[0, 6].Text := 'Par b';

      with TIWUserSession(WebApplication.Data) do // ALTERADO EM 20.02.22
        for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
        begin
          grdCritInfo.Cell[i + 1, 0].Text := TIWUserSession(WebApplication.Data)
            .NomeCrit[i];
          if TIWUserSession(WebApplication.Data).Typeofcrit[i] > 1 then
          begin
            grdCritInfo.Cell[i + 1, 1].Text := 'Discrete';
            grdCritInfo.Cell[i + 1, 2].Text := inttostr
              (TIWUserSession(WebApplication.Data).Niveis[i]);
          end // end do   if TIWUserSession(WebApplication.Data).Typeofcrit[I] > 1
          else
          begin
            grdCritInfo.Cell[i + 1, 1].Text := 'Continuous';
            grdCritInfo.Cell[i + 1, 2].Text := '---';
          end; // end do else

          { ALTERÃÇÃO FEITA EM 20.02 }
          // quando Ftipo=0 (elic intra), escreve direto na célula do grid. Sem utilizar AuxTipo
          if FTipo[i] = 0 then
          begin
            grdCritInfo.Cell[i + 1, 3].Text := '---';
            TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
          end
          else
            grdCritInfo.Cell[i + 1, 3].Text := AuxTipo
              [TIWUserSession(WebApplication.Data).FTipo[i] - 1];

          /// //COLOCANDO TYPECRIT SEPARADO, DEVIDO QUESTÃO DE EDTS VAZIOS//////
          if (Typeofcrit[i] = 1) then // CONT MAX
          begin
            grdCritInfo.Cell[i + 1, 4].Text := 'Maximization';

            if (edtparA.Text = '') then
            begin
              grdCritInfo.Cell[i + 1, 5].Text := '---';
              if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
            end;

            if (edtparB.Text = '') then
            begin
              grdCritInfo.Cell[i + 1, 6].Text := '---';
              if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
            end // se vazio

            else
            begin
              grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
              grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
              if (parA[i] = 0) then
                grdCritInfo.Cell[i + 1, 5].Text := '---';
              if (parB[i] = 0) then
                grdCritInfo.Cell[i + 1, 6].Text := '---';
            end; // se possui valor
          end; // type=1

          if (Typeofcrit[i] = 3) then // DISC MAX
          begin
            grdCritInfo.Cell[i + 1, 4].Text := 'Maximization';
            grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
            grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
            if (parA[i] = 0) then
              grdCritInfo.Cell[i + 1, 5].Text := '---';
            if (parB[i] = 0) then
              grdCritInfo.Cell[i + 1, 6].Text := '---';
          end
          else

            if (Typeofcrit[i] = 0) then // CONT MIN
          begin
            grdCritInfo.Cell[i + 1, 4].Text := 'Minimization';

            if (edtparA.Text = '') then
            begin
              grdCritInfo.Cell[i + 1, 5].Text := '---';
              if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
            end;

            if (edtparB.Text = '') then
            begin
              grdCritInfo.Cell[i + 1, 6].Text := '---';
              if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
            end // se vazio

            else
            begin
              grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
              grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
              if (parA[i] = 0) then
                grdCritInfo.Cell[i + 1, 5].Text := '---';
              if (parB[i] = 0) then
                grdCritInfo.Cell[i + 1, 6].Text := '---';
            end; // se possui valor
          end; // type= 0

          if (Typeofcrit[i] = 2) then // DISC MIN
          begin
            grdCritInfo.Cell[i + 1, 4].Text := 'Minimization';
            grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
            grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
            if (parA[i] = 0) then
              grdCritInfo.Cell[i + 1, 5].Text := '---';
            if (parB[i] = 0) then
              grdCritInfo.Cell[i + 1, 6].Text := '---';
          end;
        end;
      // FINAL ALTERAÇÃO EM 10.02.22//

      { if (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 1) OR
        (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 3) then
        begin
        grdCritInfo.Cell[i + 1, 4].Text := 'Maximization';
        grdCritInfo.Cell[i + 1, 5].Text := FloatToStr
        (TIWUserSession(WebApplication.Data).parA[i]);
        grdCritInfo.Cell[i + 1, 6].Text := FloatToStr
        (TIWUserSession(WebApplication.Data).parB[i]);
        if (TIWUserSession(WebApplication.Data).parA[i] = 0) then
        grdCritInfo.Cell[i + 1, 5].Text := '---';
        if (TIWUserSession(WebApplication.Data).parB[i] = 0) then
        grdCritInfo.Cell[i + 1, 6].Text := '---';
        end // end do  if (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 1)...
        else if (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 0) OR
        (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 2) then
        begin
        grdCritInfo.Cell[i + 1, 4].Text := 'Minimization';
        grdCritInfo.Cell[i + 1, 5].Text := FloatToStr
        (TIWUserSession(WebApplication.Data).parA[i]);
        grdCritInfo.Cell[i + 1, 6].Text := FloatToStr
        (TIWUserSession(WebApplication.Data).parB[i]);
        if (TIWUserSession(WebApplication.Data).parA[i] = 0) then
        grdCritInfo.Cell[i + 1, 5].Text := '---';
        if (TIWUserSession(WebApplication.Data).parB[i] = 0) then
        grdCritInfo.Cell[i + 1, 6].Text := '---';
        end; // end do else
        end; } // end do for I := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do

      // criando o grid com as alternativas
      grdinformaalt.ColumnCount := 2;
      grdinformaalt.RowCount := TIWUserSession(WebApplication.Data).NumAlt;
      for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
      begin
        grdinformaalt.Cell[i, 0].Text := 'Alt' + ' ' + inttostr(i + 1);
        grdinformaalt.Cell[i, 1].Text := TIWUserSession(WebApplication.Data)
          .NomeAlt[i];
      end; // referente a for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do

    end; // end do -1

    if TIWUserSession(WebApplication.Data).estagio = 6 then
    begin
      rgnLoginproject.Visible := false;
      rgnConseqvalues.Visible := true;
      TIWUserSession(WebApplication.Data).Nomeprojeto := TIWUserSession
        (WebApplication.Data).auxnome;
      // procedimento para retornar dados da elicitação
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 31;
      SalvarBD;

      with TIWUserSession(WebApplication.Data) do
      begin
        Setlength(ArrayCrt, NumCrit);
        Setlength(ArrayAlt, NumAlt);
        for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
        begin
          ArrayAlt[i] := TIWLabel.Create(Self);
          ArrayAlt[i].parent := rgnconsman;
          ArrayAlt[i].Text := ('Alt' + inttostr(i + 1));
          ArrayAlt[i].Left := (5);
          ArrayAlt[i].Top := (50 + (i * 25));
          ArrayAlt[i].Font.Style := [fsbold];
        end;

        { setlength(ArrayCrt, TIWUserSession(WebApplication.Data).numcrit);
          for i := 0 to TIWUserSession(WebApplication.Data).numcrit - 1 do
          begin
          ArrayCrt[i] := TIWLabel.Create(Self);
          ArrayCrt[i].parent := rgnconsman;
          ArrayCrt[i].Text := ('C' + inttostr(i + 1));
          ArrayCrt[i].Left := (45 + (i * 60));
          ArrayCrt[i].Top := (20);
          ArrayCrt[i].Font.Style := [fsbold];
          end; }
        if (Problematica = 2) or (TIWUserSession(WebApplication.Data)
            .Problematica = 4) then
        begin
          Setlength(ArrayCrt, NumCrit + 1);
          Setlength(MatEdit, NumAlt, NumCrit + 1);

          lbllembretebudget.Visible := true;
          lbllembretebudget.Caption := 'Budget value: $' + FloatToStr
            (TIWUserSession(WebApplication.Data).budget);

          for i := 0 to NumCrit do

          Begin
            if i < NumCrit then
            begin
              ArrayCrt[i] := TIWLabel.Create(Self);
              ArrayCrt[i].parent := rgnconsman;
              ArrayCrt[i].Text := ('C' + inttostr(i + 1));
              ArrayCrt[i].Left := (45 + (i * 60));
              ArrayCrt[i].Top := (20);
              TIWUserSession(WebApplication.Data).ArrayCrt[i].Font.Style :=
                [fsbold];
            end
            else
            begin
              ArrayCrt[i] := TIWLabel.Create(Self);
              ArrayCrt[i].parent := rgnconsman;
              TIWUserSession(WebApplication.Data).ArrayCrt[i].Text :=
                ('Cost ($)');
              ArrayCrt[i].Left := (45 + (i * 60));
              ArrayCrt[i].Top := (20);
              ArrayCrt[i].Font.Style := [fsbold];
            end;
          end;
        end
        else
          for i := 0 to NumCrit - 1 do
          begin
            ArrayCrt[i] := TIWLabel.Create(Self);
            ArrayCrt[i].parent := rgnconsman;
            ArrayCrt[i].Text := ('C' + inttostr(i + 1));
            ArrayCrt[i].Left := (45 + (i * 60));
            ArrayCrt[i].Top := (20);
            ArrayCrt[i].Font.Style := [fsbold];
          end;
      end;
      // end;

      { setlength(TIWUserSession(WebApplication.Data).MatEdit, TIWUserSession
        (WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data)
        .numcrit); }
      if (TIWUserSession(WebApplication.Data).Problematica = 2) or
        (TIWUserSession(WebApplication.Data).Problematica = 4) then
      begin
        Setlength(TIWUserSession(WebApplication.Data).MatEdit, TIWUserSession
            (WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data)
            .NumCrit + 1);
      end
      else
      begin
        Setlength(TIWUserSession(WebApplication.Data).MatEdit, TIWUserSession
            (WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data)
            .NumCrit);
      end;

      if (TIWUserSession(WebApplication.Data).Problematica = 2) or
        (TIWUserSession(WebApplication.Data).Problematica = 4) then
      begin
        for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
        begin
          for j := 0 to TIWUserSession(WebApplication.Data).NumCrit do
          begin
            // Criando os edits que recaberão os desempenhos a partir de uma Matriz de TIWEdit's
            TIWUserSession(WebApplication.Data).MatEdit[i][j] := TIWEdit.Create
              (Self);
            TIWUserSession(WebApplication.Data).MatEdit[i][j].parent :=
              rgnconsman;
            if j < TIWUserSession(WebApplication.Data).NumCrit then
            begin
              if TIWUserSession(WebApplication.Data).ConsMatrix[i, j] <> 0 then
                TIWUserSession(WebApplication.Data).MatEdit[i][j].Text :=
                  FloatToStr(TIWUserSession(WebApplication.Data)
                    .ConsMatrix[i, j]);
              if TIWUserSession(WebApplication.Data).ConsMatrix[i, j] = 0 then
                TIWUserSession(WebApplication.Data).MatEdit[i][j].Text := '';
            end;

            if j = TIWUserSession(WebApplication.Data).NumCrit then
            Begin
              if TIWUserSession(WebApplication.Data).Cost[i] <> 0 then
                TIWUserSession(WebApplication.Data).MatEdit[i][j].Text :=
                  '$' + FloatToStr(TIWUserSession(WebApplication.Data).Cost[i]);
              if TIWUserSession(WebApplication.Data).Cost[i] = 0 then
                TIWUserSession(WebApplication.Data).MatEdit[i][j].Text := '$';
            End;

            TIWUserSession(WebApplication.Data).MatEdit[i][j].Left :=
              (45 + (j * 60));
            if j = TIWUserSession(WebApplication.Data).NumCrit then
              TIWUserSession(WebApplication.Data).MatEdit[i][j].BGColor :=
                clMenuBar;
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Top :=
              (25 + (i * 25) + 23);
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Width := 60;
          end;
        end;
      end
      else
      Begin

        // alteração feita em 21.01.22 para permitir que retomada manual com info parcial ocorra corretamente
        Setlength(TIWUserSession(WebApplication.Data).Verificaintra,
          TIWUserSession(WebApplication.Data).NumCrit);

        for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
        begin
          for j := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do

          begin
            // Criando os edits que recaberão os desempenhos a partir de uma Matriz de TIWEdit's
            TIWUserSession(WebApplication.Data).MatEdit[i][j] := TIWEdit.Create
              (Self);
            TIWUserSession(WebApplication.Data).MatEdit[i][j].parent :=
              rgnconsman;
            if TIWUserSession(WebApplication.Data).ConsMatrix[i, j] <> 0 then
              TIWUserSession(WebApplication.Data).MatEdit[i][j].Text :=
                FloatToStr(TIWUserSession(WebApplication.Data)
                  .ConsMatrix[i, j]);
            if TIWUserSession(WebApplication.Data).ConsMatrix[i, j] = 0 then
              TIWUserSession(WebApplication.Data).MatEdit[i][j].Text := '';
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Left :=
              (45 + (j * 60));
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Top :=
              (25 + (i * 25) + 23);
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Width := 60;
          end;
        end;
      end;
      // cmproblematica1.ItemIndex := -1;  // Não limpo o nome caso o usuário volte para essa página

      AuxTipo[0] := 'Linear';
      AuxTipo[1] := 'Exponential';
      AuxTipo[2] := 'Logarithmic';
      AuxTipo[3] := 'Logistic';

      // Alteração feita em 21.01.22//
      // ** Preenchendo a tabela de informações dos critérios **//
      grdCritInfo.ColumnCount := 7;
      grdCritInfo.RowCount := TIWUserSession(WebApplication.Data).NumCrit + 1;

      grdCritInfo.Cell[0, 0].Text := 'Name';
      grdCritInfo.Cell[0, 1].Text := 'Scale';
      grdCritInfo.Cell[0, 2].Text := 'Number of levels';
      grdCritInfo.Cell[0, 3].Text := 'Value Function';
      grdCritInfo.Cell[0, 4].Text := 'Direction';
      grdCritInfo.Cell[0, 5].Text := 'Par a';
      grdCritInfo.Cell[0, 6].Text := 'Par b';

      with TIWUserSession(WebApplication.Data) do
        for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
        begin
          grdCritInfo.Cell[i + 1, 0].Text := TIWUserSession(WebApplication.Data)
            .NomeCrit[i];
          if TIWUserSession(WebApplication.Data).Typeofcrit[i] > 1 then
          begin
            grdCritInfo.Cell[i + 1, 1].Text := 'Discrete';
            grdCritInfo.Cell[i + 1, 2].Text := inttostr
              (TIWUserSession(WebApplication.Data).Niveis[i]);
          end // end do   if TIWUserSession(WebApplication.Data).Typeofcrit[I] > 1
          else
          begin
            grdCritInfo.Cell[i + 1, 1].Text := 'Continuous';
            grdCritInfo.Cell[i + 1, 2].Text := '---';
          end; // end do else

          // quando Ftipo=0 (elic intra), escreve direto na célula do grid. Sem utilizar AuxTipo
          if FTipo[i] = 0 then
          begin
            grdCritInfo.Cell[i + 1, 3].Text := '---';
            TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
          end
          else
            grdCritInfo.Cell[i + 1, 3].Text := AuxTipo
              [TIWUserSession(WebApplication.Data).FTipo[i] - 1];

          /// //COLOCANDO TYPECRIT SEPARADO, DEVIDO QUESTÃO DE EDTS VAZIOSS//////
          if (Typeofcrit[i] = 1) then // CONT MAX
          begin
            grdCritInfo.Cell[i + 1, 4].Text := 'Maximization';

            if (edtparA.Text = '') then
            begin
              grdCritInfo.Cell[i + 1, 5].Text := '---';
              if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
            end;

            if (edtparB.Text = '') then
            begin
              grdCritInfo.Cell[i + 1, 6].Text := '---';
              if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
            end // se vazio

            else
            begin
              grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
              grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
              if (parA[i] = 0) then
                grdCritInfo.Cell[i + 1, 5].Text := '---';
              if (parB[i] = 0) then
                grdCritInfo.Cell[i + 1, 6].Text := '---';
            end; // se possui valor
          end; // type=1

          if (Typeofcrit[i] = 3) then // DISC MAX
          begin
            grdCritInfo.Cell[i + 1, 4].Text := 'Maximization';
            grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
            grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
            if (parA[i] = 0) then
              grdCritInfo.Cell[i + 1, 5].Text := '---';
            if (parB[i] = 0) then
              grdCritInfo.Cell[i + 1, 6].Text := '---';
          end
          else

            if (Typeofcrit[i] = 0) then // CONT MIN
          begin
            grdCritInfo.Cell[i + 1, 4].Text := 'Minimization';

            if (edtparA.Text = '') then
            begin
              grdCritInfo.Cell[i + 1, 5].Text := '---';
              if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
            end;

            if (edtparB.Text = '') then
            begin
              grdCritInfo.Cell[i + 1, 6].Text := '---';
              if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
            end // se vazio

            else
            begin
              grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
              grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
              if (parA[i] = 0) then
                grdCritInfo.Cell[i + 1, 5].Text := '---';
              if (parB[i] = 0) then
                grdCritInfo.Cell[i + 1, 6].Text := '---';
            end; // se possui valor
          end; // type= 0

          if (Typeofcrit[i] = 2) then // DISC MIN
          begin
            grdCritInfo.Cell[i + 1, 4].Text := 'Minimization';
            grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
            grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
            if (parA[i] = 0) then
              grdCritInfo.Cell[i + 1, 5].Text := '---';
            if (parB[i] = 0) then
              grdCritInfo.Cell[i + 1, 6].Text := '---';
          end;
        end;

      { grdCritInfo.Cell[i + 1, 3].Text := AuxTipo
        [TIWUserSession(WebApplication.Data).FTipo[i] - 1];

        if (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 1) OR
        (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 3) then
        begin
        grdCritInfo.Cell[i + 1, 4].Text := 'Maximization';
        grdCritInfo.Cell[i + 1, 5].Text := FloatToStr
        (TIWUserSession(WebApplication.Data).parA[i]);
        grdCritInfo.Cell[i + 1, 6].Text := FloatToStr
        (TIWUserSession(WebApplication.Data).parB[i]);
        if (TIWUserSession(WebApplication.Data).parA[i] = 0) then
        grdCritInfo.Cell[i + 1, 5].Text := '---';
        if (TIWUserSession(WebApplication.Data).parB[i] = 0) then
        grdCritInfo.Cell[i + 1, 6].Text := '---';
        end // end do  if (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 1)...
        else if (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 0) OR
        (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 2) then
        begin
        grdCritInfo.Cell[i + 1, 4].Text := 'Minimization';
        grdCritInfo.Cell[i + 1, 5].Text := FloatToStr
        (TIWUserSession(WebApplication.Data).parA[i]);
        grdCritInfo.Cell[i + 1, 6].Text := FloatToStr
        (TIWUserSession(WebApplication.Data).parB[i]);
        if (TIWUserSession(WebApplication.Data).parA[i] = 0) then
        grdCritInfo.Cell[i + 1, 5].Text := '---';
        if (TIWUserSession(WebApplication.Data).parB[i] = 0) then
        grdCritInfo.Cell[i + 1, 6].Text := '---';
        end; // end do else
        end; // end do for I := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do

        }   // fim alteração de 21.01.22

      // criando o grid com as alternativas
      grdinformaalt.ColumnCount := 2;
      grdinformaalt.RowCount := TIWUserSession(WebApplication.Data).NumAlt;
      for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
      begin
        grdinformaalt.Cell[i, 0].Text := 'Alt' + ' ' + inttostr(i + 1);
        grdinformaalt.Cell[i, 1].Text := TIWUserSession(WebApplication.Data)
          .NomeAlt[i];
      end; // referente a for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do

    end; // end do 6

    if (TIWUserSession(WebApplication.Data).estagio = 4) or
      (TIWUserSession(WebApplication.Data).estagio = 5) then
    begin
      // procedimento para pegar o round, caracteristicas dos criterios e das alternativas
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 15;
      SalvarBD;

      // ** Recupera a informação dos portfólios ** //
      if TIWUserSession(WebApplication.Data).Problematica = 4 then
      begin
        TIWUserSession(WebApplication.Data).AuxsalvarBD := 43;
        SalvarBD;
      end;

      // procedimento para retornar dados dos criterios
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 16;
      SalvarBD;

      // procedimento para retornar dados da AH
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 19;
      SalvarBD;

      // procedimento para retornar dados da elicitação
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 18;
      SalvarBD;

      // procedimento para retornar dados da elicitação
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 21;
      SalvarBD;

      // procedimento para retornar dados da avaliação holistica
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 29;
      SalvarBD;

      /// ////////////////////////////////////////////////////////////////////////////////////////
      // preenchimento de variaveis que vão depender da resposta da heuristica e de outros fatores
      if (TIWUserSession(WebApplication.Data).HeurQuest = 0) or
        (TIWUserSession(WebApplication.Data).HeurQuest = 2) then
      begin
        // se a resposta da heuristica for A ou indiferente
        TIWUserSession(WebApplication.Data).Etapa := 2;
      end; // referente a  if (TIWUserSession(WebApplication.Data).HeurQuest=0) or (TIWUserSession(WebApplication.Data).HeurQuest=2) then

      if TIWUserSession(WebApplication.Data).HeurQuest = 1 then
      begin
        // se a resposta da heuristica for igual a B

        if TIWUserSession(WebApplication.Data).NumQuest =
          (TIWUserSession(WebApplication.Data).NumCrit - TIWUserSession
            (WebApplication.Data).NumCriBlind) then
        begin
          TIWUserSession(WebApplication.Data).Etapa := 2;
        end; // referente a if TIWUserSession(WebApplication.Data).NumQuest=(TIWUserSession(WebApplication.Data).NumCrit-TIWUserSession(WebApplication.Data).NumCriBlind) then

        if TIWUserSession(WebApplication.Data).NumQuest <
          (Round(0.5 * (TIWUserSession(WebApplication.Data)
                .NumCrit - TIWUserSession(WebApplication.Data).NumCriBlind)))
          then
        begin
          TIWUserSession(WebApplication.Data).Etapa := 1;
          TIWUserSession(WebApplication.Data).Fase := 1;
        end // referente a if TIWUserSession(WebApplication.Data).NumQuest < (round(0.5*(TIWUserSession(WebApplication.Data).NumCrit - TIWUserSession(WebApplication.Data).NumCriBlind)))then
        else
        begin
          TIWUserSession(WebApplication.Data).Etapa := 1;
          TIWUserSession(WebApplication.Data).Fase := 2;
        end;
      end; // referente a if TIWUserSession(WebApplication.Data).HeurQuest=1 then

      // ** devemos retomar as informações dos limites das relações
      // ** entre os crit não adj e quais foram blindadas **//

      TIWUserSession(WebApplication.Data).AuxsalvarBD := 45;
      SalvarBD;
      /// ///////////////////////////////////////////////////////////////////////////////////////////////////

      { if TIWUserSession(WebApplication.Data).auxparacasonsaieretone=-1002525 then
        TIWUserSession(WebApplication.Data).stage:=0; }

      TIWUserSession(WebApplication.Data).RETPROBDEPOIS := -4555;
      Inicializa;
      // TIWUserSession(WebApplication.Data).stage := 3;
      TIWUserSession(WebApplication.Data).epsilon := trocapv(edtMaxdif.Text);
      rgnLoginproject.Visible := false;





      // Verificação para passar ou não direto para a unit de resultados
      if (TIWUserSession(WebApplication.Data).elicitarperfis = true) and
        (TIWUserSession(WebApplication.Data).Problematica = 3) then
      begin
        Tfrmaintercriterio.Create(WebApplication).btnprofileelicitClick(Self);
      end
      else
      begin
        Tfrmaintercriterio.Create(WebApplication).show;
      // Verificação para passar ou não direto para a unit de resultados
      { if TIWUserSession(WebApplication.Data).Sol <> -5 then
        begin }
      TIWUserSession(WebApplication.Data).RETPROBDEPOIS := -4555;
      Tfrmaintercriterio.Create(WebApplication).BtnshowresultsClick(Self);
      // end;

      end;

      // Verificação para passar ou não direto para a unit de resultados
      if TIWUserSession(WebApplication.Data).Sol = -5 then
      begin
        TIWUserSession(WebApplication.Data).RETPROBDEPOIS := -4555;
        Tfrmaintercriterio.Create(WebApplication).btnstopelicitationClick(Self);
      end;

    end; // referente a if (TIWUserSession(WebApplication.Data).estagio = 5) or (TIWUserSession(WebApplication.Data).estagio = 4) then

  end; // referente a if datainicioproblema='19/04/2021' then



  // começou e terminou antes de 19/04/2021

  if (StrToDate(TIWUserSession(WebApplication.Data).dtinicretorn) < StrToDate
      ('19/04/2021')) AND (StrToDate(TIWUserSession(WebApplication.Data)
        .dtfnretorn) < StrToDate('19/04/2021')) then
  begin
    // procedimento para retornar o número de critérios
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 32;
    SalvarBD;
    if TIWUserSession(WebApplication.Data).NumCrit <= 10 then
    begin
      if (TIWUserSession(WebApplication.Data).estagio = 0) or
        (TIWUserSession(WebApplication.Data).estagio = 1) then
      begin
        TIWUserSession(WebApplication.Data).tempinicio := TimeToStr(Now);

        // procedimento para pegar o round, caracteristicas dos criterios e das alternativas
        TIWUserSession(WebApplication.Data).AuxsalvarBD := 15;
        SalvarBD;
        // rgnInputMatrizConseq.Visible := true;
        // rgnLoginproject.Visible := false;
        // **Deixa visível a opção para seguir pra elicitação direto **//
        AlreadyRakend1.Visible := true;


        // procedimento para preencher as informações na região rgnInputMatrizConseq

        Inicializa;
        if TIWUserSession(WebApplication.Data).Flag <> -2 then
        begin
          with TIWUserSession(WebApplication.Data) do
          begin
            /// / Setando o tamanho dos grids que exibem a informação importada na região imput ////
            grdMatrizConseq.RowCount := NumAlt + 7;
            grdMatrizConseq.ColumnCount := NumCrit + 1;
            probname := edtEntername.Text;

            /// / Preenchendo a informação fixa dos grids ////
            grdMatrizConseq.Cell[0, 0].Text := 'Criteria:';
            grdMatrizConseq.Cell[1, 0].Text :=
              '0-Cont Min; 1-Cont Max; 2-Disc Min; 3- Disc Max; 4- Int Min; 5- Int Max:';
            { grdInputvalues.Cell[2, 0].text := 'Weights:'; }
            grdMatrizConseq.Cell[2, 0].Text := '';
            grdMatrizConseq.Cell[3, 0].Text := '';
            grdMatrizConseq.Cell[4, 0].Text := '';
            grdMatrizConseq.Cell[5, 0].Text :=
              'Number of levels of discrete criteria';
            grdMatrizConseq.Cell[6, 0].Text := 'Consequence Matrix:';

            /// / Preenchendo os Grids com a informação Imputada ///////////////////////
            for j := 0 to NumCrit - 1 do
            begin
              grdMatrizConseq.Cell[0, j + 1].Text := NomeCrit[j];
              if Cri_inteiro[j] then
              begin
                if Typeofcrit[j] = 0 then
                  grdMatrizConseq.Cell[1, j + 1].Text := '4';
                if Typeofcrit[j] = 1 then
                  grdMatrizConseq.Cell[1, j + 1].Text := '5';
              end
              else
                grdMatrizConseq.Cell[1, j + 1].Text := inttostr(Typeofcrit[j]);
              { grdMatrizConseq.Cell[2, j + 1].Text := inttostr(FTipo[j]);
                if FTipo[j] <> 1 then
                begin
                grdMatrizConseq.Cell[3, j + 1].Text := FloatToStr(parA[j]);
                grdMatrizConseq.Cell[4, j + 1].Text := FloatToStr(parB[j]);
                end; }
              grdMatrizConseq.Cell[5, j + 1].Text := inttostr(Niveis[j]);
            end;

            for i := 0 to NumAlt - 1 do
            begin
              grdMatrizConseq.Cell[i + 7, 0].Text := NomeAlt[i];
              for j := 0 to NumCrit - 1 do
              begin
                grdMatrizConseq.Cell[i + 7, j + 1].Text := TrocaVP
                  (ConsMatrix[i, j]);
              end;
            end;
            /// ////////////////////////////////////////////////////////////////////////
            // end;
          end;

          // após recuperar todas as informações necessárias para retomar um problema, verifica em qual estagiointra
          // a elicitação foi interrompida, redirecionando paraa tela apropriada
          if (TIWUserSession(WebApplication.Data).Estagiointra = 0) { OR
            (TIWUserSession(WebApplication.Data).Estagiointra = 1) } then
          begin
            TIWUserSession(WebApplication.Data).retornointra := 0; // var global que auxilia na determinação de qual tela intra o sis retomará
            //TTFrmIntraCriterio.Create(WebApplication).show;
            rgnLoginproject.Visible := false;
          end; // estagiointra 0

          { ALTERAÇÃO 20.02.22 }
          if (TIWUserSession(WebApplication.Data).Estagiointra = 1) then
          begin
            TIWUserSession(WebApplication.Data).retornointra := 1;
            //TTFrmIntraCriterio.Create(WebApplication).show;
            rgnLoginproject.Visible := false;

            for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
              TIWUserSession(WebApplication.Data).FTipo[i] := TIWUserSession
                (WebApplication.Data).NewFtipo[i];

          end; // estagiointra 1  - FIM ALTERAÇÃO

          if (TIWUserSession(WebApplication.Data).Estagiointra = 2) OR
            (TIWUserSession(WebApplication.Data).Estagiointra = 3) then
          begin
            rgnLoginproject.Visible := false;
            rgnInputMatrizConseq.Visible := true;
          end; // estagiointra 2 ou 3

        end;
      end;

      if (TIWUserSession(WebApplication.Data).estagio = 2) or
        (TIWUserSession(WebApplication.Data).estagio = 3) then
      begin
        // procedimento para pegar o round, caracteristicas dos criterios e das alternativas
        TIWUserSession(WebApplication.Data).AuxsalvarBD := 15;
        SalvarBD;
        // procedimento para retornar dados dos criterios
        TIWUserSession(WebApplication.Data).AuxsalvarBD := 33;
        SalvarBD;

        // procedimento para retornar dados da AH
        TIWUserSession(WebApplication.Data).AuxsalvarBD := 19;
        SalvarBD;

        // procedimento para retornar dados da elicitação
        TIWUserSession(WebApplication.Data).AuxsalvarBD := 18;
        SalvarBD;

        // procedimento para retornar dados da elicitação
        TIWUserSession(WebApplication.Data).AuxsalvarBD := 21;
        SalvarBD;

        // procedimento para retornar dados da avaliação holistica
        TIWUserSession(WebApplication.Data).AuxsalvarBD := 29;
        SalvarBD;

        /// ////////////////////////////////////////////////////////////////////////////////////////
        // preenchimento de variaveis que vão depender da resposta da heuristica e de outros fatores
        if (TIWUserSession(WebApplication.Data).HeurQuest = 0) or
          (TIWUserSession(WebApplication.Data).HeurQuest = 2) then
        begin
          // se a resposta da heuristica for A ou indiferente
          TIWUserSession(WebApplication.Data).Etapa := 2;
        end; // referente a  if (TIWUserSession(WebApplication.Data).HeurQuest=0) or (TIWUserSession(WebApplication.Data).HeurQuest=2) then

        if TIWUserSession(WebApplication.Data).HeurQuest = 1 then
        begin
          // se a resposta da heuristica for igual a B

          if TIWUserSession(WebApplication.Data).NumQuest =
            (TIWUserSession(WebApplication.Data).NumCrit - TIWUserSession
              (WebApplication.Data).NumCriBlind) then
          begin
            TIWUserSession(WebApplication.Data).Etapa := 2;
          end; // referente a if TIWUserSession(WebApplication.Data).NumQuest=(TIWUserSession(WebApplication.Data).NumCrit-TIWUserSession(WebApplication.Data).NumCriBlind) then

          if TIWUserSession(WebApplication.Data).NumQuest <
            (Round(0.5 * (TIWUserSession(WebApplication.Data)
                  .NumCrit - TIWUserSession(WebApplication.Data).NumCriBlind)))
            then
          begin
            TIWUserSession(WebApplication.Data).Etapa := 1;
            TIWUserSession(WebApplication.Data).Fase := 1;
          end // referente a if TIWUserSession(WebApplication.Data).NumQuest < (round(0.5*(TIWUserSession(WebApplication.Data).NumCrit - TIWUserSession(WebApplication.Data).NumCriBlind)))then
          else
          begin
            TIWUserSession(WebApplication.Data).Etapa := 1;
            TIWUserSession(WebApplication.Data).Fase := 2;
          end;
        end; // referente a if TIWUserSession(WebApplication.Data).HeurQuest=1 then

        /// ///////////////////////////////////////////////////////////////////////////////////////////////////

        { if TIWUserSession(WebApplication.Data).auxparacasonsaieretone=-1002525 then
          TIWUserSession(WebApplication.Data).stage:=0; }

        TIWUserSession(WebApplication.Data).RETPROBDEPOIS := -4555;
        Inicializa;
        // TIWUserSession(WebApplication.Data).stage := 3;
        TIWUserSession(WebApplication.Data).epsilon := trocapv(edtMaxdif.Text);
        rgnLoginproject.Visible := false;




        // Verificação para passar ou não direto para a unit de resultados
        if (TIWUserSession(WebApplication.Data).elicitarperfis = true) and
          (TIWUserSession(WebApplication.Data).Problematica = 3) then
        begin
          Tfrmaintercriterio.Create(WebApplication).btnprofileelicitClick(Self);
        end
        else
        begin
          Tfrmaintercriterio.Create(WebApplication).show;

        // Verificação para passar ou não direto para a unit de resultados (caso não seja necessário realizar a elicitção flexível)
        { if TIWUserSession(WebApplication.Data).Sol <> -5 then
          begin }
        Tfrmaintercriterio.Create(WebApplication).BtnshowresultsClick(Self);
        // end;

        end;

      end; // referente a if (TIWUserSession(WebApplication.Data).estagio = 2) or (TIWUserSession(WebApplication.Data).estagio = 3) then

      if TIWUserSession(WebApplication.Data).estagio = -1 then
      begin
        rgnLoginproject.Visible := false;
        rgnConseqvalues.Visible := true;
        TIWUserSession(WebApplication.Data).Nomeprojeto := TIWUserSession
          (WebApplication.Data).auxnome;
        // procedimento para retornar dados da elicitação
        TIWUserSession(WebApplication.Data).AuxsalvarBD := 25;
        SalvarBD;

        with TIWUserSession(WebApplication.Data) do
        begin
          Setlength(ArrayAlt, TIWUserSession(WebApplication.Data).NumAlt);
          for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
          begin
            ArrayAlt[i] := TIWLabel.Create(Self);
            ArrayAlt[i].parent := rgnconsman;
            ArrayAlt[i].Text := ('Alt' + inttostr(i + 1));
            ArrayAlt[i].Left := (5);
            ArrayAlt[i].Top := (50 + (i * 25));
            ArrayAlt[i].Font.Style := [fsbold];
          end;

          Setlength(ArrayCrt, TIWUserSession(WebApplication.Data).NumCrit);
          for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
          begin
            ArrayCrt[i] := TIWLabel.Create(Self);
            ArrayCrt[i].parent := rgnconsman;
            ArrayCrt[i].Text := ('C' + inttostr(i + 1));
            ArrayCrt[i].Left := (45 + (i * 60));
            ArrayCrt[i].Top := (20);
            ArrayCrt[i].Font.Style := [fsbold];
          end;
        end;

        Setlength(TIWUserSession(WebApplication.Data).MatEdit, TIWUserSession
            (WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data)
            .NumCrit);
        // alteração feita em 20.02 para permitir que retomada manual com info parcial ocorra corretamente
        Setlength(TIWUserSession(WebApplication.Data).Verificaintra,
          TIWUserSession(WebApplication.Data).NumCrit);

        for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
        begin
          for j := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do

          begin
            // Criando os edits que recaberão os desempenhos a partir de uma Matriz de TIWEdit's
            TIWUserSession(WebApplication.Data).MatEdit[i][j] := TIWEdit.Create
              (Self);
            TIWUserSession(WebApplication.Data).MatEdit[i][j].parent :=
              rgnconsman;
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Text := ('');
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Left :=
              (45 + (j * 60));
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Top :=
              (25 + (i * 25) + 23);
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Width := 60;
          end;
        end;
        // cmproblematica1.ItemIndex := -1;  // Não limpo o nome caso o usuário volte para essa página

        AuxTipo[0] := 'Linear';
        AuxTipo[1] := 'Exponential';
        AuxTipo[2] := 'Logarithmic';
        AuxTipo[3] := 'Logistic';
        // ** Preenchendo a tabela de informações dos critérios **//

        grdCritInfo.ColumnCount := 7;
        grdCritInfo.RowCount := TIWUserSession(WebApplication.Data).NumCrit + 1;

        grdCritInfo.Cell[0, 0].Text := 'Name';
        grdCritInfo.Cell[0, 1].Text := 'Scale';
        grdCritInfo.Cell[0, 2].Text := 'Number of levels';
        grdCritInfo.Cell[0, 3].Text := 'Value Function';
        grdCritInfo.Cell[0, 4].Text := 'Direction';
        grdCritInfo.Cell[0, 5].Text := 'Par a';
        grdCritInfo.Cell[0, 6].Text := 'Par b';

        with TIWUserSession(WebApplication.Data) do // alteração em 20.02
          for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
          begin
            grdCritInfo.Cell[i + 1, 0].Text := TIWUserSession
              (WebApplication.Data).NomeCrit[i];
            if TIWUserSession(WebApplication.Data).Typeofcrit[i] > 1 then
            begin
              grdCritInfo.Cell[i + 1, 1].Text := 'Discrete';
              grdCritInfo.Cell[i + 1, 2].Text := inttostr
                (TIWUserSession(WebApplication.Data).Niveis[i]);
            end // end do   if TIWUserSession(WebApplication.Data).Typeofcrit[I] > 1
            else
            begin
              grdCritInfo.Cell[i + 1, 1].Text := 'Continuous';
              grdCritInfo.Cell[i + 1, 2].Text := '---';
            end; // end do else

            { ALTERAÇÃO EM 20.02 }
            // quando Ftipo=0 (elic intra), escreve direto na célula do grid. Sem utilizar AuxTipo
            if FTipo[i] = 0 then
            begin
              grdCritInfo.Cell[i + 1, 3].Text := '---';
              TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
            end
            else
              grdCritInfo.Cell[i + 1, 3].Text := AuxTipo
                [TIWUserSession(WebApplication.Data).FTipo[i] - 1];

            /// //COLOCANDO TYPECRIT SEPARADO, DEVIDO QUESTÃO DE EDTS VAZIOS//////
            if (Typeofcrit[i] = 1) then // CONT MAX
            begin
              grdCritInfo.Cell[i + 1, 4].Text := 'Maximization';

              if (edtparA.Text = '') then
              begin
                grdCritInfo.Cell[i + 1, 5].Text := '---';
                if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                  TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
              end;

              if (edtparB.Text = '') then
              begin
                grdCritInfo.Cell[i + 1, 6].Text := '---';
                if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                  TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
              end // se vazio

              else
              begin
                grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
                grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
                if (parA[i] = 0) then
                  grdCritInfo.Cell[i + 1, 5].Text := '---';
                if (parB[i] = 0) then
                  grdCritInfo.Cell[i + 1, 6].Text := '---';
              end; // se possui valor
            end; // type=1

            if (Typeofcrit[i] = 3) then // DISC MAX
            begin
              grdCritInfo.Cell[i + 1, 4].Text := 'Maximization';
              grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
              grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
              if (parA[i] = 0) then
                grdCritInfo.Cell[i + 1, 5].Text := '---';
              if (parB[i] = 0) then
                grdCritInfo.Cell[i + 1, 6].Text := '---';
            end
            else

              if (Typeofcrit[i] = 0) then // CONT MIN
            begin
              grdCritInfo.Cell[i + 1, 4].Text := 'Minimization';

              if (edtparA.Text = '') then
              begin
                grdCritInfo.Cell[i + 1, 5].Text := '---';
                if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                  TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
              end;

              if (edtparB.Text = '') then
              begin
                grdCritInfo.Cell[i + 1, 6].Text := '---';
                if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                  TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
              end // se vazio

              else
              begin
                grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
                grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
                if (parA[i] = 0) then
                  grdCritInfo.Cell[i + 1, 5].Text := '---';
                if (parB[i] = 0) then
                  grdCritInfo.Cell[i + 1, 6].Text := '---';
              end; // se possui valor
            end; // type= 0

            if (Typeofcrit[i] = 2) then // DISC MIN
            begin
              grdCritInfo.Cell[i + 1, 4].Text := 'Minimization';
              grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
              grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
              if (parA[i] = 0) then
                grdCritInfo.Cell[i + 1, 5].Text := '---';
              if (parB[i] = 0) then
                grdCritInfo.Cell[i + 1, 6].Text := '---';
            end;
          end;
        // fim alteração 20.02.22

        { if (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 1) OR
          (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 3) then
          begin
          grdCritInfo.Cell[i + 1, 4].Text := 'Maximization';
          grdCritInfo.Cell[i + 1, 5].Text := FloatToStr
          (TIWUserSession(WebApplication.Data).parA[i]);
          grdCritInfo.Cell[i + 1, 6].Text := FloatToStr
          (TIWUserSession(WebApplication.Data).parB[i]);
          if (TIWUserSession(WebApplication.Data).parA[i] = 0) then
          grdCritInfo.Cell[i + 1, 5].Text := '---';
          if (TIWUserSession(WebApplication.Data).parB[i] = 0) then
          grdCritInfo.Cell[i + 1, 6].Text := '---';
          end // end do  if (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 1)...
          else if (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 0) OR
          (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 2) then
          begin
          grdCritInfo.Cell[i + 1, 4].Text := 'Minimization';
          grdCritInfo.Cell[i + 1, 5].Text := FloatToStr
          (TIWUserSession(WebApplication.Data).parA[i]);
          grdCritInfo.Cell[i + 1, 6].Text := FloatToStr
          (TIWUserSession(WebApplication.Data).parB[i]);
          if (TIWUserSession(WebApplication.Data).parA[i] = 0) then
          grdCritInfo.Cell[i + 1, 5].Text := '---';
          if (TIWUserSession(WebApplication.Data).parB[i] = 0) then
          grdCritInfo.Cell[i + 1, 6].Text := '---';
          end; // end do else
          end; // end do for I := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do

          // criando o grid com as alternativas
          grdinformaalt.ColumnCount := 2;
          grdinformaalt.RowCount := TIWUserSession(WebApplication.Data).NumAlt;
          for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
          begin
          grdinformaalt.Cell[i, 0].Text := 'Alt' + ' ' + inttostr(i + 1);
          grdinformaalt.Cell[i, 1].Text := TIWUserSession(WebApplication.Data)
          .NomeAlt[i];
          end; }// referente a for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do

      end; // end do -1

      if TIWUserSession(WebApplication.Data).estagio = 6 then
      begin
        rgnLoginproject.Visible := false;
        rgnConseqvalues.Visible := true;
        TIWUserSession(WebApplication.Data).Nomeprojeto := TIWUserSession
          (WebApplication.Data).auxnome;
        // procedimento para retornar dados da elicitação
        TIWUserSession(WebApplication.Data).AuxsalvarBD := 31;
        SalvarBD;

        with TIWUserSession(WebApplication.Data) do
        begin
          Setlength(ArrayAlt, TIWUserSession(WebApplication.Data).NumAlt);
          for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
          begin
            ArrayAlt[i] := TIWLabel.Create(Self);
            ArrayAlt[i].parent := rgnconsman;
            ArrayAlt[i].Text := ('Alt' + inttostr(i + 1));
            ArrayAlt[i].Left := (5);
            ArrayAlt[i].Top := (50 + (i * 25));
            ArrayAlt[i].Font.Style := [fsbold];
          end;

          Setlength(ArrayCrt, TIWUserSession(WebApplication.Data).NumCrit);
          for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
          begin
            ArrayCrt[i] := TIWLabel.Create(Self);
            ArrayCrt[i].parent := rgnconsman;
            ArrayCrt[i].Text := ('C' + inttostr(i + 1));
            ArrayCrt[i].Left := (45 + (i * 60));
            ArrayCrt[i].Top := (20);
            ArrayCrt[i].Font.Style := [fsbold];
          end;
        end;

        Setlength(TIWUserSession(WebApplication.Data).MatEdit, TIWUserSession
            (WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data)
            .NumCrit);
        // alteração feita em 21.01.22 para permitir que retomada manual com info parcial ocorra corretamente
        Setlength(TIWUserSession(WebApplication.Data).Verificaintra,
          TIWUserSession(WebApplication.Data).NumCrit);

        for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
        begin
          for j := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do

          begin
            // Criando os edits que recaberão os desempenhos a partir de uma Matriz de TIWEdit's
            TIWUserSession(WebApplication.Data).MatEdit[i][j] := TIWEdit.Create
              (Self);
            TIWUserSession(WebApplication.Data).MatEdit[i][j].parent :=
              rgnconsman;
            if TIWUserSession(WebApplication.Data).ConsMatrix[i, j] <> 0 then
              TIWUserSession(WebApplication.Data).MatEdit[i][j].Text :=
                FloatToStr(TIWUserSession(WebApplication.Data)
                  .ConsMatrix[i, j]);
            if TIWUserSession(WebApplication.Data).ConsMatrix[i, j] = 0 then
              TIWUserSession(WebApplication.Data).MatEdit[i][j].Text := '';
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Left :=
              (45 + (j * 60));
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Top :=
              (25 + (i * 25) + 23);
            TIWUserSession(WebApplication.Data).MatEdit[i][j].Width := 60;
          end;
        end;
        // cmproblematica1.ItemIndex := -1;  // Não limpo o nome caso o usuário volte para essa página

        AuxTipo[0] := 'Linear';
        AuxTipo[1] := 'Exponential';
        AuxTipo[2] := 'Logarithmic';
        AuxTipo[3] := 'Logistic';


        // Alteração feita em 21.01.22//
        // ** Preenchendo a tabela de informações dos critérios **//

        grdCritInfo.ColumnCount := 7;
        grdCritInfo.RowCount := TIWUserSession(WebApplication.Data).NumCrit + 1;

        grdCritInfo.Cell[0, 0].Text := 'Name';
        grdCritInfo.Cell[0, 1].Text := 'Scale';
        grdCritInfo.Cell[0, 2].Text := 'Number of levels';
        grdCritInfo.Cell[0, 3].Text := 'Value Function';
        grdCritInfo.Cell[0, 4].Text := 'Direction';
        grdCritInfo.Cell[0, 5].Text := 'Par a';
        grdCritInfo.Cell[0, 6].Text := 'Par b';

        with TIWUserSession(WebApplication.Data) do
          for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
          begin
            grdCritInfo.Cell[i + 1, 0].Text := TIWUserSession
              (WebApplication.Data).NomeCrit[i];
            if TIWUserSession(WebApplication.Data).Typeofcrit[i] > 1 then
            begin
              grdCritInfo.Cell[i + 1, 1].Text := 'Discrete';
              grdCritInfo.Cell[i + 1, 2].Text := inttostr
                (TIWUserSession(WebApplication.Data).Niveis[i]);
            end // end do   if TIWUserSession(WebApplication.Data).Typeofcrit[I] > 1
            else
            begin
              grdCritInfo.Cell[i + 1, 1].Text := 'Continuous';
              grdCritInfo.Cell[i + 1, 2].Text := '---';
            end; // end do else

            // quando Ftipo=0 (elic intra), escreve direto na célula do grid. Sem utilizar AuxTipo
            if FTipo[i] = 0 then
            begin
              grdCritInfo.Cell[i + 1, 3].Text := '---';
              TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
            end
            else
              grdCritInfo.Cell[i + 1, 3].Text := AuxTipo
                [TIWUserSession(WebApplication.Data).FTipo[i] - 1];

            /// //COLOCANDO TYPECRIT SEPARADO, DEVIDO QUESTÃO DE EDTS VAZIOSS//////
            if (Typeofcrit[i] = 1) then // CONT MAX
            begin
              grdCritInfo.Cell[i + 1, 4].Text := 'Maximization';

              if (edtparA.Text = '') then
              begin
                grdCritInfo.Cell[i + 1, 5].Text := '---';
                if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                  TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
              end;

              if (edtparB.Text = '') then
              begin
                grdCritInfo.Cell[i + 1, 6].Text := '---';
                if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                  TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
              end // se vazio

              else
              begin
                grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
                grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
                if (parA[i] = 0) then
                  grdCritInfo.Cell[i + 1, 5].Text := '---';
                if (parB[i] = 0) then
                  grdCritInfo.Cell[i + 1, 6].Text := '---';
              end; // se possui valor
            end; // type=1

            if (Typeofcrit[i] = 3) then // DISC MAX
            begin
              grdCritInfo.Cell[i + 1, 4].Text := 'Maximization';
              grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
              grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
              if (parA[i] = 0) then
                grdCritInfo.Cell[i + 1, 5].Text := '---';
              if (parB[i] = 0) then
                grdCritInfo.Cell[i + 1, 6].Text := '---';
            end
            else

              if (Typeofcrit[i] = 0) then // CONT MIN
            begin
              grdCritInfo.Cell[i + 1, 4].Text := 'Minimization';

              if (edtparA.Text = '') then
              begin
                grdCritInfo.Cell[i + 1, 5].Text := '---';
                if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                  TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
              end;

              if (edtparB.Text = '') then
              begin
                grdCritInfo.Cell[i + 1, 6].Text := '---';
                if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
                  TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
              end // se vazio

              else
              begin
                grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
                grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
                if (parA[i] = 0) then
                  grdCritInfo.Cell[i + 1, 5].Text := '---';
                if (parB[i] = 0) then
                  grdCritInfo.Cell[i + 1, 6].Text := '---';
              end; // se possui valor
            end; // type= 0

            if (Typeofcrit[i] = 2) then // DISC MIN
            begin
              grdCritInfo.Cell[i + 1, 4].Text := 'Minimization';
              grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
              grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
              if (parA[i] = 0) then
                grdCritInfo.Cell[i + 1, 5].Text := '---';
              if (parB[i] = 0) then
                grdCritInfo.Cell[i + 1, 6].Text := '---';
            end;
          end;

        { grdCritInfo.Cell[i + 1, 3].Text := AuxTipo
          [TIWUserSession(WebApplication.Data).FTipo[i] - 1];

          if (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 1) OR
          (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 3) then
          begin
          grdCritInfo.Cell[i + 1, 4].Text := 'Maximization';
          grdCritInfo.Cell[i + 1, 5].Text := FloatToStr
          (TIWUserSession(WebApplication.Data).parA[i]);
          grdCritInfo.Cell[i + 1, 6].Text := FloatToStr
          (TIWUserSession(WebApplication.Data).parB[i]);
          if (TIWUserSession(WebApplication.Data).parA[i] = 0) then
          grdCritInfo.Cell[i + 1, 5].Text := '---';
          if (TIWUserSession(WebApplication.Data).parB[i] = 0) then
          grdCritInfo.Cell[i + 1, 6].Text := '---';
          end // end do  if (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 1)...
          else if (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 0) OR
          (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 2) then
          begin
          grdCritInfo.Cell[i + 1, 4].Text := 'Minimization';
          grdCritInfo.Cell[i + 1, 5].Text := FloatToStr
          (TIWUserSession(WebApplication.Data).parA[i]);
          grdCritInfo.Cell[i + 1, 6].Text := FloatToStr
          (TIWUserSession(WebApplication.Data).parB[i]);
          if (TIWUserSession(WebApplication.Data).parA[i] = 0) then
          grdCritInfo.Cell[i + 1, 5].Text := '---';
          if (TIWUserSession(WebApplication.Data).parB[i] = 0) then
          grdCritInfo.Cell[i + 1, 6].Text := '---';
          end; // end do else
          end; // end do for I := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
          } // fim alteração 21.01.22

        // criando o grid com as alternativas
        grdinformaalt.ColumnCount := 2;
        grdinformaalt.RowCount := TIWUserSession(WebApplication.Data).NumAlt;
        for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
        begin
          grdinformaalt.Cell[i, 0].Text := 'Alt' + ' ' + inttostr(i + 1);
          grdinformaalt.Cell[i, 1].Text := TIWUserSession(WebApplication.Data)
            .NomeAlt[i];
        end; // referente a for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do

      end; // end do 6

      if (TIWUserSession(WebApplication.Data).estagio = 4) or
        (TIWUserSession(WebApplication.Data).estagio = 5) then
      begin
        // procedimento para pegar o round, caracteristicas dos criterios e das alternativas
        TIWUserSession(WebApplication.Data).AuxsalvarBD := 15;
        SalvarBD;
        // procedimento para retornar dados dos criterios
        TIWUserSession(WebApplication.Data).AuxsalvarBD := 33;
        SalvarBD;

        // procedimento para retornar dados da AH
        TIWUserSession(WebApplication.Data).AuxsalvarBD := 19;
        SalvarBD;

        // procedimento para retornar dados da elicitação
        TIWUserSession(WebApplication.Data).AuxsalvarBD := 18;
        SalvarBD;

        // procedimento para retornar dados da elicitação
        TIWUserSession(WebApplication.Data).AuxsalvarBD := 21;
        SalvarBD;

        // procedimento para retornar dados da avaliação holistica
        TIWUserSession(WebApplication.Data).AuxsalvarBD := 29;
        SalvarBD;

        /// ////////////////////////////////////////////////////////////////////////////////////////
        // preenchimento de variaveis que vão depender da resposta da heuristica e de outros fatores
        if (TIWUserSession(WebApplication.Data).HeurQuest = 0) or
          (TIWUserSession(WebApplication.Data).HeurQuest = 2) then
        begin
          // se a resposta da heuristica for A ou indiferente
          TIWUserSession(WebApplication.Data).Etapa := 2;
        end; // referente a  if (TIWUserSession(WebApplication.Data).HeurQuest=0) or (TIWUserSession(WebApplication.Data).HeurQuest=2) then

        if TIWUserSession(WebApplication.Data).HeurQuest = 1 then
        begin
          // se a resposta da heuristica for igual a B

          if TIWUserSession(WebApplication.Data).NumQuest =
            (TIWUserSession(WebApplication.Data).NumCrit - TIWUserSession
              (WebApplication.Data).NumCriBlind) then
          begin
            TIWUserSession(WebApplication.Data).Etapa := 2;
          end; // referente a if TIWUserSession(WebApplication.Data).NumQuest=(TIWUserSession(WebApplication.Data).NumCrit-TIWUserSession(WebApplication.Data).NumCriBlind) then

          if TIWUserSession(WebApplication.Data).NumQuest <
            (Round(0.5 * (TIWUserSession(WebApplication.Data)
                  .NumCrit - TIWUserSession(WebApplication.Data).NumCriBlind)))
            then
          begin
            TIWUserSession(WebApplication.Data).Etapa := 1;
            TIWUserSession(WebApplication.Data).Fase := 1;
          end // referente a if TIWUserSession(WebApplication.Data).NumQuest < (round(0.5*(TIWUserSession(WebApplication.Data).NumCrit - TIWUserSession(WebApplication.Data).NumCriBlind)))then
          else
          begin
            TIWUserSession(WebApplication.Data).Etapa := 1;
            TIWUserSession(WebApplication.Data).Fase := 2;
          end;
        end; // referente a if TIWUserSession(WebApplication.Data).HeurQuest=1 then

        /// ///////////////////////////////////////////////////////////////////////////////////////////////////

        { if TIWUserSession(WebApplication.Data).auxparacasonsaieretone=-1002525 then
          TIWUserSession(WebApplication.Data).stage:=0; }

        TIWUserSession(WebApplication.Data).RETPROBDEPOIS := -4555;
        Inicializa;
        // TIWUserSession(WebApplication.Data).stage := 3;
        TIWUserSession(WebApplication.Data).epsilon := trocapv(edtMaxdif.Text);
        rgnLoginproject.Visible := false;



        // Verificação para passar ou não direto para a unit de resultados
        if (TIWUserSession(WebApplication.Data).elicitarperfis = true) and
          (TIWUserSession(WebApplication.Data).Problematica = 3) then
        begin
          Tfrmaintercriterio.Create(WebApplication).btnprofileelicitClick(Self);
        end
        else
        begin
          Tfrmaintercriterio.Create(WebApplication).show;

        // Verificação para passar ou não direto para a unit de resultados
        { if TIWUserSession(WebApplication.Data).Sol <> -5 then
          begin }
        TIWUserSession(WebApplication.Data).RETPROBDEPOIS := -4555;
        Tfrmaintercriterio.Create(WebApplication).BtnshowresultsClick(Self);
        // end;

        end;

        // Verificação para passar ou não direto para a unit de resultados
        if TIWUserSession(WebApplication.Data).Sol = -5 then
        begin
          TIWUserSession(WebApplication.Data).RETPROBDEPOIS := -4555;
          Tfrmaintercriterio.Create(WebApplication).btnstopelicitationClick
            (Self);
        end;

      end; // referente a if (TIWUserSession(WebApplication.Data).estagio = 5) or (TIWUserSession(WebApplication.Data).estagio = 4) then

    end; // if  TIWUserSession(WebApplication.Data).Numcrit =< 10 then
    if TIWUserSession(WebApplication.Data).NumCrit > 10 then
    begin
      TIWUserSession(WebApplication.Data).tempinicio := TimeToStr(Now);
      // procedimento para pegar o round, caracteristicas dos criterios e das alternativas
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 15;
      SalvarBD;
      rgnInputMatrizConseq.Visible := true;
      // **Deixa visível a opção para seguir pra elicitação direto **//
      AlreadyRakend1.Visible := true;
      rgnLoginproject.Visible := false;

      // procedimento para preencher as informações na região rgnInputMatrizConseq

      Inicializa;
      if TIWUserSession(WebApplication.Data).Flag <> -2 then
      begin
        with TIWUserSession(WebApplication.Data) do
        begin
          /// / Setando o tamanho dos grids que exibem a informação importada na região imput ////
          grdMatrizConseq.RowCount := NumAlt + 7;
          grdMatrizConseq.ColumnCount := NumCrit + 1;
          probname := edtEntername.Text;

          /// / Preenchendo a informação fixa dos grids ////
          grdMatrizConseq.Cell[0, 0].Text := 'Criteria:';
          grdMatrizConseq.Cell[1, 0].Text :=
            '0-Cont Min; 1-Cont Max; 2-Disc Min; 3- Disc Max; 4- Int Min; 5- Int Max:';
          { grdInputvalues.Cell[2, 0].text := 'Weights:'; }
          grdMatrizConseq.Cell[2, 0].Text := '';
          grdMatrizConseq.Cell[3, 0].Text := '';
          grdMatrizConseq.Cell[4, 0].Text := '';
          grdMatrizConseq.Cell[5, 0].Text :=
            'Number of levels of discrete criteria';
          grdMatrizConseq.Cell[6, 0].Text := 'Consequence Matrix:';

          /// / Preenchendo os Grids com a informação Imputada ///////////////////////
          for j := 0 to NumCrit - 1 do
          begin
            grdMatrizConseq.Cell[0, j + 1].Text := NomeCrit[j];
            if Cri_inteiro[j] then
            begin
              if Typeofcrit[j] = 0 then
                grdMatrizConseq.Cell[1, j + 1].Text := '4';
              if Typeofcrit[j] = 1 then
                grdMatrizConseq.Cell[1, j + 1].Text := '5';
            end
            else
              grdMatrizConseq.Cell[1, j + 1].Text := inttostr(Typeofcrit[j]);
            { grdMatrizConseq.Cell[2, j + 1].Text := inttostr(FTipo[j]);
              if FTipo[j] <> 1 then
              begin
              grdMatrizConseq.Cell[3, j + 1].Text := FloatToStr(parA[j]);
              grdMatrizConseq.Cell[4, j + 1].Text := FloatToStr(parB[j]);
              end; }
            grdMatrizConseq.Cell[5, j + 1].Text := inttostr(Niveis[j]);
          end;

          for i := 0 to NumAlt - 1 do
          begin
            grdMatrizConseq.Cell[i + 7, 0].Text := NomeAlt[i];
            for j := 0 to NumCrit - 1 do
            begin
              grdMatrizConseq.Cell[i + 7, j + 1].Text := TrocaVP
                (ConsMatrix[i, j]);
            end;
          end;
          /// ////////////////////////////////////////////////////////////////////////
          // end;
        end;
      end;
      TIWUserSession(WebApplication.Data).mudadata := -3254;
      mensageretorno := -90998;
      // WebApplication.ShowMessage('Sorry for the inconvenience, the database has been modified and it is not possible to return from where you left off');
    end; // referente a  if TIWUserSession(WebApplication.Data).Numcrit > 1
  end;

  // if (StrToDate(dtinicretorn) < StrToDate('19/04/2021')) AND (StrToDate(dtfnretorn)=StrToDa





  // retornando quando começou o problema mas não finalizou
  { if (StrToDate(TIWUserSession(WebApplication.Data).dtinicretorn) < StrToDate('19/04/2021')) AND (StrToDate(TIWUserSession(WebApplication.Data).dtfnretorn) = StrToDate('02/01/1900')) then
    begin
    // procedimento para retornar o número de critérios
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 32;
    SalvarBD;
    if  TIWUserSession(WebApplication.Data).Numcrit <= 10 then
    begin
    if (TIWUserSession(WebApplication.Data).estagio = 0) or
    (TIWUserSession(WebApplication.Data).estagio = 1) then
    begin
    TIWUserSession(WebApplication.Data).tempinicio := TimeToStr(Now);

    // procedimento para pegar o round, caracteristicas dos criterios e das alternativas
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 15;
    SalvarBD;
    rgnInputMatrizConseq.Visible := true;
    rgnLoginproject.Visible := false;

    // procedimento para preencher as informações na região rgnInputMatrizConseq

    Inicializa;
    if TIWUserSession(WebApplication.Data).Flag <> -2 then
    begin
    with TIWUserSession(WebApplication.Data) do
    begin
    /// / Setando o tamanho dos grids que exibem a informação importada na região imput ////
    grdInputvalues.RowCount := 6;
    grdInputvalues.ColumnCount := NumCrit + 1;
    grdMatrizConseq.RowCount := NumAlt;
    grdMatrizConseq.ColumnCount := NumCrit + 1;
    probname := edtEntername.text;

    /// / Preenchendo a informação fixa dos grids ////
    grdInputvalues.Cell[0, 0].text := 'Criteria:';
    grdInputvalues.Cell[1, 0].text :=
    '0-Cont Min; 1-Cont Max; 2-Disc Min; 3- Disc Max:';
    { grdInputvalues.Cell[2, 0].text := 'Weights:'; }
  { grdInputvalues.Cell[2, 0].text := 'Type:';
    grdInputvalues.Cell[3, 0].text := 'a:';
    grdInputvalues.Cell[4, 0].text := 'b:';
    grdInputvalues.Cell[5, 0].text :=
    'c: number of levels of discrete criteria';

    /// / Preenchendo os Grids com a informação Imputada ///////////////////////
    for J := 0 to NumCrit - 1 do
    begin
    lstbxInitialOrder.Items.Add(NomeCrit[J]);
    grdInputvalues.Cell[0, J + 1].text := NomeCrit[J];
    grdInputvalues.Cell[1, J + 1].text := inttostr(Typeofcrit[J]);
    grdInputvalues.Cell[2, J + 1].text := inttostr(Ftipo[J]);
    if Ftipo[J] <> 1 then
    begin
    grdInputvalues.Cell[3, J + 1].text := FloatToStr(parA[J]);
    grdInputvalues.Cell[4, J + 1].text := FloatToStr(parB[J]);
    end;
    grdInputvalues.Cell[5, J + 1].text := inttostr(Niveis[J]);
    end;

    for I := 0 to NumAlt - 1 do
    begin
    grdMatrizConseq.Cell[I, 0].text := NomeAlt[I];
    for J := 0 to NumCrit - 1 do
    begin
    grdMatrizConseq.Cell[I, J + 1].text := TrocaVP(ConsMatrix[I, J]);
    end;
    end;
    /// ////////////////////////////////////////////////////////////////////////
    // end;
    end;
    end;
    end;

    if (TIWUserSession(WebApplication.Data).estagio = 2) or
    (TIWUserSession(WebApplication.Data).estagio = 3) then
    begin
    // procedimento para pegar o round, caracteristicas dos criterios e das alternativas
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 15;
    SalvarBD;
    // procedimento para retornar dados dos criterios
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 33;
    SalvarBD;

    // procedimento para retornar dados da AH
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 19;
    SalvarBD;

    // procedimento para retornar dados da elicitação
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 18;
    SalvarBD;

    // procedimento para retornar dados da elicitação
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 21;
    SalvarBD;

    // procedimento para retornar dados da avaliação holistica
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 29;
    SalvarBD;

    /// ////////////////////////////////////////////////////////////////////////////////////////
    // preenchimento de variaveis que vão depender da resposta da heuristica e de outros fatores
    if (TIWUserSession(WebApplication.Data).HeurQuest = 0) or
    (TIWUserSession(WebApplication.Data).HeurQuest = 2) then
    begin
    // se a resposta da heuristica for A ou indiferente
    TIWUserSession(WebApplication.Data).Etapa := 2;
    end; // referente a  if (TIWUserSession(WebApplication.Data).HeurQuest=0) or (TIWUserSession(WebApplication.Data).HeurQuest=2) then

    if TIWUserSession(WebApplication.Data).HeurQuest = 1 then
    begin
    // se a resposta da heuristica for igual a B

    if TIWUserSession(WebApplication.Data).NumQuest =
    (TIWUserSession(WebApplication.Data).NumCrit - TIWUserSession
    (WebApplication.Data).NumCriBlind) then
    begin
    TIWUserSession(WebApplication.Data).Etapa := 2;
    end; // referente a if TIWUserSession(WebApplication.Data).NumQuest=(TIWUserSession(WebApplication.Data).NumCrit-TIWUserSession(WebApplication.Data).NumCriBlind) then

    if TIWUserSession(WebApplication.Data).NumQuest <
    (Round(0.5 * (TIWUserSession(WebApplication.Data)
    .NumCrit - TIWUserSession(WebApplication.Data).NumCriBlind))) then
    begin
    TIWUserSession(WebApplication.Data).Etapa := 1;
    TIWUserSession(WebApplication.Data).Fase := 1;
    end // referente a if TIWUserSession(WebApplication.Data).NumQuest < (round(0.5*(TIWUserSession(WebApplication.Data).NumCrit - TIWUserSession(WebApplication.Data).NumCriBlind)))then
    else
    begin
    TIWUserSession(WebApplication.Data).Etapa := 1;
    TIWUserSession(WebApplication.Data).Fase := 2;
    end;
    end; // referente a if TIWUserSession(WebApplication.Data).HeurQuest=1 then

    /// ///////////////////////////////////////////////////////////////////////////////////////////////////

    { if TIWUserSession(WebApplication.Data).auxparacasonsaieretone=-1002525 then
    TIWUserSession(WebApplication.Data).stage:=0; }

  { TIWUserSession(WebApplication.Data).RETPROBDEPOIS := -4555;
    Inicializa;
    // TIWUserSession(WebApplication.Data).stage := 3;
    TIWUserSession(WebApplication.Data).epsilon := trocapv(edtMaxdif.text);
    rgnLoginproject.Visible := false;
    Tfrmaintercriterio.Create(WebApplication).Show;

    // Verificação para passar ou não direto para a unit de resultados (caso não seja necessário realizar a elicitção flexível)
    if TIWUserSession(WebApplication.Data).Sol <> -5 then
    begin
    Tfrmaintercriterio.Create(WebApplication).BtnshowresultsClick(Self);
    end;

    end; // referente a if (TIWUserSession(WebApplication.Data).estagio = 2) or (TIWUserSession(WebApplication.Data).estagio = 3) then

    if TIWUserSession(WebApplication.Data).estagio = -1 then
    begin
    rgnLoginproject.Visible := false;
    rgnConseqvalues.Visible := true;
    TIWUserSession(WebApplication.Data).Nomeprojeto := TIWUserSession(WebApplication.Data).auxnome;
    // procedimento para retornar dados da elicitação
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 25;
    SalvarBD;

    with TIWUserSession(WebApplication.Data) do
    begin
    Setlength(ArrayAlt, TIWUserSession(WebApplication.Data).NumAlt);
    for I := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
    begin
    ArrayAlt[I] := TIWLabel.Create(Self);
    ArrayAlt[I].parent := rgnconsman;
    ArrayAlt[I].text := ('Alt' + inttostr(I + 1));
    ArrayAlt[I].Left := (5);
    ArrayAlt[I].Top := (50 + (I * 25));
    ArrayAlt[I].Font.Style := [fsbold];
    end;

    Setlength(ArrayCrt, TIWUserSession(WebApplication.Data).NumCrit);
    for I := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
    begin
    ArrayCrt[I] := TIWLabel.Create(Self);
    ArrayCrt[I].parent := rgnconsman;
    ArrayCrt[I].text := ('C' + inttostr(I + 1));
    ArrayCrt[I].Left := (45 + (I * 60));
    ArrayCrt[I].Top := (20);
    ArrayCrt[I].Font.Style := [fsbold];
    end;
    end;

    Setlength(TIWUserSession(WebApplication.Data).MatEdit, TIWUserSession
    (WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data)
    .NumCrit);
    for I := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
    begin
    for J := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do

    begin
    // Criando os edits que recaberão os desempenhos a partir de uma Matriz de TIWEdit's
    TIWUserSession(WebApplication.Data).MatEdit[I][J] := TIWEdit.Create
    (Self);
    TIWUserSession(WebApplication.Data).MatEdit[I][J].parent := rgnconsman;
    TIWUserSession(WebApplication.Data).MatEdit[I][J].text := ('');
    TIWUserSession(WebApplication.Data).MatEdit[I][J].Left :=
    (45 + (J * 60));
    TIWUserSession(WebApplication.Data).MatEdit[I][J].Top :=
    (25 + (I * 25) + 23);
    TIWUserSession(WebApplication.Data).MatEdit[I][J].Width := 60;
    end;
    end;
    // cmproblematica1.ItemIndex := -1;  // Não limpo o nome caso o usuário volte para essa página

    AuxTipo[0] := 'Linear';
    AuxTipo[1] := 'Exponential';
    AuxTipo[2] := 'Logarithmic';
    AuxTipo[3] := 'Logistic';
    // ** Preenchendo a tabela de informações dos critérios **//

    grdCritInfo.ColumnCount := 7;
    grdCritInfo.RowCount := TIWUserSession(WebApplication.Data).NumCrit + 1;

    grdCritInfo.Cell[0, 0].text := 'Name';
    grdCritInfo.Cell[0, 1].text := 'Scale';
    grdCritInfo.Cell[0, 2].text := 'Number of levels';
    grdCritInfo.Cell[0, 3].text := 'Value Function';
    grdCritInfo.Cell[0, 4].text := 'Direction';
    grdCritInfo.Cell[0, 5].text := 'Par a';
    grdCritInfo.Cell[0, 6].text := 'Par b';

    for I := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
    begin
    grdCritInfo.Cell[I + 1, 0].text := TIWUserSession(WebApplication.Data)
    .NomeCrit[I];
    if TIWUserSession(WebApplication.Data).Typeofcrit[I] > 1 then
    begin
    grdCritInfo.Cell[I + 1, 1].text := 'Discrete';
    grdCritInfo.Cell[I + 1, 2].text := inttostr
    (TIWUserSession(WebApplication.Data).Niveis[I]);
    end // end do   if TIWUserSession(WebApplication.Data).Typeofcrit[I] > 1
    else
    begin
    grdCritInfo.Cell[I + 1, 1].text := 'Continuous';
    grdCritInfo.Cell[I + 1, 2].text := '---';
    end; // end do else
    grdCritInfo.Cell[I + 1, 3].text := AuxTipo
    [TIWUserSession(WebApplication.Data).Ftipo[I] - 1];

    if (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 1) OR
    (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 3) then
    begin
    grdCritInfo.Cell[I + 1, 4].text := 'Maximization';
    grdCritInfo.Cell[I + 1, 5].text := FloatToStr
    (TIWUserSession(WebApplication.Data).parA[I]);
    grdCritInfo.Cell[I + 1, 6].text := FloatToStr
    (TIWUserSession(WebApplication.Data).parB[I]);
    if (TIWUserSession(WebApplication.Data).parA[I] = 0) then
    grdCritInfo.Cell[I + 1, 5].text := '---';
    if (TIWUserSession(WebApplication.Data).parB[I] = 0) then
    grdCritInfo.Cell[I + 1, 6].text := '---';
    end // end do  if (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 1)...
    else if (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 0) OR
    (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 2) then
    begin
    grdCritInfo.Cell[I + 1, 4].text := 'Minimization';
    grdCritInfo.Cell[I + 1, 5].text := FloatToStr
    (TIWUserSession(WebApplication.Data).parA[I]);
    grdCritInfo.Cell[I + 1, 6].text := FloatToStr
    (TIWUserSession(WebApplication.Data).parB[I]);
    if (TIWUserSession(WebApplication.Data).parA[I] = 0) then
    grdCritInfo.Cell[I + 1, 5].text := '---';
    if (TIWUserSession(WebApplication.Data).parB[I] = 0) then
    grdCritInfo.Cell[I + 1, 6].text := '---';
    end; // end do else
    end; // end do for I := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
    end; // end do -1


    if TIWUserSession(WebApplication.Data).estagio = 6 then
    begin
    rgnLoginproject.Visible := false;
    rgnConseqvalues.Visible := true;
    TIWUserSession(WebApplication.Data).Nomeprojeto := TIWUserSession(WebApplication.Data).auxnome;
    // procedimento para retornar dados da elicitação
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 31;
    SalvarBD;

    with TIWUserSession(WebApplication.Data) do
    begin
    Setlength(ArrayAlt, TIWUserSession(WebApplication.Data).NumAlt);
    for I := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
    begin
    ArrayAlt[I] := TIWLabel.Create(Self);
    ArrayAlt[I].parent := rgnconsman;
    ArrayAlt[I].text := ('Alt' + inttostr(I + 1));
    ArrayAlt[I].Left := (5);
    ArrayAlt[I].Top := (50 + (I * 25));
    ArrayAlt[I].Font.Style := [fsbold];
    end;

    Setlength(ArrayCrt, TIWUserSession(WebApplication.Data).NumCrit);
    for I := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
    begin
    ArrayCrt[I] := TIWLabel.Create(Self);
    ArrayCrt[I].parent := rgnconsman;
    ArrayCrt[I].text := ('C' + inttostr(I + 1));
    ArrayCrt[I].Left := (45 + (I * 60));
    ArrayCrt[I].Top := (20);
    ArrayCrt[I].Font.Style := [fsbold];
    end;
    end;

    Setlength(TIWUserSession(WebApplication.Data).MatEdit, TIWUserSession
    (WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data)
    .NumCrit);
    for I := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
    begin
    for J := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do

    begin
    // Criando os edits que recaberão os desempenhos a partir de uma Matriz de TIWEdit's
    TIWUserSession(WebApplication.Data).MatEdit[I][J] := TIWEdit.Create
    (Self);
    TIWUserSession(WebApplication.Data).MatEdit[I][J].parent := rgnconsman;
    if TIWUserSession(WebApplication.Data).ConsMatrix[I,J]<>0 then
    TIWUserSession(WebApplication.Data).MatEdit[I][J].text := Floattostr(TIWUserSession(WebApplication.Data).ConsMatrix[I,J]);
    if TIWUserSession(WebApplication.Data).ConsMatrix[I,J]=0 then
    TIWUserSession(WebApplication.Data).MatEdit[I][J].text := '';
    TIWUserSession(WebApplication.Data).MatEdit[I][J].Left :=
    (45 + (J * 60));
    TIWUserSession(WebApplication.Data).MatEdit[I][J].Top :=
    (25 + (I * 25) + 23);
    TIWUserSession(WebApplication.Data).MatEdit[I][J].Width := 60;
    end;
    end;
    // cmproblematica1.ItemIndex := -1;  // Não limpo o nome caso o usuário volte para essa página

    AuxTipo[0] := 'Linear';
    AuxTipo[1] := 'Exponential';
    AuxTipo[2] := 'Logarithmic';
    AuxTipo[3] := 'Logistic';
    // ** Preenchendo a tabela de informações dos critérios **//

    grdCritInfo.ColumnCount := 7;
    grdCritInfo.RowCount := TIWUserSession(WebApplication.Data).NumCrit + 1;

    grdCritInfo.Cell[0, 0].text := 'Name';
    grdCritInfo.Cell[0, 1].text := 'Scale';
    grdCritInfo.Cell[0, 2].text := 'Number of levels';
    grdCritInfo.Cell[0, 3].text := 'Value Function';
    grdCritInfo.Cell[0, 4].text := 'Direction';
    grdCritInfo.Cell[0, 5].text := 'Par a';
    grdCritInfo.Cell[0, 6].text := 'Par b';

    for I := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
    begin
    grdCritInfo.Cell[I + 1, 0].text := TIWUserSession(WebApplication.Data)
    .NomeCrit[I];
    if TIWUserSession(WebApplication.Data).Typeofcrit[I] > 1 then
    begin
    grdCritInfo.Cell[I + 1, 1].text := 'Discrete';
    grdCritInfo.Cell[I + 1, 2].text := inttostr
    (TIWUserSession(WebApplication.Data).Niveis[I]);
    end // end do   if TIWUserSession(WebApplication.Data).Typeofcrit[I] > 1
    else
    begin
    grdCritInfo.Cell[I + 1, 1].text := 'Continuous';
    grdCritInfo.Cell[I + 1, 2].text := '---';
    end; // end do else
    grdCritInfo.Cell[I + 1, 3].text := AuxTipo
    [TIWUserSession(WebApplication.Data).Ftipo[I] - 1];

    if (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 1) OR
    (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 3) then
    begin
    grdCritInfo.Cell[I + 1, 4].text := 'Maximization';
    grdCritInfo.Cell[I + 1, 5].text := FloatToStr
    (TIWUserSession(WebApplication.Data).parA[I]);
    grdCritInfo.Cell[I + 1, 6].text := FloatToStr
    (TIWUserSession(WebApplication.Data).parB[I]);
    if (TIWUserSession(WebApplication.Data).parA[I] = 0) then
    grdCritInfo.Cell[I + 1, 5].text := '---';
    if (TIWUserSession(WebApplication.Data).parB[I] = 0) then
    grdCritInfo.Cell[I + 1, 6].text := '---';
    end // end do  if (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 1)...
    else if (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 0) OR
    (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 2) then
    begin
    grdCritInfo.Cell[I + 1, 4].text := 'Minimization';
    grdCritInfo.Cell[I + 1, 5].text := FloatToStr
    (TIWUserSession(WebApplication.Data).parA[I]);
    grdCritInfo.Cell[I + 1, 6].text := FloatToStr
    (TIWUserSession(WebApplication.Data).parB[I]);
    if (TIWUserSession(WebApplication.Data).parA[I] = 0) then
    grdCritInfo.Cell[I + 1, 5].text := '---';
    if (TIWUserSession(WebApplication.Data).parB[I] = 0) then
    grdCritInfo.Cell[I + 1, 6].text := '---';
    end; // end do else
    end; // end do for I := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
    end; // end do 6



    if (TIWUserSession(WebApplication.Data).estagio = 4) or
    (TIWUserSession(WebApplication.Data).estagio = 5) then
    begin
    // procedimento para pegar o round, caracteristicas dos criterios e das alternativas
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 15;
    SalvarBD;
    // procedimento para retornar dados dos criterios
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 33;
    SalvarBD;

    // procedimento para retornar dados da AH
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 19;
    SalvarBD;

    // procedimento para retornar dados da elicitação
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 18;
    SalvarBD;

    // procedimento para retornar dados da elicitação
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 21;
    SalvarBD;

    // procedimento para retornar dados da avaliação holistica
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 29;
    SalvarBD;

    /// ////////////////////////////////////////////////////////////////////////////////////////
    // preenchimento de variaveis que vão depender da resposta da heuristica e de outros fatores
    if (TIWUserSession(WebApplication.Data).HeurQuest = 0) or
    (TIWUserSession(WebApplication.Data).HeurQuest = 2) then
    begin
    // se a resposta da heuristica for A ou indiferente
    TIWUserSession(WebApplication.Data).Etapa := 2;
    end; // referente a  if (TIWUserSession(WebApplication.Data).HeurQuest=0) or (TIWUserSession(WebApplication.Data).HeurQuest=2) then

    if TIWUserSession(WebApplication.Data).HeurQuest = 1 then
    begin
    // se a resposta da heuristica for igual a B

    if TIWUserSession(WebApplication.Data).NumQuest =
    (TIWUserSession(WebApplication.Data).NumCrit - TIWUserSession
    (WebApplication.Data).NumCriBlind) then
    begin
    TIWUserSession(WebApplication.Data).Etapa := 2;
    end; // referente a if TIWUserSession(WebApplication.Data).NumQuest=(TIWUserSession(WebApplication.Data).NumCrit-TIWUserSession(WebApplication.Data).NumCriBlind) then

    if TIWUserSession(WebApplication.Data).NumQuest <
    (Round(0.5 * (TIWUserSession(WebApplication.Data)
    .NumCrit - TIWUserSession(WebApplication.Data).NumCriBlind))) then
    begin
    TIWUserSession(WebApplication.Data).Etapa := 1;
    TIWUserSession(WebApplication.Data).Fase := 1;
    end // referente a if TIWUserSession(WebApplication.Data).NumQuest < (round(0.5*(TIWUserSession(WebApplication.Data).NumCrit - TIWUserSession(WebApplication.Data).NumCriBlind)))then
    else
    begin
    TIWUserSession(WebApplication.Data).Etapa := 1;
    TIWUserSession(WebApplication.Data).Fase := 2;
    end;
    end; // referente a if TIWUserSession(WebApplication.Data).HeurQuest=1 then

    /// ///////////////////////////////////////////////////////////////////////////////////////////////////

    { if TIWUserSession(WebApplication.Data).auxparacasonsaieretone=-1002525 then
    TIWUserSession(WebApplication.Data).stage:=0; }

  { TIWUserSession(WebApplication.Data).RETPROBDEPOIS := -4555;
    Inicializa;
    // TIWUserSession(WebApplication.Data).stage := 3;
    TIWUserSession(WebApplication.Data).epsilon := trocapv(edtMaxdif.text);
    rgnLoginproject.Visible := false;
    Tfrmaintercriterio.Create(WebApplication).Show;

    // Verificação para passar ou não direto para a unit de resultados
    if TIWUserSession(WebApplication.Data).Sol <> -5 then
    begin
    TIWUserSession(WebApplication.Data).RETPROBDEPOIS := -4555;
    Tfrmaintercriterio.Create(WebApplication).BtnshowresultsClick(Self);
    end;

    // Verificação para passar ou não direto para a unit de resultados
    if TIWUserSession(WebApplication.Data).Sol = -5 then
    begin
    TIWUserSession(WebApplication.Data).RETPROBDEPOIS := -4555;
    Tfrmaintercriterio.Create(WebApplication).btnstopelicitationClick(Self);
    end;

    end; // referente a if (TIWUserSession(WebApplication.Data).estagio = 5) or (TIWUserSession(WebApplication.Data).estagio = 4) then

    end; // if  TIWUserSession(WebApplication.Data).Numcrit =< 10 then

    if TIWUserSession(WebApplication.Data).Numcrit > 10  then
    begin
    if (TIWUserSession(WebApplication.Data).estagio = 0) or
    (TIWUserSession(WebApplication.Data).estagio = 1) then
    begin
    TIWUserSession(WebApplication.Data).tempinicio := TimeToStr(Now);

    // procedimento para pegar o round, caracteristicas dos criterios e das alternativas
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 15;
    SalvarBD;
    rgnInputMatrizConseq.Visible := true;
    rgnLoginproject.Visible := false;

    // procedimento para preencher as informações na região rgnInputMatrizConseq

    Inicializa;
    if TIWUserSession(WebApplication.Data).Flag <> -2 then
    begin
    with TIWUserSession(WebApplication.Data) do
    begin
    /// / Setando o tamanho dos grids que exibem a informação importada na região imput ////
    grdInputvalues.RowCount := 6;
    grdInputvalues.ColumnCount := NumCrit + 1;
    grdMatrizConseq.RowCount := NumAlt;
    grdMatrizConseq.ColumnCount := NumCrit + 1;
    probname := edtEntername.text;

    /// / Preenchendo a informação fixa dos grids ////
    grdInputvalues.Cell[0, 0].text := 'Criteria:';
    grdInputvalues.Cell[1, 0].text :=
    '0-Cont Min; 1-Cont Max; 2-Disc Min; 3- Disc Max:';
    { grdInputvalues.Cell[2, 0].text := 'Weights:'; }
  { grdInputvalues.Cell[2, 0].text := 'Type:';
    grdInputvalues.Cell[3, 0].text := 'a:';
    grdInputvalues.Cell[4, 0].text := 'b:';
    grdInputvalues.Cell[5, 0].text :=
    'c: number of levels of discrete criteria';

    /// / Preenchendo os Grids com a informação Imputada ///////////////////////
    for J := 0 to NumCrit - 1 do
    begin
    lstbxInitialOrder.Items.Add(NomeCrit[J]);
    grdInputvalues.Cell[0, J + 1].text := NomeCrit[J];
    grdInputvalues.Cell[1, J + 1].text := inttostr(Typeofcrit[J]);
    grdInputvalues.Cell[2, J + 1].text := inttostr(Ftipo[J]);
    if Ftipo[J] <> 1 then
    begin
    grdInputvalues.Cell[3, J + 1].text := FloatToStr(parA[J]);
    grdInputvalues.Cell[4, J + 1].text := FloatToStr(parB[J]);
    end;
    grdInputvalues.Cell[5, J + 1].text := inttostr(Niveis[J]);
    end;

    for I := 0 to NumAlt - 1 do
    begin
    grdMatrizConseq.Cell[I, 0].text := NomeAlt[I];
    for J := 0 to NumCrit - 1 do
    begin
    grdMatrizConseq.Cell[I, J + 1].text := TrocaVP(ConsMatrix[I, J]);
    end;
    end;
    /// ////////////////////////////////////////////////////////////////////////
    // end;
    end;
    end;
    end;

    if TIWUserSession(WebApplication.Data).estagio = 6 then
    begin
    rgnLoginproject.Visible := false;
    rgnConseqvalues.Visible := true;
    TIWUserSession(WebApplication.Data).Nomeprojeto := TIWUserSession(WebApplication.Data).auxnome;
    // procedimento para retornar dados da elicitação
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 31;
    SalvarBD;

    with TIWUserSession(WebApplication.Data) do
    begin
    Setlength(ArrayAlt, TIWUserSession(WebApplication.Data).NumAlt);
    for I := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
    begin
    ArrayAlt[I] := TIWLabel.Create(Self);
    ArrayAlt[I].parent := rgnconsman;
    ArrayAlt[I].text := ('Alt' + inttostr(I + 1));
    ArrayAlt[I].Left := (5);
    ArrayAlt[I].Top := (50 + (I * 25));
    ArrayAlt[I].Font.Style := [fsbold];
    end;

    Setlength(ArrayCrt, TIWUserSession(WebApplication.Data).NumCrit);
    for I := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
    begin
    ArrayCrt[I] := TIWLabel.Create(Self);
    ArrayCrt[I].parent := rgnconsman;
    ArrayCrt[I].text := ('C' + inttostr(I + 1));
    ArrayCrt[I].Left := (45 + (I * 60));
    ArrayCrt[I].Top := (20);
    ArrayCrt[I].Font.Style := [fsbold];
    end;
    end;

    Setlength(TIWUserSession(WebApplication.Data).MatEdit, TIWUserSession
    (WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data)
    .NumCrit);
    for I := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
    begin
    for J := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do

    begin
    // Criando os edits que recaberão os desempenhos a partir de uma Matriz de TIWEdit's
    TIWUserSession(WebApplication.Data).MatEdit[I][J] := TIWEdit.Create
    (Self);
    TIWUserSession(WebApplication.Data).MatEdit[I][J].parent := rgnconsman;
    if TIWUserSession(WebApplication.Data).ConsMatrix[I,J]<>0 then
    TIWUserSession(WebApplication.Data).MatEdit[I][J].text := Floattostr(TIWUserSession(WebApplication.Data).ConsMatrix[I,J]);
    if TIWUserSession(WebApplication.Data).ConsMatrix[I,J]=0 then
    TIWUserSession(WebApplication.Data).MatEdit[I][J].text := '';
    TIWUserSession(WebApplication.Data).MatEdit[I][J].Left :=
    (45 + (J * 60));
    TIWUserSession(WebApplication.Data).MatEdit[I][J].Top :=
    (25 + (I * 25) + 23);
    TIWUserSession(WebApplication.Data).MatEdit[I][J].Width := 60;
    end;
    end;
    // cmproblematica1.ItemIndex := -1;  // Não limpo o nome caso o usuário volte para essa página

    AuxTipo[0] := 'Linear';
    AuxTipo[1] := 'Exponential';
    AuxTipo[2] := 'Logarithmic';
    AuxTipo[3] := 'Logistic';
    // ** Preenchendo a tabela de informações dos critérios **//

    grdCritInfo.ColumnCount := 7;
    grdCritInfo.RowCount := TIWUserSession(WebApplication.Data).NumCrit + 1;

    grdCritInfo.Cell[0, 0].text := 'Name';
    grdCritInfo.Cell[0, 1].text := 'Scale';
    grdCritInfo.Cell[0, 2].text := 'Number of levels';
    grdCritInfo.Cell[0, 3].text := 'Value Function';
    grdCritInfo.Cell[0, 4].text := 'Direction';
    grdCritInfo.Cell[0, 5].text := 'Par a';
    grdCritInfo.Cell[0, 6].text := 'Par b';

    for I := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
    begin
    grdCritInfo.Cell[I + 1, 0].text := TIWUserSession(WebApplication.Data)
    .NomeCrit[I];
    if TIWUserSession(WebApplication.Data).Typeofcrit[I] > 1 then
    begin
    grdCritInfo.Cell[I + 1, 1].text := 'Discrete';
    grdCritInfo.Cell[I + 1, 2].text := inttostr
    (TIWUserSession(WebApplication.Data).Niveis[I]);
    end // end do   if TIWUserSession(WebApplication.Data).Typeofcrit[I] > 1
    else
    begin
    grdCritInfo.Cell[I + 1, 1].text := 'Continuous';
    grdCritInfo.Cell[I + 1, 2].text := '---';
    end; // end do else
    grdCritInfo.Cell[I + 1, 3].text := AuxTipo
    [TIWUserSession(WebApplication.Data).Ftipo[I] - 1];

    if (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 1) OR
    (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 3) then
    begin
    grdCritInfo.Cell[I + 1, 4].text := 'Maximization';
    grdCritInfo.Cell[I + 1, 5].text := FloatToStr
    (TIWUserSession(WebApplication.Data).parA[I]);
    grdCritInfo.Cell[I + 1, 6].text := FloatToStr
    (TIWUserSession(WebApplication.Data).parB[I]);
    if (TIWUserSession(WebApplication.Data).parA[I] = 0) then
    grdCritInfo.Cell[I + 1, 5].text := '---';
    if (TIWUserSession(WebApplication.Data).parB[I] = 0) then
    grdCritInfo.Cell[I + 1, 6].text := '---';
    end // end do  if (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 1)...
    else if (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 0) OR
    (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 2) then
    begin
    grdCritInfo.Cell[I + 1, 4].text := 'Minimization';
    grdCritInfo.Cell[I + 1, 5].text := FloatToStr
    (TIWUserSession(WebApplication.Data).parA[I]);
    grdCritInfo.Cell[I + 1, 6].text := FloatToStr
    (TIWUserSession(WebApplication.Data).parB[I]);
    if (TIWUserSession(WebApplication.Data).parA[I] = 0) then
    grdCritInfo.Cell[I + 1, 5].text := '---';
    if (TIWUserSession(WebApplication.Data).parB[I] = 0) then
    grdCritInfo.Cell[I + 1, 6].text := '---';
    end; // end do else
    end; // end do for I := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
    end; // end do 6


    if TIWUserSession(WebApplication.Data).estagio = -1 then
    begin
    rgnLoginproject.Visible := false;
    rgnConseqvalues.Visible := true;
    TIWUserSession(WebApplication.Data).Nomeprojeto := TIWUserSession(WebApplication.Data).auxnome;
    // procedimento para retornar dados da elicitação
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 25;
    SalvarBD;

    with TIWUserSession(WebApplication.Data) do
    begin
    Setlength(ArrayAlt, TIWUserSession(WebApplication.Data).NumAlt);
    for I := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
    begin
    ArrayAlt[I] := TIWLabel.Create(Self);
    ArrayAlt[I].parent := rgnconsman;
    ArrayAlt[I].text := ('Alt' + inttostr(I + 1));
    ArrayAlt[I].Left := (5);
    ArrayAlt[I].Top := (50 + (I * 25));
    ArrayAlt[I].Font.Style := [fsbold];
    end;

    Setlength(ArrayCrt, TIWUserSession(WebApplication.Data).NumCrit);
    for I := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
    begin
    ArrayCrt[I] := TIWLabel.Create(Self);
    ArrayCrt[I].parent := rgnconsman;
    ArrayCrt[I].text := ('C' + inttostr(I + 1));
    ArrayCrt[I].Left := (45 + (I * 60));
    ArrayCrt[I].Top := (20);
    ArrayCrt[I].Font.Style := [fsbold];
    end;
    end;

    Setlength(TIWUserSession(WebApplication.Data).MatEdit, TIWUserSession
    (WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data)
    .NumCrit);
    for I := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
    begin
    for J := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do

    begin
    // Criando os edits que recaberão os desempenhos a partir de uma Matriz de TIWEdit's
    TIWUserSession(WebApplication.Data).MatEdit[I][J] := TIWEdit.Create
    (Self);
    TIWUserSession(WebApplication.Data).MatEdit[I][J].parent := rgnconsman;
    TIWUserSession(WebApplication.Data).MatEdit[I][J].text := ('');
    TIWUserSession(WebApplication.Data).MatEdit[I][J].Left :=
    (45 + (J * 60));
    TIWUserSession(WebApplication.Data).MatEdit[I][J].Top :=
    (25 + (I * 25) + 23);
    TIWUserSession(WebApplication.Data).MatEdit[I][J].Width := 60;
    end;
    end;
    // cmproblematica1.ItemIndex := -1;  // Não limpo o nome caso o usuário volte para essa página

    AuxTipo[0] := 'Linear';
    AuxTipo[1] := 'Exponential';
    AuxTipo[2] := 'Logarithmic';
    AuxTipo[3] := 'Logistic';
    // ** Preenchendo a tabela de informações dos critérios **//

    grdCritInfo.ColumnCount := 7;
    grdCritInfo.RowCount := TIWUserSession(WebApplication.Data).NumCrit + 1;

    grdCritInfo.Cell[0, 0].text := 'Name';
    grdCritInfo.Cell[0, 1].text := 'Scale';
    grdCritInfo.Cell[0, 2].text := 'Number of levels';
    grdCritInfo.Cell[0, 3].text := 'Value Function';
    grdCritInfo.Cell[0, 4].text := 'Direction';
    grdCritInfo.Cell[0, 5].text := 'Par a';
    grdCritInfo.Cell[0, 6].text := 'Par b';

    for I := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
    begin
    grdCritInfo.Cell[I + 1, 0].text := TIWUserSession(WebApplication.Data)
    .NomeCrit[I];
    if TIWUserSession(WebApplication.Data).Typeofcrit[I] > 1 then
    begin
    grdCritInfo.Cell[I + 1, 1].text := 'Discrete';
    grdCritInfo.Cell[I + 1, 2].text := inttostr
    (TIWUserSession(WebApplication.Data).Niveis[I]);
    end // end do   if TIWUserSession(WebApplication.Data).Typeofcrit[I] > 1
    else
    begin
    grdCritInfo.Cell[I + 1, 1].text := 'Continuous';
    grdCritInfo.Cell[I + 1, 2].text := '---';
    end; // end do else
    grdCritInfo.Cell[I + 1, 3].text := AuxTipo
    [TIWUserSession(WebApplication.Data).Ftipo[I] - 1];

    if (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 1) OR
    (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 3) then
    begin
    grdCritInfo.Cell[I + 1, 4].text := 'Maximization';
    grdCritInfo.Cell[I + 1, 5].text := FloatToStr
    (TIWUserSession(WebApplication.Data).parA[I]);
    grdCritInfo.Cell[I + 1, 6].text := FloatToStr
    (TIWUserSession(WebApplication.Data).parB[I]);
    if (TIWUserSession(WebApplication.Data).parA[I] = 0) then
    grdCritInfo.Cell[I + 1, 5].text := '---';
    if (TIWUserSession(WebApplication.Data).parB[I] = 0) then
    grdCritInfo.Cell[I + 1, 6].text := '---';
    end // end do  if (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 1)...
    else if (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 0) OR
    (TIWUserSession(WebApplication.Data).Typeofcrit[I] = 2) then
    begin
    grdCritInfo.Cell[I + 1, 4].text := 'Minimization';
    grdCritInfo.Cell[I + 1, 5].text := FloatToStr
    (TIWUserSession(WebApplication.Data).parA[I]);
    grdCritInfo.Cell[I + 1, 6].text := FloatToStr
    (TIWUserSession(WebApplication.Data).parB[I]);
    if (TIWUserSession(WebApplication.Data).parA[I] = 0) then
    grdCritInfo.Cell[I + 1, 5].text := '---';
    if (TIWUserSession(WebApplication.Data).parB[I] = 0) then
    grdCritInfo.Cell[I + 1, 6].text := '---';
    end; // end do else
    end; // end do for I := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
    end; // end do -1

    if (TIWUserSession(WebApplication.Data).estagio = 2) OR (TIWUserSession(WebApplication.Data).estagio = 3) OR (TIWUserSession(WebApplication.Data).estagio = 4) OR (TIWUserSession(WebApplication.Data).estagio = 5) then
    begin
    TIWUserSession(WebApplication.Data).tempinicio := TimeToStr(Now);

    // procedimento para pegar o round, caracteristicas dos criterios e das alternativas
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 15;
    SalvarBD;
    rgnInputMatrizConseq.Visible := true;
    rgnLoginproject.Visible := false;

    // procedimento para preencher as informações na região rgnInputMatrizConseq

    Inicializa;
    if TIWUserSession(WebApplication.Data).Flag <> -2 then
    begin
    with TIWUserSession(WebApplication.Data) do
    begin
    /// / Setando o tamanho dos grids que exibem a informação importada na região imput ////
    grdInputvalues.RowCount := 6;
    grdInputvalues.ColumnCount := NumCrit + 1;
    grdMatrizConseq.RowCount := NumAlt;
    grdMatrizConseq.ColumnCount := NumCrit + 1;
    probname := edtEntername.text;

    /// / Preenchendo a informação fixa dos grids ////
    grdInputvalues.Cell[0, 0].text := 'Criteria:';
    grdInputvalues.Cell[1, 0].text :=
    '0-Cont Min; 1-Cont Max; 2-Disc Min; 3- Disc Max:';
    { grdInputvalues.Cell[2, 0].text := 'Weights:'; }
  { grdInputvalues.Cell[2, 0].text := 'Type:';
    grdInputvalues.Cell[3, 0].text := 'a:';
    grdInputvalues.Cell[4, 0].text := 'b:';
    grdInputvalues.Cell[5, 0].text :=
    'c: number of levels of discrete criteria';

    /// / Preenchendo os Grids com a informação Imputada ///////////////////////
    for J := 0 to NumCrit - 1 do
    begin
    lstbxInitialOrder.Items.Add(NomeCrit[J]);
    grdInputvalues.Cell[0, J + 1].text := NomeCrit[J];
    grdInputvalues.Cell[1, J + 1].text := inttostr(Typeofcrit[J]);
    grdInputvalues.Cell[2, J + 1].text := inttostr(Ftipo[J]);
    if Ftipo[J] <> 1 then
    begin
    grdInputvalues.Cell[3, J + 1].text := FloatToStr(parA[J]);
    grdInputvalues.Cell[4, J + 1].text := FloatToStr(parB[J]);
    end;
    grdInputvalues.Cell[5, J + 1].text := inttostr(Niveis[J]);
    end;

    for I := 0 to NumAlt - 1 do
    begin
    grdMatrizConseq.Cell[I, 0].text := NomeAlt[I];
    for J := 0 to NumCrit - 1 do
    begin
    grdMatrizConseq.Cell[I, J + 1].text := TrocaVP(ConsMatrix[I, J]);
    end;
    end;
    /// ////////////////////////////////////////////////////////////////////////
    // end;
    end;
    end;
    WebApplication.ShowMessage('Sorry for the inconvenience, the database has been modified and it is not possible to return from where you left off');
    end; //referente a    if (TIWUserSession(WebApplication.Data).estagio = 2)

    end;  //if TIWUserSession(WebApplication.Data).Numcrit > 10

    end; //referente a  if (StrToDate(dtinicretorn) < StrToDate('19/04/2021')) AND (StrToDate(

    }

  // ** Elementos para dar impressão de carregamento **//
  rgnLoginproject.Cursor := CRAUTO;
  tmrcontinue.Enabled := false;
  btnauxchoose.Enabled := false;

  if mensageretorno = -90998 then
    WebApplication.ShowMessage(
      'Sorry for the inconvenience, the database has been modified and it is not possible to return from where you left off');
  mensageretorno := 0;

end; // end do procedimento

procedure TFrmEntrada.btncloseIDCClick(Sender: TObject);
begin
  rgnaboutdiscrete.Visible := false;
  rgntabela.Visible := true;
  rgntabelaalter.Visible := true;
end;

procedure TFrmEntrada.btncloselaClick(Sender: TObject);
begin
  rgnlanguage.Visible := false;
end;

procedure TFrmEntrada.btnConfBackClick(Sender: TObject);
var
  i, j: integer;
begin
  // Volta para página de adição de critérios e alternativas

  rgnConfirmBackEditar.Visible := false;
  rgnConseqvalues.Visible := false;
  rgnCadprojeto.Visible := true;

  btnSaveproblem.Enabled := true;
  btnSaveStart.Enabled := true;

  // Limpar
  // edtProjectname.Clear;
  edtNamealternat.Clear;
  edtNameattribut.Clear;
  edtparA.Clear;
  edtparB.Clear;
  edtbudget.Clear;
  edtNumberalternat.Text := '0';
  edtNumberattribut.Text := '0';

  lstbxNamealternat.Clear;
  lstbxNamecriteria.Clear;

  // Preencher com os dados das variáveis globais recuperados do BD
  with TIWUserSession(WebApplication.Data) do
  begin
    Setlength(AuxNomeAlt, 100);
    Setlength(AuxNomeCrit, 100);
    Setlength(AuxFuncTipo, 100);
    Setlength(AuxNiveis, 100);
    Setlength(AuxParA, 100);
    Setlength(AuxParB, 100);
    Setlength(AuxEscalaTipo, 100);
    Setlength(AuxDirecTipo, 100);

    // Nomeprojeto := auxnome;
    edtProjectname.Text := Nomeprojeto;
    edtNumberalternat.Text := inttostr(NumAlt);
    edtNumberattribut.Text := inttostr(NumCrit);
    edtbudget.Text := FloatToStr(budget);

    for i := 0 to NumAlt - 1 do
    begin
      lstbxNamealternat.Items.Add(NomeAlt[i]);
      AuxNomeAlt[i] := NomeAlt[i];
    end;
    for j := 0 to NumCrit - 1 do
    begin
      lstbxNamecriteria.Items.Add(NomeCrit[j]);
      AuxNomeCrit[j] := NomeCrit[j];
      AuxFuncTipo[j] := FTipo[j];
      AuxNiveis[j] := Niveis[j];
      AuxParA[j] := parA[j];
      AuxParB[j] := parB[j];

      if Typeofcrit[j] = 1 then
      begin
        AuxEscalaTipo[j] := 0;
        AuxDirecTipo[j] := 0;
      end;
      if Typeofcrit[j] = 0 then
      begin
        AuxEscalaTipo[j] := 0;
        AuxDirecTipo[j] := 1;
      end;
      if Typeofcrit[j] = 3 then
      begin
        AuxEscalaTipo[j] := 1;
        AuxDirecTipo[j] := 0;
      end;
      if Typeofcrit[j] = 2 then
      begin
        AuxEscalaTipo[j] := 1;
        AuxDirecTipo[j] := 1;
      end;

    end;

    AuxNumAlt := NumAlt;
    AuxNumCrit := NumCrit;

  end;

  { --------- LIMPAR A MATRIZ DE EDITS ----------- }
  with TIWUserSession(WebApplication.Data) do
  begin
    for i := 0 to NumAlt - 1 do
    begin
      if length(ArrayAlt) > i then
        ArrayAlt[i].Destroy;

      if TIWUserSession(WebApplication.Data).Problematica = 2 then
      Begin
        for j := 0 to NumCrit do
        begin
          if length(MatEdit) > i then
            if length(MatEdit[i]) > j then
              MatEdit[i, j].Destroy;
          if i = 0 then
            if length(ArrayCrt) > j then
              TIWUserSession(WebApplication.Data).ArrayCrt[j].Destroy;
        end;
      End
      else
      begin
        for j := 0 to NumCrit - 1 do
        begin
          if length(MatEdit) > i then
            if length(MatEdit[i]) > j then
              MatEdit[i, j].Destroy;
          if i = 0 then
            if length(ArrayCrt) > j then
              ArrayCrt[j].Destroy;
        end;
      end;
    end; // for i

    Setlength(MatEdit, 0, 0);
    Setlength(ArrayAlt, 0);
    Setlength(ArrayCrt, 0);
  end; // with

  // auxiliar para indicar que retornou para editar altenativas e criterios

  TIWUserSession(WebApplication.Data).editar := -33333;
  lstbxNamecriteria.Enabled := true; // alteração 20.02
  lstbxNamealternat.Enabled := true; // alteração 20.02

end;

procedure TFrmEntrada.btnConfNaoClick(Sender: TObject);
begin
  rgnConfirmBackEditar.Visible := false;

  btnSaveproblem.Enabled := true;
  btnSaveStart.Enabled := true;

end;

procedure TFrmEntrada.btnContinueregistClick(Sender: TObject);
var
  i: integer;
begin
  AlreadyRakend1.Visible := false;
  { inicialmente pegar atraves do email e senha os problemas que aquele usuario
    já cadastrou e apresenta-los no listbox da região rgnLoginproject }

  // pegar os nomes dos problemas e ids relacionados ao usuario que foi feito o login
  TIWUserSession(WebApplication.Data).AuxsalvarBD := 13;
  SalvarBD;

  if length(TIWUserSession(WebApplication.Data).nameproblem) = 0 then
  begin
    WebApplication.ShowMessage('There is no problem to be continued');
    Exit;
  end;

  rgnOpcaoInput.Visible := false;
  rgnLoginproject.Visible := true;
  btnauxchoose.Enabled := true;

  // colocar o vetor do nome no listbox
  lstbxProjects.Items.Clear;
  for i := 0 to length(TIWUserSession(WebApplication.Data).nameproblem) - 1 do
  begin
    lstbxProjects.Items.Add(TIWUserSession(WebApplication.Data)
        .nameproblem[i] + ' - (' + TIWUserSession(WebApplication.Data)
        .datainicioproblema[i] + ')');
  end;

  { // salvar banco de dados na tabela de problem
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 12;
    SalvarBD; }

  // auxiliar ans_seq
  TIWUserSession(WebApplication.Data).auxiliar_ans_seq_bd := true;

end;

procedure TFrmEntrada.btndeletealternatClick(Sender: TObject); // procedimento para deletar alternativa inserida no listbox de conferência
var
  auxnome: string;
  indicealtdaux, valor, c: integer;
  nome: string;

begin

  // verifica se um nome realmente foi selecionado
  if (lstbxNamealternat.ItemIndex = -1) AND
    (TIWUserSession(WebApplication.Data).EditaAlt <> 1) then
  begin
    WebApplication.ShowMessage(
      'Please select an alternative to exclude from the list.');
  end
  else
  begin

    /// /// Procedimento semelhante ao que está no botão (-)
    // Capta o valor inteiro da string presente no edit
    valor := StrToInt(edtNumberalternat.Text) - 1; // a variável é inicializada em zero e tem que incrementar a cada clique
    edtNumberalternat.Text := inttostr(valor);
    // mostra pro usuario a qntd de alt

    // captura o nome selecionado do listbox e o indice
    nome := lstbxNamealternat.Items[lstbxNamealternat.ItemIndex];
    indicealtdaux := lstbxNamealternat.ItemIndex;

    // Limpando o edit caso a alternativa tenha sido selecionada para edição e em seguida excluida
    edtNamealternat.Clear;

    with TIWUserSession(WebApplication.Data) do
    begin
      AuxAltApagadas := AuxAltApagadas + 1;

      for c := 0 to length(AuxNomeAlt) do
      begin
        if AuxNomeAlt[c] = nome then
        begin
          TIWUserSession(WebApplication.Data).AuxNomeAlt[c] := '';
          // Limpa o nome do vetor auxiliar
          Break;
        end;
      end;
    end;

    // deleta item selecionado do listbox
    lstbxNamealternat.Items.Delete(indicealtdaux);
    lstbxNamealternat.ItemIndex := -1; // tira a seleção
    btnAddalternative.Caption := 'Add';
    lstbxNamealternat.Enabled := true;
    TIWUserSession(WebApplication.Data).EditaAlt := 0;
  end;
end;

procedure TFrmEntrada.btndeletecriteriaClick(Sender: TObject); // procedimento para deletar critério inserido no listbox de conferência
var
  c, valor: integer;
  nome: string;
begin

  // verifica se um nome realmente foi selecionado
  if lstbxNamecriteria.ItemIndex = -1 then
  begin
    WebApplication.ShowMessage(
      'Please select a criterion to exclude from the list.');
    Exit;
  end;

  with TIWUserSession(WebApplication.Data) do
  begin
    /// /// Procedimento semelhante ao que está no botão (-)
    // E capta o valor inteiro da string presente no edit
    valor := StrToInt(edtNumberattribut.Text) - 1; // a variável é inicializada em zero e tem que incrementar a cada clique
    edtNumberattribut.Text := inttostr(valor);
    // mostra pro usuario a qntd de crit

    nome := lstbxNamecriteria.Items[lstbxNamecriteria.ItemIndex];
    for c := 0 to length(AuxNomeCrit) - 1 do
      if AuxNomeCrit[c] = nome then
      begin
        AuxNomeCrit[c] := ''; // limpa do vetor
        Break;
      end;

    lstbxNamecriteria.Items.Delete(lstbxNamecriteria.ItemIndex);
    // tira do listbox o critério selecionado

    AuxCritApagados := AuxCritApagados + 1; // conta a apagada
    lstbxNamecriteria.ItemIndex := -1; // tira a seleção do listbox
    lstbxNamecriteria.Enabled := true;
    EditaCrit := 0;

  end; // with

  // Limpando caso a alternativa tenha sido selecionada para edição e em seguida excluida
  edtNameattribut.Clear;
  edtparA.Clear;
  edtparB.Clear;
  edtNumberlevels.Clear;
  btnAddattribut.Caption := 'Add';

  cmbxScaletype.ItemIndex := -1;
  cmbxFO.ItemIndex := -2;
  cmbxCriteriondirection.ItemIndex := -1;

  lblfunc.Visible := false;
  lblNumberlevels.Visible := false;
  edtNumberlevels.Visible := false;
  txtparametros.Visible := false;
  lblparA.Visible := false;
  lblparB.Visible := false;
  edtparA.Visible := false;
  edtparB.Visible := false;

end;

procedure TFrmEntrada.btnDiscrete0Click(Sender: TObject);
begin
  { rgnaboutdiscrete0.Visible := true;
    edtNameattribut.Visible := true; }
  WebApplication.ShowMessage(txthelpdisc.Lines.Text);
  { rgnaux1.top := 176;
    rgnaux1.Left := 8;
    rgnaux1.Height := 513;
    rgnaux1.Width := 1223;
    rgnaboutdiscrete0.Left := 329;
    rgnaboutdiscrete0.Top := 165; }
end;

procedure TFrmEntrada.btneditalternativeClick(Sender: TObject);
// procedimento para editar alternativa já existente no listbox
// var
// auxnome: string; // apagar (variável local para verificação)
begin

  lstbxNamealternat.Enabled := false; // Desabilita o listbox para que o usuário não altere a alternativa selecionada para edição a medida que a edita

  if lstbxNamealternat.ItemIndex = -1 then
  begin
    WebApplication.ShowMessage('Select an alternative to edit');
    Exit;
  end;
  // else
  // begin
  with TIWUserSession(WebApplication.Data) do
  begin
    EditaAlt := 1; // usado no procedimento AddAlt pra indicar Edição
    btnAddalternative.Caption := 'Save'; // Altera o nome do botão para ficar claro para o usuário
    TIWUserSession(WebApplication.Data).NomeAlternativaEditada :=
      lstbxNamealternat.Items[lstbxNamealternat.ItemIndex];
    // captura o nome selecionado do listbox
    TIWUserSession(WebApplication.Data).AuxIndiceAlt :=
      lstbxNamealternat.ItemIndex;
    // Indice *no listbox* da alternativa a ser editada
    edtNamealternat.Text := lstbxNamealternat.Items.Strings[AuxIndiceAlt];
    // mostra nome no edit
    // end;
  end;
end;

procedure TFrmEntrada.btneditcriteriaClick(Sender: TObject);
// procedimento para editar critério já existente no listbox
var
  auxnome: string;
  c: integer;

begin
  // verifica se um nome realmente foi selecionado
  if lstbxNamecriteria.ItemIndex = -1 then
  begin
    WebApplication.ShowMessage('Select a criterion to edit');
    Exit;
  end;

  with TIWUserSession(WebApplication.Data) do
  begin
    EditaCrit := 1; // editar critério

    TIWUserSession(WebApplication.Data).NomeCriterioEditado :=
      lstbxNamecriteria.Items[lstbxNamecriteria.ItemIndex];
    for c := 0 to length(AuxNomeCrit) - 1 do
    begin
      if AuxNomeCrit[c] = NomeCriterioEditado then
      begin
        IndiceCriterioEditado := c;
        Break;
      end;
    end;

    lstbxNamecriteria.Enabled := false;
    btnAddattribut.Caption := 'Save'; // Altera o nome do botão para ficar claro para o usuário

    // Alterações no ListBox
    AuxIndice := lstbxNamecriteria.ItemIndex; // índice do critério *no listbox*  a ser editado
    edtNameattribut.Text := lstbxNamecriteria.Items.Strings[AuxIndice];
    // nome no edit

    // Com base nos dados/índice do vetor
    { deixar os combobox e os nomes de acordo com o criterio selecionado }

    begin
      case AuxFuncTipo[IndiceCriterioEditado] of
        1:
          cmbxFO.ItemIndex := 1; // Linear
        2:
          cmbxFO.ItemIndex := 3; // Exponencial
        3:
          cmbxFO.ItemIndex := 0; // Logaritmica
        4:
          cmbxFO.ItemIndex := 2; // Logistica
        -1:
          cmbxFO.ItemIndex := -1; // optou realizar elic intra
      end;

      cmbxScaletype.ItemIndex := AuxEscalaTipo[IndiceCriterioEditado];
      cmbxCriteriondirection.ItemIndex := AuxDirecTipo[IndiceCriterioEditado];

      if (AuxFuncTipo[IndiceCriterioEditado] <> 1) OR
        (AuxFuncTipo[IndiceCriterioEditado] <> -1) then
      begin
        lblfunc.Visible := true;
        txtparametros.Visible := true;
        lblparA.Visible := true;
        lblparB.Visible := true;
        edtparA.Visible := true;
        edtparB.Visible := true;
        edtparA.Text := FloatToStr(AuxParA[IndiceCriterioEditado]);
        edtparB.Text := FloatToStr(AuxParB[IndiceCriterioEditado]);

        // Combobox referente a AuxFuncTipo Não-Linear
        if cmbxFO.ItemIndex = 0 then
          lblfunc.Caption := 'F(x) = ln(ax) + b';
        if cmbxFO.ItemIndex = 2 then
          lblfunc.Caption := 'F(x) = e^(-a/x) + b';
        if cmbxFO.ItemIndex = 3 then
          lblfunc.Caption := 'F(x) = e^(ax) + b';

      end
      else
      begin
        lblfunc.Visible := false;
        txtparametros.Visible := false;
        lblparA.Visible := false;
        lblparB.Visible := false;
        edtparA.Visible := false;
        edtparB.Visible := false;
      end;

      if Cri_inteiro[IndiceCriterioEditado] = true then
      begin
        chckbxinteger.Checked := true;
        chckbxinteger.Visible := true;
      end
      else
      begin
        chckbxinteger.Checked := false;
        chckbxinteger.Visible := false;
      end;

      if AuxEscalaTipo[IndiceCriterioEditado] = 1 then
      begin
        lblNumberlevels.Visible := true;
        edtNumberlevels.Visible := true;
        edtNumberlevels.Text := inttostr(AuxNiveis[IndiceCriterioEditado]);
        rdbtndeclaretype.Checked := true;
        if true then

        end
      else
      begin
        lblNumberlevels.Visible := false;
        edtNumberlevels.Visible := false;

        // alteração 20/01. Ver se influencia na função do btnedit
        rdbtnelictfunction.Checked := false;
        rdbtndeclaretype.Checked := false;
        lblfunc.Visible := false;
        txtparametros.Visible := false;
        lblparA.Visible := false;
        lblparB.Visible := false;
        edtparA.Visible := false;
        edtparB.Visible := false;

      end;

    end;
  end;
end;

procedure TFrmEntrada.btnEnglishClick(Sender: TObject);
begin
  WebApplication.SendFile(
    'C:\inetpub\fitradeoff\Images\User guide FITradeoff web' + '.pdf', true);
  rgnlanguage.Visible := false;
end;

procedure TFrmEntrada.btnEnterClick(Sender: TObject);
// procedimento de login no sistema
var
  email, id_email: string;
  senha, validation: string;
  quantregis: integer;

begin
  IWLabel9.Visible := false;
  IWLabel8.Visible := false;
  TIWUserSession(WebApplication.Data).primeiroacesso := false;
  // inicialmente fazendo um selectcount para verificar se já existe algum cadastro com esse email
  {with UserSession.ZQuery1 do
  begin

    SQL.Text := 'SELECT count(*) FROM user WHERE email="' + (edtemail.Text)
      + '"';
    ExecSQL;
    open;
    quantregis := FieldByName('count(*)').AsInteger;
    close;
  end; }
  if quantregis = 0 then
  begin
    IWLabel9.Visible := true;
    IWButton1.Visible := true;
  end
  else
  begin
    TIWUserSession(WebApplication.Data).email := edtemail.Text;
   { with UserSession.ZQuery1 do
    begin

      SQL.Text := 'SELECT validacao FROM user WHERE email ="' + (edtemail.Text)
        + '"';
      ExecSQL;
      open;
      validation := FieldByName('validacao').AsString;
      close;

      SQL.Text := 'SELECT password FROM user WHERE email ="' + (edtemail.Text)
        + '"';
      ExecSQL;
      open;
      senha := FieldByName('password').AsString;
      close;
    end;  }
    if senha = edtSenha.Text then
    begin
      {with UserSession.ZQuery1 do
      begin
        SQL.Text := 'SELECT id FROM user WHERE email="' + (edtemail.Text) + '"';
        ExecSQL;
        open;
        TIWUserSession(WebApplication.Data).ID_user := FieldByName('id')
          .AsString;
        close;

        SQL.Text := 'SELECT analyst FROM user WHERE email="' + (edtemail.Text)
          + '"';
        ExecSQL;
        open;
        TIWUserSession(WebApplication.Data).analyst := FieldByName('analyst')
          .AsInteger;
        close;

        if TIWUserSession(WebApplication.Data).analyst = 0 then
        begin
          TIWUserSession(WebApplication.Data).email := edtemail.Text;
          TIWUserSession(WebApplication.Data).senha := edtSenha.Text;
        end;

      end;    }

      // * Caso o usuário esteja trocando a senha de cadastro *//
      if validation = 'password' then
      begin
        rgnvalidationpassword.show;
        rgnnewpassword.show;
        Exit;
      end;

      // * Caso seja o primeiro acesso *//
      if validation <> 'validado' then
      begin
        rgnvalidationpassword.show;
        rgnvalidation.show;
        edtSenha.Clear;
        // edtemail.Clear;
        Exit;
      end;

      rgnEnterLogin.Visible := false;
      rgnproblematica.Visible := true;
      rgnOpcaoInput.Visible := false;
    end
    else
    begin
      IWLabel8.Visible := true;
      IWButton1.Visible := true;
    end;

  end;
  //** Salvando informação de data e horário
          TIWUserSession(WebApplication.Data).AuxsalvarBD := 46;
          SalvarBD;

  // end;

end;

procedure TFrmEntrada.btnExporttemplateClick(Sender: TObject);
// procedimento para exportação modelo de planilha
var
  i: integer;
Begin
  with TIWUserSession(WebApplication.Data) do
  begin
    if Problematica = 2 then
      WebApplication.SendFile(
        'C:\inetpub\fitradeoff\Images\Fitradeoff_Portfólio_Tamplate' + '.xls',
        true)
    else
      WebApplication.SendFile
        ('C:\inetpub\fitradeoff\Images\Fitradeoff_Tamplate' + '.xls', true);
    { grdImportacao.TotalRows := 12;
      if (Problematica = 2) then // portfolio
      begin
      ExporttemplatePortfolioCB;
      Exit;
      end;
      for i := 0 to 4 do
      begin
      grdImportacao.Columns.Add
      end; { no componente advwebgrid o endereçamento de celulas segue o padrão AdvWebGrid.Cells[coluna, linha]:= }

    { grdImportacao.Cells[0, 0] := 'Criteria:';
      grdImportacao.Cells[0, 1] :=
      '0-Cont Min; 1-Cont Max; 2-Disc Min; 3-Disc Max';
      // grdImportacao.Cells[0, 2] := 'Weights';
      grdImportacao.Cells[0, 3] := 'Type:';
      grdImportacao.Cells[0, 4] := 'a:';
      grdImportacao.Cells[0, 5] := 'b:';
      grdImportacao.Cells[0, 6] := 'c:';
      grdImportacao.Cells[0, 7] := 'Alternatives:';
      grdImportacao.Cells[0, 8] := 'Alternative 1';
      grdImportacao.Cells[0, 9] := 'Alternative 2';
      grdImportacao.Cells[0, 10] := 'Alternative 3';
      grdImportacao.Cells[0, 11] := 'Alternative 4';
      grdImportacao.Cells[1, 0] := 'Criteria 1';
      grdImportacao.Cells[2, 0] := 'Criteria 2';
      grdImportacao.Cells[3, 0] := 'Criteria 3';
      grdImportacao.Cells[4, 0] := 'Criteria 4';
      grdImportacao.Cells[1, 7] := 'Consequence Matrix:';
      TIWAdvWebGridExcelIO2.AdvWebGrid := grdImportacao;
      TIWAdvWebGridExcelIO2.XLSExport('FITradeoff_model' + '.xls',
      WebApplication);

      end; // conexão user session }
  end;
end; // procedimento

procedure TFrmEntrada.English1Click(Sender: TObject);
begin
  WebApplication.SendFile(
    'C:\inetpub\fitradeoff\Images\User guide FITradeoff web' + '.pdf', true);
end;

Procedure TFrmEntrada.ExporttemplatePortfolioCB;
// procedimento para exportação modelo de planilha
var
  i: integer;
Begin
  with TIWUserSession(WebApplication.Data) do
  begin
    grdImportacao.TotalRows := 12;
    for i := 0 to 5 do
    begin
      grdImportacao.Columns.Add
    end; { no componente advwebgrid o endereçamento de celulas segue o padrão AdvWebGrid.Cells[coluna, linha]:= }

    grdImportacao.Cells[0, 0] := 'Criteria:';
    grdImportacao.Cells[0, 1] :=
      '0-Cont Min; 1-Cont Max; 2-Disc Min; 3-Disc Max; 4- Int Min; 5- Int Max:';
    // grdImportacao.Cells[0, 2] := 'Weights';
    grdImportacao.Cells[0, 3] := '';
    grdImportacao.Cells[0, 4] := '';
    grdImportacao.Cells[0, 5] := '';
    grdImportacao.Cells[0, 6] := 'Number of levels of discrete criteria:';
    grdImportacao.Cells[0, 7] := 'Alternatives:';
    grdImportacao.Cells[0, 8] := 'Alternative 1';
    grdImportacao.Cells[0, 9] := 'Alternative 2';
    grdImportacao.Cells[0, 10] := 'Alternative 3';
    grdImportacao.Cells[0, 11] := 'Alternative 4';
    grdImportacao.Cells[1, 0] := 'Criteria 1';
    grdImportacao.Cells[2, 0] := 'Criteria 2';
    grdImportacao.Cells[3, 0] := 'Criteria 3';
    grdImportacao.Cells[4, 0] := 'Criteria 4';
    grdImportacao.Cells[5, 0] := 'Budget';
    grdImportacao.Cells[1, 7] := 'Consequence Matrix:';
    TIWAdvWebGridExcelIO2.AdvWebGrid := grdImportacao;
    TIWAdvWebGridExcelIO2.XLSExport('FITradeoff_model' + '.xls',
      WebApplication);
  end; // conexão user session
end; // procedimento

procedure TFrmEntrada.btnImportClick(Sender: TObject);

begin
  AlreadyRakend1.Visible := false;
  rgnOpcaoInput.Visible := false;
  rgnImportExport.Visible := true;

  if TIWUserSession(WebApplication.Data).Problematica = 2 then
  begin
    cmbbxproblem.Visible := false;
    lblchosenproblematic.Visible := true;
    lblchosenproblematic.Caption := 'Portfolio';
  end; // refrente a if TIWUserSession(WebApplication.Data).problematica =2 then

  if TIWUserSession(WebApplication.Data).Problematica = 0 then
  begin
    cmbbxproblem.Visible := false;
    lblchosenproblematic.Visible := true;
    lblchosenproblematic.Caption := 'Choice';
  end; // refrente a if TIWUserSession(WebApplication.Data).problematica =0 then

  if TIWUserSession(WebApplication.Data).Problematica = 3 then
  begin
    cmbbxproblem.Visible := false;
    lblchosenproblematic.Visible := true;
    lblchosenproblematic.Caption := 'Sorting';
  end; // refrente a if TIWUserSession(WebApplication.Data).problematica =1 then

  if TIWUserSession(WebApplication.Data).Problematica = 1 then
  begin
    cmbbxproblem.Visible := false;
    lblchosenproblematic.Visible := true;
    lblchosenproblematic.Caption := 'Ranking';
  end; // refrente a if TIWUserSession(WebApplication.Data).problematica =1 then

  // flImportarproblema.Visible := true;
  // btnslotherfile.Enabled := true;

  // salvar banco de dados na tabela de problema
  TIWUserSession(WebApplication.Data).ID_chave := idUnique;

end;

procedure TFrmEntrada.btnImporttabelaClick(Sender: TObject);
// procedimento que direciona para o modulo correspondnete a problemática e importação do arquivo xls
begin
  with TIWUserSession(WebApplication.Data) do
  begin
    if (Problematica = 2) then
      ImporttabelaPortfolioCB;

    if (Problematica = 0) or (Problematica = 1) then
      ImporttabelaChoiRank;

    if (Problematica = 3) then
      ImporttabelaClassificacao;

    if (Problematica = 4) then
      ImporttabelaPortfolioC;

  end; // ifs
end; // procedimento

procedure TFrmEntrada.ImporttabelaChoiRank;
// procedimento para importação de planilha .xls em escolha e ordenação
var
  tipoarquivo: string;
  i, j, k, VerificaAlt, ACP: integer;
  NumCri, NumAlt, Numcolimport: integer;
  vtrtipo, tcrit, niveisdisc: array of integer;
  vtrnomecrit, vtrpesos, vtrnomealternativa: array of string;
  vtrparametroA, vtrparametroB: array of Double;
  consreal1: array of array of Double;

  /// //////////////////////////////////////////////////////////////////////////
  /// / Procedimento especial para evitar erros ao apagar linhas e colunas da
  /// planilha do excel sem excluí-las ////////////////////////////////////////
  procedure validar_planilha;
  /// / Esse procedimento valida o número real de alternativas e critérios introdu-
  /// zidos na planilha de input //////////////////////////////////////////////
  var
    i, j, k: integer;
    StrAux: string;
    anterior: boolean;
    /// / Verdadeiro quando o nome de um critério ou alternativa
    /// não foi preenchido //////////////////////////////////////////////////////
  begin
    k := 0;
    anterior := false;
    ACP := 0;

    for i := 0 to NumCri - 1 do
    begin
      StrAux := grdImportproblema.Cells[i + 1, 0];
      if StrAux <> '' then
      begin
        INC(k);
        /// / Caso o critério atual teve o nome preenchido e o anterior não, então
        /// deve-se alertar ao Decisor que existe um problema na matriz de input
        if anterior = true then
        begin
          ACP := 1;
          anterior := false;
          WebApplication.ShowMessage(
            'There is at least one criterion name missing, please verify the Excel spreadsheet before continuing.');
          Break;
        end; // Anterior
      end // strAux
      else
        anterior := true;
    end; // for i
    if ACP = 0 then
      NumCri := k;
    k := 0;
    anterior := false;
    for j := 0 to NumAlt - 1 do
    begin
      StrAux := grdImportproblema.Cells[0, j + 8];
      if StrAux <> '' then
      begin
        INC(k);

        for i := 0 to NumCri - 1 do
        begin
          StrAux := grdImportproblema.Cells[i + 1, j + 8];
          if StrAux = '' then
            // ** Neste caso está faltando alguma consequência **//
            anterior := true;
        end;

        /// / Caso a alternativa atual teve o nome preenchido e a anterior não, então
        /// deve-se alertar ao Decisor que existe um problema na matriz de input
        if anterior = true then
        begin
          anterior := false;
          // ** Caso já haja problema nos critérios **//
          if ACP = 1 then
            ACP := 3;
          // ** Caso os critérios estejam ok **//
          if ACP = 0 then
          begin
            ACP := 2;
            WebApplication.ShowMessage(
              'There is at least one alternative name or consequence missing, please verify the Excel spreadsheet before continuing.');
          end;
          Break;
        end; // anterior = true
      end // atrAux
      else
        anterior := true;
    end; // for j
    if ACP = 0 then
      NumAlt := k;
  end;

/// ///////////////////////////////////////////////////////////////////////////
begin // procedure
  IWLabel13.Visible := true;
  IWLabel14.Visible := true;
  IWLabel15.Visible := true;
  TIWUserSession(WebApplication.Data).ValDiscreto := 0;
  if edtEntername.Text = '' then // verifica se foi inserido um nome pro arquivo
  begin // se não tiver nome do problema
    WebApplication.ShowMessage('Please, enter a problem''s name!');
    Exit;
  end // se não tiver nome do problema

  else
  begin // se tiver nome do problema
    if flImportarproblema.Filename = '' then
    begin // se o file estiver preenchido
      WebApplication.ShowMessage('You should choose an excel file.');
    end // se o file não estiver preenchido

    else
    begin // se o file estiver preenchido
      tipoarquivo := flImportarproblema.ContentType;
      if (tipoarquivo = ' application/vnd.ms-excel') then
      // verifica se o arquivo é compatível

      begin
        { se o arquivo for compatível é criado um arquivo de excel no diretório especificado, igual ao
          selecionado pelo usuário. É de la que o advwebgridexcel irá pegar as informações para exibir no delphi. No final este arquivo é deletado }
        lblsheet.Caption := flImportarproblema.Filename;
        flImportarproblema.SaveToFile
          ('C:\Users\Public\Documents' + edtEntername.Text + '.xls');
        // salva o arquivo na pasta

        TIWAdvWebGridExcelIO1.XLSImport
          ('C:\Users\Public\Documents' + edtEntername.Text + '.xls');

        TIWAdvWebGridExcelIO1.AdvWebGrid := grdImportproblema;
        // realiza a importação

        /// /////////////////////////////////////////////////////////////////////
        /// / Aplica uma correção no número de alternativas e critérios com base
        /// na possibilidade de erros quando da exclusão de dados de uma linha
        /// da tabela sem excluir a própria linha ////////////////////////////
        NumCri := (grdImportproblema.Columns.Count - 1);
        NumAlt := (grdImportproblema.TotalRows - 8);
        validar_planilha;
        if ACP = 0 then
        begin
          j := grdImportproblema.Columns.Count;
          // ** Deletando as colunas em branco **//
          for i := j - 1 Downto NumCri + 1 do
            grdImportproblema.Columns.Delete(i);
          grdImportproblema.TotalRows := NumAlt + 8;
        end
        else
          Exit;
        /// /////////////////////////////////////////////////////////////////////
        begin
          // setando os vetores que vão receber os dados da planilha importada
          Setlength(vtrnomecrit, grdImportproblema.Columns.Count - 1);
          Setlength(consreal1, grdImportproblema.TotalRows - 8,
            grdImportproblema.Columns.Count - 1);
          Setlength(vtrtipo, grdImportproblema.Columns.Count - 1);
          Setlength(tcrit, grdImportproblema.Columns.Count - 1);
          Setlength(vtrparametroA, grdImportproblema.Columns.Count - 1);
          Setlength(vtrparametroB, grdImportproblema.Columns.Count - 1);
          Setlength(niveisdisc, grdImportproblema.Columns.Count - 1);
          Setlength(vtrnomealternativa, grdImportproblema.TotalRows - 8);
          Setlength(TIWUserSession(WebApplication.Data).Verificaintra,
            grdImportproblema.Columns.Count - 1); // intracriterio
          Setlength(TIWUserSession(WebApplication.Data).DinamVerificaIntra,
            grdImportproblema.Columns.Count - 1); // intracriterio

          flImportarproblema.Visible := false;
          btnslotherfile.Enabled := true;

          // TIWAdvWebGridExcelIO1.AdvWebGrid := grdImportacao;
          // grdImportproblema.RowCount := grdImportacao.TotalRows;
          // grdImportproblema.Columns.Count := grdImportacao.ColumnCount;
          // passa as informações do grdImportação para o grdDadosImportados

          /// ////// VERIFICAÇÃO ALTERNATIVAS IGUAIS ////////////////////
          { with TIWUserSession(WebApplication.Data) do
            begin
            for i := 0 to (grdImportproblema.TotalRows - 8) - 1 do
            // quantidade de linhas ou NumAlt
            begin
            for k := 0 to (grdImportproblema.TotalRows - 8) - 1 do // Numalt
            begin
            VerificaAlt := 0;
            for j := 0 to (grdImportproblema.Columns.Count - 1) - 1 do
            // quantidade de colunas ou NumCrit
            begin
            if (k > i) AND (grdImportproblema.Cells[j + 1, i + 8]
            = grdImportproblema.Cells[j + 1, k + 8]) then
            begin
            VerificaAlt := VerificaAlt + 1;
            end;
            end; // for j
            // Depois de percorrer todas as colunas
            if VerificaAlt = grdImportproblema.Columns.Count - 1 then
            // Numcrit
            begin
            WebApplication.ShowMessage(
            'All consequences are the same in alternative ' +
            grdImportproblema.Cells[0, i + 8] { NomeAlt[i] }
          { + ' and ' + grdImportproblema.Cells[0, k + 8]
            { NomeAlt[K]
            +
            '. Please, verify the input data, make sure these alternatives are not the same, and then try again.');
            flImportarproblema.Visible := true;
            btnslotherfile.Enabled := true;
            Exit;
            end; // if
            end; // for k
            end;
            end;
            /// /////////////////////////////////////////////////////// }

          /// Recebendo o nome das alternativas ////
          for i := 0 to grdImportproblema.TotalRows - 9 do
          begin
            vtrnomealternativa[i] := grdImportproblema.Cells[0, i + 8];

            if (grdImportproblema.Cells[0, i + 8] = '') then

            Begin
              WebApplication.ShowMessage(
                'Please check the inserted alternatives names');

              flImportarproblema.Visible := true;
              btnslotherfile.Enabled := true;
              Exit;
            end; // end referente à mensagem e componentes

          end;

          /// / Recebendo os nomes dos critérios ////
          for i := 0 to grdImportproblema.Columns.Count - 2 do
          begin
            vtrnomecrit[i] := grdImportproblema.Cells[i + 1, 0];

            if (grdImportproblema.Cells[i + 1, 0] = '') then
            Begin
              WebApplication.ShowMessage(
                'Please check the inserted criteria names');

              flImportarproblema.Visible := true;
              btnslotherfile.Enabled := true;
              Exit;
            end; // end referente à mensagem e componentes
          end;

          Setlength(TIWUserSession(WebApplication.Data).Cri_inteiro, NumCri);
          /// / Recebendo os tipos de escala dos critérios ////
          for i := 0 to grdImportproblema.Columns.Count - 2 do
          begin
            tcrit[i] := StrToInt(grdImportproblema.Cells[i + 1, 1]);
            TIWUserSession(WebApplication.Data).Cri_inteiro[i] := false;
            if tcrit[i] = 4 then
            begin
              // ** Inteiro de minimização **//
              tcrit[i] := 0;
              TIWUserSession(WebApplication.Data).Cri_inteiro[i] := true;
            end;
            if tcrit[i] = 5 then
            begin
              // ** Inteiro de maximização **//
              tcrit[i] := 1;
              TIWUserSession(WebApplication.Data).Cri_inteiro[i] := true;
            end;
          end;

          { Numcolimport:= grdImportproblema.Columns.Count; //verificando se imputou modelo errdo de planilha
            for i := 0 to grdImportproblema.Columns.Count - 3 do
            if (grdImportproblema.Cells[Numcolimport + 1,1] = Floattostr(TIWUserSession(WebApplication.Data).Budget)) then
            Begin
            WebApplication.ShowMessage(
            'Please check the spreadsheet inserted');

            flImportarproblema.Visible := true;
            btnslotherfile.Enabled := true;
            Exit;
            End; //message }

          /// Recebendo os tipos de função valor  dos critérios ////

          for i := 0 to grdImportproblema.Columns.Count - 2 do
          begin
            if (tcrit[i] = 0) OR (tcrit[i] = 1) then
              if grdImportproblema.Cells[i + 1, 3] = '' then
              begin
                TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
                // se tem função sem declarar, boolean é verdadeira e auxilia nas proximas telas  exibidas na elicintra
                INC(TIWUserSession(WebApplication.Data).Contavazio); // se em algum momento verificaintra é verdadeira, incrementa e usa esse contador para habilitar o btn da rgnMenuIntra
              end
              else
                vtrtipo[i] := StrToInt(grdImportproblema.Cells[i + 1, 3]);
          end;

          /// / Recebendo as consequências das alternativas ////
          for i := 0 to grdImportproblema.TotalRows - 9 do
          begin
            for j := 0 to grdImportproblema.Columns.Count - 2 do
            begin // atribui os valores do grid impotação ao MatrizInput
              consreal1[i, j] := trocapv(grdImportproblema.Cells[j + 1, i + 8]);
              if (consreal1[i, j] = 0) and (vtrtipo[j] = 4) then
              begin
                if TIWUserSession(WebApplication.Data).Idioma = 1 then
                  WebApplication.ShowMessage(
                    'Como você está considerando uma função valor logística, as consequências no critério ' + vtrnomecrit[j]
                      + ' devem ser diferentes de 0. Por favor, verifique a matriz de consequências antes de continuar.')
                else
                  WebApplication.ShowMessage(
                    'As you are considering a Logistic Value Fuction, the consequences in the criterion ' + vtrnomecrit[j]
                      + ' cannot be equal to 0. Please, verify the consequences matrix before continuing.');
                flImportarproblema.Visible := true;
                btnslotherfile.Enabled := true;
                Exit;
              end;
            end; // for j
          end; // for i

          { if (tcrit[i] = 2) OR (tcrit[i] = 3) then // CritDisc
            begin
            if grdImportproblema.Cells[i + 1, 3] = '' then
            begin
            { if TIWUserSession(WebApplication.Data).Idioma = 1 then
            WebApplication.ShowMessage(
            'Por favor, informe a função valor para os critérios discretos.')
            else
            WebApplication.ShowMessage(
            'Please, fill the value functions for discrete criteria');

            Exit; }
          { vtrtipo[i] := 1;
            end
            else
            vtrtipo[i] := StrToInt(grdImportproblema.Cells[i + 1, 3]);
            end; }

          with TIWUserSession(WebApplication.Data) do
            /// Recebendo os valores parametro A dos criterios ñ lineares//
            for i := 0 to grdImportproblema.Columns.Count - 2 do
            begin
              if (grdImportproblema.Cells[i + 1, 4] = '') then
              begin
                if vtrtipo[i] <> 1 then
                begin
                  { WebApplication.ShowMessage(
                    'Please, fill the non-linear value functions parameters A and B in the input spreadsheet. In case of doubt check the user guide.');
                    btnslotherfileClick(Self);
                    Exit; }
                  Verificaintra[i] := true; // falta info, vai ter que elicitar
                  INC(Contavazio); // se  verificaintra=true, incrementa e usa esse contador para habilitar o btn da rgnMenuIntra

                end;
                if vtrtipo[i] = 1 then
                  vtrparametroA[i] := 0;
              end
              else
              begin
                if (trocapv(grdImportproblema.Cells[i + 1, 4]) = 0) or
                  (trocapv(grdImportproblema.Cells[i + 1, 5]) = 0) then
                  if (vtrtipo[i] <> 1) then
                  begin
                    if TIWUserSession(WebApplication.Data).Idioma = 1 then
                      WebApplication.ShowMessage(
                        'Os parâmetros "A" e "B" das funções valor não-lineares não podem assumir valor zero.  Em caso de dúvidas, consulte o guia do usuário.')
                    else
                      WebApplication.ShowMessage(
                        'The non-linear value functions parameters A and B cannot assume the value "0", please check it and try again. In case of doubt check the user guide.');
                    btnslotherfileClick(Self);
                    Exit;
                  end;
                vtrparametroA[i] := trocapv(grdImportproblema.Cells[i + 1, 4]);
              end;
            end;

          /// Recebendo os valores parametro B dos criterios ñ lineares//
          for i := 0 to grdImportproblema.Columns.Count - 2 do
          begin
            if (grdImportproblema.Cells[i + 1, 5] = '') then
            begin
              if vtrtipo[i] <> 1 then
              begin
                { WebApplication.ShowMessage(
                  'Please, fill the non-linear value functions parameters A and B in the input spreadsheet. In case of doubt check the user guide.');
                  btnslotherfileClick(Self);
                  Exit; }
                TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
                // falta info, vai ter que elicitar
                INC(TIWUserSession(WebApplication.Data).Contavazio); // se  verificaintra=true, incrementa e usa esse contador para habilitar o btn da rgnMenuIntra

              end;
              if vtrtipo[i] = 1 then
                vtrparametroB[i] := 0;
            end
            else
              vtrparametroB[i] := trocapv(grdImportproblema.Cells[i + 1, 5]);
          end;
          /// / Recebendo o número de níveis discretos dos critérios ////
          for i := 0 to grdImportproblema.Columns.Count - 2 do
          begin

            if (tcrit[i] = 2) or (tcrit[i] = 3) then
            begin
              if (StrToInt(grdImportproblema.Cells[i + 1, 6]) < 2) OR
                (StrToInt(grdImportproblema.Cells[i + 1, 6]) > 7) then
              begin

                if TIWUserSession(WebApplication.Data).Idioma = 1 then
                  WebApplication.ShowMessage(
                    'O parâmetro "C", que indica os níveis dos critérios discretos não pode assumir valores menores que 2 ou maiores que 7. Em caso de dúvidas, consulte o guia do usuário.')
                else
                  WebApplication.ShowMessage(
                    'The discrete criteria levels, parameter "C" cannot assume values beyond 2 and 7. In case of doubt check the user guide.');
                Exit;
              end;
              niveisdisc[i] := StrToInt(grdImportproblema.Cells[i + 1, 6]);
              // * Indica ao sistema que existem critérios discretos a serem validados *//
              TIWUserSession(WebApplication.Data).ValDiscreto := 5;
            end
            else
              niveisdisc[i] := 0;
          end;

          with TIWUserSession(WebApplication.Data) do
          begin
            NumCrit := length(tcrit);
            NumAlt := length(vtrnomealternativa);

            Setlength(NomeCrit, NumCrit);
            Setlength(FTipo, NumCrit);
            Setlength(Typeofcrit, NumCrit);
            Setlength(Niveis, NumCrit);
            Setlength(parA, NumCrit);
            Setlength(parB, NumCrit);
            Setlength(NomeAlt, NumAlt);
            Setlength(ConsMatrix, NumAlt, NumCrit);

            for i := 0 to NumCrit - 1 do
            begin
              NomeCrit[i] := vtrnomecrit[i];
              FTipo[i] := vtrtipo[i];
              Typeofcrit[i] := tcrit[i];
              Niveis[i] := niveisdisc[i];
              parA[i] := vtrparametroA[i];
              parB[i] := vtrparametroB[i];
            end;

            for j := 0 to NumAlt - 1 do
              NomeAlt[j] := vtrnomealternativa[j];

            for i := 0 to NumCrit - 1 do
              for j := 0 to NumAlt - 1 do
                ConsMatrix[j, i] := consreal1[j, i];
          end;

          for i := 0 to length(TIWUserSession(WebApplication.Data)
              .Verificaintra) - 1 do
            TIWUserSession(WebApplication.Data).DinamVerificaIntra[i] :=
              TIWUserSession(WebApplication.Data).Verificaintra[i];

          DeleteFile('C:\Users\Public\Documents' + edtEntername.Text + '.xls');

          // ** Permite a seleção de um novo problema ** //
          IWLabel13.Visible := true;
          IWLabel14.Visible := true;
          IWLabel15.Visible := true;
          // * Chama a validação dos discretos *//
          if TIWUserSession(WebApplication.Data).ValDiscreto = 5 then
            validardiscretos(TIWUserSession(WebApplication.Data).ValDiscreto);
          if TIWUserSession(WebApplication.Data).ValDiscreto <> -1 then
          begin
            AreThereDiscCriteria;
            lblsheet.Visible := true;
            if TIWUserSession(WebApplication.Data).Idioma = 1 then
              WebApplication.ShowMessage(
                'Planilha carregada com sucesso. Clique em "Continuar".')
            else
              WebApplication.ShowMessage(
                'Worksheet successfully loaded! Click on "Continue".');
            btnauxstartelicitation.Font.Style := [fsbold];
            btnauxstartelicitation.Enabled := true;
            btnImporttabela.Enabled := false;
            // ativa botão para inicio de elicitação

            // deleta o arquivo salvo na pasta
            // flImportarproblema.Destroy;

          end;
        end; // se o arquivo for compatível
      end // importação excel quando escolha ou ordenação
      else
      begin // se o arquivo não for compatível com a extensão .xls
        if TIWUserSession(WebApplication.Data).Idioma = 1 then
          WebApplication.ShowMessage('Verifique o tipo do arquivo.')
        else
          WebApplication.ShowMessage('Check the type of the file.');
      end;
    end;
  end; // user session
end; // procedimento

procedure TFrmEntrada.ImporttabelaClassificacao;
var
  tipoarquivo: string;
  i, j, k, VerificaAlt, ACP: integer;
  NumCri, NumAlt, Numcolimport, numero_alt, numero_crit: integer;
  vtrtipo, tcrit, niveisdisc: array of integer;
  vtrnomecrit, vtrpesos, vtrnomealternativa: array of string;
  vtrparametroA, vtrparametroB: array of Double;
  consreal1: array of array of Double;

  /// //////////////////////////////////////////////////////////////////////////
  /// / Procedimento especial para evitar erros ao apagar linhas e colunas da
  /// planilha do excel sem excluí-las ////////////////////////////////////////
  procedure validar_planilha;
  /// / Esse procedimento valida o número real de alternativas e critérios introdu-
  /// zidos na planilha de input //////////////////////////////////////////////
  var
    i, j, k: integer;
    StrAux: string;
    anterior: boolean;
    /// / Verdadeiro quando o nome de um critério ou alternativa
    /// não foi preenchido //////////////////////////////////////////////////////
  begin
    k := 0;
    anterior := false;
    ACP := 0;

    for i := 0 to NumCri - 1 do
    begin
      StrAux := grdImportproblema.Cells[i + 1, 0];
      if StrAux <> '' then
      begin
        INC(k);
        /// / Caso o critério atual teve o nome preenchido e o anterior não, então
        /// deve-se alertar ao Decisor que existe um problema na matriz de input
        if anterior = true then
        begin
          ACP := 1;
          anterior := false;
          WebApplication.ShowMessage(
            'There is at least one criterion name missing, please verify the Excel spreadsheet before continuing.');
          Break;
        end; // Anterior
      end // strAux
      else
        anterior := true;
    end; // for i
    if ACP = 0 then
      NumCri := k;
    k := 0;
    anterior := false;
    for j := 0 to NumAlt - 1 do
    begin
      StrAux := grdImportproblema.Cells[0, j + 8];
      if StrAux <> '' then
      begin
        INC(k);

        for i := 0 to NumCri - 1 do
        begin
          StrAux := grdImportproblema.Cells[i + 1, j + 8];
          if StrAux = '' then
            // ** Neste caso está faltando alguma consequência **//
            anterior := true;
        end;

        /// / Caso a alternativa atual teve o nome preenchido e a anterior não, então
        /// deve-se alertar ao Decisor que existe um problema na matriz de input
        if anterior = true then
        begin
          anterior := false;
          // ** Caso já haja problema nos critérios **//
          if ACP = 1 then
            ACP := 3;
          // ** Caso os critérios estejam ok **//
          if ACP = 0 then
          begin
            ACP := 2;
            WebApplication.ShowMessage(
              'There is at least one alternative name or consequence missing, please verify the Excel spreadsheet before continuing.');
          end;
          Break;
        end; // anterior = true
      end // atrAux
      else
        anterior := true;
    end; // for j
    if ACP = 0 then
      NumAlt := k;
  end;

/// ///////////////////////////////////////////////////////////////////////////
begin // procedure
  IWLabel13.Visible := true;
  IWLabel14.Visible := true;
  IWLabel15.Visible := true;
  TIWUserSession(WebApplication.Data).elicitarperfis := true;
  TIWUserSession(WebApplication.Data).ValDiscreto := 0;
  if edtEntername.Text = '' then // verifica se foi inserido um nome pro arquivo
  begin // se não tiver nome do problema
    WebApplication.ShowMessage('Please, enter a problem''s name!');
    Exit;
  end // se não tiver nome do problema

  else
  begin // se tiver nome do problema
    if flImportarproblema.Filename = '' then
    begin // se o file estiver preenchido
      WebApplication.ShowMessage('You should choose an excel file.');
    end // se o file não estiver preenchido

    else
    begin // se o file estiver preenchido
      tipoarquivo := flImportarproblema.ContentType;
      if (tipoarquivo = ' application/vnd.ms-excel') then
      // verifica se o arquivo é compatível

      begin
        { se o arquivo for compatível é criado um arquivo de excel no diretório especificado, igual ao
          selecionado pelo usuário. É de la que o advwebgridexcel irá pegar as informações para exibir no delphi. No final este arquivo é deletado }
        lblsheet.Caption := flImportarproblema.Filename;
        flImportarproblema.SaveToFile
          ('C:\Users\Public\Documents' + edtEntername.Text + '.xls');
        // salva o arquivo na pasta

        TIWAdvWebGridExcelIO1.XLSImport
          ('C:\Users\Public\Documents' + edtEntername.Text + '.xls');

        TIWAdvWebGridExcelIO1.AdvWebGrid := grdImportproblema;
        // realiza a importação

        { numero_alt := grdImportproblema.TotalRows - 8;
          numero_crit := grdImportproblema.Columns.Count - 1; }
        { try
          /// verifica se na primeira celula da ultima coluna tem um numero real
          /// caso tenha, preenche o numero de criterios e o vetor de perfis
          StrToFloat
          (grdImportproblema.Cells[grdImportproblema.Columns.Count - 1, 0]);
          numero_crit := grdImportproblema.Columns.Count - 2;
          for i := 0 to numero_alt - 1 do
          if grdImportproblema.Cells[grdImportproblema.Columns.Count - 1, i]
          <> '' then
          with TIWUserSession(WebApplication.Data) do
          begin
          setlength(perfis, length(perfis) + 1);
          for j := length(perfis) - 1 downto 1 do
          perfis[j] := perfis[j - 1];
          perfis[0] := StrToFloat
          (grdImportproblema.Cells[grdImportproblema.Columns.Count - 1,
          i]);
          end
          else
          begin
          break;
          end;
          except
          numero_crit := grdImportproblema.Columns.Count - 1;
          end; }
        NumCri := (grdImportproblema.Columns.Count - 1);
        NumAlt := (grdImportproblema.TotalRows - 8);
        validar_planilha;

        if ACP = 0 then
        begin
          j := grdImportproblema.Columns.Count;
          // ** Deletando as colunas em branco **//
          for i := j - 1 Downto NumCri + 1 do
            grdImportproblema.Columns.Delete(i);
          grdImportproblema.TotalRows := NumAlt + 8;
        end
        else
          Exit;

        /// /////////////////////////////////////////////////////////////////////
        begin
          // setando os vetores que vão receber os dados da planilha importada
          Setlength(vtrnomecrit, grdImportproblema.Columns.Count - 1);
          Setlength(consreal1, grdImportproblema.TotalRows - 8,
            grdImportproblema.Columns.Count - 1);
          Setlength(vtrtipo, grdImportproblema.Columns.Count - 1);
          Setlength(tcrit, grdImportproblema.Columns.Count - 1);
          Setlength(vtrparametroA, grdImportproblema.Columns.Count - 1);
          Setlength(vtrparametroB, grdImportproblema.Columns.Count - 1);
          Setlength(niveisdisc, grdImportproblema.Columns.Count - 1);
          Setlength(vtrnomealternativa, grdImportproblema.TotalRows - 8);
          Setlength(TIWUserSession(WebApplication.Data).Verificaintra,
            grdImportproblema.Columns.Count - 1); // intracriterio
          Setlength(TIWUserSession(WebApplication.Data).DinamVerificaIntra,
            grdImportproblema.Columns.Count - 1); // intracriterio

          flImportarproblema.Visible := false;
          btnslotherfile.Enabled := true;

          // TIWAdvWebGridExcelIO1.AdvWebGrid := grdImportacao;
          // grdImportproblema.RowCount := grdImportacao.TotalRows;
          // grdImportproblema.Columns.Count := grdImportacao.ColumnCount;
          // passa as informações do grdImportação para o grdDadosImportados

          /// ////// VERIFICAÇÃO ALTERNATIVAS IGUAIS ////////////////////
          { with TIWUserSession(WebApplication.Data) do
            begin
            for i := 0 to (grdImportproblema.TotalRows - 8) - 1 do
            // quantidade de linhas ou NumAlt
            begin
            for k := 0 to (grdImportproblema.TotalRows - 8) - 1 do // Numalt
            begin
            VerificaAlt := 0;
            for j := 0 to (numero_crit) - 1 do
            // quantidade de colunas ou NumCrit
            begin
            if (k > i) AND (grdImportproblema.Cells[j + 1, i + 8]
            = grdImportproblema.Cells[j + 1, k + 8]) then
            begin
            VerificaAlt := VerificaAlt + 1;
            end;
            end; // for j
            // Depois de percorrer todas as colunas
            if VerificaAlt = numero_crit then
            // Numcrit
            begin
            WebApplication.ShowMessage(
            'All consequences are the same in alternative ' +
            grdImportproblema.Cells[0, i + 8] { NomeAlt[i]
            + ' and ' + grdImportproblema.Cells[0, k + 8]
            { NomeAlt[K]
            +
            '. Please, verify the input data, make sure these alternatives are not the same, and then try again.');
            flImportarproblema.Visible := true;
            btnslotherfile.Enabled := true;
            Exit;
            end; // if
            end; // for k
            end;
            end;
            /// /////////////////////////////////////////////////////// }

          /// Recebendo o nome das alternativas ////
          for i := 0 to grdImportproblema.TotalRows - 9 do
          begin
            vtrnomealternativa[i] := grdImportproblema.Cells[0, i + 8];

            if (grdImportproblema.Cells[0, i + 8] = '') then

            Begin
              WebApplication.ShowMessage(
                'Please check the inserted alternatives names');

              flImportarproblema.Visible := true;
              btnslotherfile.Enabled := true;
              Exit;
            end; // end referente à mensagem e componentes

          end;

          /// / Recebendo os nomes dos critérios ////
          for i := 0 to grdImportproblema.Columns.Count - 2 do
          begin
            vtrnomecrit[i] := grdImportproblema.Cells[i + 1, 0];

            if (grdImportproblema.Cells[i + 1, 0] = '') then
            Begin
              WebApplication.ShowMessage(
                'Please check the inserted criteria names');

              flImportarproblema.Visible := true;
              btnslotherfile.Enabled := true;
              Exit;
            end; // end referente à mensagem e componentes
          end;

          { Numcolimport:= grdImportproblema.Columns.Count; //verificando se imputou modelo errdo de planilha
            for i := 0 to grdImportproblema.Columns.Count - 3 do
            if (grdImportproblema.Cells[Numcolimport + 1,1] = Floattostr(TIWUserSession(WebApplication.Data).Budget)) then
            Begin
            WebApplication.ShowMessage(
            'Please check the spreadsheet inserted');

            flImportarproblema.Visible := true;
            btnslotherfile.Enabled := true;
            Exit;
            End; //message }

          /// /////////////////////////////////////////////////////// }
          Setlength(TIWUserSession(WebApplication.Data).Cri_inteiro, NumCri);
          /// / Recebendo os tipos de escala dos critérios ////
          for i := 0 to grdImportproblema.Columns.Count - 2 do
          begin
            tcrit[i] := StrToInt(grdImportproblema.Cells[i + 1, 1]);
            TIWUserSession(WebApplication.Data).Cri_inteiro[i] := false;
            if tcrit[i] = 4 then
            begin
              // ** Inteiro de minimização **//
              tcrit[i] := 0;
              TIWUserSession(WebApplication.Data).Cri_inteiro[i] := true;
            end;
            if tcrit[i] = 5 then
            begin
              // ** Inteiro de maximização **//
              tcrit[i] := 1;
              TIWUserSession(WebApplication.Data).Cri_inteiro[i] := true;
            end;
          end;

          /// Recebendo os tipos de função valor  dos critérios ////

          for i := 0 to NumCri - 1 do
          begin

            if grdImportproblema.Cells[i + 1, 3] = '' then
            begin
              TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
              // se tem função sem declarar, boolean é verdadeira e auxilia nas proximas telas  exibidas na elicintra
              INC(TIWUserSession(WebApplication.Data).Contavazio); // se em algum momento verificaintra é verdadeira, incrementa e usa esse contador para habilitar o btn da rgnMenuIntra
            end
            else
              vtrtipo[i] := StrToInt(grdImportproblema.Cells[i + 1, 3]);

          END;

          with TIWUserSession(WebApplication.Data) do
            /// Recebendo os valores parametro A dos criterios ñ lineares//
            for i := 0 to NumCri - 1 do
            begin
              if (grdImportproblema.Cells[i + 1, 4] = '') then
              begin
                if vtrtipo[i] <> 1 then
                begin
                  { WebApplication.ShowMessage(
                    'Please, fill the non-linear value functions parameters A and B in the input spreadsheet. In case of doubt check the user guide.');
                    btnslotherfileClick(Self);
                    Exit; }
                  Verificaintra[i] := true; // falta info, vai ter que elicitar
                  INC(Contavazio); // se  verificaintra=true, incrementa e usa esse contador para habilitar o btn da rgnMenuIntra

                end;
                if vtrtipo[i] = 1 then
                  vtrparametroA[i] := 0;
              end
              else
              begin
                { if (trocapv(grdImportproblema.Cells[i + 1, 4]) = 0) or
                  (trocapv(grdImportproblema.Cells[i + 1, 5]) = 0) then
                  if (vtrtipo[i] <> 1) then
                  begin
                  WebApplication.ShowMessage(
                  'The non-linear value functions parameters A and B cannot assume the value "0", please check it and try again. In case of doubt check the user guide.');
                  btnslotherfileClick(Self);
                  Exit;
                  end; }
                vtrparametroA[i] := trocapv(grdImportproblema.Cells[i + 1, 4]);
              end;
            end;

          /// Recebendo os valores parametro B dos criterios ñ lineares//
          for i := 0 to NumCri - 1 do
          begin
            if (grdImportproblema.Cells[i + 1, 5] = '') then
            begin
              if vtrtipo[i] <> 1 then
              begin
                { WebApplication.ShowMessage(
                  'Please, fill the non-linear value functions parameters A and B in the input spreadsheet. In case of doubt check the user guide.');
                  btnslotherfileClick(Self);
                  Exit; }
                TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
                // falta info, vai ter que elicitar
                INC(TIWUserSession(WebApplication.Data).Contavazio); // se  verificaintra=true, incrementa e usa esse contador para habilitar o btn da rgnMenuIntra

              end;
              if vtrtipo[i] = 1 then
                vtrparametroB[i] := 0;
            end
            else
              vtrparametroB[i] := trocapv(grdImportproblema.Cells[i + 1, 5]);
          end;

          /// / Recebendo as consequências das alternativas ////
          for i := 0 to grdImportproblema.TotalRows - 9 do
          begin
            for j := 0 to grdImportproblema.Columns.Count - 2 do
            begin // atribui os valores do grid impotação ao MatrizInput
              consreal1[i, j] := trocapv(grdImportproblema.Cells[j + 1, i + 8]);
              if (consreal1[i, j] = 0) and (vtrtipo[j] = 4) then
              begin
                if TIWUserSession(WebApplication.Data).Idioma = 1 then
                  WebApplication.ShowMessage(
                    'Como você está considerando uma função valor logística, as consequências no critério ' + vtrnomecrit[j]
                      + ' devem ser diferentes de 0. Por favor, verifique a matriz de consequências antes de continuar.')
                else
                  WebApplication.ShowMessage(
                    'As you are considering a Logistic Value Fuction, the consequences in the criterion ' + vtrnomecrit[j]
                      + ' cannot be equal to 0. Please, verify the consequences matrix before continuing.');
                flImportarproblema.Visible := true;
                btnslotherfile.Enabled := true;
                Exit;
              end;
            end; // for j
          end; // for i

          with TIWUserSession(WebApplication.Data) do
            /// / Recebendo o número de níveis discretos dos critérios ////
            for i := 0 to NumCri - 1 do
            begin
              if (tcrit[i] = 2) or (tcrit[i] = 3) then
              begin
                niveisdisc[i] := StrToInt(grdImportproblema.Cells[i + 1, 6]);
                // * Indica ao sistema que existem critérios discretos a serem validados *//
                TIWUserSession(WebApplication.Data).ValDiscreto := 5;
              end
              else
                niveisdisc[i] := 0;
            end;

          with TIWUserSession(WebApplication.Data) do
          begin
            NumCrit := NumCri;
            NumAlt := (grdImportproblema.TotalRows - 8);

            Setlength(NomeCrit, NumCrit);
            Setlength(FTipo, NumCrit);
            Setlength(Typeofcrit, NumCrit);
            Setlength(Niveis, NumCrit);
            Setlength(parA, NumCrit);
            Setlength(parB, NumCrit);
            Setlength(NomeAlt, NumAlt);
            Setlength(ConsMatrix, NumAlt, NumCrit);

            for i := 0 to NumCrit - 1 do
            begin
              NomeCrit[i] := vtrnomecrit[i];
              FTipo[i] := vtrtipo[i];
              Typeofcrit[i] := tcrit[i];
              Niveis[i] := niveisdisc[i];
              parA[i] := vtrparametroA[i];
              parB[i] := vtrparametroB[i];
            end;

            for j := 0 to NumAlt - 1 do
              NomeAlt[j] := vtrnomealternativa[j];

            for i := 0 to NumCrit - 1 do
              for j := 0 to NumAlt - 1 do
                ConsMatrix[j, i] := consreal1[j, i];
          end;

          for i := 0 to length(TIWUserSession(WebApplication.Data)
              .Verificaintra) - 1 do
            TIWUserSession(WebApplication.Data).DinamVerificaIntra[i] :=
              TIWUserSession(WebApplication.Data).Verificaintra[i];

          DeleteFile('C:\Users\Public\Documents' + edtEntername.Text + '.xls');

          // ** Permite a seleção de um novo problema ** //
          IWLabel13.Visible := true;
          IWLabel14.Visible := true;
          IWLabel15.Visible := true;
          // * Chama a validação dos discretos *//
          if TIWUserSession(WebApplication.Data).ValDiscreto = 5 then
            validardiscretos(TIWUserSession(WebApplication.Data).ValDiscreto);
          if TIWUserSession(WebApplication.Data).ValDiscreto <> -1 then
          begin
            AreThereDiscCriteria;
            lblsheet.Visible := true;
            WebApplication.ShowMessage(
              'Worksheet successfully loaded! Click on "Continue".');
            btnauxstartelicitation.Font.Style := [fsbold];
            btnauxstartelicitation.Enabled := true;
            btnImporttabela.Enabled := false;
            // ativa botão para inicio de elicitação

            // deleta o arquivo salvo na pasta
            // flImportarproblema.Destroy;

          end;
        end; // se o arquivo for compatível
      end // importação excel quando escolha ou ordenação
      else
      begin // se o arquivo não for compatível com a extensão .xls
        WebApplication.ShowMessage('check the type of the file');
      end;
    end;
  end; // user session

end;

procedure TFrmEntrada.ImporttabelaPortfolioCB; // procedimento para importação de excel na problematica de portfolio CB
var
  tipoarquivo: string;
  i, j, k, VerificaAlt, ACP: integer;
  NumCri, NumAlt1: integer;
  vtrtipo, tcrit, niveisdisc: array of integer;
  vtrnomecrit, vtrpesos, vtrnomealternativa: array of string;
  vtrparametroA, vtrparametroB: array of Double;
  consreal1: array of array of Double;
  TotBudget, VerificaCost: Double; // totbudget- var que receberá valor total disponível para o portfolio
  /// verificaCost- auxiliar par afazer a verificação da diferença max de 10^4, entre os custos unitarios do vetor
  Costproj: array of Double; // var que recebe custo unt por projeto
  Numcolimport: integer; // recebe o numero de colunas do grid importproblem para preencher corretamente budget
  auxcost: string;
  /// //////////////////////////////////////////////////////////////////////////
  /// / Procedimento especial para evitar erros ao apagar linhas e colunas da
  /// planilha do excel sem excluí-las ////////////////////////////////////////
  procedure validar_planilha;
  /// / Esse procedimento valida o número real de alternativas e critérios introdu-
  /// zidos na planilha de input //////////////////////////////////////////////
  var
    i, j, k: integer;
    StrAux: string;
    anterior: boolean;
    /// / Verdadeiro quando o nome de um critério ou alternativa
    /// não foi preenchido //////////////////////////////////////////////////////
  begin
    k := 0;
    anterior := false;
    ACP := 0;

    for i := 0 to NumCri - 1 do
    begin
      StrAux := grdImportproblema.Cells[i + 1, 0];
      if StrAux <> '' then
      begin
        INC(k);
        /// / Caso o critério atual teve o nome preenchido e o anterior não, então
        /// deve-se alertar ao Decisor que existe um problema na matriz de input
        if anterior = true then
        begin
          ACP := 1;
          anterior := false;
          WebApplication.ShowMessage(
            'There is at least one criterion name missing, please verify the Excel spreadsheet before continuing.');
          Break;
        end; // Anterior
      end // strAux
      else
        anterior := true;
    end; // for i
    if ACP = 0 then
      NumCri := k;
    k := 0;
    anterior := false;
    for j := 0 to NumAlt1 - 1 do
    begin
      StrAux := grdImportproblema.Cells[0, j + 8];
      if StrAux <> '' then
      begin
        INC(k);

        for i := 0 to NumCri - 1 do
        begin
          StrAux := grdImportproblema.Cells[i + 1, j + 8];
          if StrAux = '' then
            // ** Neste caso está faltando alguma consequência **//
            anterior := true;
        end;

        /// / Caso a alternativa atual teve o nome preenchido e a anterior não, então
        /// deve-se alertar ao Decisor que existe um problema na matriz de input
        if anterior = true then
        begin
          anterior := false;
          // ** Caso já haja problema nos critérios **//
          if ACP = 1 then
            ACP := 3;
          // ** Caso os critérios estejam ok **//
          if ACP = 0 then
          begin
            ACP := 2;
            WebApplication.ShowMessage(
              'There is at least one alternative name or consequence missing, please verify the Excel spreadsheet before continuing.');
          end;
          Break;
        end; // anterior = true
      end // atrAux
      else
        anterior := true;
    end; // for j
    if ACP = 0 then
      NumAlt1 := k;
  end;

/// ///////////////////////////////////////////////////////////////////////////
Begin
  with TIWUserSession(WebApplication.Data) do
  begin // procedure
    IWLabel13.Visible := true;
    IWLabel14.Visible := true;
    IWLabel15.Visible := true;
    TIWUserSession(WebApplication.Data).ValDiscreto := 0;
    if edtEntername.Text = '' then // verifica se foi inserido um nome pro arquivo

    begin // se não tiver nome do problema
      WebApplication.ShowMessage('Please, enter a problem''s name!');
      Exit;
    end // se não tiver nome do problema

    else
    begin // se tiver nome do problema
      if flImportarproblema.Filename = '' then
      begin // se o file estiver preenchido
        WebApplication.ShowMessage('You should choose an excel file.');
      end // se o file não estiver preenchido

      else
      begin // se o file estiver preenchido
        tipoarquivo := flImportarproblema.ContentType;
        if (tipoarquivo = ' application/vnd.ms-excel') then
        // verifica se o arquivo é compatível

        begin
          { se o arquivo for compatível é criado um arquivo de excel no diretório especificado, igual ao
            selecionado pelo usuário. É de la que o advwebgridexcel irá pegar as informações para exibir no delphi. No final este arquivo é deletado }
          lblsheet.Caption := flImportarproblema.Filename;
          flImportarproblema.SaveToFile
            ('C:\Users\Public\Documents' + edtEntername.Text + '.xls');
          // salva o arquivo na pasta

          TIWAdvWebGridExcelIO1.XLSImport
            ('C:\Users\Public\Documents' + edtEntername.Text + '.xls');

          TIWAdvWebGridExcelIO1.AdvWebGrid := grdImportproblema;
          // realiza a importação

          NumCri := (grdImportproblema.Columns.Count - 1);
          NumAlt1 := (grdImportproblema.TotalRows - 8);
          validar_planilha;

          if ACP = 0 then
          begin
            j := grdImportproblema.Columns.Count;
            // ** Deletando as colunas em branco **//
            for i := j - 1 Downto NumCri + 1 do
              grdImportproblema.Columns.Delete(i);
            grdImportproblema.TotalRows := NumAlt1 + 8;
          end
          else
            Exit;

          begin

            // setando os vetores que vão receber os dados da planilha importada
            Setlength(vtrnomecrit, grdImportproblema.Columns.Count - 2);
            Setlength(consreal1, grdImportproblema.TotalRows - 8,
              grdImportproblema.Columns.Count - 2);
            Setlength(vtrtipo, grdImportproblema.Columns.Count - 2);
            Setlength(tcrit, grdImportproblema.Columns.Count - 2);
            Setlength(vtrparametroA, grdImportproblema.Columns.Count - 2);
            Setlength(vtrparametroB, grdImportproblema.Columns.Count - 2);
            Setlength(niveisdisc, grdImportproblema.Columns.Count - 2);
            Setlength(vtrnomealternativa, grdImportproblema.TotalRows - 8);
            /// //// setando vetor dos custos unitário CB
            Setlength(Costproj, grdImportproblema.TotalRows - 8);
            Setlength(TIWUserSession(WebApplication.Data).Verificaintra,
              grdImportproblema.Columns.Count - 2); // intracriterio
            Setlength(TIWUserSession(WebApplication.Data).DinamVerificaIntra,
              grdImportproblema.Columns.Count - 2); // intracriterio

            flImportarproblema.Visible := false;
            btnslotherfile.Enabled := true;

            /// ////// VERIFICAÇÃO ALTERNATIVAS IGUAIS ////////////////////
            { with TIWUserSession(WebApplication.Data) do
              begin
              for i := 0 to (grdImportproblema.TotalRows - 8) - 1 do
              // quantidade de linhas ou numalt1
              begin
              for K := 0 to (grdImportproblema.TotalRows - 8) - 1 do // numalt1
              begin
              VerificaAlt := 0;
              for j := 0 to (grdImportproblema.Columns.Count - 2) - 1 do
              // quantidade de colunas ou NumCrit
              begin
              if (K > i) AND (grdImportproblema.Cells[j + 1, i + 8]
              = grdImportproblema.Cells[j + 1, K + 8]) then
              begin
              VerificaAlt := VerificaAlt + 1;
              end;
              end; // for j
              // Depois de percorrer todas as colunas
              if VerificaAlt = grdImportproblema.Columns.Count - 2 then
              // Numcrit
              begin
              WebApplication.ShowMessage(
              'All consequences are the same in alternative ' +
              grdImportproblema.Cells[0, i + 8] { NomeAlt[i] }
            { + ' and ' + grdImportproblema.Cells[0, K + 8]
              { NomeAlt[K] }
            { +
              '. Please, verify the input data, make sure these alternatives are not the same, and then try again.');
              flImportarproblema.Visible := true;
              btnslotherfile.Enabled := true;
              Exit;
              end; // if
              end; // for k
              end;
              end; }
            /// ///////////////////////////////////////////////////////

            /// / Recebendo as consequências das alternativas ////
            for i := 0 to grdImportproblema.TotalRows - 9 do
            begin
              for j := 0 to grdImportproblema.Columns.Count - 3 do
              begin // atribui os valores do grid impotação ao MatrizInput
                consreal1[i, j] := trocapv
                  (grdImportproblema.Cells[j + 1, i + 8]);

              end; // for j
            end; // for i

            /// Recebendo o nome das alternativas ////
            for i := 0 to grdImportproblema.TotalRows - 9 do
            begin
              vtrnomealternativa[i] := grdImportproblema.Cells[0, i + 8];

              if (grdImportproblema.Cells[0, i + 8] = '') then

              Begin
                WebApplication.ShowMessage(
                  'Please check the inserted alternatives names');

                flImportarproblema.Visible := true;
                btnslotherfile.Enabled := true;
                Exit;
              end; // end referente à mensagem e componentes
            end;

            /// /////////////////////////////////////////////////////// }
            Setlength(TIWUserSession(WebApplication.Data).Cri_inteiro,
              NumCri - 1);
            /// / Recebendo os tipos de escala dos critérios ////
            for i := 0 to grdImportproblema.Columns.Count - 3 do
            begin
              tcrit[i] := StrToInt(grdImportproblema.Cells[i + 1, 1]);
              TIWUserSession(WebApplication.Data).Cri_inteiro[i] := false;
              if tcrit[i] = 4 then
              begin
                // ** Inteiro de minimização **//
                tcrit[i] := 0;
                TIWUserSession(WebApplication.Data).Cri_inteiro[i] := true;
              end;
              if tcrit[i] = 5 then
              begin
                // ** Inteiro de maximização **//
                tcrit[i] := 1;
                TIWUserSession(WebApplication.Data).Cri_inteiro[i] := true;
              end;
            end;

            /// // Recebendo custos unitarios dos projetos da carteira /////
            Numcolimport := grdImportproblema.Columns.Count;

            for i := 0 to grdImportproblema.TotalRows - 9 do
            begin
              if (grdImportproblema.Cells[Numcolimport - 1, i + 8] = '') then
              // verificando se células estão preenchidas

              Begin
                WebApplication.ShowMessage(
                  'Please check the unit costs inserted in the projects');

                flImportarproblema.Visible := true;
                btnslotherfile.Enabled := true;
                Exit;
              end; // end referente à mensagem e componentes

              Costproj[i] := trocapv
                (grdImportproblema.Cells[Numcolimport - 1, i + 8]);

            end; // for i

            begin // verificação 10^4
              // ** Identificando o custos**//
              TIWUserSession(WebApplication.Data).MaxCost := Costproj[0];
              TIWUserSession(WebApplication.Data).MinCost := Costproj[0];

              for j := 0 to grdImportproblema.TotalRows - 9 do
              begin
                if Costproj[j] > TIWUserSession(WebApplication.Data)
                  .MaxCost then
                  TIWUserSession(WebApplication.Data).MaxCost := Costproj[j];
                if Costproj[j] < TIWUserSession(WebApplication.Data)
                  .MinCost then
                  TIWUserSession(WebApplication.Data).MinCost := Costproj[j];
              end;

              begin
                VerificaCost := MaxCost / MinCost;
                if (VerificaCost > 100000) then
                begin
                  WebApplication.ShowMessage(
                    'This version of the system supports a range between alternative costs of less than 10^4. We apologize for the inconvenience.Please, check your file');
                  flImportarproblema.Visible := true;
                  btnslotherfile.Enabled := true;
                  Exit;
                End; // show message
              end; // begin verificacost
            end; // verificação 10^4

            /// ///Recebendo valor do budget disponível para o portfolio////
            Numcolimport := grdImportproblema.Columns.Count;
            begin

              if (grdImportproblema.Cells[Numcolimport - 1, 1] = '') then
              // VERIFICANDO SE VALOR FOI INSERIDO
              begin
                WebApplication.ShowMessage(
                  'Please check the Budget value inserted.');
              end; // showmensage

              TotBudget := trocapv
                (grdImportproblema.Cells[Numcolimport - 1, 1]);
              budget := TotBudget; // var global recebe valor local

            end; // if

            /// / Recebendo os nomes dos critérios ////
            for i := 0 to grdImportproblema.Columns.Count - 3 do
            begin
              vtrnomecrit[i] := grdImportproblema.Cells[i + 1, 0];

              if (grdImportproblema.Cells[i + 1, 0] = '') then
              Begin
                WebApplication.ShowMessage(
                  'Please check the inserted criteria names');

                flImportarproblema.Visible := true;
                btnslotherfile.Enabled := true;
                Exit;
              end; // end referente à mensagem e componentes
            end;

            /// Recebendo os tipos de função valor  dos critérios ////

            for i := 0 to grdImportproblema.Columns.Count - 2 do
            begin
              if grdImportproblema.Cells[i + 1, 3] = '' then
              begin
                TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
                // se tem função sem declarar, boolean é verdadeira e auxilia nas proximas telas  exibidas na elicintra
                INC(TIWUserSession(WebApplication.Data).Contavazio); // se em algum momento verificaintra é verdadeira, incrementa e usa esse contador para habilitar o btn da rgnMenuIntra
              end
              else
                vtrtipo[i] := StrToInt(grdImportproblema.Cells[i + 1, 3]);
            end;

            /// Recebendo os valores parametro A dos criterios ñ lineares//
            for i := 0 to grdImportproblema.Columns.Count - 3 do
            begin
              if (grdImportproblema.Cells[i + 1, 4] = '') then
              begin
                if vtrtipo[i] <> 1 then
                begin
                  { WebApplication.ShowMessage(
                    'Please, fill the non-linear value functions parameters A and B in the input spreadsheet. In case of doubt check the user guide.');
                    btnslotherfileClick(Self);
                    Exit; }
                  Verificaintra[i] := true; // falta info, vai ter que elicitar
                  INC(Contavazio); // se  verificaintra=true, incrementa e usa esse contador para habilitar o btn da rgnMenuIntra
                end;

                if vtrtipo[i] = 1 then
                  vtrparametroA[i] := 0;
              end
              else
              begin
                if (trocapv(grdImportproblema.Cells[i + 1, 4]) = 0) or
                  (trocapv(grdImportproblema.Cells[i + 1, 5]) = 0) then
                  if (vtrtipo[i] <> 1) then
                  begin
                    WebApplication.ShowMessage(
                      'The non-linear value functions parameters A and B cannot assume the value "0", please check it and try again. In case of doubt check the user guide.');
                    btnslotherfileClick(Self);
                    Exit;
                  end;
                vtrparametroA[i] := trocapv(grdImportproblema.Cells[i + 1, 4]);
              end;
            end;

            /// Recebendo os valores parametro B dos criterios ñ lineares//
            for i := 0 to grdImportproblema.Columns.Count - 3 do
            begin
              if (grdImportproblema.Cells[i + 1, 5] = '') then
              begin
                if vtrtipo[i] <> 1 then
                begin
                  { WebApplication.ShowMessage(
                    'Please, fill the non-linear value functions parameters A and B in the input spreadsheet. In case of doubt check the user guide.');
                    btnslotherfileClick(Self);
                    Exit; }
                  TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
                  // falta info, vai ter que elicitar
                  INC(TIWUserSession(WebApplication.Data).Contavazio); // se  	verificaintra=true, incrementa e usa esse contador para habilitar o btn da rgnMenuIntra
                end;
                if vtrtipo[i] = 1 then
                  vtrparametroB[i] := 0;
              end
              else
                vtrparametroB[i] := trocapv(grdImportproblema.Cells[i + 1, 5]);
            end;

            /// / Recebendo o número de níveis discretos dos critérios ////
            for i := 0 to grdImportproblema.Columns.Count - 3 do
            begin
              if (tcrit[i] = 2) or (tcrit[i] = 3) then
              begin
                if (StrToInt(grdImportproblema.Cells[i + 1, 6]) < 2) OR
                  (StrToInt(grdImportproblema.Cells[i + 1, 6]) > 7) then
                begin

                  if TIWUserSession(WebApplication.Data).Idioma = 1 then
                    WebApplication.ShowMessage(
                      'O parâmetro "C", que indica os níveis dos critérios discretos não pode assumir valores menores que 2 ou maiores que 7. Em caso de dúvidas, consulte o guia do usuário.')
                  else
                    WebApplication.ShowMessage(
                      'The discrete criteria levels, parameter "C" cannot assume values beyond 2 and 7. In case of doubt check the user guide.');
                  Exit;
                end;
                niveisdisc[i] := StrToInt(grdImportproblema.Cells[i + 1, 6]);
                // * Indica ao sistema que existem critérios discretos a serem validados *//
                TIWUserSession(WebApplication.Data).ValDiscreto := 5;
              end
              else
                niveisdisc[i] := 0;
            end;

            with TIWUserSession(WebApplication.Data) do
            begin
              TIWUserSession(WebApplication.Data).NumCrit := length(tcrit);
              TIWUserSession(WebApplication.Data).NumAlt := length
                (vtrnomealternativa);

              Setlength(NomeCrit, NumCrit);
              Setlength(FTipo, NumCrit);
              Setlength(Typeofcrit, NumCrit);
              Setlength(Niveis, NumCrit);
              Setlength(parA, NumCrit);
              Setlength(parB, NumCrit);
              Setlength(NomeAlt, NumAlt);
              Setlength(ConsMatrix, NumAlt, NumCrit);
              Setlength(Cost, NumAlt);

              for i := 0 to NumCrit - 1 do
              begin
                TIWUserSession(WebApplication.Data).NomeCrit[i] := vtrnomecrit
                  [i];
                TIWUserSession(WebApplication.Data).FTipo[i] := vtrtipo[i];
                TIWUserSession(WebApplication.Data).Typeofcrit[i] := tcrit[i];
                TIWUserSession(WebApplication.Data).Niveis[i] := niveisdisc[i];
                TIWUserSession(WebApplication.Data).parA[i] := vtrparametroA[i];
                TIWUserSession(WebApplication.Data).parB[i] := vtrparametroB[i];
              end;

              for j := 0 to NumAlt - 1 do
              Begin
                NomeAlt[j] := vtrnomealternativa[j];
                Cost[j] := Costproj[j];
                /// global recebe local portfolio
              end;

              for i := 0 to NumCrit - 1 do
                for j := 0 to NumAlt - 1 do
                  ConsMatrix[j, i] := consreal1[j, i];
            end;

            for i := 0 to length(TIWUserSession(WebApplication.Data)
                .Verificaintra) - 1 do
              TIWUserSession(WebApplication.Data).DinamVerificaIntra[i] :=
                TIWUserSession(WebApplication.Data).Verificaintra[i];

            DeleteFile
              ('C:\Users\Public\Documents' + edtEntername.Text + '.xls');

            // ** Permite a seleção de um novo problema ** //
            IWLabel13.Visible := true;
            IWLabel14.Visible := true;
            IWLabel15.Visible := true;
            // * Chama a validação dos discretos *//
            if TIWUserSession(WebApplication.Data).ValDiscreto = 5 then
              validardiscretos(TIWUserSession(WebApplication.Data).ValDiscreto);
            if TIWUserSession(WebApplication.Data).ValDiscreto <> -1 then
            begin
              lblsheet.Visible := true;
              WebApplication.ShowMessage(
                'Worksheet successfully loaded! Click on "Continue".');
              btnauxstartelicitation.Font.Style := [fsbold];
              btnauxstartelicitation.Enabled := true;
              btnImporttabela.Enabled := false;
              // ativa botão para inicio de elicitação

              // deleta o arquivo salvo na pasta
              // flImportarproblema.Destroy;

            end;
          end;
        end // se o arquivo for compatível
        else
        begin // se o arquivo não for compatível com a extensão .xls
          WebApplication.ShowMessage('check the type of the file');
        end;

      end;
    end;
  end; // UserSession
end; // procedimento

procedure TFrmEntrada.ImporttabelaPortfolioC; // procedimento para importação de excel na problematica de portfolio CB
var
  tipoarquivo: string;
  i, j, k, VerificaAlt, ACP: integer;
  NumCri, NumAlt: integer;
  vtrtipo, tcrit, niveisdisc: array of integer;
  vtrnomecrit, vtrpesos, vtrnomealternativa: array of string;
  vtrparametroA, vtrparametroB: array of Double;
  consreal1: array of array of Double;
  TotBudget, VerificaCost: Double; // totbudget- var que receberá valor total disponível para o portfolio
  /// verificaCost- auxiliar par afazer a verificação da diferença max de 10^4, entre os custos unitarios do vetor
  Costproj: array of Double; // var que recebe custo unt por projeto
  Numcolimport: integer; // recebe o numero de colunas do grid importproblem para preencher corretamente budget

  /// //////////////////////////////////////////////////////////////////////////
  /// / Procedimento especial para evitar erros ao apagar linhas e colunas da
  /// planilha do excel sem excluí-las ////////////////////////////////////////
  procedure validar_planilha;
  /// / Esse procedimento valida o número real de alternativas e critérios introdu-
  /// zidos na planilha de input //////////////////////////////////////////////
  var
    i, j, k: integer;
    StrAux: string;
    anterior: boolean;
    /// / Verdadeiro quando o nome de um critério ou alternativa
    /// não foi preenchido //////////////////////////////////////////////////////
  begin
    k := 0;
    anterior := false;
    ACP := 0;

    for i := 0 to NumCri - 1 do
    begin
      StrAux := grdImportproblema.Cells[i + 1, 0];
      if StrAux <> '' then
      begin
        INC(k);
        /// / Caso o critério atual teve o nome preenchido e o anterior não, então
        /// deve-se alertar ao Decisor que existe um problema na matriz de input
        if anterior = true then
        begin
          ACP := 1;
          anterior := false;
          WebApplication.ShowMessage(
            'There is at least one criterion name missing, please verify the Excel spreadsheet before continuing.');
          Break;
        end; // Anterior
      end // strAux
      else
        anterior := true;
    end; // for i
    if ACP = 0 then
      NumCri := k;
    k := 0;
    anterior := false;
    for j := 0 to NumAlt - 1 do
    begin
      StrAux := grdImportproblema.Cells[0, j + 8];
      if StrAux <> '' then
      begin
        INC(k);

        for i := 0 to NumCri - 1 do
        begin
          StrAux := grdImportproblema.Cells[i + 1, j + 8];
          if StrAux = '' then
            // ** Neste caso está faltando alguma consequência **//
            anterior := true;
        end;

        /// / Caso a alternativa atual teve o nome preenchido e a anterior não, então
        /// deve-se alertar ao Decisor que existe um problema na matriz de input
        if anterior = true then
        begin
          anterior := false;
          // ** Caso já haja problema nos critérios **//
          if ACP = 1 then
            ACP := 3;
          // ** Caso os critérios estejam ok **//
          if ACP = 0 then
          begin
            ACP := 2;
            WebApplication.ShowMessage(
              'There is at least one alternative name or consequence missing, please verify the Excel spreadsheet before continuing.');
          end;
          Break;
        end; // anterior = true
      end // atrAux
      else
        anterior := true;
    end; // for j
    if ACP = 0 then
      NumAlt := k;
  end;
/// ///////////////////////////////////////////////////////////////////////////

Begin
  with TIWUserSession(WebApplication.Data) do
  begin // procedure
    IWLabel13.Visible := true;
    IWLabel14.Visible := true;
    IWLabel15.Visible := true;
    TIWUserSession(WebApplication.Data).ValDiscreto := 0;
    if edtEntername.Text = '' then // verifica se foi inserido um nome pro arquivo

    begin // se não tiver nome do problema
      WebApplication.ShowMessage('Please, enter a problem''s name!');
      Exit;
    end // se não tiver nome do problema

    else
    begin // se tiver nome do problema
      if flImportarproblema.Filename = '' then
      begin // se o file estiver preenchido
        WebApplication.ShowMessage('You should choose an excel file.');
      end // se o file não estiver preenchido

      else
      begin // se o file estiver preenchido
        tipoarquivo := flImportarproblema.ContentType;
        if (tipoarquivo = ' application/vnd.ms-excel') then
        // verifica se o arquivo é compatível

        begin
          { se o arquivo for compatível é criado um arquivo de excel no diretório especificado, igual ao
            selecionado pelo usuário. É de la que o advwebgridexcel irá pegar as informações para exibir no delphi. No final este arquivo é deletado }
          lblsheet.Caption := flImportarproblema.Filename;
          flImportarproblema.SaveToFile
            ('C:\Users\Public\Documents' + edtEntername.Text + '.xls');
          // salva o arquivo na pasta

          TIWAdvWebGridExcelIO1.XLSImport
            ('C:\Users\Public\Documents' + edtEntername.Text + '.xls');

          TIWAdvWebGridExcelIO1.AdvWebGrid := grdImportproblema;
          // realiza a importação

          // setando os vetores que vão receber os dados da planilha importada
          Setlength(vtrnomecrit, grdImportproblema.Columns.Count - 2);
          Setlength(consreal1, grdImportproblema.TotalRows - 8,
            grdImportproblema.Columns.Count - 2);
          Setlength(vtrtipo, grdImportproblema.Columns.Count - 2);
          Setlength(tcrit, grdImportproblema.Columns.Count - 2);
          Setlength(vtrparametroA, grdImportproblema.Columns.Count - 2);
          Setlength(vtrparametroB, grdImportproblema.Columns.Count - 2);
          Setlength(niveisdisc, grdImportproblema.Columns.Count - 2);
          Setlength(vtrnomealternativa, grdImportproblema.TotalRows - 8);
          /// //// setando vetor dos custos unitário CB
          Setlength(Costproj, grdImportproblema.TotalRows - 8);
          Setlength(TIWUserSession(WebApplication.Data).Verificaintra,
            grdImportproblema.Columns.Count - 2); // intracriterio
          Setlength(TIWUserSession(WebApplication.Data).DinamVerificaIntra,
            grdImportproblema.Columns.Count - 2); // intracriterio

          flImportarproblema.Visible := false;
          btnslotherfile.Enabled := true;

          /// ////// VERIFICAÇÃO ALTERNATIVAS IGUAIS ////////////////////
          { with TIWUserSession(WebApplication.Data) do
            begin
            for i := 0 to (grdImportproblema.TotalRows - 8) - 1 do
            // quantidade de linhas ou NumAlt
            begin
            for k := 0 to (grdImportproblema.TotalRows - 8) - 1 do // Numalt
            begin
            VerificaAlt := 0;
            for j := 0 to (grdImportproblema.Columns.Count - 2) - 1 do
            // quantidade de colunas ou NumCrit
            begin
            if (k > i) AND (grdImportproblema.Cells[j + 1, i + 8]
            = grdImportproblema.Cells[j + 1, k + 8]) then
            begin
            VerificaAlt := VerificaAlt + 1;
            end;
            end; // for j
            // Depois de percorrer todas as colunas
            if VerificaAlt = grdImportproblema.Columns.Count - 2 then
            // Numcrit
            begin
            WebApplication.ShowMessage(
            'All consequences are the same in alternative ' +
            grdImportproblema.Cells[0, i + 8] { NomeAlt[i] }
          { + ' and ' + grdImportproblema.Cells[0, k + 8]
            { NomeAlt[K] }
          { +
            '. Please, verify the input data, make sure these alternatives are not the same, and then try again.');
            flImportarproblema.Visible := true;
            btnslotherfile.Enabled := true;
            Exit;
            end; // if
            end; // for k
            end;
            end;
            /// /////////////////////////////////////////////////////// }

          /// Recebendo o nome das alternativas ////
          for i := 0 to grdImportproblema.TotalRows - 9 do
          begin
            vtrnomealternativa[i] := grdImportproblema.Cells[0, i + 8];

            if (grdImportproblema.Cells[0, i + 8] = '') then

            Begin
              WebApplication.ShowMessage(
                'Please check the inserted alternatives names');

              flImportarproblema.Visible := true;
              btnslotherfile.Enabled := true;
              Exit;
            end; // end referente à mensagem e componentes
          end;

          /// // Recebendo custos unitarios dos projetos da carteira /////
          Numcolimport := grdImportproblema.Columns.Count;

          for i := 0 to grdImportproblema.TotalRows - 9 do
          begin
            if (grdImportproblema.Cells[Numcolimport - 1, i + 8] = '') then
            // verificando se células estão preenchidas

            Begin
              WebApplication.ShowMessage(
                'Please check the unit costs inserted in the projects');

              flImportarproblema.Visible := true;
              btnslotherfile.Enabled := true;
              Exit;
            end; // end referente à mensagem e componentes

            Costproj[i] := trocapv
              (grdImportproblema.Cells[Numcolimport - 1, i + 8]);

          end; // for i

          { begin // verificação 10^4
            // ** Identificando o custos**//
            TIWUserSession(WebApplication.Data).MaxCost := Costproj[0];
            TIWUserSession(WebApplication.Data).MinCost := Costproj[0];

            for j := 0 to grdImportproblema.TotalRows - 9 do
            begin
            if Costproj[j] > TIWUserSession(WebApplication.Data).MaxCost then
            TIWUserSession(WebApplication.Data).MaxCost := Costproj[j];
            if Costproj[j] < TIWUserSession(WebApplication.Data).MinCost then
            TIWUserSession(WebApplication.Data).MinCost := Costproj[j];
            end;

            begin
            VerificaCost := MaxCost / MinCost;
            if (VerificaCost > 100000) then
            begin
            WebApplication.ShowMessage(
            'This version of the system supports a range between alternative costs of less than 10^4. We apologize for the inconvenience.Please, check your file');
            flImportarproblema.Visible := true;
            btnslotherfile.Enabled := true;
            Exit;
            End; // show message
            end; // begin verificacost
            end; // verificação 10^4 }

          /// ///Recebendo valor do budget disponível para o portfolio////
          Numcolimport := grdImportproblema.Columns.Count;
          begin

            if (grdImportproblema.Cells[Numcolimport - 1, 1] = '') then
            // VERIFICANDO SE VALOR FOI INSERIDO
            begin
              WebApplication.ShowMessage(
                'Please check the Budget value inserted.');
            end; // showmensage

            TotBudget := trocapv(grdImportproblema.Cells[Numcolimport - 1, 1]);
            budget := TotBudget; // var global recebe valor local

          end; // if

          Setlength(TIWUserSession(WebApplication.Data).Cri_inteiro, NumCri);
          /// / Recebendo os tipos de escala dos critérios ////
          for i := 0 to grdImportproblema.Columns.Count - 2 do
          begin
            tcrit[i] := StrToInt(grdImportproblema.Cells[i + 1, 1]);
            TIWUserSession(WebApplication.Data).Cri_inteiro[i] := false;
            if tcrit[i] = 4 then
            begin
              // ** Inteiro de minimização **//
              tcrit[i] := 0;
              TIWUserSession(WebApplication.Data).Cri_inteiro[i] := true;
            end;
            if tcrit[i] = 5 then
            begin
              // ** Inteiro de maximização **//
              tcrit[i] := 1;
              TIWUserSession(WebApplication.Data).Cri_inteiro[i] := true;
            end;
          end;

          /// / Recebendo os nomes dos critérios ////
          for i := 0 to grdImportproblema.Columns.Count - 3 do
          begin
            vtrnomecrit[i] := grdImportproblema.Cells[i + 1, 0];

            if (grdImportproblema.Cells[i + 1, 0] = '') then
            Begin
              WebApplication.ShowMessage(
                'Please check the inserted criteria names');

              flImportarproblema.Visible := true;
              btnslotherfile.Enabled := true;
              Exit;
            end; // end referente à mensagem e componentes
          end;

          /// / Recebendo os tipos de função valor  dos critérios ////
          for i := 0 to grdImportproblema.Columns.Count - 3 do
          begin
            if grdImportproblema.Cells[i + 1, 3] = '' then
            begin
              TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
              // se tem função sem declarar, boolean é verdadeira e auxilia nas proximas telas  exibidas na elicintra
              INC(TIWUserSession(WebApplication.Data).Contavazio); // se em algum momento verificaintra é verdadeira, incrementa e usa esse contador para habilitar o btn da rgnMenuIntra
            end
            else
              vtrtipo[i] := StrToInt(grdImportproblema.Cells[i + 1, 3]);

            { if (tcrit[i] = 2) OR (tcrit[i] = 3) then // CritDisc
              begin
              if grdImportproblema.Cells[i + 1, 3] = '' then
              begin
              WebApplication.ShowMessage(
              'Please, fill the value functions for discrete criteria');
              Exit;
              end
              else
              vtrtipo[i] := StrToInt(grdImportproblema.Cells[i + 1, 3]);
              end; }
          END;

          with TIWUserSession(WebApplication.Data) do
            /// Recebendo os valores parametro A dos criterios ñ lineares//
            for i := 0 to grdImportproblema.Columns.Count - 3 do
            begin
              if (grdImportproblema.Cells[i + 1, 4] = '') then
              begin
                if vtrtipo[i] <> 1 then
                begin
                  { WebApplication.ShowMessage(
                    'Please, fill the non-linear value functions parameters A and B in the input spreadsheet. In case of doubt check the user guide.');
                    btnslotherfileClick(Self);
                    Exit; }
                  Verificaintra[i] := true; // falta info, vai ter que elicitar
                  INC(Contavazio); // se  verificaintra=true, incrementa e usa esse contador para habilitar o btn da rgnMenuIntra

                end;
                if vtrtipo[i] = 1 then
                  vtrparametroA[i] := 0;
              end
              else
              begin
                { if (trocapv(grdImportproblema.Cells[i + 1, 4]) = 0) or
                  (trocapv(grdImportproblema.Cells[i + 1, 5]) = 0) then
                  if (vtrtipo[i] <> 1) then
                  begin
                  WebApplication.ShowMessage(
                  'The non-linear value functions parameters A and B cannot assume the value "0", please check it and try again. In case of doubt check the user guide.');
                  btnslotherfileClick(Self);
                  Exit;
                  end; }
                vtrparametroA[i] := trocapv(grdImportproblema.Cells[i + 1, 4]);
              end;
            end;

          /// / Recebendo as consequências das alternativas ////
          for i := 0 to grdImportproblema.TotalRows - 9 do
          begin
            for j := 0 to grdImportproblema.Columns.Count - 2 do
            begin // atribui os valores do grid impotação ao MatrizInput
              consreal1[i, j] := trocapv(grdImportproblema.Cells[j + 1, i + 8]);
              if (consreal1[i, j] = 0) and (vtrtipo[j] = 4) then
              begin
                if TIWUserSession(WebApplication.Data).Idioma = 1 then
                  WebApplication.ShowMessage(
                    'Como você está considerando uma função valor logística, as consequências no critério ' + vtrnomecrit[j]
                      + ' devem ser diferentes de 0. Por favor, verifique a matriz de consequências antes de continuar.')
                else
                  WebApplication.ShowMessage(
                    'As you are considering a Logistic Value Fuction, the consequences in the criterion ' + vtrnomecrit[j]
                      + ' cannot be equal to 0. Please, verify the consequences matrix before continuing.');
                flImportarproblema.Visible := true;
                btnslotherfile.Enabled := true;
                Exit;
              end;
            end; // for j
          end; // for i

          /// Recebendo os valores parametro B dos criterios ñ lineares//
          for i := 0 to grdImportproblema.Columns.Count - 3 do
          begin
            if (grdImportproblema.Cells[i + 1, 5] = '') then
            begin
              if vtrtipo[i] <> 1 then
              begin
                { WebApplication.ShowMessage(
                  'Please, fill the non-linear value functions parameters A and B in the input spreadsheet. In case of doubt check the user guide.');
                  btnslotherfileClick(Self);
                  Exit; }
                TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
                // falta info, vai ter que elicitar
                INC(TIWUserSession(WebApplication.Data).Contavazio); // se  verificaintra=true, incrementa e usa esse contador para habilitar o btn da rgnMenuIntra

              end;
              if vtrtipo[i] = 1 then
                vtrparametroB[i] := 0;
            end
            else
              vtrparametroB[i] := trocapv(grdImportproblema.Cells[i + 1, 5]);
          end;

          with TIWUserSession(WebApplication.Data) do
            /// / Recebendo o número de níveis discretos dos critérios ////
            for i := 0 to grdImportproblema.Columns.Count - 3 do
            begin
              if (tcrit[i] = 2) or (tcrit[i] = 3) then
              begin
                niveisdisc[i] := StrToInt(grdImportproblema.Cells[i + 1, 6]);
                // * Indica ao sistema que existem critérios discretos a serem validados *//
                TIWUserSession(WebApplication.Data).ValDiscreto := 5;
              end
              else
                niveisdisc[i] := 0;
            end;

          with TIWUserSession(WebApplication.Data) do
          begin
            TIWUserSession(WebApplication.Data).NumCrit := length(tcrit);
            TIWUserSession(WebApplication.Data).NumAlt := length
              (vtrnomealternativa);

            Setlength(NomeCrit, NumCrit);
            Setlength(FTipo, NumCrit);
            Setlength(Typeofcrit, NumCrit);
            Setlength(Niveis, NumCrit);
            Setlength(parA, NumCrit);
            Setlength(parB, NumCrit);
            Setlength(NomeAlt, NumAlt);
            Setlength(ConsMatrix, NumAlt, NumCrit);
            Setlength(Cost, NumAlt);

            for i := 0 to NumCrit - 1 do
            begin
              TIWUserSession(WebApplication.Data).NomeCrit[i] := vtrnomecrit[i];
              TIWUserSession(WebApplication.Data).FTipo[i] := vtrtipo[i];
              TIWUserSession(WebApplication.Data).Typeofcrit[i] := tcrit[i];
              TIWUserSession(WebApplication.Data).Niveis[i] := niveisdisc[i];
              TIWUserSession(WebApplication.Data).parA[i] := vtrparametroA[i];
              TIWUserSession(WebApplication.Data).parB[i] := vtrparametroB[i];
            end;

            for j := 0 to NumAlt - 1 do
            Begin
              NomeAlt[j] := vtrnomealternativa[j];
              Cost[j] := Costproj[j];
              /// global recebe local portfolio
            end;

            for i := 0 to NumCrit - 1 do
              for j := 0 to NumAlt - 1 do
                ConsMatrix[j, i] := consreal1[j, i];
          end;

          for i := 0 to length(TIWUserSession(WebApplication.Data)
              .Verificaintra) - 1 do
            TIWUserSession(WebApplication.Data).DinamVerificaIntra[i] :=
              TIWUserSession(WebApplication.Data).Verificaintra[i];

          DeleteFile('C:\Users\Public\Documents' + edtEntername.Text + '.xls');

          // ** Permite a seleção de um novo problema ** //
          IWLabel13.Visible := true;
          IWLabel14.Visible := true;
          IWLabel15.Visible := true;
          // * Chama a validação dos discretos *//
          if TIWUserSession(WebApplication.Data).ValDiscreto = 5 then
            validardiscretos(TIWUserSession(WebApplication.Data).ValDiscreto);
          if TIWUserSession(WebApplication.Data).ValDiscreto <> -1 then
          begin
            TIWUserSession(WebApplication.Data).PortfolioG := true;
            lblsheet.Visible := true;
            WebApplication.ShowMessage(
              'Worksheet successfully loaded! Click on "Continue".');
            btnauxstartelicitation.Font.Style := [fsbold];
            btnauxstartelicitation.Enabled := true;
            btnImporttabela.Enabled := false;
            // ativa botão para inicio de elicitação

            // deleta o arquivo salvo na pasta
            // flImportarproblema.Destroy;

          end;
        end // se o arquivo for compatível
        else
        begin // se o arquivo não for compatível com a extensão .xls
          WebApplication.ShowMessage('check the type of the file');
        end;

      end;
    end;
  end; // UserSession
end; // procedimento

procedure TFrmEntrada.btnlcloseClick(Sender: TObject);
begin
  rgnlanguage.Visible := false;
end;

procedure TFrmEntrada.btnLoginClick(Sender: TObject);
begin
  rgnintro.Visible := false;
  rgnEnterLogin.Visible := true;

end;

procedure TFrmEntrada.btnLogoutRgnCadProjetoClick(Sender: TObject);
begin
  Logout; // Muda as regiões

  // Limpa os edts e lisbox
  edtProjectname.Clear;
  edtNumberalternat.Text := '0';
  edtNumberattribut.Text := '0';
  edtNamealternat.Clear;
  edtNameattribut.Clear;
  edtNumberlevels.Clear;
  edtparA.Clear;
  edtparB.Clear;
  edtNumberlevels.Clear;
  edtbudget.Clear;

  lstbxNamealternat.Items.Clear;
  lstbxNamecriteria.Items.Clear;

  cmbxFO.ItemIndex := -2;
  cmbxScaletype.ItemIndex := -1;
  cmbxCriteriondirection.ItemIndex := -1;

  // Nesse momento nenhuma variável foi salva ainda (não precisa zerar)
end;

procedure TFrmEntrada.btnLogoutRgnImportExportClick(Sender: TObject);
begin
  Logout;
  edtEntername.Text := '';
  btnslotherfileClick(Self); // Limpa as variáveis
  // cmbbxproblem.ItemIndex := -1; // "tira" a seleção
end;

procedure TFrmEntrada.btnLogoutRgnInputMatrizConseqClick(Sender: TObject);
begin
  Logout; // Muda as regiões
  btnslotherfileClick(Self); // Limpa as variáveis

  grdMatrizConseq.Clear;
end;

procedure TFrmEntrada.btnLogoutRgnLoginClick(Sender: TObject);
begin
  Logout; // Muda a região
end;

procedure TFrmEntrada.btnLogoutRgnOpcaoInputClick(Sender: TObject);
begin
  Logout;

end;

procedure TFrmEntrada.btnmessageDiscClick(Sender: TObject);
begin
  WebApplication.ShowMessage(txthelpdisc.Lines.Text);
end;

procedure TFrmEntrada.btnminaltClick(Sender: TObject);
var
  E, valor: integer;

begin
  E := StrToInt(edtNumberalternat.Text); // E capta o valor inteiro da string presente no edit
  valor := E - 1; // a variável é inicializada em zero e tem que incrementar a cada clique
  edtNumberalternat.Text := inttostr(valor);
  // mostra pro usuario ver a qntd de alt

  if (StrToInt(edtNumberalternat.Text) < 1) then
  begin
    WebApplication.ShowMessage('Please enter a valid number of alternatives');
    btnminalt.Enabled := false;
  end;
end;

procedure TFrmEntrada.btnmincritClick(Sender: TObject);
var
  E, valor: integer;

begin
  E := StrToInt(edtNumberattribut.Text); // E capta o valor inteiro da string presente no edit
  valor := E - 1; // a variável é inicializada em zero e tem que incrementar a cada clique
  edtNumberattribut.Text := inttostr(valor);
  // mostra pro usuario ver a qntd de crit

  if (StrToInt(edtNumberattribut.Text) < 1) then
  begin
    WebApplication.ShowMessage('Please enter a valid number of criteria');
    btnmincrit.Enabled := false;
  end;
end;

procedure TFrmEntrada.btnnewcodeClick(Sender: TObject);
var
  P, email, nome, titulo: string;
  vetchar: array [0 .. 9] of string;
  i, a: integer;
  H: boolean;
begin
  email := TIWUserSession(WebApplication.Data).email;
  vetchar[0] := '!';
  vetchar[1] := '@';
  vetchar[2] := '#';
  vetchar[3] := '$';
  vetchar[4] := '%';
  vetchar[5] := '&';
  vetchar[6] := '*';
  vetchar[7] := '^';
  vetchar[8] := '-';
  vetchar[9] := '+';

  for i := 0 to 5 - 1 do
  begin
    P := FloatToStr(RandomRange(0, 9));
    TIWUserSession(WebApplication.Data).auxvalidacao := TIWUserSession
      (WebApplication.Data).auxvalidacao + FloatToStr
      (Round(RandomRange(0, 1000) / 10)) + vetchar[StrToInt(P)];
  end;

  {with UserSession.ZQuery1 do
  begin
    SQL.Text := 'SELECT name FROM user WHERE email ="' + (edtemail.Text) + '"';
    ExecSQL;
    open;
    nome := FieldByName('name').AsString;
    close;

    SQL.Text := 'SELECT title FROM user WHERE email ="' + (edtemail.Text) + '"';
    ExecSQL;
    open;
    titulo := FieldByName('title').AsString;
    close;
  end; }

  if titulo <> '' then
    txtnewcode.Lines[0] := txtnewcode.Lines[0] + ' ' + titulo + ' ' + nome + ','
  else
    txtnewcode.Lines[0] := txtnewcode.Lines[0] + ' ' + nome + ',';
  txtnewcode.Lines[6] := txtnewcode.Lines[6] + ' ' + TIWUserSession
    (WebApplication.Data).auxvalidacao;

  H := send_email(email, 'Confirm your FITradeoff Web Based Account',
    txtnewcode.Lines.Text);
  if H then
  begin
    WebApplication.ShowMessage(
      ' It was sent a new validation code to your e-mail. Introduce it on your next access to FITradeoff system.'
        + ' Please, check also your spam in case the e-mail does not reach your inbox.');

   { with UserSession.ZQuery1 do
    begin
      // salvando no BD
      SQL.Text := 'UPDATE user SET validacao ="' +
        (TIWUserSession(WebApplication.Data).auxvalidacao)
        + '" WHERE email="' + (TIWUserSession(WebApplication.Data).email)
        + '";';
      ExecSQL;
    end;
          }
  end
  else
    WebApplication.ShowMessage(
      'It was not possible to conclude this action. Please, try again later. If the problem persist, contact us.');

  txtnewcode.Lines[0] := 'Dear';
  txtnewcode.Lines[6] := 'Activation Code =';
  rgnvalidationpassword.Visible := false;
  rgnvalidation.Visible := false;
  rgnEnterLogin.Visible := false;
  rgnintro.Visible := true;
  edtSenha.Clear;
  edtemail.Clear;

end;

procedure TFrmEntrada.btnplusaltClick(Sender: TObject);
var
  E, valor: integer;
begin
  E := StrToInt(edtNumberalternat.Text); // E capta o valor inteiro da string presente no edit
  valor := E + 1; // a variável é inicializada em zero e tem que incrementar a cada clique
  edtNumberalternat.Text := inttostr(valor);
  // mostra pro usuario ver a qntd de alt
  btnminalt.Enabled := true;
end;

procedure TFrmEntrada.btnpluscritClick(Sender: TObject);
var
  E, valor: integer;

begin
  E := StrToInt(edtNumberattribut.Text); // E capta o valor inteiro da string presente no edit
  valor := E + 1; // a variável é inicializada em zero e tem que incrementar a cada clique
  edtNumberattribut.Text := inttostr(valor);
  // mostra pro usuario ver a qntd de crit
  btnmincrit.Enabled := true;
end;

procedure TFrmEntrada.btnportugueseClick(Sender: TObject);
begin
  WebApplication.SendFile(
    'C:\inetpub\fitradeoff\Images\Guia rapido Fitradeoff web' + '.pdf', true);
  rgnlanguage.Visible := false;
end;

procedure TFrmEntrada.btnrecoverClick(Sender: TObject);
var
  P, email, nome, titulo: string;
  vetchar: array [0 .. 9] of string;
  i, a, quantregis: integer;
  H: boolean;
begin
  email := edtemail.Text;
  if email = '' then
  begin
    WebApplication.ShowMessage(
      'Please, fill in the field with the registered e-mail and try again.');
    Exit;
  end;
  {with UserSession.ZQuery1 do
  begin

    SQL.Text := 'SELECT count(*) FROM user WHERE email="' + (edtemail.Text)
      + '"';
    ExecSQL;
    open;
    quantregis := FieldByName('count(*)').AsInteger;
    close;
  end;  }
  if quantregis = 0 then
  begin
    IWLabel9.Visible := true;
    Exit;
  end
  else
    IWLabel9.Visible := false;

  vetchar[0] := '!';
  vetchar[1] := '@';
  vetchar[2] := '#';
  vetchar[3] := '$';
  vetchar[4] := '%';
  vetchar[5] := '&';
  vetchar[6] := '*';
  vetchar[7] := '^';
  vetchar[8] := '-';
  vetchar[9] := '+';

  {with UserSession.ZQuery1 do
  begin
    SQL.Text := 'SELECT name FROM user WHERE email ="' + (edtemail.Text) + '"';
    ExecSQL;
    open;
    nome := FieldByName('name').AsString;
    close;

    SQL.Text := 'SELECT title FROM user WHERE email ="' + (edtemail.Text) + '"';
    ExecSQL;
    open;
    titulo := FieldByName('title').AsString;
    close;
  end;   }

  for i := 0 to 5 - 1 do
  begin
    P := FloatToStr(RandomRange(0, 9));
    TIWUserSession(WebApplication.Data).auxvalidacao := TIWUserSession
      (WebApplication.Data).auxvalidacao + FloatToStr
      (Round(RandomRange(0, 1000) / 10)) + vetchar[StrToInt(P)];

  end;

  if titulo <> '' then
    txtpassrecover.Lines[0] := txtpassrecover.Lines[0] + ' ' + titulo + ' ' +
      nome + ','
  else
    txtpassrecover.Lines[0] := txtpassrecover.Lines[0] + ' ' + nome + ',';
  txtpassrecover.Lines[6] := txtpassrecover.Lines[6] + ' ' + TIWUserSession
    (WebApplication.Data).auxvalidacao;

  H := false;
  H := send_email(email, 'FItradeoff Web Based Temporary Password',
    txtpassrecover.Lines.Text);

  if H then
  begin
    WebApplication.ShowMessage(
      ' It was sent a temporary password to the e-mail informed. You shall change it on your next access to FITradeoff system.'
        + ' Please, check also your spam in case the e-mail does not reach your inbox.');
   { with UserSession.ZQuery1 do
    begin
      // salvando no BD, o campo validacao atual
      SQL.Text := 'UPDATE user SET validacao ="password" WHERE email="' +
        (email) + '";';
      ExecSQL;

      // salvando no BD,a senha temporária
      SQL.Text := 'UPDATE user SET password ="' +
        (TIWUserSession(WebApplication.Data).auxvalidacao)
        + '" WHERE email="' + (email) + '";';
      ExecSQL;
    end; }
  end
  else
    WebApplication.ShowMessage(
      'It was not possible to conclude this action. Please, try again later. If the problem persist, please contact us.');
  txtpassrecover.Lines[0] := 'Dear';
  txtpassrecover.Lines[6] := 'Temporary Password =';
end;

procedure TFrmEntrada.btnRegisterClick(Sender: TObject);
// botão migrar para região de cadastro pessoal
begin
  rgnintro.Visible := false;
  rgnCadusuario.Visible := true;

end;

procedure TFrmEntrada.btnRegisternewClick(Sender: TObject);
// procedimento para cadastrar novo problema
var
  i, j: integer;
begin

  // cmproblematica1.ItemIndex := -1;
  AlreadyRakend1.Visible := false;
  rgnOpcaoInput.Visible := false;
  rgnCadprojeto.Visible := true;

  if TIWUserSession(WebApplication.Data).Problematica = 2 then
  begin
    edtbudget.Visible := true;
    lblbudget.Visible := true;
    hptpbud.Visible := true;
    lblproblematic2.Caption := 'Portfolio CB';
    lblproblematic1.Caption := 'Portfolio CB';
  end; // referente a   if TIWUserSession(WebApplication.Data).problematica =2 then

  if TIWUserSession(WebApplication.Data).Problematica = 4 then
  // ADRIANA INICIO
  begin
    edtbudget.Visible := true;
    lblbudget.Visible := true;
    hptpbud.Visible := true;
    lblproblematic2.Caption := 'Combinatorial Portfolio';
    lblproblematic1.Caption := 'Combinatorial Portfolio';
  end; // referente a   if TIWUserSession(WebApplication.Data).problematica =2 then   //ADRIANA FIM

  if TIWUserSession(WebApplication.Data).Problematica = 0 then
  begin
    lblproblematic2.Caption := 'Choice';
    lblproblematic1.Caption := 'Choice';
  end;

  if TIWUserSession(WebApplication.Data).Problematica = 1 then
  begin
    lblproblematic2.Caption := 'Ranking';
    lblproblematic1.Caption := 'Ranking';
  end;

  if TIWUserSession(WebApplication.Data).Problematica = 3 then
  begin
    lblproblematic2.Caption := 'Sorting';
    lblproblematic1.Caption := 'Sorting';
  end;

  TIWUserSession(WebApplication.Data).AuxNumAlt := 0;
  TIWUserSession(WebApplication.Data).AuxNumCrit := 0;
  TIWUserSession(WebApplication.Data).EditaCrit := 0; // não editar
  TIWUserSession(WebApplication.Data).AuxCritApagados := 0;
  TIWUserSession(WebApplication.Data).AuxAltApagadas := 0;
  TIWUserSession(WebApplication.Data).ID_chave := idUnique;

  { --- Limpando todos os elementos da página de cadastro manual ---- }

  lstbxNamealternat.Clear; // Limpa conteúdo dos list box
  lstbxNamecriteria.Clear;

  edtProjectname.Clear;
  edtbudget.Clear;
  edtNumberalternat.Text := '0';
  edtNumberattribut.Text := '0';
  lstbxNamealternat.Clear;
  lstbxNamecriteria.Clear;
  // cmbbxproblem.ItemIndex := -1;

  edtNameattribut.Clear;
  edtparA.Clear;
  edtparB.Clear;
  edtNumberlevels.Clear;

  cmbxScaletype.ItemIndex := -1;
  cmbxFO.ItemIndex := -1;
  cmbxCriteriondirection.ItemIndex := -1;

  lblfunc.Visible := false;
  lblNumberlevels.Visible := false;
  edtNumberlevels.Visible := false;
  txtparametros.Visible := false;
  lblparA.Visible := false;
  lblparB.Visible := false;
  edtparA.Visible := false;
  edtparB.Visible := false;

  { --------- LIMPAR A MATRIZ DE EDITS ----------- }
  with TIWUserSession(WebApplication.Data) do
  begin
    for i := 0 to NumAlt - 1 do
    begin
      if length(ArrayAlt) > i then
        ArrayAlt[i].Destroy;
      for j := 0 to NumCrit - 1 do
      begin
        if length(MatEdit) > i then
          if length(MatEdit[i]) > j then
            MatEdit[i, j].Destroy;
        if i = 0 then
          if length(ArrayCrt) > j then
            ArrayCrt[j].Destroy;
      end;
    end;
    Setlength(MatEdit, 0, 0);
    Setlength(ArrayAlt, 0);
    Setlength(ArrayCrt, 0);

    EditaAlt := 0;

  end; // with

end;

procedure TFrmEntrada.btnreset2Click(Sender: TObject);
var
  i, j, cont_preenchido: integer;

begin
  with TIWUserSession(WebApplication.Data) do
  begin
    cont_preenchido := 0;

    if (Problematica <> 2) or (Problematica <> 4) then // ADRIANA
    begin
      for i := 0 to NumAlt - 1 do
      begin
        for j := 0 to NumCrit - 1 do
        begin
          if (MatEdit[i][j].Text <> '') then
            INC(cont_preenchido);
        end; // for j := 0 to Numcrit - 1
      end; // for i := 0 to Numalt - 1
    end // problematica<>2
    else
    begin
      for i := 0 to NumAlt - 1 do
      begin
        for j := 0 to NumCrit do
        begin
          if (MatEdit[i][j].Text <> '') then
            INC(cont_preenchido);
        end; // for j := 0 to Numcrit - 1
      end; // for i := 0 to Numalt - 1
    end; // referente ao else

    if (cont_preenchido > 0) AND (Auxresetedit = false) then
    begin
      rgnCofirmReset.Visible := true;
      Exit;
    end;

    rgnConseqvalues.Visible := false;
    rgnOpcaoInput.Visible := true;
    // rgnproblematica.Visible := true;
    lblbudget.Visible := false;
    edtbudget.Visible := false;
    hptpbud.Visible := false;

    lstbxNamealternat.Clear; // Limpa conteúdo dos list box
    lstbxNamecriteria.Clear;

    // limpando marcação do tipo de problematica
    rgnchoice.BorderOptions.NumericWidth := 0;
    rgnrank.BorderOptions.NumericWidth := 0;
    rgnsorting.BorderOptions.NumericWidth := 0;
    rgnport.BorderOptions.NumericWidth := 0;
    // problematica:= -1;

    edtProjectname.Clear;
    edtbudget.Clear;
    edtNumberalternat.Text := '0';
    edtNumberattribut.Text := '0';
    lstbxNamealternat.Clear;
    lstbxNamecriteria.Clear;

    // cmbbxproblem.ItemIndex := -1; //atribuido problematica=-1 no final do procedimento para zerar

    { --------- LIMPAR A MATRIZ DE EDITS ----------- }
    with TIWUserSession(WebApplication.Data) do
    begin
      for i := 0 to NumAlt - 1 do
      begin
        if length(ArrayAlt) > i then
          ArrayAlt[i].Destroy;

        If (Problematica = 2) or (Problematica = 4) then // ADRIANA
        begin
          for j := 0 to NumCrit do
          begin
            if length(MatEdit) > i then
              if length(MatEdit[i]) > j then
                MatEdit[i, j].Destroy;
            if i = 0 then
              if length(ArrayCrt) > j then
                TIWUserSession(WebApplication.Data).ArrayCrt[j].Destroy;
          end;
        end
        else
        Begin
          for j := 0 to NumCrit - 1 do
          begin
            if length(MatEdit) > i then
              if length(MatEdit[i]) > j then
                MatEdit[i, j].Destroy;
            if i = 0 then
              if length(ArrayCrt) > j then
                ArrayCrt[j].Destroy;
          end;
        End;
      end;

      Setlength(MatEdit, 0, 0);
      Setlength(ArrayAlt, 0);
      Setlength(ArrayCrt, 0);

      Auxresetedit := false;

    end; // with
  end; // user session
end;

procedure TFrmEntrada.btnReset3Click(Sender: TObject);
var
  i, j: integer;
begin
  rgnOpcaoInput.Visible := true;
  rgnLoginproject.Visible := false;
end;

procedure TFrmEntrada.btnReset4Click(Sender: TObject);
begin
  rgnImportExport.Visible := false;
  rgnOpcaoInput.Visible := true;

  edtEntername.Text := '';
  // cmbbxproblem.ItemIndex := -1; // "tira" a seleção
  btnslotherfileClick(Self); // Limpa as variáveis

end;

procedure TFrmEntrada.btnResetClick(Sender: TObject);
begin
  rgnCadprojeto.Visible := false;
  rgnOpcaoInput.Visible := true;

  // Limpa os edts e lisbox
  edtProjectname.Clear;
  edtNumberalternat.Text := '0';
  edtNumberattribut.Text := '0';
  edtNamealternat.Clear;
  edtNameattribut.Clear;
  edtNumberlevels.Clear;
  edtbudget.Clear;
  lstbxNamealternat.Items.Clear;
  lstbxNamecriteria.Items.Clear;

  edtparA.Clear;
  edtparB.Clear;
  edtNumberlevels.Clear;

  cmbxFO.ItemIndex := -2;
  cmbxScaletype.ItemIndex := -1;
  cmbxCriteriondirection.ItemIndex := -1;

end;

procedure TFrmEntrada.btnResetRgnInputMatrizConClick(Sender: TObject);
begin
  rgnOpcaoInput.Visible := true;
  rgnInputMatrizConseq.Visible := false;
  rgnLoginproject.Visible := false;
  rgnConseqvalues.Visible := false;
  rgnImportExport.Visible := false;
  rgnCadprojeto.Visible := false;
  // Exit;

  // limpando marcação do tipo de problematica
  rgnchoice.BorderOptions.NumericWidth := 0;
  rgnrank.BorderOptions.NumericWidth := 0;
  rgnsorting.BorderOptions.NumericWidth := 0;
  rgnport.BorderOptions.NumericWidth := 0;

  // referente a budget
  edtbudget.Visible := false;
  hptpbud.Visible := false;
  lblbudget.Visible := false;

  btnslotherfileClick(Self); // Limpa as variáveis

  grdMatrizConseq.Clear;
  // limpando a problematica
  // TIWUserSession(WebApplication.Data).Problematica:= -1;

end;

procedure TFrmEntrada.btnSavecadastroClick(Sender: TObject);
// procedimento para salvar infos cadastro de projeto
var
  i, j, a, b, c, d, verificacrit, k { NumAlt } : integer; // numcrit var que armazena qntd de criterios e numalt var que armazena qntd de alternativas
  tipoFO, tipoEscala, tipoDirecao: array of string;
  vtrnomecrit, vtrnomealternativa: array of string;
  // ArrayAlt, ArrayCrt: array of TIWLabel; //Coloquei como variável global
  // Nomeprojeto: string;
  AuxTipo: array [0 .. 3] of string;

  // apagar depois (variáveis locais para verificação)
  nome, nomealternativa: array of string;
  afunctipo, aescala, adirec, aniveis: array of integer;
  apara, aparb: array of Double;

begin
  with TIWUserSession(WebApplication.Data) do
  begin

    if edtProjectname.Text = '' then // verificação se nome do projeto está preenchido
    begin
      WebApplication.ShowMessage('Please,fill the name of problem.');
      Exit;
    end;

    { if cmproblematica1.ItemIndex = -1 then
      begin
      WebApplication.ShowMessage('Please, select the problematic.');
      Exit;
      end; }

    /// verificação que edt do budget está preenchido e exibição do lembrete na rgn da matriz editável///
    if (Problematica = 2) or (Problematica = 4) then
    begin
      lbllembretebudget.Visible := true;
      // lbllembretebudget.Text:= 'Budget value: $' + FormatFloat('#,##0.00',StrToFloat(edtbudget.text));

      { if (strtofloat(edtbudget.Text)<0) or (strtofloat(edtbudget.Text)=0) then
        begin
        Webapplication.showmessage('Please, fill the field budget with numbers bigger than 0.');
        Exit;
        end; //referente a }

      if edtbudget.Text = '' then // verificação se declarou valor budget
      begin
        WebApplication.ShowMessage('Please,fill the budget value.');
        Exit;
      end // referente a  if edtbudget.Text= '' then
      else
      Begin
        Try // verificação para só haverem numerais
          trocapv(edtbudget.Text);
        Except
          begin
            WebApplication.ShowMessage(
              'Please, fill the field budget with numbers only.');
            Exit;
          End; // show message
        End;
        budget := StrToFloat(edtbudget.Text);
      End; // referente ao else

      lbllembretebudget.Text := 'Budget value: $' + FormatFloat
        ('#,##0.00', StrToFloat(edtbudget.Text));

      if (StrToFloat(edtbudget.Text) < 0) or (StrToFloat(edtbudget.Text) = 0)
        then // verificação para budget somente >0
      begin
        WebApplication.ShowMessage(
          'Please, fill the field budget with numbers bigger than 0.');
        Exit;
      end; // referente a
    end; // problematica=2

    with TIWUserSession(WebApplication.Data) do
    begin

      { ---- Validando as alternativas ------ }
      { recebe quantidade de alternativas, valindando ser número natural e não nulo }
      { Try
        NumAlt := strtoint(edtNumberalternat.text);
        except

        begin
        WebApplication.ShowMessage(
        'Please, fill in the field "Number of Alternatives" only with numbers.');
        Exit;
        end;
        End;

        if (edtNumberalternat.text = '') then
        begin
        WebApplication.ShowMessage('Please, fill in the number of alternatives.');
        Exit;
        end; }
      // Primeira parte da validação também está no botão de adicionar alternativas para garantir que não ocorra nenhum problema lá
      // Repetida aqui caso o usuário altere o edit depois de adicionar alternativa

      { if (lstbxNamealternat.Items.Count < 2) or
        (strtoint(edtNumberalternat.text) < 2) then
        begin
        WebApplication.ShowMessage(
        'Please enter a valid value to the number of alternatives. It should be at least 2.');
        Exit;
        end;

        if (strtoint(edtNumberalternat.text) > (AuxNumAlt - AuxAltApagadas)) then
        begin
        WebApplication.ShowMessage(
        'The informed number of alternatives does not correspond to the number of alternatives added.');
        Exit;
        end; }

      { ---- Validando os critérios ------ }
      { Try
        verificacrit := strtoint(edtNumberattribut.text);
        except

        begin
        WebApplication.ShowMessage(
        'Please, fill in the field "Number of Criteria" only with numbers.');
        Exit;
        end;
        End;

        if (edtNumberattribut.text = '') then
        begin
        WebApplication.ShowMessage('Please, fill in the number of criteria.');
        Exit;
        end;
        // Validação inicial direto no botão de adicionar critérios para evitar erros naquela parte
        // Repetida aqui caso o usuário altere o edit depois de adicionar alternativa

        if (lstbxNamealternat.Items.Count < 2) or
        (strtoint(edtNumberattribut.text) < 2) then
        begin
        WebApplication.ShowMessage(
        'Please enter a valid value to the number of criteria. It should be at least 2.');
        Exit;
        end;

        if verificacrit > (AuxNumCrit - AuxCritApagados) then
        begin
        WebApplication.ShowMessage(
        'The defined number of criteria does not correspond to the number of criteria added.');
        Exit;
        end; }

      // numero de vezes que o usuário adcionou alternativas e critérios
      // independente do valor escrito no edit
      // TIWUserSession(WebApplication.Data).NumCrit := AuxNumCrit - AuxCritApagados;
      // TIWUserSession(WebApplication.Data).NumAlt := AuxNumAlt - AuxAltApagadas;

      TIWUserSession(WebApplication.Data).NumCrit := StrToInt
        (edtNumberattribut.Text);
      TIWUserSession(WebApplication.Data).NumAlt := StrToInt
        (edtNumberalternat.Text);

      Setlength(vtrnomealternativa, NumAlt);
      Setlength(vtrnomecrit, NumCrit);
      Setlength(MatEdit, NumAlt, NumCrit);

      for i := 0 to NumAlt - 1 do
        vtrnomealternativa[i] := '';

      for j := 0 to NumCrit - 1 do
        vtrnomecrit[j] := '';

      Setlength(ArrayAlt, NumAlt);
      Setlength(ArrayCrt, NumCrit);

      // begin // Criando os Edit's

      // Setando o tamanho dos vetores que irão armazenar as informações obtidas
      { for i := 0 to NumAlt - 1 do
        begin
        vtrnomealternativa[NumAlt] := '';
        end;

        for j := 0 to NumCrit - 1 do
        begin
        vtrnomecrit[NumCrit] := '';
        end;
        end; }
      k := 0;

      for i := 0 to NumAlt - 1 do
      begin
        ArrayAlt[i] := TIWLabel.Create(Self);
        ArrayAlt[i].parent := rgnconsman;
        ArrayAlt[i].Text := ('Alt' + inttostr(i + 1));
        ArrayAlt[i].Left := (5);
        ArrayAlt[i].Top := (21 + (i * 25) + 23);
        ArrayAlt[i].Font.Style := [fsbold];
        ArrayAlt[i].Font.FontName := 'Verdana';
      end;

      if (Problematica = 2) or (TIWUserSession(WebApplication.Data)
          .Problematica = 4) then
      begin
        Setlength(ArrayCrt, NumCrit + 1);
        Setlength(MatEdit, NumAlt, NumCrit + 1);

        for i := 0 to NumCrit do

        Begin
          if i < NumCrit then
          begin
            ArrayCrt[i] := TIWLabel.Create(Self);
            ArrayCrt[i].parent := rgnconsman;
            ArrayCrt[i].Text := ('C' + inttostr(i + 1));
            ArrayCrt[i].Left := (45 + (i * 60));
            ArrayCrt[i].Top := (20);
            TIWUserSession(WebApplication.Data).ArrayCrt[i].Font.Style :=
              [fsbold];
            ArrayCrt[i].Font.FontName := 'Verdana';
          end
          else
          begin
            ArrayCrt[i] := TIWLabel.Create(Self);
            ArrayCrt[i].parent := rgnconsman;
            TIWUserSession(WebApplication.Data).ArrayCrt[i].Text :=
              ('Cost ($)');
            ArrayCrt[i].Left := (45 + (i * 60));
            ArrayCrt[i].Top := (20);
            ArrayCrt[i].Font.Style := [fsbold];
            ArrayCrt[i].Font.FontName := 'Verdana';
          end;
        end;
      end
      else
        for i := 0 to NumCrit - 1 do
        begin
          ArrayCrt[i] := TIWLabel.Create(Self);
          ArrayCrt[i].parent := rgnconsman;
          ArrayCrt[i].Text := ('C' + inttostr(i + 1));
          ArrayCrt[i].Left := (45 + (i * 60));
          ArrayCrt[i].Top := (20);
          ArrayCrt[i].Font.Style := [fsbold];
          ArrayCrt[i].Font.FontName := 'Verdana';
        end;

      for i := 0 to NumAlt - 1 do
      begin

        if (TIWUserSession(WebApplication.Data).Problematica = 2) or
          (TIWUserSession(WebApplication.Data).Problematica = 4) then
        begin
          for j := 0 to NumCrit do
          begin
            // Criando os edits que recaberão consequencias e valor dos custos na Matriz de TIWEdit's
            MatEdit[i][j] := TIWEdit.Create(Self);
            MatEdit[i][j].parent := rgnconsman;
            MatEdit[i][j].Text := ('');
            MatEdit[i][j].Left := (45 + (j * 60));
            MatEdit[i][j].Top := (25 + (i * 25) + 23);
            MatEdit[i][j].Width := 60;

            if j = NumCrit then // quando estiver na ultima coluna, coloca cifrão e destaca coluna cost
            begin
              MatEdit[i][j] := TIWEdit.Create(Self);
              MatEdit[i][j].BGColor := clMenuBar;
              MatEdit[i][j].parent := rgnconsman;
              MatEdit[i][j].Text := '$';
              MatEdit[i][j].Left := (45 + (j * 60));
              MatEdit[i][j].Top := (25 + (i * 25) + 23);
              MatEdit[i][j].Width := 60;
            end; // if j=numcrit
          end; // for j
        end
        else
          for j := 0 to NumCrit - 1 do
          begin
            // Criando os edits que recaberão os desempenhos a partir de uma Matriz de TIWEdit's
            MatEdit[i][j] := TIWEdit.Create(Self);
            MatEdit[i][j].parent := rgnconsman;
            MatEdit[i][j].Text := ('');
            MatEdit[i][j].Left := (45 + (j * 60));
            MatEdit[i][j].Top := (25 + (i * 25) + 23);
            MatEdit[i][j].Width := 60;
          end; // if prob=2

      end; // for i

      // ** Removi daqui pois agora o usuário pode precisar voltar para esta pagina **//
      { lstbxNamealternat.Clear; // Limpa conteúdo dos list box
        lstbxNamecriteria.Clear; }

      // salvando o nomedoprojeto
      Nomeprojeto := edtProjectname.Text;
      // TIWUserSession(WebApplication.Data).Nomeprojeto := Nomeprojeto;

      // salvando os dados nos vetores originais

      Setlength(NomeCrit, NumCrit); // um valor qualquer muito grande
      Setlength(FTipo, NumCrit);
      Setlength(Typeofcrit, NumCrit);
      Setlength(Niveis, NumCrit);
      Setlength(parA, NumCrit);
      Setlength(parB, NumCrit);
      Setlength(NomeAlt, NumAlt);
      Setlength(Verificaintra, NumCrit);
      // TIWUserSession(WebApplication.Data).Problematica := rgnport.BorderOptions.NumericWidth;
      b := 0;
      d := 0;

      for a := 0 to 99 do
      begin
        if TIWUserSession(WebApplication.Data).AuxNomeCrit[a] <> '' then
        begin
          // for b := 0 to NumCrit - 1 do
          // begin
          NomeCrit[b] := AuxNomeCrit[a];
          FTipo[b] := AuxFuncTipo[a];
          Niveis[b] := AuxNiveis[a];
          parA[b] := AuxParA[a];
          parB[b] := AuxParB[a];

          if (AuxEscalaTipo[a] = 0) AND (AuxDirecTipo[a] = 0) then
            Typeofcrit[b] := 1; // continuo maximização
          if (AuxEscalaTipo[a] = 0) AND (AuxDirecTipo[a] = 1) then
            Typeofcrit[b] := 0; // continuo minimização
          if (AuxEscalaTipo[a] = 1) AND (AuxDirecTipo[a] = 0) then
          begin
            Typeofcrit[b] := 3; // discreto maximização
            // ** Indica que deve-se fazer a validação dos discretos **//
            TIWUserSession(WebApplication.Data).ValDiscreto := 5;
          end;
          if (AuxEscalaTipo[a] = 1) AND (AuxDirecTipo[a] = 1) then
          begin
            Typeofcrit[b] := 2; // discreto minimização
            // ** Indica que deve-se fazer a validação dos discretos **//
            TIWUserSession(WebApplication.Data).ValDiscreto := 5;
          end;
          // break;

          b := b + 1;
        end;
      end;

      for c := 0 to 99 do
      begin
        if AuxNomeAlt[c] <> '' then
        begin
          NomeAlt[d] := AuxNomeAlt[c];
          d := d + 1;
        end;
      end;

      // neste trecho inicia troca de região para criar matriz de conseq editavel
      rgnCadprojeto.Visible := false;
      rgnConseqvalues.Visible := true;

      /// //////APAGAR (verificação)/////////////////
      /// ///////////////////////////////
      Setlength(nome, 100);
      Setlength(afunctipo, 100);
      Setlength(aescala, 100);
      Setlength(adirec, 100);
      Setlength(aniveis, 100);
      Setlength(apara, 100);
      Setlength(aparb, 100);
      Setlength(nomealternativa, 100);

      for j := 0 to NumCrit - 1 do
      begin
        nome[j] := NomeCrit[j];
        afunctipo[j] := FTipo[j];
        aniveis[j] := Niveis[j];
        aescala[j] := Typeofcrit[j];
        apara[j] := parA[j];
        aparb[j] := parB[j];
      end;

      for j := 0 to NumAlt - 1 do
        nomealternativa[j] := NomeAlt[j];
      /// ////////////////////////////////////////////////////////////////////////
      /// ///////////////////////////////////////

      // Limpar os edits
      // edtProjectname.Clear;    // Não limpo o nome caso o usuário volte para essa página
      edtNumberalternat.Text := '0';
      edtNumberattribut.Text := '0';
      // edtNumberattribut.Clear;
      edtNameattribut.Clear;
      edtparA.Clear;
      edtparB.Clear;
      edtNumberlevels.Clear;
      edtbudget.Clear;
      cmbxFO.ItemIndex := -2; // -2, porque agora o -1 está sendo usado para pegar caso de não querer declarar
      cmbxScaletype.ItemIndex := -1;
      cmbxCriteriondirection.ItemIndex := -1;
      // cmproblematica1.ItemIndex := -1;  // problematica não vem mais do combobox

      AuxTipo[0] := 'Linear';
      AuxTipo[1] := 'Exponential';
      AuxTipo[2] := 'Logarithmic';
      AuxTipo[3] := 'Logistic';
      // ** Preenchendo a tabela de informações dos critérios **//
      grdCritInfo.ColumnCount := 7;
      grdCritInfo.RowCount := NumCrit + 1;

      grdCritInfo.Cell[0, 0].Text := 'Name';
      grdCritInfo.Cell[0, 1].Text := 'Scale';
      grdCritInfo.Cell[0, 2].Text := 'Number of levels';
      grdCritInfo.Cell[0, 3].Text := 'Value Function';
      grdCritInfo.Cell[0, 4].Text := 'Direction';
      grdCritInfo.Cell[0, 5].Text := 'Par a';
      grdCritInfo.Cell[0, 6].Text := 'Par b';

      for i := 0 to NumCrit - 1 do
      begin
        grdCritInfo.Cell[i + 1, 0].Text := NomeCrit[i];
        if Typeofcrit[i] > 1 then
        begin
          grdCritInfo.Cell[i + 1, 1].Text := 'Discrete';
          grdCritInfo.Cell[i + 1, 2].Text := inttostr(Niveis[i]);
        end
        else
        begin
          grdCritInfo.Cell[i + 1, 1].Text := 'Continuous';
          grdCritInfo.Cell[i + 1, 2].Text := '---';
        end; // end do else

        // quando Ftipo=0 (elic intra), escreve direto na célula do grid. Sem utilizar AuxTipo
        if FTipo[i] = 0 then
        begin
          grdCritInfo.Cell[i + 1, 3].Text := '---';
          TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
        end
        else
          grdCritInfo.Cell[i + 1, 3].Text := AuxTipo[FTipo[i] - 1];

        /// //COLOCANDO TYPECRIT SEPARADO, DEVIDO QUESTÃO DE EDTS VAZIOSS//////
        if (Typeofcrit[i] = 1) then // CONT MAX
        begin
          grdCritInfo.Cell[i + 1, 4].Text := 'Maximization';

          if (edtparA.Text = '') then
          begin
            grdCritInfo.Cell[i + 1, 5].Text := '---';
            if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
              TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
          end;

          if (edtparB.Text = '') then
          begin
            grdCritInfo.Cell[i + 1, 6].Text := '---';
            if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
              TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
          end // se vazio

          else
          begin
            grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
            grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
            if (parA[i] = 0) then
              grdCritInfo.Cell[i + 1, 5].Text := '---';
            if (parB[i] = 0) then
              grdCritInfo.Cell[i + 1, 6].Text := '---';
          end; // se possui valor
        end; // type=1

        if (Typeofcrit[i] = 3) then // DISC MAX
        begin
          grdCritInfo.Cell[i + 1, 4].Text := 'Maximization';
          grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
          grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
          if (parA[i] = 0) then
            grdCritInfo.Cell[i + 1, 5].Text := '---';
          if (parB[i] = 0) then
            grdCritInfo.Cell[i + 1, 6].Text := '---';
        end
        else

          if (Typeofcrit[i] = 0) then // CONT MIN
        begin
          grdCritInfo.Cell[i + 1, 4].Text := 'Minimization';

          if (edtparA.Text = '') then
          begin
            grdCritInfo.Cell[i + 1, 5].Text := '---';
            if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
              TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
          end;

          if (edtparB.Text = '') then
          begin
            grdCritInfo.Cell[i + 1, 6].Text := '---';
            if FTipo[i] <> 1 then // se vazio, e não linear, elicitação intra é obrigatoria
              TIWUserSession(WebApplication.Data).Verificaintra[i] := true;
          end // se vazio

          else
          begin
            grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
            grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
            if (parA[i] = 0) then
              grdCritInfo.Cell[i + 1, 5].Text := '---';
            if (parB[i] = 0) then
              grdCritInfo.Cell[i + 1, 6].Text := '---';
          end; // se possui valor
        end; // type= 0

        if (Typeofcrit[i] = 2) then // DISC MIN
        begin
          grdCritInfo.Cell[i + 1, 4].Text := 'Minimization';
          grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
          grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
          if (parA[i] = 0) then
            grdCritInfo.Cell[i + 1, 5].Text := '---';
          if (parB[i] = 0) then
            grdCritInfo.Cell[i + 1, 6].Text := '---';
        end;
      end;
    end;

    { if (Typeofcrit[i] = 1) OR (Typeofcrit[i] = 3) then
      begin
      grdCritInfo.Cell[i + 1, 4].Text := 'Maximization';
      grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
      grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
      if (parA[i] = 0) then
      grdCritInfo.Cell[i + 1, 5].Text := '---';
      if (parB[i] = 0) then
      grdCritInfo.Cell[i + 1, 6].Text := '---';
      end
      else if (Typeofcrit[i] = 0) OR (Typeofcrit[i] = 2) then
      begin
      grdCritInfo.Cell[i + 1, 4].Text := 'Minimization';
      grdCritInfo.Cell[i + 1, 5].Text := FloatToStr(parA[i]);
      grdCritInfo.Cell[i + 1, 6].Text := FloatToStr(parB[i]);
      if (parA[i] = 0) then
      grdCritInfo.Cell[i + 1, 5].Text := '---';
      if (parB[i] = 0) then
      grdCritInfo.Cell[i + 1, 6].Text := '---';
      end;
      end;
      end; }
    /// ///////////////////////////////////////////////////////////////////////////
    // grdCritInfo.Cell[0, 0].Width := '10';

    // criando o grid com as alternativas
    grdinformaalt.ColumnCount := 2;
    grdinformaalt.RowCount := TIWUserSession(WebApplication.Data).NumAlt;
    for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
    begin
      grdinformaalt.Cell[i, 0].Text := 'Alt' + ' ' + inttostr(i + 1);
      grdinformaalt.Cell[i, 1].Text := TIWUserSession(WebApplication.Data)
        .NomeAlt[i];
    end; // referente a for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do

    if TIWUserSession(WebApplication.Data).editar = -33333 then
    begin
      // apagando registros já existentes
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 27;
      SalvarBD;
    end;

    // salvar banco de dados na tabela consequence
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 1;
    SalvarBD;

    // salvar banco de dados na tabela consequence
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 35;
    SalvarBD;

    // salvar banco de dados na tabela consequence
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 10;
    SalvarBD;

    // salvar banco de dados na tabela consequence
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 23;
    SalvarBD;

    // salvar banco de dados na tabela consequence
    TIWUserSession(WebApplication.Data).estagio := -1;
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 5;
    SalvarBD;

  end;
end;

procedure TFrmEntrada.btnSaveproblemClick(Sender: TObject); // usuario salva valores do problema e retorna para tela de opções de input
var
  i, j, k, P, VerificaAlt: integer;
  MatConseq: array of array of Double;
  aux_custo, aux_custo2: string;
begin

  // Mensagem caso não esteja preenchido
  { for I := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
    begin
    for J := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
    begin
    if TIWUserSession(WebApplication.Data).MatEdit[J][I].text = '' then
    begin
    WebApplication.ShowMessage('Please,fill in all fields.');
    Exit;
    end;

    end; // j
    end; // i }

  { for i := 0 to TIWUserSession(WebApplication.Data).numcrit  - 1 do
    begin
    for j := 0 to TIWUserSession(WebApplication.Data).NumAlt  - 1 do
    begin
    Try
    trocapv(TIWUserSession(WebApplication.Data).MatEdit[j][i].Text);
    Except
    begin
    Webapplication.showmessage('Please, fill the fields with numbers only.');
    Exit;
    End; // referente ao except
    End;
    end; //referente a for j:=0...
    end; // referente a for i:=0... }

  { if TIWUserSession(WebApplication.Data).problematica=2  then
    begin
    for j := 0 to TIWUserSession(WebApplication.Data).NumAlt  - 1 do
    begin
    Try
    trocapv(TIWUserSession(WebApplication.Data).MatEdit[j][TIWUserSession(WebApplication.Data).numcrit].Text);
    Except
    begin
    Webapplication.showmessage('Please, fill the fields with numbers only.');
    Exit;
    End; // referente ao except
    end;
    end; //referente a for j:=0...
    end; //referente a problematica=2 }

  with TIWUserSession(WebApplication.Data) do
  begin
    Setlength(MatConseq, NumAlt, NumCrit);
    Setlength(ConsMatrix, NumAlt, NumCrit);
    Setlength(Cost, NumAlt);
    Setlength(ConsNorm, NumAlt, NumCrit);

    for i := 0 to NumCrit - 1 do
    begin
      for j := 0 to NumAlt - 1 do
      Begin
        if MatEdit[j][i].Text <> '' then
        begin
          MatConseq[j, i] := trocapv(MatEdit[j][i].Text);
          // variável local para teste
          ConsMatrix[j, i] := MatConseq[j, i];
        end; // fo if
      end; // i
    End; // j

    if (TIWUserSession(WebApplication.Data).Problematica = 2) or
      (TIWUserSession(WebApplication.Data).Problematica = 4) then
    begin
      for j := 0 to NumAlt - 1 do
      begin
        if (MatEdit[j][TIWUserSession(WebApplication.Data).NumCrit].Text <> '')
          AND (MatEdit[j][TIWUserSession(WebApplication.Data).NumCrit]
            .Text <> '$') then
        begin
          aux_custo :=
            ((MatEdit[j][TIWUserSession(WebApplication.Data).NumCrit].Text));
          if aux_custo[1] = '$' then
          begin

            for P := 2 to length(aux_custo) do
            begin
              aux_custo2 := aux_custo2 + aux_custo[P];
              // TIWUserSession(WebApplication.Data).Cost[j]:= strtofloat(TrocaVP(strtofloat(aux_custo[p])));
            end; // referente a for p := 2 to Length(aux_custo) - 1 do
          end // referente a if aux_...
          else
          begin
            for P := 1 to length(aux_custo) do
            begin
              aux_custo2 := aux_custo2 + aux_custo[P];
              // TIWUserSession(WebApplication.Data).Cost[j]:= strtofloat(TrocaVP(strtofloat(aux_custo[p])));
            end; // referente a for p := 2 to Length(aux_custo) - 1 do
          end; // referente ao else
          TIWUserSession(WebApplication.Data).Cost[j] := StrToFloat
            (TrocaVP(StrToFloat(aux_custo2)));
          aux_custo2 := '';
        end; // referente a if MatEdit[j][TIWUserSession(WebApplication.Data).numcrit].Text<>''  then
      end; // referente a for j := 0 to NumAlt - 1 do
    end; // referente a  if problematica =2 then
  end; // do with

  { /// ////// VERIFICAÇÃO ALTERNATIVAS IGUAIS ////////////////////
    for I := 0 to NumAlt - 1 do // quantidade de linhas
    begin
    for K := 0 to NumAlt - 1 do
    begin
    VerificaAlt := 0;
    for J := 0 to NumCrit - 1 do // quantidade de colunas
    begin
    if (K > I) AND (MatConseq[I, J] = MatConseq[K, J]) then
    begin
    VerificaAlt := VerificaAlt + 1;
    end;
    end; // for j
    // Depois de percorrer todas as colunas
    if VerificaAlt = NumCrit then
    begin
    WebApplication.ShowMessage(
    'All consequences are the same in alternative ' + NomeAlt[I]
    + ' and ' + NomeAlt[K] +
    '. Please, verify the input data, make sure these alternatives are not the same, and then try again.');
    Exit;
    end; // if
    end; // for k
    end;
    /// ///////////////////////////////////////////////////////

    end; // with

    if TIWUserSession(WebApplication.Data).ValDiscreto = 5 then
    validardiscretos(TIWUserSession(WebApplication.Data).ValDiscreto);
    if TIWUserSession(WebApplication.Data).ValDiscreto = -1 then
    begin
    TIWUserSession(WebApplication.Data).ValDiscreto := 5;
    Exit;
    end;

    Inicializa; }

  // se flag = -2: ocorreu um erro no preenchimento da matriz identificado no procedimento Inicializa
  // sai do procedimento sem adicionar no BD já que aconteceu um erro
  if TIWUserSession(WebApplication.Data).Flag <> -2 then
  begin

    // apagar registros do banco de dados na tabela consequence
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 30;
    SalvarBD;

    // salvar banco de dados na tabela consequence
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 6;
    SalvarBD;

    { // salvar banco de dados na tabela consequence
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 24;
      SalvarBD; }

    // salvar banco de dados na tabela consequence
    TIWUserSession(WebApplication.Data).estagio := 6;
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 5;
    SalvarBD;

    if (TIWUserSession(WebApplication.Data).Problematica = 2) or
      (TIWUserSession(WebApplication.Data).Problematica = 4) then
    begin
      // se a problematica for de portfolio deve também chamar o caso que salva os valores dos custos
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 36;
      SalvarBD;
    end; // referente a  if TIWUserSession(WebApplication.Data).Problematica=2 then

    { rgnConseqvalues.Visible := false;
      rgnOpcaoInput.Visible := true; }

    /// / Após salvar o problema os listbox da pagina de cadastro manual já podem
    /// ser limpados///////////////////////////////////////////////////////////
    { lstbxNamealternat.Clear; // Limpa conteúdo dos list box
      lstbxNamecriteria.Clear;

      edtProjectname.Clear;
      edtNumberalternat.text := '0';
      edtNumberattribut.text := '0';
      lstbxNamealternat.Clear;
      lstbxNamecriteria.Clear;
      cmbbxproblem.ItemIndex := -1;

      edtNameattribut.Clear;
      edtparA.Clear;
      edtparB.Clear;
      edtNumberlevels.Clear;

      cmbxScaletype.ItemIndex := -1;
      cmbxFO.ItemIndex := -1;
      cmbxCriteriondirection.ItemIndex := -1;

      lblfunc.Visible := false;
      lblNumberlevels.Visible := false;
      edtNumberlevels.Visible := false;
      txtparametros.Visible := false;
      lblparA.Visible := false;
      lblparB.Visible := false;
      edtparA.Visible := false;
      edtparB.Visible := false; }

    { --------- LIMPAR A MATRIZ DE EDITS ----------- }
    { with TIWUserSession(WebApplication.Data) do
      begin
      for I := 0 to NumAlt - 1 do
      begin
      if Length(ArrayAlt) > I then
      ArrayAlt[I].Destroy;
      for J := 0 to NumCrit - 1 do
      begin
      if Length(MatEdit) > I then
      if Length(MatEdit[I]) > J then
      MatEdit[I, J].Destroy;
      if I = 0 then
      if Length(ArrayCrt) > J then
      ArrayCrt[J].Destroy;
      end;
      end;
      Setlength(MatEdit, 0, 0);
      Setlength(ArrayAlt, 0);
      Setlength(ArrayCrt, 0);
      end; // with }

  end; // if flag <> -2 (erro na inicialização)

  { rgnConseqvalues.Visible := false;
    rgnOpcaoInput.Visible := true; }

  TIWUserSession(WebApplication.Data).Auxresetedit := true;

end; // procedure

procedure TFrmEntrada.btnSaveStartClick(Sender: TObject); // usuário salva valores do problema e inicia procedimento de elicitação
var
  i, j, k, P, VerificaAlt: integer;
  MatConseq: array of array of Double;
  Numcolimport: integer;
  aux_custo, aux_custo2, aux_Str, aux_Str2: string;

begin

  for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
  begin
    for j := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
    begin
      Try
        trocapv(TIWUserSession(WebApplication.Data).MatEdit[j][i].Text);
      Except
        begin
          WebApplication.ShowMessage(
            'Please, fill the fields with numbers only.');
          Exit;
        End; // referente ao except
      End;
    end; // referente a for j:=0...
  end; // referente a for i:=0...

  if (TIWUserSession(WebApplication.Data).Problematica = 2) or
    (TIWUserSession(WebApplication.Data).Problematica = 4) then
  begin
    for j := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
    begin
      aux_Str := ((TIWUserSession(WebApplication.Data).MatEdit[j][TIWUserSession
            (WebApplication.Data).NumCrit].Text));
      if aux_Str[1] = '$' then
      begin

        for P := 2 to length(aux_Str) do
        begin
          aux_Str2 := aux_Str2 + aux_Str[P];
          // TIWUserSession(WebApplication.Data).Cost[j]:= strtofloat(TrocaVP(strtofloat(aux_custo[p])));
        end; // referente a for p := 2 to Length(aux_custo) - 1 do
      end // referente a if aux_...
      else
      begin
        for P := 1 to length(aux_Str) do
        begin
          aux_Str2 := aux_Str2 + aux_Str[P];
          // TIWUserSession(WebApplication.Data).Cost[j]:= strtofloat(TrocaVP(strtofloat(aux_custo[p])));
        end; // referente a for p := 2 to Length(aux_custo) - 1 do
      end; // referente ao else

      Try
        trocapv(aux_Str2);
      Except
        begin
          WebApplication.ShowMessage(
            'Please, fill the fields with numbers only.');
          Exit;
        End; // referente ao except
      end;

      aux_Str2 := ''; // limpa auxiliar da verificação $

    end; // referente a for j:=0...
  end; // referente a problematica=2

  if (TIWUserSession(WebApplication.Data).Problematica = 2) or
    (TIWUserSession(WebApplication.Data).Problematica = 4) then
  begin

    // Mensagem caso não esteja preenchido
    for i := 0 to TIWUserSession(WebApplication.Data).NumCrit do
    begin
      for j := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
      begin
        if TIWUserSession(WebApplication.Data).MatEdit[j][i].Text = '' then
        begin
          WebApplication.ShowMessage('Please,fill in all fields.');
          Exit;
        end;
      end; // j
    end; // i
  end
  Else // se problematica 0 ou 1
  begin
    // Mensagem caso não esteja preenchido
    for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
    begin
      for j := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
      begin
        if TIWUserSession(WebApplication.Data).MatEdit[j][i].Text = '' then
        begin
          WebApplication.ShowMessage('Please,fill in all fields.');
          Exit;
        end;
      end; // j
    end; // i
  end;

  with TIWUserSession(WebApplication.Data) do
  begin
    Setlength(MatConseq, NumAlt, NumCrit);
    Setlength(ConsMatrix, NumAlt, NumCrit);
    Setlength(Cost, NumAlt);

    for i := 0 to NumCrit - 1 do
    begin
      for j := 0 to NumAlt - 1 do
      begin
        try
          MatConseq[j, i] := trocapv(MatEdit[j][i].Text);
        except

          begin
            WebApplication.ShowMessage(
              'Only numbers are accepted as consequences.');
            Exit;
          end;
        End; // try
        // variável local para teste
        ConsMatrix[j, i] := MatConseq[j, i];
      end; // j

      if (TIWUserSession(WebApplication.Data).Problematica = 2) or
        (TIWUserSession(WebApplication.Data).Problematica = 4) then
      begin
        for j := 0 to NumAlt - 1 do
        begin
          if MatEdit[j][TIWUserSession(WebApplication.Data).NumCrit]
            .Text <> '' then
          begin
            aux_custo :=
              ((MatEdit[j][TIWUserSession(WebApplication.Data).NumCrit].Text));
            if aux_custo[1] = '$' then
            begin

              for P := 2 to length(aux_custo) do
              begin
                aux_custo2 := aux_custo2 + aux_custo[P];
                // TIWUserSession(WebApplication.Data).Cost[j]:= strtofloat(TrocaVP(strtofloat(aux_custo[p])));
              end; // referente a for p := 2 to Length(aux_custo) - 1 do
            end // referente a if aux_...
            else
            begin
              for P := 1 to length(aux_custo) do
              begin
                aux_custo2 := aux_custo2 + aux_custo[P];
                // TIWUserSession(WebApplication.Data).Cost[j]:= strtofloat(TrocaVP(strtofloat(aux_custo[p])));
              end; // referente a for p := 2 to Length(aux_custo) - 1 do
            end; // referente ao else
            TIWUserSession(WebApplication.Data).Cost[j] := StrToFloat
              (TrocaVP(StrToFloat(aux_custo2)));
            aux_custo2 := '';
            // TIWUserSession(WebApplication.Data).Cost[j]:= trocapv(MatEdit[j][TIWUserSession(WebApplication.Data).numcrit].Text);
          end; // referente a if MatEdit[j][TIWUserSession(WebApplication.Data).numcrit].Text<>''  then
        end; // referente a for j := 0 to NumAlt - 1 do
      end; // referente a  if problematica =2 then

    end; // i

    /// ////// VERIFICAÇÃO ALTERNATIVAS IGUAIS ////////////////////
    for i := 0 to NumAlt - 1 do // quantidade de linhas
    begin
      for k := 0 to NumAlt - 1 do
      begin
        VerificaAlt := 0;
        for j := 0 to NumCrit - 1 do // quantidade de colunas
        begin
          if (k > i) AND (MatConseq[i, j] = MatConseq[k, j]) then
          begin
            VerificaAlt := VerificaAlt + 1;
          end;
        end; // for j
        // Depois de percorrer todas as colunas
        if VerificaAlt = NumCrit then
        begin
          WebApplication.ShowMessage(
            'There are at least two alternatives with the same consequences. Please, verify the altenartives consequences before continuing.');
          Exit;
        end; // if
      end; // for k
    end;
    /// ///////////////////////////////////////////////////////

  end; // with

  if TIWUserSession(WebApplication.Data).ValDiscreto = 5 then
    validardiscretos(TIWUserSession(WebApplication.Data).ValDiscreto);
  if TIWUserSession(WebApplication.Data).ValDiscreto = -1 then
  begin
    TIWUserSession(WebApplication.Data).ValDiscreto := 5;
    Exit;
  end;

  with TIWUserSession(WebApplication.Data) do
  begin
    /// / Setando o tamanho dos grids que exibem a informação importada na região imput ////
    grdMatrizConseq.RowCount := NumAlt + 7;
    grdMatrizConseq.ColumnCount := NumCrit + 1;
    probname := edtEntername.Text;

    /// / Preenchendo a informação fixa dos grids ////
    grdMatrizConseq.Cell[0, 0].Text := 'Criteria:';
    grdMatrizConseq.Cell[1, 0].Text :=
      '0-Cont Min; 1-Cont Max; 2-Disc Min; 3- Disc Max; 4- Int Min; 5- Int Max:';
    { grdInputvalues.Cell[2, 0].text := 'Weights:'; }
    grdMatrizConseq.Cell[2, 0].Text := '';
    grdMatrizConseq.Cell[3, 0].Text := '';
    grdMatrizConseq.Cell[4, 0].Text := '';
    grdMatrizConseq.Cell[5, 0].Text := 'Number of levels of discrete criteria';
    grdMatrizConseq.Cell[6, 0].Text := 'Consequence Matrix:';

    /// / Preenchendo os Grids com a informação Imputada ///////////////////////
    for j := 0 to NumCrit - 1 do
    begin
      grdMatrizConseq.Cell[0, j + 1].Text := NomeCrit[j];
      if Cri_inteiro[j] then
      begin
        if Typeofcrit[j] = 0 then
          grdMatrizConseq.Cell[1, j + 1].Text := '4';
        if Typeofcrit[j] = 1 then
          grdMatrizConseq.Cell[1, j + 1].Text := '5';
      end
      else
        grdMatrizConseq.Cell[1, j + 1].Text := inttostr(Typeofcrit[j]);
      { grdMatrizConseq.Cell[2, j + 1].Text := inttostr(FTipo[j]);
        if FTipo[j] <> 1 then
        begin
        grdMatrizConseq.Cell[3, j + 1].Text := FloatToStr(parA[j]);
        grdMatrizConseq.Cell[4, j + 1].Text := FloatToStr(parB[j]);
        end; }
      grdMatrizConseq.Cell[5, j + 1].Text := inttostr(Niveis[j]);
    end;

    for i := 0 to NumAlt - 1 do
    begin
      grdMatrizConseq.Cell[i + 7, 0].Text := NomeAlt[i];
      for j := 0 to NumCrit - 1 do
      begin
        grdMatrizConseq.Cell[i + 7, j + 1].Text := TrocaVP(ConsMatrix[i, j]);
      end;
    end;
    /// ////////////////////////////////////////////////////////////////////////
    //
    if (Problematica = 2) or (TIWUserSession(WebApplication.Data)
        .Problematica = 4) then
    Begin
      if TIWUserSession(WebApplication.Data).Flag <> -2 then
      begin
        with TIWUserSession(WebApplication.Data) do
        begin
          /// / Setando o tamanho dos grids que exibem a informação importada na região imput ////
          grdMatrizConseq.RowCount := NumAlt + 7;
          grdMatrizConseq.ColumnCount := NumCrit + 2;
          probname := edtEntername.Text;
          Numcolimport := grdImportproblema.Columns.Count;

          /// / Preenchendo a informação fixa dos grids ////
          grdMatrizConseq.Cell[0, 0].Text := 'Criteria:';
          grdMatrizConseq.Cell[1, 0].Text :=
            '0-Cont Min; 1-Cont Max; 2-Disc Min; 3- Disc Max; 4- Int Min; 5- Int Max:';
          { grdInputvalues.Cell[2, 0].text := 'Weights:'; }
          grdMatrizConseq.Cell[2, 0].Text := '';
          grdMatrizConseq.Cell[3, 0].Text := '';
          grdMatrizConseq.Cell[4, 0].Text := '';
          grdMatrizConseq.Cell[5, 0].Text :=
            'Number of levels of discrete criteria';
          grdMatrizConseq.Cell[6, 0].Text := 'Consequence Matrix:';
          grdMatrizConseq.Cell[0, NumCrit + 1].Text := 'Budget';
          // adicioando titulo pro budget

          /// / Preenchendo o Grid com a informação Imputada ///////////////////////
          grdMatrizConseq.Cell[1, NumCrit + 1].Text := TrocaVP(budget);
          for j := 0 to NumCrit - 1 do
          begin
            grdMatrizConseq.Cell[0, j + 1].Text := NomeCrit[j];
            if Cri_inteiro[j] then
            begin
              if Typeofcrit[j] = 0 then
                grdMatrizConseq.Cell[1, j + 1].Text := '4';
              if Typeofcrit[j] = 1 then
                grdMatrizConseq.Cell[1, j + 1].Text := '5';
            end
            else
              grdMatrizConseq.Cell[1, j + 1].Text := inttostr(Typeofcrit[j]);
            { grdMatrizConseq.Cell[2, j + 1].Text := inttostr(FTipo[j]);
              if FTipo[j] <> 1 then
              begin
              grdMatrizConseq.Cell[3, j + 1].Text := FloatToStr(parA[j]);
              grdMatrizConseq.Cell[4, j + 1].Text := FloatToStr(parB[j]);
              end; }
            grdMatrizConseq.Cell[5, j + 1].Text := inttostr(Niveis[j]);
          end;

          for i := 0 to NumAlt - 1 do
          begin
            grdMatrizConseq.Cell[i + 7, 0].Text := NomeAlt[i];
            grdMatrizConseq.Cell[i + 7, NumCrit + 1].Text := TrocaVP(Cost[i]);
            // texto com vetor dos Custo Unt dos proj

            for j := 0 to NumCrit - 1 do
            begin
              grdMatrizConseq.Cell[i + 7, j + 1].Text := TrocaVP
                (ConsMatrix[i, j]);
            end;
          end;
          /// ////////////////////////////////////////////////////////////////////////
          // end;

        end;
      end;
    End;
    /// ////////////////////////////////////////////////////
    if (TIWUserSession(WebApplication.Data).Problematica = 4) then
    // ADRIANA INICIO  Tentei incluir aqui mas nada feito
    begin
      btnStartelicitationClick(Self);
      rgnConseqvalues.Hide;
      if PortfolioG then
      begin
        // ** Manda os dados para a geração dos portfólios **//
        TIWUserSession(WebApplication.Data).AuxsalvarBD := 42;
        SalvarBD;
        rgnImportExport.Hide;
        rgnPortGeneration.show;
        tmrPortfoliocombination.Enabled := true;
      end;
    end; // ADRIANA FIM

  end; // with

  Inicializa;

  if TIWUserSession(WebApplication.Data).Flag <> -2 then
  begin
    // **Deixa visível a opção para seguir pra elicitação direto **//
    AlreadyRakend1.Visible := true;
    rgnImportExport.Visible := false;
    rgnInputMatrizConseq.Visible := true;
    // **Deixa visível a opção para seguir pra elicitação direto **//
    AlreadyRakend1.Visible := true;
    // apagar registros do banco de dados na tabela consequence
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 30;
    SalvarBD;

    // salvar banco de dados na tabela consequence
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 6;
    SalvarBD;

    // salvar banco de dados na tabela consequence
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 24;
    SalvarBD;

    // salvar banco de dados na tabela consequence
    TIWUserSession(WebApplication.Data).estagio := 0;
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 5;
    SalvarBD;

    if (TIWUserSession(WebApplication.Data).Problematica = 2) or
      (TIWUserSession(WebApplication.Data).Problematica = 4) then
    begin
      // se a problematica for de portfolio deve também chamar o caso que salva os valores dos custos
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 36;
      SalvarBD;
    end; // referente a  if TIWUserSession(WebApplication.Data).Problematica=2 then

    TIWUserSession(WebApplication.Data).tempinicio := TimeToStr(Now);
    rgnConseqvalues.Visible := false;
    rgnInputMatrizConseq.Visible := true;
    // **Deixa visível a opção para seguir pra elicitação direto **//
    AlreadyRakend1.Visible := true;

    /// / Após salvar o problema os listbox da pagina de cadastro manual já podem
    /// ser limpados///////////////////////////////////////////////////////////
    lstbxNamealternat.Clear; // Limpa conteúdo dos list box
    lstbxNamecriteria.Clear;

    // edtProjectname.Clear;
    edtNumberalternat.Text := '0';
    edtNumberattribut.Text := '0';
    // cmbbxproblem.ItemIndex := -1;

    edtNameattribut.Clear;
    edtparA.Clear;
    edtparB.Clear;
    edtNumberlevels.Clear;
    edtbudget.Clear;

    cmbxScaletype.ItemIndex := -1;
    cmbxFO.ItemIndex := -2;
    cmbxCriteriondirection.ItemIndex := -1;

    lblfunc.Visible := false;
    lblNumberlevels.Visible := false;
    edtNumberlevels.Visible := false;
    txtparametros.Visible := false;
    lblparA.Visible := false;
    lblparB.Visible := false;
    edtparA.Visible := false;
    edtparB.Visible := false;

    { --------- LIMPAR A MATRIZ DE EDITS ----------- }
    with TIWUserSession(WebApplication.Data) do
    begin
      for i := 0 to NumAlt - 1 do
      begin
        if length(ArrayAlt) > i then
          ArrayAlt[i].Destroy;

        if (TIWUserSession(WebApplication.Data).Problematica = 2) or
          (TIWUserSession(WebApplication.Data).Problematica = 4) then
        Begin
          for j := 0 to NumCrit do
          begin
            if length(MatEdit) > i then
              if length(MatEdit[i]) > j then
                MatEdit[i, j].Destroy;
            if i = 0 then
              if length(ArrayCrt) > j then
                TIWUserSession(WebApplication.Data).ArrayCrt[j].Destroy;
          end;
        End
        else
        begin
          for j := 0 to NumCrit - 1 do
          begin
            if length(MatEdit) > i then
              if length(MatEdit[i]) > j then
                MatEdit[i, j].Destroy;
            if i = 0 then
              if length(ArrayCrt) > j then
                ArrayCrt[j].Destroy;
          end;
        end;
      end;

      Setlength(MatEdit, 0, 0);
      Setlength(ArrayAlt, 0);
      Setlength(ArrayCrt, 0);
    end; // with

    TIWUserSession(WebApplication.Data).Estagiointra := -1;
    // inicializando var, quando entrada manual
    TIWUserSession(WebApplication.Data).fvround := -1; // var relacionado ao campo roundIntra do BD é inicializado em -1.
    if not(TIWUserSession(WebApplication.Data).Problematica = 4) then
      //TTFrmIntraCriterio.Create(WebApplication).show // cria unit intra
    else
    begin
      btnStartelicitationClick(Self);
      rgnConseqvalues.Hide;
    end;
  end; // if flag <> -2 (erro na inicialização)
end; // procedure

procedure TFrmEntrada.btnSendClick(Sender: TObject);
var
  genero, statuscurso, nescolaridade, P, email: string;
  vetchar: array [0 .. 9] of string;
  i, a: integer;
  H: boolean;
begin
  /// ////////////////////////////////////////////////////////////////////////////
  /// /////////  Travas para correto preenchimento do formulário /////////////////

  // ** Senhas diferentes **//
  if edtpassword.Text <> edtconfirmpass.Text then
  begin
    lblpasswarning.Visible := true;
    Exit;
  end
  else
  begin
    lblpasswarning.Visible := false;
  end;

  // ** Campos obrigatórios não preenchidos **//
  if (edtcademail.Text = '') or (edtcadname.Text = '') or
    (edtcadlast.Text = '') or (cmbbxAffiliation.ItemIndex = -1) or
    (edtpassword.Text = '') or (edtcadaddcity.Text = '') or
    (edtcadaddresscoun.Text = '') or (edtcadwhats.Text = '') or
    (IWComboBox1.ItemIndex = -1) then
  begin
    WebApplication.ShowMessage('Please, fill in all fields.');
    Exit;
  end;

  /// Caso haja vínculo com alguma instituição, seu nome e cidade devem ser////
  /// informados /////////////////////////////////////////////////////////////
  if cmbbxAffiliation.ItemIndex <> 3 then
    if (edtcadinstname.Text = '') or (edtcadaddresscity.Text = '') then
    begin
      WebApplication.ShowMessage('Please, fill in all fields.');
      Exit;
    end;
  /// /////////////////////////////////////////////////////////////////////////////
  vetchar[0] := '!';
  vetchar[1] := '@';
  vetchar[2] := '#';
  vetchar[3] := '$';
  vetchar[4] := '%';
  vetchar[5] := '&';
  vetchar[6] := '*';
  vetchar[7] := '^';
  vetchar[8] := '-';
  vetchar[9] := '+';

  email := edtcademail.Text;
  a := 0;
  for i := 1 to length(email) do
  begin
    if (email[i] = '@') or (email[i] = '.') then
      INC(a);
  end;
  if a < 2 then
  begin
    WebApplication.ShowMessage('Please, introduce a valid E-mail Address.');
    Exit;
  end;

  TIWUserSession(WebApplication.Data).ID_user := '';
  {with UserSession.ZQuery1 do
  begin
    SQL.Text := 'SELECT id FROM user WHERE email ="' + (edtcademail.Text)
      + '";';
    ExecSQL;
    open;
    TIWUserSession(WebApplication.Data).ID_user := FieldByName('id').AsString;
    close;
  end;  }

  // * se o e-mail já estiver cadastrado *//
  if TIWUserSession(WebApplication.Data).ID_user <> '' then
  begin
    WebApplication.ShowMessage(
      'The e-mail adress informed has already been registered. If you do not remember your password, it is possible to create a new one in the login page.');
    Exit;
  end;

  // preenchimento do genero
  if IWRadioGroup2.ItemIndex = 0 then
  begin
    genero := 'Female';
  end;
  if IWRadioGroup2.ItemIndex = 1 then
  begin
    genero := 'Male';
  end;
  if IWRadioGroup2.ItemIndex = 2 then
  begin
    genero := 'Not report';
  end;

  // preenchimento do nível de escolaridade Elementary School, High School, University Education, Specialization, Masters Degree, Doctorate Degree
  case IWComboBox1.ItemIndex of
    0:
      begin
        nescolaridade := 'Elementary School';
      end;
    1:
      begin
        nescolaridade := 'High School';
      end;
    2:
      begin
        nescolaridade := 'University Education';
      end;
    3:
      begin
        nescolaridade := 'Specialization';
      end;
    4:
      begin
        nescolaridade := 'Masters Degree';
      end;
    5:
      begin
        nescolaridade := 'Doctorate Degree';
      end;
  end;



  // preenchimento do em progresso ou finalizado

  if IWRadioGroup1.ItemIndex = 0 then
  begin
    statuscurso := 'In Progress';
  end;

  if IWRadioGroup1.ItemIndex = 1 then
  begin
    statuscurso := 'Finished';
  end;
  for i := 0 to 5 - 1 do
  begin
    P := FloatToStr(RandomRange(0, 9));
    TIWUserSession(WebApplication.Data).auxvalidacao := TIWUserSession
      (WebApplication.Data).auxvalidacao + FloatToStr
      (Round(RandomRange(0, 1000) / 10)) + vetchar[StrToInt(P)];

  end;
  // passando para variáveis globais
  if chckbxanalyst.Checked then
    TIWUserSession(WebApplication.Data).analyst := 0
  else
    TIWUserSession(WebApplication.Data).analyst := -1;
  TIWUserSession(WebApplication.Data).genero := genero;
  TIWUserSession(WebApplication.Data).statuscurso := statuscurso;
  TIWUserSession(WebApplication.Data).nescolaridade := nescolaridade;
  if edtcadtitle.Text <> '' then
    txtconfirmation.Lines[0] := txtconfirmation.Lines[0]
      + ' ' + edtcadtitle.Text + ' ' + edtcadname.Text + ','
  else
    txtconfirmation.Lines[0] := txtconfirmation.Lines[0]
      + ' ' + edtcadname.Text + ',';
  txtconfirmation.Lines[6] := txtconfirmation.Lines[6] + ' ' + TIWUserSession
    (WebApplication.Data).auxvalidacao;
  // P := 'To confirmate your FITradeoff Web Based account, introduce the the following code during your first access. CODE = ' + TIWUserSession(WebApplication.Data).auxvalidacao;
  H := false;
  H := send_email(email, 'Confirm your FITradeoff Web Based Account',
    txtconfirmation.Lines.Text);
  if H then
  begin
    WebApplication.ShowMessage(
      'Successful registration. A confirmation code was sent to your e-mail.' +
        ' You shall inform this code on your first access to FITradeoff system.' + ' Please, check also your spam in case the e-mail does not reach your inbox.');
  end
  else
  begin
    WebApplication.ShowMessage(
      'It was not possible to conclude the registration. Please, try again later.'
        + ' If the problem persist, let us know about it.');
    Exit;
  end;

  txtconfirmation.Lines[0] := 'Dear';
  txtconfirmation.Lines[6] := 'Activation Code =';
  // salvar banco de dados na tabela consequence
  TIWUserSession(WebApplication.Data).AuxsalvarBD := 0;
  SalvarBD;
  if H then
  begin
    rgnCadusuario.Visible := false;
    rgnEnterLogin.Visible := true;
    /// /////////////////////////////////////////////////////////////////////////
    // ** Limpando todos os campos do formulário **//
    edtconfirmpass.Clear;
    edtpassword.Clear;
    edtcadlast.Clear;
    edtcadname.Clear;
    edtcadtitle.Clear;
    edtcademail.Clear;
    edtcadaddcity.Clear;
    edtcadaddresscoun.Clear;
    edtcadinstname.Clear;
    edtcadaddresscity.Clear;
    edtcadwhats.Clear;
    IWComboBox1.ItemIndex := -1;
    cmbbxAffiliation.ItemIndex := -1;
    edtcruso.Clear;
    edtdia.Clear;
    edtmes.Clear;
    edtano.Clear;
    chckbxanalyst.Checked := false;
    /// /////////////////////////////////////////////////////////////////////////
  end;

end;

procedure TFrmEntrada.btnslotherfileClick(Sender: TObject);
begin
  /// // Limpa as variáveis ///////////////////////////////////////////////////////
  with TIWUserSession(WebApplication.Data) do
  begin
    NumCrit := 0;
    NumAlt := 0;

    Setlength(NomeCrit, 0);
    Setlength(FTipo, 0);
    Setlength(Typeofcrit, 0);
    Setlength(parA, 0);
    Setlength(parB, 0);
    Setlength(Niveis, 0);
    Setlength(NomeAlt, 0);
    Setlength(ConsMatrix, 0, 0);

    { if Length(ArrayCrt) > 0 then
      for i := 0 to Length(ArrayCrt) - 1 do
      ArrayCrt[i].Destroy;
      if Length(ArrayCrt) > 0 then
      for i := 0 to Length(ArrayAlt) - 1 do
      ArrayAlt[i].Destroy; }
  end;
  lblsheet.Visible := false;
  flImportarproblema.Visible := true;
  btnImporttabela.Enabled := true;
  btnauxstartelicitation.Enabled := false;

  IWLabel13.Visible := false;
  IWLabel14.Visible := false;
  IWLabel15.Visible := false;

end;

procedure TFrmEntrada.CostNorm;
var
  i, j: integer;
  auxcost: string;
begin
  if TIWUserSession(WebApplication.Data).Problematica = 2 then
  begin
    // ** Identificando o custo maximo **//
    TIWUserSession(WebApplication.Data).MaxCost := TIWUserSession
      (WebApplication.Data).Cost[0];
    TIWUserSession(WebApplication.Data).MinCost := TIWUserSession
      (WebApplication.Data).Cost[0];
    for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
    begin
      if TIWUserSession(WebApplication.Data).Cost[i] > TIWUserSession
        (WebApplication.Data).MaxCost then
        TIWUserSession(WebApplication.Data).MaxCost := TIWUserSession
          (WebApplication.Data).Cost[i];
      if TIWUserSession(WebApplication.Data).Cost[i] < TIWUserSession
        (WebApplication.Data).MinCost then
        TIWUserSession(WebApplication.Data).MinCost := TIWUserSession
          (WebApplication.Data).Cost[i];
    end;
    // ** Verifica a ordem de grandeza do maior custo **//
    auxcost := FloatToStr(TIWUserSession(WebApplication.Data).MinCost);
    j := 0;
    for i := 0 to length(auxcost) - 1 do
      if auxcost[i] <> '' then
        if (auxcost[i] = ',') or (auxcost[i] = '.') then
        begin
          INC(j, -1);
          Break;
        end
        else
          INC(j);
    // ** Constrói um múltiplo de 10 de grandeza semelhante ao custo máximo **//
    auxcost := '1';
    for i := 0 to j - 1 do
      auxcost := auxcost + '0';

    // ** Calcula o fator multiplicativo **//
    TIWUserSession(WebApplication.Data).fmult := TIWUserSession
      (WebApplication.Data).MinCost / StrToInt(auxcost);

    // ** Transformação do vetor de custos **//
    Setlength(TIWUserSession(WebApplication.Data).CostNorm, TIWUserSession
        (WebApplication.Data).NumAlt);

    for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
      TIWUserSession(WebApplication.Data).CostNorm[i] := TIWUserSession
        (WebApplication.Data).Cost[i] / TIWUserSession(WebApplication.Data)
        .MinCost;

  end; // problematica = 2
end;

procedure TFrmEntrada.Inicializa;
var
  i, j: integer;
  Mxv, Mnv, aux: Double;
  Mxvalor, MnValor: array of Double;
  /// Var criadas para fim de verificação ///
begin
  /// / Esse procedimento é reponsável por inicializar variáveis importantes //////
  /// utilizadas ao longo do SAD e que não são informadas diretamente pelo //////
  /// usuário. //////////////////////////////////////////////////////////////////
  // ** Limpa o listox da região de input **//
  // lstbxInitialOrder.Clear;

  Checar_Siglas;
  TIWUserSession(WebApplication.Data).Epson := Power(10, -8);
  // ** Definindo limiar para as restrições de Potencial Otimalidade **//
  TIWUserSession(WebApplication.Data).Epson := Power(10, -8);
  Setlength(TIWUserSession(WebApplication.Data).Bconseqord, TIWUserSession
      (WebApplication.Data).NumCrit);
  Setlength(TIWUserSession(WebApplication.Data).Wconseqord, TIWUserSession
      (WebApplication.Data).NumCrit);

  if TIWUserSession(WebApplication.Data).Problematica = 2 then
  begin
    edtMaxdif.Text := '0';
    edtMaxdif.Enabled := false;
    CostNorm;
  end;
  if TIWUserSession(WebApplication.Data).RETPROBDEPOIS <> -4555 then
  begin

    with TIWUserSession(WebApplication.Data) do // Comando para chamar as variáveis da userssession//
    begin
      Setlength(MatPositions, 0, 0);
      /// / Setando vetores que armazenam as maiores e menores consequências de cada critério ///
      Setlength(MaxCons, NumCrit);
      Setlength(MinCons, NumCrit);
      Setlength(Mxvalor, NumCrit);
      Setlength(MnValor, NumCrit);
      Setlength(class_min, NumAlt);
      Setlength(class_max, NumAlt);
      Setlength(Max, NumCrit);
      Setlength(Min, TIWUserSession(WebApplication.Data).NumCrit);
      Setlength(Cri_Cong, NumCrit);
      /// / Stenado vetor de alternativas Potencialmente Ótimas ////
      Setlength(VetPOA, NumAlt);
      /// Setando vetores que armazenam os limites das relações entre os critérios ////
      Setlength(LimMin, NumCrit - 1);
      Setlength(LimMax, NumCrit - 1);
      Setlength(LimMinA, NumCrit - 1);
      Setlength(LimMaxA, NumCrit - 1);
      Setlength(PergNAdj, 0, 0);
      Setlength(MLimMin, 0, 0);
      Setlength(MLimMax, 0, 0);
      Setlength(MLimMinA, 0, 0);
      Setlength(MLimMaxA, 0, 0);
      Setlength(Mblindados, 0, 0);
      /// / Setando vetor de relações blindadas (indica que não serão mais////////
      /// / feitas perguntas para as relações de critérios) ////
      Setlength(vtrrelacoesblin, NumCrit - 1);
      Setlength(HEINDEX, 2);
      Setlength(Vetorpergunta, NumCrit - 1);
      TotalAH := 0;
      TADC := false;
      PergDiscFinalized := false;

      /// arrays relacionados a avaliação intracriterio - alteração em 21.01.22
      Setlength(VtrcritBisec, NumCrit);
      Setlength(vtrLinearintra, NumCrit);
      Setlength(vtrChangeLin, NumCrit);
      Setlength(TIWUserSession(WebApplication.Data).Verificaintra, NumCrit);
      Setlength(TIWUserSession(WebApplication.Data).DinamVerificaIntra,
        NumCrit);
      Setlength(matrizfixa, length(matrizfixa), 2); // alteração em 19.01.22
      Setlength(BoundMn, NumCrit);
      Setlength(BoundMx, NumCrit);
      Contavazio := 0;

      // Indica se a AS foi realizada ou não
      // Seta como falso para ela não tentar exportar os resultados da AS depois que a página for fechada
      Sensitivity := false;
      Setlength(MatDataAS, 0);
      if (Problematica = 1) or (Problematica = 2) then
      begin
        Setlength(ArmazenaParaPar, 150, NumAlt, NumAlt);
        Ciclo := 0;
        NumRel := 0;
        NDRl := 0;
        NumInd := 0;
        ClpSolve := false;
      end;

      if Problematica = 3 then
      begin
        Alt_nao_classificadas := NumAlt;
        ASProfiles := false;
        // ** Só há um caminho agora **//
        elicitarperfis := true;
      end;

      // ** Supõem que só existem critérios discretos no problema **//
      TOADC := true;

      /// / Obtendo maiores e menores consequências de cada critério ////
      for j := 0 to NumCrit - 1 do
      begin
        /// Utiliza-se valores grandes/pequenos o suficientes para que sejam /////
        /// substituidos por qualquer valor de consequência //////////////////////
        Mxv := -50000000000;
        Mnv := 50000000000;
        for i := 0 to NumAlt - 1 do
        begin
          aux := ConsMatrix[i, j];
          if Mxv < aux then
            Mxv := aux;
          if Mnv > aux then
            Mnv := aux;
        end;
        MaxCons[j] := Mxv;
        Mxvalor[j] := Mxv;
        Max[j] := Mxv;
        MinCons[j] := Mnv;
        MnValor[j] := Mnv;
        Min[j] := Mnv;
        if Niveis[j] = 0 then
          TOADC := false;

        for i := 0 to NumCrit - 1 do
          Cri_Cong[i] := 0;

        { if Maxcons[j] = Mincons[j] then
          begin
          WebApplication.ShowMessage(
          'All consequences are the same in criterion ' + NomeCrit[j] +
          '. Please, verify the input data and then try again.');
          Flag := -2;
          Exit;
          end; }
      end;

      Flag := 0;

      /// / No início todas as alternativas são potencialmente ótimas /////
      for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
        TIWUserSession(WebApplication.Data).VetPOA[i] := i;

      /// / Inicializando vetores com os limites das razões entre os pesos ////
      for j := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
      begin
        /// //////** Corrige os vetores para a escala global **////////////////////
        if TIWUserSession(WebApplication.Data).Typeofcrit[j] > 1 then
        begin
          if TIWUserSession(WebApplication.Data).Niveis[j] <> 2 then
          begin
            TIWUserSession(WebApplication.Data).MaxCons[j] := Niveis[j];
            Mxvalor[j] := Niveis[j];
            Max[j] := Niveis[j];
            TIWUserSession(WebApplication.Data).MinCons[j] := 1;
            MnValor[j] := 1;
            Min[j] := 1;
            TADC := true;
            Setlength(DiscIntBlind, NumCrit);
            DiscIntBlind[j] := false;
          end
          else
          begin
            MaxCons[j] := 1;
            Mxvalor[j] := 1;
            Max[j] := 1;
            MinCons[j] := 0;
            MnValor[j] := 0;
            Min[j] := 0;
          end;
        end;
        /// /////////////////////////////////////////////////////////////////////
        if j < NumCrit - 1 then
        begin
          LimMin[j] := 0;
          LimMinA[j] := 0;
          LimMax[j] := 1;
          LimMaxA[j] := 1;
          TIWUserSession(WebApplication.Data).vtrrelacoesblin[j] := false;
          TIWUserSession(WebApplication.Data).Vetorpergunta[j] :=
            -0.00000000000005;
        end;
      end;

      // procedimento para identificar o que é criterio de minimização e o que é criterio de maximizãção
      for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
      begin
        if (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 1) or
          (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 3) then
        begin
          TIWUserSession(WebApplication.Data).Bconseqord[i] := MaxCons[i];
          TIWUserSession(WebApplication.Data).Wconseqord[i] := MinCons[i];
        end; // referente a if (Typeofcrit[i] = 0) or (Typeofcrit[i] = 2)
        if (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 0) or
          (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 2) then
        begin
          TIWUserSession(WebApplication.Data).Bconseqord[i] := MinCons[i];
          TIWUserSession(WebApplication.Data).Wconseqord[i] := MaxCons[i];
        end; // referente a if (Typeofcrit[i] = 0) or (Typeofcrit[i] = 2)
      end; // referente a for
      /// /////////////////////////////////////////////////////////////////////
      ///
      PAns := -1;
      TIFS := true;
      NumCriCong := 0;
      NumCriBlind := 0;
      HeurQuest := -1;
      TIWUserSession(WebApplication.Data).IndiceA := 0;
      IndiceB := NumCrit - 1;
      NumQuest := 0;
      NumRQuestion := 0;
      NDA := NumAlt;
      NumAH := 0;
      // setlength(TIWUserSession(WebApplication.Data).MATAH, 0, 0);
      Sol := -5;
      Analystkey := 0;
      auxbdgraph := '-1';
      TIWUserSession(WebApplication.Data).auxas := 0;
      TIWUserSession(WebApplication.Data).auxid_ah := 0;
      contblin := 0;
      zeramatriz := false;
      MinVK := 0.000000001;
      Etapa := 0;
      // Normaliza;
      /// teste vtr
      Setlength(TIWUserSession(WebApplication.Data).vtrRecomendacao_ciclo, 1);
    end;
  end;

  if TIWUserSession(WebApplication.Data).RETPROBDEPOIS = -4555 then
  begin
    with TIWUserSession(WebApplication.Data) do // Comando para chamar as variáveis da userssession//
    begin
      /// / Setando vetores que armazenam as maiores e menores consequências de cada critério ///
      Setlength(HEINDEX, 2);
      Setlength(MatPositions, 0, 0);
      Setlength(MaxCons, NumCrit);
      Setlength(MinCons, NumCrit);
      Setlength(Mxvalor, NumCrit);
      Setlength(MnValor, NumCrit);
      Setlength(Max, NumCrit);
      Setlength(Min, NumCrit);
      Setlength(Cri_Cong, NumCrit);
      /// Setando vetores que armazenam os limites das relações entre os critérios ////
      Setlength(LimMin, NumCrit - 1);
      Setlength(LimMax, NumCrit - 1);
      Setlength(LimMinA, NumCrit - 1);
      Setlength(LimMaxA, NumCrit - 1);

      /// / Setando vetor de relações blindadas (indica que não serão mais////////
      /// / feitas perguntas para as relações de critérios) ////
      Setlength(vtrrelacoesblin, NumCrit - 1);
      Setlength(HEINDEX, 2);
      Setlength(Vetorpergunta, NumCrit - 1);
      Setlength(TIWUserSession(WebApplication.Data).class_min, NumAlt);
      Setlength(TIWUserSession(WebApplication.Data).class_max, NumAlt);

      /// arrays relacionados a avaliação intracriterio - alterado em 21.01.22
      Setlength(VtrcritBisec, NumCrit);
      Setlength(vtrLinearintra, NumCrit);
      Setlength(vtrChangeLin, NumCrit);
      Setlength(Verificaintra, NumCrit);
      Setlength(DinamVerificaIntra, NumCrit);
      // setlength(matrizfixa, 2, numcrit);
      Setlength(matrizfixa, length(matrizfixa), 2); // alteração em 19.01.22
      Setlength(BoundMn, NumCrit);
      Setlength(BoundMx, NumCrit);
      Contavazio := 0;
      TADC := false;
      Sensitivity := false;
      Setlength(MatDataAS, 0);
      EtapaAnt := 0;

      if (Problematica = 1) or (Problematica = 2) then
      begin
        Setlength(ArmazenaParaPar, 150, NumAlt, NumAlt);
        // Ciclo := 0;
        NumRel := 0;
        NDRl := 0;
        NumInd := 0;
      end;
      ASProfiles := false;

      // ** Supõem que só existem critérios discretos no problema **//
      TOADC := true;

      /// / Obtendo maiores e menores consequências de cada critério ////
      for j := 0 to NumCrit - 1 do
      begin
        /// Utiliza-se valores grandes/pequenos o suficientes para que sejam /////
        /// substituidos por qualquer valor de consequência //////////////////////
        Mxv := -50000000000;
        Mnv := 50000000000;
        for i := 0 to NumAlt - 1 do
        begin
          aux := ConsMatrix[i, j];
          if Mxv < aux then
            Mxv := aux;
          if Mnv > aux then
            Mnv := aux;
        end;
        MaxCons[j] := Mxv;
        Mxvalor[j] := Mxv;
        Max[j] := Mxv;
        MinCons[j] := Mnv;
        MnValor[j] := Mnv;
        Min[j] := Mnv;
        if Niveis[j] = 0 then
          TOADC := false;

        { if Maxcons[j] = Mincons[j] then
          begin
          WebApplication.ShowMessage(
          'All consequences are the same in criterion ' + NomeCrit[j] +
          '. Please, verify the input data and then try again.');
          Flag := -2;
          Exit;
          end; }
      end;
      Flag := 0;

      /// / Inicializando vetores com os limites das razões entre os pesos ////
      for j := 0 to NumCrit - 1 do
      begin
        /// //////** Corrige os vetores para a escala global **////////////////////
        if Typeofcrit[j] > 1 then
        begin
          if Niveis[j] <> 2 then
          begin
            TIWUserSession(WebApplication.Data).MaxCons[j] := Niveis[j];
            Mxvalor[j] := Niveis[j];
            Max[j] := Niveis[j];
            TIWUserSession(WebApplication.Data).MinCons[j] := 1;
            MnValor[j] := 1;
            Min[j] := 1;
            TADC := true;
            Setlength(DiscIntBlind, NumCrit);
            DiscIntBlind[j] := false;
          end
          else
          begin
            MaxCons[j] := 1;
            Mxvalor[j] := 1;
            Max[j] := 1;
            MinCons[j] := 0;
            MnValor[j] := 0;
            Min[j] := 0;
          end;
        end;
      end;
      /// /////////////////////////////////////////////////////////////////////
      // procedimento para identificar o que é criterio de minimização e o que é criterio de maximizãção
      for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
      begin
        if (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 1) or
          (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 3) then
        begin
          TIWUserSession(WebApplication.Data).Bconseqord[i] := MaxCons[i];
          TIWUserSession(WebApplication.Data).Wconseqord[i] := MinCons[i];
        end; // referente a if (Typeofcrit[i] = 0) or (Typeofcrit[i] = 2)
        if (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 0) or
          (TIWUserSession(WebApplication.Data).Typeofcrit[i] = 2) then
        begin
          TIWUserSession(WebApplication.Data).Bconseqord[i] := MinCons[i];
          TIWUserSession(WebApplication.Data).Wconseqord[i] := MaxCons[i];
        end; // referente a if (Typeofcrit[i] = 0) or (Typeofcrit[i] = 2)
      end; // referente a for
      /// /////////////////////////////////////////////////////////////////////

      TIFS := true;
      Alt_nao_classificadas := NumAlt;
      NumRQuestion := 0;
      TIWUserSession(WebApplication.Data).NDA := length
        (TIWUserSession(WebApplication.Data).VetPOA);
      Sol := -5;
      Analystkey := 0;
      auxbdgraph := '-1';
      TIWUserSession(WebApplication.Data).auxas := 0;
      contblin := 0;
      zeramatriz := false;
      MinVK := 0.000000001;
      // Normaliza;
      if (Etapa = 3) and (TADC) then
        // ** Se existirem critérios discretos e todos blindados, tenta retomar
        // ** a Elicitação de níveis intermediários **//
        Etapa := 4;
      // Normaliza;

    end; // REFERENTE A  with TIWUserSession(WebApplication.Data)
  end; // ende do RETPROBDEPOIS=-4555

  // ** Independente da origem dos dados, deve-se setar o número de casas dec **//
  Definir_Decimais;
end;

procedure TFrmEntrada.edtnewpassconfirClick(Sender: TObject);
var
  senha: string;
begin
 { with UserSession.ZQuery1 do
  begin
    if edtnewpassword.Text = edtnewpasswordconfirm.Text then
      senha := edtnewpassword.Text
    else
    begin
      lbldifcamp.Visible := true;
      Exit;
    end;

    with UserSession.ZQuery1 do
    begin
      // salvando no BD, o campo validacao atual
      SQL.Text := 'UPDATE user SET validacao ="validado" WHERE email="' +
        (TIWUserSession(WebApplication.Data).email) + '";';
      ExecSQL;

      // salvando no BD,a senha temporária
      SQL.Text := 'UPDATE user SET password ="' + (senha) + '" WHERE email="' +
        (TIWUserSession(WebApplication.Data).email) + '";';
      ExecSQL;
    end;


    WebApplication.ShowMessage('Password succesfuly saved.');

    edtnewpassword.Text := '';
    edtnewpasswordconfirm.Text := '';

    rgnvalidationpassword.Visible := false;
    rgnnewpassword.Visible := false;
    rgnEnterLogin.Visible := false;
    rgnproblematica.Visible := true;

  end;
      }
end;

procedure TFrmEntrada.IWButton1Click(Sender: TObject);
begin
  rgnvalidation.Visible := false;
  rgnnewpassword.Visible := false;
  rgnvalidationpassword.Visible := false;
  rgnEnterLogin.Visible := false;
  rgnintro.Visible := true;
  IWLabel9.Visible := false;
  IWLabel8.Visible := false;

  /// /////////////////////////////////////////////////////////////////////////////
  /// ////////////* Limpando edit's de login, validacao *//////////////////////////
  edtemail.Clear;
  edtSenha.Clear;
  edtvalidation.Clear;
  edtnewpassword.Clear;
  edtnewpasswordconfirm.Clear;
  /// /////////////////////////////////////////////////////////////////////////////
end;

procedure TFrmEntrada.IWButton2Click(Sender: TObject);
begin
  WebApplication.ShowMessage(
    'The number of criteria and alternatives are automatically incremented by adding them to the problem.');
end;

procedure TFrmEntrada.IWButton3Click(Sender: TObject);
begin
  rgnlanguage.Visible := true;

end;

procedure TFrmEntrada.IWButton5Click(Sender: TObject);
begin
  // chamar a região que escolhe o que deseja fazer, importar matriz, entrada manual ou retornar problema

  if (TIWUserSession(WebApplication.Data).Problematica = -1) then
  // verificação para não avaçar sem escolher um modulo
  begin
    WebApplication.ShowMessage('Please, select a module.');
    Exit;
  end // SHOW MESSAGE
  else
  begin
    rgnOpcaoInput.Visible := true;
    rgnproblematica.Visible := false;
    alinhamento;
  end;
end;

procedure TFrmEntrada.btnovoproblemaClick(Sender: TObject);
begin
TiwForm1.Create(WebApplication).Show;
end;

procedure TFrmEntrada.IWButton7Click(Sender: TObject);
/// /realiza transição de tela, caso o usuario queira trocar o tipo de modulo
/// / verificar se é necessário/interessante limpar mais algo neste (edit, label, var...)
begin
  rgnOpcaoInput.Visible := false;
  rgnproblematica.Visible := true;
  lblproblematic1.Caption := '';
  lblproblematic2.Caption := '';
  lblchosenproblematic.Caption := '';
  TIWUserSession(WebApplication.Data).Problematica := -1;
end;

procedure TFrmEntrada.IWButton8Click(Sender: TObject);
begin
  rgnCofirmReset.Visible := false;
  btnSaveproblem.Enabled := true;
  btnSaveStart.Enabled := true;
end;

procedure TFrmEntrada.btnConfResetClick(Sender: TObject);
var
  i, j { , cont_preenchido } : integer;

begin
  with TIWUserSession(WebApplication.Data) do

  begin
    rgnConseqvalues.Visible := false;
    rgnCofirmReset.Visible := false;
    rgnOpcaoInput.Visible := true;
    // rgnproblematica.Visible := true;
    lblbudget.Visible := false;
    edtbudget.Visible := false;
    hptpbud.Visible := false;

    lstbxNamealternat.Clear; // Limpa conteúdo dos list box
    lstbxNamecriteria.Clear;

    // limpando marcação do tipo de problematica
    rgnchoice.BorderOptions.NumericWidth := 0;
    rgnrank.BorderOptions.NumericWidth := 0;
    rgnsorting.BorderOptions.NumericWidth := 0;
    rgnport.BorderOptions.NumericWidth := 0;
    // problematica:= -1;

    edtProjectname.Clear;
    edtbudget.Clear;
    edtNumberalternat.Text := '0';
    edtNumberattribut.Text := '0';
    lstbxNamealternat.Clear;
    lstbxNamecriteria.Clear;

    // cmbbxproblem.ItemIndex := -1; //atribuido problematica=-1 no final do procedimento para zerar

    { --------- LIMPAR A MATRIZ DE EDITS ----------- }
    with TIWUserSession(WebApplication.Data) do
    begin
      for i := 0 to NumAlt - 1 do
      begin
        if length(ArrayAlt) > i then
          ArrayAlt[i].Destroy;

        If (Problematica = 2) or (TIWUserSession(WebApplication.Data)
            .Problematica = 4) then
        begin
          for j := 0 to NumCrit do
          begin
            if length(MatEdit) > i then
              if length(MatEdit[i]) > j then
                MatEdit[i, j].Destroy;
            if i = 0 then
              if length(ArrayCrt) > j then
                TIWUserSession(WebApplication.Data).ArrayCrt[j].Destroy;
          end;
        end
        else
        Begin
          for j := 0 to NumCrit - 1 do
          begin
            if length(MatEdit) > i then
              if length(MatEdit[i]) > j then
                MatEdit[i, j].Destroy;
            if i = 0 then
              if length(ArrayCrt) > j then
                ArrayCrt[j].Destroy;
          end;
        End;
      end;

      Setlength(MatEdit, 0, 0);
      Setlength(ArrayAlt, 0);
      Setlength(ArrayCrt, 0);

    end; // with
  end; // user session
end;

procedure TFrmEntrada.btnguideClick(Sender: TObject);
begin
  rgnlanguage.Visible := true;

end;

procedure TFrmEntrada.btnlogoutrgnconsClick(Sender: TObject);
var
  i, j: integer;
begin
  Logout;

  // **  Chama o salvamento dos dados para evitar perdas **//
  btnSaveproblemClick(Self);

  // ** Necessário excluir todos os componentes dinâmicos da tela **//
  with TIWUserSession(WebApplication.Data) do
  begin
    for i := 0 to NumAlt - 1 do
    begin
      if length(ArrayAlt) > i then
        ArrayAlt[i].Destroy;

      if (TIWUserSession(WebApplication.Data).Problematica = 2) or
        (TIWUserSession(WebApplication.Data).Problematica = 4) then
      Begin
        for j := 0 to NumCrit do
        begin
          if length(MatEdit) > i then
            if length(MatEdit[i]) > j then
              MatEdit[i, j].Destroy;
          if i = 0 then
            if length(ArrayCrt) > j then
              TIWUserSession(WebApplication.Data).ArrayCrt[j].Destroy;
        end;
      End
      else
      begin
        for j := 0 to NumCrit - 1 do
        begin
          if length(MatEdit) > i then
            if length(MatEdit[i]) > j then
              MatEdit[i, j].Destroy;
          if i = 0 then
            if length(ArrayCrt) > j then
              ArrayCrt[j].Destroy;
        end;
      end;
    end; // for i

    Setlength(MatEdit, 0, 0);
    Setlength(ArrayAlt, 0);
    Setlength(ArrayCrt, 0);
  end; // with

  rgnLoginproject.Visible := false;
end;

procedure TFrmEntrada.IWDBLabel1AsyncClick(Sender: TObject;
  EventParams: TStringList);
begin
  btnExporttemplateClick(Self);
end;

{ procedure TFrmEntrada.IWLabel47AsyncClick(Sender: TObject;
  EventParams: TStringList);
  begin

  end; }

procedure TFrmEntrada.IWLink11Click(Sender: TObject);
/// /realiza transição de tela, caso o usuario queira trocar o tipo de modulo
/// / verificar se é necessário/interessante limpar mais algo neste (edit, label, var...)
begin
  rgnOpcaoInput.Visible := false;
  rgnproblematica.Visible := true;
  lblproblematic1.Caption := '';
  lblproblematic2.Caption := '';
  lblchosenproblematic.Caption := '';
  TIWUserSession(WebApplication.Data).Problematica := -1;
end;

procedure TFrmEntrada.IWLink12Click(Sender: TObject);
begin
  rgnCadprojeto.Visible := false;
  rgnOpcaoInput.Visible := true;

  // Limpa os edts e lisbox
  edtProjectname.Clear;
  edtNumberalternat.Text := '0';
  edtNumberattribut.Text := '0';
  edtNamealternat.Clear;
  edtNameattribut.Clear;
  edtNumberlevels.Clear;
  edtbudget.Clear;
  lstbxNamealternat.Items.Clear;
  lstbxNamecriteria.Items.Clear;

  edtparA.Clear;
  edtparB.Clear;
  edtNumberlevels.Clear;

  cmbxFO.ItemIndex := -2;
  cmbxScaletype.ItemIndex := -1;
  cmbxCriteriondirection.ItemIndex := -1;
end;

procedure TFrmEntrada.IWLink13Click(Sender: TObject);
begin
  rgnEnterLogin.Visible := false;
  rgnOpcaoInput.Visible := false;
  rgnImportExport.Visible := false;
  rgnCadprojeto.Visible := false;
  rgnInputMatrizConseq.Visible := false;
  rgnLoginproject.Visible := false;
  rgnConseqvalues.Visible := false;
  rgnproblematica.Visible := false;
  rgnintro.Visible := true;
  edtemail.Clear;
  edtSenha.Clear;
  lblbudget.Visible := false; // adicionado aqui para garantir que componente não aparecerá qndo sair
  lbllembretebudget.Visible := false; // adicionado aqui para garantir que componente não aparecerá qndo sair
  edtbudget.Visible := false; // adicionado aqui para garantir que componente não aparecerá qndo sair
  hptpbud.Visible := false; // adicionado aqui para garantir que componente não aparecerá qndo sair

  rgnchoice.BorderOptions.NumericWidth := 0;
  rgnrank.BorderOptions.NumericWidth := 0;
  rgnsorting.BorderOptions.NumericWidth := 0;
  rgnport.BorderOptions.NumericWidth := 0;
  TIWUserSession(WebApplication.Data).Problematica := 0;

end;

procedure TFrmEntrada.lnkcloseaboutClick(Sender: TObject);
begin
  rgnAbout.Hide;
end;

procedure TFrmEntrada.cmproblematica1Change(Sender: TObject);
begin
  TIWUserSession(WebApplication.Data).Problematica := cmproblematica1.ItemIndex;
  if (cmproblematica1.ItemIndex <> -1) and (edtEntername.Text <> '') then
  end;

  procedure TFrmEntrada.cmbbxAffiliationChange(Sender: TObject);
  begin
    if cmbbxAffiliation.ItemIndex = 3 then
    begin
      lblAffiliation1.Visible := false;
      lblAffiliation2.Visible := false;
    end
    else
    begin
      lblAffiliation1.Visible := true;
      lblAffiliation2.Visible := true;
    end;
  end;

  procedure TFrmEntrada.cmbbxproblemChange(Sender: TObject);
  begin
    TIWUserSession(WebApplication.Data).Problematica := cmbbxproblem.ItemIndex;
    if (cmbbxproblem.ItemIndex <> -1) and (edtEntername.Text <> '') then
      flImportarproblema.Enabled := true;
  end;

  procedure TFrmEntrada.cmbxFOChange(Sender: TObject);
  begin
    lblfunc.Visible := true;
    if cmbxFO.ItemIndex <> 1 then
    begin
      txtparametros.Visible := true;
      lblparA.Visible := true;
      lblparB.Visible := true;
      edtparA.Visible := true;
      edtparB.Visible := true;
      edtparA.Clear;
      edtparB.Clear;
      if cmbxFO.ItemIndex = 0 then
        lblfunc.Caption := 'F(x) = bln(ax) + c';
      if cmbxFO.ItemIndex = 2 then
        lblfunc.Caption := 'F(x) = be^(-a/x) + c';
      if cmbxFO.ItemIndex = 3 then
        lblfunc.Caption := 'F(x) = be^(ax) + c';
    end
    else
    begin
      txtparametros.Visible := false;
      lblparA.Visible := false;
      lblparB.Visible := false;
      edtparA.Visible := false;
      edtparB.Visible := false;
      lblfunc.Visible := false;
    end;
  end;

  procedure TFrmEntrada.cmbxScaletypeChange(Sender: TObject);
  begin
    if cmbxScaletype.ItemIndex = 1 then
    begin
      lblNumberlevels.Visible := true;
      edtNumberlevels.Visible := true;
      chckbxinteger.Visible := false;
      lblfunc.Visible := false; // só tem que aparecer quando altera cmbbxFO
      lblTipoproblema.Visible := true; // se discreto, obrigatoriamente vai liberar componentes da FTipo
      cmbxFO.Visible := true;
      cmbxFO.ItemIndex := 1; // garante que sempre esteja desmarcado
      rdbtndeclaretype.Checked := true; // sempre desmarcar, pra não dar problema com visualização de componentes
      rdbtnelictfunction.Checked := false;
      rdbtnelictfunction.Enabled := false
      // rgnacaointra.Visible:= false; //não dá a aopção de escolher sobre FO

    end
    else
    begin
      lblNumberlevels.Visible := false;
      edtNumberlevels.Visible := false;
      chckbxinteger.Visible := true;
      lblTipoproblema.Visible := false; // se continuo, é facultativo
      cmbxFO.ItemIndex := -2; // garante que sempre esteja desmarcado
      cmbxFO.Visible := false;
      lblfunc.Visible := false; // a partir daqui são os componentes caso FO <>1
      txtparametros.Visible := false;
      lblparA.Visible := false;
      edtparA.Visible := false;
      lblparB.Visible := false;
      edtparB.Visible := false;
      rdbtnelictfunction.Enabled := true;
      rdbtndeclaretype.Checked := false; // sempre desmarcar, pra não dar problema com visualização de componentes
      rdbtnelictfunction.Checked := false;
      // rgnacaointra.Visible := true; //exibe rgn para decisor escolher como fornecerá FO

      edtNumberlevels.Clear;
    end;
  end;

  procedure TFrmEntrada.btnStartelicitationClick(Sender: TObject);
  // procedimento verifica tipo de prblematica e direciona para correspondnete, no recebimento do grid
  begin
    with TIWUserSession(WebApplication.Data) do
    Begin
      if (Problematica = 0) or (Problematica = 1) or (Problematica = 3) then
        StartelicitationChoiRank;

      if (Problematica = 2) then
        // portfolio
        StartelicitationPortfolioCB;

      if (Problematica = 4) then
      begin
        StartelicitationPortfolioComb;
        if PortfolioG then
        begin
          // ** Manda os dados para a geração dos portfólios **//
          TIWUserSession(WebApplication.Data).AuxsalvarBD := 42;
          SalvarBD;
          rgnImportExport.Hide;
          rgnPortGeneration.show;
          tmrPortfoliocombination.Enabled := true;
        end;
      end;

      AlreadyRakend1.Visible := true;
    end; // user
  end; // procedimento

  procedure TFrmEntrada.StartelicitationChoiRank;
  // procedimento responsável por receber excel importado e preencher o grid exibido ao usuário na região para ordenação das ctes
  var
    i, j: integer;
    Nomeprojeto: string;
  begin
    with TIWUserSession(WebApplication.Data) do

    Begin
      if (TIWUserSession(WebApplication.Data).Problematica = 1) or
        (Problematica = 0) or (Problematica = 3) then
      begin
        if TIWUserSession(WebApplication.Data).VerificaUnit <> -50 then
        begin
          { if cmbbxproblem.ItemIndex = -1 then
            begin
            WebApplication.ShowMessage(
            'Please, inform which problematic you wish to evaluate.');
            tmrcont2.Enabled := false;
            btnauxstartelicitation.Enabled := true;
            Exit;
            end; }

          // salvando o nome do projeto
          Nomeprojeto := edtEntername.Text;
          TIWUserSession(WebApplication.Data).Nomeprojeto := Nomeprojeto;
        end;

        /// / Esse procedimento é responsável por realizar a transferências dos //////
        /// dados  para a região input ///////////////////////////////////////////////

        Inicializa;

        if TIWUserSession(WebApplication.Data).Flag <> -2 then
        begin
          with TIWUserSession(WebApplication.Data) do
          begin
            /// / Setando o tamanho dos grids que exibem a informação importada na região imput ////
            grdMatrizConseq.RowCount := NumAlt + 7;
            grdMatrizConseq.ColumnCount := NumCrit + 1;
            probname := edtEntername.Text;

            /// / Preenchendo a informação fixa dos grids ////
            grdMatrizConseq.Cell[0, 0].Text := 'Criteria:';
            grdMatrizConseq.Cell[1, 0].Text :=
              '0-Cont Min; 1-Cont Max; 2-Disc Min; 3- Disc Max; 4- Int Min; 5- Int Max:';
            { grdInputvalues.Cell[2, 0].text := 'Weights:'; }
            grdMatrizConseq.Cell[2, 0].Text := '';
            grdMatrizConseq.Cell[3, 0].Text := '';
            grdMatrizConseq.Cell[4, 0].Text := '';
            grdMatrizConseq.Cell[5, 0].Text :=
              'Number of levels of discrete criteria';
            grdMatrizConseq.Cell[6, 0].Text := 'Consequence Matrix:';

            /// / Preenchendo os Grids com a informação Imputada ///////////////////////
            for j := 0 to NumCrit - 1 do
            begin
              grdMatrizConseq.Cell[0, j + 1].Text := NomeCrit[j];
              if Cri_inteiro[j] then
              begin
                if Typeofcrit[j] = 0 then
                  grdMatrizConseq.Cell[1, j + 1].Text := '4';
                if Typeofcrit[j] = 1 then
                  grdMatrizConseq.Cell[1, j + 1].Text := '5';
              end
              else
                grdMatrizConseq.Cell[1, j + 1].Text := inttostr(Typeofcrit[j]);
              { grdMatrizConseq.Cell[2, j + 1].Text := inttostr(FTipo[j]);
                if FTipo[j] <> 1 then
                begin
                grdMatrizConseq.Cell[3, j + 1].Text := FloatToStr(parA[j]);
                grdMatrizConseq.Cell[4, j + 1].Text := FloatToStr(parB[j]);
                end; }
              grdMatrizConseq.Cell[5, j + 1].Text := inttostr(Niveis[j]);
            end;

            for i := 0 to NumAlt - 1 do
            begin
              grdMatrizConseq.Cell[i + 7, 0].Text := NomeAlt[i];
              for j := 0 to NumCrit - 1 do
              begin
                grdMatrizConseq.Cell[i + 7, j + 1].Text := TrocaVP
                  (ConsMatrix[i, j]);
              end;
            end;
            /// ////////////////////////////////////////////////////////////////////////
            // end;

            rgnImportExport.Visible := false;
            // Ou o problema está sendo reiniciado ou a avaliação intra já foi performada
            if (TIWUserSession(WebApplication.Data).VerificaUnit = -50) and
              (not PortfolioG) then
              rgnInputMatrizConseq.Visible := true;

            // **Deixa visível a opção para seguir pra elicitação direto **//
            AlreadyRakend1.Visible := true;
            // **Deixa visível a opção para seguir pra elicitação direto **//
            AlreadyRakend1.Visible := true;
          end;
        end;

        // Limpa tudo caso o usuário volte para essa página depois
        flImportarproblema.Enabled := true;
        edtEntername.Text := '';
        // cmbbxproblem.ItemIndex := -1;
        IWLabel13.Visible := false;
        IWLabel14.Visible := false;
        IWLabel15.Visible := false;

        if TIWUserSession(WebApplication.Data).VerificaUnit <> -50 then
        begin

          // salvar banco de dados na tabela de problema
          // TIWUserSession(WebApplication.Data).ID_Problema := idUnique;
          // gerando um codigo para o problema
          TIWUserSession(WebApplication.Data).AuxsalvarBD := 1;
          SalvarBD;
          // salvar banco de dados na tabela de alternativas
          TIWUserSession(WebApplication.Data).AuxsalvarBD := 3;
          SalvarBD;
          // salvar banco de dados na tabela de criterios
          TIWUserSession(WebApplication.Data).AuxsalvarBD := 2;
          SalvarBD;
          // salvar banco de dados na tabela problem
          TIWUserSession(WebApplication.Data).estagio := 0;
          TIWUserSession(WebApplication.Data).AuxsalvarBD := 5;
          SalvarBD;
          // salvar banco de dados na tabela consequence
          TIWUserSession(WebApplication.Data).AuxsalvarBD := 6;
          SalvarBD;
          // salvar banco de dados na tabela consequence
          { TIWUserSession(WebApplication.Data).AuxsalvarBD := 9;
            SalvarBD; }

          TIWUserSession(WebApplication.Data).tempinicio := TimeToStr(Now);

          // salvar banco de dados na tabela  round e id_problem na tabela current_stage
          TIWUserSession(WebApplication.Data).AuxsalvarBD := 10;
          SalvarBD;

          // salvar banco de dados na tabela current_stage
          TIWUserSession(WebApplication.Data).BDstatus := '0';
          TIWUserSession(WebApplication.Data).AuxsalvarBD := 11;
          SalvarBD;
          // end;

          // ** Avaliação intracritério **//
          //TTFrmIntraCriterio.Create(WebApplication).show;
          // cria unt intracriterio
        end;

        // ** Elementos para dar a sensação de carregamento **//
        btnauxstartelicitation.Enabled := true;
        rgnImportExport.Cursor := CRAUTO;
        tmrcont2.Enabled := false;
      end;
    end; // conexão user session

    // flImportarproblema.Enabled := true;
    // lblsheet.Caption:= '';

  end; // procedimento

  Procedure TFrmEntrada.StartelicitationPortfolioCB;
  // procedimento responsável por receber excel importado e preencher o grid exibido ao usuário na região para ordenação das ctes
  var
    i, j: integer;
    Nomeprojeto: string;
    Numcolimport: integer;

  begin
    if TIWUserSession(WebApplication.Data).VerificaUnit <> -50 then
    begin
      { if cmbbxproblem.ItemIndex = -1 then
        begin
        WebApplication.ShowMessage(
        'Please, inform which problematic you wish to evaluate.');
        tmrcont2.Enabled := false;
        btnauxstartelicitation.Enabled := true;
        Exit;
        end; }

      // salvando o nome do projeto
      Nomeprojeto := edtEntername.Text;
      TIWUserSession(WebApplication.Data).Nomeprojeto := Nomeprojeto;
    end;

    /// / Esse procedimento é responsável por realizar a transferências dos //////
    /// dados  para a região input ///////////////////////////////////////////////

    Inicializa;

    if TIWUserSession(WebApplication.Data).Flag <> -2 then
    begin
      with TIWUserSession(WebApplication.Data) do
      begin
        /// / Setando o tamanho dos grids que exibem a informação importada na região imput ////
        grdMatrizConseq.RowCount := NumAlt + 7;
        grdMatrizConseq.ColumnCount := NumCrit + 2;
        probname := edtEntername.Text;
        Numcolimport := grdImportproblema.Columns.Count;

        /// / Preenchendo a informação fixa dos grids ////
        grdMatrizConseq.Cell[0, 0].Text := 'Criteria:';
        grdMatrizConseq.Cell[1, 0].Text :=
          '0-Cont Min; 1-Cont Max; 2-Disc Min; 3- Disc Max; 4- Int Min; 5- Int Max:';
        { grdInputvalues.Cell[2, 0].text := 'Weights:'; }
        grdMatrizConseq.Cell[2, 0].Text := '';
        grdMatrizConseq.Cell[3, 0].Text := '';
        grdMatrizConseq.Cell[4, 0].Text := '';
        grdMatrizConseq.Cell[5, 0].Text :=
          'Number of levels of discrete criteria';
        grdMatrizConseq.Cell[6, 0].Text := 'Consequence Matrix:';
        grdMatrizConseq.Cell[0, NumCrit + 1].Text := 'Budget';
        // adicioando titulo pro budget

        /// / Preenchendo o Grid com a informação Imputada ///////////////////////
        grdMatrizConseq.Cell[1, NumCrit + 1].Text := TrocaVP(budget);
        for j := 0 to NumCrit - 1 do
        begin
          grdMatrizConseq.Cell[0, j + 1].Text := NomeCrit[j];
          if Cri_inteiro[j] then
          begin
            if Typeofcrit[j] = 0 then
              grdMatrizConseq.Cell[1, j + 1].Text := '4';
            if Typeofcrit[j] = 1 then
              grdMatrizConseq.Cell[1, j + 1].Text := '5';
          end
          else
            grdMatrizConseq.Cell[1, j + 1].Text := inttostr(Typeofcrit[j]);
          { grdMatrizConseq.Cell[2, j + 1].Text := inttostr(FTipo[j]);
            if FTipo[j] <> 1 then
            begin
            grdMatrizConseq.Cell[3, j + 1].Text := FloatToStr(parA[j]);
            grdMatrizConseq.Cell[4, j + 1].Text := FloatToStr(parB[j]);
            end; }
          grdMatrizConseq.Cell[5, j + 1].Text := inttostr(Niveis[j]);
        end;

        for i := 0 to NumAlt - 1 do
        begin
          grdMatrizConseq.Cell[i + 7, 0].Text := NomeAlt[i];
          grdMatrizConseq.Cell[i + 7, NumCrit + 1].Text := TrocaVP(Cost[i]);
          // texto com vetor dos Custo Unt dos proj

          for j := 0 to NumCrit - 1 do
          begin
            grdMatrizConseq.Cell[i + 7, j + 1].Text := TrocaVP
              (ConsMatrix[i, j]);
          end;
        end;
        /// ////////////////////////////////////////////////////////////////////////
        // end;
        // **Deixa visível a opção para seguir pra elicitação direto **//
        AlreadyRakend1.Visible := true;

        rgnImportExport.Visible := false;
        // Ou o problema está sendo reiniciado ou a avaliação intra já foi performada
        if TIWUserSession(WebApplication.Data).VerificaUnit = -50 then
          rgnInputMatrizConseq.Visible := true;
      end;
    end;

    // Limpa tudo caso o usuário volte para essa página depois
    flImportarproblema.Enabled := true;
    edtEntername.Text := '';
    // cmbbxproblem.ItemIndex := -1;
    IWLabel13.Visible := false;
    IWLabel14.Visible := false;
    IWLabel15.Visible := false;

    if TIWUserSession(WebApplication.Data).VerificaUnit <> -50 then
    begin

      // salvar banco de dados na tabela de problema
      // TIWUserSession(WebApplication.Data).ID_Problema := idUnique;
      // gerando um codigo para o problema
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 1;
      SalvarBD;
      // salvar banco de dados na tabela de alternativas
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 3;
      SalvarBD;
      // salvar banco de dados na tabela de criterios
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 2;
      SalvarBD;
      // salvar banco de dados na tabela problem
      TIWUserSession(WebApplication.Data).estagio := 0;
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 5;
      SalvarBD;
      // salvar banco de dados na tabela consequence
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 6;
      SalvarBD;
      // salvar banco de dados na tabela consequence
      { TIWUserSession(WebApplication.Data).AuxsalvarBD := 9;
        SalvarBD; }

      TIWUserSession(WebApplication.Data).tempinicio := TimeToStr(Now);

      // salvar banco de dados na tabela  round e id_problem na tabela current_stage
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 10;
      SalvarBD;

      // salvar banco de dados na tabela current_stage
      TIWUserSession(WebApplication.Data).BDstatus := '0';
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 11;
      SalvarBD;

      // ** Avaliação intracritério **//
      //TTFrmIntraCriterio.Create(WebApplication).show; // cria unt intracriterio
    end;

    // ** Elementos para dar a sensação de carregamento **//
    btnauxstartelicitation.Enabled := true;
    rgnImportExport.Cursor := CRAUTO;
    tmrcont2.Enabled := false;

  end; // procedimento

  Procedure TFrmEntrada.StartelicitationPortfolioComb;
  // procedimento responsável por receber excel importado e preencher o grid exibido ao usuário na região para ordenação das ctes
  var
    i, j: integer;
    Nomeprojeto: string;
    Numcolimport: integer;
    auxcost: string;

  begin
    if TIWUserSession(WebApplication.Data).VerificaUnit <> -50 then
    begin
      { if cmbbxproblem.ItemIndex = -1 then
        begin
        WebApplication.ShowMessage(
        'Please, inform which problematic you wish to evaluate.');
        tmrcont2.Enabled := false;
        btnauxstartelicitation.Enabled := true;
        Exit;
        end; }

      // salvando o nome do projeto
      Nomeprojeto := edtEntername.Text;
      if Nomeprojeto <> '' then
        TIWUserSession(WebApplication.Data).Nomeprojeto := Nomeprojeto;
    end;

    /// / Esse procedimento é responsável por realizar a transferências dos //////
    /// dados  para a região input ///////////////////////////////////////////////

    Inicializa;

    if TIWUserSession(WebApplication.Data).Flag <> -2 then
    begin
      with TIWUserSession(WebApplication.Data) do
      begin
        /// / Setando o tamanho dos grids que exibem a informação importada na região imput ////
        grdMatrizConseq.RowCount := NumAlt + 7;
        grdMatrizConseq.ColumnCount := NumCrit + 2;
        probname := edtEntername.Text;
        Numcolimport := grdImportproblema.Columns.Count;

        /// / Preenchendo a informação fixa dos grids ////
        grdMatrizConseq.Cell[0, 0].Text := 'Criteria:';
        grdMatrizConseq.Cell[1, 0].Text :=
          '0-Cont Min; 1-Cont Max; 2-Disc Min; 3- Disc Max; 4- Int Min; 5- Int Max:';
        { grdInputvalues.Cell[2, 0].text := 'Weights:'; }
        grdMatrizConseq.Cell[2, 0].Text := '';
        grdMatrizConseq.Cell[3, 0].Text := '';
        grdMatrizConseq.Cell[4, 0].Text := '';
        grdMatrizConseq.Cell[5, 0].Text :=
          'Number of levels of discrete criteria';
        grdMatrizConseq.Cell[6, 0].Text := 'Consequence Matrix:';
        grdMatrizConseq.Cell[0, NumCrit + 1].Text := 'Budget';
        // adicioando titulo pro budget

        /// / Preenchendo o Grid com a informação Imputada ///////////////////////
        grdMatrizConseq.Cell[1, NumCrit + 1].Text := TrocaVP(budget);
        for j := 0 to NumCrit - 1 do
        begin
          grdMatrizConseq.Cell[0, j + 1].Text := NomeCrit[j];
          if Cri_inteiro[j] then
          begin
            if Typeofcrit[j] = 0 then
              grdMatrizConseq.Cell[1, j + 1].Text := '4';
            if Typeofcrit[j] = 1 then
              grdMatrizConseq.Cell[1, j + 1].Text := '5';
          end
          else
            grdMatrizConseq.Cell[1, j + 1].Text := inttostr(Typeofcrit[j]);
          { grdMatrizConseq.Cell[2, j + 1].Text := inttostr(FTipo[j]);
            if FTipo[j] <> 1 then
            begin
            grdMatrizConseq.Cell[3, j + 1].Text := FloatToStr(parA[j]);
            grdMatrizConseq.Cell[4, j + 1].Text := FloatToStr(parB[j]);
            end; }
          grdMatrizConseq.Cell[5, j + 1].Text := inttostr(Niveis[j]);
        end;

        for i := 0 to NumAlt - 1 do
        begin
          grdMatrizConseq.Cell[i + 7, 0].Text := NomeAlt[i];
          grdMatrizConseq.Cell[i + 7, NumCrit + 1].Text := TrocaVP(Cost[i]);
          // texto com vetor dos Custo Unt dos proj

          for j := 0 to NumCrit - 1 do
          begin
            grdMatrizConseq.Cell[i + 7, j + 1].Text := TrocaVP
              (ConsMatrix[i, j]);
          end;
        end;
        /// ////////////////////////////////////////////////////////////////////////
        // end;
        // **Deixa visível a opção para seguir pra elicitação direto **//
        AlreadyRakend1.Visible := true;

        rgnImportExport.Visible := false;
        // Ou o problema está sendo reiniciado ou a avaliação intra já foi performada
        if TIWUserSession(WebApplication.Data).VerificaUnit = -50 then
          rgnInputMatrizConseq.Visible := true;
      end;
    end;

    // Limpa tudo caso o usuário volte para essa página depois
    flImportarproblema.Enabled := true;
    edtEntername.Text := '';
    // cmbbxproblem.ItemIndex := -1;
    IWLabel13.Visible := false;
    IWLabel14.Visible := false;
    IWLabel15.Visible := false;

    if (TIWUserSession(WebApplication.Data).VerificaUnit <> -50) and
      (TIWUserSession(WebApplication.Data).PortfolioG) then
    begin

      // salvar banco de dados na tabela de problema
      // TIWUserSession(WebApplication.Data).ID_Problema := idUnique;
      // gerando um codigo para o problema
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 1;
      SalvarBD;
      // salvar banco de dados na tabela de alternativas
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 3;
      SalvarBD;
      // salvar banco de dados na tabela de criterios
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 2;
      SalvarBD;
      // salvar banco de dados na tabela problem
      TIWUserSession(WebApplication.Data).estagio := 0;
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 5;
      SalvarBD;
      // salvar banco de dados na tabela consequence
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 6;
      SalvarBD;
      // salvar banco de dados na tabela consequence
      { TIWUserSession(WebApplication.Data).AuxsalvarBD := 9;
        SalvarBD; }

      TIWUserSession(WebApplication.Data).tempinicio := TimeToStr(Now);

      // salvar banco de dados na tabela  round e id_problem na tabela current_stage
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 10;
      SalvarBD;

      // salvar banco de dados na tabela current_stage
      TIWUserSession(WebApplication.Data).BDstatus := '0';
      TIWUserSession(WebApplication.Data).AuxsalvarBD := 11;
      SalvarBD;

      // cria unt intracriterio
    end;

    if (TIWUserSession(WebApplication.Data).VerificaUnit <> -50) and
      (not TIWUserSession(WebApplication.Data).PortfolioG) then
      // ** Avaliação intracritério **//
      //TTFrmIntraCriterio.Create(WebApplication).show;

    // ** Elementos para dar a sensação de carregamento **//
    btnauxstartelicitation.Enabled := true;
    rgnImportExport.Cursor := CRAUTO;
    tmrcont2.Enabled := false;
  end; // procedimento

  procedure TFrmEntrada.btnStep1Click(Sender: TObject);
  begin
    TIWUserSession(WebApplication.Data).stage := 1;
    TIWUserSession(WebApplication.Data).epsilon := trocapv(edtMaxdif.Text);
    rgnInputMatrizConseq.Visible := false;
    Tfrmaintercriterio.Create(WebApplication).show;

    // salvar banco de dados na tabela de problema
    TIWUserSession(WebApplication.Data).estagio := 1;
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 5;
    SalvarBD;

    // salvar banco de dados na tabela consequence
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 9;
    SalvarBD;

    // salvar banco de dados na tabela current_stage
    TIWUserSession(WebApplication.Data).BDstatus := '0';
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 11;
    SalvarBD;

    // **Elementos para impressão de carregamento **//
    rgnInputMatrizConseq.Cursor := CRAUTO;
    btnauxstep2.Cursor := CRAUTO;
    btnauxstep1.Cursor := CRAUTO;
    btnauxstep1.Enabled := true;
    btnauxstep2.Enabled := true;
    tmrstep.Enabled := false;

    /// chamar botão gerar gráfico///
  end;

  procedure TFrmEntrada.btnStep2Click(Sender: TObject);
  begin
    if TIWUserSession(WebApplication.Data).RETPROBDEPOIS <> -4555 then
      TIWUserSession(WebApplication.Data).stage := 2;
    TIWUserSession(WebApplication.Data).epsilon := trocapv(edtMaxdif.Text);
    rgnInputMatrizConseq.Visible := false;

    // Verificação para passar ou não direto para a unit de resultados
    if (TIWUserSession(WebApplication.Data).Problematica = 3) and
      (TIWUserSession(WebApplication.Data).elicitarperfis = true) then
    begin
      Tfrmaintercriterio.Create(WebApplication).btnprofileelicitClick(Self);
    end;
    if (TIWUserSession(WebApplication.Data).Problematica <> 3) then
    begin
      TIWUserSession(WebApplication.Data).elicitarperfis := false;
      Tfrmaintercriterio.Create(WebApplication).BtnshowresultsClick(Self);
    end;
    // salvar banco de dados na tabela de problema
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 4;
    SalvarBD;
    // salvar banco de dados na tabela de problema
    TIWUserSession(WebApplication.Data).estagio := 2;
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 5;
    SalvarBD;
    // salvar banco de dados na tabela consequence
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 9;
    SalvarBD;

    // salvar banco de dados na tabela current_stage
    TIWUserSession(WebApplication.Data).BDstatus := '0';
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 11;
    SalvarBD;

    // salvar banco de dados na tabela consequence
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 17;
    SalvarBD;

    // salvar banco de dados na tabela consequence
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 20;
    SalvarBD;

    // salvar banco de dados na tabela consequence
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 26;
    SalvarBD;

    // salvar o tipo_ord no BD
    TIWUserSession(WebApplication.Data).auxtipoord := -1;
    TIWUserSession(WebApplication.Data).AuxsalvarBD := 28;
    SalvarBD;

    // **Elementos para impressão de carregamento **//
    rgnInputMatrizConseq.Cursor := CRAUTO;
    btnauxstep2.Cursor := CRAUTO;
    btnauxstep1.Cursor := CRAUTO;
    btnauxstep1.Enabled := true;
    btnauxstep2.Enabled := true;
    tmrstep.Enabled := false;
  end;

  procedure TFrmEntrada.btnvalidaremailClick(Sender: TObject);
  var
    validation: string;
  begin
   { with UserSession.ZQuery1 do
    begin

      SQL.Text := 'SELECT validacao FROM user WHERE email ="' + (edtemail.Text)
        + '"';
      ExecSQL;
      open;
      validation := FieldByName('validacao').AsString;
      close;
    end;
    // * Verifica se o código inserido bate com o eviado *//
    if validation = edtvalidation.Text then
    begin

      with UserSession.ZQuery1 do
      begin
        // salvando no BD, como a planilha problema já será atualizada no inicio deve-se apenas atualizar o campo Cri_Order
        SQL.Text := 'UPDATE user SET validacao ="validado" WHERE email="' +
          (TIWUserSession(WebApplication.Data).email) + '";';
        ExecSQL;
      end;

      WebApplication.ShowMessage(
        'Your FITradeoff account has been successfuly confirmed.');

      TIWUserSession(WebApplication.Data).primeiroacesso := true;
      rgnvalidationpassword.Visible := false;
      rgnvalidation.Visible := false;
      rgnEnterLogin.Visible := false;
      rgnproblematica.Visible := true;
      // rgnOpcaoInput.Visible := true;
      edtvalidation.Clear;
      edtSenha.Clear;
      edtemail.Clear;
      btnContinueregist.Enabled := false;
    end
    else
    begin
      WebApplication.ShowMessage(
        'The code introduced does not match with the sent one. Please verify your e-mail address and try again.');
      WebApplication.ShowMessage(
        'You should insert all the numbers and characters present in the code. You can ask for a new confirmation code if you need.');
    end;
    edtvalidation.Clear;
                          }
  end;

  procedure TFrmEntrada.btnVoltarinicioClick(Sender: TObject);
  begin
    rgnCadusuario.Visible := false;
    rgnintro.Visible := true;
  end;

  Function TFrmEntrada.trocapv(Num: String): Double;
  var
    S: string;
    i: integer;
  begin
    S := Num;
    for i := 2 to length(S) - 1 do
      if S[i] = '.' then
        S[i] := ',';

    Result := StrToFloat(S);
  end;

  Function TFrmEntrada.TrocaVP(Num: Double): String;
  var
    S: string;
    i: integer;
  begin
    S := FloatToStr(Num);
    for i := 2 to length(S) - 1 do
      if S[i] = ',' then
        S[i] := '.';

    Result := S;
  end;

  Function TFrmEntrada.trocavp1(Num: String): Double;
  var
    S: string;
    i: integer;
  begin
    S := Num;
    for i := 2 to length(S) - 1 do
      if S[i] = ',' then
        S[i] := '.';

    Result := StrToFloat(S);
  end;

  function TFrmEntrada.idUnique: String;
  // ATENÇÃO SÓ PODE SER CHAMDO UMA VEZ.
  var
    gID: TGuid;
  begin
    CreateGUID(gID);
    Result := GUIDToString(gID);
  end;

  procedure TFrmEntrada.imagechoiceClick(Sender: TObject);
  begin
    rgnchoice.BorderOptions.Color := clWebGAINSBORO;
    rgnchoice.BorderOptions.NumericWidth := 2;
    rgnrank.BorderOptions.NumericWidth := 0;
    rgnsorting.BorderOptions.NumericWidth := 0;
    rgnport.BorderOptions.NumericWidth := 0;
    rgnportComb.BorderOptions.NumericWidth := 0;
    TIWUserSession(WebApplication.Data).Problematica := 0;
    lblchosenproblematic.Caption := 'Choice';
    lblproblematic2.Caption := 'Choice';
    lblproblematic1.Caption := 'Choice';
    IWButton5.Enabled := true;
  end;

  procedure TFrmEntrada.imageportClick(Sender: TObject);
  begin
    rgnport.BorderOptions.Color := clWebGAINSBORO;
    rgnchoice.BorderOptions.NumericWidth := 0;
    rgnrank.BorderOptions.NumericWidth := 0;
    rgnsorting.BorderOptions.NumericWidth := 0;
    rgnportComb.BorderOptions.NumericWidth := 0;
    rgnport.BorderOptions.NumericWidth := 2;
    TIWUserSession(WebApplication.Data).Problematica := 2;
    lblchosenproblematic.Caption := 'Portfolio BC';
    lblproblematic2.Caption := 'Portfolio BC';
    lblproblematic1.Caption := 'Portfolio BC';
    IWButton5.Enabled := true;
    lblbudget.Visible := true;
    edtbudget.Visible := true;
    hptpbud.Visible := true;
  end;

  procedure TFrmEntrada.imagerankClick(Sender: TObject);
  begin
    rgnrank.BorderOptions.Color := clWebGAINSBORO;
    rgnchoice.BorderOptions.NumericWidth := 0;
    rgnrank.BorderOptions.NumericWidth := 2;
    rgnsorting.BorderOptions.NumericWidth := 0;
    rgnport.BorderOptions.NumericWidth := 0;
    rgnportComb.BorderOptions.NumericWidth := 0;
    TIWUserSession(WebApplication.Data).Problematica := 1;
    lblchosenproblematic.Caption := 'Ranking';
    lblproblematic2.Caption := 'Ranking';
    lblproblematic1.Caption := 'Ranking';
    IWButton5.Enabled := true;
  end;

  procedure TFrmEntrada.imagesortingClick(Sender: TObject);
  begin
    rgnsorting.BorderOptions.Color := clWebGAINSBORO;
    rgnchoice.BorderOptions.NumericWidth := 0;
    rgnrank.BorderOptions.NumericWidth := 0;
    rgnsorting.BorderOptions.NumericWidth := 2;
    rgnport.BorderOptions.NumericWidth := 0;
    rgnportComb.BorderOptions.NumericWidth := 0;
    TIWUserSession(WebApplication.Data).Problematica := 3;
    lblchosenproblematic.Caption := 'Sorting';
    lblproblematic2.Caption := 'Sorting';
    lblproblematic1.Caption := 'Sorting';
    IWButton5.Enabled := true;
  end;

  procedure TFrmEntrada.imgCabecalhoClick(Sender: TObject);
  begin
    WebApplication.GoToURL('http://cdsid.org.br/');
  end;

  procedure TFrmEntrada.imgPortCombClick(Sender: TObject);
  begin
    rgnportComb.BorderOptions.Color := clWebGAINSBORO;
    rgnchoice.BorderOptions.NumericWidth := 0;
    rgnrank.BorderOptions.NumericWidth := 0;
    rgnsorting.BorderOptions.NumericWidth := 0;
    rgnport.BorderOptions.NumericWidth := 0;
    rgnportComb.BorderOptions.NumericWidth := 2;
    TIWUserSession(WebApplication.Data).Problematica := 4;
    lblchosenproblematic.Caption := 'Combinatorial Portfolio';
    lblproblematic2.Caption := 'Combinatorial Portfolio';
    lblproblematic1.Caption := 'Combinatorial Portfolio';
    IWButton5.Enabled := true;
    lblbudget.Visible := true;
    edtbudget.Visible := true;
    hptpbud.Visible := true;
  end;

  procedure TFrmEntrada.SalvarBD;
  var
    AuxsalvarBD, auxcount, id_cri, id_alt, auxiniciomx, auxdecision,
      auxnumaltavaliadas, H: integer;
    Auxproblematica, id_pro, axipoaunico, auxnumresp, axirelblin,
      auxcrirelblin_, axiavali, axiavaunico: string;
    NumCrit, i, j, P, auxcontAH: integer;
    auxtype: array of string;
    auxdirection, vtrobjetivo, vtrtypo1: array of string;
    NumAlt, id_round: integer;
    time, auxordem_, auxcriblin_, auxrepre, auxrere_, axilimmx, axilimmxunico,
      auxnumquest: string;
    datainicio, auxaltavaliadas, auxid_alt_e, auxid_alt_c, axicricong,
      auxcricong_: string;
    Axuordem, idproblemaaux, axior, axiblin, axirere, axipoa1, auxpoavolta,
      auxHeurQuest, axiorunico: string;
    Auxnomesobrenome, user_, id_status: string;
    Auxdatanascimento, axprob1, datafinal, auxlimax, axilimmin, auxlimin,
      axilimminunico: string;
    ic, fn, tt: ttime;
    Auxcri_cong, Auxcriposi, axicriposi, auxcriposi_, axireprep,
      axirepunico: string;
    axilimminunicoA, auxliminA, axilimminA, axilimmxA, auxlimaxA,
      axilimmxunicoA: string;
    auxcontAHh, id_ah_total, contadorativas: integer;
    aux_ret_problmatic: string;
    aux_boolean_posi: array of boolean;
    aux_ult_nivel_posi, aux_dois_digitos: string;
    retorno_problematic: string;
    aux_versao_barra: integer;
    axiposin, axiposinunico: string;
    question_aux: integer;
    matrix_aux_domina: array of array of integer;
    aux_relations, str, StrAux, Str1, strAux1, Str2: string;

    /// auxiliares utilizadas nos casos referentes á etapa intracriterio- 38, 39, 40
    auxavalia, auxround, auxvalround, aux2norm: integer;
    auxId, auxboundmx, auxbarra, auxbarramx, auxboundmn, auxbarra1, auxbarramn,
      auxpercentual, auxbarrapec, auxbarra2, auxftipo, auxbarratipo, auxbarra3,
      auxbarrabis, auxbarra4, auxcritbisec, auxbarralin, auxvtrlin, auxbarra5,
      auxbarracl, auxvtrcl, auxbarra6, auxbarraNI, auxbarra7, auxNorm,
      auxValor, aux_norm, auxbarrapontos, auxpontos, auxbarra8, auxbarraVI,
      auxverificaI, auxbarra0, auxbarraDVI, auxDinverI, auxbarra10,
      auxbarranew, auxNftipo, auxbarra11: string;
    // ** Auxiliares Portfólio Combinatório **//
    sConsequences, sAlternatives, sCustos, sBudget, scriDirection: string;
    percentual: real;

    // ** Mapear queda ** //  Yara
   myDate: TDateTime;
   myYear, myMonth, myDay: Word;
   myHour, myMin, mySec, myMilli: Word;
   Data_completa,IP: String;

  begin
    NumAlt := TIWUserSession(WebApplication.Data).NumAlt;
    NumCrit := TIWUserSession(WebApplication.Data).NumCrit;
    // setando o tamanho dos vetores  que são locais
    Setlength(auxdirection, NumCrit);
    Setlength(auxtype, NumCrit);
    Setlength(TIWUserSession(WebApplication.Data).ID_Alternative, NumAlt);
    Setlength(TIWUserSession(WebApplication.Data).ID_criteria, NumCrit);
    Setlength(TIWUserSession(WebApplication.Data).ConsMatrix, NumAlt, NumCrit);
    // Setlength(TIWUserSession(WebApplication.Data).ConsNorm, NumAlt, NumCrit);
    AuxsalvarBD := TIWUserSession(WebApplication.Data).AuxsalvarBD;
    //with UserSession.ZQuery1 do
    //begin


    //end;

  end;

initialization

//TFrmEntrada.SetAsMainForm;

end.
