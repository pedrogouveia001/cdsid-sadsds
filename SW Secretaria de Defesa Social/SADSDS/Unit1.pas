unit Unit1;

interface

uses
  Classes, SysUtils, IWAppForm, IWApplication, IWColor, IWTypes, IWCompListbox,
  IWExtCtrls, IWCompButton, IWCompMemo, IWCompLabel, IWVCLBaseControl,
  IWBaseControl, IWBaseHTMLControl, IWControl, IWCompEdit, Controls, Forms,
  IWVCLBaseContainer, IWContainer, IWHTMLContainer, IWHTML40Container, IWRegion,
  UserSessionUnit, Dialogs, ComObj, Variants, jpeg, windows, ActiveX, IWWebGrid,
  IWAdvWebGrid, IWAdvWebGridExcel, Menus, IWCompMenu, IWCompTabControl, Math,
  pngimage, IWCompText, IWHTMLControls, chart, IWBaseComponent,
  IWBaseHTMLComponent, IWBaseHTML40Component, dateutils, TeeGDIPlus, TeEngine,
  ExtCtrls, TeeProcs, Series;

type
  TIWForm1 = class(TIWAppForm)
    rgnManualProblem: TIWRegion;
    btngoinput3: TIWButton;
    btnbacktoSA: TIWButton;
    IWImage6: TIWImage;
    IWImage7: TIWImage;
    IWImage8: TIWImage;
    IWLabel18: TIWLabel;
    lstbxcriterio: TIWListbox;
    IWLabel21: TIWLabel;
    edtnomecriterio: TIWEdit;
    IWLabel31: TIWLabel;
    rdgrpdirecaocriterio: TIWRadioGroup;
    btndeletecriterio: TIWButton;
    btnupdatecriterio: TIWButton;
    btnnewcriterio: TIWButton;
    IWLabel2: TIWLabel;
    IWLabel51: TIWLabel;
    IWLabel4: TIWLabel;
    lstbxalternativa: TIWListbox;
    IWLabel44: TIWLabel;
    edtnomealternativa: TIWEdit;
    btndeletealternativa: TIWButton;
    btnupdatealternativa: TIWButton;
    btnnewalternativa: TIWButton;
    IWRegion5: TIWRegion;
    rgnintro: TIWRegion;
    lblWarning: TIWLabel;
    btnInputManual: TIWButton;
    btnImportar: TIWButton;
    IWButton4: TIWButton;
    IWLabel46: TIWLabel;
    IWImage9: TIWImage;
    IWLabel70: TIWLabel;
    IWImage10: TIWImage;
    IWImage27: TIWImage;
    IWImage23: TIWImage;
    IWRegion2: TIWRegion;
    IWRegion16: TIWRegion;
    IWRegion3: TIWRegion;
    IWText2: TIWText;
    rgnInicial: TIWRegion;
    btnLogin: TIWButton;
    btnRegister: TIWButton;
    IWImage2: TIWImage;
    IWLink3: TIWLink;
    IWImage1: TIWImage;
    IWImage29: TIWImage;
    IWText1: TIWText;
    rgnConsequenceMatrix: TIWRegion;
    btncontinue: TIWButton;
    IWImage3: TIWImage;
    IWImage4: TIWImage;
    IWImage5: TIWImage;
    IWLabel7: TIWLabel;
    IWRegion4: TIWRegion;
    rgnDisplaymatrizconsequencia: TIWRegion;
    btnsubmit: TIWButton;
    btnAlterar: TIWButton;
    IWLabel3: TIWLabel;
    IWLink1: TIWLink;
    IWLink2: TIWLink;
    btnsalvarpesos: TIWButton;
    IWLabel5: TIWLabel;
    rgnimport: TIWRegion;
    IWLabel6: TIWLabel;
    IWLabel8: TIWLabel;
    FileArquivo: TIWFile;
    btnimport: TIWButton;
    IWLink4: TIWLink;
    grdImport: TTIWAdvWebGrid;
    TIWAdvWebGridExcelIO1: TTIWAdvWebGridExcelIO;
    IWLabel9: TIWLabel;
    IWLabel11: TIWLabel;
    Chart1: TChart;
    Series1: TBarSeries;
    IWLabel12: TIWLabel;
    procedure IWAppFormCreate(Sender: TObject);
    procedure btndeletecriterioClick(Sender: TObject);
    procedure btnupdatecriterioClick(Sender: TObject);
    procedure btnnewcriterioClick(Sender: TObject);
    procedure btndeletealternativaClick(Sender: TObject);
    procedure btnupdatealternativaClick(Sender: TObject);
    procedure btnnewalternativaClick(Sender: TObject);
    procedure btngoinput3Click(Sender: TObject);
    procedure CriarString(Linhas, Colunas: integer; Regiao: TIWRegion);
    procedure btnInputManualClick(Sender: TObject);
    procedure IWLink1Click(Sender: TObject);
    procedure IWLink2Click(Sender: TObject);
    procedure lstbxcriterioChange(Sender: TObject);
    procedure lstbxalternativaChange(Sender: TObject);
    procedure btnAlterarClick(Sender: TObject);
    procedure IWLink4Click(Sender: TObject);
    procedure btnimportClick(Sender: TObject);
    procedure ImportExcel;
    procedure btnImportarClick(Sender: TObject);
    procedure btnsubmitClick(Sender: TObject);
    procedure btnsalvarpesosClick(Sender: TObject);
    procedure btncontinueClick(Sender: TObject);
    procedure Inicializa;
    procedure Definir_Decimais;
    procedure Normaliza;
    procedure Siglas;
    procedure rgnManualProblemCreate(Sender: TObject);
    function TrocaVirgPPto(Valor: string): String;
  public
  end;

implementation

uses ServerController, lpsolve51, uAPI_SolveLP, uCode, uHasseDiagram,
  uIntercritério, uResultados, untIntro;

//uses UserSessionUnit, ServerController;

{$R *.dfm}


function TiwForm1.TrocaVirgPPto(Valor: string): String;
var
  i: integer;
begin
  if Valor <> '' then
  begin
    for i := 0 to Length(Valor) do
    begin
      if Valor[i] = ',' then
      begin
        Valor[i] := '.';
      end
      else if Valor[i] = ' .' then
      begin
        Valor[i] := '.';
      end;
    end;
  end;
  Result := Valor;
end;

procedure TiwForm1.CriarString(Linhas, Colunas: integer;
  Regiao: TIWRegion);
var
  i, j, auxi, auxj: integer;
  lbl: TIWLabel;
  Texto: string;
begin
With TIWUserSession(WebApplication.Data) do begin
Siglas;

  if auxalt <> 0 then
  begin
  for i := 0 to TIWUserSession(WebApplication.Data).auxalt - 1 do
    lblalt[i].free;
  end;

  if auxcrit <> 0 then
  begin
   for j := 0 to TIWUserSession(WebApplication.Data).auxcrit - 1 do
   begin
     lblcrit[j].free;
     lbldircrit[j].Free;
     lblauxcrit[j].Free;
     edtpesos[0,j].Free;
     edtpesos[1,j].Free;
   end;
  end;




  SetLength(lblcrit, NumCrit+1);
  SetLength(lbldircrit, NumCrit+1);
  SetLength(lblalt, NumAlt+1);


  for j := 0 to Colunas - 1 do
  begin
      //lblcrit[j].Free;
      //lbldircrit[j].Free;
      TIWUserSession(WebApplication.Data).lblcrit[j] := TIWLabel.Create(Regiao);
      lbldircrit[j] := TIWLabel.Create(Regiao);
      lbldircrit[j].Parent := Regiao;
      lbldircrit[j].Left := 95 + 54 * j;
      lbldircrit[j].Top := 15;

      lblcrit[j].Parent := Regiao;
      lblcrit[j].Left := 100 + 54 * j;
      lblcrit[j].Caption := CritSiglas[j];

      if (TIWUserSession(WebApplication.Data).Typeofcrit[j] = 0) or (TIWUserSession(WebApplication.Data).Typeofcrit[j] = 2) or (TIWUserSession(WebApplication.Data).Typeofcrit[j] = 4)  then
      lbldircrit[j].Caption := '(Min.)'
      else
      lbldircrit[j].Caption := '(Max.)';


      lblcrit[j].Font.FontFamily := 'Verdana, Helvetica, Sans-Serif';
      lbldircrit[j].Font.FontFamily := 'Verdana, Helvetica, Sans-Serif';
      //lblcrit[j].Name := 'lblCelLinha' + IntToStr(j);
  end;

  {lbl := TIWLabel.Create(rgnpesos);
  lbl.Parent := rgnpesos;
  lbl.left := 0;
  lbl.Top := 35;
  lbl.Caption := 'Limite Superior';

  lbl := TIWLabel.Create(rgnpesos);
  lbl.Parent := rgnpesos;
  lbl.left := 0;
  lbl.Top := 105;
  lbl.Caption := 'Limite Inferior';   }

  setlength(edtpesos, 2, NumCrit+1);
  setlength(lblauxcrit, NumCrit+1);

  {for j := TIWUserSession(WebApplication.Data).auxalt to Colunas - 1 do
  begin
  lbl := TIWLabel.Create(rgnpesos);
  lbl.Parent := rgnpesos;
  lbl.Left := 90 + 54 * j;
  lbl.Top :=  10;
  lbl.Caption := 'C' + inttostr(j+1);
  end; }



  for j := 0 to Colunas - 1 do
  begin
  //inc(auxmanual);
  //lblauxcrit[j].Free;
  lblauxcrit[j] := TIWLabel.Create(rgnconsequencematrix);
  lblauxcrit[j].Parent := rgnconsequencematrix;
  lblauxcrit[j].Left := 140 + 54 * (j);
  lblauxcrit[j].Top :=  600;
  lblauxcrit[j].Caption := CritSiglas[j];
  lblauxcrit[j].Name := 'lblcrit' + inttostr(j+1);
  lblauxcrit[j].Font.FontFamily := 'Verdana, Helvetica, Sans-Serif';
  end;
  for j := 0 to Colunas - 1 do
  begin
  //edtpesos[0,j].Free;
  TIWUserSession(WebApplication.Data).edtpesos[0,j] := TIWEdit.Create(rgnconsequencematrix);
  edtpesos[0,j].Parent := rgnconsequencematrix;
  edtpesos[0,j].Height := 21;
  edtpesos[0,j].Width := 54;
  edtpesos[0,j].left := 140 + 54 * j;
  edtpesos[0,j].Top := 620;
  edtpesos[0,j].caption := FloattoStr(LimPesoMax[j]);
  edtpesos[0,j].name := 'edtlscrit' + inttostr(1+j);


  //edtpesos[1,j].Free;
  TIWUserSession(WebApplication.Data).edtpesos[1,j] := TIWEdit.Create(rgnconsequencematrix);
  edtpesos[1,j].Parent := rgnconsequencematrix;
  edtpesos[1,j].Height := 21;
  edtpesos[1,j].Width := 54;
  edtpesos[1,j].left := 140 + 54 * j;
  edtpesos[1,j].Top := 660;
  edtpesos[1,j].caption := FloattoStr(LimPesoMin[j]);
  edtpesos[1,j].name := 'edtlicrit' + inttostr(1+j);
  end;







  for i := 0 to Linhas - 1 do
  begin
  //inc(auxmanual);
      //lblalt[i].Free;
      lblalt[i] := TIWLabel.Create(Regiao);
      lblalt[i].Parent := Regiao;
      lblalt[i].left := 20;
      lblalt[i].Top := 35 + 21 * i;
      lblalt[i].Caption := AltSiglas[i];
      lblalt[i].Font.FontFamily := 'Verdana, Helvetica, Sans-Serif';
      lblalt[i].Name := 'lblCelColuna' + IntToStr(i + 1);
  end;

  auxi := TIWUserSession(WebApplication.Data).auxalt+1;
  auxj := TIWUserSession(WebApplication.Data).auxcrit+1;



 if (auxi <> 1) or (auxj <> 1) then
  begin
  for i := 0 to TIWUserSession(WebApplication.Data).auxalt - 1 do
  begin
    for j := 0 to TIWUserSession(WebApplication.Data).auxcrit - 1 do
    begin
    edt[i,j].Free;
    end;
    end;
  end;

  SetLength(edt, Numalt+1, Numcrit+1);
  SetLength(TIWUserSession(WebApplication.Data).ConsMatrix, TIWUserSession(WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data).NumCrit);



  for j := 0 to Colunas - 1 do
  begin
    for i := 0 to Linhas - 1 do
    begin
      TIWUserSession(WebApplication.Data).edt[i,j] := TIWEdit.Create(Regiao);
      //edt.Name := 'edtCel' + IntToStr(i) + 'x' + IntToStr(j);
      edt[i,j].Height := 21;
      edt[i,j].Width := 54;
      edt[i,j].left := 95 + 54 * j;
      edt[i,j].Top := 32 + 21 * i;
      edt[i,j].Parent := Regiao;
      edt[i,j].text := floattostr(TIWUserSession(WebApplication.Data).ConsMatrix[i,j]);
      inc(auxj);
      {if btnbacktoSA.Visible = true then
        edt.Enabled := False;}

      // Recuperando o valor

    end;// for i
    {if auxalt = linhas then begin
    i := auxalt-1;
    edt[i,j].Height := 21;
    edt[i,j].Width := 54;
    edt[i,j].left := 85 + 54 * j;
    edt[i,j].Top := 32 + 21 * i;
    edt[i,j].Parent := Regiao;
    edt[i,j].text := floattostr(TIWUserSession(WebApplication.Data).ConsMatrix[i,j]);
    inc(auxj);
     }

    inc(auxi);
  end; // for j

  {lstbxlegendacriterios.Clear;
  lstbxlegendaalternativas.Clear;

  for i := 0 to Numcrit - 1 do
  begin
    lstbxlegendacriterios.Items.Add(CritSiglas[i]+ ': ' + NomeCrit[i])
  end;

  for j := 0 to NumAlt - 1 do
  begin
    lstbxlegendaalternativas.Items.Add(AltSiglas[j]+ ': ' + NomeAlt[j])
  end;
   }
  auxcrit := NumCrit;
  auxalt := NumAlt;
  end;
end;

procedure TIWForm1.btnAlterarClick(Sender: TObject);
var
i, j : integer;
begin
rgnConsequenceMatrix.Visible := False;


if lstbxcriterio.Items.Count = 0 then begin
 with TIWUserSession(WebApplication.Data) do
  begin
    for i := 0 to NumCrit - 1 do
    lstbxcriterio.Items.Add(NomeCrit[i]);
    for j := 0 to NumAlt - 1 do
    lstbxalternativa.Items.Add(NomeAlt[j]);
    auxalt := Numalt;
    auxcrit := Numcrit;
  end;
end;
rgnManualProblem.Visible := True;
end;

procedure TIWForm1.Definir_Decimais;
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

procedure TIWForm1.Siglas;
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

procedure TIWForm1.Inicializa;
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

  //Checar_Siglas;
  TIWUserSession(WebApplication.Data).Epson := Power(10, -8);
  // ** Definindo limiar para as restrições de Potencial Otimalidade **//
  TIWUserSession(WebApplication.Data).Epson := Power(10, -8);
  Setlength(TIWUserSession(WebApplication.Data).Bconseqord, TIWUserSession
      (WebApplication.Data).NumCrit);
  Setlength(TIWUserSession(WebApplication.Data).Wconseqord, TIWUserSession
      (WebApplication.Data).NumCrit);

  if TIWUserSession(WebApplication.Data).Problematica = 2 then
  begin
    //edtMaxdif.Text := '0';
    //edtMaxdif.Enabled := false;
    //CostNorm;
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
      Setlength(LimMax, NumCrit - 1);
      Setlength(LimMin, NumCrit - 1);
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

procedure TIWForm1.btncontinueClick(Sender: TObject);
var
i : integer;
begin
With TIWUserSession(WebApplication.Data) do
begin
  problematica := 1;
  idioma := 1;
  id_problema := '0';
  ID_user := '0';
  reiniciar := false;
  stage := 2;
  inicializa;
  normaliza;
  sol := 10;
  epsilon := 0.01;
  //ClpSolve := True;
  //Ciclo := 1;
  Setlength(vtrordem, numcrit);
      Setlength(vtrblin, numcrit);
      Setlength(vtrrelacoesblin, numcrit - 1);
      Setlength(vtrposicoes, numcrit);
      Setlength(Representantes, numcrit);
      SetLength(Cri_Inteiro, numcrit);
      for i := 0 to numcrit - 1 do
      begin
        Representantes[i] := i;
        vtrordem[i] := i;
        vtrposicoes[i] := i + 1;
        vtrblin[i] := false;
        Cri_Inteiro[i] := False;
        {lstbxtco.Items.Add(TIWUserSession(WebApplication.Data)
            .CritSiglas[vtrordem[i]] + ': ' + NomeCrit[vtrordem[i]]);}
        // Round:=0;
      end;
      for i := 0 to numcrit - 2 do
        vtrrelacoesblin[i] := false;


  Tfrmaintercriterio.Create(WebApplication).BtnshowresultsClick(Self);
  //TfrmResultados.Create(WebApplication).rgnHEQ0.Visible := False;
end;
end;

procedure TIWForm1.btndeletealternativaClick(Sender: TObject);
var
  Id_Alternativa, i,j, k: integer;
begin
    if edtnomealternativa.Text = '' then
  begin
    WebApplication.ShowMessage(
      'Preencha corretamente os dados para excluir a alternativa.');
  end
  else
  begin
    With TIWUserSession(WebApplication.Data) do
    begin
    for j := 0 to NumAlt - 1 do begin
       if NomeAlt[j]<> edtnomealternativa.Text  then
       continue
       else
       begin
      SetLength(TIWUserSession(WebApplication.Data).ConsMatrix, TIWUserSession(WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data).NumCrit);
      for i := lstbxalternativa.ItemIndex + 1 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
      begin
      NomeAlt[i-1] := NomeAlt[i];
      for k := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
        TIWUserSession(WebApplication.Data).ConsMatrix[i-1,k] := ConsMatrix[i,k]
      end; //for i
      ///rgnManualProblemCreate(Self);
      NumAlt := NumAlt - 1;
      SetLength(TIWUserSession(WebApplication.Data).ConsMatrix, TIWUserSession(WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data).NumCrit);
      SetLength(NomeAlt, NumAlt);
      lstbxalternativa.Clear;

      for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
      lstbxalternativa.Items.Add(NomeAlt[i]);
      edtnomealternativa.Clear;
      rgnManualProblemCreate(Self);
      lstbxalternativa.Refresh;
      break;
      end; // else
    end;//for j

    end;// with
  end;// else
end;

procedure TIWForm1.Normaliza;
/// Esse procedimento realiza a normalização das consequências considerando
/// o tipo de função valor. E se o critério em qurestão passou ou não pela avaliação IntraCrit
var
  I, J: integer;
  Zmax, Zmin: Double;
  matriztestesnormal: array of array of Double;

begin

  with TIWUserSession(WebApplication.Data) do
  begin
    Setlength(TIWUserSession(WebApplication.Data).ConsNorm, TIWUserSession
        (WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data)
        .Numcrit);
    // Setlength(VtrcritBisec, Numcrit);   verificar se após comentario vai dar treta

    for J := 0 to Numcrit - 1 do
      for I := 0 to NumAlt - 1 do

        if TIWUserSession(WebApplication.Data).VtrcritBisec[J] = true then
        // criterio passou pela interpolação para transformar escala
        begin
          continue;
        end

        else // false indica que crit não foi elicitado, entra na normaliza
        begin

          if (TIWUserSession(WebApplication.Data).Niveis[J] = 2) or
            (TIWUserSession(WebApplication.Data).Niveis[J] = 6) or
            (TIWUserSession(WebApplication.Data).Niveis[J] = 7) then
          begin
            TIWUserSession(WebApplication.Data).Ftipo[J] := 1;
          end;

          /// verificar se os criterios DISCRETOS devem seguir uma regra diferente de transformação da escala///
          case TIWUserSession(WebApplication.Data).Ftipo[J] of // função que armazena o vetor dos tipos de FO declaradas na planilha

            1: // procedimento realiza a normalização das consequências considerando a função valor linear ///////////////////////
              begin
                // Normalização para os critérios de maximização //////////////////////
                if (TIWUserSession(WebApplication.Data).Typeofcrit[J] = 1) or
                  (Typeofcrit[J] = 3) then
                  TIWUserSession(WebApplication.Data).ConsNorm[I, J] :=
                    (ConsMatrix[I, J] - Min[J]) / (Max[J] - Min[J]);
                /// Normalização para os critérios de minimização //////////////////////
                if (Typeofcrit[J] = 0) or (Typeofcrit[J] = 2) then
                  TIWUserSession(WebApplication.Data).ConsNorm[I, J] :=
                    (ConsMatrix[I, J] - Max[J]) / (-Max[J] + Min[J]);

              end;

            2: // procedimento realiza a normalização das consequencias considerando função valor exponencial
              begin
                if Min[J] = 0 then
                  Min[J] := Power(10, -5); // O QUE É POWER??
                Zmin := Exp(TIWUserSession(WebApplication.Data)
                    .parA[J] * Min[J]); // exponencial do parametro A vezes o valor minimo capturado no criterio
                Zmax := Exp(parA[J] * Max[J]); // exponencial do parametro A vezes o valor máximo capturado no criterio

                // normalização para os criterios de maximização//
                if (Typeofcrit[J] = 1) or (Typeofcrit[J] = 3) then
                  ConsNorm[I, J] :=
                    (Exp(parA[J] * TIWUserSession(WebApplication.Data)
                        .ConsMatrix[I, J]) - Zmin) / (Zmax - Zmin);
                // normalização para os criterios de minimização//
                if (Typeofcrit[J] = 0) or (Typeofcrit[J] = 2) then
                  ConsNorm[I, J] := (Exp(parA[J] * ConsMatrix[I, J]) - Zmax) /
                    (Zmin - Zmax);
              end;

            3: // procedimento realiza a normalização das consequencias considerando função valor logaritmica
              begin
                if Min[J] = 0 then
                  Min[J] := Power(10, -5); // O QUE É POWER??
                Zmin := Ln(parA[J] * Min[J] + 1 - parA[J] * Min[J]);
                Zmax := Ln(parA[J] * Max[J] + 1 - parA[J] * Min[J]);

                // normalização para os criterios de maximização///
                if (Typeofcrit[J] = 1) or (Typeofcrit[J] = 3) then
                  ConsNorm[I, J] :=
                    (Ln(parA[J] * ConsMatrix[I, J] + 1 - parA[J] * Min[J])
                      - Zmin) / (Zmax - Zmin);
                // normalização para os criterios de minimização//
                if (Typeofcrit[J] = 0) or (Typeofcrit[J] = 2) then
                  ConsNorm[I, J] :=
                    (Ln(parA[J] * ConsMatrix[I, J] + 1 - parA[J] * Min[J])
                      - Zmax) / (Zmin - Zmax);
                { else
                  ConsNorm[i, j] := 1; }
              end;

            4: // procedimento realiza a normalização das consequencias considerando função valor logistica (S-Shape)
              begin
                if Min[J] = 0 then
                  Min[J] := Power(10, -5);

                Zmax := Exp(-parA[J] / Max[J]);
                Zmin := Exp(-parA[J] / Min[J]);

                // normalização para os criterios de maximização///
                if (Typeofcrit[J] = 1) or (Typeofcrit[J] = 3) then
                  ConsNorm[I, J] := (Exp(-parA[J] / ConsMatrix[I, J]) - Zmin) /
                    (Zmax - Zmin);
                // normalização para os criterios de minimização//
                if (Typeofcrit[J] = 0) or (Typeofcrit[J] = 2) then
                  TIWUserSession(WebApplication.Data).ConsNorm[I, J] :=
                    (Exp(-parA[J] / ConsMatrix[I, J]) - Zmax) / (Zmin - Zmax);
                { else
                  ConsNorm[i, j] := 1; }
              end;
          end;

        end; // final dos contadores

    Setlength(matriztestesnormal, NumAlt, Numcrit);
    for I := 0 to NumAlt - 1 do
      for J := 0 to Numcrit - 1 do
        matriztestesnormal[I, J] := TIWUserSession(WebApplication.Data)
          .ConsNorm[I, J];

  end;
  // final dos casos

  //SalvarBdIntra(1); // chama caso para atualizar ConsNorm dos crits que não passaram pela interpolação

end;

procedure TIWForm1.rgnManualProblemCreate(Sender: TObject);
begin
if lstbxcriterio.Items.Count = 0 then
begin
  btndeletecriterio.Enabled := False;
  btnupdatecriterio.enabled := false;
end;
if lstbxalternativa.Items.Count = 0 then
begin
 btndeletealternativa.Enabled := False;
 btnupdatealternativa.enabled := false;
end;

end;

procedure TIWForm1.btndeletecriterioClick(Sender: TObject);
var
  ID_criterio, i,j, k: integer;
begin

  // btnsaveelicitationClick(Sender);
  if ((edtnomecriterio.Text = '') or
      (rdgrpdirecaocriterio.ItemIndex = -1) or (TIWUserSession(WebApplication.Data).NumCrit = 0)) then
  begin
    WebApplication.ShowMessage(
      'Preencha corretamente os dados para excluir o critério.');
  end
  else
  begin
    With TIWUserSession(WebApplication.Data) do
    begin
      for j := 0 to NumCrit - 1 do
      begin
        if NomeCrit[j]<> edtnomecriterio.Text then
        continue
        else
        begin
        SetLength(TIWUserSession(WebApplication.Data).ConsMatrix, TIWUserSession(WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data).NumCrit);
        for i := lstbxcriterio.ItemIndex + 1 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
        begin
        TIWUserSession(WebApplication.Data).NomeCrit[i-1] := NomeCrit[i];
        //CritSigla[i-1] := CritSigla[i];
        Typeofcrit[i-1] := Typeofcrit[i];
        for k := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
        ConsMatrix[k,i-1] := ConsMatrix[k,i]
        end; //for i
        //rgnManualProblemCreate(Self);
        NumCrit := NumCrit - 1;
        SetLength(TIWUserSession(WebApplication.Data).ConsMatrix, TIWUserSession(WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data).NumCrit);
        setlength(TIWUserSession(WebApplication.Data).NomeCrit, TIWUserSession(WebApplication.Data).NumCrit);
        //setlength(TIWUserSession(WebApplication.Data).CritSigla, TIWUserSession(WebApplication.Data).NumCrit);
        SetLength(TIWUserSession(WebApplication.Data).ftipo, TIWUserSession(WebApplication.Data).NumCrit);
        SetLength(TIWUserSession(WebApplication.Data).Niveis, TIWUserSession(WebApplication.Data).NumCrit);
        SetLength(TIWUserSession(WebApplication.Data).parA, TIWUserSession(WebApplication.Data).NumCrit);
        SetLength(TIWUserSession(WebApplication.Data).parB, TIWUserSession(WebApplication.Data).NumCrit);
        SetLength(TIWUserSession(WebApplication.Data).Typeofcrit, TIWUserSession(WebApplication.Data).NumCrit);

        lstbxcriterio.Clear;
        for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
        Lstbxcriterio.Items.Add(NomeCrit[i]);

        edtnomecriterio.Clear;
        rdgrpdirecaocriterio.ItemIndex := -1;
        lstbxcriterio.Refresh;
        rgnManualProblemCreate(Self);
        break;
        end;
        end;
        end;



  end; //with


end;

procedure TIWForm1.btngoinput3Click(Sender: TObject);
var
  i, j,k: integer;
  cred: Boolean;
  ID_problema, qntprob: integer;
  nameprob: string;
  vetchar: array [0 .. 9] of string;
begin
  // WebApplication.ShowMessage('New tool available soon', smAlert);

  if (TIWUserSession(WebApplication.Data).NumAlt = 0) or
    (TIWUserSession(WebApplication.Data).NumCrit = 0) then
  begin
    WebApplication.ShowMessage(
      'Você deve definir corretamente o número de alternativas e critérios para prosseguir.');
  end
  else
  begin
    CriarString(TIWUserSession(WebApplication.Data).NumAlt, TIWUserSession
        (WebApplication.Data).NumCrit, rgnDisplaymatrizconsequencia);
    rgnConsequenceMatrix.Visible := true;
    rgnManualProblem.Visible := False;

    //TiwForm1.Refresh;
  end;
end;

procedure TIWForm1.btnImportarClick(Sender: TObject);
begin
rgnimport.Visible := True;
btnimportar.Visible := False;
btninputmanual.Visible := False;
iwregion3.Visible := False;
end;

procedure TiwForm1.ImportExcel;
var
  i, j, idalt, idcrit, r, s, k, a: integer;
  x, y : string;
  ID_problema, qntprob: integer;
  nameprob, str, auxranking: string;

begin

// Limpando os elementos
With TIWUserSession(WebApplication.Data) do
begin
  NumAlt := 0;
  NumCrit := 0;
end;

  FileArquivo.SaveToFile('C:\Users\Public\Documents' + (FileArquivo.Filename));
  TIWAdvWebGridExcelIO1.XLSImport('C:\Users\Public\Documents' +
      (FileArquivo.Filename));

  TIWAdvWebGridExcelIO1.AdvWebGrid := grdImport;

  TIWUserSession(WebApplication.Data).vQdeColunas := grdImport.Columns.Count;
  // número total de colunas preenchidas
  TIWUserSession(WebApplication.Data).vQdeLinhas := grdImport.TotalRows;
  // número total de linhas preenchidas


   for i := 1 to TIWUserSession(WebApplication.Data).vQdeColunas - 1 do
   begin
   x := grdImport.CELLS[i, 0];
   if x <> '' then
   inc(TIWUserSession(WebApplication.Data).NumCrit)
   else
   break;
   end;

   for j := 7 to TIWUserSession(WebApplication.Data).vQdeLinhas - 1 do
   begin
   y := grdImport.CELLS[0, j];
   if y <> '' then begin
   inc(TIWUserSession(WebApplication.Data).NumAlt);
   SetLength(TIWUserSession(WebApplication.Data).NomeAlt, TIWUserSession(WebApplication.Data).NumAlt);
   TIWUserSession(WebApplication.Data).NomeAlt[TIWUserSession(WebApplication.Data).NumAlt-1] := y;
   end
   else
   break;
   end;

   for i := 1 to TIWUserSession(WebApplication.Data).NumCrit do
   begin
     x := grdimport.Cells[i,1];
     {if (x = '2') or (x = '4') then
     grdimport.Cells[i,1] := '0';

     if (x = '3') or (x = '5') then
     grdimport.Cells[i,1] := '1';
      }
     if x = '' then begin
     WebApplication.ShowMessage(
      'Please, fill the mandatory data in the file to continue your problem.');
        exit;
     end;
   end;

   for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
   begin
     for j := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
     begin
       x := grdimport.Cells[i+1,j+7];
       if x = '' then begin
        WebApplication.ShowMessage(
          'Please, fill the mandatory data in the file to continue your problem.');
        exit;
        end;
     end;
   end;

  //Setando os elementos dos criterios
   SetLength(TIWUserSession(WebApplication.Data).NomeCrit, TIWUserSession(WebApplication.Data).NumCrit);
   //SetLength(TIWUserSession(WebApplication.Data).CritSigla, TIWUserSession(WebApplication.Data).NumCrit);
   SetLength(TIWUserSession(WebApplication.Data).ftipo, TIWUserSession(WebApplication.Data).NumCrit);
   SetLength(TIWUserSession(WebApplication.Data).Niveis, TIWUserSession(WebApplication.Data).NumCrit);
   SetLength(TIWUserSession(WebApplication.Data).parA, TIWUserSession(WebApplication.Data).NumCrit);
   SetLength(TIWUserSession(WebApplication.Data).parB, TIWUserSession(WebApplication.Data).NumCrit);
   SetLength(TIWUserSession(WebApplication.Data).Typeofcrit, TIWUserSession(WebApplication.Data).NumCrit);
   SetLength(TIWUserSession(WebApplication.Data).LimPesoMax, TIWUserSession(WebApplication.Data).NumCrit);
   SetLength(TIWUserSession(WebApplication.Data).LimPesoMin, TIWUserSession(WebApplication.Data).NumCrit);
   SetLength(TIWUserSession(WebApplication.Data).ranking, TIWUserSession(WebApplication.Data).NumAlt);
   //Setlength(TIWUserSession(WebApplication.Data).niveisdisc, grdImportproblema.Columns.Count - 1);
   SetLength(TIWUserSession(WebApplication.Data).NomeAlt, TIWUserSession(WebApplication.Data).NumAlt);
   //SetLength(TIWUserSession(WebApplication.Data).AltSigla, TIWUserSession(WebApplication.Data).NumAlt);
   //Setando a matriz de consequencias
   SetLength(TIWUserSession(WebApplication.Data).ConsMatrix, TIWUserSession(WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data).NumCrit);


   With TIWUserSession(WebApplication.Data) do
   begin
   for i := 1 to TIWUserSession(WebApplication.Data).NumCrit do
   begin
   NomeCrit[i-1] := grdImport.CELLS[i, 0];
   //CritSigla[i-1] := 'Crit'  + inttostr(i);
   //lstbxcriterio.Items.Add(NomeCrit[i-1]);
   ftipo[i-1] := 1;
   TIWUserSession(WebApplication.Data).Typeofcrit[i-1] := strtoint(grdImport.CELLS[i, 1]);
   if (Typeofcrit[i-1] = 2) or (Typeofcrit[i-1] = 3) then
   Niveis[i-1] := strtoint(grdImport.CELLS[i, 5])
   else
   Niveis[i-1] := 0;
   parA[i-1] := 0;
   parB[i-1] := 0;
   //inc(idcrit);

   end; // for i

   {for j := 0 to TIWUserSession(WebApplication.Data).NumAlt do
    begin
      TIWUserSession(WebApplication.Data).NomeAlt[j] := grdImport.CELLS[0, j + 8];
      //AltSigla[j-1] := 'Alt' + inttostr(j);
      //lstbxalternativa.Items.Add(NomeAlt[i-1]);
    end; }



    for i := 1 to TIWUserSession(WebApplication.Data).NumCrit do
    begin
      for j := 1 to TIWUserSession(WebApplication.Data).NumAlt do
      begin
      TIWUserSession(WebApplication.Data).ConsMatrix[j-1,i-1] := strtofloat((grdImport.CELLS[i, j + 6]));
      end;
    end;
   end; // with

    for i := 1 to TIWUserSession(WebApplication.Data).NumCrit do
    begin
     TIWUserSession(WebApplication.Data).LimPesoMax[i-1] := strtofloat((grdImport.CELLS[i, 11+TIWUserSession(WebApplication.Data).NumAlt]));
     TIWUserSession(WebApplication.Data).LimPesoMin[i-1] := strtofloat((grdImport.CELLS[i, 12+TIWUserSession(WebApplication.Data).NumAlt]));
    end;



    With TIWUserSession(WebApplication.Data) do
    begin
    for j := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
    begin
    str := grdImport.CELLS[TIWUserSession(WebApplication.Data).Numcrit+3, j+1];
    if str <> '' then
    inc(NumAH);
    end;

    for j := 0 to TIWUserSession(WebApplication.Data).NumAH - 1 do
    begin
    auxranking := grdImport.CELLS[TIWUserSession(WebApplication.Data).Numcrit+3, j+1];
    for i := 1 to length(auxranking) do
      begin

      if auxranking[i] = '[' then
       str := '';

      if (auxranking[i] <> ',') and (auxranking[i] <> ']') and (auxranking[i] <> '[') then
      str :=  str + auxranking[i];

       if (auxranking[i] = ',') or (auxranking[i] = ']') then
       begin
         for k := 0 to NumAlt - 1 do
         begin
           if str = NomeAlt[k] then
           TIWUserSession(WebApplication.Data).ranking[k] := j+1;
         end; // for k
       end; // if
      end; // for i
    end; // for j


    {SetLength(MatAH, NumAH-1, 3);
    for i := 0 to NumAlt - 1 do
    begin
      for j := 0 to NumAlt - 1 do
        begin
          if ranking[i] = ranking[j] + 1 then
          begin
          MatAH[a,0] := 2;
          MatAH[a,1] := j;
          TIWUserSession(WebApplication.Data).MatAH[a,2] := i;
          inc(a);
          end;
        end;
    end;  }


    end; // with





    rgnimport.Visible := false;
    rgnConsequenceMatrix.Visible := True;
    rgnintro.Visible := False;
    btnimportar.Visible := true;
    btninputmanual.Visible := true;
    iwregion3.Visible := true;



   TIWUserSession(WebApplication.Data).inputmanual := false;
   CriarString(TIWUserSession(WebApplication.Data).NumAlt, TIWUserSession
        (WebApplication.Data).NumCrit, rgnDisplaymatrizconsequencia);

//DeleteFile('C:\Users\CDSID\Desktop\Pedro - CDSID\SW Secretaria de Defesa Social (25.05)\SW Secretaria de Defesa Social\' + (FileArquivo.Filename)+ '.xls');
grdimport.CleanupInstance;
end;


procedure TIWForm1.btnimportClick(Sender: TObject);
var
  i,j: integer;
  cred: Boolean;
  vetchar: array [0 .. 9] of string;
begin
    if (FileArquivo.Filename = '') then
    begin
      if TIWUserSession(WebApplication.Data).excelverif = False then
        WebApplication.ShowMessage('Você deve escolher um arquivo de Excel!');
      exit;
    end
    else if FileArquivo.Filename <> '' then
    begin
        ImportExcel;
    end;

end;

procedure TIWForm1.btnInputManualClick(Sender: TObject);
begin
rgnintro.visible := False;
rgnManualProblem.Visible := True;
end;

procedure TIWForm1.btnnewalternativaClick(Sender: TObject);
var
  i: integer;
  Texto: string;
begin
  if edtnomealternativa.Text = '' then
  begin
    WebApplication.ShowMessage(
      'Please fill the mandatory data to insert a new alternative.');
  end
  else
  begin
    Texto := edtnomealternativa.Text;
    for i := 0 to TIWUserSession(WebApplication.Data).NumAlt - 1 do
    begin
      if Texto = lstbxalternativa.Items[i] then
      begin
        WebApplication.ShowMessage(
          'Please, insert a new code/ name for the new alternative.');
        Texto := '';
        break;
      end;
    end;
     if Texto <> '' then
     begin
     With TIWUserSession(WebApplication.Data) do
     begin
       btndeletealternativa.Enabled := True;
       btnupdatealternativa.Enabled := True;
       inc(NumAlt);
       setlength(NomeAlt, NumAlt);
       //setlength(AltSigla, NumAlt);
       NomeAlt[NumAlt-1] := edtnomealternativa.Text;
       //AltSigla[NumAlt-1] := 'Alt' + inttostr(NumAlt);
       lstbxalternativa.Items.Add(edtnomealternativa.Text);
     end;
     end;
    edtnomealternativa.Clear;
    lstbxalternativa.Refresh;
  end;
end;


procedure TIWForm1.btnnewcriterioClick(Sender: TObject);
var
  i,j: integer;
  edt: TIWEdit;
  lbl: TIWLabel;
  Texto: string;
begin
  if ((edtnomecriterio.Text = '') or (rdgrpdirecaocriterio.ItemIndex = -1)) then
  begin
    WebApplication.ShowMessage(
      'Please, fill the mandatory data to insert a new criterion.');
  end
  else
  begin
    Texto := edtnomecriterio.Text;
    for i := 0 to TIWUserSession(WebApplication.Data).NumCrit - 1 do
    begin
      if Texto = lstbxcriterio.Items[i] then
      begin
        WebApplication.ShowMessage(
          'Please, insert a new code/ name for the new criterion.');
          Texto := '';
        break;
      end;
    end;
    if texto <> '' then
    begin
    btndeletecriterio.Enabled := True;
    btnupdatecriterio.Enabled := true;
    inc(TIWUserSession(WebApplication.Data).NumCrit);
    setlength(TIWUserSession(WebApplication.Data).NomeCrit, TIWUserSession(WebApplication.Data).NumCrit);
    SetLength(TIWUserSession(WebApplication.Data).ftipo, TIWUserSession(WebApplication.Data).NumCrit);
    SetLength(TIWUserSession(WebApplication.Data).Niveis, TIWUserSession(WebApplication.Data).NumCrit);
    SetLength(TIWUserSession(WebApplication.Data).parA, TIWUserSession(WebApplication.Data).NumCrit);
    SetLength(TIWUserSession(WebApplication.Data).parB, TIWUserSession(WebApplication.Data).NumCrit);
    SetLength(TIWUserSession(WebApplication.Data).Typeofcrit, TIWUserSession(WebApplication.Data).NumCrit);
    SetLength(TIWUserSession(WebApplication.Data).LimPesoMax, TIWUserSession(WebApplication.Data).NumCrit);
   SetLength(TIWUserSession(WebApplication.Data).LimPesoMin, TIWUserSession(WebApplication.Data).NumCrit);
    With TIWUserSession(WebApplication.Data) do
    begin
    NomeCrit[NumCrit-1] := edtnomecriterio.Text;
    lstbxcriterio.Items.Add(edtnomecriterio.Text);
    ftipo[NumCrit-1] := 1;
    TIWUserSession(WebApplication.Data).Typeofcrit[NumCrit-1] := rdgrpdirecaocriterio.ItemIndex;
    SetLength(TIWUserSession(WebApplication.Data).ConsMatrix, TIWUserSession(WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data).NumCrit);


    end;
    end;



    edtnomecriterio.Clear;
    rdgrpdirecaocriterio.ItemIndex := -1;
    //btngoinput2.OnClick(Sender);
  end;
end;

procedure TIWForm1.btnsubmitClick(Sender: TObject);
var
i, j : integer;
begin

  With TIWUserSession(WebApplication.Data) do
  begin
    for j := 0 to Numcrit - 1 do
      for i := 0 to Numalt - 1 do begin
        consmatrix[i,j] := strtofloat((edt[i,j].caption));
        edt[i,j].Enabled := False;
      end;
      btnsalvarpesos.Enabled := True;
      btnsubmit.enabled := False;
      btnalterar.Enabled := False;
      Siglas;

      WebApplication.ShowMessage(
      'Os dados foram coletados com sucesso!');




  end;
end;

procedure TIWForm1.btnupdatealternativaClick(Sender: TObject);
var
  Id_Alternativa, i: integer;
  Texto: string;
begin
  if edtnomealternativa.Text = '' then
  begin
    WebApplication.ShowMessage(
      'Preencha corretamente os dados para atualizar a alternativa.');
  end
  else
  begin
    With TIWUserSession(WebApplication.Data) do
    begin
    for i := 0 to NumAlt - 1 do begin
       if NomeAlt[i]<> edtnomealternativa.Text  then
       continue
       else
       begin
       //rgnManualProblemCreate(Self);
        NomeAlt[lstbxalternativa.ItemIndex] := edtnomealternativa.Text;
        lstbxalternativa.Items[lstbxalternativa.ItemIndex] := edtnomealternativa.text;
        lstbxalternativa.ItemIndex := -1;
        edtnomealternativa.Clear;
        lstbxalternativa.Refresh;
        rgnManualProblemCreate(Self);
        break;
       end;
    end;
    end;
    WebApplication.ShowMessage(
      'Please fill the mandatory data to update this alternative.');


  end;
end;

procedure TIWForm1.btnupdatecriterioClick(Sender: TObject);
var
  ID_criterio, i: integer;
  Texto: string;
begin
  // btnsaveelicitationClick(Sender);
  if ((edtnomecriterio.Text = '') or
      (rdgrpdirecaocriterio.ItemIndex = -1) or (TIWUserSession(WebApplication.Data).NumCrit = 0)) then
  begin
    WebApplication.ShowMessage(
      'Preencha corretamente os dados para atualizar o critério');
  end
  else
  begin


    With TIWUserSession(WebApplication.Data) do
    begin
      for i := 0 to NumCrit - 1 do
      begin
        if NomeCrit[i]<> edtnomecriterio.Text then
        continue
        else
        begin
        //rgnManualProblemCreate(Self);
        NomeCrit[lstbxcriterio.ItemIndex] := edtnomecriterio.Text;
        lstbxcriterio.Items[lstbxcriterio.ItemIndex] := edtnomecriterio.text;
        Typeofcrit[lstbxcriterio.ItemIndex] := rdgrpdirecaocriterio.ItemIndex;
        lstbxcriterio.ItemIndex := -1;
      edtnomecriterio.Clear;
      rdgrpdirecaocriterio.ItemIndex := -1;
      //btngoinput2.OnClick(Sender);
      lstbxcriterio.Refresh;
      rgnManualProblemCreate(Self);
        break;
        end;
      end;
    end;

  end;
end;


procedure TIWForm1.IWAppFormCreate(Sender: TObject);
begin
rgnintro.Align := AlClient;
rgnManualProblem.Align := AlClient;
rgnConsequenceMatrix.Align := AlClient;
rgnintro.Visible := True;
With TIWUserSession(WebApplication.Data) do
begin
Setlength(vtrordem, 0);
Setlength(vtrblin, 0);
Setlength(vtrrelacoesblin, 0);
Setlength(vtrposicoes, 0);
Setlength(Representantes, 0);
SetLength(Cri_Inteiro, 0);
NumAlt := 0;
Numcrit := 0;
SetLength(TIWUserSession(WebApplication.Data).ConsMatrix, TIWUserSession(WebApplication.Data).NumAlt, TIWUserSession(WebApplication.Data).NumCrit);
setlength(TIWUserSession(WebApplication.Data).NomeCrit, TIWUserSession(WebApplication.Data).NumCrit);
SetLength(TIWUserSession(WebApplication.Data).ftipo, TIWUserSession(WebApplication.Data).NumCrit);
SetLength(TIWUserSession(WebApplication.Data).Niveis, TIWUserSession(WebApplication.Data).NumCrit);
SetLength(TIWUserSession(WebApplication.Data).parA, TIWUserSession(WebApplication.Data).NumCrit);
SetLength(TIWUserSession(WebApplication.Data).parB, TIWUserSession(WebApplication.Data).NumCrit);
SetLength(TIWUserSession(WebApplication.Data).Typeofcrit, TIWUserSession(WebApplication.Data).NumCrit);
SetLength(TIWUserSession(WebApplication.Data).LimPesoMax, TIWUserSession(WebApplication.Data).NumCrit);
SetLength(TIWUserSession(WebApplication.Data).LimPesoMin, TIWUserSession(WebApplication.Data).NumCrit);
end;
end;

procedure TIWForm1.btnsalvarpesosClick(Sender: TObject);
var
i, j : integer;
auxpesomax, auxpesomin : double;
begin
With TIWUserSession(WebApplication.Data) do
  begin
      auxpesomax := 0;
      auxpesomin := 0;
      for j := 0 to NumCrit - 1 do begin
       LimPesoMin[j] := StrToFloat((edtpesos[1,j].Caption));
       LimPesoMax[j] := StrToFloat((edtpesos[0,j].Caption));
       edtpesos[1,j].Enabled := false;
       edtpesos[0,j].Enabled := False;
      end;


      btnsalvarpesos.Enabled := False;
      btnContinue.Enabled := True;


      WebApplication.ShowMessage(
      'Os pesos foram salvos com sucesso!');




  end;
end;

procedure TIWForm1.IWLink1Click(Sender: TObject);
begin
rgnManualProblem.Visible := False;
rgnintro.Visible := True;
IWAppFormCreate(self);
end;

procedure TIWForm1.IWLink2Click(Sender: TObject);
begin
rgnConsequenceMatrix.Visible := False;
rgnintro.Visible := True;
IWAppFormCreate(Self);
end;

procedure TIWForm1.IWLink4Click(Sender: TObject);
begin
rgnimport.Visible := False;
btnimportar.Visible := True;
btninputmanual.Visible := True;
iwregion3.Visible := True;
end;

procedure TIWForm1.lstbxalternativaChange(Sender: TObject);
var
  Id_Alternativa: integer;
begin

  with TIWUserSession(WebApplication.Data) do
  begin

    edtnomealternativa.Text := NomeAlt[lstbxalternativa.itemindex];
    //edtsiglaalternativa.Text := FieldByName('sigla_alternativa').AsString;
  end;
end;

procedure TIWForm1.lstbxcriterioChange(Sender: TObject);
var
  ID_criterio: integer;
begin

    With TIWUserSession(WebApplication.Data) do
    begin
    edtnomecriterio.Text := NomeCrit[lstbxcriterio.ItemIndex];
    //edtsiglacriterio.Text := FieldByName('sigla_criterio').AsString;

    if Typeofcrit[lstbxcriterio.ItemIndex] = 1 then
      rdgrpdirecaocriterio.ItemIndex := 1
    else
      rdgrpdirecaocriterio.ItemIndex := 0;
    // edtdescricaocriterio.Text := FieldByName('descricao_criterio').AsString;

    end; // With

end;

initialization
  TIWForm1.SetAsMainForm;

end.
