unit UserSessionUnit;

{
  This is a DataModule where you can add components or declare fields that are specific to 
  ONE user. Instead of creating global variables, it is better to use this datamodule. You can then
  access the it using UserSession.
}
interface

uses
  IWUserSessionBase, SysUtils, Classes, DB, DBClient, WebAdapt, WebComp,
  ZConnection, ZAbstractRODataset, ZAbstractDataset, ZDataset,
  ZAbstractConnection, IWCompEdit, IWCompLabel, comobj, dateutils, IWGrids, series,IWRegion, SimpleGraph;
type
  TIWUserSession = class(TIWUserSessionBase)
    ZQuery1: TZQuery;
    ZConnection1: TZConnection;
  private
    { Private declarations }
  public
    { Public declarations }
    inputmanual, goroc, goratio : boolean;
    qtdcolunapreenchida, auxcrit, auxalt: Integer;
    dirprob, vQdeColunas, vQdeLinhas: Integer;
    excelverif: Boolean;
    NameProblem, id, email: string;
    ranking : array of integer;
    edt, edtpesos: array of array of TiwEdit;
    lblauxcrit, lbldircrit, lblcrit: array of TIWLabel;
    lblalt: array of TIWLabel;
    auxmanual : integer;

    reiniciar : boolean;

    auxmatresultado : array of array of string;


    LimPesoMax : array of double;
    LimPesoMin : array of double;

    Idioma : Integer; //var sistema bilingue
    IdAHeur, IdBHeur: integer;
    MinVK, MinCost, MaxCost, Fmult: Double; // **Menor valor máximo admitido para uma constante de escala **//
    MatCostPortfolioCB: array of array of string; // matriz que armazenará vetor cost e custos acuulados exibidos no grid final dos resultados CB
    CostNorm, Cost: Array of Double; // cost- armazena var local Costproj, com custos de cada proj
    Budget, CumulativeCost: Double; // global que armazena a var local Totbudget
    /// CumulativeCost- global que recebe CustoA da rgntabelafinalportifolio. E auxilia na criação da planilha de exportação
    Portfoliofinal: string; //global que armazena info da variável local Str, utilizando na construção da planilha de exportação
    vtrRecomendacao_ciclo: array of Double; // vetor dinamico que armazena a cada resposta do decisor a recomendação atual
    primeiroacesso, zeramatriz: boolean;
    VetorPergunta: array of Double; // Armazena a última consequência real

    // de cada critério exibida ao decisor
    auxauxa, editar: integer;
    VTRAlternatives: array of integer; // Esse vetor irá armazenar os índices das alternativas a serem exibidos
    // nos gráficos e na tabela de resultados.
    Auxall: integer; // auxiliar para caso o decisor decida variar todos os criterios na AS de uma só vez
    Avaliable, elicitarperfis: boolean; // Indica se o relatório está ou não disponível; Indica que o decisor irá realizar a elicitação de perfis
    //email: string;
    senha: string;
    AUXLOGOUT: integer;
    AnswerInc: integer;
    RETPROBDEPOIS: integer;
    flag, Problematica, Rlevels, NDRl: integer;
    Ans_seq: string;
    ID_Problema: string;
    ID_user: string;
    ID_Alternative: array of integer;
    ID_criteria: array of integer;
    Cri_Cong: array of integer;
    probname, name_, email_: string;
    Altglobalvalue: array of Double;
    Estagio: integer;
    parA, parB, Sconst: array of Double;
    stage: integer;
    Nomeprojeto: string;
    NomeCrit, NomeAlt: array of string;
    NumCrit, NumAlt, NumAH, NDA: integer;
    NumCriBlind, NumCriCong: integer;
    LimMin, LimMax: Array of Double;
    ConsNorm, ConsMatrix: Array of Array of Double;
    ValDiscreto: integer; // utilizado para indicar que os critérios discretos foram validados
    VetPOA, IAltAH: Array of integer; // HE
    OPAH, CriteriaName: Array of string;
    IndiceA, IndiceB, Indice, IA, IR, IB, Fase, PAns: integer;
    WDP1: Double;
    CriBlind: Array of boolean;
    HeurQuest, Etapa, NumQuest, NumQP, NumRQuest: integer;
    VtrOrdem, Typeofcrit: Array of integer;
    VtrRelacoesBlin: Array of boolean;
    MaxCons, MinCons: Array of Double;
    MatrizFaixa: Array of Array of integer;
    Niveis, Ftipo, CriTipo: Array of integer;
    FlagDiscBlin, AuxAns, PAH: boolean;
    ConsReal, Epsilon: Double;
    Max, Min, MaxPeso, MinPeso: Array of Double;
    ArmazenaParaPar: array of array of array of integer;
    Ciclo, NumRel, NumInd, Sol: integer;
    MatResultOrd: Array of array of string;
    Analyst: integer; // -1: a conta não é de analista, 0 conta de analista
    ConsRazao: array of array of Double; // Valores da matriz em escala de razão
    Escala: integer; // Indica escala intervalar ou de razão para construção dos gráficos dos resultados
    TRE: array of boolean; // auxiliar para o checklistbox da região dos gráficos (barras)
    // variaveis cadastro
    genero, statuscurso, nescolaridade: string;
    // variaveis globais da ordenação dos criterios
    m, n: integer; // contadores da ordenação dos critérios
    contblin: integer; // contador dos criterios blindados
    vtrblin: array of boolean; // vetor que vai armazenar os criterios blindados
    auxnomeblin: array of string;
    vtrposicoes: array of integer; // vetor que vai colocar as posições do criterio no ranqueamento
    auxiAPAGAR: integer;
    MATAH, HEINDEX: Array of array of integer;
    TotalAH: integer; // auxiliar que guarda o total de Avaliações holísticas realizadas
    qtitensselected: integer; // conta quantos itens foram selecionados no listbox inicial
    vetorgrafico: array of integer;
    // variaveis globais da ordenação dos criterios aos pares
    numalocado: integer; // indica o número de critérios que já foram alocados pela avaliação pareada no vtrposições, já inicia com 2
    numrespA: integer; // indica se a resposta A dada anteriormente foi a primeira ou não, se foi a primeira 0 se não foi 1
    posicrit: integer; // armazena a posição dos critérios no vetor de preferencia
    indice1: integer; // relacionado a posição que vai sendo modificada dos critérios do primeiro gráfico da comparação pareda
    indice2: integer; // relacionado a posição que vai sendo modificada dos critérios do segundo gráfico da comparação pareada
    respanterior: integer; // a reposta anterior da pergunta de comparação pareada o se for A e 1 se for B
    VerificaUnit, RgnReset: integer; // Auxiliar para trocar de unit e abrir em região diferente
    AHType, NumAltAH, NumRAH: integer; // He
    TIFS, CBT, Aux0501: boolean; // TIFS indica se há inconsitência HExDE/ CBT indica se uma cmbinação de restrições estão gerando inconsistência
    AuxRSel: boolean;
    TINFS, VetAuxOrd, VetVisualization: array of integer;
    NumRQuestion: integer; // número real de perguntas respondidas
    LimMaxA, LimMinA: Array of Double;
    Visualization, Analystkey, indexHE: integer;
    MatrizGrupo: Array of array of integer;
    vetpesos: array of array of Double; // Vetor que armazena as constantes de escala das POAS
    Representantes: array of integer; // Armazena os critérios representantes de cada posição do raking
    vordem: array of integer; // auxiliar que pega a ordem dos criterios
    AuxValidacao: string; // Utilizado para gerar o código de validação eviado ao usuário
    MatIndiferenca: array of array of string; // Utilizada para armazenar alternativas indiferentes na prob de escolha
    /// ///////// Variáveis utilizadas na Análise de Sensibilidade //////////////
    START, FINISH: boolean;
    Sensitivity: boolean; // Se True , Indica que a análise de sensibilidade foi realizada
    VauxSel: integer; // Armazena o índice do critério selecionado
    pmax, pmin: array of Double; // Amazena os percentuais de variáção
    crisel: array of boolean; // Indica se o critério foi ou não selecionado
    nsel: integer; // Número de critérios selecionados
    ninst: integer; // Número de instâncias
    ncv: integer; // Número de critérios variados
    vtrASordchange: array of integer;
    perchange, percigual: Double;
    MconsN, MatrizConsR: array of array of Double;
    // Armazena as consequências normalizadas
    // utilizadas na Análise de Sensibilidade
    MatrizGruposAS: Array of Array of integer; // Armazena as alternativas nas
    /// / posições do ranking
    VetSol: array of integer; // Vetor de POAS
    MatrizSol: array of array of integer; // MatrizParaPar
    numlevels: integer; // Número de posições no ranking
    marizpercent: array of array of integer;
    auxmarizpercent: array of array of Double;
    VAltPositions, VAltPositionsO: array of integer; // Armazena a posição de cada alternativa no ranking
    MatrizprofileOrdenacao: array of array of array of string; // Utilizada para
    // armazenar a informação das alternativas incluídas e excluídas em cada
    // nível do ranking bem como os perfis das alternativas em cada cenário
    AuxFInc, AuxFexc: array of integer;
    AuxMatrizprofile: array of array of string;
    Flevelsc: array of integer;
    auxfsol: array [0 .. 1] of integer;
    auxrchange, auxas: integer;
    /// ////////////////////////////////////////////////////////////////////////
    /// ///////////////Var utilizadas para a construção do diagrama de Hasse /////
    MatrizParaPar2: Array of Array of Double;
    // matriz utilizada para construção
    MatrizHolistica: array of array of integer;
    // do grafo de dominancia
    Node1: Array of integer; // guarda quais alternativas estão na camada analisada
    Node2: Array of integer; // quarda quais alternativas ainda serão alocadas em níveis restantes
    Node2Ant: Array of integer; // quarda quais alternativas ainda restaram após camada anterior
    QntCam: integer; // guarda o número de camadas existentes no grafo
    CamAlt: Array of array of integer; // mostra em qual camada estão as alternativas
    Node3: Array of integer; // guarda o número de alt de cada nível
    contador: integer; // variavel de auxilio para montar a CamAlt
    contador2: integer; // primeira vez que roda o procedure AMD
    Node4: Array of integer; // guarda, em ordem, qual os valores de i, j, k, da última
    // ligação entre nós que irá rodar o procedure DomAlt2
    Node5: Array of array of integer; // matriz que guarda quais alt dominam cada alt
    Node6: Array of integer; // guarda a ordem das alternativas estão preenchidas
    // no grafo e utiliza para fazer as ligações entre as alt no procedure NumNode
    Node7: Array of integer;
    Node8: Array of integer;
    Node8Ant: Array of integer;
    Node9: Array of array of integer;
    AltOrdemGrafo: Array of integer; // guarda a ordem das alternativas a serem preenchidas
    // no grafo
    CountWords: Array of String; // Guarda o nome da alt analisada e a separa de acordo com
    // os espaços
    axm2: array of array of String; // matriz auxiliar para visualização tabular
    auxContador: integer; // para saber se o botão de avaliação holistica foi acionado
    // variáveis usadas no procedure DomAlt2, extraídas do NumNode
    auxcontador2: integer; // para saber que as duas alternativas foram selecionadas para o par a par
    aux_Nodei: integer;
    aux_NodeiAnt: integer;
    aux_Nodej: integer;
    aux_Nodek: integer;
    SimpleGraph: TSimpleGraph;
    Bconseqord, Wconseqord: array of Real; // auaxiliares que pegam os valores maximos e minimos das consequencias dos criterios de acordo com sua direção
    aux_i: integer;
    /// /////////////////////////////////////////////////////////////////////////

    /// ////////// Variáveis adicionadas para o BD //////////////////////////////
    tempinicio, tempofinal, tempexec, tempexecant: string;
    Num_Rpos, AuxsalvarBD: integer;
    BDMatrizParaPar: array of array of integer;
    Statusdoprocesso: integer;
    BDstatus: string;
    auxbdgraph, BDgraph: string;
    nround, auxid_ah: integer;
    auxNumAH: integer;
    //nameproblem: array of string;
    password: string;
    emailsbd: string;
    auxnome: string;
    idpoblemasuser, datainicioproblema, chavesproblemas,
      datafinalproblema: array of string;
    dtinicretorn, dtfnretorn: string;
    mudadata: integer;
    auxnoans: integer;
    auxinconsi, auxgoback, contholisticglobal: integer;
    auxtipoord, contord: integer;
    id_chave: string;
    tipoordretomada, auxparacasonsaieretone: integer;
    auxbd_id_ah_incos: array of string;
    vet_rpos_ciclo: array of string;
    consrealAnt: Double; // pega o valor anterior da consreal
    WDP1ant: Double; // pega o valor de wdp1 anterior
    Auxinconsistency: integer; // impedir que no BD seja salvo duas vezes o caso 0
    as_in_progress: integer; // verifica se a AS do desk está rodando
    qp_ret_exp, vet_poa_ciclo, vet_preferences_ciclo, vet_idA_ciclo,
      vet_idB_ciclo, vet_real_consA_ciclo: array of string;
    auxiliar_ans_seq_bd: boolean; // auxiliar que indica se o problema foi retornado e auxilia no proeenchimento do ans_seq
    auxquantdeah: integer; // armazena quantas holisticas foram feitas para aquele problema
    vtr_id_ah_sum, vtr_alt_avaliadas_sum, vtr_decision_sum, vtr_id_alt_e_sum,
      vtr_id_alt_c_sum, vtr_poa_sum: array of string;
    vet_ret_hc: array of string; // verifica se foi realizada Ah
    vtr_num_rlevels_sum: array of string;
    vet_recomend_ciclo: array of string; //recebe as recomendações do CB vindas do BD para o sumario
    /// /////////////////////////////////////////////////////////////////////////

    /// ////////// VARIAVEIS AUXILIARES PARA ENTRADA DE DADOS MANUALMENTE
    AuxNomeCrit, AuxNomeAlt: array of string;
    AuxNumAlt, AuxNumCrit, AuxIndice, AuxIndiceAlt, EditaCrit, EditaAlt,
    AuxCritApagados, AuxAltApagadas: integer;
    AuxEscalaTipo, AuxDirecTipo, AuxFuncTipo: array of integer;
    AuxNiveis: array of integer;
    AuxParA, AuxParB: array of Double;
    NomeAlternativaEditada, NomeCriterioEditado: string;
    IndiceCriterioEditado: integer;

    /// /////////// Variável para matriz de edits do cadastro manual
    MatEdit: array of array of TIWEdit;
    ArrayAlt, ArrayCrt: array of TIWLabel;
    AuxResetedit: boolean; //auxiliar que verifica se edit da matriz está preenchido e btn "save prob" ñ foi acionado

    /// ///// Variaveis auxiliares para dados do diagrama de hasse
    DomMat: string;
    auxDM: integer;
    HolMat, auxHM: string;
    alt1, alt2: integer;
    Mat_HM: array of array of string;
    NomAlt: string;
    levels_numalt: string;

       //PARA CLASSIFICAÇÃO (MOACIR)
    VetPesosMax, VetPesosMin : array of array of real; //vetores de pesos resultantes da maximização e minimização
    perfis : array of real; //perfis do problema de classificação; são alterados na AS com variação dos perfis
    class_min, class_max : array of integer; //classes mínima e máxima possíveis para cada alternativa
    num_cat : integer; // número de categorias selecionadas na elicitação de perfis

      //PARA CLASSIFICAÇÃO (JÚLIA TOMÉ)
    indice_altp: array of integer; // Esse vetor irá armazenar os índices das alternativas perfil a serem exibidos
    indice_alt: array of integer; // Esse vetor irá armazenar os índices das alternativas a serem exibidos
    altperfil: array of array of Double; //Armazena a matriz não normalizada com o valor dos perfis pra cada critério
    indice_altnaoclassificadas: array of integer; // Esse vetor irá armazenar os índices das alternativas do problema não classificadas
    ValAltMax, ValAltMin: array of real;
    Consrazaop: array of array of Double; // Valores da matriz perfil em escala de razão
    bestAlt, notBestAlt: integer; //auxiliares que possuem os índices das alt avaliadas holisticamente
    indiceCmbxBest: array of array of Double;
      // indiceCmbxBest[0,0]: Pega a informação da alternativa do CmbxAlternativa1
      // indiceCmbxBest[1,1]: Pega a informação da alternativa2 se for perfil
      // indiceCmbxBest[1,0]: Pega a informação da alternativa2 se nao for perfil

    //PARA CLASSIFICAÇÃO (PEDRO)
    solucao_encontrada : boolean; // Variável booleana que indica se a solução do problema foi ou não encontrada
    auxiliar_perfil : array of TiwLabel; // Vetor de labels utilizado para
    auxiliar_vp : array of TiwEdit;


    //PARA ANALISE DE SENSIBILIDADE DA CLASSIFICAÇÃO (MOACIR 03.08.2021)
    class_final : array of integer; //classes finais para cada alternativa
    contadorAS : array of array of integer; // contador do resultado da AS; primeira dimensão
    /// é alternativa, segunda é classe (contadorAS[i,0] será o número de instâncias em que
    /// a alternativa i não foi classificada
    escala_AS : integer; // 0 - intervalar ; 1 - razão
    perfis_inseridos : array of real; //perfis inseridos na elicitação de perfis; NÃO SÃO ALTERADOS pela AS com variação de perfis
    Pmax_perfil, Pmin_perfil: array of double;
    valor_nadir_minimo : real; // valor mínimo da NADIR; vem da uPerfis com escala de razão
    ValAltMax_final, ValAltMin_final : array of real; //valores globais das alternativas
    Alt_nao_classificadas: integer; //Indica o número de alternativas não classificadas
    ClassRestart, ASProfiles: boolean; //indica se um problema de classificação vai ser reiniciado
    tipoAS : integer; // 0 - variando consequencias; 1 - variando perfis


////////////////////// variáveis elicitação intracritério- 16.02.22 /////////////////
    NumCont: integer; // numero de critérios contínuos na etapa intraelic
    BoundMn, BoundMx: array of Double; // vetor que armazena os limites que serão usados "localmente" na bisec
    LimSup, LimInf, LimMed, BoundMd: Double; // var que armazenarão pior, melhor e medio valores para atualização do range durante bisec
    ValorA, ValorB, Xn: Double; // var que armazenam ultimos valores da elicitação. constroi plote graph
    DA: Double; // da= desvio utilizado para validação do range;
    parada: boolean; // var que indicará se criterio de parada da bisec foi atendido
    vtrpontosbisec, Perc: array of Double; // vetor que armazenará dinamicamente os 3 pontos elicitados
    // perc =vtr que armazena os valores percentuais selelcionados, para cada criterio elicitado na intra
    contadorponto: integer; // contador que será incrementado a cada início de elicitação para novo ponto da func
    matrizfixa, matrizdinamica: array of array of Double; // matrizes utilizadas no plote dos graficos intra
    matrizpontosreais: array of array of Double; // matriz com todos os pontos elicitados no robolema
    verificaintra, DinamVerificaIntra: array of boolean; // verificação do preenchimento na planilha excel. true= info faltando
    Contavazio: integer; // contador que se >0 significa que o boolean verificaintra foi verdadeiro alguma vez
    AuxSalvai, retornointra: integer; //ALTERAÇÃO EM 14.02.22- retornointra: auxilia na determinação de qual tela intra o sis retomará
    AuxConsNorm: array of array of Double; // auxiliar utilizada para salvar a cada criterio valores normalizados por emio da interpolação.
    vtrNomeNumCont: array of string; // vetor que armazena nome dos criterios continuos do problema
    vtrLinearintra, vtrChangeLin: array of integer;
  //vtrlinearintra indica que crit entrou para etapa de elcitação, mas já era FTipo=1
  //vtrChangeLin salva indice dos criterios que FTipo foi alterado para linear na bisec
    VtrcritBisec: array of boolean; // criterios continuos que foram elicitados com bisec = true
    auxVtrCL, auxVtrLI: integer; // variável auxiliar para o vetor vtrChangeLin, indicando a posição nesse vetor que deve ser inserido o índice do critério alterado
  //AuxMatFixaNormaliza: array of array of Double; //matriz que armazena pontosd normalizados para utilizar na inversa
  //VerificaLin: integer; //verificação se link que declara linear, na intracriterio, foi apertado
    VerificaDefPoint: integer;  //verifica se decisor já respondeu uma pergunta na bisseção, liberando o lnk "Define point"
    fvround: integer; //var para salvar no BD qual o round de um problema para a etapa Intracriterio
    auxvaloresplotados: array of double; //auxiliar pra pegar valor do rgd do grafico
    EstagioIntra: integer; //var utilizada para atualizar stage do problema na etapa intracriterio
    vtrrangeintra: array of double; //auxilia no plote do grafico para perguntas bisec
    Strpadrao: string; //str que analisa se o range de um critério é >10 ou não. padronizando a visualização em alguns graficos
    newftipo: array of integer; //vtr que recebe mesmos valores de ftipo, mas é atualizado quando algum crit é declarado como linear durante elicitação. auxilia no retonorno quando estagiointra=1
   //** Variáveis Portfólio Combinatório **//
   PortfolioG : Boolean; //** Indica se o sistema está realizando combinações de projetos **//

    // ** Larissa - Alteração intracrit
    OPrange: array of Real;
    NumPergRange: Integer;
    AuxPonto, NaoElicita: Real;
    TipoCrit: array of Real;
    ParadaDisc: Boolean;
    elicita05, elicita25, elicita75, AuxParada: Boolean;
    aux3niveis: Integer;
    seqresp: string;
    ETFimIntra, ETInicioIntra, ETSelIntra, EAuxNomeCrit, EindCrit, Eseqintra : string;
    VetFixAuxiliar: array of real;
    //** Criado por Manoel para a AS da Intra **//
    //** Armazena os valores utilizados na AS da função valor **//
    MvarIntra : array of array of array of real;
    //** Armazena o índice do ponto sendo elicitado na intra **//
    PontoIntra : Integer;
    //** Armazenam o espaço de pesos atualizado pela AS a cada instância **//
    LimMinASI, LimMaxASI : array of real;
    //** Armazenam o espaço de pesos atualizado pela AS a cada instância **//
    MLimMinASI, MLimMaxASI : array of array of real;
    //** Indica que o tipo de AS é das funções valor **//
    AsVF : boolean;
    //** Indica que o espaço foi inviabilizado **//
    UnfeasibleR : boolean;
    WDPHEURAS : Real;
    //** Armazena todos os dados das ASs **//
    MatDataAS : array of array of array of string;

    // PARA CB (MARIA ELVIRA)
    nvlint : integer;
    BudResult, nproj : Double;
    existimcomp: boolean;
    ele_comuns, eledif1, eledif2 : string;
    matriz_combport : Array of Array of real; // armazena as combinações de portfólios
    matriz_vetop, matrizf_vetop : array of array of integer;
    vtrcost_combport, vtrcombcost: array of real;
    vtrnome_combport, vtrcombname: array of string;
    vtraltcomb,vetop: array of integer;
    //** Descrever **//
    GComb : boolean;
    //** Armazena os índices dos portfólios presentes no combobox1 e 2 respectivamente **//
    vetorport1, vetorport2 : array of integer;
    {//** Armazena as consequências dos portfólios a serem avaliados na holística considerando apenas
     os projetos distintos **//}
    MatrizAvaliacaoPort : array of array of Real;
    //** Armazena os índices dos elementos comuns a dois portfólios selecionados **//
    ElementosComuns : Array of Integer;
    //**Armazena os índices dos portfólios selecionados para a holística **//
    projsel: Array [0 .. 1] of integer;
    //** Indica se o critério deve ser escondido na plotagem das alternativas do gráfico, ficará verdadeiro se todas as alternativas exibidas tiver o mesmo valor de consequência no critério **//
    esconderCrit: array of Boolean;

           // ** Cores - YARA **//
    //** vetor que armazena as cores das alternativas do problema  **//
    Coresalt: array of Integer;
    //**Indica que uma combinação ótima já foi obtida **//
    Coptimal : boolean;
    // ** Variáveis para a nova pairwise **//
    MatPositions: Array of array of Integer; // armazena o índice dos critérios nas posições que cada um ocupa
    MatPositionsA: Array of array of Integer; // armazena a MatPositions do ciclo imediatamente anterior
    Position: Integer; // Indica a posição pleiteada pelos critérios
    PositionA: array of Integer; // Indica a posição pleiteada pelos critérios na rodada anterior
    PositionAA: array of Integer;
    NextCri: Boolean; // Indica se já é possível passar para o próximo critério
    NextCriA: Boolean;
    RMax, RMin: Integer; // Indica a posição máxima e a mínima que o critério sendo perguntado pode assumir
    RMaxA, RMinA: Integer;
    Crepresentante: Integer; // indica qual é o critério representante quando há indiferença

    // ** Variáveis para a captação dos valores globais max e min de todas as alternativas **//
    MatOverallValues: Array of array of Real;

    // ** Variável para definição do número de casas decimais **//
    StrDecimal: array of string;
    StrDInd: Integer;

    // ** Variáveis ligadas à declaração de critérios inteiros **//
    Cri_inteiro: array of Boolean; // indica os critérios declarados como inteiro

    ClpSolve: Boolean; // ** Indica que é hora de usar o LPSolve para a resolução do problema de ordenação
    //** Variável auxiliar para indicar que o uduário está revendo um par **//
    reverPar: Boolean;
    //** Indica que após a mudança do limiar de equivalência, vai ser necessário dar mais informação **//
    NovaPerg : boolean;
    //** Auxilia na atualização dos limites dos não intermediários **//
    AvaliandoInd : boolean;
    ConseqA, ConseqB: Double;

    /// ////////////////////** Critérios não adjacentes **///////////////////////
    MLimMin, MLimMax: Array of array of double; // Matriz de limites: utilizada para perguntas entre crit não adjacentes
    MLimMinA, MLimMaxA: Array of array of double;
    PergNAdj: array of array of Integer;
    Mblindados: array of array of Boolean; // Matriz de critérios blindados: utilizada para perguntas entre crit não adjacentes
    // ** Matrizes que armazenam o máxmimo e o mínimo limite imposto pelas perguntas entre critérios não adjacentes **//
    MLimMinImp: array of array of array of Real;
    MLimMaxImp: array of array of array of Real;
    MLimInd: array of array of array of Real;
    //** Armazena o desempenho da consequência usada na perg da Heurística **//
    WDPHEUR : Double;
    /// ////////////////////////////////////////////////////////////////////////

    /// ////////////////////** Geração de Siglas **///////////////////////
    CritSiglas, AltSiglas: array of string;
    /// ////////////////////** Máximas diferenças  **///////////////////////
     maxDiff: array of array of Real; // Armazena a matriz de máximas diferenças entre as alternativas
    vetMatDiff: array of TIWGrid; // Criação de grids dinâmicos na tela de resultados parciais
    vetregions: array of TIWRegion; // Criação de regiões dinâmicas na tela de resultados parciais
    vetLabelsDm: array of TIWLabel; // Criação de labels dinâmicas na tela de resultados parciais
    RegMaxDifDim: Integer; // Armazena a altura da região de máxima diferença
   /// ////////////////////** Holística Ordenação **///////////////////////
    Matcombination: array of array of Integer;
   // esconderCrit: array of Boolean; // Indica se o critério deve ser escondido na plotagem das alternativas do gráfico, ficará verdadeiro se todas as alternativas exibidas tiver o mesmo valor de consequência no critério
    IndiceCBHE: Integer;
    SavedPairs: array of Integer; // ** Armazena os pares de alternativas salvos para posterior análise na Holística de Ordenação
    PergDiscFinalized: Boolean; // ** Indica que a elicitação especial para critérios de escala construída foi finalizada **//
    ParHolisticaOrd: Integer; // ** Salva o índice do par de alternativas que está sendo visualizada
    PHOChange: Boolean; // ** Indica se deve alterar a variável ParHolisticaOrd quando execura o onchange dos comboboxes das alternativas
    VtrAlternativesND: array of Integer; //** Indica as alternativas que podem ser comparadas com a primeira selecionada **//
    /// ////////////////////** Constante para PL de Escolha**///////////////////////
    Epson: Real; // ** Limiar da potencial otimalidade **//

   /// ////////////////////** Variáveis da Etapa 4 **///////////////////////
    TOADC: Boolean; // ** Indica que só existem critérios discretos no problema
    DCLab: Boolean; // ** Indica que a label da elicitação dos níveis intermediários deve estar visível **//
    VGMax, VGMin: Array of Real; // ** Armazena o valor global maximizado e minimizado das POAS **//
    DiscIntBlind: Array of Boolean; // ** Indica quais critérios não podem mais ser perguntados na Etapa 4**//
    TADC: Boolean; // ** Indica que existe ao menos 1 critério discreto com Nível <> 2, usado para iniciar a etapa 4 **//
    VMinG: Integer; // ** Armazena o X referente ao V(x) = LimMin do critério discreto naetapa 4
    NextRelDisc: Boolean; // ** Na etapa 4, esta variável aponta se devemos ou não passar para uma nova relação
    NRDTest: Boolean; // ** Na etapa 4, esta variável aponta que tentamos fazer perguntas para uma determinada relação de critérios e não foi possível, assim estamos testando uma nova relação
    NiveisIntermediarios: array [0 .. 2] of Boolean; // ** Indica quais níveis intermediários ainda podem ser elicitados **//
    EtapaAnt: Integer; // ** Indica a etapa da elicitação anterior a última pergunta respondida **//

    /// ////////////////////** Variáveis da Etapa 3 **///////////////////////
    BIndex : Integer;

    //** Outras variáveis incluídas para inclusão das features do SAD de Carla **//
    valor_razao: array of double; // vetor que recebe os resultado das equações best e worst no procedimento intra_razão

  end;

implementation

uses Unit1, ServerController;

{$R *.dfm}

end.