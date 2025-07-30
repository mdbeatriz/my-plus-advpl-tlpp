#include "protheus.ch"
#include "topconn.ch" 
#include "tbiconn.ch"
#include 'rwmake.ch'
#include 'colors.ch'
#include 'font.ch'

/*/{Protheus.doc} ESTSLD010
(long_description)
@type user function
@author user
@since 12/04/2025
@version version
@param param_name, param_type, param_descr
@return return_var, return_type, return_description
@example
(examples)
@see (links_or_references)
/*/
User Function ESTSLD010()

Private cFile		:= GetNextAlias() 
Private aPR	   		:= {}    
Private aLin 		:= {}
Private nTotLin 	:= 0
Private nTotSKU 	:= 0
Private nTotFil		:= 0     
Private __cProduto	:= Space(999)
//Private coTblPR	
Private oTempTable  := Nil

SetPrvt("oFont1","oDlgPR","oSay1","oSay2","oSay3","oSay4","oSay5","oBtn1","oBtn2")

oFont1     := TFont():New( "Courier New",0,-11,,.F.,0,,400,.F.,.F.,,,,,, )
oDlgPR      := MSDialog():New( 092,232,274,905,"Importação de arquivo .csv para criação de lote e lançamento de saldos iniciais.",,,.F.,,,,,,.T.,,oFont1,.T. )
oSay1      := TSay():New( 012,012,{||"Esta rotina tem como objetivo gerar a Importação de Planilha Excel para criação de lote e lançamento de saldos iniciais. "},oDlgPR,,oFont1,.F.,.F.,.F.,.T.,CLR_BLACK,CLR_WHITE,308,008)
oSay2      := TSay():New( 020,012,{||"na extensão *.csv."},oDlgPR,,oFont1,.F.,.F.,.F.,.T.,CLR_BLACK,CLR_WHITE,308,008)
oSay3      := TSay():New( 028,012,{||"Após a importação, será possível também realizar ajustes nos valores modificados "},oDlgPR,,oFont1,.F.,.F.,.F.,.T.,CLR_BLACK,CLR_WHITE,308,008)
oSay4      := TSay():New( 036,012,{||"anteriormente na planilha antes de sua execução."},oDlgPR,,oFont1,.F.,.F.,.F.,.T.,CLR_BLACK,CLR_WHITE,308,008)
oSay5      := TSay():New( 044,012,{||"Escolha o caminho do arquivo que deseja importar ao confirmar."},oDlgPR,,oFont1,.F.,.F.,.F.,.T.,CLR_BLACK,CLR_WHITE,312,008)

oBtn1      := TButton():New( 063,228,"Confirmar",oDlgPR,,043,014,,oFont1,,.T.,,"",,,,.F. )
oBtn1:bAction := {|| MsAguarde({|| sfImpArqPR()},"Abrindo diretorio...") }

oBtn2      := TButton():New( 063,276,"Cancelar",oDlgPR,,043,014,,oFont1,,.T.,,"",,,,.F. )
oBtn2:bAction := {||Close(oDlgPR)}

oDlgPR:Activate(,,,.T.)

Return()


/*/{Protheus.doc} nomeStaticFunction
    (long_description)
    @type  Static Function
    @author user
    @since 12/04/2025
    @version version
    @param param_name, param_type, param_descr
    @return return_var, return_type, return_description
    @example
    (examples)
    @see (links_or_references)
/*/
Static Function sfImpArqPR()

Local cMascara  := "Arquivos de excel|*.csv"
Local cTitulo   := "Escolha o arquivo para criação de lote e lançamento de saldos iniciais"
Local nMascpad  := 0
Local cDirini   := "C:\"
Local lSalvar   := .F. /*.T. = Salva || .F. = Abre*/
Local nOpcoes   := GETF_LOCALHARD
Local lArvore   := .F. /*.T. = apresenta o árvore do servidor || .F. = não apresenta*/

Private targetDir := cGetFile( cMascara, cTitulo, nMascpad, cDirIni, lSalvar, nOpcoes, lArvore) //cGetFile(,"Importação para criação de lote e lançamento de saldos iniciais ",0,"\",.T., nOR( GETF_LOCALHARD, GETF_LOCALFLOPPY))
Private cArqcsv := targetDir
Private nHdl    := fOpen(cArqcsv,68) 
Private cEOL    := "CHR(13)+CHR(10)"

If Empty(cEOL)
	cEOL := CHR(13)+CHR(10)
Else
	cEOL := Trim(cEOL)
	cEOL := &cEOL
Endif

If (nHdl == -1) .OR. (Alltrim(SubStr(targetDir, RAT(".", targetDir))) <> ".csv")
	MsgAlert("O arquivo de nome "+cArqcsv+" nao pode ser aberto! Verifique se extensão é compatível ou se o arquivo se encontra aberto.","Atencao!")
	Return
Endif

Processa({|| sfImpPR() }, "Aguarde...", OEMTOANSI("Processando Planilha para criação de lote e lançamento de saldos iniciais..."),.F.)

Return()


/*/{Protheus.doc} nomeStaticFunction
    (long_description)
    @type  Static Function
    @author user
    @since 12/04/2025
    @version version
    @param param_name, param_type, param_descr
    @return return_var, return_type, return_description
    @example
    (examples)
    @see (links_or_references)
/*/
Static Function sfImpPR()

Local nY		 := 0
Local cLinha	 := ""
Local aDados	 := {}
Local cColumns   := ""
Private aSldProd := {}

If !File(cArqcsv)
	MsgStop("O arquivo " +cArqcsv+" não foi encontrado. A importação será abortada!","[QG_ATF] - ATENCAO")
	Return
EndIf

SD5->(DbSetOrder(2)) //D5_FILIAL+D5_PRODUTO+D5_LOCAL+D5_LOTECTL+D5_NUMLOTE+D5_NUMSEQ

FT_FUSE(cArqcsv)
ProcRegua(FT_FLASTREC())
FT_FGOTOP()

While !FT_FEOF() //!FT_FEOF(cArqcsv) 

	nY:=nY+1
	
	IncProc()         		   
	//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
	//³Cabecalho            ³
	//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ0,
	If nY = 1
		cLinha := FT_FREADLN()
		FT_FSKIP()
	EndIf
	
	cLinha := FT_FREADLN() 

	If Empty(Alltrim(cLinha)) //Pula Linha Vazia.
		FT_FSKIP()
		Loop
	EndIf     		
	
	Aadd(aDados,Separa(cLinha,";",.T.))

	If  Len(aDados[nY])  <> 8 
		MsgStop("Impossível importar planilha, pois é Obrigatório que existam exatamente 8 (COD_PRODUTO, ARMAZEM, LOTE, QTD, DT_FABRIC, DT_VENCTO, TIPO_CONTROLE, ANOS) colunas de informações."+cEOL+"Favor, Corrija Planilha antes de Importa-la.")
		Return() 
	EndIf
	
	//valida se existe qtd. negativA
	If AT("-", Alltrim(aDados[nY,4]) ) == 1  
		MsgStop("Impossível importar planilha com Quantidade negativa!"+cEOL+"Favor, Corrija Planilha antes de Importa-la.")
		Return() 
	Elseif !sfVerEsp(Alltrim(aDados[nY,4]))
		Return() 	
	EndIf

    //verifica se data de validade está vazia.
    if Empty(aDados[nY][1])  
        cColumns := "Codigo do Produto"
    ElseIf Empty(aDados[nY][2]) 
        cColumns := "Armazem"
    ElseIf Empty(aDados[nY][3])
        cColumns := "Lote"
    ElseIf Empty(aDados[nY][4]) 
        cColumns := "Quantidade"
    ElseIf Empty(aDados[nY][6]) 
        cColumns := "Data de Vencimento"
    Endif	

    If !Empty(cColumns)
        MsgStop("Não foi possível importar a planilha: há registros na coluna '"+cColumns+"' em branco."+cEOL+"Favor, Corrija Planilha antes de Importa-la.")
		Return()
    EndIf
    
	If AT("'", Alltrim(aDados[nY,1]) ) >= 1
        aDados[nY][1] := StrTran(Alltrim(aDados[nY,1]), "'", "")
	EndIf	
    
    If AT("'", Alltrim(aDados[nY,2]) ) >= 1
        aDados[nY][2]  := StrTran(Alltrim(aDados[nY,2]), "'", "")
    Endif
    
    If AT("'", Alltrim(aDados[nY,3]) ) >= 1
        aDados[nY][3]  := StrTran(Alltrim(aDados[nY,3]), "'", "")
    Endif

    If SD5->(DbSeek(xFilial("SD5")+Padr(aDados[nY][1],TamSx3("D5_PRODUTO")[1])+Padr(aDados[nY][2],TamSx3("D5_LOCAL")[1])+Padr(aDados[nY][3],TamSx3("D5_LOTECTL")[1])))
        MsgStop("Não foi possível importar a planilha pois Lote "+aDados[nY][3]+" para o produto: "+aDados[nY][1]+" já existe!"+cEOL+"Favor, Corrija Planilha antes de Importa-la.")
        Return()
    Endif

	If !Empty(aDados[nY][1])
        
        nRet := aScan(aSldProd, {|x| x[1] == Alltrim(aDados[nY,1])+Alltrim(aDados[nY,2])})
        
        If  nRet <= 0 //verifica se o produto já existe no array.
            Aadd(aSldProd,{Alltrim(aDados[nY,1])+Alltrim(aDados[nY,2]), val(aDados[nY,4]), Len(Alltrim(aDados[nY,1])), Len(Alltrim(aDados[nY,2]))}) //adiciona produto, armazem e quantidade
        Else
            aSldProd[nRet][2] += val(aDados[nY][4]) //soma quantidade do produto já existente no array.
        EndIf

	    Aadd(aPR,{	aDados[nY,1],; //CODPRODUTO
	    			aDados[nY,2],; //ARMAZEM
	    			aDados[nY,3],; //LOTE
	    			aDados[nY,4],; //QTD	
					aDados[nY,5],; //DTFABRIC	
					aDados[nY,6],; //DTVENCTO	
                    aDados[nY,7],; //TPCONTROLE	
                    aDados[nY,8];  //ANOS	
					})   	

	EndIf	

	FT_FSKIP() 
		
EndDo                  

FT_FUSE()

sfPRVerPla()

Return()  

/*/{Protheus.doc} nomeStaticFunction
    (long_description)
    @type  Static Function
    @author user
    @since 12/04/2025
    @version version
    @param param_name, param_type, param_descr
    @return return_var, return_type, return_description
    @example
    (examples)
    @see (links_or_references)
/*/
Static Function sfPRVerPla() 

Local aCampos		:= {}

SetPrvt("oDlg1","oGrp1","oBrw1","oBtnVizu","oBtnAlt","oBtnCanc","oGet1","oCBoxPesq",/*"oSay1",*/"oSay2")  

oDlg1      := MSDialog():New( 175,774,676,1794,"Importação de arquivo .csv para criação de lote e lançamento de saldos iniciais.",,,.F.,,,,,,.T.,,,.F. )
oGrp1      := TGroup():New( 028,008,212,496,"LOTE - SALDOS INICIAIS",oDlg1,CLR_BLACK,CLR_WHITE,.T.,.F. )
MsAguarde({|| oTblPR()}, "Aguarde", "Localizando registros...") 
DbSelectArea(cFile)
(cFile)->(DbGotop())

Aadd( aCampos , {"CODPRODUTO"   ,"","Cod.Produto"   ,""} )
Aadd( aCampos , {"ARMAZEM"      ,"","Armazem"   	,""} )
Aadd( aCampos , {"LOTE"         ,"","Lote"	        ,""} )
Aadd( aCampos , {"QTD"          ,"","Quantidade"    ,""} )   
Aadd( aCampos , {"DTFABRIC" 	,"","Dt.Fabric"     ,""} )   
Aadd( aCampos , {"DTVENCTO"     ,"","Dt.Vencto"     ,""} )   
Aadd( aCampos , {"TPCONTROLE"   ,"","Tp.Controle"   ,""} )   
Aadd( aCampos , {"ANOS"         ,"","Anos"          ,""} )  


oBrw1 := MsSelect():New( cFile,"","",aCampos,.F.,,{038,012,206,492},,, oGrp1 ) 

oBtnImp   := TButton():New( 220,412,"Importar",oDlg1,,037,012,,,,.T.,,"",,,,.F. )
oBtnImp:bAction := {||Processa({|| Iif(MsGyEsNo( "Deseja realmente importar arquivo Excel?", "Importar Sim/Não"),sfImport(),) }, "Aguarde...", "Processando Planilha criação de lote e lançamento de saldos iniciais....",.F.)}                                 

oBtnCanc   := TButton():New( 220,458,"Cancelar",oDlg1,,037,012,,,,.T.,,"",,,,.F. )
oBtnCanc:bAction := {|| Iif(MsGyEsNo( "Deseja Realmente Sair?", "Cancelar Sim/Não"),Close(oDlg1),)}                       	

oGet1      := TGet():New( 012,008,,oDlg1,145,008,'',,CLR_BLACK,CLR_WHITE,,,,.T.,"",,,.F.,.F.,,.F.,.F.,"","",,)  
oGet1:bSetGet 	 := {|u| If(PCount()>0,__cProduto:=u,__cProduto)}

oBtnBusc   := TButton():New( 012,160,"Buscar",oDlg1,,037,010,,,,.T.,,"",,,,.F. )
oBtnBusc:bAction := {|| MsAguarde({|| sfFilterReg(.T.)}, "Aguarde...", "Localizando registros do filtro...") }         

oDlg1:Activate(,,,.T.) 

nTotLin     := 0
nTotSKU     := 0
nTotFil	    := 0
aPR		    := {}
aLin	    := {}
__cProduto  := Space(999)
oTempTable:Delete()

Return()

/*/{Protheus.doc} sfFilterReg
    (long_description)
    @type  Static Function
    @author user
    @since 12/04/2025
    @version version
    @param param_name, param_type, param_descr
    @return return_var, return_type, return_description
    @example
    (examples)
    @see (links_or_references)
/*/
Static Function sfFilterReg(lFlag)

Local nVar := y := 0                       
Local cArquivo 	:= GetNextAlias()
Local cChave 	:= ""
Local aProduto 	:= {}

cChave := "CODPRODUTO"

aProduto := sToA(__cProduto,";") //quebra pesquisa em array
    
If !lFlag
	cFor := " CODPRODUTO <> '"+PADR(__cProduto,TAMSX3("D5_PRODUTO")[1])+"'	"    
	IndRegua(cFile,cArquivo,cChave,,cFor,"Aguarde selecionando registros....")     
	
	FErase(cArquivo + GetDbExtension())  // Deletando o arquivo
	FErase(cArquivo+OrdBagExt())                        
	
	DbSelectArea(cFile)
	(cFile)->(DbGotop())
	
	Return()
EndIf     

If Empty(Alltrim(__cProduto)) 
    
	cFor := " CODPRODUTO <> '"+PADR(__cProduto,TAMSX3("D5_PRODUTO")[1])+"'	"    
	IndRegua(cFile,cArquivo,cChave,,cFor,"Aguarde selecionando registros....")
                      
	FErase(cArquivo + GetDbExtension())  // Deletando o arquivo
	FErase(cArquivo+OrdBagExt())
    
    
	DbSelectArea(cFile)
	(cFile)->(DbGotop())
	
	Return()             	
EndIf

For y:=1 to len(aProduto)     
	nVar+=1
	If nVar == 1
		cFor := "CODPRODUTO = '"+PADR(aProduto[y][1],TAMSX3("D5_PRODUTO")[1])+"' " 
	Else
		cFor += " .OR. CODPRODUTO = '"+PADR(aProduto[y][1],TAMSX3("D5_PRODUTO")[1])+"' "
	EndIf
next y                          

IndRegua(cFile,cArquivo,cChave,,cFor,"Aguarde selecionando registros....")
DbSelectArea(cFile)
(cFile)->(DbGotop())   

Return 


/*/{Protheus.doc} sfVerEsp
    (long_description)
    @type  Static Function
    @author user
    @since 12/04/2025
    @version version
    @param param_name, param_type, param_descr
    @return return_var, return_type, return_description
    @example
    (examples)
    @see (links_or_references)
/*/
Static Function sfVerEsp(cQtdRep)

//Variavel Local de Controle
Local aCarc_Esp := {} 
Local cMensagem := ""
Local i := 0

//Imputa os Caracteres Especiais no Array de Controle
AADD(aCarc_Esp,{"!", "Exclamacao"})
AADD(aCarc_Esp,{"@", "Arroba"})
AADD(aCarc_Esp,{"#", "Sustenido"})
AADD(aCarc_Esp,{"$", "Cifrao"})
AADD(aCarc_Esp,{"%", "Porcentagem"})
AADD(aCarc_Esp,{"*", "Asterisco"})
AADD(aCarc_Esp,{"/", "Barra"})
AADD(aCarc_Esp,{"(", "Parentese"})
AADD(aCarc_Esp,{")", "Parentese"})
AADD(aCarc_Esp,{"+", "Mais"})
AADD(aCarc_Esp,{"¨", ""})
AADD(aCarc_Esp,{"=", "Igual"})
AADD(aCarc_Esp,{"~", "Til"})
AADD(aCarc_Esp,{"^", "Circunflexo"})
AADD(aCarc_Esp,{"]", "Chave"})
AADD(aCarc_Esp,{"[", "Chave"}) 
AADD(aCarc_Esp,{"{", "Colchete"})
AADD(aCarc_Esp,{"}", "Colchete"})
AADD(aCarc_Esp,{";", "Ponto e Virgula"})
AADD(aCarc_Esp,{":", "Dois Pontos"})
AADD(aCarc_Esp,{">", "Maior"})
AADD(aCarc_Esp,{"<", "Menor"})
AADD(aCarc_Esp,{"?", "Interrogacao"})
AADD(aCarc_Esp,{"_", "Underline"})
AADD(aCarc_Esp,{",", "Virgula"})
AADD(aCarc_Esp,{".", "Ponto"})
AADD(aCarc_Esp,{" ", "Espaco"})
AADD(aCarc_Esp,{"-", "Traço"})
AADD(aCarc_Esp,{"'", "Aspas"})
AADD(aCarc_Esp,{"a", "Letra"})
AADD(aCarc_Esp,{"b", "Letra"})
AADD(aCarc_Esp,{"c", "Letra"})
AADD(aCarc_Esp,{"d", "Letra"})
AADD(aCarc_Esp,{"e", "Letra"})
AADD(aCarc_Esp,{"f", "Letra"})
AADD(aCarc_Esp,{"g", "Letra"})
AADD(aCarc_Esp,{"h", "Letra"})
AADD(aCarc_Esp,{"i", "Letra"})
AADD(aCarc_Esp,{"j", "Letra"}) 
AADD(aCarc_Esp,{"k", "Letra"})
AADD(aCarc_Esp,{"l", "Letra"}) 
AADD(aCarc_Esp,{"m", "Letra"})
AADD(aCarc_Esp,{"n", "Letra"})
AADD(aCarc_Esp,{"o", "Letra"})
AADD(aCarc_Esp,{"p", "Letra"})
AADD(aCarc_Esp,{"q", "Letra"})
AADD(aCarc_Esp,{"r", "Letra"})
AADD(aCarc_Esp,{"s", "Letra"})
AADD(aCarc_Esp,{"t", "Letra"})
AADD(aCarc_Esp,{"u", "Letra"})
AADD(aCarc_Esp,{"v", "Letra"})
AADD(aCarc_Esp,{"w", "Letra"})
AADD(aCarc_Esp,{"x", "Letra"})
AADD(aCarc_Esp,{"y", "Letra"})
AADD(aCarc_Esp,{"z", "Letra"}) 

//Executa o Laco ate o Tamanho Total do Array
For i:= 1 to Len(aCarc_Esp)
     //Verifica se Algum dos Caracteres Especiais foi Digitado
     If At(aCarc_Esp[i][1], LOWER(AllTrim(cQtdRep))) <> 0
          cMensagem := "Não é Permitido o Caracter " + aCarc_Esp[i][1] + " (" + aCarc_Esp[i][2]+ ") na coluna que informa Quantidade de Reposição!"+cEOL
     EndIf 
     
     If !Empty(cMensagem) 
		MsgStop(cMensagem+"Favor, Corrija Planilha antes mesmo de importa-la!")
        Return .F.    
     EndIf            
Next                     

Return .T. 

/*/{Protheus.doc} nomeStaticFunction
    (long_description)
    @type  Static Function
    @author user
    @since 12/04/2025
    @version version
    @param param_name, param_type, param_descr
    @return return_var, return_type, return_description
    @example
    (examples)
    @see (links_or_references)
/*/
Static Function oTblPR()

Local aProdutos := aPR
Local y			:= 0
Local aFds		:= {}

Aadd( aFds , {"CODPRODUTO"  ,"C",015,000} )
Aadd( aFds , {"ARMAZEM"	    ,"C",002,000} )
Aadd( aFds , {"LOTE" 	    ,"C",010,000} )
Aadd( aFds , {"QTD"	        ,"N",014,000} )
Aadd( aFds , {"DTFABRIC"	,"D",008,000} )
Aadd( aFds , {"DTVENCTO"	,"D",008,000} )
Aadd( aFds , {"TPCONTROLE"	,"D",008,000} )
Aadd( aFds , {"ANOS"	    ,"C",008,000} )

oTempTable := FWTemporaryTable():New( cFile )
oTempTable:SetFields(aFds) 
oTempTable:Create()

DbSelectArea(cFile)

For y:=1 to Len(aProdutos)
	(cFile)->(DbAppend())	 
	(cFile)->CODPRODUTO	    := aProdutos[y][1]       //COD_PRODUTO
	(cFile)->ARMAZEM	    := aProdutos[y][2]       //ARMAZEM
	(cFile)->LOTE	        := aProdutos[y][3]       //LOTE
	(cFile)->QTD            := val(aProdutos[y][4])  //QTD                           
	(cFile)->DTFABRIC 		:= ctod(aProdutos[y][5]) //DT_FABRIC  
	(cFile)->DTVENCTO 	    := ctod(aProdutos[y][6]) //DT_VENCTO  
    (cFile)->TPCONTROLE 	:= ctod(aProdutos[y][7]) //TIPO_CONTROLE  
    (cFile)->ANOS 	        := aProdutos[y][8]       //ANOS  
Next y

(cFile)->(DbGotop())

oDlg1:Refresh()

Return()

/*/{Protheus.doc} sfImport
    (long_description)
    @type  Static Function
    @author user
    @since 12/04/2025
    @version version
    @param param_name, param_type, param_descr
    @return return_var, return_type, return_description
    @example
    (examples)
    @see (links_or_references)
/*/
Static Function sfImport()
  
Local aLote     := {}
Local nOpc      := 3 // Inclusao
Local lRet      := .T.

Private lMsErroAuto := .F.
Private lMSHelpAuto := .F.

(cFile)->(DbGotop())

//Iniciando transação e executando saldos iniciais
Begin Transaction
    
    lRet := sfCriaSB9()
    
    If lRet 
        While (cFile)->(!EoF()) 

            lMsErroAuto := .F.
            lMSHelpAuto := .F.   
        
            aLote := {  {"D5_PRODUTO"  ,(cFile)->CODPRODUTO ,nil} ,;
                        {"D5_DOC"      ,"IMP"               ,nil} ,;
                        {"D5_LOCAL"    ,(cFile)->ARMAZEM    ,nil} ,;
                        {"D5_DATA"     ,dtos(dDatabase)     ,nil} ,;
                        {"D5_QUANT"    ,(cFile)->QTD        ,nil} ,;
                        {"D5_LOTECTL"  ,(cFile)->LOTE       ,nil} ,;
                        {"D5_DTVALID"  ,(cFile)->DTVENCTO   ,nil}}
                    
            MSExecAuto({|x, y| Mata390(x,y)}, aLote, nOpc)
        
            //Se houve erro, mostra mensagem
            If lMsErroAuto
                MostraErro()
                DisarmTransaction()

                lRet := .F.

                exit
            EndIf

            (cFile)->(DbSkip())

        EndDo

    EndIf

End Transaction

if lRet
    FWAlertSuccess("Processo de importação finalizado com sucesso.", "Totvs")
EndIf

(cFile)->(dbCloseArea())

Close(oDlg1)    

Return Nil


/*/{Protheus.doc} sfCriaSB9
    (long_description)
    @type  Static Function
    @author user
    @since 13/04/2025
    @version version
    @param param_name, param_type, param_descr
    @return return_var, return_type, return_description
    @example
    (examples)
    @see (links_or_references)
/*/
Static Function sfCriaSB9()

Local lRet      := .T.
Local aSldIni   := {}
Local nOpc      := 3 // Inclusao
Local n         := 0

For n := 1 to Len(aSldProd)
    
    aSldIni := {{"B9_FILIAL", FWxFilial("SB9")                                          , Nil},;
                {"B9_COD"   , SubStr( aSldProd[n][1], 1, aSldProd[n][3])                , Nil},;
                {"B9_LOCAL" , SubStr( aSldProd[n][1], aSldProd[n][3]+1, aSldProd[n][4]) , Nil},;
                {"B9_QINI"  , aSldProd[n][2]                                            , Nil} }

    MSExecAuto({|x, y| mata220(x, y)}, aSldIni, nOpc)

    //Se houve erro, mostra mensagem
    If lMsErroAuto
        MostraErro()
        DisarmTransaction()
        
        lRet := .F.
        exit
    //Else 
    //    aSldProd := {} //zera array para não repetir o mesmo produto.
    EndIf

Next n

Return lRet

/*/{Protheus.doc} nomeStaticFunction
    (long_description)
    @type  Static Function
    @author user
    @since 14/04/2025
    @version version
    @param param_name, param_type, param_descr
    @return return_var, return_type, return_description
    @example
    (examples)
    @see (links_or_references)
/*/
Static Function sToA(cString,cCter)

Local i
Local __aProd := {}
cString :=Alltrim(cString)
     
While At(cCter,cString) > 0          
	i := At(cCter,cString)     
	aAdd(__aProd,{substr(cString,1,i-1)})
	cString := substr(cString,i+1,len(cString)-i)               
enddo       
aAdd(__aProd,{cString})     
     
Return __aProd   
