- se nao exister BU contrata geral entao nao se tem que lidar com os users afetados, ou seja correr s� 1 e 4


- Nunca existem pedidos para fazer equipas NA, quando uma equipa NA n�o existe isso significa que � preciso fazer o parque, e quando se faz o parque faz-se logo a equipa de chefia

# Solved Issues

	- A cria��o da BU e da equipa proprietaria e equipa contrata j� verificam se cada um dos elementos j� existem no sistema, se sim n�o s�o criados em duplicado
	- O c�digo para criar equipas e BU's j� usa uma query que tem em conta o centor de emplazamento e o grupo planificador, a query antigar s� usava o centro e n�o encontrava o registo certo
	- O basic user retriever j� n�o tem hardcoded values e j� mostra at� 2 users para compara��o
	- O basica user retriever j� alinha corretamente os users lado a lado quando o 1� user tem uma equipa com o nome demasiado comprido
	- No normalizer so diz se deu o edpr inspections nao diz nada sobre a templates quando � para dar RESCO
	- Quando se est� a dar equipas no assign team, por checks para ter a certeza que n�o h� espa�os em branco a mais em lado nenhum
	- Quando estou a criar equipas o c�digo devolve multiplos centros de planifica��o, a query que est� a ser feita n�o � especifica o suficiente (update: a raz�o era que os nomes de alguns campos queried � bd estavam errados)
	- No create BU n�o est� a verificar se as queries para o grupo planificador e centro de trabalho devolvem 1 s� resultado v�lido
	- A cria��o das notifica��es est� com a logica errada, os 3 campos tem que estar dentro do "Y" e nao dentro de um OR
	- Notification & workorders view creator est�o a usar o codigo completo do posto de trabalho quando para a query funcionar so podem usar os 1�s 4 chars
	- Notification & workorders view creator est�o a mostrar os 4 chars de todos os codigos mesmo apos o distinct, isto porque o distinct � feito depois do slice
	- Notification & workorders view creator est�o a colocar cada grupo planificador e codigo de posto de trabalho na mesma linha da query quando o xrm funcionar melhor quando cada um tem a sua linha
	- No workorder e avisos ele coloca as condi��es de forma incorreta e portanto caso haja equipas diferentes com grupos planificadores oup
	- Mudado o normalizer para permitir configurar BU's EDPR mas avisar dos riscos
	- Mudar o user normalizer para quando o user escreve 0 n�o mostrar texto flash na consola e deixar na duvida se algo aconteceu que nao devia, � feito uma results>0 para verificar se algum user foi passado tanto para o normalizer como para o SAP credentials


# Unsolved Issues
	
	- No normalizer ele faz sempre update � regi�o sem verificar se o user j� tem a correta
	- O codigo para criar equipas s� faz equipas EU por agora
	- verificar se o ficheiro excel esta aberto ou a ser usado por  outro processo e recusar avan�ar (ja tentei demasiado dificil, no handholding here)
	- colocar uma biblioteca que apanha todos os logs e poe no ficheiro de logs, existe uma qualquer
	- quando copiar um user, deveria perguntar se quer tambem normalizar s� por percau��o
	- quando o codigo esta a tentar atribuir a equipa criada aos users e nao encontra a equipa ainda assim cria o excel sem nada, isso � redudante
	- quando atribuo no passo 3 uma equipa a um user ele diz algo como j� ter essa equipa mas depois diz que atribuiu, isso tem que ser visto e corrigido
	- na opcao 1 ele tem que validar se todos os campos na folha estao preenchidos, por exemplo eu vi quando nao tinha nada no contractor ele deixou prosseguir
	- temos que tirar do runflow workorders e aviso o hardcoded name do nome do workflow e por no ficheiro commons
	- o codigo que copia permisoes de um user para o outro, quando copia de um interno para EX nao devia deixar levar o role edpr personal interno nem equipas que digam interno ou pelo menos deve avisar, por agora so copia cegamente tudo
	- o codigo que copia permiss�es n�o lida bem com acentos
	- o usernormalizer quando o user tem a BU edpr tem que mostrar uma msg mais clara e quando o user nao tem E ou EX no nome a msg tem que ser mais clara
	- no modulo que copia teams, bu, etc, podia oferecer-se para copiar tambem a regi�o
	- no modulo de copiar os roles teams e etc os roles podiam estar por ordem alfabetica
	- o codigo do run new workflow tem demasiados console read keys, pedir para ser optmizado
	- colocar o nome do fluxo do novo user no codes and roles
	- dentro do user copier devia perguntar se o user quer chamar o normalizer