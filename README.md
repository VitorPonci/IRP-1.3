# IRP-1.3
Macro VBA utilizada para tratamento de dados do indicador 1.3 do PMCRP 
IRP 1.3 – Consolidação e Cálculo da Proporção da Produção por Estado (PMCRP)
Descrição Geral

Este repositório contém a macro em VBA (Visual Basic for Applications) desenvolvida no âmbito do Programa Macrorregional de Caracterização de Rendas Petrolíferas (PMCRP) para o processamento do Indicador IRP 1.3, que mensura a participação da produção de petróleo e gás natural dos Estados de Santa Catarina, São Paulo, Rio de Janeiro e Espírito Santo em relação à produção nacional total.

A macro tem finalidade documental, metodológica e informativa, sendo disponibilizada publicamente para reforçar a transparência, rastreabilidade e padronização dos procedimentos adotados pelo PMCRP.

Indicador Associado

IRP 1.3 – Proporção da produção representada pelos Estados de Santa Catarina, São Paulo, Rio de Janeiro e Espírito Santo

O indicador é calculado conforme a expressão:

𝐼
𝑅
𝑃
1.3
=
∑
𝑛
𝑃
𝐸
𝑛
𝐼
𝑅
𝑃
1.1
IRP1.3=
IRP1.1
∑
n
	​

PE
n
	​

	​


Onde:

𝑃
𝐸
𝑛
PE
n
	​

 corresponde à produção anual de petróleo e gás natural de cada estado considerado;

𝐼
𝑅
𝑃
1.1
IRP1.1 representa a produção nacional total anual de petróleo e gás natural.

Fonte dos Dados

Fonte primária: Agência Nacional do Petróleo, Gás Natural e Biocombustíveis (ANP)

Publicação: Boletim Mensal da Produção de Petróleo e Gás Natural

Edição utilizada: Dezembro de cada ano (dados consolidados anuais)

Tabela de origem:
Distribuição da produção de petróleo e gás natural por estado

Os dados são utilizados exatamente conforme publicados pela ANP, sem estimativas externas ou ajustes estatísticos independentes.

Estrutura Esperada da Planilha
Aba de origem (entrada de dados)

Nome da aba: Produção por Estado

Coluna	Conteúdo
A	Ano
B	Estado
E	Produção total (petróleo + gás natural)

⚠️ A macro pressupõe que cada bloco anual contenha uma linha identificada como “Total Geral”, conforme o padrão da ANP.

Aba de destino (resultados)

Nome da aba: Graficos

A macro preenche automaticamente as seguintes colunas:

Coluna	Conteúdo
C	Ano
D	Produção Total Nacional (boe/d)
E	Produção Total Nacional (boe)
F	Produção Total PMCRP (boe)
G	Produção dos Estados PMCRP (boe)
H	Proporção PMCRP (%)
I	Proporção Rio de Janeiro (%)
J	Proporção Espírito Santo (%)
K	Proporção São Paulo (%)
L	Proporção Santa Catarina (%)
Funcionamento da Macro
Macro: AtualizarProporcaoPMCRP

A macro executa, de forma automatizada, as seguintes etapas:

Identifica os blocos anuais de dados na aba Produção por Estado;

Soma a produção anual dos Estados:

Rio de Janeiro

Espírito Santo

São Paulo

Santa Catarina

Identifica o valor da produção nacional total a partir da linha “Total Geral”;

Calcula:

o volume total anual dos estados do PMCRP (boe);

a proporção desse volume em relação à produção nacional;

as proporções individuais de cada estado;

Preenche automaticamente a aba Graficos, respeitando:

padronização de colunas;

formatação numérica (volumes e percentuais);

consistência temporal da série histórica.

Premissas Metodológicas

Utilização exclusiva dos dados oficiais da ANP;

Consideração apenas da edição de dezembro de cada ano;

Consolidação anual da produção;

Integração metodológica com o indicador IRP 1.1, que atua como denominador;

Ausência de ajustes probabilísticos, inferências externas ou reprocessamento estatístico.

Limitações

A macro depende da estrutura específica dos arquivos internos do PMCRP;

A reprodução integral dos resultados por terceiros não é garantida, em razão da ausência de acesso aos mesmos arquivos-base;

Eventuais revisões dos dados pela ANP podem alterar resultados históricos.

Finalidade do Repositório

Este código é disponibilizado com finalidade documental e informativa, permitindo:

auditoria metodológica;

compreensão do processo de cálculo do IRP 1.3;

replicação conceitual da metodologia em outros contextos institucionais.

Autor

Nascimento, Vitor Luiz Ponciano
Programa Macrorregional de Caracterização de Rendas Petrolíferas – PMCRP
FIA/USP
2025
