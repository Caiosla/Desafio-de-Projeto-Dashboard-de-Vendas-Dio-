# Desafio-de-Projeto-Dashboard-de-Vendas-Dio-
Dashboard de Vendas — Excel Online
Dashboard de vendas construído do zero no Excel Online, sem uso de macros ou ferramentas pagas. O projeto organiza dados brutos de vendas em visualizações interativas com métricas automáticas.

 Estrutura do Projeto
O arquivo Excel é dividido em 3 abas:
Aba        Função
Dados      Inserção dos registros de venda
Calculos   Tabelas dinâmicas e Slicers de filtro
Dashboard  Visualização final com cards e gráficos


 Dados Utilizados
Os dados são fictícios e foram criados manualmente para fins de demonstração.
Colunas da aba Dados:
Campo       Descrição
Data        Data da venda
Vendedor    Nome do vendedor responsável
Produto     Nome do produto vendido
Categoria   Categoria do produto
Quantidade  Unidades vendidas
Preço Unit  Preço unitário do produto
Total       Calculado automaticamente (Quantidade × Preço)

Vendedores: Ana Souza, Carlos Lima, Juliana Reis, Pedro Alves
Categorias: Informática, Periféricos, Áudio, Armazenamento, Eletrônicos, Móveis, Redes
Período: Maio de 2026

 Componentes do Dashboard
Cards de Métricas

Total de Vendas — soma geral de todas as vendas
Ticket Médio — total dividido pelo número de transações
Melhor Vendedor — vendedor com maior volume de vendas

Gráficos

Vendas por Vendedor — gráfico de colunas comparando desempenho individual
Vendas por Produto — gráfico de barras horizontais com ranking de produtos
Evolução das Vendas por Mês — gráfico de linha com tendência temporal

Filtros (Slicers)

Filtro por Vendedor
Filtro por Categoria


 Fórmulas Utilizadas
# Total de Vendas
=Calculos!H6

# Ticket Médio
=Calculos!H6/CONT.NÚM(Dados!E2:E21)

# Melhor Vendedor
=ÍNDICE(Calculos!G2:G5;CORRESP(MÁXIMO(Calculos!H2:H5);Calculos!H2:H5;0))

#Como Reproduzir
Requisitos

Excel Online (gratuito via office.com) ou Excel Desktop

Passo a Passo

Crie um novo arquivo Excel com 3 abas: Dados, Calculos, Dashboard
Na aba Dados, crie a tabela com as colunas:
Data | Vendedor | Produto | Categoria | Quantidade | Preço Unit. | Total

Converta o intervalo em Tabela oficial: Inserir → Tabela
A coluna Total usa a fórmula =Quantidade*PreçoUnit


Na aba Calculos, crie 3 Tabelas Dinâmicas apontando para a aba Dados:

Tabela 1: Vendedor × Valor das Vendas
Tabela 2: Produto × Valor dos Produtos
Tabela 3: Data × Valor arrecadado do dia


Adicione os Slicers clicando numa Tabela Dinâmica na aba Calculos → Análise → Inserir Segmentação de Dados

Crie um Slicer de Vendedor e um de Categoria
Os Slicers funcionam na aba Calculos — filtrar por vendedor ou categoria atualiza as tabelas dinâmicas
Limitação do Excel Online: não é possível mover Slicers entre abas nem conectá-los via Conexões de Relatório. Para isso, é necessário o Excel Desktop com Microsoft 365


Na aba Dashboard, insira os gráficos:

Clique em cada Tabela Dinâmica → Inserir Gráfico
Mova os gráficos para o Dashboard via Design → Mover Gráfico


Adicione os Cards de métricas com as fórmulas acima nas células desejadas
Formate o Dashboard: remova linhas de grade, adicione cores e título
Proteja as abas Calculos e Dashboard via botão direito → Proteger Planilha

Atualização dos Dados

Adicione novas linhas na aba Dados
Vá na aba Calculos e clique com botão direito → Atualizar em qualquer Tabela Dinâmica
O Dashboard atualiza automaticamente


 Limitações do Excel Online

Slicers não podem ser movidos entre abas (limitação da versão gratuita)
Algumas funcionalidades avançadas requerem Excel Desktop com assinatura Microsoft 365

Projeto desenvolvido como exercício prático de criação de dashboards no Excel.


Projeto desenvolvido como exercício prático de criação de dashboards no Excel.
