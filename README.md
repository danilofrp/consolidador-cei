# Consolidação de extratos do CEI para imposto de renda (IRPF)

Estes script visa facilitar a declaração do imposto de renda da pessoa física (IRPF), através da consolidação dos extratos gerados pelo Canal Eletrônico do Investidor (CEI), da B3. Com geração automática de reports como declaração de bens e posição na data desejada, deixar suas Ações, Opções e FIIs em dia com o leão fica muito mais fácil. O script também retorna um histórico de toda as compras e vendas do ativo, para acompanhamento e discriminação de maneira simples.

### Requisitos:
 - [Python 3.7.x](https://www.python.org)
 - [pandas](https://pandas.pydata.org)

### Preparo:
Os extratos de negociação do [CEI](https://cei.b3.com.br) (Etratos e Informativos > Negociação de ativos) devem estar em formato excel, localizados na pasta `extratos_cei`, e seguir o seguinte padrão de nomeclatura: `[ano]_negociacoes_cei_[corretora].xls` (Ex.: 2019_negociacoes_cei_clear.xls). Não deve haver sobreposição entre as datas dos arquivos, isto é, dados da mesma transição que se encontrem em dois arquivos diferentes serão considerados como duas transações iguais. A recomendação é que se gere um único arquivo por ao/corretora, de 1º de Janeiro a 31 de Dezembro.

### Modo de uso:
No momento, há dois modos de uso possíveis:

- Geração de relatório de posição em data à escolha.
```
python process_transactions.py --posicao yyyy-mm-dd
```

- Geração de relatório de posição no ano base do imposto de renda e no ano anterior (para a declaração de bens e direitos). Esta planilha também inclui o lucro/prejuízo realizado mensalmente, no ano base.
```
python process_transactions.py --declaracao yyyy
```

Adicionalmente, pode-se criar uma planilha consolidando todos os extratos em um único arquivo:
```
python consolidate_cei.py
```

#### Exemplo de uso:
Na pasta `extratos_cei` se encontram duas planilhas, `2018_negociacoes_cei_clear.xls` e `2019_negociacoes_cei_clear.xls`. O conteúdo destas planilhas pode ser visto abaixo:

**2018_negociacoes_cei_clear.xls**<br>
![2018_negociacoes_cei_clear.xls](https://github.com/danilofrp/consolidador-cei/blob/master/img/2018_extrato_cei_clear.png "2018_negociacoes_cei_clear.xls")

**2019_negociacoes_cei_clear.xls**<br>
![2019_negociacoes_cei_clear.xls](https://github.com/danilofrp/consolidador-cei/blob/master/img/2019_extrato_cei_clear.png "2019_negociacoes_cei_clear.xls")

Ao rodar o comando
```
python process_transactions.py --posicao 2020-01-01
```
o script nos dá todas as negociações feitas e a posição até no dia 01/01/2020. Estas informaçoes ficam salvas na planilha `posicoes_2020-01-01.xlsx`, que pode ser vista abaixo:

**posicoes_2020-01-01.xlsx**<br>
![posicoes_2020-01-01.xlsx](https://github.com/danilofrp/consolidador-cei/blob/master/img/posicao.png "posicoes_2020-01-01.xlsx")


Rodando o comando
```
python process_transactions.py --declaracao 2019
```
o script consolida informações compra e venda de ações para imposto de renda, nos dando todas as posições em 31 de dezembro de 2018 e de 2019, e também o lucro/prejuízo realizado a cada mês, em diferentes abas. Note que, nesta planilha, o histórico mostra as transações **apenas** no ano base. Estas informaçoes ficam salvas na planilha `declaracao_2019.xlsx`, que pode ser vista abaixo:

**declaracao_2019.xlsx -> Declaração de bens**<br>
![declaracao_2019.xlsx -> Declaração de bens](https://github.com/danilofrp/consolidador-cei/blob/master/img/declaracao.png "declaracao_2019.xlsx -> Declaração de bens")

**declaracao_2019.xlsx -> Lucro Realizado**<br>
![declaracao_2019.xlsx -> Lucro Realizado](https://github.com/danilofrp/consolidador-cei/blob/master/img/realizado.png "declaracao_2019.xlsx -> Lucro Realizado")


### Melhorias previstas
Atualmente, o script não processa as seguintes informações:

- Exercício de opções
- Vencimento de opções sem exercício

A compra e venda de opções é processada normalmente. O desenvolvimento destas e outras funcionalidades é esperado no futuro


### Colaboração

Pull requests com melhorias e novas funcionalidades são sempre bem vindos!
