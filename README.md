# Consolidação de extratos do CEI para imposto de renda (IRPF)

Estes script visa facilitar a declaração do imposto de renda da pessoa física (IRPF), através da consolidação dos extratos gerados pelo Canal Eletrônico do Investidor (CEI), da B3. Com geração automática de reports como declaração de bens e posição na data desejada, deixar suas Ações, Opções e FIIs em dia com o leão fica muito mais simples.

### Requisitos:
 - [Python 3.7.x](https://www.python.org)
 - [pandas](https://pandas.pydata.org)

### Condições:
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

### Melhorias previstas
Atualmente, o script não processa as seguintes informações:

- Exercício de opções
- Vencimento de opções sem exercício

O desenvolvimento destas e outras funcionalidades é esperado no futuro
