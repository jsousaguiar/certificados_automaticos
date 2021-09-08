# CERTIFICADOS AUTOMÁTICOS
Este programa gera certificados automáticos no formato ```.pdf```.

A lista de inscritos deve estar no formato ```.xlsx```.

A planilha com o nome dos inscritos, deverá ser a primeira da esquerda para a direita e o nome das colunas devem ser configurados em ```config.hjson```.

O modelo do certificado deverá ser editado para inserir as informações de sua empresa.

Todas as informações personalizadas como nome da empresa, nome do curso, data de início e fim devem ser editadas no arquivo ```config.hjson```.

Para inserir o CPF no certificado, basta colocar a chave ```{{ cpf }}``` no local desejado.

Antes de executar o programa, faça as configurações em ```config.hjson``` e instale os requerimentos digitando o comando ```pip install -r requeriments.txt``` no terminal.



