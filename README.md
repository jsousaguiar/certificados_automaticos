# CERTIFICADOS AUTOMÁTICOS
Este programa gera certificados automáticos no formato ```.pdf``` a partir de um modelo ```.docx``` e uma lista de inscritos no formato ```.xlsx```.

A planilha com o nome dos inscritos, deverá ser a primeira da esquerda para a direita e o nome das colunas devem ser configurados no arquivo ```config.hjson```.

O modelo do certificado deverá ser editado para inserir as informações de sua empresa.

Todas as informações personalizadas como nome da empresa, nome do curso, data de início e fim devem ser editadas no arquivo ```config.hjson```, a exceção do logo da empresa, que deve ser inserido diretamente no certificado modelo ("modelo_certificado.docx").

Retirar ou mudar o local do CPF certificado, basta apagar colocar a chave ```{{ cpf }}``` no local desejado.

Antes de executar o programa, faça as configurações em ```config.hjson``` e instale os requerimentos digitando o comando ```pip install -r requeriments.txt``` no terminal.



