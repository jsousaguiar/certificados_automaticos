# CERTIFICADOS AUTOMÁTICOS
Este programa gera certificados automáticos no formato ```.pdf``` a partir de um modelo ```.docx``` e uma lista de inscritos no formato ```.xlsx```.

A planilha com o nome dos inscritos, deverá ser a primeira da esquerda para a direita e o nome das colunas devem ser configurados no arquivo ```config.hjson```.

O modelo do certificado deverá ser editado para inserir as informações de sua empresa.

Todas as informações personalizadas como nome da empresa, nome do curso, data de início e fim devem ser editadas no arquivo ```config.hjson```, a exceção do logo da empresa, que deve ser inserido diretamente no certificado modelo ("modelo_certificado.docx").

Para inserir o CPF no certificado, basta colocar a chave ```{{ cpf }}``` no local desejado.

Antes de executar o programa, faça as configurações em ```config.hjson``` e instale os requerimentos digitando o comando ```pip install -r requeriments.txt``` no terminal.


<img width="561" alt="Screenshot_2" src="https://user-images.githubusercontent.com/68362578/132575581-27efae72-0569-456b-8aee-25c06edc50b5.png"> <img width="561" alt="Screenshot_3" src="https://user-images.githubusercontent.com/68362578/132576312-7c1b3c32-cf90-4739-98e6-145dfc007a39.png">



