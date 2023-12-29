# README

Script desenvolvido em PYTHON, que tem a seguinte finalidade
Atualmente identificamos que os relatórios diários enviados pelo Bitdefender é de dificil visualização já que é enviar em .CSV e se faz necessária a configuração

Como este relatório é enviado a gestores, então pensamos em uma forma de melhorar e automatizar essas informções.

Desta forma o script faz:
- Faz a deleção de arquivos baixados anteriormente
- Faz a leitura de uma caixa de e-mail para verificar os emails enviados no dia
- Baixa os anexos
- Envia as informações para um servidor Elastic para tratativa e visualização de relatórios
