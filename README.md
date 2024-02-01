<div align="center">
  <h1> :rocket: Bem vindo ao meu repositório:rocket: </h1>
</div>

<br>
<br>

<div>
  <h1> Minhas redes sociais</h1>
  <a href="https://www.youtube.com/channel/UC88QEmxaSyY_V2vXn1RMgQQ" target="_blank"><img src="https://img.shields.io/badge/YouTube-FF0000?style=for-the-badge&logo=youtube&logoColor=white" target="_blank"></a>
<a href="https://www.instagram.com/_anthonny_michael_dev/" target="_blank"><img src="https://img.shields.io/badge/-Instagram-%23E4405F?style=for-the-badge&logo=instagram&logoColor=white" target="_blank"></a>
<a href="https://www.linkedin.com/in/anthonny-michael-64450a206/" target="_blank"><img src="https://img.shields.io/badge/-LinkedIn-%230077B5?style=for-the-badge&logo=linkedin&logoColor=white" target="_blank"></a> 
</div>



# Como baixar esse repositório? :sassy_man:

1. Baixando diretamente pelo github.

    <img src="Email-Python/readme/Github Download Repo.png" />

2.  Baixando pelo "gitclone" no seu prompt de comando. (Sintaxe: git clone "https://www.site.com")

    <img src="Email-Python/readme/Git clone.png" />
    
<br>

# O que é esse repositório?
Envio de e-mails com leitura de arquivo excel. 


# Email-Python Aplicação em Python. 

Sua funcionalidade é ler um arquivo excel, e formalizar um e-mail e realizar seu envio com os dados e critérios de formatação da mensagem no código. Sua lógica é fazer a leitura de um arquivo excel e montar um dataframe através da biblioteca Pandas, onde faz o agrupamento de colunas e linhas e cria um dataframe separado com essa linha e coluna que especifiquei no código e realiza uma contagem de quantas vezes cada nome de supervisor e coordenador se repetem dentro do arquivo excel e salva em uma variável a contagem e o dataframe agrupado com as linhas e colunas especificadas. A partir disso, cria um corpo de e-mail no formato de HTML usando a plataforma Gmail, e trás as variáveis no corpo do e-mail com as contagens de cada nome de supervisor e coordenador e sua contagem de repetição dentro do arquivo excel e envia um e-mail para os coordenadores de operação e o gerente de operação, com intuito de relatório de acompanhamento. 

A aplicação também faz uma leitura do arquivo excel para não trazer repetidamente os nomes de operador, supervisor e coordenador, sua lógica por trás faz com que cada pessoa receba somente um e-mail com todas as suas monitorias pendentes de assinatura e/ou aplicação de feedback. O supervisor recebe somente um e-mail com todas as suas monitorias pendentes no arquivo excel em forma de tabela no corpo do e-mail, onde somente um e-mail é disparado para o supervisor e no corpo do e-mail vem uma tabela, que é a mesma do arquivo excel, mas detalhe, eu precisei fazer em uma parte do código, algumas linhas, em que trazesse somente as suas pendências consolidadas em um dataframe só, para não repetir indevidamente a contagem de cada supervisor e afins, valendo também para coordenador e operador. 

O operador recebe o mesmo e-mail no intuito do supervisor, o código envia somente um e-mail para o operador com todas as monitorias pendentes para o mesmo dentro do arquivo excel, que no corpo do e-mail é um dataframe que criei separadamente com o agrupamento de linhas e colunas do meu arquivo em excel lido com Pandas, e com isso, somente adicionei a contagem de cada nome presente do arquivo excel. Formatei as linhas "Coordenador" e "Quantidade de monitorias" separadamente do meu dataframe do meu corpo de e-mail enviado ao operador, supervisor e coordenador e gerente, para poder estilizar com propriedades do CSS3, já que é um corpo de e-mail no formato de HTML. 

Com isso, juntando tudo, unificando e contando os valores, trago as variáveis do meu dataframe personalizado com somente os dados que preciso e com a contagem lado, no e-mail, monto o corpo do e-mail e coloco as variáveis dentro e realizo o envio para todos, uma única vez. 

# OBS: Email-Python 

Esse código pode ser alterado para qualquer realidade de qualquer empresa, seja para um relatório que todos os funcionários precisam acompanhar diariamente, desde cobranças e dentre outros motivos e interesses. 

# Como funciona o disparo de e-mails? 

1. <strong>Segue imagem de exemplo de e-mail enviado para o operador, notificando-o da pendência de monitoria de assinatura na ferramenta de monitoria da empresa de telemarketing:</strong>
<img src="/Email-Python/readme/operador.png" />
<br>

2. <strong>Segue imagem de exemplo de e-mail enviado para o supervisor, notificando-o da pendência de monitoria de aplicação de feedback na ferramenta de monitoria da empresa de telemarketing:</strong>
<img src="/Email-Python/readme/supervisor.png" />
<br>

3. <strong>Segue imagem de exemplo de e-mail enviado para os coordenadores e gerente, enviando um e-mail com contagem de monitorias pendentes de assinatura e feedback por supervisor e coordenador para acompanhamento:</strong>
<br>
<img src="/Email-Python/readme/assinatura.png" />
<br>
<img src="/Email-Python/readme/feedback.png" />

