---
layout: post
title: Integração MS Project e SharePoint
date: 2016-01-08 16:31:00
description: Mapeando campos do MS Project no SharePoint
img: 
fig-caption: 
tags: [SharePoint, Microsoft Project]
---

Hoje eu vou falar sobre algo bem simples. Todos nós sabemos que o SharePoint se integra com o MS Project e com qualquer outro produto da família Office. O grande X da questão é saber como explorar esta integração entre os produtos e a plataforma SharePoint.

O exemplo que vou mencionar aqui veio de um problema vivido por um cliente.


### Cenário

O cliente possui uma lista de tarefas do SharePoint, esta lista é preenchida pelo cliente através do MS Project.
O cliente abre a lista com o Project, preenche as tarefas do projeto e depois sincroniza a lista preenchida com o SharePoint. Até aí nada de mais...

O problema é que há um coluna chamada **Duração** que não está visivel no SharePoint e por este simples motivo o cliente acredita que não há integração entre o Project e o SharePoint.


### Solução
 Para a coluna em questão ou qualquer outra coluna que seja adicionada no Project, é necessário mapear o campo para que ele fique visível no SharePoint. Acompanhe o post e veja como é simples realizar este mapeamento.

#### Passo 1
 Crie uma lista do tipo tarefa. No meu caso eu estarei utilizando uma já existente chamada **Cronograma**.

<center><img src="/assets/img/MSProject/img01.jpg" alt="Lista de tarefas" ></center>


#### Passo 2

Selecione a opção **Lista** menu ribbon localizado no topo do site e depois clique na opção **Abrir no Project**.

<center><img src="/assets/img/MSProject/Img02.jpg" alt="Abrindo a lista no MS Project" ></center>


#### Passo 3

Toda vez que você tentar abrir um documento que está no SharePoint com qualquer produto Office será necessário informar as suas credencias de acesso ao SharePoint, mesmo que você já esteja logado.

Está é uma politica de segurança que verifica se você possui permissão de acesso ao documento e verifica qual é o seu nível de acesso. Insira o seu login e senha para continuar.

<center><img src="/assets/img/MSProject/Img03.jpg" alt="Janela de credenciais de acesso" ></center>


#### Passo 4

No meu exemplo eu não irei adicionar nenhuma coluna nova no meu doc. do Project, pois a coluna **Duração** não é mapeada no SharePoint.
Caso queira, você pode adicionar uma nova coluna para realizar o passo a passo.

<center><img src="/assets/img/MSProject/Img04.jpg" alt="Colunas padrão do Project" ></center>

Preencha o documento do MS Project com o cronograma do projeto ou uma informação fake só para realizar o teste.

Se você salvar o documento preenchido irá perceber que os dados foram sincronizados com o SharePoint mas a coluna Duração não aparece na view do SharePoint.

<center><img src="/assets/img/MSProject/Img05.jpg" alt="Exemplo de um doc. Project" ></center>


#### Passo 5

O próximo passo é mapear os campos (colunas) que não existem no SharePoint, neste nosso exemplo é o **Duração**.

Clique na opção **Arquivo**.

<center><img src="/assets/img/MSProject/Img06.jpg" alt="Opção Arquivo do menu ribbon do MS Project" ></center>


Depois clique no **Mapear campos**.
<center><img src="/assets/img/MSProject/Img07.jpg" alt="Mapeando os campos do Project" ></center>


#### Passo 6

Sincronizar os campos do MS Project com o SharePoint. Clique no botão **Mapear Campo**.
<center><img src="/assets/img/MSProject/Img08.jpg" alt="Janela de mapeamento" ></center>


Selecione o campo (coluna) do Project que deseja soncronizar com o SharePoint. Neste exemplo é Duração e depois de um nome a coluna que será adicionada no SharePoint. O ideal é que a coluna do Project e do SharePoint tenham o mesmo nome.

<center><img src="/assets/img/MSProject/Img09.jpg" alt="Janela adicionar campos" ></center>

Veja que agora a coluna Duração aparece na janela Mapear Campos, e ela deverá estar marcada da mesma forma que na figura abaixo.

<center><img src="/assets/img/MSProject/Img10.jpg" alt="Janela mapear campos" ></center>


#### Passo 7
Salve o documento, aguarde o término da sincronização com o SharePoint e feche o documento.

**Obs:** Ao verificar a lista perceba que o campo Duração ainda não está visível (mas ele existe).

<center><img src="/assets/img/MSProject/Img11.jpg" alt="Lista do tipo tarefa do SharePoint" ></center>


#### Passo 8

Agora temos que adicionar a coluna **Duração** na nossa View Padrão da lista de tarefas para que ele apareça.

Acesse o menu ribbon e na aba **Lista** clique na opção **Configurações da Lista** na seção de configurações.

<center><img src="/assets/img/MSProject/Img12.jpg" alt="Acessando as configurações da biblioteca" ></center>

Vá até o final da página, na opção **Exibições (Views)** encontre a view que está definida como padrão e clique nela.

<center><img src="/assets/img/MSProject/Img13.jpg" alt="Lista de Exibições" ></center>

Na lista de colunas disponível, selecione a coluna Duração, indique a posição dele na view (caso ache necessário) e depois clique em **OK**.

<center><img src="/assets/img/MSProject/Img14.jpg" alt="Listas de colunas disponíveis" ></center>


#### Passo 9
Você será redirecionado para a lista de tarefas, verifique se a coluna Duração está disponível na view.

<center><img src="/assets/img/MSProject/Img15.jpg" alt="Lista de tarefas com o campo de duração" ></center>

Isto é tudo...caso tenha alguma dúvida ou sugestão deixe seu comentário. Obrigado e até a próxima.

Abraço!!
