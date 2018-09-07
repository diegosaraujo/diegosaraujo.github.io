---
layout: post
title: Como abrir campos do tipo Hiperlinks em uma nova janela no SharePoint
date: 2015-11-23 11:45:00
description: Algo simples que não há nativamente :(
img: 
fig-caption: 
tags: [SharePoint, JavaScript]
---

### Cenário:
Um cliente solicitou a criação de uma lista no SharePoint com duas colunas, uma coluna do tipo texto com uma linha e outra do tipo hiperLink. A ideia dele era simples, a lista seria utilizada para armazenar links úteis para ele. Ao clicar na coluna do tipo hiperlink o link deverá ser aberto em uma nova aba do navegador.


Simples né? Sim, se não fosse pelo fato de nativamente o SharePoint não abrir os links dos campos do tipo hiperlink em uma página nova(target _blank).

<center><img src="/assets/img/HiperLinkSharePoint/im01.png" alt="Exemplo lista de Links" style="padding: 0; border: none !important; background:none; text-align:center;">
</center>
<center>*Exemplo da lista solicitada pelo cliente com um campo do tipo HiperLink*</center>

<br />
### O que fiz para solucionar?

Depois de quebrar um pouco a cabeça e realizar algumas pesquisas na internet achei uma solução bem interessante que me agradou bastante e ao cliente também.

Adicionei uma webpart do tipo `Editor de Script` e inseri o trecho de código JavaScript abaixo na página em que a lista de links úteis está.

Veja o código abaixo:


``` javascript
<script language="JavaScript">
_spBodyOnLoadFunctionNames.push("rewriteLinks");

function rewriteLinks() {
    var anchors = document.getElementsByTagName("a");

    for (var x=0; x<anchors.length; x++)
    {
         if (anchors[x].outerHTML.indexOf('#novajanela')>0)
         {
            oldText = anchors[x].outerHTML;
            newText = oldText.replace(/#novajanela/,'" target="_blank');
            anchors[x].outerHTML = newText;
         }
    }
}
</script>
```
<center> *Trecho de código JavaScript utilizado para realizar a abertura do link em _blank*</center>

Depois bastou adicionar um item normalamente e no campo "Endereço do Link"(exemplo da imagem abaixo) informar o link acompanhado de `#novajanela`. O javascript se encarregada de tudo :)

<br />
<center><img src="/assets/img/HiperLinkSharePoint/im02.png" alt="Adicionando a chave novajanela no endereço do link" style="padding: 0; border: none !important; background:none; text-align:center;">
</center>

<center> *Adicionando a chave novajanela no endereço do link* </center>
 
<br />
**Obs**: Se quiser alterar a "chave" fique à vontade, o código JavaScript nem é tão complicado de entender. Outro detalhe importante é que se no endereço do link não for inserida a "chave" o link será aberta na própria página e não em uma nova, ou seja, o JavaScript só irá funcionar nos itens que tiverem a "chave".


A lista é exibida normalmente no SharePoint. A diferença só é notada quando o usuário clica no item da lista.

Espero que tenha entendido. Qualquer dúvida ou caso queira contribuir com o post deixe seu comentário abaixo.

Forte abraço e até a próxima!!!
