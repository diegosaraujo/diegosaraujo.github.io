---
layout: post
title: CSOM SharePoint
date: 2016-01-28 15:31:00
description: Operações básicas no SharePoint usando JavaScript
img: 
fig-caption: 
tags: [SharePoint, JavaScript]
---

### Por que o CSOM?

CSOM – Client Side Object Model. O CSOM é utilizado para executar operações simples no SharePoint, poupando assim muitas vezes o trabalho que temos para criar uma webpart customizada.
Algumas operações básicas são bem simples de fazer, basta ter conhecimento de JavaScript e deixar que a biblioteca de JS do SharePoint faça o resto.
O link que eu utilizei como referência para este artigo foi retirado do site da [MSDN].
Antes de começar eu gostaria de citar uma API JavaScript que utilizamos em um outro projeto aqui na empresa, [SharePoint Plus]. Funciona igualmente bem, e cabe a você escolher o que utilizar, mas vale dar uma olhada no projeto do SharePoint Plus. O projeto está no github.


### Cenário

Eu irei utilizar o CSOM para ler uma lista e exibir os 10 primeiros itens dela, certo? Coisa simples, só para entender o funcionamento.
Você pode até achar, que é bobeira ou para o que eu utilizaria isto, mas já te adianta que com essa brincadeira dá para fazer bastante coisa.


### Como implementar?
Em primeiro lugar precisaremos de uma lista personalizada para recuperar os dados.
Eu criei uma lista chamada Contatos e criei as seguintes colunas:  Empresa (do tipo texto simples), Status (do tipo escolha “choice” e adicionei duas opções: Ativo e Inativo) e Título (vou deixar padrão).

<center><img src="/assets/img/CSOM/img01.jpg" alt="Lista de Contatos" ></center>


Adicione alguns contatos na lista para vermos os resultados do teste.
O próximo passo é criar o código.

Crie um arquivo JS no seu editor preferido. E adicione o seguinte código.

**Obs**: O código Html e o código JS deverão estar no mesmo arquivo, pois o botão chamará a função javascript.

**Código HTML**
```html
<div>
<button onclick="exibeItensLista('Digite aqui a URL do seu Site SharePoint'); return false;">Exibe 10 itens da Lista</button>
</div>
<br />
<h2>Exibe os 10 primeiros itens da lista Contatos</h2>
<div id="exibeItens"></div>
```

<br />
**Código JavaScript**
```javascript
<script type="text/javascript">
	function exibeItensLista(siteUrl){
	var clientContext = new SP.ClientContext(siteUrl);
	var oList = clientContext.get_web().get_lists().getByTitle('Contatos');
	 var camlQuery = new SP.CamlQuery();
     	camlQuery.set_viewXml(
            '<View><Query><Where><Geq><FieldRef Name=\'ID\'/>' + '<Value Type=\'Number\'>1</Value></Geq></Where></Query>'
            + '<RowLimit>10</RowLimit></View>'
     		);
        		this.collListItem = oList.getItems(camlQuery);
		clientContext.load(collListItem);
		clientContext.executeQueryAsync(
		Function.createDelegate(this, this.onQuerySucceeded),
		Function.createDelegate(this, this.onQueryFailed)
		);
	}

	function onQuerySucceeded(sender, args){
		var listInfo = '';
		var listEnumerator = collListItem.getEnumerator();
		while (listEnumerator.moveNext()) {
			var itemAtual = listEnumerator.get_current();
			listInfo += '<br />ID: ' + itemAtual.get_id() + '<br /> Nome: ' + itemAtual.get_item('Title') + '<br /> Empresa: ' +
			itemAtual.get_item('Empresa') + '<br /> Status do Cliente: ' + itemAtual.get_item('Status') + '<br />' ;
		}

		document.getElementById('exibeItens').innerHTML = listInfo;
	}

	function onQueryFailed(sender, args){
	document.getElementById('exibeItens').innerHTML = '<br />Falha de Requisição' + args.get_message() + '<br />' + args.get_stackTrace();
	}
</script>
```

Dê uma olhada com calma no código, mas não é complicado de entender. Os detalhes que são válidos mencionar são:

**Html?** - Sim eu adicionei tags no início do código para que executássemos o teste.
Então repare que o HTML nada mais é que uma div com um botão que executa a minha função em JS (exibeItensLista) ao receber um clique. Após o clique há uma div chamada `ExibeItens` que recebe os itens da lista.
Bem simples.

**Coluna Title?** – Se você reparar a function onQuerySucceeded possui um while que faz o percurso na lista e retorna os itens. Para que os itens sejam guardados na variável `listInfo` eu preciso informar os nomes das colunas (ex: itemAtual.get_item('Title')). Aí você deve estar se perguntando, mas porquê campo `Title`? O meu SharePoint Server é inglês e eu utilizo um language pack para português. Basta ficar atento a isto, mas na dúvida crie uma outra coluna se achar necessário.

Detalhe “chato”, mas é o tipo de coisa que te faz bater cabeça.

**Query CAML** - Eu estou utilizando query CAML somente para realizar filtro e dar uma ideia do que pode ser feito, neste exemplo é filtrado somente os 10 primeiros itens, mas eu poderia filtra pelo Status do cliente.
Referência Query CAML [https://msdn.microsoft.com/pt-br/library/office/ms462365.aspx]
Agora com o arquivo JS criado basta importar ele para o SharePoint.

<center><img src="/assets/img/CSOM/img02.jpg" alt="Biblioteca que guarda o arquivo JS" ></center>

Eu adicionei o arquivo JS na pasta /SiteAssets/JS, mas você pode adicionar em qualquer outra pasta do SharePoint.
Para finalizar basta adicionar uma webpart do tipo `Editor de Conteúdo` na página que deseja exibir o CSOM.


<center><img src="/assets/img/CSOM/img03.jpg" alt="Adicionando a WebPart Editor de Conteúdo" ></center>

Depois clique em **Adicionar**.

Após adicionar o Editor de Conteúdo clique na seta no canto superior direito da webpart e clique na opção **Editar Web Part**.

<center><img src="/assets/img/CSOM/img04.jpg" alt="Editando a WePart" ></center>

No campo **Link de Conteúdo** adicione a url aonde está localizado o arquivo JS e depois clique em **OK**.

<center><img src="/assets/img/CSOM/img05.jpg" alt="Inserindo a URL no campo link de conteúdo" ></center>

Após finalizado o resultado deverá ser como o da imagem abaixo.

<center><img src="/assets/img/CSOM/img06.jpg" alt="WebPart adicionada à página" ></center>

Agora é só testar e ver os 10 primeiros itens da lista sendo apresentados.

<center><img src="/assets/img/CSOM/img07.jpg" alt="Testando o CSOM na página" ></center>

Valeu galera, caso tenha alguma dúvida ou sugestão deixe o seu comentário abaixo.
Abraço!!!

[https://msdn.microsoft.com/pt-br/library/office/ms462365.aspx]: https://msdn.microsoft.com/pt-br/library/office/ms462365.aspx
[SharePoint Plus]: http://aymkdn.github.io/SharepointPlus/
[MSDN]: https://msdn.microsoft.com/pt-br/library/office/jj163201.aspx
