# Calendário de Datas Comemorativas para o Google Sheets (Apps Script)

Crie e mantenha um calendário de datas comemorativas para diversos clientes facilmente, aumente sua produtividade e simplifique o trabalho com a ajuda desta ferramenta para o Google Sheets!

---

## Demonstração 📅

Com a ferramenta de Calendário de Datas Comemorativas para o Google Sheets é possível criar calendários, editar e adicionar datas, inserir observações pertinentes para cada data e organizá-las facilmente com filtros!

Veja o vídeo de demonstração de uso abaixo:

[![video-de-demonstração](/instruções/capa-demonstração.png)](/instruções/demonstração.webm)
_[Baixar vídeo](/instruções/demonstração.webm)_

### Benefícios

Essa ferramenta ajuda você a aumentar a sua produtividade e manter a consistência no trabalho de todos os seus clientes! Além disso, veja os principais benefícios dessa ferramenta:

- 🤩 Fácil e intuitiva.
- 🚀 Diversos calendários em uma única planilha.
- 🪄 Consistência de layout e informações.
- 📄 Baixe como PDF facilmente.
- 🤝 Compartilhe acesso com seus colegas de trabalho.
- 🏅 Licença de domínio público.
  - Um conhecimento básico de JavaScript é recomendado para realizar alterações no código.

---

## Utilizando a ferramenta 🪄

O Calendário de Datas Comemorativas conta com diversas funções úteis e intuitivas para que você possa começar a trabalhar imediatamente após a configuração inicial! Veja abaixo os primeiros passos:

### Primeiro passo: Adicionando um calendário

Para começar, primeiramente precisamos criar um novo calendário (planilha) no documento! Mas você não precisa se preocupar, a ferramenta dá conta de todo o processo!

Veja como criar um novo calendário no documento:

![criar-novo-calendario](/instruções/criar-novo-calendario.gif)

Você também pode seguir as instruções abaixo:
1. No menu superior, clique em "Calendário";
1. Dentre as opções disponíveis, clique em "Criar novo calendário de datas";
1. Siga as instruções apresentadas para criar um novo calendário;
1. Ao final do processo, você terá uma nova página em sua planilha com o nome escolhido para o calendário.

### Segundo passo: Adicionando uma nova data comemorativa

Agora que já temos um calendário pronto, basta adicionarmos as datas que precisamos. É recomendado seguir as seguintes instruções para adicionar novas datas (linhas) ao calendário:

![adicionar-data](/instruções/adicionar-data.gif)

Você também pode seguir as instruções abaixo:
1. No menu superior, clique em "Calendário";
1. Dentre as opções disponíveis, clique em "Adicionar data";
1. Uma nova linha será adicionada à tabela. Preencha-a com as informações necessárias;
    - Evite adicionar uma linha manualmente, pois assim você terá de configurar as dimensões e funcionalidades.
    - Atenção: Certifique-se de que todas as datas estão no mesmo ano, caso contrário, a organização automática poderá não funcionar como esperado.

### Terceiro passo: Organizando as datas no calendário

Por fim, com todas as datas adicionadas, podemos agora reorganizar o calendário, para que as datas sejam exibidas de maneira ascendente, isso facilita a visualização!

Veja como é simples organizar um calendário:

![organizar-datas](/instruções/organizar-datas.gif)

Você também pode seguir as instruções abaixo:
1. No menu superior, clique em "Calendário";
1. Dentre as opções disponíveis, clique em "Organizar datas do calendário";
1. Aguarde enquanto as datas são organizadas em ordem ascendente.

> Você também pode usar os filtros da primeira linha da planilha para filtrar ou ordenar a exibição da maneira que preferir.

---

## Instruções de uso em projeto único 📝

Esta ferramenta pode ser adicionada manualmente a um novo documento ou a um documento já existente de maneira simples!

A ferramenta é composta por um único código de Apps Script (```mains.js```), uma linguagem de programação feita para os serviços do Google, como a suíte do Google Workspace (Docs, Slides, Sheets e Forms).

Seguindo as instruções abaixo, você poderá começar a utilizar essa ferramenta em instantes:

[![video-passo-a-passo](/instruções/capa-passo-a-passo.png)](/instruções/passo-a-passo.webm)
_[Baixar vídeo](/instruções/passo-a-passo.webm)_

> Caso você tenha várias contas do Google conectadas, recomendo que abra uma nova Janela Anônima e conecte apenas uma conta do Google.

### Passo a passo

Acompanhe o passo a passo com as instruções por escrito abaixo:

1. Crie um novo documento no [Google Sheets 🡕](sheets.new);
    - Caso queira utilizar em um documento já existente, pule para o próximo passo.
1. No menu superior, clique em "Extensões", depois, clique em "Apps Script";
1. Na nova aba, você deve encontrar o Apps Script aberto já com o seguinte código:
    ```JavaScript
    function myFunction() {
      
    }
    ```

1. Nomeie o projeto com o nome que preferir;
    - Exemplo: "Calendário de Datas Comemorativas".
1. Copie o código de ```main.js``` para o editor do projeto;
    - Use o atalho Ctrl + A para copiar todo o código e Ctrl + V para colar no editor do projeto.
1. Clique em "Salvar projeto" (ou aperte Ctrl + S);
1. Clique em "Executar" para executar a função ```onOpen```.
1. Você deve encontrar a seguinte mensagem: "Autorização obrigatória. O projeto requer sua permissão para acessar seus dados". Clique em "Revisar permissões";
1. Na tela seguinte, faça login na sua conta do Google e prossiga;
1. Na tela seguinte, você deve encontrar a seguinte mensagem: "O Google não verificou este app", clique em "Avançado" e, em seguida, "Acessar [Nome do seu projeto] (não seguro)";
    - Esta autorização é necessária, pois o código não foi enviado ao Google para aprovação e posterior publicação no Google Workspace Marketplace. Como você tem acesso ao código em sua íntegra, pode conferir todas as ações que são executadas e aferir sua legitimidade.
    - Como "desenvolvedor", o Google está se referindo a você, veja o endereço de e-mail exibido.
    - Como "app", o Google está se referindo ao código disponibilizado neste repositório (```main.js```).
1. A seguir, confirme que o seu projeto tenha acesso às seguintes informações: "Ver, editar, criar e excluir todos os seus arquivos do Planilhas Google", clicando em "Permitir";
    - Apesar dessa abrangência, o código irá apenas funcionar neste documento em específico que você criou. Essa é uma permissão genérica requerida automaticamente para execução do código.
1. Após isso, o "Registro de execução" deve apresentar dois logs, "Notificação:	Execução iniciada" e "Notificação:	Execução concluída";
1. Volte ao documento para conferir que, na barra superior, um novo menu está disponível: "Calendário";
1. Agora você já pode fechar o editor do projeto (aba do Apps Script).

Pronto, seu projeto está configurado corretamente e já pode ser utilizado!

> Lembrete: Caso você compartilhe a planilha com colegas de trabalho, eles também precisarão fazer login e autorizar o funcionamento do código para poder utilizá-lo. Essa autorização é individual e única para cada usuário.
> > Os usuários que não autorizarem o funcionamento do código ainda poderão acessar a planilha de acordo com as permissões fornecidas a eles, porém não poderão utilizar o código.

---

👨🏽‍💻 Criado por [André Moura Santos 🡕](https://andremourasantos.com.br/?utm_source=github&utm_medium=repositories&utm_campaign=calendario-de-datas-comemorativas), 2024, licença MIT.