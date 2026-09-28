# Ativo · Endereço incorreto

Página que pega a planilha do JMS, separa os pacotes com **Endereço incorreto**, mostra o ranking por motorista,
busca o telefone completo de cada cliente no JMS e deixa tudo pronto para copiar / mandar no WhatsApp.

Tudo roda no navegador do próprio computador. A planilha e os telefones **não são enviados para lugar nenhum**.

## Uso diário

1. Abra a página e clique em **Escolher arquivo** → selecione a planilha exportada do JMS.
2. **(Só na primeira vez em cada computador)** arraste o botão **📞 Ativo JMS** para a barra de favoritos do Chrome
   (Ctrl+Shift+B mostra a barra).
3. Clique em **Copiar códigos pendentes**.
4. No JMS: **Operação → Rastreamento do pacote**. Clique no favorito **Ativo JMS**, cole os códigos e clique **Iniciar**.
   Deixe essa aba do JMS **visível** enquanto roda (em segundo plano o Chrome deixa lento). ~4 segundos por pacote.
5. Ao terminar, clique **Copiar resultado**, volte à página, cole no campo e clique **Importar resultado**.
6. Use a lista por motorista: botões **Cód / Tel / End / Msg / WhatsApp / Enviado**. **Exportar Excel** gera o arquivo.

## Avisos que a página mostra

- **Sem telefone cadastrado no JMS:** o JMS devolveu "0". Não há número para esse cliente.
- **⚠ 10 dígitos:** provável falta do 9 na frente (ou telefone fixo). O WhatsApp pode não abrir.
- **⚠ mesmo telefone em clientes diferentes:** número provavelmente genérico do sistema; confira antes de mandar.

## Se o favorito parar de funcionar

O JMS pode mudar o layout. O código do favorito está em `bookmarklet.js` (seletores dos campos no topo do arquivo).
Depois de ajustar, basta recriar o favorito.

## Onde ficam os dados

Somente no `localStorage` do navegador deste computador. **Limpar tudo** apaga. Cada computador tem o seu.
