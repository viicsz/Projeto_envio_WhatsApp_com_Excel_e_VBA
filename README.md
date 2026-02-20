# Projeto_envio_WhatsApp_com_Excel_e_VBA
Projeto feito para envio de mensagens via WhatsApp, utilizando automação do VBA do Excel.

***

# WhatsApp Sender VBA 📱  

Este projeto permite o **envio automatizado de mensagens pelo WhatsApp Web** utilizando **VBA no Excel**.  
Ideal para comunicações em massa — como mensagens personalizadas de aviso, lembretes, marketing ou notificações — de forma simples, sem necessidade de APIs externas.

***

## ⚙️ Como Funciona

O script usa automação de teclado com `Application.SendKeys` para interagir com o navegador e enviar mensagens via WhatsApp Web.  
O processo é inteiramente controlado a partir de uma planilha Excel, contendo os **contatos**, os **intervalos de tempo (delay)** entre mensagens e a **mensagem** a ser enviada.

***

## 🧩 Estrutura do Projeto

### 1. Planilha Principal (Contatos)

| Coluna | Conteúdo                       | Exemplo           |
|--------|--------------------------------|-------------------|
| B      | Número de telefone (com DDI/DD) | `+5511999999999` |
| E2     | Tempo mínimo entre mensagens (segundos) | `5` |
| F2     | Tempo máximo entre mensagens (segundos) | `10` |

A partir da linha 5 (`B5`), cada linha representa um contato para envio.

***

### 2. Planilha “Mensagem”

| Célula | Conteúdo                                        |
|--------|--------------------------------------------------|
| A2     | Texto da mensagem a ser enviada                 |

*Essa mensagem será enviada para cada número da lista.*

***

## 💻 Principais Funções

### `CopiaTextoPuro(Celula As Range)`
Função utilitária que copia texto de uma célula para a área de transferência (evita problemas de formatação ou links internos do Excel).

### `EnviarWhatsApp()`
Macro principal que:
1. Abre automaticamente o **Microsoft Edge** com o WhatsApp Web.  
   - (Pode ser alterado para o Chrome, modificando o caminho do executável).
2. Aguarda o carregamento da página (30 segundos padrão).
3. Itera pelos contatos listados a partir da célula `B5`:
   - Copia o número do contato e simula teclas para localizar.
   - Cola e envia a mensagem definida na planilha “Mensagem”.
   - Aguarda um tempo aleatório entre as mensagens (dentro do intervalo definido em `E2` e `F2`).

***

## 🧠 Lógica do Delay

O intervalo entre mensagens é sorteado aleatoriamente entre dois valores (mínimo e máximo):

\[
\text{Delay} = \text{Int}((\text{MaxTempo} - \text{MinTempo} + 1) * Rnd + \text{MinTempo})
\]

Isso ajuda a simular comportamento humano e evitar bloqueios automáticos.

***

## 🪟 Pré-Requisitos

- **Excel (com suporte a macros habilitado)**
- **Microsoft Edge** (ou outro navegador ajustado no script)
- Estar **logado no WhatsApp Web** no momento da execução

***

## 🚀 Como Usar

1. Abra o arquivo `.xlsm` e habilite as macros.
2. Preencha:
   - A planilha de **contatos** (coluna B)
   - A planilha **Mensagem** (célula A2)
   - Os intervalos de tempo (E2 e F2)
3. Pressione `ALT + F8` → selecione `EnviarWhatsApp` → clique em **Executar**.
4. Aguarde o envio automático das mensagens.

***

## ⚠️ Observações Importantes

- O código **simula pressionamento de teclas**, portanto **não use o computador** enquanto o envio estiver em andamento.  
- Certifique-se de que o WhatsApp Web esteja autenticado e aberto no navegador selecionado.
- O uso excessivo pode causar restrição temporária da conta — utilize com moderação.

***

## 🧾 Exemplo de Configuração

| Contato (B) | Tempo Mínimo (E2) | Tempo Máximo (F2) | Mensagem (Mensagem!A2)             |
|--------------|------------------|-------------------|------------------------------------|
| +5511988888888 | 5 | 10 | “Olá! Esta é uma mensagem automática de teste.” |

***

## 🧑‍💻 Autor
**Victor Almeida Araújo**  
Automação VBA + Integração WhatsApp Web  

***
