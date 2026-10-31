# Contas do Agrupamento

Duas apps para telemóvel que registam todo o dinheiro do agrupamento: o que entra, o que sai,
**com quem está** e o que falta acertar entre os chefes e o tesoureiro. Custo: zero.

- **App dos chefes**: `https://<utilizador>.github.io/<repositório>/`
- **App da tesouraria**: `https://<utilizador>.github.io/<repositório>/tesouraria.html`

```
Telemóveis ──► GitHub Pages (docs/) ──► Apps Script ──► Google Sheet "Contas do Agrupamento"
                                                    └─► Google Drive (fotos das faturas)
```

## Como funciona

Tudo é um **movimento de dinheiro**, e cada chefe tem um **saldo por secção**:

| Saldo | Significa |
|---|---|
| positivo | o chefe tem consigo dinheiro da secção que ainda não entregou ao fundo |
| negativo | o chefe adiantou dinheiro; o agrupamento deve-lhe esse valor |
| zero | contas acertadas |

| Movimento | Quem regista | Efeito |
|---|---|---|
| 💰 Recebi dinheiro | chefe | saldo do chefe sobe (ou o fundo, se foi para a conta do agrupamento*) |
| 🧾 Paguei uma despesa (com foto) | chefe | saldo do chefe desce (ou o fundo, se pago com cartão do agrupamento*) |
| 🏦 Entreguei ao fundo | chefe | saldo do chefe desce e o fundo da secção sobe* |
| 🤝 Passei dinheiro a outro chefe | chefe | passa saldo de um chefe para outro |
| Reembolso / recebimento | tesoureiro | acerta o saldo do chefe a zero, a partir do fundo da secção |

\* Só conta depois de o **tesoureiro confirmar**.

Cada chefe trabalha com a sua secção e com o Agrupamento. O fundo de cada secção (e o do Agrupamento)
é guardado pelo tesoureiro; os reembolsos saem do fundo da secção a que o dinheiro pertence.

Na app dos chefes há também **Contas de uma atividade**: o que entrou (em dinheiro e eletrónico), o que se gastou,
e quanto sobrou ou faltou, com um botão para enviar as contas por WhatsApp/email.

---

## Instalação / atualização do backend

Na conta Google do agrupamento, em <https://script.google.com>, no projeto que já existe:

1. Substitui o conteúdo de `Código.gs` pelo de [`apps-script/Code.gs`](apps-script/Code.gs) e guarda.
   (`appsscript.json` igual a [`apps-script/appsscript.json`](apps-script/appsscript.json).)
2. Corre a função **`setup`**. Cria a folha **Contas do Agrupamento** na pasta *Faturas do Agrupamento*
   e mostra no registo os dois PINs. A folha da versão anterior fica com "(arquivo)" no nome, não é apagada.
   Pode ser corrido mais vezes sem estragar nada.
3. **Implementar → Gerir implementações → ✏️ → Versão: Nova versão → Implementar.** (O URL mantém-se.)
4. Na folha **Contas do Agrupamento**:
   - aba **Pessoas**: um chefe por linha, com a secção e `Sim` em *Ativo*.
     Quem sair: muda *Ativo* para `Não` (não apagues a linha, por causa do histórico);
   - aba **Fundos**: na coluna amarela *Saldo inicial*, quanto tem hoje cada fundo.

Instalação de raiz (projeto novo): igual, mas antes de *Implementar* vai a **Implementar → Nova implementação →
App Web**, *Executar como: Eu*, *Quem tem acesso: Qualquer pessoa*, e cola o URL `/exec` em `URL_API`
no início do `<script>` de `docs/index.html` e `docs/tesouraria.html`.
O GitHub Pages serve a pasta `/docs` do ramo `main` (Settings → Pages).

> **Mudar PINs:** Definições do projeto → Propriedades do script → `PIN` (chefes) e `PIN_TESOURARIA`.

## A folha de cálculo

| Aba | Conteúdo |
|---|---|
| **Fundos** | por secção: saldo inicial, no fundo, com os chefes, em dívida aos chefes, total real |
| **Movimentos** | todos os movimentos (uma linha cada), com link para a fatura e o estado |
| **Saldos dos chefes** | saldo de cada chefe em cada secção (calculado automaticamente) |
| **Pessoas** | lista de chefes (editada à mão) |

Estados de um movimento: `Confirmado`, `Por confirmar` (à espera do tesoureiro), `Recusado`, `Anulado`.
Para corrigir um erro diretamente na folha, muda o *Estado* para `Anulado` em vez de apagar a linha.

## Partilhar

- **Chefes:** link da app + PIN. Na primeira vez escolhem o nome da lista; fica memorizado.
- **Tesoureiro:** prepara-lhe o telemóvel (link `tesouraria.html`, PIN, *Adicionar ao ecrã principal*).

Adicionar ao ecrã principal: **iPhone** (Safari) Partilhar → *Adicionar ao ecrã principal*;
**Android** (Chrome) menu ⋮ → *Adicionar ao ecrã principal*.

## Cábula para o tesoureiro

A app tem três botões em cima:

**Confirmar** — os chefes dizem que lhe entregaram dinheiro, ou que entrou dinheiro na conta, ou que pagaram
com o cartão do agrupamento. Aparece um de cada vez.
1. Ver o valor e quem foi.
2. Carregar em **Sim** (recebi / entrou / está certa) ou **Não**.
3. Confirmar com **Sim, confirmo**. Se se enganou no botão: **Não, voltar atrás**.
4. Se ainda não sabe: **Decidir mais tarde**.

**Acertar contas** — a lista do que tem de pagar ou receber de cada chefe.
- *"O agrupamento deve 10,00 € a Marta"* → depois de lhe pagar, carregar em **Paguei em dinheiro** ou **Paguei por transferência**.
- *"Ana tem 30,00 € consigo"* → depois de receber, carregar em **Recebi em dinheiro** ou **Recebi por transferência**.
- Confirmar com **Sim, confirmo**.

**Quanto há** — quanto dinheiro tem cada secção.

Se se enganou depois de confirmar: **Enganei-me — ver o que já fiz** → **Enganei-me, desfazer**.
