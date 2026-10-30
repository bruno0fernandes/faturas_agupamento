# Faturas do Agrupamento

App para o telemóvel onde os dirigentes registam as faturas: tiram foto, escolhem a secção, valor e descrição.
As fotos ficam numa pasta da Google Drive e os dados numa Google Sheet, organizados por **ano letivo**
(1 de setembro a 31 de agosto, pela data da fatura). Custo: zero.

```
Telemóvel ──► app (GitHub Pages) ──► Apps Script ──► Google Drive (fotos)
                                                 └─► Google Sheet (uma aba por ano + resumo)
```

- `apps-script/` — o "servidor", que corre na conta Google do agrupamento
- `docs/` — a app que o pessoal abre no telemóvel

---

## 1. Backend (Google Apps Script) — ~10 min

Usa a **conta Google do agrupamento**, não uma pessoal.

1. Vai a <https://script.google.com> → **Novo projeto**. Dá-lhe o nome *Faturas do Agrupamento*.
2. Apaga o conteúdo de `Código.gs` e cola o conteúdo de [`apps-script/Code.gs`](apps-script/Code.gs). Guarda (Ctrl/Cmd+S).
3. **Definições do projeto** (⚙️) → ativa *"Mostrar ficheiro de manifesto appsscript.json no editor"*.
   Volta ao editor, abre `appsscript.json` e substitui pelo conteúdo de [`apps-script/appsscript.json`](apps-script/appsscript.json).
4. No editor, escolhe a função **`setup`** na lista do topo e carrega em **Executar**.
   - A Google pede autorização → *Rever permissões* → escolhe a conta → *Avançadas* → *Aceder a Faturas do Agrupamento (não seguro)* → *Permitir*.
     (O aviso aparece porque o script é teu e não foi verificado pela Google; é normal.)
   - No **Registo de execução** aparecem os links da pasta, da folha de cálculo e o **PIN** gerado.
5. (Opcional) Corre a função **`testeSubmissao`** e confirma que aparece uma linha na Sheet e uma imagem na pasta. Depois apaga essa linha e o ficheiro.
6. **Implementar → Nova implementação** → tipo *App Web*:
   - *Executar como:* **Eu**
   - *Quem tem acesso:* **Qualquer pessoa**
   - Carrega em *Implementar* e **copia o URL da app web** (termina em `/exec`).

> **Mudar o PIN:** Definições do projeto → Propriedades do script → edita `PIN`. Tem efeito imediato;
> quem tiver o PIN antigo é-lhe pedido o novo no próximo envio (sem perder o que já preencheu).

> **Se alterares o `Code.gs` mais tarde:** Implementar → Gerir implementações → ✏️ → Versão: *Nova versão* → Implementar.
> Assim o URL mantém-se. (Criar uma *nova implementação* muda o URL.)

## 2. App (GitHub Pages) — ~5 min

1. Abre [`docs/index.html`](docs/index.html) e substitui `COLAR_AQUI_O_URL_DA_APP_WEB` pelo URL do passo 1.6.
2. Cria um repositório no GitHub (pode ser numa conta do agrupamento) e envia esta pasta.
3. No repositório: **Settings → Pages** → *Source: Deploy from a branch* → branch `main`, pasta **`/docs`** → *Save*.
4. Ao fim de 1–2 minutos a app fica em `https://<utilizador>.github.io/<repositório>/`.

## 3. Partilhar com os dirigentes

Envia o link e o PIN (de preferência em mensagens separadas). Sugere que adicionem a app ao ecrã principal:

- **iPhone (Safari):** botão Partilhar → *Adicionar ao ecrã principal*
- **Android (Chrome):** menu ⋮ → *Adicionar ao ecrã principal* / *Instalar app*

Fica com ícone próprio e abre como uma app normal. O PIN e o nome de quem pagou ficam memorizados no telemóvel.

## Para a tesouraria

- Partilha a pasta **Faturas do Agrupamento** da Drive com quem gere as contas.
- A Sheet tem, por ano letivo, uma aba com todas as faturas (com link *Ver fatura*) e uma aba **Resumo** com totais por secção.
- Ao mudar o ano letivo (1 de setembro), a aba e a pasta novas são criadas automaticamente na primeira fatura. Os anos anteriores ficam arquivados.
- As faturas são arrumadas pela **data da fatura**: um recibo do acampamento de julho carregado em setembro vai para o ano anterior, como deve ser.
