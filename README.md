# Dashboard de Indicadores — Agrícola e Indústria

Automação que lê as planilhas da pasta do OneDrive, atualiza o `index.html`
do dashboard e publica no GitHub Pages. Depois de configurado, basta dar
duplo-clique em `update_dashboard.bat` sempre que quiser atualizar os dados.

Este guia parte do zero: criar o repositório no GitHub, deixar o site no ar
(GitHub Pages) e configurar a automação na sua máquina.

---

## 1. Pré-requisitos (instalar uma vez só)

1. **Git para Windows** — baixe em https://git-scm.com/download/win e instale
   com as opções padrão.
2. **Python 3** — baixe em https://www.python.org/downloads/ e, na tela de
   instalação, marque a caixa **"Add python.exe to PATH"** antes de clicar
   em Install.
3. Abra o **Prompt de Comando** (tecla Windows → digite `cmd` → Enter) e
   confirme que os dois foram instalados:
   ```
   git --version
   python --version
   ```
   Se algum comando não for reconhecido, feche e abra o Prompt de novo (às
   vezes o PATH só atualiza depois de reabrir).

---

## 2. Criar o repositório no GitHub

1. Entre em https://github.com e faça login (crie uma conta se ainda não tiver).
2. Clique no **+** no canto superior direito → **New repository**.
3. Preencha:
   - **Repository name**: por exemplo `dashboard-indicadores`
   - **Visibility**: escolha **Private** se os dados não podem ficar públicos,
     ou **Public** se puderem — isso importa para o passo do GitHub Pages
     (veja a observação abaixo).
   - Não marque "Add a README file" (vamos subir os arquivos já prontos).
4. Clique em **Create repository**.
5. Na página que abrir, copie a URL do repositório (algo como
   `https://github.com/SEU-USUARIO/dashboard-indicadores.git`) — você vai
   usar essa URL no passo 3.

> **Sobre repositório privado + GitHub Pages:** publicar um site a partir de
> um repositório **privado** com GitHub Pages exige um plano GitHub Pro/Team
> (ou similar) para contas pessoais. Se o seu plano for gratuito e o repo
> for privado, o Pages não vai funcionar. Nesse caso, ou o repositório fica
> **público** (qualquer pessoa com o link acessa o dashboard), ou você
> publica de outra forma (ex.: hospedar em outro lugar, ou usar Pages sem
> tornar os dados sensíveis públicos). Se os números são sensíveis, me avise
> que ajustamos a estratégia.

---

## 3. Clonar o repositório e colocar os arquivos

1. Escolha uma pasta no seu PC para o repositório local (pode ser dentro de
   Documentos, não precisa ser dentro do OneDrive da planilha).
2. Abra o Prompt de Comando nessa pasta (ou `cd` até ela) e rode:
   ```
   git clone https://github.com/SEU-USUARIO/dashboard-indicadores.git
   cd dashboard-indicadores
   ```
3. Copie estes 4 arquivos (que estão junto com este README) para dentro
   dessa pasta do repositório:
   - `index.html`
   - `update_dashboard.py`
   - `update_dashboard.bat`
   - `requirements.txt`
4. Instale a dependência Python (só precisa fazer isso uma vez):
   ```
   pip install -r requirements.txt
   ```
5. Suba os arquivos para o GitHub pela primeira vez:
   ```
   git add .
   git commit -m "Primeira publicação do dashboard"
   git push
   ```
   Na primeira vez que você der `push`, o Windows deve abrir uma janela do
   navegador pedindo para você autorizar o Git com sua conta do GitHub — só
   aceitar. Se aparecer um pedido de usuário/senha em vez disso, use um
   **Personal Access Token** no lugar da senha (GitHub não aceita mais senha
   normal): crie um em
   https://github.com/settings/tokens → **Generate new token (classic)** →
   marque o escopo `repo` → gere e copie o token → cole no lugar da senha
   quando pedido.

---

## 4. Ativar o GitHub Pages (deixar o dashboard acessível por link)

1. No GitHub, entre no repositório → aba **Settings**.
2. No menu lateral, clique em **Pages**.
3. Em **Build and deployment → Source**, escolha **Deploy from a branch**.
4. Em **Branch**, escolha `main` (ou `master`) e a pasta `/ (root)` → **Save**.
5. Espere 1–2 minutos. O link do dashboard vai aparecer no topo dessa mesma
   página, algo como:
   `https://SEU-USUARIO.github.io/dashboard-indicadores/`

Esse link é o que você compartilha com o time.

---

## 5. Configurar o caminho das planilhas (só se mudar)

O script já está configurado com o caminho que você me passou:
```
C:\Users\raulribeiro\OneDrive - CLEALCO AÇÚCAR E ÁLCOOL S.A\Projeto Confiar Excelência\Excelência em Processo\Governanças\Pasta Indicadores
```
com os arquivos `Agrícola.xlsx` e `Industria.xlsx` dentro dela.

Se a pasta ou o nome dos arquivos mudar algum dia, abra
`update_dashboard.py` em qualquer editor de texto (Bloco de Notas serve) e
edite estas linhas no topo do arquivo:

```python
PASTA_PLANILHAS = Path(
    r"C:\Users\raulribeiro\OneDrive - CLEALCO AÇÚCAR E ÁLCOOL S.A"
    r"\Projeto Confiar Excelência\Excelência em Processo\Governanças\Pasta Indicadores"
)
ARQ_AGRICOLA = PASTA_PLANILHAS / "Agrícola.xlsx"
ARQ_INDUSTRIA = PASTA_PLANILHAS / "Industria.xlsx"
```

---

## 6. Uso do dia a dia

Sempre que as planilhas forem atualizadas e você quiser refletir isso no
dashboard publicado:

1. Dê duplo-clique em **`update_dashboard.bat`** (dentro da pasta do repositório).
2. Uma janela preta vai abrir mostrando o progresso: quantos indicadores
   foram lidos de cada planilha, se algum indicador novo apareceu sem estar
   mapeado (avisa em vez de adivinhar errado), e por fim se a publicação no
   GitHub deu certo.
3. Quando terminar, a janela pede para apertar uma tecla e fecha.
4. O dashboard publicado (o link do passo 4) atualiza sozinho em 1–2 minutos.

Se nada mudou nos dados desde a última vez, o script avisa
"Nenhuma alteração nos dados" e não sobe nada — não tem problema rodar
quantas vezes quiser.

---

## 7. Se aparecer um indicador novo na planilha

O script reconhece a **estrutura** das planilhas automaticamente (então
números novos, semanas novas, etc. são pegos sozinhos). O que ele não
adivinha sozinho é o **nome bonito / categoria** de um indicador
**totalmente novo** que nunca existiu antes — nesse caso ele:

- Ainda assim mostra o indicador no dashboard (com um nome/categoria padrão
  gerado automaticamente a partir do texto da planilha).
- Imprime um aviso na janela preta, por exemplo:
  ```
  [AGRÍCOLA] Indicador não mapeado em AGRICOLA_META: título='NOVO INDICADOR X' ...
  ```

Para deixar bonito, abra `update_dashboard.py`, ache o dicionário
`AGRICOLA_META` (ou `INDUSTRIA_META`) perto do topo do arquivo, e adicione
uma linha seguindo o padrão das que já existem, com o nome/categoria/unidade
que preferir. Da próxima vez que rodar, ele já usa o que você configurou.

---

## Arquivos deste pacote

| Arquivo                 | Para quê serve                                              |
|--------------------------|--------------------------------------------------------------|
| `index.html`             | O dashboard em si — é o que fica publicado no GitHub Pages. |
| `update_dashboard.py`    | O script que lê as planilhas e atualiza o `index.html`.      |
| `update_dashboard.bat`   | Atalho de duplo-clique que roda o script acima.               |
| `requirements.txt`       | Lista de pacotes Python necessários (`pip install -r ...`).  |
| `README.md`              | Este guia.                                                    |
