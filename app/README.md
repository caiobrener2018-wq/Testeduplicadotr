# Verificador de Fotos Duplicadas — Sebrae na sua Empresa

Aplicação web que recebe uma planilha de atendimentos, baixa as fotos
informadas, e identifica quais imagens foram reenviadas entre atendimentos
diferentes (mesmo arquivo). Gera um Excel com as duplicatas destacadas,
indicando de qual linha (original) cada cópia veio.

## Como a detecção funciona

Cada imagem é baixada e tem seu **hash SHA-256** calculado. Duas fotos com o
mesmo hash são o **mesmo arquivo** (duplicata exata, sem falso positivo). A
primeira ocorrência de uma imagem (linha mais acima na planilha) é tratada
como **original**; as repetições seguintes, em atendimentos diferentes, são
marcadas como **duplicadas**.

Fotos repetidas dentro do mesmo atendimento (mesma linha) são ignoradas por
padrão, por serem provável reenvio acidental — não uso indevido. Isso é
configurável (`ignorar_mesma_linha`).

## Instalação

Requer Python 3.9+.

```bash
cd app
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Abra http://localhost:5000 no navegador, envie a planilha `.xlsx`, escolha a
coluna das fotos (padrão: D) e clique em **Verificar duplicatas**.

## Estrutura da planilha esperada

- Linha 1: cabeçalho.
- A partir da linha 2: um atendimento por linha.
- Coluna de fotos (padrão **D**): um ou mais links de imagem. Se houver mais de
  um link na mesma célula, coloque um por linha (quebra de linha) — a
  ferramenta separa automaticamente.

A coluna C (Razão Social) é usada nos rótulos do relatório, se existir.

## O que sai no Excel gerado

- Células de fotos **duplicadas** pintadas de vermelho, com uma coluna de
  detalhe indicando a linha da original.
- Originais que têm cópias pintadas de amarelo.
- Coluna "Status Verificação" (OK / DUPLICADA / ORIGINAL / NÃO VERIFICADA).
- Aba **"Resumo Duplicatas"** com a lista completa dos pares cópia → original.

## Colocando no ar de graça (Render)

O Render oferece um plano gratuito com acesso de saída irrestrito (essencial
aqui, pois a ferramenta baixa fotos de qualquer domínio) e HTTPS grátis.
Único ponto de atenção: o serviço "dorme" após 15 min sem acesso e leva
30-60s para acordar no próximo uso — sem custo, e aceitável para uso interno.

**Passo a passo:**

1. Crie um repositório no GitHub e suba o conteúdo desta pasta (`app.py`,
   `detector.py`, `templates/`, `requirements.txt`, `Procfile`).
2. Crie uma conta gratuita em [render.com](https://render.com) (sem cartão).
3. **New +** → **Web Service** → conecte o repositório.
4. Configure:
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** deixe em branco (o `Procfile` já define isso)
   - **Instance Type:** Free
5. Clique em **Create Web Service**. Em poucos minutos a URL
   `https://seu-app.onrender.com` estará no ar.

**Por que 1 worker no `Procfile`:** o progresso dos jobs fica em memória.
Com mais de 1 worker, cada processo teria sua própria memória e o
acompanhamento do progresso quebraria. Para o volume de uma ferramenta
interna, 1 worker com várias threads (já configurado) dá conta bem.

**Arquivos são temporários:** a cada reinício do serviço gratuito, os
arquivos enviados/gerados são apagados. Isso é normal — o usuário baixa o
resultado logo após o processamento, então não há perda prática.

## Ajustes comuns (arquivo `detector.py`)

- `FOTOS_COL` — coluna padrão das fotos (4 = D).
- `MAX_WORKERS` — nº de downloads em paralelo (padrão 8).
- `DOWNLOAD_TIMEOUT` — tempo máximo por imagem, em segundos.
- `AUTH_HEADERS` — **importante:** se os links do Sebrae passarem a exigir
  login, preencha aqui o token/cookie de sessão, ex.:
  `AUTH_HEADERS = {"Authorization": "Bearer <seu_token>"}`.

## Observação sobre acesso às imagens

A ferramenta precisa conseguir baixar cada URL. Se os links forem públicos,
funciona direto. Se exigirem autenticação, use `AUTH_HEADERS` acima. Fotos que
não puderem ser baixadas aparecem como "NÃO VERIFICADA" no relatório, sem
travar o processo.

## Melhorias possíveis (fase 2)

- **Duplicata visual (pHash):** detectar a mesma foto mesmo recomprimida,
  redimensionada ou levemente recortada — pega tentativas de disfarçar a cópia.
  Hoje a ferramenta detecta apenas arquivos idênticos, conforme definido.
- Cache de hashes por ID de arquivo, para não rebaixar imagens já vistas em
  execuções anteriores.
- Detecção entre planilhas/períodos diferentes (histórico acumulado).
