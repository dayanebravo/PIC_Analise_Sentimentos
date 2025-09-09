import praw
import pandas as pd
import re
import emoji
from langdetect import detect, LangDetectException
import random

# Regex para remover caracteres inválidos no Excel
ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')

# Subreddits selecionado
subreddits_alvo = ["brasil"]

# Configurações da API do Reddit
reddit = praw.Reddit(
    client_id='oLY45wsXCcFsgG1DpVdtdw',
    client_secret='daw37L0pgiERHsct1XhC8TJND_1BAw',
    user_agent='ICAnaliseSentimentosBot/0.1 by AnaliseSentimentosIC'
)

contador = 1
postsEncontrados = []
textos_vistos = set()
ids_vistos = set()  # Novo conjunto para IDs únicos

# Palavras-chave por categoria
palavrasChavesGrupos = {
    "positivo": ["amo", "feliz", "alegre", "adoro"],
    "negativo": ["raiva", "triste", "ódio", "ansioso"],
    "neutro": ["terapia",  "autoestima", "sentimento", "apoio"]
}


# Loop para selecionar cada palavra chave e realizar varredura na API reddit
for categoria, palavras in palavrasChavesGrupos.items():
    total_categoria = 50000
    num_palavras = len(palavras)
    limite_por_palavra = total_categoria // num_palavras

    for palavra in palavras:

            prev_coletados = 0
            posts_coletados_palavra = 0
            after = None
            while posts_coletados_palavra < 200:
                novos_posts = 0
                for nome_subreddit in subreddits_alvo:
                    subreddit = reddit.subreddit(nome_subreddit)
                    search_params = {'limit': 100, 'sort': 'new'}
                    if after:
                        search_params['params'] = {'after': after}
                    results = list(subreddit.search(palavra, **search_params))
                    if not results:
                        break  # Não há mais resultados
                    for submission in results:
                        if posts_coletados_palavra >= 200:
                            break
                        texto = f"{submission.title} {submission.selftext}".replace("\n", " ").strip()
                        texto = emoji.demojize(texto)
                        texto = ILLEGAL_CHARACTERS_RE.sub("", texto)
                        texto.lower()
                        try:
                            if detect(texto) != "pt":
                                continue
                        except LangDetectException:
                            continue
                        post_id = submission.id
                        if post_id in ids_vistos:
                            continue
                        ids_vistos.add(post_id)
                        print(f"Buscando post: {contador} ({categoria.upper()} - {palavra})")
                        postsEncontrados.append({
                            "data": pd.to_datetime(submission.created_utc, unit='s'),
                            "texto": texto,
                            "autor": submission.author.name if submission.author else "Desconhecido",
                            "palavraChave": palavra,
                            "categoria": categoria,
                            palavra: posts_coletados_palavra
                        })
                        posts_coletados_palavra += 1
                        contador += 1
                        novos_posts += 1
                    # Atualiza 'after' para próxima página
                    if results:
                        after = results[-1].fullname if hasattr(results[-1], 'fullname') else None
                    else:
                        after = None
                print(f"{palavra} - {posts_coletados_palavra}")
                # Se não encontrou nenhum post novo nesta rodada, interrompe o loop
                if novos_posts == 0:
                    print(f"Não há mais posts únicos para a palavra '{palavra}'. Parando busca.")
                    break




# Coleta de posts aleatórios após a coleta principal
subreddit_aleatorio = reddit.subreddit("conversas")  # Pode trocar por "all" ou outro
posts_recentes = list(subreddit_aleatorio.new(limit=1000))  # Busca 1000 para sortear depois

postsAleatorios = []
contador_aleatorio = 1

for submission in random.sample(posts_recentes, k=200):
    texto = f"{submission.title} {submission.selftext}".replace("\n", " ").strip()
    texto = emoji.demojize(texto)
    texto = ILLEGAL_CHARACTERS_RE.sub("", texto)

    try:
        if detect(texto) != "pt":
            continue
    except LangDetectException:
        continue

    postsAleatorios.append({
        "data": str(submission.created_utc),
        "texto": texto,
        "autor": submission.author.name if submission.author else "Desconhecido",
        "subreddit": submission.subreddit.display_name
    })

    print(f"[{contador_aleatorio}/200] Post aleatório coletado.")
    contador_aleatorio += 1

# Converte para DataFrame e salva por categoria
df = pd.DataFrame(postsEncontrados)

# Ajusta colunas: mantém até 'categoria' e renomeia a coluna da palavra-chave para 'quantidade'

nome_arquivo = r"C:\Users\fabri\Desktop\IC Analise de sentimentos\post_saudeReddit.xlsx"
with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:

    for categoria in df['categoria'].unique():
            df_filtrado = df[df['categoria'] == categoria].copy()
            # Ajusta colunas para cada aba
            colunas_principais = ['data', 'texto', 'autor', 'palavraChave', 'categoria']
            palavra_col = None
            for palavra in palavrasChavesGrupos['positivo'] + palavrasChavesGrupos['negativo'] + palavrasChavesGrupos['neutro']:
                if palavra in df_filtrado.columns:
                    palavra_col = palavra
                    break
            if palavra_col:
                df_filtrado = df_filtrado.rename(columns={palavra_col: 'quantidade'})
                colunas_principais.append('quantidade')
                df_filtrado = df_filtrado[[col for col in colunas_principais if col in df_filtrado.columns]]
            # Adiciona coluna 'quantidade' com o número de posts por palavra-chave
            if 'palavraChave' in df_filtrado.columns:
                df_filtrado['quantidade'] = df_filtrado.groupby('palavraChave').cumcount() + 1
            aba = categoria[:31]
            df_filtrado.to_excel(writer, sheet_name=aba, index=False)

    if postsAleatorios:
        df_aleatorios = pd.DataFrame(postsAleatorios)
        df_aleatorios.to_excel(writer, sheet_name="aleatorios", index=False)

