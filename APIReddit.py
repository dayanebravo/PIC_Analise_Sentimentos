import praw
import pandas as pd
import re
import emoji
from langdetect import detect, LangDetectException
import random

# Regex para remover caracteres inválidos no Excel
ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')

# Subreddits selecionado
subreddits_alvo = ["conversas"]

# Configurações da API do Reddit
reddit = praw.Reddit(
    client_id='oLY45wsXCcFsgG1DpVdtdw',
    client_secret='daw37L0pgiERHsct1XhC8TJND_1BAw',
    user_agent='ICAnaliseSentimentosBot/0.1 by AnaliseSentimentosIC'
)

contador = 1
postsEncontrados = []
textos_vistos = set()

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
        posts_coletados_palavra = 0

        for nome_subreddit in subreddits_alvo:
            subreddit = reddit.subreddit(nome_subreddit)

            # Busca nos títulos e corpos das submissões
            for submission in subreddit.search(palavra, sort='new', limit=limite_por_palavra):
                # Solução temporária para o problema da quantidade de posts
                if posts_coletados_palavra >= 200:
                    break
                
                # Tratamento dos posts selecionados
                texto = f"{submission.title} {submission.selftext}".replace("\n", " ").strip()
                texto = emoji.demojize(texto)
                texto = ILLEGAL_CHARACTERS_RE.sub("", texto)
                texto.lower()

                try:
                    if detect(texto) != "pt":
                        continue
                except LangDetectException:
                    continue

                if texto in textos_vistos:
                    continue
                textos_vistos.add(texto)

                print(f"Buscando post: {contador} ({categoria.upper()} - {palavra})")
                
                #Salvando posts em lista
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
                
            print(f"{palavra} - {posts_coletados_palavra}" )

            if posts_coletados_palavra >= limite_por_palavra:
                break  # se já coletou o suficiente dessa palavra, pula para a próxima




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

nome_arquivo = r"C:\Users\fabri\Desktop\IC Analise de sentimentos\post_saudeReddit.xlsx"
with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
    for categoria in df['categoria'].unique():
        df_filtrado = df[df['categoria'] == categoria]
        aba = categoria[:31]
        df_filtrado.to_excel(writer, sheet_name=aba, index=False)

    if postsAleatorios:
        df_aleatorios = pd.DataFrame(postsAleatorios)
        df_aleatorios.to_excel(writer, sheet_name="aleatorios", index=False);