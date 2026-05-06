# ✨ Guia de Publicação - Aura

Siga estes passos para colocar seu aplicativo na internet e compartilhar com o mundo!

## 1. Criar conta no GitHub
- Acesse [github.com](https://github.com/) e crie sua conta.
- Crie um novo repositório chamado `aura`.

## 2. Enviar Arquivos
- Arraste todos os arquivos da pasta `data-dashboard` para dentro do seu repositório no GitHub.
- Clique em "Commit changes" para salvar.

## 3. Publicar no Streamlit Cloud
- Acesse [share.streamlit.io](https://share.streamlit.io/).
- Conecte com sua conta do GitHub.
- Clique em **"New app"**.
- Selecione o repositório `aura` e o arquivo `app.py`.
- No campo **Main path**, escolha `app.py`.
- No campo **URL**, tente colocar apenas `aura`. Se já estiver em uso, tente algo como `aura-diag` ou `portal-aura`.

## 4. Configurar Persistência (Google Sheets)
Para que os dados não sejam apagados, você precisará conectar ao Google Sheets:
1. Crie uma Planilha Google chamada `Aura_Dados`.
2. No menu do Streamlit Cloud, vá em **Settings** > **Secrets**.
3. Peça para o seu desenvolvedor (eu!) as credenciais para colar lá.

## 5. Pronto!
Seu link oficial será algo como: **https://aura.streamlit.app**

---
*Desenvolvido com carinho para transformar a educação.*
