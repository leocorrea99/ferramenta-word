#!/bin/bash
# ─────────────────────────────────────────────
#  Ferramenta Word — Instalador para Mac
# ─────────────────────────────────────────────

# URL do manifest hospedado (atualizar após publicar)
MANIFEST_URL="https://leocorrea99.github.io/ferramenta-word/manifest.xml"
MANIFEST_DEST="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef/ferramenta-word.xml"

clear
echo ""
echo "  ┌──────────────────────────────────┐"
echo "  │      Ferramenta Word             │"
echo "  │      Instalador — Mac            │"
echo "  └──────────────────────────────────┘"
echo ""

# Verifica se o Word está instalado
if [ ! -d "$HOME/Library/Containers/com.microsoft.Word" ]; then
    echo "  ✗  Microsoft Word não encontrado."
    echo "     Instale o Word antes de continuar."
    echo ""
    read -n 1 -s -p "  Pressione qualquer tecla para fechar..."
    exit 1
fi

# Fecha o Word se estiver aberto
if pgrep -x "Microsoft Word" > /dev/null; then
    echo "  → Fechando o Microsoft Word..."
    osascript -e 'tell application "Microsoft Word" to quit saving yes'
    sleep 2
fi

# Cria a pasta se não existir
mkdir -p "$(dirname "$MANIFEST_DEST")"

# Baixa o manifest
echo "  → Baixando manifest..."
curl -sf -o "$MANIFEST_DEST" "$MANIFEST_URL"

if [ $? -ne 0 ]; then
    echo ""
    echo "  ✗  Erro ao baixar. Verifique sua internet."
    read -n 1 -s -p "  Pressione qualquer tecla para fechar..."
    exit 1
fi

echo ""
echo "  ✓  Instalado com sucesso!"
echo ""
echo "  Como abrir:"
echo "  Word → Inserir → Suplementos"
echo "  Aba \"Suplementos do Desenvolvedor\""
echo "  → Ferramenta Word"
echo ""

# Pergunta se quer abrir o Word
read -r -p "  Abrir o Word agora? (s/n): " resposta
if [[ "$resposta" =~ ^[Ss]$ ]]; then
    open -a "Microsoft Word"
fi

echo ""
read -n 1 -s -p "  Pressione qualquer tecla para fechar..."
echo ""
