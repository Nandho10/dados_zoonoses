# Vigilância e Controle do Aedes aegypti — Dashboard

Painel interativo para acompanhamento de visitas domiciliares por situação do imóvel e bairro (jan/2024 – ago/2026), atualizável diretamente via arquivos **CSV** ou **XLSX**.

## Funcionalidades

- **Indicadores Principais (Contadores):** Total de imóveis visitados, trabalhados, fechados e bairros atendidos.
- **Medidor de Aproveitamento (Ponteiro):** Percentual de imóveis trabalhados sobre os visitados, com suporte a definição de meta interativa.
- **Motivos de Não Trabalho (Rosca):** Distribuição detalhada entre Fechados, Desocupados, Recusa e Parciais/Temporada.
- **Concentração por Bairro (Barra):** Top 8 bairros por imóveis visitados com filtro interativo ao clicar.
- **Série Histórica Mensal (Linha):** Evolução mensal de visitas com destaque para os picos sazonais (fevereiro/março).
- **Bairros com Mais Imóveis Fechados (Tabela):** Ranking de bairros críticos por percentual de imóveis fechados.
- **Atualização via CSV ou XLSX:** Importação por botão ou *drag-and-drop* (substituindo a base ou acrescentando novos registros), com exportação integrada para `.xlsx` e `.csv`.

## Como Executar

### Opção 1: Direto no Navegador (ou GitHub Pages)
Basta abrir o arquivo `index.html` em qualquer navegador moderno.

### Opção 2: Servidor Local Node.js
1. Instale as dependências (caso necessário):
   ```bash
   npm install
   ```
2. Inicie o servidor:
   ```bash
   npm start
   ```
   Ou execute um duplo clique em `Abrir_Dashboard.bat` no Windows e acesse `http://localhost:3000`.
