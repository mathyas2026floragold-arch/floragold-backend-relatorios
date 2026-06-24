# Correção Playwright no Render

Substitua o package.json e envie também o arquivo render-build.sh para a raiz do repositório.

No Render, configure:

Build Command:
./render-build.sh

Start Command:
npm start

Environment Variables:
PLAYWRIGHT_BROWSERS_PATH=0
DEMO_MODE=false
EXPORT_MODE=native
AUTO_PAUSE=true
SISTEMA_URL=https://floragold.acessocloud.com/index.php
HEADLESS=true
REPORT_DIR=./relatorios
LINHAS_POR_PAGINA=3000

Depois rode: Manual Deploy > Clear build cache & deploy.
