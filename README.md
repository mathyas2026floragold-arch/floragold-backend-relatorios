# FloraGold Backend v13 - Correção Pausa 0800

Atualização focada em destravar a primeira etapa quando o sistema não entra em pausa.

Mudanças:
- o botão "Validar acesso e pausa" agora também tenta ativar a pausa 0800;
- seleção de "Ligação 0800" mais robusta;
- clique no botão amarelo de Iniciar pausa procurando o botão no mesmo bloco do select;
- suporte a modal de confirmação: Confirmar, Sim ou OK;
- mensagens de progresso mostram o status detectado.

Variáveis recomendadas no Render:
DEMO_MODE=false
EXPORT_MODE=native
AUTO_PAUSE=true
SISTEMA_URL=https://floragold.acessocloud.com/index.php
HEADLESS=true
REPORT_DIR=./relatorios
LINHAS_POR_PAGINA=3000
PLAYWRIGHT_BROWSERS_PATH=0

Se ainda não clicar no botão correto, configure manualmente:
PAUSE_BUTTON_SELECTOR=[title*="Iniciar pausa"]
CONFIRM_PAUSE_SELECTOR=button:has-text("Confirmar")
