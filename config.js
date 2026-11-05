// =============================================================
// Configuração do Monitor SLA Telefônica/Vivo
// Edite este arquivo antes de publicar (veja o README.md).
// Nada aqui é secreto: o login acontece no navegador de quem usa.
// =============================================================
window.MONITOR_CONFIG = {
  // ID do aplicativo registrado no Microsoft Entra (Azure AD).
  // Troque pelo "Application (client) ID" do registro que você criar.
  clientId: "9ff07a0f-8433-4cbe-a1a9-80b96374df67",

  // ID do tenant da Comm (Microsoft Entra > Visão geral > Tenant ID).
  tenantId: "0760b291-f27e-4201-ab06-3e7e68daa846",

  // Caixa de correio monitorada (compartilhada). O usuário que faz login
  // precisa ter permissão de leitura nela (Acesso Total / Full Access).
  mailbox: "infra@commcenter.com.br",

  // Remetentes/destinatários monitorados.
  parceiros: [
    "iranilde.santos@telefonica.com",
    "gestao_vgr.br@telefonica.com",
    "servicosdigitais@vivo.com.br"
  ],

  // Domínios considerados "lado Telefônica/Vivo" e "lado Comm".
  dominiosParceiro: ["telefonica.com", "vivo.com.br"],
  dominiosInternos: ["commcenter.com.br"],

  // Mensagens automáticas que não contam como "última resposta".
  ignorarRemetentes: ["vivob2b@service-now.com", "postmaster@commcenter.com.br"],

  // Janela de busca e regras de SLA (em horas).
  diasRetroativos: 7,
  limiteHoras: 4,     // abaixo disso não entra na lista de prioridade
  alta: [4, 8],       // 4–8h  = ALTA
  critica: [8, 24],   // 8–24h = CRÍTICA (> 24h = ESCALAR)

  // Palavras que sobem um nível e marcam com ⚡ (assunto ou resumo).
  palavrasUrgencia: ["urgente", "prazo", "hoje", "escalonamento", "crítico", "critico",
    "bloqueio", "multa", "parado", "sem acesso", "down", "inoperante"],

  // Atualização automática (minutos). 0 = só manual.
  atualizarACadaMin: 5,

  // Fuso horário de exibição.
  fuso: "America/Sao_Paulo"
};
