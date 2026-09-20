/**
 * IBGESF Valuation Extractor – Worker v3
 * Arquitetura: recebe PDF(s)/DOCX como base64 → Claude lê nativamente
 * Múltiplos PDFs: consolida resultados (usa o de maior confiança por indicador)
 * Para textos estruturados (XLSX exportado como CSV) → recebe como texto
 *
 * Endpoint: POST /extract
 * Headers: X-Api-Key: ibgesf2024secure
 * Body (PDF único):   { tipo:"pdf", arquivo_base64:"...", mime_type:"application/pdf", nome:"arquivo.pdf" }
 * Body (múltiplos):   { tipo:"pdf", arquivos:[{base64:"...", mime_type:"...", nome:"..."}, ...] }
 * Body (texto):       { tipo:"texto", texto:"...", nome:"arquivo.xlsx" }
 *
 * Response: { ok:true, indicadores:{...}, modelo:"...", modo:"pdf-nativo"|"texto", arquivos_processados:N }
 */

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "POST,OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type,X-Api-Key,X-API-Key",
};

const PROMPT_PDF = `Você é um analista financeiro sênior com 20 anos de experiência em M&A e due diligence de empresas privadas brasileiras. Seu trabalho é ler demonstrações financeiras (DRE, DFC, Balanço Patrimonial) e extrair valores com precisão absoluta.

TAREFA: Extraia exatamente 7 indicadores financeiros deste documento e retorne JSON puro.

PROCESSO OBRIGATÓRIO antes de responder:
1. Identifique se os valores estão em R$ (unidade), R$ mil ou R$ milhões — isso define a escala
2. Para receita_bruta: procure "Receita Bruta", "Receita Operacional Bruta", "Faturamento Bruto" na DRE — é SEMPRE a primeira linha positiva da DRE, antes de deduções
3. Para receita_liquida: logo abaixo da receita bruta, após deduções de impostos
4. Para ebitda: procure "EBITDA", "LAJIDA", ou calcule Lucro Operacional + Depreciação + Amortização
5. Para fco: procure apenas na DFC (Demonstrativo de Fluxo de Caixa), seção "Atividades Operacionais" — NUNCA use lucro líquido como substituto
6. Para capex: na DFC, seção "Atividades de Investimento", linhas de aquisição de imobilizado/intangível
7. Para divida_financeira: no Balanço, soma de empréstimos + financiamentos + debêntures (CP + LP) — NUNCA o passivo total
8. Para caixa: no Balanço, "Caixa e Equivalentes de Caixa" ou "Disponibilidades"

ATENÇÃO ESPECIAL PARA ESCALA — LEIA COM CUIDADO:
A escala é definida APENAS pelo cabeçalho/rodapé da tabela, NUNCA pela magnitude dos números.

REGRA 1 — Valores com centavos (vírgula decimal): ex: "7.142.758,82" ou "4.200.000,00"
→ Já estão em R$ UNITÁRIOS. NÃO multiplique. valor_normalizado = 7142758.82

REGRA 2 — Cabeçalho diz "R$ mil" ou "Em milhares": ex: você vê "7.142" na tabela
→ Multiplique por 1.000. valor_normalizado = 7142000

REGRA 3 — Cabeçalho diz "R$ milhões" ou "Em milhões": ex: você vê "7,1" na tabela
→ Multiplique por 1.000.000. valor_normalizado = 7100000

REGRA 4 — Sem cabeçalho de escala e sem centavos: ex: você vê "7.142.758"
→ Já estão em R$ unitários. valor_normalizado = 7142758

EXEMPLOS CRÍTICOS:
- Tabela sem escala, linha mostra "7.142.758,82" → valor_normalizado = 7142758.82 (escala: unidade)
- Tabela "R$ mil", linha mostra "7.143" → valor_normalizado = 7143000 (escala: milhares)
- Tabela "R$ milhões", linha mostra "7,1" → valor_normalizado = 7100000 (escala: milhões)

RETORNE APENAS este JSON (sem markdown, sem explicações, sem \`\`\`):

{"receita_bruta":{"valor_normalizado":<R$ absolutos ou null>,"escala":"unidade|milhares|milhões","periodo":"2024|2023-2024|etc","linha_evidencia":"trecho literal ≤120 chars","pagina":<N ou null>,"confianca":"alta|media|baixa"},"receita_liquida":{"valor_normalizado":<R$ ou null>,"escala":"...","periodo":"...","linha_evidencia":"...","pagina":<N ou null>,"confianca":"..."},"ebitda":{"valor_normalizado":<R$ ou null>,"escala":"...","periodo":"...","linha_evidencia":"...","pagina":<N ou null>,"confianca":"..."},"fco":{"valor_normalizado":<R$ ou null>,"escala":"...","periodo":"...","linha_evidencia":"...","pagina":<N ou null>,"confianca":"..."},"capex":{"valor_normalizado":<R$ ou null>,"escala":"...","periodo":"...","linha_evidencia":"...","pagina":<N ou null>,"confianca":"..."},"divida_financeira":{"valor_normalizado":<R$ ou null>,"escala":"...","periodo":"...","linha_evidencia":"...","pagina":<N ou null>,"confianca":"..."},"caixa":{"valor_normalizado":<R$ ou null>,"escala":"...","periodo":"...","linha_evidencia":"...","pagina":<N ou null>,"confianca":"..."}}

REGRAS ABSOLUTAS:
- valor_normalizado SEMPRE em R$ absolutos: se tabela é R$ mil e valor é 4.200 → escreva 4200000
- Parênteses = negativo: (1.500) → -1500000 (se escala R$ mil)
- Período: use sempre o mais recente disponível no documento
- Se indicador não existe neste documento: valor_normalizado deve ser null
- linha_evidencia: copie literalmente do documento, máximo 120 caracteres`;

const PROMPT_TEXTO = `Você é um analista financeiro sênior especializado em empresas privadas brasileiras.
Analise este texto financeiro (extraído de planilha Excel ou CSV) e extraia os indicadores.

RETORNE APENAS um objeto JSON válido com exatamente estas 7 chaves:

{
  "receita_bruta": { "valor_normalizado": <número em R$ ou null>, "escala": "unidade|milhares|milhões", "periodo": "2023|2022-2023|etc", "linha_evidencia": "linha do CSV/planilha", "pagina": null, "confianca": "alta|media|baixa" },
  "receita_liquida": { "valor_normalizado": <número ou null>, "escala": "...", "periodo": "...", "linha_evidencia": "...", "pagina": null, "confianca": "..." },
  "ebitda": { "valor_normalizado": <número ou null>, "escala": "...", "periodo": "...", "linha_evidencia": "...", "pagina": null, "confianca": "..." },
  "fco": { "valor_normalizado": <número ou null>, "escala": "...", "periodo": "...", "linha_evidencia": "...", "pagina": null, "confianca": "..." },
  "capex": { "valor_normalizado": <número ou null>, "escala": "...", "periodo": "...", "linha_evidencia": "...", "pagina": null, "confianca": "..." },
  "divida_financeira": { "valor_normalizado": <número ou null>, "escala": "...", "periodo": "...", "linha_evidencia": "...", "pagina": null, "confianca": "..." },
  "caixa": { "valor_normalizado": <número ou null>, "escala": "...", "periodo": "...", "linha_evidencia": "...", "pagina": null, "confianca": "..." }
}

REGRAS CRÍTICAS:
1. valor_normalizado = valor em R$ absolutos (ex: 4.200.000 → 4200000, não 4200)
2. Escalas comuns em planilhas: "R$ mil" = multiplique por 1.000; "R$ milhões" = por 1.000.000
3. Parênteses indicam negativo: (1.500,00) = -1500
4. fco = Fluxo de Caixa Operacional — nunca lucro
5. divida = empréstimos/financiamentos — nunca passivo total
6. Use o período mais recente disponível
7. Não invente valores ausentes — use null

TEXTO DA PLANILHA:
`;

const CONFIANCA_PESO = { alta: 3, media: 2, baixa: 1, "nao-encontrado": 0 };
const CHAVES = ["receita_bruta", "receita_liquida", "ebitda", "fco", "capex", "divida_financeira", "caixa"];

// Lista de emails de associados (fallback hardcoded; sobreposto por env.MEMBER_EMAILS se definido)
const MEMBER_EMAILS_DEFAULT = ["hcayuela@gmail.com","silvana.vallim@hotmail.com","silvana.vallim@ibgovernancaeestrategia.com.br","danielareisregina@gmail.com","sabrinagoncalves.0112@gmail.com","anaritauchoa@gmail.com","dianaalves1923@gmail.com","medinalarabeatriz@gmail.com","ggpolotto@terra.com.br","nataliahackme@hotmail.com","hcayuela@icloud.com","marina@priolligaluppo.com.br","hugocayuela@hotmail.com","camontenegro@gmail.com","vitor.cayuela@hotmail.com","valravanixs@gmail.com","kauerizk@gmail.com","renata.dalmaso@vion.services","renatacorotti@gmail.com","ritasouza.rh@gmail.com","polonioelenice@gmail.com","gs.rigoleto@gmail.com","darlipalmacunha@gmail.com","paulorebello@hotmail.com","adm.damasceno@uol.com.br","elidemendes@yahoo.com.br","k.tomaz1@hotmail.com","paulafmpassos@hotmail.com","luciaperes123@hotmail.com","alexandrinadias@icloud.com","lucianorachman@gmail.com"];

export default {
  async fetch(req, env) {
    // CORS preflight
    if (req.method === "OPTIONS") return new Response(null, { headers: CORS });

    const url = new URL(req.url);

    // ── Endpoint: verificar associado ──────────────────────────────────────────
    if (url.pathname === "/check-member" && req.method === "POST") {
      let body;
      try { body = await req.json(); } catch { return json({ ok:false, erro:"JSON inválido" }, 400); }
      const email = (body.email || "").trim().toLowerCase();
      if (!email) return json({ ok:false, erro:"email obrigatório" }, 400);
      // Aceita lista sobreposta via env (CSV) ou usa o default hardcoded
      const lista = env.MEMBER_EMAILS
        ? env.MEMBER_EMAILS.split(",").map(e => e.trim().toLowerCase())
        : MEMBER_EMAILS_DEFAULT;
      return json({ ok:true, membro: lista.includes(email) });
    }

    if (url.pathname !== "/extract" || req.method !== "POST") {
      return new Response("Not found", { status: 404, headers: CORS });
    }

    // Autenticação
    const apiKey = req.headers.get("X-Api-Key") || req.headers.get("X-API-Key");
    if (apiKey !== env.IBGESF_API_KEY) {
      return json({ ok: false, erro: "Unauthorized" }, 401);
    }

    // Parse do body
    let body;
    try {
      body = await req.json();
    } catch (e) {
      return json({ ok: false, erro: "JSON inválido no body" }, 400);
    }

    const tipo = (body.tipo || "texto").toLowerCase();

    // ──────────────────────────────────────────
    // MODO 1: PDF/DOCX nativo — múltiplos arquivos suportados
    // ──────────────────────────────────────────
    if (tipo === "pdf" || tipo === "doc" || tipo === "docx") {

      // Monta lista de arquivos a processar
      let arquivos = [];
      if (body.arquivos && Array.isArray(body.arquivos) && body.arquivos.length > 0) {
        // Múltiplos PDFs
        arquivos = body.arquivos.map(a => ({
          b64: a.base64 || a.arquivo_base64,
          mimeType: a.mime_type || "application/pdf",
          nome: a.nome || "documento.pdf",
        }));
      } else {
        // PDF único (retrocompatível)
        const b64 = body.arquivo_base64 || body.base64;
        const nome = body.nome || body.nomeArquivo || "documento.pdf";
        const mimeType = body.mime_type || (tipo === "pdf" ? "application/pdf" : "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        arquivos = [{ b64, mimeType, nome }];
      }

      // Valida e filtra arquivos
      arquivos = arquivos.filter(a => a.b64 && a.b64.length >= 100);
      if (arquivos.length === 0) {
        return json({ ok: false, erro: "arquivo_base64 ausente ou vazio" }, 400);
      }

      // Limite de tamanho por arquivo (~32MB PDF = ~44MB base64)
      for (const a of arquivos) {
        if (a.b64.length > 44_000_000) {
          return json({ ok: false, erro: `Arquivo "${a.nome}" muito grande (máx ~32MB). Comprima e tente novamente.` }, 413);
        }
      }

      // Processa cada arquivo e acumula resultados
      const resultados = [];
      const erros = [];
      for (const arquivo of arquivos) {
        const claudeBody = {
          model: "claude-haiku-4-5-20251001",
          max_tokens: 2000,
          messages: [{
            role: "user",
            content: [
              {
                type: "document",
                source: {
                  type: "base64",
                  media_type: arquivo.mimeType,
                  data: arquivo.b64,
                },
              },
              {
                type: "text",
                text: PROMPT_PDF + `\n\nArquivo: ${arquivo.nome}`,
              },
            ],
          }],
        };

        const r = await callClaudeRaw(claudeBody, env);
        if (r.ok) {
          resultados.push({ indicadores: r.indicadores, nome: arquivo.nome, usage: r.usage, modelo: r.modelo });
        } else {
          erros.push({ nome: arquivo.nome, erro: r.erro });
        }
      }

      if (resultados.length === 0) {
        return json({ ok: false, erro: "Falha ao processar todos os arquivos", detalhes: erros }, 502);
      }

      // Consolida múltiplos resultados: por indicador, escolhe o de maior confiança
      const indicadoresFinais = consolidar(resultados.map(r => r.indicadores));
      const totalTokens = resultados.reduce((s, r) => s + (r.usage?.input_tokens || 0) + (r.usage?.output_tokens || 0), 0);

      return json({
        ok: true,
        indicadores: indicadoresFinais,
        modelo: resultados[0].modelo,
        modo: "pdf-nativo",
        arquivos_processados: resultados.length,
        arquivos_com_erro: erros.length > 0 ? erros : undefined,
        total_tokens: totalTokens || undefined,
      });
    }

    // ──────────────────────────────────────────
    // MODO 2: Texto estruturado (XLSX→CSV, TXT, etc.)
    // ──────────────────────────────────────────
    const texto = (body.texto || body.text || "").slice(0, 50000);
    const nome = body.nome || body.nomeArquivo || "documento";
    if (texto.length < 50) {
      return json({ ok: false, erro: "Texto muito curto (mín 50 chars)" }, 400);
    }

    const claudeBodyTexto = {
      model: "claude-haiku-4-5-20251001",
      max_tokens: 1500,
      messages: [{
        role: "user",
        content: PROMPT_TEXTO + texto + `\n\nArquivo: ${nome}`,
      }],
    };

    const r = await callClaudeRaw(claudeBodyTexto, env);
    if (!r.ok) return json({ ok: false, erro: r.erro, detalhe: r.detalhe }, 502);

    return json({
      ok: true,
      indicadores: r.indicadores,
      modelo: r.modelo,
      modo: "texto",
      arquivos_processados: 1,
    });
  }
};

// ─── Consolida resultados de múltiplos PDFs ──────────────────────────────────

function consolidar(lista) {
  if (lista.length === 1) return lista[0];
  const resultado = {};
  for (const chave of CHAVES) {
    let melhor = null;
    let melhorPeso = -1;
    for (const ind of lista) {
      const item = ind[chave];
      if (!item) continue;
      const peso = CONFIANCA_PESO[item.confianca] ?? 0;
      if (item.valor_normalizado !== null && peso > melhorPeso) {
        melhor = item;
        melhorPeso = peso;
      }
    }
    resultado[chave] = melhor || { valor_normalizado: null, confianca: "nao-encontrado" };
  }
  return resultado;
}

// ─── helpers ─────────────────────────────────────────────────────────────────

async function callClaudeRaw(claudeBody, env) {
  let claudeResp;
  try {
    const r = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": env.ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify(claudeBody),
    });
    claudeResp = await r.json();
  } catch (e) {
    return { ok: false, erro: "Falha ao contactar API Anthropic: " + e.message };
  }

  const raw = claudeResp.content?.[0]?.text || "";
  if (!raw) {
    const apiErr = claudeResp.error?.message || JSON.stringify(claudeResp).slice(0, 200);
    return { ok: false, erro: "Resposta vazia da API", detalhe: apiErr };
  }

  let indicadores = null;
  try {
    indicadores = JSON.parse(raw);
  } catch (e) {
    const m = raw.match(/\{[\s\S]*\}/);
    if (m) {
      try { indicadores = JSON.parse(m[0]); } catch (e2) {}
    }
  }

  if (!indicadores) {
    return { ok: false, erro: "JSON não encontrado na resposta", raw: raw.slice(0, 300) };
  }

  // Garante que todas as 7 chaves existem
  for (const k of CHAVES) {
    if (!(k in indicadores)) indicadores[k] = { valor_normalizado: null, confianca: "nao-encontrado" };
  }

  return { ok: true, indicadores, modelo: claudeResp.model || claudeBody.model, usage: claudeResp.usage || null };
}

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { ...CORS, "Content-Type": "application/json" },
  });
}
