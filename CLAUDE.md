# SISTEMA ARI — GeraLaudo

App desktop em Python para geração automática de laudos periciais e impugnações no estilo de Ari Vladimir Copesco Junior.

## Estrutura do projeto

```
GeraLaudo/
├── app.py              # App principal (GUI + lógica)
├── requirements.txt    # Dependências Python
├── instalar.bat        # Instala dependências (rodar 1x)
├── iniciar.bat         # Inicia o app
├── CLAUDE.md           # Este arquivo
└── .vscode/
    ├── settings.json   # Interpetador Python + formatação
    └── extensions.json # Extensões recomendadas
```

## Como rodar

```
python3.12 app.py
```

## Dependências instaladas

- `anthropic` — SDK Claude API
- `customtkinter` — GUI desktop moderna
- `python-docx` — leitura e escrita de .docx
- `Pillow` — processamento de imagens
- `pdfplumber` — extração de texto de PDF

## Variáveis importantes

- `MODEL = "claude-sonnet-4-6"` — modelo Claude usado (linha ~29 do app.py)
- `SISTEMA_ARI_PATH` — caminho base dos arquivos de perfil do Ari
- `MAX_PHOTOS = 20` — limite de fotos por laudo
- `MAX_IMAGE_SIZE = (1280, 1280)` — resize das fotos antes de enviar
- `CONFIG_FILE` — salvo em `~/.sistema_ari/config.json` (chave API + pasta padrão)

## Fluxo atual — Geração de Laudo

1. Usuário seleciona pré-laudo (secretárias) + anotações de campo
2. Usuário seleciona pasta de fotos (nome do arquivo = descrição da foto)
3. App lê os documentos, codifica as fotos em base64
4. Envia tudo ao Claude com o AGENTE_HUMANIZADOR.md + PERFIL_ESCRITA_ARI.md como system prompt
5. Claude gera o laudo completo nas 12 seções no estilo do Ari
6. App salva como .docx formatado na pasta escolhida

## Arquivos de perfil do Ari (lidos automaticamente)

Localizados em `C:\Users\ari\OneDrive\Documentos\SISTEMA ARI\`:
- `AGENTE_HUMANIZADOR.md` — regras de estilo de escrita
- `PERFIL_ESCRITA_ARI.md` — estrutura detalhada dos laudos

## PRÓXIMAS FUNCIONALIDADES A IMPLEMENTAR

### 1. Módulo de Impugnações
- Receber o laudo do perito adversário (upload de .docx/.pdf)
- Identificar automaticamente o tipo de enquadramento (ruído, vibração, calor, RNI, hidrocarbonetos)
- Gerar impugnação no estilo padrão do Ari como Assistente Técnico da Reclamada
- Estrutura padrão da impugnação:
  - Cabeçalho com identificação (EXMO SR. JUIZ..., Processo n°, Reclamante, Reclamada)
  - Identificação do erro técnico do Expert
  - Argumentação com citação de normas (NR, NHO, ACGIH)
  - Quesitos Complementares numerados
  - Encerramento: "OBS: A Reclamada se reserva o direito..." + "Termos em que, P. Deferimento."
- Salvar como .docx

### 2. Melhorias no Laudo
- Pré-visualização do laudo dentro do app antes de salvar
- Edição inline de campos (número do processo, data, partes)
- Histórico dos últimos laudos gerados

### 3. Interface
- Abas separadas: [Laudo Pericial] | [Impugnação]
- Barra lateral com histórico de processos

## Contexto do usuário

- **Nome:** Ari Vladimir Copesco Junior
- **CREA:** 060097553-3
- **Cargo:** Engenheiro de Segurança do Trabalho | Perito Judicial TRT 15
- **Vara:** 3ª Vara do Trabalho de Ribeirão Preto
- **Honorários:** R$ 4.500,00 por laudo
- **Sistema base:** `C:\Users\ari\OneDrive\Documentos\SISTEMA ARI\`

## Fluxo de trabalho do Ari

1. Secretárias preparam o **pré-laudo** (dados da inicial, agentes alegados, documentos do processo)
2. Ari vai ao local com o pré-laudo e coleta dados de campo (medições, fotos, relatos)
3. Pré-laudo + anotações de campo → app gera o laudo final

## Observações técnicas

- Python instalado via Microsoft Store: executável em `python3.12`
- Pacotes instalados em modo usuário (sem permissão de administrador)
- GUI usa customtkinter 5.2.2 com tema claro e cores azul-marinho (#1a3a6b)
- Saída .docx usa fonte Arial 12pt, margens 1.2" laterais
