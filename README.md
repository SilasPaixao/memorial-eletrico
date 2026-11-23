# Memorial Elétrico — Protótipo 0.1.0

Aplicativo desktop offline para Windows, com relação de cargas, estimativas elétricas e exportação de **rascunhos** DOCX/PDF.

## O que já existe
- Interface em português para consumidor, fornecimento e cargas.
- Equipamentos iniciais iguais aos tipos do memorial de referência: motores elétricos e iluminação.
- Novos equipamentos adicionados manualmente.
- Conversões CV/HP/W/kW/VA/kVA, distinção entre potência mecânica, ativa e aparente.
- Premissas opcionais de rendimento e fator de potência, sem inventar dados de placa.
- Exportação DOCX, impressão para PDF, backup/importação JSON.
- Salvamento automático local via localStorage (ainda **não SQLite**).
- Testes unitários e workflow de compilação Windows.

## O que ainda NÃO existe
SQLite, cadastro de múltiplos projetos, biblioteca persistente independente, relatório de cálculo completo, dimensionamento normativo, regras Coelba versionadas, aprovação, assinatura por imagem, PDF direto sem diálogo de impressão, validação por engenheiro eletricista. Os documentos são sempre RASCUNHO.

O conteúdo e a estrutura do memorial de referência orientam o protótipo, mas **não há reprodução fiel de sua diagramação**. Não reproduzimos como correta a estimativa de 32,02 kVA presente no documento: 37 CV, rendimento 90% e FP 0,85 implicam aproximadamente 35,57 kVA antes de demanda.

## Como executar
Requer Node.js 22+, Rust e dependências Tauri 2 (no Windows: Microsoft C++ Build Tools e WebView2).

```bash
npm ci
npm test
npm run dev
npm run tauri dev
```

## Como gerar .exe pelo GitHub
1. Extraia o ZIP e envie **o conteúdo da pasta** à raiz do repositório `SilasPaixao/memorial-eletrico`.
2. Acesse GitHub → Actions → **Compilar instalador Windows** → Run workflow.
3. Quando a execução concluir com sucesso, abra a execução e baixe o artefato `memorial-eletrico-installer-windows`.
4. Dentro do artefato estará o instalador NSIS `.exe`.

**Não há garantia de que o instalador funcione em todo Windows.** Testar em Windows 10/11 x64 e verificar pré-requisitos e assinatura de código antes de distribuir ao cliente.

## Responsabilidade técnica
Não usar este protótipo como projeto aprovado, laudo, ART/TRT, nem para selecionar automaticamente cabos, disjuntores ou medição. Profissional habilitado deve validar dados e regras técnicas vigentes.
