# v0.5.0 — Memória de cálculo detalhada

- Conversão CV/HP para kW e explicitação de rendimento e fator de potência por extenso.
- Motores: potências mecânica, ativa, aparente e reativa. Se os parâmetros variam entre motores, somam-se P e Q e calcula-se S resultante por √(P²+Q²).
- Iluminação: fator de demanda Fd editável (padrão 100%), FpI editável e fórmula Dm = C × Fd / FpI.
- Outros equipamentos: fórmulas gerais para kW/W e kVA/VA, conforme parâmetros disponíveis.
- Memória de cálculo compartilhada entre PDF e DOCX.
- Valores técnicos são estimativas; Fd de 100% não implica adequação normativa a todas as instalações.
