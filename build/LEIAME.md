# Tabelão Yuny: como montar

O site publicado (`index.html` na raiz) é gerado. Não edite o `index.html` à mão: edite os dados ou o modelo e rode a montagem.

| Arquivo | O que é |
|---|---|
| `dados.json` | Fonte única: menu, empreendimentos (nome, endereço, fase, logo, foto, botões, tours, ficha `fx`), tabelas de preço, widget de disponibilidade e datas. |
| `template.html` | Visual (estilo yuny.com.br) e o JavaScript do Tabelão. |
| `montar.py` | `dados.json` + `template.html` → `../index.html`. |
| `atualizar_mes.py` | PDFs do mês + VIMOB → atualiza tabelas, widget e datas no `dados.json`. |

## Atualização mensal
1. Copie os PDFs do mês e as planilhas do VIMOB para uma pasta de trabalho.
2. Em `atualizar_mes.py`, ajuste o bloco `CFG` (pasta, mês, arquivos, datas) e os rótulos de fluxo `STD`/`PAY` conforme o cabeçalho de cada PDF.
3. `python3 atualizar_mes.py`: leia o log. `pdf_vs_plan` = status diferente entre PDF e VIMOB (vale o VIMOB); `disp_sem_pdf` = livre no VIMOB sem preço no PDF (avisar); `CONFERIR` = área × R$/m² não fecha com o total (corrigir o campo errado no dados.json).
4. `python3 montar.py`, testar (desktop e 390 px) e publicar com `git push` no `main`.

## Regras
- Público: nada de comissão, política, prêmio, coordenador ou telefone.
- Fora: Quaddra Jesuíno. 1125 Joaquim só como "100% vendido". Natury com valores do VIMOB até haver tabela oficial.
- Fichas (`fx`): ficha técnica do site yuny.com.br; diferenciais e plantas do Portal do Corretor. Plantas usam as imagens públicas `yuny.com.br/media/...`.
