# Organizacao da pasta PythonSetorial

Os scripts estao separados pelo setor que atualizam. Cada pipeline permanece na
mesma pasta dos scripts que executa, preservando a resolucao relativa existente.

## Pastas

- `BeefBR/`: atualizacoes e pipeline do BeefBR.
- `BeefUS/`: atualizacoes e pipeline do BeefUS. A pasta `data_usda_ers/` fica
  aqui porque e gerada pelo coletor de forecasts.
- `FrangoBR/`: atualizacoes e pipeline do FrangoBR.
- `FrangoUS/`: atualizacoes e pipeline do FrangoUS.
- `Outros/`: rotinas gerais, como dashboard e atualizacao da Selic.
- `documentacao/`: notas e referencias de apoio.
- `_arquivo/`: logs, relatorios e artefatos antigos preservados por seguranca.

## Pipelines principais

Execute cada pipeline a partir de qualquer pasta; ele resolve seus scripts pela
localizacao do proprio arquivo.

```powershell
python .\BeefBR\atualizar_beefbr_pipeline.py
python .\BeefUS\atualizar_beefus_pipeline.py
python .\FrangoBR\atualizar_frangobr_pipeline.py
python .\FrangoUS\atualizar_frangous_pipeline.py
```

## Observacoes

- `requirements.txt` permanece na raiz por ser compartilhado.
- Nenhum script foi apagado durante a reorganizacao.
- Os caminhos das planilhas continuam sendo resolvidos pelo nome do arquivo
  abaixo da raiz estavel do Google Drive.
