# Sets Sync Tool

## Visão geral

Esta ferramenta converte planilhas Excel em arquivos XML compatíveis com o Navisworks.

Ela oferece dois fluxos de trabalho principais:
- `Appearance Profiler`: exporta regras de aparência para XML.
- `Search Sets`: converte um template de Search Sets em XML.

## Uso

### Appearance Profiler

1. Selecione o arquivo Excel de regras.
2. Informe o nome do arquivo de saída XML.
3. Opcionalmente ajuste o nome do perfil.
4. Clique em `Exportar XML`.

### Search Sets

1. Selecione o arquivo Excel com o template de Search Sets.
2. Informe o arquivo de saída XML.
3. Clique em `Exportar XML`.

## Modelos de arquivo

O repositório inclui os arquivos de modelo:
- `Template_Appearence.xlsx`
- `Template_SearchSet.xlsx`

Use-os como referência para formatar seus próprios arquivos Excel.

## Saída

Os arquivos XML gerados podem ser importados no Navisworks como Appearance Profiler ou Search Sets.

## Observações

- O diretório `Output/` está disponível para armazenar arquivos gerados.
- Caso prefira rodar a conversão de Search Sets sem a interface gráfica, o script `excel_to_nw_search.py` contém a função `convert_excel_to_xml`.