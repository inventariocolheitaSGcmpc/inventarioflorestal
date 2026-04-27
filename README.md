# Inventario Florestal no GitHub Pages

Este projeto foi adaptado do app original em R Shiny para um site estatico compativel com GitHub Pages.

## O que esta funcionando online

- mapa interativo dos talhoes;
- filtro por fazenda;
- selecao de talhoes clicando no mapa;
- sequencia visual com numeracao e linha logistica;
- resumo de area e volume;
- exportacao da sequencia em Excel no navegador;
- botao para preparar um e-mail local com assunto e texto prontos.

## Limitacao importante

Como o site roda apenas no GitHub Pages, ele nao tem um servidor R por tras. Por isso:

- o envio automatico de e-mail com anexo, igual ao Shiny, nao existe no GitHub puro;
- em vez disso, o site baixa o arquivo Excel e abre um rascunho de e-mail no programa padrao do usuario.

Se um dia voce quiser voltar ao envio automatico real, sera preciso usar um backend ou um servico externo de e-mail.

## Estrutura principal

- `index.html`: pagina principal do site.
- `style.css`: visual do site.
- `script.js`: logica do mapa, filtro, selecao e exportacao.
- `site-data/base_info.geojson`: dados geograficos e indicadores usados no mapa.
- `site-data/summary.json`: resumo para a interface.
- `convert_data.R`: script que reconstrui os dados web a partir do Excel e do GPKG originais.
- `data/`: arquivos fonte originais.

## Como atualizar os dados

Quando voce trocar a planilha ou o geopackage, execute:

```r
source("convert_data.R")
```

Se preferir pelo terminal:

```bash
"C:\Program Files\R\R-4.4.0\bin\Rscript.exe" convert_data.R
```

## Como publicar no GitHub

Na pasta do projeto, rode:

```bash
git init
git add .
git commit -m "Versao web estatica do inventario florestal"
git branch -M main
git remote add origin https://github.com/pietropmorales/inventarioflorestal.git
git push -u origin main
```

## Como ativar o GitHub Pages

Depois de subir o repositorio:

1. Abra o repositorio `inventarioflorestal` no GitHub.
2. Entre em `Settings`.
3. Clique em `Pages`.
4. Em `Build and deployment`, escolha:
   `Source` = `Deploy from a branch`
5. Selecione:
   `Branch` = `main`
   `Folder` = `/ (root)`
6. Salve.

O GitHub vai gerar um link publico parecido com:

`https://pietropmorales.github.io/inventarioflorestal/`

## Observacoes

- O primeiro carregamento pode ser um pouco pesado porque o `GeoJSON` e grande.
- O arquivo `app.R` foi mantido apenas como referencia da logica original.
- Para melhorar ainda mais a performance no futuro, eu recomendo simplificar a geometria ou dividir os dados por fazenda.
