# Agrupador-de-Arquivos

Aqui está a descrição ajustada, destacando a disponibilidade do código e do arquivo para uso por outros:

---

O **Agrupador de Planilhas** é uma ferramenta desenvolvida para organizar e consolidar arquivos Excel em grande escala, útil tanto para relatórios já salvos quanto para dados obtidos por **web scraping**. Ele agrupa arquivos com o padrão de nome *settlement_detail_report_batch_*_*.xlsx*, onde "*" representa números de referência sequenciais. O agrupador permite que o usuário selecione uma pasta, localize automaticamente todos os arquivos correspondentes, e consolide os dados em um único arquivo Excel.

### Funcionalidades principais:

1. **Aplicação para Web Scraping**: A ferramenta é ideal para gerenciar arquivos coletados por web scraping, desde que o padrão de nome dos arquivos esteja de acordo com a estrutura usada no código. Ela automatiza o agrupamento e organização de relatórios sequenciais.
2. **Nome Personalizável para Arquivos em Série**: O código permite que o padrão de nome seja facilmente ajustado, desde que a lógica de identificação de série seja mantida.
3. **Seleção de Pasta**: O usuário pode escolher a pasta onde os arquivos estão armazenados, facilitando a organização.
4. **Identificação Automática e Ignoração de Duplicatas**: A ferramenta encontra e organiza os arquivos no padrão definido, ignorando arquivos repetidos para evitar duplicação de dados.
5. **Verificação de Arquivos Faltantes**: Em caso de arquivos ausentes na sequência numérica, um alerta é emitido, recomendando a recuperação do(s) arquivo(s) faltante(s) antes de refazer o agrupamento.
6. **Combinação de Dados Consistente**: Todos os arquivos são combinados em um DataFrame, preservando o cabeçalho e organizando os dados de maneira estruturada.
7. **Exportação Final**: O arquivo consolidado é salvo como "resultado_agrupado.xlsx" na mesma pasta dos arquivos originais.

Disponibilizo tanto o **código-fonte** quanto o **arquivo do agrupador**, caso alguém tenha interesse em utilizar ou adaptar para suas próprias necessidades.
