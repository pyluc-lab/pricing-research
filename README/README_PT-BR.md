# Pesquisa de Preço com Selenium

## Descrição

Este programa foi desenvolvido para funcionar nas versões brasileiras dos sites. O Selenium é extremamente sensível a mudanças feitas pelos desenvolvedores dos sites. Qualquer mudança nos sites que altere o XPATH, nomes de classes, IDs, etc., exigirá uma pequena revisão neste código para também atualizar os XPATH, nomes de classes, IDs, etc., nas funções correspondentes ao site.

Este projeto foi desenvolvido para realizar buscas automatizadas de preços de produtos em plataformas online como Google Shopping, Mercado Livre e Amazon. Utiliza Selenium, uma ferramenta robusta para automação de navegadores, permitindo uma navegação direta nas páginas de interesse. Embora a tarefa de comparação de preços pudesse ser realizada via APIs, o uso do Selenium foi intencional para testar e aprimorar habilidades com essa biblioteca, focando em manipulações de DOM, extração de dados e automação de processos em ambiente real de navegação.

O sistema capta informações que são configuradas em um arquivo Excel (`search.xlsx`), onde o usuário pode especificar o produto, faixa de preços desejada e termos que devem ser excluídos para refinar a pesquisa. Os resultados são exportados como arquivos `.xlsx` e encaminhados automaticamente para um e-mail designado.

## Funcionalidades

* **Automação com Selenium:** Pesquisa em tempo real acessando as páginas do navegador e interagindo diretamente com os elementos do site.
* **Configuração de Pesquisa via Excel:** Personalize produtos de interesse, preços limites e condições de exclusão de termos.
* **Exportação Integrada:** Os resultados são exportados em Excel.
* **Envio Automático de E-mail:** Integração com Outlook para enviar resultados diretamente por e-mail.
* **Log Detalhado:** Sistema de logging que monitora erros e fornece informações sobre o status da execução.

```shell
> data_base
    search.xlsx
> Logs
    app.log
> modules
    __init__.py
    constants.py
> README
    readme eng
    readme port
> results
    (arquivos serão salvos aqui após a pesquisa)
main.py
Copy
```

### Descrição dos Diretórios e Arquivos

* **data_base/** : Contém arquivos cruciais para a definição da pesquisa, exemplo: `search.xlsx`.
* **Logs/** : Mantém um registro das atividades da aplicação no arquivo `app.log`.
* **modules/** : Inclui módulos essenciais como `__init__.py` e `constants.py`.
* **constants.py** : O usuário deve inserir um e-mail válido para receber os arquivos de resultados.
* **README/** : Inclui a documentação completa do projeto disponível em inglês e português.
* **results/** : Diretório designado para armazenar os arquivos de resultados `.xlsx` após a execução da pesquisa.
* **main.py** : Arquivo principal que inicia a execução do projeto.

## Instruções de Uso

1. **Configuração do E-mail:**

   * Abra `modules/constants.py`.
   * Insira um endereço de e-mail válido para garantir o recebimento dos arquivos gerados.
2. **Configuração da Pesquisa no `search.xlsx`:**

   * Abra o arquivo `search.xlsx` localizado em `data_base/`.
   * Insira o nome do produto que deseja pesquisar.
   * Defina o preço mínimo e máximo desejado para a compra.
   * Adicione palavras-chave que, se encontradas, desconsiderarão a pesquisa para evitar resultados irrelevantes.
3. **Execução do Projeto:**

   * Execute `main.py` para iniciar a pesquisa de preços.
   * Verifique o diretório `results/` para acessar os arquivos `.xlsx` gerados.
   * O log das operações estará disponível em `Logs/app.log`.
