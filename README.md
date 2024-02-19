
# Melhorias no Cadastro de Ativos - PS

Solução de melhoria desenvolvida para uma planilha com histórico de ativos.



## Novas Funcionalidades

- Formulário CRUD desenvolvido em VBA
    - Cadastro de um novo ativo
                    
            * Definição de campos obrigatórios a serem preenchidos
                    - Número do Imobilizado
                    - Denominação
                    - Local
                    - Responsável Atual

            * Adição do novo registro na planilha "Base de Dados"
            * Validação para não permitir o registro de um Ativo sem o número do Imobilizado
    - Consulta de um Ativo
            * Validação para verificar se o ativo já está cadastrado na planilha "Base de Dados"

            * Busca realizada apenas através do número do Imobilizado
    - Atualização de um Ativo

            * Busca realizada através do número do Imobilizado
            * Alteração autorizada de todos os campos
            * Adição do Registro Atualizado na planilha "Base de Dados"
    - Remoção de um Ativo
            
            * Para remoção do ativo é necessário realizar processos e obter autorizações através do SAP,
              entende-se que, o ativo será removido da planilha apenas após todas autorizações.
            * Remoção através do número do Imobilizado
            * O ativo não é removido de forma definitiva, ele solicita o nome de um responsável,
              remove na planilhaBase de Dados e leva para a planilha "Ativos Removidos", mantendo, além dos dados do Ativo,
              o responsável pela remoção e a data/horário.


## Como rodar?

