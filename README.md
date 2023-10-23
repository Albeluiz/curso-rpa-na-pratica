# curso-rpa-na-pratica
Automação de Lançamento de Notas Fiscais (Treinamento RPA na Prática)
Este projeto é uma automação desenvolvida em Python para simplificar o processo de lançamento de notas fiscais em um site de treinamento específico oferecido por Alex Diogo. Ele foi criado para fins de aprendizado e demonstração de automação de processos robóticos (RPA).

Descrição
A automação foi projetada para realizar as seguintes tarefas:

Iniciar um navegador da web (Google Chrome) e acessar o site de treinamento RPA na Prática de Alex Diogo.
Fazer login no site com credenciais pré-definidas.
Carregar um arquivo Excel contendo dados de notas fiscais a serem lançadas.
Preencher formulários da web com informações das notas fiscais a partir do arquivo Excel.
Realizar validações e seleções dinâmicas com base nas informações fornecidas.
Emitir as notas fiscais no site.
Registrar qualquer erro ou problema no próprio arquivo Excel.
A automação utiliza bibliotecas como Selenium para controle do navegador, PyAutoGUI para interações na interface gráfica, e Pandas para manipulação de dados em formato de tabela. O projeto é acompanhado de uma interface gráfica mínima que exibe mensagens informativas no início e no final da execução.

Pré-requisitos
Antes de executar o projeto, você precisará instalar as seguintes bibliotecas Python:

selenium para automação do navegador.
pyautogui para interações com a interface gráfica.
pandas para manipulação de dados em formato de tabela.
tkinter para a interface gráfica.
Pillow (PIL) para manipulação de imagens.
Certifique-se de que o driver do Chrome WebDriver esteja instalado e configurado corretamente.

Como Usar
Clone este repositório em sua máquina local.
Instale as bibliotecas Python necessárias.
Configure o Chrome WebDriver ou verifique se ele está no PATH do sistema.
Coloque seu arquivo Excel com os dados das notas fiscais no mesmo diretório do projeto.
Execute o script Python automacao_lancamento_notas.py.
Contribuições
Este projeto é aberto a contribuições e melhorias. Sinta-se à vontade para enviar pull requests ou relatar problemas.

Licença
Este projeto é distribuído sob a Licença MIT. Consulte o arquivo LICENSE para obter detalhes.

Esta descrição reflete as revisões solicitadas e destaca o site de treinamento de Alex Diogo como o alvo da automação. Certifique-se de adaptar o conteúdo acima ao seu projeto e suas informações específicas.





