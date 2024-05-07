document.addEventListener('DOMContentLoaded', function() {
    // Selecionando elementos do DOM
    const entradaArquivo = document.getElementById('entradaArquivo');
    const modal = document.getElementById('modal');
    const textoModal = document.getElementById('texto-modal');
    const span = document.getElementsByClassName("fechar")[0];

    // Event listener para quando um arquivo é selecionado
    entradaArquivo.addEventListener('change', function(e) {
        const arquivo = e.target.files[0];
        if (arquivo) {
            const leitor = new FileReader();
            leitor.onload = function(evento) {
                try {
                    // Lendo o arquivo Excel
                    const dados = evento.target.result;
                    const planilha = XLSX.read(dados, {
                        type: 'binary'
                    });
                    const nomePlanilha = planilha.SheetNames[0];
                    const folha = planilha.Sheets[nomePlanilha];
                    const linhas = XLSX.utils.sheet_to_row_object_array(folha);
                    // Exibindo os resultados na página
                    exibirResultados(linhas);
                } catch (erro) {
                    alert('Erro ao ler a planilha: ' + erro.message);
                }
            };
            leitor.readAsBinaryString(arquivo);
        } else {
            alert('Por favor, selecione uma planilha para carregar.');
        }
    });

    // Função para exibir os resultados na página
    function exibirResultados(linhas) {
        const resultados = document.getElementById('resultados');
        let totalOrdens = 0; // Contador de total de ordens de serviço

        resultados.innerHTML = ''; // Limpa resultados anteriores
        const assuntos = {}; // Objeto para contar os assuntos
        const colaboradoresEspecificos = {
            // Lista de colaboradores específicos e suas ordens de serviço
            'ALAN NORONHA FERREIRA': [], 'BRUNO MARTINS CARDOSO': [], 'ISMAEL ALVES GRAIA': [],
            'LUCAS FIGUEIREDO DOS REIS': [], 'MARCELO APARECIDO PEREIRA': [], 'MASSIVA': [],
            'PAULO CESAR MARCELLINO': [], 'PEDRO GOMES DE LIMA': [], 'ROBISON RAMOS': [],
            'RODRIGO SANTOS GUIMARAES': [], 'RONNY DA SILVA LUZ': [], 'SAMUEL SANTOS DE ARAUJO': [],
            'THIAGO PEREIRA CAMARGOS': [], 'WENDER RENS MIRANDA BELISARIO': [], 'PAULO RODRIGUES DOS SANTOS': [], 'JOSÉ DE FREITAS DA SILVA NETO': []
        };

        // Iterando sobre cada linha da planilha
        linhas.forEach(linha => {
            if (colaboradoresEspecificos.hasOwnProperty(linha.Colaborador)) {
                // Construindo os detalhes da ordem de serviço
                let detalhes = `Assunto: ${linha.Assunto}, Cliente: ${linha.Cliente}, Endereço: ${linha.Endereço}, Fechamento: ${linha.Fechamento ? numeroSerieParaData(linha.Fechamento) : 'Data de Fechamento não consta na planilha IXC'}`;
                colaboradoresEspecificos[linha.Colaborador].push(detalhes); // Adicionando detalhes às ordens do colaborador
                totalOrdens++; // Incrementando o total de ordens
                assuntos[linha.Assunto] = (assuntos[linha.Assunto] || 0) + 1; // Contando os assuntos
            }
        });

        // Ordenando colaboradores por quantidade de ordens
        const colaboradoresOrdenados = Object.entries(colaboradoresEspecificos)
            .filter(([nome, ordens]) => ordens.length > 0)
            .sort((a, b) => b[1].length - a[1].length);

        // Criando div para resumo por colaborador
        const divResumoColaborador = document.createElement('div');
        divResumoColaborador.classList.add('resumo-colaborador');

        // Construindo resumo por colaborador
        if (colaboradoresOrdenados.length > 0) {
            const divColaboradores = document.createElement('div');
            divColaboradores.innerHTML = '<h3>Resumo por Colaborador</h3>';
            colaboradoresOrdenados.forEach(([nome, ordens]) => {
                const divColaborador = document.createElement('div');
                divColaborador.innerHTML = `<span class="nome">${nome}:</span> <span class="ordens">${ordens.length} ordens de serviço </span>`;
                const divBotao = document.createElement('div');
                divBotao.classList.add('botao-colaborador');
                const botao = document.createElement('button');
                botao.innerText = 'Ver Detalhes';
                botao.onclick = function() {
                    // Exibindo detalhes no modal
                    modal.style.display = 'block';
                    const conteudoTabela = `<table><tr><th>Assunto</th><th>Cliente</th><th>Endereço</th><th>           </th><th>Fechamento</th></tr>` +
                        ordens.map(ordem => {
                            const detalhesOrdem = ordem.split(', ').map(o => o.split(': '));
                            return `<tr>${detalhesOrdem.map(([chave, valor]) => `<td>${valor || ''}</td>`).join('')}</tr>`;
                        }).join('') + '</table>';
                    textoModal.innerHTML = `<strong>Ordens de Serviço para ${nome}:</strong>${conteudoTabela}`;
                };
                divBotao.appendChild(botao);
                divColaborador.appendChild(divBotao);
                divColaboradores.appendChild(divColaborador);
                divColaboradores.appendChild(document.createElement('br'));
            });
            divResumoColaborador.style.padding = '20px';
            divResumoColaborador.appendChild(divColaboradores);
        }

        resultados.appendChild(divResumoColaborador);

        // Criando div para resumo por assunto
        const divResumoAssuntos = document.createElement('div');
        divResumoAssuntos.classList.add('resumo-assuntos');

        // Construindo resumo por assunto
        if (Object.keys(assuntos).length > 0) {
            const divAssuntos = document.createElement('div');
            divAssuntos.innerHTML = '<h3>Resumo por Assunto</h3>';
            Object.entries(assuntos)
                .sort((a, b) => b[1] - a[1])
                .forEach(([assunto, contagem]) => {
                    if (contagem > 0) {
                        const divAssunto = document.createElement('div');
                        divAssunto.innerHTML = `<span class="nome">${assunto}:</span> <span class="ordens">${contagem} vez(es) </span>`;
                        const divBotao = document.createElement('div');
                        divBotao.classList.add('botao-colaborador');
                        const botao = document.createElement('button');
                        botao.innerText = 'Ver Detalhes';
                        botao.onclick = function() {
                            // Exibindo detalhes no modal
                            modal.style.display = 'block';
                            const ordensAssunto = linhas.filter(linha => linha.Assunto === assunto && colaboradoresEspecificos.hasOwnProperty(linha.Colaborador));
                            const conteudoTabela = `<table><tr><th>Colaborador</th><th>Cliente</th><th>Endereço</th></tr>` +
                                ordensAssunto.map(ordem => {
                                    return `<tr><td>${ordem.Colaborador}</td><td>${ordem.Cliente}</td><td>${ordem.Endereço}</td><td>${ordem.Fechamento ? numeroSerieParaData(ordem.Fechamento) : 'Data de fechamento não consta'}</td></tr>`;
                                }).join('') + '</table>';
                            textoModal.innerHTML = `<strong>Detalhes para o Assunto "${assunto}":</strong>${conteudoTabela}`;
                        };
                        divBotao.appendChild(botao);
                        divAssunto.appendChild(divBotao);
                        divAssuntos.appendChild(divAssunto);
                        divAssuntos.appendChild(document.createElement('br'));
                    }
                });
            divResumoAssuntos.style.padding = '20px';
            divResumoAssuntos.appendChild(divAssuntos);
        }

        resultados.appendChild(divResumoAssuntos);

        // Adicionando o total de ordens de serviço
        const divTotalOrdens = document.createElement('p');
        divTotalOrdens.classList.add('total-ordens');
        divTotalOrdens.innerHTML = `<strong>Total de ordens de serviço:</strong> ${totalOrdens}`;
        resultados.appendChild(divTotalOrdens);

        // Evento de fechar modal pelo botão
        span.onclick = function() {
            modal.style.display = 'none';
        };

        // Evento de fechar modal clicando fora dele
        window.onclick = function(evento) {
            if (evento.target == modal) {
                modal.style.display = 'none';
            }
        };
    }

    // Evento de fechar modal pelo teclado (tecla Esc)
    document.addEventListener('keydown', function(evento) {
        if (evento.key === 'Escape') {
            modal.style.display = 'none';
        }
    });

    // Função para converter número de série em data
    function numeroSerieParaData(numeroSerie) {
        const dataBaseExcel = new Date(1900, 0, 1);
        const data = new Date(dataBaseExcel.getTime() + numeroSerie * 24 * 60 * 60 * 1000);
        return data.toLocaleDateString();
    }
});
