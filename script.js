document.addEventListener('DOMContentLoaded', function() {
    const entradaArquivo = document.getElementById('entradaArquivo');
    const modal = document.getElementById('modal');
    const textoModal = document.getElementById('texto-modal');
    const span = document.getElementsByClassName("fechar")[0];

    entradaArquivo.addEventListener('change', function(e) {
        const arquivo = e.target.files[0];
        if (arquivo) {
            const leitor = new FileReader();
            leitor.onload = function(evento) {
                try {
                    const dados = evento.target.result;
                    const planilha = XLSX.read(dados, {
                        type: 'binary'
                    });
                    const nomePlanilha = planilha.SheetNames[0];
                    const folha = planilha.Sheets[nomePlanilha];
                    const linhas = XLSX.utils.sheet_to_row_object_array(folha);
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

    function exibirResultados(linhas) {
        const resultados = document.getElementById('resultados');
        let totalOrdens = 0;
        resultados.innerHTML = '';
        const assuntos = {};
        const colaboradoresEspecificos = {
            'COLABORADOR 1': [], 'COLABORADOR 2': [], 'COLABORADOR 3': [],
            'COLABORADOR 4': [], 'COLABORADOR 5': [], 'COLABORADOR 6': [],
            'COLABORADOR 7': [], 'COLABORADOR 8': [], 'COLABORADOR 9': [],
            'COLABORADOR 10': [], 'COLABORADOR 11': [], 'COLABORADOR 12': [],
            'COLABORADOR 13': [], 'COLABORADOR 14': [], 'COLABORADOR 15': [], 'COLABORADOR 16': []
        };

        linhas.forEach(linha => {
            if (colaboradoresEspecificos.hasOwnProperty(linha.Colaborador)) {
                let detalhes = `Assunto: ${linha.Assunto}, Cliente: ${linha.Cliente}, Endereço: ${linha.Endereço}, Fechamento: ${linha.Fechamento ? numeroSerieParaData(linha.Fechamento) : 'Data de Fechamento não consta na planilha IXC'}`;
                colaboradoresEspecificos[linha.Colaborador].push(detalhes);
                totalOrdens++;
                assuntos[linha.Assunto] = (assuntos[linha.Assunto] || 0) + 1;
            }
        });

        const colaboradoresOrdenados = Object.entries(colaboradoresEspecificos)
            .filter(([nome, ordens]) => ordens.length > 0)
            .sort((a, b) => b[1].length - a[1].length);

        const divResumoColaborador = document.createElement('div');
        divResumoColaborador.classList.add('resumo-colaborador');

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

        const divResumoAssuntos = document.createElement('div');
        divResumoAssuntos.classList.add('resumo-assuntos');

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
                        const botaoVerDetalhes = document.createElement('button');
                        botaoVerDetalhes.innerText = 'Ver Detalhes';
                        botaoVerDetalhes.onclick = function() {
                            modal.style.display = 'block';
                            const ordensAssunto = linhas.filter(linha => linha.Assunto === assunto && colaboradoresEspecificos.hasOwnProperty(linha.Colaborador));
                            const conteudoTabela = `<table><tr><th>Colaborador</th><th>Cliente</th><th>Endereço</th><th>Fechamento</th></tr>` +
                                ordensAssunto.map(ordem => {
                                    return `<tr><td>${ordem.Colaborador}</td><td>${ordem.Cliente}</td><td>${ordem.Endereço}</td><td>${ordem.Fechamento ? numeroSerieParaData(ordem.Fechamento) : 'Data de fechamento não consta'}</td></tr>`;
                                }).join('') + '</table>';
                            textoModal.innerHTML = `<strong>Detalhes para o Assunto "${assunto}":</strong>${conteudoTabela}`;
                        };
                        divBotao.appendChild(botaoVerDetalhes);
                        
                        const botaoExportar = document.createElement('button');
                        botaoExportar.innerText = 'Exportar Planilha';
                        botaoExportar.onclick = function() {
                            const ordensAssunto = linhas.filter(linha => linha.Assunto === assunto && colaboradoresEspecificos.hasOwnProperty(linha.Colaborador));
                            exportarPlanilha(ordensAssunto, assunto);
                        };
                        divBotao.appendChild(botaoExportar);

                        divAssunto.appendChild(divBotao);
                        divAssuntos.appendChild(divAssunto);
                        divAssuntos.appendChild(document.createElement('br'));
                    }
                });
            divResumoAssuntos.style.padding = '20px';
            divResumoAssuntos.appendChild(divAssuntos);
        }

        resultados.appendChild(divResumoAssuntos);

        const divTotalOrdens = document.createElement('p');
        divTotalOrdens.classList.add('total-ordens');
        divTotalOrdens.innerHTML = `<strong>Total de ordens de serviço:</strong> ${totalOrdens}`;
        resultados.appendChild(divTotalOrdens);

        span.onclick = function() {
            modal.style.display = 'none';
        };

        window.onclick = function(evento) {
            if (evento.target == modal) {
                modal.style.display = 'none';
            }
        };
    }

    document.addEventListener('keydown', function(evento) {
        if (evento.key === 'Escape') {
            modal.style.display = 'none';
        }
    });

    function numeroSerieParaData(numeroSerie) {
        const dataBaseExcel = new Date(1899, 11, 30);
        const data = new Date(dataBaseExcel.getTime() + numeroSerie * 24 * 60 * 60 * 1000);
        return data.toLocaleDateString();
    }

    // Função para exportar os dados para uma planilha
    function exportarPlanilha(ordensAssunto, assunto) {
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(ordensAssunto.map(ordem => ({
            Colaborador: ordem.Colaborador,
            Cliente: ordem.Cliente,
            Endereço: ordem.Endereço,
            Fechamento: ordem.Fechamento ? numeroSerieParaData(ordem.Fechamento) : 'Data de fechamento não consta'
        })));

        XLSX.utils.book_append_sheet(wb, ws, assunto + "_Detalhes");

        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });

        function s2ab(s) {
            const buf = new ArrayBuffer(s.length);
            const view = new Uint8Array(buf);
            for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
            return buf;
        }

        const blob = new Blob([s2ab(wbout)], { type: "application/octet-stream" });

        const link = document.createElement("a");
        link.href = window.URL.createObjectURL(blob);
        link.download = assunto + "_Detalhes.xlsx";
        link.click();
    }
});
