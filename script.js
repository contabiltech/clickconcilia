let dadosExtrato = [];
let dadosRazao = [];
let dadosFornecedoresProcessados = [];
let filtroFornecedoresAtual = 'todos';

function mudarAba(modulo) {
    document.querySelectorAll('.nav-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.modulo').forEach(mod => mod.classList.add('hidden'));

    if (modulo === 'bancaria') {
        event.target.classList.add('active');
        document.getElementById('mod-bancaria').classList.remove('hidden');
    } else {
        event.target.classList.add('active');
        document.getElementById('mod-fornecedores').classList.remove('hidden');
    }
}

// Helper para formatar qualquer entrada de data no padrão dd/mm/aaaa
function formatarDataBR(dataEntrada) {
    if (!dataEntrada || dataEntrada === 'N/A') return 'N/A';

    let str = String(dataEntrada).trim();

    // Tratamento para datas no formato AAAA-MM-DD (ISO)
    if (str.includes('-')) {
        const partes = str.split('T')[0].split('-');
        if (partes.length === 3) {
            const [ano, mes, dia] = partes;
            if (ano.length === 4) {
                return `${dia.padStart(2, '0')}/${mes.padStart(2, '0')}/${ano}`;
            }
        }
    }

    // Tratamento para datas com barra
    if (str.includes('/')) {
        const partes = str.split('/');
        if (partes.length === 3) {
            let [dia, mes, ano] = partes;
            if (ano.length === 2) ano = `20${ano}`;
            return `${dia.padStart(2, '0')}/${mes.padStart(2, '0')}/${ano}`;
        }
    }

    // Tenta parsing via objeto Date do JS se for número serial/outro formato
    const dt = new Date(dataEntrada);
    if (!isNaN(dt.getTime())) {
        const dia = String(dt.getDate()).padStart(2, '0');
        const mes = String(dt.getMonth() + 1).padStart(2, '0');
        const ano = dt.getFullYear();
        return `${dia}/${mes}/${ano}`;
    }

    return str;
}

function lerArquivoPlanilha(file, origem) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // ATUALIZADO: Utiliza a função flexível para preencher células mescladas e limpar quebras
                const json = window.lerPlanilhaFlexivel(workbook);
                resolve(normalizarDadosPlanilha(json, origem));
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = (error) => reject(error);
        reader.readAsArrayBuffer(file);
    });
}

function normalizarDadosPlanilha(linhas, origem) {
    let listaNormalizada = [];
    if (!linhas || linhas.length === 0) return listaNormalizada;

    let idxData = 0, idxDesc = 1, idxValor = 2, idxConta = -1;

    // Se o retorno veio como objetos (formato gerado pelo sheet_to_json limpo), extraímos as chaves da primeira linha
    if (typeof linhas[0] === 'object' && !Array.isArray(linhas[0])) {
        const chaves = Object.keys(linhas[0]);
        
        chaves.forEach((chave, idx) => {
            const cLower = chave.toLowerCase();
            if (cLower.includes('data') || cLower.includes('date')) idxData = idx;
            if (cLower.includes('desc') || cLower.includes('historico') || cLower.includes('histó')) idxDesc = idx;
            if (cLower.includes('valor') || cLower.includes('monto') || cLower.includes('quant')) idxValor = idx;
            if (cLower.includes('conta') || cLower.includes('cto') || cLower.includes('codigo') || cLower.includes('cód')) idxConta = idx;
        });

        // Converte o array de objetos de volta para matriz de valores alinhados com as colunas detectadas
        linhas = linhas.map(obj => chaves.map(k => obj[k]));
    } else {
        // Fallback para detecção por cabeçalho em matriz padrão
        for (let i = 0; i < Math.min(5, linhas.length); i++) {
            const linhaHeader = linhas[i].map(c => String(c || '').toLowerCase().trim());
            const dataFind = linhaHeader.findIndex(c => c.includes('data') || c.includes('date'));
            const descFind = linhaHeader.findIndex(c => c.includes('desc') || c.includes('historico') || c.includes('histó'));
            const valorFind = linhaHeader.findIndex(c => c.includes('valor') || c.includes('monto') || c.includes('quant'));
            const contaFind = linhaHeader.findIndex(c => c.includes('conta') || c.includes('cto') || c.includes('codigo') || c.includes('cód'));

            if (dataFind !== -1) idxData = dataFind;
            if (descFind !== -1) idxDesc = descFind;
            if (valorFind !== -1) idxValor = valorFind;
            if (contaFind !== -1) idxConta = contaFind;
        }
    }

    for (let i = 0; i < linhas.length; i++) {
        let linha = linhas[i];
        if (!linha || linha.length === 0) continue;

        const strData = String(linha[idxData] || '').trim();
        const strColHeader = strData.toLowerCase();
        if (strColHeader.includes('data') || strColHeader.includes('date')) continue;

        let dataFormatada = formatarDataBR(strData);
        let descricao = String(linha[idxDesc] || 'Sem histórico').trim();
        let contaContabil = idxConta !== -1 && linha[idxConta] ? String(linha[idxConta]).trim() : 'N/A';
        
        let valorBruto = linha[idxValor];
        
        // ATUALIZADO: Usa o conversor universal para lidar com sufixos 'D' e 'C' e sinais monetários
        let valor = window.parseValorUniversal(valorBruto);

        if (!isNaN(valor) && valor !== 0) {
            listaNormalizada.push({
                id: `${origem}_${i}`,
                data: dataFormatada,
                descricao: descricao || 'Sem Histórico',
                contaContabil: contaContabil,
                valor: Math.round(valor * 100) / 100,
                conciliado: false
            });
        }
    }
    return listaNormalizada;
}

async function processarConciliacaoBancaria() {
    const extratoInput = document.getElementById('extratoFile').files[0];
    const razaoInput = document.getElementById('razaoFile').files[0];

    if (!extratoInput || !razaoInput) {
        alert("Por favor, selecione os dois arquivos para comparar!");
        return;
    }

    const spinner = document.getElementById('loadingSpinner');
    const resultadoContainer = document.getElementById('resultadoBancaria');

    spinner.classList.remove('hidden');
    resultadoContainer.classList.add('hidden');

    try {
        await new Promise(resolve => setTimeout(resolve, 50));

        dadosExtrato = await lerArquivoPlanilha(extratoInput, 'EXT');
        dadosRazao = await lerArquivoPlanilha(razaoInput, 'RAZ');

        let conciliadosCount = 0;

        dadosExtrato.forEach(itemExtrato => {
            const matchIndex = dadosRazao.findIndex(itemRazao => 
                !itemRazao.conciliado && Math.abs(itemRazao.valor - itemExtrato.valor) < 0.001
            );

            if (matchIndex !== -1) {
                itemExtrato.conciliado = true;
                dadosRazao[matchIndex].conciliado = true;
                conciliadosCount++;
            }
        });

        const faltamNoRazao = dadosExtrato
            .filter(item => !item.conciliado)
            .map(item => ({ ...item, status: 'Não localizado no razão' }));

        const sobramNoRazao = dadosRazao
            .filter(item => !item.conciliado)
            .map(item => ({ ...item, status: 'Não localizado no extrato' }));

        const todasDivergencias = [...faltamNoRazao, ...sobramNoRazao];

        document.getElementById('qtdTotalExtrato').innerText = dadosExtrato.length;
        document.getElementById('qtdTotalRazao').innerText = dadosRazao.length;
        document.getElementById('qtdPendentes').innerText = todasDivergencias.length;

        renderizarTabelaUnica('tblExtrato', todasDivergencias);

        spinner.classList.add('hidden');
        resultadoContainer.classList.remove('hidden');

    } catch (erro) {
        spinner.classList.add('hidden');
        alert("Erro no processamento dos arquivos. Verifique o console para mais detalhes.");
        console.error("Detalhes do erro:", erro);
    }
}

function renderizarTabelaUnica(idTabela, lista) {
    const tabela = document.getElementById(idTabela);
    if (!tabela) return;

    const tbody = tabela.querySelector('tbody');
    tbody.innerHTML = '';

    if (lista.length === 0) {
        tbody.innerHTML = `<tr><td colspan="4" style="text-align: center; color: #16a34a; font-weight: bold; padding: 1rem;">✅ Perfeito! Todos os lançamentos foram conciliados.</td></tr>`;
        return;
    }

    lista.forEach(item => {
        const tr = document.createElement('tr');
        const valorFormatado = item.valor.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
        
        tr.innerHTML = `
            <td>${formatarDataBR(item.data)}</td>
            <td>${item.descricao}</td>
            <td style="color: ${item.valor < 0 ? '#dc2626' : '#16a34a'}; font-weight: bold;">${valorFormatado}</td>
            <td><span style="background: #fef2f2; color: #dc2626; border: 1px solid #fecaca; padding: 3px 8px; border-radius: 4px; font-size: 0.8em; font-weight: bold;">❌ ${item.status}</span></td>
        `;
        tbody.appendChild(tr);
    });
}

// =========================================================================
// MÓDULO DE FORNECEDORES
// =========================================================================

function extrairNumeroNF(texto) {
    const match = texto.match(/(?:nf|nfe|nota|nota\s*fiscal)\s*[-:]?\s*(\d+)/i);
    return match ? match[1] : null;
}

async function processarFornecedores() {
    const fornecedorInput = document.getElementById('fornecedorFile').files[0];

    if (!fornecedorInput) {
        alert("Por favor, selecione o arquivo do razão de fornecedores!");
        return;
    }

    const spinner = document.getElementById('loadingSpinnerFornecedores');
    const resultadoContainer = document.getElementById('resultadoFornecedores');

    if (spinner) spinner.classList.remove('hidden');
    if (resultadoContainer) resultadoContainer.classList.add('hidden');

    try {
        await new Promise(resolve => setTimeout(resolve, 50));

        const lancamentos = await lerArquivoPlanilha(fornecedorInput, 'FORN');

        if (lancamentos.length === 0) {
            alert("Nenhum lançamento válido encontrado no arquivo.");
            if (spinner) spinner.classList.add('hidden');
            return;
        }

        const notasFiscais = {};

        lancamentos.forEach(item => {
            const numNF = extrairNumeroNF(item.descricao);
            const chave = numNF ? `NF ${numNF}` : 'Sem NF Identificada';
            const ehPagamento = item.descricao.toLowerCase().includes('pagamento') || item.valor < 0;

            if (!notasFiscais[chave]) {
                notasFiscais[chave] = {
                    data: formatarDataBR(item.data),
                    numeroNF: numNF ? `NF ${numNF}` : 'N/A',
                    descricao: item.descricao,
                    contaContabil: item.contaContabil || 'N/A',
                    compras: 0,
                    pagamentos: 0,
                    saldo: 0,
                    listaPagamentos: []
                };
            }

            const valorAbsoluto = Math.abs(item.valor);

            if (ehPagamento) {
                notasFiscais[chave].pagamentos += valorAbsoluto;
                notasFiscais[chave].listaPagamentos.push({
                    data: formatarDataBR(item.data),
                    descricao: item.descricao,
                    valor: valorAbsoluto
                });
            } else {
                notasFiscais[chave].compras += valorAbsoluto;
            }

            notasFiscais[chave].saldo = Math.round((notasFiscais[chave].compras - notasFiscais[chave].pagamentos) * 100) / 100;
        });

        dadosFornecedoresProcessados = Object.values(notasFiscais);
        const emAberto = dadosFornecedoresProcessados.filter(n => Math.abs(n.saldo) > 0.01);
        const quitadas = dadosFornecedoresProcessados.filter(n => Math.abs(n.saldo) <= 0.01);

        document.getElementById('qtdTotalFornecedores').innerText = lancamentos.length;
        document.getElementById('qtdComSaldo').innerText = emAberto.length;
        document.getElementById('qtdQuitados').innerText = quitadas.length;

        filtrarTabelaFornecedores('todos');

        if (spinner) spinner.classList.add('hidden');
        if (resultadoContainer) resultadoContainer.classList.remove('hidden');

    } catch (erro) {
        if (spinner) spinner.classList.add('hidden');
        alert("Erro ao processar o arquivo de fornecedores.");
        console.error("Detalhes do erro:", erro);
    }
}

function filtrarTabelaFornecedores(tipoFiltro) {
    filtroFornecedoresAtual = tipoFiltro;
    let listaFiltrada = [];
    const tituloTabela = document.getElementById('tituloTabelaFornecedores');

    if (tipoFiltro === 'aberto') {
        listaFiltrada = dadosFornecedoresProcessados.filter(n => Math.abs(n.saldo) > 0.01);
        if (tituloTabela) tituloTabela.innerText = '📋 Títulos em Aberto';
    } else if (tipoFiltro === 'quitados') {
        listaFiltrada = dadosFornecedoresProcessados.filter(n => Math.abs(n.saldo) <= 0.01);
        if (tituloTabela) tituloTabela.innerText = '📋 Títulos Quitados';
    } else {
        listaFiltrada = [...dadosFornecedoresProcessados];
        if (tituloTabela) tituloTabela.innerText = '📋 Resumo de Títulos e Saldos por Fornecedor (Todos)';
    }

    renderizarTabelaFornecedores('tblFornecedores', listaFiltrada);
}

function renderizarTabelaFornecedores(idTabela, lista) {
    const tabela = document.getElementById(idTabela);
    if (!tabela) return;

    const tbody = tabela.querySelector('tbody');
    tbody.innerHTML = '';

    if (lista.length === 0) {
        tbody.innerHTML = `<tr><td colspan="9" style="text-align: center; padding: 1rem;">Nenhum registro encontrado para este filtro.</td></tr>`;
        return;
    }

    lista.forEach((item, index) => {
        const formatBRL = (v) => v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
        const temSaldoAberto = Math.abs(item.saldo) > 0.01;
        
        const temMultiplosPagamentos = item.listaPagamentos.length > 1;

        const badgeStatus = temSaldoAberto 
            ? `<span style="background: #fef2f2; color: #dc2626; border: 1px solid #fecaca; padding: 4px 8px; border-radius: 4px; font-size: 0.8em; font-weight: bold;">⚠️ Saldo em Aberto</span>`
            : `<span style="background: #f0fdf4; color: #15803d; border: 1px solid #bbf7d0; padding: 4px 8px; border-radius: 4px; font-size: 0.8em; font-weight: bold;">✅ Quitado</span>`;

        const btnExpandir = temMultiplosPagamentos 
            ? `<button onclick="alternarAgrupamento(${index})" id="btn-toggle-${index}" title="Ver múltiplos pagamentos (${item.listaPagamentos.length})" style="background: #e2e8f0; border: 1px solid #cbd5e1; border-radius: 4px; width: 24px; height: 24px; cursor: pointer; font-weight: bold; line-height: 1;">+</button>`
            : `<span style="color: #cbd5e1;">-</span>`;

        const trPrincipal = document.createElement('tr');
        if (temSaldoAberto) trPrincipal.style.backgroundColor = '#fffbe2';

        trPrincipal.innerHTML = `
            <td style="text-align: center;">${btnExpandir}</td>
            <td>${formatarDataBR(item.data)}</td>
            <td>${item.contaContabil}</td>
            <td>${item.descricao}</td>
            <td><strong>${item.numeroNF}</strong></td>
            <td>${formatBRL(item.pagamentos)}</td>
            <td>${formatBRL(item.compras)}</td>
            <td style="font-weight: bold; color: ${temSaldoAberto ? '#dc2626' : '#16a34a'};">${formatBRL(item.saldo)}</td>
            <td>${badgeStatus}</td>
        `;
        tbody.appendChild(trPrincipal);

        if (temMultiplosPagamentos) {
            const trDetalhes = document.createElement('tr');
            trDetalhes.id = `detalhes-${index}`;
            trDetalhes.className = 'hidden';
            trDetalhes.style.backgroundColor = '#f8fafc';

            let tabelaInternaHTML = `
                <td colspan="9" style="padding: 10px 20px 10px 50px; border-top: 1px dashed #cbd5e1; border-bottom: 1px solid #cbd5e1;">
                    <strong style="color: #475569; font-size: 0.9em;">↳ Desmembramento dos ${item.listaPagamentos.length} Pagamentos Vinculados:</strong>
                    <table style="width: 100%; margin-top: 5px; font-size: 0.85em; background: white; border: 1px solid #e2e8f0; border-radius: 4px;">
                        <thead>
                            <tr style="background: #f1f5f9; text-align: left;">
                                <th style="padding: 6px;">Data do Pagamento</th>
                                <th style="padding: 6px;">Descrição do Lançamento</th>
                                <th style="padding: 6px; text-align: right;">Valor Pago</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            item.listaPagamentos.forEach(pag => {
                tabelaInternaHTML += `
                    <tr style="border-bottom: 1px solid #f1f5f9;">
                        <td style="padding: 6px;">${formatarDataBR(pag.data)}</td>
                        <td style="padding: 6px;">${pag.descricao}</td>
                        <td style="padding: 6px; text-align: right; color: #16a34a; font-weight: bold;">${formatBRL(pag.valor)}</td>
                    </tr>
                `;
            });

            tabelaInternaHTML += `
                        </tbody>
                    </table>
                </td>
            `;

            trDetalhes.innerHTML = tabelaInternaHTML;
            tbody.appendChild(trDetalhes);
        }
    });
}

function alternarAgrupamento(index) {
    const trDetalhes = document.getElementById(`detalhes-${index}`);
    const btn = document.getElementById(`btn-toggle-${index}`);

    if (trDetalhes && btn) {
        if (trDetalhes.classList.contains('hidden')) {
            trDetalhes.classList.remove('hidden');
            btn.innerText = '-';
            btn.style.background = '#cbd5e1';
        } else {
            trDetalhes.classList.add('hidden');
            btn.innerText = '+';
            btn.style.background = '#e2e8f0';
        }
    }
}

function exportarRelatorioFornecedoresXLSX() {
    let listaFiltrada = [];

    if (filtroFornecedoresAtual === 'aberto') {
        listaFiltrada = dadosFornecedoresProcessados.filter(n => Math.abs(n.saldo) > 0.01);
    } else if (filtroFornecedoresAtual === 'quitados') {
        listaFiltrada = dadosFornecedoresProcessados.filter(n => Math.abs(n.saldo) <= 0.01);
    } else {
        listaFiltrada = [...dadosFornecedoresProcessados];
    }

    if (listaFiltrada.length === 0) {
        alert("Não há dados para exportar no filtro atual.");
        return;
    }

    const dadosExportacao = listaFiltrada.map(item => ({
        "Data": formatarDataBR(item.data),
        "Conta Contábil": item.contaContabil,
        "Descrição do Lançamento": item.descricao,
        "Nº da NF": item.numeroNF,
        "Total Pagamentos (Débitos)": item.pagamentos,
        "Total Compras (Créditos)": item.compras,
        "Saldo Remanescente": item.saldo,
        "Status": Math.abs(item.saldo) > 0.01 ? "Saldo em Aberto" : "Quitado"
    }));

    const worksheet = XLSX.utils.json_to_sheet(dadosExportacao);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Conciliacao_Fornecedores");

    const nomeArquivo = `Relatorio_Fornecedores_${filtroFornecedoresAtual}.xlsx`;
    XLSX.writeFile(workbook, nomeArquivo);
}
