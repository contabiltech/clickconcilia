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

// Helper para formatar qualquer entrada de data no padrão dd/mm/aaaa sem problemas de timezone
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

    return str;
}

// Parser universal para valores monetários com suporte a sufixos D/C (Débito/Crédito)
function parseValorUniversal(valorBruto) {
    if (typeof valorBruto === 'number') return valorBruto;
    if (!valorBruto) return 0.0;

    let str = String(valorBruto).trim().toUpperCase();
    let multiplicador = 1;

    if (str.endsWith('D')) {
        multiplicador = 1; // Débito padrão extrato/razão
        str = str.slice(0, -1).trim();
    } else if (str.endsWith('C')) {
        multiplicador = -1; // Crédito
        str = str.slice(0, -1).trim();
    }

    // Remove símbolos de moeda e trata pontos de milhar e vírgula decimal
    str = str.replace(/[R$\s]/g, '');
    if (str.includes(',') && str.includes('.')) {
        str = str.replace(/\./g, '').replace(',', '.');
    } else if (str.includes(',')) {
        str = str.replace(',', '.');
    }

    let num = parseFloat(str);
    return isNaN(num) ? 0.0 : num * multiplicador;
}

// Leitura de planilha com suporte a XLSX / XLS
function lerArquivoPlanilha(file, origem) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                resolve(normalizarDadosPlanilha(json, origem));
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = (error) => reject(error);
        reader.readAsArrayBuffer(file);
    });
}

// Função auxiliar para leitura de PDF com OCR automático forçado
async function lerArquivoPDFComOCR(file, origem) {
    return new Promise(async (resolve, reject) => {
        try {
            const reader = new FileReader();
            reader.onload = async function() {
                try {
                    const typedarray = new Uint8Array(this.result);
                    const pdf = await pdfjsLib.getDocument({ data: typedarray }).promise;
                    let textoCompleto = "";
                    
                    for (let i = 1; i <= pdf.numPages; i++) {
                        const page = await pdf.getPage(i);
                        const textContent = await page.getTextContent();
                        let textoPagina = textContent.items.map(item => item.str).join(" ");
                        
                        if (textoPagina.trim().length < 20 && typeof Tesseract !== 'undefined') {
                            const viewport = page.getViewport({ scale: 1.5 });
                            const canvas = document.createElement('canvas');
                            const context = canvas.getContext('2d');
                            canvas.height = viewport.height;
                            canvas.width = viewport.width;
                            await page.render({ canvasContext: context, viewport: viewport }).promise;
                            
                            const resultadoOCR = await Tesseract.recognize(canvas, 'por');
                            textoPagina = resultadoOCR.data.text;
                        }
                        textoCompleto += textoPagina + "\n";
                    }
                    
                    resolve(normalizarTextoRazaoOuExtrato(textoCompleto, origem));
                } catch (e) {
                    reject(e);
                }
            };
            reader.readAsArrayBuffer(file);
        } catch (err) {
            reject(err);
        }
    });
}

function normalizarTextoRazaoOuExtrato(texto, origem) {
    let linhasNormalizadas = [];
    const linhas = texto.split('\n');
    
    linhas.forEach((linha, index) => {
        let linhaStr = linha.trim();
        const regexData = /\d{2}\/\d{2}\/\d{4}/;
        if (regexData.test(linhaStr)) {
            linhasNormalizadas.push({
                id: `${origem}_${index}`,
                data: formatarDataBR(linhaStr.match(regexData)[0]),
                descricao: linhaStr,
                contaContabil: 'N/A',
                valor: parseValorUniversal(linhaStr),
                conciliado: false
            });
        }
    });
    return linhasNormalizadas;
}

function normalizarDadosPlanilha(linhas, origem) {
    let listaNormalizada = [];
    if (!linhas || linhas.length === 0) return listaNormalizada;

    let idxData = 0, idxDesc = 1, idxValor = 2, idxConta = -1;

    for (let i = 0; i < Math.min(5, linhas.length); i++) {
        const linhaHeader = (linhas[i] || []).map(c => String(c || '').toLowerCase().trim());
        const dataFind = linhaHeader.findIndex(c => c.includes('data') || c.includes('date'));
        const descFind = linhaHeader.findIndex(c => c.includes('desc') || c.includes('historico') || c.includes('histó'));
        const valorFind = linhaHeader.findIndex(c => c.includes('valor') || c.includes('débito') || c.includes('crédito') || c.includes('monto'));
        const contaFind = linhaHeader.findIndex(c => c.includes('conta') || c.includes('cto') || c.includes('codigo') || c.includes('cód'));

        if (dataFind !== -1) idxData = dataFind;
        if (descFind !== -1) idxDesc = descFind;
        if (valorFind !== -1) idxValor = valorFind;
        if (contaFind !== -1) idxConta = contaFind;
    }

    for (let i = 0; i < linhas.length; i++) {
        let linha = linhas[i];
        if (!linha || linha.length === 0) continue;

        const strData = String(linha[idxData] || '').trim();
        if (strData.toLowerCase().includes('data') || strData.toLowerCase().includes('date')) continue;

        let dataFormatada = formatarDataBR(strData);
        let descricao = String(linha[idxDesc] || 'Sem histórico').trim();
        let contaContabil = idxConta !== -1 && linha[idxConta] ? String(linha[idxConta]).trim() : 'N/A';
        
        let valor = 0;
        if (linha.length > 4) {
            let debito = parseValorUniversal(linha[3]);
            let credito = parseValorUniversal(linha[4]);
            valor = debito !== 0 ? debito : -Math.abs(credito);
        } else {
            valor = parseValorUniversal(linha[idxValor]);
        }

        if (!isNaN(valor) && valor !== 0 && dataFormatada !== 'N/A') {
            listaNormalizada.push({
                id: `${origem}_${i}`,
                data: dataFormatada,
                descricao: descricao,
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
        const lerArquivo = async (file, origem) => {
            if (file.type === 'application/pdf' || file.name.endsWith('.pdf')) {
                return await lerArquivoPDFComOCR(file, origem);
            } else {
                return await lerArquivoPlanilha(file, origem);
            }
        };

        dadosExtrato = await lerArquivo(extratoInput, 'EXT');
        dadosRazao = await lerArquivo(razaoInput, 'RAZ');

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

        const faltamNoRazao = dadosExtrato.filter(item => !item.conciliado).map(item => ({ ...item, status: 'Não localizado no razão' }));
        const sobramNoRazao = dadosRazao.filter(item => !item.conciliado).map(item => ({ ...item, status: 'Não localizado no extrato' }));
        const todasDivergencias = [...faltamNoRazao, ...sobramNoRazao];

        document.getElementById('qtdTotalExtrato').innerText = dadosExtrato.length;
        document.getElementById('qtdTotalRazao').innerText = dadosRazao.length;
        document.getElementById('qtdPendentes').innerText = todasDivergencias.length;

        renderizarTabelaUnica('tblExtrato', todasDivergencias);

        spinner.classList.add('hidden');
        resultadoContainer.classList.remove('hidden');

    } catch (erro) {
        spinner.classList.add('hidden');
        alert("Erro no processamento dos arquivos. Verifique o console.");
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
            <td>${item.data}</td>
            <td>${item.descricao}</td>
            <td style="color: ${item.valor < 0 ? '#dc2626' : '#16a34a'}; font-weight: bold;">${valorFormatado}</td>
            <td><span style="background: #fef2f2; color: #dc2626; border: 1px solid #fecaca; padding: 3px 8px; border-radius: 4px; font-size: 0.8em; font-weight: bold;">❌ ${item.status}</span></td>
        `;
        tbody.appendChild(tr);
    });
}

// =========================================================================
// MÓDULO DE FORNECEDORES (ADICIONADO)
// =========================================================================

window.processarFornecedores = async function() {
    const fileInput = document.getElementById('fornecedorFile');
    if (!fileInput.files || fileInput.files.length === 0) {
        alert("Por favor, selecione o arquivo de Razão de Fornecedores.");
        return;
    }

    const spinner = document.getElementById('loadingSpinnerFornecedores');
    const resultado = document.getElementById('resultadoFornecedores');

    spinner.classList.remove('hidden');
    resultado.classList.add('hidden');

    try {
        const file = fileInput.files[0];
        let dadosBrutos = [];

        if (file.type === 'application/pdf' || file.name.endsWith('.pdf')) {
            dadosBrutos = await lerArquivoPDFComOCR(file, 'FORN');
        } else {
            dadosBrutos = await lerArquivoPlanilha(file, 'FORN');
        }

        // Processamento e agrupamento por NF / Fornecedor
        dadosFornecedoresProcessados = dadosBrutos.map((item, index) => {
            // Tenta extrair número de NF da descrição se houver (ex: "NF 12345" ou "NFE 12345")
            let matchNF = item.descricao.match(/(?:NF|NFC-e|NFE|NF-e)[\s:]*(\d+)/i);
            let numNF = matchNF ? matchNF[1] : `NF-${index + 1}`;
            
            let valor = item.valor;
            let totalCompras = valor > 0 ? valor : 0;
            let totalPagamentos = valor < 0 ? Math.abs(valor) : 0;
            let saldo = totalCompras - totalPagamentos;
            let quitado = Math.abs(saldo) < 0.01;

            return {
                ...item,
                numNF: numNF,
                totalCompras: totalCompras,
                totalPagamentos: totalPagamentos,
                saldo: saldo,
                quitado: quitado,
                status: quitado ? 'Quitado' : 'Em Aberto'
            };
        });

        atualizarResumoEExibirFornecedores();

        spinner.classList.add('hidden');
        resultado.classList.remove('hidden');

    } catch (erro) {
        spinner.classList.add('hidden');
        alert("Erro ao processar o arquivo de fornecedores. Verifique o console.");
        console.error("Erro fornecedores:", erro);
    }
};

function atualizarResumoEExibirFornecedores() {
    const totalItens = dadosFornecedoresProcessados.length;
    const emAberto = dadosFornecedoresProcessados.filter(i => !i.quitado).length;
    const quitados = dadosFornecedoresProcessados.filter(i => i.quitado).length;

    document.getElementById('qtdTotalFornecedores').innerText = totalItens;
    document.getElementById('qtdComSaldo').innerText = emAberto;
    document.getElementById('qtdQuitados').innerText = quitados;

    filtrarTabelaFornecedores(filtroFornecedoresAtual);
}

window.filtrarTabelaFornecedores = function(filtro) {
    filtroFornecedoresAtual = filtro;
    
    // Atualiza classes visuais dos cards
    document.getElementById('cardFiltroTodos').style.border = filtro === 'todos' ? '2px solid #2563eb' : 'none';
    document.getElementById('cardFiltroAberto').style.border = filtro === 'aberto' ? '2px solid #d97706' : 'none';
    document.getElementById('cardFiltroQuitados').style.border = filtro === 'quitados' ? '2px solid #16a34a' : 'none';

    let dadosFiltrados = dadosFornecedoresProcessados;
    let tituloFiltroTexto = "Todos";

    if (filtro === 'aberto') {
        dadosFiltrados = dadosFornecedoresProcessados.filter(i => !i.quitado);
        tituloFiltroTexto = "Títulos em Aberto";
    } else if (filtro === 'quitados') {
        dadosFiltrados = dadosFornecedoresProcessados.filter(i => i.quitado);
        tituloFiltroTexto = "Títulos Quitados";
    }

    document.getElementById('tituloTabelaFornecedores').innerText = `📋 Resumo de Títulos e Saldos por Fornecedor (${tituloFiltroTexto})`;

    const tabela = document.getElementById('tblFornecedores');
    const tbody = tabela.querySelector('tbody');
    tbody.innerHTML = '';

    if (dadosFiltrados.length === 0) {
        tbody.innerHTML = `<tr><td colspan="9" style="text-align: center; color: #64748b; font-weight: bold; padding: 1rem;">Nenhum registro encontrado para este filtro.</td></tr>`;
        return;
    }

    dadosFiltrados.forEach((item, index) => {
        const tr = document.createElement('tr');
        const corStatus = item.quitado ? '#16a34a' : '#d97706';
        const bgStatus = item.quitado ? '#f0fdf4' : '#fffbeb';
        const borderStatus = item.quitado ? '#bbf7d0' : '#fef3c7';

        tr.innerHTML = `
            <td style="text-align: center;">${index + 1}</td>
            <td style="text-align: center;">${item.data}</td>
            <td style="text-align: center;">${item.contaContabil}</td>
            <td>${item.descricao}</td>
            <td style="text-align: center; font-weight: bold;">${item.numNF}</td>
            <td style="text-align: right;">${item.totalPagamentos.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</td>
            <td style="text-align: right;">${item.totalCompras.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</td>
            <td style="text-align: right; font-weight: bold;">${item.saldo.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</td>
            <td style="text-align: center;"><span class="status-badge" style="background: ${bgStatus}; color: ${corStatus}; border: 1px solid ${borderStatus}; padding: 3px 8px; border-radius: 4px; font-size: 0.8em; font-weight: bold;">${item.status}</span></td>
        `;
        tbody.appendChild(tr);
    });
};

window.exportarRelatorioFornecedoresXLSX = function() {
    if (!dadosFornecedoresProcessados || dadosFornecedoresProcessados.length === 0) {
        alert("Não há dados processados para exportar.");
        return;
    }

    let dadosFiltrados = dadosFornecedoresProcessados;
    if (filtroFornecedoresAtual === 'aberto') {
        dadosFiltrados = dadosFornecedoresProcessados.filter(i => !i.quitado);
    } else if (filtroFornecedoresAtual === 'quitados') {
        dadosFiltrados = dadosFornecedoresProcessados.filter(i => i.quitado);
    }

    const dadosExportacao = dadosFiltrados.map((item, idx) => ({
        "Item": idx + 1,
        "Data": item.data,
        "Conta Contábil": item.contaContabil,
        "Descrição": item.descricao,
        "Nº da NF": item.numNF,
        "Total Pagamentos": item.totalPagamentos,
        "Total Compras": item.totalCompras,
        "Saldo": item.saldo,
        "Status": item.status
    }));

    const worksheet = XLSX.utils.json_to_sheet(dadosExportacao);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Relatorio_Fornecedores");
    XLSX.writeFile(workbook, `Relatorio_Fornecedores_${filtroFornecedoresAtual}.xlsx`);
};
