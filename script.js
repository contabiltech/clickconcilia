let dadosExtrato = [];
let dadosRazao = [];
let dadosFornecedoresProcessados = [];
let filtroFornecedoresAtual = 'todos';

pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';

// --- NAVEGAÇÃO ---
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

// --- TRATAMENTO DE DATAS ---
function formatarDataBR(dataEntrada) {
    if (dataEntrada == null || dataEntrada === '') return null;

    // Se for número serial do Excel (ex: 44348)
    if (typeof dataEntrada === 'number') {
        const dataData = new Date((dataEntrada - 25569) * 86400 * 1000);
        const dia = String(dataData.getUTCDate()).padStart(2, '0');
        const mes = String(dataData.getUTCMonth() + 1).padStart(2, '0');
        const ano = dataData.getUTCFullYear();
        return `${dia}/${mes}/${ano}`;
    }

    let str = String(dataEntrada).trim();
    if (str.includes('-')) {
        const partes = str.split('T')[0].split('-');
        if (partes.length === 3 && partes[0].length === 4) return `${partes[2]}/${partes[1]}/${partes[0]}`;
    }
    if (str.includes('/')) {
        const partes = str.split('/');
        if (partes.length === 3) {
            let ano = partes[2].length === 2 ? `20${partes[2]}` : partes[2];
            return `${partes[0].padStart(2, '0')}/${partes[1].padStart(2, '0')}/${ano}`;
        }
    }
    return str;
}

// --- LEITURA DE ARQUIVOS (EXCEL, .XLS HTML DO DOMÍNIO E PDF) ---
async function processarArquivoBruto(file, origem) {
    const ext = file.name.split('.').pop().toLowerCase();
    
    if (ext === 'pdf') {
        return await lerArquivoPDF(file, origem);
    } else {
        return await lerArquivoExcel(file, origem);
    }
}

function lerArquivoExcel(file, origem) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const buffer = e.target.result;
                let workbook;
                
                // Tenta ler como binário padrão do Excel
                try {
                    const data = new Uint8Array(buffer);
                    workbook = XLSX.read(data, { type: 'array', cellDates: true });
                } catch (exBinario) {
                    // Fallback inteligente: se falhar, o .xls do Domínio costuma ser HTML/Texto
                    const textDecoder = new TextDecoder('windows-1252');
                    const textContent = textDecoder.decode(buffer);
                    workbook = XLSX.read(textContent, { type: 'string', cellDates: true });
                }

                const primeiraAba = workbook.SheetNames[0];
                const aba = workbook.Sheets[primeiraAba];
                
                const linhas = XLSX.utils.sheet_to_json(aba, { header: 1, raw: true, defval: null });
                resolve(normalizarDadosInteligente(linhas, origem));
            } catch (err) { 
                reject(err); 
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

async function lerArquivoPDF(file, origem) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    let linhasTexto = [];
    
    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        
        let rows = {};
        textContent.items.forEach(item => {
            let y = Math.round(item.transform[5]); 
            if (!rows[y]) rows[y] = [];
            rows[y].push({ text: item.str, x: Math.round(item.transform[4]) });
        });

        const yKeys = Object.keys(rows).map(Number).sort((a, b) => b - a);
        yKeys.forEach(y => {
            rows[y].sort((a, b) => a.x - b.x);
            let rowValues = rows[y].map(obj => obj.text.trim()).filter(t => t !== '');
            if (rowValues.length > 0) linhasTexto.push(rowValues);
        });
    }
    return normalizarDadosInteligente(linhasTexto, origem);
}

// --- ALGORITMO INTELIGENTE DE NORMALIZAÇÃO ---
function normalizarDadosInteligente(linhas, origem) {
    let listaNormalizada = [];
    let currentConta = "N/A";
    
    let idxData = -1, idxDesc = -1, idxDebito = -1, idxCredito = -1;

    for (let i = 0; i < Math.min(25, linhas.length); i++) {
        let rowStr = linhas[i].map(c => String(c || '').toLowerCase().trim());
        if (rowStr.includes('débito') || rowStr.includes('debito') || rowStr.includes('crédito') || rowStr.includes('credito') || rowStr.includes('valor')) {
            idxData = rowStr.findIndex(c => c.includes('data'));
            idxDesc = rowStr.findIndex(c => c.includes('históric') || c.includes('historic') || c.includes('descri'));
            idxDebito = rowStr.findIndex(c => c.includes('débito') || c.includes('debito'));
            idxCredito = rowStr.findIndex(c => c.includes('crédito') || c.includes('credito'));
            break;
        }
    }

    if (idxData === -1) idxData = 0; 
    if (idxDesc === -1) idxDesc = 2; 

    for (let i = 0; i < linhas.length; i++) {
        let row = linhas[i];
        if (!row || row.length === 0) continue;

        let rowTextoCompleto = row.map(String).join(" ").toUpperCase();
        
        if (rowTextoCompleto.includes("CONTA:") || String(row[0]).toUpperCase().includes("CONTA:")) {
            let pedacos = row.filter(c => typeof c === 'string' && c.length > 3);
            currentConta = pedacos[pedacos.length - 1] || "Conta Detectada";
            continue;
        }

        let possibleData = row[idxData];
        let dataFormatada = formatarDataBR(possibleData);

        if (!dataFormatada || dataFormatada === 'N/A' || dataFormatada.includes('NaN') || rowTextoCompleto.includes("SALDO") || rowTextoCompleto.includes("TOTAL") || rowTextoCompleto.includes("RAZÃO")) {
            continue;
        }

        let descricao = "";
        for (let c = idxDesc; c < row.length; c++) {
            if (c === idxDebito || c === idxCredito) break;
            if (row[c] && isNaN(row[c]) && typeof row[c] === 'string') {
                descricao += row[c] + " ";
            }
        }
        descricao = descricao.trim() || "Lançamento sem histórico";

        const parseValor = (val) => {
            if (typeof val === 'number') return val;
            if (!val) return 0;
            let str = String(val).replace('R$', '').replace(/\s/g, '').trim();
            if (str.includes(',') && str.includes('.')) str = str.replace(/\./g, '').replace(',', '.');
            else if (str.includes(',')) str = str.replace(',', '.');
            let num = parseFloat(str);
            return isNaN(num) ? 0 : num;
        };

        let debito = 0, credito = 0, valorUnico = 0;

        if (idxDebito !== -1 && idxCredito !== -1) {
            debito = parseValor(row[idxDebito]);
            credito = parseValor(row[idxCredito]);
            valorUnico = credito > 0 ? credito : (debito * -1); 
        } else {
            let numeros = row.filter(c => typeof c === 'number' || (!isNaN(parseValor(c)) && parseValor(c) > 0));
            if (numeros.length > 0) {
                valorUnico = parseValor(numeros[numeros.length - 1]);
                if (rowTextoCompleto.includes('PAG') || rowTextoCompleto.includes('DEB')) valorUnico = -Math.abs(valorUnico);
                if (valorUnico < 0) debito = Math.abs(valorUnico);
                else credito = Math.abs(valorUnico);
            }
        }

        if (debito > 0 || credito > 0 || valorUnico !== 0) {
            listaNormalizada.push({
                id: `${origem}_${i}`,
                data: dataFormatada,
                contaContabil: currentConta,
                descricao: descricao,
                debito: debito,
                credito: credito,
                valor: valorUnico,
                conciliado: false
            });
        }
    }
    return listaNormalizada;
}

// =========================================================================
// MÓDULO: CONCILIAÇÃO BANCÁRIA
// =========================================================================
async function processarConciliacaoBancaria() {
    const extratoInput = document.getElementById('extratoFile').files[0];
    const razaoInput = document.getElementById('razaoFile').files[0];

    if (!extratoInput || !razaoInput) return alert("Selecione os dois arquivos (Extrato e Razão)!");

    document.getElementById('loadingSpinner').classList.remove('hidden');
    document.getElementById('resultadoBancaria').classList.add('hidden');

    try {
        dadosExtrato = await processarArquivoBruto(extratoInput, 'EXT');
        dadosRazao = await processarArquivoBruto(razaoInput, 'RAZ');

        dadosExtrato.forEach(ext => {
            const matchIndex = dadosRazao.findIndex(raz => 
                !raz.conciliado && Math.abs(Math.abs(raz.valor) - Math.abs(ext.valor)) < 0.01
            );
            if (matchIndex !== -1) {
                ext.conciliado = true;
                dadosRazao[matchIndex].conciliado = true;
            }
        });

        const faltamNoRazao = dadosExtrato.filter(item => !item.conciliado).map(i => ({...i, origem: 'Falta no Razão (Sobrou no Extrato)'}));
        const sobramNoRazao = dadosRazao.filter(item => !item.conciliado).map(i => ({...i, origem: 'Falta no Extrato (Sobrou no Razão)'}));
        
        const divergencias = [...faltamNoRazao, ...sobramNoRazao];

        document.getElementById('qtdTotalExtrato').innerText = dadosExtrato.length;
        document.getElementById('qtdTotalRazao').innerText = dadosRazao.length;
        document.getElementById('qtdPendentes').innerText = divergencias.length;

        const tbody = document.querySelector('#tblExtrato tbody');
        tbody.innerHTML = '';
        if (divergencias.length === 0) {
            tbody.innerHTML = `<tr><td colspan="5" style="text-align:center; padding: 2rem;">✅ Conciliação 100% exata. Nenhuma divergência!</td></tr>`;
        } else {
            const formatBRL = (v) => v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
            divergencias.forEach(item => {
                const badgeOrigem = item.origem.includes('Falta no Razão') ? `<span class="badge badge-alerta">No Extrato</span>` : `<span class="badge badge-aberto">No Razão</span>`;
                tbody.innerHTML += `
                    <tr>
                        <td>${badgeOrigem}</td>
                        <td>${item.data}</td>
                        <td>${item.descricao}</td>
                        <td style="font-weight:bold; color: ${item.valor < 0 ? '#dc2626' : '#16a34a'};">${formatBRL(Math.abs(item.valor))}</td>
                        <td><span class="badge badge-aberto">❌ Pendente</span></td>
                    </tr>`;
            });
        }

        document.getElementById('loadingSpinner').classList.add('hidden');
        document.getElementById('resultadoBancaria').classList.remove('hidden');

    } catch (erro) {
        document.getElementById('loadingSpinner').classList.add('hidden');
        alert("Erro ao ler os arquivos. Verifique se não estão corrompidos.");
        console.error(erro);
    }
}

// =========================================================================
// MÓDULO: CONCILIAÇÃO DE FORNECEDORES
// =========================================================================
function extrairNumeroNF(texto) {
    const match = texto.match(/(?:nf|nfe|nota|nota\s*fiscal)\s*[-:#]?\s*(\d+)/i);
    return match ? match[1] : null;
}

async function processarFornecedores() {
    const fornecedorInput = document.getElementById('fornecedorFile').files[0];
    if (!fornecedorInput) return alert("Selecione o arquivo do razão de fornecedores!");

    document.getElementById('loadingSpinnerFornecedores').classList.remove('hidden');
    document.getElementById('resultadoFornecedores').classList.add('hidden');

    try {
        const lancamentos = await processarArquivoBruto(fornecedorInput, 'FORN');
        const notasFiscais = {};

        lancamentos.forEach(item => {
            const numNF = extrairNumeroNF(item.descricao);
            const chave = numNF ? `NF ${numNF}` : `Sem NF - ${item.contaContabil} - ${item.data}`;
            
            const ehPagamento = item.debito > 0 || item.descricao.toLowerCase().includes('pagto') || item.descricao.toLowerCase().includes('pagamento');

            if (!notasFiscais[chave]) {
                notasFiscais[chave] = {
                    data: item.data,
                    numeroNF: numNF ? `NF ${numNF}` : 'Não Identificada',
                    descricao: item.descricao,
                    contaContabil: item.contaContabil,
                    compras: 0,
                    pagamentos: 0,
                    saldo: 0,
                    listaPagamentos: []
                };
            }

            if (ehPagamento) {
                let valorPago = item.debito > 0 ? item.debito : Math.abs(item.valor);
                notasFiscais[chave].pagamentos += valorPago;
                notasFiscais[chave].listaPagamentos.push({ data: item.data, descricao: item.descricao, valor: valorPago });
            } else {
                let valorCompra = item.credito > 0 ? item.credito : Math.abs(item.valor);
                notasFiscais[chave].compras += valorCompra;
            }

            notasFiscais[chave].saldo = Math.round((notasFiscais[chave].compras - notasFiscais[chave].pagamentos) * 100) / 100;
        });

        dadosFornecedoresProcessados = Object.values(notasFiscais);
        filtrarTabelaFornecedores('todos');

        document.getElementById('loadingSpinnerFornecedores').classList.add('hidden');
        document.getElementById('resultadoFornecedores').classList.remove('hidden');

    } catch (erro) {
        document.getElementById('loadingSpinnerFornecedores').classList.add('hidden');
        alert("Erro ao processar o arquivo. Verifique o console.");
        console.error(erro);
    }
}

function filtrarTabelaFornecedores(tipoFiltro) {
    filtroFornecedoresAtual = tipoFiltro;
    let lista = [];

    if (tipoFiltro === 'aberto') {
        lista = dadosFornecedoresProcessados.filter(n => Math.abs(n.saldo) > 0.01);
        document.getElementById('tituloTabelaFornecedores').innerText = '📋 Títulos Apenas Em Aberto';
    } else if (tipoFiltro === 'quitados') {
        lista = dadosFornecedoresProcessados.filter(n => Math.abs(n.saldo) <= 0.01);
        document.getElementById('tituloTabelaFornecedores').innerText = '📋 Títulos Liquidados';
    } else {
        lista = [...dadosFornecedoresProcessados];
        document.getElementById('tituloTabelaFornecedores').innerText = '📋 Resumo de Títulos (Todos)';
    }

    const abertos = dadosFornecedoresProcessados.filter(n => Math.abs(n.saldo) > 0.01);
    const quitados = dadosFornecedoresProcessados.filter(n => Math.abs(n.saldo) <= 0.01);
    const totalAbertoValor = abertos.reduce((acc, curr) => acc + curr.saldo, 0);

    document.getElementById('qtdTotalFornecedores').innerText = dadosFornecedoresProcessados.length;
    document.getElementById('qtdComSaldo').innerText = abertos.length;
    document.getElementById('qtdQuitados').innerText = quitados.length;
    document.getElementById('valorTotalAberto').innerText = totalAbertoValor.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });

    const tbody = document.querySelector('#tblFornecedores tbody');
    tbody.innerHTML = '';

    if(lista.length === 0) {
        tbody.innerHTML = `<tr><td colspan="9" style="text-align: center; padding: 2rem;">Nenhum registro encontrado.</td></tr>`;
        return;
    }

    const formatBRL = (v) => v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });

    lista.forEach((item, index) => {
        const saldoAberto = Math.abs(item.saldo) > 0.01;
        const multiplosPag = item.listaPagamentos.length > 1;

        const badge = saldoAberto ? `<span class="badge badge-aberto">⚠️ EM ABERTO</span>` : `<span class="badge badge-liquidado">✅ LIQUIDADO</span>`;
        const btnExpandir = multiplosPag 
            ? `<button class="btn-expand" onclick="alternarAgrupamento(${index})" id="btn-toggle-${index}" title="Ver ${item.listaPagamentos.length} pagamentos">+</button>` 
            : `-`;

        tbody.innerHTML += `
            <tr style="${saldoAberto ? 'background-color: #fffbeb;' : ''}">
                <td style="text-align:center;">${btnExpandir}</td>
                <td>${item.data}</td>
                <td><strong>${item.contaContabil}</strong></td>
                <td><strong>${item.numeroNF}</strong></td>
                <td>${item.descricao}</td>
                <td style="color: #475569;">${formatBRL(item.compras)}</td>
                <td style="color: #475569;">${formatBRL(item.pagamentos)}</td>
                <td style="font-weight: bold; color: ${saldoAberto ? '#dc2626' : '#15803d'};">${formatBRL(item.saldo)}</td>
                <td>${badge}</td>
            </tr>
        `;

        if (multiplosPag) {
            let pagamentosHTML = item.listaPagamentos.map(p => `
                <tr>
                    <td>${p.data}</td><td>${p.descricao}</td>
                    <td style="text-align:right; color:#16a34a; font-weight:bold;">${formatBRL(p.valor)}</td>
                </tr>
            `).join('');

            tbody.innerHTML += `
                <tr id="detalhes-${index}" class="hidden tr-detalhes">
                    <td colspan="9" style="padding: 10px 40px;">
                        <strong style="color: #475569; font-size: 0.85em;">↳ Desmembramento de Pagamentos (Débitos):</strong>
                        <table class="table-interna">
                            <thead style="background: #f1f5f9;"><tr><th>Data</th><th>Histórico</th><th style="text-align:right;">Valor Pago</th></tr></thead>
                            <tbody>${pagamentosHTML}</tbody>
                        </table>
                    </td>
                </tr>
            `;
        }
    });
}

function alternarAgrupamento(index) {
    const trDetalhes = document.getElementById(`detalhes-${index}`);
    const btn = document.getElementById(`btn-toggle-${index}`);
    if (trDetalhes && btn) {
        trDetalhes.classList.toggle('hidden');
        btn.innerText = trDetalhes.classList.contains('hidden') ? '+' : '-';
    }
}

function exportarRelatorioFornecedoresXLSX() {
    let lista = filtroFornecedoresAtual === 'aberto' ? dadosFornecedoresProcessados.filter(n => Math.abs(n.saldo) > 0.01) : 
                filtroFornecedoresAtual === 'quitados' ? dadosFornecedoresProcessados.filter(n => Math.abs(n.saldo) <= 0.01) : 
                [...dadosFornecedoresProcessados];

    if (lista.length === 0) return alert("Não há dados para exportar.");

    const dados = lista.map(item => ({
        "Data Inicial": item.data,
        "Conta Contábil": item.contaContabil,
        "Nº da Nota Fiscal": item.numeroNF,
        "Histórico da Compra": item.descricao,
        "Compras (Créditos)": item.compras,
        "Pagamentos (Débitos)": item.pagamentos,
        "Saldo Final": item.saldo,
        "Status": Math.abs(item.saldo) > 0.01 ? "EM ABERTO" : "LIQUIDADO"
    }));

    const ws = XLSX.utils.json_to_sheet(dados);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Fornecedores");
    XLSX.writeFile(wb, `Relatorio_Fornecedores_${filtroFornecedoresAtual}.xlsx`);
}
