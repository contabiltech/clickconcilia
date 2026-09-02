let dadosExtrato = [];
let dadosRazao = [];
let dadosFornecedoresProcessados = [];
let filtroFornecedoresAtual = 'todos';

// Necessário para o pdf.js localizar o worker
if (typeof pdfjsLib !== 'undefined') {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
}

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
    if (str.includes('-')) {
        const partes = str.split('T')[0].split('-');
        if (partes.length === 3) {
            const [ano, mes, dia] = partes;
            if (ano.length === 4) {
                return `${dia.padStart(2, '0')}/${mes.padStart(2, '0')}/${ano}`;
            }
        }
    }
    if (str.includes('/')) {
        const partes = str.split('/');
        if (partes.length === 3) {
            let [dia, mes, ano] = partes;
            if (ano.length === 2) ano = `20${ano}`;
            return `${dia.padStart(2, '0')}/${mes.padStart(2, '0')}/${ano}`;
        }
    }
    const dt = new Date(dataEntrada);
    if (!isNaN(dt.getTime())) {
        const dia = String(dt.getDate()).padStart(2, '0');
        const mes = String(dt.getMonth() + 1).padStart(2, '0');
        const ano = dt.getFullYear();
        return `${dia}/${mes}/${ano}`;
    }
    return str;
}

function formatBRL(v) {
    return (v || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

// =========================================================================
// DISPATCHER: escolhe o leitor certo pela extensão do arquivo
// =========================================================================
function lerArquivo(file, origem) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext === 'pdf') return lerArquivoPDF(file, origem);
    if (ext === 'ofx') return lerArquivoOFX(file, origem);
    return lerArquivoPlanilha(file, origem); // xlsx / csv
}

// =========================================================================
// LEITOR: XLSX / CSV (tabular, com cabeçalho identificável)
// =========================================================================
function lerArquivoPlanilha(file, origem) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const primeiraAba = workbook.SheetNames[0];
                const aba = workbook.Sheets[primeiraAba];
                const json = XLSX.utils.sheet_to_json(aba, { header: 1, raw: false });
                resolve(normalizarDadosPlanilha(json, origem));
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = (error) => reject(error);
        reader.readAsArrayBuffer(file);
    });
}

// Converte um valor de célula em número, detectando sufixo D/C (débito/crédito)
// quando presente (ex: "1.500,00D") e aceitando vírgula ou ponto decimal.
function parseValorComTipo(bruto) {
    if (bruto === undefined || bruto === null) return { valor: 0, tipo: null };
    let str = String(bruto).trim();
    if (!str) return { valor: 0, tipo: null };
    let tipo = null;
    if (/[DC]$/i.test(str) && /\d/.test(str)) {
        tipo = str.slice(-1).toUpperCase();
        str = str.slice(0, -1).trim();
    }
    str = str.replace('R$', '').replace(/\s/g, '');
    if (str.includes(',') && str.includes('.')) {
        str = str.replace(/\./g, '').replace(',', '.');
    } else if (str.includes(',')) {
        str = str.replace(',', '.');
    }
    let v = parseFloat(str);
    if (isNaN(v)) return { valor: 0, tipo };
    if (v < 0 && !tipo) tipo = 'D';
    return { valor: Math.abs(v), tipo };
}

// Procura, nas primeiras linhas da planilha, qual coluna corresponde a cada
// campo. Cobre nomenclaturas comuns usadas por ERPs (Domínio, Sismade,
// SAP, Totvs, planilhas exportadas de bancos etc.).
function detectarColunasPlanilha(linhaHeader) {
    const col = linhaHeader.map(c => String(c || '').toLowerCase().trim());
    const achar = (...termos) => col.findIndex(c => termos.some(t => c.includes(t)));
    return {
        data: achar('data', 'dt.', 'date'),
        historico: achar('históric', 'historic', 'descri', 'lançamento', 'lancamento', 'complemento', 'memo'),
        conta: achar('conta contáb', 'conta contab', 'cta cont', 'cód. conta', 'codigo conta', 'plano de conta', 'conta'),
        // Débito/Crédito separados: o padrão mais comum em razões contábeis de ERP
        debito: achar('débito', 'debito', 'valor débito', 'valor debito'),
        credito: achar('crédito', 'credito', 'valor crédito', 'valor credito'),
        // Coluna única de valor (extratos bancários, planilhas simples)
        valor: achar('valor', 'montante', 'quantia', 'monto')
    };
}

function normalizarDadosPlanilha(linhas, origem) {
    let listaNormalizada = [];
    if (!linhas || linhas.length === 0) return listaNormalizada;

    // trata CSV que caiu inteiro numa única "coluna" separada por ;
    linhas = linhas.map(linha => {
        if (linha.length === 1 && typeof linha[0] === 'string' && linha[0].includes(';')) {
            return linha[0].split(';');
        }
        return linha;
    });

    // Procura a linha de cabeçalho (até as 10 primeiras linhas) que tenha,
    // no mínimo, coluna de Data + (Valor único OU Débito/Crédito)
    let colunas = null;
    for (let i = 0; i < Math.min(10, linhas.length); i++) {
        if (!linhas[i]) continue;
        const teste = detectarColunasPlanilha(linhas[i]);
        if (teste.data !== -1 && (teste.valor !== -1 || teste.debito !== -1 || teste.credito !== -1)) {
            colunas = teste;
            break;
        }
    }
    // Se não achou cabeçalho reconhecível, assume o layout posicional clássico
    if (!colunas) {
        colunas = { data: 0, historico: 1, valor: 2, conta: -1, debito: -1, credito: -1 };
    }

    let chaveAnterior = null; // usada para descartar linhas duplicadas (células mescladas)
    for (let i = 0; i < linhas.length; i++) {
        const linha = linhas[i];
        if (!linha || linha.length === 0) continue;

        const strData = String(linha[colunas.data] || '').trim();
        if (!strData || /^data$|^date$/i.test(strData)) continue;

        const dataFormatada = formatarDataBR(strData);
        const descricao = String(linha[colunas.historico] ?? '').trim() || 'Sem histórico';
        const contaContabil = colunas.conta !== -1 && linha[colunas.conta]
            ? String(linha[colunas.conta]).trim() : 'N/A';

        let valor = 0, tipo = null;
        if (colunas.debito !== -1 || colunas.credito !== -1) {
            // Layout com colunas separadas de Débito e Crédito
            const deb = colunas.debito !== -1 ? parseValorComTipo(linha[colunas.debito]) : { valor: 0 };
            const cred = colunas.credito !== -1 ? parseValorComTipo(linha[colunas.credito]) : { valor: 0 };
            if (deb.valor > 0) { valor = deb.valor; tipo = 'D'; }
            else if (cred.valor > 0) { valor = cred.valor; tipo = 'C'; }
        } else if (colunas.valor !== -1) {
            // Layout com coluna única de valor (pode vir com sinal ou sufixo D/C)
            const r = parseValorComTipo(linha[colunas.valor]);
            valor = r.valor;
            tipo = r.tipo;
        }

        if (!valor || isNaN(valor)) continue;

        // Evita contar duas vezes a mesma linha quando o export do ERP repete
        // o conteúdo por causa de células mescladas
        const chaveLinha = `${dataFormatada}|${descricao}|${valor}|${tipo}`;
        if (chaveLinha === chaveAnterior) continue;
        chaveAnterior = chaveLinha;

        listaNormalizada.push({
            id: `${origem}_${i}`,
            data: dataFormatada,
            descricao: descricao,
            contaContabil: contaContabil,
            valor: Math.round(valor * 100) / 100,
            tipo: tipo, // 'D' = Débito/Pagamento, 'C' = Crédito/Compra, null = indefinido
            conciliado: false
        });
    }
    return listaNormalizada;
}

// =========================================================================
// LEITOR: OFX (extrato bancário padrão)
// =========================================================================
function lerArquivoOFX(file, origem) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const texto = e.target.result;
                const transacoes = [];
                const blocos = texto.split(/<STMTTRN>/i).slice(1);
                blocos.forEach((bloco, i) => {
                    const dataMatch = bloco.match(/<DTPOSTED>(\d{8})/i);
                    const valorMatch = bloco.match(/<TRNAMT>(-?[\d.,]+)/i);
                    const memoMatch = bloco.match(/<MEMO>([^<\r\n]+)/i) || bloco.match(/<NAME>([^<\r\n]+)/i);
                    if (dataMatch && valorMatch) {
                        const d = dataMatch[1];
                        const dataFormatada = `${d.substr(6, 2)}/${d.substr(4, 2)}/${d.substr(0, 4)}`;
                        const valor = parseFloat(valorMatch[1].replace(',', '.'));
                        if (!isNaN(valor) && valor !== 0) {
                            transacoes.push({
                                id: `${origem}_ofx_${i}`,
                                data: dataFormatada,
                                descricao: (memoMatch ? memoMatch[1].trim() : 'Sem histórico'),
                                contaContabil: 'N/A',
                                valor: Math.round(valor * 100) / 100,
                                tipo: valor < 0 ? 'D' : 'C',
                                conciliado: false
                            });
                        }
                    }
                });
                resolve(transacoes);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = (err) => reject(err);
        reader.readAsText(file);
    });
}

// =========================================================================
// LEITOR: PDF (Razão contábil — layout "Data Lote Histórico Cta.C.Part.")
// =========================================================================

// Reconstrói as linhas do PDF agrupando os itens de texto por posição (Y/X),
// já que o pdf.js entrega os textos soltos, sem quebras de linha reais.
async function extrairLinhasPDF(file) {
    const buffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
    const linhas = [];
    for (let p = 1; p <= pdf.numPages; p++) {
        const page = await pdf.getPage(p);
        const content = await page.getTextContent();
        const porLinha = {};
        content.items.forEach(item => {
            // agrupa textos que estão na mesma "altura" (tolerância de 2px)
            const y = Math.round(item.transform[5] / 2) * 2;
            if (!porLinha[y]) porLinha[y] = [];
            porLinha[y].push({ x: item.transform[4], texto: item.str });
        });
        Object.keys(porLinha)
            .map(Number)
            .sort((a, b) => b - a) // de cima para baixo
            .forEach(y => {
                const linha = porLinha[y]
                    .sort((a, b) => a.x - b.x) // da esquerda para a direita
                    .map(o => o.texto)
                    .join(' ')
                    .replace(/\s+/g, ' ')
                    .trim();
                if (linha) linhas.push(linha);
            });
    }
    return linhas;
}

// Pega o último token no formato de valor monetário (ex: 1.500,00 / 1.500,00C)
// dentro de uma string, pois nas linhas de transação o valor sempre vem no final.
function extrairUltimoValor(str) {
    const matches = str.match(/\d{1,3}(?:\.\d{3})*,\d{2}[DC]?/g);
    if (!matches || matches.length === 0) return null;
    return matches[matches.length - 1];
}

// Converte as linhas de texto extraídas do PDF em lançamentos.
// Layout esperado (Razão contábil):
//   "Conta: 15382 - 2.1.01.02.01.0001 Fabio Ramos De Alencar"   -> cabeçalho da conta
//   "PRESTAÇÃO DE SERVIÇOS CONF. NF # 76 ..."                    -> descrição
//   "03/06/2026 11522 874 1.500,00"                               -> data + lote + valor
function normalizarRazaoPDF(linhas, origem) {
    const lancamentos = [];
    const regexConta = /^Conta:\s*(\d+)\s*-\s*[\d.]+\s*(.*)$/i;
    const regexIgnorar = /^Total do m[eê]s|^SALDO ANTERIOR|^RAZ[ÃA]O$|^Folha:|^Período:|^C\.N\.P\.J\.:|^Empresa:|^Data Lote Hist/i;
    const regexTransacao = /^(\d{2}\/\d{2}\/\d{4})\s+(\d+)\s+(.+)$/;

    let contaAtual = 'N/A';
    let ultimaDescricao = 'Sem histórico';

    linhas.forEach(linhaBruta => {
        const linha = linhaBruta.trim();
        if (!linha) return;

        const mConta = linha.match(regexConta);
        if (mConta) {
            contaAtual = `${mConta[1]}${mConta[2] ? ' - ' + mConta[2].trim() : ''}`.trim();
            return;
        }
        if (regexIgnorar.test(linha)) return;

        const mTrans = linha.match(regexTransacao);
        if (mTrans) {
            const data = mTrans[1];
            const resto = mTrans[3];
            const valorToken = extrairUltimoValor(resto);
            if (valorToken) {
                const tipo = /D$/.test(valorToken) ? 'D' : (/C$/.test(valorToken) ? 'C' : null);
                const valorLimpo = valorToken.replace(/[DC]$/, '');
                const valor = parseFloat(valorLimpo.replace(/\./g, '').replace(',', '.'));

                // Se sobrar texto além do valor nessa mesma linha, é complemento da descrição
                const descExtra = resto.replace(valorToken, '').trim();
                const descricaoFinal = (descExtra && descExtra.length > 3 && !/^\d+\s*$/.test(descExtra))
                    ? `${ultimaDescricao} ${descExtra}`.trim()
                    : ultimaDescricao;

                if (!isNaN(valor) && valor !== 0) {
                    lancamentos.push({
                        id: `${origem}_${lancamentos.length}`,
                        data: formatarDataBR(data),
                        descricao: descricaoFinal,
                        contaContabil: contaAtual,
                        valor: Math.round(valor * 100) / 100,
                        tipo: tipo, // 'D' = Débito/Pagamento, 'C' = Crédito/Compra, null = indefinido
                        conciliado: false
                    });
                }
            }
            return;
        }

        // Linha que não é cabeçalho de conta nem linha de transação:
        // guarda como candidata a descrição do próximo lançamento,
        // removendo valores monetários (ex: "1.500,00C") que às vezes
        // vazam para a mesma linha por causa do layout do PDF.
        if (!/^\d/.test(linha)) {
            const semValor = linha.replace(/\d{1,3}(?:\.\d{3})*,\d{2}[DC]?\s*$/i, '').trim();
            ultimaDescricao = semValor || linha;
        }
    });

    return lancamentos;
}

async function lerArquivoPDF(file, origem) {
    const linhas = await extrairLinhasPDF(file);
    return normalizarRazaoPDF(linhas, origem);
}

// =========================================================================
// MÓDULO: CONCILIAÇÃO BANCÁRIA
// =========================================================================

// Classifica um lançamento como Pagamento (saída) ou Recebimento (entrada)
function classificarMovimento(item) {
    if (item.tipo === 'D') return 'Pagamento';
    if (item.tipo === 'C') return 'Recebimento';
    return item.valor < 0 ? 'Pagamento' : 'Recebimento';
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
        dadosExtrato = await lerArquivo(extratoInput, 'EXT');
        dadosRazao = await lerArquivo(razaoInput, 'RAZ');

        dadosExtrato.forEach(itemExtrato => {
            const matchIndex = dadosRazao.findIndex(itemRazao =>
                !itemRazao.conciliado && Math.abs(itemRazao.valor - itemExtrato.valor) < 0.001
            );
            if (matchIndex !== -1) {
                itemExtrato.conciliado = true;
                dadosRazao[matchIndex].conciliado = true;
            }
        });

        // Lançamentos que estão no extrato (banco) mas ainda faltam ser lançados no razão
        const faltamNoRazao = dadosExtrato
            .filter(item => !item.conciliado)
            .map(item => ({ ...item, status: `${classificarMovimento(item)} não lançado no Razão` }));

        // Lançamentos que estão no razão mas ainda não aparecem no extrato (a compensar / verificar)
        const sobramNoRazao = dadosRazao
            .filter(item => !item.conciliado)
            .map(item => ({ ...item, status: `${classificarMovimento(item)} lançado no Razão sem correspondência no Extrato — verificar` }));

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
        tbody.innerHTML = `<tr><td colspan="5" style="text-align: center; color: #16a34a; font-weight: bold; padding: 1rem;">✅ Perfeito! Todos os lançamentos foram conciliados.</td></tr>`;
        return;
    }
    lista.forEach(item => {
        const tr = document.createElement('tr');
        const valorFormatado = formatBRL(item.valor);
        const movimento = classificarMovimento(item);
        tr.innerHTML = `
            <td>${formatarDataBR(item.data)}</td>
            <td>${item.descricao}</td>
            <td style="color: ${item.valor < 0 ? '#dc2626' : '#16a34a'}; font-weight: bold;">${valorFormatado}</td>
            <td><span style="font-weight:bold; color: ${movimento === 'Pagamento' ? '#dc2626' : '#16a34a'};">${movimento === 'Pagamento' ? '↑ Pagamento' : '↓ Recebimento'}</span></td>
            <td><span style="background: #fef2f2; color: #dc2626; border: 1px solid #fecaca; padding: 3px 8px; border-radius: 4px; font-size: 0.8em; font-weight: bold;">❌ ${item.status}</span></td>
        `;
        tbody.appendChild(tr);
    });
}

// =========================================================================
// MÓDULO DE FORNECEDORES
// =========================================================================
function extrairNumeroNF(texto) {
    const match = texto.match(/(?:nf|nfe|nota|nota\s*fiscal)\s*[-:#]?\s*(\d+)/i);
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
        const lancamentos = await lerArquivo(fornecedorInput, 'FORN');
        if (lancamentos.length === 0) {
            alert("Nenhum lançamento válido encontrado no arquivo.");
            if (spinner) spinner.classList.add('hidden');
            return;
        }
        const notasFiscais = {};
        lancamentos.forEach(item => {
            const numNF = extrairNumeroNF(item.descricao);
            const chave = numNF ? `NF ${numNF}` : 'Sem NF Identificada';

            // Prioriza o indicador D/C extraído do PDF; se não houver, usa palavras-chave
            let ehPagamento;
            if (item.tipo === 'D') {
                ehPagamento = true;
            } else if (item.tipo === 'C') {
                ehPagamento = false;
            } else {
                const descLower = item.descricao.toLowerCase();
                ehPagamento = descLower.includes('pagto') || descLower.includes('pagamento') || item.valor < 0;
            }

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

        // saldo > 0  -> ainda deve (em aberto)
        // saldo < 0  -> pagou a mais do que a nota (verificar)
        // saldo == 0 -> quitado
        const emAberto = dadosFornecedoresProcessados.filter(n => n.saldo > 0.01);
        const quitadas = dadosFornecedoresProcessados.filter(n => Math.abs(n.saldo) <= 0.01);
        const pagosAMaior = dadosFornecedoresProcessados.filter(n => n.saldo < -0.01);
        const valorTotalAberto = emAberto.reduce((soma, n) => soma + n.saldo, 0);

        document.getElementById('qtdTotalFornecedores').innerText = lancamentos.length;
        document.getElementById('qtdComSaldo').innerText = emAberto.length;
        document.getElementById('qtdQuitados').innerText = quitadas.length;
        document.getElementById('valorTotalAberto').innerText = formatBRL(valorTotalAberto);
        document.getElementById('qtdVerificar').innerText = pagosAMaior.length;

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
        listaFiltrada = dadosFornecedoresProcessados.filter(n => n.saldo > 0.01);
        if (tituloTabela) tituloTabela.innerText = '📋 Títulos em Aberto';
    } else if (tipoFiltro === 'quitados') {
        listaFiltrada = dadosFornecedoresProcessados.filter(n => Math.abs(n.saldo) <= 0.01);
        if (tituloTabela) tituloTabela.innerText = '📋 Títulos Quitados';
    } else if (tipoFiltro === 'verificar') {
        listaFiltrada = dadosFornecedoresProcessados.filter(n => n.saldo < -0.01);
        if (tituloTabela) tituloTabela.innerText = '📋 Pagamentos a Maior — Verificar';
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
        const temSaldoAberto = item.saldo > 0.01;
        const temPagoAMaior = item.saldo < -0.01;
        const temMultiplosPagamentos = item.listaPagamentos.length > 1;

        let badgeStatus;
        let corLinha = '';
        if (temSaldoAberto) {
            badgeStatus = `<span style="background: #fef2f2; color: #dc2626; border: 1px solid #fecaca; padding: 4px 8px; border-radius: 4px; font-size: 0.8em; font-weight: bold;">⚠️ Saldo em Aberto</span>`;
            corLinha = '#fffbe2';
        } else if (temPagoAMaior) {
            badgeStatus = `<span style="background: #fff7ed; color: #c2410c; border: 1px solid #fed7aa; padding: 4px 8px; border-radius: 4px; font-size: 0.8em; font-weight: bold;">🔎 Verificar (pago a maior)</span>`;
            corLinha = '#fff2e6';
        } else {
            badgeStatus = `<span style="background: #f0fdf4; color: #15803d; border: 1px solid #bbf7d0; padding: 4px 8px; border-radius: 4px; font-size: 0.8em; font-weight: bold;">✅ Quitado</span>`;
        }

        const btnExpandir = temMultiplosPagamentos
            ? `<button onclick="alternarAgrupamento(${index})" id="btn-toggle-${index}" title="Ver múltiplos pagamentos (${item.listaPagamentos.length})" style="background: #e2e8f0; border: 1px solid #cbd5e1; border-radius: 4px; width: 24px; height: 24px; cursor: pointer; font-weight: bold; line-height: 1;">+</button>`
            : `<span style="color: #cbd5e1;">-</span>`;

        const trPrincipal = document.createElement('tr');
        if (corLinha) trPrincipal.style.backgroundColor = corLinha;
        trPrincipal.innerHTML = `
            <td style="text-align: center;">${btnExpandir}</td>
            <td>${formatarDataBR(item.data)}</td>
            <td>${item.contaContabil}</td>
            <td>${item.descricao}</td>
            <td><strong>${item.numeroNF}</strong></td>
            <td>${formatBRL(item.pagamentos)}</td>
            <td>${formatBRL(item.compras)}</td>
            <td style="font-weight: bold; color: ${temSaldoAberto ? '#dc2626' : (temPagoAMaior ? '#c2410c' : '#16a34a')};">${formatBRL(item.saldo)}</td>
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
        listaFiltrada = dadosFornecedoresProcessados.filter(n => n.saldo > 0.01);
    } else if (filtroFornecedoresAtual === 'quitados') {
        listaFiltrada = dadosFornecedoresProcessados.filter(n => Math.abs(n.saldo) <= 0.01);
    } else if (filtroFornecedoresAtual === 'verificar') {
        listaFiltrada = dadosFornecedoresProcessados.filter(n => n.saldo < -0.01);
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
        "Status": item.saldo > 0.01 ? "Saldo em Aberto" : (item.saldo < -0.01 ? "Verificar (pago a maior)" : "Quitado")
    }));
    const worksheet = XLSX.utils.json_to_sheet(dadosExportacao);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Conciliacao_Fornecedores");
    const nomeArquivo = `Relatorio_Fornecedores_${filtroFornecedoresAtual}.xlsx`;
    XLSX.writeFile(workbook, nomeArquivo);
}
