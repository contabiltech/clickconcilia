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

// Leitura de planilha com suporte a XLSX
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

// Função auxiliar para leitura de PDF com OCR automático forçado (caso utilize PDF.js / Tesseract)
async function lerArquivoPDFComOCR(file, origem) {
    return new Promise(async (resolve, reject) => {
        try {
            // Exemplo de integração com leitor PDF e fallback/forçamento de OCR se necessário
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
                        
                        // Forçar OCR automático se o texto extraído for insuficiente (ex: PDF escaneado/imagem)
                        if (textoPagina.trim().length < 20 && typeof Tesseract !== 'undefined') {
                            console.log(`[OCR Automático] Página ${i} parece ser imagem. Acionando OCR...`);
                            // Renderiza página em canvas para OCR
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
                    
                    // Processar linhas extraídas do PDF para o formato padrão
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
    let i = 0;
    
    linhas.forEach((linha, index) => {
        let linhaStr = linha.trim();
        // Regex básica para capturar datas DD/MM/AAAA e valores monetários
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
        
        // Verifica se há colunas separadas de Débito e Crédito no razão
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
        // Verifica se é PDF e aplica OCR se necessário
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
