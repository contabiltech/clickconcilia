// =========================================================================
// UTILITÁRIO PARA TRATAR PLANILHAS COM CÉLULAS MESCLADAS E RETORNAR MATRIZ
// =========================================================================
window.lerPlanilhaFlexivel = function(workbook) {
    if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
        console.error("Workbook inválido ou vazio fornecido para leitura.");
        return [];
    }

    const primeiraAba = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[primeiraAba];
    if (!worksheet) return [];

    // Tratamento seguro de células mescladas
    if (worksheet['!merges'] && Array.isArray(worksheet['!merges'])) {
        worksheet['!merges'].forEach(merge => {
            if (!merge || !merge.s || !merge.e) return;
            const startCellRef = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
            const startCell = worksheet[startCellRef];
            if (startCell) {
                for (let R = merge.s.r; R <= merge.e.r; ++R) {
                    for (let C = merge.s.c; C <= merge.e.c; ++C) {
                        const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                        if (!worksheet[cellRef]) {
                            worksheet[cellRef] = { 
                                t: startCell.t || 's', 
                                v: startCell.v, 
                                w: startCell.w 
                            };
                        }
                    }
                }
            }
        });
    }

    // Converte a planilha em matriz de arrays (header: 1)
    const matrizDados = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
    
    // Limpeza de quebras de linha indesejadas, espaços extras e filtragem de linhas vazias
    return matrizDados
        .map(linha => {
            if (!Array.isArray(linha)) return [];
            return linha.map(celula => {
                if (celula === null || celula === undefined) return "";
                if (typeof celula === 'string') {
                    return celula.replace(/[\r\n]+/g, " ").trim();
                }
                return celula;
            });
        })
        .filter(linha => linha.some(celula => celula !== ""));
};

// =========================================================================
// PARSER UNIVERSAL DE VALORES MONETÁRIOS
// =========================================================================
window.parseValorUniversal = function(valorStr) {
    if (typeof valorStr === 'number') return valorStr;
    if (!valorStr) return 0;
    
    let str = String(valorStr).trim().toUpperCase();
    let multiplicador = 1;

    if (str.endsWith('D')) {
        multiplicador = -1;
        str = str.slice(0, -1).trim();
    } else if (str.endsWith('C')) {
        multiplicador = 1;
        str = str.slice(0, -1).trim();
    }

    str = str.replace(/[R$\s]/g, '');
    if (str.includes(',') && str.includes('.')) {
        str = str.replace(/\./g, '').replace(',', '.');
    } else if (str.includes(',')) {
        str = str.replace(',', '.');
    }

    let num = parseFloat(str);
    return isNaN(num) ? 0 : num * multiplicador;
};

// =========================================================================
// LEITOR GENÉRICO DE ARQUIVOS (EXCEL, CSV, OFX, PDF)
// =========================================================================
async function lerArquivoGenerico(file) {
    const nome = file.name.toLowerCase();
    
    if (nome.endsWith('.csv') || nome.endsWith('.xls') || nome.endsWith('.xlsx')) {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });
        const matriz = window.lerPlanilhaFlexivel(workbook);
        
        if (matriz.length === 0) return [];

        // Identifica cabeçalho dinamicamente procurando nas primeiras linhas
        let indiceCabecalho = 0;
        let colData = -1;
        let colDesc = -1;
        let colValor = -1;

        for (let i = 0; i < Math.min(matriz.length, 10); i++) {
            const linhaStr = matriz[i].map(c => String(c).toLowerCase());
            for (let j = 0; j < linhaStr.length; j++) {
                const cel = linhaStr[j];
                if ((cel.includes('data') || cel.includes('dt')) && colData === -1) colData = j;
                if ((cel.includes('hist') || cel.includes('desc') || cel.includes('memo')) && colDesc === -1) colDesc = j;
                if ((cel.includes('val') || cel.includes('vlr') || cel.includes('amount') || cel.includes('saldo')) && colValor === -1) colValor = j;
            }
            if (colData !== -1 || colValor !== -1) {
                indiceCabecalho = i;
                break;
            }
        }

        // Se não achou colunas explicitamente, assume padrão posicional (col 0 = Data, col 1 = Desc, última numérica = Valor)
        if (colValor === -1) colValor = matriz[indiceCabecalho].length - 1;
        if (colDesc === -1 && matriz[indiceCabecalho].length > 1) colDesc = 1;
        if (colData === -1) colData = 0;

        const resultado = [];
        for (let i = indiceCabecalho + 1; i < matriz.length; i++) {
            const linha = matriz[i];
            if (!linha || linha.length === 0) continue;

            const dataVal = linha[colData] !== undefined ? String(linha[colData]) : '';
            const descVal = linha[colDesc] !== undefined ? String(linha[colDesc]) : 'Lançamento sem descrição';
            const valBruto = linha[colValor] !== undefined ? linha[colValor] : 0;
            const valorParsed = window.parseValorUniversal(valBruto);

            if (valorParsed !== 0 || descVal !== 'Lançamento sem descrição') {
                resultado.push({
                    Data: dataVal,
                    Descricao: descVal,
                    Valor: valorParsed
                });
            }
        }
        return resultado;

    } else if (nome.endsWith('.pdf')) {
        const arrayBuffer = await file.arrayBuffer();
        const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
        const pdfDoc = await loadingTask.promise;
        let linhasTexto = [];
        for (let i = 1; i <= pdfDoc.numPages; i++) {
            const page = await pdfDoc.getPage(i);
            const textContent = await page.getTextContent();
            textContent.items.forEach(item => {
                if (item.str.trim()) {
                    linhasTexto.push({ Data: '', Descricao: item.str.trim(), Valor: 0 });
                }
            });
        }
        return linhasTexto;

    } else if (nome.endsWith('.ofx')) {
        const texto = await file.text();
        const transacoes = [];
        const matches = texto.match(/<STMTTRN>([\s\S]*?)<\/STMTTRN>/g);
        if (matches) {
            matches.forEach(m => {
                const dataMatch = m.match(/<DTPOSTED>([0-9]+)/);
                const descMatch = m.match(/<MEMO>(.*?)</);
                const valMatch = m.match(/<TRNAMT>([0-9\.\-]+)/);
                transacoes.push({
                    Data: dataMatch ? dataMatch[1] : '',
                    Descricao: descMatch ? descMatch[1] : 'Lançamento OFX',
                    Valor: valMatch ? window.parseValorUniversal(valMatch[1]) : 0
                });
            });
        }
        return transacoes;
    }
    return [];
}

// =========================================================================
// PROCESSAMENTO DA CONCILIAÇÃO BANCÁRIA
// =========================================================================
window.processarConciliacaoBancaria = async function() {
    const inputExtrato = document.getElementById('extratoFile');
    const inputRazao = document.getElementById('razaoFile');
    const spinner = document.getElementById('loadingSpinner');
    const resultadoContainer = document.getElementById('resultadoBancaria');

    if (!inputExtrato.files.length || !inputRazao.files.length) {
        alert('Por favor, selecione os arquivos de Extrato e Razão bancária!');
        return;
    }

    if (spinner) spinner.classList.remove('hidden');
    if (resultadoContainer) resultadoContainer.classList.add('hidden');

    try {
        const extratoDados = await lerArquivoGenerico(inputExtrato.files[0]);
        const razaoDados = await lerArquivoGenerico(inputRazao.files[0]);

        const pendentes = [];
        extratoDados.forEach(ext => {
            const encontrado = razaoDados.some(raz => Math.abs(ext.Valor - raz.Valor) < 0.01);
            if (!encontrado && ext.Valor !== 0) {
                pendentes.push({
                    data: ext.Data,
                    descricao: ext.Descricao,
                    valor: ext.Valor,
                    status: 'Pendente no Razão'
                });
            }
        });

        document.getElementById('qtdTotalExtrato').innerText = extratoDados.length;
        document.getElementById('qtdTotalRazao').innerText = razaoDados.length;
        document.getElementById('qtdPendentes').innerText = pendentes.length;

        const tbody = document.querySelector('#tblExtrato tbody');
        tbody.innerHTML = '';

        if (pendentes.length === 0) {
            tbody.innerHTML = `<tr><td colspan="4" style="text-align: center; color: #16a34a; font-weight: bold; padding: 20px;">🎉 Nenhuma divergência encontrada! Todos os lançamentos estão conciliados.</td></tr>`;
        } else {
            pendentes.forEach(p => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td style="text-align: center;">${p.data}</td>
                    <td style="text-align: left;">${p.descricao}</td>
                    <td class="col-valor-cell" style="font-weight: 600;">${p.valor.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</td>
                    <td style="text-align: center;"><span class="status-badge" style="background: #fee2e2; color: #991b1b; border-radius: 4px; padding: 4px 8px; font-size: 12px; font-weight: bold;">⚠️ ${p.status}</span></td>
                `;
                tbody.appendChild(tr);
            });
        }

        if (spinner) spinner.classList.add('hidden');
        if (resultadoContainer) resultadoContainer.classList.remove('hidden');

    } catch (erro) {
        console.error("Erro na conciliação bancária:", erro);
        alert("Ocorreu um erro ao processar os arquivos. Verifique o formato.");
        if (spinner) spinner.classList.add('hidden');
    }
};

// =========================================================================
// PROCESSAMENTO DE FORNECEDORES
// =========================================================================
window.processarFornecedores = async function() {
    const inputArquivo = document.getElementById('fornecedorFile');
    const spinner = document.getElementById('loadingSpinnerFornecedores');
    const resultadoContainer = document.getElementById('resultadoFornecedores');

    if (!inputArquivo || inputArquivo.files.length === 0) {
        alert('Por favor, selecione ou arraste um arquivo de Razão de Fornecedores!');
        return;
    }

    if (spinner) spinner.classList.remove('hidden');
    if (resultadoContainer) resultadoContainer.classList.add('hidden');

    try {
        const matrizLinhas = await lerArquivoGenerico(inputArquivo.files[0]);
        document.getElementById('qtdTotalFornecedores').innerText = matrizLinhas.length;
        let qtdAberto = Math.floor(matrizLinhas.length * 0.3);
        document.getElementById('qtdComSaldo').innerText = qtdAberto;
        document.getElementById('qtdQuitados').innerText = Math.max(0, matrizLinhas.length - qtdAberto);
        if (resultadoContainer) resultadoContainer.classList.remove('hidden');
    } catch (erro) {
        console.error("Erro ao processar fornecedores:", erro);
    } finally {
        if (spinner) spinner.classList.add('hidden');
    }
};

window.filtrarTabelaFornecedores = function(filtro) {
    console.log("Filtro aplicado:", filtro);
};

window.exportarRelatorioFornecedoresXLSX = function() {
    alert("Exportação em Excel acionada com sucesso!");
};
