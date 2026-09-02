function normalizarDadosPlanilha(linhas, origem) {
    let listaNormalizada = [];
    if (!linhas || linhas.length === 0) return listaNormalizada;

    // 1. Trata linhas separadas por ponto e vírgula (.CSV com coluna única)
    linhas = linhas.map(linha => {
        if (linha.length === 1 && typeof linha[0] === 'string' && linha[0].includes(';')) {
            return linha[0].split(';');
        }
        return linha;
    });

    // 2. Mapeamento Dinâmico de Colunas (Varre até 20 linhas para achar o cabeçalho real)
    let idxData = -1, idxDesc = -1, idxDebito = -1, idxCredito = -1, idxValor = -1, idxConta = -1;
    let linhaCabecalhoIndex = -1;

    for (let i = 0; i < Math.min(20, linhas.length); i++) {
        const linhaHeader = linhas[i].map(c => String(c || '').toLowerCase().trim());
        
        const dataFind = linhaHeader.findIndex(c => c.includes('data') || c.includes('date') || c === 'dt');
        const descFind = linhaHeader.findIndex(c => c.includes('desc') || c.includes('historico') || c.includes('histó') || c.includes('detalhe'));
        const debitoFind = linhaHeader.findIndex(c => c.includes('débito') || c.includes('debito') || c === 'deb');
        const creditoFind = linhaHeader.findIndex(c => c.includes('crédito') || c.includes('credito') || c === 'cred');
        const valorFind = linhaHeader.findIndex(c => c.includes('valor') || c.includes('monto') || c.includes('quant') || c.includes('saldo'));
        const contaFind = linhaHeader.findIndex(c => c.includes('conta') || c.includes('cto') || c.includes('codigo') || c.includes('cód'));

        // Se encontrou ao menos a coluna de Data ou Descrição, assume este índice como início dos dados
        if (dataFind !== -1 || descFind !== -1) {
            idxData = dataFind;
            idxDesc = descFind;
            idxDebito = debitoFind;
            idxCredito = creditoFind;
            idxValor = valorFind;
            idxConta = contaFind;
            linhaCabecalhoIndex = i;
            break;
        }
    }

    // Fallbacks padrão caso não encontre cabeçalhos explícitos
    if (idxData === -1) idxData = 0;
    if (idxDesc === -1) idxDesc = 1;
    if (idxValor === -1 && idxDebito === -1) idxValor = 2;

    let contaAtualContexto = 'N/A';

    // Helper interno para converter texto de moeda em float
    const parseValor = (valStr) => {
        if (!valStr) return 0;
        let limpo = String(valStr).replace('R$', '').replace(/\s/g, '').trim();
        if (limpo.includes(',') && limpo.includes('.')) {
            limpo = limpo.replace(/\./g, '').replace(',', '.');
        } else if (limpo.includes(',')) {
            limpo = limpo.replace(',', '.');
        }
        const num = parseFloat(limpo);
        return isNaN(num) ? 0 : num;
    };

    // 3. Processamento das Linhas de Dados
    for (let i = linhaCabecalhoIndex + 1; i < linhas.length; i++) {
        let linha = linhas[i];
        if (!linha || linha.length === 0) continue;

        const conteudoLinha = linha.map(c => String(c || '').trim()).join(' ');
        const conteudoLower = conteudoLinha.toLowerCase();

        // Ignora linhas de total, rodapé ou saldos que distorcem a conciliação
        if (conteudoLower.includes('total') || 
            conteudoLower.includes('saldo anterior') || 
            conteudoLower.includes('saldo atual') || 
            conteudoLower.includes('transporte')) {
            continue;
        }

        // Detecta quando o ERP Domínio coloca a Conta em uma linha isolada de agrupamento
        if (conteudoLower.includes('conta:') || conteudoLower.includes('código:')) {
            contaAtualContexto = conteudoLinha;
            continue;
        }

        const strData = idxData !== -1 ? String(linha[idxData] || '').trim() : '';

        // Ignora reaparições do cabeçalho
        if (strData.toLowerCase().includes('data') || strData.toLowerCase().includes('date')) continue;

        let dataFormatada = formatarDataBR(strData);
        let descricao = idxDesc !== -1 ? String(linha[idxDesc] || '').trim() : '';
        let contaContabil = idxConta !== -1 && linha[idxConta] ? String(linha[idxConta]).trim() : contaAtualContexto;

        // Cálculo de valor considerando Débito/Crédito em colunas separadas ou coluna única
        let valorFinal = 0;
        if (idxDebito !== -1 || idxCredito !== -1) {
            const valDebito = idxDebito !== -1 ? parseValor(linha[idxDebito]) : 0;
            const valCredito = idxCredito !== -1 ? parseValor(linha[idxCredito]) : 0;
            valorFinal = valCredito - valDebito; // Convenção: Crédito (positivo), Débito (negativo)
        } else if (idxValor !== -1) {
            valorFinal = parseValor(linha[idxValor]);
        }

        // Adiciona apenas se houver data válida ou descrição consistente com valor não zerado
        if ((dataFormatada !== 'N/A' || descricao.length > 2) && Math.abs(valorFinal) > 0.001) {
            listaNormalizada.push({
                id: `${origem}_${i}`,
                data: dataFormatada,
                descricao: descricao || 'Sem Histórico',
                contaContabil: contaContabil,
                valor: Math.round(valorFinal * 100) / 100,
                conciliado: false
            });
        }
    }

    return listaNormalizada;
}
