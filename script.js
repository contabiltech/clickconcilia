function normalizarDadosPlanilha(linhas, origem) {
    let listaNormalizada = [];
    if (!linhas || linhas.length === 0) return listaNormalizada;

    console.log("Linhas brutas lidas da planilha:", linhas);

    let idxData = -1, idxDesc = -1, idxValor = -1, idxConta = -1;

    // Procura o cabeçalho real varrendo até as primeiras 15 linhas
    for (let i = 0; i < Math.min(15, linhas.length); i++) {
        const linhaHeader = (linhas[i] || []).map(c => String(c || '').toLowerCase().trim());
        
        const dataFind = linhaHeader.findIndex(c => c.includes('data') || c.includes('dt') || c.includes('emissao'));
        const descFind = linhaHeader.findIndex(c => c.includes('desc') || c.includes('historico') || c.includes('hist') || c.includes('obs'));
        const valorFind = linhaHeader.findIndex(c => c.includes('valor') || c.includes('saldo') || c.includes('total') || c.includes('débito') || c.includes('credito'));
        const contaFind = linhaHeader.findIndex(c => c.includes('conta') || c.includes('cto') || c.includes('codigo') || c.includes('fornecedor'));

        if (dataFind !== -1 && idxData === -1) idxData = dataFind;
        if (descFind !== -1 && idxDesc === -1) idxDesc = descFind;
        if (valorFind !== -1 && idxValor === -1) idxValor = valorFind;
        if (contaFind !== -1 && idxConta === -1) idxConta = contaFind;
    }

    // Se não encontrou cabeçalho explícito por nome, assume colunas padrão (0: Data, 1: Conta, 2: Descrição, etc.)
    if (idxData === -1) idxData = 0;
    if (idxDesc === -1) idxDesc = 2;
    if (idxValor === -1) idxValor = 5; // ou coluna de valor padrão

    console.log(`Colunas detectadas -> Data: ${idxData}, Descrição: ${idxDesc}, Valor: ${idxValor}, Conta: ${idxConta}`);

    for (let i = 0; i < linhas.length; i++) {
        let linha = linhas[i];
        if (!linha || linha.length === 0) continue;

        const strData = String(linha[idxData] || '').trim();
        
        // Pula linhas de cabeçalho ou totais gerais do relatório
        if (strData.toLowerCase().includes('data') || strData.toLowerCase().includes('total') || strData.toLowerCase().includes('saldo anterior')) {
            continue;
        }

        let dataFormatada = formatarDataBR(strData);
        let descricao = idxDesc !== -1 ? String(linha[idxDesc] || 'Sem histórico').trim() : 'Sem histórico';
        let contaContabil = idxConta !== -1 && linha[idxConta] ? String(linha[idxConta]).trim() : 'N/A';
        
        let valor = 0;
        // Se a linha tem valores separados de débito e crédito nas colunas à frente
        if (linha.length > 5 && idxValor !== -1) {
            let valBruto = linha[idxValor];
            valor = parseValorUniversal(valBruto);
        } else {
            // Tenta pegar o primeiro valor numérico válido da linha se a coluna exata falhar
            valor = parseValorUniversal(linha[idxValor !== -1 ? idxValor : linha.length - 1]);
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

    console.log("Linhas normalizadas com sucesso:", listaNormalizada.length);
    return listaNormalizada;
}
