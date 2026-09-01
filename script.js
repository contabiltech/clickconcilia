function normalizarDadosPlanilha(linhas, origem) {
    let listaNormalizada = [];
    if (!linhas || linhas.length === 0) return listaNormalizada;

    console.log("Linhas brutas lidas da planilha:", linhas.length);

    let idxData = -1, idxDesc = -1, idxDebito = -1, idxCredito = -1, idxValorUnico = -1, idxConta = -1;

    // 1. Procura o cabeçalho real varrendo até as primeiras 15 linhas
    for (let i = 0; i < Math.min(15, linhas.length); i++) {
        const linhaHeader = (linhas[i] || []).map(c => String(c || '').toLowerCase().trim());
        
        const dataFind = linhaHeader.findIndex(c => c === 'data' || c === 'dt' || c.includes('data'));
        const descFind = linhaHeader.findIndex(c => c.includes('histórico') || c.includes('historico') || c.includes('desc') || c.includes('hist'));
        const debFind = linhaHeader.findIndex(c => c.includes('débito') || c.includes('debito'));
        const credFind = linhaHeader.findIndex(c => c.includes('crédito') || c.includes('credito'));
        const valFind = linhaHeader.findIndex(c => c === 'valor' || c === 'saldo');
        const contaFind = linhaHeader.findIndex(c => c.includes('conta') || c.includes('cto') || c.includes('c.c.part'));

        if (dataFind !== -1 && idxData === -1) idxData = dataFind;
        if (descFind !== -1 && idxDesc === -1) idxDesc = descFind;
        if (debFind !== -1 && idxDebito === -1) idxDebito = debFind;
        if (credFind !== -1 && idxCredito === -1) idxCredito = credFind;
        if (valFind !== -1 && idxValorUnico === -1) idxValorUnico = valFind;
        if (contaFind !== -1 && idxConta === -1) idxConta = contaFind;

        // Se encontrou os principais, pode parar de procurar o cabeçalho
        if (idxData !== -1 && idxDesc !== -1 && (idxDebito !== -1 || idxValorUnico !== -1)) {
            break;
        }
    }

    // Fallbacks inteligentes baseados no padrão comum de relatórios contábeis
    if (idxData === -1) idxData = 0;
    if (idxDesc === -1) idxDesc = 2;
    if (idxConta === -1) idxConta = 1;

    console.log(`Colunas detectadas -> Data: ${idxData}, Descrição: ${idxDesc}, Débito: ${idxDebito}, Crédito: ${idxCredito}, Valor: ${idxValorUnico}, Conta: ${idxConta}`);

    // 2. Processamento das linhas de dados
    for (let i = 0; i < linhas.length; i++) {
        let linha = linhas[i];
        if (!linha || linha.length === 0) continue;

        const strData = String(linha[idxData] || '').trim();
        
        // Pula cabeçalhos repetidos, rodapés ou linhas de saldo anterior
        const strLower = strData.toLowerCase();
        if (strLower === 'data' || strLower.includes('total') || strLower.includes('saldo anterior') || strLower.includes('empresa:')) {
            continue;
        }

        let dataFormatada = formatarDataBR(strData);
        // Se a data for inválida (ex: linha de texto solta ou código), pula
        if (dataFormatada === 'N/A' || dataFormatada.length < 10) continue;

        let descricao = idxDesc !== -1 && linha[idxDesc] ? String(linha[idxDesc]).trim() : 'Sem histórico';
        let contaContabil = idxConta !== -1 && linha[idxConta] ? String(linha[idxConta]).trim() : 'N/A';
        
        let valorFinal = 0;
        let tipoMovimento = 'D'; // Padrão

        // Tratamento para colunas separadas de Débito e Crédito (Razão Contábil)
        if (idxDebito !== -1 || idxCredito !== -1) {
            let valDebito = idxDebito !== -1 ? parseValorUniversal(linha[idxDebito]) : 0;
            let valCredito = idxCredito !== -1 ? parseValorUniversal(linha[idxCredito]) : 0;

            if (valDebito !== 0 && !isNaN(valDebito)) {
                valorFinal = -Math.abs(valDebito); // Débito normalmente reduz saldo ou é negativo no fluxo
                tipoMovimento = 'D';
            } else if (valCredito !== 0 && !isNaN(valCredito)) {
                valorFinal = Math.abs(valCredito); // Crédito positivo
                tipoMovimento = 'C';
            }
        } else {
            // Coluna única de valor
            let colAlvo = idxValorUnico !== -1 ? idxValorUnico : linha.length - 1;
            let valBruto = parseValorUniversal(linha[colAlvo]);
            if (!isNaN(valBruto)) {
                valorFinal = valBruto;
            }
        }

        // Adiciona à lista se o valor for válido (permitindo zero apenas se estritamente necessário, mas idealmente diferente de zero)
        if (!isNaN(valorFinal) && dataFormatada !== 'N/A') {
            listaNormalizada.push({
                id: `${origem}_${i}`,
                data: dataFormatada,
                descricao: descricao,
                contaContabil: contaContabil,
                valor: Math.round(valorFinal * 100) / 100,
                tipo: tipoMovimento,
                conciliado: false
            });
        }
    }

    console.log("Linhas normalizadas com sucesso:", listaNormalizada.length);
    return listaNormalizada;
}
