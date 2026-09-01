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
        .filter(linha => linha.some(celula => celula !== "")); // Remove linhas totalmente vazias
};
