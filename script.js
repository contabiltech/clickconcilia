// =========================================================================
        // UTILITÁRIO PARA TRATAR PLANILHAS COM CÉLULAS MESCLADAS E RETORNAR MATRIZ
        // =========================================================================
        window.lerPlanilhaFlexivel = function(workbook) {
            const primeiraAba = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[primeiraAba];

            // Tratamento de células mescladas
            if (worksheet['!merges']) {
                worksheet['!merges'].forEach(merge => {
                    const startCellRef = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
                    const startCell = worksheet[startCellRef];
                    if (startCell) {
                        for (let R = merge.s.r; R <= merge.e.r; ++R) {
                            for (let C = merge.s.c; C <= merge.e.c; ++C) {
                                const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                                if (!worksheet[cellRef]) {
                                    worksheet[cellRef] = { t: startCell.t, v: startCell.v, w: startCell.w };
                                }
                            }
                        }
                    }
                });
            }

            // IMPORTANTE: header: 1 retorna um array de arrays (linhas e colunas por índice)
            const matrizDados = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
            
            return matrizDados.map(linha => {
                return linha.map(celula => {
                    if (typeof celula === 'string') {
                        return celula.replace(/[\r\n]+/g, " ").trim();
                    }
                    return celula;
                });
            });
        };
