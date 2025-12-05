class GeradorEspecificacoes {
    constructor() {
        this.baseEspecificacoes = [];
        this.budgetData = [];
        this.currentPage = 1;
        this.itemsPerPage = 10;
        this.currentSearch = '';
        
        this.initializeEventListeners();
        this.setCurrentDate();
        this.loadLocalDatabase();
    }

    initializeEventListeners() {
        // Database upload
        const loadDbBtn = document.getElementById('load-db-btn');
        const dbFileInput = document.getElementById('db-file-input');
        const viewDbBtn = document.getElementById('view-db-btn');
        
        loadDbBtn.addEventListener('click', () => dbFileInput.click());
        dbFileInput.addEventListener('change', (e) => this.loadDatabaseFromFile(e));
        viewDbBtn.addEventListener('click', () => this.showDatabaseModal());

        // Budget upload
        const loadBudgetBtn = document.getElementById('load-budget-btn');
        const budgetFileInput = document.getElementById('budget-file-input');
        const clearBudgetBtn = document.getElementById('clear-budget-btn');
        
        loadBudgetBtn.addEventListener('click', () => budgetFileInput.click());
        budgetFileInput.addEventListener('change', (e) => this.loadBudgetFromFile(e));
        clearBudgetBtn.addEventListener('click', () => this.clearBudget());

        // Generate buttons
        const generateBtn = document.getElementById('generate-btn');
        const previewBtn = document.getElementById('preview-btn');
        const refreshPreviewBtn = document.getElementById('refresh-preview-btn');
        
        generateBtn.addEventListener('click', () => this.generateDocument());
        previewBtn.addEventListener('click', () => this.generatePreview());
        refreshPreviewBtn.addEventListener('click', () => this.generatePreview());

        // Modal controls
        const closeDbModal = document.getElementById('close-db-modal');
        const dbModal = document.getElementById('db-modal');
        const tabBtns = document.querySelectorAll('.tab-btn');
        const searchInput = document.getElementById('db-search');
        const prevPageBtn = document.getElementById('prev-page');
        const nextPageBtn = document.getElementById('next-page');
        const exportDbBtn = document.getElementById('export-db-btn');
        const importDbBtn = document.getElementById('import-db-btn');

        closeDbModal.addEventListener('click', () => dbModal.classList.add('hidden'));
        
        tabBtns.forEach(btn => {
            btn.addEventListener('click', () => this.switchTab(btn.dataset.tab));
        });

        searchInput.addEventListener('input', (e) => {
            this.currentSearch = e.target.value;
            this.currentPage = 1;
            this.updateDatabaseView();
        });

        prevPageBtn.addEventListener('click', () => {
            if (this.currentPage > 1) {
                this.currentPage--;
                this.updateDatabaseView();
            }
        });

        nextPageBtn.addEventListener('click', () => {
            const totalPages = Math.ceil(this.filteredItems.length / this.itemsPerPage);
            if (this.currentPage < totalPages) {
                this.currentPage++;
                this.updateDatabaseView();
            }
        });

        exportDbBtn.addEventListener('click', () => this.exportDatabase());
        importDbBtn.addEventListener('click', () => {
            document.getElementById('db-file-input').click();
        });

        // Download button
        const downloadBtn = document.getElementById('download-btn');
        downloadBtn.addEventListener('click', () => this.downloadGeneratedDocument());
    }

    setCurrentDate() {
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('project-date').value = today;
    }

    async loadDatabaseFromFile(event) {
        const file = event.target.files[0];
        if (!file) return;

        try {
            const data = await this.readExcelFile(file);
            this.baseEspecificacoes = data;
            
            // Salvar no localStorage
            localStorage.setItem('especificacoes_db', JSON.stringify(this.baseEspecificacoes));
            localStorage.setItem('db_last_updated', new Date().toISOString());
            localStorage.setItem('db_name', file.name);
            
            this.updateDatabaseStatus(true);
            this.updateDatabaseInfo();
            this.updateDatabaseView();
            
            alert('Base de dados carregada com sucesso!');
        } catch (error) {
            console.error('Erro ao carregar base de dados:', error);
            alert('Erro ao carregar base de dados. Verifique o formato do arquivo.');
        }
    }

    // script.js - Localize a função loadBudgetFromFile(event)

    async loadBudgetFromFile(event) {
        const file = event.target.files[0];
        if (!file) return;

        this.setMessage('Processando arquivo de orçamento...', 'loading');

        try {
            // 1. Lê o arquivo binário
            const data = await new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (e) => resolve(new Uint8Array(e.target.result));
                reader.onerror = reject;
                reader.readAsArrayBuffer(file);
            });
            
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0]; // Assume a primeira planilha
            const worksheet = workbook.Sheets[sheetName];

            // 2. DETECÇÃO DINÂMICA DO CABEÇALHO (Simulando a lógica do Python)
            let start_row = 1; // 1-indexed (XLSX.js)
            const max_rows_to_check = 10;
            
            // Lê as primeiras linhas como um array de arrays para checar os títulos
            const sheet_data = XLSX.utils.sheet_to_json(worksheet, { 
                header: 1, 
                range: 'A1:Z' + max_rows_to_check, 
                raw: false 
            });

            // Procura a linha que contém "Item" e "Descrição"
            for (let i = 0; i < sheet_data.length; i++) {
                const row = sheet_data[i];
                if (row && row.length > 0) {
                    const first_cell = String(row[0] || '').trim();
                    // Busca por 'Descrição' ou 'Descricao' em qualquer coluna da linha
                    const contains_desc = row.some(cell => String(cell || '').includes('Descrição') || String(cell || '').includes('Descricao'));
                    
                    // Condição: 'Item' na primeira célula E 'Descrição' em algum lugar da linha
                    if (first_cell === 'Item' && contains_desc) {
                        // i é 0-indexed, o cabeçalho do XLSX é 1-indexed
                        start_row = i + 1; 
                        break;
                    }
                }
            }
            
            // 3. CONVERTE PARA JSON usando a linha de cabeçalho encontrada
            const rawBudgetData = XLSX.utils.sheet_to_json(worksheet, { 
                header: start_row, // Usa a linha dinâmica
                raw: false // Garante que os valores vêm como strings/números formatados
            });
            
            // 4. PROCESSA/MAPEIA DADOS (usando a função robusta)
            this.budgetData = this.processBudgetData(rawBudgetData);
            
            if (this.budgetData.length === 0) {
                 throw new Error("O arquivo foi lido, mas não contém itens válidos ou as colunas esperadas não foram encontradas. Verifique as colunas 'Item', 'Código' e 'Quant.'.");
            }

            this.showBudgetPreview();
            this.updateBudgetStatus(true);
            this.checkReadyToGenerate();
            this.setMessage(`Orçamento carregado com sucesso! ${this.budgetData.length} itens encontrados.`, 'success');
            
            // Gerar prévia automática se base estiver carregada
            if (this.baseEspecificacoes.length > 0) {
                this.generatePreview();
            }
        } catch (error) {
            console.error('Erro ao carregar orçamento:', error);
            this.showError(`Erro ao carregar orçamento: ${error.message}.`);
        }
    }

    // script.js - Localize esta função

    async readExcelFile(file, sheetToJsonOptions = {}) { // <--- Adiciona o parâmetro de opções
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // Converter todas as planilhas para JSON
                    const result = [];
                    workbook.SheetNames.forEach(sheetName => {
                        const worksheet = workbook.Sheets[sheetName];
                        // Usa o parâmetro 'sheetToJsonOptions' aqui
                        const jsonData = XLSX.utils.sheet_to_json(worksheet, sheetToJsonOptions); 
                        result.push(...jsonData);
                    });
                    
                    resolve(result);
                } catch (error) {
                    reject(error);
                }
            };
            
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    processBudgetData(data) {
        // Mapeamento flexível de chaves (simulando a lógica do Python)
        // A ordem é importante: a primeira chave encontrada será usada.
        const keyMap = {
            'Código': ['Código', 'Codigo'],
            'Descrição': ['Descrição', 'Descricao'],
            'Quant.': ['Quant.', 'Quant', 'QTDE', 'Quantidade'], 
            'Item': ['Item'],
            'Banco': ['Banco'],
            'Und': ['Und', 'UNID']
        };

        // Função auxiliar para encontrar o valor correto dada uma lista de possíveis chaves
        const findKey = (item, potentialKeys) => {
            for (const key of potentialKeys) {
                if (item.hasOwnProperty(key)) {
                    return item[key];
                }
            }
            return undefined;
        };

        // Filtra e processa dados do orçamento
        return data.map(item => {
            // Usa o mapeamento flexível para obter os valores
            const itemValue = findKey(item, keyMap['Item']);
            const codigoValue = findKey(item, keyMap['Código']);
            const descricaoValue = findKey(item, keyMap['Descrição']);
            const quantidadeValue = findKey(item, keyMap['Quant.']);
            
            const codigo = codigoValue ? String(codigoValue).trim() : '';
            const descricao = descricaoValue ? String(descricaoValue).trim() : '';

            // 1. Filtros (Remove linhas vazias, nulas ou totais, como no Python)
            if (!itemValue || 
                String(itemValue).toLowerCase().includes('total') || 
                String(itemValue).toLowerCase().includes('nan')) {
                return null;
            }
            
            // 2. Filtro principal (Deve ter Código e Descrição)
            if (codigo && descricao) {
                // Limpeza do Código (remove caracteres não-alfanuméricos)
                const cleanedCodigo = codigo.replace(/[^a-zA-Z0-9]/g, ''); 
                
                // Processamento da Quantidade
                let quantidade = 0;
                if (quantidadeValue !== undefined && quantidadeValue !== null) {
                    // Tenta converter para float, tratando o separador decimal (se for string)
                    let quantStr = String(quantidadeValue);
                    // Remove ponto como separador de milhar e usa ponto como decimal (se houver vírgula)
                    quantStr = quantStr.replace(/\./g, '').replace(',', '.'); 
                    quantidade = parseFloat(quantStr) || 0;
                }
                
                // Retorna o objeto padronizado
                return {
                    item: String(itemValue).trim(),
                    codigo: cleanedCodigo,
                    descricao: descricao,
                    banco: findKey(item, keyMap['Banco']) || '',
                    und: findKey(item, keyMap['Und']) || '',
                    quantidade: quantidade, 
                    itemOriginal: item 
                };
            }
            return null;
        }).filter(item => item !== null && item.quantidade > 0); // Filtra itens nulos e com quantidade zero
    }

    loadLocalDatabase() {
        try {
            const savedDb = localStorage.getItem('especificacoes_db');
            if (savedDb) {
                this.baseEspecificacoes = JSON.parse(savedDb);
                this.updateDatabaseStatus(true);
                this.updateDatabaseInfo();
            } else {
                // Criar base de dados de exemplo se não existir
                this.createSampleDatabase();
            }
        } catch (error) {
            console.error('Erro ao carregar base local:', error);
        }
    }

    createSampleDatabase() {
        this.baseEspecificacoes = [
            {
                "COMPOSIÇÃO": "101",
                "Banco": "SINAPI",
                "DESCRIÇÃO": "CABO DE COBRE ISOLADO PVC 750V - 2,5 MM²",
                "ESPECIFICAÇÃO TÉCNICA": "Cabo flexível de cobre, isolamento em PVC 750V/1000V, seção 2,5mm², coloração conforme NBR 7286. Deve possuir certificado INMETRO e laudo técnico."
            },
            {
                "COMPOSIÇÃO": "102",
                "Banco": "SINAPI",
                "DESCRIÇÃO": "DISJUNTOR TERMOMAGNÉTICO MONOPOLAR 10A CURVA C",
                "ESPECIFICAÇÃO TÉCNICA": "Disjuntor termomagnético DIN, monopolar, corrente nominal 10A, curva C, tensão 127/220V, conforme NBR NM 60898. Deve possuir certificado INMETRO."
            }
        ];
        
        localStorage.setItem('especificacoes_db', JSON.stringify(this.baseEspecificacoes));
        this.updateDatabaseStatus(true);
        this.updateDatabaseInfo();
    }

    updateDatabaseStatus(loaded) {
        const dbStatus = document.getElementById('db-status');
        const dbInfo = document.querySelector('.db-info');
        const viewDbBtn = document.getElementById('view-db-btn');
        const statusDb = document.getElementById('status-db');
        
        if (loaded) {
            dbStatus.className = 'status-badge status-success';
            dbStatus.innerHTML = '<i class="fas fa-check-circle"></i> Base Carregada';
            dbInfo.textContent = `${this.baseEspecificacoes.length} itens disponíveis`;
            viewDbBtn.disabled = false;
            statusDb.className = 'status-value status-success';
            statusDb.innerHTML = '✅ Carregada';
        } else {
            dbStatus.className = 'status-badge status-warning';
            dbStatus.innerHTML = '<i class="fas fa-exclamation-circle"></i> Base Não Carregada';
            dbInfo.textContent = 'Carregue a base de dados para começar';
            viewDbBtn.disabled = true;
            statusDb.className = 'status-value status-error';
            statusDb.innerHTML = '❌ Não carregada';
        }
    }

    updateBudgetStatus(loaded) {
        const statusBudget = document.getElementById('status-budget');
        const previewBtn = document.getElementById('preview-btn');
        
        if (loaded && this.budgetData.length > 0) {
            statusBudget.className = 'status-value status-success';
            statusBudget.innerHTML = `✅ ${this.budgetData.length} itens`;
            previewBtn.disabled = false;
        } else {
            statusBudget.className = 'status-value status-error';
            statusBudget.innerHTML = '❌ Não carregado';
            previewBtn.disabled = true;
        }
    }

    checkReadyToGenerate() {
        const generateBtn = document.getElementById('generate-btn');
        const statusReady = document.getElementById('status-ready');
        
        const isReady = this.baseEspecificacoes.length > 0 && this.budgetData.length > 0;
        
        if (isReady) {
            generateBtn.disabled = false;
            statusReady.className = 'status-value status-success';
            statusReady.innerHTML = '✅ Pronto';
        } else {
            generateBtn.disabled = true;
            statusReady.className = 'status-value status-error';
            statusReady.innerHTML = '❌ Aguardando';
        }
    }

    showBudgetPreview() {
        const previewCard = document.getElementById('budget-preview-card');
        const tableBody = document.querySelector('#budget-preview-table tbody');
        const itemCount = document.getElementById('budget-item-count');
        
        previewCard.classList.remove('hidden');
        tableBody.innerHTML = '';
        
        this.budgetData.slice(0, 20).forEach(item => {
            const tr = document.createElement('tr');
            
            ['Item', 'Código', 'Descrição', 'Und', 'Quant.', 'Banco'].forEach(field => {
                const td = document.createElement('td');
                td.textContent = item[field] || '';
                tr.appendChild(td);
            });
            
            tableBody.appendChild(tr);
        });
        
        itemCount.textContent = `${this.budgetData.length} itens`;
    }

    clearBudget() {
        this.budgetData = [];
        document.getElementById('budget-preview-card').classList.add('hidden');
        document.getElementById('budget-file-input').value = '';
        this.updateBudgetStatus(false);
        this.checkReadyToGenerate();
    }

    showDatabaseModal() {
        const modal = document.getElementById('db-modal');
        modal.classList.remove('hidden');
        this.updateDatabaseView();
    }

    switchTab(tabName) {
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.classList.remove('active');
        });
        
        document.querySelectorAll('.tab-content').forEach(content => {
            content.classList.remove('active');
        });

        document.querySelector(`.tab-btn[data-tab="${tabName}"]`).classList.add('active');
        document.getElementById(`tab-${tabName}`).classList.add('active');
        
        if (tabName === 'info') {
            this.updateDatabaseInfo();
        }
    }

    updateDatabaseView() {
        // Filtrar itens com base na busca
        this.filteredItems = this.baseEspecificacoes.filter(item => {
            if (!this.currentSearch) return true;
            
            const searchLower = this.currentSearch.toLowerCase();
            return (
                (item.COMPOSIÇÃO && item.COMPOSIÇÃO.toLowerCase().includes(searchLower)) ||
                (item.DESCRIÇÃO && item.DESCRIÇÃO.toLowerCase().includes(searchLower)) ||
                (item.Banco && item.Banco.toLowerCase().includes(searchLower)) ||
                (item['ESPECIFICAÇÃO TÉCNICA'] && item['ESPECIFICAÇÃO TÉCNICA'].toLowerCase().includes(searchLower))
            );
        });

        // Calcular paginação
        const totalPages = Math.ceil(this.filteredItems.length / this.itemsPerPage);
        const startIndex = (this.currentPage - 1) * this.itemsPerPage;
        const endIndex = startIndex + this.itemsPerPage;
        const pageItems = this.filteredItems.slice(startIndex, endIndex);

        // Atualizar tabela
        const tbody = document.getElementById('db-tbody');
        tbody.innerHTML = '';

        pageItems.forEach(item => {
            const tr = document.createElement('tr');
            
            const compTd = document.createElement('td');
            compTd.textContent = item.COMPOSIÇÃO || '';
            
            const bancoTd = document.createElement('td');
            bancoTd.textContent = item.Banco || '';
            
            const descTd = document.createElement('td');
            descTd.textContent = item.DESCRIÇÃO || '';
            
            const especTd = document.createElement('td');
            especTd.textContent = (item['ESPECIFICAÇÃO TÉCNICA'] || '').substring(0, 80) + '...';
            especTd.title = item['ESPECIFICAÇÃO TÉCNICA'] || '';
            
            const actionsTd = document.createElement('td');
            actionsTd.innerHTML = `
                <button class="btn btn-sm" onclick="gerador.useSpecification('${item.COMPOSIÇÃO}')">
                    <i class="fas fa-eye"></i>
                </button>
            `;
            
            tr.appendChild(compTd);
            tr.appendChild(bancoTd);
            tr.appendChild(descTd);
            tr.appendChild(especTd);
            tr.appendChild(actionsTd);
            
            tbody.appendChild(tr);
        });

        // Atualizar controles de paginação
        document.getElementById('db-result-count').textContent = 
            `${this.filteredItems.length} itens encontrados`;
        
        document.getElementById('page-info').textContent = 
            `Página ${this.currentPage} de ${totalPages}`;
        
        document.getElementById('prev-page').disabled = this.currentPage <= 1;
        document.getElementById('next-page').disabled = this.currentPage >= totalPages;
    }

    updateDatabaseInfo() {
        document.getElementById('info-total-items').textContent = this.baseEspecificacoes.length;
        
        const banks = new Set(this.baseEspecificacoes.map(item => item.Banco).filter(Boolean));
        document.getElementById('info-banks').textContent = banks.size;
        
        const lastUpdate = localStorage.getItem('db_last_updated');
        if (lastUpdate) {
            const date = new Date(lastUpdate);
            document.getElementById('info-last-update').textContent = 
                date.toLocaleDateString('pt-BR');
        }
        
        // Calcular tamanho aproximado
        const dbSize = JSON.stringify(this.baseEspecificacoes).length;
        document.getElementById('info-db-size').textContent = 
            `${(dbSize / 1024).toFixed(2)} KB`;
    }

    useSpecification(composicao) {
        const spec = this.baseEspecificacoes.find(item => item.COMPOSIÇÃO === composicao);
        if (spec) {
            alert(`Especificação: ${spec['ESPECIFICAÇÃO TÉCNICA'].substring(0, 200)}...`);
        }
    }

    async generatePreview() {
        if (this.budgetData.length === 0) {
            alert('Carregue um orçamento primeiro.');
            return;
        }

        const container = document.getElementById('spec-preview-container');
        container.innerHTML = '<div class="loading">Gerando prévia...</div>';
        
        document.getElementById('spec-preview-card').classList.remove('hidden');

        try {
            const previewItems = this.budgetData.slice(0, 5);
            let html = '';
            
            previewItems.forEach((item, index) => {
                const spec = this.getSpecification(item);
                html += `
                    <div class="spec-item">
                        <h4>${item.Item} - ${item.Descrição}</h4>
                        <div class="spec-content">
                            ${spec.split('\n').map(line => `<p>${line}</p>`).join('')}
                        </div>
                        <div class="spec-meta">
                            <span>Código: ${item.Código || 'N/A'}</span>
                            <span>Quantidade: ${item['Quant.']} ${item.Und}</span>
                        </div>
                    </div>
                `;
                
                if (index < previewItems.length - 1) {
                    html += '<hr>';
                }
            });
            
            container.innerHTML = html;
        } catch (error) {
            console.error('Erro ao gerar prévia:', error);
            container.innerHTML = '<p class="error">Erro ao gerar prévia.</p>';
        }
    }

    getSpecification(item) {
        // Primeiro tenta buscar da base de dados
        let especificacao = this.getSpecificationFromDatabase(item);
        
        if (!especificacao) {
            // Se não encontrou, gera baseado na descrição
            especificacao = this.generateSpecificationFromDescription(item);
        }
        
        return especificacao;
    }

    getSpecificationFromDatabase(item) {
        // Buscar por código
        if (item.Código) {
            const found = this.baseEspecificacoes.find(dbItem => 
                dbItem.COMPOSIÇÃO.toString() === item.Código.toString()
            );
            if (found) return found['ESPECIFICAÇÃO TÉCNICA'];
        }

        // Buscar por descrição
        if (item.Descrição) {
            const descClean = item.Descrição.split('(')[0].split('-')[0].split('AF_')[0].trim();
            const found = this.baseEspecificacoes.find(dbItem => 
                dbItem.DESCRIÇÃO && dbItem.DESCRIÇÃO.toLowerCase().includes(descClean.toLowerCase())
            );
            if (found) return found['ESPECIFICAÇÃO TÉCNICA'];
        }

        return null;
    }

    generateSpecificationFromDescription(item) {
        const descUpper = item.Descrição.toUpperCase();
        
        if (descUpper.includes('CABO') && descUpper.includes('COBRE')) {
            return this.generateCableSpecification(item);
        } else if (descUpper.includes('DISJUNTOR')) {
            return this.generateBreakerSpecification(item);
        } else if (descUpper.includes('QUADRO') && descUpper.includes('ENERGIA')) {
            return this.generatePanelSpecification(item);
        } else if (descUpper.includes('TOMADA')) {
            return this.generateOutletSpecification(item);
        } else if (descUpper.includes('INTERRUPTOR')) {
            return this.generateSwitchSpecification(item);
        } else {
            return this.generateDefaultSpecification(item);
        }
    }

    generateCableSpecification(item) {
        return `CABO DE COBRE - Deve ser cabo flexível de cobre eletrolítico, isolamento em PVC 750V/1000V, conforme NBR 7286. 
        Bitola conforme especificado em projeto. Todos os materiais devem ser novos e de primeira qualidade. 
        Deve possuir certificado de garantia e laudos técnicos quando aplicável. 
        Quantidade: ${item['Quant.']} ${item.Und}.`;
    }

    generateBreakerSpecification(item) {
        return `DISJUNTOR - Disjuntor termomagnético DIN, curva C, conforme NBR NM 60898. 
        Deve possuir certificado INMETRO. Tensão nominal 127/220V - 60Hz. 
        Todos os materiais devem ser novos e de primeira qualidade. 
        Quantidade: ${item['Quant.']} ${item.Und}.`;
    }

    generatePanelSpecification(item) {
        return `QUADRO DE DISTRIBUIÇÃO - Quadro em chapa de aço galvanizado, pintura epóxi eletrostática. 
        Conforme NBR IEC 61439-1. Deve incluir barramentos em cobre, trilhos DIN e sistema de aterramento. 
        Todos os materiais devem ser novos e de primeira qualidade. 
        Quantidade: ${item['Quant.']} ${item.Und}.`;
    }

    generateOutletSpecification(item) {
        return `TOMADA ELÉTRICA - Tomada 2P+T 10A, material policarbonato autoextinguível. 
        Conforme NBR 14136. Bornes de aperto torque controlado. 
        Todos os materiais devem ser novos e de primeira qualidade. 
        Quantidade: ${item['Quant.']} ${item.Und}.`;
    }

    generateSwitchSpecification(item) {
        return `INTERRUPTOR - Interruptor simples 10A, material policarbonato autoextinguível. 
        Conforme NBR 14136. Deve suportar no mínimo 40.000 ciclos de acionamento. 
        Todos os materiais devem ser novos e de primeira qualidade. 
        Quantidade: ${item['Quant.']} ${item.Und}.`;
    }

    generateDefaultSpecification(item) {
        return `O item ${item.Descrição} deve ser fornecido e instalado conforme especificações técnicas do fabricante 
        e normas técnicas aplicáveis, em especial a NBR 5410. Deve possuir certificado de garantia e laudos técnicos quando aplicável. 
        Todos os materiais devem ser novos e de primeira qualidade. 
        Quantidade: ${item['Quant.']} ${item.Und}.`;
    }

    async generateDocument() {
        if (!this.budgetData.length) {
            alert('Por favor, carregue um orçamento primeiro.');
            return;
        }

        const format = document.querySelector('input[name="format"]:checked').value;
        const projectName = document.getElementById('project-name').value || 'Projeto Elétrico';
        const projectCode = document.getElementById('project-code').value || '';
        const clientName = document.getElementById('client-name').value || '';
        const date = document.getElementById('project-date').value;
        const formattedDate = date ? new Date(date).toLocaleDateString('pt-BR') : new Date().toLocaleDateString('pt-BR');

        // Mostrar progresso
        this.showProgress(true);

        try {
            let content, fileName;
            
            if (format === 'docx') {
                // Gerar documento Word
                fileName = `Especificacoes_Tecnicas_${projectName.replace(/\s+/g, '_')}.docx`;
                content = await this.generateWordDocument(projectName, projectCode, clientName, formattedDate);
            } else {
                // Gerar arquivo TXT
                fileName = `Especificacoes_Tecnicas_${projectName.replace(/\s+/g, '_')}.txt`;
                content = this.generateTextDocument(projectName, projectCode, clientName, formattedDate);
            }

            // Salvar conteúdo gerado
            this.generatedContent = content;
            this.generatedFileName = fileName;

            // Mostrar seção de download
            this.showDownloadSection(fileName);

        } catch (error) {
            console.error('Erro ao gerar documento:', error);
            alert('Erro ao gerar documento. Tente novamente.');
        } finally {
            this.showProgress(false);
        }
    }

    generateTextDocument(projectName, projectCode, clientName, date) {
        let content = '='.repeat(80) + '\n';
        content += 'ESPECIFICAÇÕES TÉCNICAS\n';
        content += '='.repeat(80) + '\n\n';
        
        if (projectCode) content += `Código do Projeto: ${projectCode}\n`;
        content += `Projeto: ${projectName}\n`;
        if (clientName) content += `Cliente: ${clientName}\n`;
        content += `Data de emissão: ${date}\n`;
        content += '-'.repeat(80) + '\n\n';
        
        content += `Este documento contém as especificações técnicas dos materiais e equipamentos `;
        content += `previstos no orçamento, baseadas nas descrições dos itens e normas técnicas aplicáveis.\n\n`;
        content += '-'.repeat(80) + '\n\n';

        // Organizar itens por hierarquia
        const organizedItems = this.organizeItemsByHierarchy();

        organizedItems.forEach((item, index) => {
            const nivel = item.nivel || 0;
            const indent = '  '.repeat(nivel * 2);
            
            content += `${indent}${item.Item} - ${item.Descrição}\n`;
            
            if (this.isLowestLevelItem(item, organizedItems)) {
                content += `${indent}Código: ${item.Código || 'Não informado'} | `;
                content += `Banco: ${item.Banco || 'Não informado'} | `;
                content += `Quantidade: ${item['Quant.'] || 'N/A'} ${item.Und || 'N/A'}\n`;
                
                const especificacao = this.getSpecification(item);
                content += `${indent}${especificacao.replace(/\n/g, '\n' + indent)}\n\n`;
            }
        });

        content += '\n' + '='.repeat(80) + '\n';
        content += 'FIM DO DOCUMENTO\n';
        content += '='.repeat(80) + '\n';

        return content;
    }

    async generateWordDocument(projectName, projectCode, clientName, date) {
        // Nota: Esta é uma implementação simplificada
        // Para uma implementação completa, use a biblioteca docx
        return this.generateTextDocument(projectName, projectCode, clientName, date);
    }

    organizeItemsByHierarchy() {
        return this.budgetData.map(item => {
            const itemStr = item.Item.toString();
            const nivel = itemStr.split('.').length - 1;
            return { ...item, nivel };
        }).sort((a, b) => {
            const aParts = a.Item.split('.').map(Number);
            const bParts = b.Item.split('.').map(Number);
            
            for (let i = 0; i < Math.max(aParts.length, bParts.length); i++) {
                const aVal = aParts[i] || 0;
                const bVal = bParts[i] || 0;
                if (aVal !== bVal) return aVal - bVal;
            }
            return 0;
        });
    }

    isLowestLevelItem(item, items) {
        const itemPrefix = item.Item + '.';
        return !items.some(other => other.Item.startsWith(itemPrefix));
    }

    showProgress(show) {
        const progressContainer = document.getElementById('progress-container');
        const generateBtn = document.getElementById('generate-btn');
        
        if (show) {
            progressContainer.classList.remove('hidden');
            generateBtn.disabled = true;
            
            // Simular progresso
            let progress = 0;
            const interval = setInterval(() => {
                progress += 10;
                document.getElementById('progress-fill').style.width = `${progress}%`;
                document.getElementById('progress-percent').textContent = `${progress}%`;
                
                if (progress >= 100) {
                    clearInterval(interval);
                }
            }, 200);
        } else {
            progressContainer.classList.add('hidden');
            generateBtn.disabled = false;
        }
    }

    showDownloadSection(fileName) {
        const downloadSection = document.getElementById('download-section');
        const generatedFileName = document.getElementById('generated-file-name');
        const downloadStats = document.getElementById('download-stats');
        
        downloadSection.classList.remove('hidden');
        generatedFileName.textContent = `Arquivo: ${fileName}`;
        downloadStats.textContent = `${this.budgetData.length} itens processados`;
    }

    downloadGeneratedDocument() {
        if (!this.generatedContent || !this.generatedFileName) return;
        
        const blob = new Blob([this.generatedContent], { 
            type: this.generatedFileName.endsWith('.docx') ? 
                'application/vnd.openxmlformats-officedocument.wordprocessingml.document' : 
                'text/plain;charset=utf-8' 
        });
        
        saveAs(blob, this.generatedFileName);
    }

    exportDatabase() {
        const workbook = XLSX.utils.book_new();
        const worksheet = XLSX.utils.json_to_sheet(this.baseEspecificacoes);
        XLSX.utils.book_append_sheet(workbook, worksheet, "Base de Dados");
        
        const fileName = `BASE_DE_DADOS_ESPECIFICACAO_TECNICA_${new Date().toISOString().split('T')[0]}.xlsx`;
        XLSX.writeFile(workbook, fileName);
    }
}

// Instanciar o gerador quando o DOM estiver carregado
let gerador;
document.addEventListener('DOMContentLoaded', () => {
    gerador = new GeradorEspecificacoes();
});
