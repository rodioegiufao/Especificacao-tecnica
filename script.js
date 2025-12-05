class GeradorEspecificacoes {
    constructor() {
        this.baseEspecificacoes = [];
        this.currentFile = null;
        this.budgetData = [];
        this.initializeEventListeners();
        this.setCurrentDate();
        this.loadLocalDatabase();
    }

    initializeEventListeners() {
        // File upload
        const fileInput = document.getElementById('file-input');
        const browseBtn = document.getElementById('browse-btn');
        const dropArea = document.getElementById('drop-area');
        const removeFileBtn = document.getElementById('remove-file');
        const generateBtn = document.getElementById('generate-btn');
        const viewDbBtn = document.getElementById('view-db-btn');
        const closeDbModal = document.getElementById('close-db-modal');
        const dbModal = document.getElementById('db-modal');
        const tabBtns = document.querySelectorAll('.tab-btn');
        const downloadTemplate = document.getElementById('download-template');

        browseBtn.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', (e) => this.handleFileSelect(e));
        
        removeFileBtn.addEventListener('click', () => this.removeFile());

        // Drag and drop
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropArea.addEventListener(eventName, preventDefaults, false);
        });

        function preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }

        ['dragenter', 'dragover'].forEach(eventName => {
            dropArea.addEventListener(eventName, highlight, false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            dropArea.addEventListener(eventName, unhighlight, false);
        });

        function highlight() {
            dropArea.classList.add('drag-over');
        }

        function unhighlight() {
            dropArea.classList.remove('drag-over');
        }

        dropArea.addEventListener('drop', (e) => this.handleDrop(e));

        // Generate document
        generateBtn.addEventListener('click', () => this.generateDocument());

        // Database modal
        viewDbBtn.addEventListener('click', () => this.showDatabaseModal());
        closeDbModal.addEventListener('click', () => dbModal.classList.add('hidden'));

        // Tab switching
        tabBtns.forEach(btn => {
            btn.addEventListener('click', () => this.switchTab(btn.dataset.tab));
        });

        // Download template
        downloadTemplate.addEventListener('click', (e) => {
            e.preventDefault();
            this.downloadTemplate();
        });

        // Preview toggle
        const togglePreview = document.getElementById('toggle-preview');
        togglePreview.addEventListener('click', () => this.togglePreview());
    }

    setCurrentDate() {
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('project-date').value = today;
    }

    async loadLocalDatabase() {
        try {
            // Tentar carregar do localStorage
            const savedDb = localStorage.getItem('especificacoes_db');
            if (savedDb) {
                this.baseEspecificacoes = JSON.parse(savedDb);
                this.updateDatabaseView();
                return;
            }

            // Se não existir, criar uma base de dados de exemplo
            this.baseEspecificacoes = [
                {
                    "COMPOSIÇÃO": "101",
                    "Banco": "SINAPI",
                    "DESCRIÇÃO": "CABO DE COBRE ISOLADO PVC 750V - 2,5 MM²",
                    "ESPECIFICAÇÃO TÉCNICA": "Cabo flexível de cobre, isolamento em PVC 750V/1000V, seção 2,5mm², coloração conforme NBR..."
                },
                {
                    "COMPOSIÇÃO": "102",
                    "Banco": "SINAPI",
                    "DESCRIÇÃO": "DISJUNTOR TERMOMAGNÉTICO MONOPOLAR 10A CURVA C",
                    "ESPECIFICAÇÃO TÉCNICA": "Disjuntor termomagnético DIN, monopolar, corrente nominal 10A, curva C, tensão 127/220V..."
                },
                {
                    "COMPOSIÇÃO": "103",
                    "Banco": "SINAPI",
                    "DESCRIÇÃO": "TOMADA 2P+T 10A SOBREPOR",
                    "ESPECIFICAÇÃO TÉCNICA": "Tomada 2P+T 10A, tipo sobrepor, material policarbonato autoextinguível, bornes de aperto..."
                }
            ];

            localStorage.setItem('especificacoes_db', JSON.stringify(this.baseEspecificacoes));
            this.updateDatabaseView();
        } catch (error) {
            console.error('Erro ao carregar base de dados:', error);
        }
    }

    updateDatabaseView() {
        const tbody = document.getElementById('db-tbody');
        tbody.innerHTML = '';

        this.baseEspecificacoes.forEach(item => {
            const tr = document.createElement('tr');
            
            const compTd = document.createElement('td');
            compTd.textContent = item.COMPOSIÇÃO;
            
            const bancoTd = document.createElement('td');
            bancoTd.textContent = item.Banco;
            
            const descTd = document.createElement('td');
            descTd.textContent = item.DESCRIÇÃO;
            
            const especTd = document.createElement('td');
            especTd.textContent = item['ESPECIFICAÇÃO TÉCNICA'].substring(0, 100) + '...';
            especTd.title = item['ESPECIFICAÇÃO TÉCNICA'];
            
            tr.appendChild(compTd);
            tr.appendChild(bancoTd);
            tr.appendChild(descTd);
            tr.appendChild(especTd);
            
            tbody.appendChild(tr);
        });
    }

    handleFileSelect(event) {
        const file = event.target.files[0];
        this.processFile(file);
    }

    handleDrop(event) {
        const dt = event.dataTransfer;
        const file = dt.files[0];
        this.processFile(file);
    }

    async processFile(file) {
        if (!file || !file.name.match(/\.(xlsx|xls)$/i)) {
            alert('Por favor, selecione um arquivo Excel (.xlsx ou .xls)');
            return;
        }

        this.currentFile = file;

        // Mostrar informações do arquivo
        document.getElementById('file-info').classList.remove('hidden');
        document.getElementById('file-name').textContent = file.name;
        document.getElementById('file-size').textContent = this.formatFileSize(file.size);

        // Habilitar botão de gerar
        document.getElementById('generate-btn').disabled = false;

        // Ler o arquivo Excel
        try {
            const data = await this.readExcelFile(file);
            this.budgetData = data;
            this.showDataPreview(data);
        } catch (error) {
            console.error('Erro ao processar arquivo:', error);
            alert('Erro ao processar o arquivo. Verifique se é um Excel válido.');
        }
    }

    async readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // Pegar a primeira planilha
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    
                    // Converter para JSON
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    // Processar os dados
                    const processedData = this.processExcelData(jsonData);
                    resolve(processedData);
                } catch (error) {
                    reject(error);
                }
            };
            
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    processExcelData(data) {
        // Encontrar o cabeçalho
        let headerRow = 0;
        for (let i = 0; i < Math.min(10, data.length); i++) {
            const row = data[i];
            if (row && row.some(cell => 
                cell && typeof cell === 'string' && 
                (cell.includes('Item') || cell.includes('Descrição'))
            )) {
                headerRow = i;
                break;
            }
        }

        // Extrair cabeçalhos
        const headers = data[headerRow].map(h => h ? h.toString().trim() : '');

        // Mapear colunas
        const columnMapping = {
            'Item': headers.findIndex(h => h.includes('Item')),
            'Código': headers.findIndex(h => h.includes('Código') || h.includes('Codigo')),
            'Descrição': headers.findIndex(h => h.includes('Descrição') || h.includes('Descricao')),
            'Und': headers.findIndex(h => h.includes('Und') || h.includes('UNID')),
            'Quant.': headers.findIndex(h => h.includes('Quant') || h.includes('QTDE')),
            'Banco': headers.findIndex(h => h.includes('Banco'))
        };

        // Processar linhas de dados
        const processed = [];
        for (let i = headerRow + 1; i < data.length; i++) {
            const row = data[i];
            if (!row || row.every(cell => !cell)) continue;

            const item = {
                'Item': columnMapping['Item'] >= 0 ? this.safeGetCell(row, columnMapping['Item']) : '',
                'Código': columnMapping['Código'] >= 0 ? this.safeGetCell(row, columnMapping['Código']) : '',
                'Descrição': columnMapping['Descrição'] >= 0 ? this.safeGetCell(row, columnMapping['Descrição']) : '',
                'Und': columnMapping['Und'] >= 0 ? this.safeGetCell(row, columnMapping['Und']) : '',
                'Quant.': columnMapping['Quant.'] >= 0 ? this.safeGetCell(row, columnMapping['Quant.']) : '',
                'Banco': columnMapping['Banco'] >= 0 ? this.safeGetCell(row, columnMapping['Banco']) : ''
            };

            // Filtrar linhas vazias ou totais
            if (item.Item && !item.Item.toString().includes('Total')) {
                processed.push(item);
            }
        }

        return processed;
    }

    safeGetCell(row, index) {
        return row[index] !== undefined && row[index] !== null ? row[index].toString().trim() : '';
    }

    showDataPreview(data) {
        const previewSection = document.getElementById('preview-section');
        const tbody = document.querySelector('#data-preview tbody');
        
        previewSection.classList.remove('hidden');
        tbody.innerHTML = '';

        // Mostrar apenas as primeiras 10 linhas
        const displayData = data.slice(0, 10);
        
        displayData.forEach(item => {
            const tr = document.createElement('tr');
            
            Object.values(item).forEach(value => {
                const td = document.createElement('td');
                td.textContent = value || '';
                tr.appendChild(td);
            });
            
            tbody.appendChild(tr);
        });

        // Atualizar contador
        document.getElementById('item-count').textContent = `${data.length} itens carregados`;
    }

    removeFile() {
        this.currentFile = null;
        this.budgetData = [];
        
        document.getElementById('file-info').classList.add('hidden');
        document.getElementById('preview-section').classList.add('hidden');
        document.getElementById('generate-btn').disabled = true;
        document.getElementById('file-input').value = '';
    }

    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    togglePreview() {
        const previewSection = document.getElementById('preview-section');
        const toggleBtn = document.getElementById('toggle-preview');
        
        if (previewSection.classList.contains('expanded')) {
            previewSection.classList.remove('expanded');
            toggleBtn.innerHTML = '<i class="fas fa-expand-alt"></i>';
        } else {
            previewSection.classList.add('expanded');
            toggleBtn.innerHTML = '<i class="fas fa-compress-alt"></i>';
        }
    }

    showDatabaseModal() {
        const modal = document.getElementById('db-modal');
        modal.classList.remove('hidden');
        this.updateDatabaseView();
    }

    switchTab(tabName) {
        // Remover classe active de todas as tabs
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.classList.remove('active');
        });
        
        document.querySelectorAll('.tab-content').forEach(content => {
            content.classList.remove('active');
        });

        // Ativar tab selecionada
        document.querySelector(`.tab-btn[data-tab="${tabName}"]`).classList.add('active');
        document.getElementById(`tab-${tabName}`).classList.add('active');
    }

    async generateDocument() {
        if (!this.budgetData.length) {
            alert('Por favor, carregue um orçamento primeiro.');
            return;
        }

        const projectName = document.getElementById('project-name').value || 'Projeto Elétrico';
        const generateBtn = document.getElementById('generate-btn');
        const progressContainer = document.getElementById('progress-container');
        const progressFill = document.getElementById('progress-fill');
        const progressText = document.getElementById('progress-text');

        // Mostrar progresso
        generateBtn.disabled = true;
        progressContainer.classList.remove('hidden');

        try {
            // Simular progresso
            for (let i = 0; i <= 100; i += 10) {
                await this.sleep(100);
                progressFill.style.width = `${i}%`;
                progressText.textContent = `${i}%`;
            }

            // Gerar conteúdo do documento
            const docContent = this.generateDocumentContent(projectName);

            // Criar e baixar o documento
            const fileName = `Especificacoes_Tecnicas_${projectName.replace(/\s+/g, '_')}.docx`;
            this.createAndDownloadDoc(docContent, fileName);

            // Mostrar seção de download
            this.showDownloadSection(fileName);

        } catch (error) {
            console.error('Erro ao gerar documento:', error);
            alert('Erro ao gerar documento. Tente novamente.');
        } finally {
            generateBtn.disabled = false;
        }
    }

    generateDocumentContent(projectName) {
        const date = document.getElementById('project-date').value;
        const formattedDate = date ? new Date(date).toLocaleDateString('pt-BR') : new Date().toLocaleDateString('pt-BR');

        let content = `ESPECIFICAÇÕES TÉCNICAS\n\n`;
        content += `Projeto: ${projectName}\n`;
        content += `Data de emissão: ${formattedDate}\n\n`;
        content += `Este documento contém as especificações técnicas dos materiais e equipamentos `;
        content += `previstos no orçamento, baseadas nas descrições dos itens e normas técnicas aplicáveis.\n\n`;

        // Organizar itens por hierarquia
        const organizedItems = this.organizeItemsByHierarchy(this.budgetData);

        organizedItems.forEach((item, index) => {
            const nivel = item.nivel || 0;
            const indent = '  '.repeat(nivel);
            
            content += `\n${indent}${item.Item} - ${item.Descrição}\n`;
            
            // Adicionar especificação apenas para itens de nível mais baixo
            if (this.isLowestLevelItem(item, organizedItems)) {
                content += `${indent}Código: ${item.Código || 'Não informado'} | `;
                content += `Banco: ${item.Banco || 'Não informado'} | `;
                content += `Quantidade: ${item['Quant.'] || 'N/A'} ${item.Und || 'N/A'}\n`;
                
                const especificacao = this.getSpecification(item);
                content += `${indent}${especificacao}\n`;
            }
        });

        return content;
    }

    organizeItemsByHierarchy(data) {
        // Adicionar nível hierárquico baseado no número de pontos
        const itemsWithLevel = data.map(item => {
            const itemStr = item.Item.toString();
            const nivel = itemStr.split('.').length - 1;
            return { ...item, nivel };
        });

        // Ordenar pela hierarquia
        itemsWithLevel.sort((a, b) => {
            const aParts = a.Item.split('.').map(Number);
            const bParts = b.Item.split('.').map(Number);
            
            for (let i = 0; i < Math.max(aParts.length, bParts.length); i++) {
                const aVal = aParts[i] || 0;
                const bVal = bParts[i] || 0;
                if (aVal !== bVal) return aVal - bVal;
            }
            return 0;
        });

        return itemsWithLevel;
    }

    isLowestLevelItem(item, items) {
        const itemPrefix = item.Item + '.';
        return !items.some(other => other.Item.startsWith(itemPrefix));
    }

    getSpecification(item) {
        // Primeiro tenta buscar da base de dados
        const fromDb = this.getSpecificationFromDatabase(item);
        if (fromDb) return fromDb;

        // Se não encontrou, gera baseado na descrição
        const descricao = item.Descrição.toUpperCase();
        
        if (descricao.includes('CABO')) {
            return this.generateCableSpecification(item);
        } else if (descricao.includes('DISJUNTOR')) {
            return this.generateBreakerSpecification(item);
        } else if (descricao.includes('QUADRO')) {
            return this.generatePanelSpecification(item);
        } else if (descricao.includes('TOMADA')) {
            return this.generateOutletSpecification(item);
        } else if (descricao.includes('INTERRUPTOR')) {
            return this.generateSwitchSpecification(item);
        } else {
            return this.generateDefaultSpecification(item);
        }
    }

    getSpecificationFromDatabase(item) {
        // Buscar por código
        if (item.Código) {
            const found = this.baseEspecificacoes.find(dbItem => 
                dbItem.COMPOSIÇÃO === item.Código.toString()
            );
            if (found) return found['ESPECIFICAÇÃO TÉCNICA'];
        }

        // Buscar por descrição
        if (item.Descrição) {
            const descClean = item.Descrição.split('(')[0].split('-')[0].trim();
            const found = this.baseEspecificacoes.find(dbItem => 
                dbItem.DESCRIÇÃO.toLowerCase().includes(descClean.toLowerCase())
            );
            if (found) return found['ESPECIFICAÇÃO TÉCNICA'];
        }

        return null;
    }

    generateCableSpecification(item) {
        return `Cabo de cobre flexível, isolamento em PVC 750V/1000V, conforme NBR 7286. 
        Deve possuir certificado de garantia e laudos técnicos. 
        Quantidade: ${item['Quant.']} ${item.Und}.`;
    }

    generateBreakerSpecification(item) {
        return `Disjuntor termomagnético DIN, conforme NBR NM 60898. 
        Deve possuir certificado INMETRO. 
        Quantidade: ${item['Quant.']} ${item.Und}.`;
    }

    generatePanelSpecification(item) {
        return `Quadro de distribuição em chapa de aço galvanizado, pintura epóxi. 
        Conforme NBR IEC 61439-1. 
        Quantidade: ${item['Quant.']} ${item.Und}.`;
    }

    generateOutletSpecification(item) {
        return `Tomada 2P+T 10A, material policarbonato autoextinguível. 
        Conforme NBR 14136. 
        Quantidade: ${item['Quant.']} ${item.Und}.`;
    }

    generateSwitchSpecification(item) {
        return `Interruptor simples 10A, material policarbonato autoextinguível. 
        Conforme NBR 14136. 
        Quantidade: ${item['Quant.']} ${item.Und}.`;
    }

    generateDefaultSpecification(item) {
        return `O item ${item.Descrição} deve ser fornecido e instalado conforme especificações 
        técnicas do fabricante e normas técnicas aplicáveis, em especial a NBR 5410. 
        Deve possuir certificado de garantia. 
        Quantidade: ${item['Quant.']} ${item.Und}.`;
    }

    createAndDownloadDoc(content, fileName) {
        // Para uma implementação real, você precisaria de uma biblioteca
        // como docx.js ou fazer uma requisição para um backend que gera o DOCX
        // Esta é uma versão simplificada que cria um arquivo TXT
        
        const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
        saveAs(blob, fileName);
    }

    showDownloadSection(fileName) {
        const downloadSection = document.getElementById('download-section');
        const generatedFileName = document.getElementById('generated-file-name');
        const downloadBtn = document.getElementById('download-btn');
        const progressContainer = document.getElementById('progress-container');

        // Esconder barra de progresso
        progressContainer.classList.add('hidden');

        // Mostrar seção de download
        downloadSection.classList.remove('hidden');
        generatedFileName.textContent = `Arquivo: ${fileName}`;

        // Configurar botão de download
        downloadBtn.onclick = () => {
            const content = this.generateDocumentContent(
                document.getElementById('project-name').value || 'Projeto Elétrico'
            );
            const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
            saveAs(blob, fileName);
        };
    }

    downloadTemplate() {
        // Criar dados de exemplo para o template
        const templateData = [
            ['Item', 'Código', 'Descrição', 'Und', 'Quant.', 'Banco'],
            ['1', '101', 'CABO DE COBRE ISOLADO PVC 750V - 2,5 MM²', 'm', '100', 'SINAPI'],
            ['1.1', '102', 'DISJUNTOR TERMOMAGNÉTICO MONOPOLAR 10A CURVA C', 'un', '5', 'SINAPI'],
            ['1.2', '103', 'TOMADA 2P+T 10A SOBREPOR', 'un', '20', 'SINAPI'],
            ['2', '201', 'QUADRO DE DISTRIBUIÇÃO 12 DISJUNTORES', 'un', '2', 'SINAPI'],
            ['2.1', '202', 'INTERRUPTOR SIMPLES 10A', 'un', '15', 'SINAPI']
        ];

        // Criar worksheet
        const ws = XLSX.utils.aoa_to_sheet(templateData);
        
        // Criar workbook
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Orçamento");

        // Gerar arquivo
        XLSX.writeFile(wb, "Template_Orcamento.xlsx");
    }

    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}

// Inicializar aplicação quando o DOM estiver carregado
document.addEventListener('DOMContentLoaded', () => {
    new GeradorEspecificacoes();
});