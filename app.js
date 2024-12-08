class Task {
    constructor(name, resource, connectedTask, duration, startDate, optimistic = null, pessimistic = null) {
        this.name = name;
        this.resource = resource;
        this.connectedTask = connectedTask;
        this.mostLikelyDuration = duration;
        this.startDate = startDate;
        this.optimisticDuration = optimistic || Math.max(1, Math.floor(duration * 0.6));
        this.pessimisticDuration = pessimistic || Math.ceil(duration * 1.6);
        this.updatePertEstimate();
    }

    updatePertEstimate() {
        this.pertEstimate = (this.optimisticDuration + 4 * this.mostLikelyDuration + this.pessimisticDuration) / 6;
    }

    calculateFinishDate(startDate) {
        const date = new Date(startDate);
        date.setDate(date.getDate() + this.mostLikelyDuration);
        return date;
    }

    updateFromEditable(field, value) {
        switch(field) {
            case 'name':
                this.name = value;
                break;
            case 'resource':
                this.resource = value;
                break;
            case 'connectedTask':
                this.connectedTask = value;
                break;
            case 'startDate':
                this.startDate = value;
                break;
            case 'optimistic':
                this.optimisticDuration = parseInt(value) || this.optimisticDuration;
                break;
            case 'mostLikely':
                this.mostLikelyDuration = parseInt(value) || this.mostLikelyDuration;
                break;
            case 'pessimistic':
                this.pessimisticDuration = parseInt(value) || this.pessimisticDuration;
                break;
        }
        this.updatePertEstimate();
    }
}

class ProjectSimulator {
    constructor() {
        this.tasks = [];
        this.setupEventListeners();
        this.simulationResults = [];
    }

    setupEventListeners() {
        // Import Excel button
        document.getElementById('importExcel').addEventListener('click', () => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.xlsx, .xls';
            input.onchange = (e) => this.handleExcelImport(e);
            input.click();
        });

        // Download Template button
        document.getElementById('downloadTemplate').addEventListener('click', () => {
            this.downloadTemplate();
        });

        // Task Form submission
        document.getElementById('taskForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.addTask();
        });

        // Run Simulation button
        document.getElementById('runSimulation').addEventListener('click', () => {
            this.runSimulation();
        });

        // Refresh Durations button
        document.getElementById('refreshDurations').addEventListener('click', () => {
            this.refreshDurations();
        });

        // Export to Excel button
        document.getElementById('exportToExcel').addEventListener('click', () => {
            this.exportToExcel();
        });
    }

    addTask() {
        const name = document.getElementById('taskName').value;
        const resource = document.getElementById('resourceName').value;
        const connectedTask = document.getElementById('connectedTask').value;
        const duration = parseInt(document.getElementById('duration').value);
        const startDate = document.getElementById('startDate').value;

        if (!name || !resource || !duration || !startDate) {
            alert('Please fill in all required fields');
            return;
        }

        const task = new Task(name, resource, connectedTask, duration, startDate);
        this.tasks.push(task);
        this.updateTaskTable();
        this.updateConnectedTaskDropdown();
        
        // Reset form
        document.getElementById('taskForm').reset();
    }

    updateTaskTable() {
        const tbody = document.getElementById('taskTableBody');
        tbody.innerHTML = '';

        this.tasks.forEach((task, index) => {
            const finishDate = task.calculateFinishDate(task.startDate);
            const row = document.createElement('tr');
            row.innerHTML = `
                <td><input type="text" class="form-control form-control-sm" data-field="name" value="${task.name}"></td>
                <td><input type="text" class="form-control form-control-sm" data-field="resource" value="${task.resource}"></td>
                <td>
                    <select class="form-control form-control-sm" data-field="connectedTask">
                        <option value="None" ${task.connectedTask === 'None' ? 'selected' : ''}>None</option>
                        ${this.tasks.map(t => t.name !== task.name ? 
                            `<option value="${t.name}" ${task.connectedTask === t.name ? 'selected' : ''}>${t.name}</option>` : 
                            '').join('')}
                    </select>
                </td>
                <td><input type="date" class="form-control form-control-sm" data-field="startDate" value="${task.startDate}"></td>
                <td><input type="text" class="form-control form-control-sm" readonly value="${finishDate.toISOString().split('T')[0]}"></td>
                <td><input type="number" class="form-control form-control-sm" data-field="optimistic" value="${task.optimisticDuration}" min="1"></td>
                <td><input type="number" class="form-control form-control-sm" data-field="mostLikely" value="${task.mostLikelyDuration}" min="1"></td>
                <td><input type="number" class="form-control form-control-sm" data-field="pessimistic" value="${task.pessimisticDuration}" min="1"></td>
                <td>${task.pertEstimate.toFixed(1)}</td>
                <td>
                    <button class="btn btn-danger btn-sm" onclick="window.projectSimulator.deleteTask(${index})">
                        <i class="bi bi-trash"></i> Delete
                    </button>
                </td>
            `;

            // Add event listeners for all input fields
            row.querySelectorAll('input, select').forEach(input => {
                input.addEventListener('change', (e) => {
                    const field = e.target.dataset.field;
                    const value = e.target.value;
                    task.updateFromEditable(field, value);
                    this.updateTaskTable();
                    this.updateConnectedTaskDropdown();
                });
            });

            tbody.appendChild(row);
        });
    }

    updateConnectedTaskDropdown() {
        const select = document.getElementById('connectedTask');
        select.innerHTML = '<option value="None">None</option>';
        this.tasks.forEach(task => {
            const option = document.createElement('option');
            option.value = task.name;
            option.textContent = task.name;
            select.appendChild(option);
        });
    }

    deleteTask(index) {
        this.tasks.splice(index, 1);
        this.updateTaskTable();
        this.updateConnectedTaskDropdown();
    }

    getPERTDuration(task) {
        // PERT Beta distribution parameters
        const a = parseFloat(task.optimisticDuration) || 0;    // Optimistic
        const m = parseFloat(task.mostLikelyDuration) || a;    // Most likely
        const b = parseFloat(task.pessimisticDuration) || a;   // Pessimistic
        
        console.log('PERT Calculation:', { 
            optimistic: a, 
            mostLikely: m, 
            pessimistic: b 
        });
        
        // Generate a beta-distributed random number
        const u1 = Math.random();
        const u2 = Math.random();
        
        // Box-Muller transform to get normal distribution
        const z = Math.sqrt(-2 * Math.log(u1)) * Math.cos(2 * Math.PI * u2);
        
        // Convert to beta distribution using mean and standard deviation
        const mean = (a + 4*m + b) / 6;
        const sd = (b - a) / 6;
        
        console.log('PERT Calculation Details:', { 
            mean, 
            standardDeviation: sd, 
            zScore: z 
        });
        
        // Apply the transformation
        let duration = mean + (z * sd);
        
        // Clamp the result between optimistic and pessimistic
        const finalDuration = Math.min(Math.max(duration, a), b);
        
        console.log('Final Duration:', finalDuration);
        
        return finalDuration;
    }

    runSimulation() {
        const simulationCount = parseInt(document.getElementById('simulationCount').value);
        
        // Show results container
        document.getElementById('simulationResults').style.display = 'block';
        
        // Run simulation
        this.simulationResults = [];
        const startDate = new Date('2024-12-07');

        for (let i = 0; i < simulationCount; i++) {
            let totalDuration = 0;
            this.tasks.forEach(task => {
                // Use PERT distribution instead of uniform
                totalDuration += this.getPERTDuration(task);
            });
            this.simulationResults.push(totalDuration);
        }

        // Sort results for percentile calculations
        const sortedResults = [...this.simulationResults].sort((a, b) => a - b);
        
        // Calculate confidence levels with adjusted percentiles
        const p20Index = Math.floor(sortedResults.length * 0.2);
        const p50Index = Math.floor(sortedResults.length * 0.5);
        const p90Index = Math.floor(sortedResults.length * 0.9);
        
        const confidenceLevels = {
            p20: sortedResults[p20Index],
            p50: sortedResults[p50Index],
            p90: sortedResults[p90Index],
            p20Date: new Date(startDate.getTime() + (sortedResults[p20Index] * 24 * 60 * 60 * 1000)),
            p50Date: new Date(startDate.getTime() + (sortedResults[p50Index] * 24 * 60 * 60 * 1000)),
            p90Date: new Date(startDate.getTime() + (sortedResults[p90Index] * 24 * 60 * 60 * 1000))
        };

        // Update confidence boxes
        this.updateSimulationResults();
        
        // Create histogram bins
        const binCount = 10;
        const min = Math.min(...sortedResults);
        const max = Math.max(...sortedResults);
        const binSize = (max - min) / binCount;

        // Create bins
        const bins = new Array(binCount).fill(0);
        const binLabels = [];

        for (let i = 0; i < binCount; i++) {
            const binStart = min + (i * binSize);
            const binEnd = binStart + binSize;
            binLabels.push(`${Math.round(binStart)}-${Math.round(binEnd)} days`);
        }

        // Populate bins
        sortedResults.forEach(value => {
            const binIndex = Math.min(
                Math.floor((value - min) / binSize), 
                binCount - 1
            );
            bins[binIndex]++;
        });

        // Create chart with confidence levels
        window.createChart('simulationChart', binLabels, bins, confidenceLevels);
    }

    updateSimulationResults() {
        const startDate = new Date('2024-12-07');
        
        // Sort the results to get correct percentile values
        const sortedResults = [...this.simulationResults].sort((a, b) => a - b);
        
        const results = {
            p20: sortedResults[Math.floor(sortedResults.length * 0.2)],
            p50: sortedResults[Math.floor(sortedResults.length * 0.5)],
            p90: sortedResults[Math.floor(sortedResults.length * 0.9)]
        };

        console.log('Simulation Results:', {
            rawResults: this.simulationResults,
            sortedResults: sortedResults,
            percentiles: results
        });

        // Calculate end dates
        const p20Date = new Date(startDate);
        p20Date.setDate(p20Date.getDate() + Math.ceil(results.p20));
        
        const p50Date = new Date(startDate);
        p50Date.setDate(p50Date.getDate() + Math.ceil(results.p50));
        
        const p90Date = new Date(startDate);
        p90Date.setDate(p90Date.getDate() + Math.ceil(results.p90));

        // Update confidence boxes
        const confidenceBoxes = document.querySelectorAll('.confidence-box');
        
        // 20% Confidence Box
        confidenceBoxes[0].querySelector('.days-result').textContent = `${Math.ceil(results.p20)} days`;
        confidenceBoxes[0].querySelector('small').textContent = `End: ${p20Date.toLocaleDateString('en-US', { 
            month: 'short', 
            day: 'numeric', 
            year: 'numeric' 
        })}`;

        // 50% Confidence Box
        confidenceBoxes[1].querySelector('.days-result').textContent = `${Math.ceil(results.p50)} days`;
        confidenceBoxes[1].querySelector('small').textContent = `End: ${p50Date.toLocaleDateString('en-US', { 
            month: 'short', 
            day: 'numeric', 
            year: 'numeric' 
        })}`;

        // 90% Confidence Box
        confidenceBoxes[2].querySelector('.days-result').textContent = `${Math.ceil(results.p90)} days`;
        confidenceBoxes[2].querySelector('small').textContent = `End: ${p90Date.toLocaleDateString('en-US', { 
            month: 'short', 
            day: 'numeric', 
            year: 'numeric' 
        })}`;

        console.log('Confidence Levels:', {
            p20: { days: Math.ceil(results.p20), date: p20Date },
            p50: { days: Math.ceil(results.p50), date: p50Date },
            p90: { days: Math.ceil(results.p90), date: p90Date }
        });
    }

    handleExcelImport(event) {
        const file = event.target.files[0];
        const reader = new FileReader();

        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet);

            this.tasks = jsonData.map(row => new Task(
                row['Task Name'],
                row['Resource Name'],
                row['Connected Task'] || 'None',
                parseInt(row['Duration (days)']),
                row['Start Date']
            ));

            this.updateTaskTable();
            this.updateConnectedTaskDropdown();
        };

        reader.readAsArrayBuffer(file);
    }

    downloadTemplate() {
        const template = [
            {
                'Task Name': 'Example Task',
                'Resource Name': 'Resource 1',
                'Connected Task': 'None',
                'Duration (days)': 5,
                'Start Date': '2024-12-07'
            }
        ];

        const ws = XLSX.utils.json_to_sheet(template);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Tasks');
        XLSX.writeFile(wb, 'task_template.xlsx');
    }

    exportToExcel() {
        const data = this.tasks.map(task => ({
            'Task Name': task.name,
            'Resource Name': task.resource,
            'Connected Task': task.connectedTask,
            'Duration (days)': task.mostLikelyDuration,
            'Start Date': task.startDate,
            'Optimistic': task.optimisticDuration,
            'Pessimistic': task.pessimisticDuration,
            'PERT Estimate': task.pertEstimate
        }));

        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Tasks');
        XLSX.writeFile(wb, 'project_tasks.xlsx');
    }

    refreshDurations() {
        this.tasks.forEach(task => {
            task.optimisticDuration = Math.max(1, Math.floor(task.mostLikelyDuration * 0.6));
            task.pessimisticDuration = Math.ceil(task.mostLikelyDuration * 1.6);
            task.updatePertEstimate();
        });
        this.updateTaskTable();
    }
}

// Initialize the application
document.addEventListener('DOMContentLoaded', () => {
    window.projectSimulator = new ProjectSimulator();
});
