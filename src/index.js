(function () {

    function jsonToInsertSQL(jsonData, tableName, { valuesClause }) {
        if (!jsonData || jsonData.length === 0)
            return '';
        const columns = Object.keys(jsonData[0]);
        const colStr = columns.join(', ');

        if (valuesClause) {
            const valuesStr = jsonData.map(row => {
                const vals = formatRowValues(row, columns);
                return `(${vals.join(', ')})`;
            }).join(',\n');
            return `insert into ${tableName} (${colStr}) values ${valuesStr};`;
        }
        const sqlLines = jsonData.map(row => {
            const vals = formatRowValues(row, columns);
            return `insert into ${tableName} (${colStr}) values (${vals.join(', ')});`;
        });
        return sqlLines.join('\n');
    }

    function formatRowValues(row, columns) {
        return columns.map(col => {
            const val = row[col];
            if (val === null || val === undefined)
                return 'NULL';
            if (typeof val === 'string')
                return `N'${val.replace(/'/g, "''")}'`;
            return val;
        });
    }

    async function handleFile(file) {
        const data = new Uint8Array(await file.arrayBuffer());
        const workbook = XLSX.read(data, { type: 'array' });
        const sb = [];
        const options = {
            valuesClause: document.getElementById('valuesClause').checked
        };
        for (let sheetName of Object.keys(workbook.Sheets)) {
            const sheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(sheet);
            const sql = jsonToInsertSQL(jsonData, sheetName, options);
            sb.push(sql);
        }
        document.getElementById('output').textContent = sb.join("\n\n");
    }

    document.addEventListener('DOMContentLoaded', () => {

        // 檔案選取
        document.getElementById('excel-file').addEventListener('change', function (e) {
            const file = e.target.files[0];
            if (!file)
                return;
            handleFile(file);
        });

        // 複製
        document.getElementById('copy-btn').addEventListener('click', function () {
            const output = document.getElementById('output').textContent;
            if (!output || output.trim() === '')
                return;
            navigator.clipboard.writeText(output)
                .then(() => {
                    this.textContent = 'copied!';
                    setTimeout(() => { this.textContent = 'copy SQL'; }, 1200);
                })
                .catch(() => {
                    alert('copy failed, please manually select and copy the content.');
                });
        });

        // 拖曳
        const uploadSection = document.getElementById('upload-section');
        const handleEnter = (e) => {
            e.preventDefault();
            e.stopPropagation();
            uploadSection.style.background = '#e0f7fa';
            uploadSection.style.borderColor = '#00bcd4';
        };
        const handleLeave = (e) => {
            e.preventDefault();
            e.stopPropagation();
            uploadSection.style.background = '';
            uploadSection.style.borderColor = '#ccc';
        };
        uploadSection.addEventListener('dragenter', handleEnter);
        uploadSection.addEventListener('dragover', handleEnter);
        uploadSection.addEventListener('dragleave', handleLeave);
        uploadSection.addEventListener('drop', e => {
            handleLeave(e);
            const files = e.dataTransfer.files;
            if (files && files.length > 0)
                handleFile(files[0]);
        });
    });

})();
