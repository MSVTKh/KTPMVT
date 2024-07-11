document.getElementById('searchButton').addEventListener('click', function() {
    // Fetch data from the Excel sheet
    fetch('data.xlsx')
        .then(response => response.arrayBuffer())
        .then(data => {
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);

            // Get search values
            const searchValues = {
                hming: document.getElementById('hming').value.trim().toLowerCase(),
                chhungkawHotu: document.getElementById('chhungkawHotu').value.trim().toLowerCase(),
                upaBial: document.getElementById('upaBial').value.trim().toLowerCase(),
                group: document.getElementById('group').value.trim().toLowerCase(),
                committeeAwmna: document.getElementById('committeeAwmna').value.trim().toLowerCase(),
                chanvoNeihTheih: document.getElementById('chanvoNeihTheih').value.trim().toLowerCase(),
                awmnaHmun: document.getElementById('awmnaHmun').value.trim().toLowerCase()
            };

            // Filter data based on search values
            const filteredData = jsonData.filter(row => {
                return (!searchValues.hming || (row.HMING || '').toLowerCase().includes(searchValues.hming)) &&
                       (!searchValues.chhungkawHotu || (row['CHHUNGKAW HOTU'] || '').toLowerCase().includes(searchValues.chhungkawHotu)) &&
                       (!searchValues.upaBial || (row['UPA BIAL'] || '').toLowerCase().includes(searchValues.upaBial)) &&
                       (!searchValues.group || (row.GROUP || '').toLowerCase().includes(searchValues.group)) &&
                       (!searchValues.committeeAwmna || (row['COMMITTEE AWMNA'] || '').toLowerCase().includes(searchValues.committeeAwmna)) &&
                       (!searchValues.chanvoNeihTheih || (row['CHANVO NEIH THEIH'] || '').toLowerCase().includes(searchValues.chanvoNeihTheih)) &&
                       (!searchValues.awmnaHmun || (row['AWMNA HMUN'] || '').toLowerCase().includes(searchValues.awmnaHmun));
            });

            // Display results
            const resultsTableBody = document.querySelector('#resultsTable tbody');
            resultsTableBody.innerHTML = '';

            if (filteredData.length === 0) {
                resultsTableBody.innerHTML = '<tr><td colspan="7">I mi zawn chu a awm tlat lo mai</td></tr>';
            } else {
                filteredData.forEach(row => {
                    const tr = document.createElement('tr');
                    tr.innerHTML = `
                        <td>${row.HMING || ''}</td>
                        <td>${row['CHHUNGKAW HOTU'] || ''}</td>
                        <td>${row['UPA BIAL'] || ''}</td>
                        <td>${row.GROUP || ''}</td>
                        <td>${row['COMMITTEE AWMNA'] || ''}</td>
                        <td>${row['CHANVO NEIH THEIH'] || ''}</td>
                        <td>${row['AWMNA HMUN'] || ''}</td>
                    `;
                    resultsTableBody.appendChild(tr);
                });
            }
        })
        .catch(error => console.error('Error fetching data:', error));
});
